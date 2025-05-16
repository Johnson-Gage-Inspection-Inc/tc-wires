from dotenv import load_dotenv
from io import BytesIO
from msal import ConfidentialClientApplication
from pdf2image import convert_from_bytes
from pytesseract import pytesseract, image_to_string
from tqdm import tqdm
from qualer_sdk import (
    ApiClient,
    AssetServiceRecordsApi,
    Configuration,
    ServiceOrderItemsApi,
    ServiceOrderDocumentsApi,
)
import hashlib
import logging
import os
import pandas as pd
import re
import requests
import sys
import time

load_dotenv()

DRIVE_ID = os.environ["SHAREPOINT_DRIVE_ID"]
DRIVE = f'https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/root:/'


def initialize_logging():
    """Set up logging to log to a file"""
    log_file_path = os.path.join(os.getcwd(), 'tc-wires.log')
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file_path),
            logging.StreamHandler(sys.stdout)
        ]
    )

    logging.debug(f"Current working directory: {os.getcwd()}")


def hash_df(df):
    return hashlib.md5(
        pd.util.hash_pandas_object(df, index=True).values
        ).hexdigest()


def retrieve_wire_roll_cert_number(cert_guid):
    certificate_document_pdf = SOD_api.service_order_documents_get_document(
        guid=cert_guid, _preload_content=False
        )
    if not certificate_document_pdf:
        raise ValueError(f"Failed to retrieve document with GUID: {cert_guid}")

    images = convert_from_bytes(certificate_document_pdf.data, dpi=300)
    pattern = r"The above expendable wireset was made from wire roll\s+(.*?)\.\s"  # noqa: E501
    for i, img in enumerate(images):
        text = image_to_string(img)
        if match := re.search(pattern, text, re.IGNORECASE | re.DOTALL):
            return match.group(1).strip()


def save_to_sharepoint(df, headers):
    buffer = BytesIO()
    df.to_excel(buffer, index=False)

    url = f"{DRIVE}Pyro/WireSetCerts.xlsx:/content"
    attempts = 0
    max_attempts = 5
    wait = 5  # seconds
    while attempts < max_attempts:
        attempts += 1
        try:
            buffer.seek(0)
            upload_resp = requests.put(url, headers=headers, data=buffer)
            upload_resp.raise_for_status()
        except requests.exceptions.RequestException as e:
            logging.error(f"Error uploading file: {e}")
            wait *= 2  # Exponential backoff
            if attempts >= max_attempts:
                logging.error("Max attempts reached. Exiting.")
                raise
            time.sleep(wait)  # Wait before retrying
        else:
            break

    logging.info("Successfully uploaded updated Excel to SharePoint.")


def perform_lookups():
    """Perform lookups and update the Excel file with wire roll certificate
    numbers."""

    azure_token = acquire_azure_access_token()
    headers = {"Authorization": f"Bearer {azure_token}"}
    # Load existing output if it exists
    download_url = f"{DRIVE}Pyro/WireSetCerts.xlsx:/content"
    resp = requests.get(download_url, headers=headers)
    resp.raise_for_status()

    df = pd.read_excel(BytesIO(resp.content), dtype={"asset_id": "Int64"}, parse_dates=['service_date', 'next_service_date'])  # noqa: E501

    before_hash = hash_df(df.copy())

    start_time = time.time()

    for idx, row in tqdm(df.iterrows(),
                         total=len(df),
                         desc="Processing assets",
                         **{'file': sys.stdout}):
        asset_id = row.get("asset_id")
        if pd.isna(asset_id):
            continue

        # Get latest service record
        ASR_api = AssetServiceRecordsApi(client)
        service_records = ASR_api.asset_service_records_get_asset_service_records_by_asset(asset_id=asset_id)  # noqa: E501
        if not service_records:
            tqdm.write(f"No service records found for asset ID: {asset_id}")
            continue

        latest = service_records[-1]

        if str(latest.asset_tag) != row.get("asset_tag"):
            tqdm.write(f"Overwriting Asset tag for asset ID: {asset_id}.")

        if str(latest.serial_number) != row.get("serial_number"):
            tqdm.write(f"Overwriting Serial Number for asset ID: {asset_id}.")

        existing_date = row["service_date"]
        if pd.notna(existing_date) and existing_date == latest.service_date:
            continue

        df.at[idx, "wire_roll_cert_number"] = None
        df.at[idx, "serial_number"] = latest.serial_number
        df.at[idx, "asset_tag"] = latest.asset_tag
        df.at[idx, "custom_order_number"] = latest.custom_order_number
        df.at[idx, "service_date"] = latest.service_date
        df.at[idx, "next_service_date"] = latest.next_service_date

        # Find related service order item
        SOI_api = ServiceOrderItemsApi(client)
        service_order_items = SOI_api.service_order_items_get_work_items_0(
            work_item_number=latest.custom_order_number,
        )
        service_order_id = None
        for item in service_order_items:
            if int(item.asset_id) == int(asset_id):
                service_order_id = item.service_order_id
                df.at[idx, 'certificate_number'] = item.certificate_number
                break

        if not service_order_id:
            tqdm.write(f"No matching work item for asset ID: {asset_id}")
            continue

        # Find certificate document
        order_documents = SOD_api.service_order_documents_get_documents_list(
            service_order_id=service_order_id
            )
        certificate_document = None
        for document in order_documents:
            prefix = latest.asset_tag.replace(" ", "")
            if (document.document_name.startswith(prefix)
                    and document.document_name.endswith('.pdf')):
                certificate_document = document
                break

        if certificate_document:
            WRollSN = retrieve_wire_roll_cert_number(certificate_document.guid)
            df.at[idx, 'wire_roll_cert_number'] = WRollSN
        else:
            tqdm.write(f"No certificate  found for asset ID: {asset_id}")

    after_hash = hash_df(df)
    if before_hash != after_hash:
        save_to_sharepoint(df, headers)
    else:
        logging.info("No changes detected; skipping upload.")

    duration = time.time() - start_time
    m, s = divmod(duration, 60)
    logging.debug(
        f"Script completed in {int(m)} minutes and {int(s)} seconds."
        )


def get_qualer_token():
    azure_token = acquire_azure_access_token()
    headers = {"Authorization": f"Bearer {azure_token}"}
    url = f"{DRIVE}General/apikey.txt:/content"
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()
    return resp.text.strip()


def acquire_azure_access_token():
    TENANT = os.environ["AZURE_TENANT_ID"]
    app = ConfidentialClientApplication(
        client_id=os.environ["AZURE_CLIENT_ID"],
        client_credential=os.environ["AZURE_CLIENT_SECRET"],
        authority=f"https://login.microsoftonline.com/{TENANT}",
        )
    result = app.acquire_token_for_client(scopes=[
        "https://graph.microsoft.com/.default",
    ])
    if "access_token" not in result:
        error_description = result.get('error_description')
        raise Exception(f"Failed to acquire token: {error_description}")
    return result['access_token']


if __name__ == "__main__":
    initialize_logging()
    tesseract_path = r"C:/Program Files/Tesseract-OCR/tesseract.exe"
    pytesseract.tesseract_cmd = tesseract_path

    config = Configuration()
    config.host = "https://jgiquality.qualer.com"

    client = ApiClient(configuration=config)
    client.default_headers["Authorization"] = get_qualer_token()

    SOD_api = ServiceOrderDocumentsApi(client)

    # Loop until 5 PM
    while time.localtime().tm_hour < 17:
        current_timestamp = time.strftime('%Y-%m-%d %H:%M:%S')
        logging.info(f"Starting update at {current_timestamp}")
        try:
            perform_lookups()
        except Exception as e:
            logging.error(f"An error occurred: {e}")
        # wait for 10 minutes
        time.sleep(600)
    logging.info("Script finished running at 5 PM.")
