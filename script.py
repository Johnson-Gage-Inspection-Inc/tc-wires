from dotenv import load_dotenv
from io import BytesIO
from msal import ConfidentialClientApplication
from pdf2image import convert_from_bytes
from tqdm import tqdm
from qualer_sdk import (
    ApiClient,
    AssetsApi,
    AssetServiceRecordsApi,
    Configuration,
    ServiceOrderItemDocumentsApi,
    ServiceOrderItemsApi,
    ServiceOrderDocumentsApi,
)
import hashlib
import logging
import os
import pandas as pd
import pytesseract
import re
import requests
import sys
import time

load_dotenv()

# Set up logging
logging.basicConfig(level=logging.DEBUG,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# DEBUG log the working directory
logging.debug(f"Current working directory: {os.getcwd()}")


def hash_df(df):
    return hashlib.md5(
        pd.util.hash_pandas_object(df, index=True).values
        ).hexdigest()


def retrieve_wire_roll_cert_number(cert_guid):
    certificate_document_pdf = service_order_documents_api.service_order_documents_get_document(  # noqa: E501
        guid=cert_guid, _preload_content=False
        )
    if not certificate_document_pdf:
        raise ValueError(f"Failed to retrieve document with GUID: {cert_guid}")

    images = convert_from_bytes(certificate_document_pdf.data, dpi=300)
    pattern = r"The above expendable wireset was made from wire roll\s+(.*?)\.\s"  # noqa: E501
    for i, img in enumerate(images):
        text = pytesseract.image_to_string(img)
        if match := re.search(pattern, text, re.IGNORECASE | re.DOTALL):
            return match.group(1).strip()


def save_to_sharepoint(existing_df):
    buffer = BytesIO()
    existing_df.to_excel(buffer, index=False)
    buffer.seek(0)

    upload_url = f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/root:/Pyro/WireSetCerts.xlsx:/content"  # noqa: E501
    upload_resp = requests.put(upload_url, headers=headers, data=buffer)
    upload_resp.raise_for_status()

    logging.info("Successfully uploaded updated Excel to SharePoint.")


def perform_lookups():
    """Perform lookups and update the Excel file with wire roll certificate
    numbers."""
    # Load existing output if it exists
    download_url = f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/root:/Pyro/WireSetCerts.xlsx:/content"  # noqa: E501
    resp = requests.get(download_url, headers=headers)
    resp.raise_for_status()

    existing_df = pd.read_excel(BytesIO(resp.content), dtype={"asset_id": "Int64"}, parse_dates=['service_date', 'next_service_date'])  # noqa: E501

    before_hash = hash_df(existing_df.copy())

    start_time = time.time()

    for idx, row in tqdm(existing_df.iterrows(),
                         total=len(existing_df),
                         desc="Processing assets",
                         **{'file': sys.stdout}):
        asset_id = row.get("asset_id")
        if pd.isna(asset_id):
            continue

    # Get latest service record
        service_records = asset_service_records_api.asset_service_records_get_asset_service_records_by_asset(asset_id=asset_id)  # noqa: E501
        if not service_records:
            tqdm.write(f"No service records found for asset ID: {asset_id}")
            continue

        latest = service_records[-1]

        if str(latest.asset_tag) != row.get("asset_tag"):
            tqdm.write(f"Asset tag mismatch for asset ID: {asset_id}. Overwriting.")  # noqa: E501

        if str(latest.serial_number) != row.get("serial_number"):
            tqdm.write(f"Serial Number mismatch for asset ID: {asset_id}. Overwriting.")  # noqa: E501

        existing_date = row["service_date"]
        if pd.notna(existing_date) and existing_date == latest.service_date:
            continue

        existing_df.at[idx, "wire_roll_cert_number"] = None
        existing_df.at[idx, "serial_number"] = latest.serial_number
        existing_df.at[idx, "asset_tag"] = latest.asset_tag
        existing_df.at[idx, "custom_order_number"] = latest.custom_order_number
        existing_df.at[idx, "service_date"] = latest.service_date
        existing_df.at[idx, "next_service_date"] = latest.next_service_date

    # Find related service order item
        service_order_items = service_order_items_api.service_order_items_get_work_items_0(  # noqa: E501
            work_item_number=latest.custom_order_number,
        )
        service_order_id = None
        for item in service_order_items:
            if int(item.asset_id) == int(asset_id):
                service_order_id = item.service_order_id
                existing_df.at[idx, 'certificate_number'] = item.certificate_number  # noqa: E501
                break

        if not service_order_id:
            tqdm.write(f"No matching service order item for asset ID: {asset_id}")  # noqa: E501
            continue

    # Find certificate document
        order_documents = service_order_documents_api.service_order_documents_get_documents_list(service_order_id=service_order_id)  # noqa: E501
        certificate_document = None
        for document in order_documents:
            prefix = latest.asset_tag.replace(" ", "")
            if document.document_name.startswith(prefix) and document.document_name.endswith('.pdf'):  # noqa: E501
                certificate_document = document
                break

        if certificate_document:
            existing_df.at[idx, 'wire_roll_cert_number'] = retrieve_wire_roll_cert_number(certificate_document.guid)  # noqa: E501
        else:
            tqdm.write(f"No certificate document found for asset ID: {asset_id}")  # noqa: E501

    after_hash = hash_df(existing_df)
    if before_hash != after_hash:
        save_to_sharepoint(existing_df)
    else:
        logging.info("No changes detected; skipping upload.")

    duration = time.time() - start_time
    m, s = divmod(duration, 60)
    logging.debug(
        f"Script completed in {int(m)} minutes and {int(s)} seconds."
        )


def get_qualer_token():
    url = f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/root:/General/apikey.txt:/content"  # noqa: E501
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()
    return resp.text.strip()


TENANT = os.environ["AZURE_TENANT_ID"]
CLIENT_ID = os.environ["AZURE_CLIENT_ID"]
CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]
DRIVE_ID = os.environ["SHAREPOINT_DRIVE_ID"]

authority = f"https://login.microsoftonline.com/{TENANT}"
scope = ["https://graph.microsoft.com/.default"]

app = ConfidentialClientApplication(
    client_id=CLIENT_ID,
    client_credential=CLIENT_SECRET,
    authority=authority,
    )
result = app.acquire_token_for_client(scopes=scope)
if "access_token" not in result:
    raise Exception(f"Failed to acquire token: {result.get('error_description')}")  # noqa: E501
headers = {"Authorization": f"Bearer {result['access_token']}"}

tesseract_path = r"C:/Program Files/Tesseract-OCR/tesseract.exe"
pytesseract.pytesseract.tesseract_cmd = tesseract_path

config = Configuration()
config.host = "https://jgiquality.qualer.com"

client = ApiClient(configuration=config)
client.default_headers["Authorization"] = get_qualer_token()

assets_api = AssetsApi(client)
asset_service_records_api = AssetServiceRecordsApi(client)
service_order_items_api = ServiceOrderItemsApi(client)
service_order_documents_api = ServiceOrderDocumentsApi(client)
service_order_item_documents_api = ServiceOrderItemDocumentsApi(client)


if __name__ == "__main__":
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
