import hashlib
import logging
import os
import re
import sys
import time
import uuid
from io import BytesIO

import pandas as pd
import requests
from dotenv import load_dotenv
from msal import ConfidentialClientApplication
from pdf2image import convert_from_bytes
from pytesseract import image_to_string, pytesseract
from qualer_sdk.client import AuthenticatedClient
from qualer_sdk.api.client_asset_service_records import (
    get_asset_service_records_by_asset_get_2,
)
from qualer_sdk.api.service_order_items import get_work_items_workitems
from qualer_sdk.api.service_order_documents import (
    get_documents_list,
    get_document,
)
from tqdm import tqdm

load_dotenv()

DRIVE_ID = os.environ["SHAREPOINT_DRIVE_ID"]
DRIVE = f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/root:/"


def get_latest_service_record(asset_id, client):
    """Get the latest service record for a given asset ID."""
    records = get_asset_service_records_by_asset_get_2.sync(
        asset_id=asset_id, client=client
    )
    if not records:
        return None
    return max(records, key=lambda x: x.service_date)


def initialize_logging():
    """Set up logging to log to a file"""
    log_file_path = os.path.join(os.getcwd(), "tc-wires.log")
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        handlers=[
            logging.FileHandler(log_file_path),
            logging.StreamHandler(sys.stdout),
        ],
    )

    logging.debug(f"Current working directory: {os.getcwd()}")


def hash_df(df):
    return hashlib.md5(
        pd.util.hash_pandas_object(df, index=True).values
    ).hexdigest()


def retrieve_wire_roll_SN(client, cert_guid):
    """
    Retrieve the serial number of a wire roll from a certificate document.

    Parameters:
        client: The authenticated client for API calls.
        cert_guid (str): The GUID of the certificate document to retrieve.

    Returns:
        str: The serial number of the wire roll if found in the document.

    Raises:
        ValueError: If the document cannot be retrieved or the serial number
            cannot be found in the document.
    """
    # Convert string GUID to UUID
    cert_uuid = uuid.UUID(cert_guid)
    
    response = get_document.sync(guid=cert_uuid, client=client)
    if not response:
        raise ValueError(f"Failed to retrieve document with GUID: {cert_guid}")

    # The response should be binary content (PDF data)
    images = convert_from_bytes(response, dpi=300)
    patn = r"The above expendable wireset was made from wire roll\s+(.*?)\.\s"
    for i, img in enumerate(images):
        text = image_to_string(img)
        if match := re.search(patn, text, re.IGNORECASE | re.DOTALL):
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


def perform_lookups(client):
    """Perform lookups and update the Excel file with wire roll certificate
    numbers."""

    azure_token = acquire_azure_access_token()
    headers = {"Authorization": f"Bearer {azure_token}"}
    # Load existing output if it exists
    download_url = f"{DRIVE}Pyro/WireSetCerts.xlsx:/content"
    resp = requests.get(download_url, headers=headers)
    resp.raise_for_status()

    df = pd.read_excel(
        BytesIO(resp.content),
        dtype={"asset_id": "Int64"},
        parse_dates=["service_date", "next_service_date"],
    )

    before_hash = hash_df(df.copy())

    start_time = time.time()

    for idx, row in tqdm(
        df.iterrows(),
        total=len(df),
        desc="Processing assets",
        **{"file": sys.stdout},
        dynamic_ncols=True,
    ):
        asset_id = row.get("asset_id")
        if pd.isna(asset_id):
            continue

        if not (latest := get_latest_service_record(asset_id, client)):
            tqdm.write(f"No service records found for asset ID: {asset_id}")
            continue

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
        service_order_items = get_work_items_workitems.sync(
            client=client,
            work_item_number=latest.custom_order_number,
        )
        service_order_id = None
        for item in service_order_items:
            if int(item.asset_id) == int(asset_id):
                service_order_id = item.service_order_id
                df.at[idx, "certificate_number"] = item.certificate_number
                break

        if not service_order_id:
            tqdm.write(f"No matching work item for asset: {asset_id}")
            continue

        # Find certificate document
        order_documents = get_documents_list.sync(
            service_order_id=service_order_id, client=client
        )
        certificate_document = None
        for document in order_documents:
            prefix = latest.asset_tag.replace(" ", "")
            if document.document_name.startswith(
                prefix
            ) and document.document_name.endswith(".pdf"):
                certificate_document = document
                break

        if certificate_document:
            roll_sn = retrieve_wire_roll_SN(client, certificate_document.guid)
            df.at[idx, "wire_roll_cert_number"] = roll_sn
        else:
            tqdm.write(f"No certificate found for asset ID: {asset_id}")

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
    result = app.acquire_token_for_client(
        scopes=[
            "https://graph.microsoft.com/.default",
        ]
    )
    if "access_token" not in result:
        error_description = result.get("error_description")
        raise Exception(f"Failed to acquire token: {error_description}")
    return result["access_token"]


if __name__ == "__main__":
    initialize_logging()
    tesseract_path = r"C:/Program Files/Tesseract-OCR/tesseract.exe"
    pytesseract.tesseract_cmd = tesseract_path

    client = AuthenticatedClient(
        base_url="https://jgiquality.qualer.com",
        token=get_qualer_token()
    )

    # Loop until 5 PM
    while time.localtime().tm_hour < 17:
        current_timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        logging.info(f"Starting update at {current_timestamp}")
        try:
            perform_lookups(client)
        except Exception as e:
            logging.error(f"An error occurred: {e}")
        # wait for 10 minutes
        time.sleep(600)
    logging.info("Script finished running at 5 PM.")
