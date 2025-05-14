from dotenv import load_dotenv
from pdf2image import convert_from_bytes
from qualer_sdk import (
    ApiClient,
    AssetsApi,
    AssetServiceRecordsApi,
    Configuration,
    ServiceOrderItemDocumentsApi,
    ServiceOrderItemsApi,
    ServiceOrderDocumentsApi,
)
from tqdm import tqdm
import os
import pandas as pd
import pytesseract
import re
import sys
import time


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


tesseract_path = r"C:/Program Files/Tesseract-OCR/tesseract.exe"
pytesseract.pytesseract.tesseract_cmd = tesseract_path

start_time = time.time()
output_path = 'wire_roll_cert_numbers.csv'

load_dotenv()

token = os.environ.get('QUALER_API_KEY')
if not token:
    print("Please set the QUALER_API_KEY environment variable.")
    sys.exit(1)

print(f"Using token: {repr(token)}")
config = Configuration()
config.host = "https://jgiquality.qualer.com"

client = ApiClient(configuration=config)
client.default_headers["Authorization"] = f"Api-Token {token}"

assets_api = AssetsApi(client)
asset_service_records_api = AssetServiceRecordsApi(client)
service_order_items_api = ServiceOrderItemsApi(client)
service_order_documents_api = ServiceOrderDocumentsApi(client)
service_order_item_documents_api = ServiceOrderItemDocumentsApi(client)

# Load existing output if it exists
existing_df = pd.read_csv(output_path, dtype={"asset_id": "Int64"}, parse_dates=['service_date', 'next_service_date'])  # noqa: E501

tqdm_kwargs = {'file': sys.stdout}

for idx, row in tqdm(existing_df.iterrows(), total=len(existing_df), desc="Processing assets", **tqdm_kwargs):  # noqa: E501
    asset_id = row.get("asset_id")
    if pd.isna(asset_id):
        continue
    asset_id = int(asset_id)

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

    # Check against existing data
    existing_date = row["service_date"]  # already a Timestamp or NaT
    if pd.notna(existing_date) and existing_date == latest.service_date:
        continue  # skip unchanged

    # Clear wire_roll_cert_number
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
        if int(item.asset_id) == asset_id:
            service_order_id = item.service_order_id
            existing_df.at[idx, 'certificate_number'] = item.certificate_number
            break

    if not service_order_id:
        tqdm.write(f"No matching service order item for asset ID: {asset_id}")
        existing_df.to_csv(output_path, index=False)
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
        tqdm.write(f"No certificate document found for asset ID: {asset_id}")
    existing_df.to_csv(output_path, index=False)

duration = time.time() - start_time
m, s = divmod(duration, 60)
print(f"Script completed in {int(m)} minutes and {int(s)} seconds.")
