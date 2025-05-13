from pdf2image import convert_from_bytes
from qualer_sdk import ApiClient, AssetsApi, AssetServiceRecordsApi
from qualer_sdk import Configuration, ServiceOrderItemDocumentsApi
from qualer_sdk import ServiceOrderItemsApi, ServiceOrderDocumentsApi
from tqdm import tqdm
import os
import pandas as pd
import pytesseract
import re
import sys
import time
pytesseract.pytesseract.tesseract_cmd = r"C:/Program Files/Tesseract-OCR/tesseract.exe"  # noqa: E501

start_time = time.time()

token = os.environ.get('QUALER_API_KEY')
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

# List of asset IDs to collect
asset_ids = [
    1235344, 1235555, 1235388, 1235682, 1235426, 1235686, 1235630, 1235561,
    1235543, 1235622, 1235403, 1235569, 1235743, 2550412, 1446161, 2701517,
    2336905, 1235639, 1235626, 1235522, 1235770, 1235767, 1235772, 1235660,
    1235502, 1235505, 1235646, 1235659, 1235504, 1235506, 1235510, 1235590,
    1235489, 1235508, 1235777, 1235400, 1235661, 1235507, 1235500, 1235509,
    1235498, 1235401, 1235402, 1235610, 1235673, 1235526, 2635568, 1235345,
    1235444, 1235564, 1235501, 1235704, 1235563, 1235428, 2822437, 1235598,
    1235517, 2822784, 1235416, 1235503, 1235768, 1235769, 1235540
]

# Collect the assets
assets_api.assets_clear_collected_assets([])
assets_api.assets_collect_assets(asset_ids)
collected_assets = assets_api.assets_get_asset_manager_list(
    model_filter_type="CollectedAssets"
    )
assets_api.assets_clear_collected_assets([])

df = pd.DataFrame([asset.to_dict() for asset in collected_assets])

records = []
tqdm_kwargs = {'file': sys.stdout}  # Ensure it flushes to visible console
for index, row in tqdm(df.iterrows(), total=len(df), desc="Processing assets", unit="asset", **tqdm_kwargs):  # noqa: E501
    asset_id = row['asset_id']
    tqdm.write(f"Asset ID: {asset_id}, Name: {row['asset_name']}")

    # Get asset service records
    service_records = asset_service_records_api.asset_service_records_get_asset_service_records_by_asset(asset_id=asset_id)  # noqa: E501
    if not service_records:
        tqdm.write(f"No service records found for asset ID: {asset_id}")
        continue
    latest_service_record = service_records[-1]
    latest_service_record_date = latest_service_record.service_date
    next_service_date = latest_service_record.next_service_date
    custom_order_number = latest_service_record.custom_order_number
    asset_service_record_id = latest_service_record.asset_service_record_id
    service_order_items = service_order_items_api.service_order_items_get_work_items_0(  # noqa: E501
        work_item_number=custom_order_number,
    )
    for item in service_order_items:
        if item.asset_id == asset_id:
            service_order_item = item
            break
    work_item_id = item.work_item_id
    service_order_id = item.service_order_id
    certificate_number = item.certificate_number

    order_documents = service_order_documents_api.service_order_documents_get_documents_list(  # noqa: E501
        service_order_id=service_order_id,
    )
    certificate_document = None
    for document in order_documents:
        filename = document.document_name
        if filename.startswith(row['asset_tag']) and filename.endswith('.pdf'):
            certificate_document = document
            break

    if not certificate_document:
        tqdm.write(f"No certificate document found for asset ID: {asset_id}")
        continue

    cert_guid = certificate_document.guid

    # This should now work without a UnicodeDecodeError
    certificate_document_pdf = service_order_documents_api.service_order_documents_get_document(  # noqa: E501
        guid=cert_guid, _preload_content=False
    )
    if not certificate_document_pdf:
        tqdm.write(f"No PDF found for asset ID: {asset_id}")
        continue

    # Convert PDF (bytes) to image(s)
    images = convert_from_bytes(certificate_document_pdf.data, dpi=300)

    # OCR each page
    pattern = r"The above expendable wireset was made from wire roll\s+(.*?)\.\s"  # noqa: E501

    for i, img in enumerate(images):
        text = pytesseract.image_to_string(img)
        match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
        if match:
            wire_roll_cert_number = match.group(1).strip()
            tqdm.write(f"OCR match (page {i+1}): {wire_roll_cert_number}")
            records.append({
                "asset_id": asset_id,
                "serial_number": row['serial_number'],
                "asset_tag": row['asset_tag'],
                "asset_service_record_id": asset_service_record_id,
                "custom_order_number": custom_order_number,
                "work_item_id": work_item_id,
                "latest_service": latest_service_record_date,
                "next_service_date": next_service_date,
                "wire_roll_cert_number": wire_roll_cert_number
            })
            break

df = pd.DataFrame(records)
duration = time.time() - start_time
m, s = divmod(duration, 60)
print(f"Script completed in {int(m)} minutes and {int(s)} seconds.")

# Save the DataFrame to a CSV file
df.to_csv('wire_roll_cert_numbers.csv', index=False)

pass
