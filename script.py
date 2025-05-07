from qualer_sdk import ApiClient, AssetsApi, AssetServiceRecordsApi
from qualer_sdk import Configuration, rest, ServiceOrderItemDocumentsApi
from qualer_sdk import ServiceOrderItemsApi, ServiceOrderDocumentsApi
import os
import pandas as pd
from tqdm import tqdm
import sys
tqdm_kwargs = {'file': sys.stdout}  # Ensure it flushes to visible console


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
asset_ids = [  # FIXME: Add High Temp wires (J, K, Ultra K, Ultra N)
    1235400, 1235401, 1235402, 1235498, 1235500, 1235502,
    1235504, 1235505, 1235506, 1235507, 1235508, 1235509,
    1235510, 1235526, 1235646, 1235659, 1235660, 1235661,
    1235673, 1235704, 1235770, 1235777, 2635568
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
for index, row in tqdm(df.iterrows(), total=len(df), desc="Processing assets", unit="asset", **tqdm_kwargs):  # noqa: E501
    tqdm.write(f"Asset ID: {row['asset_id']}, Name: {row['asset_name']}")

    print(f"Asset ID: {row['asset_id']}, Name: {row['asset_name']}")

    # Get asset service records
    service_records = asset_service_records_api.asset_service_records_get_asset_service_records_by_asset(asset_id=row['asset_id'])  # noqa: E501
    latest_service_record = service_records[-1] if service_records else None
    # document_list = asset_service_records_api.asset_service_records_document_list(  # noqa: E501
    #     asset_service_record_id=latest_service_record.asset_service_record_id,
    # )
    # assert document_list == [], "Whoa! We have documents in the service record. This is not expected."  # noqa: E501
    keepsies = latest_service_record.to_dict()

    try:
        service_order_items = service_order_items_api.service_order_items_get_work_items_0(  # noqa: E501
            work_item_number=latest_service_record.certificate_number
        )
        if len(service_order_items) == 0:
            raise Exception("No service order items found")
        service_order_item = service_order_items[0]
    except (rest.ApiException, Exception):
        try:
            service_order_items = service_order_items_api.service_order_items_get_work_items_0(  # noqa: E501
                work_item_number=latest_service_record.custom_order_number
            )
            if len(service_order_items) == 1:
                service_order_item = service_order_items[0]
            else:
                pass
        except rest.ApiException as e:
            print(f"Error: {e}")
            continue

    documents_list = service_order_documents_api.service_order_documents_get_documents_list(  # noqa: E501
            service_order_id=service_order_item.service_order_id
        )
    excel_files = []
    for document in documents_list:
        if document.document_name.endswith(".xlsx"):
            excel_files.append(document.document_name)
    if excel_files:
        pass
    else:
        work_item_documents = service_order_item_documents_api.service_order_item_documents_get_documents_list(  # noqa: E501
                service_order_item_id=service_order_item.work_item_id
            )
        for document in work_item_documents:
            if document.document_name.endswith(".xlsx"):
                excel_files.append(document.document_name)
    if excel_files:
        pass
    records.append(keepsies)

records_df = pd.DataFrame([record.to_dict() for record in records])


# Remove columns that are all None
records_df = records_df.dropna(axis=1, how='all')
# Remove columns that are have no variance

records_df = records_df.loc[:, (records_df != records_df.iloc[0]).any()]

# Save the DataFrame to a CSV file
records_df.to_csv('service_records.csv', index=False)
