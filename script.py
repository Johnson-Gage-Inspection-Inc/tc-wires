from qualer_sdk import ApiClient, AssetsApi, AssetServiceRecordsApi
from qualer_sdk import Configuration, ServiceOrderItemDocumentsApi
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
for index, row in tqdm(df.iterrows(), total=len(df), desc="Processing assets", unit="asset", **tqdm_kwargs):  # noqa: E501
    tqdm.write(f"Asset ID: {row['asset_id']}, Name: {row['asset_name']}")

    print(f"Asset ID: {row['asset_id']}, Name: {row['asset_name']}")

    # Get asset service records
    service_records = asset_service_records_api.asset_service_records_get_asset_service_records_by_asset(asset_id=row['asset_id'])  # noqa: E501
    latest_service_record = service_records[-1] if service_records else None
    records.append(latest_service_record)

records_df = pd.DataFrame([record.to_dict() for record in records])

# Remove columns that are all None
records_df = records_df.dropna(axis=1, how='all')
# Remove columns that are have no variance

records_df = records_df.loc[:, (records_df != records_df.iloc[0]).any()]

# Save the DataFrame to a CSV file
records_df.to_csv('service_records.csv', index=False)
