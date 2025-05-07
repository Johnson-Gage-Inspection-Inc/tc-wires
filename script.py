
from qualer_sdk import Configuration, ApiClient, AssetsApi, AssetServiceRecordsApi  # noqa: E501
import os
import pandas as pd

token = os.environ.get('QUALER_API_KEY')
print(f"Using token: {repr(token)}")
config = Configuration()
config.host = "https://jgiquality.qualer.com"

client = ApiClient(configuration=config)
client.default_headers["Authorization"] = f"Api-Token {token}"

assets_api = AssetsApi(client)
asset_service_records_api = AssetServiceRecordsApi(client)

# List of asset IDs to collect
asset_ids = [
    1235400, 1235401, 1235402, 1235498, 1235500, 1235502,
    1235504, 1235505, 1235506, 1235507, 1235508, 1235509,
    1235510, 1235526, 1235646, 1235659, 1235660, 1235661,
    1235673, 1235704, 1235770, 1235777, 2635568
]

# Collect the assets
assets_api.assets_collect_assets(asset_ids)
collected_assets = assets_api.assets_get_asset_manager_list(
    model_filter_type="CollectedAssets"
    )

df = pd.DataFrame([asset.to_dict() for asset in collected_assets])

# Example: Print asset names
for index, row in df.iterrows():
    print(f"Asset ID: {row['asset_id']}, Name: {row['asset_name']}")

    # Get asset service records
    service_records = asset_service_records_api.asset_service_records_get_asset_service_records_by_asset(asset_id=row['asset_id'])  # noqa: E501
    latest_service_record = service_records[-1] if service_records else None
    pass