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

tesseract_path = r"C:/Program Files/Tesseract-OCR/tesseract.exe"
pytesseract.tesseract_cmd = tesseract_path

DRIVE_ID = os.environ["SHAREPOINT_DRIVE_ID"]
DRIVE = f'https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/root:/'


class SharePointClient:
    def __init__(self):
        self.tenant_id = os.environ["AZURE_TENANT_ID"]
        self.client_id = os.environ["AZURE_CLIENT_ID"]
        self.client_secret = os.environ["AZURE_CLIENT_SECRET"]
        self.token = self._acquire_token()
        self.headers = {"Authorization": f"Bearer {self.token}"}

    def _acquire_token(self):
        app = ConfidentialClientApplication(
            client_id=self.client_id,
            client_credential=self.client_secret,
            authority=f"https://login.microsoftonline.com/{self.tenant_id}",
        )
        scopes = ["https://graph.microsoft.com/.default"]
        result = app.acquire_token_for_client(scopes=scopes)
        if "access_token" not in result:
            err_desc = result.get('error_description')
            raise Exception(f"Failed to acquire token: {err_desc}")
        return result['access_token']

    def download_excel(self, filename):
        url = f"{DRIVE}{filename}:/content"
        resp = requests.get(url, headers=self.headers)
        resp.raise_for_status()
        return pd.read_excel(
            BytesIO(resp.content),
            dtype={"asset_id": "Int64"},
            parse_dates=['service_date', 'next_service_date']
        )

    def upload_excel(self, df, filename):
        buffer = BytesIO()
        df.to_excel(buffer, index=False)
        url = f"{DRIVE}{filename}:/content"
        attempts = 0
        wait = 5
        while attempts < 5:
            try:
                buffer.seek(0)
                resp = requests.put(url, headers=self.headers, data=buffer)
                resp.raise_for_status()
                logging.info("Successfully updated on SharePoint.")
                break
            except requests.exceptions.RequestException as e:
                logging.error(f"Error uploading file: {e}")
                wait *= 2
                time.sleep(wait)

    def get_qualer_token(self):
        url = f"{DRIVE}General/apikey.txt:/content"
        resp = requests.get(url, headers=self.headers)
        resp.raise_for_status()
        return resp.text.strip()


class ExtendedAssetServiceRecordsApi(AssetServiceRecordsApi):
    def get_latest(cls, asset_id):
        records = cls.asset_service_records_get_asset_service_records_by_asset(
            asset_id=asset_id)
        return max(records, key=lambda x: x.service_date) if records else None


class ExtendedServiceOrderDocumentsApi(ServiceOrderDocumentsApi):
    patn = r"The above expendable wireset was made from wire roll\s+(.*?)\.\s"

    def retrieve_wire_roll_SN(self, guid):
        pdf = self.service_order_documents_get_document(guid=guid,
                                                        _preload_content=False)
        if not pdf:
            raise ValueError(f"Failed to retrieve document with GUID: {guid}")
        images = convert_from_bytes(pdf.data, dpi=300)
        for img in images:
            text = image_to_string(img)
            if match := re.search(self.patn, text, re.IGNORECASE | re.DOTALL):
                return match.group(1).strip()
        return None


class WireRollUpdater:
    def __init__(self):
        initialize_logging()
        self.sp_client = SharePointClient()
        self.config = Configuration()
        self.config.host = "https://jgiquality.qualer.com"
        self.client = ApiClient(configuration=self.config)
        qualer_token = self.sp_client.get_qualer_token()
        self.client.default_headers["Authorization"] = qualer_token
        self.asr_api = ExtendedAssetServiceRecordsApi(self.client)
        self.soi_api = ServiceOrderItemsApi(self.client)
        self.sod_api = ExtendedServiceOrderDocumentsApi(self.client)

    def update_wire_rolls(self):
        df = self.sp_client.download_excel("Pyro/WireSetCerts.xlsx")
        before_hash = self.hash_df(df.copy())
        start_time = time.time()

        for idx, row in tqdm(df.iterrows(),
                             total=len(df),
                             desc="Processing assets",
                             file=sys.stdout,
                             dynamic_ncols=True):
            adi = row.get("asset_id")
            if pd.isna(adi):
                continue
            latest = self.asr_api.get_latest(adi)
            if not latest:
                tqdm.write(f"No service records found for asset ID: {adi}")
                continue
            if str(latest.asset_tag) != row.get("asset_tag"):
                tqdm.write(f"Overwriting Asset tag for asset ID: {adi}.")
            if str(latest.serial_number) != row.get("serial_number"):
                tqdm.write(f"Overwriting Serial Number for asset ID: {adi}.")
            dates_match = row["service_date"] == latest.service_date
            if pd.notna(row["service_date"]) and dates_match:
                continue

            df.at[idx, "wire_roll_cert_number"] = None
            df.at[idx, "serial_number"] = latest.serial_number
            df.at[idx, "asset_tag"] = latest.asset_tag
            df.at[idx, "custom_order_number"] = latest.custom_order_number
            df.at[idx, "service_date"] = latest.service_date
            df.at[idx, "next_service_date"] = latest.next_service_date

            service_order_items = self.get_service_order_items(latest)
            service_order_id = None
            for item in service_order_items:
                if int(item.asset_id) == int(adi):
                    service_order_id = item.service_order_id
                    df.at[idx, 'certificate_number'] = item.certificate_number
                    break

            if not service_order_id:
                tqdm.write(f"No matching work item for asset ID: {adi}")
                continue

            order_documents = self.get_work_order_documents(service_order_id)
            prefix = latest.asset_tag.replace(" ", "")
            for doc in order_documents:
                if doc.document_name.startswith(prefix) and is_pdf(doc):
                    WRollSN = self.sod_api.retrieve_wire_roll_SN(doc.guid)
                    df.at[idx, 'wire_roll_cert_number'] = WRollSN
                    break
            else:
                tqdm.write(f"No certificate found for asset ID: {adi}")

        if self.hash_df(df) != before_hash:
            self.sp_client.upload_excel(df, "Pyro/WireSetCerts.xlsx")
        else:
            logging.info("No changes detected; skipping upload.")

        duration = time.time() - start_time
        m, s = divmod(duration, 60)
        logging.debug(f"Script completed in {int(m)}m {int(s)}s.")

    def get_work_order_documents(self, service_order_id):
        return self.sod_api.service_order_documents_get_documents_list(
            service_order_id=service_order_id)

    def get_service_order_items(self, latest):
        return self.soi_api.service_order_items_get_work_items_0(
            work_item_number=latest.custom_order_number)

    @staticmethod
    def hash_df(df):
        return hashlib.md5(
            pd.util.hash_pandas_object(df, index=True).values
            ).hexdigest()


def is_pdf(doc):
    return doc.document_name.endswith('.pdf')


def initialize_logging():
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


if __name__ == "__main__":
    updater = WireRollUpdater()
    while time.localtime().tm_hour < 17:
        logging.info(f"Starting update at {time.strftime('%Y-%m-%d %H:%M')}")
        try:
            updater.update_wire_rolls()
        except Exception as e:
            logging.error(f"An error occurred: {e}")
        time.sleep(600)
    logging.info("Script finished running at 5 PM.")
