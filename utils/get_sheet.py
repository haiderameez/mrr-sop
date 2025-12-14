import os
import json

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv

from utils.load_sheet_id import get_sheet_ids

load_dotenv()

credentials_json = os.getenv("GOOGLE_CREDENTIALS_JSON")
creds_dict = json.loads(credentials_json)

scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
credentials = Credentials.from_service_account_info(creds_dict, scopes=scopes)
client = gspread.authorize(credentials)

def get_all_sheets(sheet_id):
    sh = client.open_by_key(sheet_id)
    worksheets = sh.worksheets()
    dataframes = {}
    for ws in worksheets:
        values = ws.get_all_values()
        headers = []
        for i, h in enumerate(values[0]):
            headers.append(h if h.strip() != "" else f"col_{i}")
        df = pd.DataFrame(values[1:], columns=headers)
        dataframes[ws.title] = df
    return dataframes

def get_sheet():
    invoice_sheet_id, mis_sheet_id, master_sheet_id = get_sheet_ids()

    invoice_sheets = get_all_sheets(invoice_sheet_id)
    mis_sheets = get_all_sheets(mis_sheet_id)
    master_sheets = get_all_sheets(master_sheet_id)

    return invoice_sheets, mis_sheets, master_sheets