import os
import json

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv

from utils.mrr import generate_financial_reports

def sync_mis_sheets_to_gsheet(
    google_sheet_id,
    month_str,
    keep_worksheet="Invoices"
):
    load_dotenv()

    creds_json = os.getenv("GOOGLE_CREDENTIALS_JSON")
    creds_dict = json.loads(creds_json)

    scopes = ["https://www.googleapis.com/auth/spreadsheets"]

    creds = Credentials.from_service_account_info(
        creds_dict,
        scopes=scopes
    )

    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(google_sheet_id)

    # Let generate_financial_reports perform its own sync when provided
    # a google_sheet_id; this avoids divergence between sync behaviors.
    mis_sheets = generate_financial_reports(month_str, google_sheet_id=google_sheet_id)

    existing_worksheets = spreadsheet.worksheets()

    # Delete non-kept worksheets but ensure we never remove the last sheet
    for ws in existing_worksheets:
        if ws.title == keep_worksheet:
            continue
        try:
            # Refresh current worksheets to check live count before deleting
            current_ws = spreadsheet.worksheets()
            if len(current_ws) <= 1:
                # cannot delete the last remaining sheet in a spreadsheet
                break
            spreadsheet.del_worksheet(ws)
        except Exception:
            # ignore deletion errors and continue (robustness)
            continue

    existing_titles = {ws.title for ws in spreadsheet.worksheets()}

    for sheet_name, df in mis_sheets.items():
        if sheet_name in existing_titles:
            continue

        rows, cols = df.shape
        new_ws = spreadsheet.add_worksheet(
            title=sheet_name[:100],
            rows=max(rows + 1, 1),
            cols=max(cols, 1)
        )

        values = [df.columns.tolist()] + df.fillna("").astype(str).values.tolist()
        new_ws.update(values)

    return mis_sheets
