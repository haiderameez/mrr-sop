import os
import json
from datetime import datetime

import pandas as pd
import numpy as np
import gspread
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv
from dateutil.relativedelta import relativedelta

from utils.churnout import churnout


load_dotenv()


def clean_currency(val):
    if pd.isna(val):
        return 0.0
    s = str(val).strip()
    if s in ["", "-"]:
        return 0.0
    try:
        return float(s.replace(",", ""))
    except:
        return 0.0


def find_nature(row, invoice_nature_map):
    invoice_val = str(row["Invoice"])
    for sep in [",", "/", "\n"]:
        invoice_val = invoice_val.replace(sep, ",")
    invoice_list = [x.strip() for x in invoice_val.split(",") if x.strip()]

    for inv in invoice_list:
        if inv in invoice_nature_map:
            val = invoice_nature_map[inv]
            if pd.notna(val) and str(val).lower() != "nan":
                return val

    if "B2C" in str(row.get("Customer Name", "")).upper():
        return "B2C"
    return "B2B"


def generate_revenue_report(df, value_columns, month_names):
    b2b_df = df[df["Nature"] == "B2B"].copy()

    for col in value_columns:
        if col not in b2b_df.columns:
            b2b_df[col] = 0.0

    b2b_grouped = b2b_df.groupby("Payment Cycle ")[value_columns].sum()
    b2b_grouped.columns = month_names

    total_b2b = b2b_grouped.sum()

    b2c_df = df[df["Nature"] == "B2C"].copy()
    for col in value_columns:
        if col not in b2c_df.columns:
            b2c_df[col] = 0.0

    total_b2c = b2c_df[value_columns].sum()
    total_b2c.index = month_names

    total_revenue = total_b2b + total_b2c

    exclude_keywords = [
        "one time", "days", "10-days", "12-days",
        "15-days", "19-days", "38 -days"
    ]

    def is_recurring(x):
        s = str(x).lower()
        return not any(k in s for k in exclude_keywords)

    b2b_recurring = b2b_df[b2b_df["Payment Cycle "].apply(is_recurring)]

    mrr_b2b = b2b_recurring[value_columns].sum()
    mrr_b2b.index = month_names

    active_subs = pd.Series(
        [b2b_recurring[b2b_recurring[c] > 1].shape[0] for c in value_columns],
        index=month_names
    )

    arpu = mrr_b2b / active_subs.replace(0, 1)
    arr = mrr_b2b * 12

    rows = []

    rows.append({**total_b2b.to_dict(), "Particulars": "Total monthly revenue from B2B subscribers excluding GST (A)"})

    for cycle, data in b2b_grouped.iterrows():
        rows.append({**data.to_dict(), "Particulars": cycle})

    rows.append({"Particulars": ""})
    rows.append({**total_b2c.to_dict(), "Particulars": "Total monthly revenue from B2C subscribers (B)"})
    rows.append({**total_revenue.to_dict(), "Particulars": "Total monthly revenue from B2B and B2C C=A+B"})
    rows.append({"Particulars": ""})

    rows.append({**mrr_b2b.to_dict(), "Particulars": "Monthly Recurring revenue (excluding one time and others)-B2B"})
    rows.append({**active_subs.to_dict(), "Particulars": "No. of active subscribers-B2B"})
    rows.append({**arpu.to_dict(), "Particulars": "Monthly recurring revenue per subscriber"})
    rows.append({**arr.to_dict(), "Particulars": "Total Annual Recurring revenue (excluding one time and others)-B2B"})

    report = pd.DataFrame(rows)
    return report[["Particulars"] + month_names]


def sync_to_google_sheet(mis_sheets, google_sheet_id, keep_worksheet="Invoices"):
    creds_json = os.getenv("GOOGLE_CREDENTIALS_JSON")
    creds_dict = json.loads(creds_json)

    creds = Credentials.from_service_account_info(
        creds_dict,
        scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )

    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(google_sheet_id)

    for ws in spreadsheet.worksheets():
        if ws.title == keep_worksheet:
            continue
        if len(spreadsheet.worksheets()) > 1:
            spreadsheet.del_worksheet(ws)

    for sheet_name, df in mis_sheets.items():
        if sheet_name == keep_worksheet:
            continue

        rows, cols = df.shape
        ws = spreadsheet.add_worksheet(
            title=sheet_name[:100],
            rows=max(rows + 1, 1),
            cols=max(cols, 1)
        )

        values = [df.columns.tolist()] + df.fillna("").astype(str).values.tolist()
        ws.update(values)


def generate_financial_reports(month_name, google_sheet_id=None):
    mis_sheets = churnout(month_name)

    df_invoices = mis_sheets["Invoices"].iloc[2:].copy()
    df_invoices.columns = mis_sheets["Invoices"].iloc[1]
    df_invoices.reset_index(drop=True, inplace=True)

    invoice_nature_map = dict(
        zip(
            df_invoices["Invoice Number"].astype(str).str.strip(),
            df_invoices["Nature"]
        )
    )

    df_accrual = mis_sheets["FY 25-26-Accrual"].iloc[2:].copy()
    df_accrual.columns = mis_sheets["FY 25-26-Accrual"].iloc[1]
    df_accrual.reset_index(drop=True, inplace=True)

    df_accrual["Invoice"] = df_accrual["Invoice"].astype(str).str.strip()
    df_accrual["Nature"] = df_accrual.apply(
        lambda r: find_nature(r, invoice_nature_map),
        axis=1
    )

    ar_columns = [
        "Invoiced amount for April",
        " Invoiced amount for MAY ",
        "June in months",
        " July in months ",
        "Aug Sales in months",
        "Sept Sales in months"
    ]

    accrual_columns = [
        "April Sales (Days Wise)",
        "May Sales (Days Wise)",
        "June in days",
        " July in days ",
        " Aug Sales in days ",
        " Sept Sales in days "
    ]

    month_names = ["Apr-25", "May-25", "Jun-25", "Jul-25", "Aug-25", "Sep-25"]

    target_date = datetime.strptime(month_name, "%b-%y")
    current_date = datetime(2025, 10, 1)

    while current_date <= target_date:
        m = current_date.strftime("%b")
        month_names.append(current_date.strftime("%b-%y"))
        ar_columns.append(f"{m} Sales in months")
        accrual_columns.append(f"{m} Sales in days")
        current_date += relativedelta(months=1)

    for col in ar_columns + accrual_columns:
        if col in df_accrual.columns:
            df_accrual[col] = df_accrual[col].apply(clean_currency)
        else:
            df_accrual[col] = 0.0

    mis_sheets["MR-AR"] = generate_revenue_report(
        df_accrual,
        ar_columns,
        month_names
    )

    mis_sheets["MR Accrual"] = generate_revenue_report(
        df_accrual,
        accrual_columns,
        month_names
    ).iloc[:-4]

    if google_sheet_id:
        sync_to_google_sheet(mis_sheets, google_sheet_id)

    return mis_sheets