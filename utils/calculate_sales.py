from calendar import monthrange
from datetime import datetime

import pandas as pd
import numpy as np
from dateutil.relativedelta import relativedelta

from utils.attach import attach

def clean_currency(val):
    if pd.isna(val) or str(val).strip() in ["", "-", "CN", "#REF!"]:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    try:
        return float(str(val).replace(",", "").replace('"', "").strip())
    except:
        return 0.0

def parse_date(val):
    if pd.isna(val) or str(val).strip() == "":
        return pd.NaT
    return pd.to_datetime(val, dayfirst=True, errors="coerce")

def get_total_months(start_date, end_date):
    if pd.isna(start_date) or pd.isna(end_date):
        return 0
    rd = relativedelta(end_date + pd.Timedelta(days=1), start_date)
    months = rd.years * 12 + rd.months
    return max(1, months)

def calculate_sales(month_name):
    df, mis_sheets = attach(month_name)

    try:
        target_date = datetime.strptime(month_name, "%b-%y")
    except:
        return df, mis_sheets

    target_month = target_date.month
    target_year = target_date.year
    month_abbr = target_date.strftime("%b")

    month_start = datetime(target_year, target_month, 1)
    month_end = datetime(
        target_year,
        target_month,
        monthrange(target_year, target_month)[1]
    )

    col_days = f"Days in {month_abbr}"
    col_sales_month = f"{month_abbr} Sales in months"
    col_sales_day = f"{month_abbr} Sales in days"

    df[col_days] = 0
    df[col_sales_month] = 0.0
    df[col_sales_day] = 0.0

    for i, row in df.iterrows():
        start_date = parse_date(row["Start Date "])
        end_date = parse_date(row["End Date"])

        if pd.isna(start_date) or pd.isna(end_date):
            continue

        contract_amount = clean_currency(row[" Contract Amount "])

        active_start = max(start_date, month_start)
        active_end = min(end_date, month_end)

        if active_start > active_end:
            continue

        days_in_month = (active_end - active_start).days + 1
        df.at[i, col_days] = days_in_month

        total_contract_days = (end_date - start_date).days + 1
        total_contract_months = get_total_months(start_date, end_date)

        if total_contract_days > 0:
            val = (contract_amount / total_contract_days) * days_in_month
            df.at[i, col_sales_day] = round(val, 2)

        is_closing_month = (
            end_date.year == target_year and end_date.month == target_month
        )

        if is_closing_month:
            df.at[i, col_sales_month] = 0.0
        else:
            if total_contract_months > 0:
                val = contract_amount / total_contract_months
                df.at[i, col_sales_month] = round(val, 2)

    full_sheet = mis_sheets["FY 25-26-Accrual"]

    header_rows = full_sheet.iloc[:2].copy()

    df_to_stack = df.copy()
    df_to_stack.columns = range(len(df.columns))

    if header_rows.shape[1] != df_to_stack.shape[1]:
        new_header = pd.DataFrame(
            np.nan,
            index=header_rows.index,
            columns=range(df_to_stack.shape[1])
        )
        new_header.iloc[:, :header_rows.shape[1]] = header_rows.values
        header_rows = new_header

    header_rows.iloc[1] = df.columns.tolist()

    final_sheet = pd.concat([header_rows, df_to_stack], ignore_index=True)

    mis_sheets["FY 25-26-Accrual"] = final_sheet

    return df, mis_sheets