from datetime import datetime

import pandas as pd
import numpy as np
from dateutil.relativedelta import relativedelta

from utils.calculate_sales import calculate_sales

def clean_currency_col(val):
    if pd.isna(val) or val == '' or str(val).strip() in ['-', 'CN', '#REF!']:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    try:
        return float(str(val).replace(',', '').replace('"', '').strip())
    except ValueError:
        return 0.0

def get_column_name(df, search_month_names, search_keywords, exclude_keywords=None):
    for col in df.columns:
        c = col.lower()
        if not any(m.lower() in c for m in search_month_names):
            continue
        if not any(k.lower() in c for k in search_keywords):
            continue
        if exclude_keywords and any(e.lower() in c for e in exclude_keywords):
            continue
        return col
    return None

def generate_pivots(month_name):
    df, mis_sheets = calculate_sales(month_name)
    df['Payment Cycle '] = df['Payment Cycle '].fillna('Unknown').astype(str).str.strip()

    start_date = datetime(2025, 4, 1)
    try:
        end_date = datetime.strptime(month_name, "%b-%y")
    except ValueError:
        end_date = datetime(2025, 10, 1)

    month_wise_map = {}
    day_wise_map = {}

    cur = start_date
    while cur <= end_date:
        full = cur.strftime("%B")
        abbr = cur.strftime("%b")

        m_col = get_column_name(
            df,
            [full, abbr],
            ['invoiced amount', 'in months'],
            ['days', 'wise']
        )
        if m_col:
            month_wise_map[full] = m_col
            df[m_col] = df[m_col].apply(clean_currency_col)

        d_col = get_column_name(
            df,
            [full, abbr],
            ['days wise', 'in days', 'sales in days'],
            ['months']
        )
        if d_col:
            day_wise_map[full] = d_col
            df[d_col] = df[d_col].apply(clean_currency_col)

        cur += relativedelta(months=1)

    if month_wise_map:
        pivot_month = (
            df.groupby('Payment Cycle ')[list(month_wise_map.values())]
            .sum()
            .reset_index()
            .rename(columns={v: k for k, v in month_wise_map.items()})
        )
        total_row = pivot_month.drop(columns=['Payment Cycle ']).sum(numeric_only=True)
        total_row['Payment Cycle '] = 'Grand Total'
        pivot_month = pd.concat([pivot_month, pd.DataFrame([total_row])], ignore_index=True)
        pivot_month['Grand Total'] = pivot_month[list(month_wise_map.keys())].sum(axis=1)
    else:
        pivot_month = pd.DataFrame(columns=['Payment Cycle '])

    if day_wise_map:
        pivot_day = (
            df.groupby('Payment Cycle ')[list(day_wise_map.values())]
            .sum()
            .reset_index()
            .rename(columns={v: k for k, v in day_wise_map.items()})
        )
        total_row = pivot_day.drop(columns=['Payment Cycle ']).sum(numeric_only=True)
        total_row['Payment Cycle '] = 'Grand Total'
        pivot_day = pd.concat([pivot_day, pd.DataFrame([total_row])], ignore_index=True)
        pivot_day['Grand Total'] = pivot_day[list(day_wise_map.keys())].sum(axis=1)
    else:
        pivot_day = pd.DataFrame(columns=['Payment Cycle '])

    if "Pivot" in mis_sheets:
        mis_sheets.pop("Pivot", None)
        mis_sheets["Pivot Month Wise"] = pivot_month
        mis_sheets["Pivot Day Wise"] = pivot_day
    else:
        mis_sheets["Pivot Month Wise"] = pivot_month
        mis_sheets["Pivot Day Wise"] = pivot_day

    return mis_sheets, df