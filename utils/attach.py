import pandas as pd
import numpy as np

from utils.extract_data import extract_data
from utils.normalize import normalize_invoice_column

def normalize_mis_invoices(df):
    df = normalize_invoice_column(df, col_name="Invoice", default_year="25-26")
    return df

def attach(month):
    matched_df, mis_sheets = extract_data()

    original_sheet = mis_sheets["FY 25-26-Accrual"]

    header_rows = original_sheet.iloc[:2].copy()

    df = original_sheet.iloc[2:].copy()
    df.columns = original_sheet.iloc[1].values
    df = df.reset_index(drop=True)

    df = normalize_mis_invoices(df)
    target_df = df

    existing_invoices = set(
        target_df["Invoice"]
        .dropna()
        .astype(str)
        .str.split(",")
        .explode()
        .str.strip()
    )

    new_rows_source = matched_df[
        ~matched_df["Invoice Number"]
        .astype(str)
        .str.strip()
        .isin(existing_invoices)
    ].copy()

    column_mapping = {
        "Invoice Number": "Invoice",
        "Customer Name": "Customer Name",
        "Start Date ": "Start Date ",
        "End Date": "End Date",
        "Payment Cycle ": "Payment Cycle ",
        "Contract Amount": " Contract Amount "
    }

    data_to_append = new_rows_source[list(column_mapping.keys())].rename(
        columns=column_mapping
    )

    if not data_to_append.empty:
        data_to_append["Months"] = month

    final_df = pd.concat([target_df, data_to_append], ignore_index=True)

    month_name = month.split("-")[0] if "-" in month else month

    col_days = f"Days in {month_name}"
    col_sales_months = f"{month_name} Sales in months"
    col_sales_days = f"{month_name} Sales in days"

    final_df[col_days] = ""
    final_df[col_sales_months] = ""
    final_df[col_sales_days] = ""

    start = pd.to_datetime(final_df["Start Date "], dayfirst=True, errors="coerce")
    end = pd.to_datetime(final_df["End Date"], dayfirst=True, errors="coerce")

    calculated_period = (end - start).dt.days + 1
    final_df["Contract period"] = final_df["Contract period"].fillna(calculated_period)

    new_column_names = final_df.columns.tolist()

    reconstructed_header = pd.DataFrame(
        np.nan,
        index=header_rows.index,
        columns=range(len(new_column_names))
    )

    orig_width = header_rows.shape[1]
    reconstructed_header.iloc[:, :orig_width] = header_rows.values
    reconstructed_header.iloc[1] = new_column_names

    final_df_for_sheet = final_df.copy()
    final_df_for_sheet.columns = range(len(new_column_names))

    full_sheet_df = pd.concat(
        [reconstructed_header, final_df_for_sheet],
        ignore_index=True
    )

    mis_sheets["FY 25-26-Accrual"] = full_sheet_df

    return final_df, mis_sheets