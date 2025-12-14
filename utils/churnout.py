from datetime import datetime

import pandas as pd
import numpy as np

from dateutil.relativedelta import relativedelta

from utils.generate_pivot import generate_pivots

def get_previous_month_str(month_name):
    try:
        dt = datetime.strptime(month_name, "%b-%y")
        return (dt - relativedelta(months=1)).strftime("%b-%y")
    except:
        return None

def churnout(month_name):
    mis_sheets, sales_data = generate_pivots(month_name)

    last_col = sales_data.columns[-1]
    sales_work_df = sales_data.copy()

    sales_work_df[last_col] = (
        sales_work_df[last_col]
        .astype(str)
        .str.replace(",", "", regex=False)
        .str.strip()
    )
    sales_work_df[last_col] = pd.to_numeric(
        sales_work_df[last_col], errors="coerce"
    ).fillna(0)

    active_companies = (
        sales_work_df.loc[sales_work_df[last_col] > 0, "Customer Name"]
        .dropna()
        .astype(str)
        .str.strip()
        .drop_duplicates()
        .reset_index(drop=True)
    )

    if "Active Subscriber" not in mis_sheets:
        mis_sheets["Active Subscriber"] = pd.DataFrame()

    active_df = mis_sheets["Active Subscriber"]

    if len(active_companies) > len(active_df):
        active_df = active_df.reindex(range(len(active_companies)))
        mis_sheets["Active Subscriber"] = active_df

    mis_sheets["Active Subscriber"][month_name] = pd.Series(active_companies)

    current_df = mis_sheets["Active Subscriber"]
    additions = []
    deletions = []

    if current_df.shape[1] >= 2:
        current = set(current_df.iloc[:, -1].dropna().astype(str).str.strip())
        previous = set(current_df.iloc[:, -2].dropna().astype(str).str.strip())
        additions = list(current - previous)
        deletions = list(previous - current)
    elif current_df.shape[1] == 1:
        additions = current_df.iloc[:, 0].dropna().tolist()
        deletions = []

    if "Addition" in mis_sheets:
        df_add = mis_sheets["Addition"]
        header_row_idx = 2
        data_start_idx = 3

        needed_length = data_start_idx + len(additions)
        if len(df_add) < needed_length:
            df_add = df_add.reindex(range(needed_length))

        new_column_data = [np.nan] * len(df_add)
        new_column_data[header_row_idx] = month_name

        for i, company in enumerate(additions):
            new_column_data[data_start_idx + i] = company

        df_add[month_name] = new_column_data
        mis_sheets["Addition"] = df_add

    if "Deletions" in mis_sheets:
        df_del = mis_sheets["Deletions"]
        header_row_idx = 2
        data_start_idx = 3

        needed_length = data_start_idx + len(deletions)
        if len(df_del) < needed_length:
            df_del = df_del.reindex(range(needed_length))

        new_column_data = [np.nan] * len(df_del)
        new_column_data[header_row_idx] = month_name

        for i, company in enumerate(deletions):
            new_column_data[data_start_idx + i] = company

        df_del[month_name] = new_column_data
        mis_sheets["Deletions"] = df_del

    if "Customer churnout" in mis_sheets:
        df_churn = mis_sheets["Customer churnout"]

        r_header = 3
        r_begin = 4
        r_add = 5
        r_less = 6
        r_end = 7

        prev_month_name = get_previous_month_str(month_name)
        header_values = df_churn.iloc[r_header].astype(str).str.strip().tolist()

        try:
            target_col_idx = header_values.index(month_name)
        except ValueError:
            valid_indices = [
                i for i, v in enumerate(header_values)
                if v.lower() not in ["nan", "none", "", "particulars"] and i > 0
            ]
            target_col_idx = valid_indices[-1] + 1 if valid_indices else 1
            if target_col_idx >= df_churn.shape[1]:
                df_churn[f"col_{target_col_idx}"] = np.nan

        try:
            prev_col_idx = header_values.index(prev_month_name) if prev_month_name else -1
        except ValueError:
            prev_col_idx = -1

        begin_count = 0
        if prev_col_idx != -1:
            try:
                begin_count = float(str(df_churn.iloc[r_end, prev_col_idx]).replace(",", ""))
            except:
                begin_count = 0

        add_count = len(additions)
        del_count = len(deletions)
        end_count = len(active_companies)

        df_churn.iloc[r_header, target_col_idx] = month_name
        df_churn.iloc[r_begin, target_col_idx] = begin_count
        df_churn.iloc[r_add, target_col_idx] = add_count
        df_churn.iloc[r_less, target_col_idx] = del_count
        df_churn.iloc[r_end, target_col_idx] = end_count

        mis_sheets["Customer churnout"] = df_churn

    return mis_sheets