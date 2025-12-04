# financial_reports.py
import pandas as pd
import numpy as np
import re
from config import *
from utils import extract_all

def generate_pivot_and_mrar():
    invoices = pd.read_excel(MIS_PATH, sheet_name="Invoices", skiprows=2)
    fy_acc = pd.read_excel(MIS_PATH, sheet_name=MIS_SHEET_NAME, skiprows=2)

    # --- B2B Logic ---
    b2b_df = invoices[invoices["Nature"].astype(str).str.upper() == "B2B"]
    b2b_companies = b2b_df[["Customer Name"]].drop_duplicates().reset_index(drop=True)
    b2b_data = fy_acc[fy_acc["Customer Name"].isin(b2b_companies["Customer Name"])].reset_index(drop=True)

    # 1. Pivot Month (Sales in months)
    fixed_month_cols = [
        "Invoiced amount for April", "Invoiced amount for MAY",
        "June in months", "July in months", "Aug Sales in months", "Sept Sales in months"
    ]
    all_dynamic = [
        c for c in b2b_data.columns
        if c not in fixed_month_cols and (
            c.endswith("Sales in months") or c.endswith("in months") or c.lower().startswith("invoiced amount for")
        )
    ]
    month_order_map = {"jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
                       "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12}
    
    def extract_month_index(col):
        txt = col.lower()
        for m in month_order_map:
            if m in txt:
                return month_order_map[m]
        return 999

    dynamic_month_cols = sorted(all_dynamic, key=extract_month_index)
    month_cols = fixed_month_cols + dynamic_month_cols

    pivot_month = b2b_data.pivot_table(index="Payment Cycle", values=month_cols, aggfunc="sum").reset_index()
    grand_total_pm = pivot_month[month_cols].sum().to_frame().T
    grand_total_pm.insert(0, "Payment Cycle", "Grand Total")
    pivot_month = pd.concat([pivot_month, grand_total_pm], ignore_index=True)

    # 2. Pivot Days (Sales in days)
    fixed_date_cols = [
        "April Sales (Days Wise)", "May Sales (Days Wise)",
        "June in days", "July in days", "Aug Sales in days", "Sept Sales in days"
    ]
    all_dynamic_days = [
        c for c in b2b_data.columns
        if c not in fixed_date_cols and (
            c.endswith("Sales in days") or c.endswith("in days") or "(days" in c.lower()
        )
    ]
    dynamic_date_cols = sorted(all_dynamic_days, key=extract_month_index)
    date_cols = fixed_date_cols + dynamic_date_cols

    pivot_days = b2b_data.pivot_table(index="Payment Cycle", values=date_cols, aggfunc="sum").reset_index()
    grand_total_pd = pivot_days[date_cols].sum().to_frame().T
    grand_total_pd.insert(0, "Payment Cycle", "Grand Total")
    pivot_days = pd.concat([pivot_days, grand_total_pd], ignore_index=True)

    # --- B2C Logic ---
    b2c_df = invoices[invoices["Nature"].astype(str).str.upper() == "B2C"].copy()
    b2c_df["Norm Inv"] = b2c_df["Invoice Number"].apply(extract_all)
    b2c_df = b2c_df.explode("Norm Inv")

    fy_acc["Norm Inv"] = fy_acc["Invoice Number"].apply(extract_all)
    fy_acc = fy_acc.explode("Norm Inv")
    b2c_data = fy_acc[fy_acc["Norm Inv"].isin(b2c_df["Norm Inv"])].reset_index(drop=True)

    # --- MR-AR Report Construction ---
    pivot = pivot_days.copy()
    fixed_day_map = {
        "April Sales (Days Wise)": "April", "May Sales (Days Wise)": "May",
        "June in days": "June", "July in days": "July",
        "Aug Sales in days": "August", "Sept Sales in days": "September"
    }
    available_fixed = {k: v for k, v in fixed_day_map.items() if k in pivot.columns}
    dynamic_day_cols = [c for c in pivot.columns if c not in available_fixed and ("days" in c.lower())]

    def extract_month_name(col):
        t = col.lower()
        for m in month_order_map:
            if m in t:
                return m.capitalize()
        return None

    dynamic_day_rename = {c: extract_month_name(c) for c in dynamic_day_cols if extract_month_name(c)}
    rename_map_all = {}
    rename_map_all.update(available_fixed)
    rename_map_all.update(dynamic_day_rename)

    pivot = pivot.rename(columns=rename_map_all)
    
    final_month_cols = [rename_map_all[c] for c in available_fixed] + sorted(
        [dynamic_day_rename[c] for c in dynamic_day_rename],
        key=lambda x: month_order_map[x[:3].lower()]
    )

    mr_ar_b2b = pivot[["Payment Cycle"] + final_month_cols].copy()
    mr_ar_b2b = mr_ar_b2b.rename(columns={"Payment Cycle": "Particulars"})

    # B2C Totals
    available_month_fixed = [c for c in fixed_month_cols if c in b2c_data.columns]
    dynamic_month_cols_b2c = [
        c for c in b2c_data.columns
        if c not in available_month_fixed and (
            c.endswith("Sales in months") or c.endswith("in months") or c.lower().startswith("invoiced amount for")
        )
    ]
    dynamic_month_rename = {c: extract_month_name(c) for c in dynamic_month_cols_b2c if extract_month_name(c)}
    rename_map_months = {c: extract_month_name(c) for c in available_month_fixed}
    rename_map_months.update(dynamic_month_rename)

    all_month_cols_b2c = available_month_fixed + dynamic_month_cols_b2c
    b2c_totals = b2c_data[all_month_cols_b2c].sum().to_frame().T
    b2c_totals.insert(0, "Particulars", "B2C")
    b2c_totals = b2c_totals.rename(columns=rename_map_months)

    final_months = sorted(rename_map_months.values(), key=lambda x: month_order_map[x[:3].lower()])

    mr_ar = pd.concat([mr_ar_b2b, b2c_totals[["Particulars"] + final_months]], ignore_index=True)
    
    # Correction for B2C short names if present
    mr_ar.loc[mr_ar["Particulars"]=="B2C", ["April","June","July","August","September"]] = \
        mr_ar.loc[mr_ar["Particulars"]=="B2C", ["Apr","Jun","Jul","Aug","Sep"]].values
    mr_ar = mr_ar.drop(columns=["Apr", "Jun", "Jul", "Aug", "Sep"], errors="ignore")

    return mr_ar, pivot_days, pivot_month