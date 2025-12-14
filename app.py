import os
import traceback
import pandas as pd
import streamlit as st
from supabase import create_client, Client
from dotenv import load_dotenv

from utils.mrr import generate_financial_reports

load_dotenv()

url = os.getenv("SUPABASE_URL")
key = os.getenv("SUPABASE_KEY")

supabase: Client = create_client(url, key)

st.set_page_config(
    page_title="MRR Calculator",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.title("📊 MRR Calculator")
st.caption("End-to-end MRR, ARR & MIS automation")

st.divider()

def fetch_sheet_id(table_name):
    res = supabase.table(table_name).select("*").limit(1).execute()
    if res.data:
        return res.data[0]["sheet_id"]
    return ""

def update_sheet_id(table_name, new_sheet_id):
    res = supabase.table(table_name).select("*").limit(1).execute()
    if res.data:
        row = res.data[0]
        row_id = row.get("id")
        if row_id is not None:
            supabase.table(table_name).update(
                {"sheet_id": new_sheet_id}
            ).eq("id", row_id).execute()
            return
    supabase.table(table_name).insert(
        {"sheet_id": new_sheet_id}
    ).execute()

invoice_sheet_id = fetch_sheet_id("coro_invoices_sheet")
mis_sheet_id = fetch_sheet_id("coro_mis_sheet")
master_sheet_id = fetch_sheet_id("coro_master_sheet")

st.subheader("🔧 Google Sheet Configuration")
st.caption("These Sheet IDs control where data is read from and written to")

with st.form("sheet_id_form"):
    col1, col2, col3 = st.columns(3)

    with col1:
        invoice_input = st.text_input(
            "Invoices Sheet ID",
            value=invoice_sheet_id,
            placeholder="Paste Google Sheet ID"
        )

    with col2:
        mis_input = st.text_input(
            "MIS Sheet ID",
            value=mis_sheet_id,
            placeholder="Paste Google Sheet ID"
        )

    with col3:
        master_input = st.text_input(
            "Master Sheet ID",
            value=master_sheet_id,
            placeholder="Paste Google Sheet ID"
        )

    save = st.form_submit_button("💾 Save Configuration")

    if save:
        update_sheet_id("coro_invoices_sheet", invoice_input)
        update_sheet_id("coro_mis_sheet", mis_input)
        update_sheet_id("coro_master_sheet", master_input)
        st.success("Sheet configuration updated successfully")

st.divider()

st.subheader("🚀 Run MRR Workflow")
st.caption("Generates MIS, churn, additions, deletions, ARR & MRR reports")

with st.form("run_mrr_form"):
    left, right = st.columns([3, 1])

    with left:
        month_input = st.text_input(
            "Target Month (Mon-YY)",
            value="Oct-25"
        )

    with right:
        st.markdown("<div style='height: 28px'></div>", unsafe_allow_html=True)
        run = st.form_submit_button(
            "▶ Run Workflow",
            use_container_width=True
        )

if run:
    try:
        with st.spinner("Running MRR calculations…"):
            mis_sheets = generate_financial_reports(
                month_input,
                google_sheet_id=mis_input
            )

        st.success(f"Workflow completed — {len(mis_sheets)} sheets generated & synced")

        st.divider()
        st.subheader("📄 Generated Sheets")

        tabs = st.tabs(list(mis_sheets.keys()))

        for tab, (name, df) in zip(tabs, mis_sheets.items()):
            with tab:
                st.markdown(f"### {name}")
                st.dataframe(
                    df.fillna("").reset_index(drop=True),
                    use_container_width=True,
                    height=600
                )

    except Exception:
        st.error("Workflow failed")
        st.code(traceback.format_exc(), language="text")

