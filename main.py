import shutil
import tempfile
from pathlib import Path

import streamlit as st

from app import WorkflowConfig, orchestrate_workflow


def init_session_state():
    st.session_state.setdefault("temp_dir", None)
    st.session_state.setdefault("final_file", None)
    st.session_state.setdefault("workflow_status", "")


def cleanup_temp_dir():
    temp_dir = st.session_state.get("temp_dir")
    if temp_dir and Path(temp_dir).exists():
        shutil.rmtree(temp_dir, ignore_errors=True)
    st.session_state["temp_dir"] = None
    st.session_state["final_file"] = None
    st.session_state["workflow_status"] = ""


def save_uploaded_file(uploaded_file, destination: Path):
    destination.write_bytes(uploaded_file.getbuffer())
    return destination


def run_workflow(target_month, target_year, invoice_file, master_file, mis_file):
    config = WorkflowConfig(
        target_month=int(target_month),
        target_year=int(target_year),
        invoice_file=str(invoice_file),
        master_file=str(master_file),
        mis_path=str(mis_file),
    )
    orchestrate_workflow(config)
    return mis_file


def main():
    st.set_page_config(page_title="MRR Workflow Orchestrator", layout="centered")
    init_session_state()

    st.title("MRR Workflow Orchestrator")
    st.caption("Process Invoice, Master, and MIS workbooks completely offline and download the refreshed MIS.")

    with st.expander("How it works", expanded=True):
        st.markdown(
            "1. Upload the three Excel workbooks (Invoice, Master, MIS).\n"
            "2. Choose the reporting month/year.\n"
            "3. Run the workflow – progress is shown below.\n"
            "4. Download the processed MIS; temporary files are deleted automatically."
        )

    col_month, col_year = st.columns(2)
    with col_month:
        target_month = st.number_input("Target Month", min_value=1, max_value=12, value=10, step=1)
    with col_year:
        target_year = st.number_input("Target Year", min_value=2000, max_value=2100, value=2025, step=1)

    st.markdown("### Upload Workbooks")
    invoice_upload = st.file_uploader("Invoice Workbook (.xlsx)", type=["xlsx"], key="invoice_upload")
    master_upload = st.file_uploader("Master Workbook (.xlsx)", type=["xlsx"], key="master_upload")
    mis_upload = st.file_uploader("MIS Workbook (.xlsx)", type=["xlsx"], key="mis_upload")

    run_clicked = st.button("Run Workflow", type="primary", use_container_width=True)

    if run_clicked:
        if not all([invoice_upload, master_upload, mis_upload]):
            st.error("Please upload Invoice, Master, and MIS workbooks before running the workflow.")
        else:
            cleanup_temp_dir()
            temp_dir = Path(tempfile.mkdtemp(prefix="mrr_sop_"))
            status_placeholder = st.empty()
            try:
                invoice_path = save_uploaded_file(invoice_upload, temp_dir / invoice_upload.name)
                master_path = save_uploaded_file(master_upload, temp_dir / master_upload.name)
                mis_path = save_uploaded_file(mis_upload, temp_dir / mis_upload.name)

                status_placeholder.info("Running workflow. This may take a moment...")
                final_path = run_workflow(target_month, target_year, invoice_path, master_path, mis_path)

                status_placeholder.success("Workflow complete! Download the processed MIS below.")
                st.session_state["temp_dir"] = str(temp_dir)
                st.session_state["final_file"] = str(final_path)
                st.session_state["workflow_status"] = "ready"
            except Exception as exc:
                cleanup_temp_dir()
                status_placeholder.error(f"Workflow failed: {exc}")

    final_file = st.session_state.get("final_file")
    if final_file and Path(final_file).exists():
        st.markdown("### Download Processed MIS")
        with open(final_file, "rb") as f:
            file_data = f.read()
        download_clicked = st.download_button(
            label="Download Processed MIS",
            data=file_data,
            file_name=Path(final_file).name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        if download_clicked:
            st.success("Download started. Cleaning up temporary files.")
            cleanup_temp_dir()
    elif final_file:
        cleanup_temp_dir()


if __name__ == "__main__":
    main()
