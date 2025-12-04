import argparse

from app import WorkflowConfig, UploadTarget, orchestrate_workflow


def parse_upload_target(value: str) -> UploadTarget:
    """
    Parse an upload target definition.
    Format: table_name:worksheet_name[:source_sheet_name]
    """
    parts = value.split(":")
    if len(parts) < 2 or len(parts) > 3:
        raise argparse.ArgumentTypeError(
            "Upload target must be table:worksheet[:source_sheet]"
        )
    table_name, worksheet = parts[0], parts[1]
    source_sheet = parts[2] if len(parts) == 3 else None
    return UploadTarget(
        table_name=table_name,
        worksheet_name=worksheet,
        source_sheet_name=source_sheet,
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Invoke MRR workflow orchestrator.")
    parser.add_argument("--target-month", type=int, required=True, help="Reporting month (1-12).")
    parser.add_argument("--target-year", type=int, required=True, help="Reporting year (e.g., 2025).")
    parser.add_argument("--invoice-file", required=True, help="Path to the invoice workbook.")
    parser.add_argument("--master-file", required=True, help="Path to the customer master workbook.")
    parser.add_argument("--mis-path", required=True, help="Path to the MIS workbook.")
    parser.add_argument(
        "--master-sheet-name",
        default="FY 25-26",
        help="Sheet name in the master workbook (default: FY 25-26).",
    )
    parser.add_argument(
        "--mis-sheet-name",
        default="FY 25-26-Accrual",
        help="Accrual sheet name in the MIS workbook (default: FY 25-26-Accrual).",
    )
    parser.add_argument(
        "--upload-target",
        action="append",
        type=parse_upload_target,
        help="Optional upload target definition table:worksheet[:source_sheet]. Can repeat.",
    )
    return parser


def invoke():
    parser = build_parser()
    args = parser.parse_args()

    config = WorkflowConfig(
        target_month=args.target_month,
        target_year=args.target_year,
        invoice_file=args.invoice_file,
        master_file=args.master_file,
        mis_path=args.mis_path,
        master_sheet_name=args.master_sheet_name,
        mis_sheet_name=args.mis_sheet_name,
        upload_targets=args.upload_target or [],
    )

    orchestrate_workflow(config)


if __name__ == "__main__":
    invoke()
