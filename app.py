# app.py
import logging
from dataclasses import dataclass, field
from typing import Optional, Sequence

from utils.accrual_logic import process_accrual_update
from utils.financial_reports import (
    generate_pivot_and_mrar,
    save_financial_reports,
)
from utils.subscriber_logic import process_subscribers_and_churn

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


@dataclass
class UploadTarget:
    table_name: str
    worksheet_name: str
    source_sheet_name: Optional[str] = None


@dataclass
class WorkflowConfig:
    target_month: int
    target_year: int
    invoice_file: str
    master_file: str
    mis_path: str
    master_sheet_name: str = "FY 25-26"
    mis_sheet_name: str = "FY 25-26-Accrual"
    upload_targets: list[UploadTarget] = field(default_factory=list)


def orchestrate_workflow(config: WorkflowConfig):
    """Run the accrual -> subscriber -> financial workflow."""
    logger.info("Starting workflow for %s/%s", config.target_month, config.target_year)

    logger.info("Stage 1: Accrual logic")
    process_accrual_update(config)

    logger.info("Stage 2: Subscriber logic")
    process_subscribers_and_churn(config)

    logger.info("Stage 3: Financial reports")
    mr_ar, pivot_days, pivot_month = generate_pivot_and_mrar(config)
    save_financial_reports(config, mr_ar, pivot_days, pivot_month)

    if config.upload_targets:
        logger.info("Uploading workbook sheets")
    else:
        logger.info("No upload targets configured; skipping upload")

    logger.info("Workflow complete")


def build_config(
    target_month: int = 10,
    target_year: int = 2025,
    invoice_file: str = "data/Invoice (2).xlsx",
    master_file: str = "data/Copy of Master List of Customers_FY 25-26.xlsx",
    mis_path: str = "data/MIS_Coresphere100 Apr 25 to Oct 25.xlsx",
    master_sheet_name: str = "FY 25-26",
    mis_sheet_name: str = "FY 25-26-Accrual",
    upload_targets: Optional[Sequence[UploadTarget]] = None,
) -> WorkflowConfig:
    """Helper for building a WorkflowConfig with optional overrides."""
    targets = list(upload_targets) if upload_targets else []
    return WorkflowConfig(
        target_month=target_month,
        target_year=target_year,
        invoice_file=invoice_file,
        master_file=master_file,
        mis_path=mis_path,
        master_sheet_name=master_sheet_name,
        mis_sheet_name=mis_sheet_name,
        upload_targets=targets,
    )


def main():
    config = build_config()
    orchestrate_workflow(config)


if __name__ == "__main__":
    main()
