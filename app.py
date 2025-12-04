# app.py
from accrual_logic import process_accrual_update
from financial_reports import generate_pivot_and_mrar
from subscriber_logic import process_subscribers_and_churn

def main():
    
    # Step 1: Update Accrual Logic (Invoices -> Master -> MIS)
    # This updates the 'FY 25-26-Accrual' sheet in the Excel file
    process_accrual_update()
    
    # Step 2: Generate Report Dataframes
    # This reads the updated Excel and calculates Pivots and MR-AR
    mr_ar, pivot_days, pivot_month = generate_pivot_and_mrar()
    
    # Step 3: Handle Subscriber Lists, Churn, and Final Saving
    # This calculates additions/deletions and writes all auxiliary sheets back to Excel
    process_subscribers_and_churn(mr_ar, pivot_days, pivot_month)

if __name__ == "__main__":
    main()