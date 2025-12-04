# config.py
import os

# --- Configuration Settings ---
TARGET_MONTH = 10  # October
TARGET_YEAR = 2025

# File Paths
INVOICE_FILE = 'Invoice (2).xlsx'
MASTER_FILE = 'Copy of Master List of Customers_FY 25-26.xlsx'
MIS_PATH = 'MIS_Coresphere100 Apr 25 to Oct 25.xlsx'

# Sheet Names
MASTER_SHEET_NAME = 'FY 25-26'
MIS_SHEET_NAME = 'FY 25-26-Accrual'

# Derived Constants
MONTH_NAMES = {1: 'Jan', 2: 'Feb', 3: 'Mar', 4: 'Apr', 5: 'May', 6: 'Jun',
               7: 'Jul', 8: 'Aug', 9: 'Sep', 10: 'Oct', 11: 'Nov', 12: 'Dec'}
MONTH_NAME = MONTH_NAMES[TARGET_MONTH]