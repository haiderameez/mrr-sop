import pandas as pd

from utils.get_sheet import get_sheet
from utils.normalize import normalize_invoice_column

def normalize_master_invoices(df):
    df = normalize_invoice_column(df, col_name="Invoice", default_year="25-26")
    return df

def match_invoices(invoice_df, master_df):
    master_cols = [
        "Nature of service", 
        "Start Date ",
        "End Date", 
        "Payment Cycle ",
        "Contract Amount"
    ]
    
    existing_master_cols = [col for col in master_cols if col in master_df.columns]
    
    expanded_rows = []
    
    for idx, row in master_df.iterrows():
        inv_val = row.get('Invoice', '')
        
        if pd.isna(inv_val):
            continue
        
        invoices = [x.strip() for x in str(inv_val).split(',') if x.strip()]
        
        for inv in invoices:
            data = {col: row[col] for col in existing_master_cols}
            data['Matched_Invoice_Key'] = inv
            expanded_rows.append(data)
            
    master_lookup = pd.DataFrame(expanded_rows)
    
    if master_lookup.empty:
        return pd.DataFrame()
        
    merged = pd.merge(
        invoice_df[['Invoice Number', 'Customer Name']],
        master_lookup,
        left_on='Invoice Number',
        right_on='Matched_Invoice_Key',
        how='inner'
    )
    
    final_cols = ['Invoice Number', 'Customer Name'] + existing_master_cols
    
    return merged[final_cols]

def extract_data():
    invoice_sheets, mis_sheets, master_sheets = get_sheet()

    invoice_sheets['Invoice (2)'] = invoice_sheets['Invoice (2)'].drop_duplicates()

    new_header = master_sheets["FY 25-26"].iloc[1]
    master_sheets["FY 25-26"] = master_sheets["FY 25-26"].iloc[2:]
    master_sheets["FY 25-26"].columns = new_header
    master_sheets["FY 25-26"] = master_sheets["FY 25-26"].reset_index(drop=True)
    
    master_sheets["FY 25-26"] = normalize_master_invoices(master_sheets["FY 25-26"])

    matched_df = match_invoices(invoice_sheets['Invoice (2)'], master_sheets["FY 25-26"])

    return matched_df, mis_sheets