import re

import pandas as pd

def normalize_invoice_column(df, col_name="Invoice", default_year="25-26"):
    def _parse_row(text):
        if pd.isna(text) or str(text).strip() == "":
            return ""

        text = str(text).strip()

        text = re.sub(r'\bCN[\s\-]*(?:\d+(?:\s*to\s*\d+)?)', '', text, flags=re.IGNORECASE)

        text = re.sub(r'INV-(\d{2})-(\d{2})-', r'INV~\1~\2~', text)
        text = re.sub(r'INV-', r'INV~', text)
        text = re.sub(r"[,_'\-()\[\]/\n\r]", " ", text)

        tokens = text.split()
        clean_invoices = []
        current_prefix = f"INV-{default_year}-"

        for token in tokens:
            token = token.replace("~", "-")
            if not any(char.isdigit() for char in token):
                continue

            match_full = re.match(r'INV-(\d{2}-\d{2})-?(\d*)', token)
            match_short = re.match(r'INV-?(\d+)', token)

            if match_full:
                year_part = match_full.group(1)
                current_prefix = f"INV-{year_part}-"
                number_part = match_full.group(2)
                if number_part:
                    clean_invoices.append(f"{current_prefix}{number_part.zfill(6)}")
            elif match_short:
                number_part = match_short.group(1)
                clean_invoices.append(f"{current_prefix}{number_part.zfill(6)}")
            elif token.isdigit():
                clean_invoices.append(f"{current_prefix}{token.zfill(6)}")

        return ", ".join(clean_invoices)

    if col_name in df.columns:
        df[col_name] = df[col_name].apply(_parse_row)
    return df