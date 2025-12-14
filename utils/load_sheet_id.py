import os

from supabase import create_client, Client
import pandas as pd
from dotenv import load_dotenv

load_dotenv()

url = os.getenv("SUPABASE_URL")
key = os.getenv("SUPABASE_KEY")

supabase: Client = create_client(url, key)

def get_sheet_ids():
    invoices = supabase.table("coro_invoices_sheet").select("*").execute()
    mis = supabase.table("coro_mis_sheet").select("*").execute()
    master = supabase.table("coro_master_sheet").select("*").execute()

    df_invoices = pd.DataFrame(invoices.data)
    df_mis = pd.DataFrame(mis.data)
    df_master = pd.DataFrame(master.data)

    invoice_sheet_id = df_invoices["sheet_id"].iloc[0]
    mis_sheet_id = df_mis["sheet_id"].iloc[0]
    master_sheet_id = df_master["sheet_id"].iloc[0]

    return invoice_sheet_id, mis_sheet_id, master_sheet_id