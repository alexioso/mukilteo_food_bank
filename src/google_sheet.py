import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv
import os
load_dotenv()


def upload_csv_to_gsheet(csv_path, spreadsheet_id, worksheet_name):
    df = pd.read_csv(csv_path)

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_file("../service_account.json", scopes=scopes)
    gc = gspread.authorize(creds)

    sh = gc.open_by_key(spreadsheet_id)
    ws = sh.worksheet(worksheet_name)

    ws.clear()
    ws.update([df.columns.tolist()] + df.fillna("").values.tolist())

# monthly distribution upsert
upload_csv_to_gsheet(
    "../data/processed/df_monthly.csv",
    os.getenv("GOOGLE_SPREADSHEET_ID"),
    "Monthly Distribution Raw"
)

# daily distribution upsert
upload_csv_to_gsheet(
    "../data/raw/total_report_daily.csv",
    os.getenv("GOOGLE_SPREADSHEET_ID"),
    "Daily Distribution Raw"
)