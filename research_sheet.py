import os
import json
import gspread
from google.oauth2.service_account import Credentials
import base64

SPREADSHEET_ID = '1tQ5wMw3m7ZdlQx7BuNBeinokDjCWx8xc8LT6YL16TYI'

def get_google_client():
    creds_raw = os.environ.get('GOOGLE_CREDENTIALS_JSON', '').strip()
    if not creds_raw:
        raise ValueError("GOOGLE_CREDENTIALS_JSON environment variable not set")
    
    try:
        decoded = base64.b64decode(creds_raw).decode('utf-8')
        if decoded.strip().startswith('{'):
            creds_dict = json.loads(decoded)
    except Exception:
        creds_dict = json.loads(creds_raw)

    scopes = ['https://www.googleapis.com/auth/spreadsheets.readonly']
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(creds)

def research():
    gc = get_google_client()
    sh = gc.open_by_key(SPREADSHEET_ID)
    
    # 1. Get Target Date from D.count sheet cell Z2
    try:
        ws_count = sh.worksheet("D.count")
        target_date = ws_count.acell('Z2').value
        print(f"Target Date (D.count!Z2): {target_date}")
    except Exception as e:
        print(f"Error reading D.count: {e}")

    # 2. Get Headers and first few rows from Natural Trends sheet
    try:
        ws_trends = sh.worksheet("Natural Trends")
        header_row = 1 # Assuming headers are at row 1
        data = ws_trends.get('A1:Z5')
        print("Natural Trends Data (first 5 rows):")
        for row in data:
            print(row)
    except Exception as e:
        print(f"Error reading Natural Trends: {e}")

if __name__ == "__main__":
    research()
