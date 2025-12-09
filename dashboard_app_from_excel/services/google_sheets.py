import os
import json
import gspread
from oauth2client.service_account import ServiceAccountCredentials

SERVICE_ACCOUNT_JSON = os.getenv("flask-dashboard-service@dashboard-app-480314.iam.gserviceaccount.com")
GOOGLE_SHEET_ID = os.getenv("https://docs.google.com/spreadsheets/d/1JY0XOEHqy400a4o1oHRr5U-eAIH-Fr_tSBStV4mBaL0/edit?usp=sharing")

def _get_client():
    if not SERVICE_ACCOUNT_JSON:
        raise FileNotFoundError("SERVICE_ACCOUNT_JSON env var missing")

    info = json.loads(SERVICE_ACCOUNT_JSON)

    scope = [
        "https://www.googleapis.com/auth/spreadsheets.readonly",
        "https://www.googleapis.com/auth/drive.readonly"
    ]

    creds = ServiceAccountCredentials.from_json_keyfile_dict(info, scope)
    return gspread.authorize(creds)
