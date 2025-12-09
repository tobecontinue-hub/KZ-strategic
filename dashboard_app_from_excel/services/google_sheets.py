# services/google_sheets.py
import time
import os
from typing import Optional
import gspread
import pandas as pd
from google.oauth2.service_account import Credentials

from config import GOOGLE_SHEET_ID, SERVICE_ACCOUNT_FILE, SHEET_CACHE_TTL

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

_cache = {"ts": 0, "spreadsheet": None}


def _get_client():
    if not os.path.exists(SERVICE_ACCOUNT_FILE):
        raise FileNotFoundError(f"Service account JSON missing: {SERVICE_ACCOUNT_FILE}")
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    return gspread.authorize(creds)


def _get_spreadsheet(force_refresh=False):
    now = time.time()
    if (
        not force_refresh
        and _cache["spreadsheet"] is not None
        and now - _cache["ts"] < SHEET_CACHE_TTL
    ):
        return _cache["spreadsheet"]

    ss = _get_client().open_by_key(GOOGLE_SHEET_ID)
    _cache["spreadsheet"] = ss
    _cache["ts"] = now
    return ss


def sheet_to_df(sheet_name_or_index: Optional[str] = None, worksheet_index: int = 0):
    ss = _get_spreadsheet()
    ws = None

    if sheet_name_or_index:
        try:
            ws = ss.worksheet(sheet_name_or_index)
        except Exception:
            pass

    if ws is None:
        ws = ss.get_worksheet(worksheet_index)

    return pd.DataFrame(ws.get_all_records())


# --------- SPECIFIC SHEET HELPERS ----------
def get_exe_summary():
    return sheet_to_df("exe_summary")

def get_brand_promise():
    return sheet_to_df("brand_promise")

def get_value_map():
    return sheet_to_df("value_map")

def get_trajectories():
    return sheet_to_df("trajectories")

def get_segments():
    return sheet_to_df("segments")

def get_swot():
    return sheet_to_df("swot")

def get_dna():
    return sheet_to_df("dna")

def get_roadmap():
    return sheet_to_df("roadmap")

def get_top_product():
    return sheet_to_df("top_product")

def get_fna_performance():
    return sheet_to_df("fna_performance")

def get_operation_health():
    return sheet_to_df("operation_health")

def get_key_decisions():
    return sheet_to_df("key_desicions")
