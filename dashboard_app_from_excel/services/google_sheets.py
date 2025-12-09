# services/google_sheets.py
"""Local Excel data access helpers.

The original project pulled data directly from Google Sheets. For offline or
credentials-free deployments we now source everything from the bundled
`strategic_insight.xlsx` workbook (path configurable via `LOCAL_EXCEL_FILE`).

All existing helper functions (`sheet_to_df`, `get_dna`, etc.) remain so the
rest of the Flask app does not need to change.
"""

from __future__ import annotations

import os
import time
from functools import lru_cache
from typing import Optional

import pandas as pd

from config import LOCAL_EXCEL_FILE, SHEET_CACHE_TTL

_excel_cache: dict[str, pd.DataFrame] = {}
_excel_timestamp: float = 0.0
_excel_mtime: float | None = None


def _should_reload() -> bool:
    global _excel_timestamp, _excel_mtime
    if not os.path.exists(LOCAL_EXCEL_FILE):
        raise FileNotFoundError(
            f"Excel data source not found: {LOCAL_EXCEL_FILE}. "
            "Update LOCAL_EXCEL_FILE or place the workbook next to the code."
        )

    current_mtime = os.path.getmtime(LOCAL_EXCEL_FILE)
    if _excel_mtime != current_mtime:
        _excel_mtime = current_mtime
        return True

    if SHEET_CACHE_TTL <= 0:
        return False

    now = time.time()
    if now - _excel_timestamp > SHEET_CACHE_TTL:
        _excel_timestamp = now
        return True
    return False


@lru_cache(maxsize=1)
def _load_workbook() -> pd.ExcelFile:
    return pd.ExcelFile(LOCAL_EXCEL_FILE)


def _refresh_workbook_if_needed():
    global _excel_cache
    if not _excel_cache or _should_reload():
        _excel_cache = {}
        _load_workbook.cache_clear()
        _load_workbook()


ALIAS_MAP = {
    "value_map": "Full price value_map",
    "value_map_promo": "promo price value_map",
    "top_product": "top_product_full_price",
    "top_product_promo": "top_product_promo",
    "okr": "2026 OKR",
    "2025 okr": "2025 OKR",
    "profit_n_loss": "P&L",
    "profit per x": "Profit per X",
    "cost per x": "Cost per X",
    "core -new segments": "Core -New segments",
}


def _normalize_sheet_name(book: pd.ExcelFile, key: str) -> Optional[str]:
    if key in book.sheet_names:
        return key

    lowered = {name.lower(): name for name in book.sheet_names}
    alias_key = ALIAS_MAP.get(key.lower(), key)
    if alias_key in book.sheet_names:
        return alias_key

    return lowered.get(alias_key.lower())


def sheet_to_df(sheet_name_or_index: Optional[str] = None, worksheet_index: int = 0) -> pd.DataFrame:
    _refresh_workbook_if_needed()
    book = _load_workbook()

    sheet_name: Optional[str] = None
    if sheet_name_or_index:
        sheet_name = _normalize_sheet_name(book, sheet_name_or_index)
    else:
        try:
            sheet_name = book.sheet_names[worksheet_index]
        except IndexError:
            sheet_name = None

    if not sheet_name:
        return pd.DataFrame()

    if sheet_name not in _excel_cache:
        _excel_cache[sheet_name] = book.parse(sheet_name).fillna("")
    return _excel_cache[sheet_name].copy()


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
