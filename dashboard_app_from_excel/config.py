# config.py
import os

# Google Sheet ID (the one you shared)
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID", "1JY0XOEHqy400a4o1oHRr5U-eAIH-Fr_tSBStV4mBaL0")

# Path to service account JSON (relative to project root)
SERVICE_ACCOUNT_FILE = os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE", "service_account.json")

# Simple cache TTL in seconds (0 = always fetch)
SHEET_CACHE_TTL = int(os.getenv("SHEET_CACHE_TTL", "30"))