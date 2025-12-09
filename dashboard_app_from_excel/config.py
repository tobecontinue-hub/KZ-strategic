# config.py
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Optional legacy Google Sheets settings (kept for backward compatibility)
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID")
SERVICE_ACCOUNT_FILE = os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE", "service_account.json")

# Local Excel dashboard source
LOCAL_EXCEL_FILE = os.getenv(
    "LOCAL_EXCEL_FILE",
    os.path.join(BASE_DIR, "strategic_insight.xlsx"),
)

# Cache TTL in seconds (0 = always fetch)
SHEET_CACHE_TTL = int(os.getenv("SHEET_CACHE_TTL", "30"))