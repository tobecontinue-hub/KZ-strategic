# KZ Strategic Dashboard

Flask dashboard that reads strategic-performance data directly from Google Sheets and renders multiple executive views (executive summary, value maps, OKRs, etc.).

## Local development

```bash
python -m venv venv
source venv/bin/activate           # Windows: venv\Scripts\activate
pip install -r requirements.txt

# Provide credentials & sheet config
set GOOGLE_SHEET_ID=...            # or export on macOS/Linux
set GOOGLE_SERVICE_ACCOUNT_FILE=service_account.json

python run.py
# http://127.0.0.1:5000
```

## Required environment variables

| Name | Description |
|------|-------------|
| `GOOGLE_SHEET_ID` | ID of the Google Spreadsheet containing all tabs used by the app |
| `GOOGLE_SERVICE_ACCOUNT_FILE` | Path to the service-account JSON that has read access |
| `SHEET_CACHE_TTL` | Optional cache TTL (seconds). Defaults to 30 |

The default `config.py` falls back to `service_account.json` in the repository root if the env vars are not provided.

## Deployment (Render free tier example)

1. Push the project to GitHub.
2. Create a free account on [Render](https://render.com) → "New" → "Web Service".
3. Connect the GitHub repository that contains this app.
4. Use the following settings:
   - **Runtime:** Python
   - **Build command:** `pip install -r requirements.txt`
   - **Start command:** `gunicorn app:app` (or `gunicorn dashboard_app_from_excel.app:app` if the repo root contains additional files)
5. Add Environment Variables in the Render dashboard:
   - `GOOGLE_SHEET_ID`
   - `GOOGLE_SERVICE_ACCOUNT_FILE` (point to the secret-file path, see below)
6. Upload the Google service-account JSON via **Secret Files** in Render (Settings → Secret Files) and set `GOOGLE_SERVICE_ACCOUNT_FILE` to that generated path.
7. Deploy. Render will build the image and provide a live URL.

## Repository structure

```
dashboard_app_from_excel/
├── app.py                  # Flask routes
├── config.py
├── requirements.txt
├── run.py                  # Local entry point
├── services/
│   └── google_sheets.py    # Google Sheets integration helpers
├── templates/
├── static/
└── ...
```

## Notes
- **Do not commit** `service_account.json`. The `.gitignore` already ignores it.
- Ensure the service account has at least read access on every tab referenced in the code (e.g., `exe_summary`, `value_map`, `org_chart`, etc.).
- Each time you push to the main branch, Render (or any connected platform) can auto-deploy the latest version.
