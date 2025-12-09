# KZ Strategic Dashboard

Flask dashboard that renders executive views from the `strategic_insight.xlsx`
workbook bundled with the project (one worksheet per dashboard section).

## Local development

```bash
python -m venv venv
source venv/bin/activate           # Windows: venv\Scripts\activate
pip install -r requirements.txt

# Optional: use a different workbook
set LOCAL_EXCEL_FILE="D:\\data\\custom_dashboard.xlsx"

python run.py
# http://127.0.0.1:5000
```

## Configuration

| Name | Description |
|------|-------------|
| `LOCAL_EXCEL_FILE` | Absolute/relative path to the Excel workbook (defaults to `strategic_insight.xlsx` inside the project) |
| `SHEET_CACHE_TTL` | Optional cache TTL (seconds). Defaults to 30 |

You no longer need Google credentials—the data is read directly from the local
Excel file.

## Deployment (Render free tier example)

1. Push the project to GitHub.
2. Create a free account on [Render](https://render.com) → "New" → "Web Service".
3. Connect the GitHub repository that contains this app.
4. Use the following settings:
   - **Runtime:** Python
   - **Build command:** `pip install -r requirements.txt`
   - **Start command:** `gunicorn app:app` (or `gunicorn dashboard_app_from_excel.app:app` if the repo root contains additional files)
5. (Optional) Add Environment Variables in the Render dashboard:
   - `LOCAL_EXCEL_FILE` if the workbook lives somewhere else (otherwise the bundled `strategic_insight.xlsx` is used)
   - `SHEET_CACHE_TTL` if you want to adjust caching
6. Deploy. Render will build the image and provide a live URL.

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
- Ensure `strategic_insight.xlsx` stays in sync with the dashboard tabs (worksheet names must match those referenced in the code).
- Each time you push to the main branch, Render (or any connected platform) can auto-deploy the latest version.
