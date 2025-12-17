# app.py
from flask import Flask, render_template, abort
import plotly.express as px
import plotly
import json
import re
from collections import defaultdict
from collections import OrderedDict
from collections import Counter
from flask import redirect, url_for
from jinja2 import Environment, FileSystemLoader
from pathlib import Path
import pandas as pd

from services import google_sheets as gs

app = Flask(__name__, template_folder="templates", static_folder="static")

# ===== HELPER FUNCTIONS (PLACE AT TOP) =====

def safe_get_first(dlist, key_candidates):
    """Return first matching key value from a dict list's first record."""
    if not dlist:
        return None
    row = dlist[0]
    for k in key_candidates:
        if k in row and row[k] not in (None, ""):
            return row[k]
    for k in row:
        if k.lower() in [c.lower() for c in key_candidates]:
            return row[k]
    return None

def clean_number(val):
    """Convert '15,750' or '$\\mathbf{5,250}$' to float 15750.0"""
    if pd.isna(val) or val == "":
        return None
    s = str(val)
    s = re.sub(r'[^0-9.]', '', s)
    return float(s) if s else None

def clean_latex_math(text):
    """Clean LaTeX math notation for display: $17\%$ → 17%, \rightarrow → →, etc."""
    if not isinstance(text, str):
        return ""
    text = re.sub(r'\\mathbf\{([^}]*)\}', r'\1', text)
    text = re.sub(r'\\rightarrow', '→', text)
    text = re.sub(r'\$(.*?)\$', r'\1', text)
    return text.strip()

@app.route("/profit_n_loss")
def profit_n_loss():
    try:
        df = gs.sheet_to_df("P&L")
    except Exception as e:
        print(f"❌ Error loading P&L: {e}")
        df = pd.DataFrame()

    if df.empty:
        return render_template(
            "profit_n_loss.html",
            rows=[],
            title="Profit & Loss"
        )

    df.columns = (
        df.columns.astype(str)
        .str.strip()
        .str.replace("\u00A0", " ", regex=False)
    )

    def to_float(val):
        if pd.isna(val) or val == "":
            return 0.0
        try:
            return float(str(val).replace(",", "").strip())
        except Exception:
            return 0.0

    records = []
    for _, row in df.iterrows():
        date_val = str(row.get("Date", "")).strip()
        year = str(row.get("Year", "")).strip()
        month = str(row.get("Month", "")).strip()
        if date_val and (not year or not month):
            parts = re.split(r"[-/ ]", date_val)
            if len(parts) >= 2:
                month = month or parts[0]
                year = year or parts[1]
        entry = {
            "Year": year or "",
            "Month": month or "",
            "Label": f"{month} {year}".strip(),
            "Revenue": to_float(row.get("Revenue")),
            "Cost of Sales": to_float(row.get("Cost of Sales")),
            "Gross Profit": to_float(row.get("Gross Profit")),
            "Expense": to_float(row.get("Expense")),
            "Net Profit": to_float(row.get("Net Profit"))
        }
        records.append(entry)

    records.sort(key=lambda r: (r["Year"], r["Month"]))

    return render_template(
        "profit_n_loss.html",
        rows=records,
        title="Profit & Loss"
    )
    
@app.route("/financial_review")
def financial_review():
    try:
        df = gs.sheet_to_df("Financial Review 25")
    except Exception as e:
        print("Error loading sheet:", e)
        df = pd.DataFrame()

    # Clean column names (safe)
    df.columns = [c.strip() for c in df.columns]

    # Group by Section
    sections = {}
    if "Section" in df.columns:
        for section, group in df.groupby("Section"):
            sections[section] = group.to_dict(orient="records")

    return render_template("financial_review.html", sections=sections)

    
@app.template_filter("money")
def money(value):
    """Format numbers like 1,234,567.00"""
    try:
        value = float(str(value).replace(",", ""))
        return f"{value:,.2f}"
    except:
        return value

def parse_number(x):
    if pd.isna(x) or x == "":
        return None
    s = str(x)
    s = s.replace(",", "").replace(" ", "")
    s = re.sub(r"[^\d\.\-]", "", s)   # keep digits, dot, minus
    try:
        return float(s)
    except:
        return None

def fmt_money(x):
    # format as 0,000,000,000.00
    if x is None:
        return ""
    return "{:,.2f}".format(x)

def read_local_excel_sheet(sheet_name):
    base_dir = Path(__file__).resolve().parent
    primary = base_dir / "strategic_insight.xlsx"
    shadow = base_dir / "~$strategic_insight.xlsx"

    if primary.exists():
        try:
            return pd.read_excel(primary, sheet_name=sheet_name)
        except Exception as e:
            print(f"❌ Error loading {sheet_name} from strategic_insight.xlsx: {e}")
            return pd.DataFrame()

    if shadow.exists():
        try:
            return pd.read_excel(shadow, sheet_name=sheet_name)
        except Exception as e:
            print(f"❌ Error loading {sheet_name} from ~$strategic_insight.xlsx: {e}")

    return pd.DataFrame()

def format_number(value, decimals=0):
    if value in (None, ""):
        return "-"
    try:
        value = float(str(value).replace(",", ""))
    except Exception:
        return str(value)
    fmt = f"{{:,.{decimals}f}}"
    return fmt.format(value)

def format_percent(value):
    num = parse_number(value)
    if num is None:
        return "-"
    pct = num * 100 if abs(num) <= 1 else num
    return f"{pct:.1f}%"

def parse_review_text(text):
    if not text or (isinstance(text, float) and pd.isna(text)):
        return []
    chunks = []
    for raw in str(text).splitlines():
        line = raw.strip()
        if not line:
            continue
        if re.match(r"^(-|•|→|->|–)", line):
            cleaned = re.sub(r"^(-|•|→|->|–)\s*", "", line)
            chunks.append({"type": "bullet", "text": cleaned})
        else:
            chunks.append({"type": "text", "text": line})
    return chunks

@app.route("/ecom")
def ecom():
    try:
        df = gs.sheet_to_df("2026 Ecom Target")
    except Exception as e:
        print("Error loading sheet '2026 Ecom Target':", e)
        df = pd.DataFrame()

    # normalize columns
    if not df.empty:
        df.columns = (
            df.columns.astype(str)
            .str.replace("\u00A0", " ", regex=False)
            .str.replace("\t", " ", regex=False)
            .str.replace("  ", " ", regex=False)
            .str.strip()
        )

    # find insight row: either a row where first col contains 'Insight' (case-insensitive)
    insight_text = ""
    if not df.empty:
        first_col = df.columns[0]
        # rows where first column contains 'insight'
        mask = df[first_col].astype(str).str.strip().str.lower().str.contains(r"insight", na=False)
        if mask.any():
            # gather cell values from that row (join other columns)
            insight_row = df[mask].iloc[0]
            parts = []
            for c in df.columns[1:]:
                val = str(insight_row.get(c, "")) or ""
                if val.strip():
                    parts.append(val.strip())
            insight_text = " ".join(parts).strip()
        else:
            # fallback: try last non-empty row under any 'Insight' like label (sometimes it's below table)
            # look for any row where any cell contains 'Insight' or long paragraph
            long_text = ""
            for i in range(len(df)-1, -1, -1):
                row = df.iloc[i].astype(str).fillna("")
                joined = " ".join([x for x in row.tolist() if x and x.strip()])
                # heuristic: if joined length > 40 and doesn't look like a brand row (no numeric columns)
                if len(joined) > 40 and not re.search(r"\d{3,}", joined):
                    long_text = joined
                    break
            insight_text = long_text

    # Format numeric columns (detect columns with words like 'Amount','Target','Sales','Moonshot')
    formatted_rows = []
    if not df.empty:
        # decide columns order: keep original
        cols = list(df.columns)

        for _, r in df.iterrows():
            row_out = {}
            for c in cols:
                val = r.get(c, "")
                # numeric detection by column name
                if any(k in c.lower() for k in ["amount", "target", "sales", "moonshot", "fulfillment"]):
                    num = parse_number(val)
                    row_out[c] = fmt_money(num) if num is not None else ""
                else:
                    # keep original cell
                    row_out[c] = val if (val is not None and str(val).strip() != "nan") else ""
            formatted_rows.append(row_out)
    else:
        cols = []

    return render_template(
        "ecom.html",
        active_tab="target",
        target_columns=cols,
        target_rows=formatted_rows,
        target_insight=insight_text,
        comp_rows=[],
        comp_summary={},
        title="E-commerce Performance"
    )


@app.route("/ecom_comp")
def ecom_comparison():
    try:
        df = gs.sheet_to_df("ecom 2024 vs 2025")
    except Exception as e:
        print("Error loading sheet 'ecom 2024 vs 2025':", e)
        df = pd.DataFrame()

    df.columns = df.columns.astype(str).str.strip() if not df.empty else []
    if not df.empty:
        keep = [c for c in ["Months", "2024", "2025"] if c in df.columns]
        df = df[keep]

    def fmt_int(val):
        try:
            return f"{float(val):,.0f}"
        except Exception:
            return "0"

    records = []
    total_2024 = total_2025 = 0
    max_row = min_row = None

    if not df.empty:
        for _, row in df.iterrows():
            month = str(row.get("Months", "")).strip()
            val_2024 = parse_number(row.get("2024")) or 0
            val_2025 = parse_number(row.get("2025")) or 0
            entry = {
                "Months": month,
                "2024": val_2024,
                "2025": val_2025,
                "2024_fmt": fmt_int(val_2024),
                "2025_fmt": fmt_int(val_2025),
                "delta": val_2025 - val_2024,
                "delta_fmt": fmt_int(val_2025 - val_2024)
            }
            records.append(entry)
            total_2024 += val_2024
            total_2025 += val_2025

        if records:
            max_row = max(records, key=lambda r: r["2025"])
            eligible_min = [r for r in records if r["Months"].lower() not in ("dec", "december")]
            min_source = eligible_min if eligible_min else records
            min_row = min(min_source, key=lambda r: r["2025"])

    count = len(records) if records else 1
    comp_summary = {
        "total_2024": fmt_int(total_2024),
        "total_2025": fmt_int(total_2025),
        "avg_2024": fmt_int(total_2024 / count),
        "avg_2025": fmt_int(total_2025 / count),
        "max_month": max_row["Months"] if max_row else "-",
        "max_value": fmt_int(max_row["2025"]) if max_row else "0",
        "min_month": min_row["Months"] if min_row else "-",
        "min_value": fmt_int(min_row["2025"]) if min_row else "0",
    }

    return render_template(
        "ecom.html",
        active_tab="comparison",
        target_columns=[],
        target_rows=[],
        target_insight="",
        comp_rows=records,
        comp_summary=comp_summary,
        title="E-commerce Performance"
    )




@app.route("/strategy_plan")
def strategy_plan():
    try:
        df = gs.sheet_to_df("2026 Strategy plan")
    except Exception as e:
        print("Error loading 2026 Strategy plan:", e)
        df = pd.DataFrame()

    if df.empty:
        return render_template(
            "strategy_plan.html",
            goal_text="2026 Strategy Plan",
            pillars={},
            title="2026 Strategy Plan"
        )

    df.columns = df.columns.astype(str).str.strip()
    required = ["Goal", "Strategy Pillar", "Phase", "Quarter", "Action", "Photo_URL 1", "Photo_URL 2", "Photo_URL 3"]
    for col in required:
        if col not in df.columns:
            df[col] = ""

    df = df.fillna("")

    current_goal = ""
    pillars = OrderedDict()

    for _, row in df.iterrows():
        goal_raw = str(row.get("Goal", "")).strip()
        if goal_raw:
            current_goal = goal_raw
        pillar = str(row.get("Strategy Pillar", "")).strip() or "General"
        entry = {
            "goal": current_goal,
            "phase": row.get("Phase", ""),
            "quarter": row.get("Quarter", ""),
            "action": row.get("Action", ""),
            "photos": [
                url for url in [row.get("Photo_URL 1", ""), row.get("Photo_URL 2", ""), row.get("Photo_URL 3", "")]
                if url and str(url).strip()
            ]
        }
        pillars.setdefault(pillar, []).append(entry)

    goal_text = current_goal or "2026 Strategy Plan"

    return render_template(
        "strategy_plan.html",
        goal_text=goal_text,
        pillars=pillars,
        title="2026 Strategy Plan"
    )

@app.route("/org_structure")
def org_structure():
    """
    Display the organizational structure as a hierarchical org chart.

    Loads org chart data from the 'org_chart' Google Sheet, builds a tree structure based on the 'Reports_To' field, and renders the org_structure.html template with the hierarchy.
    """
    excel_path = Path(__file__).resolve().parent / "strategic_insight.xlsx"
    try:
        df = pd.read_excel(excel_path, sheet_name="org_chart")
    except Exception as e:
        print(f"❌ Error loading org_chart from Excel: {e}")
        df = pd.DataFrame()

    rows = df.to_dict(orient="records") if not df.empty else []

    employees = {}
    vacant_map = {}     # maps "(Vacant)" -> "Vacant_3"
    root_nodes = []

    # First pass: create all employee nodes
    for row in rows:
        original_name = str(row.get("Name", "")).strip()

        # Standardize name
        if not original_name or original_name.lower() == "(vacant)":
            generated_name = f"Vacant_{len(vacant_map)+1}"
            vacant_map[original_name] = generated_name
            name = generated_name
        else:
            name = original_name

        employees[name] = {
            "name": name,
            "original_name": original_name,
            "level": str(row.get("Level", "")).strip(),
            "department": row.get("Department", ""),
            "role": row.get("Role", ""),
            "status": row.get("Status", ""),
            "photo_url": row.get("Photo_URL", ""),
            "reports_to_raw": str(row.get("Reports_To", "")).strip(),
            "children": []
        }

    # Second pass: normalize Reports_To
    for name, emp in employees.items():
        reports_to = emp["reports_to_raw"]

        # Map vacant Reports_To to generated names
        if not reports_to or reports_to == "0":
            emp["reports_to"] = ""
        elif reports_to in vacant_map:
            emp["reports_to"] = vacant_map[reports_to]
        elif reports_to in employees:
            emp["reports_to"] = reports_to
        else:
            emp["reports_to"] = ""

    # Build hierarchy
    for name, emp in employees.items():
        parent = emp["reports_to"]

        if not parent or parent not in employees or parent == name:
            root_nodes.append(emp)
        else:
            employees[parent]["children"].append(emp)

    # Department color mapping (for department-wise card color)
    dept_colors = {
        "CEO": "#003366",
        "KHIT ZAY": "#007BFF",
        "LAST-MILE OPERATION TEAM": "#17a2b8",
        "BOB & CUSTOMER CARE TEAM": "#ffc107",
        "BUSINESS INTELLIGENCE TEAM": "#dc3545",
        "UI/UX TEAM": "#6f42c1",
        "DATA-BASED MARKETING TEAM": "#28a745",
    }
    # Role color mapping (for left border)
    role_colors = {
        "CEO": "#003366",
        "Dep. Head of Ecommerce": "#007BFF",
        "ASST MANAGER": "#17a2b8",
        "Executive": "#6f42c1",
        "JUNIOR": "#28a745",
        "DEVELOPER": "#6f42c1",
        "SNR DESIGNER": "#ffc107",
        "STAFF": "#dc3545",
        "BOB SALE DRIVE SUPERVISOR": "#ff9800",
        "BOB": "#ff9800",
        "CC Agent-VIP & Loyalty": "#ff9800",
        "CC Agent-Complaint": "#dc3545",
        "Vacant": "#b0b8c1",
        "EXECUTIVE (Shopper Marketing)": "#6f42c1",
        "EXECUTIVE (Buyer Marketing)": "#6f42c1",
        "STAFF (Data Analyst) + Virtual Fast Cash": "#28a745",
    }

    return render_template(
        "org_structure.html",
        root_nodes=root_nodes,
        dept_colors=dept_colors,
        role_colors=role_colors,
        title="Organizational Structure"
    )

import pandas as pd
from urllib.parse import quote # Added to handle spaces and special characters in URLs

@app.route("/home")
def home_page():
    return render_template("home.html")

# ===== ROUTES =====
@app.route("/")
def index():
    return redirect(url_for("home_page"))


@app.route("/executive_summary")
def executive_summary_page():
    # Load exe_summary data
    try:
        exe_df = gs.sheet_to_df("exe_summary")
        exe_rows = exe_df.to_dict(orient="records") if not exe_df.empty else []
    except Exception:
        exe_rows = []

    # Load brand_promise data
    try:
        brand_df = gs.sheet_to_df("brand_promise")
        brand_rows = brand_df.to_dict(orient="records") if not brand_df.empty else []
    except Exception:
        brand_rows = []

    # Parse brand promise data
    brand_promise = ""
    mission_statement = ""
    dashboard_subtitle = ""
    strategic_insight = ""
    summary_override = None

    for row in brand_rows:
        key = row.get("Content_Key", "")
        value = row.get("Content_Value", "")
        if "Brand_Promise" in key:
            brand_promise = value
        elif "Mission_Statement" in key:
            mission_statement = value
        elif "Dashboard_Subtitle" in key:
            dashboard_subtitle = value
        elif "Strategic Insight" in key:
            strategic_insight = value
        elif "Summary" in key:
            summary_override = value

    # Parse executive summary data
    summary_text = None
    kpi_points = []

    for row in exe_rows:
        category = str(row.get("Category", "")).strip()
        if category.lower() == "summary":
            summary_text = row.get("Key_Insight", "")
        else:
            kpi_points.append(row)

    return render_template(
        "executive_summary.html",
        brand_promise=brand_promise,
        mission_statement=mission_statement,
        dashboard_subtitle=dashboard_subtitle,
        strategic_insight=strategic_insight,
        summary=summary_override or summary_text,
        points=kpi_points,
        title="Executive Summary"
    )
    
@app.route("/top_product_promo")
def top_product_promo():
    try:
        df = gs.sheet_to_df('top_product_promo')
        # Normalize headers
        if not df.empty:
            df.columns = (
                df.columns.astype(str)
                .str.replace("\u00A0", " ", regex=False)
                .str.replace("\t", " ", regex=False)
                .str.replace("  ", " ", regex=False)
                .str.strip()
            )

            # Rename sheet columns to template fields
            rename_map = {}
            for c in df.columns:
                lc = c.lower()
                if "2025 qty" in lc or "qty" in lc:
                    rename_map[c] = "Current_Qty"
                elif "insight" in lc:
                    rename_map[c] = "Insight"
                elif "action" in lc:
                    rename_map[c] = "Strategy_Focus"
                elif lc == "no":
                    rename_map[c] = "No"
                elif lc == "product":
                    rename_map[c] = "Product"
                elif "sale" in lc:
                    rename_map[c] = "Sale_Total"
                elif "q1" in lc:
                    rename_map[c] = "Q1_Forecast"
                elif "q2" in lc:
                    rename_map[c] = "Q2_Forecast"
                elif "q3" in lc:
                    rename_map[c] = "Q3_Forecast"
                elif "q4" in lc:
                    rename_map[c] = "Q4_Forecast"
                elif "image_url" in lc or "image" in lc:
                    rename_map[c] = "Image_URL"
            if rename_map:
                df = df.rename(columns=rename_map)

            # Clean numeric-like strings; keep format for display later
            def numberish_to_float(x):
                if pd.isna(x) or x == "":
                    return None
                s = str(x).strip()
                s = s.replace(",", "")
                # allow letters like M/B to stay as-is for KPI?
                # For sold qty and forecast we expect plain numbers.
                try:
                    return float(s)
                except:
                    return None

            # Convert relevant numeric columns
            for col in ["Current_Qty", "Sale_Total", "Q1_Forecast", "Q2_Forecast", "Q3_Forecast", "Q4_Forecast"]:
                if col in df.columns:
                    df[col] = df[col].apply(numberish_to_float)

            # Ensure required columns exist
            for col in ["No", "Product", "Current_Qty", "Sale_Total", "Insight", "Strategy_Focus", "Q1_Forecast", "Q2_Forecast", "Q3_Forecast", "Q4_Forecast", "Image_URL"]:
                if col not in df.columns:
                    df[col] = "" if col in ["Insight", "Strategy_Focus", "Product", "Image_URL"] else None

    except Exception as e:
        print("Error loading top_product_promo:", e)
        df = pd.DataFrame()

    rows = df.to_dict(orient="records") if not df.empty else []
    return render_template("top_product_promo.html", rows=rows, title="Top 10 Product Promo Forecast")
@app.route("/value_map")
def value_map():
    df = gs.sheet_to_df("Full price value_map")

    # Clean headers
    df.columns = (
        df.columns.astype(str)
        .str.replace("\u00A0", " ", regex=False)
        .str.replace("\t", " ", regex=False)
        .str.replace("  ", " ", regex=False)
        .str.strip()
        .str.replace(" ", "_")
    )

    print("CLEANED COLUMNS:", df.columns.tolist())

    # Auto-map columns to UI fields
    df = df.rename(columns={
        "Point": "Key_Identifier",
        "Key_Insight": "Value/Headline",
        "Highlight": "Details/Rationale",
        "Current_Status": "Details/Rationale",   # fallback if Highlight empty
        "Category": "Content_Category",
    })

    # Fill missing fields
    for col in ["Key_Identifier", "Value/Headline", "Details/Rationale"]:
        if col not in df.columns:
            df[col] = ""

    print("UNIQUE CATEGORIES:", df["Content_Category"].unique().tolist())

    # Normalize content category for robust filtering
    cat = df["Content_Category"].astype(str).str.strip().str.lower()

    pains = df[cat == "pain"].to_dict(orient='records')
    relievers = df[cat == "pain reliever"].to_dict(orient='records')
    gains = df[cat == "gain"].to_dict(orient='records')
    creators = df[cat == "gain creator"].to_dict(orient='records')
    activities = df[cat == "activity"].to_dict(orient='records')

    products = df[cat == "top-performing full-price brands"].to_dict(orient='records')

    services = df[cat.isin([
        "delivery",
        "return",
        "customer service",
        "delivery tracking",
        "service",
    ])].to_dict(orient='records')

    demo = df[cat.isin([
        "core customer demo",
        "core customer geo",
        "core customer demo + geo",
    ])].to_dict(orient='records')

    # "Just to be done" handling (JTBD): match common variants
    jtbd_mask = (
        cat.str.contains(r"\bjust\b", na=False) & cat.str.contains(r"done|to be done|tbd|jtbd", na=False)
    ) | (cat == "just") | (cat == "just to be done") | (cat.str.contains(r"jtbd", na=False))
    justtobedone = df[jtbd_mask].to_dict(orient='records')

    return render_template(
        "value_map.html",
        pains=pains,
        relievers=relievers,
        gains=gains,
        creators=creators,
        activities=activities,
        products=products,
        services=services,
        demo=demo,
        justtobedone=justtobedone,
        title="Value Proposition (Full Price)"
    )

@app.route("/value_map_promo")
def value_map_promo():
    df = gs.sheet_to_df("promo price value_map")

    # Clean headers
    df.columns = (
        df.columns.astype(str)
        .str.replace("\u00A0", " ", regex=False)
        .str.replace("\t", " ", regex=False)
        .str.replace("  ", " ", regex=False)
        .str.strip()
        .str.replace(" ", "_")
    )

    print("CLEANED COLUMNS:", df.columns.tolist())

    # Auto-map columns to UI fields
    df = df.rename(columns={
        "Point": "Key_Identifier",
        "Key_Insight": "Value/Headline",
        "Highlight": "Details/Rationale",
        "Current_Status": "Details/Rationale",   # fallback if Highlight empty
        "Category": "Content_Category",
    })

    # Fill missing fields
    for col in ["Key_Identifier", "Value/Headline", "Details/Rationale"]:
        if col not in df.columns:
            df[col] = ""

    print("UNIQUE CATEGORIES:", df["Content_Category"].unique().tolist())

    # Normalize content category for robust filtering
    cat = df["Content_Category"].astype(str).str.strip().str.lower()

    pains = df[cat == "pain"].to_dict(orient='records')
    relievers = df[cat == "pain reliever"].to_dict(orient='records')
    gains = df[cat == "gain"].to_dict(orient='records')
    creators = df[cat == "gain creator"].to_dict(orient='records')
    activities = df[cat == "activity"].to_dict(orient='records')

    products = df[cat == "top promo brands (actual performance):"].to_dict(orient='records')

    services = df[cat.isin([
        "cod",
        "return",
        "customer service",
        "delivery tracking",
        "service",
    ])].to_dict(orient='records')

    demo = df[cat.isin([
        "new customer demo",
        "new customer geo",
        "new customer demo + geo",
    ])].to_dict(orient='records')

    # "Just to be done" handling (JTBD): match common variants
    jtbd_mask = (
        cat.str.contains(r"\bjust\b", na=False) & cat.str.contains(r"done|to be done|tbd|jtbd", na=False)
    ) | (cat == "just") | (cat == "just to be done") | (cat.str.contains(r"jtbd", na=False))
    justtobedone = df[jtbd_mask].to_dict(orient='records')

    return render_template(
        "value_map_promo.html",
        pains=pains,
        relievers=relievers,
        gains=gains,
        creators=creators,
        activities=activities,
        products=products,
        services=services,
        demo=demo,
        justtobedone=justtobedone,
        title="Value Proposition (Promo Price)"
    )

@app.route("/trajectories")
def trajectories_page():
    try:
        df = gs.sheet_to_df("trajectories")
    except Exception as e:
        print(f"❌ Error loading trajectories: {e}")
        df = pd.DataFrame()

    rows = df.to_dict(orient="records") if not df.empty else []

    # Group by Section_ID
    sections = {
        "I. Trajectories": [],
        "II. Summary": []
    }

    for row in rows:
        section_id = row.get("Section_ID", "").strip()
        if not section_id:
            continue

        # Trajectories: T1, T2, T3, T4, T5
        if section_id.startswith("T") and section_id[1:].isdigit():
            sections["I. Trajectories"].append(row)
        # Summary: S1, S2, S3, ...
        elif section_id.startswith("S") and section_id[1:].isdigit():
            sections["II. Summary"].append(row)
        # Optional: Handle other sections if needed

    return render_template(
        "trajectories.html",
        sections=sections,
        title="Trajectories & Strategic Insights"
    )
@app.route("/retail_swift_online")
def retail_swift_online():
    try:
        df = gs.sheet_to_df("Retail_Swift_Online")
    except Exception as e:
        print("Error loading Retail Swift Online data:", e)
        df = pd.DataFrame()

    # Clean NaNs to avoid template errors
    df = df.fillna("")
    
    return render_template("retail_swift.html", data=df.to_dict("records"))

@app.route("/segments")
def segments_page():
    try:
        df = gs.sheet_to_df("Core -New segments")
    except Exception as e:
        print("GS ERROR:", e)
        df = pd.DataFrame()

    # Clean NaNs to avoid template errors
    df = df.fillna("")

    return render_template("segments.html", data=df.to_dict("records"))


def clean_html_breaks(text):
    if not isinstance(text, str):
        return text
    return (
        text.replace("\\n", "<br>")
            .replace("\n", "<br>")
            .replace("&lt;br&gt;", "<br>")
            .replace("<br/>", "<br>")
    )


@app.route("/profit_x")
@app.route("/profit_per_x")
def profit_x():
    try:
        df = gs.sheet_to_df("Profit per X")
    except Exception as e:
        print("GS error:", e)
        df = pd.DataFrame()

    # Ensure missing columns do not break template
    required_cols = ["Section", "Segment", "Data", "Insight", "What to Improve More? (2026 Actions)"]
    for col in required_cols:
        if col not in df.columns:
            df[col] = ""

    # Clean HTML breaks in relevant columns
    for col in ["Data", "Insight", "What to Improve More? (2026 Actions)"]:
        if col in df.columns:
            df[col] = df[col].apply(clean_html_breaks)
    
    # Trim spaces
    df.columns = df.columns.str.strip()
    df["Section"] = df["Section"].fillna("").str.strip()

    # Split into sections
    section_groups = {}
    for sec in df["Section"].unique():
        if sec == "": 
            continue
        section_groups[sec] = df[df["Section"] == sec]

    return render_template("profit_per_x.html", sections=section_groups)


# ✅ SINGLE /top_product route
@app.route("/top_product")
def top_product():
    try:
        df = gs.sheet_to_df('top_product_full_price')
        # Normalize headers
        if not df.empty:
            df.columns = (
                df.columns.astype(str)
                .str.replace("\u00A0", " ", regex=False)
                .str.replace("\t", " ", regex=False)
                .str.replace("  ", " ", regex=False)
                .str.strip()
            )

            # Rename sheet columns to template fields
            rename_map = {}
            for c in df.columns:
                lc = c.lower()
                if "2025 qty" in lc or "qty" in lc:
                    rename_map[c] = "Current_Qty"
                elif "insight" in lc:
                    rename_map[c] = "Insight"
                elif "action" in lc:
                    rename_map[c] = "Strategy_Focus"
                elif lc == "no":
                    rename_map[c] = "No"
                elif lc == "product":
                    rename_map[c] = "Product"
                elif "sale" in lc:
                    rename_map[c] = "Sale_Total"
                elif "q1" in lc:
                    rename_map[c] = "Q1_Forecast"
                elif "q2" in lc:
                    rename_map[c] = "Q2_Forecast"
                elif "q3" in lc:
                    rename_map[c] = "Q3_Forecast"
                elif "q4" in lc:
                    rename_map[c] = "Q4_Forecast"
                elif "image_url" in lc or "image" in lc:
                    rename_map[c] = "Image_URL"
            if rename_map:
                df = df.rename(columns=rename_map)

            # Clean numeric-like strings; keep format for display later
            def numberish_to_float(x):
                if pd.isna(x) or x == "":
                    return None
                s = str(x).strip()
                s = s.replace(",", "")
                # allow letters like M/B to stay as-is for KPI?
                # For sold qty and forecast we expect plain numbers.
                try:
                    return float(s)
                except:
                    return None

            # Convert relevant numeric columns
            for col in ["Current_Qty", "Sale_Total", "Q1_Forecast", "Q2_Forecast", "Q3_Forecast", "Q4_Forecast"]:
                if col in df.columns:
                    df[col] = df[col].apply(numberish_to_float)

            # Ensure required columns exist
            for col in ["No", "Product", "Current_Qty", "Sale_Total", "Insight", "Strategy_Focus", "Q1_Forecast", "Q2_Forecast", "Q3_Forecast", "Q4_Forecast", "Image_URL"]:
                if col not in df.columns:
                    df[col] = "" if col in ["Insight", "Strategy_Focus", "Product", "Image_URL"] else None

    except Exception as e:
        print("Error loading top_product:", e)
        df = pd.DataFrame()

    rows = df.to_dict(orient="records") if not df.empty else []

    # helper to normalize product names for matching between sheets
    def normalize_name(name):
        return str(name).strip().lower()

    # attach normalized key to each top product row
    for r in rows:
        r["ProductKey"] = normalize_name(r.get("Product", ""))

    # Load three_offer sheet for per-product offers
    offers_by_product = defaultdict(list)
    try:
        offers_df = gs.sheet_to_df("three_offer")
    except Exception as e:
        print("Error loading three_offer:", e)
        offers_df = pd.DataFrame()

    if not offers_df.empty:
        offers_df.columns = offers_df.columns.astype(str).str.strip()

        for col in ["Product", "Offer_Product", "Offer", "Photo_URL"]:
            if col not in offers_df.columns:
                offers_df[col] = ""

        for _, r in offers_df.iterrows():
            product_name = str(r.get("Product", "")).strip()
            if not product_name:
                continue
            key = normalize_name(product_name)
            offer_text = str(r.get("Offer", "")).strip()
            offer_label = str(r.get("Offer_Product", "")).strip()
            photo_url = str(r.get("Photo_URL", "")).strip()

            offers_by_product[key].append({
                "title": offer_text or "Offer details coming soon",
                "img": photo_url,
                "label": offer_label or "Offer",
            })

        # only keep first three items per product to match UI expectation
        for product_key, product_offers in offers_by_product.items():
            offers_by_product[product_key] = product_offers[:3]

    # convert defaultdict to plain dict for safe JSON serialization in the template
    offers_by_product = dict(offers_by_product)

    return render_template("top_product.html", rows=rows, offers_by_product=offers_by_product, title="Top 10 Product Forecast")




def clean_annual_total(value):
    """Remove LaTeX like \\mathbf{5250} or $...$ and return int."""
    if not isinstance(value, str):
        return value
    value = re.sub(r'\\mathbf\{([^}]*)\}', r'\1', value)   # remove \mathbf{}
    value = re.sub(r'\$(.*?)\$', r'\1', value)             # remove $...$
    value = value.replace(",", "")
    try:
        return int(value)
    except:
        return None


# ✅ /dna AFTER clean_latex_math is defined
@app.route("/dna")
def dna_page():
    try:
        df = gs.get_dna()
    except Exception as e:
        print(f"❌ Error loading DNA: {e}")
        df = pd.DataFrame()

    rows = df.to_dict(orient="records") if not df.empty else []

    # Clean LaTeX math
    def clean_latex_math(text):
        if not isinstance(text, str):
            return ""
        text = re.sub(r'\\mathbf\{([^}]*)\}', r'\1', text)
        text = re.sub(r'\\rightarrow', '→', text)
        text = re.sub(r'\$(.*?)\$', r'\1', text)
        return text.strip()

    # Group by Content_Area
    sections = {
        "I. Core Values": [],
        "II. Hygiene Factors": [],
        "III. Motivation Factors": [],
        "Strategic Insight": []
    }

    area_to_section = {
        "core values": "I. Core Values",
        "hygiene factors": "II. Hygiene Factors",
        "motivation factors": "III. Motivation Factors",
        "strategic insight": "Strategic Insight"
    }

    for row in rows:
        content_area = str(row.get("Content_Area", "")).strip()
        section_key = content_area.lower()
        if section_key in area_to_section:
            section = area_to_section[section_key]
            cleaned_row = {
                "Point_ID": row.get("Point_ID", ""),
                "Key_Item": row.get("Key_Item", ""),
                "DNA": clean_latex_math(row.get("DNA", "")),
                "Details/Data_Alignment": clean_latex_math(row.get("Details/Data_Alignment", "")),
                "type": content_area
            }
            sections[section].append(cleaned_row)

    # Sort each section by Point_ID
    def sort_key(item):
        pid = item["Point_ID"]
        if pid.startswith("V"):
            return (0, int(pid[1:]) if pid[1:].isdigit() else 0)
        elif pid.startswith("H"):
            return (1, int(pid[1:]) if pid[1:].isdigit() else 0)
        elif pid.startswith("M"):
            return (2, int(pid[1:]) if pid[1:].isdigit() else 0)
        else:
            return (3, pid)

    for section in sections:
        sections[section].sort(key=sort_key)

    # Print for debugging
    print(f"Sections: {sections}")

    return render_template(
        "dna.html",
        sections=sections,
        title="Organizational DNA",
        description="The foundational values, hygiene factors, and motivation drivers that shape our culture and performance."
    )
    
@app.route("/roadmap")
def roadmap_page():
    try:
        df = gs.get_roadmap()
    except Exception as e:
        print(f"❌ Error loading roadmap: {e}")
        df = pd.DataFrame()

    if df.empty:
        return render_template("roadmap.html", quarters={}, quarter_order=[])

    df.columns = [c.strip() for c in df.columns]

    records = df.to_dict(orient="records")

    def get_value(row, *keys):
        for key in keys:
            key_lower = key.strip().lower()
            for actual_key, value in row.items():
                if actual_key and actual_key.strip().lower() == key_lower:
                    return "" if value in (None, "") else str(value)
        return ""

    quarters = {}

    for row in records:
        q = get_value(row, "Quarter") or "Unassigned"
        quarters.setdefault(q, []).append({
            "Activity_ID": get_value(row, "Activity_ID", "Activity ID"),
            "Topic": get_value(row, "Key Topic", "Key_Topic", "Key Activity"),
            "Owner": get_value(row, "Owner"),
        })

    quarter_order = sorted(quarters.keys())

    return render_template(
        "roadmap.html",
        quarters=quarters,
        quarter_order=quarter_order
    )




    
@app.route("/swot")
def swot_page():
    try:
        df = gs.sheet_to_df("swot")
    except Exception:
        df = pd.DataFrame()

    rows = []
    try:
        rows = df.to_dict(orient="records") if hasattr(df, "to_dict") else []
    except Exception:
        rows = []

    sections = OrderedDict()
    key_insights = []

    for row in rows:
        category = (row.get("Category") or "").strip()
        point_id = row.get("Point_ID") or ""
        key_item = row.get("Key_Item") or row.get("Key Item") or ""
        insight_2025 = row.get("2025") or row.get("2025 Insight") or ""
        strategy_2026 = row.get("2026") or row.get("2026 Strategy") or ""

        if not category:
            continue

        if category.lower() == "key insight":
            key_insights.append({
                "title": key_item,
                "content": insight_2025 or strategy_2026 or point_id
            })
            continue

        sections.setdefault(category, []).append({
            "id": point_id,
            "title": key_item,
            "details_2025": insight_2025,
            "details_2026": strategy_2026
        })

    return render_template(
        "swot.html",
        sections=sections,
        key_insights=key_insights
    )


@app.route("/cost_per_x")
def cost_per_x():
    try:
        df = gs.sheet_to_df("Cost per X")
    except Exception as e:
        print("GS error (Cost per X):", e)
        df = pd.DataFrame()

    rows = []
    if not df.empty:
        df.columns = (
            df.columns.astype(str)
            .str.replace("\u00A0", " ", regex=False)
            .str.replace("\t", " ", regex=False)
            .str.replace("  ", " ", regex=False)
            .str.strip()
        )

        # Map probable headers to canonical names
        rename_map = {}
        for c in df.columns:
            lc = c.lower()
            if "cost per x" in lc:
                rename_map[c] = "Cost per X"
            elif lc.startswith("facts"):
                rename_map[c] = "Facts"
            elif lc.startswith("why"):
                rename_map[c] = "Why?"
            elif "what to improve" in lc or "improve" in lc:
                rename_map[c] = "What to Improve More?"

        if rename_map:
            df = df.rename(columns=rename_map)

        # Ensure required columns exist
        for col in ["Cost per X", "Facts", "Why?", "What to Improve More?"]:
            if col not in df.columns:
                df[col] = ""

        # Normalize newlines for display
        for col in ["Facts", "Why?", "What to Improve More?"]:
            df[col] = (
                df[col]
                .fillna("")
                .astype(str)
                .str.replace("\r\n", "\n")
                .str.replace("\r", "\n")
            )

        rows = (
            df[["Cost per X", "Facts", "Why?", "What to Improve More?"]]
            .fillna("")
            .to_dict(orient="records")
        )

    return render_template("cost_per_x.html", rows=rows, title="Cost per X")


@app.route("/okr")
def okr_page():
    try:
        df = gs.sheet_to_df("okr")  # ← Sheet name = "okr"
    except Exception as e:
        print(f"❌ Error loading OKR: {e}")
        df = pd.DataFrame()

    rows = df.to_dict(orient="records") if not df.empty else []

    # Separate 2025 and 2026
    okr_2025 = []
    okr_2026 = []

    for row in rows:
        year = str(row.get("Years", "")).strip()
        if "2025" in year:
            okr_2025.append(row)
        elif "2026" in year:
            okr_2026.append(row)

    # Group by Functional Team
    teams_2025 = defaultdict(list)
    teams_2026 = defaultdict(list)

    for row in okr_2025:
        team = row.get("Functional POVs", "Other").strip()
        teams_2025[team].append(row)

    for row in okr_2026:
        team = row.get("Functional POVs", "Other").strip()
        teams_2026[team].append(row)

    # Merge all teams (union of keys)
    all_teams = sorted(set(teams_2025.keys()) | set(teams_2026.keys()))

    # Build comparison structure
    comparison = []

    def compute_avg(rows, field="Average"):
        values = []
        for r in rows:
            v = r.get(field)
            if v in (None, ""):
                continue
            s = str(v).replace("%", "").strip()
            try:
                # Treat sheet values like 0.7 as 70%
                values.append(float(s) * 100.0)
            except Exception:
                continue
        if not values:
            return None
        return round(sum(values) / len(values), 1)

    for team in all_teams:
        items_2025 = teams_2025.get(team, [])
        items_2026 = teams_2026.get(team, [])

        # Group by Objective
        obj_map = {}
        for item in items_2025:
            obj = item.get("Objective", "No Objective")
            if obj not in obj_map:
                obj_map[obj] = {"2025": [], "2026": []}
            obj_map[obj]["2025"].append(item)

        for item in items_2026:
            obj = item.get("Objective", "No Objective")
            if obj not in obj_map:
                obj_map[obj] = {"2025": [], "2026": []}
            obj_map[obj]["2026"].append(item)

        comparison.append({
            "team": team,
            "objectives": [
                {
                    "objective": obj,
                    "items_2025": data["2025"],
                    "items_2026": data["2026"]
                }
                for obj, data in obj_map.items()
            ],
            "avg_2025": compute_avg(items_2025),
            "avg_2026": compute_avg(items_2026),
        })

    return render_template(
        "okr.html",
        comparison=comparison,
        title="OKR Dashboard: 2025 vs 2026"
    )

def clean_number(val):
    """Convert '3.7 B' or '180 L' to float"""
    if not isinstance(val, str):
        return val
    val = val.strip()
    if not val:
        return val
    
    # Handle "3.7 B" → 3700000000
    match = re.search(r'([\d.]+)\s*([BM]?)', val)
    if match:
        num = float(match.group(1))
        unit = match.group(2)
        if unit == 'B':
            return num * 1_000_000_000
        elif unit == 'M':
            return num * 1_000_000
        elif unit == 'L':
            return num * 100_000
        else:
            return num
    return val


def load_fna_performance_from_excel():
    excel_path = Path(__file__).resolve().parent / "strategic_insight.xlsx"
    try:
        df = pd.read_excel(excel_path, sheet_name="fna_performance")
    except Exception as e:
        print(f"❌ Error reading FNA performance sheet: {e}")
        return [], Counter()

    if df.empty:
        return [], Counter()

    df.columns = (
        df.columns.astype(str)
        .str.strip()
        .str.replace("\u00A0", " ", regex=False)
    )

    records = df.fillna("").to_dict(orient="records")

    category_counter = Counter()
    for row in records:
        category = str(row.get("KPI Category", "General")).strip() or "General"
        category_counter[category] += 1
        for key, value in row.items():
            if isinstance(value, str):
                cleaned = re.sub(r"\\mathbf\{([^}]*)\}", r"\1", value)
                cleaned = re.sub(r"\\rightarrow", "→", cleaned)
                cleaned = re.sub(r"\$(.*?)\$", r"\1", cleaned)
                row[key] = cleaned.strip()

    return records, category_counter


@app.route("/fna_performance")
def fna_performance_page():
    fna_rows, category_counter = load_fna_performance_from_excel()

    palette = [
        "#1976d2", "#ef6c00", "#2e7d32",
        "#6a1b9a", "#00838f", "#c62828"
    ]
    category_meta = {}
    for idx, (category, count) in enumerate(category_counter.items()):
        category_meta[category] = {
            "color": palette[idx % len(palette)],
            "count": count
        }

    return render_template(
        "fna_performance.html",
        fna=fna_rows,
        category_meta=category_meta,
        title="FNA Performance"
    )


def split_operations(rows):
    insights = []
    stages = defaultdict(list)
    status_counter = Counter()
    for row in rows:
        stage = str(row.get("Funnel Stage", "")).strip()
        status = str(row.get("Status", "")).strip()
        if stage.lower() == "insight":
            insights.append(row)
            continue
        stages[stage or "Unassigned"].append(row)
        if status:
            status_counter[status] += 1
    return insights, stages, status_counter


@app.route("/operation_health")
def operation_health_page():
    try:
        df = gs.get_operation_health()
    except Exception as e:
        print(f"❌ Error loading operation health: {e}")
        df = pd.DataFrame()

    if df.empty:
        return render_template(
            "operation_health.html",
            grouped_ops={},
            insights=[],
            status_counter={},
            title="Operations Health"
        )

    df.columns = df.columns.astype(str).str.strip()
    records = df.fillna("").to_dict(orient="records")
    insights, grouped_ops, status_counter = split_operations(records)

    return render_template(
        "operation_health.html",
        grouped_ops=grouped_ops,
        insights=insights,
        status_counter=status_counter,
        title="Operations Health"
    )


@app.route("/bob")
def bob_page():
    bob_df = read_local_excel_sheet("BOB")
    review_df = read_local_excel_sheet("BOB_review")

    def normalize_columns(df):
        if df.empty:
            return df
        df.columns = df.columns.astype(str).str.strip()
        return df

    bob_df = normalize_columns(bob_df)
    review_df = normalize_columns(review_df)

    # Ensure numeric table columns exist even if sheet uses variants like 'Grand_Total '
    bob_column_map = {
        "months": "Months",
        "month": "Months",
        "bob order": "BOB Order",
        "bob": "BOB Order",
        "boborder": "BOB Order",
        "self order": "Self Order",
        "self": "Self Order",
        "grand total": "Grand Total",
        "total": "Grand Total",
        "cs%": "CS%",
        "cs %": "CS%",
        "cs percentage": "CS%"
    }

    if not bob_df.empty:
        rename_map = {}
        for col in bob_df.columns:
            key = col.strip().lower()
            if key in bob_column_map:
                rename_map[col] = bob_column_map[key]
        if rename_map:
            bob_df = bob_df.rename(columns=rename_map)

    bob_rows = []
    chart_data = {"months": [], "bob": [], "self": [], "cs": []}
    totals = {"bob": 0, "self": 0, "grand": 0, "cs": []}
    best_month = {"label": "-", "value": 0}

    if not bob_df.empty:
        for _, row in bob_df.iterrows():
            month = str(row.get("Months", "")).strip()
            bob_val = parse_number(row.get("BOB Order")) or 0
            self_val = parse_number(row.get("Self Order")) or 0
            grand_val = parse_number(row.get("Grand Total")) or 0
            cs_val = parse_number(row.get("CS%"))

            bob_rows.append({
                "Months": month,
                "BOB Order": format_number(bob_val),
                "Self Order": format_number(self_val),
                "Grand Total": format_number(grand_val),
                "CS%": format_percent(cs_val)
            })

            chart_data["months"].append(month)
            chart_data["bob"].append(bob_val)
            chart_data["self"].append(self_val)
            chart_data["cs"].append((cs_val * 100) if (cs_val is not None and abs(cs_val) <= 1) else (cs_val or 0))

            totals["bob"] += bob_val
            totals["self"] += self_val
            totals["grand"] += grand_val
            if cs_val is not None:
                totals["cs"].append(cs_val if abs(cs_val) <= 1 else cs_val / 100)

            if grand_val > best_month["value"]:
                best_month = {"label": month, "value": grand_val}

    review_sections = [
        {"key": "worked", "label": "What Worked?"},
        {"key": "scale", "label": "What needs to scale?"},
        {"key": "not_work", "label": "What did not work?"},
        {"key": "lesson", "label": "What is the lesson learned?"},
        {"key": "next_goal", "label": "What is the next goal for BOB?"}
    ]

    reviews = []
    if not review_df.empty:
        column_map = {
            "what worked?": "worked",
            "what needs to scale?": "scale",
            "what did not work?": "not_work",
            "what is the lesson learned?": "lesson",
            "what is the next goal for bob?": "next_goal"
        }
        keys_defaults = {sec["key"]: [] for sec in review_sections}
        for _, row in review_df.iterrows():
            entry = dict(keys_defaults)
            for col in review_df.columns:
                normalized = str(col).strip().lower()
                if normalized in column_map:
                    entry[column_map[normalized]] = parse_review_text(row.get(col, ""))
            if any(entry.values()):
                reviews.append(entry)

    summary = {
        "total_bob": format_number(totals["bob"]),
        "total_self": format_number(totals["self"]),
        "total_grand": format_number(totals["grand"]),
        "avg_cs": format_percent(sum(totals["cs"]) / len(totals["cs"]) if totals["cs"] else None),
        "best_month": best_month["label"],
        "best_month_value": format_number(best_month["value"])
    }

    return render_template(
        "bob.html",
        rows=bob_rows,
        reviews=reviews,
        review_sections=review_sections,
        summary=summary,
        chart_data=chart_data,
        title="BOB Performance",
        description="Monthly BOB volume split with qualitative learnings"
    )


if __name__ == "__main__":  
    app.run(debug=True, host="0.0.0.0", port=5000)
    