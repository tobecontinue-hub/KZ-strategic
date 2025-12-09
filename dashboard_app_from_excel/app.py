# app.py
from flask import Flask, render_template, abort
import plotly.express as px
import plotly
import json
import re
from collections import defaultdict
from collections import OrderedDict
from flask import redirect, url_for
from jinja2 import Environment, FileSystemLoader
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
        df = gs.sheet_to_df("P&L")  # ← Must match your sheet tab name exactly
    except Exception as e:
        print(f"❌ Error loading P&L: {e}")
        df = pd.DataFrame()

    rows = df.to_dict(orient="records") if not df.empty else []

    for row in rows:
        row["Date"] = str(row.get("Date", "")).strip()
        row["Revenue"] = float(row.get("Revenue", 0))
        row["Cost of Sales"] = float(row.get("Cost of Sales", 0))
        row["Gross Profit"] = float(row.get("Gross Profit", 0))
        row["Expense"] = float(row.get("Expense", 0))
        row["Net Profit"] = float(row.get("Net Profit", 0))

    return render_template(
        "profit_n_loss.html",
        rows=rows,
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
        columns=cols,
        rows=formatted_rows,
        insight=insight_text,
        title="2026 E-commerce Target"
    )

    
    
    
@app.route("/org_structure")
def org_structure():
    """
    Display the organizational structure as a hierarchical org chart.

    Loads org chart data from the 'org_chart' Google Sheet, builds a tree structure based on the 'Reports_To' field, and renders the org_structure.html template with the hierarchy.
    """
    try:
        df = gs.sheet_to_df("org_chart")
    except Exception as e:
        print(f"❌ Error loading org_chart: {e}")
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


# ===== ROUTES =====
@app.route("/")


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
        summary=summary_text,
        points=kpi_points,
        title="Executive Summary"
    )
    
import pandas as pd
from urllib.parse import quote # Added to handle spaces and special characters in URLs

@app.route("/home")
def home_page():
    return render_template("home.html")

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

    return render_template("profit_x.html", sections=section_groups)


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
    return render_template("top_product.html", rows=rows, title="Top 10 Product Forecast")




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
        "III. Motivation Factors": []
    }

    area_to_section = {
        "Core Values": "I. Core Values",
        "Hygiene Factors": "II. Hygiene Factors",
        "Motivation Factors": "III. Motivation Factors"
    }

    for row in rows:
        content_area = row.get("Content_Area", "").strip()
        if content_area in area_to_section:
            section = area_to_section[content_area]
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

    quarters = {}

    for _, row in df.iterrows():
        q = row["Quarter"]

        if q not in quarters:
            quarters[q] = []

        quarters[q].append({
            "Theme": row["Theme"],
            "Activity_ID": row["Activity_ID"],
            "Key_Activity": row["Key_Activity"]
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
        # Assuming gs.sheet_to_df("swot") returns the DataFrame with columns
        # including '2025' and '2026' as per the image.
        df = gs.sheet_to_df("swot")
    except Exception:
        # Fallback to an empty DataFrame if data fetching fails
        df = pd.DataFrame()

    rows = []
    try:
        # Convert DataFrame rows to a list of dictionaries
        rows = df.to_dict(orient="records") if hasattr(df, "to_dict") else df
    except Exception:
        rows = []

    # Initialize sections for S-W-O-T
    sections = {
        "Strength": [],
        "Opportunity": [],
        "Weakness": [],
        "Threat": []
    }

    for r in rows:
        cat = (r.get("Category") or "").strip()
        if not cat:
            continue

        # Normalize category
        cat_lower = cat.lower().strip()
        if "strength" in cat_lower:
            cat = "Strength"
        elif "opportun" in cat_lower:
            cat = "Opportunity"
        elif "weak" in cat_lower:
            cat = "Weakness"
        elif "threat" in cat_lower:
            cat = "Threat"
        else:
            continue  # Skip unknown categories

        # Extract data, including the separate 2025 and 2026 details
        item = {
            "id": r.get("Point_ID") or "",
            "title": r.get("Key_Item") or r.get("Key Item") or "",
            # Explicitly extract the content from the '2025' and '2026' columns
            "details_2025": r.get("2025", "") or "",
            "details_2026": r.get("2026", "") or "",
        }

        if cat in sections:
            sections[cat].append(item)

    return render_template("swot.html", sections=sections)


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
            ]
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



@app.route("/strategy_hub")
def strategy_hub():
    try:
        fna_df = gs.get_fna_performance()
        op_df = gs.get_operation_health()
        decisions_df = gs.get_key_decisions()
    except Exception as e:
        print("Error loading strategy data:", e)
        fna_df = op_df = decisions_df = pd.DataFrame()

    def clean_latex(text):
        if pd.isna(text) or not isinstance(text, str):
            return text
        text = re.sub(r'\\rightarrow', '→', text)
        text = re.sub(r'\$(.*?)\$', r'\1', text)
        text = re.sub(r'\\mathbf\{([^}]*)\}', r'\1', text)
        return text

    fna_rows = []
    if not fna_df.empty:
        for _, row in fna_df.iterrows():
            row = row.to_dict()
            for k in ["YTD Actual (2025)", "2026 Target", "Variance", "Rationale"]:
                row[k] = clean_latex(row.get(k, ""))
            fna_rows.append(row)

    op_rows = []
    if not op_df.empty:
        for _, row in op_df.iterrows():
            row = row.to_dict()
            row["Metric Value"] = clean_latex(row.get("Metric Value", ""))
            op_rows.append(row)

    decisions_rows = []
    if not decisions_df.empty:
        for _, row in decisions_df.iterrows():
            row = row.to_dict()
            for k in ["Decision/Activity", "Rationale (Linked to SWOT/Roadmap)"]:
                row[k] = clean_latex(row.get(k, ""))
            decisions_rows.append(row)

    return render_template(
        "strategy_hub.html",
        fna=fna_rows,
        operations=op_rows,
        decisions=decisions_rows,
        title="Strategy Hub: Performance, Ops & Decisions"
    )

@app.route("/charts")
def charts():
    try:
        df = gs.get_fna_performance()
        if df.empty:
            graphJSON = None
        else:
            numeric_cols = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
            if numeric_cols:
                fig = px.line(df, x=df.columns[0], y=numeric_cols, markers=True, title="FNA Performance Overview")
                graphJSON = json.dumps(fig, cls=plotly.utils.PlotlyJSONEncoder)
            else:
                graphJSON = None
    except Exception:
        graphJSON = None
    return render_template("charts.html", graphJSON=graphJSON)

if __name__ == "__main__":  
    app.run(debug=True, host="0.0.0.0", port=5000)
    