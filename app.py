import os
import io
import re
import zipfile
import streamlit as st
import pandas as pd
import pdfplumber
from datetime import datetime

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="Alcorn PDF Extractor", layout="wide")

DEBUG = False  # set True only for local debugging

# =========================
# AUTH (safe for GitHub)
# Prefer Streamlit secrets:
# [auth]
# username="matt"
# password="Interlynx123"
# =========================
def get_valid_credentials():
    u = None
    p = None

    if "auth" in st.secrets:
        u = st.secrets["auth"].get("username", u)
        p = st.secrets["auth"].get("password", p)

    if u is None:
        u = os.getenv("APP_USERNAME")
    if p is None:
        p = os.getenv("APP_PASSWORD")

    # fallback (your requested creds)
    if u is None or p is None:
        u = "matt"
        p = "Interlynx123"
    return u, p

VALID_USERNAME, VALID_PASSWORD = get_valid_credentials()

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.title("Alcorn PDF Extractor — Login")
    u = st.text_input("Username")
    p = st.text_input("Password", type="password")
    if st.button("Login"):
        if u == VALID_USERNAME and p == VALID_PASSWORD:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("Invalid username or password.")
    st.stop()

# =========================
# HELPERS
# =========================
MONTHS = {
    "jan": 1, "january": 1,
    "feb": 2, "february": 2,
    "mar": 3, "march": 3,
    "apr": 4, "april": 4,
    "may": 5,
    "jun": 6, "june": 6,
    "jul": 7, "july": 7,
    "aug": 8, "august": 8,
    "sep": 9, "september": 9,
    "oct": 10, "october": 10,
    "nov": 11, "november": 11,
    "dec": 12, "december": 12,
}

US_STATE_RE = r"(AL|AK|AZ|AR|CA|CO|CT|DE|FL|GA|HI|IA|ID|IL|IN|KS|KY|LA|MA|MD|ME|MI|MN|MO|MS|MT|NC|ND|NE|NH|NJ|NM|NV|NY|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VA|VT|WA|WI|WV|WY|DC)"

ITEM_CODE_RE = re.compile(r"^[A-Z0-9]+[A-Z0-9\-\._/]*$", re.IGNORECASE)

def clean_money(x: str) -> float:
    x = (x or "").strip()
    x = x.replace(",", "")
    try:
        return float(x)
    except:
        return 0.0

def parse_date_to_mmddyyyy(text: str) -> str:
    """
    Handles: "Nov 21, 2025" -> 11/21/2025
    """
    t = (text or "").strip()
    if not t:
        return ""

    # try direct parse
    try:
        dt = pd.to_datetime(t, errors="coerce", infer_datetime_format=True)
        if pd.notna(dt):
            return dt.strftime("%m/%d/%Y")
    except:
        pass

    # manual parse: "Nov 21, 2025"
    m = re.search(r"([A-Za-z]+)\s+(\d{1,2}),\s*(\d{4})", t)
    if m:
        mon = MONTHS.get(m.group(1).lower(), None)
        day = int(m.group(2))
        year = int(m.group(3))
        if mon:
            return f"{mon:02d}/{day:02d}/{year}"
    return t

def lines_from_pdf(pdf_bytes: bytes):
    """
    Extract readable lines from the first page (and optionally subsequent pages).
    Alcorn quote format is consistent.
    """
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        all_lines = []
        for page in pdf.pages:
            txt = page.extract_text() or ""
            if txt:
                all_lines.extend(txt.splitlines())
        # fallback using words if needed
        if not all_lines:
            page = pdf.pages[0]
            words = page.extract_words()
            words = sorted(words, key=lambda w: (w["top"], w["x0"]))
            # group by y
            groups = []
            y_tol = 2
            for w in words:
                placed = False
                for g in groups:
                    if abs(g["top"] - w["top"]) <= y_tol:
                        g["words"].append(w)
                        placed = True
                        break
                if not placed:
                    groups.append({"top": w["top"], "words": [w]})
            for g in groups:
                g["words"] = sorted(g["words"], key=lambda w: w["x0"])
                all_lines.append(" ".join(w["text"] for w in g["words"]))
        return [l.strip() for l in all_lines if l and l.strip()]

def extract_header_fields(lines):
    """
    Extract:
    - QuoteNumber
    - QuoteDate
    - ReferralManagerCode (Salesperson code like JZ/MR)
    - Customer Number/ID
    - Company, Address, City, State, Zip, Country
    """
    quote_number = ""
    quote_date = ""
    salesperson = ""
    customer_no = ""
    company = ""
    address = ""
    city = ""
    state = ""
    zipcode = ""
    country = ""

    # Quote date appears near top: "Nov 21, 2025"
    for l in lines[:10]:
        if re.search(r"[A-Za-z]{3,9}\s+\d{1,2},\s*\d{4}", l):
            quote_date = parse_date_to_mmddyyyy(l)
            break

    # Quote number: line often includes "... QT000171" or "... QT569MR25"
    for l in lines[:40]:
        m = re.search(r"\b(QT[0-9A-Z]+)\b", l, re.IGNORECASE)
        if m:
            quote_number = m.group(1).upper()
            break

    # Row with "Customer No." and "Salesperson"
    for l in lines:
        if "Customer No." in l and "Salesperson" in l:
            # next line usually has values
            idx = lines.index(l)
            if idx + 1 < len(lines):
                nxt = lines[idx + 1]
                # example: "11007-4 JZ Nov 21, 2025 UPSPPA NET30"
                # or: "Brock Beehler 2026-1 MR Nov 21, 2025 BRAUN NET30"
                tokens = nxt.split()
                # Find CustomerNo as token containing digit or dash
                # Find Salesperson as 2-3 letters near it
                # We'll detect salesperson as token of letters length 2-3
                # and customer_no as token containing digits and dash.
                for t in tokens:
                    if re.search(r"\d", t) and "-" in t and not customer_no:
                        customer_no = t.strip()
                for t in tokens:
                    if re.fullmatch(r"[A-Za-z]{2,3}", t) and not salesperson:
                        salesperson = t.upper()
                break

    # Company/address block (Ship To)
    # In your sample:
    # line has "Sold To: Ship To:" then next lines contain company/address.
    for i, l in enumerate(lines):
        if l.strip() == "Sold To: Ship To:":
            # next lines: company, maybe email, street, city/state/zip, country
            # We'll take Ship To side by using rightmost portion if duplicated.
            cand = lines[i+1:i+7]
            # Company often in first line after header
            if cand:
                # handle duplicates like: "Kenworth ... Kenworth ..."
                first = cand[0]
                # keep only left half before duplicated repeat (simple heuristic)
                company = first.split("  ")[0].strip()
                # normalize company example "Kenworth- Paccar Company Sainte Therese" => "Kenworth"
                company = company.split("-")[0].strip()
                if company.lower().endswith("corp"):
                    company = company.replace("Corporation", "Corp").strip()

            # Address line usually contains street
            # Find a line that has a number + street word
            street_line = ""
            city_line = ""
            for c in cand:
                if re.search(r"\d", c) and any(w in c.lower() for w in ["st", "street", "rue", "ave", "road", "rd", "blvd", "dr", "ln", "hwy"]):
                    street_line = c.strip()
                if re.search(rf"\b{US_STATE_RE}\b", c) or re.search(r"\b[A-Z]\d[A-Z]\d[A-Z]\d\b", c):  # Canada postal like J7E4K9
                    city_line = c.strip()

            if street_line:
                address = street_line

            if city_line:
                # Canada example: "Sainte Therese, QC, J7E4K9" or "Sainte Therese, QC J7E4K9"
                m_ca = re.search(r"^(.*?),\s*([A-Z]{2})[,\s]+([A-Z]\d[A-Z]\d[A-Z]\d)$", city_line)
                if m_ca:
                    city = m_ca.group(1).strip()
                    state = m_ca.group(2).strip()
                    zipcode = m_ca.group(3).strip()
                else:
                    # US example: "Winamac, IN, 46996"
                    m_us = re.search(rf"^(.*?),\s*({US_STATE_RE})[,\s]+(\d{{5}}(?:-\d{{4}})?)$", city_line)
                    if m_us:
                        city = m_us.group(1).strip()
                        state = m_us.group(2).strip()
                        zipcode = m_us.group(3).strip()

            # Country line is typically "Canada" or "USA"
            for c in cand:
                if c.strip().lower() in ("canada", "usa", "united states", "united states of america"):
                    country = "Canada" if "canada" in c.strip().lower() else "USA"
                    break

            break

    return {
        "QuoteNumber": quote_number,
        "QuoteDate": quote_date,
        "ReferralManagerCode": salesperson,
        "Customer Number/ID": customer_no,
        "Company": company,
        "Address": address,
        "City": city,
        "State": state,
        "ZipCode": zipcode,
        "Country": country,
    }

def extract_line_items(lines):
    """
    Parse the line-item table:
    Detect start at "Qty." and read rows beginning with a number.
    Handles wrapped description lines.
    """
    items = []
    start_idx = None

    for i, l in enumerate(lines):
        if l.startswith("Qty.") and "Item Number" in l and "Description" in l:
            start_idx = i
            break

    if start_idx is None:
        return items

    # table data starts after header rows (next 1-2 lines)
    data_lines = lines[start_idx+1:]

    # stop when summary section begins
    stop_words = ("Comments:", "Subtotal", "Tax Summary:", "Total order", "Total sales tax")
    cleaned = []
    for l in data_lines:
        if any(sw in l for sw in stop_words):
            break
        cleaned.append(l)

    # Build records by detecting lines that start with qty integer
    current = None
    for l in cleaned:
        if re.match(r"^\d+\s+", l):
            if current:
                items.append(current)
            current = {"raw": l}
        else:
            # continuation line
            if current:
                current["raw"] += " " + l.strip()

    if current:
        items.append(current)

    parsed = []
    for rec in items:
        row = rec["raw"]

        # Extract last two money values (unit price, extended)
        money = re.findall(r"(\d{1,3}(?:,\d{3})*\.\d{2})", row)
        unit_price = money[-2] if len(money) >= 2 else ""
        ext_price = money[-1] if len(money) >= 1 else ""

        # Remove those prices from the row for easier token parsing
        row_wo_prices = row
        if ext_price:
            row_wo_prices = row_wo_prices.rsplit(ext_price, 1)[0].strip()
        if unit_price:
            row_wo_prices = row_wo_prices.rsplit(unit_price, 1)[0].strip()

        # Now parse: qty + maybe customer item + item code + description
        # Example:
        # "2 PARTS & MISC ALCJA-13ST AlcornTCBoltTool 16mm 315rpm stdtrig"
        # "1 MISC 273711-0B07 ATE400,60w,Configured-0B07 ATE4001860... WYTB"
        tokens = row_wo_prices.split()
        if not tokens:
            continue

        qty = tokens[0]
        rest = tokens[1:]

        # Find first token that looks like an item code (has digits/letters, often contains '-' or is long)
        item_code_idx = None
        for idx, t in enumerate(rest):
            # strong signals: contains '-' OR digit+letter combination OR long alnum
            if ("-" in t) or (re.search(r"\d", t) and re.search(r"[A-Za-z]", t)) or (len(t) >= 6 and ITEM_CODE_RE.match(t)):
                item_code_idx = idx
                break

        if item_code_idx is None:
            # fallback: just keep everything
            item_id = ""
            desc = " ".join(rest)
        else:
            item_id = rest[item_code_idx].strip()
            desc = " ".join(rest[item_code_idx+1:]).strip()

        parsed.append({
            "item_id": item_id,
            "item_desc": desc,
            "Quantity": qty,
            "Unit Price": clean_money(unit_price),
            "TotalSales": clean_money(ext_price),
        })

    return parsed

def load_template_columns(template_file) -> list:
    df = pd.read_excel(template_file, dtype=str, keep_default_na=False)
    return list(df.columns)

def make_output_rows(template_cols, header, line_items, pdf_filename):
    rows = []
    for it in line_items:
        r = {c: "" for c in template_cols}

        # constant brand
        r["Brand"] = "Alcorn Industrial Inc"

        # header fields
        r["ReferralManagerCode"] = header.get("ReferralManagerCode", "")
        r["QuoteNumber"] = header.get("QuoteNumber", "")
        r["QuoteDate"] = header.get("QuoteDate", "")
        r["Customer Number/ID"] = header.get("Customer Number/ID", "")
        r["Company"] = header.get("Company", "")
        r["Address"] = header.get("Address", "")
        r["City"] = header.get("City", "")
        r["State"] = header.get("State", "")
        r["ZipCode"] = header.get("ZipCode", "")
        r["Country"] = header.get("Country", "")

        # line item fields
        r["item_id"] = it.get("item_id", "")
        r["item_desc"] = it.get("item_desc", "")
        r["Quantity"] = it.get("Quantity", "")
        r["Unit Price"] = it.get("Unit Price", "")
        r["TotalSales"] = it.get("TotalSales", "")

        # pdf column (renamed output)
        r["PDF"] = pdf_filename

        rows.append(r)

    return rows

# =========================
# UI
# =========================
st.title("Alcorn PDF → Excel Extractor")

template_file = st.file_uploader("Step 1 — Upload Template Excel (Alcron.xlsx)", type=["xlsx"], key="tmpl")
pdf_files = st.file_uploader("Step 2 — Upload Alcorn Quote PDFs (up to 100)", type=["pdf"], accept_multiple_files=True, key="pdfs")

run = st.button("Extract & Build Excel")

if run:
    if not template_file or not pdf_files:
        st.error("Please upload the template Excel and at least 1 PDF.")
        st.stop()

    try:
        template_cols = load_template_columns(template_file)

        all_rows = []
        renamed_pdf_bytes = []  # (newname, bytes)

        for f in pdf_files:
            pdf_bytes = f.read()
            lines = lines_from_pdf(pdf_bytes)

            header = extract_header_fields(lines)
            quote = header.get("QuoteNumber", "").strip()
            if not quote:
                quote = os.path.splitext(f.name)[0].strip()

            # rename PDF to match QuoteNumber
            new_pdf_name = f"{quote}.pdf"
            renamed_pdf_bytes.append((new_pdf_name, pdf_bytes))

            line_items = extract_line_items(lines)
            all_rows.extend(make_output_rows(template_cols, header, line_items, new_pdf_name))

        out_df = pd.DataFrame(all_rows, columns=template_cols)

        # Enforce QuoteDate MM/DD/YYYY if present
        if "QuoteDate" in out_df.columns:
            out_df["QuoteDate"] = out_df["QuoteDate"].astype(str).apply(parse_date_to_mmddyyyy)

        # Small preview only (updated data)
        st.success(f"Processed PDFs: {len(pdf_files)} | Output rows (line items): {len(out_df)}")
        st.dataframe(out_df.head(25), use_container_width=True, height=260)

        # Download Excel
        excel_buf = io.BytesIO()
        with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
            out_df.to_excel(writer, index=False, sheet_name="Extracted")
        excel_buf.seek(0)

        st.download_button(
            "Download Extracted Excel",
            data=excel_buf,
            file_name="alcorn_extracted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # Download renamed PDFs ZIP
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
            for name, b in renamed_pdf_bytes:
                z.writestr(name, b)
        zip_buf.seek(0)

        st.download_button(
            "Download Renamed PDFs (ZIP)",
            data=zip_buf,
            file_name="alcorn_renamed_pdfs.zip",
            mime="application/zip",
        )

    except Exception as e:
        if DEBUG:
            st.exception(e)
        else:
            st.error("Processing failed. Please ensure the PDFs are Alcorn quote format and the template is correct.")
