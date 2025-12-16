import streamlit as st
import pandas as pd
import pdfplumber
import re
import io
from datetime import datetime

# -------------------------
# STREAMLIT CONFIG
# -------------------------
st.set_page_config(page_title="Alcorn PDF → Excel", layout="wide")
st.title("Alcorn PDF → Excel Extractor")

# -------------------------
# OUTPUT HEADERS (FIXED)
# -------------------------
OUTPUT_COLUMNS = [
    "ReferralManagerCode","ReferralManager","ReferralEmail","Brand",
    "QuoteNumber","QuoteVersion","QuoteDate","QuoteValidDate",
    "Customer Number/ID","Company","Address","County","City","State",
    "ZipCode","Country","FirstName","LastName","ContactEmail",
    "ContactPhone","Webaddress","item_id","item_desc","UOM","Quantity",
    "Unit Price","List Price","TotalSales","Manufacturer_ID",
    "manufacturer_Name","Writer Name","CustomerPONumber","PDF",
    "DemoQuote","Duns","SIC","NAICS","LineOfBusiness","LinkedinProfile",
    "PhoneResearched","PhoneSupplied","ParentName"
]

MONEY_RE = re.compile(r"\d{1,3}(?:,\d{3})*\.\d{2}")

# -------------------------
# HELPERS
# -------------------------
def blank_row():
    return {c: "" for c in OUTPUT_COLUMNS}

def money(val):
    return float(str(val).replace(",", ""))

def format_date(text):
    # input: Nov 21, 2025 -> output: 11/21/2025
    try:
        return datetime.strptime(text.strip(), "%b %d, %Y").strftime("%m/%d/%Y")
    except:
        return ""

def clean_text(x):
    return str(x or "").strip()

def norm_header(s):
    s = (s or "")
    s = re.sub(r"\s+", " ", s.strip().lower())
    s = s.replace("qty.", "qty").replace("ord.", "ord")
    return s

# -------------------------
# HEADER FIELD EXTRACTION (CORRECT HIGHLIGHTS)
# -------------------------
def extract_quote_number(text: str) -> str:
    # highlighted light green: Order Number QTxxxx
    m = re.search(r"Order Number\s+(QT[0-9A-Z]+)", text)
    return m.group(1).strip() if m else ""

def extract_customer_no(text: str) -> str:
    # highlighted blue: Customer No. ####
    m = re.search(r"Customer No\.\s+([0-9\-]+)", text)
    return m.group(1).strip() if m else ""

def extract_salesperson_code(text: str) -> str:
    # highlighted silver: Salesperson CR/JZ/etc.
    m = re.search(r"Salesperson\s+([A-Z0-9]{1,8})\b", text)
    return m.group(1).strip() if m else ""

def extract_order_date(text: str) -> str:
    # highlighted orange: Order Date Nov 21, 2025
    m = re.search(r"Order Date\s+([A-Za-z]+\s+\d{1,2},\s+\d{4})", text)
    if m:
        return format_date(m.group(1))
    # fallback to top Date if needed
    m = re.search(r"\bDate\s+([A-Za-z]+\s+\d{1,2},\s+\d{4})", text)
    return format_date(m.group(1)) if m else ""

# -------------------------
# SHIP TO EXTRACTION (GREEN + RED)
# -------------------------
def extract_ship_to(text: str):
    """
    Company = first line inside Ship To box (GREEN)
    Address = street line (RED)
    City/State/Zip = city line (RED)
    Skip lines like 'ATTN:'
    """
    ship = {"Company":"","Address":"","City":"","State":"","ZipCode":"","Country":"USA"}

    if not text:
        return ship

    # Capture Ship To block until next known section
    m = re.search(
        r"Ship To:\s*(.+?)(?:\n\s*Reference|\n\s*PO\s+Number|\n\s*Customer\s+No\.|\n\s*Salesperson|\n\s*Order Date|\n\s*Qty)",
        text,
        flags=re.IGNORECASE | re.DOTALL
    )
    if not m:
        return ship

    lines = [ln.strip() for ln in m.group(1).splitlines() if ln.strip()]
    # remove ATTN lines
    lines = [ln for ln in lines if not ln.lower().startswith("attn")]

    if not lines:
        return ship

    ship["Company"] = lines[0]  # GREEN
    if len(lines) >= 2:
        ship["Address"] = lines[1]  # RED street line

    if len(lines) >= 3:
        cityline = lines[2]
        us = re.search(r"^(.*?),\s*([A-Z]{2})\s*(\d{5}(?:-\d{4})?)$", cityline)
        if us:
            ship["City"] = us.group(1).strip()
            ship["State"] = us.group(2).strip()
            ship["ZipCode"] = us.group(3).strip()
            ship["Country"] = "USA"
        else:
            # Canada fallback (if needed)
            ca = re.search(r"^(.*?),\s*([A-Z]{2})\s*([A-Z]\d[A-Z]\s?\d[A-Z]\d)$", cityline)
            if ca:
                ship["City"] = ca.group(1).strip()
                ship["State"] = ca.group(2).strip()
                ship["ZipCode"] = ca.group(3).replace(" ", "").strip()
                ship["Country"] = "Canada"
            else:
                # better than wrong split
                ship["City"] = cityline.strip()

    # if another line explicitly says Canada
    for ln in lines:
        if ln.strip().lower() == "canada":
            ship["Country"] = "Canada"

    return ship

# -------------------------
# TABLE ITEM EXTRACTION (BEST + NO MORE "0 ROWS")
# -------------------------
def detect_table_header_indices(row):
    """
    Finds column positions by header names in a table row.
    Returns dict of indices or None if not header.
    """
    cells = [norm_header(c) for c in row]
    joined = " | ".join(cells)

    if "item number" not in joined or "description" not in joined:
        return None

    idx = {}
    for i, c in enumerate(cells):
        if c.startswith("qty") or c.startswith("qty ord"):
            idx["qty"] = i
        if "item number" in c and "customer" not in c:
            idx["item"] = i
        if c.startswith("description"):
            idx["desc"] = i
        if "unit price" in c:
            idx["unit"] = i
        if "extended price" in c:
            idx["ext"] = i

    # Qty/Item/Desc/Unit/Ext must exist
    needed = ["qty","item","desc","unit","ext"]
    if all(k in idx for k in needed):
        return idx
    return None

def extract_items_from_tables(page):
    """
    Extract line items using pdfplumber tables.
    This matches your highlighted yellow (item_id) and brown/pink (item_desc).
    """
    items = []

    # Try stronger table settings first (more reliable on forms)
    table_settings = {
        "vertical_strategy": "lines",
        "horizontal_strategy": "lines",
        "snap_tolerance": 3,
        "join_tolerance": 3,
        "edge_min_length": 20,
        "min_words_vertical": 1,
        "min_words_horizontal": 1,
        "intersection_tolerance": 3,
    }

    tables = []
    try:
        t1 = page.extract_tables(table_settings=table_settings)
        if t1:
            tables.extend(t1)
    except:
        pass

    try:
        t2 = page.extract_tables()
        if t2:
            tables.extend(t2)
    except:
        pass

    if not tables:
        return []

    header_idx = None

    for table in tables:
        for row in table:
            if not row:
                continue
            row = [clean_text(c) for c in row]

            # detect header row once
            if header_idx is None:
                maybe = detect_table_header_indices(row)
                if maybe:
                    header_idx = maybe
                continue

            # after header found: parse rows
            try:
                qty = clean_text(row[header_idx["qty"]])
                item_id = clean_text(row[header_idx["item"]])
                desc = clean_text(row[header_idx["desc"]])
                unit_txt = clean_text(row[header_idx["unit"]]).replace(",", "")
                ext_txt  = clean_text(row[header_idx["ext"]]).replace(",", "")

                # skip blanks
                if not qty or not item_id:
                    continue

                # price must exist
                um = MONEY_RE.search(unit_txt)
                em = MONEY_RE.search(ext_txt)
                if not (um and em):
                    continue

                items.append({
                    "Quantity": qty,
                    "item_id": item_id,     # YELLOW
                    "item_desc": desc,      # BROWN/PINK
                    "Unit Price": money(um.group(0)),
                    "TotalSales": money(em.group(0)),
                    "UOM": ""
                })
            except:
                continue

    return items

# -------------------------
# TEXT FALLBACK (only if tables fail)
# -------------------------
def extract_items_fallback_text(text: str):
    """
    LAST resort parser: reads lines and tries to build item rows.
    Use only if extract_tables returns nothing.
    """
    if not text:
        return []

    lines = [l.strip() for l in text.splitlines() if l.strip()]

    start = None
    for i, l in enumerate(lines):
        if ("Item Number" in l) and ("Description" in l):
            start = i + 1
            break
    if start is None:
        return []

    items = []
    for l in lines[start:]:
        if l.lower().startswith("comments"):
            break
        if not re.match(r"^\d+\s", l):
            continue

        prices = MONEY_RE.findall(l)
        if len(prices) < 2:
            continue

        unit_price = money(prices[-2])
        ext_price = money(prices[-1])

        tmp = l.rsplit(prices[-1], 1)[0].strip()
        tmp = tmp.rsplit(prices[-2], 1)[0].strip()

        parts = tmp.split()
        if len(parts) < 3:
            continue

        qty = parts[0]
        item_id = parts[1]
        desc = " ".join(parts[2:]).strip()

        items.append({
            "Quantity": qty,
            "item_id": item_id,
            "item_desc": desc,
            "Unit Price": unit_price,
            "TotalSales": ext_price,
            "UOM": ""
        })
    return items

# -------------------------
# CORE EXTRACTION LOGIC
# -------------------------
def extract_from_pdf(pdf_bytes, pdf_name):
    rows = []

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        # Many quotes are 1 page, but some can be multi-page. We'll scan all pages for items.
        first_page = pdf.pages[0]
        text0 = first_page.extract_text() or ""

        quote = extract_quote_number(text0)
        quote_date = extract_order_date(text0)
        cust = extract_customer_no(text0)
        salesperson = extract_salesperson_code(text0)
        ship = extract_ship_to(text0)

        all_items = []
        for page in pdf.pages:
            # primary: tables (best)
            items = extract_items_from_tables(page)
            if items:
                all_items.extend(items)

        # fallback: if tables gave nothing
        if not all_items:
            all_text = "\n".join([(p.extract_text() or "") for p in pdf.pages])
            all_items = extract_items_fallback_text(all_text)

        if not all_items:
            return []

        for it in all_items:
            r = blank_row()

            r["Brand"] = "Alcorn Industrial Inc"
            r["QuoteNumber"] = quote
            r["QuoteDate"] = quote_date
            r["Customer Number/ID"] = cust
            r["ReferralManagerCode"] = salesperson

            r["Company"] = ship["Company"]
            r["Address"] = ship["Address"]
            r["City"] = ship["City"]
            r["State"] = ship["State"]
            r["ZipCode"] = ship["ZipCode"]
            r["Country"] = ship["Country"]

            r["item_id"] = it["item_id"]
            r["item_desc"] = it["item_desc"]
            r["UOM"] = it.get("UOM","")
            r["Quantity"] = it["Quantity"]
            r["Unit Price"] = it["Unit Price"]
            r["TotalSales"] = it["TotalSales"]
            r["PDF"] = pdf_name

            rows.append(r)

    return rows

# -------------------------
# UI
# -------------------------
pdfs = st.file_uploader("Upload Alcorn Quote PDFs", type=["pdf"], accept_multiple_files=True)

if st.button("Extract"):
    if not pdfs:
        st.error("Upload at least one PDF")
        st.stop()

    all_rows = []
    for f in pdfs:
        data = extract_from_pdf(f.read(), f.name)
        all_rows.extend(data)

    df = pd.DataFrame(all_rows, columns=OUTPUT_COLUMNS)

    st.success(f"Extracted {len(df)} rows")
    st.dataframe(df, height=300, use_container_width=True)

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Extracted")
    out.seek(0)

    st.download_button(
        "Download Excel",
        data=out,
        file_name="alcorn_extracted.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
