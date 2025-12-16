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
    try:
        return datetime.strptime(text.strip(), "%b %d, %Y").strftime("%m/%d/%Y")
    except:
        return ""

def group_words_into_lines(words, y_tol=3):
    """Group pdfplumber words into text lines using y coordinate clustering."""
    if not words:
        return []
    words = sorted(words, key=lambda w: (w["top"], w["x0"]))
    lines = []
    current = [words[0]]
    for w in words[1:]:
        if abs(w["top"] - current[-1]["top"]) <= y_tol:
            current.append(w)
        else:
            lines.append(current)
            current = [w]
    lines.append(current)
    out = []
    for ln in lines:
        ln = sorted(ln, key=lambda w: w["x0"])
        out.append(" ".join(w["text"] for w in ln).strip())
    return [x for x in out if x]

def find_word(words, target, case_insensitive=True):
    """Find first word exactly matching target."""
    for w in words:
        t = w["text"]
        if case_insensitive:
            if t.strip().lower() == target.strip().lower():
                return w
        else:
            if t.strip() == target.strip():
                return w
    return None

def find_phrase_bbox(words, phrase_tokens):
    """
    Find bbox for a phrase like ["Ship", "To:"] by matching sequential tokens
    on same line (approx).
    """
    tokens = [p.lower() for p in phrase_tokens]
    ws = sorted(words, key=lambda w: (w["top"], w["x0"]))
    for i in range(len(ws) - len(tokens) + 1):
        chunk = ws[i:i+len(tokens)]
        if all(chunk[j]["text"].strip().lower() == tokens[j] for j in range(len(tokens))):
            # ensure same line
            if max(c["top"] for c in chunk) - min(c["top"] for c in chunk) <= 3:
                x0 = min(c["x0"] for c in chunk)
                x1 = max(c["x1"] for c in chunk)
                top = min(c["top"] for c in chunk)
                bottom = max(c["bottom"] for c in chunk)
                return {"x0": x0, "x1": x1, "top": top, "bottom": bottom}
    return None

# -------------------------
# EXTRACTION: HEADER FIELDS
# -------------------------
def extract_quote_number(text):
    # Order Number QT...
    m = re.search(r"Order Number\s+(QT[0-9A-Z]+)", text)
    return m.group(1).strip() if m else ""

def extract_customer_no(text):
    m = re.search(r"Customer No\.\s+([0-9\-]+)", text)
    return m.group(1).strip() if m else ""

def extract_salesperson_code(text):
    m = re.search(r"Salesperson\s+([A-Z]{1,4})\b", text)
    return m.group(1).strip() if m else ""

def extract_order_date(text):
    # Prefer "Order Date" (orange)
    m = re.search(r"Order Date\s+([A-Za-z]+\s+\d{1,2},\s+\d{4})", text)
    if m:
        return format_date(m.group(1))
    # Fallback to top "Date"
    m = re.search(r"\bDate\s+([A-Za-z]+\s+\d{1,2},\s+\d{4})", text)
    return format_date(m.group(1)) if m else ""

# -------------------------
# EXTRACTION: SHIP TO (GREEN/RED) – robust
# -------------------------
def extract_ship_to_from_page(page):
    """
    Pull Ship To box using coordinates so Sold To doesn't bleed in.
    - Company = first non-ATTN line (GREEN)
    - Address = next non-ATTN line (RED street)
    - City/State/Zip = next line (RED city line)
    """
    ship = {"Company":"","Address":"","City":"","State":"","ZipCode":"","Country":"USA"}

    words = page.extract_words(keep_blank_chars=False, use_text_flow=True)
    if not words:
        return ship

    # Find "Ship To:" label bbox
    bbox = find_phrase_bbox(words, ["Ship", "To:"])
    if not bbox:
        # sometimes it's "Ship" "To" without colon
        bbox = find_phrase_bbox(words, ["Ship", "To"])
    if not bbox:
        return ship

    # Ship-to box is to the right half and below the label
    x0 = bbox["x0"] - 5
    x1 = page.width
    y0 = bbox["bottom"] + 2
    y1 = y0 + 130  # enough to include company+attn+address+cityline

    block_words = [
        w for w in words
        if (w["x0"] >= x0 and w["x1"] <= x1 and w["top"] >= y0 and w["bottom"] <= y1)
    ]

    lines = group_words_into_lines(block_words, y_tol=3)

    # Remove obvious non-address lines and skip ATTN
    cleaned = []
    for ln in lines:
        t = ln.strip()
        if not t:
            continue
        if t.lower().startswith("attn"):
            continue
        cleaned.append(t)

    if not cleaned:
        return ship

    ship["Company"] = cleaned[0]

    if len(cleaned) >= 2:
        ship["Address"] = cleaned[1]

    if len(cleaned) >= 3:
        cityline = cleaned[2]

        # USA: City, ST 12345 or 12345-6789
        us = re.search(r"^(.*?),\s*([A-Z]{2})\s*(\d{5}(?:-\d{4})?)$", cityline)
        if us:
            ship["City"] = us.group(1).strip()
            ship["State"] = us.group(2).strip()
            ship["ZipCode"] = us.group(3).strip()
            ship["Country"] = "USA"
        else:
            # Canada: City, PR A1A 1A1
            ca = re.search(r"^(.*?),\s*([A-Z]{2})\s*([A-Z]\d[A-Z]\s?\d[A-Z]\d)$", cityline)
            if ca:
                ship["City"] = ca.group(1).strip()
                ship["State"] = ca.group(2).strip()
                ship["ZipCode"] = ca.group(3).replace(" ", "").strip()
                ship["Country"] = "Canada"
            else:
                # Last resort: keep raw in City
                ship["City"] = cityline.strip()

    # If any line explicitly says Canada/USA
    for ln in lines:
        if ln.strip().lower() == "canada":
            ship["Country"] = "Canada"
        if ln.strip().lower() in ["usa", "united states", "united states of america"]:
            ship["Country"] = "USA"

    return ship

# -------------------------
# EXTRACTION: LINE ITEMS (yellow/brown) – robust columns
# -------------------------
def extract_items_from_page(page):
    """
    Extract items from the table by column positions:
    Qty | Item Number | (Customer Item Number) | Description | Unit Price | Extended Price
    """
    words = page.extract_words(keep_blank_chars=False, use_text_flow=True)
    if not words:
        return []

    # Find table header words
    hdr_qty = find_phrase_bbox(words, ["Qty."])
    if not hdr_qty:
        hdr_qty = find_phrase_bbox(words, ["Qty."])  # same call, kept for clarity

    hdr_item = find_phrase_bbox(words, ["Item", "Number"])
    hdr_desc = find_phrase_bbox(words, ["Description"])
    hdr_unit = find_phrase_bbox(words, ["Unit", "Price"])
    hdr_ext  = find_phrase_bbox(words, ["Extended", "Price"])

    if not (hdr_item and hdr_desc and hdr_unit and hdr_ext):
        return []

    # Column boundaries (x)
    x_qty0 = 0
    x_qty1 = hdr_item["x0"] - 4

    x_item0 = hdr_item["x0"] - 2
    x_item1 = hdr_desc["x0"] - 6  # includes Customer Item Number column if present

    x_desc0 = hdr_desc["x0"] - 2
    x_desc1 = hdr_unit["x0"] - 6

    x_unit0 = hdr_unit["x0"] - 2
    x_unit1 = hdr_ext["x0"] - 6

    x_ext0  = hdr_ext["x0"] - 2
    x_ext1  = page.width

    # Rows start below header line
    y_start = max(hdr_item["bottom"], hdr_desc["bottom"], hdr_unit["bottom"], hdr_ext["bottom"]) + 2
    y_end = page.height  # stop naturally when no prices found

    body_words = [w for w in words if w["top"] >= y_start and w["bottom"] <= y_end]

    # Group into row lines by y
    row_lines = []
    body_words = sorted(body_words, key=lambda w: (w["top"], w["x0"]))
    current = []
    last_top = None
    for w in body_words:
        if last_top is None or abs(w["top"] - last_top) <= 3:
            current.append(w)
            last_top = w["top"] if last_top is None else last_top
        else:
            row_lines.append(current)
            current = [w]
            last_top = w["top"]
    if current:
        row_lines.append(current)

    items = []
    for row in row_lines:
        # Build cell texts by column range
        def cell_text(x0, x1):
            ws = [w for w in row if w["x0"] >= x0 and w["x1"] <= x1]
            ws = sorted(ws, key=lambda w: w["x0"])
            return " ".join(w["text"] for w in ws).strip()

        qty_txt  = cell_text(x_qty0, x_qty1)
        item_txt = cell_text(x_item0, x_item1)
        desc_txt = cell_text(x_desc0, x_desc1)
        unit_txt = cell_text(x_unit0, x_unit1)
        ext_txt  = cell_text(x_ext0,  x_ext1)

        # stop at footer / blanks
        if not (qty_txt or item_txt or desc_txt or unit_txt or ext_txt):
            continue
        if "comments" in (desc_txt or "").lower():
            break

        # must have prices to be a valid item row
        unit_m = MONEY_RE.search(unit_txt or "")
        ext_m = MONEY_RE.search(ext_txt or "")
        if not (unit_m and ext_m):
            continue

        # Quantity should be numeric
        qty_m = re.search(r"\d+", qty_txt or "")
        if not qty_m:
            continue

        qty = qty_m.group(0)

        # item_id: take full item number cell (yellow)
        # normalize multiple spaces
        item_id = " ".join((item_txt or "").split()).strip()

        # item_desc: take description cell (brown)
        item_desc = " ".join((desc_txt or "").split()).strip()

        if not item_id or not item_desc:
            continue

        items.append({
            "Quantity": qty,
            "item_id": item_id,
            "item_desc": item_desc,
            "Unit Price": money(unit_m.group(0)),
            "TotalSales": money(ext_m.group(0)),
            "UOM": ""
        })

    return items

# -------------------------
# CORE EXTRACTION LOGIC
# -------------------------
def extract_from_pdf(pdf_bytes, pdf_name):
    rows = []

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        page = pdf.pages[0]
        text = page.extract_text() or ""

        quote = extract_quote_number(text)             # light green
        quote_date = extract_order_date(text)          # orange (Order Date preferred)
        cust = extract_customer_no(text)               # blue
        salesperson = extract_salesperson_code(text)   # silver

        ship = extract_ship_to_from_page(page)         # green/red from Ship To box
        items = extract_items_from_page(page)          # yellow/brown from table columns

        if not items:
            return []

        for it in items:
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
