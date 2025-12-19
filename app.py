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

def find_phrase_bbox(words, phrase_tokens):
    """Find bbox for a phrase (best-effort, tolerant to punctuation)."""
    tokens = [p.lower().strip() for p in phrase_tokens]
    ws = sorted(words, key=lambda w: (w["top"], w["x0"]))

    def norm(t):  # normalize punctuation
        return re.sub(r"[^a-z0-9]+", "", t.lower())

    tok_norm = [norm(t) for t in tokens]

    for i in range(len(ws) - len(tokens) + 1):
        chunk = ws[i:i+len(tokens)]
        if max(c["top"] for c in chunk) - min(c["top"] for c in chunk) > 4:
            continue
        chunk_norm = [norm(c["text"]) for c in chunk]
        if chunk_norm == tok_norm:
            x0 = min(c["x0"] for c in chunk)
            x1 = max(c["x1"] for c in chunk)
            top = min(c["top"] for c in chunk)
            bottom = max(c["bottom"] for c in chunk)
            return {"x0": x0, "x1": x1, "top": top, "bottom": bottom}
    return None

def find_word_bbox(words, target_word):
    """Find bbox for a single header word like Description / Extended / Unit."""
    t = re.sub(r"[^a-z0-9]+", "", target_word.lower())
    for w in words:
        wn = re.sub(r"[^a-z0-9]+", "", w["text"].lower())
        if wn == t:
            return {"x0": w["x0"], "x1": w["x1"], "top": w["top"], "bottom": w["bottom"]}
    return None

# -------------------------
# HEADER FIELDS FROM TEXT
# -------------------------
def extract_quote_number(text):
    m = re.search(r"Order Number\s+(QT[0-9A-Z]+)", text)
    return m.group(1).strip() if m else ""

def extract_customer_no(text):
    m = re.search(r"Customer No\.\s+([0-9\-]+)", text)
    return m.group(1).strip() if m else ""

def extract_salesperson_code(text):
    m = re.search(r"Salesperson\s+([A-Z]{1,6})\b", text)
    return m.group(1).strip() if m else ""

def extract_order_date(text):
    # Prefer "Order Date"
    m = re.search(r"Order Date\s+([A-Za-z]+\s+\d{1,2},\s+\d{4})", text)
    if m:
        return format_date(m.group(1))
    # Fallback to top "Date"
    m = re.search(r"\bDate\s+([A-Za-z]+\s+\d{1,2},\s+\d{4})", text)
    return format_date(m.group(1)) if m else ""

# -------------------------
# SHIP TO: PRIMARY (coords) + FALLBACK (text)
# -------------------------
def extract_ship_to_from_page(page):
    ship = {"Company":"","Address":"","City":"","State":"","ZipCode":"","Country":"USA"}
    words = page.extract_words(keep_blank_chars=False, use_text_flow=True)
    if not words:
        return ship

    bbox = find_phrase_bbox(words, ["Ship", "To:"]) or find_phrase_bbox(words, ["Ship", "To"])
    if not bbox:
        return ship

    # right-side box area under label
    x0 = bbox["x0"] - 5
    x1 = page.width
    y0 = bbox["bottom"] + 2
    y1 = y0 + 150

    block_words = [w for w in words if (w["x0"] >= x0 and w["x1"] <= x1 and w["top"] >= y0 and w["bottom"] <= y1)]
    lines = group_words_into_lines(block_words, y_tol=3)

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
        us = re.search(r"^(.*?),\s*([A-Z]{2})\s*(\d{5}(?:-\d{4})?)$", cityline)
        if us:
            ship["City"] = us.group(1).strip()
            ship["State"] = us.group(2).strip()
            ship["ZipCode"] = us.group(3).strip()
            ship["Country"] = "USA"
        else:
            ca = re.search(r"^(.*?),\s*([A-Z]{2})\s*([A-Z]\d[A-Z]\s?\d[A-Z]\d)$", cityline)
            if ca:
                ship["City"] = ca.group(1).strip()
                ship["State"] = ca.group(2).strip()
                ship["ZipCode"] = ca.group(3).replace(" ", "").strip()
                ship["Country"] = "Canada"
            else:
                ship["City"] = cityline.strip()

    for ln in lines:
        if ln.strip().lower() == "canada":
            ship["Country"] = "Canada"
        if ln.strip().lower() in ["usa", "united states", "united states of america"]:
            ship["Country"] = "USA"
    return ship

def extract_ship_to_from_text(text):
    ship = {"Company":"","Address":"","City":"","State":"","ZipCode":"","Country":"USA"}
    if not text:
        return ship

    m = re.search(
        r"Ship To:\s*(.+?)(?:\n\s*Reference|\n\s*PO\s+Number|\n\s*Customer\s+No\.|\n\s*Salesperson|\n\s*Order Date|\n\s*Qty\.)",
        text, flags=re.IGNORECASE | re.DOTALL
    )
    if not m:
        return ship

    lines = [ln.strip() for ln in m.group(1).splitlines() if ln.strip()]
    # skip ATTN lines
    cleaned = [ln for ln in lines if not ln.lower().startswith("attn")]
    if not cleaned:
        return ship

    ship["Company"] = cleaned[0]
    if len(cleaned) >= 2:
        ship["Address"] = cleaned[1]
    if len(cleaned) >= 3:
        cityline = cleaned[2]
        us = re.search(r"^(.*?),\s*([A-Z]{2})\s*(\d{5}(?:-\d{4})?)$", cityline)
        if us:
            ship["City"] = us.group(1).strip()
            ship["State"] = us.group(2).strip()
            ship["ZipCode"] = us.group(3).strip()
            ship["Country"] = "USA"
        else:
            ship["City"] = cityline.strip()
    return ship

# -------------------------
# ITEMS: PRIMARY (coords) + FALLBACK (text line parse)
# -------------------------
def extract_items_from_page(page):
    words = page.extract_words(keep_blank_chars=False, use_text_flow=True)
    if not words:
        return []

    # Find key headers (robust)
    hdr_item = find_phrase_bbox(words, ["Item", "Number"])
    hdr_desc = find_word_bbox(words, "Description") or find_phrase_bbox(words, ["Description"])
    hdr_unit = find_phrase_bbox(words, ["Unit", "Price"]) or find_word_bbox(words, "Unit")
    hdr_ext  = find_phrase_bbox(words, ["Extended", "Price"]) or find_word_bbox(words, "Extended")

    if not (hdr_item and hdr_desc and hdr_unit and hdr_ext):
        return []

    # Column boundaries
    x_item0 = hdr_item["x0"] - 2
    x_item1 = hdr_desc["x0"] - 6

    x_desc0 = hdr_desc["x0"] - 2
    x_desc1 = hdr_unit["x0"] - 6

    x_unit0 = hdr_unit["x0"] - 2
    x_unit1 = hdr_ext["x0"] - 6

    x_ext0  = hdr_ext["x0"] - 2
    x_ext1  = page.width

    y_start = max(hdr_item["bottom"], hdr_desc["bottom"], hdr_unit["bottom"], hdr_ext["bottom"]) + 2
    y_end = page.height

    body_words = [w for w in words if w["top"] >= y_start and w["bottom"] <= y_end]
    if not body_words:
        return []

    # Group into lines by y
    body_words = sorted(body_words, key=lambda w: (w["top"], w["x0"]))
    rows = []
    current = []
    last_top = None
    for w in body_words:
        if last_top is None or abs(w["top"] - last_top) <= 3:
            current.append(w)
            last_top = w["top"] if last_top is None else last_top
        else:
            rows.append(current)
            current = [w]
            last_top = w["top"]
    if current:
        rows.append(current)

    items = []
    for row in rows:
        def cell_text(x0, x1):
            ws = [w for w in row if (w["x0"] >= x0 and w["x1"] <= x1)]
            ws = sorted(ws, key=lambda w: w["x0"])
            return " ".join(w["text"] for w in ws).strip()

        item_txt = cell_text(x_item0, x_item1)
        desc_txt = cell_text(x_desc0, x_desc1)
        unit_txt = cell_text(x_unit0, x_unit1)
        ext_txt  = cell_text(x_ext0,  x_ext1)

        unit_m = MONEY_RE.search(unit_txt or "")
        ext_m = MONEY_RE.search(ext_txt or "")

        if not (unit_m and ext_m):
            continue

        item_id = " ".join((item_txt or "").split()).strip()
        item_desc = " ".join((desc_txt or "").split()).strip()

        # If table line contains no item/desc, skip
        if not item_id or not item_desc:
            continue

        # Quantity often appears in far-left; easiest fallback: read leading digit from entire row line
        row_line = " ".join(w["text"] for w in sorted(row, key=lambda w: w["x0"]))
        qm = re.match(r"\s*(\d+)\b", row_line)
        qty = qm.group(1) if qm else ""

        items.append({
            "Quantity": qty,
            "item_id": item_id,
            "item_desc": item_desc,
            "Unit Price": money(unit_m.group(0)),
            "TotalSales": money(ext_m.group(0)),
            "UOM": ""
        })

    return items

def extract_items_from_text(text):
    if not text:
        return []

    lines = [l.strip() for l in text.splitlines() if l.strip()]

    # Find start of items table
    start = None
    for i, l in enumerate(lines):
        if "Item Number" in l and ("Qty." in l or "Qty" in l):
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

        # Item Number might be multi-word (PARTS & MISC)
        if len(parts) >= 4 and parts[1] == "PARTS" and parts[2] == "&" and parts[3] == "MISC":
            item_id = "PARTS & MISC"
            rest = parts[4:]
        else:
            item_id = parts[1]
            rest = parts[2:]

        item_desc = " ".join(rest).strip()
        if not item_id or not item_desc:
            continue

        items.append({
            "Quantity": qty,
            "item_id": item_id,
            "item_desc": item_desc,
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
        page = pdf.pages[0]
        text = page.extract_text() or ""

        # Required fields (from highlighted areas)
        quote = extract_quote_number(text)              # Order Number (light green)
        quote_date = extract_order_date(text)           # Order Date (orange) -> MM/DD/YYYY
        cust = extract_customer_no(text)                # Customer No. (blue)
        salesperson = extract_salesperson_code(text)    # Salesperson (silver)

        # Ship To: try coords first, fallback to text
        ship = extract_ship_to_from_page(page)
        if not ship["Company"]:
            ship = extract_ship_to_from_text(text)

        # Items: try coords first, fallback to text
        items = extract_items_from_page(page)
        if not items:
            items = extract_items_from_text(text)

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
