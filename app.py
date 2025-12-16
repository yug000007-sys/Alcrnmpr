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
        return datetime.strptime(text, "%b %d, %Y").strftime("%m/%d/%Y")
    except:
        return ""

def safe_text(s):
    return (s or "").strip()

# -------------------------
# SHIP TO (GREEN + RED) - FIXED
# -------------------------
def extract_ship_to(text: str):
    """
    Company  = Ship To first line (GREEN)
    Address  = Ship To second line (RED street)
    City/State/Zip = Ship To third line (RED city line)
    Country default USA (set Canada if detected)
    """
    ship = {"Company":"","Address":"","City":"","State":"","ZipCode":"","Country":"USA"}

    if not text:
        return ship

    # Capture Ship To block until next known section
    m = re.search(
        r"Ship To:\s*(.+?)(?:\n\s*Reference|\n\s*PO\s+Number|\n\s*Customer\s+No\.|\n\s*Salesperson|\n\s*Order Date)",
        text,
        flags=re.IGNORECASE | re.DOTALL
    )
    if not m:
        # fallback (your old method)
        m = re.search(r"Ship To:\s*(.+?)\n\s*\n", text, re.DOTALL)

    if not m:
        return ship

    lines = [ln.strip() for ln in m.group(1).splitlines() if ln.strip()]
    if not lines:
        return ship

    # GREEN
    ship["Company"] = lines[0]

    # RED street
    if len(lines) >= 2:
        ship["Address"] = lines[1]

    # RED city/state/zip line
    if len(lines) >= 3:
        cityline = lines[2]

        # USA format: City, ST 12345 or City, ST 12345-6789
        us = re.search(r"^(.*?),\s*([A-Z]{2})\s*(\d{5}(?:-\d{4})?)$", cityline)
        if us:
            ship["City"] = us.group(1).strip()
            ship["State"] = us.group(2).strip()
            ship["ZipCode"] = us.group(3).strip()
            ship["Country"] = "USA"
        else:
            # Canada format: City, PR A1A 1A1  (optional space)
            ca = re.search(r"^(.*?),\s*([A-Z]{2})\s*([A-Z]\d[A-Z]\s?\d[A-Z]\d)$", cityline)
            if ca:
                ship["City"] = ca.group(1).strip()
                ship["State"] = ca.group(2).strip()
                ship["ZipCode"] = ca.group(3).replace(" ", "").strip()
                ship["Country"] = "Canada"
            else:
                # If not matched, keep raw in City (better than wrong split)
                ship["City"] = cityline.strip()

    # If a later line explicitly says Canada/USA, respect it
    for ln in lines:
        if ln.lower() == "canada":
            ship["Country"] = "Canada"
        if ln.lower() in ["usa", "united states", "united states of america"]:
            ship["Country"] = "USA"

    return ship

# -------------------------
# LINE ITEMS (YELLOW + BROWN) - FIXED
# -------------------------
def parse_line_item_row(line: str):
    """
    Handles both patterns:
    1) Qty + Item Number + Customer Item Number + Description + Unit + Extended
    2) Qty + Item Number + Description + Unit + Extended

    Also handles multi-word Item Number like: PARTS & MISC
    """
    prices = MONEY_RE.findall(line)
    if len(prices) < 2:
        return None

    unit_price = money(prices[-2])
    ext_price = money(prices[-1])

    # remove last two money values from the end safely
    tmp = line.rsplit(prices[-1], 1)[0].strip()
    tmp = tmp.rsplit(prices[-2], 1)[0].strip()

    parts = tmp.split()
    if len(parts) < 2:
        return None

    qty = parts[0]

    # Item Number (YELLOW) can be "PARTS & MISC"
    item_id = ""
    rest = []

    if len(parts) >= 4 and parts[1] == "PARTS" and parts[2] == "&" and parts[3] == "MISC":
        item_id = "PARTS & MISC"
        rest = parts[4:]
    else:
        item_id = parts[1]
        rest = parts[2:]

    # Sometimes there is a Customer Item Number column (often has dash)
    customer_item = ""
    if rest and ("-" in rest[0] or re.match(r"^[A-Z0-9]{3,}$", rest[0])):
        # We only treat it as customer item if the PDF actually uses that column
        # Many Alcorn PDFs show it, and it's usually a short code.
        # But we must not steal the first word of the description if there is no column.
        # Rule: if next token looks like a code AND there are still words after it, treat as customer item.
        if len(rest) >= 2:
            customer_item = rest[0]
            rest = rest[1:]

    # Description (BROWN)
    item_desc = " ".join(([customer_item] if customer_item else []) + rest).strip()

    return {
        "Quantity": qty,
        "item_id": item_id,
        "item_desc": item_desc,
        "Unit Price": unit_price,
        "TotalSales": ext_price,
        "UOM": ""
    }

def extract_items(text: str):
    if not text:
        return []

    lines = [l.strip() for l in text.splitlines() if l.strip()]

    # Find start of items table
    start = None
    for i, l in enumerate(lines):
        if l.startswith("Qty.") and "Item Number" in l:
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

        parsed = parse_line_item_row(l)
        if parsed:
            items.append(parsed)

    return items

# -------------------------
# CORE EXTRACTION LOGIC
# -------------------------
def extract_from_pdf(pdf_bytes, pdf_name):
    rows = []

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        page = pdf.pages[0]
        text = page.extract_text() or ""

        # -------------------------
        # QUOTE NUMBER
        # -------------------------
        quote = ""
        m = re.search(r"Order Number\s+(QT[0-9A-Z]+)", text)
        if m:
            quote = m.group(1)

        # -------------------------
        # QUOTE DATE
        # -------------------------
        quote_date = ""
        d = re.search(r"Date\s+([A-Za-z]+\s+\d{1,2},\s+\d{4})", text)
        if d:
            quote_date = format_date(d.group(1))

        # -------------------------
        # CUSTOMER NUMBER / SALESPERSON
        # -------------------------
        cust = ""
        salesperson = ""

        m = re.search(r"Customer No\.\s+([0-9\-]+)", text)
        if m:
            cust = m.group(1)

        s = re.search(r"Salesperson\s+([A-Z]+)", text)
        if s:
            salesperson = s.group(1)

        # -------------------------
        # SHIP TO (GREEN + RED) FIXED
        # -------------------------
        ship = extract_ship_to(text)

        # -------------------------
        # LINE ITEMS (YELLOW + BROWN) FIXED
        # -------------------------
        items = extract_items(text)
        if not items:
            return []

        for it in items:
            r = blank_row()
            r["Brand"] = "Alcorn Industrial Inc"
            r["QuoteNumber"] = quote
            r["QuoteDate"] = quote_date
            r["Customer Number/ID"] = cust
            r["ReferralManagerCode"] = salesperson

            # Ship To mapping
            r["Company"] = ship["Company"]
            r["Address"] = ship["Address"]
            r["City"] = ship["City"]
            r["State"] = ship["State"]
            r["ZipCode"] = ship["ZipCode"]
            r["Country"] = ship["Country"]

            # Line mapping
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
