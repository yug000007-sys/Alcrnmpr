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

# -------------------------
# HELPERS
# -------------------------
def blank_row():
    return {c: "" for c in OUTPUT_COLUMNS}

def money(val):
    return float(val.replace(",", ""))

def format_date(text):
    try:
        return datetime.strptime(text, "%b %d, %Y").strftime("%m/%d/%Y")
    except:
        return ""

# -------------------------
# CORE EXTRACTION LOGIC
# -------------------------
def extract_from_pdf(pdf_bytes, pdf_name):
    rows = []

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        page = pdf.pages[0]
        text = page.extract_text()

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
        # SHIP TO BLOCK (GREEN + RED)
        # -------------------------
        ship = {"Company":"","Address":"","City":"","State":"","ZipCode":"","Country":"USA"}

        m = re.search(r"Ship To:\s*(.+?)\n\s*\n", text, re.DOTALL)
        if m:
            lines = [l.strip() for l in m.group(1).splitlines() if l.strip()]
            if len(lines) >= 1:
                ship["Company"] = lines[0]                # GREEN
            if len(lines) >= 2:
                ship["Address"] = lines[1]                # RED
            if len(lines) >= 3:
                cityline = lines[2]
                cm = re.search(r"(.*?),\s*([A-Z]{2})\s*(\d{5})", cityline)
                if cm:
                    ship["City"] = cm.group(1)
                    ship["State"] = cm.group(2)
                    ship["ZipCode"] = cm.group(3)

        # -------------------------
        # LINE ITEMS (YELLOW + BROWN)
        # -------------------------
        lines = [l.strip() for l in text.splitlines() if l.strip()]
        start = None
        for i, l in enumerate(lines):
            if l.startswith("Qty.") and "Item Number" in l:
                start = i + 1
                break

        if start is None:
            return []

        for l in lines[start:]:
            if l.lower().startswith("comments"):
                break

            if not re.match(r"^\d+\s", l):
                continue

            prices = re.findall(r"\d{1,3}(?:,\d{3})*\.\d{2}", l)
            if len(prices) < 2:
                continue

            unit = money(prices[-2])
            total = money(prices[-1])

            core = l.replace(prices[-2], "").replace(prices[-1], "").strip()
            parts = core.split()

            qty = parts[0]
            item_id = parts[1]                   # YELLOW
            item_desc = " ".join(parts[2:])      # BROWN

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
            r["item_id"] = item_id
            r["item_desc"] = item_desc
            r["Quantity"] = qty
            r["Unit Price"] = unit
            r["TotalSales"] = total
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

    # DOWNLOAD
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
