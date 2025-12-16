import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO
from datetime import datetime

# ---------------- CONFIG ----------------
st.set_page_config(page_title="Alcorn Quote Extractor", layout="wide")
st.title("Alcorn PDF → Excel Extractor")

# ---------------- HELPERS ----------------
def clean_date(text):
    try:
        return datetime.strptime(text.strip(), "%b %d, %Y").strftime("%m/%d/%Y")
    except:
        return None

def parse_address(lines):
    """
    Parses:
    10 Rue Sicard
    Sainte Therese, QC J7E4K9
    Canada
    """
    address = lines[0] if len(lines) > 0 else ""
    city, state, zip_code, country = "", "", "", ""

    if len(lines) > 1:
        m = re.search(r"(.*),\s*([A-Z]{2})\s*([\w\d]+)", lines[1])
        if m:
            city = m.group(1).strip()
            state = m.group(2).strip()
            zip_code = m.group(3).strip()

    if len(lines) > 2:
        country = lines[2].strip()

    return address, city, state, zip_code, country

# ---------------- PDF PARSER ----------------
def extract_pdf(pdf_file):
    rows = []

    with pdfplumber.open(pdf_file) as pdf:
        text = "\n".join(page.extract_text() for page in pdf.pages if page.extract_text())

    # HEADER FIELDS
    quote_number = re.search(r"QT[\w\d]+", text)
    quote_number = quote_number.group(0) if quote_number else None

    date_match = re.search(r"Nov\s+\d{1,2},\s+\d{4}", text)
    quote_date = clean_date(date_match.group(0)) if date_match else None

    cust_match = re.search(r"Customer No\.\s*([\w\-]+)", text)
    customer_id = cust_match.group(1) if cust_match else None

    sales_match = re.search(r"Customer No\.\s*[\w\-]+\s+([A-Z]{1,3})", text)
    salesperson = sales_match.group(1) if sales_match else None

    # SHIP TO BLOCK
    ship_block = re.search(
        r"Ship To:\n(.*?)\nCustomer",
        text,
        re.S
    )

    company = address = city = state = zip_code = country = ""

    if ship_block:
        ship_lines = [l.strip() for l in ship_block.group(1).split("\n") if l.strip()]
        company = ship_lines[0]
        address, city, state, zip_code, country = parse_address(ship_lines[1:])

    # LINE ITEMS
    item_pattern = re.compile(
        r"(\d+)\s+([A-Z0-9& ]+)\s+([A-Z0-9\-]+)\s+(.*?)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})"
    )

    for m in item_pattern.finditer(text):
        qty = int(m.group(1))
        item_id = m.group(3)
        desc = m.group(4).strip()
        unit_price = float(m.group(5).replace(",", ""))
        total = float(m.group(6).replace(",", ""))

        rows.append({
            "ReferralManagerCode": salesperson,
            "QuoteNumber": quote_number,
            "QuoteDate": quote_date,
            "Customer Number/ID": customer_id,
            "Company": company,
            "Address": address,
            "City": city,
            "State": state,
            "ZipCode": zip_code,
            "Country": country,
            "item_id": item_id,
            "item_desc": desc,
            "Quantity": qty,
            "Unit Price": unit_price,
            "TotalSales": total
        })

    return rows

# ---------------- UI ----------------
uploaded_pdfs = st.file_uploader(
    "Upload Alcorn Quote PDFs",
    type=["pdf"],
    accept_multiple_files=True
)

if uploaded_pdfs:
    all_rows = []
    for pdf in uploaded_pdfs:
        all_rows.extend(extract_pdf(pdf))

    df = pd.DataFrame(all_rows)

    st.success(f"Extracted {len(df)} rows")
    st.dataframe(df.head(10), use_container_width=True)

    # Download
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Extracted")

    st.download_button(
        "Download Excel",
        data=output.getvalue(),
        file_name="alcorn_extracted.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
