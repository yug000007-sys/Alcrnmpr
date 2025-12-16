import os
import io
import re
import zipfile
from datetime import datetime

import streamlit as st
import pandas as pd
import pdfplumber

# -----------------------------
# PAGE
# -----------------------------
st.set_page_config(page_title="Alcorn PDF Extractor", layout="wide")

# -----------------------------
# AUTH (your creds)
# -----------------------------
USERNAME = "matt"
PASSWORD = "Interlynx123"

if "auth" not in st.session_state:
    st.session_state.auth = False

if not st.session_state.auth:
    st.title("Alcorn PDF Extractor — Login")
    u = st.text_input("Username")
    p = st.text_input("Password", type="password")
    if st.button("Login"):
        if u == USERNAME and p == PASSWORD:
            st.session_state.auth = True
            st.rerun()
        else:
            st.error("Invalid username or password.")
    st.stop()

# -----------------------------
# HELPERS
# -----------------------------
US_STATE_RE = r"(AL|AK|AZ|AR|CA|CO|CT|DE|FL|GA|HI|IA|ID|IL|IN|KS|KY|LA|MA|MD|ME|MI|MN|MO|MS|MT|NC|ND|NE|NH|NJ|NM|NV|NY|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VA|VT|WA|WI|WV|WY|DC)"
CAN_POSTAL_RE = r"\b([A-Z]\d[A-Z]\s?\d[A-Z]\d)\b"

def mmddyyyy_any(s: str) -> str:
    s = (s or "").strip()
    if not s:
        return ""
    dt = pd.to_datetime(s, errors="coerce", infer_datetime_format=True)
    if pd.notna(dt):
        return dt.strftime("%m/%d/%Y")

    # handle "Nov 21, 2025"
    m = re.search(r"\b([A-Za-z]{3,9})\s+(\d{1,2}),\s*(\d{4})\b", s)
    if m:
        mon = m.group(1).lower()[:3]
        mon_map = {"jan":1,"feb":2,"mar":3,"apr":4,"may":5,"jun":6,"jul":7,"aug":8,"sep":9,"oct":10,"nov":11,"dec":12}
        if mon in mon_map:
            return f"{mon_map[mon]:02d}/{int(m.group(2)):02d}/{int(m.group(3))}"
    return s

def clean_money(val: str) -> float:
    if val is None:
        return 0.0
    t = str(val).strip().replace(",", "")
    try:
        return float(t)
    except:
        return 0.0

def read_template_cols(template_file) -> list[str]:
    df = pd.read_excel(template_file, dtype=str, keep_default_na=False)
    return list(df.columns)

def extract_full_text(pdf: pdfplumber.PDF) -> str:
    txts = []
    for page in pdf.pages:
        t = page.extract_text() or ""
        if t.strip():
            txts.append(t)
    return "\n".join(txts)

def extract_ship_to_block(full_text: str) -> dict:
    """
    Extract the Ship To block using regex from the full text.
    Expected block:
    Ship To:
      <Company>
      <Street>
      <City, ST/PR, ZIP>
      <Country>
    """
    out = {"Company":"", "Address":"", "City":"", "State":"", "ZipCode":"", "Country":""}

    # capture from "Ship To:" until "Reference" row (or next big section)
    m = re.search(r"Ship To:\s*(.+?)\n\s*Reference\b", full_text, flags=re.IGNORECASE | re.DOTALL)
    if not m:
        # fallback: until "Customer No."
        m = re.search(r"Ship To:\s*(.+?)\n\s*Customer No\.", full_text, flags=re.IGNORECASE | re.DOTALL)
    if not m:
        return out

    block = m.group(1).strip()
    lines = [l.strip() for l in block.splitlines() if l.strip()]

    if not lines:
        return out

    # Company = first line (blue)
    out["Company"] = lines[0]

    # Address line (brown) = first line containing a digit (street)
    street = ""
    cityline = ""
    country = ""
    for l in lines[1:]:
        if not street and re.search(r"\d", l):
            street = l
            continue
        # city line often has comma + state/province + zip/postal
        if not cityline and ("," in l) and (re.search(r"\b[A-Z]{2}\b", l) or re.search(CAN_POSTAL_RE, l)):
            cityline = l
            continue
        if l.lower() in ("canada", "usa", "united states", "united states of america"):
            country = "Canada" if "canada" in l.lower() else "USA"

    out["Address"] = street
    out["Country"] = country

    # parse city/state/zip
    if cityline:
        # Canada example: "Sainte Therese, QC, J7E4K9"
        mca = re.search(r"^(.*?),\s*([A-Z]{2})\s*,?\s*([A-Z]\d[A-Z]\s?\d[A-Z]\d)\s*$", cityline)
        if mca:
            out["City"] = mca.group(1).strip()
            out["State"] = mca.group(2).strip()
            out["ZipCode"] = mca.group(3).replace(" ", "").strip()
        else:
            mus = re.search(rf"^(.*?),\s*({US_STATE_RE})\s*,?\s*(\d{{5}}(?:-\d{{4}})?)\s*$", cityline)
            if mus:
                out["City"] = mus.group(1).strip()
                out["State"] = mus.group(2).strip()
                out["ZipCode"] = mus.group(3).strip()

    return out

def extract_header_fields(full_text: str) -> dict:
    """
    Extract red/yellow/date/quote# using full-text regex (robust).
    - ReferralManagerCode from Salesperson (RED)
    - Customer Number/ID from Customer No. (YELLOW)
    - QuoteNumber from Order Number QT... (top-right)
    - QuoteDate from Date Nov 21, 2025 (top-right)
    """
    out = {
        "ReferralManagerCode": "",
        "Customer Number/ID": "",
        "QuoteNumber": "",
        "QuoteDate": ""
    }

    # QuoteNumber: Order Number QT000171
    m = re.search(r"\bOrder Number\s*(QT[0-9A-Z]+)\b", full_text, flags=re.IGNORECASE)
    if m:
        out["QuoteNumber"] = m.group(1).upper()
    else:
        m2 = re.search(r"\b(QT[0-9A-Z]+)\b", full_text, flags=re.IGNORECASE)
        if m2:
            out["QuoteNumber"] = m2.group(1).upper()

    # QuoteDate: Date Nov 21, 2025
    md = re.search(r"\bDate\s*([A-Za-z]{3,9}\s+\d{1,2},\s*\d{4})\b", full_text, flags=re.IGNORECASE)
    if md:
        out["QuoteDate"] = mmddyyyy_any(md.group(1))

    # Customer No.
    mc = re.search(r"\bCustomer No\.\s*([0-9A-Z\-]+)\b", full_text, flags=re.IGNORECASE)
    if mc:
        out["Customer Number/ID"] = mc.group(1).strip()

    # Salesperson
    ms = re.search(r"\bSalesperson\s*([A-Z]{2,3})\b", full_text)
    if ms:
        out["ReferralManagerCode"] = ms.group(1).strip()

    return out

def extract_items_by_coordinates(page: pdfplumber.page.Page) -> list[dict]:
    """
    Coordinate-based line item extraction:
    Uses word positions and column x-ranges to build rows.
    Matches your screenshot columns:
      Qty. Ord. | Item Number | Customer Item Number | Description | Unit Price | Extended Price
    item_id MUST come from Customer Item Number column (light green).
    item_desc from Description column, including wrapped lines (pink).
    """
    words = page.extract_words(x_tolerance=2, y_tolerance=2, keep_blank_chars=False)
    if not words:
        return []

    # Find the header row "Qty." to locate table start area
    header_words = [w for w in words if w["text"] in ("Qty.", "Qty")]
    if not header_words:
        return []

    header_top = min(w["top"] for w in header_words)

    # Filter table area below header
    table_words = [w for w in words if w["top"] > header_top + 5]

    # Column x ranges (tuned for Alcorn layout; works across QT PDFs)
    # NOTE: These may vary slightly but are stable for this template.
    COL_QTY = (0, 70)
    COL_ITEMNUM = (70, 170)
    COL_CUSTITEM = (170, 290)
    COL_DESC = (290, 520)
    COL_UNIT = (520, 610)
    COL_EXT = (610, 900)

    def in_col(w, col):
        return col[0] <= w["x0"] < col[1]

    # Group by line (y coordinate buckets)
    table_words = sorted(table_words, key=lambda w: (w["top"], w["x0"]))

    lines = []
    y_tol = 3
    for w in table_words:
        placed = False
        for line in lines:
            if abs(line["top"] - w["top"]) <= y_tol:
                line["words"].append(w)
                placed = True
                break
        if not placed:
            lines.append({"top": w["top"], "words": [w]})

    # Sort words inside each line
    for line in lines:
        line["words"] = sorted(line["words"], key=lambda w: w["x0"])

    items = []
    current = None

    def line_text_in(col, ws):
        part = [w["text"] for w in ws if in_col(w, col)]
        return " ".join(part).strip()

    money_re = re.compile(r"^\d{1,3}(?:,\d{3})*\.\d{2}$")

    for line in lines:
        ws = line["words"]
        qty_txt = line_text_in(COL_QTY, ws)

        # A "new row" starts when the qty column begins with an integer
        if re.match(r"^\d+$", qty_txt):
            # save previous
            if current:
                # final cleanup
                current["item_desc"] = current["item_desc"].strip()
                items.append(current)

            itemnum = line_text_in(COL_ITEMNUM, ws)
            custitem = line_text_in(COL_CUSTITEM, ws)  # <-- item_id must come from here
            desc = line_text_in(COL_DESC, ws)
            unit = line_text_in(COL_UNIT, ws)
            ext = line_text_in(COL_EXT, ws)

            # Extract unit/ext as last money tokens if multiple
            unit_price = ""
            ext_price = ""
            unit_tokens = unit.split()
            ext_tokens = ext.split()

            # Prefer money-looking tokens
            for t in reversed(unit_tokens):
                if money_re.match(t):
                    unit_price = t
                    break
            for t in reversed(ext_tokens):
                if money_re.match(t):
                    ext_price = t
                    break

            current = {
                "Quantity": qty_txt,
                "item_id": custitem,         # LIGHT GREEN
                "item_desc": desc,           # PINK (starts here, may wrap)
                "Unit Price": clean_money(unit_price),
                "TotalSales": clean_money(ext_price),
            }
        else:
            # wrapped description line: append to current description if it has text in description col
            if current:
                wrap_desc = line_text_in(COL_DESC, ws)
                if wrap_desc:
                    current["item_desc"] += " " + wrap_desc

    if current:
        current["item_desc"] = current["item_desc"].strip()
        items.append(current)

    # Filter out junk rows where item_id is empty and desc empty
    items = [it for it in items if (it.get("item_id") or it.get("item_desc"))]
    return items

# -----------------------------
# UI
# -----------------------------
st.title("Alcorn PDF → Excel Extractor")

template_file = st.file_uploader("Step 1 — Upload Alcorn Template Excel", type=["xlsx"])
pdf_files = st.file_uploader("Step 2 — Upload Alcorn Quote PDFs (up to 100)", type=["pdf"], accept_multiple_files=True)

if st.button("Extract"):
    if not template_file or not pdf_files:
        st.error("Please upload the template Excel and at least 1 PDF.")
        st.stop()

    template_cols = read_template_cols(template_file)

    output_rows = []
    renamed_pdfs = []  # (new_name, bytes)

    for f in pdf_files:
        pdf_bytes = f.read()
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            full_text = extract_full_text(pdf)
            header = extract_header_fields(full_text)

            ship_to = extract_ship_to_block(full_text)
            header.update(ship_to)

            quote_num = header.get("QuoteNumber") or os.path.splitext(f.name)[0]
            quote_num = quote_num.strip().replace("'", "").replace('"', "")
            new_pdf_name = f"{quote_num}.pdf"
            renamed_pdfs.append((new_pdf_name, pdf_bytes))

            # Line items from first page (table is on page 1)
            items = extract_items_by_coordinates(pdf.pages[0])

            # If still no items, avoid a blank export row
            if not items:
                continue

            for it in items:
                row = {c: "" for c in template_cols}

                # Fill the template columns if they exist
                def put(col, val):
                    if col in row:
                        row[col] = val

                put("Brand", "Alcorn Industrial Inc")
                put("QuoteNumber", header.get("QuoteNumber", ""))
                put("QuoteDate", header.get("QuoteDate", ""))
                put("Customer Number/ID", header.get("Customer Number/ID", ""))
                put("Company", header.get("Company", ""))
                put("Address", header.get("Address", ""))
                put("City", header.get("City", ""))
                put("State", header.get("State", ""))
                put("ZipCode", header.get("ZipCode", ""))
                put("Country", header.get("Country", ""))

                put("ReferralManagerCode", header.get("ReferralManagerCode", ""))

                # item mapping (green/pink)
                put("item_id", it.get("item_id", ""))
                put("item_desc", it.get("item_desc", ""))

                put("Quantity", it.get("Quantity", ""))
                put("Unit Price", it.get("Unit Price", ""))
                put("TotalSales", it.get("TotalSales", ""))

                put("PDF", new_pdf_name)

                output_rows.append(row)

    df_out = pd.DataFrame(output_rows, columns=template_cols)

    st.success(f"Processed PDFs: {len(pdf_files)} | Extracted rows: {len(df_out)}")

    # Preview updated output only
    st.dataframe(df_out.head(25), use_container_width=True, height=260)

    # Download Excel
    excel_buf = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
        df_out.to_excel(writer, index=False, sheet_name="Extracted")
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
        for name, b in renamed_pdfs:
            z.writestr(name, b)
    zip_buf.seek(0)

    st.download_button(
        "Download Renamed PDFs (ZIP)",
        data=zip_buf,
        file_name="alcorn_renamed_pdfs.zip",
        mime="application/zip",
    )
