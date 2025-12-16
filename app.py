import io
import re
import zipfile
from datetime import datetime

import streamlit as st
import pandas as pd
import pdfplumber

# -----------------------------
# SETTINGS
# -----------------------------
st.set_page_config(page_title="Alcorn PDF Extractor", layout="wide")

USERNAME = "matt"
PASSWORD = "Interlynx123"

# -----------------------------
# AUTH
# -----------------------------
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
# REQUIRED OUTPUT HEADER (EXACT)
# -----------------------------
OUTPUT_COLUMNS = [
    "ReferralManagerCode", "ReferralManager", "ReferralEmail", "Brand",
    "QuoteNumber", "QuoteVersion", "QuoteDate", "QuoteValidDate",
    "Customer Number/ID", "Company", "Address", "County", "City", "State",
    "ZipCode", "Country", "FirstName", "LastName", "ContactEmail",
    "ContactPhone", "Webaddress", "item_id", "item_desc", "UOM", "Quantity",
    "Unit Price", "List Price", "TotalSales", "Manufacturer_ID",
    "manufacturer_Name", "Writer Name", "CustomerPONumber", "PDF",
    "DemoQuote", "Duns", "SIC", "NAICS", "LineOfBusiness", "LinkedinProfile",
    "PhoneResearched", "PhoneSupplied", "ParentName"
]

US_STATE_RE = r"(AL|AK|AZ|AR|CA|CO|CT|DE|FL|GA|HI|IA|ID|IL|IN|KS|KY|LA|MA|MD|ME|MI|MN|MO|MS|MT|NC|ND|NE|NH|NJ|NM|NV|NY|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VA|VT|WA|WI|WV|WY|DC)"
CAN_POSTAL_RE = r"\b([A-Z]\d[A-Z]\s?\d[A-Z]\d)\b"

def clean_money(s: str) -> float:
    s = (s or "").strip().replace(",", "")
    try:
        return float(s)
    except:
        return 0.0

def to_mmddyyyy(s: str) -> str:
    s = (s or "").strip()
    if not s:
        return ""
    dt = pd.to_datetime(s, errors="coerce", infer_datetime_format=True)
    if pd.notna(dt):
        return dt.strftime("%m/%d/%Y")

    m = re.search(r"\b([A-Za-z]{3,9})\s+(\d{1,2}),\s*(\d{4})\b", s)
    if m:
        mon = m.group(1).lower()[:3]
        mon_map = {"jan":1,"feb":2,"mar":3,"apr":4,"may":5,"jun":6,"jul":7,"aug":8,"sep":9,"oct":10,"nov":11,"dec":12}
        if mon in mon_map:
            return f"{mon_map[mon]:02d}/{int(m.group(2)):02d}/{int(m.group(3))}"
    return s

def extract_full_text(pdf) -> str:
    parts = []
    for page in pdf.pages:
        txt = page.extract_text() or ""
        if txt.strip():
            parts.append(txt)
    return "\n".join(parts)

def extract_header(full_text: str) -> dict:
    """
    Extract fields based on Alcorn layout:
    - QuoteNumber from "Order Number QTxxxxx"
    - QuoteDate from "Date Nov 21, 2025"
    - Customer No. (yellow)
    - Salesperson (red)
    """
    out = {
        "QuoteNumber": "",
        "QuoteDate": "",
        "Customer Number/ID": "",
        "ReferralManagerCode": "",
    }

    m = re.search(r"\bOrder Number\s*(QT[0-9A-Z]+)\b", full_text, flags=re.IGNORECASE)
    if m:
        out["QuoteNumber"] = m.group(1).upper()
    else:
        m2 = re.search(r"\b(QT[0-9A-Z]+)\b", full_text, flags=re.IGNORECASE)
        if m2:
            out["QuoteNumber"] = m2.group(1).upper()

    md = re.search(r"\bDate\s*([A-Za-z]{3,9}\s+\d{1,2},\s*\d{4})\b", full_text, flags=re.IGNORECASE)
    if md:
        out["QuoteDate"] = to_mmddyyyy(md.group(1))

    mc = re.search(r"\bCustomer No\.\s*([0-9A-Z\-]+)\b", full_text, flags=re.IGNORECASE)
    if mc:
        out["Customer Number/ID"] = mc.group(1).strip()

    ms = re.search(r"\bSalesperson\s*([A-Z]{2,3})\b", full_text)
    if ms:
        out["ReferralManagerCode"] = ms.group(1).strip()

    return out

def extract_ship_to(full_text: str) -> dict:
    """
    Extract ship-to block (blue + brown):
    Ship To:
      <Company>
      <Street>
      <City, ST/PR, ZIP/Postal>
      <Country>
    """
    out = {"Company":"", "Address":"", "City":"", "State":"", "ZipCode":"", "Country":""}

    m = re.search(r"Ship To:\s*(.+?)\n\s*Reference\b", full_text, flags=re.IGNORECASE | re.DOTALL)
    if not m:
        m = re.search(r"Ship To:\s*(.+?)\n\s*Customer No\.", full_text, flags=re.IGNORECASE | re.DOTALL)
    if not m:
        return out

    block = m.group(1).strip()
    lines = [l.strip() for l in block.splitlines() if l.strip()]
    if not lines:
        return out

    out["Company"] = lines[0]

    # street line: first line with a digit after company
    street = ""
    cityline = ""
    country = ""

    for l in lines[1:]:
        if not street and re.search(r"\d", l):
            street = l
            continue
        if not cityline and ("," in l) and (re.search(r"\b[A-Z]{2}\b", l) or re.search(CAN_POSTAL_RE, l)):
            cityline = l
            continue
        if l.lower() in ("canada", "usa", "united states", "united states of america"):
            country = "Canada" if "canada" in l.lower() else "USA"

    out["Address"] = street
    out["Country"] = country

    if cityline:
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

def extract_items(page: pdfplumber.page.Page) -> list[dict]:
    """
    Coordinate-based item extraction for Alcorn table:
      Qty. Ord. | Item Number | Customer Item Number | Description | Unit Price | Extended Price

    We will set:
      item_id = Customer Item Number (your green highlight)
      item_desc = Description (pink, incl wrapped lines)
    """
    words = page.extract_words(x_tolerance=2, y_tolerance=2)
    if not words:
        return []

    # Locate header row by "Qty." text
    header_candidates = [w for w in words if w["text"] in ("Qty.", "Qty")]
    if not header_candidates:
        return []

    header_top = min(w["top"] for w in header_candidates)
    table_words = [w for w in words if w["top"] > header_top + 5]
    table_words = sorted(table_words, key=lambda w: (w["top"], w["x0"]))

    # Column x ranges tuned for Alcorn PDFs
    COL_QTY = (0, 70)
    COL_CUSTITEM = (170, 290)
    COL_DESC = (290, 520)
    COL_UNIT = (520, 610)
    COL_EXT = (610, 900)

    def in_col(w, col):
        return col[0] <= w["x0"] < col[1]

    # Group words into visual lines
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

    for line in lines:
        line["words"] = sorted(line["words"], key=lambda w: w["x0"])

    def line_text(col, ws):
        return " ".join([w["text"] for w in ws if in_col(w, col)]).strip()

    money_re = re.compile(r"^\d{1,3}(?:,\d{3})*\.\d{2}$")

    items = []
    current = None

    for line in lines:
        ws = line["words"]
        qty_txt = line_text(COL_QTY, ws)

        if re.fullmatch(r"\d+", qty_txt):
            # New item row
            if current:
                current["item_desc"] = current["item_desc"].strip()
                items.append(current)

            cust_item = line_text(COL_CUSTITEM, ws)
            desc = line_text(COL_DESC, ws)
            unit_txt = line_text(COL_UNIT, ws)
            ext_txt = line_text(COL_EXT, ws)

            unit_price = ""
            ext_price = ""
            for t in reversed(unit_txt.split()):
                if money_re.match(t):
                    unit_price = t
                    break
            for t in reversed(ext_txt.split()):
                if money_re.match(t):
                    ext_price = t
                    break

            current = {
                "Quantity": qty_txt,
                "item_id": cust_item,
                "item_desc": desc,
                "Unit Price": clean_money(unit_price),
                "TotalSales": clean_money(ext_price),
            }
        else:
            # Wrap line → description continuation
            if current:
                wrap = line_text(COL_DESC, ws)
                if wrap:
                    current["item_desc"] += " " + wrap

    if current:
        current["item_desc"] = current["item_desc"].strip()
        items.append(current)

    # Remove empty rows
    items = [it for it in items if it.get("item_id") or it.get("item_desc")]
    return items

def make_row() -> dict:
    return {c: "" for c in OUTPUT_COLUMNS}

# -----------------------------
# UI
# -----------------------------
st.title("Alcorn PDF → Excel Extractor")

pdfs = st.file_uploader("Upload Alcorn Quote PDFs (up to 100)", type=["pdf"], accept_multiple_files=True)

if st.button("Extract"):
    if not pdfs:
        st.error("Please upload at least 1 PDF.")
        st.stop()

    all_rows = []
    renamed_pdfs = []

    for f in pdfs:
        pdf_bytes = f.read()
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            full_text = extract_full_text(pdf)

            header = extract_header(full_text)
            ship = extract_ship_to(full_text)

            quote = header.get("QuoteNumber", "").strip() or f.name.replace(".pdf", "")
            safe_quote = quote.replace("'", "").replace('"', "")
            renamed_name = f"{safe_quote}.pdf"
            renamed_pdfs.append((renamed_name, pdf_bytes))

            # items typically on page 1
            items = extract_items(pdf.pages[0])

            for it in items:
                r = make_row()

                # Fill columns we can extract now
                r["ReferralManagerCode"] = header.get("ReferralManagerCode", "")
                r["Brand"] = "Alcorn Industrial Inc"
                r["QuoteNumber"] = safe_quote
                r["QuoteDate"] = header.get("QuoteDate", "")
                r["Customer Number/ID"] = header.get("Customer Number/ID", "")

                r["Company"] = ship.get("Company", "")
                r["Address"] = ship.get("Address", "")
                r["City"] = ship.get("City", "")
                r["State"] = ship.get("State", "")
                r["ZipCode"] = ship.get("ZipCode", "")
                r["Country"] = ship.get("Country", "")

                r["item_id"] = it.get("item_id", "")
                r["item_desc"] = it.get("item_desc", "")
                r["Quantity"] = it.get("Quantity", "")
                r["Unit Price"] = it.get("Unit Price", "")
                r["TotalSales"] = it.get("TotalSales", "")

                r["PDF"] = renamed_name

                all_rows.append(r)

    df = pd.DataFrame(all_rows, columns=OUTPUT_COLUMNS)

    st.success(f"Processed PDFs: {len(pdfs)} | Output rows: {len(df)}")
    st.dataframe(df.head(25), use_container_width=True, height=260)

    # Excel download (openpyxl)
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Extracted")
    out.seek(0)

    st.download_button(
        "Download Extracted Excel",
        data=out,
        file_name="alcorn_extracted.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # Renamed PDF ZIP
    zbuf = io.BytesIO()
    with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as z:
        for name, b in renamed_pdfs:
            z.writestr(name, b)
    zbuf.seek(0)

    st.download_button(
        "Download Renamed PDFs (ZIP)",
        data=zbuf,
        file_name="alcorn_renamed_pdfs.zip",
        mime="application/zip",
    )
