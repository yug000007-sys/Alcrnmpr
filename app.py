import io
import re
import zipfile
import streamlit as st
import pandas as pd
import pdfplumber

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

# -----------------------------
# AUTH
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

st.title("Alcorn PDF → Excel (Auto-Formatted)")

# -----------------------------
# EXACT HEADER REQUIRED
# -----------------------------
OUTPUT_COLUMNS = [
    "ReferralManagerCode","ReferralManager","ReferralEmail","Brand","QuoteNumber","QuoteVersion",
    "QuoteDate","QuoteValidDate","Customer Number/ID","Company","Address","County","City","State",
    "ZipCode","Country","FirstName","LastName","ContactEmail","ContactPhone","Webaddress","item_id",
    "item_desc","UOM","Quantity","Unit Price","List Price","TotalSales","Manufacturer_ID",
    "manufacturer_Name","Writer Name","CustomerPONumber","PDF","DemoQuote","Duns","SIC","NAICS",
    "LineOfBusiness","LinkedinProfile","PhoneResearched","PhoneSupplied","ParentName"
]

US_STATE_RE = r"(AL|AK|AZ|AR|CA|CO|CT|DE|FL|GA|HI|IA|ID|IL|IN|KS|KY|LA|MA|MD|ME|MI|MN|MO|MS|MT|NC|ND|NE|NH|NJ|NM|NV|NY|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VA|VT|WA|WI|WV|WY|DC)"
CAN_POSTAL_RE = r"\b([A-Z]\d[A-Z]\s?\d[A-Z]\d)\b"
MONEY_RE = re.compile(r"\d{1,3}(?:,\d{3})*\.\d{2}")

def make_row():
    return {c: "" for c in OUTPUT_COLUMNS}

def norm(s):
    return "" if s is None else str(s).strip().replace("\u00a0"," ")

def quote_norm(s):
    return norm(s).replace("'", "").replace('"', "").upper()

def to_mmddyyyy(val):
    s = norm(val)
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

def money_float(s):
    t = norm(s).replace(",", "")
    try:
        return float(t)
    except:
        return 0.0

def pdf_text(pdf_bytes):
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        parts = []
        for p in pdf.pages:
            t = p.extract_text() or ""
            if t.strip():
                parts.append(t)
        return "\n".join(parts)

def extract_header(text):
    out = {
        "ReferralManagerCode": "",
        "QuoteNumber": "",
        "QuoteDate": "",
        "Customer Number/ID": "",
        "Brand": "Alcorn Industrial Inc",
    }

    m = re.search(r"\bOrder Number\s*(QT[0-9A-Z]+)\b", text, flags=re.IGNORECASE)
    out["QuoteNumber"] = quote_norm(m.group(1)) if m else quote_norm(re.search(r"\b(QT[0-9A-Z]+)\b", text, re.IGNORECASE).group(1)) if re.search(r"\b(QT[0-9A-Z]+)\b", text, re.IGNORECASE) else ""

    md = re.search(r"\bDate\s*([A-Za-z]{3,9}\s+\d{1,2},\s*\d{4})\b", text, flags=re.IGNORECASE)
    out["QuoteDate"] = to_mmddyyyy(md.group(1)) if md else ""

    mc = re.search(r"\bCustomer No\.\s*([0-9A-Z\-]+)\b", text, flags=re.IGNORECASE)
    out["Customer Number/ID"] = norm(mc.group(1)) if mc else ""

    ms = re.search(r"\bSalesperson\s*([A-Z]{2,3})\b", text)
    out["ReferralManagerCode"] = norm(ms.group(1)) if ms else ""

    return out

def extract_ship_to(text):
    out = {"Company":"", "Address":"", "City":"", "State":"", "ZipCode":"", "Country":""}
    m = re.search(r"Ship To:\s*\n(.*?)(?:\nCustomer\s+Item Number|\nCustomer Item Number|\nPlease send your order|\nTax Summary:)",
                  text, flags=re.IGNORECASE | re.DOTALL)
    if not m:
        return out
    block = m.group(1).strip()
    lines = [l.strip() for l in block.splitlines() if l.strip()]
    if not lines:
        return out

    out["Company"] = lines[0]

    street, cityline, country = "", "", ""
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

def looks_like_item(tok):
    t = tok.strip()
    if "-" in t and any(ch.isdigit() for ch in t):
        return True
    if len(t) >= 4 and any(ch.isdigit() for ch in t) and any(ch.isalpha() for ch in t):
        return True
    if t.isdigit() and len(t) >= 4:
        return True
    return False

def extract_items(text):
    m = re.search(r"Please send your order to:.*?\n(.*?)(?:\nTax Summary:)", text, flags=re.IGNORECASE | re.DOTALL)
    if not m:
        return []
    lines = [l.strip() for l in m.group(1).splitlines() if l.strip()]

    items = []
    current = None
    for line in lines:
        if re.match(r"^\d+\s+", line):
            if current:
                current["item_desc"] = current["item_desc"].strip()
                items.append(current)

            monies = MONEY_RE.findall(line)
            unit = monies[-2] if len(monies) >= 2 else ""
            ext = monies[-1] if len(monies) >= 1 else ""

            line_wo = line
            if ext:
                line_wo = line_wo.rsplit(ext, 1)[0].strip()
            if unit:
                line_wo = line_wo.rsplit(unit, 1)[0].strip()

            toks = line_wo.split()
            qty = toks[0]
            rest = toks[1:]

            idx = None
            for i, tok in enumerate(rest):
                if looks_like_item(tok):
                    idx = i
                    break

            item_id = rest[idx] if idx is not None else ""
            desc = " ".join(rest[idx+1:]).strip() if idx is not None else " ".join(rest).strip()

            current = {
                "Quantity": qty,
                "item_id": item_id,
                "item_desc": desc,
                "Unit Price": money_float(unit),
                "TotalSales": money_float(ext),
            }
        else:
            if current:
                current["item_desc"] += " " + line

    if current:
        current["item_desc"] = current["item_desc"].strip()
        items.append(current)

    return [it for it in items if it.get("item_id") or it.get("item_desc")]

def auto_format_excel(xlsx_bytes: bytes) -> bytes:
    """
    Formats:
    - Freeze header row, add filter
    - Force TEXT columns (IDs)
    - Date format MM/DD/YYYY
    - Currency columns
    - Column widths
    """
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment, Font
    from openpyxl.utils import get_column_letter

    wb = load_workbook(io.BytesIO(xlsx_bytes))
    ws = wb.active

    # Freeze & filter
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    header_font = Font(bold=True)
    for cell in ws[1]:
        cell.font = header_font
        cell.alignment = Alignment(wrap_text=True, vertical="center")

    # Column categories
    TEXT_COLS = {
        "ReferralManagerCode","QuoteNumber","QuoteVersion","Customer Number/ID","ZipCode",
        "item_id","Manufacturer_ID","Duns","SIC","NAICS","CustomerPONumber","PDF"
    }
    DATE_COLS = {"QuoteDate","QuoteValidDate"}
    CURR_COLS = {"Unit Price","List Price","TotalSales"}
    NUM_COLS = {"Quantity"}

    # Map header -> col index
    header_map = {}
    for j, c in enumerate(ws[1], start=1):
        header_map[str(c.value).strip()] = j

    # Apply formats row by row
    max_row = ws.max_row
    for col_name, j in header_map.items():
        col_letter = get_column_letter(j)

        if col_name in TEXT_COLS:
            # Force text
            for r in range(2, max_row + 1):
                ws[f"{col_letter}{r}"].number_format = "@"

        if col_name in DATE_COLS:
            for r in range(2, max_row + 1):
                ws[f"{col_letter}{r}"].number_format = "mm/dd/yyyy"

        if col_name in CURR_COLS:
            for r in range(2, max_row + 1):
                ws[f"{col_letter}{r}"].number_format = '"$"#,##0.00'

        if col_name in NUM_COLS:
            for r in range(2, max_row + 1):
                ws[f"{col_letter}{r}"].number_format = "0"

        # Width heuristic
        ws.column_dimensions[col_letter].width = min(45, max(12, len(col_name) + 2))

    # Make description wider
    if "item_desc" in header_map:
        ws.column_dimensions[get_column_letter(header_map["item_desc"])].width = 55
    if "Company" in header_map:
        ws.column_dimensions[get_column_letter(header_map["Company"])].width = 35
    if "Address" in header_map:
        ws.column_dimensions[get_column_letter(header_map["Address"])].width = 35

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()

# ---------------- UI ----------------
pdfs = st.file_uploader("Upload Alcorn Quote PDFs (up to 100)", type=["pdf"], accept_multiple_files=True)

if st.button("Extract + Auto-Format Excel"):
    if not pdfs:
        st.error("Please upload at least 1 PDF.")
        st.stop()

    rows = []
    renamed = []

    for f in pdfs:
        pdf_bytes = f.read()
        text = pdf_text(pdf_bytes)

        header = extract_header(text)
        ship = extract_ship_to(text)
        items = extract_items(text)

        quote = quote_norm(header.get("QuoteNumber") or f.name.replace(".pdf",""))
        pdf_name = f"{quote}.pdf"
        renamed.append((pdf_name, pdf_bytes))

        for it in items:
            r = make_row()
            r["Brand"] = header.get("Brand", "Alcorn Industrial Inc")
            r["QuoteNumber"] = quote
            r["QuoteDate"] = header.get("QuoteDate", "")
            r["Customer Number/ID"] = header.get("Customer Number/ID", "")
            r["ReferralManagerCode"] = header.get("ReferralManagerCode", "")

            r["Company"] = ship.get("Company", "")
            r["Address"] = ship.get("Address", "")
            r["City"] = ship.get("City", "")
            r["State"] = ship.get("State", "")
            r["ZipCode"] = ship.get("ZipCode", "")
            r["Country"] = ship.get("Country", "")

            r["item_id"] = norm(it.get("item_id",""))
            r["item_desc"] = norm(it.get("item_desc",""))
            r["Quantity"] = norm(it.get("Quantity",""))
            r["Unit Price"] = it.get("Unit Price", 0.0)
            r["TotalSales"] = it.get("TotalSales", 0.0)

            r["PDF"] = pdf_name
            rows.append(r)

    df = pd.DataFrame(rows, columns=OUTPUT_COLUMNS)
    st.success(f"Extracted rows: {len(df)}")
    st.dataframe(df.head(25), height=260, use_container_width=True)

    # Write Excel
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Extracted")
    buf.seek(0)

    # Auto-format
    formatted = auto_format_excel(buf.getvalue())

    st.download_button(
        "Download Auto-Formatted Excel",
        data=formatted,
        file_name="alcorn_extracted_autoformatted.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # Renamed PDFs ZIP
    zbuf = io.BytesIO()
    with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as z:
        for name, b in renamed:
            z.writestr(name, b)
    zbuf.seek(0)

    st.download_button(
        "Download Renamed PDFs (ZIP)",
        data=zbuf.getvalue(),
        file_name="alcorn_renamed_pdfs.zip",
        mime="application/zip",
    )
