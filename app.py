import streamlit as st
import pandas as pd
import pdfplumber
import io
import re
from datetime import datetime

# ---------------- AUTH ----------------
USERNAME = "matt"
PASSWORD = "Interlynx123"

if "auth" not in st.session_state:
    st.session_state.auth = False

if not st.session_state.auth:
    st.title("Alcorn PDF Extractor Login")
    u = st.text_input("Username")
    p = st.text_input("Password", type="password")
    if st.button("Login"):
        if u == USERNAME and p == PASSWORD:
            st.session_state.auth = True
            st.rerun()
        else:
            st.error("Invalid username or password")
    st.stop()

# ---------------- HELPERS ----------------
def mmddyyyy(date_text):
    try:
        return pd.to_datetime(date_text).strftime("%m/%d/%Y")
    except:
        return ""

def extract_lines(pdf):
    lines = []
    for p in pdf.pages:
        txt = p.extract_text()
        if txt:
            lines.extend([l.strip() for l in txt.splitlines() if l.strip()])
    return lines

# ---------------- HEADER EXTRACTION ----------------
def extract_header(lines):
    data = {}

    for i, l in enumerate(lines):
        if l == "Customer No.":
            data["Customer Number/ID"] = lines[i+1].strip()

        if l == "Salesperson":
            data["ReferralManagerCode"] = lines[i+1].strip()

        if re.match(r"[A-Za-z]{3} \d{1,2}, \d{4}", l):
            data["QuoteDate"] = mmddyyyy(l)

        if l.startswith("QT"):
            data["QuoteNumber"] = l.strip()

    # Ship To block
    for i, l in enumerate(lines):
        if l == "Ship To:":
            data["Company"] = lines[i+1]
            data["Address"] = lines[i+2]

            city_line = lines[i+3]
            city, state, zipc = "", "", ""

            if "," in city_line:
                parts = city_line.split(",")
                city = parts[0].strip()
                rest = parts[1].strip().split()
                state = rest[0]
                zipc = rest[1] if len(rest) > 1 else ""

            data["City"] = city
            data["State"] = state
            data["ZipCode"] = zipc
            data["Country"] = lines[i+4]
            break

    return data

# ---------------- LINE ITEMS ----------------
def extract_items(lines):
    items = []
    in_table = False
    current = None

    for l in lines:
        if l.startswith("Qty.") and "Customer Item Number" in l:
            in_table = True
            continue

        if in_table:
            if re.match(r"^\d+", l):
                if current:
                    items.append(current)
                current = {
                    "item_id": "",
                    "item_desc": "",
                    "Quantity": l.split()[0],
                    "Unit Price": "",
                    "TotalSales": ""
                }
            elif current and re.search(r"[A-Z0-9\-]{5,}", l):
                current["item_id"] = l.strip()
            elif current and re.search(r"\d{1,3}(,\d{3})*\.\d{2}", l):
                nums = re.findall(r"\d{1,3}(?:,\d{3})*\.\d{2}", l)
                if len(nums) >= 2:
                    current["Unit Price"] = nums[-2].replace(",", "")
                    current["TotalSales"] = nums[-1].replace(",", "")
            elif current:
                current["item_desc"] += " " + l.strip()

    if current:
        items.append(current)

    return items

# ---------------- UI ----------------
st.title("Alcorn PDF → Excel Extractor")

template = st.file_uploader("Upload Alcorn Template Excel", type=["xlsx"])
pdfs = st.file_uploader("Upload Alcorn Quote PDFs", type=["pdf"], accept_multiple_files=True)

if st.button("Extract"):
    if not template or not pdfs:
        st.error("Upload template and PDFs")
        st.stop()

    cols = pd.read_excel(template).columns.tolist()
    rows = []

    for f in pdfs:
        with pdfplumber.open(io.BytesIO(f.read())) as pdf:
            lines = extract_lines(pdf)
            header = extract_header(lines)
            items = extract_items(lines)

            for it in items:
                row = {c: "" for c in cols}
                row.update(header)
                row.update(it)
                row["Brand"] = "Alcorn Industrial Inc"
                row["PDF"] = f.name
                rows.append(row)

    df = pd.DataFrame(rows, columns=cols)
    st.success(f"Extracted {len(df)} rows")
    st.dataframe(df.head(20), height=300)

    out = io.BytesIO()
    df.to_excel(out, index=False)
    out.seek(0)

    st.download_button(
        "Download Excel",
        out,
        "alcorn_extracted.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
