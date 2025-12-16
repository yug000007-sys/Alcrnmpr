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
    st.title("Alcorn Extractor — Login")
    u = st.text_input("Username")
    p = st.text_input("Password", type="password")
    if st.button("Login"):
        if u == USERNAME and p == PASSWORD:
            st.session_state.auth = True
            st.rerun()
        else:
            st.error("Invalid username or password.")
    st.stop()


st.title("Alcorn PDF → Excel (Auto Format)")


# -----------------------------
# OUTPUT HEADER (YOUR REQUIRED)
# -----------------------------
OUTPUT_COLUMNS = [
    "ReferralManagerCode","ReferralManager","ReferralEmail","Brand","QuoteNumber","QuoteVersion",
    "QuoteDate","QuoteValidDate","Customer Number/ID","Company","Address","County","City","State",
    "ZipCode","Country","FirstName","LastName","ContactEmail","ContactPhone","Webaddress","item_id",
    "item_desc","UOM","Quantity","Unit Price","List Price","TotalSales","Manufacturer_ID",
    "manufacturer_Name","Writer Name","CustomerPONumber","PDF","DemoQuote","Duns","SIC","NAICS",
    "LineOfBusiness","LinkedinProfile","PhoneResearched","PhoneSupplied","ParentName"
]

MONEY_RE = re.compile(r"\d{1,3}(?:,\d{3})*\.\d{2}")
US_STATE_RE = r"(AL|AK|AZ|AR|CA|CO|CT|DE|FL|GA|HI|IA|ID|IL|IN|KS|KY|LA|MA|MD|ME|MI|MN|MO|MS|MT|NC|ND|NE|NH|NJ|NM|NV|NY|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VA|VT|WA|WI|WV|WY|DC)"
CAN_POSTAL_RE = r"\b([A-Z]\d[A-Z]\s?\d[A-Z]\d)\b"


def make_row():
    return {c: "" for c in OUTPUT_COLUMNS}


def norm(x):
    if x is None:
        return ""
    if isinstance(x, float) and pd.isna(x):
        return ""
    return str(x).replace("\u00a0", " ").strip()


def quote_norm(x):
    return norm(x).replace("'", "").replace('"', "").upper()


def money_float(s):
    s = norm(s).replace(",", "")
    try:
        return float(s)
    except:
        return 0.0


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


def extract_words_band(words, top_min, top_max):
    return [w for w in words if top_min <= w["top"] <= top_max]


def get_value_under_label(words, label_text):
    """
    Find label word(s) and get the value located slightly below it (box value).
    Works for Customer No and Salesperson (as seen in QT000171).
    """
    label_text_low = label_text.lower()
    label_words = [w for w in words if w["text"].lower().startswith(label_text_low)]
    if not label_words:
        return ""

    # take the left-most instance (usually correct)
    label = sorted(label_words, key=lambda x: (x["top"], x["x0"]))[0]
    lx0, lx1, ltop = label["x0"], label["x1"], label["top"]

    # candidates below the label area
    candidates = [
        w for w in words
        if (w["top"] > ltop + 5 and w["top"] < ltop + 25)
        and (w["x0"] >= lx0 - 5 and w["x0"] <= lx1 + 80)
    ]
    # choose left-most
    candidates = sorted(candidates, key=lambda x: x["x0"])
    # return first "value-like"
    for c in candidates:
        t = c["text"].strip()
        if t and t.lower() != "no.":
            return t
    return ""


def extract_header_fields(pdf):
    """
    Extract:
    - QuoteNumber (Order Number QTxxxx)
    - QuoteDate
    - Customer Number/ID
    - ReferralManagerCode (Salesperson code)
    """
    p = pdf.pages[0]
    text = p.extract_text() or ""
    words = p.extract_words(x_tolerance=2, y_tolerance=2)

    # Quote number: "Order Number" then QTxxxxx (often on next line)
    qn = ""
    m = re.search(r"Order Number\s*\n?\s*(QT[0-9A-Z]+)", text, flags=re.IGNORECASE)
    if m:
        qn = quote_norm(m.group(1))
    else:
        m2 = re.search(r"\b(QT[0-9A-Z]{5,})\b", text, flags=re.IGNORECASE)
        if m2:
            qn = quote_norm(m2.group(1))

    # Quote date: line usually like "Nov 21, 2025 1"
    qd = ""
    md = re.search(r"\b([A-Za-z]{3,9}\s+\d{1,2},\s*\d{4})\b", text)
    if md:
        qd = to_mmddyyyy(md.group(1))

    # Customer Number/ID (boxed under label "Customer")
    cust = ""
    # In QT000171, the value "11007-4" is near the "Customer" label.
    # We can locate that by scanning the band around the labels row (~230-250).
    band = extract_words_band(words, 225, 255)
    # Look for the value that matches digit pattern and is near "Customer"
    customer_label = [w for w in band if w["text"].lower() == "customer"]
    if customer_label:
        cl = customer_label[0]
        cand = [
            w for w in band
            if (w["top"] > cl["top"] + 5 and w["top"] < cl["top"] + 25)
            and (w["x0"] > cl["x0"] - 10 and w["x0"] < cl["x0"] + 120)
            and re.search(r"\d", w["text"])
        ]
        cand = sorted(cand, key=lambda x: x["x0"])
        if cand:
            cust = cand[0]["text"].strip()

    # Salesperson code (boxed under label "Salesperson")
    sales = ""
    sp = [w for w in band if w["text"].lower() == "salesperson"]
    if sp:
        sl = sp[0]
        cand = [
            w for w in band
            if (w["top"] > sl["top"] + 5 and w["top"] < sl["top"] + 25)
            and (w["x0"] > sl["x0"] - 10 and w["x0"] < sl["x0"] + 80)
            and len(w["text"].strip()) <= 6
        ]
        cand = sorted(cand, key=lambda x: x["x0"])
        if cand:
            sales = cand[0]["text"].strip()

    return {
        "QuoteNumber": qn,
        "QuoteDate": qd,
        "Customer Number/ID": cust,
        "ReferralManagerCode": sales,
        "Brand": "Alcorn Industrial Inc"
    }


def extract_ship_to(text):
    """
    Ship To block typically:
    Ship To:
    <Company>
    <Street>
    <City, ST Zip>
    <Country>
    """
    out = {"Company":"", "Address":"", "City":"", "State":"", "ZipCode":"", "Country":""}

    m = re.search(
        r"Ship To:\s*(.+?)(?:\n\s*Reference|\n\s*PO\s+Number|\n\s*Customer|\n\s*Salesperson)",
        text, flags=re.IGNORECASE | re.DOTALL
    )
    if not m:
        return out

    block = m.group(1).strip()
    lines = [l.strip() for l in block.splitlines() if l.strip()]
    if not lines:
        return out

    out["Company"] = lines[0]

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


def parse_items_from_text(text):
    """
    Works for Alcorn PDFs like QT000171 where rows look like:
    2 PARTS & MISC ALCJA-13ST AlcornTCBoltTool ... 21,775.00 43,550.00
    1 IEC4EGV Controller w/ Touch Screen 7,507.00 7,507.00
    """
    lines = [l.strip() for l in (text or "").splitlines() if l.strip()]

    start_idx = None
    for i, l in enumerate(lines):
        if "Qty." in l and "Extended" in l:
            start_idx = i + 1

    # fallback: after "Ord." line
    if start_idx is None:
        for i, l in enumerate(lines):
            if l.startswith("Ord.") and "Price" in l:
                start_idx = i + 1
                break

    if start_idx is None:
        return []

    items = []
    for l in lines[start_idx:]:
        if l.lower().startswith("comments"):
            break

        if not re.match(r"^\d+\s", l):
            continue

        monies = MONEY_RE.findall(l)
        if len(monies) < 2:
            continue

        unit = monies[-2]
        ext = monies[-1]

        # remove trailing money tokens
        tmp = l.rsplit(ext, 1)[0].strip()
        tmp = tmp.rsplit(unit, 1)[0].strip()

        tokens = tmp.split()
        if len(tokens) < 2:
            continue

        qty = tokens[0]

        # item_id rule:
        # special case "PARTS & MISC"
        if len(tokens) >= 4 and tokens[1] == "PARTS" and tokens[2] == "&" and tokens[3] == "MISC":
            item_id = "PARTS & MISC"
            rest = tokens[4:]
        else:
            item_id = tokens[1]
            rest = tokens[2:]

        # optional customer item number (like ALCJA-13ST)
        cust_item = ""
        if rest and "-" in rest[0]:
            cust_item = rest[0]
            rest = rest[1:]

        item_desc = " ".join(([cust_item] if cust_item else []) + rest).strip()

        items.append({
            "Quantity": qty,
            "UOM": "",  # not present reliably in this PDF format
            "item_id": item_id,
            "item_desc": item_desc,
            "Unit Price": money_float(unit),
            "TotalSales": money_float(ext),
        })

    return items


def auto_format_excel(xlsx_bytes: bytes) -> bytes:
    wb = load_workbook(io.BytesIO(xlsx_bytes))
    ws = wb.active

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    header_font = Font(bold=True)
    for cell in ws[1]:
        cell.font = header_font
        cell.alignment = Alignment(wrap_text=True, vertical="center")

    TEXT_COLS = {
        "ReferralManagerCode","QuoteNumber","QuoteVersion","Customer Number/ID",
        "ZipCode","item_id","Manufacturer_ID","Duns","SIC","NAICS","CustomerPONumber","PDF"
    }
    DATE_COLS = {"QuoteDate","QuoteValidDate"}
    CURR_COLS = {"Unit Price","List Price","TotalSales"}
    NUM_COLS = {"Quantity"}

    header_map = {str(c.value).strip(): i for i, c in enumerate(ws[1], start=1)}
    max_row = ws.max_row

    for col_name, j in header_map.items():
        col_letter = get_column_letter(j)

        if col_name in TEXT_COLS:
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

        ws.column_dimensions[col_letter].width = min(45, max(12, len(col_name) + 2))

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

if st.button("Extract"):
    if not pdfs:
        st.error("Please upload at least 1 PDF.")
        st.stop()

    rows = []
    renamed = []

    for f in pdfs:
        pdf_bytes = f.read()

        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            p0 = pdf.pages[0]
            text = p0.extract_text() or ""

            header = extract_header_fields(pdf)
            ship = extract_ship_to(text)

            quote = quote_norm(header.get("QuoteNumber") or f.name.replace(".pdf", ""))
            pdf_name = f"{quote}.pdf"
            renamed.append((pdf_name, pdf_bytes))

            items = parse_items_from_text(text)

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

                r["item_id"] = norm(it.get("item_id", ""))
                r["item_desc"] = norm(it.get("item_desc", ""))
                r["UOM"] = norm(it.get("UOM", ""))
                r["Quantity"] = norm(it.get("Quantity", ""))
                r["Unit Price"] = it.get("Unit Price", 0.0)
                r["TotalSales"] = it.get("TotalSales", 0.0)

                r["PDF"] = pdf_name
                rows.append(r)

    df = pd.DataFrame(rows, columns=OUTPUT_COLUMNS)

    st.success(f"Extracted rows: {len(df)}")

    # show small preview only
    st.dataframe(df.head(30), height=260, use_container_width=True)

    # Write Excel then format
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Extracted")
    buf.seek(0)

    formatted = auto_format_excel(buf.getvalue())

    st.download_button(
        "Download Excel",
        data=formatted,
        file_name="alcorn_extracted.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # ZIP renamed PDFs
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
