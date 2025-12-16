import io
import re
import zipfile
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st
import pdfplumber

# =========================
# AUTH (your requested creds)
# =========================
USERNAME = "matt"
PASSWORD = "Interlynx123"

# =========================
# OUTPUT HEADER (EXACT)
# =========================
OUTPUT_COLUMNS = [
    "ReferralManagerCode", "ReferralManager", "ReferralEmail", "Brand", "QuoteNumber",
    "QuoteVersion", "QuoteDate", "QuoteValidDate", "Customer Number/ID", "Company",
    "Address", "County", "City", "State", "ZipCode", "Country", "FirstName", "LastName",
    "ContactEmail", "ContactPhone", "Webaddress", "item_id", "item_desc", "UOM", "Quantity",
    "Unit Price", "List Price", "TotalSales", "Manufacturer_ID", "manufacturer_Name",
    "Writer Name", "CustomerPONumber", "PDF", "DemoQuote", "Duns", "SIC", "NAICS",
    "LineOfBusiness", "LinkedinProfile", "PhoneResearched", "PhoneSupplied", "ParentName"
]

US_STATE_RE = r"(AL|AK|AZ|AR|CA|CO|CT|DE|FL|GA|HI|IA|ID|IL|IN|KS|KY|LA|MA|MD|ME|MI|MN|MO|MS|MT|NC|ND|NE|NH|NJ|NM|NV|NY|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VA|VT|WA|WI|WV|WY|DC)"
CAN_POSTAL_RE = r"\b([A-Z]\d[A-Z]\s?\d[A-Z]\d)\b"

MONEY_RE = re.compile(r"\d{1,3}(?:,\d{3})*\.\d{2}")

def make_empty_row() -> Dict[str, str]:
    return {c: "" for c in OUTPUT_COLUMNS}

def normalize_text(x) -> str:
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    return str(x).strip().replace("\u00a0", " ")

def normalize_quote(x) -> str:
    s = normalize_text(x)
    s = s.replace("'", "").replace('"', "")
    return s.upper()

def to_mmddyyyy(val: str) -> str:
    s = normalize_text(val)
    if not s:
        return ""
    dt = pd.to_datetime(s, errors="coerce", infer_datetime_format=True)
    if pd.notna(dt):
        return dt.strftime("%m/%d/%Y")
    # Handle "Nov 21, 2025"
    m = re.search(r"\b([A-Za-z]{3,9})\s+(\d{1,2}),\s*(\d{4})\b", s)
    if m:
        mon = m.group(1).lower()[:3]
        mon_map = {"jan":1,"feb":2,"mar":3,"apr":4,"may":5,"jun":6,"jul":7,"aug":8,"sep":9,"oct":10,"nov":11,"dec":12}
        if mon in mon_map:
            return f"{mon_map[mon]:02d}/{int(m.group(2)):02d}/{int(m.group(3))}"
    return s

def money_to_float(s: str) -> float:
    t = normalize_text(s).replace(",", "")
    try:
        return float(t)
    except:
        return 0.0

def read_entries_excel(uploaded) -> pd.DataFrame:
    uploaded.seek(0)
    df = pd.read_excel(uploaded, dtype=str, keep_default_na=False)
    df.columns = [c.strip() for c in df.columns]

    # Force exact output columns (keep anything extra out)
    for col in OUTPUT_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    df = df[OUTPUT_COLUMNS].copy()

    # Normalize key columns
    df["QuoteNumber"] = df["QuoteNumber"].apply(normalize_quote)
    df["Customer Number/ID"] = df["Customer Number/ID"].astype(str).str.strip()
    df["ReferralManagerCode"] = df["ReferralManagerCode"].astype(str).str.strip()

    # Ensure numeric-ish are still text unless you want otherwise
    df["QuoteDate"] = df["QuoteDate"].apply(to_mmddyyyy)
    df["QuoteValidDate"] = df["QuoteValidDate"].apply(to_mmddyyyy)

    return df

def pdf_full_text(pdf_bytes: bytes) -> str:
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        chunks = []
        for p in pdf.pages:
            t = p.extract_text() or ""
            if t.strip():
                chunks.append(t)
        return "\n".join(chunks)

def extract_ship_to_block(text: str) -> Dict[str, str]:
    """
    Extract Ship To (Company + Address lines).
    Your PDFs show Ship To block repeated; we take first Ship To block.
    """
    out = {"Company":"", "Address":"", "City":"", "State":"", "ZipCode":"", "Country":""}

    m = re.search(r"Ship To:\s*\n(.*?)(?:\nCustomer\s*\nItem Number|\nCustomer\s+Item Number|\nPlease send your order|\nTax Summary:)",
                  text, flags=re.IGNORECASE | re.DOTALL)
    if not m:
        return out

    block = m.group(1).strip()
    lines = [l.strip() for l in block.splitlines() if l.strip()]

    if not lines:
        return out

    # Company = first line (blue)
    out["Company"] = lines[0]

    # Street = first line containing a digit after company (brown)
    street = ""
    cityline = ""
    country = ""
    for l in lines[1:]:
        if not street and re.search(r"\d", l):
            street = l
            continue
        # City line often has "City, ST ..." or "City, QC J7E4K9"
        if not cityline and ("," in l) and (re.search(r"\b[A-Z]{2}\b", l) or re.search(CAN_POSTAL_RE, l)):
            cityline = l
            continue
        if l.lower() in ("canada", "usa", "united states", "united states of america"):
            country = "Canada" if "canada" in l.lower() else "USA"

    out["Address"] = street
    out["Country"] = country

    if cityline:
        # Canada: "Sainte Therese, QC J7E4K9" or "Sainte Therese, QC, J7E4K9"
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

def extract_header_fields(text: str) -> Dict[str, str]:
    """
    Based on your real PDFs:
    - "Order Number QTxxxxx"
    - "Date Nov 21, 2025"
    - "Customer No. 11007-4"
    - "Salesperson JZ"
    """
    out = {
        "QuoteNumber": "",
        "QuoteDate": "",
        "Customer Number/ID": "",
        "ReferralManagerCode": "",
        "Brand": "Alcorn Industrial Inc",
    }

    m = re.search(r"\bOrder Number\s*(QT[0-9A-Z]+)\b", text, flags=re.IGNORECASE)
    if m:
        out["QuoteNumber"] = normalize_quote(m.group(1))
    else:
        m2 = re.search(r"\b(QT[0-9A-Z]+)\b", text, flags=re.IGNORECASE)
        if m2:
            out["QuoteNumber"] = normalize_quote(m2.group(1))

    md = re.search(r"\bDate\s*([A-Za-z]{3,9}\s+\d{1,2},\s*\d{4})\b", text, flags=re.IGNORECASE)
    if md:
        out["QuoteDate"] = to_mmddyyyy(md.group(1))

    mc = re.search(r"\bCustomer No\.\s*([0-9A-Z\-]+)\b", text, flags=re.IGNORECASE)
    if mc:
        out["Customer Number/ID"] = normalize_text(mc.group(1))

    ms = re.search(r"\bSalesperson\s*([A-Z]{2,3})\b", text)
    if ms:
        out["ReferralManagerCode"] = normalize_text(ms.group(1))

    return out

def looks_like_item_code(tok: str) -> bool:
    t = tok.strip()
    if not t:
        return False
    # Alcorn customer item numbers often contain "-" (ALCJA-13ST, 273711-0B07)
    if "-" in t and any(ch.isdigit() for ch in t):
        return True
    # Some can be like IEC4EGV or 24320: alnum with digits, length >= 4
    if len(t) >= 4 and any(ch.isdigit() for ch in t) and any(ch.isalpha() for ch in t):
        return True
    # Pure digits (24320) also can be item codes
    if t.isdigit() and len(t) >= 4:
        return True
    return False

def extract_line_items(text: str) -> List[Dict[str, object]]:
    """
    Extract between "Please send your order to:" and "Tax Summary:"
    Handles wrapped lines: if a line doesn't start with qty, it appends to previous description.
    """
    items: List[Dict[str, object]] = []

    # Isolate item block
    m = re.search(r"Please send your order to:.*?\n(.*?)(?:\nTax Summary:)", text, flags=re.IGNORECASE | re.DOTALL)
    if not m:
        return items

    block = m.group(1).strip()
    lines = [l.strip() for l in block.splitlines() if l.strip()]

    current = None
    for line in lines:
        # New row starts with qty integer
        if re.match(r"^\d+\s+", line):
            if current:
                current["item_desc"] = normalize_text(current["item_desc"])
                items.append(current)
            # Pull money at end: unit + extended are usually last 2 money values
            monies = MONEY_RE.findall(line)
            unit = monies[-2] if len(monies) >= 2 else ""
            ext = monies[-1] if len(monies) >= 1 else ""

            # Remove trailing money strings for parsing tokens
            line_wo_money = line
            if ext:
                line_wo_money = line_wo_money.rsplit(ext, 1)[0].strip()
            if unit:
                line_wo_money = line_wo_money.rsplit(unit, 1)[0].strip()

            tokens = line_wo_money.split()
            qty = tokens[0]
            rest = tokens[1:]

            # Find item_id token (light green mapping)
            idx_code = None
            for i, tok in enumerate(rest):
                if looks_like_item_code(tok):
                    idx_code = i
                    break

            item_id = rest[idx_code] if idx_code is not None else ""
            desc = " ".join(rest[idx_code+1:]).strip() if idx_code is not None else " ".join(rest).strip()

            current = {
                "Quantity": qty,
                "item_id": item_id,
                "item_desc": desc,
                "Unit Price": money_to_float(unit),
                "TotalSales": money_to_float(ext),
            }
        else:
            # wrapped line -> append to description
            if current:
                current["item_desc"] = f"{current['item_desc']} {line}".strip()

    if current:
        current["item_desc"] = normalize_text(current["item_desc"])
        items.append(current)

    # Remove empties
    items = [it for it in items if it.get("item_id") or it.get("item_desc")]
    return items

def build_extracted_df(pdfs: List) -> Tuple[pd.DataFrame, bytes, bytes]:
    """
    Returns:
      - extracted rows as DataFrame (OUTPUT_COLUMNS)
      - excel bytes
      - renamed PDFs zip bytes
    """
    rows = []
    renamed_pdfs: List[Tuple[str, bytes]] = []

    for f in pdfs:
        pdf_bytes = f.read()
        text = pdf_full_text(pdf_bytes)

        header = extract_header_fields(text)
        ship = extract_ship_to_block(text)

        quote = normalize_quote(header.get("QuoteNumber") or f.name.replace(".pdf", ""))
        pdf_name = f"{quote}.pdf"

        renamed_pdfs.append((pdf_name, pdf_bytes))

        items = extract_line_items(text)
        for it in items:
            r = make_empty_row()
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

            r["item_id"] = normalize_text(it.get("item_id", ""))
            r["item_desc"] = normalize_text(it.get("item_desc", ""))
            r["Quantity"] = normalize_text(it.get("Quantity", ""))
            r["Unit Price"] = it.get("Unit Price", "")
            r["TotalSales"] = it.get("TotalSales", "")

            r["PDF"] = pdf_name
            rows.append(r)

    df = pd.DataFrame(rows, columns=OUTPUT_COLUMNS)

    # Excel bytes
    excel_buf = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Extracted")
    excel_buf.seek(0)

    # ZIP bytes
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as z:
        for name, b in renamed_pdfs:
            z.writestr(name, b)
    zip_buf.seek(0)

    return df, excel_buf.getvalue(), zip_buf.getvalue()

def merge_extracted_into_entries(entries: pd.DataFrame, extracted: pd.DataFrame) -> pd.DataFrame:
    """
    Merge strategy:
    - Match QuoteNumber
    - For line items: match QuoteNumber + item_id + Quantity + Unit Price + TotalSales
    - Fill blanks in entries from extracted
    - Append any extracted rows not matched
    """
    ent = entries.copy()
    ext = extracted.copy()

    # Normalize keys
    ent["QuoteNumber"] = ent["QuoteNumber"].apply(normalize_quote)
    ext["QuoteNumber"] = ext["QuoteNumber"].apply(normalize_quote)

    def key_df(df: pd.DataFrame) -> pd.Series:
        return (
            df["QuoteNumber"].astype(str).fillna("") + "||" +
            df["item_id"].astype(str).fillna("") + "||" +
            df["Quantity"].astype(str).fillna("") + "||" +
            df["Unit Price"].astype(str).fillna("") + "||" +
            df["TotalSales"].astype(str).fillna("")
        )

    ent["_k"] = key_df(ent)
    ext["_k"] = key_df(ext)

    ext_map = ext.set_index("_k", drop=False)

    # Fill blanks in entries for matched keys
    fill_cols = [c for c in OUTPUT_COLUMNS if c not in ("QuoteNumber", "item_id", "Quantity", "Unit Price", "TotalSales")]
    for i, row in ent.iterrows():
        k = row["_k"]
        if k in ext_map.index:
            src = ext_map.loc[k]
            # if duplicate keys in extracted (rare), take first
            if isinstance(src, pd.DataFrame):
                src = src.iloc[0]
            for col in fill_cols:
                if normalize_text(ent.at[i, col]) == "" and normalize_text(src.get(col, "")) != "":
                    ent.at[i, col] = src.get(col, "")

    # Append extracted rows not present in entries
    ent_keys = set(ent["_k"].tolist())
    ext_new = ext[~ext["_k"].isin(ent_keys)].copy()

    out = pd.concat([ent.drop(columns=["_k"]), ext_new.drop(columns=["_k"])], ignore_index=True)
    out = out[OUTPUT_COLUMNS].copy()
    return out

# =========================
# UI
# =========================
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.title("Alcorn PDF → Excel Mapper — Login")
    u = st.text_input("Username")
    p = st.text_input("Password", type="password")
    if st.button("Login"):
        if u == USERNAME and p == PASSWORD:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("Invalid username or password.")
    st.stop()

st.title("Alcorn PDF → Excel Mapper")

c1, c2 = st.columns(2)
with c1:
    entries_file = st.file_uploader("Step 1 — Upload Excel Entries (your header)", type=["xlsx"], key="entries")
with c2:
    pdfs = st.file_uploader("Step 2 — Upload Alcorn PDFs (up to 100)", type=["pdf"], accept_multiple_files=True, key="pdfs")

run = st.button("Run Mapping")

if run:
    if not entries_file or not pdfs:
        st.error("Please upload BOTH the Excel entries file and the PDFs.")
        st.stop()

    try:
        entries_df = read_entries_excel(entries_file)
        extracted_df, extracted_excel_bytes, renamed_zip_bytes = build_extracted_df(pdfs)

        final_df = merge_extracted_into_entries(entries_df, extracted_df)

        st.success(f"PDF rows extracted: {len(extracted_df)} | Final rows: {len(final_df)}")
        st.dataframe(final_df.head(25), use_container_width=True, height=260)

        # Download final mapped file
        out_buf = io.BytesIO()
        with pd.ExcelWriter(out_buf, engine="openpyxl") as writer:
            final_df.to_excel(writer, index=False, sheet_name="Mapped")
        out_buf.seek(0)

        st.download_button(
            "Download Mapped Excel",
            data=out_buf.getvalue(),
            file_name="alcorn_mapped.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # Optional downloads
        st.download_button(
            "Download Extracted-only Excel",
            data=extracted_excel_bytes,
            file_name="alcorn_extracted_only.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.download_button(
            "Download Renamed PDFs (ZIP)",
            data=renamed_zip_bytes,
            file_name="alcorn_renamed_pdfs.zip",
            mime="application/zip",
        )

    except Exception as e:
        st.error("Processing failed. Please verify the Excel header matches exactly and PDFs are Alcorn format.")
