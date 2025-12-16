from __future__ import annotations

import io
import re
import zipfile
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st
from dateutil import parser as dtparser

# Use PyMuPDF for reliable Streamlit Cloud installs
import fitz  # PyMuPDF


# -------------------------
# Cleaning / Normalization
# -------------------------
_WS = re.compile(r"\s+")
_BAD_FILENAME = re.compile(r"[^A-Za-z0-9._-]+")
_QT_RE = re.compile(r"\bQT[0-9A-Z]+\b")

_HEADER_LINE_RE = re.compile(
    r"^(?P<prefix>.+?)\s+(?P<cust>\d{2,}-\d+)\s+(?P<code>[A-Z]{1,3})\s+"
    r"(?P<date>[A-Za-z]{3}\s+\d{1,2},\s+\d{4})\s+.+\s+NET\d+$"
)

_ITEM_LINE_RE = re.compile(
    r"^(?P<qty>\d+)\s+(.+?)\s+(?P<unit>\d{1,3}(?:,\d{3})*\.\d{2})\s+(?P<ext>\d{1,3}(?:,\d{3})*\.\d{2})$"
)

def clean_spaces(s: Optional[str]) -> str:
    if not s:
        return ""
    s = s.replace("\u00a0", " ")
    return _WS.sub(" ", s).strip()

def clean_value(s: Optional[str]) -> str:
    """
    Remove excess spacing + remove special chars.
    Keeps: letters/digits/space and . , - / ( ) # : &
    """
    s = clean_spaces(s)
    s = "".join(ch for ch in s if ch.isprintable())
    s = re.sub(r"[^A-Za-z0-9 \.,\-\/\(\)#:&]", "", s)
    return clean_spaces(s)

def safe_filename(name: str) -> str:
    name = clean_spaces(name)
    name = _BAD_FILENAME.sub("_", name)
    name = name.strip("._-")
    return name or "file"

def parse_money(s: str) -> float:
    return float(s.replace(",", "").strip())

def format_date_mmddyyyy(raw: str) -> str:
    if not raw:
        return ""
    dt = dtparser.parse(raw, fuzzy=True)
    return dt.strftime("%m/%d/%Y")


# -------------------------
# PDF Text Extraction (PyMuPDF)
# -------------------------
def pdf_to_text(pdf_bytes: bytes) -> str:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    pages = []
    for page in doc:
        pages.append(page.get_text("text") or "")
    doc.close()
    return "\n".join(pages)


# -------------------------
# Data Models
# -------------------------
@dataclass
class LineItem:
    qty: int
    item_id: str
    desc: str
    unit_price: float
    ext_price: float

@dataclass
class QuoteData:
    quote_number: str
    quote_date: str
    customer_id: str
    writer_name: str
    company: str
    address: str
    city: str
    state: str
    zipcode: str
    country: str
    items: List[LineItem]


# -------------------------
# Parsing Functions
# -------------------------
def extract_header_fields(lines: List[str], full_text: str) -> Tuple[str, str, str, str]:
    qt = ""
    m = _QT_RE.search(full_text)
    if m:
        qt = m.group(0)

    raw_date = ""
    for l in lines[:30]:
        mm = re.search(r"([A-Za-z]{3}\s+\d{1,2},\s+\d{4})", l)
        if mm:
            raw_date = mm.group(1)
            break
    quote_date = format_date_mmddyyyy(raw_date) if raw_date else ""

    customer_id = ""
    writer_name = ""
    for l in lines[:160]:
        l2 = clean_spaces(l)
        mm = _HEADER_LINE_RE.match(l2)
        if mm:
            customer_id = mm.group("cust")
            prefix = clean_spaces(mm.group("prefix"))
            writer_name = "" if re.match(r"^\d", prefix) else prefix
            break

    return clean_value(qt), quote_date, clean_value(customer_id), clean_value(writer_name)

def extract_ship_to_block(lines: List[str]) -> Tuple[str, str, str, str, str, str]:
    company = address = city = state = zipcode = country = ""

    try:
        idx = next(i for i, l in enumerate(lines) if "Sold To:" in l and "Ship To:" in l)
    except StopIteration:
        return company, address, city, state, zipcode, country

    block = lines[idx + 1 : idx + 10]
    block = [clean_spaces(b) for b in block if clean_spaces(b)]

    def right_half(s: str) -> str:
        if "  " not in s:
            return s
        parts = re.split(r"\s{2,}", s)
        return parts[-1] if len(parts) >= 2 else s

    ship_lines = [right_half(b) for b in block]

    if ship_lines:
        company = ship_lines[0]
    if len(ship_lines) >= 2:
        address = ship_lines[1]

    for i in range(len(ship_lines) - 1, -1, -1):
        v = ship_lines[i].strip()
        if v.lower() in ("canada", "usa", "united states", "united states of america"):
            country = v
            break

    if len(ship_lines) >= 3:
        loc = ship_lines[2].replace(", ", ",")
        parts = [p for p in loc.split(",") if p]
        if len(parts) >= 2:
            city = parts[0]
            rest_join = " ".join(parts[1:]).strip()
            mm = re.match(r"^(?P<st>[A-Z]{2})\s+(?P<zip>[A-Za-z0-9-]+)$", rest_join)
            if mm:
                state = mm.group("st")
                zipcode = mm.group("zip")
            else:
                toks = rest_join.split()
                if len(toks) >= 2:
                    state = toks[0]
                    zipcode = toks[-1]

    return (
        clean_value(company),
        clean_value(address),
        clean_value(city),
        clean_value(state),
        clean_value(zipcode),
        clean_value(country),
    )

def extract_line_items(lines: List[str]) -> List[LineItem]:
    items: List[LineItem] = []

    start_idx = None
    for i, l in enumerate(lines):
        if "Please send your order to:" in l:
            start_idx = i
            break
    if start_idx is None:
        return items

    table_lines = []
    for l in lines[start_idx + 1 :]:
        if "Tax Summary" in l:
            break
        table_lines.append(clean_spaces(l))

    current: Optional[Dict] = None

    for l in table_lines:
        if not l:
            continue

        m = _ITEM_LINE_RE.match(l)
        if m and re.match(r"^\d+\s+", l):
            if current:
                items.append(LineItem(**current))

            qty = int(m.group("qty"))
            unit = parse_money(m.group("unit"))
            ext = parse_money(m.group("ext"))

            core = re.sub(r"^\d+\s+", "", l)
            core = re.sub(
                r"\s+\d{1,3}(?:,\d{3})*\.\d{2}\s+\d{1,3}(?:,\d{3})*\.\d{2}$",
                "",
                core
            ).strip()

            toks = core.split()
            item_id = ""
            desc_tokens = toks[:]
            for j, t in enumerate(toks):
                if any(ch.isdigit() for ch in t) or "-" in t:
                    item_id = t
                    desc_tokens = toks[j + 1 :]
                    break

            desc = " ".join(desc_tokens)

            current = {
                "qty": qty,
                "item_id": clean_value(item_id),
                "desc": clean_value(desc),
                "unit_price": unit,
                "ext_price": ext,
            }
        else:
            if current and l and not re.match(r"^(Qty\.|Ord\.|Item Number|Customer\b|Reference\b)", l):
                current["desc"] = clean_value(current["desc"] + " " + l)

    if current:
        items.append(LineItem(**current))

    return items

def extract_quote(pdf_bytes: bytes) -> QuoteData:
    text = pdf_to_text(pdf_bytes)
    lines = [clean_spaces(l) for l in text.splitlines()]

    quote_number, quote_date, customer_id, writer_name = extract_header_fields(lines, text)
    company, address, city, state, zipcode, country = extract_ship_to_block(lines)
    items = extract_line_items(lines)

    return QuoteData(
        quote_number=quote_number,
        quote_date=quote_date,
        customer_id=customer_id,
        writer_name=writer_name,
        company=company,
        address=address,
        city=city,
        state=state,
        zipcode=zipcode,
        country=country,
        items=items,
    )


# -------------------------
# Mapping to Template Columns
# -------------------------
def build_rows(template_cols: List[str], quote: QuoteData, renamed_pdf_name: str, default_brand: str, country_fallback: str) -> List[Dict]:
    rows: List[Dict] = []

    for it in quote.items:
        r = {c: "" for c in template_cols}

        # Header fields
        if "Brand" in r: r["Brand"] = default_brand
        if "QuoteNumber" in r: r["QuoteNumber"] = quote.quote_number
        if "QuoteDate" in r: r["QuoteDate"] = quote.quote_date
        if "Customer Number/ID" in r: r["Customer Number/ID"] = quote.customer_id
        if "Writer Name" in r: r["Writer Name"] = quote.writer_name

        # Address block
        if "Company" in r: r["Company"] = quote.company
        if "Address" in r: r["Address"] = quote.address
        if "City" in r: r["City"] = quote.city
        if "State" in r: r["State"] = quote.state
        if "ZipCode" in r: r["ZipCode"] = quote.zipcode
        if "Country" in r: r["Country"] = quote.country or country_fallback

        # Line item fields
        if "item_id" in r: r["item_id"] = it.item_id
        if "item_desc" in r: r["item_desc"] = it.desc
        if "Quantity" in r: r["Quantity"] = it.qty
        if "Unit Price" in r: r["Unit Price"] = float(it.unit_price)
        if "TotalSales" in r: r["TotalSales"] = float(it.ext_price)

        # PDF name column (if exists)
        if "PDF" in r: r["PDF"] = renamed_pdf_name

        rows.append(r)

    # If no items, still output one row
    if not rows:
        r = {c: "" for c in template_cols}
        if "Brand" in r: r["Brand"] = default_brand
        if "QuoteNumber" in r: r["QuoteNumber"] = quote.quote_number
        if "QuoteDate" in r: r["QuoteDate"] = quote.quote_date
        if "Customer Number/ID" in r: r["Customer Number/ID"] = quote.customer_id
        if "Company" in r: r["Company"] = quote.company
        if "PDF" in r: r["PDF"] = renamed_pdf_name
        rows.append(r)

    return rows

def strict_cleanup_df(df: pd.DataFrame) -> pd.DataFrame:
    for c in df.columns:
        if df[c].dtype == "object":
            df[c] = df[c].astype(str).map(clean_value)
    return df


# -------------------------
# Streamlit UI
# -------------------------
st.set_page_config(page_title="PDF → Excel Extractor", layout="wide")

st.title("PDF → Excel Extractor (Streamlit Cloud / GitHub)")
st.caption("Upload template Excel + 100+ PDFs → export extracted Excel + renamed PDFs ZIP. No disk writes.")

with st.expander("Privacy & Safety", expanded=False):
    st.markdown(
        """
- No database, no external APIs, no network calls
- PDFs are processed in-memory only
- Output files are generated in-memory only (Excel + ZIP)
- Close/reload app to clear session memory
"""
    )

c1, c2 = st.columns(2)
with c1:
    template_file = st.file_uploader("Upload template Excel (.xlsx)", type=["xlsx"])
with c2:
    pdf_files = st.file_uploader("Upload PDFs (100+ at once)", type=["pdf"], accept_multiple_files=True)

default_brand = st.text_input("Default Brand (if your sheet has Brand column)", value="Alcorn Industrial Inc")
country_fallback = st.text_input("Country fallback if missing", value="")
do_cleanup = st.checkbox("Strict cleanup (remove special chars + extra spaces)", value=True)

run = st.button("Extract", type="primary", disabled=not (template_file and pdf_files))

if run:
    template_df = pd.read_excel(io.BytesIO(template_file.getvalue()))
    template_cols = list(template_df.columns)

    all_rows: List[Dict] = []
    renamed_pdf_blobs: List[Tuple[str, bytes]] = []

    prog = st.progress(0)
    for i, up in enumerate(pdf_files):
        pdf_bytes = up.getvalue()
        quote = extract_quote(pdf_bytes)

        if quote.quote_number:
            new_name = safe_filename(f"Alcorn_{quote.quote_number}.pdf")
        else:
            new_name = safe_filename(up.name)

        renamed_pdf_blobs.append((new_name, pdf_bytes))
        all_rows.extend(build_rows(template_cols, quote, new_name, default_brand, country_fallback))

        prog.progress(int((i + 1) / max(len(pdf_files), 1) * 100))

    out_df = pd.DataFrame(all_rows, columns=template_cols)
    if do_cleanup:
        out_df = strict_cleanup_df(out_df)

    st.success(f"Done. Extracted {len(out_df)} row(s) from {len(pdf_files)} PDF(s).")
    st.dataframe(out_df, use_container_width=True, height=420)

    # Excel export (in-memory)
    excel_buf = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine="xlsxwriter") as writer:
        out_df.to_excel(writer, index=False, sheet_name="Extracted")
    excel_buf.seek(0)

    st.download_button(
        "Download extracted Excel",
        data=excel_buf.getvalue(),
        file_name="extracted_quotes.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # Renamed PDFs ZIP export (in-memory)
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        for fname, blob in renamed_pdf_blobs:
            z.writestr(fname, blob)
    zip_buf.seek(0)

    st.download_button(
        "Download renamed PDFs (ZIP)",
        data=zip_buf.getvalue(),
        file_name="renamed_pdfs.zip",
        mime="application/zip",
    )
