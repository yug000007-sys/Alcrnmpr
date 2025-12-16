import io
import re
import zipfile
import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
from dateutil import parser as dtparser

st.set_page_config(page_title="PDF Extractor", layout="wide")

# -----------------------
# Helpers
# -----------------------
def clean(s):
    if not s:
        return ""
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"[^A-Za-z0-9 \.,\-\/()#:]", "", s)
    return s.strip()

def format_date(d):
    try:
        return dtparser.parse(d, fuzzy=True).strftime("%m/%d/%Y")
    except:
        return ""

def extract_text(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    text = ""
    for page in doc:
        text += page.get_text()
    doc.close()
    return text

# -----------------------
# Streamlit UI
# -----------------------
st.title("PDF → Excel Extractor (Streamlit Cloud Safe)")

template = st.file_uploader("Upload template Excel", type=["xlsx"])
pdfs = st.file_uploader("Upload PDFs (100+ allowed)", type=["pdf"], accept_multiple_files=True)

if st.button("Extract") and template and pdfs:
    template_df = pd.read_excel(template)
    columns = template_df.columns.tolist()

    rows = []
    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(zip_buffer, "w") as zipf:
        for pdf in pdfs:
            text = extract_text(pdf.read())

            quote = re.search(r"QT[0-9A-Z]+", text)
            date = re.search(r"[A-Za-z]{3}\s+\d{1,2},\s+\d{4}", text)

            quote = quote.group(0) if quote else ""
            date = format_date(date.group(0)) if date else ""

            row = {c: "" for c in columns}
            if "QuoteNumber" in row:
                row["QuoteNumber"] = quote
            if "QuoteDate" in row:
                row["QuoteDate"] = date

            rows.append(row)

            new_name = f"{quote or pdf.name}.pdf"
            zipf.writestr(new_name, pdf.getvalue())

    df = pd.DataFrame(rows, columns=columns)

    excel_buffer = io.BytesIO()
    df.to_excel(excel_buffer, index=False)
    excel_buffer.seek(0)

    zip_buffer.seek(0)

    st.success("Extraction complete")

    st.download_button(
        "Download Excel",
        excel_buffer,
        "extracted.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.download_button(
        "Download Renamed PDFs (ZIP)",
        zip_buffer,
        "pdfs.zip",
        mime="application/zip"
    )
