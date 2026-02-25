import streamlit as st
import pandas as pd
import pdfplumber
import re
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Blinkit PO Consolidation", layout="wide")

st.title("Blinkit PO Consolidation Tool")

st.markdown("Upload matching PDF and Excel/CSV files (same file names).")

# Upload section
pdf_files = st.file_uploader(
    "Upload PDF Files",
    type=["pdf"],
    accept_multiple_files=True
)

excel_files = st.file_uploader(
    "Upload Excel/CSV Files",
    type=["xlsx", "xls", "csv"],
    accept_multiple_files=True
)

def normalize_filename(name):
    return name.split(".")[0].strip().lower()

def extract_pdf_data(file):
    text_data = ""

    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            text_data += page.extract_text() or ""

    # Extract required fields using regex
    po_no = re.search(r"PO\s*NO[:\-]?\s*(\S+)", text_data, re.IGNORECASE)
    po_date = re.search(r"PO\s*DATE[:\-]?\s*(\S+)", text_data, re.IGNORECASE)
    expiry = re.search(r"EXPIRY\s*DATE[:\-]?\s*(\S+)", text_data, re.IGNORECASE)
    gst = re.search(r"GST\s*NO[:\-]?\s*(\S+)", text_data, re.IGNORECASE)

    def format_date(date_str):
        try:
            return datetime.strptime(date_str, "%d.%m.%Y").strftime("%d-%m-%Y")
        except:
            return date_str

    return {
        "PO NO": po_no.group(1) if po_no else None,
        "PO DATE": format_date(po_date.group(1)) if po_date else None,
        "PO EXPIRY DATE": format_date(expiry.group(1)) if expiry else None,
        "GST NO": gst.group(1) if gst else None
    }

def process_excel(file):
    if file.name.endswith(".csv"):
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file, engine="openpyxl")

    df.columns = df.columns.str.strip()
    return df

if pdf_files and excel_files:

    pdf_dict = {normalize_filename(f.name): f for f in pdf_files}
    excel_dict = {normalize_filename(f.name): f for f in excel_files}

    # Validation check
    pdf_keys = set(pdf_dict.keys())
    excel_keys = set(excel_dict.keys())

    missing_excels = pdf_keys - excel_keys
    missing_pdfs = excel_keys - pdf_keys

    if missing_excels or missing_pdfs:
        st.error("Mismatch detected between PDF and Excel files.")

        if missing_excels:
            st.error(f"Missing Excel for: {', '.join(missing_excels)}")

        if missing_pdfs:
            st.error(f"Missing PDF for: {', '.join(missing_pdfs)}")

        st.stop()

    consolidated_rows = []

    for key in pdf_keys:
        pdf_data = extract_pdf_data(pdf_dict[key])
        excel_df = process_excel(excel_dict[key])

        # Attach PDF header info to each Excel row
        for col, value in pdf_data.items():
            excel_df[col] = value

        excel_df["SOURCE FILE"] = key
        consolidated_rows.append(excel_df)

    final_df = pd.concat(consolidated_rows, ignore_index=True)

    st.success("Files processed successfully.")

    # Download button
    output = BytesIO()
    final_df.to_excel(output, index=False, engine="openpyxl")
    output.seek(0)

    st.download_button(
        label="Download Consolidated File",
        data=output,
        file_name="Blinkit_PO_Consolidated.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("Upload both PDF and Excel files to begin.")