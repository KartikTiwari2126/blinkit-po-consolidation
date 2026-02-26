import streamlit as st
import pdfplumber
import re
import io
from datetime import datetime
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Blinkit PO Consolidation", layout="wide")
st.title("Blinkit PO Consolidation Tool")
st.markdown("Upload matching Excel and PDF files.")

# ================= ALWAYS VISIBLE RESET =================
col1, col2 = st.columns([8, 1])
with col2:
    if st.button("Reset"):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()

# ================= FILE UPLOAD =================

excel_files = st.file_uploader(
    "Upload EXCEL files",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

pdf_files = st.file_uploader(
    "Upload PDF files",
    type=["pdf"],
    accept_multiple_files=True
)

# ================= FILE COUNT DISPLAY =================

if excel_files or pdf_files:
    col1, col2 = st.columns(2)

    with col1:
        st.metric("Excel Files Uploaded", len(excel_files) if excel_files else 0)

    with col2:
        st.metric("PDF Files Uploaded", len(pdf_files) if pdf_files else 0)

# ================= MAIN PROCESS =================

if excel_files and pdf_files:

    if st.button("Consolidate"):

        with st.spinner("Processing files..."):

            uploaded_excel = {f.name: f.read() for f in excel_files}
            uploaded_pdf = {f.name: f.read() for f in pdf_files}

            # ================= VALIDATION =================
            excel_keys = {name.rsplit(".", 1)[0] for name in uploaded_excel.keys()}
            pdf_keys = {name.rsplit(".", 1)[0] for name in uploaded_pdf.keys()}

            missing_excels = pdf_keys - excel_keys
            missing_pdfs = excel_keys - pdf_keys

            if missing_excels or missing_pdfs:
                st.error("Mismatch detected between uploaded files.")

                if missing_excels:
                    st.error(f"Missing Excel for: {', '.join(missing_excels)}")

                if missing_pdfs:
                    st.error(f"Missing PDF for: {', '.join(missing_pdfs)}")

                st.stop()

            # ================= SHIPPING ADDRESS =================
            def extract_shipping_address(filedata):

                with pdfplumber.open(io.BytesIO(filedata)) as pdf:
                    text = "\n".join([(page.extract_text() or "") for page in pdf.pages])

                lines = [l.strip() for l in text.splitlines() if l.strip()]

                company = ""
                address_lines = []
                capture_address = False

                for line in lines:

                    lower = line.lower()

                    if lower.startswith("delivered"):
                        parts = line.split(":", 1)
                        if len(parts) > 1:
                            temp = parts[1]
                            temp = re.split(r"gst", temp, flags=re.IGNORECASE)[0]
                            company = temp.strip()
                        continue

                    if lower.startswith("to "):
                        capture_address = True
                        addr_part = line[3:].strip()
                        addr_part = re.split(r"reference", addr_part, flags=re.IGNORECASE)[0]
                        if addr_part:
                            address_lines.append(addr_part)
                        continue

                    if capture_address:
                        if line.startswith("#"):
                            break

                        clean_line = re.split(r"reference", line, flags=re.IGNORECASE)[0].strip()

                        if clean_line:
                            address_lines.append(clean_line)

                full_address = company + "\n" + "\n".join(address_lines)
                return full_address.strip()

            # ================= PDF LOGIC =================

            def clean_spaces(x):
                return re.sub(r"\s+", " ", x).strip()

            def to_ddmmyyyy(val):
                if not val:
                    return ""
                val = val.replace(".", "")
                if "," in val:
                    parts = val.split(",")
                    val = ",".join(parts[:2])
                for fmt in ["%b %d, %Y", "%B %d, %Y", "%d %b %Y", "%d %B %Y"]:
                    try:
                        return datetime.strptime(val.strip(), fmt).strftime("%d-%m-%Y")
                    except:
                        pass
                return val

            pdf_records = []

            for filename, filedata in uploaded_pdf.items():

                with pdfplumber.open(io.BytesIO(filedata)) as pdf:
                    text = "\n".join([(page.extract_text() or "") for page in pdf.pages])

                rec = {}

                m = re.search(r"P\.?O\.?\s*Number\s*:\s*(.*)", text, re.IGNORECASE)
                rec["PO NO"] = clean_spaces(m.group(1)) if m else ""

                m = re.search(r"\bDate\s*:\s*(.*)", text, re.IGNORECASE)
                rec["PO DATE"] = to_ddmmyyyy(clean_spaces(m.group(1))) if m else ""

                m = re.search(r"PO\s*expiry\s*date\s*:\s*(.*)", text, re.IGNORECASE)
                rec["PO EXPIRY DATE"] = to_ddmmyyyy(clean_spaces(m.group(1))) if m else ""

                # CLIENT NAME LOGIC
                lines = [l.strip() for l in text.splitlines() if l.strip()]
                client = ""

                for l in lines:
                    if re.search(r"(feeder\s*warehouse|feeder)\s*$", l, re.IGNORECASE):
                        client = clean_spaces(l)
                        break

                if not client:
                    try:
                        start = next(i for i, l in enumerate(lines) if l.lower().startswith("cin"))
                        end   = next(i for i, l in enumerate(lines) if l.lower().startswith("contact name"))
                        block = lines[start:end]

                        for l in block:
                            ll = l.lower()
                            if "pan" not in ll and "cin" not in ll:
                                client = clean_spaces(l)
                                break
                    except:
                        pass

                rec["client name"] = client
                rec["SHIPPING ADDRESS"] = extract_shipping_address(filedata)

                matches = re.findall(r"GST\s*No\.?\s*:\s*([A-Z0-9]{15})", text, re.IGNORECASE)
                rec["GST Number"] = matches[-1] if matches else ""

                rec["base"] = filename.rsplit(".", 1)[0]
                rec["base_key"] = rec["base"].split("_")[0]

                pdf_records.append(rec)

            df_pdf = pd.DataFrame(pdf_records)

            # ================= EXCEL LOGIC =================

            required_cols = [
                "#","Item Code","HSN Code","Product UPC","Product Description",
                "Grammage","Basic Cost Price","CGST %","SGST %","IGST %",
                "CESS %","Additional CES","Tax Amount","Landing Rate",
                "Quantity","MRP","Margin %","Total Amount"
            ]

            excel_rows = []

            for filename, filedata in uploaded_excel.items():
                xls = pd.ExcelFile(io.BytesIO(filedata))
                sheet = xls.sheet_names[0]

                df = pd.read_excel(xls, sheet_name=sheet, header=None)

                header_row = 0
                for i in range(min(15, len(df))):
                    row_text = " ".join(df.iloc[i].astype(str)).lower()
                    if "item" in row_text and "code" in row_text:
                        header_row = i
                        break

                df = pd.read_excel(xls, sheet_name=sheet, header=header_row)
                df.columns = [str(c).strip() for c in df.columns]

                keep = [c for c in required_cols if c in df.columns]
                df = df[keep]

                if "#" in df.columns:
                    df = df[df["#"].notna()]

                df["base"] = filename.rsplit(".", 1)[0]
                df["base_key"] = df["base"].str.split("_").str[0]

                excel_rows.append(df)

            df_excel = pd.concat(excel_rows, ignore_index=True)

            # ================= MERGE =================

            final_df = df_excel.merge(
                df_pdf,
                on="base_key",
                how="left"
            )

            first_cols = [
                "PO NO","PO DATE","PO EXPIRY DATE",
                "client name","SHIPPING ADDRESS","GST Number"
            ]

            cols = first_cols + [c for c in final_df.columns if c not in first_cols and c != "base_key"]
            final_df = final_df[cols]

            st.success("Processing completed successfully.")

            output = BytesIO()
            final_df.to_excel(output, index=False)
            output.seek(0)

            st.download_button(
                label="Download Consolidated File",
                data=output,
                file_name="final_consolidated.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # Memory cleanup
            del final_df
            del df_excel
            del df_pdf
            excel_rows.clear()
            pdf_records.clear()

else:
    st.info("Please upload both Excel and PDF files.")

