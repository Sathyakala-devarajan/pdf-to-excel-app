
import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import re
import io

# --------- PDF Extraction Logic ---------
def extract_product_details_from_pdf(pdf_file):
    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
    text = "".join([page.get_text() for page in doc])
    lines = text.splitlines()

    product_data = {}
    i = 0

    while i < len(lines) - 6:
        line = lines[i].strip()

        compact_match = re.match(r"^([A-Z0-9]{4,10})\s+.*?\s+(\d+(\.\d+)?)\s+(\d+(\.\d+)?)\s+1\s+\d+$", line)
        if compact_match:
            code = compact_match.group(1)
            pack = float(compact_match.group(2))
            price = float(compact_match.group(4))
            product_data[code] = {"pack": pack, "price": price}
            i += 1
            continue

        if re.fullmatch(r"[A-Z0-9]{4,10}", line):
            try:
                code = line
                next_lines = [lines[i + j].strip() for j in range(1, 7)]
                numeric_values = [l for l in next_lines if re.match(r"^\d+(\.\d+)?$", l)]
                if len(numeric_values) >= 2:
                    pack_size = float(numeric_values[0])
                    price = float(numeric_values[1])
                    product_data[code] = {"pack": pack_size, "price": price}
                i += 7
                continue
            except Exception:
                pass
        i += 1

    return product_data

# --------- Excel Matching Logic ---------
def process_files(excel_file, pdf_file):
    product_data = extract_product_details_from_pdf(pdf_file)
    df = pd.read_excel(excel_file, sheet_name="AAH DATA")
    df.columns = df.columns.map(str)

    col_product_code = next((c for c in df.columns if "product cod" in c.lower() and "aah" not in c.lower()), None)
    col_aah_code = next((c for c in df.columns if "aah product co" in c.lower()), None)
    col_aah_desc = next((c for c in df.columns if "desc" in c.lower()), None)

    if not all([col_product_code, col_aah_code, col_aah_desc]):
        raise ValueError("Required columns not found. Please check Excel headers.")

    df[col_aah_code] = df[col_aah_code].astype(str).str.strip().str.upper()
    df_matched = df[df[col_aah_code].isin(product_data.keys())].copy()
    df_matched["Pack Size"] = df_matched[col_aah_code].map(lambda x: product_data[x]["pack"])
    df_matched["Price"] = df_matched[col_aah_code].map(lambda x: product_data[x]["price"])

    output_df = df_matched[[col_product_code, "Pack Size", "Price"]]
    output_df.columns = ["SKU Code", "Quantity", "Price"]
    return output_df

# --------- Streamlit UI ---------
st.set_page_config(page_title="📦 Orders Matching Tool", layout="wide")
st.title("📦 Product Matching from PDF and Excel")

with st.sidebar:
    st.header("Upload Files")
    pdf_file = st.file_uploader("📄 Upload Product PDF", type=["pdf"])
    excel_file = st.file_uploader("📊 Upload Excel File", type=["xls", "xlsx"])
    generate = st.button("🔍 Generate Output")

with st.expander("ℹ️ How It Works"):
    st.markdown("""
    - Upload your **product list in PDF** and **order data in Excel**
    - It extracts the pack size and price from the PDF
    - Matches them with the **AAH Product Code** in Excel
    - Downloads final matched output as Excel
    """)

if generate:
    if not pdf_file or not excel_file:
        st.warning("⚠️ Please upload both PDF and Excel files.")
    else:
        try:
            progress = st.progress(0)
            with st.spinner("Processing... Please wait."):
                progress.progress(10)
                output_df = process_files(excel_file, pdf_file)
                progress.progress(100)

            st.success("✅ Matching completed!")

            tab1, tab2 = st.tabs(["📋 View Output", "📥 Download"])

            with tab1:
                st.dataframe(output_df, use_container_width=True)

            with tab2:
                buffer = io.BytesIO()
                output_df.to_excel(buffer, index=False)
                buffer.seek(0)
                st.download_button(
                    label="📥 Download Matched Excel",
                    data=buffer,
                    file_name="matched_output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"❌ Error occurred: {e}")

# --------- File Previews ---------
if excel_file:
    st.sidebar.markdown("✅ Excel uploaded:")
    try:
        df_preview = pd.read_excel(excel_file)
        st.sidebar.dataframe(df_preview.head(5))
    except Exception:
        st.sidebar.warning("Couldn't preview Excel.")

if pdf_file:
    st.sidebar.markdown("✅ PDF uploaded:")
    try:
        pdf_preview = fitz.open(stream=pdf_file.read(), filetype="pdf")
        st.sidebar.text_area("PDF Page 1 Preview", pdf_preview[0].get_text(), height=150)
    except Exception:
        st.sidebar.warning("Couldn't preview PDF.")
