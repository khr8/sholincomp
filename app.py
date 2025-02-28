import os
import time
import io
import logging
import zipfile
import pandas as pd
import streamlit as st

# Setup logging (optional – logs will be written to a file in the user's Documents folder)
LOG_DIR = os.path.join(os.path.expanduser("~"), "Documents", "MyAppLogs")
os.makedirs(LOG_DIR, exist_ok=True)
log_file = os.path.join(LOG_DIR, "comparison_log.txt")
logging.basicConfig(
    filename=log_file,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)
logging.info("Log file created or appended successfully on Streamlit launch.")

# Required columns lists
REQUIRED_COLUMNS_ISBN = [
    "ISBN13", "TITLE", "AUTHOR", "DISCOUNT", "STOCK", "CUR",
    "DIM1", "DIM2", "DIM3", "WEIGHT", "PUBLISHER", "IMPRINT"
]
REQUIRED_COLUMNS_EAN = [
    "EAN #", "TITLE", "QTYAV", "CUR", "PRICE", "AUTHOR",
    "PUBLISHER", "WGT OZS", "LENGTH", "WIDTH", "HEIGHT", "CD"
]

def clean_isbn(isbn):
    if pd.isna(isbn):
        return ""
    x = str(isbn).strip().replace("\u200b", "").replace("\xa0", "").replace("\ufeff", "")
    return x.zfill(13) if x.isnumeric() and len(x) < 13 else x

def detect_header(file_obj):
    # Reset pointer
    file_obj.seek(0)
    filename = file_obj.name.lower() if hasattr(file_obj, "name") else ""
    if filename.endswith(".csv"):
        df = pd.read_csv(file_obj, encoding="ISO-8859-1", dtype=str, errors="replace", nrows=20, header=None)
    else:
        df = pd.read_excel(file_obj, dtype=str, nrows=20, header=None)
    for i, row in df.iterrows():
        vals = row.astype(str).str.upper().str.replace(r"[^\w\s]", "", regex=True)
        if any("ISBN13" in v.replace(" ", "") or "EAN" in v.replace(" ", "") for v in vals):
            return i
    raise KeyError("No valid header row found.")

def process_file(file_obj):
    # Use the file's name attribute to determine extension
    filename = file_obj.name.lower() if hasattr(file_obj, "name") else ""
    hdr = detect_header(file_obj)
    file_obj.seek(0)
    if filename.endswith(".csv"):
        df = pd.read_csv(file_obj, encoding="ISO-8859-1", dtype=str, errors="replace", header=hdr)
    else:
        df = pd.read_excel(file_obj, dtype=str, header=hdr)
    # Clean column names
    df.columns = df.columns.str.strip().str.upper().str.replace(r"[^\w\s]", "", regex=True)
    # Identify the ISBN/EAN column
    c = next((col for col in df.columns if "ISBN13" in col.replace(" ", "") or "EAN" in col.replace(" ", "")), None)
    if not c:
        raise KeyError("No ISBN/EAN column found.")
    df[c] = df[c].apply(clean_isbn)
    # Convert stock / quantity columns to numeric if present
    s = next((col for col in df.columns if "STOCK" in col.replace(" ", "") or "QTYAV" in col.replace(" ", "")), None)
    if s:
        df[s] = pd.to_numeric(df[s], errors="coerce")
    return df, c

def extract_isbns(file_obj):
    try:
        file_obj.seek(0)
        df = pd.read_excel(file_obj, header=None, dtype=str)
        vals = df.values.flatten()
        return {clean_isbn(v) for v in vals if str(v).isnumeric()}
    except Exception as e:
        logging.error(f"Error reading ISBN removal file: {e}")
        return set()

def clean_file(file_obj, cur, rem_obj=None):
    df, c = process_file(file_obj)
    # Decide which required columns to use
    if "ISBN13" in df.columns:
        req = REQUIRED_COLUMNS_ISBN
    else:
        req = REQUIRED_COLUMNS_EAN
    if rem_obj:
        rset = extract_isbns(rem_obj)
        df = df[~df[c].isin(rset)]
    # Keep only required columns if they exist
    df = df[[x for x in req if x in df.columns]]
    if "CUR" not in df.columns:
        df["CUR"] = ""
    df["CUR"] = cur
    df = df.reindex(columns=req, fill_value="")
    return df

def to_excel_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

def prepare_file(uploaded_file):
    """
    Read the uploaded file into a BytesIO buffer and set its name attribute.
    """
    file_bytes = uploaded_file.read()
    file_obj = io.BytesIO(file_bytes)
    file_obj.name = uploaded_file.name
    return file_obj

# ------------------ Streamlit App UI ------------------

st.title("File Cleaner & Comparator")

st.write("Upload **2 comparison files** (CSV or Excel) and an optional **ISBN Removal File** (Excel).")

col1, col2 = st.columns(2)
with col1:
    uploaded_file1 = st.file_uploader("Upload First Comparison File", type=["csv", "xlsx"])
with col2:
    uploaded_file2 = st.file_uploader("Upload Second Comparison File", type=["csv", "xlsx"])

uploaded_removal = st.file_uploader("Upload ISBN Removal File (Optional)", type=["xlsx"])

# Currency selection
currency = st.radio("Select Currency", options=["USD", "GBP"], index=0)

if st.button("Start Cleaning & Comparison"):
    if not uploaded_file1 or not uploaded_file2:
        st.error("Please upload both comparison files.")
    else:
        start_time = time.time()
        progress_bar = st.progress(0)
        status_text = st.empty()
        try:
            # Prepare files from uploaders
            file1 = prepare_file(uploaded_file1)
            file2 = prepare_file(uploaded_file2)
            rem_file = prepare_file(uploaded_removal) if uploaded_removal else None

            progress_bar.progress(10)
            status_text.text("Cleaning first file...")
            d1 = clean_file(file1, currency, rem_file)

            progress_bar.progress(30)
            status_text.text("Cleaning second file...")
            d2 = clean_file(file2, currency, rem_file)

            progress_bar.progress(50)
            status_text.text("Comparing files...")
            # Get key columns from each cleaned file (first column)
            k1 = d1.columns[0]
            k2 = d2.columns[0]
            new_items = d2[~d2[k2].isin(d1[k1])]
            inactive_items = d1[~d1[k1].isin(d2[k2])]

            progress_bar.progress(80)
            elapsed = round(time.time() - start_time, 2)
            status_text.text(f"✅ Processing completed in {elapsed} seconds.")

            # Convert DataFrames to Excel bytes
            cleaned1_bytes = to_excel_bytes(d1)
            cleaned2_bytes = to_excel_bytes(d2)
            new_items_bytes = to_excel_bytes(new_items)
            inactive_items_bytes = to_excel_bytes(inactive_items)

            # Create an in-memory zip file with the folder structure
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
                # Add cleaned files into a "cleaned" folder
                zipf.writestr("cleaned/Cleaned_File1.xlsx", cleaned1_bytes)
                zipf.writestr("cleaned/Cleaned_File2.xlsx", cleaned2_bytes)
                # Add comparison files into a "comparison" folder
                zipf.writestr("comparison/New_Items.xlsx", new_items_bytes)
                zipf.writestr("comparison/Inactive_Items.xlsx", inactive_items_bytes)
            zip_buffer.seek(0)
            zip_data = zip_buffer.getvalue()

            st.success(f"Processing completed in {elapsed} seconds.")
            st.download_button("Download All Files (ZIP)", zip_data, file_name="Comparison_Output.zip", mime="application/zip")
            progress_bar.progress(100)
            logging.info(f"--- Comparison finished in {elapsed} seconds.")
        except Exception as e:
            st.error(f"An error occurred: {e}")
            logging.error(f"Processing error: {e}")
