import streamlit as st
import pandas as pd
import io
import msoffcrypto

st.set_page_config(page_title="Universal Excel & Data Cleaner", layout="centered")
st.title("📊 Universal File Cleaner")
st.write("Upload any data file (Excel, Legacy Excel, Encrypted, or CSV) to automatically remove duplicate records.")

def robust_load_file(uploaded_file, password=None):
    """
    Attempts to safely read an uploaded file regardless of whether it's
    a true .xlsx, legacy .xls, disguised text file, or encrypted workbook.
    """
    # Read raw bytes into memory
    file_bytes = uploaded_file.read()
    file_stream = io.BytesIO(file_bytes)
    
    # --- STEP 1: Handle Potential Encryption / Passwords ---
    try:
        office_file = msoffcrypto.OfficeFile(file_stream)
        if office_file.is_encrypted():
            decrypted_stream = io.BytesIO()
            if password:
                office_file.load_key(password=password)
            else:
                # Try default/empty password fallback used by some systems
                office_file.load_key(password="VelvetSweatshop") 
            
            office_file.decrypt(decrypted_stream)
            decrypted_stream.seek(0)
            file_stream = decrypted_stream  # Use decrypted bytes going forward
    except Exception as crypto_err:
        # File isn't a standard OLE/OOXML encrypted structure, proceed normally
        file_stream.seek(0)

    # --- STEP 2: Try Parsing Frameworks (xlsx -> xls -> csv -> tsv) ---
    # Attempt 1: Standard Modern Excel (.xlsx)
    try:
        return pd.read_excel(file_stream, engine='openpyxl')
    except Exception:
        file_stream.seek(0)
        
    # Attempt 2: Legacy Excel Binary (.xls)
    try:
        return pd.read_excel(file_stream, engine='xlrd')
    except Exception:
        file_stream.seek(0)
        
    # Attempt 3: Standard Comma-Separated Values (.csv)
    try:
        # Decode byte stream safely to text handling mixed encoding anomalies
        text_data = file_stream.read().decode('utf-8', errors='ignore')
        return pd.read_csv(io.StringIO(text_data))
    except Exception:
        file_stream.seek(0)
        
    # Attempt 4: Tab-Separated Values (.tsv / txt log formatting)
    try:
        text_data = file_stream.read().decode('utf-8', errors='ignore')
        return pd.read_csv(io.StringIO(text_data), sep='\t')
    except Exception:
        raise ValueError(
            "Unsupported or heavily corrupted file format. Verify the file opens "
            "locally or check your decryption password."
        )

# Optional Password Input Field UI
with st.sidebar:
    st.header("File Settings")
    file_password = st.text_input(
        "File Password (Leave blank if not encrypted):", 
        type="password", 
        help="If the uploaded Excel workbook is encrypted/password-protected, type the password here."
    )

# File Uploader Widget
uploaded_file = st.file_uploader(
    "Choose a data file", 
    type=["xlsx", "xls", "csv", "tsv", "txt"]
)

if uploaded_file is not None:
    try:
        # Load using the universal parser framework
        df = robust_load_file(uploaded_file, password=file_password if file_password else None)
        
        st.success(f"Successfully loaded file! Rows found: {len(df)}")
        
        # Select matching columns or default to checking the entire row
        all_columns = df.columns.tolist()
        selected_cols = st.multiselect(
            "Select columns to check for duplicates (Leave empty to check entire rows):", 
            options=all_columns
        )
        
        # Deduplicate
        subset_param = selected_cols if selected_cols else None
        df_cleaned = df.drop_duplicates(subset=subset_param, keep="first")
        
        removed_count = len(df) - len(df_cleaned)
        st.info(f"Removed {removed_count} duplicate rows. Remaining unique rows: {len(df_cleaned)}.")
        
        # Convert clean dataset back to clean modern Excel bytes
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_cleaned.to_excel(writer, index=False, sheet_name='Cleaned Data')
        
        # Download Button
        st.download_button(
            label="📥 Download Cleaned Excel File",
            data=buffer.getvalue(),
            file_name=f"deduped_{uploaded_file.name if '.' in uploaded_file.name else 'data.xlsx'}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"Error processing file: {e}")
