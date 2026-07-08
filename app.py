import streamlit as st
import pandas as pd
import io
import msoffcrypto
import urllib.parse
from sqlalchemy import create_engine

st.set_page_config(page_title="Universal Excel Database Suppressor", layout="centered")
st.title("📊 Database Suppression & Cleaner")
st.write("Upload a file to automatically remove duplicates and cross-check phone numbers against your local MySQL database.")

# --- SIDEBAR: Database & File Settings ---
with st.sidebar:
    st.header("🗄️ MySQL Connection Settings")
    db_host = st.text_input("DB Host:", value="localhost")
    db_user = st.text_input("DB User:", value="root")
    db_pass = st.text_input("DB Password:", type="password", value="")
    db_name = st.text_input("Database Name:", value="lead_generation")
    db_table = st.text_input("Table Name:", value="contacts")
    db_phone_col = st.text_input("DB Phone Column Name:", value="phone_number")
    
    st.markdown("---")
    st.header("🔐 File Decryption")
    file_password = st.text_input(
        "File Password (If encrypted):", 
        type="password", 
        help="If the uploaded Excel workbook is password-protected, type the password here."
    )

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
            office_file.load_key(password=password if password else "VelvetSweatshop")
            office_file.decrypt(decrypted_stream)
            decrypted_stream.seek(0)
            file_stream = decrypted_stream  # Use decrypted bytes going forward
    except Exception:
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
        text_data = file_stream.read().decode('utf-8', errors='ignore')
        return pd.read_csv(io.StringIO(text_data))
    except Exception:
        file_stream.seek(0)
        
    # Attempt 4: Tab-Separated Values (.tsv / txt)
    try:
        text_data = file_stream.read().decode('utf-8', errors='ignore')
        return pd.read_csv(io.StringIO(text_data), sep='\t')
    except Exception:
        raise ValueError(
            "Unsupported or heavily corrupted file format. Verify the file opens "
            "locally or check your decryption password."
        )

# --- CORE PROCESSING WORKFLOW ---
uploaded_file = st.file_uploader(
    "Choose a data file to process", 
    type=["xlsx", "xls", "csv", "tsv", "txt"]
)

if uploaded_file is not None:
    try:
        # 1. Load File Robustly
        df = robust_load_file(uploaded_file, password=file_password if file_password else None)
        st.success(f"Successfully loaded file! Rows found: {len(df)}")
        
        # 2. Select Phone Column from Uploaded File
        all_columns = df.columns.tolist()
        
        # Auto-detect column containing 'phone' or 'tel' to assist non-tech users
        default_idx = 0
        for i, col in enumerate(all_columns):
            if 'phone' in str(col).lower() or 'tel' in str(col).lower():
                default_idx = i
                break
                
        selected_phone_col = st.selectbox(
            "Select the 'Phone Number' column from the uploaded file:", 
            options=all_columns,
            index=default_idx
        )
        
        if st.button("🚀 Run Database Suppression"):
            with st.spinner("Connecting to MySQL and pulling existing records..."):
                # Clean phone values in memory to avoid formatting mismatches (removes spaces, decimals, etc.)
                df[selected_phone_col] = df[selected_phone_col].astype(str).str.replace(r'\s+|\.0$', '', regex=True).str.strip()
                
                # Deduplicate the file internally first to minimize downstream overhead
                df_internal_clean = df.drop_duplicates(subset=[selected_phone_col], keep="first")
                internal_dupes = len(df) - len(df_internal_clean)
                
                # Safe URL Encoding for passwords containing special characters like '@'
                safe_password = urllib.parse.quote_plus(db_pass)
                
                # Connect to MySQL Database
                connection_str = f"mysql+pymysql://{db_user}:{safe_password}@{db_host}/{db_name}"
                engine = create_engine(connection_str)
                
                # Securely stream just the targeted phone column from the DB
                query = f"SELECT `{db_phone_col}` FROM `{db_table}`"
                db_phones_df = pd.read_sql(query, con=engine)
                
                # Format database numbers uniformly to match the clean uploaded data structure
                db_phones_set = set(db_phones_df[db_phone_col].astype(str).str.replace(r'\s+|\.0$', '', regex=True).str.strip())
                
                # 3. Filter data: keep only if phone is NOT present in database set
                df_final = df_internal_clean[~df_internal_clean[selected_phone_col].isin(db_phones_set)]
                
                db_suppressed_count = len(df_internal_clean) - len(df_final)
                
                # --- Metrics Breakdown Layout ---
                st.markdown("### 📊 Metrics Summary")
                col1, col2, col3 = st.columns(3)
                col1.metric("Internal File Duplicates Removed", f"{internal_dupes}")
                col2.metric("Existing Records Matched in DB", f"{db_suppressed_count}")
                col3.metric("Net New Unique Records", f"{len(df_final)}")
                
                if len(df_final) == 0:
                    st.warning("All records in this file already exist within your database!")
                else:
                    # Convert remaining fresh entries back to Excel bytes for download
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                        df_final.to_excel(writer, index=False, sheet_name='New Uniques Only')
                    
                    st.download_button(
                        label="📥 Download Cleaned & Suppressed Excel File",
                        data=buffer.getvalue(),
                        file_name=f"suppressed_{uploaded_file.name if '.' in uploaded_file.name else 'data.xlsx'}",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
    except Exception as e:
        st.error(f"Error Processing Pipeline: {e}")
