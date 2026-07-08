import streamlit as st
import pandas as pd
import io
import msoffcrypto
import urllib.parse
from sqlalchemy import create_engine, text

st.set_page_config(page_title="Universal Excel Database Suppressor", layout="centered")
st.title("📊 Database Suppression & Cleaner")
st.write("Automatically remove duplicates and cross-check phone numbers against the secure database.")

# --- BACKGROUND SECRETS CONFIGURATION ---
try:
    db_host = st.secrets["database"]["host"]
    db_port = str(st.secrets["database"]["port"])
    db_user = st.secrets["database"]["user"]
    db_pass = st.secrets["database"]["password"]
    db_name = st.secrets["database"]["database_name"]
    db_table = st.secrets["database"]["table_name"]
    db_phone_col = st.secrets["database"]["phone_column"]
except KeyError:
    st.error("🔒 App Secrets are missing! Please configure the database credentials in the Streamlit Cloud Dashboard.")
    st.stop()

# --- SIDEBAR: Now Only Contains Decryption Options ---
with st.sidebar:
    st.header("🔐 File Decryption")
    file_password = st.text_input(
        "File Password (If encrypted):", 
        type="password", 
        help="If the uploaded Excel workbook is password-protected, type the password here."
    )
    st.markdown("---")
    st.caption("Database credentials are encrypted and securely managed in the background.")

# --- ROBUST NATIVE CONNECTION BUILDER ---
def get_clean_engine(user, password, host, port, dbname):
    safe_password = urllib.parse.quote_plus(password)
    connection_str = f"mysql+mysqlconnector://{user}:{safe_password}@{host}:{port}/{dbname}"
    return create_engine(connection_str, connect_args={"connection_timeout": 5})

# --- UI WORKFLOW: Step 1 (Auto-connects behind the scenes) ---
st.markdown("### 🔍 Step 1: Database Status Verification")

if "db_verified" not in st.session_state:
    try:
        engine = get_clean_engine(db_user, db_pass, db_host, db_port, db_name)
        with engine.connect() as conn:
            conn.execute(text("SELECT 1"))
        st.session_state["db_verified"] = True
    except Exception as e:
        st.session_state["db_verified"] = False

if st.session_state.get("db_verified", False):
    st.success("✅ Secure database connection established automatically!")
    st.markdown("---")
    st.markdown("### 📥 Step 2: Upload Data File")
    
    def robust_load_file(uploaded_file, password=None):
        file_bytes = uploaded_file.read()
        file_stream = io.BytesIO(file_bytes)
        error_logs = []
        
        try:
            office_file = msoffcrypto.OfficeFile(file_stream)
            if office_file.is_encrypted():
                decrypted_stream = io.BytesIO()
                office_file.load_key(password=password if password else "VelvetSweatshop")
                office_file.decrypt(decrypted_stream)
                decrypted_stream.seek(0)
                file_stream = decrypted_stream
        except Exception as e:
            error_logs.append(f"Decryption check skipped: {e}")
            file_stream.seek(0)

        try: return pd.read_excel(file_stream, engine='openpyxl')
        except Exception as e: error_logs.append(f"Openpyxl engine failed: {e}"); file_stream.seek(0)
        try: return pd.read_excel(file_stream, engine='xlrd')
        except Exception as e: error_logs.append(f"Xlrd engine failed: {e}"); file_stream.seek(0)
        try:
            html_tables = pd.read_html(file_stream)
            if html_tables: return html_tables[0]
        except Exception as e: error_logs.append(f"HTML table parse failed: {e}"); file_stream.seek(0)
        try:
            text_data = file_stream.read().decode('utf-8', errors='ignore')
            return pd.read_csv(io.StringIO(text_data))
        except Exception as e: error_logs.append(f"CSV parse failed: {e}"); file_stream.seek(0)
        try:
            text_data = file_stream.read().decode('utf-8', errors='ignore')
            return pd.read_csv(io.StringIO(text_data), sep='\t')
        except Exception as e: error_logs.append(f"TSV/TXT parse failed: {e}")
            
        detailed_error_msg = "\n".join([f"- {log}" for log in error_logs])
        raise ValueError(f"Unsupported or heavily corrupted file format.\n\n**Engine Diagnostic Breakdown:**\n{detailed_error_msg}")

    uploaded_file = st.file_uploader("Choose a data file to process", type=["xlsx", "xls", "csv", "tsv", "txt"])

    if uploaded_file is not None:
        try:
            df = robust_load_file(uploaded_file, password=file_password if file_password else None)
            st.success(f"Successfully loaded file! Rows found: {len(df)}")
            
            all_columns = df.columns.tolist()
            default_idx = 0
            for i, col in enumerate(all_columns):
                if 'phone' in str(col).lower() or 'tel' in str(col).lower():
                    default_idx = i
                    break
                    
            selected_phone_col = st.selectbox("Select the 'Phone Number' column from the uploaded file:", options=all_columns, index=default_idx)
            
            if st.button("🚀 Run Database Suppression"):
                with st.spinner("Processing data layers and generating sheets..."):
                    
                    # 1. Clean formatting and standardize string types
                    df[selected_phone_col] = df[selected_phone_col].astype(str).str.replace(r'\s+|\.0$', '', regex=True).str.strip()
                    
                    # 2. Filter out blanks completely
                    initial_count = len(df)
                    is_blank = df[selected_phone_col].isin(['', 'nan', 'None', 'NaN']) | df[selected_phone_col].isna()
                    df_valid = df[~is_blank].copy()
                    blank_records_count = initial_count - len(df_valid)
                    
                    # 3. Separate Internal File Duplicates
                    df_internal_clean = df_valid.drop_duplicates(subset=[selected_phone_col], keep="first")
                    df_duplicates = df_valid[df_valid.duplicated(subset=[selected_phone_col], keep="first")]
                    internal_dupes = len(df_duplicates)
                    
                    # 4. Fetch database records for cross-checking
                    engine = get_clean_engine(db_user, db_pass, db_host, db_port, db_name)
                    query = f"SELECT `{db_phone_col}` FROM `{db_table}`"
                    db_phones_df = pd.read_sql(query, con=engine)
                    
                    db_phones_df[db_phone_col] = db_phones_df[db_phone_col].astype(str).str.replace(r'\s+|\.0$', '', regex=True).str.strip()
                    db_phones_set = set(db_phones_df[~db_phones_df[db_phone_col].isin(['', 'nan', 'None', 'NaN'])][db_phone_col])
                    
                    # 5. Separate Matches in DB from Net Unique Output
                    in_db_mask = df_internal_clean[selected_phone_col].isin(db_phones_set)
                    
                    df_matched_db = df_internal_clean[in_db_mask]
                    df_net_unique = df_internal_clean[~in_db_mask]
                    db_suppressed_count = len(df_matched_db)
                    
                    # --- Display Metrics Breakdown ---
                    st.markdown("### 📊 Metrics Summary")
                    col1, col2, col3, col4 = st.columns(4)
                    col1.metric("Skipped Blanks", f"{blank_records_count}")
                    col2.metric("File Duplicates", f"{internal_dupes}")
                    col3.metric("Matched in DB", f"{db_suppressed_count}")
                    col4.metric("Net Unique Output", f"{len(df_net_unique)}")
                    
                    # --- Generate Multi-Sheet Excel Output ---
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                        df_net_unique.to_excel(writer, index=False, sheet_name='Net Unique Output')
                        df_matched_db.to_excel(writer, index=False, sheet_name='Matched in DB')
                        df_duplicates.to_excel(writer, index=False, sheet_name='File Duplicates')
                    
                    st.download_button(
                        label="📥 Download Complete Suppressed Package",
                        data=buffer.getvalue(),
                        file_name=f"suppressed_package_{uploaded_file.name if '.' in uploaded_file.name else 'data.xlsx'}",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                        
        except Exception as e:
            st.error(f"Error Processing Pipeline: {e}")
else:
    st.error(f"❌ Connection Failed using background credentials. Verify your Pinggy tunnel is active in your terminal and secrets are up to date.")
    if st.button("🔄 Retry Connection"):
        st.rerun()
