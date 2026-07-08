import streamlit as st
import pandas as pd
import io
import msoffcrypto
import urllib.parse
from sqlalchemy import create_engine, text

st.set_page_config(page_title="Universal Excel Database Suppressor", layout="centered")
st.title("📊 Database Suppression & Cleaner")
st.write("Upload a file to automatically remove duplicates and cross-check phone numbers against your local MySQL database.")

# --- SIDEBAR: Database & File Settings ---
with st.sidebar:
    st.header("🗄️ MySQL Connection Settings")
    # Using 127.0.0.1 forces the OS loopback layer directly instead of relying on local domain files
    db_host = st.text_input("DB Host:", value="127.0.0.1", help="Use 127.0.0.1 for local instances.")
    db_port = st.text_input("DB Port:", value="3306")
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

# --- ROBUST CONNECTION BUILDER ---
def get_clean_engine(user, password, host, port, dbname):
    """
    Builds a connection string explicitly forcing standard TCP/IP loopback,
    avoiding both IPv6 resolution failures and missing local socket file bugs.
    """
    safe_password = urllib.parse.quote_plus(password)
    connection_str = f"mysql+pymysql://{user}:{safe_password}@{host}:{port}/{dbname}"
    
    return create_engine(
        connection_str, 
        connect_args={
            "connect_timeout": 5,
            "use_pure": True # Forces PyMySQL to handle pure Python socket allocation
        }
    )

def test_db_connection():
    try:
        engine = get_clean_engine(db_user, db_pass, db_host, db_port, db_name)
        with engine.connect() as conn:
            conn.execute(text("SELECT 1"))
        return True, "Connection successful! Network pipeline established."
    except Exception as e:
        return False, str(e)

# --- UI WORKFLOW ---
st.markdown("### 🔍 Step 1: Database Status Verification")

if st.button("⚡ Test DB Connection"):
    with st.spinner("Verifying database credentials..."):
        is_valid, msg = test_db_connection()
        if is_valid:
            st.success(msg)
            st.session_state["db_verified"] = True
        else:
            st.error(f"Database Connection Failed:\n\n`{msg}`\n\n💡 Tip: Verify your host is '127.0.0.1', your password is exact, and your local firewall allows connection lookups.")
            st.session_state["db_verified"] = False

# Only unlock the workflow when the database check returns True
if st.session_state.get("db_verified", False):
    st.markdown("---")
    st.markdown("### 📥 Step 2: Upload Data File")
    
    def robust_load_file(uploaded_file, password=None):
        """Safely decompresses and parses files from different formats and versions."""
        file_bytes = uploaded_file.read()
        file_stream = io.BytesIO(file_bytes)
        
        try:
            office_file = msoffcrypto.OfficeFile(file_stream)
            if office_file.is_encrypted():
                decrypted_stream = io.BytesIO()
                office_file.load_key(password=password if password else "VelvetSweatshop")
                office_file.decrypt(decrypted_stream)
                decrypted_stream.seek(0)
                file_stream = decrypted_stream
        except Exception:
            file_stream.seek(0)

        try: return pd.read_excel(file_stream, engine='openpyxl')
        except Exception: file_stream.seek(0)
            
        try: return pd.read_excel(file_stream, engine='xlrd')
        except Exception: file_stream.seek(0)
            
        try:
            text_data = file_stream.read().decode('utf-8', errors='ignore')
            return pd.read_csv(io.StringIO(text_data))
        except Exception:
            file_stream.seek(0)
            
        try:
            text_data = file_stream.read().decode('utf-8', errors='ignore')
            return pd.read_csv(io.StringIO(text_data), sep='\t')
        except Exception:
            raise ValueError("Unsupported or heavily corrupted file format.")

    uploaded_file = st.file_uploader(
        "Choose a data file to process", 
        type=["xlsx", "xls", "csv", "tsv", "txt"]
    )

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
                    
            selected_phone_col = st.selectbox(
                "Select the 'Phone Number' column from the uploaded file:", 
                options=all_columns,
                index=default_idx
            )
            
            if st.button("🚀 Run Database Suppression"):
                with st.spinner("Processing file data and pulling database indexes..."):
                    # Cast column values to a standardized string baseline (clears spaces, floating decimals)
                    df[selected_phone_col] = df[selected_phone_col].astype(str).str.replace(r'\s+|\.0$', '', regex=True).str.strip()
                    
                    # Run target validation deduction step internally first
                    df_internal_clean = df.drop_duplicates(subset=[selected_phone_col], keep="first")
                    internal_dupes = len(df) - len(df_internal_clean)
                    
                    # Connect via network configuration logic
                    engine = get_clean_engine(db_user, db_pass, db_host, db_port, db_name)
                    
                    # Stream just the single designated column into memory
                    query = f"SELECT `{db_phone_col}` FROM `{db_table}`"
                    db_phones_df = pd.read_sql(query, con=engine)
                    
                    # Transform search list into an instantaneous lookup hash set
                    db_phones_set = set(db_phones_df[db_phone_col].astype(str).str.replace(r'\s+|\.0$', '', regex=True).str.strip())
                    
                    # Drop existing values
                    df_final = df_internal_clean[~df_internal_clean[selected_phone_col].isin(db_phones_set)]
                    db_suppressed_count = len(df_internal_clean) - len(df_final)
                    
                    # --- Metrics Breakdown Output ---
                    st.markdown("### 📊 Metrics Summary")
                    col1, col2, col3 = st.columns(3)
                    col1.metric("Internal File Duplicates", f"{internal_dupes}")
                    col2.metric("Matched in DB", f"{db_suppressed_count}")
                    col3.metric("Net New Unique Records", f"{len(df_final)}")
                    
                    if len(df_final) == 0:
                        st.warning("All records in this file already exist within your database!")
                    else:
                        buffer = io.BytesIO()
                        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                            df_final.to_excel(writer, index=False, sheet_name='New Uniques Only')
                        
                        st.download_button(
                            label="📥 Download Clean & Suppressed Excel File",
                            data=buffer.getvalue(),
                            file_name=f"suppressed_{uploaded_file.name if '.' in uploaded_file.name else 'data.xlsx'}",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        
        except Exception as e:
            st.error(f"Error Processing Pipeline: {e}")
else:
    st.info("⚠️ Please verify your MySQL connection using the 'Test DB Connection' button to unlock the file processor workflow.")
