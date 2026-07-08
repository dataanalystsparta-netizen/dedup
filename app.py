import streamlit as st
import pandas as pd

st.set_page_config(page_title="Excel Deduplicator", layout="centered")
st.title("📊 Excel File Cleaner")
st.write("Upload your Excel file below to instantly remove duplicate records.")

# File Uploader Widget
uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # Read Excel Data
        df = pd.read_excel(uploaded_file)
        st.success(f"Successfully loaded file with {len(df)} rows!")
        
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
        
        # Convert back to Excel format in-memory for download
        import io
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_cleaned.to_excel(writer, index=False, sheet_name='Cleaned Data')
        
        # Download Button
        st.download_button(
            label="📥 Download Cleaned Excel File",
            data=buffer.getvalue(),
            file_name=f"deduped_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"Error processing file: {e}")
