import streamlit as st
import pandas as pd
import extract_msg
import io
import re
from datetime import datetime

st.set_page_config(page_title="Email .msg Data Extractor", page_icon="📧", layout="wide")

st.title("📧 Email Data Extractor (.msg files)")
st.markdown("---")

def extract_info(email_body):
    info = {
        'User_ID': None,
        'User_Name': None,
        'Manager': None,
        'Special_Permission': None,
        'Permission_Code': None,
        'End_Date': None,
        'Link': None
    }

    text = email_body.replace("\r", "").strip()

    # User ID and Name
    match_user = re.search(r'User:\s*([A-Z0-9]+)\s*/\s*([^\r\n]+)', text, re.IGNORECASE)
    if match_user:
        info['User_ID'] = match_user.group(1).strip()
        info['User_Name'] = match_user.group(2).strip()

    # Manager
    match_manager = re.search(r'Manager:\s*([^\r\n]+)', text, re.IGNORECASE)
    if match_manager:
        info['Manager'] = match_manager.group(1).strip()

    # Special permission and code
    match_permission = re.search(r'Special permission:\s*([^\(]+)\((\d+)\)', text, re.IGNORECASE)
    if match_permission:
        info['Special_Permission'] = match_permission.group(1).strip()
        info['Permission_Code'] = match_permission.group(2).strip()

    # End date
    match_date = re.search(r'Planned End date:\s*(\d{2}[-/]\d{2}[-/]\d{4})', text, re.IGNORECASE)
    if match_date:
        info['End_Date'] = match_date.group(1).strip()

    # Link (URL or plain text)
    match_link = re.search(r'Link:\s*<?([^>\r\n]+)>?', text, re.IGNORECASE)
    if match_link:
        link = match_link.group(1).strip()
        link = re.sub(r'[<>]', '', link).strip()
        info['Link'] = link

    return info


# --- Streamlit UI ---
st.sidebar.header("📁 Upload Files")
uploaded_files = st.sidebar.file_uploader(
    "Select your .msg files",
    type=['msg'],
    accept_multiple_files=True
)

if uploaded_files:
    st.success(f"✅ {len(uploaded_files)} file(s) uploaded")

    if st.button("🚀 EXTRACT DATA", type="primary"):
        data = []
        progress_bar = st.progress(0)
        status_text = st.empty()

        for i, uploaded_file in enumerate(uploaded_files):
            try:
                msg = extract_msg.Message(io.BytesIO(uploaded_file.read()))
                info = extract_info(msg.body)
                data.append(info)
                msg.close()

                progress_bar.progress((i + 1) / len(uploaded_files))
                status_text.text(f"Processing: {i+1}/{len(uploaded_files)}")

            except Exception as e:
                st.warning(f"⚠️ Error with {uploaded_file.name}: {str(e)}")

        if data:
            df = pd.DataFrame(data)
            column_order = [
                'User_ID', 'User_Name', 'Manager',
                'Special_Permission', 'Permission_Code',
                'End_Date', 'Link'
            ]
            df = df[column_order]

            st.markdown("---")
            st.subheader("📊 Extracted Data")
            st.dataframe(df, use_container_width=True)

            st.markdown("---")
            st.subheader("📥 Download Results")

            col1, col2 = st.columns(2)

            # --- CSV (plain text link) ---
            with col1:
                csv = df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📄 Download CSV",
                    data=csv,
                    file_name=f"permissions_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv"
                )

            # --- Excel (clickable links) ---
            with col2:
                df_excel = df.copy()
                df_excel['Link'] = df_excel['Link'].apply(
                    lambda x: f'=HYPERLINK("{x}", "{x}")' if isinstance(x, str) and x.startswith("http") else x
                )

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_excel.to_excel(writer, index=False)
                excel_data = output.getvalue()

                st.download_button(
                    label="📊 Download Excel (clickable links)",
                    data=excel_data,
                    file_name=f"permissions_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            st.success(f"✅ Extraction complete! {len(data)} email(s) processed.")
        else:
            st.error("❌ No data extracted")

else:
    st.info("👈 Upload your .msg files using the sidebar to get started.")

    st.markdown("---")
    st.subheader("📖 Instructions")
    st.markdown("""
    1. Click **"Browse files"** in the sidebar  
    2. Select one or more `.msg` email files  
    3. Click **"EXTRACT DATA"**  
    4. Download the results in CSV or Excel format  
    """)

    st.markdown("---")
    st.subheader("📋 Example Email Format")
    st.code("""
Dear Approver, 
Please evaluate the special user permission below: 
User: Z99SKM / SAMPADA KUMARI
Manager: GEOFFROY DE PUYT 
Special permission: SP- Partial installation permission - Yes (7979) 
Planned End date: 22-11-2025 
Link: https://intranet.company.com/permissions?id=7979
Regards, 
PRIAM
    """)
