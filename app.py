import streamlit as st
import pandas as pd
import extract_msg
import io
import re
from datetime import datetime
from openpyxl.utils import get_column_letter

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="AI4IAM - Special Permissions", page_icon="📧", layout="wide")

# --- HEADER ---
col1, col2 = st.columns([9,1])
with col1:
    st.title("📧 AI4IAM - Special Permissions")
with col2:
    st.markdown("""
    <div style="white-space: nowrap; text-align: right; font-weight: bold;">
    Created by Loïc Persyn with the help of AI
    </div>
    """, unsafe_allow_html=True)

st.markdown("---")

# --- FUNCTION TO EXTRACT INFO FROM EMAIL BODY ---
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

    match_user = re.search(r'User:\s*([A-Z0-9]+)\s*/\s*([^\r\n]+)', text, re.IGNORECASE)
    if match_user:
        info['User_ID'] = match_user.group(1).strip()
        info['User_Name'] = match_user.group(2).strip()

    match_manager = re.search(r'Manager:\s*([^\r\n]+)', text, re.IGNORECASE)
    if match_manager:
        info['Manager'] = match_manager.group(1).strip()

    match_permission = re.search(r'Special permission:\s*([^\(]+)\((\d+)\)', text, re.IGNORECASE)
    if match_permission:
        info['Special_Permission'] = match_permission.group(1).strip()
        info['Permission_Code'] = match_permission.group(2).strip()

    match_date = re.search(r'Planned End date:\s*(\d{2}[-/]\d{2}[-/]\d{4})', text, re.IGNORECASE)
    if match_date:
        info['End_Date'] = match_date.group(1).strip()

    match_link = re.search(r'Link:\s*<?([^>\r\n]+)>?', text, re.IGNORECASE)
    if match_link:
        link_text = match_link.group(1).strip()
        url_match = re.search(r'(https?://[^\s>]+)', link_text)
        if url_match:
            link_text = url_match.group(1).strip()
        info['Link'] = link_text

    return info

# --- SIDEBAR FILE UPLOAD WITH LOGO ---
st.sidebar.image("logo.png", width=240)  # <-- Votre logo ici
st.sidebar.header("📁 Upload Files")
uploaded_files = st.sidebar.file_uploader(
    "Select your .msg files",
    type=['msg'],
    accept_multiple_files=True
)

# --- MAIN LOGIC ---
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
            # --- CREATE DATAFRAME ---
            df = pd.DataFrame(data)
            df.insert(6, 'Needs Extension ? [y/n]', "")  # Insert empty column before Link

            column_order = [
                'User_ID', 'User_Name', 'Manager',
                'Special_Permission', 'Permission_Code',
                'End_Date', 'Needs Extension ? [y/n]', 'Link'
            ]
            df = df[column_order]

            # --- DISPLAY RESULTS ---
            st.markdown("---")
            st.subheader("📊 Extracted Data")
            st.dataframe(df, use_container_width=True)

            # --- DOWNLOAD RESULTS ---
            st.markdown("---")
            st.subheader("📥 Download Results")

            df_excel = df.copy()
            df_excel['Link'] = df_excel['Link'].apply(
                lambda x: f'=HYPERLINK("{x}", "{x}")' if isinstance(x, str) and x.startswith("http") else x
            )

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_excel.to_excel(writer, index=False)
                worksheet = writer.sheets['Sheet1']

                # Auto-adjust column width
                for col_idx, col in enumerate(df_excel.columns, 1):
                    max_length = max(df_excel[col].astype(str).map(len).max(), len(col))
                    worksheet.column_dimensions[get_column_letter(col_idx)].width = max_length + 2

            excel_data = output.getvalue()

            st.download_button(
                label="📘 Download Excel – for manual review",
                data=excel_data,
                file_name=f"PRIAM - Special Permissions - {datetime.now().strftime('%Y-%m-%d')}.xlsx",
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
    1. Upload one or more `.msg` email files using the sidebar.  
    2. Click **"🚀 EXTRACT DATA"** to process the emails.  
    3. Download the results in Excel format using the button provided.  
    4. Review the "Needs Extension ? [y/n]" column if applicable.
    """)

    st.markdown("---")
    st.subheader("📋 Example Email Format (Anonymized)")
    st.code("""
Dear Approver, 
Please evaluate the special user permission below: 
User: ABC123 / John Doe
Manager: Jane Smith
Special permission: SP- Partial installation permission - Yes (XXXX)
Planned End date: 01-01-2025
Link: https://company.com/permissions?id=XXXX
Regards, 
PRIAM
    """)
