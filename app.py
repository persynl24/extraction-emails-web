import streamlit as st
import pandas as pd
import extract_msg
import io
import re
from datetime import datetime

# --- PAGE CONFIG ---
st.set_page_config(page_title="AI4IAM - Special Permissions", page_icon="📧", layout="wide")

# --- CUSTOM STYLES ---
st.markdown("""
    <style>
    body {
        background-color: #f5f9ff;
        color: #0d1b2a;
    }
    .stButton>button {
        background-color: #1d4ed8;
        color: white;
        border-radius: 8px;
        padding: 0.5rem 1rem;
        font-weight: 500;
    }
    .stButton>button:hover {
        background-color: #2563eb;
    }
    .blue-container {
        background-color: #e8f0ff;
        padding: 1rem;
        border-radius: 10px;
        border: 1px solid #cddbf7;
        margin-bottom: 1rem;
    }
    </style>
""", unsafe_allow_html=True)

# --- HEADER WITH TOP-RIGHT CREDIT ---
col1, col2 = st.columns([9, 1])
with col1:
    st.title("📧 AI4IAM - Special Permissions")
with col2:
    st.markdown("**Created by Loïc Persyn with the help of AI**")

st.markdown("---")

# --- SESSION STATE ---
if "uploaded_files" not in st.session_state:
    st.session_state.uploaded_files = None
if "current_step" not in st.session_state:
    st.session_state.current_step = 1
if "data" not in st.session_state:
    st.session_state.data = None

# --- HELPER FUNCTION ---
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

# --- STEP 1: UPLOAD FILES ---
step1_expanded = st.session_state.current_step == 1
with st.expander("🗂 Step 1: Upload Files", expanded=step1_expanded):
    with st.container():
        st.markdown('<div class="blue-container">', unsafe_allow_html=True)
        uploaded_files = st.file_uploader(
            "Select your .msg files",
            type=['msg'],
            accept_multiple_files=True,
            label_visibility="collapsed"
        )

        if uploaded_files:
            st.session_state.uploaded_files = uploaded_files
            st.success(f"✅ {len(uploaded_files)} file(s) uploaded")
        else:
            st.info("👈 Upload one or more .msg files to begin")

        # Clear all button
        if st.button("🧹 Clear all"):
            st.session_state.uploaded_files = None
            st.session_state.data = None
            st.session_state.current_step = 1
            st.experimental_rerun()
        st.markdown('</div>', unsafe_allow_html=True)

# --- STEP 2: EXTRACT & VALIDATE ---
step2_expanded = st.session_state.current_step == 2
if st.session_state.uploaded_files:
    with st.expander("⚙️ Step 2: Extract & Validate", expanded=step2_expanded):
        with st.container():
            st.markdown('<div class="blue-container">', unsafe_allow_html=True)
            if st.button("🚀 EXTRACT DATA"):
                data = []
                progress_bar = st.progress(0)
                status_text = st.empty()

                for i, uploaded_file in enumerate(st.session_state.uploaded_files):
                    try:
                        msg = extract_msg.Message(io.BytesIO(uploaded_file.read()))
                        info = extract_info(msg.body)
                        data.append(info)
                        msg.close()
                        progress_bar.progress((i + 1) / len(st.session_state.uploaded_files))
                        status_text.text(f"Processing: {i+1}/{len(st.session_state.uploaded_files)}")
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

                    df.insert(df.columns.get_loc('Link'), 'Needs Extension ? [y/n]', "")
                    st.session_state.data = df
                    st.session_state.current_step = 3  # Move to step 3 automatically
                    st.success(f"✅ Extraction complete! {len(data)} email(s) processed.")
                else:
                    st.error("❌ No data extracted")
            st.markdown('</div>', unsafe_allow_html=True)

# --- STEP 3: DOWNLOAD RESULTS ---
step3_expanded = st.session_state.current_step == 3
if st.session_state.data is not None:
    with st.expander("💾 Step 3: Download Results", expanded=step3_expanded):
        with st.container():
            st.markdown('<div class="blue-container">', unsafe_allow_html=True)
            st.markdown("### 🎯 Export your results (Excel only)")

            df_excel = st.session_state.data.copy()
            df_excel['Link'] = df_excel['Link'].apply(
                lambda x: f'=HYPERLINK("{x}", "{x}")' if isinstance(x, str) and x.startswith("http") else x
            )

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_excel.to_excel(writer, index=False)
                worksheet = writer.sheets['Sheet1']
                for column_cells in worksheet.columns:
                    max_length = max(len(str(cell.value)) if cell.value else 0 for cell in column_cells)
                    worksheet.column_dimensions[column_cells[0].column_letter].width = max_length + 2
            excel_data = output.getvalue()

            st.download_button(
                label="📘 Download Excel – for manual review",
                data=excel_data,
                file_name=f"PRIAM - Special Permissions - {datetime.now().strftime('%Y-%m-%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.markdown('</div>', unsafe_allow_html=True)

# --- INSTRUCTIONS WHEN NO FILES ---
if st.session_state.uploaded_files is None:
    st.markdown("---")
    st.subheader("📖 Instructions")
    st.markdown("""
    1. Expand **Step 1** and upload one or more `.msg` email files  
    2. Click **"🚀 EXTRACT DATA"** under Step 2  
    3. Download the results in Excel format under Step 3  
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
