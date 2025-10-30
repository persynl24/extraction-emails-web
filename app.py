import streamlit as st
import pandas as pd
import extract_msg
import io
import re
from datetime import datetime

st.set_page_config(page_title="Extracteur Emails .msg", page_icon="📧", layout="wide")

st.title("📧 Extracteur de données depuis fichiers .msg")
st.markdown("---")

def extraire_infos(corps_email):
    infos = {
        'User_ID': None,
        'Nom_Utilisateur': None,
        'Manager': None,
        'Permission_Speciale': None,
        'Code_Permission': None,
        'Date_Fin': None,
        'Lien': None
    }

    # Nettoyage de base
    texte = corps_email.replace("\r", "").strip()

    # User et Nom : capture jusqu'à la fin de la ligne
    match_user = re.search(r'User:\s*([A-Z0-9]+)\s*/\s*([^\r\n]+)', texte, re.IGNORECASE)
    if match_user:
        infos['User_ID'] = match_user.group(1).strip()
        infos['Nom_Utilisateur'] = match_user.group(2).strip()

    # Manager : capture jusqu'à la fin de la ligne
    match_manager = re.search(r'Manager:\s*([^\r\n]+)', texte, re.IGNORECASE)
    if match_manager:
        infos['Manager'] = match_manager.group(1).strip()

    # Permission spéciale et code (tout avant et dans parenthèses)
    match_permission = re.search(r'Special permission:\s*([^\(]+)\((\d+)\)', texte, re.IGNORECASE)
    if match_permission:
        infos['Permission_Speciale'] = match_permission.group(1).strip()
        infos['Code_Permission'] = match_permission.group(2).strip()

    # Date fin
    match_date = re.search(r'Planned End date:\s*(\d{2}[-/]\d{2}[-/]\d{4})', texte, re.IGNORECASE)
    if match_date:
        infos['Date_Fin'] = match_date.group(1).strip()

    # Lien propre (avec ou sans chevrons)
    match_lien = re.search(r'Link:\s*<?(https?://[^\s>]+)>?', texte, re.IGNORECASE)
    if match_lien:
        infos['Lien'] = match_lien.group(1).strip()

    return infos


# --- Interface Streamlit ---
st.sidebar.header("📁 Upload des fichiers")
uploaded_files = st.sidebar.file_uploader(
    "Sélectionnez vos fichiers .msg",
    type=['msg'],
    accept_multiple_files=True
)

if uploaded_files:
    st.success(f"✅ {len(uploaded_files)} fichier(s) uploadé(s)")

    if st.button("🚀 EXTRAIRE LES DONNÉES", type="primary"):
        donnees = []
        progress_bar = st.progress(0)
        status_text = st.empty()

        for i, uploaded_file in enumerate(uploaded_files):
            try:
                msg = extract_msg.Message(io.BytesIO(uploaded_file.read()))
                infos = extraire_infos(msg.body)
                donnees.append(infos)
                msg.close()

                progress_bar.progress((i + 1) / len(uploaded_files))
                status_text.text(f"Traitement : {i+1}/{len(uploaded_files)}")

            except Exception as e:
                st.warning(f"⚠️ Erreur avec {uploaded_file.name}: {str(e)}")

        if donnees:
            df = pd.DataFrame(donnees)
            colonnes_ordre = [
                'User_ID', 'Nom_Utilisateur', 'Manager',
                'Permission_Speciale', 'Code_Permission',
                'Date_Fin', 'Lien'
            ]
            df = df[colonnes_ordre]

            st.markdown("---")
            st.subheader("📊 Résultats")
            st.dataframe(df, use_container_width=True)

            st.markdown("---")
            st.subheader("📥 Téléchargement")

            col1, col2 = st.columns(2)

            # --- CSV : lien texte simple ---
            with col1:
                csv = df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📄 Télécharger CSV",
                    data=csv,
                    file_name=f"permissions_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv"
                )

            # --- Excel : lien cliquable ---
            with col2:
                df_excel = df.copy()
                df_excel['Lien'] = df_excel['Lien'].apply(
                    lambda x: f'=HYPERLINK("{x}", "{x}")' if isinstance(x, str) and x.startswith("http") else x
                )

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_excel.to_excel(writer, index=False)
                excel_data = output.getvalue()

                st.download_button(
                    label="📊 Télécharger Excel (liens cliquables)",
                    data=excel_data,
                    file_name=f"permissions_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            st.success(f"✅ Extraction terminée ! {len(donnees)} emails traités.")
        else:
            st.error("❌ Aucune donnée extraite")

else:
    st.info("👈 Uploadez vos fichiers .msg dans la barre latérale pour commencer")

    st.markdown("---")
    st.subheader("📖 Instructions")
    st.markdown("""
    1. Cliquez sur **"Browse files"** dans la barre latérale  
    2. Sélectionnez un ou plusieurs fichiers .msg  
    3. Cliquez sur **"EXTRAIRE LES DONNÉES"**  
    4. Téléchargez le résultat en CSV ou Excel  
    """)

    st.markdown("---")
    st.subheader("📋 Exemple de format d'email")
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
