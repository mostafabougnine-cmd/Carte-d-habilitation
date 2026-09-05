# -*- coding: utf-8 -*-
import os
import io
import streamlit as st
import openpyxl
from openpyxl.drawing.image import Image as XLImage

st.set_page_config(page_title="Générateur de Cartes d'Habilitation - ONCF", layout="centered")

# خريطة الخانات الموحدة
CELL_MAPPING = {
    "nom": "F5",
    "prenom": "J5",
    "matricule": "F6",
    "centre": "F7",
    "antenne": "J7",
    "date_aut": "F8",
    "date_prof": "F9",
    "date_med": "F10",
    "date_psy": "F11",
    "photo_cell": "B5"
}

# قائمة الملفات في مجلد data/
TEMPLATES = {
    "CFT (Chef Formation Trains)": "CFT.xlsx",
    "CL (Conducteur de Ligne)": "CL.xlsx",
    "CTR (Chef de Train)": "CTR.xlsx",
    "CRMV (Conducteur de Manœuvre)": "CRMV.xlsx"
}

def generate_card(template_path, data, photo_bytes):
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    # تعبئة الخانات الفارغة فقط من دون المساس بالعبارات المسجلة سلفاً
    for key, cell in CELL_MAPPING.items():
        if key in data and key != "photo_cell":
            user_value = data[key]
            current_cell_value = ws[cell].value
            
            # الكتابة فقط إذا كانت الخانة فارغة أو تحتوي على مسافات فقط
            if (current_cell_value is None or str(current_cell_value).strip() == "") and user_value and str(user_value).strip() != "":
                ws[cell] = user_value

    # إضافة الصورة الشخصية
    if photo_bytes:
        img_file = io.BytesIO(photo_bytes)
        img = XLImage(img_file)
        img.width = 110
        img.height = 125
        ws.add_image(img, CELL_MAPPING["photo_cell"])

    output_stream = io.BytesIO()
    wb.save(output_stream)
    output_stream.seek(0)
    return output_stream.getvalue()

st.title("🎴 Générateur de Cartes d'Habilitation")

# اختيار نوع البطاقة
selected_label = st.selectbox("Choisissez le modèle de carte :", list(TEMPLATES.keys()))
template_filename = TEMPLATES[selected_label]

# رفع الصورة الشخصية
uploaded_photo = st.file_uploader("Photo d'identité (JPG / PNG)", type=["jpg", "jpeg", "png"])

# استمارة إدخال البيانات
with st.form("agent_form"):
    st.subheader("Informations de l'Agent")
    col1, col2 = st.columns(2)
    with col1:
        nom = st.text_input("Nom", "")
        matricule = st.text_input("Matricule", "")
        centre = st.text_input("Centre (laisser vide si déjà rempli)", "")
        date_aut = st.text_input("Date d'autorisation", "")
        date_med = st.text_input("Date examen médical", "")
    with col2:
        prenom = st.text_input("Prénom", "")
        antenne = st.text_input("Antenne (laisser vide si déjà rempli)", "")
        date_prof = st.text_input("Date examen professionnel", "")
        date_psy = st.text_input("Date examen psychotechnique", "")

    submit = st.form_submit_button("⚡ Générer la Carte")

if submit:
    template_path = os.path.join("data", template_filename)

    if not os.path.exists(template_path):
        st.error(f"⚠️ Fichier introuvable : '{template_path}'. Assurez-vous d'avoir placé le fichier dans le dossier 'data/'.")
    else:
        photo_bytes = uploaded_photo.read() if uploaded_photo else None
        data = {
            'nom': nom, 'prenom': prenom, 'matricule': matricule,
            'centre': centre, 'antenne': antenne,
            'date_aut': date_aut, 'date_prof': date_prof,
            'date_med': date_med, 'date_psy': date_psy
        }

        try:
            excel_out = generate_card(template_path, data, photo_bytes)
            st.success(f"La carte {selected_label.split(' ')[0]} a été générée avec succès ! 🎉")
            st.download_button(
                label="📥 Télécharger la Carte (Excel)",
                data=excel_out,
                file_name=f"Carte_{selected_label.split(' ')[0]}_{nom if nom else 'Agent'}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Erreur lors de la génération : {e}")
