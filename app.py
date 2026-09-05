# -*- coding: utf-8 -*-
import os
import io
import streamlit as st
import openpyxl
from openpyxl.drawing.image import Image as XLImage

st.set_page_config(page_title="Générateur Carte Habilitation CFT", layout="centered")

# الخانات الحقيقية والمضبوطة فـ ملف CFT.xlsx
CELL_MAPPING = {
    "nom": "D8",
    "prenom": "F8",
    "matricule": "D10",
    "centre": "D12",
    "antenne": "F12",
    "date_aut": "D14",
    "date_prof": "D16",
    "date_med": "D18",
    "date_psy": "D20",
    "photo_cell": "A8"
}

def fill_excel_template(template_file, cell_mapping, data, photo_bytes):
    wb = openpyxl.load_workbook(template_file)
    ws = wb.active

    # 1. كتابة المعطيات فـ الخانات الصحيحة
    for key, cell in cell_mapping.items():
        if key in data and key != "photo_cell":
            ws[cell] = data[key]

    # 2. ضبط أبعاد وتطبيق التصويرة فـ الخانة A8
    if photo_bytes and "photo_cell" in cell_mapping:
        img_file = io.BytesIO(photo_bytes)
        img = XLImage(img_file)
        # العبار المقاد للتصويرة باش تجي فـ المربع تماماً
        img.width = 90
        img.height = 110
        ws.add_image(img, cell_mapping["photo_cell"])

    output_stream = io.BytesIO()
    wb.save(output_stream)
    output_stream.seek(0)
    return output_stream.getvalue()

st.title("🎴 Générateur Carte Habilitation CFT")

st.subheader("1. Modèle Excel CFT")
uploaded_template = st.file_uploader("Chargez le fichier modèle Excel (ex: CFT.xlsx)", type=["xlsx"])

st.subheader("2. Photo d'identité")
uploaded_photo = st.file_uploader("Photo de l'Agent (JPG / PNG)", type=["jpg", "jpeg", "png"])

with st.form("card_form"):
    st.subheader("3. Informations Personnelles")
    col1, col2 = st.columns(2)
    with col1:
        nom = st.text_input("Nom", "AIT BAHALI")
        matricule = st.text_input("Matricule", "47607A")
        centre = st.text_input("Centre", "CCFTC Kénitra")
    with col2:
        prenom = st.text_input("Prénom", "BRAHIM")
        antenne = st.text_input("Antenne", "ACFTC Kénitra")

    st.subheader("4. Dates")
    col3, col4 = st.columns(2)
    with col3:
        date_aut = st.text_input("Date d'autorisation", "01/03/2021")
        date_med = st.text_input("Date examen médical", "02/03/2023")
    with col4:
        date_prof = st.text_input("Date examen professionnel", "02/09/2026")
        date_psy = st.text_input("Date examen psychotechnique", "03/04/2024")

    submit = st.form_submit_button("⚡ Générer la Carte Excel")

if submit:
    # استعمال CFT.xlsx الموجود فـ data/CFT.xlsx أو المرفوع
    template_to_use = uploaded_template if uploaded_template is not None else "data/CFT.xlsx"

    if isinstance(template_to_use, str) and not os.path.exists(template_to_use):
        st.error("⚠️ Le fichier 'data/CFT.xlsx' est introuvable. Veuillez le charger ci-dessus.")
    else:
        photo_bytes = uploaded_photo.read() if uploaded_photo else None
        data = {
            'nom': nom, 'prenom': prenom, 'matricule': matricule,
            'centre': centre, 'antenne': antenne,
            'date_aut': date_aut, 'date_prof': date_prof,
            'date_med': date_med, 'date_psy': date_psy
        }

        try:
            excel_out = fill_excel_template(template_to_use, CELL_MAPPING, data, photo_bytes)
            st.success("Carte générée avec succès ! 🎉")
            st.download_button(
                label="📥 Télécharger la Carte (Excel .xlsx)",
                data=excel_out,
                file_name=f"Carte_CFT_{nom}_{matricule}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Erreur lors de la génération: {e}")
