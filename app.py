# -*- coding: utf-8 -*-
import os
import io
import streamlit as st
import openpyxl
from openpyxl.drawing.image import Image as XLImage

st.set_page_config(page_title="Générateur Carte Habilitation CFT", layout="centered")

# تعيين الخانات الخاصة بنموذج CFT
CELL_MAPPING = {
    "nom": "C8",
    "prenom": "F8",
    "matricule": "C10",
    "centre": "C12",
    "antenne": "F12",
    "date_aut": "C14",
    "date_prof": "C16",
    "date_med": "C18",
    "date_psy": "C20",
    "engins": "H8",
    "sites": "J8",
    "photo_cell": "A8"
}

def fill_excel_template(template_file, cell_mapping, data, photo_bytes):
    wb = openpyxl.load_workbook(template_file)
    ws = wb.active

    # 1. كتابة المعطيات فـ الخانات
    for key, cell in cell_mapping.items():
        if key in data and key != "photo_cell":
            ws[cell] = data[key]

    # 2. إضافة الصورة الشخصية
    if photo_bytes and "photo_cell" in cell_mapping:
        img_file = io.BytesIO(photo_bytes)
        img = XLImage(img_file)
        img.width = 110
        img.height = 130
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
    st.subheader("3. Informations Personnelles & Dates")
    col1, col2 = st.columns(2)
    with col1:
        nom = st.text_input("Nom", "AIT BAHALI")
        matricule = st.text_input("Matricule", "47607A")
        centre = st.text_input("Centre", "CCFTC Kénitra")
        date_aut = st.text_input("Date d'autorisation", "01/03/2021")
        date_med = st.text_input("Date examen médical", "02/03/2023")
    with col2:
        prenom = st.text_input("Prénom", "BRAHIM")
        antenne = st.text_input("Antenne", "ACFTC Kénitra")
        date_prof = st.text_input("Date examen professionnel", "02/09/2026")
        date_psy = st.text_input("Date examen psychotechnique", "03/04/2024")

    st.subheader("4. Informations Complémentaires")
    engins = st.text_area("Engins autorisés", "E1450, E1400, Z2M")
    sites = st.text_area("Sites autorisés", "Site Voyageurs Kénitra")

    submit = st.form_submit_button("⚡ Générer la Carte Excel")

if submit:
    # استعمال الملف المرفوع أو البحث عن ملف افتراضي CFT.xlsx
    template_to_use = uploaded_template if uploaded_template is not None else "CFT.xlsx"

    if isinstance(template_to_use, str) and not os.path.exists(template_to_use):
        st.error("⚠️ Veuillez charger votre fichier modèle Excel (CFT) ci-dessus.")
    else:
        photo_bytes = uploaded_photo.read() if uploaded_photo else None
        data = {
            'nom': nom, 'prenom': prenom, 'matricule': matricule,
            'centre': centre, 'antenne': antenne,
            'date_aut': date_aut, 'date_prof': date_prof,
            'date_med': date_med, 'date_psy': date_psy,
            'engins': engins, 'sites': sites
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
