# -*- coding: utf-8 -*-
import os
import io
import streamlit as st
import openpyxl
from openpyxl.drawing.image import Image as XLImage
import subprocess  # لتحويل Excel إلى PDF عبر LibreOffice إذا كان متوفراً

st.set_page_config(page_title="Générateur Carte Habilitation ONCF", layout="centered")

# خريطة ملفات الموديل الخانات المخصصة لكل معلومة f-Excel
TEMPLATE_CONFIG = {
    "Chef Formation Trains": {
        "template": "data/template_cft.xlsx",
        "cells": {
            "nom": "C8",
            "prenom": "F8",
            "matricule": "C10",
            "date_aut": "C14",
            "date_prof": "C16",
            "date_med": "C18",
            "date_psy": "C20",
            "photo_cell": "A8"
        }
    },
    "Conducteur de Ligne": {
        "template": "data/template_cl.xlsx",
        "cells": {
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
    }
}

def fill_excel_template(template_path, cell_mapping, data, photo_bytes):
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    # 1. كتابة المعطيات فـ الخانات المحددة بالضبط
    for key, cell in cell_mapping.items():
        if key in data and key != "photo_cell":
            ws[cell] = data[key]

    # 2. إدراج الصورة الشخصية فـ الخانة المخصصة ليها
    if photo_bytes and "photo_cell" in cell_mapping:
        img_file = io.BytesIO(photo_bytes)
        img = XLImage(img_file)
        # تحديد عرض وارتفاع الصورة داخل Excel
        img.width = 110
        img.height = 130
        ws.add_image(img, cell_mapping["photo_cell"])

    output_stream = io.BytesIO()
    wb.save(output_stream)
    output_stream.seek(0)
    return output_stream.getvalue()

# Streamlit UI
st.title("🎴 Générateur Carte Habilitation ONCF (Via Excel)")

selected_fonction = st.selectbox("Choix de la Fonction", list(TEMPLATE_CONFIG.keys()))
config = TEMPLATE_CONFIG[selected_fonction]

uploaded_photo = st.file_uploader("Photo de l'Agent (JPG / PNG)", type=["jpg", "jpeg", "png"])

with st.form("card_form"):
    col1, col2 = st.columns(2)
    with col1:
        nom = st.text_input("Nom", "AIT BAHALI")
        matricule = st.text_input("Matricule", "47607A")
        centre = st.text_input("Centre", "CCFTC Kénitra")
    with col2:
        prenom = st.text_input("Prénom", "BRAHIM")
        antenne = st.text_input("Antenne", "ACFTC Kénitra")

    col3, col4 = st.columns(2)
    with col3:
        date_aut = st.text_input("Date d'autorisation", "01/03/2021")
        date_med = st.text_input("Date examen médical", "02/03/2023")
    with col4:
        date_prof = st.text_input("Date examen professionnel", "02/09/2026")
        date_psy = st.text_input("Date examen psychotechnique", "03/04/2024")

    engins = st.text_area("Engins autorisés", "E1450, E1400, Z2M")
    sites = st.text_area("Sites autorisés", "Site Voyageurs Kénitra")

    submit = st.form_submit_button("⚡ Générer le fichier Excel rempli")

if submit:
    photo_bytes = uploaded_photo.read() if uploaded_photo else None

    data = {
        'nom': nom, 'prenom': prenom, 'matricule': matricule,
        'centre': centre, 'antenne': antenne,
        'date_aut': date_aut, 'date_prof': date_prof,
        'date_med': date_med, 'date_psy': date_psy,
        'engins': engins, 'sites': sites
    }

    if os.path.exists(config['template']):
        excel_out = fill_excel_template(config['template'], config['cells'], data, photo_bytes)
        
        st.success("Carte générée dans le modèle Excel avec succès ! 🎉")
        st.download_button(
            label="📥 Télécharger la Carte (Excel .xlsx)",
            data=excel_out,
            file_name=f"Carte_{nom}_{matricule}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error(f"Le fichier modèle '{config['template']}' est introuvable.")
