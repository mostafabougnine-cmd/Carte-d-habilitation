# -*- coding: utf-8 -*-
import os
import io
import streamlit as st
import openpyxl
from openpyxl.drawing.image import Image as XLImage

st.set_page_config(page_title="Générateur de Cartes d'Habilitation - ONCF", layout="centered")

# خريطة الخانات المطابقة للصورة تماماً
CELL_MAPPING = {
    "nom": "F6",
    "prenom": "J6",
    "matricule": "F7",
    "centre": "F8",
    "antenne": "J8",
    "date_aut": "F9",
    "date_prof": "F10",
    "date_med": "F11",
    "date_psy": "F12",
    "materiel": "L4",
    "lines_sites": "Q4",
    "photo_cell": "B6"
}

TEMPLATES = {
    "CFT (Chef Formation Trains)": "CFT.xlsx",
    "CL (Conducteur de Ligne)": "CL.xlsx",
    "CTR (Chef de Train)": "CTR.xlsx",
    "CRMV (Conducteur de Manœuvre)": "CRMV.xlsx"
}

def safe_write_cell(ws, cell_address, value):
    """كتابة القيمة وتفريغ الخلية إذا كانت القيمة فارغة"""
    cell = ws[cell_address]
    target_cell = cell
    
    # التعامل مع الخلايا المدمجة
    if type(cell).__name__ == 'MergedCell':
        for rng in ws.merged_cells.ranges:
            if cell_address in rng:
                target_cell = ws.cell(row=rng.min_row, column=rng.min_col)
                break
                
    # إذا كانت هناك قيمة مدخلة نكتبها، وإذا تركها المستخدم فارغة نفرغ الخلية القديمة
    if value and str(value).strip() != "":
        target_cell.value = value
    else:
        target_cell.value = None

def generate_card(template_path, data, photo_bytes):
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    # كتابة أو تعديل الخانات
    for key, cell_address in CELL_MAPPING.items():
        if key in data and key != "photo_cell":
            safe_write_cell(ws, cell_address, data[key])

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

# 1. اختيار نوع البطاقة
selected_label = st.selectbox("Choisissez le modèle de carte :", list(TEMPLATES.keys()))
template_filename = TEMPLATES[selected_label]

# 2. تحميل الصورة
uploaded_photo = st.file_uploader("Photo d'identité (JPG / PNG)", type=["jpg", "jpeg", "png"])

# 3. استمارة البيانات
with st.form("agent_form"):
    st.subheader("Informations de l'Agent")
    col1, col2 = st.columns(2)
    with col1:
        nom = st.text_input("Nom", "")
        matricule = st.text_input("Matricule", "")
        centre = st.text_input("Centre", "")
        date_aut = st.text_input("Date d'autorisation", "")
        date_med = st.text_input("Date examen médical", "")
        materiel = st.text_input("Matériel / Locos / Rames", "")
    with col2:
        prenom = st.text_input("Prénom", "")
        antenne = st.text_input("Antenne", "")
        date_prof = st.text_input("Date examen professionnel", "")
        date_psy = st.text_input("Date examen psychotechnique", "")
        lines_sites = st.text_input("Lignes / Sites autorisés", "")

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
            'date_med': date_med, 'date_psy': date_psy,
            'materiel': materiel, 'lines_sites': lines_sites
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
