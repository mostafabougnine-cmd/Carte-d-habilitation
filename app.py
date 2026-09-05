# -*- coding: utf-8 -*-
import os
import io
import streamlit as st
import openpyxl
from openpyxl.drawing.image import Image as XLImage

st.set_page_config(page_title="Générateur de Cartes d'Habilitation - ONCF", layout="centered")

# خريطة الخانات المطابقة لهيكل Excel الأصلي
CELL_MAPPING = {
    "nom": "F5",
    "prenom": "J5",
    "matricule": "F6",
    "date_aut": "F8",
    "date_prof": "F9",
    "date_med": "F10",
    "date_psy": "F11",
    "materiel": "L4",
    "lines_sites": "Q4",
    "photo_cell": "B5"
}

TEMPLATES = {
    "CFT (Chef Formation Trains)": "CFT.xlsx",
    "CL (Conducteur de Ligne)": "CL.xlsx",
    "CTR (Chef de Train)": "CTR.xlsx",
    "CRMV (Conducteur de Manœuvre)": "CRMV.xlsx"
}

# الآلات الافتراضية لـ CTR و CFT
DEFAULT_MACHINES = "E1450 , E1400 ,E1250 ,DH400,Z2M"

def get_valid_template_path(filename):
    """البحث عن الملف في المجلد الرئيسي أو داخل مجلد data/"""
    paths_to_check = [
        filename,                         # المجلد الرئيسي (Root)
        os.path.join("data", filename)    # داخل مجلد data/
    ]
    for path in paths_to_check:
        if os.path.exists(path):
            return path
    return None

def safe_write_cell(ws, cell_address, value):
    """كتابة القيمة في الخلية العادية أو المدمجة مع الحفاظ على التنسيق الأصلي"""
    cell = ws[cell_address]
    target_cell = cell
    
    if type(cell).__name__ == 'MergedCell':
        for rng in ws.merged_cells.ranges:
            if cell_address in rng:
                target_cell = ws.cell(row=rng.min_row, column=rng.min_col)
                break
                
    if value and str(value).strip() != "":
        target_cell.value = value

def generate_card(template_path, data, photo_bytes):
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    # تعبئة الخانات المتغيرة (Centre و Antenne والـ Titres تظل ثابته من القالب الأصلي)
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

# تعيين القيمة الافتراضية للآلات حسب نوع البطاقة
default_materiel_val = ""
if "CTR" in selected_label or "CFT" in selected_label:
    default_materiel_val = DEFAULT_MACHINES

# 2. تحميل الصورة
uploaded_photo = st.file_uploader("Photo d'identité (JPG / PNG)", type=["jpg", "jpeg", "png"])

# 3. استمارة البيانات
with st.form("agent_form"):
    st.subheader("Informations de l'Agent")
    col1, col2 = st.columns(2)
    with col1:
        nom = st.text_input("Nom", "")
        matricule = st.text_input("Matricule", "")
        date_aut = st.text_input("Date d'autorisation", "")
        date_med = st.text_input("Date examen médical", "")
        materiel = st.text_input("Matériel / Locos / Rames", value=default_materiel_val)
    with col2:
        prenom = st.text_input("Prénom", "")
        date_prof = st.text_input("Date examen professionnel", "")
        date_psy = st.text_input("Date examen psychotechnique", "")
        lines_sites = st.text_input("Lignes / Sites autorisés", "")

    submit = st.form_submit_button("⚡ Générer la Carte")

if submit:
    template_path = get_valid_template_path(template_filename)

    if not template_path:
        st.error(f"⚠️ Fichier introuvable : '{template_filename}'. Assurez-vous que le fichier est présent dans votre projet.")
    else:
        photo_bytes = uploaded_photo.read() if uploaded_photo else None
        data = {
            'nom': nom, 'prenom': prenom, 'matricule': matricule,
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
