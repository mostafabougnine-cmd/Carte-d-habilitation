# -*- coding: utf-8 -*-
import os
import io
import streamlit as st
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles.borders import Border, Side

st.set_page_config(page_title="Générateur de Cartes d'Habilitation - ONCF", layout="centered")

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

# الآلات الافتراضية لكل قالب
CFT_MACHINES = "E1450 , E1400 ,E1250 ,Z2M, DH400, DM600"
CTR_MACHINES = "E1450 , E1400 ,E1250 ,DH400,Z2M"
DEFAULT_SITE = "Site Voyageurs Kénitra"

def get_valid_template_path(filename):
    paths_to_check = [
        filename,
        os.path.join("data", filename)
    ]
    for path in paths_to_check:
        if os.path.exists(path):
            return path
    return None

def safe_write_cell(ws, cell_address, value):
    cell = ws[cell_address]
    target_cell = cell
    
    if type(cell).__name__ == 'MergedCell':
        for rng in ws.merged_cells.ranges:
            if cell_address in rng:
                target_cell = ws.cell(row=rng.min_row, column=rng.min_col)
                break
                
    if value and str(value).strip() != "":
        target_cell.value = value

def remove_inner_title_borders(ws):
    """إزالة الحدود الداخلية في السطر الثاني لكي لا تظهر خطوط عرضية في العنوان المدمج"""
    no_bottom = Side(border_style=None)
    for col in range(12, 20):
        cell = ws.cell(row=2, column=col)
        current_border = cell.border
        cell.border = Border(
            left=current_border.left,
            right=current_border.right,
            top=current_border.top,
            bottom=no_bottom
        )

def generate_card(template_path, data, photo_bytes):
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    # إزالة الخطوط الداخلية غير المرغوبة في رأس الجدول
    remove_inner_title_borders(ws)

    # تعبئة البيانات
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

selected_label = st.selectbox("Choisissez le modèle de carte :", list(TEMPLATES.keys()))
template_filename = TEMPLATES[selected_label]

# تعيين القيم الافتراضية
default_materiel_val = ""
default_sites_val = ""

if "CFT" in selected_label:
    default_materiel_val = CFT_MACHINES
elif "CTR" in selected_label:
    default_materiel_val = CTR_MACHINES

if "CRMV" in selected_label or "CFT" in selected_label:
    default_sites_val = DEFAULT_SITE

uploaded_photo = st.file_uploader("Photo d'identité (JPG / PNG)", type=["jpg", "jpeg", "png"])

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
        lines_sites = st.text_input("Lignes / Sites autorisés", value=default_sites_val)

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
