# -*- coding: utf-8 -*-
import os
import io
import streamlit as st
import openpyxl
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics

st.set_page_config(page_title="Générateur Carte Habilitation ONCF", layout="centered")

def load_excel_data(file_path):
    if not os.path.exists(file_path):
        return {"engins": "", "sites": "", "manoeuvre": ""}
    
    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = wb.active
    
    engins, sites, manoeuvre = "", "", ""
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            val = str(ws.cell(row=r, column=c).value or "").strip()
            if "E1450" in val or "DH" in val or "Z2M" in val:
                engins = val
            elif "Site" in val or "Lignes" in val or "Réseau" in val:
                sites = val
            elif "Matériel" in val:
                manoeuvre = val
                
    return {"engins": engins, "sites": sites, "manoeuvre": manoeuvre}

FUNCTIONS_CONFIG = {
    "Chef Formation Trains": {
        "excel": "data/Cartes d'habilitation CFT.xlsx",
        "bg_image": "photos/carte_cft.png"
    },
    "Chef de Train": {
        "excel": "data/carte d'habilitation CTR.xlsx",
        "bg_image": "photos/carte_ctr.png"
    },
    "Conducteur de Manœuvre": {
        "excel": "data/carte d'habilitation CRMV.xlsx",
        "bg_image": "photos/carte_crmv.png"
    },
    "Conducteur de Ligne": {
        "excel": "data/carte d'habilitation CL.xlsx",
        "bg_image": "photos/carte_cl.png"
    }
}

def register_font():
    candidates = [
        r"C:\Windows\Fonts\arial.ttf",
        r"C:\Windows\Fonts\calibri.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ]
    for p in candidates:
        if os.path.exists(p):
            try:
                pdfmetrics.registerFont(TTFont("MyFont", p))
                return "MyFont"
            except Exception:
                pass
    return "Helvetica"

FONT = register_font()

def wrap_text(text, max_chars):
    if not text:
        return [""]
    words = str(text).split()
    lines, current = [], ""
    for word in words:
        test = word if not current else current + " " + word
        if len(test) <= max_chars:
            current = test
        else:
            if current:
                lines.append(current)
            current = word
    if current:
        lines.append(current)
    return lines or [""]

def draw_wrapped(c, text, x, y, width_chars=32, leading=18, size=12):
    c.setFont(FONT, size)
    lines = wrap_text(text, width_chars)
    yy = y
    for line in lines:
        c.drawString(x, yy, line)
        yy -= leading

def generate_pdf_bytes(data, photo_bytes, bg_path, fonction_name):
    buffer = io.BytesIO()
    CARD_W, CARD_H = 1100, 550
    c = canvas.Canvas(buffer, pagesize=(CARD_W, CARD_H))

    # 1. رسم الصورة الخلفية
    if os.path.exists(bg_path):
        c.drawImage(bg_path, 0, 0, width=CARD_W, height=CARD_H)

    # 2. إضافة صورة الشخص فـ الإطار الخاوي بالضبط
    px, py, pw, ph = 85, CARD_H - 280, 105, 125
    if photo_bytes:
        try:
            c.drawImage(ImageReader(io.BytesIO(photo_bytes)), px, py, width=pw, height=ph, preserveAspectRatio=True)
        except Exception:
            pass

    c.setFont(FONT, 13)
    c.setFillColorRGB(0, 0, 0)

    # 3. كتابة المعلومات الشخصية فـ الأماكن الخاوية فقط
    # Nom & Prénom
    c.drawString(270, CARD_H - 180, data['nom'])
    c.drawString(400, CARD_H - 180, data['prenom'])

    # Matricule
    c.drawString(270, CARD_H - 212, data['matricule'])

    # Centre & Antenne (إيلا ما كانوش مكتوبين فـ الخلفية)
    if "CCFTC" not in fonction_name: 
        c.drawString(270, CARD_H - 245, data['centre'])
        c.drawString(450, CARD_H - 245, data['antenne'])

    # Dates
    c.drawString(270, CARD_H - 275, data['date_aut'])
    c.drawString(270, CARD_H - 308, data['date_prof'])
    c.drawString(270, CARD_H - 340, data['date_med'])
    c.drawString(270, CARD_H - 372, data['date_psy'])

    # 4. طباعة Engins / Sites فقط إيلا كان الإطار خاوي فـ صورة الوظيفة
    # بالنسبة لـ Conducteur de Ligne و Conducteur de Manœuvre (الجهات الخاوية):
    if fonction_name in ["Conducteur de Ligne", "Conducteur de Manœuvre"]:
        if data['engins']:
            draw_wrapped(c, data['engins'], 560, CARD_H - 140, width_chars=32, leading=20, size=12)
        if data['sites'] and fonction_name == "Conducteur de Ligne":
            draw_wrapped(c, data['sites'], 820, CARD_H - 140, width_chars=20, leading=20, size=12)

    c.save()
    buffer.seek(0)
    return buffer.getvalue()

st.title("🎴 Générateur Carte Habilitation ONCF")

selected_fonction = st.selectbox(
    "Choix de la Fonction",
    list(FUNCTIONS_CONFIG.keys())
)

config = FUNCTIONS_CONFIG[selected_fonction]
excel_data = load_excel_data(config['excel'])

if os.path.exists(config['bg_image']):
    st.image(config['bg_image'], caption=f"Modèle: {selected_fonction}", use_container_width=True)

with st.form("card_form"):
    st.subheader("1. Photo de l'Agent")
    uploaded_photo = st.file_uploader("Photo d'identité (JPG / PNG)", type=["jpg", "jpeg", "png"])

    st.subheader("2. Informations Personnelles")
    col1, col2 = st.columns(2)
    with col1:
        nom = st.text_input("Nom", "")
        matricule = st.text_input("Matricule", "")
        centre = st.text_input("Centre", "CCFTC Kénitra")
    with col2:
        prenom = st.text_input("Prénom", "")
        antenne = st.text_input("Antenne", "ACFTC Kénitra")

    st.subheader("3. Dates")
    col3, col4 = st.columns(2)
    with col3:
        date_aut = st.text_input("Date d'autorisation", "")
        date_med = st.text_input("Date examen médical", "")
    with col4:
        date_prof = st.text_input("Date examen professionnel", "")
        date_psy = st.text_input("Date examen psychotechnique", "")

    # إظهار الخانات للتعديل فقط للوظائف اللي محتاجة تعمير
    st.subheader("4. Données de la fonction")
    engins = st.text_area("Engins autorisés", excel_data['engins'])
    sites = st.text_area("Sites / Lignes autorisés", excel_data['sites'])
    manoeuvre = st.text_input("Autorisé pour la Manœuvre du", excel_data['manoeuvre'])

    submit = st.form_submit_button("⚡ Générer la Carte PDF")

if submit:
    photo_bytes = uploaded_photo.read() if uploaded_photo else None

    data = {
        'nom': nom, 'prenom': prenom, 'matricule': matricule,
        'fonction': selected_fonction, 'centre': centre, 'antenne': antenne,
        'date_aut': date_aut, 'date_prof': date_prof,
        'date_med': date_med, 'date_psy': date_psy,
        'engins': engins, 'manoeuvre': manoeuvre, 'sites': sites
    }

    pdf_data = generate_pdf_bytes(data, photo_bytes, config['bg_image'], selected_fonction)

    st.success("Carte générée avec succès ! 🎉")
    st.download_button(
        label="📥 Télécharger le PDF de la Carte",
        data=pdf_data,
        file_name=f"Carte_{nom}_{matricule}.pdf",
        mime="application/pdf"
    )
