# -*- coding: utf-8 -*-
import os
import io
import streamlit as st
import openpyxl
from PIL import Image, ImageDraw, ImageFont
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

st.set_page_config(page_title="Générateur Carte Habilitation ONCF", layout="centered")

def load_excel_data(file_path):
    if not os.path.exists(file_path):
        return {"engins": "", "sites": "", "manoeuvre": ""}
    try:
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
    except Exception:
        return {"engins": "", "sites": "", "manoeuvre": ""}

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

POSITIONS = {
    "Chef Formation Trains": {
        "photo": (80, 140, 100, 120),
        "nom": (230, 145), "prenom": (400, 145),
        "matricule": (230, 180),
        "dates": (250, 245, 280, 315, 350)
    },
    "Chef de Train": {
        "photo": (80, 140, 100, 120),
        "nom": (230, 145), "prenom": (400, 145),
        "matricule": (230, 180),
        "dates": (250, 245, 280, 315, 350)
    },
    "Conducteur de Manœuvre": {
        "photo": (80, 140, 100, 120),
        "nom": (230, 145), "prenom": (400, 145),
        "matricule": (230, 180),
        "centre": (230, 215), "antenne": (400, 215),
        "dates": (250, 245, 280, 315, 350),
        "engins": (520, 140)
    },
    "Conducteur de Ligne": {
        "photo": (80, 140, 100, 120),
        "nom": (230, 145), "prenom": (400, 145),
        "matricule": (230, 180),
        "centre": (230, 215), "antenne": (400, 215),
        "dates": (250, 245, 280, 315, 350),
        "engins": (520, 140), "sites": (750, 140)
    }
}

def load_system_font(size=18):
    candidates = [
        r"C:\Windows\Fonts\arial.ttf",
        r"C:\Windows\Fonts\calibri.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
    ]
    for p in candidates:
        if os.path.exists(p):
            try:
                return ImageFont.truetype(p, size)
            except Exception:
                pass
    return ImageFont.load_default()

def draw_card_image(bg_input, photo_bytes, data, fonction):
    # فتح خلفية الكارطة سوا كانت مسار ملف ولا رفع مباشر
    if isinstance(bg_input, str) and os.path.exists(bg_input):
        card_img = Image.open(bg_input).convert("RGB")
    elif hasattr(bg_input, 'read'):
        card_img = Image.open(bg_input).convert("RGB")
    else:
        # إنشاء خلفية بيضاء فـ حالة عدم وجود خلفية
        card_img = Image.new("RGB", (1000, 450), color=(255, 255, 255))

    draw = ImageDraw.Draw(card_img)
    font = load_system_font(size=20)
    pos = POSITIONS.get(fonction, POSITIONS["Chef Formation Trains"])

    # 1. تركيب صورة الشخص
    if photo_bytes:
        try:
            user_photo = Image.open(io.BytesIO(photo_bytes)).convert("RGB")
            px, py, pw, ph = pos["photo"]
            user_photo = user_photo.resize((pw, ph))
            card_img.paste(user_photo, (px, py))
        except Exception as e:
            st.error(f"Erreur photo: {e}")

    # 2. الكتابة فوق الصورة
    black = (0, 0, 0)
    draw.text(pos["nom"], str(data['nom']), fill=black, font=font)
    draw.text(pos["prenom"], str(data['prenom']), fill=black, font=font)
    draw.text(pos["matricule"], str(data['matricule']), fill=black, font=font)

    if "centre" in pos:
        draw.text(pos["centre"], str(data['centre']), fill=black, font=font)
        draw.text(pos["antenne"], str(data['antenne']), fill=black, font=font)

    d_x, d_aut, d_prof, d_med, d_psy = pos["dates"]
    draw.text((d_x, d_aut), str(data['date_aut']), fill=black, font=font)
    draw.text((d_x, d_prof), str(data['date_prof']), fill=black, font=font)
    draw.text((d_x, d_med), str(data['date_med']), fill=black, font=font)
    draw.text((d_x, d_psy), str(data['date_psy']), fill=black, font=font)

    if "engins" in pos and data['engins']:
        draw.text(pos["engins"], str(data['engins']), fill=black, font=font)
    if "sites" in pos and data['sites']:
        draw.text(pos["sites"], str(data['sites']), fill=black, font=font)

    return card_img

def convert_pil_to_pdf(pil_img):
    buffer = io.BytesIO()
    w, h = pil_img.size
    c = canvas.Canvas(buffer, pagesize=(w, h))
    img_bytes = io.BytesIO()
    pil_img.save(img_bytes, format='PNG')
    img_bytes.seek(0)
    c.drawImage(ImageReader(img_bytes), 0, 0, width=w, height=h)
    c.save()
    buffer.seek(0)
    return buffer.getvalue()

# Streamlit UI
st.title("🎴 Générateur Carte Habilitation ONCF")

selected_fonction = st.selectbox("Choix de la Fonction", list(FUNCTIONS_CONFIG.keys()))
config = FUNCTIONS_CONFIG[selected_fonction]
excel_data = load_excel_data(config['excel'])

st.subheader("1. Modèle de la carte (Image de fond)")
uploaded_bg = st.file_uploader("Si le modèle n'apparaît pas, chargez l'image de la carte ici (PNG/JPG)", type=["png", "jpg", "jpeg"])

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

    st.subheader("5. Données de la fonction")
    engins = st.text_area("Engins autorisés", excel_data['engins'])
    sites = st.text_area("Sites / Lignes autorisés", excel_data['sites'])

    submit = st.form_submit_button("⚡ Générer et Visualiser la Carte")

if submit:
    bg_to_use = uploaded_bg if uploaded_bg is not None else config['bg_image']
    photo_bytes = uploaded_photo.read() if uploaded_photo else None

    data = {
        'nom': nom, 'prenom': prenom, 'matricule': matricule,
        'centre': centre, 'antenne': antenne,
        'date_aut': date_aut, 'date_prof': date_prof,
        'date_med': date_med, 'date_psy': date_psy,
        'engins': engins, 'sites': sites
    }

    generated_img = draw_card_image(bg_to_use, photo_bytes, data, selected_fonction)

    st.subheader("🖼️ Aperçu de la Carte Générée :")
    st.image(generated_img, caption="Carte Finale", use_container_width=True)

    pdf_bytes = convert_pil_to_pdf(generated_img)
    st.download_button(
        label="📥 Télécharger la Carte en PDF",
        data=pdf_bytes,
        file_name=f"Carte_{nom}_{matricule}.pdf",
        mime="application/pdf"
    )
