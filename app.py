# -*- coding: utf-8 -*-
import os
import io
import streamlit as st
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics

st.set_page_config(page_title="Générateur Carte Habilitation ONCF", layout="centered")

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

def draw_wrapped(c, text, x, y, width_chars=35, leading=20, size=13):
    c.setFont(FONT, size)
    lines = wrap_text(text, width_chars)
    yy = y
    for line in lines:
        c.drawString(x, yy, line)
        yy -= leading
    return yy

def generate_pdf_bytes(data, photo_bytes, logo_bytes, tampon_bytes):
    buffer = io.BytesIO()
    CARD_W, CARD_H = 1100, 550
    c = canvas.Canvas(buffer, pagesize=(CARD_W, CARD_H))

    # Cadre extérieur
    c.setLineWidth(1.5)
    c.rect(10, 10, CARD_W - 20, CARD_H - 20)

    # Séparations
    c.line(550, 10, 550, CARD_H - 10)
    c.line(550, 360, CARD_W - 10, 360)
    c.line(550, 170, CARD_W - 10, 170)
    c.line(850, 170, 850, CARD_H - 10)

    # Logo
    if logo_bytes:
        try:
            c.drawImage(ImageReader(io.BytesIO(logo_bytes)), 25, CARD_H - 90, width=130, height=65, preserveAspectRatio=True, mask='auto')
        except Exception:
            c.setFont(FONT, 22)
            c.drawString(30, CARD_H - 60, "ONCF")
    else:
        c.setFont(FONT, 22)
        c.setFillColorRGB(0.85, 0.35, 0.1)
        c.drawString(30, CARD_H - 60, "ONCF")
        c.setFillColorRGB(0, 0, 0)

    c.setFont(FONT, 11)
    c.drawString(30, CARD_H - 110, "PV/DTV/EPTCN")

    c.setFont(FONT, 28)
    c.drawString(240, CARD_H - 60, "Titre d’habilitation")

    c.setFont(FONT, 18)
    c.setFillColorRGB(0.85, 0.35, 0.1)
    c.drawString(260, CARD_H - 95, data['fonction'])
    c.setFillColorRGB(0, 0, 0)

    # Photo
    px, py, pw, ph = 25, CARD_H - 330, 130, 160
    c.rect(px, py, pw, ph)
    if photo_bytes:
        try:
            c.drawImage(ImageReader(io.BytesIO(photo_bytes)), px + 2, py + 2, width=pw - 4, height=ph - 4, preserveAspectRatio=True)
        except Exception:
            pass

    # Infos Agent
    lx, ly, gap = 170, CARD_H - 150, 35

    c.setFont(FONT, 14)
    c.drawString(lx, ly, "Nom")
    c.drawString(lx + 80, ly, ":  " + data['nom'])
    c.drawString(lx + 230, ly, "Prénom  :  " + data['prenom'])

    ly -= gap
    c.drawString(lx, ly, "Matricule")
    c.drawString(lx + 80, ly, ":  " + data['matricule'])

    ly -= gap
    c.drawString(lx, ly, "Centre")
    c.drawString(lx + 80, ly, ":  " + data['centre'])
    c.drawString(lx + 230, ly, "Antenne : " + data['antenne'])

    ly -= gap
    c.drawString(lx - 145, ly, "Date d’autorisation")
    c.drawString(lx + 100, ly, ":  " + data['date_aut'])

    ly -= gap
    c.drawString(lx - 145, ly, "Date de l’examen professionnel")
    c.drawString(lx + 100, ly, ":  " + data['date_prof'])

    ly -= gap
    c.drawString(lx - 145, ly, "Date de l’examen médical")
    c.drawString(lx + 100, ly, ":  " + data['date_med'])

    ly -= gap
    c.drawString(lx - 145, ly, "Date de l’examen psychotechniqu")
    c.drawString(lx + 100, ly, ":  " + data['date_psy'])

    # Tampon
    if tampon_bytes:
        try:
            c.drawImage(ImageReader(io.BytesIO(tampon_bytes)), 360, 40, width=170, height=130, preserveAspectRatio=True, mask='auto')
        except Exception:
            pass

    # Partie Droite
    c.setFont(FONT, 15)
    c.drawCentredString(700, CARD_H - 40, "Autorisé à arrêter les locos et rames")
    c.drawCentredString(700, CARD_H - 60, "suivantes")
    draw_wrapped(c, data['engins'], 565, CARD_H - 100, width_chars=35, leading=22, size=13)

    c.setFont(FONT, 15)
    c.drawCentredString(975, CARD_H - 40, "Autorisé aux sites suivants")
    draw_wrapped(c, data['sites'], 860, CARD_H - 110, width_chars=18, leading=20, size=13)

    c.setFont(FONT, 15)
    c.drawCentredString(700, 320, "Autorisé pour la Manœuvre du")
    c.setFont(FONT, 14)
    c.drawCentredString(700, 230, data['manoeuvre'])

    # Red Texts
    c.setFillColorRGB(0.9, 0, 0)
    c.setFont(FONT, 9)
    wy = 145
    c.drawString(515, wy, "• Cette carte doit être présentée à tout contrôle ;")
    c.drawString(515, wy - 20, "• Doit être restituée en cas de retrait temporaire ou définitif des fonctions de sécurité ;")
    c.drawString(515, wy - 40, "• La fonction de sécurité à laquelle vous êtes habilité ne peut s’exercer qu’en pleine possession de vos")
    c.drawString(515, wy - 52, "facultés et moyens.")

    c.save()
    buffer.seek(0)
    return buffer.getvalue()

st.title("🎴 Générateur Manuel de Carte ONCF")

with st.form("card_form"):
    st.subheader("1. Informations de l'Agent")
    col1, col2 = st.columns(2)
    with col1:
        nom = st.text_input("Nom", "OUGUI")
        matricule = st.text_input("Matricule", "45892H")
        centre = st.text_input("Centre", "CCFTC Kénitra")
    with col2:
        prenom = st.text_input("Prénom", "HAMZA")
        fonction = st.text_input("Fonction (titre)", "Chef Formation Trains")
        antenne = st.text_input("Antenne", "ACFTC Kénitra")

    st.subheader("2. Dates")
    col3, col4 = st.columns(2)
    with col3:
        date_aut = st.text_input("Date d'autorisation", "17/02/2022")
        date_med = st.text_input("Date examen médical", "17/05/2022")
    with col4:
        date_prof = st.text_input("Date examen professionnel", "08/01/2026")
        date_psy = st.text_input("Date examen psychotechnique", "05/08/2026")

    st.subheader("3. Autorisations & Engins")
    engins = st.text_area("Engins autorisés", "E1450 , E1400 , E1250 , Z2M , DH350, DH400 , DM600")
    manoeuvre = st.text_input("Autorisé pour la Manœuvre du", "Matériel à Voyageurs")
    sites = st.text_area("Sites autorisés", "Site Voyageurs Kénitra")

    st.subheader("4. Images (Logo, Tampon, Photo)")
    uploaded_photo = st.file_uploader("Photo d'identité (JPG / PNG)", type=["jpg", "jpeg", "png"])
    uploaded_logo = st.file_uploader("Logo ONCF (Optionnel)", type=["png", "jpg"])
    uploaded_tampon = st.file_uploader("Tampon / Signature (Optionnel)", type=["png", "jpg"])

    submit = st.form_submit_button("⚡ Générer la Carte PDF")

if submit:
    photo_bytes = uploaded_photo.read() if uploaded_photo else None
    
    logo_bytes = uploaded_logo.read() if uploaded_logo else (
        open("photos/oncf_logo.png", "rb").read() if os.path.exists("photos/oncf_logo.png") else None
    )
    tampon_bytes = uploaded_tampon.read() if uploaded_tampon else (
        open("photos/tampon.png", "rb").read() if os.path.exists("photos/tampon.png") else None
    )

    data = {
        'nom': nom, 'prenom': prenom, 'matricule': matricule,
        'fonction': fonction, 'centre': centre, 'antenne': antenne,
        'date_aut': date_aut, 'date_prof': date_prof,
        'date_med': date_med, 'date_psy': date_psy,
        'engins': engins, 'manoeuvre': manoeuvre, 'sites': sites
    }

    pdf_data = generate_pdf_bytes(data, photo_bytes, logo_bytes, tampon_bytes)

    st.success("Carte générée avec succès ! 🎉")
    st.download_button(
        label="📥 Télécharger le PDF de la Carte",
        data=pdf_data,
        file_name=f"Carte_{nom}_{matricule}.pdf",
        mime="application/pdf"
    )