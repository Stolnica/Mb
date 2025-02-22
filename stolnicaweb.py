import streamlit as st
from datetime import datetime
from PIL import Image
import docx
import os
from zipfile import ZipFile
from streamlit_option_menu import option_menu
import fitz  # PyMuPDF za delo s PDF datotekami
import base64

# --- Določanje izbir ---
izbor1 = "Oznanila"
izbor2 = "Vabila na dogodke"
izbor3 = "Obvestila"
izbor4 = "Sveto leto"
izbor5 = "Obnova stolnice"
izbor6 = "Pesmi"

# Preverimo dan in čas
dan_v_tednu = datetime.today().weekday()
ura = datetime.now().hour

# Nastavimo vrstni red menija in ikon
if dan_v_tednu == 6 and ura < 12:  # Nedelja dopoldan
    mozni_izbori = [izbor6, izbor2, izbor3, izbor4, izbor5, izbor1]  # Pesmi na vrhu
    ikone = ['music-note-list', 'calendar2-event', 'brightness-high', 'book', 'buildings', 'card-text']
else:  # Ostali dnevi + nedelja popoldan
    mozni_izbori = [izbor1, izbor2, izbor3, izbor4, izbor5, izbor6]  # Privzeti vrstni red
    ikone = ['card-text', 'calendar2-event', 'brightness-high', 'book', 'buildings', 'music-note-list']

# Funkcija za branje besedila iz Word dokumenta
def preberi_docx(datoteka):
    if not os.path.exists(datoteka):
        return None
    doc = docx.Document(datoteka)
    odstavki = [para.text.strip() for para in doc.paragraphs if para.text.strip()]
    return "\n\n".join(odstavki)

# Funkcija za pridobitev slik iz Word dokumenta
def pridobi_slike(datoteka, izhodna_mapa="slike"):
    if not os.path.exists(datoteka):
        return []
    slike_poti = []
    with ZipFile(datoteka, "r") as docx_zip:
        for datoteka in docx_zip.namelist():
            if datoteka.startswith("word/media/"):
                slika_pot = os.path.join(izhodna_mapa, os.path.basename(datoteka))
                if not os.path.exists(izhodna_mapa):
                    os.makedirs(izhodna_mapa)
                with open(slika_pot, "wb") as f:
                    f.write(docx_zip.read(datoteka))
                slike_poti.append(slika_pot)
    return slike_poti

# Funkcija za pretvorbo PDF datoteke v slike
def pdf_v_slike(pdf_pot, izhodna_mapa="slike_pdf", resolucija=3):
    if not os.path.exists(pdf_pot):
        return []
    if not os.path.exists(izhodna_mapa):
        os.makedirs(izhodna_mapa)

    slike_poti = []
    pdf_dokument = fitz.open(pdf_pot)

    for stran_num, stran in enumerate(pdf_dokument):
        matrika = fitz.Matrix(resolucija, resolucija)
        pix = stran.get_pixmap(matrix=matrika)
        slika_pot = os.path.join(izhodna_mapa, f"stran_{stran_num + 1}.png")
        pix.save(slika_pot)
        slike_poti.append(slika_pot)

    return slike_poti

# Dodajte CSS za povečavo glavne vsebine
st.markdown("""
<style>
    .main-content {
        transform-origin: top left; /* Določimo izhodišče povečave */
        transition: transform 0.3s ease; /* Dodamo animacijo za gladko povečavo */
    }
    .main-content:hover { /* Opcija: Povečava ob prehodu miške čez vsebino */
        transform: scale(1.1); /* Povečamo za 10% */
    }

    /* Zagotovite, da se vsebina ne preliva izven zaslona */
    body {
        overflow-x: hidden;
    }

    /* Prilagodite velikost pisave za manjše zaslone */
    @media (max-width: 768px) {
        .font {
            font-size: 40px;
        }
        .main-content {
            transform: scale(1); /* Onemogočite povečavo ob prehodu miške na mobilnih napravah */
        }
    }

    /* Zagotovite, da so slike odzivne */
    img {
        max-width: 100%;
        height: auto;
    }
</style>
""", unsafe_allow_html=True)

# Naložimo slike logo in sidebarimage
logo = Image.open(r'./slomsek.jpg')
sidebarimage = Image.open(r'./cerkev.jpg')

# Pot do dokumentov in zvočnih datotek
pot_dokumentov = "./"
pot_zvoka = "./"

# --- Glavni meni ---
with st.sidebar:
    choose = option_menu("Izberi:", mozni_izbori,
                         icons=ikone,
                         menu_icon="list", default_index=0)

    # Gumb za "Domov"
    st.markdown(
        f"""
        <a href="https://stolnicamaribor.si" target="_blank">
            <button style="width: 100%; padding: 10px; background-color: #4CAF50; color: white; border: none; border-radius: 5px; font-size: 16px;">
                <span>Domov</span>
            </button>
        </a>
        """,
        unsafe_allow_html=True
    )
    st.markdown("---")
    st.sidebar.image(sidebarimage, width=290)

# Stil za poudarjeno besedilo
st.markdown("""
<head>
    <link href="https://fonts.googleapis.com/css2?family=Nunito:wght@300&display=swap" rel="stylesheet">
    <style>
        .font { 
            font-size:60px; 
            font-family: 'Lato', sans-serif; 
            color: #800000;
            font-weight: bold; 
        }
    </style>
</head>
""", unsafe_allow_html=True)

# Meta tag za zoomanje na mobilnih napravah
st.markdown("""
<head>
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=5.0, user-scalable=yes">
</head>
""", unsafe_allow_html=True)

# Funkcija za prikaz naslova in logotipa
def prikazi_naslov_in_logo(naslov):
    col1, col2 = st.columns([0.8, 0.2])
    with col1:
        st.markdown(f'<p class="font">{naslov}</p>', unsafe_allow_html=True)
    with col2:
        st.image(logo, width=130)

# Funkcija za predvajanje zvoka
def predvajaj_zvok(pot_do_zvocne_datoteke):
    if os.path.exists(pot_do_zvocne_datoteke):
        with open(pot_do_zvocne_datoteke, "rb") as f:
            data = f.read()
            b64 = base64.b64encode(data).decode()
            st.markdown(f"""
                <audio autoplay controls>
                <source src="data:audio/mpeg;base64,{b64}" type="audio/mpeg">
                Tvoj brskalnik ne podpira predvajanja zvoka.
                </audio>
                """, unsafe_allow_html=True)

# Glavna vsebina
if choose in mozni_izbori:
    # Dodajte razred "main-content" na div, ki obdaja glavno vsebino
    st.markdown('<div class="main-content">', unsafe_allow_html=True)
    prikazi_naslov_in_logo(choose)

    # Prikaz besedila in slik iz DOCX
    docx_pot = os.path.join(pot_dokumentov, f"{choose}.docx")
    besedilo = preberi_docx(docx_pot)
    slike = pridobi_slike(docx_pot)

    if besedilo:
        for odstavek in besedilo.split("\n\n"):
            st.markdown(f'<p class="bold-text">{odstavek}</p>', unsafe_allow_html=True)

    if slike:
        st.markdown("---")
        for slika_pot in slike:
            st.image(slika_pot, width=700)
            st.markdown("<br>", unsafe_allow_html=True)

    # Preverjanje in prikaz dodatnih DOCX dokumentov
    stevilka = 1
    while os.path.exists(os.path.join(pot_dokumentov, f"{choose}{stevilka}.docx")):
        dodatni_pot = os.path.join(pot_dokumentov, f"{choose}{stevilka}.docx")
        dodatno_besedilo = preberi_docx(dodatni_pot)
        dodatne_slike = pridobi_slike(dodatni_pot)

        if dodatno_besedilo:
            st.markdown("---")
            for odstavek in dodatno_besedilo.split("\n\n"):
                st.markdown(f'<p class="bold-text">{odstavek}</p>', unsafe_allow_html=True)

        if dodatne_slike:
            for slika_pot in dodatne_slike:
                st.image(slika_pot, width=700)

        stevilka += 1

    # Prikaz PDF kot slike
    pdf_pot = os.path.join(pot_dokumentov, f"{choose}.pdf")
    for slika_pot in pdf_v_slike(pdf_pot, resolucija=4):
        st.image(slika_pot, width=700)

    # Predvajanje zvoka
    predvajaj_zvok(os.path.join(pot_zvoka, f"{choose}.mp3"))

    st.markdown('</div>', unsafe_allow_html=True)  # Zaključimo div