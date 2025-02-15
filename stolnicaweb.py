import streamlit as st
from PIL import Image
import docx
import os
from zipfile import ZipFile
#import webbrowser
from streamlit_option_menu import option_menu
import sys


print(sys.executable)
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

# Naložimo logo
#logo = Image.open(r'slomsek.png')   # Image.open(r'slomsek.png')
logo = Image.open(r'./slomsek.png')
# Pot do dokumentov
pot_dokumentov = ""   # "c:/Users/evgen/Documents/2024/"
datoteke = {
    "Obvestila": "Obvestila.docx",
    "Dogodki": "Dogodki.docx",
    "Odpustki": "Odpustki.docx",
    "Skupine": "Skupine.docx"
}

# --- Glavni meni ---
with st.sidebar:
    choose = option_menu("Izberi:", ["Obvestila", "Dogodki", "Odpustki", "Skupine", "Domov"],
                         icons=['card-text', 'calendar2-event', 'brightness-high', 'person lines fill', 'house'],
                         menu_icon="list", default_index=0)

# Stil za poudarjeno besedilo
st.markdown("""
<style>
.bold-text {
    font-weight: bold;
}
.font {
    font-size:35px;
    font-family: 'Cooper Black';
    color: #FF9633;
}
</style>
""", unsafe_allow_html=True)

# Funkcija za prikaz naslova menija in logotipa na vsaki strani
def prikazi_naslov_in_logo(naslov):
    col1, col2 = st.columns([0.8, 0.2])
    with col1:
        st.markdown(f'<p class="font">{naslov}</p>', unsafe_allow_html=True)
    with col2:
        st.image(logo, width=130)
    st.sidebar.image(logo, width=290)

# Stran "Domov" - preusmeritev na spletno stran
if choose == "Domov":
    webbrowser.open_new_tab("https://stolnicamaribor.si")

# Strani menija (Obvestila, Dogodki, Odpustki, Skupine)
elif choose in datoteke:
    prikazi_naslov_in_logo(choose)  # Prikaz naslova in logotipa

    # Prikaz glavnega dokumenta
    datoteka_pot = os.path.join(pot_dokumentov, datoteke[choose])
    besedilo = preberi_docx(datoteka_pot)
    slike = pridobi_slike(datoteka_pot)

    if besedilo:
        odstavki = besedilo.split("\n\n")
        for odstavek in odstavki:
            st.markdown(f'<p class="bold-text">{odstavek}</p>', unsafe_allow_html=True)

    if slike:
        st.markdown("---")
        for slika_pot in slike:
            st.image(slika_pot, width=700)
            st.markdown("<br>", unsafe_allow_html=True)

    # Dodatno besedilo in slike iz datotek Ime1.docx, Ime2.docx ...
    stevilka = 1
    while True:
        dodatna_datoteka = os.path.join(pot_dokumentov, f"{choose}{stevilka}.docx")
        if not os.path.exists(dodatna_datoteka):
            break  # Če datoteka ne obstaja, nehamo iskati naprej
        dodatno_besedilo = preberi_docx(dodatna_datoteka)
        dodatne_slike = pridobi_slike(dodatna_datoteka)  # Pridobi slike tudi iz dodatnih dokumentov

        if dodatno_besedilo:
            st.markdown("---")  # Ločilna črta pred dodatnim besedilom
            odstavki = dodatno_besedilo.split("\n\n")
            for odstavek in odstavki:
                st.markdown(f'<p class="bold-text">{odstavek}</p>', unsafe_allow_html=True)

        if dodatne_slike:
            st.markdown("---")  # Ločilna črta pred dodatnimi slikami
            for slika_pot in dodatne_slike:
                st.image(slika_pot, width=700)
                st.markdown("<br>", unsafe_allow_html=True)

        stevilka += 1  # Nadaljuj s preverjanjem naslednje datoteke
