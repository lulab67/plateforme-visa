# ============================================================================
#  portail.py — point d'entrée de « Plateforme visa partenaires », le portail des entreprises.
#  Application Streamlit SÉPARÉE de la « Plateforme visa interne » (app.py) :
#  même base et même stockage, mais une entreprise connectée ne voit que ses
#  lots et ne peut que déposer. Lancement :
#     python -m streamlit run portail.py --server.port 8502
#  Toute la logique est dans core.py (section 16) ; l'écran dans ui/page_portail.py.
# ============================================================================

import streamlit as st
import core
from ui.style import appliquer_css
from ui import demo_en_ligne, page_portail

PORTAIL_VERSION = "1.6"
st.set_page_config(page_title=f"{core.NOM_PORTAIL} — Solares Bauen", page_icon=":material/upload_file:",
                   layout="wide", initial_sidebar_state="expanded")
core.init_db()
appliquer_css()
demo_en_ligne.preparer()

entreprise = page_portail.connexion()      # arrête la page tant que l'entreprise n'est pas connectée
page_portail.afficher(entreprise, PORTAIL_VERSION)
