# ============================================================================
#  app.py — point d'entrée Streamlit : connexion, barre latérale, puis
#  affichage de la page choisie. Chaque page vit dans ui/page_*.py ;
#  toute la logique métier est dans core.py.
# ============================================================================

import streamlit as st
import core
from ui.style import appliquer_css
from ui.auth import authentifier
from ui.sidebar import barre_laterale
from ui import demo_en_ligne, page_suivi, page_dashboard, page_depots, page_base, page_trames, page_admin

APP_VERSION = "8.3"
st.set_page_config(page_title=f"{core.NOM_PLATEFORME} — Solares Bauen", page_icon=":material/fact_check:",
                   layout="wide")
core.init_db()
appliquer_css()
demo_en_ligne.preparer()

authenticator = authentifier()
ctx = barre_laterale(authenticator, APP_VERSION)

PAGES = {
    "Suivi des visas": page_suivi,
    "Dashboard": page_dashboard,
    "Dépôts reçus": page_depots,
    "Base de données": page_base,
    "Trames de visa": page_trames,
    "Administration": page_admin,
}
PAGES[ctx["page"]].afficher(ctx)
