# ============================================================================
#  ui/demo_en_ligne.py — démonstration publiée sur Streamlit Community Cloud.
#  Active seulement si les secrets de l'appli contiennent demo_en_ligne = true.
#  En local (pas de secrets), rien ne change.
#  En ligne, le disque est effacé aux redémarrages et chaque appli a le sien :
#  la démo est rechargée si la base est vide, et un message le dit.
# ============================================================================

import streamlit as st
import core
import demo


def lire_secret(nom):
    """Valeur d'un secret Streamlit, ou None (le fichier de secrets n'existe pas en local)."""
    try:
        if nom in st.secrets:
            return st.secrets[nom]
    except FileNotFoundError:
        pass
    return None


@st.cache_resource
def charger_une_seule_fois(chemin_base):
    """Une seule fois par démarrage et par base : si plusieurs exécutions arrivent en même
    temps, les autres attendent la fin du chargement. Sans effet si la base est déjà remplie."""
    return demo.charger_demo()


def preparer():
    """À appeler après core.init_db() : charge la démo si besoin et affiche le message."""
    if not lire_secret("demo_en_ligne"):
        return
    charger_une_seule_fois(core.DB_PATH)
    st.caption(":material/info: Démonstration : données fictives, effacées à chaque redémarrage. "
               f"La {core.NOM_PLATEFORME} et la {core.NOM_PORTAIL} ne communiquent pas ici : "
               "un dépôt fait sur l'une n'arrive pas sur l'autre.")
