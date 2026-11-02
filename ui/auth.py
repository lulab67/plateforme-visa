# ============================================================================
#  ui/auth.py — connexion (login partagé défini dans config.yaml).
#  Sera remplacé par la connexion Microsoft 365 à l'itération 1.
# ============================================================================

import os
import streamlit as st
import yaml
from yaml.loader import SafeLoader
import streamlit_authenticator as stauth
from ui.style import marque
from ui.demo_en_ligne import lire_secret

HERE = os.path.dirname(os.path.dirname(__file__))
# Libellés du formulaire de la bibliothèque de connexion (anglais par défaut).
CHAMPS_CONNEXION = {"Form name": "Connexion", "Username": "Identifiant",
                    "Password": "Mot de passe", "Login": "Se connecter"}


def _config_connexion():
    """Compte et cookie : les secrets Streamlit en ligne (config.yaml n'est pas publié),
    sinon config.yaml. La connexion a besoin de copies modifiables, d'où to_dict()."""
    if lire_secret("credentials"):
        return st.secrets.to_dict()
    with open(os.path.join(HERE, "config.yaml")) as f:
        return yaml.load(f, Loader=SafeLoader)


def authentifier():
    """Affiche le formulaire de connexion et arrête la page tant que
    l'utilisateur n'est pas connecté. Renvoie l'objet authenticator
    (utilisé ensuite pour le bouton de déconnexion)."""
    config = _config_connexion()
    authenticator = stauth.Authenticate(config["credentials"], config["cookie"]["name"],
                                        config["cookie"]["key"], config["cookie"]["expiry_days"])
    # Logo et formulaire centrés, comme sur la Plateforme visa partenaires ; effacés dès que la connexion est faite
    # (y compris la reconnexion automatique par cookie).
    zone_connexion = st.empty()
    with zone_connexion.container():
        _, milieu, _ = st.columns([1, 1.2, 1])
        with milieu:
            if not st.session_state.get("authentication_status"):
                st.markdown(marque("interne"), unsafe_allow_html=True)
            authenticator.login(location="main", fields=CHAMPS_CONNEXION)
    status = st.session_state.get("authentication_status")
    if status:
        zone_connexion.empty()
    if status is False: milieu.error("Identifiants incorrects."); st.stop()
    elif status is None: milieu.info("Connectez-vous (login partagé — voir README)."); st.stop()
    return authenticator
