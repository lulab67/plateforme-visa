# ============================================================================
#  ui/sidebar.py — barre latérale : logo, bandeau « Portefeuille → Projet →
#  Lot » (choisi une fois, suivi par toutes les pages), menu principal,
#  déconnexion.
#  Renvoie un "contexte" (dictionnaire) partagé par toutes les pages.
# ============================================================================

import streamlit as st
import core
from ui.style import marque

# Pages du menu principal, avec leur icône. Le Dashboard vient en premier.
PAGES = {
    "Dashboard": ":material/space_dashboard:",
    "Suivi des visas": ":material/fact_check:",
    "Dépôts reçus": ":material/move_to_inbox:",
    "Base de données": ":material/table_view:",
    "Trames de visa": ":material/checklist:",
    "Administration": ":material/settings:",
}

MODE_CHEF = "Chef de projet"
MODE_EQUIPE = "Équipe"
MODE_SOLARES = "Solares"      # toute l'entreprise ; son Dashboard = délais de visa


def libelle_projet(projet):
    """Libellé d'un projet dans le sélecteur « Projet »."""
    return f"{projet['code']} — {projet['nom']}"


def libelle_lot_menu(lot):
    """Libellé d'un lot dans le sélecteur « Lot »."""
    return f"{core.libelle_lot(lot)} — {lot['entreprise'] or 'titulaire n.d.'}"


def demander_selection(projet=None, lot=None, page=None, chef_nom=None):
    """Demande à la barre latérale de se positionner sur un chef / un projet /
    un lot / une page au prochain affichage (un sélecteur déjà affiché ne peut
    pas être modifié dans le même tour : la demande est appliquée ici, au
    début du tour suivant)."""
    demande = st.session_state.setdefault("_selection_demandee", {})
    if chef_nom:
        demande["sel_mode"] = MODE_CHEF
        demande["sel_chef"] = chef_nom
    if projet:
        demande["sel_projet"] = libelle_projet(projet)
    if lot:
        demande["sel_lot"] = libelle_lot_menu(lot)
    if page:
        demande["page"] = page


def _codes_chefs_choisis(chefs, equipes):
    """Codes des chefs de projet du portefeuille choisi au tour précédent
    (pour la pastille « Dépôts reçus », calculée avant l'affichage des sélecteurs)."""
    if st.session_state.get("sel_mode") == MODE_SOLARES:
        return [chef["code"] for chef in chefs]
    if st.session_state.get("sel_mode") == MODE_EQUIPE:
        for equipe in equipes:
            if equipe["nom"] == st.session_state.get("sel_equipe"):
                return equipe["chefs"]
    for chef in chefs:
        if chef["nom"] == st.session_state.get("sel_chef"):
            return [chef["code"]]
    return [chefs[0]["code"]] if chefs else []


def _pastille_depots(nb_depots):
    """Pastille rouge du menu « Dépôts reçus », dessinée en CSS : le libellé du
    menu ne doit jamais changer, sinon Streamlit recrée le menu et revient à la
    première page."""
    if not nb_depots:
        return
    rang = list(PAGES.keys()).index("Dépôts reçus") + 1
    st.markdown(
        "<style>section[data-testid='stSidebar'] [data-testid='stRadioGroup'] > :nth-child"
        f"({rang}) label p::after{{content:'{nb_depots}';margin-left:10px;background:#E60116;"
        "color:#fff;border-radius:999px;font-size:12px;font-weight:700;padding:1px 9px;"
        "vertical-align:middle;}</style>", unsafe_allow_html=True)


def _choix_portefeuille(chefs, equipes):
    """Portefeuille = un chef de projet, une équipe ou tout Solares Bauen.
    Renvoie un dictionnaire {mode, nom, chefs (codes), projets, archives, equipe}."""
    modes = [MODE_CHEF, MODE_EQUIPE, MODE_SOLARES] if equipes else [MODE_CHEF, MODE_SOLARES]
    if st.session_state.get("sel_mode") not in modes:
        st.session_state["sel_mode"] = MODE_CHEF
    mode = st.segmented_control("Portefeuille", modes, key="sel_mode") or MODE_CHEF
    if mode == MODE_SOLARES:
        return {"mode": mode, "nom": "Solares Bauen", "chefs": [c["code"] for c in chefs], "equipe": None,
                "projets": core.get_projets(), "archives": core.get_projets(archive=1)}
    if mode == MODE_EQUIPE:
        noms = [e["nom"] for e in equipes]
        nom = st.selectbox("Équipe", noms, key="sel_equipe")
        equipe = next(e for e in equipes if e["nom"] == nom)
        return {"mode": mode, "nom": nom, "chefs": equipe["chefs"], "equipe": equipe,
                "projets": core.get_projets_team(equipe["id"]),
                "archives": core.get_projets_team(equipe["id"], archive=1)}
    noms = [c["nom"] for c in chefs]
    nom = st.selectbox("Chef de projet", noms, key="sel_chef")
    code = next(c["code"] for c in chefs if c["nom"] == nom)
    equipe = next((e for e in equipes if code in e["chefs"]), None)
    return {"mode": mode, "nom": nom, "chefs": [code], "equipe": equipe,
            "projets": core.get_projets(code), "archives": core.get_projets(code, archive=1)}


def barre_laterale(authenticator, app_version):
    """Construit la barre latérale et renvoie le contexte :
    {page, chefs, chef_nom, chef_code, proj_code, lot, portefeuille}.
    chef_code / chef_nom : le chef de projet choisi, ou (portefeuille
    « Équipe ») le chef du projet choisi."""
    chefs = core.get_chefs()
    equipes = core.get_teams()
    # Sélection demandée par une autre page au tour précédent (ex. projet créé).
    for cle, valeur in st.session_state.pop("_selection_demandee", {}).items():
        st.session_state[cle] = valeur
    nb_depots = core.nb_lignes_en_attente(_codes_chefs_choisis(chefs, equipes)) if chefs else 0
    _pastille_depots(nb_depots)

    portefeuille = None
    chef_nom = chef_code = proj_code = None
    lot = None
    with st.sidebar:
        st.markdown(marque("interne"), unsafe_allow_html=True)

        # --- Bandeau : portefeuille → projet → lot ---
        if chefs:
            portefeuille = _choix_portefeuille(chefs, equipes)
            projets = portefeuille["projets"]
            proj_map = {libelle_projet(p): p for p in projets}
            if proj_map:
                proj_label = st.selectbox("Projet", list(proj_map.keys()), key="sel_projet")
                projet = proj_map[proj_label]
                proj_code = projet["code"]
                lots = core.get_lots(proj_code)
                if lots:
                    lot_map = {libelle_lot_menu(l): l["id"] for l in lots}
                    lot_label = st.selectbox("Lot", list(lot_map.keys()), key="sel_lot")
                    lot = core.get_lot(lot_map[lot_label])
            else:
                st.caption("Aucun projet actif dans ce portefeuille.")
            # Chef de projet de référence (nouveau projet, filtres, interlocuteur).
            if portefeuille["mode"] == MODE_CHEF:
                chef_code = portefeuille["chefs"][0]
            elif proj_code:
                chef_code = projet["chef"]
            elif portefeuille["chefs"]:
                chef_code = portefeuille["chefs"][0]
            chef_nom = next((c["nom"] for c in chefs if c["code"] == chef_code), None)
        else:
            st.caption("Aucun projet : commencez par la page Administration.")

        # --- Menu principal ---
        st.markdown("<div class='sb-titre-nav'>Menu</div>", unsafe_allow_html=True)
        noms_pages = list(PAGES.keys())
        if not chefs:
            st.session_state["page"] = "Administration"
        page = st.radio("Menu", noms_pages, key="page", format_func=lambda nom: f"{PAGES[nom]} {nom}",
                        label_visibility="collapsed")

        st.divider()
        # Déconnexion et version sur une seule ligne : la barre reste courte.
        gauche, droite = st.columns([1.7, 1], vertical_alignment="center")
        with gauche:
            authenticator.logout("Se déconnecter", location="main")
        droite.markdown(f"<div class='sb-pied'>Version {app_version}</div>", unsafe_allow_html=True)

    return {"page": page, "chefs": chefs, "chef_nom": chef_nom, "chef_code": chef_code,
            "proj_code": proj_code, "lot": lot, "portefeuille": portefeuille}
