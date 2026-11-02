# ============================================================================
#  ui/page_portail.py — écrans de « Plateforme visa partenaires » (portail des entreprises).
#    - connexion : un compte par entreprise (identifiant + mot de passe) ;
#    - menu latéral : Positions CCTP/Matériel, Plans / notes de calcul, Historique des
#      envois, Chargés d'affaires ;
#    - en haut : choix du projet, puis du lot, et chargé d'affaires du lot
#      (modifiable ; Solares Bauen peut aussi le changer en Administration) ;
#    - tableau miroir : état, dernier visa, observations, motif de refus ;
#    - volet « Votre envoi » :
#        CCTP : un mode par envoi, « Un fichier commun à plusieurs positions »
#        (un gros PDF / ZIP, puis pages ou fichier du ZIP par ligne) ou
#        « Fichier(s) par position » ; pièces concernées + « Autre » (libellé
#        à préciser) ; « Tout cocher / Tout décocher » au-dessus du tableau ;
#        Plans / NdC : un seul fichier par ligne, sans pièce à choisir ;
#    - sous le volet : « document non prévu » (soumission libre, hors des lignes
#      du tableau), que Solares Bauen rattache à la réception.
#  Le brouillon d'envoi vit dans la session : il est perdu si la page est fermée
#  ou si l'on change de page.
#  Sécurité : l'identifiant de l'entreprise vient UNIQUEMENT de la connexion ;
#  core.enregistrer_envoi revérifie que le lot lui appartient.
# ============================================================================

import datetime as dt
import streamlit as st
import core
from ui import bordereaux as bordereaux_ui
from ui.style import badge_avis, entete, esc, marque, sous_titre, tuiles


# Pages du menu latéral, avec leur icône.
PAGE_CCTP = "Positions CCTP/Matériel"
PAGES = {
    PAGE_CCTP: ":material/fact_check:",
    "Plans / notes de calcul": ":material/architecture:",
    "Historique des envois": ":material/history:",
    "Chargés d'affaires": ":material/contacts:",
}

MODE_COMMUN = "Un fichier commun à plusieurs positions"
MODE_PAR_POSITION = "Fichier(s) par position"

CSS_PORTAIL = """<style>
.block-container{max-width:1680px;}
.p-etat{display:inline-block;border-radius:999px;padding:3px 10px;font-size:12px;font-weight:700;}
.e-afournir{background:#F1F1F1;color:#5C5C5C;} .e-transmis{background:#E3EEFB;color:#1D4F91;}
.e-examen{background:#FFF2D6;color:#8A5A00;} .e-vise{background:#E2F2E5;color:#1E5E2E;}
.p-sous{font-size:12px;color:#5C5C5C;margin-top:4px;}
.p-nouvelle{font-size:12px;color:#C00000;font-weight:700;margin-top:4px;}
.p-refus{font-size:12px;color:#C00000;background:#FDECEE;border-radius:6px;padding:4px 8px;margin-top:6px;}
.p-obs{font-size:12px;color:#3D3D3D;background:#F4F4F4;border-radius:6px;padding:4px 8px;margin-top:6px;}
.p-des{font-weight:600;font-size:14px;} .p-moe{font-size:12px;color:#5C5C5C;}
.p-pieces{display:flex;flex-wrap:wrap;gap:4px;margin-top:6px;}
.p-piece{display:inline-flex;align-items:center;gap:4px;font-size:12px;border-radius:12px;
  padding:3px 9px;background:#ECE6F6;color:#4B2E83;max-width:100%;overflow-wrap:anywhere;}
.p-piece svg{width:11px;height:11px;flex:none;}
.p-piece.ok{background:#E2F2E5;color:#1E5E2E;} .p-piece.bleu{background:#E3EEFB;color:#1D4F91;}
.p-pos{font-weight:700;font-size:14px;white-space:nowrap;}
/* Écran étroit : le volet d'envoi passe sous le tableau (à côté, le tableau serait illisible). */
@media (max-width:1100px){
  .st-key-p_corps > div > [data-testid="stHorizontalBlock"] > [data-testid="stColumn"]{min-width:100%;}
}
.p-ligne{margin:-8px 0 -4px 0;border:none;border-top:1px solid #EEEEEE;}
.st-key-p_dossier{border:1px solid #F2C9CE;background:#FFF6F6;border-radius:10px;
  padding:10px 16px 4px 16px;margin-bottom:12px;}
.p-dossier-titre{font-size:12px;font-weight:700;letter-spacing:.06em;color:#C00000;text-transform:uppercase;}
.p-bloc{font-size:12px;font-weight:700;letter-spacing:.06em;color:#C00000;text-transform:uppercase;margin:6px 0 2px 0;}
.st-key-p_envoyer button{min-height:58px;font-size:16px;}
</style>"""

# Écran de connexion : pas de barre latérale.
CSS_SANS_MENU = """<style>
section[data-testid="stSidebar"], button[data-testid="stExpandSidebarButton"],
div[data-testid="stSidebarCollapsedControl"]{display:none;}
</style>"""


def _date_fr(texte_iso):
    if not texte_iso:
        return ""
    return dt.date.fromisoformat(texte_iso[:10]).strftime("%d/%m/%Y")


# ---------- connexion ----------

def connexion():
    """Formulaire de connexion de l'entreprise. Arrête la page tant qu'elle
    n'est pas connectée ; renvoie ensuite sa fiche (dictionnaire)."""
    st.markdown(CSS_PORTAIL, unsafe_allow_html=True)
    entreprise_id = st.session_state.get("portail_entreprise_id")
    if entreprise_id:
        entreprise = core.get_entreprise(entreprise_id)
        if entreprise and entreprise["actif"]:
            return entreprise
        st.session_state.pop("portail_entreprise_id", None)

    st.markdown(CSS_SANS_MENU, unsafe_allow_html=True)
    _, milieu, _ = st.columns([1, 1.2, 1])
    with milieu:
        st.markdown(marque("partenaires"), unsafe_allow_html=True)
        with st.form("portail_connexion"):
            identifiant = st.text_input("Identifiant de l'entreprise")
            mot_de_passe = st.text_input("Mot de passe", type="password")
            valider = st.form_submit_button("Se connecter", type="primary", width="stretch")
        if valider:
            entreprise = core.connexion_entreprise(identifiant, mot_de_passe)
            if entreprise:
                st.session_state["portail_entreprise_id"] = entreprise["id"]
                st.rerun()
            # Blocage tenu par le serveur (core.minutes_de_blocage) : il résiste au rechargement.
            minutes = core.minutes_de_blocage(identifiant)
            if minutes:
                st.error(f"Trop de tentatives pour cet identifiant : réessayez dans {minutes} minute(s), "
                         "ou contactez votre interlocuteur Solares Bauen.", icon=":material/lock_clock:")
            else:
                # Message volontairement vague : on ne dit pas lequel des deux est faux.
                st.error("Identifiant ou mot de passe incorrect.")
        st.caption("Accès réservé aux entreprises titulaires de lots suivis par Solares Bauen. "
                   "Identifiants fournis par votre interlocuteur Solares Bauen.")
    st.stop()


def _deconnecter():
    for cle in list(st.session_state.keys()):
        if cle.startswith("p_") or cle.startswith("portail_"):
            del st.session_state[cle]


# ---------- tableau miroir ----------

def _etat_html(position):
    miroir = position["miroir"]
    etat = miroir["etat"]
    html = f"<span class='p-etat e-{etat}'>"
    if etat == core.ETAT_VISE:
        html += f"Visé le {_date_fr(miroir['date_visa'])}</span>"
    else:
        html += f"{esc(core.LIBELLES_ETAT[etat])}</span>"
    if etat == core.ETAT_TRANSMIS:
        html += f"<div class='p-sous'>le {_date_fr(miroir['depuis'])}</div>"
    if etat == core.ETAT_EXAMEN:
        html += f"<div class='p-sous'>depuis le {_date_fr(miroir['depuis'])}</div>"
    if etat == core.ETAT_VISE and miroir["avis"] in ("VAOB", "REF", "MQ"):
        html += "<div class='p-nouvelle'>Nouvelle soumission attendue</div>"
    refus = miroir["refus"]
    if refus:
        html += (f"<div class='p-refus'><b>Envoi du {_date_fr(refus['date_envoi'])} refusé le "
                 f"{_date_fr(refus['date'])}</b> (mail à {esc(refus['contact'])}) : {esc(refus['motif'])}</div>")
    return html


# Petites icônes des pièces : SVG à trait de 2,5 px (pas de glyphes Unicode).
ICONE_RECUE = ("<svg viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2.5' "
               "stroke-linecap='round' stroke-linejoin='round' aria-hidden='true'><path d='M5 12.5l4.5 4.5L19 7.5'/></svg>")
ICONE_TRANSMISE = ("<svg viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2.5' "
                   "stroke-linecap='round' stroke-linejoin='round' aria-hidden='true'><path d='M7 17L17 7M9 7h8v8'/></svg>")


def _pieces_html(position):
    """Pièces attendues : reçue et validée (vert), transmise à vérifier (bleu), à fournir (violet)."""
    morceaux = []
    for piece in position["att"]:
        if piece["s"] == "Reçu et ok":
            morceaux.append(f"<span class='p-piece ok' title='Pièce reçue'>{ICONE_RECUE}{esc(piece['n'])}</span>")
        elif piece["s"] == core.STATUT_A_VERIFIER:
            morceaux.append(f"<span class='p-piece bleu' title='Pièce transmise'>{ICONE_TRANSMISE}{esc(piece['n'])}</span>")
        else:
            morceaux.append(f"<span class='p-piece' title='Pièce à fournir'>{esc(piece['n'])}</span>")
    return f"<div class='p-pieces'>{''.join(morceaux)}</div>"


def _dernier_visa(position):
    miroir = position["miroir"]
    if not miroir["avis"]:
        st.markdown("<span class='p-sous'>—</span>", unsafe_allow_html=True)
        return
    html = f"{badge_avis(miroir['avis'])} <span class='p-sous'>{_date_fr(miroir['date_visa'])}</span>"
    if miroir["observations"]:
        html += f"<div class='p-obs'>{esc(miroir['observations'])}</div>"
    st.markdown(html, unsafe_allow_html=True)
    bordereau = core.get_bordereau(miroir["bordereau_id"]) if miroir["bordereau_id"] else None
    if bordereau is None:
        return
    prefixe = f"p{position['id']}"
    if not bordereaux_ui.boutons_telechargement(bordereau, prefixe=prefixe):
        if st.button(miroir["bordereau"], key=f"{prefixe}_bx", icon=":material/description:",
                     type="tertiary"):
            with st.spinner("Préparation du bordereau…"):
                bordereaux_ui.preparer_exports(bordereau)
            st.rerun()


def _raison_blocage(position):
    etat = position["miroir"]["etat"]
    if etat == core.ETAT_TRANSMIS:
        return "Un envoi est déjà en attente de prise en charge par Solares Bauen."
    if etat == core.ETAT_EXAMEN:
        return "Examen en cours : vous pourrez déposer une nouvelle version après le visa."
    return "Cochez pour ajouter cette ligne à votre envoi."


def _cocher_tout(positions, cle, valeur):
    """Coche (ou décoche) d'un coup toutes les lignes qui peuvent être envoyées.
    Appelé par les boutons, avant l'affichage des cases (on_click)."""
    for position in positions:
        if core.peut_deposer(position["miroir"]):
            st.session_state[f"{cle}_{position['id']}"] = valeur


def _tableau(positions, cle, plans=False):
    """Lignes du tableau miroir ; la case à cocher ajoute la ligne à l'envoi."""
    if not positions:
        st.caption("Aucune ligne.")
        return
    nb_possibles = sum(1 for p in positions if core.peut_deposer(p["miroir"]))
    b1, b2 = st.columns(2)
    b1.button("Tout cocher", key=f"{cle}_tout", icon=":material/check_box:", width="stretch",
              on_click=_cocher_tout, args=(positions, cle, True), disabled=nb_possibles == 0)
    b2.button("Tout décocher", key=f"{cle}_rien", icon=":material/check_box_outline_blank:",
              width="stretch", on_click=_cocher_tout, args=(positions, cle, False),
              disabled=nb_possibles == 0)
    titres = ["Pos.", "Désignation et pièces attendues", "État", "Dernier visa"]
    if plans:
        titres = ["", "Plan / note de calcul", "État", "Dernier visa"]
    tailles = [0.35, 0.95, 3.0, 1.7, 1.9]
    tete = st.columns(tailles)
    for colonne, titre in zip(tete[1:], titres):
        colonne.markdown(f"<span class='t-lab'>{titre}</span>", unsafe_allow_html=True)
    for position in positions:
        st.markdown("<hr class='p-ligne'>", unsafe_allow_html=True)
        c0, c1, c2, c3, c4 = st.columns(tailles)
        c0.checkbox("Ajouter à l'envoi", key=f"{cle}_{position['id']}", label_visibility="collapsed",
                    disabled=not core.peut_deposer(position["miroir"]), help=_raison_blocage(position))
        c1.markdown(f"<span class='p-pos'>{esc(position.get('pos') or '')}</span>", unsafe_allow_html=True)
        moe = ""
        if position.get("marque_moe") or position.get("type_moe"):
            moe = (f"<div class='p-moe'>Prescription MOE : {esc(position.get('marque_moe'))} "
                   f"{esc(position.get('type_moe'))}</div>")
        pieces = "" if plans else _pieces_html(position)
        c2.markdown(f"<div class='p-des'>{esc(position['des'])}</div>{moe}{pieces}", unsafe_allow_html=True)
        c3.markdown(_etat_html(position), unsafe_allow_html=True)
        with c4:
            _dernier_visa(position)


# ---------- volet « Votre envoi » ----------

def _sens_extrait(nom):
    """Ce que veut dire la case à côté d'un fichier commun : pages, fichier du ZIP, ou rien."""
    extension = core.extension_fichier(nom)
    if extension == "pdf":
        return "Pages, ex. 12-30", True
    if extension == "zip":
        return "Fichier dans le ZIP", True
    return "Fichier entier", False


def _pieces_ligne(position, cle):
    """Pièces concernées (pièces attendues + « Autre ») et libellé de « Autre ».
    Renvoie (pièces, libellé, problèmes)."""
    pid = position["id"]
    noms = []
    for piece in position["att"]:
        noms.append(piece["n"])
    noms.append(core.PIECE_AUTRE)
    pieces = st.pills("Pièces concernées", noms, selection_mode="multi", key=f"{cle}_pieces_{pid}") or []
    autre = ""
    problemes = []
    if core.PIECE_AUTRE in pieces:
        autre = st.text_input("Pièce « Autre » : laquelle ? *", key=f"{cle}_autre_{pid}",
                              placeholder="ex. PV d'essai en usine")
        if not autre.strip():
            problemes.append("préciser la pièce « Autre »")
    if not pieces:
        problemes.insert(0, "choisir au moins une pièce")
    return pieces, autre, problemes


def _titre_ligne(position):
    reference = position.get("pos") or ""
    if reference:
        return f"**{esc(reference)}** — {esc(position['des'])}"
    return f"**{esc(position['des'])}**"


def _a_completer(problemes):
    if problemes:
        st.markdown(f"<div class='p-refus'>À compléter : {esc(', '.join(problemes))}.</div>",
                    unsafe_allow_html=True)


def _ligne_commun(position, cle, communs):
    """Ligne cochée en mode « Fichier commun » : pièces, fichier du dossier (si
    plusieurs), pages ou fichier du ZIP, commentaire. Renvoie (ligne, problèmes)."""
    pid = position["id"]
    with st.container(border=True):
        st.markdown(_titre_ligne(position))
        pieces, autre, problemes = _pieces_ligne(position, cle)
        noms = list(communs.keys())
        a, b = st.columns([1.6, 1.2])
        nom = noms[0] if noms else ""
        if len(noms) > 1:
            nom = a.selectbox("Fichier commun concerné", [""] + noms, key=f"{cle}_extf_{pid}",
                              format_func=lambda n: n or "— choisir —")
        elif nom:
            a.markdown(f"<div class='p-sous' style='margin-top:30px'>Dans « {esc(nom)} »</div>",
                       unsafe_allow_html=True)
        aide, actif = _sens_extrait(nom) if nom else ("Pages, ex. 12-30", True)
        texte = b.text_input(aide, key=f"{cle}_extp_{pid}", placeholder="ex. 12-30" if actif else "",
                             disabled=not actif)
        commentaire = st.text_input("Commentaire pour Solares Bauen (facultatif)", key=f"{cle}_comm_{pid}")
        extraits = []
        if nom:
            extraits.append({"nom": nom, "extrait": texte if actif else ""})
        else:
            problemes.append("choisir le fichier commun concerné")
        _a_completer(problemes)
    ligne = {"position_id": pid, "pieces": pieces, "autre": autre, "commentaire": commentaire,
             "fichiers": [], "extraits": extraits}
    return ligne, problemes


def _ligne_par_position(position, cle):
    """Ligne cochée en mode « Fichier(s) par position » : pièces, fichiers,
    commentaire. Renvoie (ligne, problèmes)."""
    pid = position["id"]
    with st.container(border=True):
        st.markdown(_titre_ligne(position))
        pieces, autre, problemes = _pieces_ligne(position, cle)
        envoyes = st.file_uploader("Fichier(s) : PDF, ZIP, Excel, Word, DWG, images…", accept_multiple_files=True,
                                   key=f"{cle}_fichiers_{pid}")
        fichiers = []
        for envoye in envoyes or []:
            fichiers.append({"nom": envoye.name, "contenu": envoye.getvalue()})
        if not fichiers:
            problemes.append("déposer le fichier")
        commentaire = st.text_input("Commentaire pour Solares Bauen (facultatif)", key=f"{cle}_comm_{pid}")
        _a_completer(problemes)
    ligne = {"position_id": pid, "pieces": pieces, "autre": autre, "commentaire": commentaire,
             "fichiers": fichiers, "extraits": []}
    return ligne, problemes


def _ligne_plan(position, cle):
    """Plan / note de calcul coché : un seul fichier (PDF ou ZIP) + commentaire."""
    pid = position["id"]
    problemes = []
    with st.container(border=True):
        st.markdown(_titre_ligne(position))
        envoye = st.file_uploader("Fichier du plan / de la note (PDF, ZIP…)", key=f"{cle}_fichier_{pid}")
        fichiers = []
        if envoye is not None:
            fichiers.append({"nom": envoye.name, "contenu": envoye.getvalue()})
        else:
            problemes.append("déposer le fichier")
        commentaire = st.text_input("Commentaire pour Solares Bauen (facultatif)", key=f"{cle}_comm_{pid}")
        _a_completer(problemes)
    ligne = {"position_id": pid, "pieces": [], "commentaire": commentaire,
             "fichiers": fichiers, "extraits": []}
    return ligne, problemes


def _bouton_envoi(entreprise, lot, lignes, communs, problemes, contact):
    """Liste de ce qui manque, puis gros bouton d'envoi."""
    if problemes:
        st.markdown("<div class='t-lab'>Avant d'envoyer</div>" + "".join(
            f"<div class='p-refus' style='margin-top:4px'>{esc(p)}</div>" for p in problemes),
            unsafe_allow_html=True)
    else:
        st.success("Tout est complet : vous pouvez envoyer.", icon=":material/check_circle:")
    if st.button("Envoyer vos documents à viser à Solares Bauen", type="primary", key="p_envoyer",
                 width="stretch", disabled=bool(problemes)):
        with st.spinner("Envoi en cours…"):
            ok, resultat = core.enregistrer_envoi(entreprise["id"], lot["id"], lignes, communs)
        if ok:
            st.session_state["p_message"] = (
                f"{len(lignes)} ligne(s) transmise(s) à Solares Bauen. L'état passera à « En cours "
                f"d'examen » dès la prise en charge ; en cas de refus, {core.nom_avec_civilite(contact)} recevra un mail.")
            st.session_state["p_envoi_n"] = st.session_state.get("p_envoi_n", 0) + 1   # vide l'envoi
            st.rerun()
        for erreur in resultat:
            st.error(erreur)
    st.caption("Tant que l'envoi n'est pas fait, rien n'est transmis. Un envoi non terminé est perdu "
               "si vous fermez la page ou changez de page.")


def _debut_volet(lot, cochees, consigne):
    """Titre du volet ; renvoie False (et affiche la consigne) si rien n'est coché."""
    st.markdown("### Votre envoi")
    if not cochees:
        st.caption(consigne)
        return False
    st.caption(f"{len(cochees)} ligne(s) sélectionnée(s) · {core.libelle_lot(lot)}")
    return True


def _volet_cctp(entreprise, lot, positions, cle, contact):
    """Volet de droite de la page Positions CCTP/Matériel : un mode par envoi."""
    with st.container(border=True):
        cochees = [p for p in positions if st.session_state.get(f"{cle}_sel_{p['id']}")]
        if not _debut_volet(lot, cochees, "Cochez une ou plusieurs lignes du tableau, puis choisissez "
                                          "comment vous transmettez vos documents."):
            return
        mode = st.segmented_control("Mode d'envoi", [MODE_COMMUN, MODE_PAR_POSITION], key=f"{cle}_mode",
                                    default=MODE_PAR_POSITION if len(cochees) == 1 else MODE_COMMUN)
        if mode is None:
            st.caption("Choisissez le mode d'envoi.")
            return
        lignes = []
        problemes = []
        if contact is None:
            problemes.append("Désigner le chargé d'affaires du lot (en haut de la page)")
        communs = {}
        if mode == MODE_COMMUN:
            st.markdown("<div class='p-bloc'>1. Fichier commun</div>", unsafe_allow_html=True)
            envoyes = st.file_uploader("PDF ou ZIP qui couvre plusieurs positions", accept_multiple_files=True,
                                       key=f"{cle}_commun")
            for envoye in envoyes or []:
                communs[envoye.name] = {"nom": envoye.name, "contenu": envoye.getvalue()}
            st.markdown("<div class='p-bloc'>2. Positions concernées : ce que contient le fichier pour chacune</div>",
                        unsafe_allow_html=True)
            if not communs:
                problemes.append("Déposer le fichier commun")
        for position in cochees:
            if mode == MODE_COMMUN:
                ligne, manques = _ligne_commun(position, cle, communs)
            else:
                ligne, manques = _ligne_par_position(position, cle)
            lignes.append(ligne)
            for manque in manques:
                problemes.append(f"{core._libelle_position(position)} : {manque}")
        for nom in communs:
            if not any(x["nom"] == nom for ligne in lignes for x in ligne["extraits"]):
                problemes.append(f"« {nom} » n'est rattaché à aucune ligne")
        _bouton_envoi(entreprise, lot, lignes, list(communs.values()), problemes, contact)


def _volet_plans(entreprise, lot, positions, cle, contact):
    """Volet de droite de la page Plans / notes de calcul : un fichier par ligne."""
    with st.container(border=True):
        cochees = [p for p in positions if st.session_state.get(f"{cle}_sel_{p['id']}")]
        if not _debut_volet(lot, cochees, "Cochez un ou plusieurs plans / notes de calcul, puis déposez "
                                          "un fichier (PDF ou ZIP) pour chacun."):
            return
        lignes = []
        problemes = []
        if contact is None:
            problemes.append("Désigner le chargé d'affaires du lot (en haut de la page)")
        for position in cochees:
            ligne, manques = _ligne_plan(position, cle)
            lignes.append(ligne)
            for manque in manques:
                problemes.append(f"{core._libelle_position(position)} : {manque}")
        _bouton_envoi(entreprise, lot, lignes, [], problemes, contact)


def _soumission_libre(entreprise, lot, contact, plans):
    """Document non prévu dans la liste : un envoi d'une seule ligne, sans
    position. Solares Bauen décide à la réception (nouvelle position ou position
    existante). Clé propre : envoyer ce document ne vide pas « Votre envoi »."""
    cle = f"p_{lot['id']}_{'plans' if plans else 'cctp'}_libre_{st.session_state.get('p_libre_n', 0)}"
    if plans:
        titre = "Déposer un plan ou une note de calcul non prévu(e)"
    else:
        titre = "Déposer un document non prévu dans la liste"
    with st.expander(titre, icon=":material/add_circle:"):
        st.caption("Pour un document qui ne correspond à aucune ligne du tableau. Solares Bauen l'examine à "
                   "réception : il crée la position, ou rattache le document à une position existante.")
        reference = ""
        nature = ""
        if plans:
            designation = st.text_input("Désignation *", key=f"{cle}_des",
                                        placeholder="ex. Plan de réservations R+1 — indice A")
            envoye = st.file_uploader("Fichier (PDF, ZIP…) *", key=f"{cle}_fichier")
            envoyes = [envoye] if envoye is not None else []
        else:
            a, b = st.columns([1.4, 2.4])
            reference = a.text_input("Réf. CCTP *", key=f"{cle}_ref", placeholder="ex. 3.4.2",
                                     help="Référence de la position au CCTP ; « HC » si le matériel est hors CCTP.")
            designation = b.text_input("Désignation *", key=f"{cle}_des",
                                       placeholder="ex. Pompe de relevage du local technique")
            nature = st.text_input("Nature du document (facultatif)", key=f"{cle}_nature",
                                   placeholder="ex. Fiche technique")
            envoyes = st.file_uploader("Fichier(s) *", accept_multiple_files=True, key=f"{cle}_fichiers") or []
        commentaire = st.text_input("Commentaire pour Solares Bauen (facultatif)", key=f"{cle}_comm")
        problemes = []
        if contact is None:
            problemes.append("désigner le chargé d'affaires du lot (en haut de la page)")
        if not plans and not reference.strip():
            problemes.append("indiquer la référence CCTP (« HC » si hors CCTP)")
        if not designation.strip():
            problemes.append("indiquer la désignation")
        if not envoyes:
            problemes.append("déposer le fichier")
        ligne = {"libre": {"type": core.LIBRE_PLAN if plans else core.LIBRE_CCTP,
                           "pos": reference, "des": designation},
                 "pieces": [nature.strip()] if nature.strip() else [], "commentaire": commentaire,
                 "fichiers": [{"nom": f.name, "contenu": f.getvalue()} for f in envoyes], "extraits": []}
        if problemes:
            st.caption("À compléter : " + " ; ".join(problemes) + ".")
        if st.button("Envoyer ce document à Solares Bauen", key=f"{cle}_envoyer", type="primary",
                     icon=":material/send:", disabled=bool(problemes)):
            with st.spinner("Envoi en cours…"):
                ok, resultat = core.enregistrer_envoi(entreprise["id"], lot["id"], [ligne])
            if ok:
                st.session_state["p_message"] = (
                    f"« {designation.strip()} » transmis à Solares Bauen (document non prévu) : à suivre dans "
                    "« Historique des envois ». Il apparaîtra dans le tableau une fois pris en charge.")
                st.session_state["p_libre_n"] = st.session_state.get("p_libre_n", 0) + 1
                st.rerun()
            for erreur in resultat:
                st.error(erreur)


# ---------- historique et contacts ----------

LIBELLES_LIGNE = {core.LIGNE_TRANSMISE: "Document transmis à Solares Bauen",
                  core.LIGNE_TRANSFEREE: "Pris en charge (en cours d'examen)",
                  core.LIGNE_REFUSEE: "Refusé par Solares Bauen"}


def _historique(entreprise, lot):
    sous_titre(f"Mes envois sur ce lot — {core.libelle_lot(lot)}")
    envois = core.get_historique_envois(entreprise["id"], lot["id"])
    if not envois:
        st.caption("Aucun envoi pour l'instant.")
    for envoi in envois:
        motif = f" — « {esc(envoi['motif_refus'])} »" if envoi["statut"] == core.LIGNE_REFUSEE else ""
        st.markdown(f"<div style='font-size:13px;margin:6px 0'><span class='p-sous'>"
                    f"{dt.datetime.fromisoformat(envoi['date_envoi']).strftime('%d/%m/%Y à %H:%M')} · "
                    f"{esc(envoi['contact_nom'])}</span><br>{esc(envoi['libelle'])} — {envoi['nb_fichiers']} "
                    f"fichier(s) · <b>{LIBELLES_LIGNE[envoi['statut']]}</b>{motif}</div>",
                    unsafe_allow_html=True)


def _champs_contact(prefixe, contact=None):
    """Champs civilité / nom / mail / téléphone d'un contact (dans un formulaire)."""
    contact = contact or {}
    civilites = ["M.", "Mme"]
    a, b = st.columns([1, 3])
    civ = a.selectbox("Civilité", civilites, key=f"{prefixe}_civ",
                      index=civilites.index(contact.get("civ")) if contact.get("civ") in civilites else 0)
    nom = b.text_input("Prénom NOM *", value=contact.get("nom") or "", key=f"{prefixe}_nom",
                       placeholder="ex. Thomas BERNARD")
    c, d = st.columns(2)
    email = c.text_input("Mail *", value=contact.get("email") or "", key=f"{prefixe}_mail",
                         placeholder="ex. t.bernard@entreprise.fr")
    telephone = d.text_input("Téléphone", value=contact.get("telephone") or "", key=f"{prefixe}_tel")
    return civ, nom, email, telephone


def _contacts(entreprise):
    sous_titre("Chargés d'affaires / contacts")
    st.caption("Chaque lot a son chargé d'affaires (choisi en haut de la page) : c'est l'interlocuteur du "
               "bordereau de visa, et Solares Bauen lui écrit en cas de refus d'un dépôt. "
               "Solares Bauen voit et peut modifier cette liste.")
    lots = core.get_lots_entreprise(entreprise["id"])
    for contact in core.get_contacts(entreprise["id"]):
        ses_lots = [f"{l['projet']} · {core.libelle_lot(l)}" for l in lots if l.get("contact_id") == contact["id"]]
        with st.container(border=True):
            a, b = st.columns([4, 1.2], vertical_alignment="center")
            telephone = f" · {esc(contact['telephone'])}" if contact.get("telephone") else ""
            lots_html = (f"<div class='p-sous'>Chargé d'affaires de : {esc(', '.join(ses_lots))}</div>"
                         if ses_lots else "<div class='p-sous'>Chargé d'affaires d'aucun lot</div>")
            nom = core.nom_avec_civilite(contact)
            a.markdown(f"<div style='font-size:14px'><b>{esc(nom)}</b> · {esc(contact['email'])}{telephone}</div>"
                       f"{lots_html}", unsafe_allow_html=True)
            with b.popover("Modifier", icon=":material/edit:", width="stretch"):
                with st.form(f"p_modif_{contact['id']}"):
                    civ, nom, email, telephone = _champs_contact(f"p_mc{contact['id']}", contact)
                    if st.form_submit_button("Enregistrer", type="primary"):
                        ok, message = core.update_contact(contact["id"], nom, email, civ, telephone)
                        if ok:
                            st.rerun()
                        st.error(message)
                if st.button("Retirer de la liste", key=f"p_retirer_{contact['id']}", icon=":material/delete:"):
                    core.retirer_contact(contact["id"])
                    st.rerun()
    st.markdown("<div class='t-lab' style='margin-top:12px'>Ajouter un contact</div>", unsafe_allow_html=True)
    with st.form("p_nouveau_contact", clear_on_submit=True):
        civ, nom, email, telephone = _champs_contact("p_nc")
        if st.form_submit_button("Ajouter", type="primary"):
            ok, resultat = core.add_contact(entreprise["id"], nom, email, civ, telephone)
            if ok:
                st.rerun()
            st.error(resultat)


# ---------- haut de page : dossier choisi ----------

def _menu(entreprise, version):
    """Barre latérale : marque, menu, entreprise connectée. Renvoie la page choisie."""
    with st.sidebar:
        st.markdown(marque("partenaires"), unsafe_allow_html=True)
        page = st.radio("Menu", list(PAGES.keys()), key="p_page", label_visibility="collapsed",
                        format_func=lambda nom: f"{PAGES[nom]} {nom}")
        st.divider()
        st.markdown(f"<div style='font-size:13px;color:#5C5C5C'>Connecté :<br>"
                    f"<b style='color:#1E1E1E'>{esc(entreprise['nom'])}</b></div>", unsafe_allow_html=True)
        if st.button("Se déconnecter", key="p_deconnexion", icon=":material/logout:"):
            _deconnecter()
            st.rerun()
        st.caption(f"Version {version}")
    return page


def _choix_dossier(lots):
    """Bandeau du dossier : projet, puis lot (deux listes séparées). Renvoie le lot."""
    projets = {}
    for lot in lots:
        projets[lot["projet"]] = f"{lot['projet']} · {lot['projet_nom']}"
    boite = st.container(key="p_dossier")
    boite.markdown("<div class='p-dossier-titre'>Dossier</div>", unsafe_allow_html=True)
    a, b, c = boite.columns([1.4, 1.4, 1.4])
    projet = a.selectbox("Projet", list(projets.keys()), format_func=lambda code: projets[code], key="p_projet")
    lots_du_projet = [l for l in lots if l["projet"] == projet]
    libelles = {}
    for lot in lots_du_projet:
        libelles[lot["id"]] = core.libelle_lot(lot)
    lot_id = b.selectbox("Lot", list(libelles.keys()), format_func=lambda i: libelles[i], key=f"p_lot_{projet}")
    lot = next(l for l in lots_du_projet if l["id"] == lot_id)
    return lot, c


def _choix_charge_affaires(entreprise, lot, colonne):
    """Chargé d'affaires du lot (modifiable ici et par Solares Bauen). Renvoie le contact ou None."""
    contact = core.get_contact_lot(lot["id"])
    contacts = core.get_contacts(entreprise["id"])
    noms = {}
    for candidat in contacts:
        noms[candidat["id"]] = f"{core.nom_avec_civilite(candidat)} · {candidat['email']}"
    options = [None] + list(noms.keys())
    actuel = contact["id"] if contact else None
    # La clé suit la valeur enregistrée : une modification faite par Solares Bauen est reprise.
    choix = colonne.selectbox("Chargé d'affaires du lot *", options, index=options.index(actuel),
                              format_func=lambda i: noms.get(i, "— à désigner —"),
                              key=f"p_ca_{lot['id']}_{actuel}",
                              help="Interlocuteur du bordereau de visa ; reçoit le mail en cas de refus. "
                                   "Nouveau contact : page « Chargés d'affaires ».")
    if choix != actuel:
        core.set_lot_contact(lot["id"], choix)
        st.rerun()
    return contact


# ---------- page ----------

def afficher(entreprise, version):
    page = _menu(entreprise, version)
    if page == "Chargés d'affaires":
        _contacts(entreprise)
        return

    lots = core.get_lots_entreprise(entreprise["id"])
    if not lots:
        st.info("Aucun lot en cours pour votre entreprise. Contactez votre interlocuteur Solares Bauen.")
        return
    lot, colonne_ca = _choix_dossier(lots)
    contact = _choix_charge_affaires(entreprise, lot, colonne_ca)

    message = st.session_state.pop("p_message", None)
    if message:
        st.success(message, icon=":material/send:")

    chef = core.get_chef(lot["projet_chef"]) or {}
    entete(lot["projet"], lot["projet_nom"],
           [("Maître d'ouvrage", lot.get("projet_moa")), ("Lot", core.libelle_lot(lot)),
            ("Votre chargé d'affaires", core.nom_avec_civilite(contact) if contact else "à désigner"),
            ("Votre interlocuteur Solares Bauen", chef.get("nom"))], fil=page)

    if page == "Historique des envois":
        _historique(entreprise, lot)
        return

    positions = core.get_miroir_lot(lot["id"])
    if page == PAGE_CCTP:
        positions = [p for p in positions if not core.est_plan(p)]
    else:
        positions = [p for p in positions if core.est_plan(p)]
    nombre = {}
    for etat in core.LIBELLES_ETAT:
        nombre[etat] = sum(1 for p in positions if p["miroir"]["etat"] == etat)
    tuiles([
        {"lab": "À fournir", "val": nombre[core.ETAT_A_FOURNIR], "sub": "aucun document en cours"},
        {"lab": "Transmis", "val": nombre[core.ETAT_TRANSMIS], "sub": "en attente de prise en charge"},
        {"lab": "En cours d'examen", "val": nombre[core.ETAT_EXAMEN], "sub": "visa en préparation"},
        {"lab": "Visés", "val": nombre[core.ETAT_VISE], "sub": "voir le dernier bordereau"},
    ])

    # Clé de l'envoi en cours : change avec le lot, la page et après chaque envoi (tout est alors vidé).
    plans = page != PAGE_CCTP
    cle = f"p_{lot['id']}_{'plans' if plans else 'cctp'}_{st.session_state.get('p_envoi_n', 0)}"
    with st.container(key="p_corps"):
        gauche, droite = st.columns([1.5, 1], gap="medium")
        with gauche:
            _tableau(positions, f"{cle}_sel", plans=plans)
        with droite:
            if plans:
                _volet_plans(entreprise, lot, positions, cle, contact)
            else:
                _volet_cctp(entreprise, lot, positions, cle, contact)
            _soumission_libre(entreprise, lot, contact, plans)
