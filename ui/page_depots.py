# ============================================================================
#  ui/page_depots.py — page « Dépôts reçus » de la Plateforme visa interne.
#  Les documents déposés par les entreprises sur la Plateforme visa partenaires arrivent ici. Rien
#  n'entre dans le Suivi des visas tant que le chef de projet n'a pas coché
#  les lignes et cliqué « Transférer » (ou « Refuser », avec un motif et un
#  brouillon de mail au chargé d'affaires). Un document non prévu (soumission
#  libre) devient, au transfert, une nouvelle position ou rejoint une position
#  existante du lot, au choix. Toute la logique : core, section 16.
# ============================================================================

import datetime as dt
import streamlit as st
import core
import calendrier
from ui.style import badge, badge_avis, entete, esc, sous_titre, tuiles
from ui.sidebar import MODE_SOLARES

MIME_EML = "message/rfc822"


def _date_fr(texte_iso, avec_heure=False):
    if not texte_iso:
        return ""
    if avec_heure:
        return dt.datetime.fromisoformat(texte_iso).strftime("%d/%m/%Y à %H:%M")
    return dt.date.fromisoformat(texte_iso[:10]).strftime("%d/%m/%Y")


def _niveau(echeance_iso):
    """Badge « à viser avant le » : en retard, échéance proche (≤ 2 j ouvrés) ou dans les temps."""
    restant = calendrier.jours_ouvres_restants(dt.date.today(), dt.date.fromisoformat(echeance_iso))
    if restant < 0:
        return "retard", "en retard"
    if restant <= 2:
        return "proche", "échéance proche"
    return "ok", "dans les temps"


def _chefs_du_portefeuille(ctx, choix):
    """Codes des chefs de projet du filtre (None = tous les projets).
    « Portefeuille » = celui du bandeau de la barre latérale (chef de projet ou
    équipe) : c'est aussi lui que compte la pastille du menu."""
    if choix == "Tous" or not ctx.get("portefeuille"):
        return None
    return ctx["portefeuille"]["chefs"]


# ---------- fenêtres de confirmation ----------

def _destination_libre(envoi, ligne):
    """Document non prévu : nouvelle position (trame proposée, modifiable) ou
    position existante du lot. Renvoie la destination pour core.transferer_lignes."""
    libre = ligne["position"]
    plan = libre["libre"] == core.LIBRE_PLAN
    existantes = {}
    for position in core.get_positions(envoi["projet"], envoi["lot_id"]):
        if core.est_plan(position) == plan:
            existantes[position["id"]] = core._libelle_position(position)
    nouvelle = f"Nouvelle position : {core._libelle_position(libre)}"
    choix = st.selectbox("Destination", [None] + list(existantes), key=f"dr_dest_{ligne['id']}",
                         format_func=lambda i: nouvelle if i is None else existantes[i],
                         help="Document déposé hors des lignes prévues : créez la position, ou "
                              "rattachez-le à une position existante du lot.")
    if choix is not None:
        return {"position_id": choix}
    if plan:
        return {"trame": core.TRAME_AUTRE}
    noms = core.trame_names()
    proposee = core.suggest_trame(libre["des"])
    trame = st.selectbox("Trame de la nouvelle position", noms, key=f"dr_trame_{ligne['id']}",
                         index=noms.index(proposee) if proposee in noms else 0,
                         help="Proposée d'après la désignation ; génère pièces attendues et points de contrôle.")
    return {"trame": trame}


@st.dialog("Transférer vers la Plateforme visa interne", width="large")
def _dialogue_transfert(envoi, lignes, traite_par):
    st.caption(f"{envoi['projet']} · {core.libelle_lot(envoi['lot'])} · {envoi['entreprise']}")
    destinations = {}
    for ligne in lignes:
        position = ligne["position"]
        if position.get("libre"):
            st.markdown(f"**{esc(core._libelle_position(position))}** {badge('Document non prévu', 'sb')}",
                        unsafe_allow_html=True)
            destinations[ligne["id"]] = _destination_libre(envoi, ligne)
            st.caption(f"{len(ligne['fichiers'])} fichier(s) ajouté(s) aux Documents ; rappel SB : à viser "
                       f"avant le {_date_fr(envoi['echeance'])}. Côté entreprise : « En cours d'examen ».")
            continue
        if ligne["visa_emis"]:
            effet = (f"Visa {ligne['visa_emis']['avis']} ({ligne['visa_emis']['libelle']}) archivé dans "
                     "« Anciens visas » ; nouveau visa ouvert.")
        else:
            effet = f"Pas de visa émis : les documents rejoignent le visa en cours (avis {position['avis']})."
        extraits = " (extraits aux pages indiquées)" if any(f["extrait"] for f in ligne["fichiers"]) else ""
        st.markdown(f"**{esc(core._libelle_position(position))}**  \n{esc(effet)}  \n"
                    f"{len(ligne['fichiers'])} fichier(s) ajouté(s) aux Documents{extraits} ; pièces "
                    f"{esc(', '.join(ligne['pieces']))} → « À vérifier ».  \n"
                    f"Rappel SB : à viser avant le **{_date_fr(envoi['echeance'])}** (date d'envoi + délai du lot). "
                    f"Côté entreprise : « En cours d'examen ».")
    a, b = st.columns(2)
    if a.button("Annuler", width="stretch"):
        st.rerun()
    if b.button("Confirmer le transfert", type="primary", width="stretch"):
        comptes_rendus = core.transferer_lignes([l["id"] for l in lignes], traite_par=traite_par,
                                                destinations=destinations)
        erreurs = [c for c in comptes_rendus if c["erreur"]]
        nouvelles = sum(1 for c in comptes_rendus if c["nouvelle"])
        st.session_state["_depots_message"] = (
            f"{len(comptes_rendus) - len(erreurs)} ligne(s) transférée(s) vers le Suivi des visas "
            f"({envoi['projet']}) ; l'entreprise voit « En cours d'examen »."
            + (f" {nouvelles} nouvelle(s) position(s) créée(s)." if nouvelles else "")
            + (f" {len(erreurs)} ligne(s) non transférée(s) : {erreurs[0]['erreur']}" if erreurs else ""))
        st.rerun()


@st.dialog("Refuser des documents", width="large")
def _dialogue_refus(envoi, lignes, traite_par):
    st.caption("Rien n'entre dans la Plateforme visa interne. L'entreprise voit la ligne revenir à « À fournir » avec votre "
               "motif, et peut renvoyer. Ce n'est pas un visa REF : aucun bordereau, aucun ancien visa.")
    for ligne in lignes:
        st.markdown(f"- {esc(core._libelle_position(ligne['position']))}")
    motif = st.text_area(f"Motif transmis à {envoi['entreprise']} (obligatoire)",
                         placeholder="Ex. : les pages indiquées ne correspondent pas à la CTA ; merci de préciser l'extrait.")
    st.caption(f"Un brouillon de mail (Outlook) sera prêt pour {envoi['contact_nom']} <{envoi['contact_email']}>, "
               "chargé d'affaires qui a fait l'envoi (personne en copie).")
    a, b = st.columns(2)
    if a.button("Annuler", width="stretch"):
        st.rerun()
    if b.button("Refuser et préparer le mail", type="primary", width="stretch"):
        ok, resultat = core.refuser_lignes([l["id"] for l in lignes], motif, traite_par)
        if not ok:
            st.error(resultat)
            return
        st.session_state["_depots_eml"] = {
            "octets": core.build_eml_refus([l["id"] for l in lignes]),
            "nom": f"refus_{envoi['projet']}_{envoi['entreprise']}.eml".replace(" ", "_").replace("'", ""),
            "contact": envoi["contact_nom"], "nombre": resultat}
        st.rerun()


# ---------- une carte par envoi ----------

def _fichier(fichier):
    """Nom, pages et téléchargement d'un fichier déposé (contenu lu seulement à la demande ;
    pour un extrait de dossier commun, seules les pages indiquées sont téléchargées)."""
    # Pages de l'extrait sous le nom (un nom long passe à la ligne sans chevaucher la pastille).
    pages = f"<br><span class='bdg b-proche'>p. {esc(fichier['extrait'])}</span>" if fichier["extrait"] else ""
    st.markdown(f"<div class='nom-fichier'>{esc(fichier['nom'])}{pages}</div>", unsafe_allow_html=True)
    cle = f"dr_f_{fichier['id']}"
    pret = st.session_state.get(cle)
    if pret:
        st.download_button("Télécharger", pret["octets"], file_name=pret["nom"], key=f"{cle}_dl",
                           icon=":material/download:", type="tertiary")
    elif st.button("Préparer le fichier", key=f"{cle}_prep", icon=":material/description:", type="tertiary"):
        nom, octets = core.lire_fichier_depot(fichier["id"])
        if fichier["commun"] and fichier["extrait"] and core.extension_fichier(nom) == "pdf":
            extrait, erreur = core.extraire_pages(octets, fichier["extrait"])
            if extrait is not None:
                nom, octets = nom[:-4] + f" p.{fichier['extrait']}.pdf", extrait
        st.session_state[cle] = {"nom": nom, "octets": octets}
        st.rerun()


def _carte_envoi(envoi, traite_par):
    genre, texte_niveau = _niveau(envoi["echeance"])
    with st.container(border=True):
        g, d = st.columns([1.2, 1])
        g.markdown(f"<div style='font-weight:700;font-size:15px'><span class='ent-code' style='font-size:12px;"
                   f"margin-right:8px'>{esc(envoi['projet'])}</span>{esc(core.libelle_lot(envoi['lot']))} · "
                   f"{esc(envoi['entreprise'])}</div><div style='font-size:12.5px;color:#5C5C5C;margin-top:4px'>"
                   f"{esc(envoi['projet_nom'])} · envoyé le {_date_fr(envoi['date_envoi'], avec_heure=True)}</div>"
                   f"<div style='font-size:12.5px;color:#5C5C5C'>Chargé d'affaires : <b>{esc(envoi['contact_nom'])}</b>"
                   f" · {esc(envoi['contact_email'])}</div>", unsafe_allow_html=True)
        echeance = badge(f"À viser avant le {_date_fr(envoi['echeance'])} · {texte_niveau}", genre, long=True)
        d.markdown(f"<div style='text-align:right'>{echeance}</div>", unsafe_allow_html=True)

        tailles = [0.3, 1.9, 1.5, 2.1, 2.1, 1.9]
        tete = st.columns(tailles)
        for colonne, titre in zip(tete[1:], ["Position", "Visa actuel", "Pièces déclarées",
                                              "Fichiers et pages", "Commentaire"]):
            colonne.markdown(f"<span class='t-lab'>{titre}</span>", unsafe_allow_html=True)
        cochees = []
        for ligne in envoi["lignes"]:
            position = ligne["position"]
            st.markdown("<hr class='rowsep'>", unsafe_allow_html=True)
            c0, c1, c2, c3, c4, c5 = st.columns(tailles)
            if c0.checkbox("Sélectionner", key=f"dr_sel_{ligne['id']}", label_visibility="collapsed"):
                cochees.append(ligne)
            non_prevu = f"<br>{badge('Document non prévu', 'sb')}" if position.get("libre") else ""
            c1.markdown(f"<b>{esc(position.get('pos') or '')}</b><br><span style='font-size:13px'>"
                        f"{esc(position['des'])}</span>{non_prevu}", unsafe_allow_html=True)
            if position.get("libre"):
                visa = ("<div style='font-size:12px;color:#5C5C5C'>Hors liste → à créer ou à rattacher "
                        "au transfert</div>")
            elif ligne["visa_emis"]:
                visa = (f"{badge_avis(ligne['visa_emis']['avis'])}<div style='font-size:12px;color:#5C5C5C;margin-top:4px'>"
                        f"<b>{esc(ligne['visa_emis']['libelle'])}</b> du {_date_fr(ligne['visa_emis']['date'])}"
                        " → passera dans « Anciens visas »</div>")
            else:
                visa = (f"{badge_avis(position['avis'])}<div style='font-size:12px;color:#5C5C5C;margin-top:4px'>"
                        "Visa en cours (non émis) → reçoit les documents</div>")
            c2.markdown(visa, unsafe_allow_html=True)
            # Seules les pièces déclarées par l'entreprise ; une pièce hors liste vient de « Autre ».
            attendues = [piece["n"] for piece in position["att"]]
            pastilles = []
            for nom in ligne["pieces"]:
                suffixe = "" if nom in attendues or position.get("libre") else " <i>(autre)</i>"
                pastilles.append(f"<span class='bdg long' style='background:#E3EEFB;color:#1D4F91'>{esc(nom)}{suffixe}</span>")
            if not pastilles:
                if position.get("libre") == core.LIBRE_CCTP:
                    texte = "Nature du document non précisée"
                else:
                    texte = "Plan / note de calcul"
                pastilles.append(f"<span style='font-size:12px;color:#5C5C5C'>{texte}</span>")
            c3.markdown(" ".join(pastilles), unsafe_allow_html=True)
            with c4:
                for fichier in ligne["fichiers"]:
                    _fichier(fichier)
            commentaire = f"<i>« {esc(ligne['commentaire'])} »</i>" if ligne["commentaire"] else "—"
            c5.markdown(f"<div style='font-size:13px'>{commentaire}</div>", unsafe_allow_html=True)

        st.markdown("<hr class='rowsep'>", unsafe_allow_html=True)
        p0, p1, p2 = st.columns([0.7, 1.4, 2.3], vertical_alignment="center")
        p0.caption(f"{len(cochees)} ligne(s) cochée(s) sur {len(envoi['lignes'])}")
        if p1.button("Refuser la sélection…", key=f"dr_refus_{envoi['id']}", disabled=not cochees,
                     width="stretch", icon=":material/block:"):
            _dialogue_refus(envoi, cochees, traite_par)
        if p2.button(f"Transférer vers la {core.NOM_PLATEFORME}", key=f"dr_transfert_{envoi['id']}", type="primary",
                     disabled=not cochees, width="stretch", icon=":material/move_to_inbox:"):
            _dialogue_transfert(envoi, cochees, traite_par)


# ---------- page ----------

def afficher(ctx):
    entete("", "Dépôts reçus des entreprises",
           [("", f"Documents déposés sur la {core.NOM_PORTAIL}, en attente d'acquisition sur la {core.NOM_PLATEFORME}. "
                 "Rien n'est créé tant que vous n'avez pas transféré : contrôlez pièces, fichiers et pages, "
                 "puis transférez.")])

    message = st.session_state.pop("_depots_message", None)
    if message:
        st.success(message, icon=":material/task_alt:")
    eml = st.session_state.get("_depots_eml")
    if eml:
        with st.container(border=True):
            a, b, c = st.columns([3, 1.3, 0.8], vertical_alignment="center")
            a.markdown(f"**{eml['nombre']} ligne(s) refusée(s).** Brouillon de mail prêt pour "
                       f"{esc(eml['contact'])} : téléchargez-le puis envoyez-le depuis Outlook.")
            b.download_button("Brouillon de mail", eml["octets"], file_name=eml["nom"], mime=MIME_EML,
                              icon=":material/mail:", type="primary", width="stretch")
            if c.button("Fermer", width="stretch"):
                st.session_state.pop("_depots_eml", None)
                st.rerun()

    f1, f2 = st.columns([1.4, 2], vertical_alignment="bottom")
    nom_portefeuille = ctx["portefeuille"]["nom"] if ctx.get("portefeuille") else "Portefeuille"
    if ctx.get("portefeuille") and ctx["portefeuille"]["mode"] == MODE_SOLARES:
        choix = "Tous"   # portefeuille « Solares » : déjà tous les projets
    else:
        choix = f1.segmented_control("Projets", [nom_portefeuille, "Tous"], default=nom_portefeuille,
                                     key=f"dr_portefeuille_{nom_portefeuille}") or nom_portefeuille
    chefs = _chefs_du_portefeuille(ctx, choix)
    envois = core.get_depots_en_attente(chefs)
    codes = sorted({e["projet"] for e in envois})
    projet = f2.selectbox("Projet", ["Tous les projets"] + codes, key="dr_projet")
    if projet != "Tous les projets":
        envois = [e for e in envois if e["projet"] == projet]

    nb_lignes = sum(len(e["lignes"]) for e in envois)
    echeances = sorted(e["echeance"] for e in envois)
    en_retard = sum(1 for e in envois if _niveau(e["echeance"])[0] == "retard")
    tuiles([
        {"lab": "Envois en attente", "val": len(envois), "sub": "envois d'entreprises"},
        {"lab": "Lignes à transférer", "val": nb_lignes, "sub": "positions / plans concernés"},
        {"lab": "Échéance la plus proche", "val": _date_fr(echeances[0]) if echeances else "—",
         "sub": "à viser avant le (date d'envoi + délai du lot)", "alerte": en_retard > 0},
        {"lab": "En retard", "val": en_retard, "alerte": en_retard > 0, "sub": "envois dont l'échéance est passée"},
    ])

    sous_titre(f"Documents déposés sur la {core.NOM_PORTAIL} en attente d'acquisition sur la {core.NOM_PLATEFORME}")
    if not envois:
        st.success("Aucun document en attente pour ce portefeuille.", icon=":material/task_alt:")
    for envoi in envois:
        _carte_envoi(envoi, ctx["chef_code"] or "")

    sous_titre("Derniers transferts et refus")
    journal = core.get_journal_depots(chefs)
    if not journal:
        st.caption("Aucun transfert ni refus pour l'instant.")
    for ligne in journal:
        if ligne["statut"] == core.LIGNE_TRANSFEREE:
            quoi = "Transféré"
            suite = ""
        else:
            quoi = "Refusé"
            suite = f" — « {esc(ligne['motif_refus'])} » — mail préparé pour {esc(ligne['contact_nom'])}"
        st.markdown(f"<div style='font-size:13px;margin:4px 0'><span style='color:#6B6B6B'>"
                    f"{_date_fr(ligne['date_traitement'], avec_heure=True)}</span> · <b>{quoi}</b> : "
                    f"{esc(ligne['projet'])} Lot {esc(ligne['lot_numero'] or '')} · {esc(ligne['libelle'])} "
                    f"({esc(ligne['entreprise'])}){suite}</div>", unsafe_allow_html=True)
