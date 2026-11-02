# ============================================================================
#  ui/page_admin.py — page « Administration » : projets, lots, entreprises
#  (fiche, compte Plateforme visa partenaires, chargés d'affaires), chefs de projet, équipes,
#  import DPGF, archivage et suppressions.
#  Toutes les données saisies ici (MOA, lot, titulaire, interlocuteurs) sont
#  reportées automatiquement dans le bordereau de visa.
# ============================================================================

import datetime as dt
import secrets
import pandas as pd
import streamlit as st
import core
import demo
from ui.style import entete, esc, sous_titre
from ui.sidebar import demander_selection

CIVILITES = ["M.", "Mme"]

# Rubriques de la page, affichées en grands boutons (une seule à la fois).
RUBRIQUES = {"projets": ":material/folder: Projets",
             "lots": ":material/view_list: Lots",
             "entreprises": ":material/apartment: Entreprises",
             "chefs": ":material/person: Chefs de projet",
             "equipes": ":material/groups: Équipes"}


def _selectionner(chef_nom=None, projet=None, lot_label=None):
    """Demande à la barre latérale de se positionner sur un chef / un projet /
    un lot au prochain affichage (un sélecteur déjà affiché ne peut pas être
    modifié dans le même tour : la demande est appliquée par ui/sidebar.py)."""
    demander_selection(projet=projet, chef_nom=chef_nom)
    demande = st.session_state.setdefault("_selection_demandee", {})
    if lot_label:
        demande["sel_lot"] = lot_label


# ---------------------------------------------------------------- projets

def _onglet_projets(ctx):
    chefs = ctx["chefs"]
    noms_chefs = [c["nom"] for c in chefs]
    code_par_nom = {c["nom"]: c["code"] for c in chefs}
    nom_par_code = {c["code"]: c["nom"] for c in chefs}

    sous_titre("Nouveau projet")
    if not chefs:
        st.info("Créez d'abord un chef de projet (onglet « Chefs de projet »).")
    else:
        with st.form("nouveau_projet", clear_on_submit=False):
            c1, c2 = st.columns([1, 2])
            code = c1.text_input("Code projet *", placeholder="ex. 2026008_GS_KUNHEIM")
            chef = c2.selectbox("Chef de projet *", noms_chefs,
                                index=noms_chefs.index(ctx["chef_nom"]) if ctx["chef_nom"] in noms_chefs else 0)
            moa = st.text_input("Maître d'ouvrage *", placeholder="ex. Eurométropole de Strasbourg")
            nom = st.text_input("Désignation du projet *", placeholder="ex. Construction d'une école de 14 classes")
            adresse = st.text_area("Adresse du maître d'ouvrage *", height=80,
                                   placeholder="1 parc de l'Étoile\n67000 Strasbourg",
                                   help="Deux lignes : rue, puis code postal et ville.")
            if st.form_submit_button("Créer le projet", type="primary", icon=":material/add:"):
                ok, message = core.add_projet(code, nom, code_par_nom[chef], moa, adresse)
                if ok:
                    _selectionner(chef, core.get_projet(code.strip()))
                    st.session_state["admin_msg"] = f"Projet {code.strip()} créé."
                    st.rerun()
                else:
                    st.error(message)

    projet = core.get_projet(ctx["proj_code"]) if ctx["proj_code"] else None
    if projet:
        sous_titre(f"Modifier le projet {projet['code']}")
        with st.form(f"modif_projet_{projet['code']}"):
            c1, c2 = st.columns([1, 2])
            c1.text_input("Code projet", value=projet["code"], disabled=True)
            chef_actuel = nom_par_code.get(projet.get("chef"), noms_chefs[0] if noms_chefs else "")
            chef = c2.selectbox("Chef de projet *", noms_chefs,
                                index=noms_chefs.index(chef_actuel) if chef_actuel in noms_chefs else 0)
            moa = st.text_input("Maître d'ouvrage *", value=projet.get("moa") or "")
            nom = st.text_input("Désignation du projet *", value=projet.get("nom") or "")
            adresse = st.text_area("Adresse du maître d'ouvrage *", value=projet.get("adresse") or "", height=80)
            if st.form_submit_button("Enregistrer les modifications", type="primary", icon=":material/save:"):
                ok, message = core.update_projet(projet["code"], nom, code_par_nom[chef], moa, adresse)
                if ok:
                    _selectionner(chef, core.get_projet(projet["code"]))
                    st.session_state["admin_msg"] = "Projet enregistré."
                    st.rerun()
                else:
                    st.error(message)

        c1, c2 = st.columns(2)
        with c1:
            sous_titre("Archiver")
            st.caption("Le projet sort des listes de sélection ; rien n'est supprimé.")
            if st.button("Archiver ce projet", icon=":material/archive:"):
                core.set_projet_archive(projet["code"], 1); st.rerun()
        with c2:
            sous_titre("Supprimer")
            with st.popover("Supprimer ce projet", icon=":material/delete_forever:"):
                st.warning(f"Suppression définitive de {projet['code']} : lots, positions, "
                           "relances et rappels.")
                if st.button("Confirmer la suppression", type="primary", key="conf_del_projet"):
                    core.delete_projet(projet["code"]); st.rerun()

    archives = core.get_projets(ctx["chef_code"], archive=1) if ctx["chef_code"] else []
    if archives:
        sous_titre("Projets archivés")
        for projet_archive in archives:
            a1, a2 = st.columns([4, 1])
            a1.write(f"{projet_archive['code']} — {projet_archive['nom']}")
            if a2.button("Restaurer", key=f"rest_{projet_archive['code']}", icon=":material/unarchive:"):
                core.set_projet_archive(projet_archive["code"], 0); st.rerun()


# ------------------------------------------------------------------- lots

def _champs_lot(prefixe, lot=None):
    """Champs communs aux formulaires de création et de modification d'un lot.
    Renvoie un dictionnaire des valeurs saisies."""
    lot = lot or {}
    c1, c2 = st.columns([1, 3])
    numero = c1.text_input("N° de lot *", value=lot.get("lot_numero") or "", placeholder="05",
                           key=f"{prefixe}_num")
    denomination = c2.text_input("Dénomination du lot *", value=lot.get("lot_denomination") or "",
                                 placeholder="CHAUFFAGE VENTILATION", key=f"{prefixe}_den")
    return {"lot_numero": numero.strip(), "lot_denomination": denomination.strip()}


def _titulaire_du_lot(LOT):
    """Entreprise titulaire et chargé d'affaires du lot, choisis dans des listes
    (fiches de l'onglet « Entreprises »). Enregistrés dès qu'on les change ; le
    chargé d'affaires peut aussi être changé par l'entreprise sur la Plateforme visa partenaires."""
    entreprises = core.get_entreprises()
    noms = {e["id"]: e["nom"] for e in entreprises}
    c1, c2 = st.columns(2)
    actuelle = LOT.get("entreprise_id")
    options = list(noms.keys())
    if actuelle not in options:
        options = [None] + options
    choix = c1.selectbox("Entreprise titulaire *", options, index=options.index(actuelle),
                         format_func=lambda i: noms.get(i, "— choisir —"), key=f"lot_ent_{LOT['id']}_{actuelle}",
                         help="Nouvelle entreprise : onglet « Entreprises ». L'adresse du bordereau vient de la fiche.")
    if choix is not None and choix != actuelle:
        core.set_lot_entreprise(LOT["id"], choix)
        st.session_state["admin_msg"] = f"Titulaire du lot : {noms[choix]} (chargé d'affaires à désigner)."
        st.rerun()
    contacts = core.get_contacts(actuelle) if actuelle else []
    libelles = {c["id"]: f"{core.nom_avec_civilite(c)} · {c['email']}" for c in contacts}
    contact = core.get_contact_lot(LOT["id"])
    contact_id = contact["id"] if contact else None
    options = [None] + list(libelles.keys())
    choix = c2.selectbox("Chargé d'affaires du lot (interlocuteur titulaire)", options,
                         index=options.index(contact_id), format_func=lambda i: libelles.get(i, "— à désigner —"),
                         key=f"lot_ca_{LOT['id']}_{contact_id}",
                         help="Interlocuteur du bordereau (E12), destinataire des relances et des mails de refus. "
                              "Nouveau contact : onglet « Entreprises ».")
    if choix != contact_id:
        core.set_lot_contact(LOT["id"], choix)
        st.session_state["admin_msg"] = "Chargé d'affaires du lot enregistré."
        st.rerun()
    adresse = (LOT.get("contact_adr") or "").replace("\n", " · ")
    st.caption(f"Adresse du bordereau : {adresse or '— à compléter dans la fiche entreprise —'}")


def _onglet_lots(ctx):
    if not ctx["proj_code"]:
        st.info("Sélectionnez ou créez d'abord un projet.")
        return
    projet = core.get_projet(ctx["proj_code"])
    sous_titre(f"Nouveau lot — {projet['code']}")
    entreprises = core.get_entreprises()
    fiches = {e["id"]: e for e in entreprises}
    with st.form("nouveau_lot"):
        valeurs = _champs_lot("nl")
        entreprise_id = st.selectbox("Entreprise titulaire *", [None] + list(fiches.keys()),
                                     format_func=lambda i: fiches[i]["nom"] if i else "— choisir —",
                                     help="Nouvelle entreprise : onglet « Entreprises ». Le chargé d'affaires "
                                          "se désigne ensuite dans la fiche du lot (ou par l'entreprise sur la Plateforme visa partenaires).")
        if st.form_submit_button("Créer le lot", type="primary", icon=":material/add:"):
            if not (valeurs["lot_numero"] and valeurs["lot_denomination"] and entreprise_id):
                st.error("Champs obligatoires : n° de lot, dénomination et entreprise.")
            else:
                fiche = fiches[entreprise_id]
                lot_id = core.add_lot(projet["code"], valeurs["lot_numero"], valeurs["lot_denomination"],
                                      fiche["nom"], fiche.get("adresse") or "", "M.", "", "", "")
                core.set_lot_entreprise(lot_id, entreprise_id)
                _selectionner(lot_label=f"{core.libelle_lot(valeurs)} — {fiche['nom']}")
                st.session_state["admin_msg"] = f"{core.libelle_lot(valeurs)} créé."
                st.rerun()

    LOT = ctx["lot"]
    if not LOT:
        return
    sous_titre(f"Modifier — {core.libelle_lot(LOT)}")
    _titulaire_du_lot(LOT)
    with st.form(f"modif_lot_{LOT['id']}"):
        valeurs = _champs_lot(f"ml{LOT['id']}", LOT)
        st.markdown("**Bordereau et relances**")
        c2, c3, c4 = st.columns(3)
        indice = c2.text_input("Indice de révision", value=LOT.get("indice") or "",
                               help="Ligne 20 du bordereau. Le n° de visa (ligne 19) est attribué "
                                    "automatiquement à chaque bordereau émis.")
        freq = c3.number_input("Relance titulaire (jours)", 3, 90, value=int(LOT.get("freq_jours") or 14),
                               help="Fréquence des relances au titulaire, en jours calendaires.")
        delai = c4.number_input("Délai de visa SB (jours ouvrés)", 1, 60,
                                value=int(LOT.get("delai_visa_sb") or core.DELAI_VISA_SB),
                                help="Échéance « à viser avant le » après chaque dépôt de document.")
        if st.form_submit_button("Enregistrer les modifications", type="primary", icon=":material/save:"):
            if not (valeurs["lot_numero"] and valeurs["lot_denomination"]):
                st.error("Champs obligatoires : n° de lot et dénomination.")
            else:
                LOT.update(valeurs)
                LOT.update({"indice": indice, "freq_jours": freq, "delai_visa_sb": delai})
                core.save_lot(LOT)
                _selectionner(lot_label=f"{core.libelle_lot(LOT)} — {LOT['entreprise'] or 'titulaire n.d.'}")
                st.session_state["admin_msg"] = "Lot enregistré."
                st.rerun()

    historique = core.get_relances(projet["code"], LOT["id"])
    if historique:
        sous_titre("Journal des relances titulaire")
        for entree in historique[:8]:
            st.caption(f"{dt.date.fromisoformat(entree['date']).strftime('%d/%m/%Y')} — "
                       f"{entree['nb_positions']} position(s)")

    sous_titre("Import marque/type (DPGF)")
    _import_dpgf(projet["code"], LOT)

    sous_titre("Supprimer le lot")
    with st.popover("Supprimer ce lot", icon=":material/delete_forever:"):
        st.warning(f"Suppression définitive de « {core.libelle_lot(LOT)} » : positions, relances et rappels.")
        if st.button("Confirmer la suppression", type="primary", key="conf_del_lot"):
            core.delete_lot(LOT["id"]); st.rerun()


def _import_dpgf(proj_code, LOT):
    """Import de l'annexe marque/type : aperçu des lignes retenues, à cocher /
    décocher avant de créer les positions."""
    up = st.file_uploader("Annexe CCTP (.xlsx)", type=["xlsx"], key="dpgf")
    if not up:
        st.caption("Seules les lignes dont la référence est une numérotation (3, 3.1, 3.1.2, 3.1.2a…) "
                   "sont proposées : les titres comme « CHAPITRE 3 » sont écartés.")
        return
    only_leaf = st.checkbox("Ignorer les sous-titres de section (« 3 » quand « 3.1 » existe)", value=True)
    retenues, ecartees = core.lire_dpgf(up.getvalue())
    if only_leaf:
        retenues = core.feuilles_dpgf(retenues)
    refs_existantes = {p["pos"] for p in core.get_positions(proj_code, LOT["id"])}
    tableau = pd.DataFrame([
        {"Importer": ref not in refs_existantes, "Réf.": ref, "Désignation": des,
         "Marque MOE": marque, "Type MOE": typ,
         "État": "déjà dans le lot" if ref in refs_existantes else "nouvelle"}
        for ref, des, marque, typ in retenues])
    if tableau.empty:
        st.warning("Aucune ligne de position reconnue dans ce fichier.")
        return
    st.caption(f"{len(retenues)} ligne(s) reconnue(s) · {len(ecartees)} titre(s) ou ligne(s) écartée(s). "
               "Décochez les lignes à ne pas importer.")
    edite = st.data_editor(tableau, hide_index=True, width="stretch", key=f"apercu_dpgf_{up.name}",
                           disabled=["Réf.", "Désignation", "Marque MOE", "Type MOE", "État"],
                           column_config={"Importer": st.column_config.CheckboxColumn(width="small")})
    if ecartees:
        with st.expander(f"Lignes écartées ({len(ecartees)})"):
            st.dataframe(pd.DataFrame(ecartees, columns=["Référence", "Désignation"]), hide_index=True,
                         width="stretch")
    a_importer = edite[edite["Importer"]]
    if st.button(f"Importer {len(a_importer)} position(s)", type="primary", icon=":material/download_done:",
                 disabled=a_importer.empty):
        nb_ajoutees = 0
        for indice, ligne in enumerate(a_importer.itertuples(index=False)):
            ref, des, marque_moe, type_moe = ligne[1], ligne[2], ligne[3], ligne[4]
            if ref in refs_existantes:
                continue
            # Si le MOE a proposé une marque ou un type, on verrouille
            # ces champs (affichés grisés, modifiables via un interrupteur).
            verrou_moe = 1 if (marque_moe or type_moe) else 0
            core.add_position(proj_code, LOT["id"], ref, des, core.TRAME_A_DEFINIR,
                              ordre=1000 + indice, default_export=0, marque_moe=marque_moe,
                              type_moe=type_moe, mt_verrou=verrou_moe)
            nb_ajoutees = nb_ajoutees + 1
        st.session_state["admin_msg"] = (f"{nb_ajoutees} position(s) importée(s) — trame à définir, "
                                         "propositions MOE grisées lorsque le fichier les fournit.")
        st.rerun()


# ------------------------------------------------------------------ chefs

AUCUNE_EQUIPE = "— aucune équipe —"


def _choix_equipe(cle, team_id=None):
    """Liste déroulante des équipes. Renvoie l'id choisi (ou None)."""
    equipes = core.get_teams()
    noms = [AUCUNE_EQUIPE] + [e["nom"] for e in equipes]
    actuel = next((e["nom"] for e in equipes if e["id"] == team_id), AUCUNE_EQUIPE)
    choix = st.selectbox("Équipe", noms, index=noms.index(actuel), key=cle)
    return next((e["id"] for e in equipes if e["nom"] == choix), None)


def _onglet_chefs(ctx):
    sous_titre("Nouveau chef de projet")
    st.caption("En attendant la connexion Microsoft 365, le chef de projet du projet est "
               "l'interlocuteur écrit dans le bordereau et le signataire des relances.")
    with st.form("nouveau_chef", clear_on_submit=True):
        c1, c2, c3, c4 = st.columns([2, 2, 1.4, 1.6])
        nom = c1.text_input("Prénom NOM *")
        email = c2.text_input("Email", placeholder="prenom.nom@solares-bauen.fr")
        tel = c3.text_input("Téléphone")
        with c4:
            team_id = _choix_equipe("nc_team")
        if st.form_submit_button("Ajouter", type="primary", icon=":material/person_add:"):
            if nom.strip():
                code = core.add_chef(nom.strip(), email.strip(), tel.strip())
                core.set_chef_team(code, team_id)
                _selectionner(nom.strip())
                st.rerun()
            else:
                st.error("Le nom est obligatoire.")

    equipe_par_id = {e["id"]: e["nom"] for e in core.get_teams()}
    for chef in ctx["chefs"]:
        equipe = equipe_par_id.get(chef.get("team_id"), "sans équipe")
        with st.expander(f"{chef['nom']}  ·  {chef['code']}  ·  {equipe}", icon=":material/person:"):
            with st.form(f"chef_{chef['code']}"):
                c1, c2, c3, c4 = st.columns([2, 2, 1.4, 1.6])
                nom = c1.text_input("Prénom NOM", value=chef["nom"] or "")
                email = c2.text_input("Email", value=chef.get("email") or "")
                tel = c3.text_input("Téléphone", value=chef.get("telephone") or "")
                with c4:
                    team_id = _choix_equipe(f"team_{chef['code']}", chef.get("team_id"))
                if st.form_submit_button("Enregistrer", icon=":material/save:"):
                    core.update_chef(chef["code"], nom.strip() or chef["nom"], email.strip(), tel.strip())
                    core.set_chef_team(chef["code"], team_id)
                    if st.session_state.get("sel_chef") == chef["nom"]:
                        _selectionner(nom.strip() or chef["nom"])
                    st.rerun()
            with st.popover("Supprimer", icon=":material/delete:"):
                if st.button("Confirmer la suppression", key=f"del_chef_{chef['code']}", type="primary"):
                    ok, message = core.delete_chef(chef["code"])
                    if ok:
                        st.rerun()
                    else:
                        st.error(message)


# ---------------------------------------------------------------- équipes

def _onglet_equipes(ctx):
    sous_titre("Nouvelle équipe")
    st.caption("Chaque chef de projet est rattaché à une équipe (onglet « Chefs de projet ») ; "
               "le Dashboard peut alors afficher le portefeuille de toute l'équipe.")
    with st.form("nouvelle_equipe", clear_on_submit=True):
        c1, c2 = st.columns([3, 1], vertical_alignment="bottom")
        nom = c1.text_input("Nom de l'équipe *", placeholder="ex. Team Fluides Est")
        if c2.form_submit_button("Créer l'équipe", type="primary", icon=":material/group_add:",
                                 width="stretch"):
            ok, resultat = core.add_team(nom)
            if ok:
                st.session_state["admin_msg"] = f"Équipe « {nom.strip()} » créée."
                st.rerun()
            else:
                st.error(resultat)

    chefs_par_code = {c["code"]: c["nom"] for c in ctx["chefs"]}
    sans_equipe = [c["nom"] for c in ctx["chefs"] if not c.get("team_id")]
    for equipe in core.get_teams():
        membres = ", ".join(chefs_par_code.get(c, c) for c in equipe["chefs"]) or "aucun chef de projet"
        with st.expander(f"{equipe['nom']}  ·  {len(equipe['chefs'])} chef(s) de projet", icon=":material/groups:"):
            st.caption(f"Membres : {membres}")
            with st.form(f"equipe_{equipe['id']}"):
                c1, c2 = st.columns([3, 1], vertical_alignment="bottom")
                nouveau = c1.text_input("Nom de l'équipe", value=equipe["nom"])
                if c2.form_submit_button("Renommer", icon=":material/edit:", width="stretch"):
                    ok, message = core.rename_team(equipe["id"], nouveau)
                    if ok:
                        st.rerun()
                    else:
                        st.error(message)
            with st.popover("Supprimer l'équipe", icon=":material/delete:"):
                st.caption("Les chefs de projet de l'équipe restent, sans équipe.")
                if st.button("Confirmer la suppression", key=f"del_team_{equipe['id']}", type="primary"):
                    core.delete_team(equipe["id"]); st.rerun()
    if sans_equipe:
        st.caption("Sans équipe : " + ", ".join(sans_equipe))


# ------------------------------------------------------------------- page

# -------------------------------------------------------------- entreprises

def _champs_contact(prefixe, contact=None):
    """Civilité, nom, mail, téléphone d'un chargé d'affaires (dans un formulaire)."""
    contact = contact or {}
    c1, c2 = st.columns([1, 3])
    civ = c1.selectbox("Civilité", CIVILITES, key=f"{prefixe}_civ",
                       index=CIVILITES.index(contact.get("civ")) if contact.get("civ") in CIVILITES else 0)
    nom = c2.text_input("Prénom NOM *", value=contact.get("nom") or "", key=f"{prefixe}_nom")
    c1, c2 = st.columns(2)
    email = c1.text_input("Mail *", value=contact.get("email") or "", key=f"{prefixe}_mail")
    telephone = c2.text_input("Téléphone", value=contact.get("telephone") or "", key=f"{prefixe}_tel")
    return civ, nom, email, telephone


def _compte_portail(fiche):
    """Compte Plateforme visa partenaires de l'entreprise : identifiant, mot de passe (créer /
    réinitialiser), activer / désactiver."""
    eid = fiche["id"]
    if fiche["a_compte"]:
        etat = "actif" if fiche["actif"] else "désactivé"
        st.markdown(f"Compte **{esc(fiche['identifiant'])}** — {etat}")
    else:
        st.caption(f"Pas encore de compte {core.NOM_PORTAIL}.")
    nouveau = st.session_state.pop(f"mdp_affiche_{eid}", None)
    if nouveau:
        st.info(f"Mot de passe à transmettre à l'entreprise (il ne sera plus affiché) : **{nouveau}**",
                icon=":material/key:")
    with st.form(f"compte_{eid}"):
        c1, c2 = st.columns(2)
        identifiant = c1.text_input("Identifiant *", value=fiche.get("identifiant") or "",
                                    placeholder="ex. thermi-est")
        mot_de_passe = c2.text_input("Nouveau mot de passe (vide = généré)", type="password",
                                     help=f"{core.LONGUEUR_MINI_MOT_DE_PASSE} caractères minimum.")
        libelle = "Réinitialiser le mot de passe" if fiche["a_compte"] else "Créer le compte"
        if st.form_submit_button(libelle, type="primary", icon=":material/key:"):
            mot_de_passe = mot_de_passe or secrets.token_urlsafe(9)
            ok, message = core.set_compte_entreprise(eid, identifiant, mot_de_passe)
            if ok:
                st.session_state[f"mdp_affiche_{eid}"] = mot_de_passe
                st.rerun()
            st.error(message)
    if fiche["a_compte"]:
        if fiche["actif"]:
            if st.button("Désactiver le compte", key=f"desact_{eid}", icon=":material/block:"):
                core.set_compte_actif(eid, 0); st.rerun()
        elif st.button("Réactiver le compte", key=f"react_{eid}", icon=":material/check_circle:"):
            core.set_compte_actif(eid, 1); st.rerun()


def _contacts_entreprise(fiche):
    """Chargés d'affaires : même liste que la page « Chargés d'affaires » de la Plateforme visa partenaires."""
    eid = fiche["id"]
    lots = core.get_lots_entreprise(eid)
    for contact in core.get_contacts(eid):
        ses_lots = [f"{l['projet']} · {core.libelle_lot(l)}" for l in lots if l.get("contact_id") == contact["id"]]
        a, b = st.columns([4, 1.3], vertical_alignment="center")
        nom = core.nom_avec_civilite(contact)
        telephone = f" · {esc(contact['telephone'])}" if contact.get("telephone") else ""
        a.markdown(f"<div style='font-size:14px'><b>{esc(nom)}</b> · {esc(contact['email'])}{telephone}</div>"
                   f"<div style='font-size:12px;color:#5C5C5C'>Chargé d'affaires de : "
                   f"{esc(', '.join(ses_lots)) or 'aucun lot'}</div>", unsafe_allow_html=True)
        with b.popover("Modifier", icon=":material/edit:", width="stretch"):
            with st.form(f"modif_contact_{contact['id']}"):
                civ, nom, email, telephone = _champs_contact(f"mc{contact['id']}", contact)
                if st.form_submit_button("Enregistrer", type="primary"):
                    ok, message = core.update_contact(contact["id"], nom, email, civ, telephone)
                    if ok:
                        st.rerun()
                    st.error(message)
            if st.button("Retirer de la liste", key=f"ret_contact_{contact['id']}", icon=":material/delete:"):
                core.retirer_contact(contact["id"]); st.rerun()
    with st.form(f"nouveau_contact_{eid}", clear_on_submit=True):
        st.markdown("**Ajouter un chargé d'affaires**")
        civ, nom, email, telephone = _champs_contact(f"nc{eid}")
        if st.form_submit_button("Ajouter", icon=":material/person_add:"):
            ok, resultat = core.add_contact(eid, nom, email, civ, telephone)
            if ok:
                st.rerun()
            st.error(resultat)


def _onglet_entreprises():
    sous_titre("Nouvelle entreprise")
    with st.form("nouvelle_entreprise", clear_on_submit=True):
        c1, c2 = st.columns([1, 1.4])
        nom = c1.text_input("Nom de l'entreprise *", placeholder="ex. Thermi Est")
        adresse = c2.text_area("Adresse", height=80, placeholder="12 rue du Rhin\n67000 Strasbourg",
                               help="Deux lignes : rue, puis code postal et ville (destinataire du bordereau).")
        if st.form_submit_button("Créer l'entreprise", type="primary", icon=":material/add:"):
            ok, resultat = core.add_entreprise(nom, adresse)
            if ok:
                st.session_state["admin_msg"] = f"Entreprise {nom.strip()} créée."
                st.rerun()
            st.error(resultat)

    sous_titre("Entreprises titulaires")
    st.caption(f"Fiche, compte {core.NOM_PORTAIL} et chargés d'affaires. L'entreprise gère aussi ses "
               f"chargés d'affaires sur la {core.NOM_PORTAIL} : c'est la même liste.")
    for fiche in core.get_entreprises():
        if fiche["a_compte"]:
            compte = "compte actif" if fiche["actif"] else "compte désactivé"
        else:
            compte = "sans compte"
        with st.expander(f"{fiche['nom']} — {fiche['nb_lots']} lot(s) · {compte}"):
            g, d = st.columns(2, gap="large")
            with g:
                st.markdown("<div class='t-lab'>Fiche</div>", unsafe_allow_html=True)
                with st.form(f"fiche_{fiche['id']}"):
                    nom = st.text_input("Nom *", value=fiche["nom"])
                    adresse = st.text_area("Adresse", value=fiche.get("adresse") or "", height=80)
                    if st.form_submit_button("Enregistrer la fiche", icon=":material/save:"):
                        ok, message = core.update_entreprise(fiche["id"], nom, adresse)
                        if ok:
                            st.session_state["admin_msg"] = "Fiche enregistrée (recopiée dans ses lots)."
                            st.rerun()
                        st.error(message)
                st.markdown(f"<div class='t-lab'>Compte {core.NOM_PORTAIL}</div>", unsafe_allow_html=True)
                _compte_portail(fiche)
            with d:
                st.markdown("<div class='t-lab'>Chargés d'affaires / contacts</div>", unsafe_allow_html=True)
                _contacts_entreprise(fiche)
                lots = core.get_lots_entreprise(fiche["id"])
                if lots:
                    st.markdown("<div class='t-lab' style='margin-top:10px'>Lots</div>", unsafe_allow_html=True)
                    for lot in lots:
                        st.caption(f"{lot['projet']} · {core.libelle_lot(lot)} — chargé d'affaires : "
                                   f"{core.nom_avec_civilite({'civ': lot.get('contact_civ'), 'nom': lot.get('contact_nom')}) if lot.get('contact_nom') else 'à désigner'}")
            if fiche["nb_lots"] == 0:
                with st.popover("Supprimer cette entreprise", icon=":material/delete_forever:"):
                    if st.button("Confirmer la suppression", type="primary", key=f"del_ent_{fiche['id']}"):
                        ok, message = core.delete_entreprise(fiche["id"])
                        if ok:
                            st.rerun()
                        st.error(message)


def afficher(ctx):
    entete("", "Projets, lots, entreprises et équipes",
           [("", "Les informations saisies ici alimentent automatiquement le bordereau de visa.")])
    message = st.session_state.pop("admin_msg", None)
    if message:
        st.success(message, icon=":material/check_circle:")

    if not ctx["chefs"]:
        with st.container(border=True):
            st.markdown("**Base vide.** Chargez un jeu d'exemple pour découvrir l'outil "
                        "(2 chefs de projet, 4 projets, documents déposés, relances et rappels).")
            if st.button("Charger la démo", type="primary", icon=":material/play_arrow:"):
                demo.charger_demo(); st.rerun()

    # Rubriques en grands boutons (même style que les volets du Suivi) ; seule la
    # rubrique choisie est affichée.
    with st.container(key="volets_admin"):
        rubrique = st.segmented_control("Rubrique", list(RUBRIQUES), format_func=lambda cle: RUBRIQUES[cle],
                                        default="projets", required=True, key="rubrique_admin",
                                        label_visibility="collapsed", width="stretch") or "projets"
    if rubrique == "projets":
        _onglet_projets(ctx)
    elif rubrique == "lots":
        _onglet_lots(ctx)
    elif rubrique == "entreprises":
        _onglet_entreprises()
    elif rubrique == "chefs":
        _onglet_chefs(ctx)
    else:
        _onglet_equipes(ctx)
