# ============================================================================
#  ui/page_suivi.py — page « Suivi des visas » : travail sur les positions
#  d'un lot (avis, documents, attendus, contrôles, relances, bordereau).
#
#  Organisation :
#    - callbacks d'enregistrement (appelés par Streamlit à chaque saisie)
#    - en-tête, bandeaux d'alerte, cartes d'indicateurs et actions
#    - sélection rapide (relance / export) et ajout de position
#    - rendu d'une position : ligne repliée + détail (aussi pour les anciens
#      visas, en lecture seule)
#    - afficher() : assemble la page
# ============================================================================

import datetime as dt
import json
import mimetypes
import os
from collections import Counter
import streamlit as st
import core
from ui import bordereaux as bordereaux_ui
from ui.style import (STAT_KIND, STAT_COLOR, badge, badge_avis, barre_actions, entete, esc,
                      sous_titre, tuiles)

TEMPLATE_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), "templates", "VISA_TYPE.xlsx")
MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _date_fr(texte_iso):
    return dt.date.fromisoformat(texte_iso).strftime("%d/%m/%Y") if texte_iso else ""


# ---------- callbacks d'enregistrement ----------
# POS (liste des positions du lot) est passé en argument : les callbacks
# s'exécutent au début de l'affichage suivant, avant que la page ne relise la base.

def save_field(POS, idx, field, key):
    val = st.session_state[key]
    if isinstance(val, dt.date): val = val.isoformat()
    if isinstance(val, bool): val = int(val)
    POS[idx][field] = val
    core.save_position(POS[idx])


def save_ctrl(POS, idx, ci, field, key):
    val = st.session_state[key]
    if isinstance(val, bool): val = int(val)
    POS[idx]["ctrl"][ci][field] = val
    if field == "so" and val:
        POS[idx]["ctrl"][ci]["v"] = 0
    core.save_position(POS[idx])


def save_att(POS, idx, ai, field, key):
    val = st.session_state[key]
    if isinstance(val, dt.date): val = val.isoformat()
    POS[idx]["att"][ai][field] = val
    if field == "s" and val not in core.STATUTS_FOURNIS: POS[idx]["att"][ai]["d"] = ""
    core.save_position(POS[idx])


def apply_relance_pills(POS, proj_code, lot_id):
    sel = list(st.session_state.get("qsel") or [])
    core.set_relance_bulk(proj_code, lot_id, None, 0)
    if sel: core.set_relance_bulk(proj_code, lot_id, sel, 1)
    for p in POS:
        p["relance"] = 1 if p["avis"] in sel else 0
        st.session_state[f"rel_{p['id']}"] = bool(p["relance"])


# ---------- en-tête, bandeaux, indicateurs ----------

def _bandeaux(proj_code, LOT, POS, rappels):
    due, last, nxt, ncand = core.relance_status(proj_code, LOT, POS)
    if ncand and due:
        lasttxt = f"dernière relance le {_date_fr(last)}" if last else "aucune relance envoyée"
        st.warning(f"**Relance titulaire à faire** — {ncand} position(s) en MQ/VAOB/REF, {lasttxt} "
                   f"(fréquence : {LOT.get('freq_jours', 14)} j).", icon=":material/outgoing_mail:")
    elif ncand:
        st.info(f"Prochaine relance entreprise le {_date_fr(nxt)} — {ncand} position(s) en MQ/VAOB/REF.",
                icon=":material/event_available:")
    en_retard = [r for r in rappels if r["niveau"] == "retard"]
    if en_retard:
        st.error(f"**{len(en_retard)} document(s) à viser par Solares Bauen en retard** — échéance "
                 f"dépassée depuis le {_date_fr(en_retard[0]['echeance'])}.", icon=":material/alarm:")


def _indicateurs(LOT, POS, rappels):
    """Rangée de tuiles : visas validés, relance titulaire, à viser par SB,
    bordereaux émis."""
    hors_so = [p for p in POS if p["avis"] != "SO"]
    valides = [p for p in POS if p["avis"] in ("VSO", "VAO")]
    taux = round(len(valides) / len(hors_so) * 100) if hors_so else 0
    a_relancer = core.positions_a_relancer(POS)
    n_rel = sum(1 for p in POS if p["relance"])
    retard = sum(1 for r in rappels if r["niveau"] == "retard")
    if rappels:
        sous_sb = f"prochaine échéance le {_date_fr(rappels[0]['echeance'])}"
        if retard:
            sous_sb = f"{retard} en retard · " + sous_sb
    else:
        sous_sb = "aucun document en attente"
    bordereaux = core.get_bordereaux(LOT["id"])
    if bordereaux:
        dernier = bordereaux[0]
        sous_bx = f"dernier : {dernier['libelle']} du {dt.datetime.fromisoformat(dernier['date']).strftime('%d/%m/%Y')}"
    else:
        sous_bx = f"prochain : {core.format_numero(core.prochain_numero(LOT['id']))}"
    tuiles([
        {"lab": "Visas validés", "val": f"{len(valides)} / {len(hors_so)}", "unite": "pos.", "jauge": taux,
         "sub": "VSO + VAO, hors SO"},
        {"lab": "Relance titulaire", "val": len(a_relancer), "unite": "pos.",
         "sub": f"MQ / VAOB / REF · {n_rel} cochée(s) pour la relance"},
        {"lab": "À viser par SB", "val": len(rappels), "unite": "doc.", "alerte": retard > 0, "sub": sous_sb},
        {"lab": "Bordereaux émis", "val": len(bordereaux), "sub": sous_bx},
    ])


def _barre_actions(projet, proj_code, LOT, POS, signataire):
    """Actions du lot : brouillon de relance, émission du bordereau numéroté,
    export du statut des visas, registre des bordereaux émis."""
    barre_actions("Actions du lot")
    n_rel = sum(1 for p in POS if p["relance"])
    # Conteneur nommé : style.py y agrandit les boutons (texte sur 2 lignes, sans « … »).
    a1, a2, a3, a4 = st.container(key="actions_lot").columns(4)
    with a1:
        if n_rel:
            clicked = st.download_button("Générer le brouillon de relance",
                core.build_eml(projet, LOT, POS),
                file_name=f"relance_{proj_code}_Lot{LOT.get('lot_numero') or LOT['id']}.eml",
                mime="message/rfc822", width="stretch", icon=":material/mail:")
            if clicked:
                core.log_relance(proj_code, LOT["id"], 1, [(p["pos"] or p["des"]) for p in POS if p["relance"]])
                st.toast("Relance enregistrée au journal"); st.rerun()
        else:
            st.button("Générer le brouillon de relance", disabled=True, width="stretch",
                      icon=":material/mail:", help="Cochez « Relance » sur au moins une position.")
    with a2:
        n_export = sum(1 for p in POS if p.get("export", 1))
        prochain = core.format_numero(core.prochain_numero(LOT["id"]))
        if st.button("Générer le bordereau de visa", type="primary", width="stretch",
                     icon=":material/description:", disabled=not (os.path.exists(TEMPLATE_PATH) and n_export),
                     help=f"Émet le bordereau {prochain} avec les {n_export} position(s) cochées « Au bordereau »."):
            bordereau = core.emettre_bordereau(projet, LOT, POS, signataire)
            with st.spinner("Génération du bordereau…"):
                bordereaux_ui.preparer_exports(bordereau)
            st.session_state[f"dernier_bx_{LOT['id']}"] = bordereau["id"]
            st.rerun()   # tuile « Bordereaux émis » et numéro suivant à jour
    with a3:
        st.download_button("Exporter statut Visas", core.export_statut_visas(projet, LOT, POS),
                           file_name=f"statut_visas_{proj_code}_Lot{LOT.get('lot_numero') or ''}.xlsx",
                           mime=MIME_XLSX, width="stretch", icon=":material/download:",
                           help="Tableau Excel du lot : avis, pièces, dates, remarques, échéances de visa SB.")
    with a4:
        bordereaux_ui.bouton_registre(f"Bordereaux émis ({len(core.get_bordereaux(LOT['id']))})",
                                      LOT["id"], key="bx_suivi")

    dernier = core.get_bordereau(st.session_state.get(f"dernier_bx_{LOT['id']}") or 0)
    if dernier:
        with st.container(border=True):
            c1, c2 = st.columns([2, 1.4], vertical_alignment="center")
            nb = len(dernier["contenu"].get("positions", []))
            c1.markdown(f"**{esc(dernier['libelle'])} émis** — {nb} position(s). "
                        "Retrouvez-le à tout moment dans « Bordereaux émis ».")
            with c2:
                if not bordereaux_ui.boutons_telechargement(dernier, prefixe="dernier"):
                    if st.button("Préparer les fichiers", key="prep_dernier", width="stretch"):
                        bordereaux_ui.preparer_exports(dernier); st.rerun()
            pret = st.session_state.get(f"bx_{dernier['id']}")
            if pret and not pret["pdf"]:
                st.caption("PDF indisponible : " + pret["message"])


# ---------- sélection rapide et ajout ----------

def _selection_rapide(proj_code, LOT, POS):
    q0, q1, q2 = st.columns([0.8, 1.7, 1.7], vertical_alignment="bottom")
    with q0:
        st.caption("Relance titulaire")
        if st.button("Tous", width="stretch", help="Coche MQ + VAOB + REF d'un coup"):
            st.session_state["qsel"] = ["MQ","VAOB","REF"]
            apply_relance_pills(POS, proj_code, LOT["id"]); st.rerun()
    with q1:
        st.pills("Par avis (combinables)", ["MQ","VAOB","REF"], selection_mode="multi",
                 key="qsel", on_change=apply_relance_pills, args=(POS, proj_code, LOT["id"]))
    with q2:
        st.caption("Bordereau")
        e = st.columns(2)
        if e[0].button("Tout cocher", width="stretch"):
            core.set_export_bulk(proj_code, LOT["id"], 1)
            for p in POS: st.session_state[f"exp_{p['id']}"] = True
            st.rerun()
        if e[1].button("Tout décocher", width="stretch"):
            core.set_export_bulk(proj_code, LOT["id"], 0)
            for p in POS: st.session_state[f"exp_{p['id']}"] = False
            st.rerun()
    n_rel = sum(1 for p in POS if p["relance"])
    if n_rel:
        cnt = Counter(p["avis"] for p in POS if p["relance"])
        st.caption("Sélection relance : " + f"{n_rel} position(s) — " + " · ".join(f"{v} {k}" for k,v in cnt.items()))


def _ajout_position(proj_code, LOT):
    with st.expander("Ajouter une position / un document", icon=":material/add_circle:"):
        with st.form("addpos", clear_on_submit=True):
            a1,a2 = st.columns([1, 3])
            ref = a1.text_input("Référence CCTP (facultatif)", help="Sans référence, la position ira dans l'onglet Plans/NdC du bordereau.")
            des = a2.text_input("Désignation")
            st.caption("La trame d'attendus sera proposée automatiquement sur la position (à valider).")
            if st.form_submit_button("Ajouter au bordereau", type="primary") and des:
                core.add_position(proj_code, LOT["id"], (ref or "").strip(), des, core.TRAME_A_DEFINIR); st.rerun()


# Les trois volets du lot, affichés en grands boutons (un seul volet à la fois).
VOLETS = {"cctp": ":material/list_alt: Positions CCTP/Matériel",
          "plans": ":material/architecture: Plans / NdC",
          "anciens": ":material/history: Anciens visas"}


def _choix_volet(vus, tous, filtre_actif):
    """Grand sélecteur natif (st.segmented_control) des volets Positions CCTP/Matériel /
    Plans-NdC / Anciens visas ; un volet est toujours choisi. Le nombre
    d'éléments de chaque volet est ajouté en CSS (::after) : un libellé qui
    change recréerait le sélecteur et ramènerait au premier volet.
    Renvoie la clé du volet choisi."""
    regles = []
    for rang, cle in enumerate(VOLETS, start=1):
        nombre = f"{len(vus[cle])} / {len(tous[cle])}" if filtre_actif else str(len(tous[cle]))
        regles.append(f".st-key-volets_suivi [data-testid='stButtonGroup'] button:nth-of-type({rang}) p::after"
                      f"{{content:'{nombre}';}}")
    st.markdown("<style>" + "".join(regles) + "</style>", unsafe_allow_html=True)
    return st.segmented_control("Volet", list(VOLETS), format_func=lambda cle: VOLETS[cle], default="cctp",
                                required=True, key="volet_suivi", label_visibility="collapsed",
                                width="stretch") or "cctp"


def _filtres(POS, rappels_par_pos):
    """Barre de filtres de la liste (composants natifs) : recherche, avis, et
    « à viser par SB » (rappel ouvert). Renvoie la fonction qui dit si une
    position est affichée, et un texte qui résume le filtre (vide = aucun)."""
    # Recherche + interrupteur sur une ligne, avis en dessous : rien ne se chevauche
    # même quand la page est étroite.
    f1, f2 = st.columns([3, 1.3], vertical_alignment="center")
    texte = f1.text_input("Rechercher une position", key="filtre_txt",
                          placeholder="Rechercher : référence ou désignation", label_visibility="collapsed")
    a_viser = f2.toggle("À viser par SB", key="filtre_a_viser",
                        help="Seulement les positions qui ont des documents en attente de visa.")
    avis = st.pills("Filtrer par avis", core.AVIS, selection_mode="multi", key="filtre_avis")
    recherche = core._norm((texte or "").strip())   # minuscules, sans accents

    def visible(p):
        if recherche and recherche not in core._norm(f"{p.get('pos') or ''} {p.get('des') or ''}"):
            return False
        if avis and p["avis"] not in avis:
            return False
        if a_viser and p["id"] not in rappels_par_pos:
            return False
        return True

    actifs = []
    if recherche:
        actifs.append(f"« {texte.strip()} »")
    if avis:
        actifs.append(" / ".join(avis))
    if a_viser:
        actifs.append("à viser par SB")
    return visible, " · ".join(actifs)


# ---------- rendu d'une position ----------

def _texte_position(p, rappel):
    """Ligne de titre : référence, désignation et proposition entreprise dans
    la même police, puis petites pastilles d'information."""
    ref = (p.get("pos") or "").strip()
    is_plan = not ref
    prop = "" if is_plan else " · ".join(x for x in [p.get("marque"), p.get("typeprop")] if x)
    if is_plan:
        titre = f"<b>{esc(p['des'])}</b>"
    else:
        titre = f"<span class='pos-ref'>{esc(ref)}</span>{esc(p['des'])}"
    if prop:
        titre += f"<span class='pos-prop'> — {esc(prop)}</span>"
    pastilles = []
    if p.get("famille") == core.TRAME_A_DEFINIR:
        pastilles.append("<span class='pastille'>trame à définir</span>")
    if rappel:
        classe = "pastille rouge" if rappel["niveau"] in ("retard", "proche") else "pastille"
        pastilles.append(f"<span class='{classe}'>à viser avant le {_date_fr(rappel['echeance'])}</span>")
    if (p.get("rem") or "").strip():
        pastilles.append("<span class='pastille'>remarque</span>")
    indices = f"<div class='pos-indices'>{''.join(pastilles)}</div>" if pastilles else ""
    return f"<div class='pos-texte'>{titre}{indices}</div>"


def _bordereau_du_visa(p, LOT, archive):
    """Bordereau émis du visa affiché (visa actuel ou ancien visa), ou None :
    un visa = un bordereau du lot ; le n° n'existe qu'une fois le bordereau émis."""
    if archive:
        bordereau_id = p.get("visa_bordereau_id")
    else:
        emis = core.visa_actuel(p["id"], LOT["id"])
        bordereau_id = emis["id"] if emis else None
    return core.get_bordereau(bordereau_id) if bordereau_id else None


def _numero_visa_html(bordereau):
    """Petit libellé sous l'avis : « VISA N°03 » ou « en cours »."""
    if bordereau:
        return f"<div class='num-visa'>{esc(bordereau['libelle'])}</div>"
    return "<div class='num-visa gris'>visa en cours</div>"


def _bloc_visa(bordereau, pid, archive):
    """En tête du détail : n° et date du visa, bouton « Voir bordereau » (grisé
    tant que le visa n'est pas émis)."""
    a, b = st.columns([2.4, 1.6], vertical_alignment="center")
    if bordereau:
        date_emission = _date_fr(bordereau["date"][:10])
        quel = "Ancien visa" if archive else "Visa actuel"
        a.markdown(f"<div style='font-size:14px'>{quel} : <b>{esc(bordereau['libelle'])}</b> "
                   f"<span style='color:#5C5C5C'>— bordereau émis le {date_emission}</span></div>",
                   unsafe_allow_html=True)
    else:
        texte = ("Ancien visa — aucun bordereau émis enregistré pour ce visa" if archive
                 else "Visa en cours — n° attribué à l'émission du bordereau")
        a.markdown(f"<div style='font-size:14px;color:#6B6B6B'>{texte}</div>", unsafe_allow_html=True)
    with b:
        if bordereau is None:
            st.button("Voir bordereau", key=f"vb_{pid}", disabled=True, icon=":material/description:",
                      width="stretch", help="Aucun bordereau émis pour ce visa.")
        elif not bordereaux_ui.boutons_telechargement(bordereau, prefixe=f"pos{pid}"):
            if st.button(f"Voir bordereau {bordereau['libelle']}", key=f"vb_{pid}", icon=":material/description:",
                         width="stretch"):
                with st.spinner("Préparation du bordereau…"):
                    bordereaux_ui.preparer_exports(bordereau)
                st.rerun()


def render_row(i, p, POS, projet, LOT, rappels_par_pos, archive=False):
    pid = p["id"]
    rappel = None if archive else rappels_par_pos.get(pid)
    bordereau = _bordereau_du_visa(p, LOT, archive)
    if archive:
        h = st.columns([4.0, 0.6, 1.2, 1.1], vertical_alignment="center")
        h[0].markdown(_texte_position(p, None), unsafe_allow_html=True)
        h[1].markdown(badge_avis(p["avis"]), unsafe_allow_html=True)
        numero = f"<b>{esc(bordereau['libelle'])}</b> · " if bordereau else ""
        # Sans date de visa saisie, on prend la date d'émission du bordereau.
        date_visa = p.get("datevisa") or (bordereau["date"][:10] if bordereau else "")
        h[2].markdown(f"<span style='color:#5C5C5C;font-size:13px'>{numero}Visa du "
                      f"{_date_fr(date_visa) or '—'}</span>", unsafe_allow_html=True)
        bouton = h[3]
    else:
        h = st.columns([3.0, 0.85, 1.05, 1.4, 1.2], vertical_alignment="center")
        h[0].markdown(_texte_position(p, rappel), unsafe_allow_html=True)
        h[1].markdown(badge_avis(p["avis"]) + _numero_visa_html(bordereau), unsafe_allow_html=True)
        h[2].checkbox("Relance", value=bool(p["relance"]), key=f"rel_{pid}",
                      on_change=save_field, args=(POS,i,"relance",f"rel_{pid}"))
        h[3].checkbox("Au bordereau", value=bool(p.get("export",1)), key=f"exp_{pid}",
                      help="Cochée : la position est intégrée au prochain bordereau de visa.",
                      on_change=save_field, args=(POS,i,"export",f"exp_{pid}"))
        bouton = h[4]
    if "open" not in st.session_state: st.session_state.open = set()
    if bouton.button("Détail", key=f"tg_{pid}", width="stretch",
                     icon=":material/expand_less:" if pid in st.session_state.open else ":material/expand_more:",
                     type="secondary"):
        st.session_state.open.symmetric_difference_update({pid}); st.rerun()

    if pid in st.session_state.open:
        with st.container(border=True):
            _bloc_visa(bordereau, pid, archive)
            _detail_position(i, p, POS, projet, LOT, rappel, archive)
    st.markdown("<hr class='rowsep'>", unsafe_allow_html=True)


def _detail_position(i, p, POS, projet, LOT, rappel, ro):
    """Détail d'une position. ro=True : ancien visa, tout est en lecture seule
    (seules la date de l'ancien visa et sa suppression restent possibles)."""
    pid = p["id"]
    is_plan = not (p.get("pos") or "").strip()
    ep = st.session_state.get(f"ep_{pid}", 0)

    # --- actions en tête ---
    if ro:
        st.markdown("<div class='note-ro'>Ancien visa — copie figée, en lecture seule.</div>", unsafe_allow_html=True)
        t0 = st.columns([1.3, 3, 1.3])
        dv = dt.date.fromisoformat(p["datevisa"]) if p.get("datevisa") else None
        t0[0].date_input("Date de l'ancien visa", value=dv, key=f"dv_{pid}", format="DD/MM/YYYY",
                         on_change=save_field, args=(POS, i, "datevisa", f"dv_{pid}"))
        with t0[2].popover("Supprimer cet ancien visa", icon=":material/delete:", width="stretch"):
            if st.button("Confirmer la suppression", key=f"delc_{pid}", type="primary"):
                core.delete_position(pid); st.session_state.open.discard(pid); st.rerun()
    else:
        t0 = st.columns(4)
        if t0[0].button("Nouveau visa", key=f"nv_{pid}", icon=":material/archive:", width="stretch",
                        help="Archive une copie de l'état actuel dans « Anciens visas »."):
            core.archive_position(pid); st.toast("Copie archivée dans Anciens visas"); st.rerun()
        dv = dt.date.fromisoformat(p["datevisa"]) if p.get("datevisa") else None
        t0[1].date_input("Date du visa", value=dv, key=f"dv_{pid}", format="DD/MM/YYYY",
                         on_change=save_field, args=(POS,i,"datevisa",f"dv_{pid}"))
        with t0[2]:
            if rappel:
                genre = {"retard": "retard", "proche": "proche"}.get(rappel["niveau"], "sb")
                st.markdown(badge(f"À viser par SB avant le {_date_fr(rappel['echeance'])}", genre, long=True),
                            unsafe_allow_html=True)
                st.caption("Dépôt à transférer (page Dépôts reçus)." if rappel["avant_transfert"]
                           else "Clos à l'émission du bordereau.")
            else:
                st.caption("Aucun document en attente de visa SB.")
        with t0[3].popover("Supprimer", icon=":material/delete:", width="stretch"):
            st.warning("Supprimer définitivement cette position ?")
            if st.button("Confirmer la suppression", key=f"del_{pid}", type="primary"):
                core.delete_position(pid); st.session_state.open.discard(pid); st.rerun()

        # --- association de trame ---
        need_trame = (p.get("famille") == core.TRAME_A_DEFINIR)
        if need_trame or not p.get("ctrl"):
            tnames = core.trame_names()
            sug = core.suggest_trame(p["des"])
            idx_sug = tnames.index(sug) if sug in tnames else 0
            tb1, tb2 = st.columns([2.4, 1])
            choix = tb1.selectbox("Trame d'attendus" if need_trame else "Trame (points de contrôle à charger)",
                                  tnames, index=idx_sug, key=f"tsel_{pid}",
                                  help="Présélection automatique (mots-clés + historique des validations).")
            if tb2.button("Valider la trame", key=f"tval_{pid}", width="stretch", icon=":material/verified:"):
                core.apply_trame(POS[i], choix)
                core.record_assoc(p["des"], choix)
                st.session_state[f"ep_{pid}"] = st.session_state.get(f"ep_{pid}", 0) + 1
                st.rerun()

    d = st.columns([1, 3])
    d[0].selectbox("Avis", core.AVIS, index=core.AVIS.index(p["avis"]), disabled=ro,
                   key=f"avis_{pid}", on_change=save_field, args=(POS,i,"avis",f"avis_{pid}"))
    d[1].markdown(f"<div style='padding-top:30px;color:#5C5C5C;font-size:13px'>Trame : {esc(p['famille'])}</div>",
                  unsafe_allow_html=True)
    if not is_plan:
        _propositions_marque_type(i, p, POS, ro)

    _documents_pdf(i, p, POS, projet, LOT, ro)
    _pieces_attendues(i, p, POS, ep, ro)
    _points_controle(i, p, POS, ep, ro)

    sous_titre("Remarques")
    st.text_area("Remarques / observations (libre) — reportées au bordereau", value=p.get("rem") or "",
                 height=200, disabled=ro, key=f"rm2_{pid}", on_change=save_field, args=(POS,i,"rem",f"rm2_{pid}"))
    st.text_area("Commentaire interne SB — non présent dans le bordereau", value=p.get("rem_histo") or "",
                 disabled=ro, key=f"rh_{pid}", on_change=save_field, args=(POS,i,"rem_histo",f"rh_{pid}"))


def _propositions_marque_type(i, p, POS, ro):
    pid = p["id"]
    # --- Proposition du Maître d'oeuvre (issue de l'import DPGF) ---
    # Affichée grisée par défaut lorsqu'elle provient du fichier
    # importé (mt_verrou = 1). Un interrupteur "Modifier" permet de
    # la corriger à la volée.
    a_une_proposition_moe = bool(p.get("marque_moe") or p.get("type_moe"))
    if a_une_proposition_moe or p.get("mt_verrou"):
        sous_titre("Proposition maître d'oeuvre (CCTP)")
        moe_cols = st.columns([2, 2, 1.1])
        modifier_moe = moe_cols[2].toggle("Modifier", value=not bool(p.get("mt_verrou")) and not ro,
                                          key=f"mtu_{pid}", disabled=ro,
                                          help="Déverrouille les champs de la proposition MOE.")
        champ_desactive = (not modifier_moe) or ro
        moe_cols[0].text_input("Marque MOE", value=p.get("marque_moe") or "",
                               key=f"mqmoe_{pid}", disabled=champ_desactive,
                               on_change=save_field, args=(POS, i, "marque_moe", f"mqmoe_{pid}"))
        moe_cols[1].text_input("Type MOE", value=p.get("type_moe") or "",
                               key=f"tpmoe_{pid}", disabled=champ_desactive,
                               on_change=save_field, args=(POS, i, "type_moe", f"tpmoe_{pid}"))
        # Si l'utilisateur active "Modifier", on lève le verrou en base
        # pour que le champ reste éditable aux prochains affichages.
        if modifier_moe and p.get("mt_verrou") and not ro:
            POS[i]["mt_verrou"] = 0
            core.save_position(POS[i])

    # --- Proposition de l'ENTREPRISE (seule exportée au bordereau) ---
    sous_titre("Proposition entreprise (retenue au visa)")
    m = st.columns(2)
    m[0].text_input("Marque", value=p.get("marque") or "", key=f"mq_{pid}", disabled=ro,
                    on_change=save_field, args=(POS,i,"marque",f"mq_{pid}"))
    m[1].text_input("Type", value=p.get("typeprop") or "", key=f"tp_{pid}", disabled=ro,
                    on_change=save_field, args=(POS,i,"typeprop",f"tp_{pid}"))
    if not ro:
        _deja_vise_ailleurs(p)


# Titre de l'encart : une information, pas une règle (pictogramme « information »).
ICONE_INFO = ("<svg class='dv-info' viewBox='0 0 16 16' aria-hidden='true'>"
              "<circle cx='8' cy='8' r='6.75' fill='none' stroke='currentColor' stroke-width='1.5'/>"
              "<path d='M8 7.2v4.3' stroke='currentColor' stroke-width='1.6' stroke-linecap='round'/>"
              "<circle cx='8' cy='4.9' r='1' fill='currentColor'/></svg>")
TITRE_DEJA_VISE = f"{ICONE_INFO}Pour info · Déjà visé ailleurs"


def _deja_vise_ailleurs(p):
    """Encart « Déjà visé ailleurs » : visas émis sur d'autres positions pour le
    même produit proposé (marque + type), ou un produit approchant (core.visas_du_meme_produit).
    Toujours affiché, pour qu'on sache où le trouver ; sans résultat, il dit pourquoi.
    Avec résultat, un avertissement rappelle de prendre du recul (chaque projet est différent)."""
    if not (p.get("marque") or "").strip() or not (p.get("typeprop") or "").strip():
        message = ("Renseignez la marque et le type proposés : les visas déjà rendus sur ce produit, "
                   "dans les autres projets et lots, s'afficheront ici.")
        visas = []
    else:
        message = "Aucun visa émis sur ce produit (ni sur un produit approchant) dans les autres projets et lots."
        visas = core.visas_du_meme_produit(p)
    if not visas:
        with st.container(border=True):
            st.markdown(f"<div class='dv-titre'>{TITRE_DEJA_VISE}</div><div class='dv-sous'>{message}</div>",
                        unsafe_allow_html=True)
        return
    lignes = []
    for visa in visas:
        if visa["ressemblance"] == "identique":
            produit = "même marque, même type"
        else:
            produit = f"produit approchant : {esc(visa['marque'])} {esc(visa['type'])}"
        remarque = f" · remarque : « {esc(visa['remarque'])} »" if visa["remarque"].strip() else ""
        # Responsable du visa précédent, à consulter : chef de projet et équipe du projet.
        responsable = f"Chef de projet : <b>{esc(visa['chef'] or '—')}</b>"
        if visa["equipe"]:
            responsable = responsable + f" · Équipe : <b>{esc(visa['equipe'])}</b>"
        lignes.append(f"<div class='dv-item'><div class='dv-ligne'>{badge_avis(visa['avis'])} "
                      f"<b>{esc(visa['projet'])}</b> · {esc(visa['lot'])} — {esc(visa['numero'])} "
                      f"du {_date_fr(visa['date'])}</div>"
                      f"<div class='dv-sous'>{responsable}</div>"
                      f"<div class='dv-sous'>{produit}{remarque}</div></div>")
    with st.container(border=True):
        st.markdown(f"<div class='dv-titre'>{TITRE_DEJA_VISE} — {len(visas)} "
                    f"{'visa' if len(visas) == 1 else 'visas'}, du plus récent au plus ancien</div>"
                    "<div class='dv-recul'>Prendre du recul : chaque projet est différent.</div>"
                    + "".join(lignes), unsafe_allow_html=True)


def _documents_pdf(i, p, POS, projet, LOT, ro):
    pid = p["id"]
    sous_titre("Documents")
    if not ro:
        with st.form(f"upf_{pid}", clear_on_submit=True):
            fu = st.columns([3, 2, 1])
            ups = fu[0].file_uploader("Glissez vos PDF ici", type=["pdf"], accept_multiple_files=True,
                                      key=f"up_{pid}", label_visibility="collapsed")
            att_sel = fu[1].multiselect("Pièces concernées (optionnel)", [x["n"] for x in p["att"]],
                                        key=f"ua_{pid}", label_visibility="collapsed",
                                        placeholder="Pièces concernées (sinon : toutes)")
            if fu[2].form_submit_button("Déposer", width="stretch", type="primary") and ups:
                attjson = json.dumps(att_sel, ensure_ascii=False) if att_sel else None
                for up_f in ups:
                    res = core.save_fichier(projet, LOT, p, up_f.name, up_f.getvalue(), attendu=attjson)
                    POS[i]["intitule"] = res["nom"]
                # Pièces rattachées (ou, à défaut, toutes celles pas encore reçues) : « À vérifier ».
                core.marquer_a_verifier(POS[i], att_sel)
                core.save_position(POS[i])
                rappel = core.rappel_ouvert(pid)
                if rappel:
                    st.session_state["toast"] = f"Rappel de visa SB : à viser avant le {_date_fr(rappel['echeance'])}"
                st.session_state[f"ep_{pid}"] = st.session_state.get(f"ep_{pid}", 0) + 1
                st.rerun()
        st.caption("Chaque dépôt passe les pièces concernées en « À vérifier » et ouvre un rappel de "
                   f"visa côté Solares Bauen ({LOT.get('delai_visa_sb') or core.DELAI_VISA_SB} jours ouvrés).")
    fichiers = core.get_fichiers(pid)
    if not fichiers:
        st.caption("Aucun document déposé.")
    for f_ in fichiers:
        fr = st.columns([4.35, 0.5, 0.5], vertical_alignment="center")
        try:
            attlist = json.loads(f_.get("attendu") or "[]")
            if isinstance(attlist, str): attlist = [attlist]
        except Exception:
            attlist = [f_["attendu"]] if f_.get("attendu") else []
        attc = "".join(f" <span class='bdg b-ok long'>{esc(x)}</span>" for x in attlist)
        version = f" · v{f_['version']}" if (f_.get("version") or 1) > 1 else ""
        fr[0].markdown(f"<div class='nom-fichier' style='font-size:14px;padding-top:4px'>{esc(f_['nom'])}{attc} "
                       f"<span style='color:#6B6B6B;font-size:12.5px'>· déposé le {_date_fr(f_.get('date_upload'))}{version}</span></div>",
                       unsafe_allow_html=True)
        data_f = core.read_fichier(f_)
        if data_f is not None:
            fr[1].download_button("", data_f, file_name=f_["nom"],
                                  mime=mimetypes.guess_type(f_["nom"])[0] or "application/octet-stream",
                                  key=f"dlf_{f_['id']}", help="Télécharger", icon=":material/download:")
        else:
            fr[1].markdown("<span title='Fichier introuvable' style='color:#E60116'>indisponible</span>", unsafe_allow_html=True)
        if not ro and fr[2].button("", key=f"delf_{f_['id']}", icon=":material/close:",
                                   help="Retirer de la liste (le contenu reste conservé)"):
            core.delete_fichier(f_["id"]); st.rerun()


def _pieces_attendues(i, p, POS, ep, ro):
    pid = p["id"]
    sous_titre("Pièces attendues")
    if not p["att"]:
        st.caption("Aucune pièce attendue.")
    for ai,a in enumerate(p["att"]):
        kx = f"{pid}_{ai}_{ep}"
        cc = st.columns([0.18, 1.3, 3, 2, 0.5], vertical_alignment="center")
        cc[0].markdown(f"<div><span class='dot' style='background:{STAT_COLOR[STAT_KIND.get(a['s'],'red')]}'></span></div>", unsafe_allow_html=True)
        cc[1].selectbox("Statut", core.STATUTS, index=core.STATUTS.index(a["s"]) if a["s"] in core.STATUTS else core.STATUTS.index("Manquant"),
                        key=f"as_{kx}", label_visibility="collapsed", disabled=ro,
                        on_change=save_att, args=(POS,i,ai,"s",f"as_{kx}"))
        cc[2].text_input("Pièce", value=a["n"], key=f"an_{kx}", label_visibility="collapsed", disabled=ro,
                         on_change=save_att, args=(POS,i,ai,"n",f"an_{kx}"))
        if a["s"] in core.STATUTS_FOURNIS:
            dval = dt.date.fromisoformat(a["d"]) if a.get("d") else None
            cc[3].date_input("Date reçue", value=dval, key=f"ad_{kx}", label_visibility="collapsed",
                             format="DD/MM/YYYY", disabled=ro, on_change=save_att, args=(POS,i,ai,"d",f"ad_{kx}"))
        else:
            cc[3].markdown("<div style='color:#BDBDBD;padding-top:6px'>—</div>", unsafe_allow_html=True)
        if not ro and cc[4].button("", key=f"rm_{pid}_{ai}", icon=":material/delete:", help="Retirer cette pièce"):
            POS[i]["att"].pop(ai); core.save_position(POS[i])
            # Nouvelles clés de widgets : sinon la pièce suivante, qui prend la
            # place de la pièce retirée, afficherait encore le nom et le statut de celle-ci.
            st.session_state[f"ep_{pid}"] = st.session_state.get(f"ep_{pid}", 0) + 1
            st.rerun()
    if not ro and st.button("Ajouter un attendu", key=f"aa_{pid}", icon=":material/add:"):
        POS[i]["att"].append({"n":"Nouvel attendu","s":"Manquant","d":""}); core.save_position(POS[i]); st.rerun()


def _points_controle(i, p, POS, ep, ro):
    pid = p["id"]
    if not p.get("ctrl"):
        return
    st.markdown("<div class='sous-titre'>Points de contrôle</div> "
                "<span style='color:#6B6B6B;font-size:12px'>interne — jamais exporté</span>", unsafe_allow_html=True)
    for ci, cpt in enumerate(p["ctrl"]):
        kc2 = f"{pid}_{ci}_{ep}"
        so = bool(cpt.get("so"))
        cr = st.columns([0.85, 0.75, 3.1, 2.3], vertical_alignment="center")
        cr[0].checkbox("Vérifié", value=bool(cpt.get("v")), key=f"cv_{kc2}", disabled=so or ro,
                       on_change=save_ctrl, args=(POS, i, ci, "v", f"cv_{kc2}"))
        cr[1].toggle("SO", value=so, key=f"cso_{kc2}", help="Sans objet", disabled=ro,
                     on_change=save_ctrl, args=(POS, i, ci, "so", f"cso_{kc2}"))
        style = "color:#BDBDBD;text-decoration:line-through" if so else ("color:#1E5E2E" if cpt.get("v") else "")
        cr[2].markdown(f"<div style='font-size:14px;{style}'>{esc(cpt['n'])}</div>", unsafe_allow_html=True)
        cr[3].text_input("Commentaire", value=cpt.get("c") or "", key=f"cc_{kc2}",
                         label_visibility="collapsed", placeholder="Commentaire…", disabled=so or ro,
                         on_change=save_ctrl, args=(POS, i, ci, "c", f"cc_{kc2}"))


# ---------- assemblage de la page ----------

def afficher(ctx):
    proj_code, LOT = ctx["proj_code"], ctx["lot"]
    if not ctx["chefs"]:
        st.info("Bienvenue. Ouvrez la page **Administration** pour créer votre premier chef de "
                "projet, ou charger les données de démonstration.")
        return
    if not proj_code:
        st.info("Créez un projet (page Administration → Projets) ou sélectionnez-en un.")
        return
    if not LOT:
        st.info("Créez un lot pour ce projet (page Administration → Lots).")
        return

    message = st.session_state.pop("toast", None)
    if message:
        st.toast(message, icon=":material/schedule:")

    projet = core.get_projet(proj_code)
    signataire = core.interlocuteur(projet)
    POS = core.get_positions(proj_code, LOT["id"])
    # Un document non prévu (sans position) reste sur la page Dépôts reçus jusqu'au transfert.
    rappels = [r for r in core.get_rappels_sb(lot_id=LOT["id"]) if r["position_id"]]
    rappels_par_pos = {r["position_id"]: r for r in rappels}

    contact = " ".join(x for x in [LOT.get("contact_civ"), LOT.get("contact_nom")] if x and LOT.get("contact_nom"))
    entete(proj_code, projet["nom"],
           [("Maître d'ouvrage", projet.get("moa")),
            ("Titulaire", LOT.get("entreprise")),
            ("Interlocuteur titulaire", contact),
            ("Interlocuteur Solares Bauen", signataire["nom"])],
           fil=core.libelle_lot(LOT))
    _bandeaux(proj_code, LOT, POS, rappels)
    _indicateurs(LOT, POS, rappels)
    _barre_actions(projet, proj_code, LOT, POS, signataire)
    _selection_rapide(proj_code, LOT, POS)
    _ajout_position(proj_code, LOT)

    cctp = [(i,p) for i,p in enumerate(POS) if (p.get("pos") or "").strip()]
    noref = [(i,p) for i,p in enumerate(POS) if not (p.get("pos") or "").strip()]
    ARCH = core.get_positions(proj_code, LOT["id"], archive=1)
    anciens = list(enumerate(ARCH))

    # Volet choisi d'abord (grands boutons), puis les filtres, qui s'appliquent au volet.
    sous_titre("Volets du lot")
    zone_volets = st.container(key="volets_suivi")
    visible, filtre_actif = _filtres(POS, rappels_par_pos)
    vus = {"cctp": [(i, p) for i, p in cctp if visible(p)],
           "plans": [(i, p) for i, p in noref if visible(p)],
           "anciens": [(i, p) for i, p in anciens if visible(p)]}
    tous = {"cctp": cctp, "plans": noref, "anciens": anciens}
    with zone_volets:
        volet = _choix_volet(vus, tous, filtre_actif)

    if volet == "cctp":
        if vus["cctp"]:
            for i,p in vus["cctp"]: render_row(i, p, POS, projet, LOT, rappels_par_pos)
        elif cctp:
            st.caption(f"Aucune position ne correspond au filtre ({filtre_actif}).")
        else:
            st.caption("Aucune position référencée — importez la DPGF ou ajoutez une position.")
    elif volet == "plans":
        if vus["plans"]:
            for i,p in vus["plans"]: render_row(i, p, POS, projet, LOT, rappels_par_pos)
        elif noref:
            st.caption(f"Aucun plan ni note de calcul ne correspond au filtre ({filtre_actif}).")
        else:
            st.caption("Aucun document sans référence — ajoutez un plan ou une note de calcul (référence vide).")
    else:
        if vus["anciens"]:
            for i, pa in vus["anciens"]: render_row(i, pa, ARCH, projet, LOT, {}, archive=True)
        elif ARCH:
            st.caption(f"Aucun ancien visa ne correspond au filtre ({filtre_actif}).")
        else:
            st.caption("Aucun visa archivé. Le bouton « Nouveau visa » d'une position en dépose une copie ici.")
