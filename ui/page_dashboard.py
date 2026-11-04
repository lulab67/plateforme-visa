# ============================================================================
#  ui/page_dashboard.py — page « Dashboard » du chef de projet, en deux
#  sections bien distinctes :
#    1. À viser par Solares Bauen : documents déposés en attente de NOTRE visa
#       (échéance « à viser avant le », clôture manuelle) ;
#    2. Relances titulaires : pièces manquantes / à corriger côté entreprises.
#  Survoler une étiquette affiche la liste des éléments concernés.
#  Le portefeuille s'affiche par chef de projet ou par équipe ; le portefeuille
#  « Solares » (toute l'entreprise) a son propre Dashboard : ui/page_solares.py.
# ============================================================================

import datetime as dt
import streamlit as st
import core
from ui import bordereaux as bordereaux_ui
from ui.style import badge, entete, esc, mini_jauge, sous_titre, survol, tuiles
from ui.sidebar import MODE_SOLARES, demander_selection
from ui import page_solares

GCOL = {"VSO":"#4C9A5B","VAO":"#E2B33C","VAOB":"#E5793B","REF":"#E60116","MQ":"#7E5BB5","SO":"#B5B5B5"}
NATURE = {"Plan": "Plan", "NdC": "NdC", "Document": "Document", "Matériel": "Position"}


def _date_fr(texte_iso, court=False):
    if not texte_iso:
        return ""
    return dt.date.fromisoformat(texte_iso).strftime("%d/%m" if court else "%d/%m/%Y")


def _fix_last_date(lot_id, key):
    """Callback : corrige la date de dernière relance saisie au dashboard."""
    val = st.session_state.get(key)
    if val: core.update_last_relance_date(lot_id, val.isoformat())


def _libelle_element(rappel):
    """« Position 3.2.2 — CTA restauration » ou « Plan — Plan chaufferie »."""
    nature = NATURE.get(rappel.get("ptype") or "", "Position")
    ref = (rappel.get("pos") or "").strip()
    if ref:
        return f"{nature} {esc(ref)} — {esc(rappel.get('des'))}"
    return f"{nature} — {esc(rappel.get('des'))}"


# ---------- section 1 : à viser par Solares Bauen ----------

def _section_visa_sb(projets):
    sous_titre("À viser par Solares Bauen")
    st.caption("Documents déposés par les titulaires, en attente de notre visa. "
               "Survolez l'échéance pour voir la liste. L'émission du bordereau clôt le rappel.")
    lignes = 0
    for pr in projets:
        for l in core.get_lots(pr["code"]):
            rappels = core.get_rappels_sb(lot_id=l["id"])
            if not rappels:
                continue
            lignes += 1
            plus_urgent = rappels[0]
            genre = plus_urgent["niveau"] if plus_urgent["niveau"] != "ok" else "sb"
            en_retard = sum(1 for r in rappels if r["niveau"] == "retard")
            elements = []
            for r in rappels:
                retard = " <b style='color:#FF6B6B'>· en retard</b>" if r["niveau"] == "retard" else ""
                a_transferer = " · à transférer" if r["avant_transfert"] else ""
                elements.append(f"{_libelle_element(r)}<br><span class='b-gris'>déposé le {_date_fr(r['date_depot'])}"
                                f" · à viser avant le {_date_fr(r['echeance'])}{a_transferer}</span>{retard}")
            etiquette = badge(f"{len(rappels)} à viser · avant le {_date_fr(plus_urgent['echeance'], court=True)}", genre)
            with st.container(border=True):
                c1, c2, c3 = st.columns([2.6, 1.6, 0.9])
                c1.markdown(f"<div style='font-weight:700;font-size:15px'>{esc(core.libelle_lot(l))}</div>"
                            f"<div style='color:#5C5C5C;font-size:13px'>{esc(pr['code'])} — {esc(pr['nom'])} · "
                            f"{esc(l.get('entreprise') or 'titulaire n.d.')}</div>", unsafe_allow_html=True)
                retard_txt = f"<div style='color:#C00000;font-size:12.5px;margin-top:4px'>{en_retard} en retard</div>" if en_retard else ""
                c2.markdown(survol(etiquette, f"Documents à viser — {core.libelle_lot(l)}", elements) + retard_txt,
                            unsafe_allow_html=True)
                if c3.button("Ouvrir", key=f"ouvrir_sb_{l['id']}", icon=":material/fact_check:",
                             width="stretch", help="Ouvrir ce lot dans le Suivi des visas"):
                    demander_selection(projet=pr, lot=l, page="Suivi des visas")
                    st.rerun()
    if lignes == 0:
        st.success("Aucun document en attente de visa côté Solares Bauen.", icon=":material/task_alt:")


# ---------- section 2 : relances titulaires ----------

def _jauge(posl):
    tot = len(posl)
    segs = []
    for av in ["VSO","VAO","VAOB","REF","MQ","SO"]:
        n = sum(1 for p in posl if p["avis"] == av)
        if not n: continue
        pct = n/tot*100
        inner = f"{av} {pct:.0f}%" if pct >= 14 else ""
        segs.append(
            f"<div title=\"{av} : {n} position(s) — {pct:.0f} %\" style='width:{pct:.2f}%;background:{GCOL[av]};"
            f"display:flex;align-items:center;justify-content:center;color:#fff;"
            f"font-size:11px;font-weight:700;white-space:nowrap;overflow:hidden'>{inner}</div>")
    return ("<div style='display:flex;height:24px;border-radius:6px;overflow:hidden;"
            "margin-top:8px;border:1px solid #E4E4E4'>" + "".join(segs) + "</div>")


def _bouton_relancer(pr, l, posl):
    """Bouton « Relancer » : brouillon .eml de toutes les positions MQ/VAOB/REF
    du lot (celles listées au survol), inscrit au journal au téléchargement."""
    candidates = core.positions_a_relancer(posl)
    if st.download_button("Relancer", core.build_eml(pr, l, core.copies_pour_relance(posl)),
                          file_name=f"relance_{pr['code']}_Lot{l.get('lot_numero') or l['id']}.eml",
                          mime="message/rfc822", key=f"rel_dash_{l['id']}", icon=":material/mail:",
                          width="stretch", help="Brouillon de relance des positions MQ / VAOB / REF du lot."):
        core.log_relance(pr["code"], l["id"], 1, [(p["pos"] or p["des"]) for p in candidates])
        st.toast(f"Relance {core.libelle_lot(l)} enregistrée au journal"); st.rerun()


def _section_relances(projets, chefs_par_code=None):
    sous_titre("Relances titulaires")
    st.caption("Pièces manquantes (MQ), à corriger (VAOB) ou refusées (REF) côté entreprises. "
               "Survolez l'état de relance pour voir les positions concernées.")
    for pr in projets:
        lots_pr = core.get_lots(pr["code"])
        chef = f" · {esc(chefs_par_code.get(pr.get('chef'), ''))}" if chefs_par_code else ""
        st.markdown(f"<div style='font-size:16px;font-weight:700;margin:16px 0 6px 0'>"
                    f"<span class='ent-code' style='margin-right:8px;font-size:12px'>{esc(pr['code'])}</span>"
                    f"{esc(pr['nom'])}<span style='color:#6B6B6B;font-weight:400;font-size:14px'>{chef}</span></div>",
                    unsafe_allow_html=True)
        if not lots_pr:
            st.caption("— aucun lot —")
        for l in lots_pr:
            posl = core.get_positions(pr["code"], l["id"])
            tot = len(posl)
            with st.container(border=True):
                cA, cB, cC = st.columns([2.4, 1.3, 1.3], vertical_alignment="top")
                with cA:
                    valid_l = sum(1 for p in posl if p["avis"] in ("VSO","VAO"))
                    horsso = sum(1 for p in posl if p["avis"] != "SO")
                    taux = round(valid_l/horsso*100) if horsso else 0
                    # Titre du lot à gauche, mini jauge « % validé » toujours alignée à droite.
                    st.markdown("<div style='display:flex;justify-content:space-between;align-items:center;gap:12px'>"
                                f"<div style='font-weight:700;font-size:14.5px'>{esc(core.libelle_lot(l))} "
                                f"<span style='color:#5C5C5C;font-weight:400'>· {esc(l.get('entreprise') or 'titulaire n.d.')}"
                                f" · {tot} position(s)</span></div>"
                                + (mini_jauge(taux) if horsso else "")
                                + "</div>"
                                + (_jauge(posl) if tot else ""), unsafe_allow_html=True)
                    if not tot:
                        st.caption("Aucune position dans ce lot.")
                with cB:
                    due_l, last_l, nxt_l, ncand_l = core.relance_status(pr["code"], l, posl)
                    elements = []
                    for p in core.positions_a_relancer(posl):
                        ref = (p.get("pos") or "").strip()
                        elements.append(f"<b>{esc(p['avis'])}</b> · {esc(ref) + ' — ' if ref else ''}{esc(p['des'])}")
                    if not ncand_l:
                        st.markdown(badge("Rien à relancer", "ok"), unsafe_allow_html=True)
                    else:
                        if due_l:
                            etiquette = badge(f"Relance à faire · {ncand_l} pos.", "retard")
                        else:
                            etiquette = badge(f"Prochaine relance entreprise le {_date_fr(nxt_l, court=True)} · {ncand_l} pos.", "proche")
                        st.markdown(survol(etiquette, f"À relancer — {core.libelle_lot(l)}", elements),
                                    unsafe_allow_html=True)
                    if last_l:
                        st.date_input("Dernière relance", value=dt.date.fromisoformat(last_l),
                                      key=f"dl_{l['id']}", format="DD/MM/YYYY",
                                      on_change=_fix_last_date, args=(l["id"], f"dl_{l['id']}"))
                    else:
                        st.caption("Aucune relance enregistrée.")
                with cC:
                    if ncand_l:
                        _bouton_relancer(pr, l, posl)
                    nb_bx = len(core.get_bordereaux(l["id"]))
                    bordereaux_ui.bouton_registre(f"Bordereaux ({nb_bx})", l["id"], key=f"bxd_{l['id']}")


# ---------- page ----------

def _choisir_portefeuille(ctx):
    """Portefeuille choisi dans le bandeau de la barre latérale (chef de projet
    ou équipe). Renvoie (titre, méta de l'en-tête, projets actifs, projets
    archivés, chefs par code — None pour un seul chef)."""
    portefeuille = ctx["portefeuille"]
    chefs_par_code = {c["code"]: c["nom"] for c in ctx["chefs"]}
    equipe = portefeuille["equipe"]
    if portefeuille["mode"] == "Équipe":
        membres = ", ".join(chefs_par_code[c] for c in equipe["chefs"]) or "aucun chef de projet"
        return (f"Portefeuille — {equipe['nom']}", [("Chefs de projet", membres)],
                portefeuille["projets"], portefeuille["archives"], chefs_par_code)
    meta = [("Équipe", equipe["nom"])] if equipe else []
    return (f"Portefeuille — {portefeuille['nom']}", meta, portefeuille["projets"],
            portefeuille["archives"], None)


def afficher(ctx):
    # Portefeuille « Solares » (toute l'entreprise) : un autre Dashboard, celui des délais de visa.
    if ctx["portefeuille"] and ctx["portefeuille"]["mode"] == MODE_SOLARES:
        page_solares.afficher(ctx)
        return
    if not ctx["chefs"]:
        entete("", "Portefeuille")
        st.info("Aucun projet — créez un chef de projet puis un projet (page Administration).")
        return
    titre, meta, projs, arch_projs, chefs_par_code = _choisir_portefeuille(ctx)
    entete("", titre, meta + [("Projets actifs", str(len(projs)))])
    if not projs:
        st.info("Aucun projet actif dans ce portefeuille.")
        return

    # --- indicateurs de synthèse ---
    rappels = []
    a_relancer = 0
    for pr in projs:
        rappels.extend(core.get_rappels_sb(projet=pr["code"]))
        for l in core.get_lots(pr["code"]):
            if core.relance_status(pr["code"], l, core.get_positions(pr["code"], l["id"]))[0]:
                a_relancer += 1
    en_retard = sum(1 for r in rappels if r["niveau"] == "retard")
    chefs_du_portefeuille = sorted({pr["chef"] for pr in projs})
    a_transferer = core.nb_lignes_en_attente(chefs_du_portefeuille)
    tuiles([
        {"lab": "Dépôts à transférer", "val": a_transferer, "unite": "lignes", "alerte": a_transferer > 0,
         "sub": f"reçus sur la {core.NOM_PORTAIL} — page Dépôts reçus"},
        {"lab": "À viser par SB", "val": len(rappels), "unite": "doc.", "sub": "documents déposés en attente de visa"},
        {"lab": "Visas SB en retard", "val": en_retard, "alerte": en_retard > 0, "sub": "échéance dépassée"},
        {"lab": "Entreprises à relancer", "val": a_relancer, "unite": "lots", "alerte": a_relancer > 0,
         "sub": "fréquence de relance atteinte"},
    ])

    _section_visa_sb(projs)
    st.divider()
    _section_relances(projs, chefs_par_code)

    if arch_projs:
        st.divider()
        if st.toggle(f"Afficher les projets archivés ({len(arch_projs)})", value=False):
            _section_relances(arch_projs, chefs_par_code)
