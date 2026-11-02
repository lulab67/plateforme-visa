# ============================================================================
#  ui/bordereaux.py — registre des bordereaux émis d'un lot (fenêtre modale)
#  Pour chaque bordereau : numéro, date, contenu (positions et avis), et les
#  actions renuméroter / réexporter (Excel, PDF) / supprimer.
#  Utilisé depuis le Dashboard et depuis la page Suivi des visas.
# ============================================================================

import datetime as dt
import os
import streamlit as st
import core
from ui.style import badge_avis, esc

TEMPLATE_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), "templates", "VISA_TYPE.xlsx")
MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _date_heure(texte_iso):
    if not texte_iso:
        return ""
    return dt.datetime.fromisoformat(texte_iso).strftime("%d/%m/%Y à %H:%M")


def nom_fichier(bordereau):
    """Nom de fichier d'un bordereau : VISA_<projet>_Lot<n>_N<numéro>."""
    lot = bordereau["contenu"].get("lot", {})
    return f"VISA_{bordereau['projet']}_Lot{lot.get('lot_numero') or ''}_N{int(bordereau['numero']):02d}"


def preparer_exports(bordereau):
    """Génère l'Excel (et le PDF si LibreOffice est présent) d'un bordereau et
    les garde en session pour les boutons de téléchargement."""
    with open(TEMPLATE_PATH, "rb") as gabarit:
        xlsx = core.export_bordereau(bordereau, gabarit.read())
    pdf, message = core.convertir_en_pdf(xlsx, nom_fichier(bordereau))
    st.session_state[f"bx_{bordereau['id']}"] = {"xlsx": xlsx, "pdf": pdf, "message": message,
                                                "numero": bordereau["numero"]}


def boutons_telechargement(bordereau, prefixe="reg"):
    """Boutons Excel / PDF d'un bordereau déjà préparé (rien sinon).
    'prefixe' distingue les clés quand les mêmes boutons apparaissent à deux
    endroits (encart « dernier bordereau » de la page et registre)."""
    pret = st.session_state.get(f"bx_{bordereau['id']}")
    if not pret or pret["numero"] != bordereau["numero"]:
        return False
    nom = nom_fichier(bordereau)
    c1, c2 = st.columns(2)
    c1.download_button("Excel", pret["xlsx"], file_name=f"{nom}.xlsx", mime=MIME_XLSX,
                       width="stretch", icon=":material/table:", key=f"{prefixe}_dlx_{bordereau['id']}")
    if pret["pdf"]:
        c2.download_button("PDF", pret["pdf"], file_name=f"{nom}.pdf", mime="application/pdf",
                           width="stretch", icon=":material/picture_as_pdf:", key=f"{prefixe}_dlp_{bordereau['id']}")
    else:
        c2.button("PDF", disabled=True, width="stretch", icon=":material/picture_as_pdf:",
                  help=pret["message"], key=f"{prefixe}_nop_{bordereau['id']}")
    return True


def _contenu(bordereau):
    """Tableau des positions visées dans le bordereau."""
    lignes = []
    for p in bordereau["contenu"].get("positions", []):
        ref = esc((p.get("pos") or "").strip()) or "—"
        date_visa = dt.date.fromisoformat(p["datevisa"]).strftime("%d/%m/%Y") if p.get("datevisa") else "—"
        lignes.append(f"<tr><td style='white-space:nowrap;font-weight:700'>{ref}</td><td>{esc(p.get('des'))}</td>"
                      f"<td>{badge_avis(p.get('avis'))}</td><td style='white-space:nowrap'>{date_visa}</td></tr>")
    if not lignes:
        st.caption("Aucune position dans ce bordereau.")
        return
    st.markdown("<table style='width:100%;font-size:13px;border-collapse:collapse'>"
                "<tr style='color:#6B6B6B;text-align:left'><th>Pos.</th><th>Désignation</th><th>Avis</th>"
                "<th>Date du visa</th></tr>" + "".join(lignes) + "</table>", unsafe_allow_html=True)


def bouton_registre(libelle, lot_id, key):
    """Bouton qui ouvre le registre d'un lot. Le registre se rouvre aussi tout
    seul après une renumérotation ou une suppression (voir _apres_modification)."""
    clique = st.button(libelle, key=key, icon=":material/history:", width="stretch")
    if clique or st.session_state.get("_registre_a_rouvrir") == lot_id:
        st.session_state.pop("_registre_a_rouvrir", None)
        registre(lot_id)


def _apres_modification(lot_id):
    """Relance toute la page (tuiles et compteurs à jour) puis rouvre le registre."""
    st.session_state["_registre_a_rouvrir"] = lot_id
    st.rerun()


@st.dialog("Bordereaux émis", width="large")
def registre(lot_id):
    """Fenêtre du registre des bordereaux d'un lot."""
    lot = core.get_lot(lot_id)
    st.markdown(f"**{esc(core.libelle_lot(lot))}** — {esc(lot.get('entreprise') or '')}")
    st.caption(f"Prochain numéro : {core.format_numero(core.prochain_numero(lot_id))} "
               "(dernier bordereau émis + 1).")
    bordereaux = core.get_bordereaux(lot_id)
    if not bordereaux:
        st.info("Aucun bordereau émis pour ce lot. « Générer le bordereau de visa » (page Suivi) en émet un.")
        return
    for b in bordereaux:
        with st.container(border=True):
            h1, h2 = st.columns([3, 2], vertical_alignment="center")
            nb = len(b["contenu"].get("positions", []))
            h1.markdown(f"<div style='font-size:17px;font-weight:700'>{esc(b['libelle'])}</div>"
                        f"<div style='color:#5C5C5C;font-size:13px'>émis le {_date_heure(b['date'])} · "
                        f"{nb} position(s)</div>", unsafe_allow_html=True)
            with h2:
                if not boutons_telechargement(b):
                    if st.button("Réexporter", key=f"rex_{b['id']}", icon=":material/download:",
                                 width="stretch"):
                        preparer_exports(b)
                        st.rerun(scope="fragment")
            with st.expander("Contenu visé"):
                _contenu(b)
            a1, a2, a3, a4 = st.columns([1.1, 1.2, 1.6, 1.2], vertical_alignment="bottom")
            nouveau = a1.number_input("Numéro", min_value=1, value=int(b["numero"]), step=1, key=f"num_{b['id']}")
            if a2.button("Renuméroter", key=f"ren_{b['id']}", disabled=nouveau == b["numero"],
                         width="stretch"):
                ok, message = core.set_numero_bordereau(b["id"], nouveau)
                if ok:
                    _apres_modification(lot_id)
                else:
                    st.error(message)
            confirmer = a3.checkbox("Confirmer la suppression", key=f"cdel_{b['id']}")
            if a4.button("Supprimer", key=f"del_{b['id']}", disabled=not confirmer, icon=":material/delete:",
                         width="stretch"):
                core.delete_bordereau(b["id"])
                _apres_modification(lot_id)
