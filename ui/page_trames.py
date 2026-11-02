# ============================================================================
#  ui/page_trames.py — page « Trames de visa » : référentiel des documents
#  attendus, points de contrôle et mots-clés d'association.
#  Le nom d'une trame est modifiable ; le renommage est reporté sur les
#  positions qui l'utilisent. Toutes les trames peuvent être remplacées par
#  un tableau Excel (colonnes A-D, voir core.lire_trames_excel).
# ============================================================================

import streamlit as st
import core
from ui.style import entete, sous_titre


def afficher(ctx):
    entete("", "Trames de visa",
           [("", "Documents attendus (reportés au bordereau) · points de contrôle (internes) · "
                 "mots-clés d'association automatique.")])
    message = st.session_state.pop("trame_msg", None)
    if message:
        st.success(message, icon=":material/check_circle:")

    trames = core.get_trames()
    noms = [t["nom"] for t in trames]
    choix_demande = st.session_state.pop("trame_choisie", None)
    if choix_demande in noms:
        st.session_state["sel_trame"] = choix_demande
    # Liste des trames et ses trois actions à côté ; actions plus larges que la liste
    # pour que « Supprimer » et sa flèche ne se chevauchent pas sur un écran étroit.
    c_sel, c_btn = st.columns([1.6, 3], vertical_alignment="bottom")
    tsel = c_sel.selectbox("Trame", noms, key="sel_trame")
    T = next(t for t in trames if t["nom"] == tsel)
    systeme = T["nom"] in core.TRAMES_SYSTEME
    with c_btn:
        b = st.columns([1, 1, 1.3])   # « Supprimer » porte aussi la flèche du menu
        if b[0].button("Nouvelle", icon=":material/add:", width="stretch"):
            st.session_state["trame_choisie"] = core.add_trame("Nouvelle trame"); st.rerun()
        if b[1].button("Dupliquer", icon=":material/content_copy:", width="stretch"):
            core.duplicate_trame(T["id"]); st.rerun()
        with b[2].popover("Supprimer", icon=":material/delete:", width="stretch", disabled=systeme):
            st.caption("Les positions existantes conservent leurs listes.")
            if st.button("Confirmer la suppression", type="primary"):
                ok, msg = core.delete_trame(T["id"])
                if ok: st.rerun()
                else: st.error(msg)

    with st.form(f"trame_{T['id']}"):
        nomt = st.text_input("Nom de la trame", value=T["nom"], disabled=systeme,
                             help="Trame utilisée par l'application : non renommable." if systeme
                             else "Le nouveau nom est reporté sur les positions qui utilisent cette trame.")
        c1, c2 = st.columns(2)
        docs_txt = c1.text_area("Documents attendus — un par ligne", value="\n".join(T["docs"]), height=320,
                                help="Alimente les « Pièces attendues » (reportées au bordereau).")
        ctrl_txt = c2.text_area("Points de contrôle — un par ligne",
                                value="\n".join(c["n"] for c in T["controles"]), height=320,
                                help="Checklist interne (vérifié / sans objet + commentaire). Jamais exportée.")
        kw_txt = st.text_area("Mots-clés d'association — un par ligne", value="\n".join(T.get("keywords") or []),
                              height=100, help="Présélectionnent cette trame quand la désignation d'une position les contient.")
        if st.form_submit_button("Enregistrer la trame", type="primary", icon=":material/save:"):
            docs = [l.strip() for l in docs_txt.splitlines() if l.strip()]
            ctrls = [{"n": l.strip()} for l in ctrl_txt.splitlines() if l.strip()]
            kws = [l.strip() for l in kw_txt.splitlines() if l.strip()]
            nouveau_nom = T["nom"] if systeme else nomt
            ok, msg = core.save_trame(T["id"], nouveau_nom, docs, ctrls, kws)
            if ok:
                st.session_state["trame_choisie"] = nouveau_nom.strip()
                st.session_state["trame_msg"] = "Trame enregistrée."
                st.rerun()
            else:
                st.error(msg)
    st.caption("Réordonner = déplacer les lignes. Les positions déjà créées conservent leurs listes "
               "(seul le nom de trame est mis à jour).")
    _import_excel(trames)


def _import_excel(trames):
    """Remplacement de toutes les trames par un tableau Excel (core.lire_trames_excel)."""
    with st.expander("Remplacer toutes les trames par un tableau Excel", icon=":material/upload_file:"):
        st.markdown("Une ligne d'en-tête, puis une trame par ligne. Colonnes lues : **A** nom de la trame, "
                    "**B** documents attendus, **C** points de contrôle, **D** mots-clés d'association, "
                    "un élément par ligne de la cellule (Alt+Entrée). Onglet « Trames » s'il existe, sinon "
                    "le premier ; les colonnes suivantes sont ignorées.")
        st.caption("« Autre document » et « À définir » sont conservées. Les positions déjà créées gardent "
                   "leurs pièces attendues et leurs points de contrôle.")
        # Clé renouvelée après un import : le fichier déposé disparaît du sélecteur.
        essai = st.session_state.get("trames_import_n", 0)
        fichier = st.file_uploader("Tableau des trames (.xlsx)", type=["xlsx"], key=f"trames_xlsx_{essai}")
        if fichier is None:
            return
        try:
            lues = core.lire_trames_excel(fichier.getvalue())
        except Exception:
            st.error("Ce fichier n'a pas pu être lu comme un tableau Excel.", icon=":material/error:")
            return
        if not lues:
            st.warning("Aucune trame trouvée : la colonne A (nom de la trame) est vide.", icon=":material/warning:")
            return
        actuelles = [t for t in trames if t["nom"] not in core.TRAMES_SYSTEME]
        apercu = ", ".join(t["nom"] for t in lues[:4])
        st.caption(f"{len(lues)} trames lues : {apercu}…")
        if st.button(f"Remplacer les {len(actuelles)} trames actuelles par ces {len(lues)} trames",
                     type="primary", icon=":material/sync:"):
            nombre = core.remplacer_trames(lues)
            st.session_state["trame_choisie"] = lues[0]["nom"]
            st.session_state["trames_import_n"] = essai + 1
            st.session_state["trame_msg"] = f"{nombre} trames importées depuis « {fichier.name} »."
            st.rerun()
