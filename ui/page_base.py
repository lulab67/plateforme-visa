# ============================================================================
#  ui/page_base.py — page « Base de données » : vue transversale des visas
#  CCTP de tous les projets, avec filtres cumulatifs et export CSV.
# ============================================================================

import pandas as pd
import streamlit as st
import core
from ui.style import entete


def _choix(df, colonne, cle):
    """Choix proposés par un filtre : valeurs non vides de la colonne, plus
    celles déjà retenues (elles restent visibles quand un autre filtre les écarte)."""
    valeurs = set(x for x in df[colonne].unique() if x)
    valeurs.update(st.session_state.get(cle) or [])
    return sorted(valeurs)


def _annees(df):
    """Années des codes projet, la plus récente d'abord (un code sans année n'en propose pas)."""
    return sorted((a for a in df["Année"].unique() if a), reverse=True)


def _contient(df, colonne, recherche):
    """Lignes dont la colonne contient le texte cherché (accents et majuscules ignorés)."""
    masque = df[colonne].map(lambda texte: core.contient_texte(texte, recherche)).astype(bool)
    return df[masque]


def afficher(ctx):
    entete("", "Visas — positions CCTP/Matériel, tous projets",
           [("", "Dernier visa de chaque position (les copies « Anciens visas » et les plans/NdC sont exclus). Filtres cumulatifs.")])
    rows = core.db_visas()
    if not rows:
        st.info("Aucune position CCTP/Matériel en base pour le moment.")
        return
    df = pd.DataFrame(rows)
    df["Date visa"] = pd.to_datetime(df["Date visa"], errors="coerce")
    with st.expander("Filtres", icon=":material/filter_alt:", expanded=True):
        r1 = st.columns(4)
        f_annee = r1[0].multiselect("Année (code projet)", _annees(df), key="bd_annee", placeholder="Toutes",
                                    help="4 premiers chiffres du code projet : 2026008_GS_KUNHEIM → 2026.")
        df_annees = df[df["Année"].isin(f_annee)] if f_annee else df
        f_proj = r1[1].multiselect("Projet", _choix(df_annees, "Code projet", "bd_projet"), key="bd_projet",
                                   placeholder="Tous")
        f_chef = r1[2].multiselect("Chef de projet", _choix(df, "Chef de projet", "bd_chef"), key="bd_chef",
                                   placeholder="Tous")
        f_arch = r1[3].selectbox("Projets archivés", ["Inclure", "Exclure", "Uniquement"], index=0, key="bd_archives")
        r2 = st.columns(4)
        f_lot = r2[0].multiselect("Lot", _choix(df, "Lot", "bd_lot"), key="bd_lot", placeholder="Tous")
        f_ent = r2[1].multiselect("Entreprise", _choix(df, "Entreprise", "bd_entreprise"), key="bd_entreprise",
                                  placeholder="Toutes")
        f_trame = r2[2].multiselect("Trame", _choix(df, "Trame", "bd_trame"), key="bd_trame", placeholder="Toutes")
        f_statut = r2[3].multiselect("Statut visa", ["VSO","VAO","VAOB","REF","MQ","SO"], key="bd_statut",
                                     placeholder="Tous")
        # Prescription du maître d'oeuvre (import DPGF) puis proposition de l'entreprise ;
        # les types proposés se limitent aux marques retenues juste à gauche.
        r3 = st.columns(4)
        f_marque_moe = r3[0].multiselect("Marque prescrite", _choix(df, "Marque prescrite", "bd_marque_moe"),
                                         key="bd_marque_moe", placeholder="Toutes",
                                         help="Prescription du maître d'oeuvre (CCTP / DPGF importée).")
        df_marques_moe = df[df["Marque prescrite"].isin(f_marque_moe)] if f_marque_moe else df
        f_type_moe = r3[1].multiselect("Type prescrit", _choix(df_marques_moe, "Type prescrit", "bd_type_moe"),
                                       key="bd_type_moe", placeholder="Tous",
                                       help="Prescription du maître d'oeuvre ; limité aux marques prescrites retenues.")
        f_marque = r3[2].multiselect("Marque proposée", _choix(df, "Marque proposée", "bd_marque"), key="bd_marque",
                                     placeholder="Toutes", help="Proposition de l'entreprise (celle du bordereau).")
        df_marques = df[df["Marque proposée"].isin(f_marque)] if f_marque else df
        f_type = r3[3].multiselect("Type proposé", _choix(df_marques, "Type proposé", "bd_type"), key="bd_type",
                                   placeholder="Tous",
                                   help="Proposition de l'entreprise ; limité aux marques proposées retenues.")
        r4 = st.columns(2)
        f_txt = r4[0].text_input("Intitulé position contient…", key="bd_intitule")
        f_com = r4[1].text_input("Commentaire contient…", key="bd_commentaire")
    if f_annee: df = df[df["Année"].isin(f_annee)]
    if f_proj: df = df[df["Code projet"].isin(f_proj)]
    if f_chef: df = df[df["Chef de projet"].isin(f_chef)]
    if f_lot: df = df[df["Lot"].isin(f_lot)]
    if f_ent: df = df[df["Entreprise"].isin(f_ent)]
    if f_marque_moe: df = df[df["Marque prescrite"].isin(f_marque_moe)]
    if f_type_moe: df = df[df["Type prescrit"].isin(f_type_moe)]
    if f_marque: df = df[df["Marque proposée"].isin(f_marque)]
    if f_type: df = df[df["Type proposé"].isin(f_type)]
    if f_trame: df = df[df["Trame"].isin(f_trame)]
    if f_statut: df = df[df["Statut visa"].isin(f_statut)]
    if f_txt: df = _contient(df, "Intitulé position", f_txt)
    if f_com: df = _contient(df, "Commentaire", f_com)
    if f_arch == "Exclure": df = df[df["Projet archivé"] == "Non"]
    elif f_arch == "Uniquement": df = df[df["Projet archivé"] == "Oui"]
    st.caption(f"{len(df)} position(s) affichée(s).")
    st.dataframe(df, width="stretch", hide_index=True, height=520,
                 column_config={
                     "Année": st.column_config.TextColumn("Année", width="small"),
                     "Date visa":st.column_config.DateColumn("Date visa", format="DD/MM/YYYY"),
                     "Nb relances": st.column_config.NumberColumn("Nb relances", width="small"),
                     "Nb itérations visa": st.column_config.NumberColumn("Nb itérations visa", width="small",
                         help="Nombre de cycles de visa archivés (indicateur de qualité de soumission)."),
                 })
    # Export : un texte commençant par = + - @ n'est jamais lu comme une formule par le tableur.
    export = df.copy()
    for colonne in export.columns:
        if export[colonne].dtype == object:
            export[colonne] = export[colonne].map(core.texte_sur_pour_tableur)
    st.download_button("Exporter la sélection (CSV)", export.to_csv(index=False, sep=";").encode("utf-8-sig"),
                       file_name="base_visas.csv", mime="text/csv", icon=":material/download:")
