# ============================================================================
#  ui/page_solares.py — Dashboard du portefeuille « Solares » (toute
#  l'entreprise) : délais de visa de Solares Bauen, du dépôt d'un document à
#  l'émission du bordereau qui le vise, en jours ouvrés d'Alsace-Moselle
#  (core.delais_de_visa). Tuiles, puis, mois par mois sur 12 mois : taux de
#  visa dans le délai et délai médian.
#  Toute l'entreprise seulement : ni filtre ni comparaison par équipe ou par
#  chef de projet (jugé trop sensible par l'utilisateur).
# ============================================================================

import datetime as dt
import altair as alt
import pandas as pd
import streamlit as st
import core
from ui.style import entete, sous_titre, tuiles

# Série unique des graphiques : un bleu neutre (le rouge SB signale déjà les retards).
BLEU = "#2A78D6"
GRIS = "#5C5C5C"
MOIS = ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc."]


def _jours(valeur):
    """Nombre de jours en français : 9,5 ; 10 ; « — » si inconnu."""
    if valeur is None:
        return "—"
    texte = f"{valeur:.1f}"
    if texte.endswith(".0"):
        texte = texte[:-2]
    return texte.replace(".", ",")


def _douze_derniers_mois():
    """Clés « AAAA-MM » des 12 derniers mois, le mois en cours en dernier."""
    aujourd_hui = dt.date.today()
    annee, mois = aujourd_hui.year, aujourd_hui.month
    cles = []
    for _ in range(12):
        cles.append(f"{annee}-{mois:02d}")
        mois = mois - 1
        if mois == 0:
            annee, mois = annee - 1, 12
    cles.reverse()
    return cles


def _par_mois(delais):
    """Une ligne par mois des 12 derniers mois (mois du visa) : délai médian,
    taux de visa dans le délai, et le détail du mois en texte français pour le
    survol (« 9,5 j ouvrés », « 75 % (3 sur 4) »). Mois sans visa : valeurs vides."""
    lignes = []
    for cle in _douze_derniers_mois():
        elements = [d for d in delais if d["mois"] == cle]
        stats = core.statistiques_delais(elements)
        annee, mois = cle.split("-")
        if stats["nb"]:
            taux_texte = f"{round(stats['pct_dans_delai'])} % ({stats['nb_dans_delai']} sur {stats['nb']})"
        else:
            taux_texte = "—"
        lignes.append({"mois": f"{MOIS[int(mois) - 1]} {annee[2:]}",
                       "mediane": stats["mediane"], "taux": stats["pct_dans_delai"],
                       "Documents visés": stats["nb"],
                       "Délai médian": f"{_jours(stats['mediane'])} j ouvrés",
                       "Délai moyen": f"{_jours(stats['moyenne'])} j ouvrés",
                       "Visés dans le délai": taux_texte})
    return pd.DataFrame(lignes)


def _axe_mois(donnees):
    """Axe horizontal des 12 mois, dans l'ordre chronologique."""
    return alt.X("mois:N", sort=list(donnees["mois"]), title=None, scale=alt.Scale(paddingInner=0.45),
                 axis=alt.Axis(labelAngle=0, labelColor=GRIS, ticks=False, domainColor="#CFCFCF"))


INFOBULLE = [alt.Tooltip("mois:N", title="Mois du visa"), alt.Tooltip("Documents visés:Q"),
             alt.Tooltip("Délai médian:N"), alt.Tooltip("Délai moyen:N"), alt.Tooltip("Visés dans le délai:N")]


def _barres(donnees, champ, axe_y, echelle_y, texte_zero):
    """Barres mois par mois (une seule série, bleu neutre) ; le détail du mois
    apparaît au survol. Un mois mesuré à zéro (barre invisible) porte le texte
    texte_zero, pour ne pas être confondu avec un mois sans visa."""
    barres = alt.Chart(donnees).mark_bar(color=BLEU, cornerRadiusTopLeft=4, cornerRadiusTopRight=4).encode(
        x=_axe_mois(donnees), y=alt.Y(f"{champ}:Q", scale=echelle_y, axis=axe_y), tooltip=INFOBULLE)
    zeros = alt.Chart(donnees).transform_filter(f"datum.{champ} === 0").mark_text(
        dy=-8, color=GRIS, fontSize=12, text=texte_zero).encode(
        x=_axe_mois(donnees), y=alt.Y(f"{champ}:Q", scale=echelle_y), tooltip=INFOBULLE)
    return barres + zeros


def _afficher_graphique(graphique):
    st.altair_chart(graphique.properties(height=260).configure(font="Inter").configure_view(strokeWidth=0),
                    width="stretch")


def _evolution_delai(donnees):
    """Délai médian par mois, avec le délai par défaut en pointillé (légende
    sous le titre de la section : un texte dans le graphique chevauchait les barres)."""
    axe_y = alt.Axis(title="Délai médian (jours ouvrés)", labelColor=GRIS, titleColor=GRIS,
                     gridColor="#EEEEEE", domain=False, ticks=False)
    barres = _barres(donnees, "mediane", axe_y, alt.Scale(), "0 j")
    repere = pd.DataFrame({"delai": [core.DELAI_VISA_SB]})
    ligne = alt.Chart(repere).mark_rule(color=GRIS, strokeDash=[4, 4], strokeWidth=1.5).encode(y="delai:Q")
    _afficher_graphique(barres + ligne)


def _evolution_taux(donnees):
    """Taux de visa dans le délai par mois (axe de 0 à 100 %)."""
    axe_y = alt.Axis(title="Visés dans le délai (%)", labelColor=GRIS, titleColor=GRIS,
                     gridColor="#EEEEEE", domain=False, ticks=False, values=[0, 25, 50, 75, 100],
                     labelExpr="datum.value + ' %'")
    _afficher_graphique(_barres(donnees, "taux", axe_y, alt.Scale(domain=[0, 100]), "0 %"))


def afficher(ctx):
    entete("", "Délais de visa — Solares Bauen",
           [("Périmètre", "toute l'entreprise, tous projets (archivés compris)"),
            ("Mesure", "du dépôt du document à l'émission du bordereau, en jours ouvrés (Alsace-Moselle)")])
    delais = core.delais_de_visa()
    stats = core.statistiques_delais(delais)
    if not stats["nb"]:
        st.info("Aucun document visé pour l'instant : un délai est mesuré quand le bordereau "
                "qui vise un document déposé est émis.", icon=":material/info:")
        return
    tuiles([
        {"lab": "Documents visés", "val": stats["nb"], "unite": "doc.",
         "sub": "rappels « à viser avant le » clos par un bordereau"},
        {"lab": "Délai moyen", "val": _jours(stats["moyenne"]), "unite": "j ouvrés",
         "sub": "du dépôt au bordereau"},
        {"lab": "Délai médian", "val": _jours(stats["mediane"]), "unite": "j ouvrés",
         "sub": f"la moitié des documents visés en {_jours(stats['mediane'])} j ouvrés ou moins"},
        {"lab": "Taux dans le délai", "val": round(stats["pct_dans_delai"]), "unite": "%",
         "alerte": stats["nb_retard"] > 0,
         "sub": f"{stats['nb_retard']} en retard sur l'échéance du lot (10 j ouvrés par défaut)"},
    ])
    donnees = _par_mois(delais)
    sous_titre("Taux de visa dans le délai, mois par mois")
    st.caption("Part des documents visés au plus tard à l'échéance de leur lot. Détail de chaque mois au survol.")
    _evolution_taux(donnees)
    sous_titre("Délai médian, mois par mois")
    st.caption(f"Pointillé : délai par défaut de {core.DELAI_VISA_SB} jours ouvrés. Détail de chaque mois au survol.")
    _evolution_delai(donnees)
