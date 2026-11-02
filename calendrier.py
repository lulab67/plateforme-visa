# ============================================================================
#  calendrier.py — JOURS OUVRÉS (calendrier d'Alsace-Moselle)
# ----------------------------------------------------------------------------
#  Sert à calculer la date limite des rappels de visa côté Solares Bauen
#  ("à viser avant le ..."). Jours ouvrés = du lundi au vendredi, hors jours
#  fériés. L'Alsace-Moselle a deux jours fériés de plus que le reste de la
#  France : le Vendredi saint et le 26 décembre (Saint-Étienne).
# ============================================================================

import datetime as dt


def paques(annee):
    """Date du dimanche de Pâques (algorithme grégorien dit « anonyme »)."""
    a = annee % 19
    b = annee // 100
    c = annee % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    mois = (h + l - 7 * m + 114) // 31
    jour = ((h + l - 7 * m + 114) % 31) + 1
    return dt.date(annee, mois, jour)


def jours_feries(annee):
    """Ensemble des jours fériés d'Alsace-Moselle pour une année."""
    dimanche_paques = paques(annee)
    return {
        dt.date(annee, 1, 1),                              # Jour de l'an
        dimanche_paques - dt.timedelta(days=2),            # Vendredi saint (Alsace-Moselle)
        dimanche_paques + dt.timedelta(days=1),            # Lundi de Pâques
        dt.date(annee, 5, 1),                              # Fête du travail
        dt.date(annee, 5, 8),                              # Victoire 1945
        dimanche_paques + dt.timedelta(days=39),           # Ascension
        dimanche_paques + dt.timedelta(days=50),           # Lundi de Pentecôte
        dt.date(annee, 7, 14),                             # Fête nationale
        dt.date(annee, 8, 15),                             # Assomption
        dt.date(annee, 11, 1),                             # Toussaint
        dt.date(annee, 11, 11),                            # Armistice
        dt.date(annee, 12, 25),                            # Noël
        dt.date(annee, 12, 26),                            # Saint-Étienne (Alsace-Moselle)
    }


def est_ouvre(jour):
    """Vrai si le jour est un jour ouvré (lundi-vendredi, non férié)."""
    if jour.weekday() >= 5:  # 5 = samedi, 6 = dimanche
        return False
    return jour not in jours_feries(jour.year)


def ajouter_jours_ouvres(depart, nombre):
    """Date obtenue en avançant de 'nombre' jours ouvrés après 'depart'.
    Exemple : un dépôt le vendredi + 1 jour ouvré = le lundi suivant."""
    jour = depart
    restant = nombre
    while restant > 0:
        jour = jour + dt.timedelta(days=1)
        if est_ouvre(jour):
            restant = restant - 1
    return jour


def jours_ouvres_restants(aujourd_hui, echeance):
    """Nombre de jours ouvrés entre aujourd'hui (exclu) et l'échéance (incluse).
    Négatif si l'échéance est dépassée (nombre de jours ouvrés de retard)."""
    if echeance >= aujourd_hui:
        compte = 0
        jour = aujourd_hui
        while jour < echeance:
            jour = jour + dt.timedelta(days=1)
            if est_ouvre(jour):
                compte = compte + 1
        return compte
    compte = 0
    jour = echeance
    while jour < aujourd_hui:
        jour = jour + dt.timedelta(days=1)
        if est_ouvre(jour):
            compte = compte + 1
    return -compte
