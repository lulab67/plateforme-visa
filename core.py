# ============================================================================
#  core.py — LOGIQUE MÉTIER de la plateforme de suivi des visas
# ----------------------------------------------------------------------------
#  Ce fichier ne contient AUCUN élément d'interface : uniquement la base de
#  données SQLite, les règles métier (avis, relances, rappels de visa),
#  les imports/exports Excel/PDF et la gestion des fichiers déposés.
#  L'interface (app.py et ui/) appelle les fonctions définies ici.
#
#  Organisation du fichier (dans l'ordre) :
#    1. Constantes métier (avis, statuts, trames par défaut)
#    2. Connexion et initialisation de la base (avec migrations)
#    3. Chefs de projet et équipes (créer / modifier / supprimer / lister)
#    4. Projets           (créer / modifier / archiver / supprimer / lister)
#    5. Lots              (créer / modifier / supprimer / lister)
#   5b. Entreprises       (fiche, compte du portail, chargés d'affaires)
#    6. Trames de visa    (référentiel configurable)
#    7. Positions         (lecture / écriture / archives "anciens visas")
#    8. Association automatique position -> trame (avec apprentissage)
#    9. Relances titulaire (journal, fréquence, mail groupé .eml)
#   10. Rappels de visa Solares Bauen (à viser avant J + N jours ouvrés)
#   11. Import DPGF       (annexe marque/type du CCTP)
#   12. Fichiers PDF      (contenu stocké via stockage.py)
#   13. Exports           (bordereau Excel + PDF, statut des visas)
#   14. Registre des bordereaux émis (numérotation par lot)
#   15. Base de données   (vue transversale pour les statistiques)
#   16. Portail entreprises : envois, transfert, refus, état miroir
# ============================================================================

import sqlite3          # Base de données locale, fournie avec Python.
import json             # Les listes (attendus, contrôles...) sont stockées en JSON.
import io               # Manipulation de fichiers en mémoire (exports Excel).
import os               # Chemins (conversion PDF, anciens fichiers sur disque).
import re               # Nettoyage des noms de fichiers (caractères interdits).
import shutil           # Recherche du programme LibreOffice.
import subprocess       # Lancement de LibreOffice pour la conversion PDF.
import tempfile         # Dossier de travail temporaire pour la conversion PDF.
import unicodedata      # Suppression des accents pour comparer des textes.
import difflib          # Mesure de ressemblance entre deux textes (association).
import textwrap         # Découpage du texte en lignes (hauteur des lignes du bordereau).
import datetime as dt   # Dates (visa, relances, rappels).
import zipfile          # Contrôle du contenu des ZIP déposés sur le portail.
import bcrypt           # Hachage des mots de passe des comptes entreprises.
from pypdf import PdfReader, PdfWriter                           # Extraits de pages PDF.
from openpyxl import load_workbook, Workbook                      # Lecture/écriture Excel.
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.formatting import ConditionalFormattingList  # Couleurs des avis.
from openpyxl.worksheet.pagebreak import Break, RowBreak             # Sauts de page.
from openpyxl.worksheet.cell_range import CellRange
from openpyxl.formatting.rule import Rule                            # Couleurs des avis
from openpyxl.styles.differential import DifferentialStyle           # (mise en forme conditionnelle).

import calendrier       # Jours ouvrés d'Alsace-Moselle.
import stockage         # Où vit le contenu des fichiers déposés.

# ============================================================================
#  1. CONSTANTES MÉTIER
# ============================================================================

# Nom du fichier de base de données, créé à côté de l'application.
DB_PATH = "visa.db"

# Les six avis possibles sur un document, dans l'ordre d'affichage.
AVIS = ["VSO", "VAO", "VAOB", "REF", "MQ", "SO"]

# Les statuts possibles d'une pièce attendue. « À vérifier » est posé
# automatiquement quand un document est déposé : la pièce est arrivée mais
# n'a pas encore été contrôlée.
STATUT_A_VERIFIER = "À vérifier"
STATUTS = ["Reçu et ok", STATUT_A_VERIFIER, "MAJ attendue", "Manquant"]

# Statuts qui signifient « la pièce a été fournie » (date de réception, relances).
STATUTS_FOURNIS = ("Reçu et ok", STATUT_A_VERIFIER)

# Anciens libellés de statut -> nouveaux (conversion à la lecture).
STATUT_MIGRATION = {"Reçu": "Reçu et ok", "En attente": "MAJ attendue"}

# Étiquette courte écrite dans le bordereau exporté, devant chaque pièce.
TAG = {"Reçu et ok": "[reçu]", STATUT_A_VERIFIER: "[à vérifier]",
       "MAJ attendue": "[MAJ attendue]", "Manquant": "[manquant]"}

# Forme d'une référence de position CCTP/DPGF : une numérotation, éventuellement
# suivie d'une lettre. Exemples acceptés : "3", "3.1", "3.1.2", "03.01", "3.1.2a".
# Exemples refusés : "CHAPITRE 3", "Lot 05", "Article 3.1" (titres, pas des positions).
REF_POSITION = re.compile(r"^\d+(?:[.\-]\d+)*[a-zA-Z]?\.?$")

# Délai par défaut pour viser un document déposé (jours ouvrés).
DELAI_VISA_SB = 10

# Fréquence par défaut des relances au titulaire (jours calendaires).
FREQ_RELANCE = 14

# Trames indispensables au fonctionnement : toujours présentes, non renommables.
TRAME_A_DEFINIR = "À définir (aucun attendu)"
TRAME_AUTRE = "Autre document"
TRAMES_SYSTEME = (TRAME_A_DEFINIR, TRAME_AUTRE)

# ----------------------------------------------------------------------------
# Trames installées dans une base neuve :
#   - les deux trames système ci-dessous ;
#   - toutes celles du tableau templates/TRAMES_VISA.xlsx (CVC et sanitaire,
#     voir lire_trames_excel), une trame par type de position du CCTP.
# Chaque trame contient :
#   "docs" : les documents attendus pour viser (exportés au bordereau)
#   "ctrl" : points de contrôle (checklist interne, jamais exportée)
#   "kw"   : mots-clés servant à présélectionner la trame automatiquement
# Elles ne sont installées qu'une fois (base vide) : l'utilisateur peut
# ensuite les renommer, les modifier, les supprimer, ou toutes les remplacer
# par un autre tableau (page « Trames de visa »).
# ----------------------------------------------------------------------------
DEFAULT_TRAMES = {
    TRAME_A_DEFINIR: {"docs": [], "ctrl": [], "kw": []},
    TRAME_AUTRE: {"docs": ["Fiche technique"], "ctrl": [], "kw": []},
}
TRAMES_FICHIER = os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates", "TRAMES_VISA.xlsx")

# Correspondance trame -> type de document (information affichée).
TYPE_PAR_TRAME = {
    "Plan d'exécution": "Plan",
    "Note de calcul": "NdC",
    TRAME_AUTRE: "Document",
}


def type_of(nom_trame):
    """Renvoie le type de document ("Plan", "NdC", "Document" ou "Matériel")
    correspondant à une trame. Par défaut : "Matériel"."""
    if nom_trame in TYPE_PAR_TRAME:
        return TYPE_PAR_TRAME[nom_trame]
    return "Matériel"


def lignes_adresse(texte):
    """Découpe une adresse saisie sur plusieurs lignes en (ligne 1, ligne 2).
    La ligne 2 regroupe tout ce qui suit (ex. "67000 Strasbourg").
    Exemple : "5 rue des Lilas\\n67000 Strasbourg" -> ("5 rue des Lilas", "67000 Strasbourg")."""
    lignes = []
    for ligne in (texte or "").splitlines():
        if ligne.strip():
            lignes.append(ligne.strip())
    if len(lignes) == 0:
        return "", ""
    return lignes[0], " ".join(lignes[1:])


# ============================================================================
#  2. CONNEXION ET INITIALISATION DE LA BASE
# ============================================================================

def conn():
    """Ouvre une connexion à la base SQLite.
    row_factory = Row permet d'accéder aux colonnes par leur nom."""
    connexion = sqlite3.connect(DB_PATH)
    connexion.row_factory = sqlite3.Row
    return connexion


def init_db():
    """Crée les tables si elles n'existent pas, puis applique les migrations
    (ajout de colonnes sur les bases plus anciennes). À appeler au démarrage."""
    connexion = conn()
    curseur = connexion.cursor()

    # --- Création des tables principales (si absentes) ---
    curseur.executescript("""
    CREATE TABLE IF NOT EXISTS chefs(
      code TEXT PRIMARY KEY,          -- ex. "CM" (généré depuis les initiales)
      nom  TEXT);                     -- ex. "Christian Mahler"

    CREATE TABLE IF NOT EXISTS projets(
      code TEXT PRIMARY KEY,          -- ex. "MEINAU-01"
      nom TEXT, chef TEXT,            -- nom = désignation ; chef = code du chef de projet
      moa TEXT, operation TEXT, adresse TEXT);  -- adresse = adresse du MOA (2 lignes)

    CREATE TABLE IF NOT EXISTS lots(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      projet TEXT,                    -- code du projet parent
      lot TEXT, entreprise TEXT,      -- libellé complet du lot et titulaire
      contact_civ TEXT, contact_nom TEXT, contact_email TEXT, contact_adr TEXT,
      visa_no TEXT, indice TEXT,      -- n° de bordereau et indice de révision
      freq_jours INTEGER DEFAULT 14); -- fréquence de relance en jours

    CREATE TABLE IF NOT EXISTS positions(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      projet TEXT, lot_id INTEGER, ordre INTEGER,
      pos TEXT, des TEXT,             -- référence CCTP (peut être vide) + désignation
      ptype TEXT, famille TEXT,       -- type de document + nom de la trame
      marque TEXT, typeprop TEXT,     -- proposition de l'ENTREPRISE
      avis TEXT, etat TEXT, pct INTEGER,  -- etat / pct : anciennes colonnes, plus utilisées
      relance INTEGER,                -- 1 = inclure dans la prochaine relance
      datevisa TEXT, intitule TEXT, rem TEXT,
      attendus TEXT,                  -- liste JSON des pièces attendues
      export INTEGER DEFAULT 1);      -- 1 = inclure dans le bordereau exporté

    CREATE TABLE IF NOT EXISTS relances(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      projet TEXT, lot_id INTEGER, date TEXT,
      niveau INTEGER, nb_positions INTEGER,
      positions TEXT);                -- liste JSON des références relancées

    CREATE TABLE IF NOT EXISTS fichiers(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      position_id INTEGER,            -- position à laquelle le PDF est rattaché
      attendu TEXT,                   -- liste JSON des attendus couverts (ou NULL)
      nom TEXT, chemin TEXT,          -- nom du fichier (+ chemin disque des anciens dépôts)
      version INTEGER, courant INTEGER DEFAULT 1,
      date_upload TEXT, taille INTEGER);

    CREATE TABLE IF NOT EXISTS rappels_sb(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      projet TEXT, lot_id INTEGER, position_id INTEGER,
      date_depot TEXT,                -- date du dépôt qui a ouvert le rappel
      echeance TEXT,                  -- "à viser avant le"
      fichiers TEXT,                  -- liste JSON des fichiers déposés
      cloture INTEGER DEFAULT 0,      -- 1 = visé (clôture manuelle)
      date_cloture TEXT);

    CREATE TABLE IF NOT EXISTS teams(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      nom TEXT UNIQUE);               -- ex. "Team Fluides Est"

    CREATE TABLE IF NOT EXISTS bordereaux(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      projet TEXT, lot_id INTEGER,
      numero INTEGER,                 -- n° du bordereau dans le lot (VISA N°01, 02...)
      date TEXT,                      -- date d'émission
      contenu TEXT);                  -- photo JSON : projet, lot, signataire, positions

    -- ---- Portail entreprises (itération 1c) ----
    CREATE TABLE IF NOT EXISTS entreprises(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      nom TEXT UNIQUE,                -- ex. "Thermi Est" (recopié dans lots.entreprise)
      adresse TEXT,                   -- 2 lignes (recopiée dans lots.contact_adr)
      identifiant TEXT,               -- identifiant du compte portail (un par entreprise)
      mdp_hache TEXT,                 -- mot de passe haché (bcrypt) ; jamais en clair
      actif INTEGER DEFAULT 1);       -- 0 = compte portail désactivé

    CREATE TABLE IF NOT EXISTS contacts_entreprise(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      entreprise_id INTEGER,
      nom TEXT, email TEXT,           -- chargé d'affaires (renseigné par l'entreprise)
      actif INTEGER DEFAULT 1);       -- 0 = retiré (gardé pour l'historique des envois)

    CREATE TABLE IF NOT EXISTS depots(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      entreprise_id INTEGER, projet TEXT, lot_id INTEGER,
      contact_id INTEGER,             -- chargé d'affaires choisi à l'envoi
      contact_nom TEXT, contact_email TEXT,  -- recopiés : restent justes si le contact change
      date_envoi TEXT);               -- point de départ du rappel de visa SB

    CREATE TABLE IF NOT EXISTS connexions_echouees(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      identifiant TEXT,               -- identifiant saisi sur le portail (minuscules)
      date TEXT);                     -- date et heure de l'échec (ISO)

    CREATE TABLE IF NOT EXISTS depot_lignes(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      depot_id INTEGER, position_id INTEGER,
      pieces TEXT,                    -- liste JSON des pièces attendues déclarées
      commentaire TEXT,               -- commentaire de l'entreprise
      statut TEXT DEFAULT 'transmis', -- transmis | transfere | refuse
      date_traitement TEXT,           -- date du transfert ou du refus
      traite_par TEXT,                -- chef de projet qui a traité la ligne
      motif_refus TEXT,
      libre_type TEXT DEFAULT '',     -- soumission libre : cctp | plan ('' = ligne d'une position)
      libre_pos TEXT DEFAULT '',      -- soumission libre : référence saisie par l'entreprise
      libre_des TEXT DEFAULT '');     -- soumission libre : désignation saisie par l'entreprise

    CREATE TABLE IF NOT EXISTS depot_fichiers(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      depot_ligne_id INTEGER,
      nom TEXT, taille INTEGER,
      contenu_id INTEGER,             -- référence stockage.py (un fichier commun = un seul contenu)
      extrait TEXT,                   -- pages d'un PDF ("12-30"), fichier dans un ZIP, ou vide
      commun INTEGER DEFAULT 0);      -- 1 = vient du dossier commun de l'envoi
    """)
    stockage.creer_tables(connexion)

    # --- Migrations : colonnes ajoutées au fil des versions ---
    # On tente chaque ALTER ; s'il échoue (colonne déjà présente), on ignore.
    migrations = [
        "ALTER TABLE positions ADD COLUMN archive INTEGER DEFAULT 0",   # anciens visas
        "ALTER TABLE positions ADD COLUMN rem_histo TEXT",              # commentaire interne SB
        "ALTER TABLE positions ADD COLUMN controles TEXT",              # points de contrôle
        "ALTER TABLE positions ADD COLUMN marque_moe TEXT",             # proposition MOE (import)
        "ALTER TABLE positions ADD COLUMN type_moe TEXT",               # proposition MOE (import)
        "ALTER TABLE positions ADD COLUMN mt_verrou INTEGER DEFAULT 0", # 1 = MOE grisé/verrouillé
        "ALTER TABLE lots ADD COLUMN chemin_fichiers TEXT",             # ancien stockage disque
        "ALTER TABLE lots ADD COLUMN chemin_opc TEXT",                  # ancien export OPC
        "ALTER TABLE projets ADD COLUMN archive INTEGER DEFAULT 0",     # projets archivés
        "ALTER TABLE chefs ADD COLUMN email TEXT",                      # interlocuteur bordereau
        "ALTER TABLE chefs ADD COLUMN telephone TEXT",
        "ALTER TABLE lots ADD COLUMN lot_numero TEXT",                  # ex. "05"
        "ALTER TABLE lots ADD COLUMN lot_denomination TEXT",            # ex. "CHAUFFAGE VENTILATION"
        "ALTER TABLE lots ADD COLUMN contact_tel TEXT",
        "ALTER TABLE lots ADD COLUMN delai_visa_sb INTEGER DEFAULT 10", # jours ouvrés pour viser
        "ALTER TABLE fichiers ADD COLUMN contenu_id INTEGER",           # référence stockage.py
        "ALTER TABLE chefs ADD COLUMN team_id INTEGER",                 # équipe du chef de projet
        "ALTER TABLE lots ADD COLUMN entreprise_id INTEGER",            # fiche entreprise du titulaire
        "ALTER TABLE lots ADD COLUMN contact_id INTEGER",               # chargé d'affaires du lot
        "ALTER TABLE contacts_entreprise ADD COLUMN civ TEXT",          # M. / Mme
        "ALTER TABLE contacts_entreprise ADD COLUMN telephone TEXT",
        "ALTER TABLE positions ADD COLUMN visa_bordereau_id INTEGER",   # ancien visa : son bordereau
        "ALTER TABLE positions ADD COLUMN origine_id INTEGER",          # ancien visa : position d'origine
        "ALTER TABLE depot_lignes ADD COLUMN libre_type TEXT DEFAULT ''",  # soumission libre (portail)
        "ALTER TABLE depot_lignes ADD COLUMN libre_pos TEXT DEFAULT ''",
        "ALTER TABLE depot_lignes ADD COLUMN libre_des TEXT DEFAULT ''",
    ]
    for requete_alter in migrations:
        try:
            curseur.execute(requete_alter)
        except sqlite3.OperationalError:
            pass  # La colonne existe déjà : rien à faire.
    _rattacher_lots_sans_fiche(connexion)
    _rattacher_contacts_lots(connexion)
    # Chargés d'affaires sans civilité (créés avant qu'elle existe) : celle
    # recopiée dans un de leurs lots, sinon « M. ».
    connexion.execute(
        """UPDATE contacts_entreprise SET civ = COALESCE(
               (SELECT l.contact_civ FROM lots l WHERE l.contact_id = contacts_entreprise.id
                AND IFNULL(l.contact_civ, '') <> '' LIMIT 1), 'M.')
           WHERE IFNULL(civ, '') = ''""")

    # --- Table des trames + table d'apprentissage des associations ---
    curseur.execute("""CREATE TABLE IF NOT EXISTS trames(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      nom TEXT UNIQUE, docs TEXT, controles TEXT)""")
    try:
        curseur.execute("ALTER TABLE trames ADD COLUMN keywords TEXT")
    except sqlite3.OperationalError:
        pass
    curseur.execute("""CREATE TABLE IF NOT EXISTS assoc(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      des TEXT, trame TEXT, date TEXT)""")  # mémoire des validations de trame

    # --- Trames par défaut : seulement dans une base neuve, pour qu'une trame
    #     renommée ou supprimée par l'utilisateur ne réapparaisse pas. ---
    nb_trames = curseur.execute("SELECT COUNT(*) FROM trames").fetchone()[0]
    for nom_trame in DEFAULT_TRAMES:
        contenu = DEFAULT_TRAMES[nom_trame]
        _inserer_trame(curseur, nom_trame, contenu["docs"], contenu["ctrl"], contenu["kw"])
    if nb_trames == 0:
        for trame in trames_par_defaut():
            if trame["nom"] not in TRAMES_SYSTEME:
                _inserer_trame(curseur, trame["nom"], trame["docs"], trame["ctrl"], trame["kw"])
    # Trame « Autre document » restée telle qu'installée : son document attendu
    # « Document à fournir » devient « Fiche technique » (retour du 30/10/2026).
    curseur.execute("UPDATE trames SET docs=? WHERE nom=? AND docs=?",
                    (json.dumps(["Fiche technique"], ensure_ascii=False), TRAME_AUTRE,
                     json.dumps(["Document à fournir"], ensure_ascii=False)))

    connexion.commit()
    connexion.close()


# ============================================================================
#  3. CHEFS DE PROJET
# ============================================================================

def get_chefs():
    """Renvoie la liste des chefs de projet, triée par nom.
    Chaque élément est un dictionnaire {code, nom, email, telephone}."""
    connexion = conn()
    lignes = connexion.execute("SELECT * FROM chefs ORDER BY nom").fetchall()
    connexion.close()
    resultat = []
    for ligne in lignes:
        resultat.append(dict(ligne))
    return resultat


def get_chef(code):
    """Renvoie un chef de projet (dictionnaire) à partir de son code, ou None."""
    connexion = conn()
    ligne = connexion.execute("SELECT * FROM chefs WHERE code=?", (code,)).fetchone()
    connexion.close()
    if ligne is None:
        return None
    return dict(ligne)


def add_chef(nom_complet, email="", telephone=""):
    """Ajoute un chef de projet. Le code est généré à partir des initiales
    (ex. "Marie Dupont" -> "MD") ; un chiffre est ajouté en cas de doublon.
    Renvoie le code créé."""
    # Construction des initiales : première lettre de chaque mot du nom.
    initiales = ""
    for mot in nom_complet.split():
        if mot:
            initiales = initiales + mot[0].upper()
    initiales = initiales[:3]
    if initiales == "":
        initiales = "CP"
    # Recherche d'un code libre (MD, puis MD2, MD3...).
    connexion = conn()
    code = initiales
    numero = 1
    while connexion.execute("SELECT 1 FROM chefs WHERE code=?", (code,)).fetchone():
        numero = numero + 1
        code = f"{initiales}{numero}"
    connexion.execute("INSERT INTO chefs(code, nom, email, telephone) VALUES(?,?,?,?)",
                      (code, nom_complet, email, telephone))
    connexion.commit()
    connexion.close()
    return code


def update_chef(code, nom, email="", telephone=""):
    """Modifie le nom, l'email et le téléphone d'un chef (le code ne change pas)."""
    connexion = conn()
    connexion.execute("UPDATE chefs SET nom=?, email=?, telephone=? WHERE code=?",
                      (nom, email, telephone, code))
    connexion.commit()
    connexion.close()


def delete_chef(code):
    """Supprime un chef de projet, UNIQUEMENT s'il n'a plus de projet rattaché.
    Renvoie (True, "") en cas de succès, sinon (False, message d'explication)."""
    connexion = conn()
    nb_projets = connexion.execute(
        "SELECT COUNT(*) FROM projets WHERE chef=?", (code,)).fetchone()[0]
    if nb_projets > 0:
        connexion.close()
        return False, f"{nb_projets} projet(s) rattaché(s) — réaffectez ou supprimez-les d'abord."
    connexion.execute("DELETE FROM chefs WHERE code=?", (code,))
    connexion.commit()
    connexion.close()
    return True, ""


def set_chef_team(code, team_id):
    """Rattache un chef de projet à une équipe (team_id=None : aucune équipe)."""
    connexion = conn()
    connexion.execute("UPDATE chefs SET team_id=? WHERE code=?", (team_id, code))
    connexion.commit()
    connexion.close()


def get_teams():
    """Liste des équipes, triées par nom. Chaque équipe : {id, nom, chefs}
    où chefs est la liste des codes des chefs de projet rattachés."""
    connexion = conn()
    equipes = []
    for ligne in connexion.execute("SELECT * FROM teams ORDER BY nom").fetchall():
        equipe = dict(ligne)
        equipe["chefs"] = []
        for chef in connexion.execute(
                "SELECT code FROM chefs WHERE team_id=? ORDER BY nom", (equipe["id"],)).fetchall():
            equipe["chefs"].append(chef["code"])
        equipes.append(equipe)
    connexion.close()
    return equipes


def add_team(nom):
    """Crée une équipe. Renvoie (True, id) ou (False, message)."""
    nom = (nom or "").strip()
    if nom == "":
        return False, "Le nom de l'équipe est obligatoire."
    connexion = conn()
    if connexion.execute("SELECT 1 FROM teams WHERE nom=?", (nom,)).fetchone():
        connexion.close()
        return False, f"L'équipe « {nom} » existe déjà."
    curseur = connexion.execute("INSERT INTO teams(nom) VALUES(?)", (nom,))
    nouvel_id = curseur.lastrowid
    connexion.commit()
    connexion.close()
    return True, nouvel_id


def rename_team(team_id, nom):
    """Renomme une équipe. Renvoie (True, "") ou (False, message)."""
    nom = (nom or "").strip()
    if nom == "":
        return False, "Le nom de l'équipe est obligatoire."
    connexion = conn()
    if connexion.execute("SELECT 1 FROM teams WHERE nom=? AND id<>?", (nom, team_id)).fetchone():
        connexion.close()
        return False, f"L'équipe « {nom} » existe déjà."
    connexion.execute("UPDATE teams SET nom=? WHERE id=?", (nom, team_id))
    connexion.commit()
    connexion.close()
    return True, ""


def delete_team(team_id):
    """Supprime une équipe ; ses chefs de projet restent, sans équipe."""
    connexion = conn()
    connexion.execute("UPDATE chefs SET team_id=NULL WHERE team_id=?", (team_id,))
    connexion.execute("DELETE FROM teams WHERE id=?", (team_id,))
    connexion.commit()
    connexion.close()


def get_projets_team(team_id, archive=0):
    """Projets de tous les chefs de projet d'une équipe."""
    connexion = conn()
    lignes = connexion.execute(
        """SELECT p.* FROM projets p JOIN chefs c ON c.code = p.chef
           WHERE c.team_id=? AND IFNULL(p.archive,0)=? ORDER BY p.code""",
        (team_id, archive)).fetchall()
    connexion.close()
    resultat = []
    for ligne in lignes:
        resultat.append(dict(ligne))
    return resultat


def interlocuteur(projet):
    """Interlocuteur Solares Bauen écrit dans le bordereau et signataire des
    relances. En attendant la connexion Microsoft 365, c'est le chef de projet
    du projet. Renvoie {nom, email, telephone} (valeurs vides si inconnu)."""
    chef = get_chef(projet.get("chef")) if projet else None
    if chef is None:
        return {"nom": "", "email": "", "telephone": ""}
    return {"nom": chef.get("nom") or "", "email": chef.get("email") or "",
            "telephone": chef.get("telephone") or ""}


# ============================================================================
#  4. PROJETS
# ============================================================================

def get_projets(chef=None, archive=0):
    """Liste les projets. Par défaut : uniquement les projets ACTIFS (archive=0).
    - chef : si fourni, filtre sur ce chef de projet.
    - archive=1 : renvoie au contraire les projets archivés."""
    connexion = conn()
    if chef:
        lignes = connexion.execute(
            "SELECT * FROM projets WHERE IFNULL(archive,0)=? AND chef=? ORDER BY code",
            (archive, chef)).fetchall()
    else:
        lignes = connexion.execute(
            "SELECT * FROM projets WHERE IFNULL(archive,0)=? ORDER BY code",
            (archive,)).fetchall()
    connexion.close()
    resultat = []
    for ligne in lignes:
        resultat.append(dict(ligne))
    return resultat


def get_projet(code):
    """Renvoie un projet (dictionnaire) à partir de son code, ou None."""
    connexion = conn()
    ligne = connexion.execute("SELECT * FROM projets WHERE code=?", (code,)).fetchone()
    connexion.close()
    if ligne is None:
        return None
    return dict(ligne)


def _verifier_projet(code, nom, moa, adresse):
    """Contrôle les champs obligatoires d'un projet. Renvoie "" si tout va bien,
    sinon le message à afficher."""
    manquants = []
    if not (code or "").strip():
        manquants.append("code projet")
    if not (moa or "").strip():
        manquants.append("maître d'ouvrage")
    if not (nom or "").strip():
        manquants.append("désignation du projet")
    if not (adresse or "").strip():
        manquants.append("adresse du maître d'ouvrage")
    if manquants:
        return "Champs obligatoires manquants : " + ", ".join(manquants) + "."
    return ""


def add_projet(code, nom, chef, moa, adresse):
    """Crée un projet rattaché à un chef de projet.
    - nom     : désignation du projet (écrite dans le bordereau) ;
    - moa     : nom du maître d'ouvrage ;
    - adresse : adresse du maître d'ouvrage (2 lignes).
    Refuse un code déjà utilisé (aucun projet n'est jamais écrasé).
    Renvoie (True, "") ou (False, message)."""
    code = (code or "").strip()
    erreur = _verifier_projet(code, nom, moa, adresse)
    if erreur:
        return False, erreur
    connexion = conn()
    existant = connexion.execute("SELECT nom FROM projets WHERE code=?", (code,)).fetchone()
    if existant:
        connexion.close()
        return False, f"Le code {code} est déjà utilisé par le projet « {existant['nom']} »."
    connexion.execute(
        """INSERT INTO projets(code, nom, chef, moa, operation, adresse, archive)
           VALUES(?,?,?,?,?,?,0)""",
        (code, nom.strip(), chef, moa.strip(), nom.strip(), adresse.strip()))
    connexion.commit()
    connexion.close()
    return True, ""


def update_projet(code, nom, chef, moa, adresse):
    """Modifie un projet existant (le code ne change pas).
    Renvoie (True, "") ou (False, message)."""
    erreur = _verifier_projet(code, nom, moa, adresse)
    if erreur:
        return False, erreur
    connexion = conn()
    connexion.execute(
        "UPDATE projets SET nom=?, chef=?, moa=?, operation=?, adresse=? WHERE code=?",
        (nom.strip(), chef, moa.strip(), nom.strip(), adresse.strip(), code))
    connexion.commit()
    connexion.close()
    return True, ""


def set_projet_archive(code, valeur):
    """Archive (valeur=1) ou restaure (valeur=0) un projet.
    Rien n'est supprimé : le projet sort simplement des listes de sélection."""
    connexion = conn()
    connexion.execute("UPDATE projets SET archive=? WHERE code=?", (int(valeur), code))
    connexion.commit()
    connexion.close()


def delete_projet(code):
    """Supprime DÉFINITIVEMENT un projet et tout ce qu'il contient :
    lots, positions, références de fichiers, relances et rappels."""
    connexion = conn()
    # 1. On récupère les identifiants de toutes les positions du projet.
    identifiants = []
    for ligne in connexion.execute("SELECT id FROM positions WHERE projet=?", (code,)).fetchall():
        identifiants.append(ligne[0])
    # 2. On supprime les références de fichiers liées à ces positions.
    if identifiants:
        marqueurs = ",".join("?" * len(identifiants))
        connexion.execute(f"DELETE FROM fichiers WHERE position_id IN ({marqueurs})", identifiants)
    # 3. Puis les positions, relances, rappels, lots, et le projet lui-même.
    connexion.execute("DELETE FROM positions WHERE projet=?", (code,))
    connexion.execute("DELETE FROM relances WHERE projet=?", (code,))
    connexion.execute("DELETE FROM rappels_sb WHERE projet=?", (code,))
    connexion.execute("DELETE FROM bordereaux WHERE projet=?", (code,))
    _supprimer_depots(connexion, "projet=?", (code,))
    connexion.execute("DELETE FROM lots WHERE projet=?", (code,))
    connexion.execute("DELETE FROM projets WHERE code=?", (code,))
    connexion.commit()
    connexion.close()


# ============================================================================
#  5. LOTS
# ============================================================================

def libelle_lot(lot):
    """Libellé complet d'un lot, tel qu'écrit dans le bordereau.
    Exemple : numéro "05" + dénomination "CHAUFFAGE VENTILATION"
    -> "Lot 05 - CHAUFFAGE VENTILATION". Les anciens lots (sans numéro ni
    dénomination séparés) gardent leur libellé d'origine."""
    numero = (lot.get("lot_numero") or "").strip()
    denomination = (lot.get("lot_denomination") or "").strip()
    if numero and denomination:
        return f"Lot {numero} - {denomination}"
    if numero:
        return f"Lot {numero}"
    if denomination:
        return denomination
    return lot.get("lot") or ""


def get_lots(projet):
    """Liste les lots d'un projet (dictionnaires), triés par numéro puis libellé."""
    connexion = conn()
    lignes = connexion.execute(
        "SELECT * FROM lots WHERE projet=? ORDER BY lot_numero, lot", (projet,)).fetchall()
    connexion.close()
    resultat = []
    for ligne in lignes:
        resultat.append(dict(ligne))
    return resultat


def get_lot(lot_id):
    """Renvoie un lot (dictionnaire) à partir de son identifiant, ou None."""
    connexion = conn()
    ligne = connexion.execute("SELECT * FROM lots WHERE id=?", (lot_id,)).fetchone()
    connexion.close()
    if ligne is None:
        return None
    return dict(ligne)


def add_lot(projet, numero, denomination, entreprise="", adresse="", civ="M.", nom="",
            email="", telephone=""):
    """Crée un lot pour un projet et renvoie son identifiant.
    - numero / denomination : ex. "05" / "CHAUFFAGE VENTILATION" ;
    - entreprise, adresse   : destinataire du bordereau (adresse sur 2 lignes) ;
    - civ, nom, email, telephone : interlocuteur côté titulaire."""
    libelle = libelle_lot({"lot_numero": numero, "lot_denomination": denomination})
    connexion = conn()
    curseur = connexion.cursor()
    curseur.execute(
        """INSERT INTO lots(projet, lot, lot_numero, lot_denomination, entreprise,
                            contact_civ, contact_nom, contact_email, contact_tel, contact_adr,
                            visa_no, indice, freq_jours, delai_visa_sb)
           VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        (projet, libelle, numero, denomination, entreprise, civ, nom, email, telephone,
         adresse, "VISA N°01", "rev0", FREQ_RELANCE, DELAI_VISA_SB))
    nouvel_id = curseur.lastrowid
    # Le titulaire saisi est rattaché à sa fiche entreprise (créée si besoin).
    entreprise_id = _trouver_ou_creer_entreprise(connexion, entreprise, adresse)
    connexion.execute("UPDATE lots SET entreprise_id=? WHERE id=?", (entreprise_id, nouvel_id))
    connexion.commit()
    connexion.close()
    return nouvel_id


def save_lot(lot_dict):
    """Enregistre toutes les informations modifiables d'un lot."""
    libelle = libelle_lot(lot_dict)
    connexion = conn()
    connexion.execute(
        """UPDATE lots SET lot=?, lot_numero=?, lot_denomination=?, entreprise=?,
                 contact_civ=?, contact_nom=?, contact_email=?, contact_tel=?, contact_adr=?,
                 visa_no=?, indice=?, freq_jours=?, delai_visa_sb=?
           WHERE id=?""",
        (libelle,
         lot_dict.get("lot_numero") or "",
         lot_dict.get("lot_denomination") or "",
         lot_dict.get("entreprise") or "",
         lot_dict.get("contact_civ") or "M.",
         lot_dict.get("contact_nom") or "",
         lot_dict.get("contact_email") or "",
         lot_dict.get("contact_tel") or "",
         lot_dict.get("contact_adr") or "",
         lot_dict.get("visa_no") or "",
         lot_dict.get("indice") or "",
         int(lot_dict.get("freq_jours") or FREQ_RELANCE),
         int(lot_dict.get("delai_visa_sb") or DELAI_VISA_SB),
         lot_dict["id"]))
    # Le titulaire saisi (texte) est rattaché à sa fiche entreprise.
    # (En 1c.6, l'Administration choisira directement la fiche dans une liste.)
    entreprise_id = _trouver_ou_creer_entreprise(
        connexion, lot_dict.get("entreprise"), lot_dict.get("contact_adr"))
    connexion.execute("UPDATE lots SET entreprise_id=? WHERE id=?", (entreprise_id, lot_dict["id"]))
    connexion.commit()
    connexion.close()


def delete_lot(lot_id):
    """Supprime DÉFINITIVEMENT un lot : positions, références de fichiers,
    relances et rappels."""
    connexion = conn()
    identifiants = []
    for ligne in connexion.execute("SELECT id FROM positions WHERE lot_id=?", (lot_id,)).fetchall():
        identifiants.append(ligne[0])
    if identifiants:
        marqueurs = ",".join("?" * len(identifiants))
        connexion.execute(f"DELETE FROM fichiers WHERE position_id IN ({marqueurs})", identifiants)
    connexion.execute("DELETE FROM positions WHERE lot_id=?", (lot_id,))
    connexion.execute("DELETE FROM relances WHERE lot_id=?", (lot_id,))
    connexion.execute("DELETE FROM rappels_sb WHERE lot_id=?", (lot_id,))
    connexion.execute("DELETE FROM bordereaux WHERE lot_id=?", (lot_id,))
    _supprimer_depots(connexion, "lot_id=?", (lot_id,))
    connexion.execute("DELETE FROM lots WHERE id=?", (lot_id,))
    connexion.commit()
    connexion.close()


def _supprimer_depots(connexion, condition, valeurs):
    """Supprime les dépôts du portail qui répondent à la condition (ex. "lot_id=?"),
    avec leurs lignes et leurs fichiers. Utilisé à la suppression d'un lot ou d'un projet."""
    for depot in connexion.execute(f"SELECT id FROM depots WHERE {condition}", valeurs).fetchall():
        for ligne in connexion.execute(
                "SELECT id FROM depot_lignes WHERE depot_id=?", (depot["id"],)).fetchall():
            connexion.execute("DELETE FROM depot_fichiers WHERE depot_ligne_id=?", (ligne["id"],))
        connexion.execute("DELETE FROM depot_lignes WHERE depot_id=?", (depot["id"],))
        connexion.execute("DELETE FROM depots WHERE id=?", (depot["id"],))


# ============================================================================
#  5b. ENTREPRISES : fiche, compte du portail, chargés d'affaires
# ----------------------------------------------------------------------------
#  Une fiche par entreprise titulaire (nom + adresse). Chaque lot pointe vers
#  sa fiche (lots.entreprise_id) ; le nom et l'adresse restent recopiés dans le
#  lot (lots.entreprise, lots.contact_adr) car le bordereau et les pages les
#  lisent là. Une entreprise a UN compte sur le portail (identifiant + mot de
#  passe haché) et voit tous les lots qui pointent vers sa fiche.
#  Les chargés d'affaires (nom + mail) forment UNE liste par entreprise,
#  modifiable par l'entreprise (Plateforme visa partenaires) et par Solares Bauen (Administration).
#  Chaque lot désigne SON chargé d'affaires (lots.contact_id) : deux lots d'une
#  même entreprise peuvent en avoir deux différents. Ses civilité, nom, mail et
#  téléphone sont recopiés dans le lot : c'est l'« interlocuteur titulaire » du
#  bordereau et le destinataire des relances et des mails de refus.
# ============================================================================

LONGUEUR_MINI_MOT_DE_PASSE = 8

# Forme simple d'une adresse mail : quelque chose@domaine.extension
# « \Z » et non « $ » : « $ » accepterait un retour à la ligne final, qui
# permettrait d'ajouter des en-têtes au brouillon de mail (.eml).
FORME_EMAIL = re.compile(r"[^@\s]+@[^@\s]+\.[^@\s]+\Z")


def _cle_entreprise(nom):
    """Clé de comparaison d'un nom d'entreprise : sans accents, minuscules,
    espaces réduits. "Clim'Est " et "CLIM'EST" donnent la même clé."""
    return " ".join(_norm(nom).split())


def _trouver_entreprise(connexion, nom, sauf_id=None):
    """Identifiant de la fiche portant ce nom (à la casse et aux accents près), ou None."""
    cle = _cle_entreprise(nom)
    for ligne in connexion.execute("SELECT id, nom FROM entreprises").fetchall():
        if _cle_entreprise(ligne["nom"]) == cle and ligne["id"] != sauf_id:
            return ligne["id"]
    return None


def _trouver_ou_creer_entreprise(connexion, nom, adresse=""):
    """Identifiant de la fiche de cette entreprise ; la crée si elle n'existe pas.
    Renvoie None si le nom est vide (lot sans titulaire)."""
    nom = (nom or "").strip()
    if nom == "":
        return None
    existante = _trouver_entreprise(connexion, nom)
    if existante is not None:
        return existante
    curseur = connexion.execute("INSERT INTO entreprises(nom, adresse, actif) VALUES(?,?,1)",
                                (nom, (adresse or "").strip()))
    return curseur.lastrowid


def _rattacher_lots_sans_fiche(connexion):
    """Migration : chaque lot dont le titulaire n'est encore qu'un texte libre est
    rattaché à une fiche entreprise (créée avec l'adresse du premier lot trouvé)."""
    lots_sans_fiche = connexion.execute(
        "SELECT id, entreprise, contact_adr FROM lots WHERE entreprise_id IS NULL ORDER BY id").fetchall()
    for lot in lots_sans_fiche:
        entreprise_id = _trouver_ou_creer_entreprise(connexion, lot["entreprise"], lot["contact_adr"])
        if entreprise_id is not None:
            connexion.execute("UPDATE lots SET entreprise_id=? WHERE id=?", (entreprise_id, lot["id"]))


def _rattacher_contacts_lots(connexion):
    """Migration : l'interlocuteur titulaire saisi en texte dans un lot devient le
    chargé d'affaires du lot. On reprend le contact de l'entreprise qui a le même
    mail, ou on le crée dans la liste de l'entreprise."""
    lots = connexion.execute(
        """SELECT id, entreprise_id, contact_civ, contact_nom, contact_email, contact_tel FROM lots
           WHERE contact_id IS NULL AND entreprise_id IS NOT NULL""").fetchall()
    for lot in lots:
        email = (lot["contact_email"] or "").strip()
        nom = (lot["contact_nom"] or "").strip()
        if nom == "" or not FORME_EMAIL.match(email):
            continue
        existant = connexion.execute(
            """SELECT id FROM contacts_entreprise
               WHERE entreprise_id=? AND lower(email)=lower(?) AND actif=1""",
            (lot["entreprise_id"], email)).fetchone()
        if existant is not None:
            contact_id = existant["id"]
        else:
            curseur = connexion.execute(
                """INSERT INTO contacts_entreprise(entreprise_id, civ, nom, email, telephone, actif)
                   VALUES(?,?,?,?,?,1)""",
                (lot["entreprise_id"], lot["contact_civ"] or "M.", nom, email, lot["contact_tel"] or ""))
            contact_id = curseur.lastrowid
        connexion.execute("UPDATE lots SET contact_id=? WHERE id=?", (contact_id, lot["id"]))


def get_entreprises():
    """Liste des fiches entreprises, triées par nom. Chaque élément contient en plus :
    nb_lots (lots rattachés) et a_compte (True si un compte portail existe)."""
    connexion = conn()
    resultat = []
    for ligne in connexion.execute("SELECT * FROM entreprises ORDER BY nom").fetchall():
        entreprise = dict(ligne)
        entreprise["nb_lots"] = connexion.execute(
            "SELECT COUNT(*) FROM lots WHERE entreprise_id=?", (entreprise["id"],)).fetchone()[0]
        entreprise["a_compte"] = bool(entreprise["identifiant"] and entreprise["mdp_hache"])
        del entreprise["mdp_hache"]   # le hachage ne sort jamais de core
        resultat.append(entreprise)
    connexion.close()
    return resultat


def get_entreprise(entreprise_id):
    """Renvoie une fiche entreprise (sans le mot de passe haché), ou None."""
    connexion = conn()
    ligne = connexion.execute("SELECT * FROM entreprises WHERE id=?", (entreprise_id,)).fetchone()
    connexion.close()
    if ligne is None:
        return None
    entreprise = dict(ligne)
    entreprise["a_compte"] = bool(entreprise["identifiant"] and entreprise["mdp_hache"])
    del entreprise["mdp_hache"]
    return entreprise


def add_entreprise(nom, adresse=""):
    """Crée une fiche entreprise. Refuse un nom déjà pris (à la casse et aux
    accents près). Renvoie (True, id) ou (False, message)."""
    nom = (nom or "").strip()
    if nom == "":
        return False, "Le nom de l'entreprise est obligatoire."
    connexion = conn()
    if _trouver_entreprise(connexion, nom) is not None:
        connexion.close()
        return False, f"L'entreprise « {nom} » existe déjà."
    curseur = connexion.execute("INSERT INTO entreprises(nom, adresse, actif) VALUES(?,?,1)",
                                (nom, (adresse or "").strip()))
    nouvel_id = curseur.lastrowid
    connexion.commit()
    connexion.close()
    return True, nouvel_id


def update_entreprise(entreprise_id, nom, adresse=""):
    """Modifie le nom et l'adresse d'une fiche, et les recopie dans tous ses lots
    (le bordereau suivant portera donc la nouvelle adresse).
    Renvoie (True, "") ou (False, message)."""
    nom = (nom or "").strip()
    adresse = (adresse or "").strip()
    if nom == "":
        return False, "Le nom de l'entreprise est obligatoire."
    connexion = conn()
    if _trouver_entreprise(connexion, nom, sauf_id=entreprise_id) is not None:
        connexion.close()
        return False, f"L'entreprise « {nom} » existe déjà."
    connexion.execute("UPDATE entreprises SET nom=?, adresse=? WHERE id=?", (nom, adresse, entreprise_id))
    connexion.execute("UPDATE lots SET entreprise=?, contact_adr=? WHERE entreprise_id=?",
                      (nom, adresse, entreprise_id))
    connexion.commit()
    connexion.close()
    return True, ""


def delete_entreprise(entreprise_id):
    """Supprime une fiche entreprise, UNIQUEMENT si aucun lot ni aucun dépôt n'y
    est rattaché. Ses chargés d'affaires sont supprimés avec elle.
    Renvoie (True, "") ou (False, message)."""
    connexion = conn()
    nb_lots = connexion.execute(
        "SELECT COUNT(*) FROM lots WHERE entreprise_id=?", (entreprise_id,)).fetchone()[0]
    nb_depots = connexion.execute(
        "SELECT COUNT(*) FROM depots WHERE entreprise_id=?", (entreprise_id,)).fetchone()[0]
    if nb_lots > 0 or nb_depots > 0:
        connexion.close()
        return False, (f"{nb_lots} lot(s) et {nb_depots} dépôt(s) rattaché(s) — "
                       "changez le titulaire de ces lots d'abord.")
    connexion.execute("DELETE FROM contacts_entreprise WHERE entreprise_id=?", (entreprise_id,))
    connexion.execute("DELETE FROM entreprises WHERE id=?", (entreprise_id,))
    connexion.commit()
    connexion.close()
    return True, ""


def set_lot_entreprise(lot_id, entreprise_id):
    """Choisit la fiche entreprise titulaire d'un lot ; son nom et son adresse
    sont recopiés dans le lot (destinataire du bordereau). Si l'entreprise
    change, le chargé d'affaires du lot est à désigner à nouveau."""
    connexion = conn()
    fiche = connexion.execute("SELECT nom, adresse FROM entreprises WHERE id=?",
                              (entreprise_id,)).fetchone()
    lot = connexion.execute("SELECT entreprise_id FROM lots WHERE id=?", (lot_id,)).fetchone()
    if fiche is not None and lot is not None:
        connexion.execute(
            "UPDATE lots SET entreprise_id=?, entreprise=?, contact_adr=? WHERE id=?",
            (entreprise_id, fiche["nom"], fiche["adresse"] or "", lot_id))
        if lot["entreprise_id"] != entreprise_id:
            _vider_contact_lot(connexion, lot_id)
        connexion.commit()
    connexion.close()


def get_lots_entreprise(entreprise_id):
    """Lots visibles par l'entreprise sur le portail : tous les lots dont elle est
    titulaire, dans les projets NON archivés. Chaque lot porte en plus
    projet_nom, projet_moa et projet_chef. Triés par projet puis n° de lot."""
    connexion = conn()
    lignes = connexion.execute(
        """SELECT l.*, p.nom AS projet_nom, p.moa AS projet_moa, p.chef AS projet_chef
           FROM lots l JOIN projets p ON p.code = l.projet
           WHERE l.entreprise_id=? AND IFNULL(p.archive,0)=0
           ORDER BY l.projet, l.lot_numero, l.lot""", (entreprise_id,)).fetchall()
    connexion.close()
    resultat = []
    for ligne in lignes:
        resultat.append(dict(ligne))
    return resultat


# ---------- Compte du portail (un par entreprise) ----------

def set_compte_entreprise(entreprise_id, identifiant, mot_de_passe):
    """Crée ou réinitialise le compte portail d'une entreprise.
    Le mot de passe est haché (bcrypt) : il n'est jamais enregistré en clair.
    Renvoie (True, "") ou (False, message)."""
    identifiant = (identifiant or "").strip().lower()
    if identifiant == "":
        return False, "L'identifiant est obligatoire."
    if len(mot_de_passe or "") < LONGUEUR_MINI_MOT_DE_PASSE:
        return False, f"Le mot de passe doit faire au moins {LONGUEUR_MINI_MOT_DE_PASSE} caractères."
    connexion = conn()
    autre = connexion.execute("SELECT nom FROM entreprises WHERE identifiant=? AND id<>?",
                              (identifiant, entreprise_id)).fetchone()
    if autre is not None:
        connexion.close()
        return False, f"L'identifiant « {identifiant} » est déjà utilisé par {autre['nom']}."
    hache = bcrypt.hashpw(mot_de_passe.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
    connexion.execute("UPDATE entreprises SET identifiant=?, mdp_hache=?, actif=1 WHERE id=?",
                      (identifiant, hache, entreprise_id))
    connexion.commit()
    connexion.close()
    return True, ""


def set_compte_actif(entreprise_id, valeur):
    """Active (1) ou désactive (0) le compte portail, sans rien supprimer."""
    connexion = conn()
    connexion.execute("UPDATE entreprises SET actif=? WHERE id=?", (int(valeur), entreprise_id))
    connexion.commit()
    connexion.close()


# Blocage des connexions au portail, contrôlé CÔTÉ SERVEUR (un rechargement
# de la page ne le contourne pas) : après MAX_ECHECS_CONNEXION échecs en
# DUREE_BLOCAGE_MINUTES pour un même identifiant, il est bloqué jusqu'à ce que
# le plus ancien de ces échecs ait plus de DUREE_BLOCAGE_MINUTES.
MAX_ECHECS_CONNEXION = 5
DUREE_BLOCAGE_MINUTES = 15
# Empreinte bcrypt quelconque, comparée quand l'identifiant n'existe pas : la
# réponse prend alors le même temps, ce qui ne trahit pas les identifiants existants.
_EMPREINTE_FACTICE = None


def minutes_de_blocage(identifiant):
    """Minutes restantes de blocage d'un identifiant du portail (0 = pas bloqué)."""
    identifiant = (identifiant or "").strip().lower()
    debut = dt.datetime.now() - dt.timedelta(minutes=DUREE_BLOCAGE_MINUTES)
    connexion = conn()
    dates = []
    for ligne in connexion.execute(
            "SELECT date FROM connexions_echouees WHERE identifiant=? AND date>=? ORDER BY date",
            (identifiant, debut.isoformat(timespec="seconds"))).fetchall():
        dates.append(ligne["date"])
    connexion.close()
    if len(dates) < MAX_ECHECS_CONNEXION:
        return 0
    # Débloqué quand le premier des échecs comptés sort de la fenêtre.
    fin = dt.datetime.fromisoformat(dates[-MAX_ECHECS_CONNEXION]) + dt.timedelta(minutes=DUREE_BLOCAGE_MINUTES)
    reste = (fin - dt.datetime.now()).total_seconds() / 60
    return max(1, int(reste + 0.999))


def connexion_entreprise(identifiant, mot_de_passe):
    """Vérifie l'identifiant et le mot de passe saisis sur le portail.
    Renvoie la fiche entreprise si tout est bon (compte actif), sinon None.
    Un identifiant bloqué (voir minutes_de_blocage) est refusé sans examen.
    Le message affiché ne doit pas dire lequel des deux est faux."""
    global _EMPREINTE_FACTICE
    identifiant = (identifiant or "").strip().lower()
    if identifiant == "" or not mot_de_passe:
        return None
    if minutes_de_blocage(identifiant):
        return None
    connexion = conn()
    ligne = connexion.execute("SELECT * FROM entreprises WHERE identifiant=?", (identifiant,)).fetchone()
    connexion.close()
    if ligne is not None and ligne["mdp_hache"]:
        empreinte = ligne["mdp_hache"].encode("utf-8")
    else:
        if _EMPREINTE_FACTICE is None:
            _EMPREINTE_FACTICE = bcrypt.hashpw(b"identifiant inconnu", bcrypt.gensalt())
        empreinte = _EMPREINTE_FACTICE
    mot_de_passe_juste = bcrypt.checkpw(mot_de_passe.encode("utf-8"), empreinte)
    connexion = conn()
    if ligne is None or not ligne["mdp_hache"] or not ligne["actif"] or not mot_de_passe_juste:
        connexion.execute("INSERT INTO connexions_echouees(identifiant, date) VALUES(?, ?)",
                          (identifiant, _maintenant()))
        # Ménage : les échecs de plus d'un jour ne servent plus à rien.
        hier = (dt.datetime.now() - dt.timedelta(days=1)).isoformat(timespec="seconds")
        connexion.execute("DELETE FROM connexions_echouees WHERE date<?", (hier,))
        connexion.commit()
        connexion.close()
        return None
    connexion.execute("DELETE FROM connexions_echouees WHERE identifiant=?", (identifiant,))
    connexion.commit()
    connexion.close()
    return get_entreprise(ligne["id"])


# ---------- Chargés d'affaires / contacts ----------

def get_contacts(entreprise_id, actifs_seulement=True):
    """Chargés d'affaires d'une entreprise, triés par nom : [{id, nom, email, actif}]."""
    connexion = conn()
    requete = "SELECT * FROM contacts_entreprise WHERE entreprise_id=?"
    if actifs_seulement:
        requete = requete + " AND actif=1"
    lignes = connexion.execute(requete + " ORDER BY nom", (entreprise_id,)).fetchall()
    connexion.close()
    resultat = []
    for ligne in lignes:
        resultat.append(dict(ligne))
    return resultat


def _verifier_contact(connexion, entreprise_id, nom, email, sauf_id=None):
    """Contrôles communs à l'ajout et à la modification d'un contact.
    Renvoie "" si tout va bien, sinon le message d'erreur."""
    if nom == "" or email == "":
        return "Le nom et le mail sont obligatoires."
    if not FORME_EMAIL.match(email):
        return f"« {email} » n'est pas une adresse mail valide."
    deja = connexion.execute(
        """SELECT id FROM contacts_entreprise
           WHERE entreprise_id=? AND lower(email)=lower(?) AND actif=1 AND id<>?""",
        (entreprise_id, email, sauf_id or 0)).fetchone()
    if deja is not None:
        return f"Le contact {email} existe déjà."
    return ""


def add_contact(entreprise_id, nom, email, civ="M.", telephone=""):
    """Ajoute un chargé d'affaires (nom + mail obligatoires, mail valide, pas de
    doublon de mail dans l'entreprise). Renvoie (True, id) ou (False, message)."""
    nom = (nom or "").strip()
    email = (email or "").strip()
    connexion = conn()
    erreur = _verifier_contact(connexion, entreprise_id, nom, email)
    if erreur:
        connexion.close()
        return False, erreur
    curseur = connexion.execute(
        """INSERT INTO contacts_entreprise(entreprise_id, civ, nom, email, telephone, actif)
           VALUES(?,?,?,?,?,1)""",
        (entreprise_id, civ or "M.", nom, email, (telephone or "").strip()))
    nouvel_id = curseur.lastrowid
    connexion.commit()
    connexion.close()
    return True, nouvel_id


def _copier_contact_dans_lots(connexion, contact_id):
    """Recopie civilité, nom, mail et téléphone du contact dans les lots dont il
    est le chargé d'affaires (bordereau, relances)."""
    contact = connexion.execute("SELECT * FROM contacts_entreprise WHERE id=?", (contact_id,)).fetchone()
    if contact is None:
        return
    connexion.execute(
        """UPDATE lots SET contact_civ=?, contact_nom=?, contact_email=?, contact_tel=?
           WHERE contact_id=?""",
        (contact["civ"] or "M.", contact["nom"], contact["email"], contact["telephone"] or "", contact_id))


def update_contact(contact_id, nom, email, civ="M.", telephone=""):
    """Modifie un chargé d'affaires ; la modification est recopiée dans ses lots.
    Renvoie (True, "") ou (False, message)."""
    nom = (nom or "").strip()
    email = (email or "").strip()
    connexion = conn()
    contact = connexion.execute("SELECT entreprise_id FROM contacts_entreprise WHERE id=?",
                                (contact_id,)).fetchone()
    if contact is None:
        connexion.close()
        return False, "Contact introuvable."
    erreur = _verifier_contact(connexion, contact["entreprise_id"], nom, email, sauf_id=contact_id)
    if erreur:
        connexion.close()
        return False, erreur
    connexion.execute(
        "UPDATE contacts_entreprise SET civ=?, nom=?, email=?, telephone=? WHERE id=?",
        (civ or "M.", nom, email, (telephone or "").strip(), contact_id))
    _copier_contact_dans_lots(connexion, contact_id)
    connexion.commit()
    connexion.close()
    return True, ""


def _vider_contact_lot(connexion, lot_id):
    """Le lot n'a plus de chargé d'affaires (à désigner)."""
    connexion.execute(
        """UPDATE lots SET contact_id=NULL, contact_nom='', contact_email='', contact_tel=''
           WHERE id=?""", (lot_id,))


def retirer_contact(contact_id):
    """Retire un chargé d'affaires de la liste. Il n'est pas effacé : les envois
    qu'il a faits gardent son nom et son mail. Les lots dont il était le chargé
    d'affaires en attendent un nouveau."""
    connexion = conn()
    connexion.execute("UPDATE contacts_entreprise SET actif=0 WHERE id=?", (contact_id,))
    for lot in connexion.execute("SELECT id FROM lots WHERE contact_id=?", (contact_id,)).fetchall():
        _vider_contact_lot(connexion, lot["id"])
    connexion.commit()
    connexion.close()


def set_lot_contact(lot_id, contact_id):
    """Désigne le chargé d'affaires d'un lot (None = aucun). Le contact doit
    appartenir à l'entreprise titulaire du lot. Appelé depuis la Plateforme visa partenaires ou depuis
    l'Administration : la dernière modification l'emporte.
    Renvoie (True, "") ou (False, message)."""
    connexion = conn()
    lot = connexion.execute("SELECT entreprise_id FROM lots WHERE id=?", (lot_id,)).fetchone()
    if lot is None:
        connexion.close()
        return False, "Lot introuvable."
    if contact_id is None:
        _vider_contact_lot(connexion, lot_id)
    else:
        contact = connexion.execute(
            "SELECT entreprise_id, actif FROM contacts_entreprise WHERE id=?", (contact_id,)).fetchone()
        if contact is None or not contact["actif"] or contact["entreprise_id"] != lot["entreprise_id"]:
            connexion.close()
            return False, "Ce contact n'appartient pas à l'entreprise titulaire du lot."
        connexion.execute("UPDATE lots SET contact_id=? WHERE id=?", (contact_id, lot_id))
        _copier_contact_dans_lots(connexion, contact_id)
    connexion.commit()
    connexion.close()
    return True, ""


def nom_avec_civilite(contact):
    """« M. Marc WEBER » : civilité puis nom d'un chargé d'affaires (par politesse,
    partout où le nom est affiché ou recopié)."""
    if not contact:
        return ""
    return f"{contact.get('civ') or ''} {contact.get('nom') or ''}".strip()


def get_contact_lot(lot_id):
    """Chargé d'affaires (actif) d'un lot : {id, civ, nom, email, telephone}, ou None."""
    connexion = conn()
    ligne = connexion.execute(
        """SELECT c.* FROM lots l JOIN contacts_entreprise c ON c.id = l.contact_id
           WHERE l.id=? AND c.actif=1""", (lot_id,)).fetchone()
    connexion.close()
    if ligne is None:
        return None
    return dict(ligne)


# ============================================================================
#  6. TRAMES DE VISA (référentiel configurable)
# ============================================================================

def get_trames():
    """Renvoie toutes les trames, avec leurs listes décodées :
    docs (liste de textes), controles (liste de {n}), keywords (liste)."""
    connexion = conn()
    lignes = connexion.execute("SELECT * FROM trames ORDER BY nom").fetchall()
    connexion.close()
    resultat = []
    for ligne in lignes:
        trame = dict(ligne)
        trame["docs"] = json.loads(trame["docs"] or "[]")
        # Les contrôles peuvent être d'ancien format (simples textes ou avec un
        # niveau majeur/mineur) : on ne garde que le libellé.
        controles_propres = []
        for element in json.loads(trame["controles"] or "[]"):
            if isinstance(element, dict):
                controles_propres.append({"n": element.get("n", "")})
            else:
                controles_propres.append({"n": element})
        trame["controles"] = controles_propres
        # La colonne keywords peut être absente sur de vieilles bases.
        if "keywords" in trame.keys() and trame["keywords"]:
            trame["keywords"] = json.loads(trame["keywords"])
        else:
            trame["keywords"] = []
        resultat.append(trame)
    return resultat


def trame_names():
    """Renvoie simplement la liste des noms de trames (pour les menus)."""
    noms = []
    for trame in get_trames():
        noms.append(trame["nom"])
    return noms


def get_trame_by_nom(nom):
    """Renvoie la trame portant ce nom, ou None si elle n'existe pas."""
    for trame in get_trames():
        if trame["nom"] == nom:
            return trame
    return None


def add_trame(nom):
    """Crée une trame vide (l'utilisateur la remplit ensuite dans l'interface).
    Si le nom est déjà pris, un numéro est ajouté. Renvoie le nom retenu."""
    connexion = conn()
    nom_libre = nom
    numero = 2
    while connexion.execute("SELECT 1 FROM trames WHERE nom=?", (nom_libre,)).fetchone():
        nom_libre = f"{nom} {numero}"
        numero = numero + 1
    connexion.execute(
        "INSERT INTO trames(nom, docs, controles, keywords) VALUES(?, '[]', '[]', '[]')", (nom_libre,))
    connexion.commit()
    connexion.close()
    return nom_libre


def duplicate_trame(trame_id):
    """Duplique une trame ("Nom (copie)", "Nom (copie 2)"...)."""
    connexion = conn()
    ligne = connexion.execute("SELECT * FROM trames WHERE id=?", (trame_id,)).fetchone()
    if ligne:
        nouveau_nom = ligne["nom"] + " (copie)"
        numero = 2
        while connexion.execute("SELECT 1 FROM trames WHERE nom=?", (nouveau_nom,)).fetchone():
            nouveau_nom = f"{ligne['nom']} (copie {numero})"
            numero = numero + 1
        if "keywords" in ligne.keys():
            mots_cles = ligne["keywords"]
        else:
            mots_cles = "[]"
        connexion.execute(
            "INSERT INTO trames(nom, docs, controles, keywords) VALUES(?,?,?,?)",
            (nouveau_nom, ligne["docs"], ligne["controles"], mots_cles))
    connexion.commit()
    connexion.close()


def save_trame(trame_id, nom, docs, controles, keywords=None):
    """Enregistre une trame complète (nom + les trois listes).
    Un renommage est reporté sur les positions et l'historique d'association
    qui utilisaient l'ancien nom. Les trames système ne se renomment pas.
    Renvoie (True, "") ou (False, message)."""
    nom = (nom or "").strip()
    connexion = conn()
    ligne = connexion.execute("SELECT nom FROM trames WHERE id=?", (trame_id,)).fetchone()
    if ligne is None:
        connexion.close()
        return False, "Trame introuvable."
    ancien_nom = ligne["nom"]
    if nom == "":
        connexion.close()
        return False, "Le nom de la trame ne peut pas être vide."
    if ancien_nom in TRAMES_SYSTEME and nom != ancien_nom:
        connexion.close()
        return False, f"La trame « {ancien_nom} » est utilisée par l'application et ne peut pas être renommée."
    if nom != ancien_nom and connexion.execute(
            "SELECT 1 FROM trames WHERE nom=?", (nom,)).fetchone():
        connexion.close()
        return False, f"Une trame s'appelle déjà « {nom} »."
    connexion.execute(
        "UPDATE trames SET nom=?, docs=?, controles=?, keywords=? WHERE id=?",
        (nom,
         json.dumps(docs, ensure_ascii=False),
         json.dumps(controles, ensure_ascii=False),
         json.dumps(keywords or [], ensure_ascii=False),
         trame_id))
    if nom != ancien_nom:
        connexion.execute("UPDATE positions SET famille=? WHERE famille=?", (nom, ancien_nom))
        connexion.execute("UPDATE assoc SET trame=? WHERE trame=?", (nom, ancien_nom))
    connexion.commit()
    connexion.close()
    return True, ""


def delete_trame(trame_id):
    """Supprime une trame (sauf trame système). Les positions déjà créées
    conservent leurs listes. Renvoie (True, "") ou (False, message)."""
    connexion = conn()
    ligne = connexion.execute("SELECT nom FROM trames WHERE id=?", (trame_id,)).fetchone()
    if ligne and ligne["nom"] in TRAMES_SYSTEME:
        connexion.close()
        return False, f"La trame « {ligne['nom']} » est utilisée par l'application et ne peut pas être supprimée."
    connexion.execute("DELETE FROM trames WHERE id=?", (trame_id,))
    connexion.commit()
    connexion.close()
    return True, ""


# ---------- tableau de trames (Excel) ----------
# Une ligne d'en-tête, puis une trame par ligne. Seules les 4 premières colonnes
# sont lues : A = nom de la trame, B = documents attendus, C = points de
# contrôle, D = mots-clés d'association. Dans B, C et D : un élément par ligne
# de la cellule (Alt+Entrée dans Excel). Les colonnes suivantes (aide) sont ignorées.
# Onglet lu : « Trames » s'il existe, sinon le premier.

def _inserer_trame(curseur, nom, docs, controles, mots_cles):
    """Ajoute une trame (si son nom n'existe pas déjà). docs, controles et
    mots_cles sont des listes de textes."""
    points = []
    for libelle in controles:
        points.append({"n": libelle})
    curseur.execute(
        "INSERT OR IGNORE INTO trames(nom, docs, controles, keywords) VALUES(?,?,?,?)",
        (nom,
         json.dumps(docs, ensure_ascii=False),
         json.dumps(points, ensure_ascii=False),
         json.dumps(mots_cles, ensure_ascii=False)))


def _elements_cellule(valeur):
    """Éléments d'une cellule du tableau : un par ligne, sans espaces superflus
    ni lignes vides. Exemple : "Fiche technique\\nNote de calcul \\n" -> 2 éléments."""
    elements = []
    for ligne in str(valeur or "").splitlines():
        if ligne.strip():
            elements.append(ligne.strip())
    return elements


def lire_trames_excel(source):
    """Lit un tableau de trames (chemin du fichier, ou son contenu en octets).
    Une ligne sans nom est ignorée, ainsi qu'un nom déjà lu plus haut.
    Renvoie une liste de {nom, docs, ctrl, kw} (listes de textes)."""
    if isinstance(source, (bytes, bytearray)):
        source = io.BytesIO(source)
    classeur = load_workbook(source, read_only=True, data_only=True)
    if "Trames" in classeur.sheetnames:
        onglet = classeur["Trames"]
    else:
        onglet = classeur.worksheets[0]
    trames = []
    noms_lus = set()
    numero_ligne = 0
    for ligne in onglet.iter_rows(min_col=1, max_col=4, values_only=True):
        numero_ligne = numero_ligne + 1
        if numero_ligne == 1:
            continue   # ligne d'en-tête
        valeurs = list(ligne) + [None, None, None, None]
        nom = " ".join(str(valeurs[0] or "").split())   # espaces en trop retirés
        if not nom or nom in noms_lus:
            continue
        noms_lus.add(nom)
        trames.append({"nom": nom, "docs": _elements_cellule(valeurs[1]),
                       "ctrl": _elements_cellule(valeurs[2]), "kw": _elements_cellule(valeurs[3])})
    classeur.close()
    return trames


_TRAMES_PAR_DEFAUT = []   # tableau livré, lu une seule fois


def trames_par_defaut():
    """Trames du tableau livré avec l'application (TRAMES_FICHIER), installées
    dans une base neuve. Liste vide si le fichier manque."""
    if not _TRAMES_PAR_DEFAUT and os.path.exists(TRAMES_FICHIER):
        _TRAMES_PAR_DEFAUT.extend(lire_trames_excel(TRAMES_FICHIER))
    return _TRAMES_PAR_DEFAUT


def remplacer_trames(trames):
    """Remplace TOUTES les trames par celles données (voir lire_trames_excel),
    sauf les trames système (« Autre document », « À définir »), conservées.
    Les positions déjà créées gardent leurs pièces attendues, leurs points de
    contrôle et le nom de leur ancienne trame. L'historique des validations
    qui menait à une trame disparue n'est plus proposé (voir suggest_trame).
    Une liste vide ne supprime rien. Renvoie le nombre de trames installées."""
    if not trames:
        return 0
    connexion = conn()
    curseur = connexion.cursor()
    marques = ",".join("?" for _ in TRAMES_SYSTEME)
    curseur.execute(f"DELETE FROM trames WHERE nom NOT IN ({marques})", TRAMES_SYSTEME)
    nombre = 0
    for trame in trames:
        if trame["nom"] in TRAMES_SYSTEME:
            continue
        _inserer_trame(curseur, trame["nom"], trame["docs"], trame["ctrl"], trame["kw"])
        nombre = nombre + 1
    connexion.commit()
    connexion.close()
    return nombre


# ============================================================================
#  7. POSITIONS (le coeur de l'outil)
# ============================================================================

def _controles_depuis_trame(trame_obj):
    """Points de contrôle d'une position, générés depuis une trame :
    ni vérifiés, ni sans objet, sans commentaire."""
    controles = []
    for point in trame_obj["controles"]:
        controles.append({"n": point["n"], "v": 0, "so": 0, "c": ""})
    return controles


def get_positions(projet, lot_id=None, archive=0):
    """Renvoie les positions d'un projet (et d'un lot si précisé).
    - archive=0 (défaut) : les positions ACTIVES, sur lesquelles on travaille.
    - archive=1 : les copies "Anciens visas".
    Chaque position est un dictionnaire ; les champs JSON sont décodés :
      "att"  : liste des pièces attendues   [{n, s, d}, ...]
      "ctrl" : liste des points de contrôle [{n, v, so, c}, ...]"""
    connexion = conn()
    if lot_id:
        lignes = connexion.execute(
            """SELECT * FROM positions
               WHERE projet=? AND lot_id=? AND IFNULL(archive,0)=?
               ORDER BY ordre, id""",
            (projet, lot_id, archive)).fetchall()
    else:
        lignes = connexion.execute(
            """SELECT * FROM positions
               WHERE projet=? AND IFNULL(archive,0)=?
               ORDER BY ordre, id""",
            (projet, archive)).fetchall()
    connexion.close()

    resultat = []
    for ligne in lignes:
        position = dict(ligne)

        # --- Décodage des pièces attendues + migration des anciens statuts ---
        # .pop(clé, None) : on retire la version texte du dictionnaire et on la
        # remplace par la version décodée.
        pieces = json.loads(position.pop("attendus", None) or "[]")
        for piece in pieces:
            ancien_statut = piece.get("s")
            piece["s"] = STATUT_MIGRATION.get(ancien_statut, ancien_statut)
        position["att"] = pieces

        # --- Décodage des points de contrôle (avec migration d'anciens formats) ---
        # Un seul type de point de contrôle : l'ancien niveau majeur/mineur est ignoré.
        controles_propres = []
        for element in json.loads(position.pop("controles", None) or "[]"):
            if isinstance(element, dict):
                # Ancien format : {"n", "s" ("Conforme"...), "c"} -> conversion.
                verifie = element.get("v")
                if verifie is None:
                    if element.get("s") == "Conforme":
                        verifie = 1
                    else:
                        verifie = 0
                controles_propres.append({
                    "n": element.get("n", ""),
                    "v": int(verifie),
                    "so": int(element.get("so", 0)),
                    "c": element.get("c", ""),
                })
            else:
                # Très ancien format : simple texte.
                controles_propres.append({"n": element, "v": 0, "so": 0, "c": ""})
        position["ctrl"] = controles_propres

        resultat.append(position)
    return resultat


def get_position(position_id):
    """Renvoie une position (dictionnaire décodé) à partir de son identifiant."""
    connexion = conn()
    ligne = connexion.execute(
        "SELECT projet, lot_id, archive FROM positions WHERE id=?", (position_id,)).fetchone()
    connexion.close()
    if ligne is None:
        return None
    for position in get_positions(ligne["projet"], ligne["lot_id"], int(ligne["archive"] or 0)):
        if position["id"] == position_id:
            return position
    return None


def save_position(position):
    """Enregistre TOUS les champs modifiables d'une position."""
    connexion = conn()
    connexion.execute(
        """UPDATE positions SET
             pos=?, des=?, ptype=?, famille=?, marque=?, typeprop=?,
             avis=?, relance=?, datevisa=?, intitule=?, rem=?,
             attendus=?, export=?, rem_histo=?, controles=?,
             marque_moe=?, type_moe=?, mt_verrou=?
           WHERE id=?""",
        (position["pos"], position["des"], position["ptype"], position["famille"],
         position.get("marque") or "", position.get("typeprop") or "",
         position["avis"], int(position["relance"]),
         position.get("datevisa") or "", position.get("intitule") or "",
         position.get("rem") or "",
         json.dumps(position["att"], ensure_ascii=False),
         int(position.get("export", 1)),
         position.get("rem_histo") or "",
         json.dumps(position.get("ctrl") or [], ensure_ascii=False),
         position.get("marque_moe") or "",
         position.get("type_moe") or "",
         int(position.get("mt_verrou") or 0),
         position["id"]))
    connexion.commit()
    connexion.close()


def add_position(projet, lot_id, pos, des, trame, ordre=999, default_export=1,
                 marque_moe="", type_moe="", mt_verrou=0):
    """Crée une position et renvoie son identifiant. Les pièces attendues et les
    points de contrôle sont générés depuis la trame indiquée (souvent
    "À définir" : rien n'est généré, l'utilisateur valide la trame ensuite).
    - marque_moe / type_moe : proposition du maître d'oeuvre (import DPGF).
    - mt_verrou=1 : ces champs seront affichés grisés dans l'interface."""
    trame_obj = get_trame_by_nom(trame)
    if trame_obj is None:
        trame_obj = {"docs": ["Fiche technique"], "controles": []}

    # Génération des pièces attendues : toutes en "Manquant" au départ.
    pieces = []
    for nom_doc in trame_obj["docs"]:
        pieces.append({"n": nom_doc, "s": "Manquant", "d": ""})

    connexion = conn()
    curseur = connexion.execute(
        """INSERT INTO positions(projet, lot_id, ordre, pos, des, ptype, famille,
             marque, typeprop, avis, relance, datevisa, intitule, rem,
             attendus, export, controles, marque_moe, type_moe, mt_verrou, archive)
           VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,0)""",
        (projet, lot_id, ordre, pos, des, type_of(trame), trame,
         "", "", "MQ", 1, "", "", "",
         json.dumps(pieces, ensure_ascii=False), int(default_export),
         json.dumps(_controles_depuis_trame(trame_obj), ensure_ascii=False),
         marque_moe or "", type_moe or "", int(mt_verrou)))
    nouvel_id = curseur.lastrowid
    connexion.commit()
    connexion.close()
    return nouvel_id


def delete_position(position_id):
    """Supprime une position, ses références de fichiers et ses rappels
    (le contenu des fichiers reste stocké : aucune destruction de pièce)."""
    connexion = conn()
    connexion.execute("DELETE FROM positions WHERE id=?", (position_id,))
    connexion.execute("DELETE FROM fichiers WHERE position_id=?", (position_id,))
    connexion.execute("DELETE FROM rappels_sb WHERE position_id=?", (position_id,))
    connexion.commit()
    connexion.close()


def set_relance_bulk(projet, lot_id, liste_avis=None, valeur=1):
    """Coche (valeur=1) ou décoche (0) la case "Relance" en masse.
    - liste_avis : ex. ["MQ","REF"] pour ne toucher que ces avis ;
      None = toutes les positions du lot."""
    connexion = conn()
    if liste_avis:
        marqueurs = ",".join("?" * len(liste_avis))
        connexion.execute(
            f"UPDATE positions SET relance=? WHERE projet=? AND lot_id=? AND avis IN ({marqueurs})",
            [valeur, projet, lot_id] + liste_avis)
    else:
        connexion.execute(
            "UPDATE positions SET relance=? WHERE projet=? AND lot_id=?",
            (valeur, projet, lot_id))
    connexion.commit()
    connexion.close()


def set_export_bulk(projet, lot_id, valeur=1):
    """Coche/décoche la case "Export" (bordereau) sur toutes les positions du lot."""
    connexion = conn()
    connexion.execute(
        "UPDATE positions SET export=? WHERE projet=? AND lot_id=?",
        (int(valeur), projet, lot_id))
    connexion.commit()
    connexion.close()


def archive_position(position_id):
    """Bouton "Nouveau visa" : dépose une COPIE figée de la position dans
    l'onglet "Anciens visas" (archive=1). La position d'origine reste active
    pour le cycle de visa suivant. Les références de fichiers sont dupliquées
    (le contenu stocké est partagé, rien n'est copié deux fois).
    La copie retient le bordereau du visa archivé (visa_bordereau_id) s'il a
    été émis : son n° s'affiche dans « Anciens visas ».
    Renvoie l'identifiant de la copie (ou None si la position n'existe pas)."""
    connexion = conn()
    ligne = connexion.execute("SELECT * FROM positions WHERE id=?", (position_id,)).fetchone()
    if ligne is None:
        connexion.close()
        return None
    source = dict(ligne)
    bordereau = _visa_emis(connexion, position_id, source["lot_id"])
    bordereau_id = bordereau["id"] if bordereau else None
    curseur = connexion.cursor()
    curseur.execute(
        """INSERT INTO positions(projet, lot_id, ordre, pos, des, ptype, famille,
             marque, typeprop, avis, relance, datevisa, intitule, rem,
             attendus, export, archive, rem_histo, controles,
             marque_moe, type_moe, mt_verrou, visa_bordereau_id, origine_id)
           VALUES(?,?,?,?,?,?,?,?,?,?,0,?,?,?,?,0,1,?,?,?,?,?,?,?)""",
        (source["projet"], source["lot_id"], source["ordre"], source["pos"], source["des"],
         source["ptype"], source["famille"],
         source.get("marque") or "", source.get("typeprop") or "",
         source["avis"],
         source.get("datevisa") or "", source.get("intitule") or "",
         source.get("rem") or "", source["attendus"],
         source.get("rem_histo") or "", source.get("controles") or "[]",
         source.get("marque_moe") or "", source.get("type_moe") or "",
         int(source.get("mt_verrou") or 0), bordereau_id, position_id))
    copie_id = curseur.lastrowid
    for fichier in connexion.execute(
            "SELECT * FROM fichiers WHERE position_id=?", (position_id,)).fetchall():
        curseur.execute(
            """INSERT INTO fichiers(position_id, attendu, nom, chemin, version,
                                    courant, date_upload, taille, contenu_id)
               VALUES(?,?,?,?,?,?,?,?,?)""",
            (copie_id, fichier["attendu"], fichier["nom"], fichier["chemin"],
             fichier["version"], fichier["courant"], fichier["date_upload"],
             fichier["taille"], fichier["contenu_id"]))
    connexion.commit()
    connexion.close()
    return copie_id


# ============================================================================
#  8. ASSOCIATION AUTOMATIQUE POSITION -> TRAME
# ============================================================================

def _norm(texte):
    """Normalise un texte pour comparaison : minuscules et sans accents.
    Exemple : "Régulation GTB" -> "regulation gtb"."""
    texte_decompose = unicodedata.normalize("NFD", (texte or "").lower())
    resultat = ""
    for caractere in texte_decompose:
        if unicodedata.category(caractere) != "Mn":  # "Mn" = les accents
            resultat = resultat + caractere
    return resultat


def record_assoc(designation, nom_trame):
    """Mémorise une validation de trame (apprentissage). La prochaine fois
    qu'une désignation SIMILAIRE apparaîtra, cette trame sera proposée."""
    connexion = conn()
    connexion.execute(
        "INSERT INTO assoc(des, trame, date) VALUES(?,?,?)",
        (_norm(designation), nom_trame, dt.date.today().isoformat()))
    connexion.commit()
    connexion.close()


def suggest_trame(designation):
    """Propose la trame la plus adaptée à une désignation, en trois étapes :
      1. HISTORIQUE : une désignation déjà validée qui ressemble (>= 72 %) ?
      2. MOTS-CLÉS  : les mots-clés configurés sur chaque trame.
      3. REPLI      : "Autre document".
    Renvoie le nom de la trame proposée."""
    designation_normalisee = _norm(designation)

    # --- Étape 1 : recherche dans l'historique des validations ---
    connexion = conn()
    historique = connexion.execute(
        "SELECT des, trame FROM assoc ORDER BY id DESC LIMIT 500").fetchall()
    connexion.close()
    noms_existants = set(trame_names())
    meilleure_trame_historique = None
    meilleur_score_historique = 0.0
    for entree in historique:
        if entree["trame"] not in noms_existants:
            continue   # trame supprimée ou remplacée depuis : plus proposée
        # ratio() renvoie un score de ressemblance entre 0 et 1.
        score = difflib.SequenceMatcher(None, designation_normalisee, entree["des"]).ratio()
        if score > meilleur_score_historique:
            meilleur_score_historique = score
            meilleure_trame_historique = entree["trame"]
    if meilleure_trame_historique is not None and meilleur_score_historique >= 0.72:
        return meilleure_trame_historique

    # --- Étape 2 : mots-clés des trames ---
    meilleure_trame_mots_cles = None
    meilleur_score_mots_cles = 0.0
    for trame in get_trames():
        score_trame = 0.0
        for mot_cle in trame["keywords"]:
            mot_normalise = _norm(mot_cle)
            if mot_normalise and mot_normalise in designation_normalisee:
                # Le mot-clé est contenu dans la désignation : score fort.
                # Un mot-clé plus long est plus discriminant, donc mieux noté.
                bonus_longueur = 0.1 * len(mot_normalise) / max(len(designation_normalisee), 1)
                score_candidat = 0.9 + bonus_longueur
            else:
                # Sinon : simple ressemblance globale, plafonnée à 0.8.
                score_candidat = difflib.SequenceMatcher(
                    None, designation_normalisee, mot_normalise).ratio() * 0.8
            if score_candidat > score_trame:
                score_trame = score_candidat
        # Le nom de la trame lui-même compte comme un mot-clé : contenu dans la
        # désignation, il l'emporte sur un mot-clé plus court d'une autre trame
        # (« Régulateur d'eau électrochimique » plutôt que « régulateur »).
        nom_normalise = _norm(trame["nom"])
        if nom_normalise and nom_normalise in designation_normalisee:
            score_nom = 0.9 + 0.1 * len(nom_normalise) / max(len(designation_normalisee), 1)
            if score_nom > score_trame:
                score_trame = score_nom
        if score_trame > meilleur_score_mots_cles:
            meilleur_score_mots_cles = score_trame
            meilleure_trame_mots_cles = trame["nom"]
    if meilleure_trame_mots_cles is not None and meilleur_score_mots_cles >= 0.45:
        return meilleure_trame_mots_cles

    # --- Étape 3 : repli ---
    if meilleure_trame_historique is not None:
        return meilleure_trame_historique
    return TRAME_AUTRE


def apply_trame(position, nom_trame):
    """Applique une trame à une position :
    - les pièces attendues sont générées SEULEMENT si la liste est vide
      (on ne détruit jamais des statuts déjà saisis) ;
    - les points de contrôle sont (re)générés depuis la trame ;
    - la famille et le type de document sont mis à jour.
    La position est enregistrée en base à la fin."""
    trame_obj = get_trame_by_nom(nom_trame)
    if trame_obj is None:
        trame_obj = {"docs": [], "controles": []}
    position["famille"] = nom_trame
    position["ptype"] = type_of(nom_trame)
    if not position.get("att"):
        pieces = []
        for nom_doc in trame_obj["docs"]:
            pieces.append({"n": nom_doc, "s": "Manquant", "d": ""})
        position["att"] = pieces
    position["ctrl"] = _controles_depuis_trame(trame_obj)
    save_position(position)


# ============================================================================
#  9. RELANCES TITULAIRE (journal, fréquence, mail groupé)
# ============================================================================

def log_relance(projet, lot_id, niveau, references):
    """Ajoute une entrée au journal des relances (appelée quand le brouillon
    .eml est généré). 'references' = liste des positions relancées."""
    connexion = conn()
    connexion.execute(
        """INSERT INTO relances(projet, lot_id, date, niveau, nb_positions, positions)
           VALUES(?,?,?,?,?,?)""",
        (projet, lot_id, dt.date.today().isoformat(), niveau,
         len(references), json.dumps(references, ensure_ascii=False)))
    connexion.commit()
    connexion.close()


def get_relances(projet, lot_id):
    """Renvoie le journal des relances d'un lot, la plus récente en premier."""
    connexion = conn()
    lignes = connexion.execute(
        "SELECT * FROM relances WHERE projet=? AND lot_id=? ORDER BY date DESC, id DESC",
        (projet, lot_id)).fetchall()
    connexion.close()
    resultat = []
    for ligne in lignes:
        resultat.append(dict(ligne))
    return resultat


def update_last_relance_date(lot_id, date_iso):
    """Corrige la date de la DERNIÈRE relance du lot (éditable au dashboard)."""
    connexion = conn()
    connexion.execute(
        """UPDATE relances SET date=? WHERE id=(
             SELECT id FROM relances WHERE lot_id=? ORDER BY date DESC, id DESC LIMIT 1)""",
        (date_iso, lot_id))
    connexion.commit()
    connexion.close()


def positions_a_relancer(positions):
    """Positions qui justifient une relance du titulaire (MQ, VAOB, REF)."""
    resultat = []
    for position in positions:
        if position["avis"] in ("MQ", "VAOB", "REF"):
            resultat.append(position)
    return resultat


def copies_pour_relance(positions):
    """Copies des positions d'un lot prêtes pour build_eml, pour le bouton
    « Relancer » du Dashboard : la case « Relance » est cochée sur toutes les
    positions MQ / VAOB / REF et décochée ailleurs, quel que soit son état
    enregistré. Les positions d'origine ne sont pas modifiées."""
    copies = []
    for position in positions:
        copie = dict(position)
        if position["avis"] in ("MQ", "VAOB", "REF"):
            copie["relance"] = 1
        else:
            copie["relance"] = 0
        copies.append(copie)
    return copies


def relance_status(projet, lot, positions):
    """Calcule l'état de relance d'un lot. Renvoie un tuple de 4 valeurs :
      (relance_a_faire : booléen,
       date_derniere   : texte ISO ou None,
       date_prochaine  : texte ISO ou None,
       nb_candidates   : nombre de positions en MQ/VAOB/REF)."""
    candidates = positions_a_relancer(positions)
    historique = get_relances(projet, lot["id"])
    if historique:
        date_derniere = historique[0]["date"]
    else:
        date_derniere = None

    if len(candidates) == 0:
        return (False, date_derniere, None, 0)
    if date_derniere is None:
        # Jamais relancé alors qu'il y a des candidates : relance à faire.
        return (True, None, None, len(candidates))
    frequence_jours = int(lot.get("freq_jours") or FREQ_RELANCE)
    date_prochaine = dt.date.fromisoformat(date_derniere) + dt.timedelta(days=frequence_jours)
    relance_a_faire = dt.date.today() >= date_prochaine
    return (relance_a_faire, date_derniere, date_prochaine.isoformat(), len(candidates))


def _grouped_lines(positions):
    """Prépare les lignes du mail de relance, groupées par nature d'avis :
    MQ (à fournir), VAOB (à corriger), REF (à re-soumettre)."""
    groupes = {"MQ": [], "VAOB": [], "REF": [], "AUTRE": []}
    for position in positions:
        if not position["relance"]:
            continue  # Seules les positions cochées "Relance" sont concernées.
        reference = (position.get("pos") or "").strip()
        if reference:
            prefixe = reference + " — "
        else:
            prefixe = ""
        if position["avis"] == "MQ":
            # Pour un MQ, on liste les pièces non reçues.
            pieces_manquantes = []
            for piece in position["att"]:
                if piece["s"] not in STATUTS_FOURNIS:
                    pieces_manquantes.append(piece["n"])
            detail = ", ".join(pieces_manquantes)
            if detail == "":
                detail = "pièces à fournir"
            groupes["MQ"].append(f"{prefixe}{position['des']} : {detail}")
        elif position["avis"] == "VAOB":
            remarque = (position.get("rem") or "").strip()
            if remarque == "":
                remarque = "voir observations du bordereau"
            groupes["VAOB"].append(f"{prefixe}{position['des']} : {remarque}")
        elif position["avis"] == "REF":
            remarque = (position.get("rem") or "").strip()
            if remarque == "":
                remarque = "document refusé"
            groupes["REF"].append(f"{prefixe}{position['des']} : {remarque}")
        else:
            groupes["AUTRE"].append(f"{prefixe}{position['des']}")
    return groupes


def _entete_mail(texte):
    """Valeur d'un en-tête de mail (To, Subject) sur UNE ligne : un retour à la
    ligne glissé dans un nom ou un mail ajouterait des en-têtes (ex. « Bcc »)."""
    return " ".join(str(texte or "").split())


def build_eml(projet, lot, positions):
    """Construit le brouillon de mail de relance (.eml, prêt à ouvrir dans
    Outlook). Le corps est structuré en trois sections MQ / VAOB / REF.
    Pas de formule de politesse ni de signature : Outlook ajoute celle de
    l'utilisateur. Renvoie le contenu du fichier en octets."""
    groupes = _grouped_lines(positions)
    civilite = lot.get("contact_civ") or "M."
    nom_contact = lot.get("contact_nom") or ""
    email_contact = lot.get("contact_email") or ""
    libelle = libelle_lot(lot)

    # --- Construction des sections, seulement si elles ont du contenu ---
    sections = []
    if groupes["MQ"]:
        lignes_puces = []
        for ligne in groupes["MQ"]:
            lignes_puces.append("   \u2022  " + ligne)
        sections.append("1. PIÈCES MANQUANTES À FOURNIR (avis MQ)\n" + "\n".join(lignes_puces))
    if groupes["VAOB"]:
        lignes_puces = []
        for ligne in groupes["VAOB"]:
            lignes_puces.append("   \u2022  " + ligne)
        sections.append(
            "2. DOCUMENTS À CORRIGER — OBSERVATIONS BLOQUANTES (avis VAOB)\n"
            "Les documents suivants ont été visés avec observations bloquantes ; "
            "une nouvelle diffusion intégrant les corrections est attendue :\n"
            + "\n".join(lignes_puces))
    if groupes["REF"]:
        lignes_puces = []
        for ligne in groupes["REF"]:
            lignes_puces.append("   \u2022  " + ligne)
        sections.append(
            "3. DOCUMENTS REFUSÉS — NOUVELLE SOUMISSION REQUISE (avis REF)\n"
            "Les documents suivants ont été refusés ; ils doivent faire l'objet "
            "d'une nouvelle soumission complète :\n" + "\n".join(lignes_puces))
    if groupes["AUTRE"]:
        lignes_puces = []
        for ligne in groupes["AUTRE"]:
            lignes_puces.append("   \u2022  " + ligne)
        sections.append("AUTRES POSITIONS SIGNALÉES\n" + "\n".join(lignes_puces))

    # --- Assemblage du corps du mail ---
    corps_sections = "\n\n".join(sections)
    if nom_contact:
        nom_court = nom_contact.split()[-1]  # "Thomas BERNARD" -> "BERNARD"
    else:
        nom_court = ""
    corps = (
        f"Bonjour {civilite} {nom_court},\n\n"
        f"Dans le cadre du suivi des visas du {libelle} de l'opération {projet['nom']}, "
        f"nous attirons votre attention sur les positions suivantes :\n\n"
        f"{corps_sections}\n\n"
        "Nous vous remercions de nous transmettre ces éléments dans les meilleurs délais.\n")
    sujet = f"{projet['nom']} — {libelle} — Relance visas — Pièces en attente / à corriger"
    # "X-Unsent: 1" indique à Outlook d'ouvrir le mail en mode brouillon.
    contenu_eml = (f"To: {_entete_mail(email_contact)}\r\nSubject: {_entete_mail(sujet)}\r\nX-Unsent: 1\r\n"
                   f"Content-Type: text/plain; charset=utf-8\r\n\r\n{corps}")
    return contenu_eml.encode("utf-8")


# ============================================================================
#  10. RAPPELS DE VISA SOLARES BAUEN
# ----------------------------------------------------------------------------
#  Quand le titulaire dépose un document, c'est à NOUS de le viser : un
#  rappel s'ouvre automatiquement avec une échéance « à viser avant le »
#  (date du dépôt + délai du lot, 10 jours ouvrés par défaut). Il reste
#  ouvert jusqu'à ce qu'on le clôture à la main (bouton « Visé »).
#  Un nouveau dépôt sur une position qui a déjà un rappel ouvert s'ajoute à
#  ce rappel, sans repousser l'échéance.
# ============================================================================

def ouvrir_rappel_sb(projet, lot, position_id, noms_fichiers, date_depot=None):
    """Ouvre (ou complète) le rappel de visa d'une position.
    - date_depot : date ISO du dépôt (aujourd'hui par défaut).
    Renvoie l'identifiant du rappel."""
    if date_depot is None:
        date_depot = dt.date.today().isoformat()
    connexion = conn()
    existant = connexion.execute(
        "SELECT id, fichiers FROM rappels_sb WHERE position_id=? AND cloture=0",
        (position_id,)).fetchone()
    if existant:
        deja = json.loads(existant["fichiers"] or "[]")
        for nom in noms_fichiers:
            if nom not in deja:
                deja.append(nom)
        connexion.execute("UPDATE rappels_sb SET fichiers=? WHERE id=?",
                          (json.dumps(deja, ensure_ascii=False), existant["id"]))
        rappel_id = existant["id"]
    else:
        delai = int(lot.get("delai_visa_sb") or DELAI_VISA_SB)
        echeance = calendrier.ajouter_jours_ouvres(dt.date.fromisoformat(date_depot), delai)
        curseur = connexion.execute(
            """INSERT INTO rappels_sb(projet, lot_id, position_id, date_depot, echeance,
                                      fichiers, cloture)
               VALUES(?,?,?,?,?,?,0)""",
            (projet, lot["id"], position_id, date_depot, echeance.isoformat(),
             json.dumps(list(noms_fichiers), ensure_ascii=False)))
        rappel_id = curseur.lastrowid
    connexion.commit()
    connexion.close()
    return rappel_id


def cloturer_rappel_sb(rappel_id):
    """Clôture d'un rappel : le document a été visé (voir aussi emettre_bordereau)."""
    connexion = conn()
    connexion.execute("UPDATE rappels_sb SET cloture=1, date_cloture=? WHERE id=?",
                      (dt.date.today().isoformat(), rappel_id))
    connexion.commit()
    connexion.close()


def rappel_ouvert(position_id):
    """Rappel ouvert d'une position (dictionnaire) ou None."""
    for rappel in get_rappels_sb(position_id=position_id):
        return rappel
    return None


def get_rappels_sb(projet=None, lot_id=None, position_id=None, ouverts_seulement=True):
    """Liste des rappels de visa, échéance la plus proche en premier.
    Chaque rappel est complété par la position (pos, des, ptype) et son état :
      "jours_restants" : jours ouvrés avant l'échéance (négatif = retard)
      "niveau"         : "retard", "proche" (≤ 2 jours ouvrés) ou "ok"."""
    conditions = []
    valeurs = []
    if ouverts_seulement:
        conditions.append("r.cloture=0")
    if projet:
        conditions.append("r.projet=?")
        valeurs.append(projet)
    if lot_id:
        conditions.append("r.lot_id=?")
        valeurs.append(lot_id)
    if position_id:
        conditions.append("r.position_id=?")
        valeurs.append(position_id)
    filtre = ""
    if conditions:
        filtre = "WHERE " + " AND ".join(conditions)
    connexion = conn()
    lignes = connexion.execute(
        f"""SELECT r.*, p.pos AS pos, p.des AS des, p.ptype AS ptype
            FROM rappels_sb r LEFT JOIN positions p ON p.id = r.position_id
            {filtre} ORDER BY r.echeance, r.id""", valeurs).fetchall()
    connexion.close()
    aujourd_hui = dt.date.today()
    resultat = []
    for ligne in lignes:
        rappel = dict(ligne)
        rappel["fichiers"] = json.loads(rappel["fichiers"] or "[]")
        restants = calendrier.jours_ouvres_restants(
            aujourd_hui, dt.date.fromisoformat(rappel["echeance"]))
        rappel["jours_restants"] = restants
        if restants < 0:
            rappel["niveau"] = "retard"
        elif restants <= 2:
            rappel["niveau"] = "proche"
        else:
            rappel["niveau"] = "ok"
        resultat.append(rappel)
    return resultat


# ============================================================================
#  11. IMPORT DPGF (annexe marque/type du CCTP)
# ============================================================================

def lire_dpgf(contenu_fichier):
    """Lit l'annexe marque/type (fichier Excel). Renvoie deux listes :
      - retenues : tuples (référence, désignation, marque_MOE, type_MOE) ;
      - ecartees : tuples (référence, désignation) des lignes dont la
        référence n'est pas une numérotation (titres de chapitre...).
    Structure attendue du fichier :
      colonne 1 = référence CCTP        colonne 2 = désignation
      colonnes 3-4 = proposition du soumissionnaire (ignorées à l'import)
      colonnes 5-6 = proposition du Maître d'oeuvre (marque / type et référence)
    Les lignes sans référence ou sans désignation sont ignorées sans être listées."""
    classeur = load_workbook(io.BytesIO(contenu_fichier), data_only=True)
    feuille = classeur[classeur.sheetnames[0]]
    retenues = []
    ecartees = []
    for numero_ligne in range(1, feuille.max_row + 1):
        reference = feuille.cell(numero_ligne, 1).value
        designation = feuille.cell(numero_ligne, 2).value
        marque_moe = feuille.cell(numero_ligne, 5).value
        type_moe = feuille.cell(numero_ligne, 6).value
        # Filtre : il faut une référence ET une désignation.
        if not reference or not designation:
            continue
        reference_texte = str(reference).strip()
        designation_texte = str(designation).strip()
        # La référence doit être une numérotation (ex. "3.1.2" oui, "CHAPITRE 3" non).
        if not REF_POSITION.match(reference_texte):
            ecartees.append((reference_texte, designation_texte))
            continue
        # Les valeurs MOE peuvent être vides : on renvoie alors "".
        marque_moe = str(marque_moe).strip() if marque_moe else ""
        type_moe = str(type_moe).strip() if type_moe else ""
        retenues.append((reference_texte, designation_texte, marque_moe, type_moe))
    return retenues, ecartees


def import_dpgf(contenu_fichier):
    """Lignes de positions retenues de l'annexe marque/type (voir lire_dpgf)."""
    return lire_dpgf(contenu_fichier)[0]


def feuilles_dpgf(lignes):
    """Retire les sous-titres de section : une référence "3" est un sous-titre
    si une référence "3.1" (ou "3.x...") existe aussi dans le fichier."""
    toutes_refs = []
    for ligne in lignes:
        toutes_refs.append(ligne[0])
    resultat = []
    for ligne in lignes:
        est_sous_titre = False
        for autre in toutes_refs:
            if autre != ligne[0] and autre.startswith(ligne[0] + "."):
                est_sous_titre = True
                break
        if not est_sous_titre:
            resultat.append(ligne)
    return resultat


# ============================================================================
#  12. FICHIERS PDF (contenu stocké via stockage.py)
# ============================================================================

def _san(segment, longueur_max=80):
    """Assainit un texte pour en faire un nom de fichier valide :
    remplace les caractères interdits par "_" et limite la longueur."""
    texte = re.sub(r'[\\/:*?"<>|]+', "_", (segment or "").strip())
    texte = texte[:longueur_max].rstrip(" .")
    if texte == "":
        texte = "SANS-REF"
    return texte


def _san_file(nom_fichier):
    """Assainit un nom de fichier en conservant son extension (en minuscules)."""
    base, extension = os.path.splitext(nom_fichier or "document.pdf")
    return _san(base, 120) + extension.lower()


def next_version(position_id, nom_fichier):
    """Numéro de version suivant pour un même nom de fichier sur une position."""
    connexion = conn()
    ligne = connexion.execute(
        "SELECT MAX(version) FROM fichiers WHERE position_id=? AND nom=?",
        (position_id, nom_fichier)).fetchone()
    connexion.close()
    version_max = ligne[0]
    if version_max is None:
        version_max = 0
    return version_max + 1


def save_fichier(projet, lot, position, nom_original, contenu, attendu=None,
                 date_depot=None):
    """Enregistre un PDF déposé sur une position :
      1. range son contenu via stockage.py ;
      2. journalise le fichier en base (avec numéro de version) ;
      3. ouvre (ou complète) le rappel de visa côté Solares Bauen.
    'attendu' = texte JSON de la liste des attendus rattachés (ou None).
    Renvoie un dictionnaire : {ok, version, nom, rappel_id}."""
    nom_propre = _san_file(nom_original)
    version = next_version(position["id"], nom_propre)
    if date_depot is None:
        date_depot = dt.date.today().isoformat()
    connexion = conn()
    reference = stockage.enregistrer(connexion, contenu)
    # La notion de "courant" reste en base pour l'historique interne.
    connexion.execute(
        "UPDATE fichiers SET courant=0 WHERE position_id=? AND nom=?",
        (position["id"], nom_propre))
    connexion.execute(
        """INSERT INTO fichiers(position_id, attendu, nom, chemin, version,
                                courant, date_upload, taille, contenu_id)
           VALUES(?,?,?,'',?,1,?,?,?)""",
        (position["id"], attendu, nom_propre, version, date_depot, len(contenu), reference))
    connexion.commit()
    connexion.close()
    rappel_id = ouvrir_rappel_sb(projet["code"], lot, position["id"], [nom_propre], date_depot)
    return {"ok": True, "version": version, "nom": nom_propre, "rappel_id": rappel_id}


def marquer_a_verifier(position, noms_attendus=None, date_reception=None):
    """Après un dépôt, passe des pièces attendues en « À vérifier » (avec la
    date de réception, aujourd'hui par défaut) :
      - noms_attendus fournis : seulement ces pièces ;
      - sinon : toutes les pièces pas encore « Reçu et ok ».
    Une pièce déjà « Reçu et ok » n'est jamais rétrogradée. Ne sauvegarde pas."""
    if date_reception is None:
        date_reception = dt.date.today().isoformat()
    for piece in position["att"]:
        if piece["s"] == "Reçu et ok":
            continue
        if noms_attendus and piece["n"] not in noms_attendus:
            continue
        piece["s"] = STATUT_A_VERIFIER
        piece["d"] = piece.get("d") or date_reception


def get_fichiers(position_id):
    """Liste les fichiers d'une position (les plus récents d'abord)."""
    connexion = conn()
    lignes = connexion.execute(
        "SELECT * FROM fichiers WHERE position_id=? ORDER BY nom, version DESC",
        (position_id,)).fetchall()
    connexion.close()
    resultat = []
    for ligne in lignes:
        resultat.append(dict(ligne))
    return resultat


def delete_fichier(fichier_id):
    """Retire un fichier de la liste (son contenu reste stocké : on ne
    détruit jamais une pièce depuis l'application)."""
    connexion = conn()
    connexion.execute("DELETE FROM fichiers WHERE id=?", (fichier_id,))
    connexion.commit()
    connexion.close()


def read_fichier(fichier):
    """Renvoie le contenu (octets) d'un fichier de la liste, ou None s'il est
    introuvable. Les dépôts antérieurs à la v7 sont relus sur le disque."""
    if fichier.get("contenu_id"):
        connexion = conn()
        contenu = stockage.lire(connexion, fichier["contenu_id"])
        connexion.close()
        return contenu
    chemin = fichier.get("chemin") or ""
    if chemin:
        try:
            with open(chemin, "rb") as flux:
                return flux.read()
        except OSError:
            return None
    return None


# ============================================================================
#  13. EXPORTS (bordereau Excel + PDF, statut des visas)
# ============================================================================

def compose_attendus(pieces):
    """Compose le texte de la colonne "Pièces attendues / statut" du bordereau.
    Exemple : "[reçu] Fiche technique · [manquant] Courbe pompe"."""
    morceaux = []
    for piece in pieces:
        etiquette = TAG.get(piece["s"], "")
        morceaux.append((etiquette + " " + piece["n"]).strip())
    if len(morceaux) == 0:
        return "—"
    return " · ".join(morceaux)


def last_recu(pieces):
    """Renvoie la date (ISO) de la pièce reçue (ou à vérifier) la plus récente, ou ""."""
    dates = []
    for piece in pieces:
        if piece["s"] in STATUTS_FOURNIS and piece.get("d"):
            dates.append(piece["d"])
    if len(dates) == 0:
        return ""
    return max(dates)


def _date(texte_iso):
    """Convertit "2026-07-08" en objet date (pour qu'Excel le formate en date).
    Renvoie None si le texte est vide ou invalide."""
    if not texte_iso:
        return None
    try:
        annee, mois, jour = str(texte_iso).split("-")
        return dt.date(int(annee), int(mois), int(jour))
    except (ValueError, TypeError):
        return None


def _ensure_rows(feuille, ligne_depart, lignes_gabarit, lignes_necessaires,
                 nb_colonnes, ligne_modele):
    """Le gabarit officiel ne prévoit qu'un petit nombre de lignes de données.
    S'il en faut davantage, on INSÈRE de vraies lignes (insert_rows) puis :
      1. on décale nous-mêmes les cellules fusionnées et les hauteurs de ligne
         situées en dessous (légende du bas) : openpyxl ne le fait pas ;
      2. on copie sur les nouvelles lignes le style, la hauteur et les
         fusions de la ligne modèle (ex. A:C, H:I, K:M de l'onglet Plans)."""
    from copy import copy as copier_style
    if lignes_necessaires <= lignes_gabarit:
        return  # Le gabarit suffit : rien à insérer.
    lignes_a_ajouter = lignes_necessaires - lignes_gabarit
    position_insertion = ligne_depart + lignes_gabarit

    # Hauteurs des lignes du bas, notées AVANT l'insertion.
    hauteurs_du_bas = {}
    for numero_ligne in range(position_insertion, feuille.max_row + 1):
        hauteurs_du_bas[numero_ligne] = feuille.row_dimensions[numero_ligne].height
    # Fusions de la ligne modèle, à reproduire sur les nouvelles lignes.
    fusions_modele = []
    for fusion in feuille.merged_cells.ranges:
        if fusion.min_row == ligne_modele and fusion.max_row == ligne_modele:
            fusions_modele.append((fusion.min_col, fusion.max_col))

    feuille.insert_rows(position_insertion, amount=lignes_a_ajouter)

    # 1. Décalage des fusions et des hauteurs situées sous le point d'insertion.
    for fusion in feuille.merged_cells.ranges:
        if fusion.min_row >= position_insertion:
            fusion.shift(row_shift=lignes_a_ajouter)
    for numero_ligne, hauteur in hauteurs_du_bas.items():
        feuille.row_dimensions[numero_ligne + lignes_a_ajouter].height = hauteur

    # 2. Nouvelles lignes : style, hauteur et fusions de la ligne modèle.
    for indice in range(lignes_a_ajouter):
        numero_ligne = position_insertion + indice
        for numero_colonne in range(1, nb_colonnes + 1):
            cellule_modele = feuille.cell(ligne_modele, numero_colonne)
            cellule_cible = feuille.cell(numero_ligne, numero_colonne)
            cellule_cible._style = copier_style(cellule_modele._style)
        feuille.row_dimensions[numero_ligne].height = feuille.row_dimensions[ligne_modele].height
        for colonne_debut, colonne_fin in fusions_modele:
            feuille.merge_cells(start_row=numero_ligne, start_column=colonne_debut,
                                end_row=numero_ligne, end_column=colonne_fin)

    # 3. Mise en page : zone d'impression, sauts de page et couleurs des avis.
    _decaler_mise_en_page(feuille, position_insertion, lignes_a_ajouter)


def _decaler_mise_en_page(feuille, position_insertion, lignes_a_ajouter):
    """Après l'insertion de lignes, openpyxl ne met pas à jour la mise en page :
      - la zone d'impression garde son ancienne fin (la légende était coupée) ;
      - les sauts de page manuels restent au même numéro de ligne ;
      - la mise en forme conditionnelle (couleur des avis) ne couvre que les
        lignes de données du gabarit.
    On corrige ces trois points ici."""
    # Zone d'impression : on recule sa dernière ligne.
    if feuille.print_area:
        zone = feuille.print_area.split("!")[-1].replace("$", "")
        plage = CellRange(zone)
        if plage.max_row >= position_insertion:
            plage.expand(down=lignes_a_ajouter)
        feuille.print_area = plage.coord

    # Sauts de page situés sous le point d'insertion : décalés eux aussi.
    anciens_sauts = list(feuille.row_breaks.brk)
    feuille.row_breaks = RowBreak()
    for saut in anciens_sauts:
        numero = saut.id
        if numero >= position_insertion:
            numero = numero + lignes_a_ajouter
        feuille.row_breaks.append(Break(id=numero))

    # Mise en forme conditionnelle : une plage qui finit juste au-dessus du point
    # d'insertion (les lignes de données) est agrandie ; une plage située en
    # dessous est décalée.
    anciennes_regles = []
    for mise_en_forme in feuille.conditional_formatting:
        for regle in mise_en_forme.rules:
            anciennes_regles.append((mise_en_forme.sqref, regle))
    feuille.conditional_formatting = ConditionalFormattingList()
    for plages, regle in anciennes_regles:
        nouvelles_plages = []
        for plage in plages.ranges:
            plage = CellRange(plage.coord)
            if plage.min_row >= position_insertion:
                plage.shift(row_shift=lignes_a_ajouter)
            elif plage.max_row >= position_insertion - 1:
                plage.expand(down=lignes_a_ajouter)
            nouvelles_plages.append(plage.coord)
        feuille.conditional_formatting.add(" ".join(nouvelles_plages), regle)


# Hauteur d'une ligne de texte en Arial 9 (police des lignes du bordereau,
# voir _habiller_bordereau), en points, et marge haut + bas.
HAUTEUR_LIGNE_TEXTE = 11.5
MARGE_CELLULE = 6
# Hauteur normale d'une ligne de données du gabarit : une ligne courte la garde.
HAUTEUR_MINI_BORDEREAU = 30.75
# Format « Date courte » intégré d'Excel (n° 14) : Excel l'affiche selon la
# langue du poste (jj/mm/aaaa en français) et le classe bien en « Date »,
# pas en « Standard » ni en « Personnalisé ».
FORMAT_DATE_EXCEL = "mm-dd-yy"


def _largeur_cellule(feuille, numero_ligne, numero_colonne):
    """Largeur (en caractères Excel) d'une cellule, fusion comprise."""
    debut = numero_colonne
    fin = numero_colonne
    for fusion in feuille.merged_cells.ranges:
        if fusion.min_row <= numero_ligne <= fusion.max_row and fusion.min_col == numero_colonne:
            fin = fusion.max_col
            break
    largeur = 0
    for colonne in range(debut, fin + 1):
        lettre = get_column_letter(colonne)
        largeur = largeur + (feuille.column_dimensions[lettre].width or 8.43)
    return largeur


def _ajuster_hauteur(feuille, numero_ligne, colonnes, hauteur_mini):
    """Agrandit une ligne de données pour que tout son texte soit visible :
    pour chaque cellule remplie, on compte les lignes nécessaires (retours à
    la ligne + coupure des mots à la largeur de la colonne), on garde le
    maximum et on en déduit la hauteur. Jamais moins que 'hauteur_mini'.
    Le retour à la ligne automatique est activé sur ces cellules."""
    from copy import copy as copier_style
    nb_lignes_max = 1
    for numero_colonne in colonnes:
        cellule = feuille.cell(numero_ligne, numero_colonne)
        if cellule.value is None or str(cellule.value).strip() == "":
            continue
        # En Arial 9, une unité de largeur Excel tient environ 1,1 caractère ;
        # on retire 1 caractère de marge pour rester du côté sûr.
        caracteres_par_ligne = max(1, int(_largeur_cellule(feuille, numero_ligne, numero_colonne) * 1.1) - 1)
        nb_lignes = 0
        for paragraphe in str(cellule.value).split("\n"):
            morceaux = textwrap.wrap(paragraphe, caracteres_par_ligne, break_long_words=True)
            nb_lignes = nb_lignes + max(1, len(morceaux))
        if nb_lignes > nb_lignes_max:
            nb_lignes_max = nb_lignes
        alignement = copier_style(cellule.alignment)
        alignement.wrap_text = True
        cellule.alignment = alignement
    hauteur = nb_lignes_max * HAUTEUR_LIGNE_TEXTE + MARGE_CELLULE
    feuille.row_dimensions[numero_ligne].height = max(hauteur_mini, hauteur)


# ---------- Habillage du bordereau (style seulement) ----------
# Le gabarit officiel fixe la STRUCTURE (onglets, cellules, fusions, textes,
# zone d'impression) : rien de cela n'est modifié ici. On ne change que
# l'apparence : polices, couleurs, fonds, bordures et hauteurs de ligne, aux
# couleurs de Solares Bauen, avec les mêmes couleurs d'avis que l'application.
ANTHRACITE = "2B2B2B"          # bandeau des titres de colonnes
GRIS_TEXTE = "5C5C5C"          # libellés et textes secondaires
GRIS_TRAIT = "D0D0D0"          # traits du tableau
GRIS_ZEBRE = "F7F7F7"          # une ligne sur deux
GRIS_BANDEAU = "F2F2F2"        # fond du bloc « lot / n° de visa / indice »
ROSE_SB = "FDECEE"             # fond de la ligne « + toute autre fiche technique »
POLICE_BORDEREAU = "Arial"
# Libellés de l'en-tête (le texte du gabarit reste tel quel ; seul son style change).
LIBELLES_ENTETE = ("Votre / vos interlocuteur", "Destinataire", "Maître d'Ouvrage",
                   "A l'attention de", "Date d'édition")


def _police(cellule, taille, gras=False, italique=False, couleur="1E1E1E"):
    """Remplace la police d'une cellule (Arial de la charte)."""
    cellule.font = Font(POLICE_BORDEREAU, taille, bold=gras, italic=italique, color=couleur)


def _alignement(cellule, horizontal=None, vertical="center", retour=None):
    """Change l'alignement en gardant le retour à la ligne existant (sauf si 'retour' est donné)."""
    ancien = cellule.alignment
    cellule.alignment = Alignment(horizontal=horizontal or ancien.horizontal, vertical=vertical,
                                  wrap_text=ancien.wrap_text if retour is None else retour)


def _cadre(cellule, couleur=GRIS_TRAIT, cotes="lrtb"):
    """Traits fins de la couleur donnée sur les côtés demandés (l, r, t, b) ; aucun ailleurs."""
    trait = Side(style="thin", color=couleur)
    vide = Side(style=None)
    cellule.border = Border(left=trait if "l" in cotes else vide, right=trait if "r" in cotes else vide,
                            top=trait if "t" in cotes else vide, bottom=trait if "b" in cotes else vide)


def _recolorer_cadre(cellule, couleur=GRIS_TRAIT):
    """Garde les traits du gabarit (mêmes côtés, même épaisseur) mais en gris clair."""
    bord = cellule.border

    def recolore(cote):
        return Side(style=cote.style, color=couleur) if cote.style else Side(style=None)
    cellule.border = Border(left=recolore(bord.left), right=recolore(bord.right),
                            top=recolore(bord.top), bottom=recolore(bord.bottom))


def _fond(cellule, couleur):
    cellule.fill = PatternFill("solid", fgColor=couleur, bgColor=couleur)


def _couleurs_avis(feuille, colonne_avis, premiere, derniere):
    """Mise en forme conditionnelle de la colonne Avis (remplace celle du gabarit) :
    fond et texte de chaque avis, les mêmes que dans l'application. Elle reste
    active si l'avis est modifié à la main dans Excel."""
    lettre = get_column_letter(colonne_avis)
    plage = f"{lettre}{premiere}:{lettre}{derniere}"
    feuille.conditional_formatting = ConditionalFormattingList()
    for avis in AVIS:
        style = DifferentialStyle(font=Font(bold=True, color=AVIS_FONT[avis]),
                                  fill=PatternFill("solid", fgColor=AVIS_FILL[avis], bgColor=AVIS_FILL[avis]))
        regle = Rule(type="expression", dxf=style, stopIfTrue=True)
        regle.formula = [f'TRIM({lettre}{premiere})="{avis}"']
        feuille.conditional_formatting.add(plage, regle)


def _habiller_bordereau(feuille, derniere_colonne, lignes_titres, premiere_donnee, derniere_donnee,
                        colonne_avis, colonnes_dates, colonnes_centrees=()):
    """Habille un onglet du bordereau rempli. 'lignes_titres' = (première, dernière)
    ligne des titres de colonnes ; données de 'premiere_donnee' à 'derniere_donnee'
    (lignes vides du gabarit comprises)."""
    colonnes = range(1, derniere_colonne + 1)
    feuille.sheet_view.showGridLines = False

    # 1. En-tête (lignes 6 à 15) : libellés en petit rouge foncé, valeurs en gris foncé.
    for numero_ligne in range(6, 16):
        for cellule in feuille[numero_ligne][:derniere_colonne]:
            if cellule.value is None or (isinstance(cellule.value, str) and cellule.value.startswith("=")
                                          and cellule.data_type == "f"):
                continue
            texte = str(cellule.value)
            if texte.startswith(LIBELLES_ENTETE):
                _police(cellule, 8.5, gras=True, couleur=ROUGE_SB_FONCE)
            else:
                _police(cellule, 10, gras=bool(cellule.font.b))
    # La date d'édition (formule =TODAY()) suit le style des valeurs.
    for numero_ligne in range(6, 16):
        for cellule in feuille[numero_ligne][:derniere_colonne]:
            if cellule.data_type == "f":
                _police(cellule, 10)

    # 2. Bloc titre (lignes 17 à 20) : bandeau rouge puis bloc gris clair.
    debut_titre = 17
    styles_titre = [(12, True, "FFFFFF", ROUGE_SB, 24),           # BORDEREAU DE VISA...
                    (11, True, "1E1E1E", GRIS_BANDEAU, 19),       # Lot
                    (14, True, ROUGE_SB_FONCE, GRIS_BANDEAU, 22),  # VISA N°xx
                    (9, False, GRIS_TEXTE, GRIS_BANDEAU, 16)]     # Indice
    for decalage, (taille, gras, couleur, fond, hauteur) in enumerate(styles_titre):
        numero_ligne = debut_titre + decalage
        for numero_colonne in colonnes:
            cellule = feuille.cell(numero_ligne, numero_colonne)
            _fond(cellule, fond)
            cellule.border = Border()
        tete = feuille.cell(numero_ligne, 1)
        _police(tete, taille, gras=gras, couleur=couleur)
        _alignement(tete, "center")
        feuille.row_dimensions[numero_ligne].height = hauteur

    # 3. Phrase d'introduction (ligne 21) : sans cadre, en italique.
    for numero_colonne in colonnes:
        feuille.cell(21, numero_colonne).border = Border()
    _police(feuille.cell(21, 1), 9.5, italique=True, couleur=GRIS_TEXTE)
    _alignement(feuille.cell(21, 1), "left", "bottom")
    feuille.row_dimensions[21].height = 22

    # 4. Titres de colonnes : bandeau anthracite, texte blanc, traits blancs.
    premiere_titre, derniere_titre = lignes_titres
    for numero_ligne in range(premiere_titre, derniere_titre + 1):
        for numero_colonne in colonnes:
            cellule = feuille.cell(numero_ligne, numero_colonne)
            _fond(cellule, ANTHRACITE)
            _cadre(cellule, "FFFFFF")
            _police(cellule, 9 if numero_ligne == premiere_titre else 8.5, gras=True, couleur="FFFFFF")
            _alignement(cellule, "center", "center", retour=True)

    # 5. Lignes de données : Arial 9, traits gris clair, une ligne sur deux grisée.
    for rang, numero_ligne in enumerate(range(premiere_donnee, derniere_donnee + 1)):
        for numero_colonne in colonnes:
            cellule = feuille.cell(numero_ligne, numero_colonne)
            _cadre(cellule)
            if rang % 2 == 1:
                _fond(cellule, GRIS_ZEBRE)
            else:
                cellule.fill = PatternFill(fill_type=None)
            if numero_colonne == colonne_avis:
                _police(cellule, 9, gras=True)
                _alignement(cellule, "center")
            elif numero_colonne == 1:
                _police(cellule, 9, gras=True)
                _alignement(cellule, "left")
            elif numero_colonne in colonnes_dates or numero_colonne in colonnes_centrees:
                _police(cellule, 9, couleur="3D3D3D")
                _alignement(cellule, "center")
            else:
                _police(cellule, 9)
                _alignement(cellule, "left")
    _couleurs_avis(feuille, colonne_avis, premiere_donnee, derniere_donnee)

    # 6. Sous le tableau : « + toute autre fiche », légende et notes.
    for numero_ligne in range(derniere_donnee + 1, feuille.max_row + 1):
        tete = feuille.cell(numero_ligne, 1)
        texte = str(tete.value or "").strip()
        for numero_colonne in colonnes:
            _recolorer_cadre(feuille.cell(numero_ligne, numero_colonne))
        if texte.startswith("+ TOUTE AUTRE"):
            for numero_colonne in colonnes:
                _fond(feuille.cell(numero_ligne, numero_colonne), ROSE_SB)
            _police(tete, 9, gras=True, italique=True, couleur=ROUGE_SB_FONCE)
            _alignement(tete, "left")
        elif texte.startswith("*L"):
            _police(tete, 9, gras=True, couleur=ROUGE_SB_FONCE)
        elif texte in AVIS:
            _fond(tete, AVIS_FILL[texte])
            _police(tete, 9, gras=True, couleur=AVIS_FONT[texte])
            _alignement(tete, "center")
            _police(feuille.cell(numero_ligne, 2), 8.5, couleur="3D3D3D")
        elif texte.startswith("NB"):
            _police(tete, 8, italique=True, couleur=GRIS_TEXTE)
        else:
            # « Présent VISA » / « VISA du jour » (case de couleur du gabarit conservée).
            voisine = feuille.cell(numero_ligne, 2)
            if voisine.value:
                _police(voisine, 8.5, gras=True, couleur="3D3D3D")



def _legende_sans_coupure(feuille, derniere_colonne, lignes_titres):
    """La légende et les notes du bas ne doivent jamais être coupées entre deux
    pages. On estime la pagination (page A4 paysage, marges et mise à l'échelle
    « 1 page en largeur », titres répétés) :
      - si la légende tient sur la page où elle commence, elle suit le tableau ;
      - sinon un saut de page est placé juste avant elle.
    Le saut manuel du gabarit avant la légende est remplacé par ce calcul."""
    debut_legende = None
    for numero_ligne in range(1, feuille.max_row + 1):
        if str(feuille.cell(numero_ligne, 1).value or "").startswith("*L"):
            debut_legende = numero_ligne
            break
    if debut_legende is None:
        return
    fin_legende = debut_legende
    for numero_ligne in range(debut_legende, feuille.max_row + 1):
        if feuille.cell(numero_ligne, 1).value not in (None, ""):
            fin_legende = numero_ligne

    # Géométrie de la page (en points) et échelle d'impression.
    marges = feuille.page_margins
    hauteur_utile = 595.3 - (marges.top + marges.bottom) * 72
    largeur_utile = 841.9 - (marges.left + marges.right) * 72
    largeur_feuille = 0
    for numero_colonne in range(1, derniere_colonne + 1):
        largeur = feuille.column_dimensions[get_column_letter(numero_colonne)].width or 8.43
        largeur_feuille = largeur_feuille + (largeur * 7 + 5) * 0.75
    echelle = min(1.0, largeur_utile / largeur_feuille)
    capacite = hauteur_utile / echelle * 0.94          # 6 % de marge de sécurité

    def hauteur(numero_ligne):
        dimension = feuille.row_dimensions[numero_ligne]
        if dimension.hidden:
            return 0
        return dimension.height or feuille.sheet_format.defaultRowHeight or 15

    premiere_titre, derniere_titre = lignes_titres
    hauteur_titres = sum(hauteur(n) for n in range(premiere_titre, derniere_titre + 1))
    sauts = [saut.id for saut in feuille.row_breaks.brk if saut.id < debut_legende - 2]
    page_de_la_ligne = {}
    page = 1
    hauteur_page = 0
    for numero_ligne in range(1, fin_legende + 1):
        if hauteur_page + hauteur(numero_ligne) > capacite and hauteur_page > 0:
            page = page + 1
            hauteur_page = hauteur_titres if numero_ligne > derniere_titre else 0
        hauteur_page = hauteur_page + hauteur(numero_ligne)
        page_de_la_ligne[numero_ligne] = page
        if numero_ligne in sauts:
            page = page + 1
            hauteur_page = hauteur_titres
    if page_de_la_ligne[debut_legende] != page_de_la_ligne[fin_legende]:
        sauts.append(debut_legende - 1)
    feuille.row_breaks = RowBreak()
    for numero in sorted(set(sauts)):
        feuille.row_breaks.append(Break(id=numero))


# En-tête de l'onglet « VISA PLANS - NDC » : (cellule de l'onglet plans,
# cellule de l'onglet matériel dont elle reprend la valeur).
EN_TETE_PLANS_DEPUIS_MATERIEL = (
    ("A7", "A7"), ("A8", "A8"), ("I7", "E7"), ("I8", "E8"), ("I9", "E9"),
    ("A12", "A12"), ("I12", "E12"), ("A13", "A13"), ("A14", "A14"), ("A15", "A15"),
    ("A18", "A18"),
)
# Marge du haut des pages du bordereau, en pouces (1,5 cm) : sans elle, les
# pages 2 et suivantes commençaient au ras du papier.
MARGE_HAUTE_BORDEREAU = 0.6
# Pied de page de gauche : « solares » en gras rouge, « bauen » puis le nom du
# fichier. Sur UNE ligne : le retour à la ligne du gabarit sortait en
# « _x000a_ » dans le PDF.
PIED_DE_PAGE_GAUCHE = '&"Arial,Gras"&KE60116solares&"Arial,Normal"&K000000 bauen  -  &F'


def _mise_en_page_bordereau(feuille, lignes_titres):
    """Marge haute, titres du tableau répétés en haut de chaque page et pied de page."""
    feuille.page_margins.top = MARGE_HAUTE_BORDEREAU
    feuille.print_title_rows = lignes_titres
    pied = feuille.oddFooter.left
    pied.text = PIED_DE_PAGE_GAUCHE
    pied.font = None
    pied.color = None
    pied.size = None


def _cellule_texte(cellule, valeur):
    """Écrit une valeur dans une cellule Excel et renvoie la cellule. Un texte
    qui commence par « = » reste du TEXTE : sinon Excel l'exécuterait comme une
    formule (un nom ou une pièce saisis sur le portail ne doivent jamais en devenir une)."""
    cellule.value = valeur
    if isinstance(valeur, str) and valeur.startswith("="):
        cellule.data_type = "s"
    return cellule


def texte_sur_pour_tableur(valeur):
    """Pour un export CSV : un texte qui commence par = + - @ (ou une
    tabulation) est précédé d'une apostrophe, pour qu'un tableur ne le prenne
    pas pour une formule. Les autres valeurs sont rendues telles quelles."""
    if isinstance(valeur, str) and valeur[:1] in ("=", "+", "-", "@", "\t", "\r"):
        return "'" + valeur
    return valeur


def export_visa(projet, lot, positions, gabarit, signataire=None):
    """Remplit le bordereau officiel "Visa type" (gabarit Excel fourni) :
      - en-tête : interlocuteur SB, destinataire (titulaire), MOA et lot ;
      - onglet "VISA MATERIEL"   : positions AVEC référence CCTP ;
      - onglet "VISA PLANS - NDC": positions SANS référence (plans, NdC...).
    Seules les positions cochées "Export" sont incluses. Un onglet sans
    aucune ligne est retiré (sauf si les deux sont vides : on garde le matériel).
    'signataire' = {nom, email, telephone} (par défaut : chef de projet).
    Renvoie le fichier Excel en octets."""
    if signataire is None:
        signataire = interlocuteur(projet)

    # Petit raccourci lisible : renvoie la valeur d'une clé, ou "" si absente.
    def lot_val(cle):
        return lot.get(cle) or ""

    from copy import copy as copier_style
    classeur = load_workbook(io.BytesIO(gabarit))

    # ------ En-tête de l'onglet MATÉRIEL (l'onglet Plans le recopie par formules) ------
    feuille = classeur["VISA MATERIEL"]
    # Votre interlocuteur : "Nom - email" puis le téléphone sur la ligne suivante.
    ligne_interlocuteur = signataire.get("nom") or ""
    if signataire.get("email"):
        ligne_interlocuteur = (ligne_interlocuteur + " - " + signataire["email"]).strip(" -")
    _cellule_texte(feuille["A7"], ligne_interlocuteur)
    _cellule_texte(feuille["A8"], ("Tél. " + signataire["telephone"]) if signataire.get("telephone") else None)
    # Destinataire : entreprise titulaire et son adresse (2 lignes).
    adresse_1, adresse_2 = lignes_adresse(lot_val("contact_adr"))
    _cellule_texte(feuille["E7"], lot_val("entreprise"))
    _cellule_texte(feuille["E8"], adresse_1)
    _cellule_texte(feuille["E9"], adresse_2)
    _cellule_texte(feuille["E12"], (lot_val("contact_civ") + " " + lot_val("contact_nom")).strip())
    # Maître d'ouvrage / opération : nom du MOA, désignation du projet, adresse du MOA.
    moa_1, moa_2 = lignes_adresse(projet.get("adresse"))
    _cellule_texte(feuille["A12"], projet.get("moa") or "")
    _cellule_texte(feuille["A13"], projet.get("nom") or "")
    _cellule_texte(feuille["A14"], moa_1)
    _cellule_texte(feuille["A15"], moa_2)
    _cellule_texte(feuille["A18"], libelle_lot(lot))
    _cellule_texte(feuille["A19"], lot_val("visa_no"))
    _cellule_texte(feuille["A20"], "Indice de révision : " + lot_val("indice"))

    # ------ Tri des positions à exporter, par onglet ------
    positions_materiel = []
    positions_plans = []
    for position in positions:
        if not position.get("export", 1):
            continue  # Décochée : pas dans le bordereau.
        if (position.get("pos") or "").strip():
            positions_materiel.append(position)  # A une référence CCTP.
        else:
            positions_plans.append(position)     # Sans référence -> Plans/NdC.

    # ------ Onglet MATÉRIEL : 7 lignes de données dans le gabarit (24 à 30) ------
    _ensure_rows(feuille, 24, 7, len(positions_materiel), 9, 25)
    numero_ligne = 24
    for position in positions_materiel:
        _cellule_texte(feuille.cell(numero_ligne, 1), position["pos"])
        _cellule_texte(feuille.cell(numero_ligne, 2), position["des"])
        _cellule_texte(feuille.cell(numero_ligne, 3), _date(last_recu(position["att"])))
        _cellule_texte(feuille.cell(numero_ligne, 4), position.get("marque") or "")
        _cellule_texte(feuille.cell(numero_ligne, 5), position.get("typeprop") or "")
        _cellule_texte(feuille.cell(numero_ligne, 6), position["avis"])
        _cellule_texte(feuille.cell(numero_ligne, 7), compose_attendus(position["att"]))
        _cellule_texte(feuille.cell(numero_ligne, 8), position.get("rem") or "")
        _cellule_texte(feuille.cell(numero_ligne, 9), _date(position.get("datevisa")))
        # Colonnes C (réception) et I (visa) : toujours au format Date.
        for numero_colonne in (3, 9):
            feuille.cell(numero_ligne, numero_colonne).number_format = FORMAT_DATE_EXCEL
        _ajuster_hauteur(feuille, numero_ligne, range(1, 10), HAUTEUR_MINI_BORDEREAU)
        numero_ligne = numero_ligne + 1
    # Lignes du gabarit non utilisées : vidées et masquées (elles restent dans le fichier).
    for numero_ligne in range(24 + len(positions_materiel), 24 + 7):
        for numero_colonne in (1, 2, 3, 4, 5, 6, 7, 8, 9):
            feuille.cell(numero_ligne, numero_colonne).value = None
        feuille.row_dimensions[numero_ligne].hidden = True

    # ------ Onglet PLANS / NDC : 8 lignes de données (23 à 30) ------
    feuille_plans = classeur["VISA PLANS - NDC"]
    _cellule_texte(feuille_plans["A19"], lot_val("visa_no"))
    _cellule_texte(feuille_plans["A20"], "Indice de révision : " + lot_val("indice"))
    _ensure_rows(feuille_plans, 23, 8, len(positions_plans), 16, 24)
    numero_ligne = 23
    for position in positions_plans:
        date_visa = _date(position.get("datevisa"))
        _cellule_texte(feuille_plans.cell(numero_ligne, 1), position["des"])
        _cellule_texte(feuille_plans.cell(numero_ligne, 6), _date(last_recu(position["att"])))
        # Le gabarit possède DEUX colonnes "Date du Visa" (G et N) : on remplit les deux.
        _cellule_texte(feuille_plans.cell(numero_ligne, 7), date_visa)
        _cellule_texte(feuille_plans.cell(numero_ligne, 14), date_visa)
        _cellule_texte(feuille_plans.cell(numero_ligne, 8), position.get("intitule") or "")
        _cellule_texte(feuille_plans.cell(numero_ligne, 10), position["avis"])
        _cellule_texte(feuille_plans.cell(numero_ligne, 11), position.get("rem") or "")
        # Colonnes F (réception), G et N (visa) : toujours au format Date.
        for numero_colonne in (6, 7, 14):
            feuille_plans.cell(numero_ligne, numero_colonne).number_format = FORMAT_DATE_EXCEL
        _ajuster_hauteur(feuille_plans, numero_ligne, (1, 6, 7, 8, 10, 11, 14), HAUTEUR_MINI_BORDEREAU)
        numero_ligne = numero_ligne + 1
    for numero_ligne in range(23 + len(positions_plans), 23 + 8):
        for numero_colonne in (1, 4, 5, 6, 7, 8, 10, 11, 14):
            feuille_plans.cell(numero_ligne, numero_colonne).value = None
        feuille_plans.row_dimensions[numero_ligne].hidden = True

    # ------ En-tête de l'onglet PLANS : recopié en valeurs, en noir ------
    # Le gabarit le recopiait par des formules (écrites en rouge) vers l'onglet
    # matériel : elles ne marcheraient plus si cet onglet est retiré.
    for cellule_plans, cellule_materiel in EN_TETE_PLANS_DEPUIS_MATERIEL:
        cellule = feuille_plans[cellule_plans]
        _cellule_texte(cellule, feuille[cellule_materiel].value)
        police = copier_style(cellule.font)
        police.color = None   # Couleur automatique (noir).
        cellule.font = police

    # ------ Mise en page et habillage des deux onglets ------
    for onglet, lignes_titres in ((feuille, "22:23"), (feuille_plans, "22:22")):
        _mise_en_page_bordereau(onglet, lignes_titres)
    # Colonnes « Date du Visa » un peu plus larges : le titre n'y tenait pas.
    feuille.column_dimensions["I"].width = 11.6
    feuille_plans.column_dimensions["N"].width = 11.6
    # Lignes de données : celles remplies, et au moins celles prévues par le gabarit.
    _habiller_bordereau(feuille, 9, (22, 23), 24, 24 + max(len(positions_materiel), 7) - 1,
                        colonne_avis=6, colonnes_dates=(3, 9), colonnes_centrees=(4, 5))
    _habiller_bordereau(feuille_plans, 14, (22, 22), 23, 23 + max(len(positions_plans), 8) - 1,
                        colonne_avis=10, colonnes_dates=(5, 6, 7, 14), colonnes_centrees=(4,))
    _legende_sans_coupure(feuille, 9, (22, 23))
    _legende_sans_coupure(feuille_plans, 14, (22, 22))

    # ------ Onglets inutiles retirés (pas de page vide dans le PDF) ------
    if not positions_plans:
        classeur.remove(feuille_plans)
    elif not positions_materiel:
        classeur.remove(feuille)
    classeur.active = 0

    # ------ Sauvegarde en mémoire ------
    tampon = io.BytesIO()
    classeur.save(tampon)
    return tampon.getvalue()


def chemin_libreoffice():
    """Chemin du programme LibreOffice (soffice), ou None s'il n'est pas installé."""
    for nom in ("soffice", "libreoffice"):
        chemin = shutil.which(nom)
        if chemin:
            return chemin
    for chemin in (r"C:\Program Files\LibreOffice\program\soffice.exe",
                   r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
                   "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if os.path.exists(chemin):
            return chemin
    return None


def convertir_en_pdf(contenu_xlsx, nom_fichier="bordereau"):
    """Convertit un classeur Excel (octets) en PDF avec LibreOffice, pour que
    le PDF soit l'image exacte du bordereau Excel.
    'nom_fichier' (sans extension) : nom donné au fichier de travail, repris
    par le pied de page (« &F »).
    Renvoie (octets du PDF, "") ou (None, message d'explication)."""
    nom_fichier = re.sub(r'[\\/:*?"<>|]', "_", nom_fichier) or "bordereau"
    programme = chemin_libreoffice()
    if programme is None:
        return None, ("LibreOffice n'est pas installé sur ce poste : seule la version "
                      "Excel du bordereau peut être générée.")
    with tempfile.TemporaryDirectory() as dossier:
        source = os.path.join(dossier, nom_fichier + ".xlsx")
        with open(source, "wb") as flux:
            flux.write(contenu_xlsx)
        # Profil LibreOffice dédié : évite les conflits avec une session ouverte.
        profil = "file:///" + os.path.join(dossier, "profil").replace("\\", "/")
        try:
            subprocess.run(
                [programme, f"-env:UserInstallation={profil}", "--headless",
                 "--convert-to", "pdf", "--outdir", dossier, source],
                capture_output=True, timeout=120, check=False)
        except (OSError, subprocess.TimeoutExpired) as erreur:
            return None, f"Conversion PDF impossible ({erreur})."
        sortie = os.path.join(dossier, nom_fichier + ".pdf")
        if not os.path.exists(sortie):
            return None, "Conversion PDF impossible : LibreOffice n'a produit aucun fichier."
        with open(sortie, "rb") as flux:
            return flux.read(), ""


# Couleurs de l'export de statut : fond et texte, alignés sur les pastilles de l'appli.
AVIS_FILL = {"VSO": "D9EAD3", "VAO": "FCEFC7", "VAOB": "FBDDC6",
             "REF": "F8CFCF", "MQ": "E3DDEE", "SO": "E6E6E6"}
AVIS_FONT = {"VSO": "1E5E2E", "VAO": "7A5A00", "VAOB": "9A3D0A",
             "REF": "A50010", "MQ": "4B2E83", "SO": "4D4D4D"}
ROUGE_SB = "E60116"
ROUGE_SB_FONCE = "C00000"


def export_statut_visas(projet, lot, positions):
    """Génère le tableau « Statut des visas » du lot : une ligne par position
    avec avis coloré, pièces attendues, dates de réception et de visa,
    remarques et échéance de visa côté Solares Bauen.
    Renvoie le fichier Excel en octets."""
    rappels_par_position = {}
    for rappel in get_rappels_sb(lot_id=lot["id"]):
        rappels_par_position[rappel["position_id"]] = rappel

    classeur = Workbook()
    feuille = classeur.active
    feuille.title = "Statut des visas"
    bordure_fine = Side(style="thin", color="BFBFBF")
    cadre = Border(left=bordure_fine, right=bordure_fine, top=bordure_fine, bottom=bordure_fine)

    # --- Titre ---
    feuille.merge_cells("A1:J1")
    cellule_titre = feuille.cell(
        1, 1, f"STATUT DES VISAS — {projet['nom']} — {libelle_lot(lot)} "
              f"({lot.get('entreprise') or ''})")
    cellule_titre.font = Font("Arial", 12, bold=True, color="FFFFFF")
    cellule_titre.fill = PatternFill("solid", fgColor=ROUGE_SB)
    cellule_titre.alignment = Alignment("left", "center", indent=1)
    feuille.row_dimensions[1].height = 22
    feuille.merge_cells("A2:J2")
    cellule_soustitre = feuille.cell(
        2, 1, f"Solares Bauen — édité le {dt.date.today().strftime('%d/%m/%Y')}")
    cellule_soustitre.font = Font("Arial", 9, italic=True, color="595959")

    # --- En-têtes de colonnes ---
    entetes = ["Pos. CCTP", "Désignation", "Marque", "Type", "Avis",
               "Pièces attendues / statut", "Date de réception", "Date du visa",
               "Remarques / observations", "À viser par SB avant le"]
    largeurs = [12, 40, 16, 18, 9, 42, 13, 13, 42, 15]
    for indice in range(len(entetes)):
        cellule = feuille.cell(4, indice + 1, entetes[indice])
        cellule.font = Font("Arial", 9, bold=True, color="FFFFFF")
        cellule.fill = PatternFill("solid", fgColor="3A3A3A")
        cellule.alignment = Alignment("center", "center", wrap_text=True)
        cellule.border = cadre
        feuille.column_dimensions[get_column_letter(indice + 1)].width = largeurs[indice]
    feuille.row_dimensions[4].height = 30

    # --- Lignes de données ---
    numero_ligne = 5
    for position in positions:
        reference = (position.get("pos") or "").strip() or "—"
        rappel = rappels_par_position.get(position["id"])
        valeurs = [reference, position["des"], position.get("marque") or "",
                   position.get("typeprop") or "", position["avis"],
                   compose_attendus(position["att"]),
                   _date(last_recu(position["att"])), _date(position.get("datevisa")),
                   position.get("rem") or "",
                   _date(rappel["echeance"]) if rappel else None]
        for indice in range(len(valeurs)):
            cellule = _cellule_texte(feuille.cell(numero_ligne, indice + 1), valeurs[indice])
            cellule.font = Font("Arial", 9)
            cellule.border = cadre
            cellule.alignment = Alignment("left", "center", wrap_text=True)
            if isinstance(valeurs[indice], dt.date):
                cellule.number_format = "DD/MM/YYYY"
        # Colonne Avis : fond coloré selon le code couleur de la plateforme.
        cellule_avis = feuille.cell(numero_ligne, 5)
        cellule_avis.fill = PatternFill("solid", fgColor=AVIS_FILL.get(position["avis"], "FFFFFF"))
        cellule_avis.font = Font("Arial", 9, bold=True, color=AVIS_FONT.get(position["avis"], "000000"))
        cellule_avis.alignment = Alignment("center", "center")
        # Échéance dépassée : en rouge.
        if rappel and rappel["niveau"] == "retard":
            feuille.cell(numero_ligne, 10).font = Font("Arial", 9, bold=True, color=ROUGE_SB_FONCE)
        numero_ligne = numero_ligne + 1

    # --- Confort d'utilisation : volets figés, filtres, impression paysage ---
    feuille.freeze_panes = "A5"
    feuille.auto_filter.ref = f"A4:J{max(numero_ligne - 1, 4)}"
    feuille.page_setup.orientation = "landscape"
    feuille.page_setup.paperSize = 9  # A4
    feuille.page_setup.fitToWidth = 1
    feuille.page_setup.fitToHeight = 0
    feuille.sheet_properties.pageSetUpPr.fitToPage = True
    feuille.sheet_view.showGridLines = False

    tampon = io.BytesIO()
    classeur.save(tampon)
    return tampon.getvalue()


# ============================================================================
#  14. REGISTRE DES BORDEREAUX ÉMIS
# ----------------------------------------------------------------------------
#  « Générer le bordereau de visa » ÉMET un bordereau : il reçoit le numéro
#  suivant du lot (VISA N°01, 02...) et une photo de son contenu (projet, lot,
#  interlocuteur, positions exportées avec leurs avis) est enregistrée. On
#  peut ensuite le réexporter à l'identique, changer son numéro ou le
#  supprimer. Le numéro suivant repart du dernier bordereau émis.
# ============================================================================

def format_numero(numero):
    """3 -> "VISA N°03" (ligne 19 du bordereau)."""
    return f"VISA N°{int(numero):02d}"


def prochain_numero(lot_id):
    """Numéro du prochain bordereau d'un lot : celui du dernier bordereau émis
    (chronologiquement) + 1, ou 1 s'il n'y en a aucun."""
    connexion = conn()
    ligne = connexion.execute(
        "SELECT numero FROM bordereaux WHERE lot_id=? ORDER BY date DESC, id DESC LIMIT 1",
        (lot_id,)).fetchone()
    connexion.close()
    if ligne is None:
        return 1
    return int(ligne["numero"]) + 1


def emettre_bordereau(projet, lot, positions, signataire=None):
    """Émet un bordereau : numéro suivant du lot + photo du contenu.
    Seules les positions cochées « Au bordereau » y figurent ; leurs rappels de visa
    SB sont clos (plus de bouton « Marquer comme visé »).
    Renvoie le bordereau (dictionnaire, voir get_bordereau)."""
    if signataire is None:
        signataire = interlocuteur(projet)
    exportees = []
    for position in positions:
        if position.get("export", 1):
            exportees.append(position)
    photo = {"projet": projet, "lot": lot, "signataire": signataire, "positions": exportees}
    numero = prochain_numero(lot["id"])
    connexion = conn()
    curseur = connexion.execute(
        "INSERT INTO bordereaux(projet, lot_id, numero, date, contenu) VALUES(?,?,?,?,?)",
        (projet["code"], lot["id"], numero, dt.datetime.now().isoformat(timespec="seconds"),
         json.dumps(photo, ensure_ascii=False)))
    bordereau_id = curseur.lastrowid
    # Le visa part chez l'entreprise (Plateforme visa partenaires) : les rappels « à viser par SB
    # avant le… » des positions du bordereau sont clos.
    for position in exportees:
        if position.get("id"):
            connexion.execute(
                "UPDATE rappels_sb SET cloture=1, date_cloture=? WHERE position_id=? AND cloture=0",
                (dt.date.today().isoformat(), position["id"]))
    connexion.commit()
    connexion.close()
    return get_bordereau(bordereau_id)


def _bordereau_depuis_ligne(ligne):
    bordereau = dict(ligne)
    bordereau["contenu"] = json.loads(bordereau["contenu"] or "{}")
    bordereau["libelle"] = format_numero(bordereau["numero"])
    return bordereau


def get_bordereau(bordereau_id):
    """Un bordereau émis : {id, projet, lot_id, numero, libelle, date, contenu}."""
    connexion = conn()
    ligne = connexion.execute("SELECT * FROM bordereaux WHERE id=?", (bordereau_id,)).fetchone()
    connexion.close()
    if ligne is None:
        return None
    return _bordereau_depuis_ligne(ligne)


def get_bordereaux(lot_id):
    """Bordereaux émis d'un lot, le plus récent en premier."""
    connexion = conn()
    lignes = connexion.execute(
        "SELECT * FROM bordereaux WHERE lot_id=? ORDER BY date DESC, id DESC", (lot_id,)).fetchall()
    connexion.close()
    resultat = []
    for ligne in lignes:
        resultat.append(_bordereau_depuis_ligne(ligne))
    return resultat


def set_numero_bordereau(bordereau_id, numero):
    """Change le numéro d'un bordereau (refusé s'il est déjà pris dans le lot).
    Renvoie (True, "") ou (False, message)."""
    try:
        numero = int(numero)
    except (TypeError, ValueError):
        return False, "Le numéro doit être un nombre entier."
    if numero < 1:
        return False, "Le numéro doit être supérieur ou égal à 1."
    connexion = conn()
    ligne = connexion.execute("SELECT lot_id FROM bordereaux WHERE id=?", (bordereau_id,)).fetchone()
    if ligne is None:
        connexion.close()
        return False, "Bordereau introuvable."
    if connexion.execute("SELECT 1 FROM bordereaux WHERE lot_id=? AND numero=? AND id<>?",
                         (ligne["lot_id"], numero, bordereau_id)).fetchone():
        connexion.close()
        return False, f"Le numéro {format_numero(numero)} est déjà utilisé dans ce lot."
    connexion.execute("UPDATE bordereaux SET numero=? WHERE id=?", (numero, bordereau_id))
    connexion.commit()
    connexion.close()
    return True, ""


def delete_bordereau(bordereau_id):
    """Supprime un bordereau du registre."""
    connexion = conn()
    connexion.execute("DELETE FROM bordereaux WHERE id=?", (bordereau_id,))
    connexion.commit()
    connexion.close()


def export_bordereau(bordereau, gabarit):
    """Réexporte un bordereau émis à l'identique (Excel), avec son numéro
    actuel en ligne 19. Renvoie le fichier Excel en octets."""
    contenu = bordereau["contenu"]
    lot = dict(contenu["lot"])
    lot["visa_no"] = format_numero(bordereau["numero"])
    return export_visa(contenu["projet"], lot, contenu["positions"], gabarit, contenu["signataire"])


# ============================================================================
#  15. BASE DE DONNÉES TRANSVERSALE (statistiques)
# ============================================================================

def annee_projet(code_projet):
    """Année d'un projet, lue dans les 4 premiers caractères de son code.
    Exemples : "2026008_GS_KUNHEIM" -> "2026" ; "2023022 39LGTS" -> "2023" ;
    "CPAM67-02" -> "" (le code ne commence pas par 19xx ou 20xx : pas d'année)."""
    debut = (code_projet or "").strip()[:4]
    if len(debut) == 4 and debut.isdigit() and debut[:2] in ("19", "20"):
        return debut
    return ""


def contient_texte(texte, recherche):
    """Vrai si `recherche` apparaît dans `texte`, sans tenir compte des
    majuscules ni des accents (« regulation » trouve « Régulation GTB »).
    Une recherche vide est toujours vraie."""
    recherche_normalisee = _norm(recherche).strip()
    if recherche_normalisee == "":
        return True
    return recherche_normalisee in _norm(texte)


def db_visas():
    """Construit la vue transversale de TOUS les visas CCTP, tous projets et
    chefs confondus, pour la page "Base de données" et les statistiques.
    Périmètre :
      - positions AVEC référence CCTP uniquement (plans/NdC exclus) ;
      - dernier visa de chaque position (les copies "Anciens visas" exclues) ;
      - les projets archivés sont inclus (filtrables dans l'interface).
    Colonnes calculées :
      - "Année"              : 4 premiers chiffres du code projet (annee_projet) ;
    Marque / type « prescrits » = proposition du maître d'oeuvre (import DPGF),
    « proposés » = proposition de l'entreprise (celle du bordereau).
      - "Nb relances"        : occurrences de la position dans le journal ;
      - "Nb itérations visa" : nombre de copies archivées (cycles de visa)."""
    connexion = conn()

    # --- Dictionnaires de correspondance (une seule lecture par table) ---
    chefs_par_code = {}
    for ligne in connexion.execute("SELECT * FROM chefs").fetchall():
        chefs_par_code[ligne["code"]] = ligne["nom"]
    projets_par_code = {}
    for ligne in connexion.execute("SELECT * FROM projets").fetchall():
        projets_par_code[ligne["code"]] = dict(ligne)
    lots_par_id = {}
    for ligne in connexion.execute("SELECT * FROM lots").fetchall():
        lots_par_id[ligne["id"]] = dict(ligne)

    # --- Comptage des relances par (lot, référence de position) ---
    relances_par_position = {}
    for ligne in connexion.execute("SELECT lot_id, positions FROM relances").fetchall():
        try:
            references = json.loads(ligne["positions"] or "[]")
        except (ValueError, TypeError):
            references = []
        for reference in references:
            cle = (ligne["lot_id"], str(reference))
            relances_par_position[cle] = relances_par_position.get(cle, 0) + 1

    # --- Comptage des itérations de visa (copies archivées par position) ---
    iterations_par_position = {}
    for ligne in connexion.execute(
            "SELECT lot_id, pos FROM positions WHERE IFNULL(archive,0)=1").fetchall():
        cle = (ligne["lot_id"], (ligne["pos"] or "").strip())
        iterations_par_position[cle] = iterations_par_position.get(cle, 0) + 1

    # --- Construction des lignes du tableau (une par position active CCTP) ---
    resultat = []
    positions_actives = connexion.execute(
        "SELECT * FROM positions WHERE IFNULL(archive,0)=0 ORDER BY projet, ordre, id").fetchall()
    for position in positions_actives:
        reference = (position["pos"] or "").strip()
        if reference == "":
            continue  # Plans / NdC : hors périmètre de la base.
        projet = projets_par_code.get(position["projet"], {})
        lot = lots_par_id.get(position["lot_id"], {})
        if projet.get("archive"):
            projet_archive = "Oui"
        else:
            projet_archive = "Non"
        resultat.append({
            "Année": annee_projet(position["projet"]),
            "Code projet": position["projet"],
            "Nom projet": projet.get("nom") or "",
            "Chef de projet": chefs_par_code.get(projet.get("chef")) or "",
            "Lot": libelle_lot(lot) if lot else "",
            "Entreprise": lot.get("entreprise") or "",
            "Intitulé position": position["des"] or "",
            "Trame": position["famille"] or "",
            "Marque prescrite": position["marque_moe"] or "",
            "Type prescrit": position["type_moe"] or "",
            "Marque proposée": position["marque"] or "",
            "Type proposé": position["typeprop"] or "",
            "Statut visa": position["avis"] or "",
            "Date visa": position["datevisa"] or "",
            "Nb relances": relances_par_position.get((position["lot_id"], reference), 0),
            "Nb itérations visa": iterations_par_position.get((position["lot_id"], reference), 0),
            "Commentaire": position["rem"] or "",
            "Projet archivé": projet_archive,
        })
    connexion.close()
    return resultat


def delais_de_visa():
    """Délais de visa de Solares Bauen, un élément par document visé, tous
    projets confondus (archivés compris), pour le Dashboard « Solares ».
    Mesure : du DÉPÔT du document (ouverture du rappel « à viser avant le »)
    au VISA (clôture du rappel, faite à l'émission du bordereau), en jours
    ouvrés d'Alsace-Moselle (un visa le jour du dépôt = 0 jour).
    Les rappels encore ouverts (documents pas encore visés) sont ignorés.
    Chaque élément : {projet, lot_id, position_id, chef (code), chef_nom,
    equipe (nom, "" sans équipe), date_depot, date_visa, mois ("2026-10"),
    echeance, delai (jours ouvrés), dans_le_delai (visé au plus tard à l'échéance)}."""
    connexion = conn()
    equipes_par_id = {}
    for ligne in connexion.execute("SELECT id, nom FROM teams").fetchall():
        equipes_par_id[ligne["id"]] = ligne["nom"]
    chefs_par_code = {}
    for ligne in connexion.execute("SELECT code, nom, team_id FROM chefs").fetchall():
        chefs_par_code[ligne["code"]] = dict(ligne)
    chef_par_projet = {}
    for ligne in connexion.execute("SELECT code, chef FROM projets").fetchall():
        chef_par_projet[ligne["code"]] = ligne["chef"]
    rappels = connexion.execute(
        """SELECT * FROM rappels_sb WHERE cloture=1 AND IFNULL(date_cloture,'')<>''
           ORDER BY date_cloture, id""").fetchall()
    connexion.close()

    resultat = []
    for rappel in rappels:
        try:
            date_depot = dt.date.fromisoformat(rappel["date_depot"][:10])
            date_visa = dt.date.fromisoformat(rappel["date_cloture"][:10])
            echeance = dt.date.fromisoformat(rappel["echeance"][:10])
        except (TypeError, ValueError):
            continue  # date illisible : document ignoré plutôt qu'un délai faux
        if date_visa < date_depot:
            delai = 0
        else:
            delai = calendrier.jours_ouvres_restants(date_depot, date_visa)
        code_chef = chef_par_projet.get(rappel["projet"])
        chef = chefs_par_code.get(code_chef, {})
        resultat.append({
            "projet": rappel["projet"], "lot_id": rappel["lot_id"], "position_id": rappel["position_id"],
            "chef": code_chef or "", "chef_nom": chef.get("nom") or "",
            "equipe": equipes_par_id.get(chef.get("team_id"), ""),
            "date_depot": date_depot.isoformat(), "date_visa": date_visa.isoformat(),
            "mois": date_visa.isoformat()[:7], "echeance": echeance.isoformat(),
            "delai": delai, "dans_le_delai": date_visa <= echeance,
        })
    return resultat


def statistiques_delais(delais):
    """Synthèse d'une liste de délais (voir delais_de_visa) :
    {nb, moyenne, mediane (jours ouvrés, None si vide), nb_dans_delai,
    pct_dans_delai (0-100, None si vide), nb_retard}."""
    valeurs = []
    nb_dans_delai = 0
    for element in delais:
        valeurs.append(element["delai"])
        if element["dans_le_delai"]:
            nb_dans_delai = nb_dans_delai + 1
    nb = len(valeurs)
    if nb == 0:
        return {"nb": 0, "moyenne": None, "mediane": None, "nb_dans_delai": 0,
                "pct_dans_delai": None, "nb_retard": 0}
    valeurs.sort()
    if nb % 2 == 1:
        mediane = valeurs[nb // 2]
    else:
        mediane = (valeurs[nb // 2 - 1] + valeurs[nb // 2]) / 2
    return {"nb": nb, "moyenne": sum(valeurs) / nb, "mediane": mediane,
            "nb_dans_delai": nb_dans_delai, "pct_dans_delai": 100 * nb_dans_delai / nb,
            "nb_retard": nb - nb_dans_delai}


# ============================================================================
#  16. PORTAIL ENTREPRISES : envois, transfert, refus, état miroir
# ----------------------------------------------------------------------------
#  Parcours d'un document déposé par une entreprise :
#   1. ENVOI (portail) : un dépôt + une ligne « transmis » par position. Rien ne
#      change encore dans les positions de la plateforme.
#   2. TRANSFERT (page Dépôts reçus, chef de projet) : si le visa actuel a été
#      émis, il passe dans « Anciens visas » ; les fichiers rejoignent la
#      position (extraits aux pages indiquées), les pièces passent
#      « À vérifier », le rappel SB démarre à la DATE D'ENVOI.
#      Ou REFUS avec un motif : brouillon de mail au chargé d'affaires.
#   3. L'émission d'un bordereau contenant la position = « Visé le … ».
#  L'état vu par l'entreprise (« miroir ») se déduit de ces trois éléments.
# ============================================================================

# Noms des deux plateformes (onglets du navigateur, textes, mails) : ceux de leurs logos.
NOM_PLATEFORME = "Plateforme visa interne"     # app.py : plateforme de Solares Bauen
NOM_PORTAIL = "Plateforme visa partenaires"       # portail.py : dépôt des documents à viser par les entreprises

# Fichiers refusés quel que soit le contexte (exécutables, scripts,
# raccourcis), y compris à l'intérieur d'un ZIP.
EXTENSIONS_INTERDITES = ("exe", "bat", "cmd", "com", "msi", "msp", "ps1", "psm1", "vbs", "vbe",
                         "js", "jse", "wsf", "wsh", "hta", "jar", "scr", "dll", "cpl", "pif",
                         "lnk", "url", "reg", "msc", "sh", "appx", "msix", "iso", "img", "vhd")
TAILLE_MAXI_FICHIER = 500 * 1024 * 1024   # 500 Mo par fichier (à confirmer)

# Pièce toujours proposée à l'entreprise en plus des pièces attendues ; elle
# précise son libellé, qui devient une nouvelle pièce de la position au transfert.
PIECE_AUTRE = "Autre"

# États d'une position vus par l'entreprise, et leur libellé à l'écran.
ETAT_A_FOURNIR = "afournir"
ETAT_TRANSMIS = "transmis"
ETAT_EXAMEN = "examen"
ETAT_VISE = "vise"
LIBELLES_ETAT = {ETAT_A_FOURNIR: "À fournir",
                 ETAT_TRANSMIS: "Document transmis à Solares Bauen",
                 ETAT_EXAMEN: "En cours d'examen",
                 ETAT_VISE: "Visé"}

# Statuts d'une ligne de dépôt.
LIGNE_TRANSMISE = "transmis"
LIGNE_TRANSFEREE = "transfere"
LIGNE_REFUSEE = "refuse"

# Soumission libre : document que l'entreprise dépose hors des lignes prévues,
# sur la page Positions CCTP/Matériel (référence + désignation) ou Plans / notes
# de calcul (désignation). La ligne n'a pas de position tant que Solares Bauen
# ne l'a pas transférée (nouvelle position, ou position existante du lot).
LIBRE_CCTP = "cctp"
LIBRE_PLAN = "plan"


def _maintenant():
    """Date et heure actuelles, au format ISO (ex. "2026-10-30T14:05:12")."""
    return dt.datetime.now().isoformat(timespec="seconds")


def _date_fr(texte_iso):
    """"2026-10-30" ou "2026-10-30T14:05:12" -> "30/10/2026"."""
    jour = _date((texte_iso or "")[:10])
    if jour is None:
        return ""
    return jour.strftime("%d/%m/%Y")


def extension_fichier(nom):
    """Extension en minuscules, sans le point : "Plan.PDF" -> "pdf".
    Les points et espaces finaux sont ignorés, comme le fait Windows
    (« outil.exe. » est un .exe)."""
    nom = (nom or "").strip().rstrip(". ")
    return os.path.splitext(nom)[1].lower().lstrip(".")


def _libelle_position(position):
    """"3.4.1 Régulation et GTB" (ou la désignation seule pour un plan)."""
    return f"{position.get('pos') or ''} {position.get('des') or ''}".strip()


# ---------- Pages d'un PDF ----------

def pages_depuis_texte(texte, nb_pages=None):
    """Lit une liste de pages saisie par l'entreprise.
    Exemples : "12-30" -> 12 à 30 ; "5, 8-9" -> 5, 8, 9 ; "p. 3" -> 3.
    nb_pages : nombre de pages du PDF (pour refuser une page qui n'existe pas).
    Renvoie (liste des pages, "") ou ([], message d'erreur)."""
    texte = (texte or "").strip()
    if texte.lower().startswith("p."):
        texte = texte[2:]
    pages = []
    for morceau in texte.replace(";", ",").split(","):
        morceau = morceau.strip()
        if morceau == "":
            continue
        if "-" in morceau:
            debut_texte, fin_texte = morceau.split("-", 1)
        else:
            debut_texte, fin_texte = morceau, morceau
        try:
            debut = int(debut_texte.strip())
            fin = int(fin_texte.strip())
        except ValueError:
            return [], f"« {morceau} » n'est pas une page ou une plage de pages (ex. 12-30)."
        if debut < 1 or fin < debut:
            return [], f"La plage « {morceau} » n'est pas valable."
        if nb_pages is not None and fin > nb_pages:
            return [], f"La page {fin} n'existe pas (le fichier a {nb_pages} pages)."
        for numero in range(debut, fin + 1):
            if numero not in pages:
                pages.append(numero)
    if not pages:
        return [], "Aucune page indiquée."
    return pages, ""


def nombre_pages_pdf(contenu):
    """Nombre de pages d'un PDF (octets), ou None s'il est illisible."""
    try:
        return len(PdfReader(io.BytesIO(contenu)).pages)
    except Exception:   # PDF abîmé ou protégé : on le signale sans planter.
        return None


def extraire_pages(contenu, texte_pages):
    """Construit un nouveau PDF avec seulement les pages indiquées.
    Renvoie (octets du PDF, "") ou (None, message d'erreur)."""
    try:
        lecteur = PdfReader(io.BytesIO(contenu))
        nb_pages = len(lecteur.pages)
    except Exception:
        return None, "Le PDF est illisible."
    pages, erreur = pages_depuis_texte(texte_pages, nb_pages)
    if erreur:
        return None, erreur
    redacteur = PdfWriter()
    for numero in pages:
        redacteur.add_page(lecteur.pages[numero - 1])
    sortie = io.BytesIO()
    redacteur.write(sortie)
    return sortie.getvalue(), ""


# ---------- État miroir d'une position ----------

def _dernier_bordereau_position(connexion, position_id, lot_id):
    """Dernier bordereau émis du lot qui contient cette position :
    {id, date, libelle, avis, observations}, ou None."""
    for ligne in connexion.execute(
            "SELECT * FROM bordereaux WHERE lot_id=? ORDER BY date DESC, id DESC",
            (lot_id,)).fetchall():
        photo = json.loads(ligne["contenu"] or "{}")
        for position in photo.get("positions", []):
            if position.get("id") == position_id:
                return {"id": ligne["id"], "date": ligne["date"],
                        "libelle": format_numero(ligne["numero"]),
                        "avis": position.get("avis") or "",
                        "observations": position.get("rem") or ""}
    return None


def _derniere_ligne_depot(connexion, position_id, statut=None):
    """Ligne de dépôt la plus récente d'une position (éventuellement d'un statut
    donné), avec la date d'envoi et le chargé d'affaires ; ou None."""
    requete = """SELECT dl.*, d.date_envoi, d.contact_nom, d.contact_email
                 FROM depot_lignes dl JOIN depots d ON d.id = dl.depot_id
                 WHERE dl.position_id=?"""
    valeurs = [position_id]
    if statut:
        requete = requete + " AND dl.statut=?"
        valeurs.append(statut)
    ligne = connexion.execute(requete + " ORDER BY d.date_envoi DESC, dl.id DESC LIMIT 1",
                              valeurs).fetchone()
    if ligne is None:
        return None
    return dict(ligne)


def _visa_emis(connexion, position_id, lot_id):
    """Le visa actuel de la position a-t-il été émis ? C'est le cas si un
    bordereau la contient ET qu'aucun transfert n'a eu lieu depuis.
    Renvoie le bordereau concerné (voir _dernier_bordereau_position) ou None."""
    bordereau = _dernier_bordereau_position(connexion, position_id, lot_id)
    if bordereau is None:
        return None
    transfert = _derniere_ligne_depot(connexion, position_id, LIGNE_TRANSFEREE)
    if transfert and (transfert["date_traitement"] or "") > bordereau["date"]:
        return None
    # Ce visa a déjà été archivé (bouton « Nouveau visa ») : le visa actuel est un nouveau.
    if connexion.execute("SELECT 1 FROM positions WHERE origine_id=? AND visa_bordereau_id=?",
                         (position_id, bordereau["id"])).fetchone():
        return None
    return bordereau


def visa_actuel(position_id, lot_id):
    """Bordereau du visa actuel d'une position s'il a été émis
    ({id, date, libelle, avis, observations}), sinon None (« Visa en cours »,
    son n° sera attribué à l'émission du prochain bordereau du lot)."""
    connexion = conn()
    bordereau = _visa_emis(connexion, position_id, lot_id)
    connexion.close()
    return bordereau


# ---------- produit déjà visé ailleurs ----------
# Quand une entreprise propose une marque et un type, on cherche si ce produit
# (ou un produit approchant) a déjà été visé sur une autre position : avis,
# projet, n° de visa, remarque. Seuls les visas ÉMIS (dans un bordereau) comptent.

PRODUITS_MAX = 5   # nombre de visas montrés dans l'encart du Suivi

# Avis qui jugent le produit. MQ (pièces manquantes) et SO (sans objet) ne disent
# rien du produit lui-même : ces visas ne sont pas montrés comme « déjà visé ».
AVIS_SUR_PRODUIT = ["VSO", "VAO", "VAOB", "REF"]


def _produit_compact(texte):
    """Texte réduit pour comparer deux produits : minuscules, sans accents, sans
    espaces ni ponctuation. Exemple : "Stratos 25/1-8" -> "stratos2518"."""
    resultat = ""
    for caractere in _norm(texte):
        if caractere.isalnum():
            resultat = resultat + caractere
    return resultat


def ressemblance_produit(marque_a, type_a, marque_b, type_b):
    """Compare deux produits (marque + type). Renvoie :
      - "identique"  : mêmes marque et type (accents, majuscules, ponctuation ignorés) ;
      - "approchant" : marques identiques ou très proches (faute de frappe) ET types dont
        l'un contient l'autre (« Stratos 25/1-8 » / « Stratos 25/1-8 PN10 ») ou très
        proches avec les mêmes nombres (« Alfea Excellia » / « Alféa Exellia ») ;
      - "" : autre produit (ou marque / type manquant).
    Les nombres d'un type désignent souvent la taille ou la puissance : « Stratos 25/1-8 »
    et « Stratos 30/1-8 » ne sont donc PAS rapprochés."""
    marque_a, marque_b = _produit_compact(marque_a), _produit_compact(marque_b)
    type_a, type_b = _produit_compact(type_a), _produit_compact(type_b)
    if not marque_a or not marque_b or not type_a or not type_b:
        return ""
    if marque_a != marque_b and difflib.SequenceMatcher(None, marque_a, marque_b).ratio() < 0.85:
        return ""
    if type_a == type_b:
        if marque_a == marque_b:
            return "identique"
        return "approchant"
    if len(type_a) <= len(type_b):
        court, long = type_a, type_b
    else:
        court, long = type_b, type_a
    if len(court) >= 4 and court in long:
        return "approchant"
    memes_nombres = re.findall(r"\d+", type_a) == re.findall(r"\d+", type_b)
    if memes_nombres and difflib.SequenceMatcher(None, type_a, type_b).ratio() >= 0.85:
        return "approchant"
    return ""


def visas_du_meme_produit(position, limite=PRODUITS_MAX):
    """Visas déjà émis, sur d'autres positions, pour le produit proposé par
    l'entreprise sur cette position (marque + type, voir ressemblance_produit).
    Les visas de la position elle-même (même lot, même référence) sont écartés ;
    pour une autre position, seul son dernier visa émis qui juge le produit compte
    (avis VSO, VAO, VAOB ou REF : voir AVIS_SUR_PRODUIT).
    Renvoie au plus `limite` éléments, du plus récent au plus ancien :
    {ressemblance, avis, remarque, date (AAAA-MM-JJ), numero ("VISA N°02"),
     projet, projet_nom, lot, marque, type, chef (nom du chef de projet du
     projet, responsable de ce visa), equipe (son équipe, "" sans équipe)}."""
    if not _produit_compact(position.get("marque")) or not _produit_compact(position.get("typeprop")):
        return []
    reference = (position.get("pos") or "").strip()
    connexion = conn()
    candidats = connexion.execute(
        """SELECT * FROM positions WHERE id<>? AND IFNULL(marque,'')<>'' AND IFNULL(typeprop,'')<>''""",
        (position.get("id") or 0,)).fetchall()
    trouves = {}
    for candidat in candidats:
        cle = (candidat["lot_id"], (candidat["pos"] or "").strip())
        if candidat["lot_id"] == position.get("lot_id") and cle[1] == reference:
            continue   # la position elle-même (ses anciens visas) : pas « ailleurs »
        ressemblance = ressemblance_produit(position.get("marque"), position.get("typeprop"),
                                            candidat["marque"], candidat["typeprop"])
        if not ressemblance:
            continue
        if candidat["archive"]:
            # Ancien visa : son bordereau est mémorisé ; avis et remarque sont ceux de la copie.
            if not candidat["visa_bordereau_id"]:
                continue
            ligne = connexion.execute("SELECT * FROM bordereaux WHERE id=?",
                                      (candidat["visa_bordereau_id"],)).fetchone()
            if ligne is None:
                continue
            date, numero = ligne["date"], format_numero(ligne["numero"])
            avis, remarque = candidat["avis"] or "", candidat["rem"] or ""
        else:
            bordereau = _visa_emis(connexion, candidat["id"], candidat["lot_id"])
            if bordereau is None:
                continue   # visa en cours : pas encore émis
            date, numero = bordereau["date"], bordereau["libelle"]
            avis, remarque = bordereau["avis"], bordereau["observations"]
        if avis not in AVIS_SUR_PRODUIT:
            continue   # MQ ou SO : pas un jugement sur le produit
        if cle in trouves and trouves[cle]["date"] >= date[:10]:
            continue   # on garde le visa le plus récent de chaque position
        # Projet, son chef de projet et son équipe : pour joindre le responsable de ce visa.
        projet = connexion.execute(
            """SELECT p.nom, c.nom AS chef_nom, t.nom AS equipe FROM projets p
               LEFT JOIN chefs c ON c.code = p.chef LEFT JOIN teams t ON t.id = c.team_id
               WHERE p.code=?""", (candidat["projet"],)).fetchone()
        lot = connexion.execute("SELECT * FROM lots WHERE id=?", (candidat["lot_id"],)).fetchone()
        trouves[cle] = {"ressemblance": ressemblance, "avis": avis, "remarque": remarque,
                        "date": date[:10], "numero": numero, "projet": candidat["projet"],
                        "projet_nom": projet["nom"] if projet else "",
                        "lot": libelle_lot(dict(lot)) if lot else "",
                        "marque": candidat["marque"], "type": candidat["typeprop"],
                        "chef": (projet["chef_nom"] or "") if projet else "",
                        "equipe": (projet["equipe"] or "") if projet else ""}
    connexion.close()
    resultat = sorted(trouves.values(), key=lambda visa: visa["date"], reverse=True)
    return resultat[:limite]


def _statut_miroir(connexion, position):
    """État d'une position vu par l'entreprise (voir statut_miroir)."""
    bordereau = _dernier_bordereau_position(connexion, position["id"], position["lot_id"])
    derniere = _derniere_ligne_depot(connexion, position["id"])
    transfert = _derniere_ligne_depot(connexion, position["id"], LIGNE_TRANSFEREE)
    rappel = connexion.execute(
        "SELECT date_depot FROM rappels_sb WHERE position_id=? AND cloture=0 ORDER BY id DESC LIMIT 1",
        (position["id"],)).fetchone()
    date_bordereau = bordereau["date"] if bordereau else ""

    statut = {"etat": ETAT_A_FOURNIR, "depuis": "", "avis": "", "date_visa": "",
              "bordereau": "", "bordereau_id": None, "observations": "", "refus": None}
    if bordereau:
        statut["avis"] = bordereau["avis"]
        statut["date_visa"] = bordereau["date"][:10]
        statut["bordereau"] = bordereau["libelle"]
        statut["bordereau_id"] = bordereau["id"]
        statut["observations"] = bordereau["observations"]

    if derniere and derniere["statut"] == LIGNE_TRANSMISE:
        # 1. Un envoi attend d'être pris en charge par Solares Bauen.
        statut["etat"] = ETAT_TRANSMIS
        statut["depuis"] = derniere["date_envoi"]
    elif transfert and (transfert["date_traitement"] or "") > date_bordereau:
        # 2. Transféré sur la plateforme, pas encore de bordereau depuis.
        statut["etat"] = ETAT_EXAMEN
        statut["depuis"] = transfert["date_traitement"]
    elif rappel and rappel["date_depot"] > date_bordereau[:10]:
        # 3. Documents reçus par un autre canal (dépôt fait par SB), en attente de visa.
        statut["etat"] = ETAT_EXAMEN
        statut["depuis"] = rappel["date_depot"]
    elif bordereau:
        # 4. Visé : date et avis du dernier bordereau émis.
        statut["etat"] = ETAT_VISE
        statut["depuis"] = bordereau["date"]

    # Dernier envoi refusé (et pas encore suivi d'un bordereau) : motif affiché.
    if (derniere and derniere["statut"] == LIGNE_REFUSEE
            and (derniere["date_traitement"] or "") > date_bordereau):
        statut["refus"] = {"date_envoi": derniere["date_envoi"],
                           "date": derniere["date_traitement"],
                           "motif": derniere["motif_refus"] or "",
                           "contact": derniere["contact_nom"] or ""}
    return statut


def statut_miroir(position_id):
    """État d'une position vu par l'entreprise sur le portail :
    {etat, depuis, avis, date_visa, bordereau, observations, refus}
      etat : afournir | transmis | examen | vise (libellés : LIBELLES_ETAT)
      avis / date_visa / bordereau / observations : dernier bordereau émis
      refus : {date_envoi, date, motif, contact} si le dernier envoi a été refusé."""
    position = get_position(position_id)
    if position is None:
        return None
    connexion = conn()
    statut = _statut_miroir(connexion, position)
    connexion.close()
    return statut


def peut_deposer(statut):
    """Une entreprise peut déposer sur une ligne « À fournir » ou déjà visée ;
    pas sur une ligne transmise ou en cours d'examen."""
    return statut["etat"] in (ETAT_A_FOURNIR, ETAT_VISE)


def get_miroir_lot(lot_id):
    """Positions ACTIVES d'un lot, chacune avec sa clé "miroir" (statut_miroir).
    C'est le tableau affiché à l'entreprise sur le portail."""
    lot = get_lot(lot_id)
    if lot is None:
        return []
    positions = get_positions(lot["projet"], lot_id)
    connexion = conn()
    for position in positions:
        position["miroir"] = _statut_miroir(connexion, position)
    connexion.close()
    return positions


# ---------- 1. Envoi depuis le portail ----------

def _verifier_fichier(nom, contenu):
    """Contrôle un fichier envoyé. Renvoie "" s'il est accepté, sinon le motif.
    Un ZIP doit être lisible et ne contenir aucun fichier d'un type interdit."""
    if extension_fichier(nom) in EXTENSIONS_INTERDITES:
        return f"« {nom} » : ce type de fichier est interdit."
    if not contenu:
        return f"« {nom} » est vide."
    if len(contenu) > TAILLE_MAXI_FICHIER:
        return f"« {nom} » dépasse {TAILLE_MAXI_FICHIER // (1024 * 1024)} Mo."
    if extension_fichier(nom) == "zip":
        try:
            with zipfile.ZipFile(io.BytesIO(contenu)) as archive:
                noms_internes = archive.namelist()
        except (zipfile.BadZipFile, OSError):
            return f"« {nom} » n'est pas un ZIP lisible."
        for nom_interne in noms_internes:
            if extension_fichier(nom_interne) in EXTENSIONS_INTERDITES:
                return f"« {nom} » contient un type de fichier interdit (« {nom_interne} »)."
    return ""


def est_plan(position):
    """Plan / note de calcul = position sans référence CCTP."""
    return not (position.get("pos") or "").strip()


def _verifier_ligne_libre(ligne):
    """Contrôles d'une soumission libre (voir enregistrer_envoi) : désignation,
    référence pour une position CCTP/Matériel, fichiers. Renvoie la liste des erreurs."""
    libre = ligne["libre"]
    type_libre = libre.get("type")
    if type_libre not in (LIBRE_CCTP, LIBRE_PLAN):
        return ["Soumission libre : type de document inconnu."]
    reference = (libre.get("pos") or "").strip()
    designation = (libre.get("des") or "").strip()
    libelle = f"Document non prévu « {designation or '…'} »"
    erreurs = []
    if not designation:
        erreurs.append("Document non prévu : indiquez sa désignation.")
    if type_libre == LIBRE_CCTP and not reference:
        erreurs.append(f"{libelle} : indiquez la référence CCTP (« HC » si hors CCTP).")
    if ligne.get("extraits"):
        erreurs.append(f"{libelle} : pas de fichier commun pour un document non prévu.")
    fichiers = ligne.get("fichiers") or []
    if type_libre == LIBRE_PLAN:
        if ligne.get("pieces"):
            erreurs.append(f"{libelle} : pas de pièce à choisir pour un plan ou une note de calcul.")
        if len(fichiers) != 1:
            erreurs.append(f"{libelle} : déposez un seul fichier (PDF, ZIP…).")
    elif not fichiers:
        erreurs.append(f"{libelle} : déposez au moins un fichier.")
    for fichier in fichiers:
        erreur = _verifier_fichier(fichier["nom"], fichier["contenu"])
        if erreur:
            erreurs.append(erreur)
    return erreurs


def _position_de_ligne(ligne):
    """Position d'une ligne de dépôt (dictionnaire). Soumission libre pas encore
    transférée : une position « à créer », sans identifiant, avec la référence et
    la désignation saisies par l'entreprise (clé "libre" = type)."""
    if ligne.get("position_id"):
        return get_position(ligne["position_id"])
    return {"id": None, "pos": ligne.get("libre_pos") or "", "des": ligne.get("libre_des") or "",
            "att": [], "avis": "", "libre": ligne.get("libre_type") or LIBRE_CCTP}


def enregistrer_envoi(entreprise_id, lot_id, lignes, fichiers_communs=None):
    """Enregistre un envoi fait sur le portail. Rien n'est enregistré si un
    seul contrôle échoue (l'entreprise corrige puis renvoie).
    - lignes : une par position cochée :
        {"position_id", "pieces": [noms des pièces attendues concernées,
                                   ou PIECE_AUTRE],
         "autre": libellé de la pièce « Autre » (obligatoire si cochée),
         "commentaire", "fichiers": [{"nom", "contenu"}],      # fichiers propres
         "extraits": [{"nom": fichier commun, "extrait": "12-30"}]}
      Soumission libre (document non prévu) : "libre": {"type": LIBRE_CCTP ou
      LIBRE_PLAN, "pos", "des"} à la place de "position_id" ; pièces = libellés
      libres (nature du document, facultative) ; fichiers propres seulement.
    - fichiers_communs : [{"nom", "contenu"}] (gros PDF ou ZIP partagé).
    Un extrait = pages d'un PDF ou fichier dans un ZIP ; obligatoire si le
    fichier commun sert à plusieurs lignes ; sans objet pour un autre type.
    Position CCTP : au moins une pièce. Plan / note de calcul : pas de pièce à
    choisir, UN seul fichier propre (PDF, ZIP…), pas de dossier commun.
    Le chargé d'affaires de l'envoi est celui du lot (il reçoit le mail de refus).
    Renvoie (True, id du dépôt) ou (False, [liste des erreurs])."""
    erreurs = []
    lot = get_lot(lot_id)
    if lot is None or lot.get("entreprise_id") != entreprise_id:
        return False, ["Ce lot n'est pas accessible à votre entreprise."]
    projet = get_projet(lot["projet"])
    if projet is None or projet.get("archive"):
        return False, ["Ce projet n'accepte plus de dépôt."]

    # --- Chargé d'affaires du lot ---
    contact = get_contact_lot(lot_id)
    if contact is None:
        erreurs.append("Désignez le chargé d'affaires du lot (en haut de la page).")
    if not lignes:
        erreurs.append("Cochez au moins une ligne.")

    # --- Fichiers communs : contrôle, nombre de pages, nombre d'utilisations ---
    communs = {}
    for fichier in fichiers_communs or []:
        erreur = _verifier_fichier(fichier["nom"], fichier["contenu"])
        if erreur:
            erreurs.append(erreur)
        communs[fichier["nom"]] = fichier
    utilisations = {}
    for ligne in lignes:
        for extrait in ligne.get("extraits") or []:
            utilisations[extrait["nom"]] = utilisations.get(extrait["nom"], 0) + 1
    nb_pages = {}
    for nom in communs:
        if extension_fichier(nom) == "pdf":
            nb_pages[nom] = nombre_pages_pdf(communs[nom]["contenu"])
            if nb_pages[nom] is None and utilisations.get(nom, 0) > 0:
                erreurs.append(f"« {nom} » n'est pas un PDF lisible.")
        if utilisations.get(nom, 0) == 0:
            erreurs.append(f"« {nom} » n'est rattaché à aucune ligne.")

    # --- Chaque ligne ---
    positions_du_lot = {}
    for position in get_miroir_lot(lot_id):
        positions_du_lot[position["id"]] = position
    deja_vues = []
    for ligne in lignes:
        if ligne.get("libre"):
            erreurs.extend(_verifier_ligne_libre(ligne))
            continue
        position = positions_du_lot.get(ligne.get("position_id"))
        if position is None:
            erreurs.append("Une des lignes n'appartient pas à ce lot.")
            continue
        libelle = _libelle_position(position)
        if position["id"] in deja_vues:
            erreurs.append(f"{libelle} : ligne présente deux fois dans l'envoi.")
        deja_vues.append(position["id"])
        if not peut_deposer(position["miroir"]):
            erreurs.append(f"{libelle} : un envoi est déjà transmis ou en cours d'examen.")
        noms_attendus = []
        for piece in position["att"]:
            noms_attendus.append(piece["n"])
        if est_plan(position):
            # Plan / note de calcul : le document EST la pièce ; un seul fichier.
            if ligne.get("pieces"):
                erreurs.append(f"{libelle} : pas de pièce à choisir pour un plan ou une note de calcul.")
            if ligne.get("extraits") or len(ligne.get("fichiers") or []) != 1:
                erreurs.append(f"{libelle} : déposez un seul fichier (PDF, ZIP…).")
        else:
            if not ligne.get("pieces"):
                erreurs.append(f"{libelle} : choisissez au moins une pièce.")
            for piece in ligne.get("pieces") or []:
                if piece == PIECE_AUTRE:
                    if not (ligne.get("autre") or "").strip():
                        erreurs.append(f"{libelle} : précisez la pièce « {PIECE_AUTRE} ».")
                elif piece not in noms_attendus:
                    erreurs.append(f"{libelle} : « {piece} » ne fait pas partie des pièces attendues.")
            if not ligne.get("fichiers") and not ligne.get("extraits"):
                erreurs.append(f"{libelle} : déposez au moins un fichier.")
        for fichier in ligne.get("fichiers") or []:
            erreur = _verifier_fichier(fichier["nom"], fichier["contenu"])
            if erreur:
                erreurs.append(erreur)
        for extrait in ligne.get("extraits") or []:
            nom = extrait["nom"]
            texte = (extrait.get("extrait") or "").strip()
            if nom not in communs:
                erreurs.append(f"{libelle} : « {nom} » ne fait pas partie du dossier commun.")
                continue
            type_fichier = extension_fichier(nom)
            if type_fichier == "pdf" and texte and nb_pages.get(nom):
                pages, erreur = pages_depuis_texte(texte, nb_pages[nom])
                if erreur:
                    erreurs.append(f"{libelle} : {erreur}")
            if type_fichier in ("pdf", "zip") and not texte and utilisations[nom] > 1:
                if type_fichier == "pdf":
                    quoi = "les pages"
                else:
                    quoi = "le fichier dans le ZIP"
                erreurs.append(f"{libelle} : indiquez {quoi} de « {nom} ».")
    if erreurs:
        return False, erreurs

    # --- Tout est bon : enregistrement (contenu des fichiers communs stocké une fois) ---
    connexion = conn()
    curseur = connexion.execute(
        """INSERT INTO depots(entreprise_id, projet, lot_id, contact_id, contact_nom,
                              contact_email, date_envoi) VALUES(?,?,?,?,?,?,?)""",
        (entreprise_id, projet["code"], lot_id, contact["id"], nom_avec_civilite(contact),
         contact["email"], _maintenant()))
    depot_id = curseur.lastrowid
    references_communes = {}
    for nom in communs:
        references_communes[nom] = stockage.enregistrer(connexion, communs[nom]["contenu"])
    for ligne in lignes:
        # La pièce « Autre » est enregistrée sous le libellé saisi par l'entreprise.
        pieces = []
        for piece in ligne.get("pieces") or []:
            if piece == PIECE_AUTRE:
                piece = ligne["autre"].strip()
            if piece not in pieces:
                pieces.append(piece)
        # Soumission libre : pas de position, la référence et la désignation saisies.
        libre = ligne.get("libre") or {}
        if libre:
            position_id = None
        else:
            position_id = ligne["position_id"]
        reference_libre = ""
        if libre.get("type") == LIBRE_CCTP:
            reference_libre = (libre.get("pos") or "").strip()
        curseur = connexion.execute(
            """INSERT INTO depot_lignes(depot_id, position_id, pieces, commentaire, statut,
                                        libre_type, libre_pos, libre_des)
               VALUES(?,?,?,?,?,?,?,?)""",
            (depot_id, position_id, json.dumps(pieces, ensure_ascii=False),
             (ligne.get("commentaire") or "").strip(), LIGNE_TRANSMISE,
             libre.get("type") or "", reference_libre, (libre.get("des") or "").strip()))
        ligne_id = curseur.lastrowid
        for fichier in ligne.get("fichiers") or []:
            reference = stockage.enregistrer(connexion, fichier["contenu"])
            connexion.execute(
                """INSERT INTO depot_fichiers(depot_ligne_id, nom, taille, contenu_id, extrait, commun)
                   VALUES(?,?,?,?,'',0)""",
                (ligne_id, _san_file(fichier["nom"]), len(fichier["contenu"]), reference))
        for extrait in ligne.get("extraits") or []:
            texte = (extrait.get("extrait") or "").strip()
            if extension_fichier(extrait["nom"]) not in ("pdf", "zip"):
                texte = ""   # Excel, DWG... : transmis en entier
            connexion.execute(
                """INSERT INTO depot_fichiers(depot_ligne_id, nom, taille, contenu_id, extrait, commun)
                   VALUES(?,?,?,?,?,1)""",
                (ligne_id, _san_file(extrait["nom"]), len(communs[extrait["nom"]]["contenu"]),
                 references_communes[extrait["nom"]], texte))
    connexion.commit()
    connexion.close()
    return True, depot_id


# ---------- 2. Page « Dépôts reçus » ----------

def _fichiers_ligne_depot(connexion, ligne_id):
    """Fichiers d'une ligne de dépôt : propres d'abord, puis extraits du dossier commun."""
    fichiers = []
    for fichier in connexion.execute(
            "SELECT * FROM depot_fichiers WHERE depot_ligne_id=? ORDER BY commun, id",
            (ligne_id,)).fetchall():
        fichiers.append(dict(fichier))
    return fichiers


def get_depots_en_attente(chefs=None):
    """Envois qui ont encore des lignes « transmis », le plus ancien en premier.
    - chefs : liste de codes de chefs de projet (Mes projets / Équipe) ;
      None = tous les projets.
    Chaque envoi : champs du dépôt + entreprise, projet_nom, chef, lot (dict),
    echeance (« à viser avant le » = date d'envoi + délai du lot en jours ouvrés)
    et lignes : [{id, position (dict), pieces, commentaire, visa_emis, fichiers}]
    où visa_emis = bordereau du visa actuel (qui sera archivé) ou None.
    Soumission libre : position « à créer » (voir _position_de_ligne)."""
    connexion = conn()
    envois = []
    for ligne_depot in connexion.execute("SELECT * FROM depots ORDER BY date_envoi, id").fetchall():
        depot = dict(ligne_depot)
        projet = connexion.execute("SELECT * FROM projets WHERE code=?", (depot["projet"],)).fetchone()
        lot = connexion.execute("SELECT * FROM lots WHERE id=?", (depot["lot_id"],)).fetchone()
        if projet is None or lot is None:
            continue
        if chefs is not None and projet["chef"] not in chefs:
            continue
        lignes = []
        for ligne in connexion.execute(
                "SELECT * FROM depot_lignes WHERE depot_id=? AND statut=? ORDER BY id",
                (depot["id"], LIGNE_TRANSMISE)).fetchall():
            lignes.append(dict(ligne))
        if not lignes:
            continue
        entreprise = connexion.execute("SELECT nom FROM entreprises WHERE id=?",
                                       (depot["entreprise_id"],)).fetchone()
        depot["entreprise"] = entreprise["nom"] if entreprise else ""
        depot["projet_nom"] = projet["nom"]
        depot["chef"] = projet["chef"]
        depot["lot"] = dict(lot)
        delai = int(lot["delai_visa_sb"] or DELAI_VISA_SB)
        depot["echeance"] = calendrier.ajouter_jours_ouvres(
            dt.date.fromisoformat(depot["date_envoi"][:10]), delai).isoformat()
        depot["lignes"] = lignes
        envois.append(depot)
    # Détail de chaque ligne.
    for depot in envois:
        for ligne in depot["lignes"]:
            ligne["pieces"] = json.loads(ligne["pieces"] or "[]")
            ligne["position"] = _position_de_ligne(ligne)
            if ligne["position_id"]:
                ligne["visa_emis"] = _visa_emis(connexion, ligne["position_id"], depot["lot_id"])
            else:
                ligne["visa_emis"] = None   # soumission libre : pas encore de position
            ligne["fichiers"] = _fichiers_ligne_depot(connexion, ligne["id"])
    connexion.close()
    return envois


def nb_lignes_en_attente(chefs=None):
    """Nombre de lignes à transférer ou refuser (compteur du menu et du Dashboard)."""
    total = 0
    for depot in get_depots_en_attente(chefs):
        total = total + len(depot["lignes"])
    return total


def lire_fichier_depot(fichier_id):
    """Contenu d'un fichier déposé : (nom, octets), ou (None, None)."""
    connexion = conn()
    fichier = connexion.execute("SELECT * FROM depot_fichiers WHERE id=?", (fichier_id,)).fetchone()
    if fichier is None:
        connexion.close()
        return None, None
    contenu = stockage.lire(connexion, fichier["contenu_id"])
    connexion.close()
    return fichier["nom"], contenu


def _position_pour_ligne_libre(ligne, destination):
    """Soumission libre au transfert : position qui reçoit les documents.
    destination = {"position_id": id} (position existante du même lot) ou
    {"trame": nom} (nouvelle position) ; sans précision : nouvelle position avec
    la trame proposée pour sa désignation (« Autre document » pour un plan ou une
    note de calcul). La ligne de dépôt est ensuite rattachée à cette position.
    Renvoie (identifiant de la position, nouvelle position ?, message d'erreur)."""
    destination = destination or {}
    if destination.get("position_id"):
        position = get_position(destination["position_id"])
        if position is None or position.get("archive") or position["lot_id"] != ligne["lot_id"]:
            return None, False, "La position choisie n'appartient pas à ce lot."
        position_id = position["id"]
        nouvelle = False
    else:
        lot = get_lot(ligne["lot_id"])
        if ligne["libre_type"] == LIBRE_PLAN:
            reference = ""
            trame = destination.get("trame") or TRAME_AUTRE
        else:
            reference = ligne["libre_pos"] or ""
            trame = destination.get("trame") or suggest_trame(ligne["libre_des"])
        position_id = add_position(lot["projet"], lot["id"], reference, ligne["libre_des"], trame)
        nouvelle = True
    connexion = conn()
    connexion.execute("UPDATE depot_lignes SET position_id=? WHERE id=?", (position_id, ligne["id"]))
    connexion.commit()
    connexion.close()
    return position_id, nouvelle, ""


def transferer_lignes(ligne_ids, traite_par="", destinations=None):
    """Transfère des lignes de dépôt vers la plateforme (bouton « Transférer »).
    destinations : {ligne_id: {"position_id": id} ou {"trame": nom}} pour les
    soumissions libres (voir _position_pour_ligne_libre).
    Pour chaque ligne encore « transmis » :
      1. si le visa actuel a été émis : copie dans « Anciens visas », et la date
         du visa est vidée (nouveau cycle) ; sinon les documents rejoignent le
         visa en cours ;
      2. les fichiers sont ajoutés aux Documents de la position ; pour un fichier
         commun avec des pages indiquées, seules ces pages sont copiées ;
      3. les pièces déclarées passent « À vérifier » (date = date d'envoi) ;
      4. le rappel SB est ouvert à la date d'envoi.
    Renvoie une liste de comptes rendus : [{ligne_id, libelle, archive,
    bordereau, nb_fichiers, nouvelle (position créée), erreur}]."""
    comptes_rendus = []
    for ligne_id in ligne_ids:
        compte_rendu = {"ligne_id": ligne_id, "libelle": "", "archive": False,
                        "bordereau": "", "nb_fichiers": 0, "nouvelle": False, "erreur": ""}
        comptes_rendus.append(compte_rendu)
        connexion = conn()
        ligne = connexion.execute(
            """SELECT dl.*, d.date_envoi, d.lot_id FROM depot_lignes dl
               JOIN depots d ON d.id = dl.depot_id WHERE dl.id=?""", (ligne_id,)).fetchone()
        connexion.close()
        if ligne is None or ligne["statut"] != LIGNE_TRANSMISE:
            compte_rendu["erreur"] = "Ligne déjà traitée ou introuvable."
            continue
        position_id = ligne["position_id"]
        if not position_id:
            # Soumission libre : nouvelle position, ou position existante choisie par SB.
            destination = (destinations or {}).get(ligne_id)
            position_id, nouvelle, erreur = _position_pour_ligne_libre(dict(ligne), destination)
            if erreur:
                compte_rendu["erreur"] = erreur
                continue
            compte_rendu["nouvelle"] = nouvelle
        position = get_position(position_id)
        if position is None or position.get("archive"):
            compte_rendu["erreur"] = "La position n'existe plus."
            continue
        compte_rendu["libelle"] = _libelle_position(position)
        lot = get_lot(ligne["lot_id"])
        date_reception = ligne["date_envoi"][:10]
        pieces = json.loads(ligne["pieces"] or "[]")

        # 1. Ancien visa émis -> « Anciens visas » (avant d'ajouter les fichiers).
        connexion = conn()
        bordereau = _visa_emis(connexion, position["id"], lot["id"])
        connexion.close()
        if bordereau is not None:
            archive_position(position["id"])
            position["datevisa"] = ""
            compte_rendu["archive"] = True
            compte_rendu["bordereau"] = bordereau["libelle"]

        # 2. Fichiers -> Documents de la position.
        connexion = conn()
        noms_ajoutes = []
        for fichier in _fichiers_ligne_depot(connexion, ligne_id):
            nom = fichier["nom"]
            reference = fichier["contenu_id"]
            taille = fichier["taille"]
            if fichier["commun"] and fichier["extrait"] and extension_fichier(nom) == "pdf":
                extrait, erreur = extraire_pages(stockage.lire(connexion, reference), fichier["extrait"])
                if extrait is not None:
                    base = os.path.splitext(nom)[0]
                    nom = _san_file(f"{base} p.{fichier['extrait'].replace(' ', '')}.pdf")
                    reference = stockage.enregistrer(connexion, extrait)
                    taille = len(extrait)
            derniere_version = connexion.execute(
                "SELECT MAX(version) FROM fichiers WHERE position_id=? AND nom=?",
                (position["id"], nom)).fetchone()[0]
            if derniere_version is None:
                derniere_version = 0
            connexion.execute("UPDATE fichiers SET courant=0 WHERE position_id=? AND nom=?",
                              (position["id"], nom))
            connexion.execute(
                """INSERT INTO fichiers(position_id, attendu, nom, chemin, version,
                                        courant, date_upload, taille, contenu_id)
                   VALUES(?,?,?,'',?,1,?,?,?)""",
                (position["id"], json.dumps(pieces, ensure_ascii=False), nom,
                 derniere_version + 1, date_reception, taille, reference))
            noms_ajoutes.append(nom)
        connexion.execute(
            "UPDATE depot_lignes SET statut=?, date_traitement=?, traite_par=? WHERE id=?",
            (LIGNE_TRANSFEREE, _maintenant(), traite_par, ligne_id))
        connexion.commit()
        connexion.close()
        compte_rendu["nb_fichiers"] = len(noms_ajoutes)

        # 3. Pièces « À vérifier » (une pièce « Autre » devient une nouvelle pièce
        #    de la position ; un plan sans pièce déclarée couvre toutes ses pièces) ;
        # 4. rappel SB depuis la date d'envoi.
        noms_existants = []
        for piece in position["att"]:
            noms_existants.append(piece["n"])
        for nom_piece in pieces:
            if nom_piece not in noms_existants:
                position["att"].append({"n": nom_piece, "s": STATUT_A_VERIFIER, "d": date_reception})
        marquer_a_verifier(position, pieces, date_reception)
        save_position(position)
        ouvrir_rappel_sb(position["projet"], lot, position["id"], noms_ajoutes, date_reception)
    return comptes_rendus


def refuser_lignes(ligne_ids, motif, traite_par=""):
    """Refuse des lignes de dépôt (rien n'entre dans la plateforme). Le motif est
    obligatoire : l'entreprise le voit sur le portail, et il est repris dans le
    brouillon de mail au chargé d'affaires (build_eml_refus).
    Renvoie (True, nombre de lignes refusées) ou (False, message)."""
    motif = (motif or "").strip()
    if motif == "":
        return False, "Le motif du refus est obligatoire."
    connexion = conn()
    nombre = 0
    for ligne_id in ligne_ids:
        curseur = connexion.execute(
            """UPDATE depot_lignes SET statut=?, date_traitement=?, traite_par=?, motif_refus=?
               WHERE id=? AND statut=?""",
            (LIGNE_REFUSEE, _maintenant(), traite_par, motif, ligne_id, LIGNE_TRANSMISE))
        nombre = nombre + curseur.rowcount
    connexion.commit()
    connexion.close()
    return True, nombre


def build_eml_refus(ligne_ids):
    """Brouillon de mail (.eml, ouvert dans Outlook) prévenant le chargé
    d'affaires que des documents n'ont pas été acquis, avec le motif.
    Destinataire : le chargé d'affaires de l'envoi ; personne en copie.
    Pas de formule de politesse ni de signature (comme la relance).
    Les lignes doivent venir du même envoi. Renvoie les octets, ou None."""
    connexion = conn()
    lignes = []
    for ligne_id in ligne_ids:
        ligne = connexion.execute("SELECT * FROM depot_lignes WHERE id=?", (ligne_id,)).fetchone()
        if ligne is not None:
            lignes.append(dict(ligne))
    if not lignes:
        connexion.close()
        return None
    depot = dict(connexion.execute("SELECT * FROM depots WHERE id=?",
                                   (lignes[0]["depot_id"],)).fetchone())
    connexion.close()
    projet = get_projet(depot["projet"])
    lot = get_lot(depot["lot_id"])
    libelle = libelle_lot(lot)

    # Un seul motif pour tout le refus (cas courant), sinon un motif par ligne.
    motifs = []
    for ligne in lignes:
        if ligne["motif_refus"] not in motifs:
            motifs.append(ligne["motif_refus"])
    puces = []
    for ligne in lignes:
        position = _position_de_ligne(ligne) or {}
        texte = "   •  " + _libelle_position(position)
        if len(motifs) > 1:
            texte = texte + " — " + (ligne["motif_refus"] or "")
        puces.append(texte)
    corps = (
        f"Bonjour {depot['contact_nom']},\n\n"
        f"Les documents que vous avez déposés le {_date_fr(depot['date_envoi'])} sur "
        f"{NOM_PORTAIL} pour le {libelle} de l'opération {projet['nom']} "
        f"n'ont pas pu être acquis pour les positions suivantes :\n\n"
        + "\n".join(puces) + "\n\n")
    if len(motifs) == 1:
        corps = corps + f"Motif : {motifs[0]}\n\n"
    corps = corps + f"Merci de déposer à nouveau ces éléments sur la {NOM_PORTAIL}.\n"
    sujet = f"{projet['nom']} — {libelle} — Documents non acquis"
    contenu_eml = (f"To: {_entete_mail(depot['contact_email'])}\r\nSubject: {_entete_mail(sujet)}\r\n"
                   f"X-Unsent: 1\r\nContent-Type: text/plain; charset=utf-8\r\n\r\n{corps}")
    return contenu_eml.encode("utf-8")


# ---------- 3. Historique et journal ----------

def _resume_ligne(connexion, ligne):
    """Complète une ligne de dépôt (dictionnaire) pour l'affichage : libellé de
    la position, nombre de fichiers, projet / lot / entreprise. Soumission libre
    pas encore transférée : référence et désignation saisies, « (document non prévu) »."""
    if ligne["position_id"]:
        position = connexion.execute("SELECT pos, des FROM positions WHERE id=?",
                                     (ligne["position_id"],)).fetchone()
        ligne["libelle"] = _libelle_position(dict(position)) if position else "(position supprimée)"
    else:
        libelle = _libelle_position({"pos": ligne.get("libre_pos"), "des": ligne.get("libre_des")})
        ligne["libelle"] = libelle + " (document non prévu)"
    ligne["nb_fichiers"] = connexion.execute(
        "SELECT COUNT(*) FROM depot_fichiers WHERE depot_ligne_id=?", (ligne["id"],)).fetchone()[0]
    return ligne


def get_historique_envois(entreprise_id, lot_id):
    """Envois d'une entreprise sur un lot (portail, « Mes envois »), du plus
    récent au plus ancien : une entrée par ligne avec date_envoi, contact_nom,
    libelle, nb_fichiers, statut, date_traitement, motif_refus."""
    connexion = conn()
    resultat = []
    for ligne in connexion.execute(
            """SELECT dl.*, d.date_envoi, d.contact_nom FROM depot_lignes dl
               JOIN depots d ON d.id = dl.depot_id
               WHERE d.entreprise_id=? AND d.lot_id=?
               ORDER BY d.date_envoi DESC, dl.id""", (entreprise_id, lot_id)).fetchall():
        resultat.append(_resume_ligne(connexion, dict(ligne)))
    connexion.close()
    return resultat


def get_journal_depots(chefs=None, limite=15):
    """Derniers transferts et refus (page Dépôts reçus), du plus récent au plus
    ancien. chefs : liste de codes de chefs de projet, None = tous."""
    connexion = conn()
    resultat = []
    for ligne in connexion.execute(
            """SELECT dl.*, d.date_envoi, d.contact_nom, d.projet, d.lot_id,
                      e.nom AS entreprise, p.chef, l.lot_numero
               FROM depot_lignes dl JOIN depots d ON d.id = dl.depot_id
               JOIN projets p ON p.code = d.projet JOIN lots l ON l.id = d.lot_id
               LEFT JOIN entreprises e ON e.id = d.entreprise_id
               WHERE dl.statut IN (?, ?)
               ORDER BY dl.date_traitement DESC, dl.id DESC""",
            (LIGNE_TRANSFEREE, LIGNE_REFUSEE)).fetchall():
        if chefs is not None and ligne["chef"] not in chefs:
            continue
        resultat.append(_resume_ligne(connexion, dict(ligne)))
        if len(resultat) >= limite:
            break
    connexion.close()
    return resultat
