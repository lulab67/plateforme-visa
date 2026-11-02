# ============================================================================
#  stockage.py — STOCKAGE DU CONTENU DES FICHIERS (PDF déposés)
# ----------------------------------------------------------------------------
#  Couche unique qui sait où vit le contenu binaire d'un fichier.
#  Aujourd'hui : dans la base SQLite (table "contenus").
#  Demain : MongoDB/GridFS, SharePoint ou un dossier du Synology — il suffira
#  de réécrire les deux fonctions ci-dessous, sans toucher au reste de l'appli.
#
#  Le reste du code ne manipule qu'une "référence" (un identifiant) renvoyée
#  par enregistrer() et relue par lire().
# ============================================================================


def creer_tables(connexion):
    """Crée la table des contenus si elle n'existe pas (appelée par core.init_db)."""
    connexion.execute("""CREATE TABLE IF NOT EXISTS contenus(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      data BLOB)""")


def enregistrer(connexion, contenu):
    """Enregistre le contenu (octets) et renvoie sa référence.
    La validation (commit) est faite par l'appelant, avec le reste de l'opération."""
    curseur = connexion.execute("INSERT INTO contenus(data) VALUES(?)", (contenu,))
    return curseur.lastrowid


def lire(connexion, reference):
    """Renvoie le contenu (octets) correspondant à la référence, ou None."""
    if reference is None:
        return None
    ligne = connexion.execute("SELECT data FROM contenus WHERE id=?", (reference,)).fetchone()
    if ligne is None:
        return None
    return bytes(ligne[0])
