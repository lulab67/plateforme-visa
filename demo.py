# ============================================================================
#  demo.py — JEU DE DONNÉES DE DÉMONSTRATION
# ----------------------------------------------------------------------------
#  Remplit une base VIDE avec un exemple réaliste : 2 équipes, 2 chefs de projet,
#  4 projets (dont 1 archivé), 5 lots CVC / plomberie, des positions à tous
#  les stades (VSO, VAO, VAOB, REF, MQ, SO), des documents déposés (qui
#  ouvrent des rappels de visa côté Solares Bauen et passent les pièces en
#  « À vérifier »), des relances titulaires datées, des bordereaux émis,
#  un ancien visa et un historique de visas sur 12 mois. Les trames sont
#  celles d'une base neuve (tableau templates/TRAMES_VISA.xlsx).
#  Ne touche jamais une base qui contient déjà des chefs de projet.
#  Les dates sont calculées par rapport à aujourd'hui, pour que les
#  échéances (en retard / proches / dans les temps) restent parlantes.
# ============================================================================

import json
import unicodedata
import datetime as dt
import calendrier
import core


def _il_y_a(jours):
    """Date ISO d'il y a N jours calendaires."""
    return (dt.date.today() - dt.timedelta(days=jours)).isoformat()


def _il_y_a_ouvres(jours):
    """Date ISO d'il y a N jours ouvrés (calendrier d'Alsace-Moselle)."""
    jour = dt.date.today()
    restant = jours
    while restant > 0:
        jour = jour - dt.timedelta(days=1)
        if calendrier.est_ouvre(jour):
            restant = restant - 1
    return jour.isoformat()


def pdf_exemple(titre):
    """Petit PDF valide d'une page portant un titre (document de démonstration)."""
    texte = unicodedata.normalize("NFD", titre).encode("ascii", "ignore").decode()
    texte = texte.replace("\\", "").replace("(", "[").replace(")", "]")
    flux = (f"BT /F1 20 Tf 60 760 Td ({texte}) Tj ET\n"
            "BT /F1 11 Tf 60 730 Td (Document de demonstration - Solares Bauen) Tj ET")
    objets = [
        "<< /Type /Catalog /Pages 2 0 R >>",
        "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Contents 4 0 R "
        "/Resources << /Font << /F1 5 0 R >> >> >>",
        f"<< /Length {len(flux)} >>\nstream\n{flux}\nendstream",
        "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    ]
    sortie = "%PDF-1.4\n"
    positions = []
    for numero, objet in enumerate(objets, start=1):
        positions.append(len(sortie))
        sortie += f"{numero} 0 obj\n{objet}\nendobj\n"
    debut_xref = len(sortie)
    sortie += f"xref\n0 {len(objets) + 1}\n0000000000 65535 f \n"
    for position in positions:
        sortie += f"{position:010d} 00000 n \n"
    sortie += f"trailer\n<< /Size {len(objets) + 1} /Root 1 0 R >>\nstartxref\n{debut_xref}\n%%EOF\n"
    return sortie.encode("latin-1")


def _position(projet, lot_id, ordre, pos, des, trame, avis="MQ", marque="", typeprop="",
              recus=0, maj=0, visa_il_y_a=None, rem="", marque_moe="", type_moe=""):
    """Crée une position puis la place dans l'état voulu :
    - recus : nombre de pièces attendues passées en « Reçu et ok » ;
    - maj   : nombre de pièces suivantes en « MAJ attendue » (le reste : Manquant) ;
    - visa_il_y_a : date du visa, en jours avant aujourd'hui.
    Renvoie l'identifiant de la position."""
    verrou = 1 if (marque_moe or type_moe) else 0
    position_id = core.add_position(projet, lot_id, pos, des, trame, ordre=ordre,
                                    marque_moe=marque_moe, type_moe=type_moe, mt_verrou=verrou)
    position = core.get_position(position_id)
    for rang, piece in enumerate(position["att"]):
        if rang < recus:
            piece["s"] = "Reçu et ok"
            piece["d"] = _il_y_a(30 - rang)
        elif rang < recus + maj:
            piece["s"] = "MAJ attendue"
    # Un document visé favorablement a ses points de contrôle vérifiés.
    if avis in ("VSO", "VAO"):
        for controle in position["ctrl"]:
            controle["v"] = 1
    position["avis"] = avis
    position["relance"] = 1 if avis in ("MQ", "VAOB", "REF") else 0
    position["marque"] = marque
    position["typeprop"] = typeprop
    position["rem"] = rem
    position["datevisa"] = _il_y_a(visa_il_y_a) if visa_il_y_a is not None else ""
    core.save_position(position)
    return position_id


def _deposer(projet_code, lot_id, position_id, nom_fichier, il_y_a_ouvres, deja_vise=False):
    """Dépose un PDF d'exemple sur une position, à une date passée (en jours
    ouvrés), ce qui ouvre le rappel de visa correspondant et passe les pièces
    pas encore « Reçu et ok » en « À vérifier » (comme un dépôt dans l'appli).
    deja_vise=True : le rappel est aussitôt clôturé (document déjà visé)."""
    projet = core.get_projet(projet_code)
    lot = core.get_lot(lot_id)
    position = core.get_position(position_id)
    resultat = core.save_fichier(projet, lot, position, nom_fichier, pdf_exemple(nom_fichier[:-4]),
                                 date_depot=_il_y_a_ouvres(il_y_a_ouvres))
    position = core.get_position(position_id)
    position["intitule"] = resultat["nom"]
    if not deja_vise:
        core.marquer_a_verifier(position)
    core.save_position(position)
    if deja_vise:
        core.cloturer_rappel_sb(resultat["rappel_id"])
        # Clôture datée du visa de la position (et non du jour du chargement de la démo).
        if position.get("datevisa"):
            connexion = core.conn()
            connexion.execute("UPDATE rappels_sb SET date_cloture=? WHERE id=?",
                              (position["datevisa"], resultat["rappel_id"]))
            connexion.commit()
            connexion.close()


def _relance(projet, lot_id, il_y_a, references):
    """Inscrit au journal une relance titulaire datée d'il y a N jours."""
    connexion = core.conn()
    connexion.execute(
        """INSERT INTO relances(projet, lot_id, date, niveau, nb_positions, positions)
           VALUES(?,?,?,?,?,?)""",
        (projet, lot_id, _il_y_a(il_y_a), 1, len(references),
         json.dumps(references, ensure_ascii=False)))
    connexion.commit()
    connexion.close()


def _bordereau(projet_code, lot_id, il_y_a):
    """Émet un bordereau de visa du lot (numéro suivant), daté d'il y a N jours."""
    bordereau = core.emettre_bordereau(core.get_projet(projet_code), core.get_lot(lot_id),
                                       core.get_positions(projet_code, lot_id))
    connexion = core.conn()
    connexion.execute("UPDATE bordereaux SET date=? WHERE id=?",
                      (_il_y_a(il_y_a) + "T10:00:00", bordereau["id"]))
    # L'émission clôt les rappels de visa SB du lot ; la démo étant construite
    # dans le désordre, on rouvre ceux des dépôts postérieurs à ce bordereau.
    connexion.execute(
        """UPDATE rappels_sb SET cloture=0, date_cloture=NULL
           WHERE lot_id=? AND date_cloture=? AND date_depot > ?""",
        (lot_id, dt.date.today().isoformat(), _il_y_a(il_y_a)))
    connexion.commit()
    connexion.close()


def charger_demo():
    """Charge le jeu de démonstration. Renvoie False (sans rien faire) si la
    base contient déjà des chefs de projet, True sinon."""
    if core.get_chefs():
        return False

    cm = core.add_chef("Christian Mahler", "c.mahler@solares-bauen.fr", "03 88 00 00 01")
    jk = core.add_chef("Julie Klein", "j.klein@solares-bauen.fr", "03 88 00 00 02")
    equipe_nord = core.add_team("Équipe Strasbourg")[1]
    equipe_sud = core.add_team("Équipe Sud Alsace")[1]
    core.set_chef_team(cm, equipe_nord)
    core.set_chef_team(jk, equipe_sud)

    # ------------------------------------------------------------------
    # Projet 1 — Groupe scolaire (2 lots).
    # ------------------------------------------------------------------
    core.add_projet("MEINAU-01", "Construction du groupe scolaire de la Meinau (14 classes)", cm,
                    "Ville de Strasbourg", "1 parc de l'Étoile\n67076 Strasbourg Cedex")
    cvc = core.add_lot("MEINAU-01", "05", "CHAUFFAGE VENTILATION", "Thermi Est",
                       "12 rue du Rhin\n67000 Strasbourg", "M.", "Thomas BERNARD",
                       "t.bernard@thermi-est.example", "06 12 34 56 78")
    pac = _position("MEINAU-01", cvc, 1, "3.1.1", "Pompe à chaleur air/eau 120 kW", "PAC air/eau",
                    avis="VSO", marque="Daikin", typeprop="EWYT120CZ", recus=5, visa_il_y_a=60,
                    marque_moe="Daikin", type_moe="EWYT-CZ ou équivalent")
    _deposer("MEINAU-01", cvc, pac, "FT_Daikin_EWYT120CZ.pdf", 50, deja_vise=True)
    _position("MEINAU-01", cvc, 2, "3.2.1", "Centrale de traitement d'air double flux salles de classe",
              "Centrale de traitement d'air double flux", avis="VAO", marque="Swegon",
              typeprop="GOLD RX 25", recus=5, visa_il_y_a=45,
              rem="Prévoir filtre ePM1 50 % au soufflage (CCTP §3.2.4).")
    cta_resto = _position("MEINAU-01", cvc, 3, "3.2.2", "Centrale de traitement d'air restauration",
                          "Centrale de traitement d'air double flux", avis="REF", marque="France Air",
                          typeprop="Ultimate 50", recus=2, maj=1, visa_il_y_a=40,
                          rem="Sélection hors CCTP : récupérateur à plaques au lieu de roue.")
    # Premier visa REF archivé, puis nouvelle soumission visée VAOB.
    core.archive_position(cta_resto)
    resto = core.get_position(cta_resto)
    resto["avis"] = "VAOB"
    resto["datevisa"] = _il_y_a(12)
    resto["rem"] = "Débit de reprise insuffisant : 4 200 m³/h au lieu de 5 000 m³/h."
    resto["rem_histo"] = f"{_il_y_a(40)} : REF — récupérateur à plaques au lieu de roue."
    core.save_position(resto)
    _deposer("MEINAU-01", cvc, cta_resto, "FT_France_Air_Ultimate50_indB.pdf", 12)   # en retard
    # Produit approchant de celui déjà visé à l'EHPAD (encart « Déjà visé ailleurs »).
    pompes = _position("MEINAU-01", cvc, 4, "3.3.1", "Pompes de circulation circuits radiateurs",
                       "Pompes de circulation", avis="MQ", marque="Grundfos",
                       typeprop="Magna3 D 40-120 F", recus=1)
    _deposer("MEINAU-01", cvc, pompes, "FT_Grundfos_Magna3.pdf", 3)                 # dans les temps
    _position("MEINAU-01", cvc, 5, "3.4.1", "Régulation et GTB chaufferie", "Analyse fonctionnelle",
              avis="MQ", marque_moe="Siemens", type_moe="Desigo PXC")
    _position("MEINAU-01", cvc, 6, "3.5.1", "Radiateurs acier panneaux", "Radiateurs panneaux à face plane",
              avis="REF", marque="Finimetal", typeprop="Reggane 3000", recus=2, visa_il_y_a=15,
              rem="Puissance calculée à ΔT 50 au lieu du régime 45/35 du CCTP.")
    _position("MEINAU-01", cvc, 7, "3.6.1", "Calorifuge des réseaux de chauffage", core.TRAME_A_DEFINIR)
    plan = _position("MEINAU-01", cvc, 8, "", "Plan d'exécution chaufferie R-1 — indice B",
                     core.TRAME_AUTRE, avis="VAO", recus=2, visa_il_y_a=20,
                     rem="Coter les vannes d'équilibrage.")
    _deposer("MEINAU-01", cvc, plan, "EXE_CVC_Chaufferie_R-1_indC.pdf", 9)         # échéance proche
    ndc = _position("MEINAU-01", cvc, 9, "", "Note de calcul des déperditions", core.TRAME_AUTRE, avis="MQ")
    _deposer("MEINAU-01", cvc, ndc, "NDC_Deperditions_indA.pdf", 1)
    _relance("MEINAU-01", cvc, 20, ["3.2.2", "3.3.1", "3.4.1", "3.5.1"])
    _bordereau("MEINAU-01", cvc, 40)
    _bordereau("MEINAU-01", cvc, 20)   # avant les dépôts en attente (sinon il les clôt)

    sanitaire = core.add_lot("MEINAU-01", "06", "PLOMBERIE SANITAIRE", "Hydro Alsace",
                             "4 rue de l'Industrie\n67400 Illkirch-Graffenstaden", "Mme",
                             "Claire MULLER", "c.muller@hydro-alsace.example", "06 98 76 54 32")
    _position("MEINAU-01", sanitaire, 1, "4.1.1", "Production ECS thermodynamique 300 L",
              core.TRAME_AUTRE, avis="VSO", marque="Atlantic", typeprop="Explorer 300",
              recus=1, visa_il_y_a=90)
    _position("MEINAU-01", sanitaire, 2, "4.2.1", "Appareils sanitaires (WC, lavabos)",
              "WC", avis="VSO", marque="Jacob Delafon", typeprop="Odéon Up",
              recus=1, visa_il_y_a=50)
    _position("MEINAU-01", sanitaire, 3, "4.3.1", "Mitigeurs thermostatiques anti-brûlure",
              "Mitigeur thermostatique centralisé ECS", avis="MQ")
    _relance("MEINAU-01", sanitaire, 5, ["4.3.1"])

    # ------------------------------------------------------------------
    # Projet 2 — CPAM : jamais relancé, un visa SB en retard.
    # ------------------------------------------------------------------
    core.add_projet("CPAM67-02", "Réhabilitation lourde du siège CPAM", cm,
                    "CPAM du Bas-Rhin", "16 rue de Lausanne\n67000 Strasbourg")
    cpam = core.add_lot("CPAM67-02", "05", "CVC", "Clim'Est", "9 rue de la Forêt\n67550 Vendenheim",
                        "M.", "Marc WEBER", "m.weber@climest.example", "06 11 22 33 44")
    _position("CPAM67-02", cpam, 1, "2.1.1", "Groupe froid 250 kW", "Production de froid",
              avis="VAOB", marque="Carrier", typeprop="30RBP 250", recus=3, maj=1, visa_il_y_a=8,
              rem="Niveau acoustique en toiture non justifié (exigence 45 dB(A) en limite de propriété).")
    _position("CPAM67-02", cpam, 2, "2.2.1", "Ventilo-convecteurs 4 tubes", "Ventilo-convecteur gainable",
              avis="VSO", marque="Aermec", typeprop="FCZ 300", recus=2, visa_il_y_a=35)
    cta_cpam = _position("CPAM67-02", cpam, 3, "2.3.1", "CTA bureaux plateaux R+1 à R+4",
                         "Centrale de traitement d'air double flux", avis="MQ")
    _deposer("CPAM67-02", cpam, cta_cpam, "FT_CTA_Wolf_KG_Top.pdf", 14)               # en retard
    _position("CPAM67-02", cpam, 4, "2.4.1", "Comptage d'énergie (option non retenue)",
              "Compteurs d'énergie thermique", avis="SO")

    # ------------------------------------------------------------------
    # Projet 3 — EHPAD (Julie Klein).
    # ------------------------------------------------------------------
    core.add_projet("EHPAD-03", "Extension de 40 lits de l'EHPAD Les Tilleuls", jk,
                    "Fondation Les Tilleuls", "5 allée des Tilleuls\n67400 Illkirch-Graffenstaden")
    ehpad = core.add_lot("EHPAD-03", "05", "CVC PLOMBERIE", "Gaz & Fluides SARL",
                         "22 route de Lyon\n67400 Illkirch-Graffenstaden", "M.", "Julien HOFFMANN",
                         "j.hoffmann@gazfluides.example", "06 55 44 33 22")
    _position("EHPAD-03", ehpad, 1, "5.1.1", "Chaudière gaz à condensation 2 × 150 kW",
              "Chaudière gaz à condensation", avis="VSO", marque="Viessmann", typeprop="Vitocrossal 200",
              recus=5, visa_il_y_a=120)
    # Même CTA que celle de la Meinau : chaque position voit le visa de l'autre.
    _position("EHPAD-03", ehpad, 2, "5.2.1", "CTA double flux circulations",
              "Centrale de traitement d'air double flux", avis="VAO", marque="Swegon", typeprop="Gold RX 25",
              recus=5, visa_il_y_a=70, rem="Fournir le certificat Eurovent de la roue de récupération.")
    _position("EHPAD-03", ehpad, 3, "5.3.1", "Pompes doubles circuit plancher chauffant",
              "Pompes de circulation", avis="VSO", marque="Grundfos", typeprop="Magna3 D 40-120",
              recus=3, visa_il_y_a=55)
    plancher = _position("EHPAD-03", ehpad, 4, "5.4.1", "Plancher chauffant basse température", "Tube PER BAO",
                         avis="REF", marque="Rehau", typeprop="Rautherm S", recus=1, visa_il_y_a=9,
                         rem="Pas d'étude de calepinage ni de note de puissance par local.")
    _deposer("EHPAD-03", ehpad, plancher, "Calepinage_Rehau_RDC.pdf", 2)
    _relance("EHPAD-03", ehpad, 10, ["5.4.1"])
    _bordereau("EHPAD-03", ehpad, 9)

    # ------------------------------------------------------------------
    # Projet 4 — archivé (livré).
    # ------------------------------------------------------------------
    core.add_projet("HALLE-00", "Halle des sports (livrée)", jk, "Commune d'Eckbolsheim",
                    "1 place de la Mairie\n67201 Eckbolsheim")
    halle = core.add_lot("HALLE-00", "05", "CVC", "Thermi Est", "12 rue du Rhin\n67000 Strasbourg",
                         "M.", "Thomas BERNARD")
    _position("HALLE-00", halle, 1, "3.1.1", "Aérothermes gaz", core.TRAME_AUTRE, avis="VSO",
              marque="Reznor", typeprop="UDAP 45", recus=1, visa_il_y_a=400)
    core.set_projet_archive("HALLE-00", 1)

    # Quelques validations mémorisées pour la suggestion automatique de trame.
    core.record_assoc("Calorifuge des réseaux eau glacée", "Calorifuge des tuyauteries – élastomère")
    core.record_assoc("Radiateurs acier panneaux", "Radiateurs panneaux à face plane")
    ajouter_portail_demo()
    ajouter_historique_demo()
    return True


# ----------------------------------------------------------------------------
# Portail Plateforme visa partenaires : comptes entreprises (identifiant, mot de passe de démo),
# chargés d'affaires, envois en attente et un envoi refusé.
# ----------------------------------------------------------------------------
COMPTES_DEMO = {
    "Thermi Est": ("thermi-est", "demo-thermi", [("M.", "Thomas BERNARD", "t.bernard@thermi-est.example"),
                                                 ("Mme", "Sophie MULLER", "s.muller@thermi-est.example")]),
    "Clim'Est": ("climest", "demo-climest", [("M.", "Marc WEBER", "m.weber@climest.example")]),
    "Gaz & Fluides SARL": ("gazfluides", "demo-gazfluides",
                           [("M.", "Julien HOFFMANN", "j.hoffmann@gazfluides.example")]),
}


def pdf_pages(titre, nb_pages):
    """PDF de démonstration de plusieurs pages (« titre — p. 1 », « p. 2 »...)."""
    import io
    from pypdf import PdfReader, PdfWriter
    redacteur = PdfWriter()
    for numero in range(1, nb_pages + 1):
        redacteur.add_page(PdfReader(io.BytesIO(pdf_exemple(f"{titre} - p. {numero}"))).pages[0])
    sortie = io.BytesIO()
    redacteur.write(sortie)
    return sortie.getvalue()


def _dater_envoi(depot_id, il_y_a_ouvres, heure):
    """Recule la date d'un envoi de démonstration (N jours ouvrés, à l'heure donnée)."""
    connexion = core.conn()
    connexion.execute("UPDATE depots SET date_envoi=? WHERE id=?",
                      (_il_y_a_ouvres(il_y_a_ouvres) + "T" + heure, depot_id))
    connexion.commit()
    connexion.close()


# ----------------------------------------------------------------------------
# Historique de visas sur 12 mois, pour le Dashboard « Solares » : trois
# documents visés par mois, dont le délai (jours ouvrés du dépôt au bordereau)
# raccourcit au fil de l'année. Chaque visa est un rappel « à viser avant le »
# clos, rattaché à une position des projets de démo.
# ----------------------------------------------------------------------------
PROJETS_DEMO = ("MEINAU-01", "EHPAD-03", "CPAM67-02", "HALLE-00")
DELAIS_HISTORIQUE = [   # un mois par ligne, du plus ancien (il y a 12 mois) au plus récent
    [8, 14, 12], [6, 13, 9], [11, 7, 15], [5, 9, 12], [8, 4, 10], [6, 11, 3],
    [9, 5, 7], [4, 8, 12], [7, 3, 6], [5, 9, 4], [6, 2, 8], [4, 7, 5]]
POSITIONS_HISTORIQUE = [("MEINAU-01", "05", "3.1.1"), ("EHPAD-03", "05", "5.1.1"),
                        ("CPAM67-02", "05", "2.2.1"), ("MEINAU-01", "06", "4.1.1"),
                        ("EHPAD-03", "05", "5.2.1"), ("HALLE-00", "05", "3.1.1"),
                        ("MEINAU-01", "05", "3.2.1"), ("EHPAD-03", "05", "5.3.1"),
                        ("MEINAU-01", "06", "4.2.1")]


def ajouter_historique_demo():
    """Ajoute l'historique de visas sur 12 mois aux projets de démo. Sans effet
    s'il est déjà là (un visa de plus de 4 mois sur un projet de démo) ou si
    les projets de démo n'existent pas. Renvoie le nombre de visas ajoutés."""
    connexion = core.conn()
    marques = ",".join("?" for _ in PROJETS_DEMO)
    deja_la = connexion.execute(
        f"SELECT 1 FROM rappels_sb WHERE cloture=1 AND projet IN ({marques}) AND date_depot < ?",
        PROJETS_DEMO + (_il_y_a(120),)).fetchone()
    connexion.close()
    if deja_la:
        return 0

    aujourd_hui = dt.date.today()
    lignes = []
    rang_position = 0
    for numero_mois, delais in enumerate(DELAIS_HISTORIQUE):
        mois_avant = len(DELAIS_HISTORIQUE) - numero_mois
        for rang, delai in enumerate(delais):
            projet, numero_lot, reference = POSITIONS_HISTORIQUE[rang_position % len(POSITIONS_HISTORIQUE)]
            rang_position = rang_position + 1
            lot, position = _position_par_ref(projet, numero_lot, reference)
            if position is None:
                continue
            # Dépôt vers le milieu du mois concerné, un jour ouvré.
            depot = aujourd_hui - dt.timedelta(days=mois_avant * 30 - 15 + rang * 8)
            while not calendrier.est_ouvre(depot):
                depot = depot - dt.timedelta(days=1)
            visa = calendrier.ajouter_jours_ouvres(depot, delai)
            if visa >= aujourd_hui:
                continue
            delai_du_lot = int(lot.get("delai_visa_sb") or core.DELAI_VISA_SB)
            echeance = calendrier.ajouter_jours_ouvres(depot, delai_du_lot)
            fichier = f"FT_{reference}_ind{'ABC'[rang]}.pdf"
            lignes.append((projet, lot["id"], position["id"], depot.isoformat(), echeance.isoformat(),
                           json.dumps([fichier]), visa.isoformat()))

    connexion = core.conn()
    for ligne in lignes:
        connexion.execute(
            """INSERT INTO rappels_sb(projet, lot_id, position_id, date_depot, echeance, fichiers,
                                      cloture, date_cloture) VALUES(?,?,?,?,?,?,1,?)""", ligne)
    connexion.commit()
    connexion.close()
    return len(lignes)


def _position_par_ref(projet, numero_lot, reference):
    for lot in core.get_lots(projet):
        if lot.get("lot_numero") == numero_lot:
            for position in core.get_positions(projet, lot["id"]):
                if position["pos"] == reference:
                    return lot, position
    return None, None


def ajouter_portail_demo():
    """Ajoute à la démo les comptes du portail et quelques envois. Sans effet si
    un compte portail existe déjà. Renvoie True si quelque chose a été ajouté."""
    entreprises = {}
    for fiche in core.get_entreprises():
        if fiche["a_compte"]:
            return False
        entreprises[fiche["nom"]] = fiche["id"]
    contacts = {}
    for nom, (identifiant, mot_de_passe, personnes) in COMPTES_DEMO.items():
        if nom not in entreprises:
            continue
        core.set_compte_entreprise(entreprises[nom], identifiant, mot_de_passe)
        ids_contacts = []
        for civilite, nom_contact, email in personnes:
            ok, contact_id = core.add_contact(entreprises[nom], nom_contact, email, civilite)
            if ok:
                contacts[nom_contact] = contact_id
                ids_contacts.append(contact_id)
        # Chargé d'affaires de chaque lot : le premier contact pour le premier lot,
        # le suivant pour le lot suivant (Thermi Est : un chargé d'affaires par lot).
        if ids_contacts:
            for rang, lot in enumerate(core.get_lots_entreprise(entreprises[nom])):
                core.set_lot_contact(lot["id"], ids_contacts[rang % len(ids_contacts)])

    # Thermi Est, MEINAU lot 05 : envoi de la régulation GTB (3 fichiers de types différents).
    lot, gtb = _position_par_ref("MEINAU-01", "05", "3.4.1")
    if gtb:
        ok, depot = core.enregistrer_envoi(entreprises["Thermi Est"], lot["id"], [
            {"position_id": gtb["id"], "pieces": ["Analyse fonctionnelle", "Schéma d'architecture"],
             "commentaire": "Synoptiques provisoires : version définitive après mise en service.",
             "fichiers": [{"nom": "AF_GTB_chaufferie_indA.pdf", "contenu": pdf_pages("Analyse fonctionnelle GTB", 6)},
                          {"nom": "Liste_points_GTB_indA.xlsx", "contenu": b"Liste des points (demo)"},
                          {"nom": "Synoptiques_GTB_indA.pdf", "contenu": pdf_pages("Synoptiques GTB", 3)}]}])
        if ok:
            _dater_envoi(depot, 1, "16:42:00")
        # Radiateurs (visa REF) : un envoi refusé, avec motif.
        lot, radiateurs = _position_par_ref("MEINAU-01", "05", "3.5.1")
        ok, depot = core.enregistrer_envoi(entreprises["Thermi Est"], lot["id"], [
            {"position_id": radiateurs["id"], "pieces": ["Note de sélection"],
             "fichiers": [{"nom": "Selection_Reggane_indB.pdf", "contenu": pdf_pages("Selection radiateurs", 2)}]}])
        if ok:
            _dater_envoi(depot, 3, "09:15:00")
            for envoi in core.get_depots_en_attente():
                if envoi["id"] == depot:
                    core.refuser_lignes([envoi["lignes"][0]["id"]], "La sélection est toujours calculée "
                                        "à ΔT 50 : merci de la refaire au régime 45/35 du CCTP.", "CM")

    # Clim'Est, CPAM lot 05 : un dossier technique commun pour deux positions.
    lot, groupe_froid = _position_par_ref("CPAM67-02", "05", "2.1.1")
    lot, ventilos = _position_par_ref("CPAM67-02", "05", "2.2.1")
    if groupe_froid and ventilos:
        ok, depot = core.enregistrer_envoi(entreprises["Clim'Est"], lot["id"], [
            {"position_id": groupe_froid["id"], "pieces": ["Fiche technique"],
             "commentaire": "Note acoustique en limite de propriété ajoutée p. 5-6.",
             "extraits": [{"nom": "Dossier_technique_CVC_indA.pdf", "extrait": "2-6"}]},
            {"position_id": ventilos["id"], "pieces": ["Note de sélection"],
             "extraits": [{"nom": "Dossier_technique_CVC_indA.pdf", "extrait": "7-9"}]},
        ], [{"nom": "Dossier_technique_CVC_indA.pdf", "contenu": pdf_pages("Dossier technique CVC", 12)}])
        if ok:
            _dater_envoi(depot, 2, "10:05:00")
        # Clim'Est, CPAM lot 05 : un document non prévu dans la liste (soumission libre).
        ok, depot = core.enregistrer_envoi(entreprises["Clim'Est"], lot["id"], [
            {"libre": {"type": core.LIBRE_CCTP, "pos": "HC", "des": "Rideau d'air chaud du hall d'entrée"},
             "pieces": ["Fiche technique"], "commentaire": "Ajout demandé par le maître d'ouvrage en réunion.",
             "fichiers": [{"nom": "FT_Rideau_air_hall.pdf", "contenu": pdf_pages("Rideau d'air chaud", 2)}]}])
        if ok:
            _dater_envoi(depot, 1, "11:30:00")
    return True
