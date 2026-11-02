# ============================================================================
#  ui/style.py — charte Solares Bauen et éléments visuels communs
#  (badges d'avis, en-tête de page, cartes d'indicateurs, infobulles).
#
#  Charte : rouge Solares Bauen #E60116 (accent) et #C00000 (foncé),
#  gris neutres, police Inter (proche d'Arial, plus moderne ; déclarée et
#  hébergée via .streamlit/config.toml et static/, Arial en secours ; le
#  bordereau Excel / PDF garde Arial). Les couleurs d'avis restent sémantiques
#  (VSO vert, REF rouge...) mais dans des teintes accordées à la charte.
# ============================================================================

import base64
import html
import os
import streamlit as st

# Logo de chaque plateforme (Solares Bauen + « Plateforme visa interne » pour
# Plateforme visa interne, « partenaires » pour la Plateforme visa partenaires), dérivé des SVG fournis : fond blanc
# rendu transparent, cadrage resserré sur le contenu (bord gauche aligné sur
# les sélecteurs de la barre latérale).
LOGOS = {
    "interne": (os.path.join(os.path.dirname(__file__), "logo_visa_interne.svg"),
                "Solares Bauen — Plateforme visa interne"),
    "partenaires": (os.path.join(os.path.dirname(__file__), "logo_visa_partenaires.svg"),
                    "Solares Bauen — Plateforme visa partenaires"),
}
_LOGOS_ENCODES = {}   # contenu base64 de chaque logo, lu une seule fois

ROUGE = "#E60116"
ROUGE_FONCE = "#C00000"

CSS = """<style>
:root{
  --sb-rouge:#E60116; --sb-rouge-fonce:#C00000; --sb-rouge-pale:#FDECEE;
  /* gris-clair : le plus clair autorisé pour du texte (contraste ≥ 4,5:1 sur le fond gris) */
  --sb-texte:#1E1E1E; --sb-gris:#5C5C5C; --sb-gris-clair:#6B6B6B;
  --sb-ligne:#E4E4E4; --sb-fond:#F6F6F6; --sb-blanc:#FFFFFF;
}
/* Inter partout (Arial en secours), sauf les icônes Material (toutes marquées translate="no").
   Taille optique automatique : Inter s'affine d'elle-même en petite taille. */
html, body, [class*="st-"]:not([translate="no"]), .stMarkdown, input, textarea, select,
button, p, div, label, span:not([translate="no"]){
  font-family: Inter, Arial, "Helvetica Neue", Helvetica, sans-serif;font-optical-sizing:auto;}
.block-container{padding-top:1.4rem;max-width:1280px;}
/* Portable (barre latérale ouverte) : marges latérales réduites, sinon à 1024 px il ne reste que
   550 px utiles et les boutons des lignes sont tronqués (« Tout décoc… »). */
@media (max-width:1366px){
  [data-testid="stMainBlockContainer"]{padding-left:1.5rem;padding-right:1.5rem;}}
/* Bouton discret (n° de bordereau) : passe à la ligne plutôt que d'être coupé. */
.stButton>button[kind="tertiary"] p{white-space:normal;overflow:visible;text-overflow:clip;}
h1,h2,h3{color:var(--sb-texte);letter-spacing:-.01em;}

/* ---------- surfaces natives du navigateur et de Streamlit (lisibilité) ---------- */
::selection{background:var(--sb-rouge-pale);color:var(--sb-texte);}
input, textarea{caret-color:var(--sb-rouge-fonce);}
/* Légendes Streamlit : gris franc plutôt qu'un texte à 60 % d'opacité (trop pâle). */
div[data-testid="stCaptionContainer"]{opacity:1;color:var(--sb-gris);}
input::placeholder, textarea::placeholder{color:var(--sb-gris-clair) !important;opacity:1;}
/* Pas de mot isolé en fin de paragraphe. */
.stMarkdown p, div[data-testid="stCaptionContainer"] p, div[data-testid="stAlert"] p{text-wrap:pretty;}
/* Choix actif des boutons groupés (Portefeuille, avis) : rouge foncé, lisible sur le fond rosé. */
div[data-testid="stButtonGroup"] button[aria-checked="true"],
div[data-testid="stButtonGroup"] button[aria-pressed="true"],
div[data-testid="stButtonGroup"] button[kind*="Active"]{color:var(--sb-rouge-fonce) !important;}
div[data-testid="stButtonGroup"] button[aria-checked="true"] p,
div[data-testid="stButtonGroup"] button[aria-pressed="true"] p,
div[data-testid="stButtonGroup"] button[kind*="Active"] p{color:var(--sb-rouge-fonce) !important;}
/* Boutons groupés de la page (pièces, mode d'envoi) : un libellé long passe à la ligne
   au lieu d'être coupé par des points de suspension. */
[data-testid="stMain"] div[data-testid="stButtonGroup"] button{height:auto;max-width:100%;}
[data-testid="stMain"] div[data-testid="stButtonGroup"] button div,
[data-testid="stMain"] div[data-testid="stButtonGroup"] button p{white-space:normal;overflow:visible;text-overflow:clip;}
/* Messages de succès : vert plus foncé que celui de Streamlit (contraste ≥ 4,5:1). */
div[data-testid="stAlertContentSuccess"], div[data-testid="stAlertContentSuccess"] p{color:#0F6B2C;}

/* ---------- barre latérale : marque + menu principal ---------- */
section[data-testid="stSidebar"]{background:var(--sb-fond);border-right:1px solid var(--sb-ligne);}
section[data-testid="stSidebar"] .block-container, section[data-testid="stSidebarContent"]{padding-top:.6rem;}
/* Portefeuille (Chef de projet / Équipe / Solares) : les trois choix sur une seule ligne. */
section[data-testid="stSidebar"] [data-testid="stButtonGroup"] [role="radiogroup"]{flex-wrap:nowrap;}
section[data-testid="stSidebar"] [data-testid="stButtonGroup"] button{flex:1 1 auto;padding:4px 8px;}
/* Logo de la plateforme : toute la largeur de la barre (les petites lignes du logo
   restent ≥ 12 px), bord gauche aligné sur les sélecteurs, filet de séparation dessous. */
.sb-marque{padding:2px 0 10px 0;margin:0 0 6px 0;border-bottom:1px solid var(--sb-ligne);}
/* Écran de connexion : le logo descend sous l'en-tête natif de la page, qui en masquait le haut. */
[data-testid="stMain"] .sb-marque{margin-top:2rem;}
/* En-tête natif de la barre (bouton de repli, à droite) resserré : le logo remonte. */
section[data-testid="stSidebar"] [data-testid="stSidebarHeader"]{height:2.75rem;}
.sb-marque img{display:block;width:100%;max-width:300px;height:auto;aspect-ratio:304.7/161.9;}
/* menu principal = st.radio restylé : grandes entrées, sans bouton rond */
section[data-testid="stSidebar"] [data-testid="stRadioGroup"]{gap:2px;width:100%;}
section[data-testid="stSidebar"] [data-testid="stRadioGroup"] > div{width:100%;}
section[data-testid="stSidebar"] label[data-testid="stRadioOption"]{
  width:100%;margin:0;padding:8px 12px;border-radius:10px;
  transition:background .15s ease-out;cursor:pointer;}
section[data-testid="stSidebar"] label[data-testid="stRadioOption"] > div > div:first-child{display:none;}
section[data-testid="stSidebar"] label[data-testid="stRadioOption"] p{
  font-size:16px;font-weight:600;color:#333;line-height:1.3;}
section[data-testid="stSidebar"] label[data-testid="stRadioOption"] p span[role="img"]{
  margin-right:10px;font-size:20px;color:#6B6B6B;}
section[data-testid="stSidebar"] label[data-testid="stRadioOption"]:hover{background:#ECECEC;}
/* Page active : carte blanche, texte rouge foncé, icône rouge (sans bande colorée à gauche). */
section[data-testid="stSidebar"] label[data-testid="stRadioOption"]:has(input:checked){
  background:var(--sb-blanc);box-shadow:0 1px 4px rgba(0,0,0,.10);}
section[data-testid="stSidebar"] label[data-testid="stRadioOption"]:has(input:checked) p{color:var(--sb-rouge-fonce);}
section[data-testid="stSidebar"] label[data-testid="stRadioOption"]:has(input:checked) p span[role="img"]{color:var(--sb-rouge);}
section[data-testid="stSidebar"] label[data-testid="stRadioOption"]:has(input:focus-visible){
  outline:2px solid var(--sb-rouge-fonce);outline-offset:2px;}
.sb-titre-nav{font-size:12px;font-weight:700;color:var(--sb-gris-clair);text-transform:uppercase;
  letter-spacing:.08em;margin:10px 0 6px 2px;}
/* Barre resserrée : portefeuille, projet, lot et menu visibles sans défiler sur un écran courant. */
section[data-testid="stSidebar"] [data-testid="stVerticalBlock"]{gap:.55rem;}
section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"]{margin-bottom:0;}
section[data-testid="stSidebar"] [data-testid="stDivider"] hr, section[data-testid="stSidebar"] hr{margin:.35rem 0;}
.sb-pied{font-size:12px;color:var(--sb-gris);text-align:right;}

/* ---------- en-tête de page : fil, titre, grille libellé / valeur ---------- */
.entete{padding:2px 0 20px 0;margin-bottom:20px;border-bottom:1px solid var(--sb-ligne);}
/* Fil d'Ariane (code du projet › lot ou page) : seulement quand il apporte une information. */
.ent-fil{display:flex;flex-wrap:wrap;align-items:center;gap:8px;font-size:13px;font-weight:600;
  color:var(--sb-gris);margin-bottom:8px;}
.ent-fil .sep{color:var(--sb-gris-clair);font-weight:400;}
.ent-code{color:var(--sb-rouge-fonce);background:var(--sb-rouge-pale);border-radius:5px;padding:2px 8px;
  font-size:12.5px;font-weight:700;letter-spacing:.02em;}
.ent-titre{font-size:28px;font-weight:700;line-height:1.2;color:var(--sb-texte);
  text-wrap:balance;}
.ent-sous{font-size:14px;color:var(--sb-gris);margin-top:6px;}
.ent-meta{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:14px 32px;margin-top:18px;}
.ent-meta .lab{display:block;font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;
  color:var(--sb-gris-clair);}
.ent-meta .val{display:block;font-size:15px;font-weight:600;color:var(--sb-texte);margin-top:4px;}

/* ---------- conteneurs bordés ---------- */
div[data-testid="stVerticalBlockBorderWrapper"]{border-radius:12px;border-color:var(--sb-ligne) !important;
  background:var(--sb-blanc);}

/* ---------- tuiles d'indicateurs : même hauteur, chiffre en tête ---------- */
.tuiles{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:14px;margin:0 0 16px 0;}
.tuile{background:var(--sb-blanc);border:1px solid var(--sb-ligne);border-radius:12px;padding:16px 18px;
  display:flex;flex-direction:column;gap:6px;min-height:118px;}
.t-lab{font-size:12px;font-weight:700;color:var(--sb-gris);text-transform:uppercase;letter-spacing:.06em;}
.t-val{font-size:32px;font-weight:700;line-height:1.05;color:var(--sb-texte);font-variant-numeric:tabular-nums;}
.t-val small{font-size:15px;font-weight:600;color:var(--sb-gris-clair);margin-left:3px;}
.t-sub{font-size:12.5px;color:var(--sb-gris);margin-top:auto;line-height:1.4;}
.tuile.alerte{border-color:#F2B9C0;background:#FFF7F8;}
.tuile.alerte .t-val{color:var(--sb-rouge);}
.mini-jauge{display:inline-flex;align-items:center;gap:8px;vertical-align:middle;white-space:nowrap;}
.mini-jauge .mj-piste{display:inline-block;width:72px;height:6px;border-radius:3px;background:#EDEDED;overflow:hidden;}
.mini-jauge .mj-rempli{display:block;height:100%;background:#4C9A5B;border-radius:3px;}
.mini-jauge .mj-val{font-size:13px;font-weight:700;color:var(--sb-texte);font-variant-numeric:tabular-nums;}
.mini-jauge .mj-lab{font-size:12px;color:var(--sb-gris);}
.mini-jauge.grande{display:flex;width:100%;gap:10px;}
.mini-jauge.grande .mj-piste{flex:1;height:12px;border-radius:6px;}
.mini-jauge.grande .mj-rempli{border-radius:6px;}
.mini-jauge.grande .mj-val{font-size:20px;}
.mini-jauge.grande .mj-lab{font-size:13px;}
.t-val.petit{font-size:16px;}
.t-val.petit small{font-size:12px;}
.st-key-actions_lot button{min-height:58px;padding:8px 12px;}
.st-key-actions_lot button p, .st-key-actions_lot button div{white-space:normal !important;
  overflow:visible !important;text-overflow:clip !important;line-height:1.25;text-align:center;}
.barre-actions{font-size:12px;font-weight:700;color:var(--sb-gris);text-transform:uppercase;
  letter-spacing:.06em;margin:2px 0 6px 0;}

/* ---------- badges ---------- */
.bdg{display:inline-block;padding:3px 11px;border-radius:999px;font-size:12px;font-weight:700;
  letter-spacing:.02em;white-space:nowrap;max-width:100%;vertical-align:middle;}
/* Pastille d'un libellé libre (pièce déclarée, pièce d'un document) : elle passe à la ligne
   dans une colonne étroite au lieu de déborder sur la voisine. */
.bdg.long{white-space:normal;overflow-wrap:anywhere;border-radius:10px;line-height:1.35;
  text-align:left;margin:2px 0;}
/* Nom de fichier (souvent sans espace) : coupé où il faut plutôt que de déborder. */
.nom-fichier{font-size:13px;line-height:1.45;overflow-wrap:anywhere;}
.nom-fichier .bdg{margin-top:4px;}
.b-VSO{background:#E3F1E4;color:#1E5E2E;} .b-VAO{background:#FFF3CD;color:#7A5A00;}
.b-VAOB{background:#FDE3D0;color:#9A3D0A;} .b-REF{background:#FBDADA;color:#A50010;}
.b-MQ{background:#ECE6F6;color:#4B2E83;} .b-SO{background:#EBEBEB;color:#4D4D4D;}
.b-retard{background:var(--sb-rouge);color:#fff;} .b-proche{background:#FDE3D0;color:#9A3D0A;}
.b-ok{background:#E3F1E4;color:#1E5E2E;} .b-neutre{background:#EFEFEF;color:#4D4D4D;}
.b-sb{background:var(--sb-rouge-pale);color:var(--sb-rouge-fonce);}
/* Étiquette longue dans une colonne étroite : elle passe à la ligne au lieu de déborder. */
.survol .bdg{white-space:normal;border-radius:10px;line-height:1.35;max-width:100%;}
.survol{max-width:100%;}

/* ---------- lignes de positions ---------- */
.pos-ligne{display:flex;align-items:center;gap:14px;min-height:40px;}
.pos-texte{flex:1;min-width:0;font-size:15px;line-height:1.45;color:var(--sb-texte);}
.pos-ref{font-weight:700;margin-right:6px;}
.pos-prop{color:var(--sb-gris);}
.pos-indices{display:flex;gap:6px;flex-wrap:wrap;margin-top:3px;}
.pastille{display:inline-block;font-size:12px;padding:1px 8px;border-radius:6px;background:#F1F1F1;color:var(--sb-gris);}
/* N° de visa sous l'avis : passe à la ligne dans la colonne étroite (« visa en cours »). */
.num-visa{font-size:12px;font-weight:700;color:var(--sb-gris);margin-top:3px;line-height:1.3;}
.num-visa.gris{font-weight:400;color:var(--sb-gris-clair);font-style:italic;}
.pastille.rouge{background:var(--sb-rouge-pale);color:var(--sb-rouge-fonce);font-weight:600;}
.dot{display:inline-block;width:9px;height:9px;border-radius:50%;margin-right:6px;vertical-align:middle;}
.rowsep{margin:-10px 0 -6px 0;border:none;border-top:1px solid #EEEEEE;}
.sous-titre{font-size:13px;font-weight:700;color:var(--sb-texte);text-transform:uppercase;
  letter-spacing:.06em;margin:14px 0 6px 0;padding-bottom:4px;border-bottom:2px solid var(--sb-rouge);
  display:inline-block;}
.note-ro{font-size:12.5px;color:var(--sb-gris);background:#F4F4F4;border-radius:8px;padding:6px 10px;
  margin-bottom:6px;}
/* Encart « Déjà visé ailleurs » du détail d'une position (Suivi). */
.dv-titre{font-size:13px;font-weight:700;color:var(--sb-texte);margin-bottom:4px;}
.dv-info{width:15px;height:15px;color:#1D4F91;margin-right:6px;vertical-align:-3px;}
.dv-recul{display:inline-block;font-size:12px;color:#8A5A00;background:#FFF2D6;border-radius:6px;
  padding:3px 8px;margin:2px 0 6px 0;}
.dv-item{padding:8px 0;border-top:1px solid #EEEEEE;}
.dv-titre + .dv-item, .dv-recul + .dv-item{border-top:none;}
.dv-ligne{font-size:14px;line-height:1.5;color:var(--sb-texte);}
.dv-sous{font-size:12.5px;line-height:1.45;color:var(--sb-gris);margin-top:2px;}

/* ---------- volets : Suivi (Positions CCTP/Matériel / Plans-NdC / Anciens visas) et Administration ----------
   st.segmented_control natif, dans un conteneur de clé « volets_… », restylé en grands boutons
   séparés, impossibles à rater. Le nombre d'éléments des volets du Suivi est posé par
   page_suivi (p::after, content) ; sans content, pas de pastille. */
[class*="st-key-volets_"] [role="radiogroup"]{gap:10px;}
[class*="st-key-volets_"] [role="radiogroup"] button{min-height:56px;height:auto;margin:0;padding:8px 16px;
  border-radius:10px;border:1px solid #CFCFCF;background:var(--sb-blanc);
  box-shadow:0 1px 2px rgba(0,0,0,.06);transition:background .15s ease-out,border-color .15s ease-out;}
[class*="st-key-volets_"] [role="radiogroup"] button > div{overflow:visible;}
[class*="st-key-volets_"] [role="radiogroup"] button p{display:inline-flex;align-items:center;gap:10px;
  overflow:visible;font-size:17px;font-weight:700;color:var(--sb-texte);}
[class*="st-key-volets_"] [role="radiogroup"] button [data-testid="stIconMaterial"]{font-size:22px;color:var(--sb-gris);}
[class*="st-key-volets_"] [role="radiogroup"] button p::after{display:inline-block;min-width:30px;padding:2px 9px;
  border-radius:999px;background:#EDEDED;color:var(--sb-texte);font-size:14px;font-weight:700;line-height:1.4;
  text-align:center;font-variant-numeric:tabular-nums;}
[class*="st-key-volets_"] [role="radiogroup"] button:hover{border-color:var(--sb-rouge);background:#FFF7F8;}
/* Volet actif : fond rouge SB, texte blanc, compteur sur pastille blanche. */
[class*="st-key-volets_"] [role="radiogroup"] button[aria-checked="true"]{background:var(--sb-rouge);
  border-color:var(--sb-rouge);box-shadow:0 2px 6px rgba(230,1,22,.25);}
[class*="st-key-volets_"] [role="radiogroup"] button[aria-checked="true"]:hover{background:var(--sb-rouge-fonce);
  border-color:var(--sb-rouge-fonce);}
[class*="st-key-volets_"] [role="radiogroup"] button[aria-checked="true"] p,
[class*="st-key-volets_"] [role="radiogroup"] button[aria-checked="true"] [data-testid="stIconMaterial"]{color:#fff !important;}
[class*="st-key-volets_"] [role="radiogroup"] button[aria-checked="true"] p::after{background:#fff;color:var(--sb-rouge-fonce);}
[class*="st-key-volets_"] [role="radiogroup"] button:focus-visible{outline:2px solid var(--sb-rouge-fonce);outline-offset:2px;}
/* Écran étroit : les volets restent sur une ligne. */
@media (max-width:1180px){
  [class*="st-key-volets_"] [role="radiogroup"]{gap:8px;}
  [class*="st-key-volets_"] [role="radiogroup"] button{padding:8px 10px;}
  [class*="st-key-volets_"] [role="radiogroup"] button p{font-size:15px;gap:8px;}
  [class*="st-key-volets_"] [role="radiogroup"] button [data-testid="stIconMaterial"]{font-size:20px;}}

/* ---------- infobulle au survol (dashboard) ---------- */
.survol{position:relative;display:inline-block;cursor:help;}
.survol .bulle{visibility:hidden;opacity:0;position:absolute;z-index:50;left:0;top:calc(100% + 6px);
  min-width:320px;max-width:460px;background:#1E1E1E;color:#fff;border-radius:10px;padding:10px 12px;
  box-shadow:0 8px 24px rgba(0,0,0,.25);font-size:12.5px;line-height:1.5;transition:opacity .12s ease;}
/* Au survol, et au clavier (Tab) : l'étiquette est focusable. */
.survol:hover .bulle, .survol:focus .bulle, .survol:focus-within .bulle{visibility:visible;opacity:1;}
.survol:focus-visible{outline:2px solid var(--sb-rouge-fonce);outline-offset:2px;border-radius:999px;}
.bulle .b-item{display:block;margin:4px 0;padding-left:12px;text-indent:-12px;}
.bulle .b-titre{display:block;font-weight:700;margin-bottom:6px;color:#fff;}
.bulle .b-gris{color:#BDBDBD;}

/* ---------- onglets, boutons ---------- */
button[data-baseweb="tab"]{font-weight:600;font-size:15px;}
button[data-baseweb="tab"][aria-selected="true"]{color:var(--sb-rouge) !important;}
div[data-baseweb="tab-highlight"]{background-color:var(--sb-rouge) !important;}
.stButton>button, .stDownloadButton>button, .stFormSubmitButton>button{border-radius:8px;font-weight:600;}
.stButton>button:hover, .stDownloadButton>button:hover{border-color:var(--sb-rouge);color:var(--sb-rouge);}
.stButton>button[kind="primary"], .stDownloadButton>button[kind="primary"], .stFormSubmitButton>button[kind="primary"]{
  background:var(--sb-rouge);border-color:var(--sb-rouge);color:#fff;}
.stButton>button[kind="primary"]:hover, .stDownloadButton>button[kind="primary"]:hover{
  background:var(--sb-rouge-fonce);border-color:var(--sb-rouge-fonce);color:#fff;}
.stButton>button:disabled, .stButton>button[kind="primary"]:disabled{opacity:.45;cursor:not-allowed;}
div[data-testid="stCheckbox"] label{white-space:nowrap;}
/* Case « Au bordereau » des lignes du Suivi : sur deux lignes plutôt que coupée (portable). */
[class*="st-key-exp_"] label, [class*="st-key-exp_"] label *{white-space:normal;overflow:visible;text-overflow:clip;}
div[data-testid="stCheckbox"] p{font-size:13px;color:var(--sb-gris);}
</style>"""

# Couleur de pastille par statut de pièce attendue.
STAT_KIND = {"Reçu et ok": "green", "À vérifier": "blue", "MAJ attendue": "amber", "Manquant": "red"}
STAT_COLOR = {"green": "#2E7D32", "blue": "#1F6FB2", "amber": "#C77700", "red": "#E60116", "gray": "#9E9E9E"}


def esc(texte):
    """Échappe un texte saisi avant de l'afficher en HTML (contenu de balise
    uniquement : les apostrophes restent intactes, car le Markdown de
    Streamlit affiche mal « &#x27; »)."""
    if texte is None:
        return ""
    return html.escape(str(texte), quote=False)


def marque(plateforme):
    """Bloc de marque (barre latérale, écran de connexion) : le logo de la
    plateforme, "interne" (Plateforme visa interne) ou "partenaires" (Plateforme visa partenaires)."""
    chemin, texte_alternatif = LOGOS[plateforme]
    if plateforme not in _LOGOS_ENCODES:
        donnees = ""
        if os.path.exists(chemin):
            with open(chemin, "rb") as fichier:
                donnees = base64.b64encode(fichier.read()).decode("ascii")
        _LOGOS_ENCODES[plateforme] = donnees
    if not _LOGOS_ENCODES[plateforme]:
        return f"<div class='sb-marque'><b>{esc(texte_alternatif)}</b></div>"
    return (f"<div class='sb-marque'><img src='data:image/svg+xml;base64,{_LOGOS_ENCODES[plateforme]}' "
            f"alt='{esc(texte_alternatif)}'></div>")


def appliquer_css():
    """Injecte la feuille de style commune (à appeler une fois par affichage)."""
    st.markdown(CSS, unsafe_allow_html=True)


def badge_avis(avis):
    """Pastille colorée d'un avis (VSO, VAO, VAOB, REF, MQ, SO)."""
    return f"<span class='bdg b-{esc(avis)}'>{esc(avis)}</span>"


def badge(texte, genre, long=False):
    """Pastille libre : genre = retard, proche, ok, neutre, sb.
    long = True : elle passe à la ligne dans une colonne étroite au lieu de déborder."""
    classes = f"bdg long b-{genre}" if long else f"bdg b-{genre}"
    return f"<span class='{classes}'>{texte}</span>"


def entete(code, titre, meta=None, fil=""):
    """En-tête de page :
      - une ligne de fil (code en pastille + texte, ex. « Lot 05 - CVC ») ;
      - le titre ;
      - une grille de couples (libellé, valeur) ; un couple sans libellé
        s'affiche en sous-titre sous le titre."""
    sous_titres = []
    cellules = []
    for libelle, valeur in (meta or []):
        if libelle:
            cellules.append(f"<div><span class='lab'>{esc(libelle)}</span>"
                            f"<span class='val'>{esc(valeur) or '—'}</span></div>")
        else:
            sous_titres.append(f"<div class='ent-sous'>{esc(valeur)}</div>")
    # Fil d'Ariane seulement s'il porte une information (code projet, lot, page du
    # portail) ; un simple rappel du nom de la page (code vide) n'est pas affiché.
    morceaux_fil = []
    if code:
        morceaux_fil.append(f"<span class='ent-code'>{esc(code)}</span>")
    if fil:
        if morceaux_fil:
            morceaux_fil.append("<span class='sep'>›</span>")
        morceaux_fil.append(f"<span>{esc(fil)}</span>")
    fil_html = f"<div class='ent-fil'>{''.join(morceaux_fil)}</div>" if morceaux_fil else ""
    grille = f"<div class='ent-meta'>{''.join(cellules)}</div>" if cellules else ""
    st.markdown(f"<div class='entete'>{fil_html}"
                f"<div class='ent-titre'>{esc(titre)}</div>{''.join(sous_titres)}{grille}</div>",
                unsafe_allow_html=True)


def mini_jauge(taux, libelle="validé", grande=False):
    """Mini jauge en ligne : petite barre + « 33 % validé ». Uniquement des
    <span>, donc utilisable à l'intérieur d'un texte ou d'un autre <span>.
    grande=True : barre pleine largeur et chiffre agrandi (tuiles)."""
    taux = max(0, min(100, int(taux)))
    classe = "mini-jauge grande" if grande else "mini-jauge"
    return (f"<span class='{classe}' title='{taux} % {esc(libelle)}'>"
            f"<span class='mj-piste'><span class='mj-rempli' style='width:{taux}%'></span></span>"
            f"<span class='mj-val'>{taux} %</span><span class='mj-lab'>{esc(libelle)}</span></span>")


def tuiles(liste):
    """Rangée de tuiles d'indicateurs de même hauteur. Chaque tuile est un
    dictionnaire : lab, val, unite (optionnel), sub, alerte (bool),
    jauge (0-100, grande jauge sous la valeur, qui est alors réduite de moitié)."""
    morceaux = []
    for t in liste:
        unite = f"<small>{esc(t['unite'])}</small>" if t.get("unite") else ""
        jauge = ""
        classe_val = "t-val"
        if t.get("jauge") is not None:
            jauge = f"<div>{mini_jauge(t['jauge'], grande=True)}</div>"
            classe_val = "t-val petit"
        classe = "tuile alerte" if t.get("alerte") else "tuile"
        morceaux.append(f"<div class='{classe}'><div class='t-lab'>{esc(t['lab'])}</div>"
                        f"<div class='{classe_val}'>{esc(t['val'])}{unite}</div>{jauge}"
                        f"<div class='t-sub'>{esc(t.get('sub'))}</div></div>")
    st.markdown(f"<div class='tuiles'>{''.join(morceaux)}</div>", unsafe_allow_html=True)


def barre_actions(titre="Actions"):
    """Petit titre au-dessus d'une rangée de boutons d'action."""
    st.markdown(f"<div class='barre-actions'>{esc(titre)}</div>", unsafe_allow_html=True)


def sous_titre(texte):
    """Petit titre de section souligné de rouge."""
    st.markdown(f"<div class='sous-titre'>{esc(texte)}</div>", unsafe_allow_html=True)


MAX_INFOBULLE = 12


def survol(etiquette_html, titre, elements):
    """Étiquette qui affiche au survol une infobulle listant des éléments
    (texte déjà échappé ou HTML simple)."""
    # Uniquement des éléments en ligne (span) : un <div> ou un <ul> dans un
    # <span> ferait sortir le contenu de l'infobulle.
    # Une infobulle reste courte : au-delà de MAX_INFOBULLE éléments, on renvoie au Suivi.
    puces = "".join(f"<span class='b-item'>• {e}</span>" for e in elements[:MAX_INFOBULLE]) \
        or "<span class='b-item b-gris'>Aucun élément</span>"
    if len(elements) > MAX_INFOBULLE:
        puces += (f"<span class='b-item b-gris'>… et {len(elements) - MAX_INFOBULLE} autre(s) : "
                  "voir le Suivi des visas</span>")
    return (f"<span class='survol' tabindex='0'>{etiquette_html}<span class='bulle' role='tooltip'>"
            f"<span class='b-titre'>{esc(titre)}</span>{puces}</span></span>")
