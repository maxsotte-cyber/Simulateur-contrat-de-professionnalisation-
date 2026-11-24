#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Générateur du simulateur de contrôle des contrats de professionnalisation
à partir de la « Grille auto contrôle » OCAPIAT (Google Sheets / Excel).

Ce que fait le script :
  1. importe la grille (téléchargement depuis Google Drive, ou fichier .xlsx local) ;
  2. repère l'onglet d'accès par branche (colonnes IDCC, rémunérations, densité…) ;
  3. reconstitue les cellules fusionnées, lit les en-têtes sur plusieurs lignes ;
  4. extrait pour chaque branche : taux de rémunération par âge / diplôme,
     densité maximale de formation, commentaires, notes et publics prioritaires ;
  5. injecte ces données dans un simulateur HTML autonome (un seul fichier,
     sans dépendance, utilisable hors ligne).

Usage :
  pip install openpyxl
  python generer_simulateur.py                          # télécharge la grille Google
  python generer_simulateur.py --xlsx grille.xlsx       # utilise un fichier local
  python generer_simulateur.py --onglet "INDEX"         # force l'onglet à lire
  python generer_simulateur.py --json donnees.json      # régénère depuis un JSON
  python generer_simulateur.py --export-json donnees.json   # sauvegarde les données lues

Les règles légales (durée, part de formation, SMIC) sont dans REGLES_LEGALES
ci-dessous : à mettre à jour lors d'une revalorisation du SMIC.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import unicodedata
import urllib.request
from datetime import datetime
from pathlib import Path

try:
    import openpyxl
except ImportError:  # pragma: no cover
    sys.exit("Il manque openpyxl : lancez  pip install openpyxl")


# --------------------------------------------------------------------------
# Paramètres
# --------------------------------------------------------------------------
ID_GOOGLE = "17mJxZ-ysBGudQO7T5yc_ZI8s6BpS_zU5"
URL_EXPORT = f"https://docs.google.com/spreadsheets/d/{ID_GOOGLE}/export?format=xlsx"
FICHIER_TELECHARGE = Path("grille_auto_controle.xlsx")
FICHIER_SORTIE = Path("simulateur_contrat_pro.html")

# Charte graphique OCAPIAT : remplacer ces trois valeurs par les codes
# officiels de la charte si besoin, tout le simulateur suit.
CHARTE = {
    "marque": "#0B3C5D",        # bleu principal (bandeau, titres, boutons)
    "marque-2": "#8CBF26",      # vert (filets, accents)
    "marque-clair": "#EAF1F6",  # fond clair des encarts
}

REGLES_LEGALES = {
    # SMIC au 1er juin 2026 (35 h hebdomadaires)
    "smic_mensuel": 1867.02,
    "smic_horaire": 12.31,
    "smic_date": "1er juin 2026",
    # Code du travail, L6325-11 et L6325-12 : durée du CDD ou de l'action
    "duree_min_mois": 6,
    "duree_max_mois": 12,
    "duree_max_accord_mois": 24,      # extension possible par accord de branche
    "duree_max_prioritaire_mois": 36,  # publics prioritaires (L6325-1-1)
    # Code du travail, L6325-13 et L6325-14 : part des actions de formation
    "formation_min_pct": 15,
    "formation_min_heures": 150,
    "formation_max_pct": 25,           # au-delà : seulement si l'accord de branche le prévoit
    # Plancher pour les 26 ans et plus : SMIC et 85 % du minimum conventionnel
    "pct_smc_26_plus": 85,
}


# --------------------------------------------------------------------------
# Outils de lecture
# --------------------------------------------------------------------------
def normaliser(texte) -> str:
    """Minuscules, sans accents, espaces simples."""
    if texte is None:
        return ""
    t = unicodedata.normalize("NFKD", str(texte))
    t = "".join(c for c in t if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", t).strip().lower()


def texte(val) -> str:
    if val is None:
        return ""
    if isinstance(val, float) and val.is_integer():
        val = int(val)
    return re.sub(r"\s+", " ", str(val)).strip()


def en_pourcentage(val):
    """0.6 -> 60 ; '60 %' -> 60 ; '25 % et 40 %' -> 40 (valeur maximale)."""
    if val is None or val == "":
        return None
    if isinstance(val, (int, float)):
        v = float(val)
        return round(v * 100, 2) if 0 < v <= 1.5 else round(v, 2)
    nombres = [float(n.replace(",", ".")) for n in re.findall(r"\d+(?:[.,]\d+)?", str(val))]
    if not nombres:
        return None
    v = max(nombres)
    return round(v * 100, 2) if 0 < v <= 1.5 else round(v, 2)


def telecharger(url: str, destination: Path) -> Path:
    print(f"Téléchargement de la grille : {url}")
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=60) as rep:
        contenu = rep.read()
    if not contenu.startswith(b"PK"):
        sys.exit("Le fichier reçu n'est pas un classeur Excel : vérifiez que la "
                 "grille est partagée « Tous les utilisateurs disposant du lien ».")
    destination.write_bytes(contenu)
    return destination


def matrice_avec_fusions(ws) -> list[list]:
    """Valeurs de la feuille, chaque cellule fusionnée recevant la valeur
    de sa cellule en haut à gauche (indispensable pour les en-têtes à deux
    niveaux et pour les taux uniques fusionnés sur plusieurs tranches)."""
    lignes = [list(r) for r in ws.iter_rows(values_only=True)]
    for plage in ws.merged_cells.ranges:
        valeur = ws.cell(plage.min_row, plage.min_col).value
        for r in range(plage.min_row, plage.max_row + 1):
            for c in range(plage.min_col, plage.max_col + 1):
                if r - 1 < len(lignes) and c - 1 < len(lignes[r - 1]):
                    lignes[r - 1][c - 1] = valeur
    return lignes


def trouver_onglet(wb, nom_force: str | None):
    if nom_force:
        for ws in wb.worksheets:
            if normaliser(ws.title) == normaliser(nom_force):
                return ws
        sys.exit(f"Onglet introuvable : {nom_force}. Onglets : {wb.sheetnames}")
    for ws in wb.worksheets:
        contenu = " ".join(normaliser(v) for r in ws.iter_rows(max_row=40, values_only=True)
                           for v in r if v is not None)
        if "idcc" in contenu and "densit" in contenu and "remuneration" in contenu:
            return ws
    sys.exit("Aucun onglet ne contient à la fois IDCC, rémunérations et densité. "
             "Précisez-le avec --onglet.")


# --------------------------------------------------------------------------
# Analyse de l'onglet
# --------------------------------------------------------------------------
def est_ligne_entete(ligne) -> bool:
    """Une ligne d'en-tête ne contient aucun nombre."""
    vals = [v for v in ligne if v not in (None, "")]
    return bool(vals) and not any(isinstance(v, (int, float)) for v in vals)


def cle_colonne(libelle: str) -> str | None:
    n = normaliser(libelle)
    oui = re.search(r"\boui\b", n)
    non = re.search(r"\bnon\b", n)
    if re.search(r"16\s*(?:a|-|/)?\s*20", n):
        return "r_16_20_oui" if oui else "r_16_20_non" if non else "r_16_20"
    if re.search(r"21\s*(?:a|-|/)?\s*25", n):
        return "r_21_25_oui" if oui else "r_21_25_non" if non else "r_21_25"
    if re.search(r"26\s*ans", n):
        return "r_26"
    if "idcc" in n:
        return "idcc"
    if "densit" in n or "maxi" in n:
        return "densite"
    if "comment" in n:
        return "commentaire"
    if "secteur" in n:
        return "secteur"
    if "type" in n:
        return "type"
    if "accord" in n:
        return "accord"
    return None


def analyser(lignes: list[list]) -> dict:
    # 1. ligne contenant « IDCC » = début du bloc d'en-têtes
    debut = next((i for i, l in enumerate(lignes)
                  if any("idcc" == normaliser(v) or normaliser(v).startswith("idcc")
                         for v in l if v is not None)), None)
    if debut is None:
        sys.exit("En-tête « IDCC » introuvable.")
    fin = debut
    # on accepte jusqu'à 3 lignes d'en-tête supplémentaires tant qu'elles
    # ressemblent à des sous-titres (oui/non, tranches d'âge, % maxi…)
    while fin + 1 < len(lignes) and fin - debut < 3 and est_ligne_entete(lignes[fin + 1]):
        suivante = " ".join(normaliser(v) for v in lignes[fin + 1] if v)
        if not re.search(r"\b(oui|non|ans|maxi|opsi)\b|%", suivante):
            break
        fin += 1

    # 2. libellé complet de chaque colonne (concaténation des niveaux)
    largeur = max(len(l) for l in lignes[debut:fin + 1])
    colonnes = {}
    for c in range(largeur):
        morceaux = []
        for r in range(debut, fin + 1):
            v = texte(lignes[r][c]) if c < len(lignes[r]) else ""
            if v and v not in morceaux:
                morceaux.append(v)
        cle = cle_colonne(" ".join(morceaux))
        if cle and cle not in colonnes:
            colonnes[cle] = c
    if "secteur" not in colonnes and "accord" not in colonnes:
        sys.exit(f"Colonnes reconnues insuffisantes : {colonnes}")

    def val(ligne, cle):
        c = colonnes.get(cle)
        return ligne[c] if c is not None and c < len(ligne) else None

    # 3. lignes de données
    branches, notes, publics = [], [], []
    zone_publics = False
    for ligne in lignes[fin + 1:]:
        cellules = [texte(v) for v in ligne if texte(v)]
        if not cellules:
            continue
        premier = cellules[0]
        n = normaliser(" ".join(cellules))

        if "public" in n and "prioritaire" in n and len(cellules) <= 2:
            zone_publics = True
            continue
        if zone_publics:
            for c in cellules:
                if c not in publics:
                    publics.append(c)
            continue
        if premier.startswith("*"):
            if premier not in notes:
                notes.append(premier)
            continue

        secteur = texte(val(ligne, "secteur")) or texte(val(ligne, "accord"))
        if not secteur:
            continue
        taux = {
            "16-20": {"non": en_pourcentage(val(ligne, "r_16_20_non") or val(ligne, "r_16_20")),
                      "oui": en_pourcentage(val(ligne, "r_16_20_oui") or val(ligne, "r_16_20"))},
            "21-25": {"non": en_pourcentage(val(ligne, "r_21_25_non") or val(ligne, "r_21_25")),
                      "oui": en_pourcentage(val(ligne, "r_21_25_oui") or val(ligne, "r_21_25"))},
            "26+": en_pourcentage(val(ligne, "r_26")),
        }
        commentaire = texte(val(ligne, "commentaire"))
        branches.append({
            "accord": texte(val(ligne, "accord")),
            "type": texte(val(ligne, "type")),
            "secteur": secteur,
            "idcc": texte(val(ligne, "idcc")),
            "taux": taux,
            "densite_max": en_pourcentage(val(ligne, "densite")),
            "commentaire": commentaire,
            "base_smc": "smc" in normaliser(commentaire),
        })

    # dédoublonnage (les fusions verticales peuvent répéter une ligne)
    vues, uniques = set(), []
    for b in branches:
        cle = (b["secteur"], b["idcc"])
        if cle not in vues:
            vues.add(cle)
            uniques.append(b)
    return {"branches": uniques, "notes": notes, "publics_prioritaires": publics}


def trouver_version(lignes) -> str:
    for l in lignes[:15]:
        for v in l:
            m = re.search(r"\bV\s?0?\d+(?:\.\d+)?\b", texte(v))
            if m:
                return m.group(0)
    return ""


# --------------------------------------------------------------------------
# Renvois : on remplace « voir l'onglet n° X » par le texte de référence
# --------------------------------------------------------------------------
MOTIF_RENVOI = re.compile(
    r"\(?(?:veuillez\s+)?(?:voir|cf\.?|consulter|utiliser|se reporter [aà])?\s*"
    r"(?:le |la |les )?(?:commentaires? (?:de |du )?)?(?:l'|la |le )?"
    r"(?:onglet|grille(?: de contr[oô]le)?)\s*(?:n°|no|n°\s|num[eé]ro)?\s*(\d+)(?:\s*\))?",
    re.IGNORECASE)


def titres_des_textes(wb) -> dict:
    """Pour chaque onglet numéroté (1, 2, …), retrouve l'intitulé du texte
    juridique qu'il décrit (première cellule citant un accord, un avenant
    ou une convention)."""
    titres = {}
    for ws in wb.worksheets:
        m = re.fullmatch(r"\s*(?:n°\s*)?(\d+)\s*", ws.title)
        if not m:
            continue
        candidats = [texte(v) for r in ws.iter_rows(max_row=15, values_only=True) for v in r if texte(v)]
        titre = next((c for c in candidats if re.search(r"accord|avenant|convention|ccn", c, re.I)),
                     candidats[0] if candidats else "")
        if titre:
            titres[m.group(1)] = titre
    return titres


def reference_branche(b: dict) -> str:
    if b.get("accord") and normaliser(b["accord"]) not in ("legal", "conventionnel"):
        return b["accord"]
    if b.get("idcc"):
        return f"la convention collective de la branche (IDCC {b['idcc']})"
    return "la convention collective de la branche"


def remplacer_renvois(chaine: str, titres: dict, defaut: str) -> str:
    if not chaine:
        return chaine
    def sub(m):
        titre = titres.get(m.group(1)) or defaut
        avant = " " if m.group(0)[:1].isspace() else ""
        return f"{avant}se référer à « {titre} »"
    return re.sub(r"\s{2,}", " ", MOTIF_RENVOI.sub(sub, chaine)).strip()


def importer_grille(chemin: Path, onglet: str | None) -> dict:
    wb = openpyxl.load_workbook(chemin, data_only=True)
    ws = trouver_onglet(wb, onglet)
    print(f"Onglet lu : « {ws.title} »")
    lignes = matrice_avec_fusions(ws)
    donnees = analyser(lignes)
    titres = titres_des_textes(wb)
    for b in donnees["branches"]:
        b["reference"] = reference_branche(b)
        b["commentaire"] = remplacer_renvois(b["commentaire"], titres, b["reference"])
    donnees["notes"] = [remplacer_renvois(n, titres, "l'accord de branche concerné")
                        for n in donnees["notes"]]
    donnees["meta"] = {
        "titre": "Grille auto contrôle — contrats de professionnalisation",
        "version": trouver_version(lignes),
        "onglet": ws.title,
        "source": chemin.name,
    }
    return donnees


# --------------------------------------------------------------------------
# Génération du HTML
# --------------------------------------------------------------------------
def generer_html(donnees: dict, sortie: Path) -> None:
    donnees = dict(donnees)
    donnees["regles"] = REGLES_LEGALES
    donnees.setdefault("meta", {})["genere_le"] = datetime.now().strftime("%d/%m/%Y %H:%M")
    for i, b in enumerate(donnees["branches"]):
        b["id"] = i
        b.setdefault("reference", reference_branche(b))
    charge = json.dumps(donnees, ensure_ascii=False).replace("</", "<\\/")
    css = " ".join(f"--{k}:{v};" for k, v in CHARTE.items())
    html = GABARIT_HTML.replace("/*__CHARTE__*/", css).replace("/*__DONNEES__*/null", charge)
    sortie.write_text(html, encoding="utf-8")
    print(f"Simulateur généré : {sortie.resolve()}  ({len(donnees['branches'])} branches)")


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--xlsx", type=Path, help="fichier Excel local (sinon téléchargement Google)")
    p.add_argument("--onglet", help="nom de l'onglet à lire (détection automatique par défaut)")
    p.add_argument("--json", type=Path, help="générer depuis un JSON déjà extrait")
    p.add_argument("--export-json", type=Path, help="enregistrer aussi les données extraites")
    p.add_argument("--sortie", type=Path, default=FICHIER_SORTIE)
    a = p.parse_args()

    if a.json:
        donnees = json.loads(a.json.read_text(encoding="utf-8"))
    else:
        chemin = a.xlsx or telecharger(URL_EXPORT, FICHIER_TELECHARGE)
        donnees = importer_grille(chemin, a.onglet)
    if a.export_json:
        a.export_json.write_text(json.dumps(donnees, ensure_ascii=False, indent=2), encoding="utf-8")
    generer_html(donnees, a.sortie)


# --------------------------------------------------------------------------
# Gabarit HTML (interface + moteur de contrôle en JavaScript)
# --------------------------------------------------------------------------
GABARIT_HTML = r"""<!doctype html>
<html lang="fr">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Simulateur contrat pro</title>
<style>
:root{
  /*__CHARTE__*/
  --bg:#f3f5f7; --surface:#ffffff; --surface-2:var(--marque-clair); --ink:#1b2733; --ink-2:#5a6672;
  --line:#d8dee4; --accent:var(--marque); --accent-ink:#ffffff;
  --ok:#2e7d4f; --ok-bg:#e3f2e8; --warn:#9a6200; --warn-bg:#fdf1db; --ko:#b3261e; --ko-bg:#fbe5e3;
  --inc:#475569; --inc-bg:#e9edf2;
  --radius:10px; --shadow:0 1px 2px rgba(11,40,70,.06),0 4px 16px rgba(11,40,70,.06);
}
@media (prefers-color-scheme: dark){
  :root:not([data-theme="light"]){
    --bg:#10161c; --surface:#18212a; --surface-2:#1f2a35; --ink:#e6ebf0; --ink-2:#9fabb7;
    --line:#2c3844; --accent:var(--marque-2); --accent-ink:#10161c;
    --ok:#7fd3a0; --ok-bg:#1c3326; --warn:#f0c06a; --warn-bg:#3a2e17; --ko:#ff8f86; --ko-bg:#3d1f1d;
    --inc:#c3ccd6; --inc-bg:#26313c; --shadow:none;
  }
}
:root[data-theme="dark"]{
  --bg:#10161c; --surface:#18212a; --surface-2:#1f2a35; --ink:#e6ebf0; --ink-2:#9fabb7;
  --line:#2c3844; --accent:var(--marque-2); --accent-ink:#10161c;
  --ok:#7fd3a0; --ok-bg:#1c3326; --warn:#f0c06a; --warn-bg:#3a2e17; --ko:#ff8f86; --ko-bg:#3d1f1d;
  --inc:#c3ccd6; --inc-bg:#26313c; --shadow:none;
}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);
  font:15px/1.5 "Montserrat",system-ui,-apple-system,"Segoe UI",Roboto,sans-serif}
.bandeau{background:var(--marque);color:#fff;border-bottom:5px solid var(--marque-2)}
header{padding:26px 16px 22px;max-width:1100px;margin:0 auto}
header .org{font-weight:800;letter-spacing:.12em;font-size:13px;color:var(--marque-2)}
header h1{margin:4px 0 0;font-size:22px;letter-spacing:-.01em}
header p{margin:4px 0 0;opacity:.8;font-size:13px}
main{max-width:1100px;margin:20px auto 0;padding:0 16px 40px;display:grid;gap:16px;
  grid-template-columns:minmax(0,1fr) minmax(0,1fr)}
@media (max-width:820px){main{grid-template-columns:minmax(0,1fr)}}
.card{background:var(--surface);border:1px solid var(--line);border-radius:var(--radius);
  box-shadow:var(--shadow);padding:18px}
.card h2{font-size:13px;text-transform:uppercase;letter-spacing:.06em;color:var(--accent);margin:0 0 12px;
  padding-bottom:8px;border-bottom:2px solid var(--marque-2)}
.full{grid-column:1/-1}
label{display:block;font-size:13px;font-weight:600;margin:12px 0 4px}
label small{font-weight:400;color:var(--ink-2)}
input,select{width:100%;padding:10px 12px;border:1px solid var(--line);border-radius:8px;
  background:var(--surface);color:var(--ink);font:inherit}
input::placeholder{color:var(--ink-2);opacity:.7}
input:focus,select:focus{outline:2px solid var(--accent);outline-offset:1px;border-color:transparent}
.row{display:grid;grid-template-columns:1fr 1fr;gap:12px}
.check{display:flex;gap:8px;align-items:flex-start;font-weight:400;margin-top:12px}
.check input{width:auto;margin-top:3px;accent-color:var(--accent)}
.seg{display:flex;border:1px solid var(--line);border-radius:8px;overflow:hidden}
.seg button{flex:1;padding:9px;border:0;background:var(--surface);color:var(--ink);font:inherit;cursor:pointer}
.seg button.on{background:var(--accent);color:var(--accent-ink);font-weight:600}
.reset{margin-top:16px;background:none;border:1px solid var(--line);color:var(--ink-2);border-radius:8px;
  padding:8px 12px;font:inherit;font-size:13px;cursor:pointer}
.infos dl{display:grid;grid-template-columns:auto 1fr;gap:6px 14px;margin:0;font-size:14px}
.infos dt{color:var(--ink-2)} .infos dd{margin:0}
table{width:100%;border-collapse:collapse;font-size:14px;margin-top:12px}
th,td{padding:7px 8px;border-bottom:1px solid var(--line);text-align:right;min-width:60px}
th:first-child,td:first-child{text-align:left}
th{color:var(--ink-2);font-weight:600;font-size:12px}
td.cur{background:var(--surface-2);font-weight:700}
.comment{margin-top:12px;padding:10px 12px;background:var(--surface-2);border-left:4px solid var(--marque-2);border-radius:6px;font-size:13px}
.vide{color:var(--ink-2);font-style:italic}
.verdict{padding:14px;border-radius:10px;font-weight:700;font-size:17px}
.verdict small{display:block;font-weight:400;font-size:13px;margin-top:2px}
.verdict.ok{background:var(--ok-bg);color:var(--ok)} .verdict.warn{background:var(--warn-bg);color:var(--warn)}
.verdict.ko{background:var(--ko-bg);color:var(--ko)} .verdict.inc{background:var(--inc-bg);color:var(--inc)}
.regles{list-style:none;margin:14px 0 0;padding:0;display:grid;gap:10px}
.regles li{border:1px solid var(--line);border-radius:10px;padding:12px 12px 12px 14px;border-left-width:5px}
.regles li.ok{border-left-color:var(--ok)} .regles li.warn{border-left-color:var(--warn)}
.regles li.ko{border-left-color:var(--ko)} .regles li.inc{border-left-color:var(--inc)}
.regles .t{display:flex;justify-content:space-between;gap:10px;font-weight:600}
.badge{font-size:12px;padding:2px 8px;border-radius:99px;white-space:nowrap;align-self:flex-start}
.badge.ok{background:var(--ok-bg);color:var(--ok)} .badge.warn{background:var(--warn-bg);color:var(--warn)}
.badge.ko{background:var(--ko-bg);color:var(--ko)} .badge.inc{background:var(--inc-bg);color:var(--inc)}
.regles .d{font-size:13px;color:var(--ink-2);margin-top:4px}
.kpis{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:10px;margin-top:14px}
.kpi{background:var(--surface-2);border-radius:10px;padding:10px}
.kpi b{display:block;font-size:18px;font-variant-numeric:tabular-nums;color:var(--accent)}
.kpi span{font-size:12px;color:var(--ink-2)}
@media (max-width:480px){.kpis{grid-template-columns:1fr 1fr}.row{grid-template-columns:1fr}}
.notes{font-size:13px;color:var(--ink-2)} .notes ul{padding-left:18px;margin:6px 0}
footer{max-width:1100px;margin:0 auto;padding:0 16px 30px;font-size:12px;color:var(--ink-2)}
</style>
</head>
<body>
<div class="bandeau"><header>
  <div class="org">OCAPIAT</div>
  <h1>Simulateur de contrôle — contrats de professionnalisation</h1>
  <p id="sousTitre"></p>
</header></div>

<main>
  <section class="card">
    <h2>Secteur et contrat</h2>
    <label for="branche">Branche / secteur</label>
    <select id="branche"></select>

    <label>Nature du contrat</label>
    <div class="seg" id="nature">
      <button type="button" data-v="CDD">CDD</button>
      <button type="button" data-v="CDI">CDI</button>
    </div>

    <div class="row">
      <div>
        <label for="age">Âge du salarié <small>au début du contrat</small></label>
        <input id="age" type="number" min="14" max="80" placeholder="ex. 19">
      </div>
      <div>
        <label for="diplome">Titulaire d'un bac pro ou diplôme / titre pro de niveau 4 et plus</label>
        <select id="diplome"><option value="">— à renseigner —</option><option value="non">Non</option><option value="oui">Oui</option></select>
      </div>
    </div>

    <label class="check"><input id="prioritaire" type="checkbox">
      <span>Public prioritaire <small>(voir liste en bas de page)</small></span></label>
    <label class="check"><input id="extension" type="checkbox">
      <span>Les conditions d'extension prévues par l'accord de branche sont remplies <small>(part de formation au-delà de 25 %, durée au-delà de 12 mois)</small></span></label>

    <div class="row">
      <div>
        <label for="duree" id="labelDuree">Durée <small>(mois)</small></label>
        <input id="duree" type="number" min="1" max="48" step="0.5" placeholder="ex. 12">
      </div>
      <div>
        <label for="heuresForm">Heures de formation</label>
        <input id="heuresForm" type="number" min="0" step="1" placeholder="ex. 400">
      </div>
    </div>
    <div class="row">
      <div>
        <label for="hebdo">Durée hebdomadaire <small>(h)</small></label>
        <input id="hebdo" type="number" min="1" max="48" step="0.5" placeholder="ex. 35">
      </div>
      <div>
        <label for="salaire">Salaire brut mensuel proposé <small>(€)</small></label>
        <input id="salaire" type="number" min="0" step="0.01" placeholder="ex. 1 300">
      </div>
    </div>
    <div class="row">
      <div>
        <label for="smic">SMIC mensuel 35 h <small>(€, valeur de référence)</small></label>
        <input id="smic" type="number" min="0" step="0.01">
      </div>
      <div>
        <label for="smc">Salaire minimum conventionnel <small>(€)</small></label>
        <input id="smc" type="number" min="0" step="0.01" placeholder="si connu">
      </div>
    </div>
    <button type="button" class="reset" id="reset">Vider les champs</button>
  </section>

  <section class="card">
    <h2>Résultat du contrôle</h2>
    <div class="verdict" id="verdict"></div>
    <div class="kpis" id="kpis"></div>
    <ul class="regles" id="regles"></ul>
  </section>

  <section class="card infos full" id="blocBranche">
    <h2>Règles de la branche sélectionnée</h2>
    <dl id="ficheBranche"></dl>
    <div style="overflow-x:auto"><table id="tableTaux"></table></div>
    <div class="comment" id="commentaire"></div>
  </section>

  <section class="card notes full">
    <h2>Notes et publics prioritaires</h2>
    <div id="notes"></div>
  </section>
</main>
<footer id="pied"></footer>

<script>
const DATA = /*__DONNEES__*/null;
const R = DATA.regles;
const $ = id => document.getElementById(id);
const eur = v => v == null || isNaN(v) ? "" : v.toLocaleString("fr-FR",{style:"currency",currency:"EUR"});
const pct = v => v == null || isNaN(v) ? "" : v.toLocaleString("fr-FR",{maximumFractionDigits:1}) + " %";
const esc = s => String(s ?? "").replace(/[&<>"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const num = id => { const v = $(id).value.trim(); return v === "" ? null : +v; };
const LIB = {ok:"Conforme", warn:"Conforme sous réserve", ko:"Non conforme", inc:"Information incomplète"};

/* ---------- initialisation ---------- */
$("sousTitre").textContent = [DATA.meta.titre, DATA.meta.version, "SMIC du " + R.smic_date].filter(Boolean).join(" · ");
$("smic").value = R.smic_mensuel;
const groupes = {};
DATA.branches.forEach(b => (groupes[b.type || "Autres branches"] ??= []).push(b));
$("branche").innerHTML = `<option value="">— Choisir une branche —</option>` +
  Object.entries(groupes).map(([g, bs]) => `<optgroup label="${esc(g)}">` + bs.map(b =>
    `<option value="${b.id}">${esc(b.secteur)}${b.idcc ? " — IDCC " + esc(b.idcc) : ""}</option>`).join("") + `</optgroup>`).join("");

let nature = null;
$("nature").addEventListener("click", e => {
  const b = e.target.closest("button"); if (!b) return;
  nature = b.dataset.v;
  [...$("nature").children].forEach(x => x.classList.toggle("on", x === b));
  $("labelDuree").innerHTML = (nature === "CDD" ? "Durée du CDD" : "Durée de l'action de professionnalisation") + " <small>(mois)</small>";
  calculer();
});
document.querySelectorAll("input,select").forEach(el => el.addEventListener("input", calculer));
$("reset").addEventListener("click", () => {
  document.querySelectorAll("input[type=number],select").forEach(el => { if (el.id !== "smic") el.value = ""; });
  document.querySelectorAll("input[type=checkbox]").forEach(el => el.checked = false);
  nature = null; [...$("nature").children].forEach(x => x.classList.remove("on"));
  $("labelDuree").innerHTML = "Durée <small>(mois)</small>";
  calculer();
});

/* ---------- moteur de règles ---------- */
const tranche = age => age <= 20 ? "16-20" : age <= 25 ? "21-25" : "26+";
const manque = liste => ({s:"inc", d:"À renseigner : " + liste.join(", ") + "."});
function texteRef(b){ return b && b.reference ? ` Se référer à : ${b.reference}.` : ""; }

function calculer(){
  const b = $("branche").value === "" ? null : DATA.branches[+$("branche").value];
  const age = num("age"), dip = $("diplome").value || null;
  const prio = $("prioritaire").checked, ext = $("extension").checked;
  const duree = num("duree"), hForm = num("heuresForm"), hebdo = num("hebdo");
  const salaire = num("salaire"), smic = num("smic"), smc = num("smc");
  const regles = [];

  /* 1. Âge */
  if (age == null) regles.push({t:"Âge du salarié", ...manque(["l'âge du salarié"])});
  else if (age < 16) regles.push({s:"ko", t:"Âge du salarié", d:"Le contrat de professionnalisation est ouvert à partir de 16 ans."});
  else regles.push({s: age >= 26 && !prio ? "warn" : "ok", t:"Âge du salarié",
    d:`Tranche ${tranche(age)} ans.` + (age >= 26 ? " À partir de 26 ans, le salarié doit être demandeur d'emploi." : "")});

  /* 2. Rémunération */
  let minimum = null;
  (function(){
    const t = "Rémunération minimale", m = [];
    if (!b) m.push("la branche"); if (age == null) m.push("l'âge");
    if (age != null && age < 26 && !dip) m.push("le diplôme");
    if (salaire == null) m.push("le salaire proposé"); if (hebdo == null) m.push("la durée hebdomadaire");
    if (smic == null && !(b && b.base_smc)) m.push("le SMIC");
    if (m.length) return regles.push({t, ...manque(m)});
    const tr = tranche(age);
    const taux = tr === "26+" ? b.taux["26+"] : b.taux[tr][dip];
    const prorata = hebdo / 35;
    if (taux == null) return regles.push({t, s:"inc",
      d:`Taux de rémunération non renseigné dans la grille pour cette branche (tranche ${tr} ans).` + texteRef(b)});
    if (b.base_smc) {
      if (smc == null) return regles.push({t, ...manque([`le salaire minimum conventionnel (contrôle à ${pct(taux)} du SMC)`])});
      minimum = smc * taux / 100 * prorata;
      return regles.push(salaire >= minimum
        ? {t, s:"ok", d:`${eur(salaire)} ≥ ${eur(minimum)} (${pct(taux)} du SMC).`}
        : {t, s:"ko", d:`${eur(salaire)} < ${eur(minimum)} (${pct(taux)} du SMC). Écart : ${eur(minimum - salaire)}.`});
    }
    minimum = smic * taux / 100 * prorata;
    let base = `${pct(taux)} du SMIC`, statutOk = "ok";
    if (tr === "26+") {
      minimum = Math.max(minimum, smic * prorata, smc ? smc * R.pct_smc_26_plus / 100 * prorata : 0);
      base = smc ? `SMIC et ${R.pct_smc_26_plus} % du SMC, le plus élevé` : "SMIC";
      if (!smc) statutOk = "inc";
    }
    if (prorata !== 1) base += `, proratisé ${hebdo} h / 35 h`;
    if (salaire < minimum) return regles.push({t, s:"ko", d:`${eur(salaire)} < ${eur(minimum)} (${base}). Écart : ${eur(minimum - salaire)}.`});
    regles.push({t, s:statutOk, d:`${eur(salaire)} ≥ ${eur(minimum)} (${base}).` +
      (statutOk === "inc" ? ` Information incomplète : renseigner le salaire minimum conventionnel pour vérifier le plancher de ${R.pct_smc_26_plus} %.` : "")});
  })();

  /* 3. Durée */
  const tDuree = nature === "CDI" ? "Durée de l'action de professionnalisation" : "Durée du contrat";
  if (duree == null || !nature) regles.push({t:tDuree, ...manque([!nature && "la nature du contrat", duree == null && "la durée"].filter(Boolean))});
  else {
    let sd;
    if (duree < R.duree_min_mois) sd = {s:"ko", d:`Minimum ${R.duree_min_mois} mois.`};
    else if (duree <= R.duree_max_mois) sd = {s:"ok", d:`Entre ${R.duree_min_mois} et ${R.duree_max_mois} mois.`};
    else if (duree <= R.duree_max_accord_mois) sd = (prio || ext)
      ? {s:"warn", d:`Au-delà de ${R.duree_max_mois} mois : admis jusqu'à ${R.duree_max_accord_mois} mois si l'accord de branche le prévoit, ou ${R.duree_max_prioritaire_mois} mois pour un public prioritaire (Code du travail, art. L6325-12).` + texteRef(b)}
      : {s:"ko", d:`Au-delà de ${R.duree_max_mois} mois sans public prioritaire ni extension prévue par l'accord de branche (Code du travail, art. L6325-11 et L6325-12).`};
    else if (duree <= R.duree_max_prioritaire_mois) sd = prio
      ? {s:"warn", d:`Jusqu'à ${R.duree_max_prioritaire_mois} mois réservé aux publics prioritaires (Code du travail, art. L6325-12) : à justifier.`}
      : {s:"ko", d:`Plus de ${R.duree_max_accord_mois} mois réservé aux publics prioritaires (Code du travail, art. L6325-12).`};
    else sd = {s:"ko", d:`Maximum ${R.duree_max_prioritaire_mois} mois.`};
    regles.push({t:tDuree, ...sd});
  }

  /* 4. Volume et part de formation */
  if (hForm == null) regles.push({t:"Volume de formation", ...manque(["les heures de formation"])});
  else regles.push(hForm >= R.formation_min_heures
    ? {t:"Volume de formation", s:"ok", d:`${hForm} h ≥ ${R.formation_min_heures} h minimum (Code du travail, art. L6325-13).`}
    : {t:"Volume de formation", s:"ko", d:`${hForm} h < ${R.formation_min_heures} h minimum (Code du travail, art. L6325-13).`});

  let ratio = null, heuresContrat = null;
  const mF = []; if (!b) mF.push("la branche"); if (duree == null) mF.push("la durée");
  if (hForm == null) mF.push("les heures de formation"); if (hebdo == null) mF.push("la durée hebdomadaire");
  if (mF.length) regles.push({t:"Part de formation (densité)", ...manque(mF)});
  else {
    heuresContrat = duree * hebdo * 52 / 12;
    ratio = hForm / heuresContrat * 100;
    const maxB = b.densite_max;
    let sf;
    if (ratio < R.formation_min_pct) sf = {s:"ko", d:`${pct(ratio)} < ${R.formation_min_pct} % de la durée (Code du travail, art. L6325-13).`};
    else if (ratio <= R.formation_max_pct) sf = {s:"ok", d:`${pct(ratio)}, entre ${R.formation_min_pct} % et ${R.formation_max_pct} % (Code du travail, art. L6325-13).`};
    else if (maxB == null) sf = {s:"inc", d:`${pct(ratio)} > ${R.formation_max_pct} %. Densité maximale de la branche non renseignée dans la grille.` + texteRef(b)};
    else if (ratio <= maxB) sf = ext
      ? {s:"warn", d:`${pct(ratio)} : au-delà de ${R.formation_max_pct} %, admis jusqu'à ${pct(maxB)} sous les conditions de l'accord de branche (Code du travail, art. L6325-14).` + texteRef(b)}
      : {s:"ko", d:`${pct(ratio)} > ${R.formation_max_pct} %. La branche admet jusqu'à ${pct(maxB)} sous conditions : cocher « conditions d'extension » si elles sont remplies.` + texteRef(b)};
    else sf = {s:"ko", d:`${pct(ratio)} > ${pct(maxB)}, densité maximale de la branche.`};
    regles.push({t:"Part de formation (densité)", ...sf});
  }

  /* ---------- affichage ---------- */
  const has = s => regles.some(r => r.s === s);
  const v = has("ko") ? "ko" : has("inc") ? "inc" : has("warn") ? "warn" : "ok";
  const nbInc = regles.filter(r => r.s === "inc").length;
  $("verdict").className = "verdict " + v;
  $("verdict").innerHTML = esc(LIB[v]) + (nbInc ? `<small>${nbInc} contrôle${nbInc > 1 ? "s" : ""} sur ${regles.length} ne peu${nbInc > 1 ? "vent" : "t"} pas être conclu${nbInc > 1 ? "s" : ""} : informations manquantes.</small>` : "");
  $("regles").innerHTML = regles.map(r =>
    `<li class="${r.s}"><div class="t"><span>${esc(r.t)}</span><span class="badge ${r.s}">${LIB[r.s]}</span></div><div class="d">${esc(r.d)}</div></li>`).join("");

  let plage = "";
  if (heuresContrat) {
    const maxAdm = ext && b.densite_max != null ? b.densite_max : R.formation_max_pct;
    plage = `${Math.max(R.formation_min_heures, Math.ceil(heuresContrat * R.formation_min_pct / 100))}–${Math.floor(heuresContrat * maxAdm / 100)} h`;
  }
  $("kpis").innerHTML = [[eur(minimum), "Salaire minimum"], [pct(ratio), "Part de formation"], [plage, "Heures de formation admises"]]
    .map(([val, l]) => `<div class="kpi"><b>${val || "&nbsp;"}</b><span>${l}</span></div>`).join("");

  afficherBranche(b, age, dip);
}

function afficherBranche(b, age, dip){
  $("blocBranche").style.display = b ? "" : "none";
  if (!b) return;
  const vide = '<span class="vide">non renseigné</span>';
  $("ficheBranche").innerHTML = [["Secteur", esc(b.secteur)], ["Texte de référence", esc(b.accord)],
    ["Type de branche", esc(b.type)], ["IDCC", esc(b.idcc)], ["Densité maximale", pct(b.densite_max)]]
    .map(([k, val]) => `<dt>${k}</dt><dd>${val || vide}</dd>`).join("");
  const t = b.taux, cur = age != null ? tranche(age) : null;
  const c = (tr, d) => cur === tr && (tr === "26+" || dip === d) ? ' class="cur"' : "";
  $("tableTaux").innerHTML = `<tr><th>Rémunération (% ${b.base_smc ? "SMC" : "SMIC"})</th><th>16–20 ans</th><th>21–25 ans</th><th>26 ans et +</th></tr>
    <tr><td>Sans bac pro / niveau 4</td><td${c("16-20","non")}>${pct(t["16-20"].non)}</td><td${c("21-25","non")}>${pct(t["21-25"].non)}</td><td rowspan="2"${c("26+")}>${pct(t["26+"])}</td></tr>
    <tr><td>Avec bac pro / niveau 4</td><td${c("16-20","oui")}>${pct(t["16-20"].oui)}</td><td${c("21-25","oui")}>${pct(t["21-25"].oui)}</td></tr>`;
  $("commentaire").textContent = b.commentaire ? b.commentaire : "Aucun commentaire dans la grille pour cette branche.";
}

/* notes */
const blocs = [];
if (DATA.notes?.length) blocs.push("<ul>" + DATA.notes.map(n => `<li>${esc(n)}</li>`).join("") + "</ul>");
if (DATA.publics_prioritaires?.length) blocs.push("<strong>Publics prioritaires</strong><ul>" + DATA.publics_prioritaires.map(n => `<li>${esc(n)}</li>`).join("") + "</ul>");
$("notes").innerHTML = blocs.join("") || "Aucune note.";
$("pied").textContent = `Généré le ${DATA.meta.genere_le} à partir de « ${DATA.meta.source || "grille d'auto-contrôle"} ». Règles légales : Code du travail, art. L6325-11 à L6325-14. Outil d'aide au contrôle, à confirmer avec le texte de référence de la branche.`;

calculer();
</script>
</body>
</html>
"""

if __name__ == "__main__":
    main()