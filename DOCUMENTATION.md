================================================================================
DOCUMENTATION — SIMULATEUR DE CONTRÔLE DES CONTRATS DE PROFESSIONNALISATION
Périmètre : Agriculture, Agroalimentaire, Pêche et Territoires (OPCO OCAPIAT)
================================================================================

1. PRÉSENTATION GÉNÉRALE
-----------------------
Ce projet est un outil autonome d'aide au contrôle de conformité réglementaire
et conventionnelle des contrats de professionnalisation (CPRO) pour les secteurs
de l'agriculture et de l'agroalimentaire (périmètre OCAPIAT).

Il s'articule autour de deux fichiers :
- generer_simulateur.py : script Python d'extraction et de compilation qui lit
  la « Grille auto contrôle » (Google Sheets ou fichier Excel .xlsx local) et
  génère l'application web.
- simulateur_contrat_pro.html : application web monopage (SPA) 100 % autonome
  (HTML/CSS/JavaScript sans dépendance externe, utilisable en local ou hors
  ligne) embarquant l'ensemble des règles légales et conventionnelles en JSON.


2. ARCHITECTURE ET FLUX DE DONNÉES
----------------------------------
[Grille Google Sheets / Fichier .xlsx]
                 |
                 v  (openpyxl : lecture, défusion, détection d'en-têtes, renvois)
       [generer_simulateur.py] ---> (optionnel) Export JSON (--export-json)
                 |
                 v  (Injection de DATA = {...} et de la charte CSS dans GABARIT_HTML)
   [simulateur_contrat_pro.html]
                 |
                 v  (Exécution côté client en JavaScript pur)
   [Interface utilisateur & Moteur de contrôle en temps réel]

Principes techniques clés :
- Zéro dépendance à l'exécution : le fichier HTML généré n'effectue aucun appel
  réseau (ni CDN, ni API, ni police distante). Toutes les données de branches
  et les paramètres légaux sont sérialisés dans la constante DATA.
- Sécurité XSS : toutes les chaînes issues du tableur sont échappées via la
  fonction esc() avant insertion dans le DOM.
- Adaptabilité visuelle : support natif des thèmes clair et sombre
  (@media (prefers-color-scheme: dark)) et mise en page responsive.


3. RÈGLES MÉTIER ET MOTEUR DE CALCUL
------------------------------------
Le moteur JavaScript (calculer()) évalue en temps réel 5 points de contrôle à
partir des paramètres saisis et des règles de la branche sélectionnée (32
branches réparties en 4 catégories : Accord de branche, Légal, Conventionnel,
Branches sans taux dans la grille).

3.1. Paramètres légaux de référence (REGLES_LEGALES)
- SMIC de référence (35 h) : 1 867,02 € mensuel / 12,31 € horaire (au 1er juin
  2026), modifiable directement dans l'interface.
- Durée du contrat ou de l'action (Code du travail, art. L6325-11 et L6325-12) :
  6 à 12 mois par défaut ; jusqu'à 24 mois si extension par accord de branche ;
  jusqu'à 36 mois pour les publics prioritaires.
- Volume et densité de formation (art. L6325-13 et L6325-14) : minimum 150 h
  et 15 % de la durée totale du contrat ; plafond de droit commun à 25 %,
  extensible jusqu'au plafond prévu par la branche (30 %, 40 % ou 50 % selon
  les accords).
- Plancher 26 ans et plus : minimum égal au plus élevé entre le SMIC et 85 % du
  Salaire Minimum Conventionnel (SMC).

3.2. Logique des 5 contrôles
1) Âge du salarié :
   - < 16 ans : Non conforme.
   - 16-20 ans ou 21-25 ans : Conforme.
   - >= 26 ans : Conforme sous réserve (obligation d'être demandeur d'emploi,
     sauf public prioritaire).

2) Rémunération minimale :
   - Détermination du taux applicable selon la tranche d'âge (16-20, 21-25, 26+)
     et la détention d'un diplôme de niveau 4 (bac pro / titre pro) ou plus.
   - Proratisation automatique selon la durée hebdomadaire saisie (hebdo / 35).
   - Cas particuliers gérés :
     * Branches indexées sur le SMC au lieu du SMIC (ex. Personnel des activités
       hippiques — IDCC 116, détecté via la mention « SMC » dans le commentaire).
     * Tranche 26+ : comparaison avec le SMIC et 85 % du SMC (statut
       « Information incomplète » si le SMC n'est pas renseigné).
     * Branches sans taux renseigné dans la grille : renvoi explicite vers le
       texte conventionnel de référence.

3) Durée du contrat (CDD) ou de l'action de professionnalisation (CDI) :
   - < 6 mois ou > 36 mois : Non conforme.
   - 6 à 12 mois : Conforme.
   - > 12 à 24 mois : Conforme sous réserve si « Public prioritaire » ou
     « Conditions d'extension » est coché, sinon Non conforme.
   - > 24 à 36 mois : Conforme sous réserve uniquement si « Public prioritaire »
     est coché.

4) Volume horaire de formation :
   - Vérification du seuil minimal légal de 150 heures.

5) Part de formation (densité) :
   - Calcul du volume d'heures du contrat :
     heuresContrat = duree (mois) * hebdo * (52 / 12)
   - Calcul du ratio :
     ratio = (heuresForm / heuresContrat) * 100
   - Comparaison aux bornes [15 % ; 25 %] puis, en cas de dépassement, au
     plafond conventionnel densite_max sous réserve que la case « Conditions
     d'extension » soit cochée.

3.3. Verdict global et indicateurs (KPI)
- Hiérarchie du verdict global :
  Non conforme (ko) > Information incomplète (inc) > Conforme sous réserve (warn) > Conforme (ok)
- KPI calculés :
  * Salaire minimum requis (€)
  * Part de formation constatée (%)
  * Plage d'heures de formation admises (min - max en heures)


4. PIPELINE DE GÉNÉRATION (generer_simulateur.py)
-------------------------------------------------
4.1. Prérequis et installation
- Python : version 3.10 ou supérieure.
- Dépendance unique : openpyxl
  Commande : pip install openpyxl

4.2. Commandes d'utilisation
- Mode standard (télécharge la grille depuis Google Drive et génère le HTML) :
  python generer_simulateur.py

- À partir d'un fichier Excel local :
  python generer_simulateur.py --xlsx grille_auto_controle.xlsx

- En forçant le nom d'un onglet spécifique :
  python generer_simulateur.py --xlsx grille.xlsx --onglet "accès par branche"

- Exporter les données extraites en JSON pour inspection ou archivage :
  python generer_simulateur.py --xlsx grille.xlsx --export-json donnees.json

- Régénérer le HTML directement depuis un fichier JSON déjà extrait :
  python generer_simulateur.py --json donnees.json --sortie simulateur_contrat_pro.html

4.3. Traitement automatisé du tableur Excel
Le script intègre plusieurs mécanismes pour gérer les mises en forme complexes
des classeurs métiers :
1) Détection de l'onglet cible (trouver_onglet) : recherche automatiquement la
   feuille contenant les mots-clés « idcc », « densit » et « remuneration » dans
   les 40 premières lignes.
2) Reconstitution des cellules fusionnées (matrice_avec_fusions) : propage la
   valeur de la cellule supérieure gauche à l'ensemble de la plage fusionnée
   (indispensable pour les en-têtes multi-niveaux et les taux identiques sur
   plusieurs tranches d'âge).
3) Reconstruction des en-têtes (analyser) : repère la ligne « IDCC » et
   concatène jusqu'à 3 lignes de sous-en-têtes (oui/non, tranches d'âge, % maxi)
   pour identifier dynamiquement l'index de chaque colonne.
4) Résolution des renvois inter-onglets (titres_des_textes, remplacer_renvois) :
   scanne les onglets numérotés (1, 2, n° 3...) pour extraire l'intitulé exact
   de l'accord ou de la convention collective, puis remplace dans les
   commentaires toute mention du type « voir l'onglet n° X » par
   « se référer à « [Intitulé de l'accord] » ».
5) Extraction des notes et publics prioritaires : isole automatiquement les
   notes de bas de tableau (lignes commençant par *) et la liste des publics
   prioritaires.


5. STRUCTURE DU SCHÉMA JSON EMBARQUÉ (DATA)
-------------------------------------------
{
  "meta": {
    "titre": "Grille auto contrôle du 22/07/2026",
    "version": "V08.2",
    "onglet": "accès par branche",
    "source": "Grille auto contrôle V8.2 du 22/07/2026",
    "genere_le": "19/10/2026 10:56"
  },
  "branches": [
    {
      "id": 0,
      "accord": "Accord secteur alimentaire du 19/09/2025",
      "type": "Accord de branche",
      "secteur": "Coopératives laitières",
      "idcc": "7004/8435",
      "taux": {
        "16-20": { "non": 60, "oui": 70 },
        "21-25": { "non": 75, "oui": 85 },
        "26+": 100
      },
      "densite_max": 40,
      "commentaire": "Plafond porté à 40 % pour certains publics",
      "base_smc": false,
      "reference": "Accord secteur alimentaire du 19/09/2025"
    }
  ],
  "notes": ["..."],
  "publics_prioritaires": ["..."],
  "regles": {
    "smic_mensuel": 1867.02,
    "smic_horaire": 12.31,
    "smic_date": "1er juin 2026",
    "duree_min_mois": 6,
    "duree_max_mois": 12,
    "duree_max_accord_mois": 24,
    "duree_max_prioritaire_mois": 36,
    "formation_min_pct": 15,
    "formation_min_heures": 150,
    "formation_max_pct": 25,
    "pct_smc_26_plus": 85
  }
}


6. MAINTENANCE ET LIMITES CONNUES
---------------------------------
- Revalorisation du SMIC : modifier les clés smic_mensuel, smic_horaire et
  smic_date dans le dictionnaire REGLES_LEGALES de generer_simulateur.py, puis
  relancer le script.
- Mise à jour de la grille conventionnelle : exécuter python generer_simulateur.py
  (si le Google Sheets reste partagé en lecture) ou pointer vers la nouvelle
  version .xlsx avec --xlsx.
- Personnalisation graphique : ajuster les trois variables hexadécimales du
  dictionnaire CHARTE (marque, marque-2, marque-clair) dans generer_simulateur.py.
- Limites connues :
  * Pour les branches listées sous « Branches sans taux dans la grille » (ex.
    Industries laitières, Cerfrance, Crédit agricole, MSA, Groupama, ONF,
    Entreprises équestres) ainsi que les Industries et exploitations
    frigorifiques (accord expiré au 31/12/2025), les taux et densités sont null :
    le simulateur signale une information incomplète et renvoie vers la
    convention collective.
  * La fonction en_pourcentage() retient la valeur numérique maximale lorsqu'une
    cellule contient plusieurs taux (ex. "25 % et 40 %" -> 40) ; les cas
    particuliers (comme la distinction CDD/CDI pour les Entreprises du paysage)
    restent visibles dans l'encart de commentaire de la branche.
