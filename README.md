# Simulateur de contrôle des contrats de professionnalisation

> **Projet en cours** : outil d'aide au contrôle, pas encore finalisé.

Un simulateur HTML qui vérifie un contrat de professionnalisation au regard des règles de la branche choisie : rémunération minimale, durée du contrat, volume et part de formation. Il est généré par un script Python qui lit la grille d'auto-contrôle (Excel / Google Sheets).

Le simulateur tient en **un seul fichier** (`simulateur_contrat_pro.html`), sans dépendance : il s'ouvre dans n'importe quel navigateur, même hors connexion.

La documentation technique complète est dans [DOCUMENTATION.md](DOCUMENTATION.md).

## Ce que le simulateur contrôle

| Contrôle | Règle appliquée |
|---|---|
| **Âge** | À partir de 16 ans. À partir de 26 ans, le salarié doit être demandeur d'emploi. |
| **Rémunération minimale** | Pourcentage du SMIC selon la tranche d'âge (16-20, 21-25, 26 ans et +) et le niveau de diplôme, propre à chaque branche, proratisé selon la durée hebdomadaire. Pour les 26 ans et +, le plus élevé du SMIC et de 85 % du salaire minimum conventionnel. Certaines branches sont contrôlées sur la base du salaire minimum conventionnel. |
| **Durée** | 6 à 12 mois. Jusqu'à 24 mois si l'accord de branche le prévoit, 36 mois pour les publics prioritaires (Code du travail, art. L6325-11 et L6325-12). |
| **Volume de formation** | 150 heures minimum (art. L6325-13). |
| **Part de formation** | 15 à 25 % de la durée du contrat. Au-delà, jusqu'au plafond de la branche si les conditions de l'accord sont remplies (art. L6325-13 et L6325-14). |

Chaque contrôle donne l'un de ces quatre résultats : **Conforme**, **Conforme sous réserve**, **Non conforme** ou **Information incomplète** (quand une donnée manque). Les renvois citent le texte de référence de la branche (accord, avenant ou convention collective).

## Utilisation

```bash
pip install -r requirements.txt

python generer_simulateur.py                            # télécharge la grille depuis Google Drive
python generer_simulateur.py --xlsx grille.xlsx         # ou utilise un fichier Excel local
python generer_simulateur.py --json donnees_grille.json # ou régénère depuis les données déjà extraites
python generer_simulateur.py --export-json donnees.json # enregistre aussi les données lues
```

Le script :
1. repère l'onglet d'accès par branche (colonnes IDCC, rémunérations, densité) ;
2. reconstitue les cellules fusionnées et lit les en-têtes sur plusieurs lignes ;
3. extrait pour chaque branche les taux de rémunération, la densité maximale de formation, les commentaires, les notes et la liste des publics prioritaires ;
4. remplace les renvois du type « voir l'onglet 1 » par l'intitulé du texte juridique correspondant ;
5. injecte le tout dans le simulateur HTML.

## Paramètres à tenir à jour

En haut de `generer_simulateur.py` :
- `REGLES_LEGALES` : montant du SMIC (au 1er juin 2026), durées et pourcentages légaux. À modifier à chaque revalorisation du SMIC, puis relancer le script.
- `CHARTE` : les trois couleurs du simulateur.

## Fichiers

```
DOCUMENTATION.md              documentation technique détaillée
generer_simulateur.py         script d'import de la grille et de génération du simulateur
simulateur_contrat_pro.html   simulateur généré (grille V8.2 du 22/07/2026)
donnees_grille.json           données extraites de la grille, pour régénérer sans le fichier Excel
```

## Limites

- Outil d'aide au contrôle : chaque résultat est à confirmer avec le texte de référence de la branche.
- Certaines branches n'ont ni taux ni densité dans la grille. Le simulateur l'indique et renvoie vers la convention collective.
