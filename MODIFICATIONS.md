# Rapport TARL : résumé des modifications

## v2.1 : rapports limités à 3 pages (d'après le modèle fourni)
- Le rapport pré-test / post-test reprend la structure du modèle `Rapport_TARL.docx` : bandeau marine, paragraphe d'introduction, section 01 (pré-test),
  02 (post-test), 03 (comparaison), cartes de chiffres, deux figures côte à côte, tableau des composantes, « قراءة النتائج / Lecture des résultats »,
  conclusion, niveaux de maîtrise, points forts / de vigilance, évolution individuelle, élèves à accompagner en priorité. **3 pages A4 exactement.**
- Mêmes statistiques que le modèle (moyenne, médiane, écart-type de population, taux, niveaux 0-2 / 3-4 / 5-6…) : valeurs vérifiées identiques sur les fichiers 3APIC-1.
- Le rapport de notes chiffrées tient lui aussi en 3 pages (listes longues tronquées / affichées en deux blocs).
- Français et anglais reprennent la même mise en page que l'arabe (RTL conservé pour l'arabe). Dans les phrases FR/EN, les composantes à intitulé arabe sont désignées par leur code (évite les mélanges de sens d'écriture) ; l'intitulé complet reste dans le tableau.
- Police Segoe UI (comme le modèle). Nouveaux modules : `layout3.py`, `report_skill.py`, `report_grade.py`. Test automatique de longueur (≤ 3 pages) avec une classe de 48 élèves.

# Rapport TARL 2.0 : résumé des modifications

## Limites identifiées dans la version précédente
1. Les colonnes étaient cherchées avec des codes fixes (`F1`, `R1`…) et des libellés arabes figés : un fichier d'arabe (« المرحلة 1…6 », codes `LTC`, `CTC`…) échouait.
2. Aucun contrôle des données, aucune correction possible par l'enseignant.
3. Textes du rapport en français uniquement ; aucune mise en page RTL.
4. Un seul modèle de calcul (compétences 0/1 en deux temps) ; impossible d'analyser des notes chiffrées.

## Ce qui a été fait
- **Architecture modulaire** `rapport_tarl/` : excel_reader, structure_analyzer, column_detector, language_detector, data_normalizer, data_validator, data_processor, report_translator, arabic_rtl, docx_helpers, charts, report_generator, pipeline, app_ui.
- **Analyse intelligente** : feuilles, dimensions, cellules fusionnées (propagées), en-têtes sur plusieurs lignes, lignes/colonnes vides, métadonnées (matière, classe, année, établissement…), zone des élèves, lignes de totaux, langue des en-têtes et des données, type de test (pré/post).
- **Reconnaissance multilingue** FR/AR/EN : dictionnaire extensible + normalisation (accents, alef/ya/ta marbuta, diacritiques, tatweel, chiffres arabes, casse, caractères spéciaux) + correspondance approximative (rapidfuzz, repli difflib) + analyse du contenu (0/1, numérique, oui/non, codes uniques, noms) + règles de cohérence (un seul nom/identifiant, blocs de colonnes, veto « moyenne composée uniquement de 0/1 »).
- **Validation humaine** : fenêtre « Vérifier les colonnes » (liste déroulante par colonne, statut Reconnu / À vérifier / Ignoré / Modifié, choix de la feuille et du barème, aperçu, contrôles).
- **Contrôles avant traitement** : identification des élèves, valeurs non numériques, cellules vides, hors barème, valeurs aberrantes, doublons, cohérence des moyennes. Rien n'est supprimé ni modifié silencieusement.
- **Structure interne standardisée** (`Dataset.to_standard()`), indépendante de la langue, des noms et de l'ordre des colonnes.
- **Deux moteurs de traitement** : compétences 0/1 pré/post (calculs conservés : mêmes résultats que la v1.x) et notes chiffrées (moyenne, médiane, écart-type, taux de réussite, niveaux, classements, évolution, absences, participation, comportement).
- **Rapport en 3 langues** (FR / AR RTL / EN), choix dans la section « Langue du rapport · Report Language · لغة التقرير ». Images de graphiques sans texte arabe ; titres et légendes en texte Word.
- **Interface** : fil d'étapes, drapeaux, mode sombre, zone défilante, barre de progression, journal, mémorisation des préférences.
- **Design system** généré avec UI/UX Pro Max (`design-system/`).

## Fichiers ajoutés / modifiés
Ajoutés : `rapport_tarl/*` (14 modules), `main.py`, `tests/`, `exemples/`, `design-system/`, `synonymes_utilisateur.json`, `MODIFICATIONS.md`.
Modifiés : `Lancer_Rapport_TARL.bat`, `Construire_EXE.bat`, `requirements.txt`, `LISEZMOI.txt`, `rapport_tarle.py` (devient un simple appel à `main.py`).
Supprimé : l'ancien `app_tarle.py` (remplacé par `rapport_tarl/app_ui.py`, mêmes fonctions + nouvelles).
