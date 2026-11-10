# Rapport TARL

Application Windows (Python / tkinter) qui lit des fichiers Excel de résultats d'élèves, **reconnaît automatiquement les colonnes**
(français, arabe, anglais), permet de les corriger, puis génère un **rapport Word de 3 pages** en français, arabe (RTL) ou anglais.

- Pré-test / post-test (compétences 0/1) ou notes chiffrées, toutes matières.
- Fonctionne hors connexion. Mode clair / sombre.
- Les données des élèves restent sur l'ordinateur.

## Utilisation
```
pip install -r requirements.txt
python main.py                      # interface graphique
python main.py pre.xlsx post.xlsx -o rapport.docx --lang ar
python tests/run_tests.py           # tests automatiques
```
## Exécutable Windows
Onglet **Actions → Build Rapport TARL.exe → Run workflow**, puis télécharger l'artefact (ou créer un tag `v2.1` pour une Release).
Détails : `LISEZMOI.txt`, `MODIFICATIONS.md`, `design-system/`.

> ⚠️ Ne publiez jamais de vrais fichiers d'élèves (noms, numéros Massar) sur GitHub : le `.gitignore` les exclut.
