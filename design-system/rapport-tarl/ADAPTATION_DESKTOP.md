# Adaptation du design system « Rapport TARL » à l'application Windows (tkinter) et au rapport Word

Source : `MASTER.md` généré avec le skill UI/UX Pro Max (style Minimalism / Swiss, densité 6, animation 3).

| Jeton du design system | Valeur | Utilisation dans l'application |
|---|---|---|
| Primary `#1E3A5F` | bleu marine | en-tête, badges d'étapes, couverture du rapport, titres |
| Secondary `#2563EB` | bleu | actions secondaires, post-test, sélection de langue |
| Accent/CTA `#059669` | vert | bouton « Générer », progression, « Prêt » |
| Background `#F8FAFC` / Card `#FFFFFF` | fonds | fond de fenêtre / cartes |
| Foreground `#0F172A`, Muted `#475569` | textes | texte principal / secondaire |
| Border `#E4E7EB` | contours | cartes, champs, tableaux |
| Destructive `#DC2626` | rouge | erreurs, régression, élèves en difficulté |
| Orange `#E07B39` | pré-test | séries « pré-test » (graphiques, légendes) |

**Mode sombre** (interrupteur en haut à gauche) : fond `#0B1220`, cartes `#111B2E`, texte `#E6EDF5`, accents éclaircis (`#60A5FA`, `#34D399`)
pour garder un contraste ≥ 4,5:1.

**Typographie** : Figtree / Noto Sans (polices web du design system) remplacées par **Segoe UI** (interface) et **Arial** (rapport Word,
latin + arabe) car l'application fonctionne hors connexion et doit s'afficher à l'identique sur tout PC Windows 10/11.

**Règles UX appliquées** : un seul bouton principal par écran ; statut jamais indiqué par la couleur seule (icône + texte : ✔ ⚠ ⛔) ;
erreurs expliquées en français simple avec une proposition de correction ; aucune donnée modifiée silencieusement ;
cibles cliquables ≥ 32 px ; animation courte (interrupteur de thème 100 ms) ; drapeaux dessinés (pas d'emoji).
