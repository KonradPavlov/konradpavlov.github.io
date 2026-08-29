# Grille de comparaison — comparegroups_py vs R (compareGroups / epitools / gtsummary)

Objectif final : un package Python documenté, équivalent à `gtsummary` +
`compareGroups`, avec OR, RR, RP, p.ratio, p.overall, IC95%, gestion des
variables à plus de 2 modalités, et export Word / HTML / Markdown — un
outil qui n'existe pas encore de façon documentée pour l'épidémio en
Python.

À remplir après avoir lancé `test_python.py` et `test_R.R` (même
numérotation TEST 1 à TEST 12).

| # | Test | Résultat Python | Résultat R | Écart constaté | Priorité | Action |
|---|------|------------------|------------|-----------------|----------|--------|
| 1 | describe_table — descriptif univarié | | | | | |
| 2 | compare_table — age (quantitatif) ~ sexe | | | | | |
| 3 | compare_table — sport (2 mod.) ~ sexe, RR | | | | | |
| 4 | compare_table — cinema (2 mod.) ~ sexe, RR | | | | | |
| 5 | compare_table — nivetud (>2 mod.) ~ sexe | | | | | |
| 6 | compare_table — trav.satisf (>2 mod.) ~ sexe, OR | | | | | |
| 7 | compare_table — age ~ qualif (groupe >2 mod.) | | | | | |
| 8 | compare_table — sport ~ qualif (groupe >2 mod.) | | | | | |
| 9 | Export Markdown | | | | | |
| 10 | Export Word | | | | | |
| 11 | Export HTML | | | | | |
| 12 | IC95% sur une moyenne | | | | | |

## Écarts déjà identifiés avant cette batterie (à confirmer/creuser)

- **Décision normal/non-normal (Shapiro-Wilk) divergente** entre Python
  et R sur `age ~ sexe` (TEST 2) : conclusion statistique opposée
  (p=0.769 non-paramétrique côté Python vs p=0.992 paramétrique côté R).
  À creuser en priorité — impact direct sur la fiabilité des résultats.
- **N affiché dans l'en-tête** : Python affiche le N total du groupe
  plutôt que le N valide (hors NA) pour la variable testée (TEST 5/6).
- **Niveau de référence par défaut** : Python trie par ordre alphabétique
  quand `ref_level=None` ; R garde l'ordre des niveaux du facteur
  d'origine. Résultat : les deux tableaux ne sont pas comparables sans
  fixer `ref_level` explicitement des deux côtés.
- **RR vs OR** : différence de mesure attendue (pas un bug), mais il faut
  vérifier si un utilisateur épidémio s'attend à voir les deux
  disponibles nativement (RP est le même calcul que RR en transversal,
  à documenter clairement pour éviter la confusion terminologique).

## Fonctionnalités absentes du package Python (déjà connues)

- [ ] `to_html()` — pas d'équivalent Python de `export2html()`
- [ ] IC95% affiché nativement dans `describe_table()`/`compare_table()`
      pour les variables quantitatives (fonction `confinterval()` existe
      déjà dans `utils.py` mais n'est pas branchée)
- [ ] Gestion explicite des valeurs manquantes dans le tableau (affichage
      "Manquant, n (%)" au lieu d'un `dropna()` silencieux)
- [ ] OR par régression logistique (comme `show.ratio=TRUE` de R) pour
      les variables quantitatives — actuellement laissé vide côté Python
- [ ] Test de Fisher exact généralisé pour tables r×c (r,c>2) —
      actuellement approximation chi² avec avertissement
- [ ] OR en méthode mid-p (comme R par défaut) en plus du Wald actuel

## Notes de session

(Ajoute ici tes observations libres au fil des tests.)
