# PavStat 📊

> **Tableaux statistiques et épidémiologiques de qualité publication en Python**  
> L'équivalent complet et documenté en Python des packages R incontournables : **`compareGroups`**, **`gtsummary`** (`tbl_summary`) et **`epitools`**.

[![Python 3.9+](https://img.shields.io/badge/python-3.9+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Biostatistics](https://img.shields.io/badge/Domain-Biostatistics%20%26%20Epidemiology-green.svg)]()

---

## 🎯 Pourquoi PavStat ?

Dans la recherche médicale, clinique et épidémiologique, la création de la fameuse **Table 1** (caractéristiques démographiques et cliniques) et des **tableaux bivariés avec Odds Ratios / Risques Relatifs / Rapports de Prévalence** nécessite souvent de passer par R.

`PavStat` apporte enfin à la communauté Python un outil tout-en-un, rigoureux et automatisé :
- **Décision statistique intelligente** : Détection automatique des variables continues vs catégorielles, et sélection robuste des tests paramétriques (Student, ANOVA) ou non-paramétriques (Mann-Whitney, Kruskal-Wallis).
- **Mesures d'association épidémiologiques** : Calcul vectorisé des **Risques Relatifs (RR)**, **Rapports de Prévalence (RP)** et **Odds Ratios (OR)** avec intervalles de confiance à 95% (méthode de Wald) et $p$-values exactes mid-$p$ (alignées sur `epitools`).
- **Prise en charge des variables continues dans les modèles logistiques** : Calcul de l'OR par unité continue comme dans `compareGroups(show.ratio=TRUE)`.
- **Préservation stricte des facteurs** : Respect de l'ordre d'origine des catégories (`pd.Categorical`) ou de l'ordre naturel d'apparition (évite l'inversion silencieuse des niveaux de référence).
- **Exportation Publication-Ready en un clic** :
  - 📝 **Markdown** (GitHub / Jupyter / Quarto) avec préservation des zéros décimaux (`0.010`, `1.00`).
  - 📄 **Microsoft Word (`.docx`)** avec en-têtes ombrés, styles médicaux et bordures professionnelles.
  - 🌐 **HTML interactif & stylisé** prêt à être intégré dans des rapports ou des applications web.

---

## 📦 Installation

```bash
pip install PavStat
```

*Ou directement depuis le dépôt source :*
```bash
git clone https://github.com/votre-compte/PavStat.git
cd PavStat
pip install .
```

---

## 🚀 Guide de démarrage rapide

### 1. Tableau Descriptif Univarié (Table 1)

```python
import pandas as pd
import Pav as pv

# Charger vos données cliniques / d'enquête
df = pd.read_excel("hdv2003.xlsx")

# Tableau 1 descriptif
tab1 = pv.describe_table(df, vars=["age", "sexe", "sport", "qualif"])
print(pv.to_markdown(tab1, title="Caractéristiques de la population"))
```

### 2. Tableau Comparatif Bivarié (avec Risques Relatifs ou Odds Ratios)

```python
# Comparaison selon le sexe avec Risques Relatifs (RR) et IC95%
tab_rr = pv.compare_table(
    df,
    group="sexe",
    vars=["sport", "cinema", "nivetud"],
    ratio="RR"
)
print(pv.to_markdown(tab_rr, title="Comparaison par sexe (RR)"))

# Comparaison avec Odds Ratios (OR) pour variables qualitatives et quantitatives
tab_or = pv.compare_table(
    df,
    group="sexe",
    vars=["age", "trav.satisf"],
    ratio="OR"
)
print(pv.to_markdown(tab_or, title="Comparaison par sexe (OR)"))
```

### 3. Comparaison multi-groupes (>2 groupes)

```python
# ANOVA / Kruskal-Wallis et Chi-2 automatiques
tab_multi = pv.compare_table(
    df,
    group="qualif",
    vars=["age", "sport"]
)
print(pv.to_markdown(tab_multi, title="Comparaison selon la qualification professionnelle"))
```

### 4. Exportation en Word (.docx), HTML et Markdown

```python
# Export vers un document Word soigné pour votre manuscrit / thèse
pv.to_docx(
    tab_rr,
    path="Tableau_1_Publication.docx",
    title="Tableau 1 : Analyse bivariée selon le sexe",
    footnote="RR = Risque Relatif (Wald) ; IC95% ; p.ratio = test exact mid-p ; p.overall = test du Chi-2."
)

# Export vers une page HTML
pv.to_html(
    tab_rr,
    path="Tableau_1.html",
    title="Tableau 1 : Analyse bivariée selon le sexe"
)
```

---

## 📊 Tableau de correspondance des tests statistiques

| Type de variable | 2 groupes (binaire) | > 2 groupes (multi-groupes) | Paramètres / Options |
| :--- | :--- | :--- | :--- |
| **Quantitative Normale** | Test $t$ de Welch (`stats.ttest_ind`) | ANOVA à 1 facteur (`stats.f_oneway`) | `method="param"` ou auto |
| **Quantitative Asymétrique** | Test de Mann-Whitney $U$ | Test de Kruskal-Wallis | `method="non-param"` ou auto |
| **Qualitative / Binaire** | Test du Chi-2 / Test exact de Fisher | Test du Chi-2 (avec Monte-Carlo optionnel) | `chisq_test2()` |
| **Mesures d'association** | **RR**, **OR**, **RP** (IC95% Wald + mid-$p$) | Tests globaux $p.	ext{overall}$ | `ratio="RR"` ou `"OR"` ou `"RP"` |

---

## 🧪 Validation & Tests

Le package est testé systématiquement contre les packages R de référence (`compareGroups 4.9.0`, `gtsummary 2.0`, `epitools 0.5-10.1`).

Pour exécuter la suite de tests automatisée :
```bash
pytest tests/
```

---

## 👤 Auteur & Conception

**SALLAN Konrad Pavlov**  
*Épidémiologiste & Data Analyst (R et Python)*  
Spécialiste de l'analyse biostatistique, des études épidémiologiques et de l'automatisation des rapports de recherche clinique.

---

## 📄 Licence

**Concepteur & Auteur principal :** SALLAN Konrad Pavlov

Ce projet est sous licence libre **MIT**. Vous êtes libre de l'utiliser, de le modifier et de l'intégrer dans vos travaux de recherche et projets professionnels.
