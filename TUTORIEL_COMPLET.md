# 📘 Guide & Tutoriel Complet de A à Z : `PavStat` 📊

> **Auteur & Concepteur :** SALLAN Konrad Pavlov (*Épidémiologiste & Data Analyst R/Python*)  
> **Version du package :** 0.3.0  
> **Licence :** MIT (Open-source)  
> **Documentation officielle pour :** Positron IDE, Jupyter, Quarto, VS Code, Python 3.9+  

---

## 📑 Sommaire
1. [Clarification des dossiers et fichiers sur votre ordinateur](#1-clarification-des-dossiers-et-fichiers)
2. [Étape 1 : Installation du package dans Positron](#2-étape-1--installation-du-package-dans-positron)
3. [Étape 2 : Chargement et vérification](#3-étape-2--chargement-et-vérification)
4. [Étape 3 : Chargement et exploration des données](#4-étape-3--chargement-et-exploration-des-données)
5. [Étape 4 : Tableau Descriptif Univarié (`describe_table`) — La "Table 1"](#5-étape-4--tableau-descriptif-univarié-describe_table)
   - [4.1 Format standard avec colonnes séparées ($n$ et %)](#41-format-standard-avec-colonnes-séparées-n-et-)
   - [4.2 Avec Intervalles de Confiance à 95% (`show_ci=True`)](#42-avec-intervalles-de-confiance-à-95-show_citrue)
   - [4.3 Affichage des valeurs manquantes (`show_na=True`)](#43-affichage-des-valeurs-manquantes-show_natrue)
   - [4.4 Forcer la méthode paramétrique ou non-paramétrique](#44-forcer-la-méthode-paramétrique-ou-non-paramétrique)
6. [Étape 5 : Tableau Comparatif Bivarié (`compare_table`)](#6-étape-5--tableau-comparatif-bivarié-compare_table)
   - [5.1 Groupe binaire avec Risques Relatifs (`ratio="RR"`)](#51-groupe-binaire-avec-risques-relatifs-ratiorr)
   - [5.2 Groupe binaire avec Odds Ratios (`ratio="OR"`)](#52-groupe-binaire-avec-odds-ratios-ratioor)
   - [5.3 Choix explicite du niveau de référence (`ref_level`)](#53-choix-explicite-du-niveau-de-référence-ref_level)
   - [5.4 Comparaison multi-groupes (> 2 groupes)](#54-comparaison-multi-groupes--2-groupes)
7. [Étape 6 : Exportations Publication-Ready (Word, HTML, Markdown)](#7-étape-6--exportations-publication-ready-word-html-markdown)
   - [6.1 Exportation Microsoft Word (`.docx`)](#61-exportation-microsoft-word-docx)
   - [6.2 Exportation HTML interactif](#62-exportation-html-interactif)
   - [6.3 Exportation Markdown](#63-exportation-markdown)
8. [Étape 7 : Script complet d'analyse prêt à l'emploi (Copier-Coller)](#8-étape-7--script-complet-danalyse-prêt-à-lemploi)
9. [Référence des tests statistiques appliqués](#9-référence-des-tests-statistiques-appliqués)

---

## 1. Clarification des dossiers et fichiers

Pour dissiper tout doute entre les différents dossiers présents sur votre Bureau :

| Emplacement | Rôle | Utilisation |
| :--- | :--- | :--- |
| **`Desktop\PavStat`** |  **LE DOSSIER OFFICIEL DE TRAVAIL** | C'est le dossier source complet contenant le package, les tests, la documentation et le code final à jour. |
| **`Desktop\...\PavStat_release_v0.3.0.zip`** | 📦 **L'ARCHIVE ZIP FINALE** | C'est le fichier compressé à déposer sur GitHub ou à partager. |
| **`Desktop\PavStat_BACKUP_...`** | 🛡️ **SAUVEGARDE HORODATÉE** | Copie de sécurité créée au tout début de notre travail. Vous pouvez la conserver comme historique. |

---

## 2. Étape 1 : Installation du package dans Positron

### 1. Ouvrir Positron et le terminal
- Dans Positron, ouvrez le terminal intégré avec : **`Ctrl + \``** (ou menu **Terminal** > **New Terminal**).

### 2. Installer le package
Dans le terminal de Positron, exécutez la commande suivante :

```bash
pip install -e "C:\Users\Administrateur\Desktop\PavStat"
```

*(L'option `-e` installe le package en mode éditable, ce qui signifie que toute modification future du code est immédiatement prise en compte sans réinstallation).*

---

## 3. Étape 2 : Chargement et vérification

Dans un script Python (`analyse.py`) ou un notebook Quarto dans Positron :

```python
import pandas as pd
import numpy as np
import Pav as pv

# Vérification de la version
print(f"Version installée : {pv.__version__}")
# Sortie attendue : Version installée : 0.3.0
```

---

## 4. Étape 3 : Chargement et exploration des données

Nous utiliserons le jeu de données épidémiologique et social **`hdv2003.xlsx`** (Enquête Histoire de Vie - Insee, 2000 individus, 20 variables) :

```python
# Chargement du fichier de données
chemin_fichier = r"C:\Users\Administrateur\Desktop\PavStat\hdv2003.xlsx"
df = pd.read_excel(chemin_fichier)

# Exploration rapide
print(f"Nombre d'observations (N) : {df.shape[0]}")
print(f"Nombre de variables : {df.shape[1]}")
print("\nAperçu des données :")
print(df[["age", "sexe", "sport", "cinema", "qualif", "trav.satisf"]].head())
```

---

## 5. Étape 4 : Tableau Descriptif Univarié (`describe_table`)

La fonction `pv.describe_table()` génère la **Table 1** classique des manuscrits médicaux et épidémiologiques.

### 4.1 Format standard avec colonnes séparées ($n$ et %)
Par défaut (`split_columns=True`), les effectifs et pourcentages sont présentés dans deux colonnes distinctes :

```python
# Sélection des variables d'intérêt
variables_etude = ["age", "sexe", "sport", "cinema", "qualif"]

# Génération de la table descriptive
tab1 = pv.describe_table(df, vars=variables_etude)

# Affichage console formaté
print(pv.to_markdown(tab1, title="Tableau 1 : Caractéristiques démographiques"))
```

#### Rendu obtenu :
| Caractéristique | Effectif ($n$) | Fréquence ($\%$) |
|:---|:---:|:---:|
| **N = 2000** | **2000** | **100.0%** |
| **age** | | |
| Moyenne $\pm$ ET | 48.2 $\pm$ 16.9 | — |
| Médiane [Q1 - Q3] | 48 [35 - 60] | — |
| Min - Max | 18 - 97 | — |
| **sexe** | | |
| Femme | 1101 | 55.0% |
| Homme | 899 | 45.0% |
| **sport** | | |
| Non | 1277 | 63.8% |
| Oui | 723 | 36.2% |

---

### 4.2 Avec Intervalles de Confiance à 95% (`show_ci=True`)
Permet d'ajouter les intervalles de confiance à 95% calculés selon la méthode paramétrique de Student pour la moyenne et binomiale exacte pour la médiane :

```python
tab1_ci = pv.describe_table(
    df,
    vars=["age", "sexe", "sport"],
    show_ci=True
)
print(pv.to_markdown(tab1_ci, title="Tableau 1 avec IC à 95%"))
```
- Affiche : `Moyenne (IC95%)` $\rightarrow$ `48.2 [47.4 - 48.9]`
- Affiche : `Médiane (IC95%)` $\rightarrow$ `48 [47 - 49]`

---

### 4.3 Affichage des valeurs manquantes (`show_na=True`)
Pour les variables contenant des données manquantes (ex: `trav.satisf` qui a 1052 non-réponses) :

```python
tab1_na = pv.describe_table(
    df,
    vars=["trav.satisf", "qualif"],
    show_na=True
)
print(pv.to_markdown(tab1_na, title="Tableau avec valeurs manquantes explicites"))
```
- Ajoute une ligne explicite : `Manquant (NA)` $\rightarrow$ `1052 (52.6%)`.

---

### 4.4 Forcer la méthode paramétrique ou non-paramétrique
Vous pouvez choisir globalement ou variable par variable si vous souhaitez afficher la moyenne ou la médiane :

```python
# Forcer en non-paramétrique (Médiane [IQR])
tab1_nonparam = pv.describe_table(df, vars=["age"], method="non-param")

# Spécifier par variable via un dictionnaire
tab1_custom = pv.describe_table(
    df,
    vars=["age", "heures_tv"],
    method={"age": "param", "heures_tv": "non-param"}
)
```

---

## 6. Étape 5 : Tableau Comparatif Bivarié (`compare_table`)

La fonction `pv.compare_table()` croise une variable d'intérêt (stratification) avec les variables explicatives, applique automatiquement les tests statistiques et calcule les ratios épidémiologiques.

---

### 5.1 Groupe binaire avec Risques Relatifs (`ratio="RR"`) ou Rapports de Prévalence (`ratio="RP"`)
Idéal pour les études de cohorte ou transversales :

```python
tab_rr = pv.compare_table(
    df,
    group="sexe",                       # Variable de groupe binaire
    vars=["sport", "cinema", "nivetud"], # Variables à croiser
    ratio="RR",                         # Risque Relatif (Wald)
    conf_level=0.95                     # IC95%
)

print(pv.to_markdown(tab_rr, title="Facteurs associés selon le sexe (Risques Relatifs)"))
```

#### Rendu obtenu :
| Variable | Homme | Femme | RR | p.ratio | p.overall |
|:---|:---:|:---:|:---:|:---:|:---:|
| | **N=899** | **N=1101** | | | |
| **sport:** | | | | | **<0.001** |
| Non | 530 (41.5%) | 747 (58.5%) | Ref. | Ref. | |
| Oui | 369 (51.0%) | 354 (49.0%) | 0.84 [0.77;0.91] | <0.001 | |
| **cinema:** | | | | | **0.208** |
| Non | 542 (46.2%) | 632 (53.8%) | Ref. | Ref. | |
| Oui | 357 (43.2%) | 469 (56.8%) | 1.05 [0.97;1.14] | 0.193 | |

---

### 5.2 Groupe binaire avec Odds Ratios (`ratio="OR"`)
Pour les études cas-témoins et modèles de régression logistique :

```python
tab_or = pv.compare_table(
    df,
    group="sexe",
    vars=["age", "sport", "trav.satisf"],
    ratio="OR"
)

print(pv.to_markdown(tab_or, title="Facteurs associés selon le sexe (Odds Ratios)"))
```
- Pour la variable continue `age` : Le package calcule automatiquement l'**OR continu par régression logistique** : `1.00 [0.99;1.01]`, $p=0.992$.
- Pour la variable qualitative `sport` : Calcule l'**OR catégoriel de Wald** : `0.68 [0.57;0.82]`, $p<0.001$.

---

### 5.3 Choix explicite du niveau de référence (`ref_level`)
Par défaut, le package respecte l'ordre naturel des facteurs ou la 1ère modalité d'apparition. Vous pouvez changer manuellement la référence :

```python
tab_ref = pv.compare_table(
    df,
    group="sexe",
    vars=["sport"],
    ratio="OR",
    ref_level="Oui"  # La modalité "Oui" devient la catégorie de référence (Ref.)
)
print(pv.to_markdown(tab_ref))
```

---

### 5.4 Comparaison multi-groupes (> 2 groupes)
Si la variable de groupe comporte 3 modalités ou plus (ex: `qualif`), le package applique :
- **ANOVA à 1 facteur** (`stats.f_oneway`) ou **Kruskal-Wallis** pour les variables quantitatives.
- **Chi-2 de Pearson** (`chisq_test2`) pour les variables qualitatives.

```python
tab_multi = pv.compare_table(
    df,
    group="qualif",  # Variable à 7 modalités
    vars=["age", "sport"]
)

print(pv.to_markdown(tab_multi, title="Comparaison selon la qualification professionnelle"))
```

---

## 7. Étape 6 : Exportations Publication-Ready (Word, HTML, Markdown)

### 6.1 Exportation Microsoft Word (`.docx`)
Génère un document Word formatté selon les standards biomédicaux :

```python
pv.to_docx(
    df=tab_rr,
    path="Tableau_2_Publication.docx",
    title="Tableau 2 : Facteurs associés à la pratique sportive selon le sexe",
    footnote="RR = Risque Relatif (méthode de Wald) ; IC95% = Intervalle de confiance à 95% ; p.ratio = test exact mid-p ; p.overall = test du Chi-2."
)
print(" Document Word créé avec succès !")
```

---

### 6.2 Exportation HTML interactif
Génère un fichier HTML responsive avec styles CSS élégants intégrés :

```python
pv.to_html(
    df=tab_rr,
    path="Tableau_2_Rapport.html",
    title="Tableau 2 : Analyse Bivariée par Sexe",
    footnote="Source : Données Enquête hdv2003 - Insee."
)
print(" Page HTML générée avec succès !")
```

---

### 6.3 Exportation Markdown
Pour vos rapports Quarto (`.qmd`), Jupyter ou GitHub :

```python
pv.save_markdown(
    df=tab1,
    path="Tableau_1_Descriptif.md",
    title="Tableau 1 : Caractéristiques démographiques"
)
```

---

## 8. Étape 7 : Script complet d'analyse prêt à l'emploi

Créez un fichier `run_analyse.py` dans Positron et collez ce script :

```python
# ==============================================================================
# SCRIPT DE RECHERCHE CLINIQUE & ÉPIDÉMIOLOGIQUE AVEC COMPAREGROUPS-PY
# ==============================================================================

import pandas as pd
import Pav as pv

# 1. Chargement des données
chemin = r"C:\Users\Administrateur\Desktop\PavStat\hdv2003.xlsx"
df = pd.read_excel(chemin)
print(f"[1/4] Jeu de données chargé : {df.shape[0]} sujets, {df.shape[1]} variables.")

# 2. Table 1 descriptive (Effectifs et Pourcentages séparés)
tab_desc = pv.describe_table(
    df,
    vars=["age", "sexe", "sport", "cinema", "qualif"],
    split_columns=True,
    show_ci=False
)
print("\n[2/4] Table 1 générée avec succès.")

# 3. Table 2 comparatif bivarié (Risques Relatifs)
tab_biv = pv.compare_table(
    df,
    group="sexe",
    vars=["age", "sport", "cinema", "nivetud"],
    ratio="RR",
    conf_level=0.95
)
print("\n[3/4] Table 2 générée avec succès.")

# 4. Exportations automatiques en Word et HTML
pv.to_docx(
    tab_desc,
    path="Tableau_1_Descriptif.docx",
    title="Tableau 1 : Caractéristiques générales de la population"
)

pv.to_docx(
    tab_biv,
    path="Tableau_2_Facteurs_Associes_Sexe.docx",
    title="Tableau 2 : Facteurs associés selon le sexe",
    footnote="RR = Risque Relatif (Wald) ; IC95% ; p.ratio = test exact mid-p ; p.overall = test du Chi-2 / test t de Welch."
)

pv.to_html(
    tab_biv,
    path="Tableau_2_Facteurs_Associes_Sexe.html",
    title="Tableau 2 : Facteurs associés selon le sexe"
)

print("\n[4/4]  Tous les rapports Word (.docx) et HTML ont été créés avec succès !")
```

---

## 9. Référence des tests statistiques appliqués

| Type de situation | Test appliqué par `PavStat` | Fonction / Équivalence R |
| :--- | :--- | :--- |
| **Continue normale (2 groupes)** | Test $t$ de Welch (variances inégales) | `stats.ttest_ind(equal_var=False)` $\approx$ `t.test(var ~ grp)` |
| **Continue asymétrique (2 groupes)** | Test de Mann-Whitney $U$ | `stats.mannwhitneyu()` $\approx$ `wilcox.test(var ~ grp)` |
| **Continue normale (> 2 groupes)** | ANOVA à 1 facteur | `stats.f_oneway()` $\approx$ `aov(var ~ grp)` |
| **Continue asymétrique (> 2 groupes)**| Test de Kruskal-Wallis | `stats.kruskal()` $\approx$ `kruskal.test(var ~ grp)` |
| **Qualitative (Effectifs $\ge 5$)** | Test du Chi-2 de Pearson avec correction de Yates | `chisq_test2()` $\approx$ `chisq.test()` |
| **Qualitative (Petits effectifs $< 5$)**| Test exact de Fisher ($2 \times 2$) ou Monte-Carlo ($r \times c$) | `stats.fisher_exact()` $\approx$ `fisher.test()` |
| **Risques Relatifs / Rapports de Prévalence** | Estimation de Wald avec IC95% asymptotique | `_riskratio_wald()` $\approx$ `epitools::riskratio()` |
| **Odds Ratios catégoriels** | Estimation de Wald avec IC95% asymptotique | `_oddsratio_wald()` $\approx$ `epitools::oddsratio()` |
| **Odds Ratios continus** | Régression logistique binaire | `statsmodels.api.Logit` $\approx$ `compareGroups(show.ratio=TRUE)` |
| **$p$-value du ratio ($p.\text{ratio}$)** | Test exact mid-$p$ | `_ratio_pvalue()` $\approx$ `epitools::tab2by2.test()` |

---
*Ce tutoriel fait partie intégrante de la distribution officielle de `PavStat` (v0.3.0).*
