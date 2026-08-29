# Documentation Technique & Référence de l'API — `PavStat` 📖

> **Package :** `PavStat` (v0.3.0)  
> **Auteur :** SALLAN Konrad Pavlov (*Épidémiologiste & Data Analyst R/Python*)  
> **Licence :** MIT  
> **Compatibilité :** Python 3.9, 3.10, 3.11, 3.12+  

---

## 📦 Guide d'Installation

### Depuis GitHub (une fois le dépôt en ligne)
```bash
pip install git+https://github.com/votre-nom-utilisateur/PavStat.git
```

### Depuis PyPI (si le package y est publié)
```bash
pip install PavStat
```
Aucune version à préciser : `pip` installe automatiquement la dernière version disponible.

### Depuis le dossier source local (mode développement)
```bash
pip install -e .
```

### Vérification rapide
```bash
python -c "import Pav as pv; print('Installation réussie :', pv.__version__, 'par', pv.__author__)"
```

---

## Table des Matières
1. [Vue d'ensemble de l'Architecture](#1-vue-densemble-de-larchitecture)
2. [Module `Pav.tables`](#2-module-pavtables)
   - [`describe_table()`](#describe_table)
   - [`compare_table()`](#compare_table)
3. [Module `Pav.export`](#3-module-pavexport)
   - [`to_markdown()`](#to_markdown)
   - [`save_markdown()`](#save_markdown)
   - [`to_docx()`](#to_docx)
   - [`to_html()`](#to_html)
4. [Module `Pav.utils`](#4-module-pavutils)
   - [`confinterval()`](#confinterval)
   - [`chisq_test2()`](#chisq_test2)
   - [`format2()`](#format2)
5. [Méthodes et Formules Statistiques](#5-méthodes-et-formules-statistiques)
6. [Exemples Pratiques Complets](#6-exemples-pratiques-complets)

---

## 1. Vue d'ensemble de l'Architecture

Le package `PavStat` est structuré en trois sous-modules complémentaires :

```
Pav/
│
├── tables.py   ──► Fonctions principales d'analyse descriptive et bivariée
├── export.py   ──► Moteurs de rendu et d'export (Markdown, Word .docx, HTML)
└── utils.py    ──► Moteur de calcul statistique (tests chi2, IC95%, formattage)
```

---

## 2. Module `Pav.tables`

### `describe_table()`

Génère un tableau descriptif univarié complet pour caractériser un échantillon d'étude (Table 1 des manuscrits scientifiques).

#### Signature
```python
Pav.describe_table(
    df: pd.DataFrame,
    vars: Optional[Iterable[str]] = None,
    method: Union[str, Dict[str, str]] = "auto",
    show_ci: bool = False,
    show_na: bool = False,
) -> pd.DataFrame
```

#### Paramètres

| Paramètre | Type | Défaut | Description |
| :--- | :--- | :--- | :--- |
| `df` | `pd.DataFrame` | *Requis* | Le jeu de données à analyser (DataFrame pandas). |
| `vars` | `list[str]` ou `None` | `None` | Liste des noms de variables à inclure. Si `None`, toutes les colonnes du DataFrame sont traitées. |
| `method` | `str` ou `dict` | `"auto"` | Stratégie de calcul ("auto", "param", "non-param"). |
| `split_columns` | `bool` | `True` | Si `True` (par défaut) : génère deux colonnes distinctes : `Effectif (n)` et `Fréquence (%)`. Si `False` : combine dans une colonne `Valeur` (ex: `1101 (55.0%)`). | Stratégie de calcul pour les variables quantitatives continues :<br>• `"auto"` : Sélection automatique (normale si \|skewness\| $\le 1$ et \|kurtosis\| $\le 2$ pour $N \ge 30$, Shapiro-Wilk si $N < 30$).<br>• `"param"` : Force Moyenne $\pm$ Écart-Type.<br>• `"non-param"` : Force Médiane [Q1 - Q3].<br>• Dictionnaire : ex. `{"age": "param", "duree_sejour": "non-param"}`. |
| `show_ci` | `bool` | `False` | Si `True`, affiche les intervalles de confiance à 95% pour les moyennes et médianes. |
| `show_na` | `bool` | `False` | Si `True`, affiche explicitement une ligne `"Manquant (NA), n (%)"` pour les variables ayant des valeurs manquantes. |

#### Valeur de retour
- **`pd.DataFrame`** : Tableau à 2 colonnes (`Caractéristique`, `Valeur`).

#### Exemple d'utilisation
```python
import pandas as pd
import Pav as pv

df = pd.read_excel("data_clinique.xlsx")
tab1 = pv.describe_table(df, vars=["age", "sexe", "tabagisme"], show_ci=True)
print(tab1)
```

---

### `compare_table()`

Génère un tableau comparatif bivarié stratifié par une variable de groupe, avec sélection automatique des tests d'hypothèse et calcul vectorisé des mesures d'association (RR, OR, RP).

#### Signature
```python
Pav.compare_table(
    df: pd.DataFrame,
    group: str,
    vars: Optional[Iterable[str]] = None,
    ratio: str = "RR",
    ref_level: Optional[Any] = None,
    conf_level: float = 0.95,
    method: Union[str, Dict[str, str]] = "auto",
    show_ratio: bool = True,
    show_ci: bool = False,
    show_na: bool = False,
) -> pd.DataFrame
```

#### Paramètres

| Paramètre | Type | Défaut | Description |
| :--- | :--- | :--- | :--- |
| `df` | `pd.DataFrame` | *Requis* | Le jeu de données à analyser. |
| `group` | `str` | *Requis* | Nom de la variable de regroupement (colonne de stratification). |
| `vars` | `list[str]` ou `None` | `None` | Liste des variables explicatives à tester par rapport à `group`. |
| `ratio` | `{"RR", "OR", "RP"}` | `"RR"` | Type de ratio d'association épidémiologique (si groupe binaire) :<br>• `"RR"` : Risque Relatif (méthode de Wald).<br>• `"OR"` : Odds Ratio (Wald pour catégoriel, régression logistique pour quantitatif).<br>• `"RP"` : Rapport de Prévalence (étude transversale). |
| `ref_level` | `Any` ou `None` | `None` | Modalité de référence de la variable explicative pour le calcul des ratios. Si `None`, la 1ère modalité d'origine est utilisée. |
| `conf_level` | `float` | `0.95` | Niveau de confiance pour les intervalles de confiance (0.95 = IC 95%). |
| `method` | `str` ou `dict` | `"auto"` | Stratégie de calcul ("auto", "param", "non-param"). |
| `split_columns` | `bool` | `True` | Si `True` (par défaut) : génère deux colonnes distinctes : `Effectif (n)` et `Fréquence (%)`. Si `False` : combine dans une colonne `Valeur` (ex: `1101 (55.0%)`). | `"auto"`, `"param"` (Welch $t$-test / ANOVA), ou `"non-param"` (Mann-Whitney / Kruskal-Wallis). |
| `show_ratio` | `bool` | `True` | Affiche ou masque les colonnes de ratio (RR/OR/RP) et $p.\text{ratio}$. |
| `show_ci` | `bool` | `False` | Affiche les intervalles de confiance sur les statistiques de chaque groupe. |
| `show_na` | `bool` | `False` | Affiche une ligne dédiée au décompte des données manquantes. |

#### Valeur de retour
- **`pd.DataFrame`** : Tableau comparatif avec colonnes de groupes, ratios épidémiologiques avec IC95%, $p.\text{ratio}$ et $p.\text{overall}$.

#### Exemple d'utilisation
```python
# Comparaison avec Risque Relatif
tab_comp = pv.compare_table(
    df,
    group="deces_hospitalier",
    vars=["age", "sexe", "comorbidite"],
    ratio="RR",
    conf_level=0.95
)
```

---

## 3. Module `Pav.export`

### `to_markdown()`
```python
Pav.to_markdown(df: pd.DataFrame, title: Optional[str] = None) -> str
```
- **Description** : Convertit un tableau au format Markdown GitHub-Flavored en préservant strictement la précision décimale (ne supprime pas les zéros finaux comme `"0.010"` ou `"1.000"`).
- **Retourne** : Une chaîne de caractères contenant le tableau Markdown.

### `save_markdown()`
```python
Pav.save_markdown(df: pd.DataFrame, path: Union[str, Path], title: Optional[str] = None) -> None
```
- **Description** : Enregistre le tableau Markdown dans un fichier texte encodé en UTF-8.

### `to_docx()`
```python
Pav.to_docx(
    df: pd.DataFrame,
    path: Union[str, Path],
    title: Optional[str] = None,
    footnote: Optional[str] = None,
) -> None
```
- **Description** : Exporte le tableau dans un document Microsoft Word (`.docx`) avec mise en page médicale professionnelle :
  - Ligne d'en-tête ombrée en gris (`#D9D9D9`) et texte en gras.
  - Lignes de titre de variable surlignées en gris très clair (`#F2F2F2`).
  - Bordures de grille fines et alignement typographique soigné.
  - Note de bas de page en italique pour expliciter les acronymes et tests.

### `to_html()`
```python
Pav.to_html(
    df: pd.DataFrame,
    path: Optional[Union[str, Path]] = None,
    title: Optional[str] = None,
    footnote: Optional[str] = None,
) -> str
```
- **Description** : Génère un document HTML responsive avec feuille de style CSS intégrée, prêt pour intégration dans un tableau de bord, une application web ou un rapport interactif.
- **Retourne** : Le code source HTML sous forme de chaîne de caractères.

---

## 4. Module `Pav.utils`

### `confinterval()`
```python
Pav.confinterval(
    x: np.ndarray,
    method: str = "param",
    conf_level: float = 0.95,
) -> ConfInterval
```
- **Description** : Calcule l'intervalle de confiance pour une série quantitative :
  - Si `method="param"` : Moyenne $\pm t_{1-\alpha/2, n-1} \times \frac{s}{\sqrt{n}}$.
  - Si `method="non-param"` : Médiane avec méthode binomiale exacte d'ordre $qbinom(\alpha/2, n, 0.5)$.
- **Retourne** : Un objet `ConfInterval(center, lower, upper)`.

### `chisq_test2()`
```python
Pav.chisq_test2(
    obj: np.ndarray,
    chisq_test_perm: bool = False,
    chisq_test_B: int = 2000,
    chisq_test_seed: Optional[int] = None,
) -> float
```
- **Description** : Test d'indépendance pour tableau de contingence avec bascule automatique :
  - Si tous les effectifs théoriques $\ge 5$ : Test du Chi-2 de Pearson (avec correction de continuité de Yates pour les tables $2 \times 2$).
  - Si un effectif théorique $< 5$ : Test exact de Fisher pour les tables $2 \times 2$, ou test du Chi-2 par permutations Monte-Carlo ($B=2000$) si demandé.
- **Retourne** : La $p$-value sous forme de flottant.

### `format2()`
```python
Pav.format2(
    x: Union[float, np.ndarray],
    digits: Optional[int] = None,
    stars: bool = False,
) -> Union[str, np.ndarray]
```
- **Description** : Formatage des nombres et $p$-values selon les conventions biomédicales :
  - Si $p < 0.001$, affiche automatiquement `"<0.001"`.
  - Maintient les zéros décimaux à droite (`1.50` reste `1.50` avec `digits=2`).
  - Ajoute optionnellement des étoiles de significativité (`*`, `**`, `***`).

---

## 5. Méthodes et Formules Statistiques

### Risque Relatif (RR) et Rapport de Prévalence (RP) — Méthode de Wald
Pour une table $2 \times 2$ :
$$RR = \frac{p_1}{p_0} = \frac{a_1 / (a_1 + b_1)}{a_0 / (a_0 + b_0)}$$
Variance asymptotique de $\ln(RR)$ :
$$\text{Var}(\ln RR) = \frac{b_1}{a_1 (a_1 + b_1)} + \frac{b_0}{a_0 (a_0 + b_0)}$$
Intervalle de confiance à $95\%$ :
$$\text{IC}_{95\%} = \exp\left(\ln(RR) \pm 1.96 \times \sqrt{\text{Var}(\ln RR)}\right)$$

### Odds Ratio (OR) — Méthode de Wald
$$OR = \frac{a_1 \cdot b_0}{a_0 \cdot b_1}$$
$$\text{Var}(\ln OR) = \frac{1}{a_1} + \frac{1}{b_1} + \frac{1}{a_0} + \frac{1}{b_0}$$
$$\text{IC}_{95\%} = \exp\left(\ln(OR) \pm 1.96 \times \sqrt{\text{Var}(\ln OR)}\right)$$

### p.ratio — Test exact mid-$p$
Équivalent exact au package R `epitools` :
$$p_{\text{mid-}p} = 2 \times \min\left(p_{\text{one-sided}}, 1 - p_{\text{one-sided}}\right)$$
où $p_{\text{one-sided}} = 0.5 \times (P(X \le a_1) - P(X \ge a_1) + 1)$.

---

## 6. Exemples Pratiques Complets

```python
import pandas as pd
import Pav as pv

# 1. Chargement des données
df = pd.read_excel("hdv2003.xlsx")

# 2. Table 1 descriptive
t1 = pv.describe_table(df, vars=["age", "sexe", "sport", "qualif"], show_ci=True)
pv.save_markdown(t1, "Tableau_1_Descriptif.md", title="Tableau 1 : Caractéristiques générales")

# 3. Analyse bivariée selon le sexe (Risques Relatifs)
t2 = pv.compare_table(df, group="sexe", vars=["sport", "cinema", "nivetud"], ratio="RR")
pv.to_docx(t2, "Tableau_2_Sexe_RR.docx", title="Tableau 2 : Facteurs associés selon le sexe (RR)")

# 4. Analyse bivariée selon la qualification (Multi-groupes)
t3 = pv.compare_table(df, group="qualif", vars=["age", "sport"])
pv.to_html(t3, path="Tableau_3_Qualif.html", title="Tableau 3 : Comparaison par Qualification")
```
