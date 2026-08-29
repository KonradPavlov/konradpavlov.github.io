"""
Pav.tables — Génération de tableaux statistiques descriptifs et comparatifs
Auteur : SALLAN Konrad Pavlov
 statistiques descriptifs et comparatifs
de qualité publication, équivalent en Python aux packages R `compareGroups`, `gtsummary` et `epitools`.

Fonctions principales :
- describe_table(df, vars=None, method="auto", split_columns=True, show_ci=False, show_na=False)
- compare_table(df, group, vars=None, ratio="RR", ref_level=None, conf_level=0.95, method="auto", show_ratio=True, show_ci=False, show_na=False)
"""
from __future__ import annotations

import warnings
from typing import Iterable, Optional, Union, Dict, Any, List

import numpy as np
import pandas as pd
from scipy import stats
import statsmodels.api as sm

from .utils import chisq_test2, confinterval, format2

RatioType = str  # "RR", "OR", ou "RP" (Rapport de Prévalence)


# ---------------------------------------------------------------------------
# Détection et typage des variables
# ---------------------------------------------------------------------------
def _is_quantitative(s: pd.Series) -> bool:
    """Détecte si une série est quantitative continue."""
    return pd.api.types.is_numeric_dtype(s) and s.nunique(dropna=True) > 5


def _is_normal(s: pd.Series, alpha: float = 0.05) -> bool:
    """
    Évalue si une variable quantitative suit une distribution approximativement normale.
    - Si n < 3 : considéré normal par défaut.
    - Si 3 <= n < 30 : Test de Shapiro-Wilk au seuil alpha (0.05).
    - Si n >= 30 (grand échantillon, Théorème Central Limite) : vérification de l'asymétrie
      (skewness) et de l'aplatissement (kurtosis).
      Une distribution avec |skewness| <= 1.0 et |kurtosis| <= 2.0 est considérée
      suffisamment symétrique et normale pour les tests paramétriques (Student / ANOVA).
    """
    x = s.dropna().values.astype(float)
    n = len(x)
    if n < 3:
        return True
    if n < 30:
        try:
            _, p = stats.shapiro(x)
            return p > alpha
        except Exception:
            return True
    else:
        skew = float(stats.skew(x, bias=False))
        kurt = float(stats.kurtosis(x, bias=False))
        return abs(skew) <= 1.0 and abs(kurt) <= 2.0


def _resolve_is_param(v: str, s: pd.Series, method: Union[str, Dict[str, str]]) -> bool:
    """Détermine si la méthode d'analyse pour la variable v doit être paramétrique ou non."""
    if isinstance(method, dict):
        m = method.get(v, "auto").lower()
    elif isinstance(method, str):
        m = method.lower()
    else:
        m = "auto"

    if m in ("param", "parametric", "normal", "1", 1):
        return True
    elif m in ("non-param", "non-parametric", "nonparam", "nonnormal", "2", 2):
        return False
    else:
        return _is_normal(s)


def _get_variable_levels(s: pd.Series, ref_level: Optional[Any] = None) -> List[Any]:
    """
    Extrait les modalités d'une variable catégorielle en préservant :
    1. L'ordre des catégories si la série est de type pd.Categorical
    2. L'ordre naturel d'apparition si la série est de type objet/texte (aligné sur R)
    3. Le niveau de référence personnalisé si spécifié via ref_level.
    """
    if isinstance(s.dtype, pd.CategoricalDtype):
        levels = [cat for cat in s.cat.categories if cat in s.dropna().values]
    else:
        levels = list(pd.unique(s.dropna()))

    if ref_level is not None:
        if ref_level not in levels:
            raise ValueError(
                f"Le niveau de référence '{ref_level}' n'existe pas parmi les modalités disponibles: {levels}"
            )
        levels = [ref_level] + [lvl for lvl in levels if lvl != ref_level]

    return levels


# ---------------------------------------------------------------------------
# Tableau descriptif univarié (style gtsummary / compareGroups)
# ---------------------------------------------------------------------------
def describe_table(
    df: pd.DataFrame,
    vars: Optional[Iterable[str]] = None,
    method: Union[str, Dict[str, str]] = "auto",
    split_columns: bool = True,
    show_ci: bool = False,
    show_na: bool = False,
) -> pd.DataFrame:
    """
    Génère un tableau descriptif univarié complet (Table 1 des caractéristiques démographiques et cliniques).
    
    Paramètres
    ----------
    df : pd.DataFrame
        Jeu de données à analyser.
    vars : list of str, optional
        Liste des variables à inclure. Par défaut, toutes les colonnes.
    method : str ou dict, default "auto"
        "auto" (sélection automatique), "param" (Moyenne ± ET), ou "non-param" (Médiane [IQR]).
    split_columns : bool, default True
        Si True (recommandé pour les publications/thèses) : sépare les Effectifs (n)
        et les Fréquences (%) en deux colonnes distinctes :
        ['Caractéristique', 'Effectif (n)', 'Fréquence (%)'].
        Si False : combine dans une seule colonne ['Caractéristique', 'Valeur'] (ex: '1101 (55.0%)').
    show_ci : bool, default False
        Si True, affiche les intervalles de confiance à 95% pour les moyennes et médianes.
    show_na : bool, default False
        Si True, affiche explicitement une ligne pour les valeurs manquantes (NA).
        
    Retourne
    -------
    pd.DataFrame
        Tableau descriptif mis en forme.
    """
    vars = list(vars) if vars is not None else list(df.columns)
    n_total = len(df)
    rows = []

    if split_columns:
        rows.append({"Caractéristique": f"**N = {n_total}**", "Effectif (n)": f"{n_total}", "Fréquence (%)": "100.0%"})
    else:
        rows.append({"Caractéristique": f"**N = {n_total}**", "Valeur": ""})

    for v in vars:
        if v not in df.columns:
            raise KeyError(f"La colonne '{v}' n'existe pas dans le DataFrame fourni.")

        s = df[v]
        n_missing = int(s.isna().sum())
        n_valid = n_total - n_missing

        if _is_quantitative(s):
            x = s.dropna().values.astype(float)
            is_param = _resolve_is_param(v, s, method)
            mean, sd = np.mean(x), np.std(x, ddof=1) if len(x) > 1 else 0.0
            med = np.median(x)
            q1, q3 = np.percentile(x, [25, 75]) if len(x) > 0 else (np.nan, np.nan)
            mn, mx = np.min(x), np.max(x) if len(x) > 0 else (np.nan, np.nan)

            title_var = f"**{v}**" if n_missing == 0 or show_na else f"**{v}** (N={n_valid})"

            if split_columns:
                rows.append({"Caractéristique": title_var, "Effectif (n)": "", "Fréquence (%)": ""})
                if show_ci:
                    ci_param = confinterval(x, method="param", conf_level=0.95)
                    rows.append({
                        "Caractéristique": "Moyenne (IC95%)",
                        "Effectif (n)": format2(mean, 1),
                        "Fréquence (%)": f"[{format2(ci_param.lower, 1)} - {format2(ci_param.upper, 1)}]",
                    })
                    ci_nonparam = confinterval(x, method="non-param", conf_level=0.95)
                    rows.append({
                        "Caractéristique": "Médiane (IC95%)",
                        "Effectif (n)": format2(med, 0),
                        "Fréquence (%)": f"[{format2(ci_nonparam.lower, 0)} - {format2(ci_nonparam.upper, 0)}]",
                    })
                else:
                    rows.append({"Caractéristique": "Moyenne ± ET", "Effectif (n)": f"{format2(mean, 1)} ± {format2(sd, 1)}", "Fréquence (%)": "—"})
                    rows.append({"Caractéristique": "Médiane [Q1 - Q3]", "Effectif (n)": f"{format2(med, 0)} [{format2(q1, 0)} - {format2(q3, 0)}]", "Fréquence (%)": "—"})
                
                rows.append({"Caractéristique": "Min - Max", "Effectif (n)": f"{format2(mn, 0)} - {format2(mx, 0)}", "Fréquence (%)": "—"})

                if show_na and n_missing > 0:
                    pct_missing = (n_missing / n_total) * 100
                    rows.append({"Caractéristique": "Manquant (NA)", "Effectif (n)": f"{n_missing}", "Fréquence (%)": f"{format2(pct_missing, 1)}%"})
            else:
                rows.append({"Caractéristique": title_var, "Valeur": ""})
                if show_ci:
                    ci_param = confinterval(x, method="param", conf_level=0.95)
                    rows.append({"Caractéristique": "Moyenne (IC95%)", "Valeur": f"{format2(mean, 1)} [{format2(ci_param.lower, 1)} - {format2(ci_param.upper, 1)}]"})
                    ci_nonparam = confinterval(x, method="non-param", conf_level=0.95)
                    rows.append({"Caractéristique": "Médiane (IC95%)", "Valeur": f"{format2(med, 0)} [{format2(ci_nonparam.lower, 0)} - {format2(ci_nonparam.upper, 0)}]"})
                else:
                    rows.append({"Caractéristique": "Moyenne ± ET", "Valeur": f"{format2(mean, 1)} ± {format2(sd, 1)}"})
                    rows.append({"Caractéristique": "Médiane [Q1 - Q3]", "Valeur": f"{format2(med, 0)} [{format2(q1, 0)} - {format2(q3, 0)}]"})
                
                rows.append({"Caractéristique": "Min - Max", "Valeur": f"{format2(mn, 0)} - {format2(mx, 0)}"})

                if show_na and n_missing > 0:
                    pct_missing = (n_missing / n_total) * 100
                    rows.append({"Caractéristique": "Manquant (NA)", "Valeur": f"{n_missing} ({format2(pct_missing, 1)}%)"})
        else:
            levels_v = _get_variable_levels(s)
            counts = s.value_counts(dropna=True)
            title_var = f"**{v}**" if n_missing == 0 or show_na else f"**{v}** (N={n_valid})"

            if split_columns:
                rows.append({"Caractéristique": title_var, "Effectif (n)": "", "Fréquence (%)": ""})
                for level in levels_v:
                    n = int(counts.get(level, 0))
                    pct = (100 * n / n_valid) if n_valid > 0 else 0.0
                    rows.append({"Caractéristique": str(level), "Effectif (n)": f"{n}", "Fréquence (%)": f"{format2(pct, 1)}%"})

                if show_na and n_missing > 0:
                    pct_missing = (n_missing / n_total) * 100
                    rows.append({"Caractéristique": "Manquant (NA)", "Effectif (n)": f"{n_missing}", "Fréquence (%)": f"{format2(pct_missing, 1)}%"})
            else:
                title_single = f"**{v}, n (%)**" if n_missing == 0 or show_na else f"**{v}, n (%)** (N={n_valid})"
                rows.append({"Caractéristique": title_single, "Valeur": ""})
                for level in levels_v:
                    n = int(counts.get(level, 0))
                    pct = (100 * n / n_valid) if n_valid > 0 else 0.0
                    rows.append({"Caractéristique": str(level), "Valeur": f"{n} ({format2(pct, 1)}%)"})

                if show_na and n_missing > 0:
                    pct_missing = (n_missing / n_total) * 100
                    rows.append({"Caractéristique": "Manquant (NA)", "Valeur": f"{n_missing} ({format2(pct_missing, 1)}%)"})

    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Calculs Épidémiologiques : Risk Ratio (RR / RP) et Odds Ratio (OR)
# ---------------------------------------------------------------------------
def _riskratio_wald(a0: float, b0: float, a1: float, b1: float, conf_level: float = 0.95):
    total0 = a0 + b0
    total1 = a1 + b1
    if total0 == 0 or total1 == 0 or a0 == 0:
        return np.nan, np.nan, np.nan
    p0 = a0 / total0
    p1 = a1 / total1
    if p1 == 0:
        return 0.0, 0.0, np.nan

    est = p1 / p0
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        var = (b1 / (a1 * total1)) + (b0 / (a0 * total0))
        alpha = 1 - conf_level
        z = stats.norm.ppf(1 - alpha / 2)
        lo = float(np.exp(np.log(est) - z * np.sqrt(var)))
        hi = float(np.exp(np.log(est) + z * np.sqrt(var)))
    return float(est), lo, hi


def _oddsratio_wald(a0: float, b0: float, a1: float, b1: float, conf_level: float = 0.95):
    if b0 == 0 or a1 == 0 or a0 == 0 or b1 == 0:
        return np.nan, np.nan, np.nan
    est = (a1 * b0) / (a0 * b1)
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        var = (1 / a1) + (1 / b1) + (1 / a0) + (1 / b0)
        alpha = 1 - conf_level
        z = stats.norm.ppf(1 - alpha / 2)
        lo = float(np.exp(np.log(est) - z * np.sqrt(var)))
        hi = float(np.exp(np.log(est) + z * np.sqrt(var)))
    return float(est), lo, hi


def _ratio_pvalue(a0: int, b0: int, a1: int, b1: int) -> float:
    x = np.array([[a1, a0], [b1, b0]], dtype=float)
    if np.any(x < 0) or x.sum() == 0:
        return np.nan
    try:
        _, lteqtoa1 = stats.fisher_exact(x, alternative="less")
        _, gteqtoa1 = stats.fisher_exact(x, alternative="greater")
        pval1 = 0.5 * (lteqtoa1 - gteqtoa1 + 1)
        one_sided = min(pval1, 1 - pval1)
        return float(2 * one_sided)
    except Exception:
        return np.nan


def _continuous_or_logit(df: pd.DataFrame, var: str, group: str, lvl0: Any, lvl1: Any, conf_level: float = 0.95):
    """
    Calcule l'Odds Ratio (OR) par unité d'une variable quantitative continue
    via régression logistique binaire (équivalent exact à compareGroups show.ratio=TRUE).
    """
    sub = df[[var, group]].dropna()
    y = (sub[group] == lvl1).astype(int)
    X = sm.add_constant(sub[var].astype(float))
    try:
        model = sm.Logit(y, X).fit(disp=False)
        beta = model.params[var]
        se = model.bse[var]
        z = stats.norm.ppf(1 - (1 - conf_level) / 2)
        est = np.exp(beta)
        lo = np.exp(beta - z * se)
        hi = np.exp(beta + z * se)
        p_val = model.pvalues[var]
        return float(est), float(lo), float(hi), float(p_val)
    except Exception:
        return np.nan, np.nan, np.nan, np.nan


# ---------------------------------------------------------------------------
# Tableau comparatif bivarié (style compareGroups / gtsummary)
# ---------------------------------------------------------------------------
def compare_table(
    df: pd.DataFrame,
    group: str,
    vars: Optional[Iterable[str]] = None,
    ratio: RatioType = "RR",
    ref_level: Optional[Any] = None,
    conf_level: float = 0.95,
    method: Union[str, Dict[str, str]] = "auto",
    show_ratio: bool = True,
    show_ci: bool = False,
    show_na: bool = False,
) -> pd.DataFrame:
    """
    Génère un tableau comparatif bivarié par groupe (avec tests statistiques et ratios).
    
    Paramètres
    ----------
    df : pd.DataFrame
        Jeu de données.
    group : str
        Nom de la variable de regroupement (colonne de stratification).
    vars : list of str, optional
        Variables explicatives à tester. Par défaut, toutes les autres colonnes.
    ratio : {"RR", "OR", "RP"}, default "RR"
        Type de ratio d'association épidémiologique à calculer (si groupe binaire).
    ref_level : any, optional
        Niveau de référence pour la variable explicative dans le calcul des ratios.
    conf_level : float, default 0.95
        Niveau de confiance pour les intervalles (0.95 pour IC95%).
    method : str ou dict, default "auto"
        Méthode statistique ("auto", "param", "non-param").
    show_ratio : bool, default True
        Si True, affiche les colonnes de ratio (RR/OR/RP) et p.ratio.
    show_ci : bool, default False
        Si True, affiche les intervalles de confiance sur les moyennes/médianes.
    show_na : bool, default False
        Affichage explicite des valeurs manquantes.
    """
    if group not in df.columns:
        raise KeyError(f"La variable de groupe '{group}' n'existe pas dans le DataFrame.")

    vars = list(vars) if vars is not None else [c for c in df.columns if c != group]
    for v in vars:
        if v not in df.columns:
            raise KeyError(f"La colonne '{v}' n'existe pas dans le DataFrame.")

    ratio_label = "RP" if ratio.upper() == "RP" else ("OR" if ratio.upper() == "OR" else "RR")

    levels = _get_variable_levels(df[group])
    if len(levels) < 2:
        raise ValueError(f"'{group}' doit avoir au moins 2 niveaux (trouvé: {levels})")
    if len(levels) > 2:
        return _compare_table_multigroup(df, group, levels, vars, method=method, show_ci=show_ci, show_na=show_na)

    lvl0, lvl1 = levels

    n0_total = int((df[group] == lvl0).sum())
    n1_total = int((df[group] == lvl1).sum())

    rows = []
    header_dict = {
        "Variable": "",
        f"{lvl0}": f"**N={n0_total}**",
        f"{lvl1}": f"**N={n1_total}**",
    }
    if show_ratio:
        header_dict[ratio_label] = ""
        header_dict["p.ratio"] = ""
    header_dict["p.overall"] = ""
    rows.append(header_dict)

    for v in vars:
        s = df[v]
        n_missing = int(s.isna().sum())
        n_valid = len(df) - n_missing

        if _is_quantitative(s):
            x0 = df.loc[df[group] == lvl0, v].dropna().values.astype(float)
            x1 = df.loc[df[group] == lvl1, v].dropna().values.astype(float)
            all_s = pd.concat([pd.Series(x0), pd.Series(x1)])
            is_param = _resolve_is_param(v, all_s, method)
            
            if is_param:
                _, p_overall = stats.ttest_ind(x0, x1, equal_var=False)
            else:
                _, p_overall = stats.mannwhitneyu(x0, x1)

            title_var = f"**{v}**" if n_missing == 0 or show_na else f"**{v}** [N={n_valid}]"
            
            if show_ci:
                ci0 = confinterval(x0, method="param" if is_param else "non-param", conf_level=conf_level)
                ci1 = confinterval(x1, method="param" if is_param else "non-param", conf_level=conf_level)
                val0_str = f"{format2(ci0.center, 1)} [{format2(ci0.lower, 1)}-{format2(ci0.upper, 1)}]"
                val1_str = f"{format2(ci1.center, 1)} [{format2(ci1.lower, 1)}-{format2(ci1.upper, 1)}]"
            else:
                val0_str = f"{format2(np.mean(x0), 1)} ± {format2(np.std(x0, ddof=1), 1)}" if is_param else f"{format2(np.median(x0), 0)} [{format2(np.percentile(x0, 25), 0)}-{format2(np.percentile(x0, 75), 0)}]"
                val1_str = f"{format2(np.mean(x1), 1)} ± {format2(np.std(x1, ddof=1), 1)}" if is_param else f"{format2(np.median(x1), 0)} [{format2(np.percentile(x1, 25), 0)}-{format2(np.percentile(x1, 75), 0)}]"

            ratio_cont_str = ""
            p_ratio_cont_str = ""
            if show_ratio and ratio_label == "OR":
                est_or, lo_or, hi_or, p_or = _continuous_or_logit(df, v, group, lvl0, lvl1, conf_level)
                if not np.isnan(est_or):
                    ratio_cont_str = f"{format2(est_or, 2)} [{format2(lo_or, 2)};{format2(hi_or, 2)}]"
                    p_ratio_cont_str = format2(p_or, 3)

            row_dict = {
                "Variable": title_var,
                f"{lvl0}": val0_str,
                f"{lvl1}": val1_str,
            }
            if show_ratio:
                row_dict[ratio_label] = ratio_cont_str
                row_dict["p.ratio"] = p_ratio_cont_str
            row_dict["p.overall"] = format2(p_overall, 3)
            rows.append(row_dict)

            if show_na and n_missing > 0:
                n0_na = int(df.loc[df[group] == lvl0, v].isna().sum())
                n1_na = int(df.loc[df[group] == lvl1, v].isna().sum())
                na_row = {
                    "Variable": "Manquant (NA)",
                    f"{lvl0}": f"{n0_na} ({format2(100 * n0_na / n0_total, 1)}%)",
                    f"{lvl1}": f"{n1_na} ({format2(100 * n1_na / n1_total, 1)}%)",
                }
                if show_ratio:
                    na_row[ratio_label] = ""
                    na_row["p.ratio"] = ""
                na_row["p.overall"] = ""
                rows.append(na_row)
            continue

        # Variable catégorielle
        levels_v = _get_variable_levels(s, ref_level=ref_level)
        ref = levels_v[0]

        ct = pd.crosstab(df[v], df[group])
        ct = ct.reindex(index=levels_v, columns=[lvl0, lvl1], fill_value=0)
        p_overall = chisq_test2(ct.values.astype(float), chisq_test_perm=False, chisq_test_B=2000)

        title_var = f"**{v}:**" if n_missing == 0 or show_na else f"**{v}:** [N={n_valid}]"
        header_row = {
            "Variable": title_var,
            f"{lvl0}": "",
            f"{lvl1}": "",
        }
        if show_ratio:
            header_row[ratio_label] = ""
            header_row["p.ratio"] = ""
        header_row["p.overall"] = format2(p_overall, 3)
        rows.append(header_row)

        b0_ref, a0_ref = int(ct.loc[ref, lvl0]), int(ct.loc[ref, lvl1])

        for level in levels_v:
            n0, n1 = int(ct.loc[level, lvl0]), int(ct.loc[level, lvl1])
            n_level = n0 + n1
            pct0 = 100 * n0 / n_level if n_level > 0 else np.nan
            pct1 = 100 * n1 / n_level if n_level > 0 else np.nan

            if level == ref:
                ratio_str, p_ratio_str = "Ref.", "Ref."
            else:
                a1, b1 = int(ct.loc[level, lvl1]), int(ct.loc[level, lvl0])
                if ratio_label in ("RR", "RP"):
                    est, lo, hi = _riskratio_wald(a0_ref, b0_ref, a1, b1, conf_level)
                else:
                    est, lo, hi = _oddsratio_wald(a0_ref, b0_ref, a1, b1, conf_level)

                if np.isnan(est):
                    ratio_str, p_ratio_str = ".", "."
                else:
                    ratio_str = f"{format2(est, 2)} [{format2(lo, 2)};{format2(hi, 2)}]"
                    p_ratio = _ratio_pvalue(a0_ref, b0_ref, a1, b1)
                    p_ratio_str = format2(p_ratio, 3)

            cat_row = {
                "Variable": str(level),
                f"{lvl0}": f"{n0} ({format2(pct0, 1)}%)",
                f"{lvl1}": f"{n1} ({format2(pct1, 1)}%)",
            }
            if show_ratio:
                cat_row[ratio_label] = ratio_str
                cat_row["p.ratio"] = p_ratio_str
            cat_row["p.overall"] = ""
            rows.append(cat_row)

        if show_na and n_missing > 0:
            n0_na = int(df.loc[df[group] == lvl0, v].isna().sum())
            n1_na = int(df.loc[df[group] == lvl1, v].isna().sum())
            na_row = {
                "Variable": "Manquant (NA)",
                f"{lvl0}": f"{n0_na} ({format2(100 * n0_na / n0_total, 1)}%)",
                f"{lvl1}": f"{n1_na} ({format2(100 * n1_na / n1_total, 1)}%)",
            }
            if show_ratio:
                na_row[ratio_label] = ""
                na_row["p.ratio"] = ""
            na_row["p.overall"] = ""
            rows.append(na_row)

    return pd.DataFrame(rows)


def _compare_table_multigroup(
    df: pd.DataFrame,
    group: str,
    levels: list,
    vars: list,
    method: Union[str, Dict[str, str]] = "auto",
    show_ci: bool = False,
    show_na: bool = False,
) -> pd.DataFrame:
    """Tableau comparatif pour un groupe à >2 modalités."""
    n_by_level = {lvl: int((df[group] == lvl).sum()) for lvl in levels}
    rows = []
    header = {"Variable": ""}
    for lvl in levels:
        header[str(lvl)] = f"**N={n_by_level[lvl]}**"
    header["p.overall"] = ""
    rows.append(header)

    for v in vars:
        s = df[v]
        n_missing = int(s.isna().sum())
        n_valid = len(df) - n_missing

        if _is_quantitative(s):
            groups_x = [df.loc[df[group] == lvl, v].dropna().values.astype(float) for lvl in levels]
            all_x = pd.Series(np.concatenate(groups_x))
            is_param = _resolve_is_param(v, all_x, method)
            
            if is_param:
                _, p_overall = stats.f_oneway(*groups_x)  # ANOVA
            else:
                _, p_overall = stats.kruskal(*groups_x)   # Kruskal-Wallis

            title_var = f"**{v}**" if n_missing == 0 or show_na else f"**{v}** [N={n_valid}]"
            row = {"Variable": title_var}
            for lvl, x in zip(levels, groups_x):
                if show_ci:
                    ci = confinterval(x, method="param" if is_param else "non-param", conf_level=0.95)
                    row[str(lvl)] = f"{format2(ci.center, 1)} [{format2(ci.lower, 1)}-{format2(ci.upper, 1)}]"
                else:
                    row[str(lvl)] = (
                        f"{format2(np.mean(x), 1)} ± {format2(np.std(x, ddof=1), 1)}" if is_param
                        else f"{format2(np.median(x), 0)} [{format2(np.percentile(x, 25), 0)}-{format2(np.percentile(x, 75), 0)}]"
                    )
            row["p.overall"] = format2(p_overall, 3)
            rows.append(row)

            if show_na and n_missing > 0:
                na_row = {"Variable": "Manquant (NA)", "p.overall": ""}
                for lvl in levels:
                    n_na = int(df.loc[df[group] == lvl, v].isna().sum())
                    tot = n_by_level[lvl]
                    na_row[str(lvl)] = f"{n_na} ({format2(100 * n_na / tot, 1)}%)" if tot > 0 else "0 (0.0%)"
                rows.append(na_row)
        else:
            levels_v = _get_variable_levels(s)
            ct = pd.crosstab(df[v], df[group]).reindex(index=levels_v, columns=levels, fill_value=0)
            p_overall = chisq_test2(ct.values.astype(float), chisq_test_perm=False, chisq_test_B=2000)

            valid_col_totals = {lvl: int(ct[lvl].sum()) for lvl in levels}

            title_var = f"**{v}:**" if n_missing == 0 or show_na else f"**{v}:** [N={n_valid}]"
            header_row = {"Variable": title_var, "p.overall": format2(p_overall, 3)}
            for lvl in levels:
                header_row[str(lvl)] = ""
            rows.append(header_row)

            for lvl_v in levels_v:
                row = {"Variable": str(lvl_v), "p.overall": ""}
                for lvl in levels:
                    n = int(ct.loc[lvl_v, lvl])
                    n_col_valid = valid_col_totals[lvl]
                    pct = 100 * n / n_col_valid if n_col_valid > 0 else np.nan
                    row[str(lvl)] = f"{n} ({format2(pct, 1)}%)"
                rows.append(row)

            if show_na and n_missing > 0:
                na_row = {"Variable": "Manquant (NA)", "p.overall": ""}
                for lvl in levels:
                    n_na = int(df.loc[df[group] == lvl, v].isna().sum())
                    tot = n_by_level[lvl]
                    na_row[str(lvl)] = f"{n_na} ({format2(100 * n_na / tot, 1)}%)" if tot > 0 else "0 (0.0%)"
                rows.append(na_row)

    return pd.DataFrame(rows)
