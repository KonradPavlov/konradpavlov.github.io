"""
Pav.utils — Fonctions statistiques et utilitaires
Auteur : SALLAN Konrad Pavlov

Traduction fidèle des fonctions utilitaires du package R `compareGroups`.

Fichiers R d'origine : format2.R, confinterval.R, chisq.test2.R,
signifdec.R, trim.R
"""
from __future__ import annotations

import re
import warnings
from dataclasses import dataclass
from typing import Optional, Union

import numpy as np
from scipy import stats


# ---------------------------------------------------------------------------
# trim.R
# ---------------------------------------------------------------------------
def trim(x: str) -> str:
    """R: trim <- function(x){ x <- gsub("^[ ]+","",x); x <- gsub("[ ]+$","",x); x }"""
    x = re.sub(r"^[ ]+", "", x)
    x = re.sub(r"[ ]+$", "", x)
    return x


# ---------------------------------------------------------------------------
# signifdec.R (signifdec.i.R non fourni dans l'extrait — implémentation
# standard : arrondi à `digits` décimales significatives, comme round())
# ---------------------------------------------------------------------------
def signifdec_i(x: float, digits: int) -> float:
    if x is None or (isinstance(x, float) and (np.isnan(x))):
        return np.nan
    return round(x, digits)


def signifdec(x, digits: int):
    """R: signifdec <- function(x,digits){ sapply(x,signifdec.i,digits=digits) }"""
    arr = np.atleast_1d(x)
    out = np.array([signifdec_i(v, digits) for v in arr])
    return out if len(out) > 1 else out[0]


# ---------------------------------------------------------------------------
# format2.R
# ---------------------------------------------------------------------------
def _fmt_fixed(x: float, digits: int) -> str:
    """Equivalent de format(round(x,digits), trim=TRUE, nsmall=digits)."""
    return f"{round(x, digits):.{digits}f}"


def format2(
    x: Union[float, np.ndarray],
    digits: Optional[int] = None,
    stars: bool = False,
) -> Union[str, np.ndarray]:
    """
    R:
    format2 <- function(x,digits=NULL,stars=FALSE,...){
      if (!is.null(digits)){
        res<-format(round(x,digits),trim=TRUE,nsmall=digits,...)
        res<-ifelse(res=="0.000..0",paste("<0.000..1"),res)   # cas 0 arrondi
        res<-ifelse(x==0, format(round(x,digits),...), res)    # vrai zéro
        res<-ifelse(is.na(x)|is.nan(x)|x==Inf|x==-Inf,".",res)
      } else {
        # digits auto : 2 déc si x<10, 1 déc si x<100, 0 déc sinon
      }
      if (stars) ajoute des étoiles de significativité (*, **, ***)
      return(res)
    }
    """
    scalar_input = np.isscalar(x) or isinstance(x, (int, float))
    arr = np.atleast_1d(np.asarray(x, dtype=float))
    res = np.empty(arr.shape, dtype=object)

    for i, v in enumerate(arr):
        if np.isnan(v) or np.isinf(v):
            res[i] = "."
            continue

        if digits is not None:
            rounded_str = _fmt_fixed(v, digits)
            # cas où l'arrondi donne 0.00..0 alors que v != 0 -> "<0.00..1"
            zero_str = f"{0:.{digits}f}"
            if rounded_str == zero_str and v != 0:
                if digits >= 1:
                    threshold = "0." + "0" * (digits - 1) + "1"
                else:
                    threshold = "0"
                rounded_str = f"<{threshold}"
            if v == 0:
                rounded_str = _fmt_fixed(v, digits)
            res[i] = rounded_str
        else:
            if v < 10:
                res[i] = _fmt_fixed(v, 2)
            elif v < 100:
                res[i] = _fmt_fixed(v, 1)
            else:
                res[i] = _fmt_fixed(v, 0)

        if stars:
            if v < 0.01:
                res[i] = f"{res[i]}*** "
            elif v < 0.05:
                res[i] = f"{res[i]}** "
            elif v < 0.1:
                res[i] = f"{res[i]}*  "
            else:
                res[i] = f"{res[i]}   "

    return res[0] if scalar_input else res


# ---------------------------------------------------------------------------
# confinterval.R
# ---------------------------------------------------------------------------
@dataclass
class ConfInterval:
    center: float   # 'Mean' ou 'Median' selon la méthode
    lower: float
    upper: float


def confinterval(x: np.ndarray, method: str, conf_level: float) -> ConfInterval:
    """
    R:
    confinterval <- function(x, method, conf.level){
      alpha <- 1-conf.level
      n <- length(x)
      if (method=="param"){
        m <- mean(x); se <- sd(x)/sqrt(n)
        low <- m + qt(alpha/2, n-1)*se
        upp <- m - qt(alpha/2, n-1)*se
        return(c('Mean'=m,'lower'=low,'upper'=upp))
      } else {
        # IC non paramétrique sur la médiane (méthode binomiale exacte)
        L <- qbinom(alpha/2, n, 0.5); U <- n-L+1
        if (L>=U) { warning("cannot compute CI"); return(NA,NA,NA) }
        order.x <- sort(x)
        c('Median'=median(x),'lower'=order.x[L],'upper'=order.x[n-L+1])
      }
    }
    """
    x = np.asarray(x, dtype=float)
    x = x[~np.isnan(x)]
    alpha = 1 - conf_level
    n = len(x)

    if method == "param":
        m = float(np.mean(x))
        se = float(np.std(x, ddof=1)) / np.sqrt(n)
        low = m + stats.t.ppf(alpha / 2, n - 1) * se
        upp = m - stats.t.ppf(alpha / 2, n - 1) * se
        return ConfInterval(center=m, lower=low, upper=upp)
    else:
        # qbinom(alpha/2, n, 0.5) -> quantile de Binomial(n, 0.5)
        L = int(stats.binom.ppf(alpha / 2, n, 0.5))
        U = n - L + 1
        if L == 0 or L >= U:
            warnings.warn("cannot compute CI")
            return ConfInterval(center=np.nan, lower=np.nan, upper=np.nan)
        order_x = np.sort(x)
        # indices R sont 1-based -> décaler de -1 en Python
        return ConfInterval(
            center=float(np.median(x)),
            lower=float(order_x[L - 1]),
            upper=float(order_x[n - L]),
        )


# ---------------------------------------------------------------------------
# chisq.test2.R
# ---------------------------------------------------------------------------
def chisq_test2(
    obj: np.ndarray,
    chisq_test_perm: bool = False,
    chisq_test_B: int = 2000,
    chisq_test_seed: Optional[int] = None,
) -> float:
    """
    R:
    chisq.test2 <- function(obj, chisq.test.perm, chisq.test.B, chisq.test.seed){
      if (any(dim(obj)<2) || is.null(dim(obj)) || sum(rowSums(obj)>0)<2 || sum(colSums(obj)>0)<2)
        return(NaN)
      obj <- obj[,colSums(obj)>0]                 # retire colonnes nulles
      expect <- outer(rowSums(obj),colSums(obj))/sum(obj)
      if (any(expect<5)){
        if (chisq.test.perm) test <- chisq.test(obj, simulate.p.value=TRUE, B=chisq.test.B)
        else test <- fisher.test(obj)
      } else test <- chisq.test(obj)
      if (inherits(test,"try-error")) return(NaN)
      return(test$p.value)
    }
    """
    obj = np.asarray(obj, dtype=float)
    if obj.ndim != 2:
        return np.nan
    if any(d < 2 for d in obj.shape):
        return np.nan
    if (obj.sum(axis=1) > 0).sum() < 2 or (obj.sum(axis=0) > 0).sum() < 2:
        return np.nan

    obj = obj[:, obj.sum(axis=0) > 0]  # retire les colonnes nulles
    row_sums = obj.sum(axis=1, keepdims=True)
    col_sums = obj.sum(axis=0, keepdims=True)
    expect = (row_sums @ col_sums) / obj.sum()

    try:
        if np.any(expect < 5):
            if chisq_test_perm:
                # Test du Chi2 par permutation (Monte-Carlo), équivalent à
                # simulate.p.value=TRUE, B=... de R
                if chisq_test_seed is not None:
                    rng = np.random.default_rng(chisq_test_seed)
                else:
                    rng = np.random.default_rng()
                res = stats.chi2_contingency(obj)
                stat_obs = res.statistic
                row_p = row_sums.flatten() / obj.sum()
                col_p = col_sums.flatten() / obj.sum()
                count = 0
                n_total = int(obj.sum())
                for _ in range(chisq_test_B):
                    sim = rng.multinomial(n_total, (row_p[:, None] * col_p[None, :]).flatten())
                    sim = sim.reshape(obj.shape)
                    try:
                        sim_stat = stats.chi2_contingency(sim, correction=False).statistic
                    except ValueError:
                        sim_stat = 0
                    if sim_stat >= stat_obs:
                        count += 1
                p_value = count / chisq_test_B
            else:
                if obj.shape == (2, 2):
                    _, p_value = stats.fisher_exact(obj)
                else:
                    # Fisher exact généralisé (table r x c) — nécessite un
                    # calcul combinatoire ; scipy ne le supporte qu'en 2x2.
                    # À raffiner plus tard (ex. via rpy2 ou R). Pour l'instant,
                    # on retombe sur un chi2 classique comme approximation.
                    warnings.warn(
                        "Fisher exact sur table r x c (r,c>2) non implémenté "
                        "en Python natif — approximation chi2 utilisée."
                    )
                    correction = obj.shape == (2, 2)
                    _, p_value, _, _ = stats.chi2_contingency(obj, correction=correction)
        else:
            # R: chisq.test(obj) applique par défaut Yates' continuity
            # correction pour les tables 2x2 uniquement (validé contre
            # p.overall=0.157 sur sexe~gravite du tableau réel de Konrad)
            correction = obj.shape == (2, 2)
            _, p_value, _, _ = stats.chi2_contingency(obj, correction=correction)
    except Exception:
        return np.nan

    return float(p_value)
