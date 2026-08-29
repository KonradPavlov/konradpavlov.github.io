# ============================================================
# Script de validation R — PavStat vs packages R de référence
# ============================================================
# Objectif : reproduire, avec de vrais packages R validés, les mêmes
# tableaux que ceux produits par ton package Python PavStat
# sur hdv2003, pour comparer les résultats chiffre à chiffre.
#
# describe_table()  -> descriptif univarié     (equivalent: compareGroups)
# compare_table()   -> comparatif par groupe,
#                      avec RR/OR (Wald) et p.ratio (mid-p exact)
#                                              (equivalent: epitools)
# p.overall (chi2/Fisher)                     (equivalent: compareGroups)
#
# NB IMPORTANT : compareGroups calcule en interne un Odds Ratio (OR) via
# régression logistique quand on demande show.ratio=TRUE, ce qui N'EST
# PAS le même calcul que le RR de Wald traduit depuis epitools_functions.R
# dans ton package Python. Pour comparer les VRAIS chiffres RR/p.ratio de
# ta sortie Python, utilise la section 4 (epitools), pas compareGroups.
# ============================================================

# ------------------------------------------------------------
# 0. Installation des packages (à lancer une seule fois)
# ------------------------------------------------------------
# install.packages("questionr")     # contient le jeu de données hdv2003
# install.packages("compareGroups") # tableaux descriptifs/comparatifs + p.overall
# install.packages("epitools")      # RR/OR (Wald) + p.ratio (mid-p exact)
# install.packages("knitr")         # pour export2md()

library(questionr)
library(compareGroups)
library(epitools)

# ------------------------------------------------------------
# 1. Chargement des données (identique au hdv2003.xlsx utilisé en Python)
# ------------------------------------------------------------
data(hdv2003)
df <- hdv2003

str(df)
dim(df)   # doit donner 2000 20, comme en Python

# ------------------------------------------------------------
# 2. Tableau descriptif univarié (équivalent describe_table)
#    -> age, sexe, sport, cinema
# ------------------------------------------------------------
desc <- compareGroups(~ age + sexe + sport + cinema, data = df)
tbl_desc <- createTable(desc)
tbl_desc
# export2md(tbl_desc)   # nécessite le package knitr

# ------------------------------------------------------------
# 3. p.overall par sexe (équivalent chi2/Fisher de compareGroups)
#    -> age, sport, cinema, nivetud, trav.satisf
# ------------------------------------------------------------
comp_sexe <- compareGroups(
  sexe ~ age + sport + cinema + nivetud + trav.satisf,
  data = df
)
tbl_comp_sexe <- createTable(comp_sexe)
tbl_comp_sexe
# export2md(tbl_comp_sexe)

# ------------------------------------------------------------
# 4. RR / OR (Wald) + p.ratio (mid-p exact) par sexe
#    -> équivalent exact de _riskratio_wald / _oddsratio_wald /
#       _ratio_pvalue de PavStat (tables.py)
# ------------------------------------------------------------
# Fonction utilitaire : reproduit le calcul RR/OR + p.ratio pour une
# variable catégorielle "var" en fonction du groupe binaire "sexe",
# avec un niveau de référence "ref".
rr_or_table <- function(df, var, group = "sexe", ref = NULL, ratio = "RR") {
  ct <- table(df[[var]], df[[group]])
  lvl0 <- colnames(ct)[1]  # ex. "Homme"
  lvl1 <- colnames(ct)[2]  # ex. "Femme" (evenement)

  levels_v <- rownames(ct)
  if (is.null(ref)) ref <- levels_v[1]

  b0_ref <- ct[ref, lvl0]  # non-evenement au niveau reference
  a0_ref <- ct[ref, lvl1]  # evenement au niveau reference

  results <- data.frame(
    niveau = levels_v,
    n_lvl0 = as.integer(ct[, lvl0]),
    n_lvl1 = as.integer(ct[, lvl1]),
    ratio = NA_character_,
    p.ratio = NA_character_,
    stringsAsFactors = FALSE
  )

  for (i in seq_along(levels_v)) {
    lvl <- levels_v[i]
    if (lvl == ref) {
      results$ratio[i] <- "Ref."
      results$p.ratio[i] <- "Ref."
      next
    }
    a1 <- ct[lvl, lvl1]
    b1 <- ct[lvl, lvl0]

    x2 <- matrix(c(a0_ref, b0_ref, a1, b1), nrow = 2, byrow = TRUE,
                 dimnames = list(c(ref, lvl), c("evenement", "non_evenement")))

    if (ratio == "RR") {
      rr <- riskratio(x2, method = "wald")
      est <- rr$measure[2, 1]; lo <- rr$measure[2, 2]; hi <- rr$measure[2, 3]
    } else {
      orr <- oddsratio(x2, method = "wald")
      est <- orr$measure[2, 1]; lo <- orr$measure[2, 2]; hi <- orr$measure[2, 3]
    }

    # p.ratio : test mid-p exact (equivalent ormidp.test / tab2by2.test)
    pmid <- ormidp.test(a0_ref, b0_ref, a1, b1)$p.value

    results$ratio[i] <- sprintf("%.2f [%.2f;%.2f]", est, lo, hi)
    results$p.ratio[i] <- sprintf("%.3f", pmid)
  }

  names(results)[2:3] <- c(lvl0, lvl1)
  results
}

cat("\n--- sport ~ sexe (RR) ---\n")
print(rr_or_table(df, "sport", ref = "Non", ratio = "RR"))

cat("\n--- cinema ~ sexe (RR) ---\n")
print(rr_or_table(df, "cinema", ref = "Non", ratio = "RR"))

cat("\n--- nivetud ~ sexe (RR) ---\n")
print(rr_or_table(df, "nivetud", ratio = "RR"))

cat("\n--- trav.satisf ~ sexe (RR) ---\n")
print(rr_or_table(df, "trav.satisf", ref = "Equilibre", ratio = "RR"))

# ------------------------------------------------------------
# 5. Export Word pour comparaison visuelle avec to_docx() (Python)
# ------------------------------------------------------------
# export2word(tbl_comp_sexe, file = "tableau_comparatif_sexe.docx")

# ------------------------------------------------------------
# 6. Points de comparaison chiffre à chiffre avec ta sortie Python
# ------------------------------------------------------------
# Python (rappel) :
#   sport:  Ref.=Non, Oui -> RR=0.84 [0.77;0.91], p.ratio<0.001, p.overall<0.001
#   cinema: Ref.=Non, Oui -> RR=1.05 [0.97;1.14], p.ratio=0.193,  p.overall=0.208
#
# -> Compare ces valeurs à la sortie de rr_or_table() ci-dessus (section 4)
#    et au p.overall de tbl_comp_sexe (section 3, colonne p-value).
