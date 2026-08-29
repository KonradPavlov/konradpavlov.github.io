# ============================================================
# Batterie de tests R — a comparer avec test_python.py
# (meme numerotation TEST 1 a TEST 12)
# ============================================================
# install.packages(c("questionr","compareGroups","epitools","gtsummary"))

library(questionr)
library(compareGroups)
library(epitools)

data(hdv2003)
df <- hdv2003
dim(df)

# ============================================================
# TEST 1 — Descriptif univarie
# ============================================================
cat("\n=== TEST 1 : describe_table ===\n")
t1 <- createTable(compareGroups(~ age + sexe + sport + cinema, data = df))
t1

# ============================================================
# TEST 2 — Comparatif, variable QUANTITATIVE, groupe binaire
# ============================================================
df$sexe <- factor(df$sexe, levels = c("Homme", "Femme"))

cat("\n=== TEST 2 : compare_table, age ~ sexe ===\n")
t2 <- createTable(compareGroups(sexe ~ age, data = df), show.ratio = TRUE)
t2
# Shapiro-Wilk sur age, pour verifier la decision normal/non-normal
shapiro.test(df$age)

# ============================================================
# TEST 3 — Comparatif, categorielle 2 modalites (OR compareGroups)
#   + RR/p.ratio via epitools (equivalent exact du ratio="RR" Python)
# ============================================================
cat("\n=== TEST 3 : compare_table, sport ~ sexe ===\n")
t3 <- createTable(compareGroups(sexe ~ sport, data = df), show.ratio = TRUE)
t3

ct_sport <- table(df$sport, df$sexe)
riskratio(ct_sport, method = "wald")

# ============================================================
# TEST 4 — Comparatif, categorielle 2 modalites
# ============================================================
cat("\n=== TEST 4 : compare_table, cinema ~ sexe ===\n")
t4 <- createTable(compareGroups(sexe ~ cinema, data = df), show.ratio = TRUE)
t4

ct_cinema <- table(df$cinema, df$sexe)
riskratio(ct_cinema, method = "wald")

# ============================================================
# TEST 5 — Comparatif, categorielle >2 modalites
#   -> verifie le niveau de reference choisi par R (1er niveau du facteur,
#      PAS l'ordre alphabetique) et le N par variable (NA exclus)
# ============================================================
cat("\n=== TEST 5 : compare_table, nivetud ~ sexe ===\n")
t5 <- createTable(compareGroups(sexe ~ nivetud, data = df), show.ratio = TRUE)
t5

# ============================================================
# TEST 6 — Comparatif, categorielle >2 modalites (OR)
# ============================================================
cat("\n=== TEST 6 : compare_table, trav.satisf ~ sexe ===\n")
t6 <- createTable(compareGroups(sexe ~ trav.satisf, data = df), show.ratio = TRUE)
t6

# ============================================================
# TEST 7 — Groupe a >2 modalites, variable QUANTITATIVE
#   -> ANOVA/Kruskal-Wallis automatique dans compareGroups
# ============================================================
t7 <- createTable(compareGroups(qualif ~ age, data = df, max.ylev = 10))
t7

# ============================================================
# TEST 8 — Groupe a >2 modalites, variable CATEGORIELLE
# ============================================================
cat("\n=== TEST 8 : compare_table, sport ~ qualif ===\n")
t8 <- createTable(compareGroups(qualif ~ sport, data = df,max.ylev = 10))
t8

# ============================================================
# TEST 9 — Export Markdown
# ============================================================
cat("\n=== TEST 9 : export2md ===\n")
export2md(t3, caption = "Comparaison par sexe")

# ============================================================
# TEST 10 — Export Word
# ============================================================
cat("\n=== TEST 10 : export2word ===\n")
export2word(t3, file = "tableau_sexe.docx", caption = "Comparaison par sexe (hdv2003)")

# ============================================================
# TEST 11 — Export HTML (deja disponible cote R, gap connu cote Python)
# ============================================================
cat("\n=== TEST 11 : export2html ===\n")
export2html(t3, file = "tableau_sexe.html", caption = "Comparaison par sexe")

# ============================================================
# TEST 12 — IC95% sur la moyenne (age), deja affichable via gtsummary
#   -> compareGroups seul ne l'affiche pas non plus par defaut ;
#      gtsummary::tbl_summary(statistic = list(age ~ "{mean} ({conf.low}-{conf.high})"))
#      est la reference si tu veux le comparer.
# ============================================================
cat("\n=== TEST 12 : IC95% moyenne age ===\n")
t.test(df$age)$conf.int
