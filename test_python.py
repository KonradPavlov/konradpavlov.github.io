"""
Batterie de tests PavStat (Python) — a executer et comparer
avec test_R.R (meme numerotation de TEST 1 a TEST 12).

Objectif : identifier precisement les ecarts entre PavStat et
les packages R de reference (compareGroups, epitools, gtsummary), pour
prioriser le travail restant vers un package Python complet
(OR, RR, RP, p.ratio, p.overall, IC95%, >2 modalites, export word/html/md).

Remplis la grille_comparaison.md au fur et a mesure des resultats.
"""
import pandas as pd
import Pav as pv

df = pd.read_excel("hdv2003.xlsx")
print(df.shape)

# ============================================================
# TEST 1 — Descriptif univarie
# ============================================================
print("\n=== TEST 1 : describe_table ===")
t1 = pv.describe_table(df, vars=["age", "sexe", "sport", "cinema"])
print(t1.to_string(index=False))

# ============================================================
# TEST 2 — Comparatif, variable QUANTITATIVE, groupe binaire
#   -> verifie la decision normal/non-normal (Shapiro-Wilk) et p.overall
# ============================================================
df["sexe"] = pd.Categorical(df["sexe"], categories=["Homme", "Femme"], ordered=True)

print("\n=== TEST 2 : compare_table, age ~ sexe ===")
t2 = pv.compare_table(df, group="sexe", vars=["age"], ratio="RR")
print(t2.to_string(index=False))

# ============================================================
# TEST 3 — Comparatif, categorielle 2 modalites, ratio=RR
# ============================================================
print("\n=== TEST 3 : compare_table, sport ~ sexe (RR) ===")
t3 = pv.compare_table(df, group="sexe", vars=["sport"], ratio="RR")
print(t3.to_string(index=False))

# ============================================================
# TEST 4 — Comparatif, categorielle 2 modalites, ratio=RR
# ============================================================
print("\n=== TEST 4 : compare_table, cinema ~ sexe (RR) ===")
t4 = pv.compare_table(df, group="sexe", vars=["cinema"], ratio="RR")
print(t4.to_string(index=False))

# ============================================================
# TEST 5 — Comparatif, categorielle >2 modalites, ref_level par defaut
#   -> verifie le choix du niveau de reference et le N par variable (NA)
# ============================================================
print("\n=== TEST 5 : compare_table, nivetud ~ sexe (RR, ref auto) ===")
t5 = pv.compare_table(df, group="sexe", vars=["nivetud"], ratio="RR")
print(t5.to_string(index=False))

# ============================================================
# TEST 6 — Comparatif, categorielle >2 modalites, ratio=OR
# ============================================================
print("\n=== TEST 6 : compare_table, trav.satisf ~ sexe (OR) ===")
t6 = pv.compare_table(df, group="sexe", vars=["trav.satisf"], ratio="OR")
print(t6.to_string(index=False))

# ============================================================
# TEST 7 — Groupe a >2 modalites, variable QUANTITATIVE
#   -> teste la branche _compare_table_multigroup (ANOVA/Kruskal-Wallis)
# ============================================================
print("\n=== TEST 7 : compare_table, age ~ qualif (>2 groupes) ===")
t7 = pv.compare_table(df, group="qualif", vars=["age"])
print(t7.to_string(index=False))

# ============================================================
# TEST 8 — Groupe a >2 modalites, variable CATEGORIELLE
#   -> teste la branche multigroup, chi2/Fisher global
# ============================================================
print("\n=== TEST 8 : compare_table, sport ~ qualif (>2 groupes) ===")
t8 = pv.compare_table(df, group="qualif", vars=["sport"])
print(t8.to_string(index=False))

# ============================================================
# TEST 9 — Export Markdown
# ============================================================
print("\n=== TEST 9 : to_markdown ===")
md = pv.to_markdown(t3, title="Comparaison par sexe")
print(md)

# ============================================================
# TEST 10 — Export Word
# ============================================================
print("\n=== TEST 10 : to_docx ===")
pv.to_docx(t3, "tableau_sexe.docx", title="Comparaison par sexe (hdv2003)",
           footnote="RR = Risque Relatif (Wald), IC95%")
print("Fichier tableau_sexe.docx genere.")

# ============================================================
# TEST 11 — Export HTML
#   -> GAP CONNU : PavStat n'a PAS de to_html() pour l'instant
# ============================================================
print("\n=== TEST 11 : to_html ===")
try:
    html = pv.to_html(t3, title="Comparaison par sexe")
    print("to_html existe :", len(html), "caracteres generes")
except AttributeError:
    print("GAP : pv.to_html n'existe pas encore dans PavStat")

# ============================================================
# TEST 12 — IC95% pour variable quantitative
#   -> GAP CONNU : confinterval() existe dans utils.py mais n'est pas
#      branche dans describe_table()/compare_table() (pas de colonne IC)
# ============================================================
print("\n=== TEST 12 : IC95% sur la moyenne (age) ===")
from Pav.utils import confinterval
ci = confinterval(df["age"].dropna().values, method="param", conf_level=0.95)
print(f"Moyenne={ci.center:.2f}  IC95%=[{ci.lower:.2f};{ci.upper:.2f}]")
print("GAP : ce resultat n'apparait dans aucune colonne de describe_table()/compare_table()")
