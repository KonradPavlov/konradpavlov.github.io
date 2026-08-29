import Pav
import pandas as pd

df = pd.read_excel("hdv2003.xlsx")
print(df.head())
print(df.shape)

tbl_descr = Pav.describe_table(df, vars=["age", "sexe", "sport", "cinema"])
print(tbl_descr.to_string(index=False))

df["sexe"] = pd.Categorical(df["sexe"], categories=["Homme", "Femme"], ordered=True)

tbl_compare = Pav.compare_table(
    df,
    group="sexe",
    vars=["age", "sport", "cinema"],
    ratio="RR",
    ref_level=None
)
print(tbl_compare.to_string(index=False))

tbl_compare2 = Pav.compare_table(
    df,
    group="sexe",
    vars=["nivetud", "trav.satisf"],
    ratio="RR"
)
print(tbl_compare2.to_string(index=False))

tbl_md = Pav.to_markdown(tbl_compare, title="Comparaison par sexe")
print(tbl_md)

Pav.to_docx(tbl_compare, "tableau_sexe.docx", title="Comparaison par sexe (hdv2003)", footnote="RR = Risque Relatif (Wald), IC95%")