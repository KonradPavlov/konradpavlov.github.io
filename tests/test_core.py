"""
Tests unitaires pour Pav.
"""
import unittest
from pathlib import Path
import pandas as pd
import numpy as np
import Pav as pv

class TestPav(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        np.random.seed(42)
        n = 200
        cls.df = pd.DataFrame({
            "age": np.random.normal(50, 15, n),
            "sexe": np.random.choice(["Homme", "Femme"], n, p=[0.45, 0.55]),
            "maladie": np.random.choice(["Non", "Oui"], n, p=[0.7, 0.3]),
            "stade": pd.Categorical(
                np.random.choice(["Stade I", "Stade II", "Stade III"], n),
                categories=["Stade I", "Stade II", "Stade III"],
                ordered=True,
            ),
            "score_skewed": np.random.exponential(2, n),
        })

    def test_describe_table_split(self):
        t = pv.describe_table(self.df, vars=["age", "sexe", "maladie"], split_columns=True)
        self.assertIsInstance(t, pd.DataFrame)
        self.assertIn("Caractéristique", t.columns)
        self.assertIn("Effectif (n)", t.columns)
        self.assertIn("Fréquence (%)", t.columns)

    def test_describe_table_single(self):
        t = pv.describe_table(self.df, vars=["age", "sexe"], split_columns=False)
        self.assertIsInstance(t, pd.DataFrame)
        self.assertIn("Caractéristique", t.columns)
        self.assertIn("Valeur", t.columns)

    def test_describe_table_ci(self):
        t = pv.describe_table(self.df, vars=["age", "score_skewed"], show_ci=True)
        self.assertIsInstance(t, pd.DataFrame)
        has_ci = any("IC95%" in str(val) for val in t["Caractéristique"])
        self.assertTrue(has_ci)

    def test_compare_table_rr(self):
        t = pv.compare_table(self.df, group="sexe", vars=["maladie"], ratio="RR")
        self.assertIsInstance(t, pd.DataFrame)
        self.assertIn("RR", t.columns)
        self.assertIn("p.ratio", t.columns)
        self.assertIn("p.overall", t.columns)

    def test_compare_table_or(self):
        t = pv.compare_table(self.df, group="sexe", vars=["maladie", "age"], ratio="OR")
        self.assertIsInstance(t, pd.DataFrame)
        self.assertIn("OR", t.columns)
        self.assertIn("p.ratio", t.columns)
        self.assertIn("p.overall", t.columns)

    def test_compare_table_multigroup(self):
        t = pv.compare_table(self.df, group="stade", vars=["age", "maladie"])
        self.assertIsInstance(t, pd.DataFrame)
        self.assertIn("Stade I", t.columns)
        self.assertIn("Stade II", t.columns)
        self.assertIn("Stade III", t.columns)
        self.assertIn("p.overall", t.columns)

    def test_exports(self):
        t = pv.describe_table(self.df, vars=["sexe", "maladie"])
        
        # Markdown
        md = pv.to_markdown(t, title="Tableau 1")
        self.assertIn("Effectif (n)", md)
        self.assertIn("Fréquence (%)", md)
        
        # HTML
        html_file = Path("test_desc.html")
        pv.to_html(t, path=html_file, title="Tableau 1 HTML")
        self.assertTrue(html_file.exists())
        html_file.unlink()

        # Word
        docx_file = Path("test_desc.docx")
        pv.to_docx(t, path=docx_file, title="Tableau 1 Word")
        self.assertTrue(docx_file.exists())
        docx_file.unlink()

if __name__ == "__main__":
    unittest.main()
