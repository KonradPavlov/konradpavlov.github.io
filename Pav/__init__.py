"""
Pav — Package Python pour l'analyse épidémiologique et biostatistique,
équivalent à compareGroups, gtsummary et epitools en R.

Auteur : SALLAN Konrad Pavlov (Épidémiologiste & Data Analyst R / Python)
Licence : MIT
"""
from .tables import describe_table, compare_table
from .utils import format2, confinterval, chisq_test2
from .export import to_markdown, save_markdown, to_docx, to_html

__author__ = "SALLAN Konrad Pavlov"
__author_title__ = "Épidémiologiste & Data Analyst (R / Python)"
__version__ = "0.3.0"
__all__ = [
    "describe_table",
    "compare_table",
    "format2",
    "confinterval",
    "chisq_test2",
    "to_markdown",
    "save_markdown",
    "to_docx",
    "to_html",
]
