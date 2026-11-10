"""Tables numbered 4.4.1 and 4.4.2 are two tables. The list check must not read both as Table 4.4."""
import os
import sys
from types import SimpleNamespace as NS

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "app"))
import check  # noqa: E402


def line(text):
    return NS(text=text, bbox_pt=lambda: (0, 0, 1, 1))


def page(no, label, kind, lines, section=None):
    return NS(no=no, kind=kind, section=section, lines=[line(t) for t in lines], label_line=line(label))


lst = page(11, "(8)", "front", ["LIST OF TABLES", "Tables Page",
                                 "Table 4.4.1 Relative change in average hemoglobin level ........ 32",
                                 "Table 4.4.2 Paired t-test for Average Hemoglobin Level ........ 34"],
           {"name": "LIST OF TABLES"})
pages = [lst,
         page(45, "32", "chapter", ["Table 4.4.1 Relative change in average hemoglobin level"]),
         page(47, "34", "chapter", ["Table 4.4.2 Paired t-test for Average Hemoglobin Level"])]
F = check.Findings()
check.check_lists(pages, {}, F)
bad = [f["message"] for f in F.numbered()]
assert not bad, bad
print("OK")
