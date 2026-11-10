"""Manuscript format checker.

Checks a thesis PDF (or a Word file, converted to PDF first) against the
measured layout of a school template and writes two files next to it:
  <name> - format report.html   the list of problems, page by page
  <name> - corrections.pdf      printable sheets: every page with a problem, boxed and numbered

Usage:
  python check.py "My Thesis.pdf"
  python check.py "My Thesis.docx" --rules rules/tu-cicm-thesis-2022.json
"""
import argparse
import html
import json
import os
import re
import shutil
import statistics
import subprocess
import sys
import tempfile
import threading
import time
from collections import Counter, defaultdict

import logging

import pdfplumber

logging.getLogger("pypdf").setLevel(logging.ERROR)
logging.getLogger("pdfminer").setLevel(logging.ERROR)

PT_TO_CM = 2.54 / 72
HERE = os.path.dirname(os.path.abspath(__file__))
BASE = HERE   # rules/ and ui/ sit next to the code, both in the installed program and in a downloaded update
DEFAULT_RULES = os.path.join(BASE, "rules", "tu-cicm-thesis-2022.json")
USER_DIR = os.path.join(os.environ.get("APPDATA") or os.path.expanduser("~"), "Manuscript Check")
USER_RULES = os.path.join(USER_DIR, "rules")      # formats learned on this computer


class ConversionError(Exception):
    pass


def rules_files():
    """Every format available: the bundled ones, then the ones learned on this computer."""
    out = []
    for folder in (os.path.join(BASE, "rules"), USER_RULES):
        if os.path.isdir(folder):
            out += [os.path.join(folder, f) for f in sorted(os.listdir(folder)) if f.endswith(".json")]
    return out


def cm(v):
    return v * PT_TO_CM


def r2(v):
    return f"{v:.2f}"


# ---------------------------------------------------------------- reading the PDF

class Line:
    __slots__ = ("page", "text", "x0", "x1", "top", "bottom", "size", "bold", "fonts", "in_table")

    def __init__(self, page, words):
        main = [w for w in words if w["size"] > 9.4] or words
        self.page = page
        self.text = " ".join(w["text"] for w in sorted(words, key=lambda w: w["x0"]))
        self.x0 = cm(min(w["x0"] for w in words))
        self.x1 = cm(max(w["x1"] for w in words))
        self.top = cm(min(w["top"] for w in main))
        self.bottom = cm(max(w["bottom"] for w in main))
        sizes = Counter()
        bold = 0
        total = 0
        for w in main:
            n = len(w["text"])
            sizes[round(w["size"], 1)] += n
            total += n
            if "bold" in w["fontname"].lower():
                bold += n
        self.size = sizes.most_common(1)[0][0]
        self.bold = bold > total / 2
        self.fonts = Counter(w["fontname"].split("+")[-1] for w in words)
        self.in_table = False

    @property
    def center(self):
        return (self.x0 + self.x1) / 2

    def bbox_pt(self):
        return (self.x0 / PT_TO_CM, self.top / PT_TO_CM, self.x1 / PT_TO_CM, self.bottom / PT_TO_CM)


def read_lines(pg, page_no):
    words = pg.extract_words(extra_attrs=["fontname", "size"], x_tolerance=1.5, keep_blank_chars=False)
    words = [w for w in words if w["text"].strip()]
    big = sorted([w for w in words if w["size"] > 9.4], key=lambda w: (w["top"], w["x0"]))
    small = [w for w in words if w["size"] <= 9.4]
    groups = []
    for w in big:
        for g in groups:
            if abs(g[0]["top"] - w["top"]) < 2.5:
                g.append(w)
                break
        else:
            groups.append([w])
    # small text: superscripts join the line they sit on, chart labels make lines of their own
    orphans = []
    for w in small:
        mid = (w["top"] + w["bottom"]) / 2
        home = None
        for g in groups:
            t = min(x["top"] for x in g)
            b = max(x["bottom"] for x in g)
            if t - 3 <= mid <= b and min(abs(w["x0"] - x["x1"]) for x in g) < 40:
                home = g
                break
        (home.append(w) if home else orphans.append(w))
    small_groups = []
    for w in sorted(orphans, key=lambda w: (w["top"], w["x0"])):
        for g in small_groups:
            if abs(g[0]["top"] - w["top"]) < 2:
                g.append(w)
                break
        else:
            small_groups.append([w])
    lines = [Line(page_no, g) for g in groups + small_groups]
    lines.sort(key=lambda l: (round(l.top, 1), l.x0))
    try:
        tables = [t.bbox for t in pg.find_tables()]
    except Exception:
        tables = []
    for l in lines:
        x0, t, x1, b = l.bbox_pt()
        for tb in tables:
            if x0 >= tb[0] - 2 and x1 <= tb[2] + 2 and t >= tb[1] - 2 and b <= tb[3] + 2:
                l.in_table = True
    return lines


ROMAN = {"i": 1, "v": 5, "x": 10, "l": 50, "c": 100}


def roman_value(s):
    s = s.lower()
    total = 0
    for i, ch in enumerate(s):
        v = ROMAN[ch]
        if i + 1 < len(s) and ROMAN[s[i + 1]] > v:
            total -= v
        else:
            total += v
    return total


def parse_label(text):
    t = text.replace(" ", "")
    if re.fullmatch(r"\(\d+\)", t):
        return "paren", int(t[1:-1])
    if re.fullmatch(r"\d+", t):
        return "arabic", int(t)
    if re.fullmatch(r"[ivxlc]+", t, re.I):
        return "roman", roman_value(t)
    return None


STYLE_NAME = {"paren": "(1), (2), (3)", "arabic": "1, 2, 3", "roman": "i, ii, iii"}


class Page:
    def __init__(self, no, pg, rules):
        self.no = no
        self.width = cm(pg.width)
        self.height = cm(pg.height)
        self.chars = [c for c in pg.chars if c["text"].strip()]
        self.images = [(cm(i["x0"]), cm(i["top"]), cm(i["x1"]), cm(i["bottom"])) for i in pg.images]
        all_lines = read_lines(pg, no)
        pn = rules["page_number"]
        if pn.get("position", "top") == "bottom":
            edge = self.height - pn.get("zone_from_bottom_cm", 2.5)
            self.header = [l for l in all_lines if l.top >= edge and l.top > self.height * 0.8]
        else:
            zone = pn["zone_bottom_cm"]
            self.header = [l for l in all_lines if l.bottom <= zone + 0.6 and l.top < zone]
        self.lines = [l for l in all_lines if l not in self.header]
        self.label_line = None
        self.label = None
        for l in self.header:
            p = parse_label(l.text)
            if p:
                self.label_line, self.label = l, p
        self.text = "\n".join(l.text for l in all_lines)
        self.ink = None  # left and right edge of anything drawn below the header, in cm
        self.kind = None          # cover, approval, front, chapter, back
        self.section = None       # heading this page belongs to
        self.heading = None       # heading that starts on this page

    @property
    def shown(self):
        return self.label_line.text.strip() if self.label_line else "no number"


# ---------------------------------------------------------------- findings

class Findings:
    def __init__(self):
        self.items = []

    def add(self, level, area, message, page=None, rule_source=None, boxes=None, pages=None):
        """boxes: [(pdf page, bbox in points)] drawn on the page. pages: extra pages that get the note with no box."""
        self.items.append({"level": level, "area": area, "message": message, "page": page,
                           "source": rule_source, "boxes": boxes or [], "pages": list(pages or [])})

    def numbered(self):
        items = sorted(self.items, key=lambda f: (f["level"] != "fix", f["page"] or 0))
        for i, f in enumerate(items, 1):
            f["n"] = i
        return items


def pg(pages, cap=False):
    """'page 4' or 'pages 3 to 5'"""
    rng = page_ranges(pages)
    word = "page" if len(set(pages)) == 1 else "pages"
    return f"{word.capitalize() if cap else word} {rng}"


def isare(pages):
    return "is" if len(set(pages)) == 1 else "are"


def join_and(items):
    items = list(items)
    return items[0] if len(items) == 1 else ", ".join(items[:-1]) + " and " + items[-1]


def page_ranges(pages):
    """[3,4,5,9] -> '3 to 5, 9'"""
    nums = sorted(set(pages))
    out = []
    start = prev = None
    for n in nums:
        if start is None:
            start = prev = n
        elif n == prev + 1:
            prev = n
        else:
            out.append((start, prev))
            start = prev = n
    if start is not None:
        out.append((start, prev))
    return ", ".join(str(a) if a == b else f"{a} to {b}" for a, b in out)


# ---------------------------------------------------------------- structure

FRONT = ["ABSTRACT", "ACKNOWLEDGEMENTS", "ACKNOWLEDGMENTS", "TABLE OF CONTENTS", "LIST OF TABLES",
         "LIST OF FIGURES", "LIST OF ABBREVIATIONS"]
BACK = ["REFERENCES", "BIBLIOGRAPHY", "APPENDICES", "BIOGRAPHY"]
TOC_NAMES = ("TABLE OF CONTENTS", "CONTENTS")
REF_NAMES = ("REFERENCES", "BIBLIOGRAPHY", "LITERATURE CITED", "WORKS CITED")
BIO_NAMES = ("BIOGRAPHY", "CURRICULUM VITAE", "VITA", "AUTHOR'S PROFILE", "BIOGRAPHICAL SKETCH")
ALIASES = (TOC_NAMES, REF_NAMES, BIO_NAMES, ("ACKNOWLEDGEMENTS", "ACKNOWLEDGMENTS"), ("APPENDIX", "APPENDICES"))
DEFAULT_FRONT = list(FRONT)
DEFAULT_BACK = list(BACK)


def norm(s):
    s = re.sub(r"(?<=\d)\s*\.\s*(?=\d)", ".", s)
    s = re.sub(r"\(\s*(\d+)\s*\)", r"(\1)", s)
    return re.sub(r"\s+", " ", s).strip()


def is_heading_line(l, rules):
    return l.size >= rules["sizes"]["heading_pt"] - 0.6 or (l.bold and l.text.isupper())


def find_heading(page, rules):
    """The heading that starts on this page: FRONT/BACK name, APPENDIX X, or CHAPTER n (+ title)."""
    ls = page.lines
    for i, l in enumerate(ls):
        t = norm(l.text).upper()
        if not is_heading_line(l, rules) or "...." in t:
            continue
        if t in FRONT or t in BACK:
            return {"name": t.replace("ACKNOWLEDGMENTS", "ACKNOWLEDGEMENTS"), "lines": [l], "index": i}
        m = re.fullmatch(r"APPENDIX\s+([A-Z0-9]+)", t)
        if m:
            nxt = ls[i + 1] if i + 1 < len(ls) else None
            return {"name": "APPENDIX", "letter": m.group(1), "lines": [l] + ([nxt] if nxt else []), "index": i}
        m = re.match(r"CHAPTER\s+(\d+)\s*(.*)$", t)
        if m:
            num = int(m.group(1))
            title = m.group(2).strip()
            used = [l]
            if not title and i + 1 < len(ls):
                nxt = ls[i + 1]
                if is_heading_line(nxt, rules) and nxt.top - l.bottom < 1.6:
                    title = norm(nxt.text).upper()
                    used.append(nxt)
                    # a long title can run onto a third line
                    if i + 2 < len(ls):
                        n2 = ls[i + 2]
                        if is_heading_line(n2, rules) and n2.top - nxt.bottom < 0.8 and not re.match(r"^\d", n2.text):
                            title += " " + norm(n2.text).upper()
                            used.append(n2)
            return {"name": "CHAPTER", "number": num, "title": title, "lines": used, "index": i,
                    "same_line": bool(m.group(2).strip())}
    return None


def classify(pages, rules, F):
    """Mark cover, approval page, and which heading each page belongs to."""
    if pages:
        pages[0].kind = "cover"
    for p in pages[1:5]:
        if "was approved" in p.text.lower():
            p.kind = "approval"
    first_appr = next((p.no for p in pages if p.kind == "approval"), None)
    if first_appr:
        for p in pages[1:first_appr - 1]:
            if not find_heading(p, rules):
                p.kind = "cover"
    current = None
    headings = []
    for p in pages:
        if p.kind in ("cover", "approval"):
            continue
        h = find_heading(p, rules)
        if h:
            h["page"] = p
            p.heading = h
            current = h
            headings.append(h)
        p.section = current
        if current is None:
            p.kind = "front"
        elif current["name"] in FRONT or current["name"] in ("ACKNOWLEDGEMENTS",):
            p.kind = "front"
        elif current["name"] == "CHAPTER":
            p.kind = "back" if current["title"].split(" ")[0:1] and current["title"] in rules["not_chapters"] else "chapter"
        else:
            p.kind = "back"
    return headings


def heading_key(h):
    if h["name"] == "CHAPTER" and h["title"].split(" ")[0] in ("REFERENCES", "BIBLIOGRAPHY", "APPENDIX", "APPENDICES", "BIOGRAPHY", "VITA"):
        return h["title"]
    if h["name"] == "CHAPTER":
        return f"CHAPTER {h['number']}"
    if h["name"] == "APPENDIX":
        return f"APPENDIX {h['letter']}"
    return h["name"]


# ---------------------------------------------------------------- checks

def check_page_size(pages, rules, F):
    r = rules["page"]
    def a4(w, h):
        return abs(w - r["width_cm"]) <= r["tolerance_cm"] and abs(h - r["height_cm"]) <= r["tolerance_cm"]
    bad = [p.no for p in pages if not a4(p.width, p.height) and not a4(p.height, p.width)]
    land = [p.no for p in pages if a4(p.height, p.width)]
    if land:
        F.add("check", "Page setup", f"{pg(land, True)} {isare(land)} landscape. The template has one landscape page, for a wide table.", land[0], r["source"])
    if bad:
        p = pages[bad[0] - 1]
        F.add("fix", "Page setup", f"{pg(bad, True)} {isare(bad)} {r2(p.width)} x {r2(p.height)} cm while the template is A4 (21.0 x 29.7 cm).",
              None, r["source"])


def font_ok(name, rules):
    n = name.replace(" ", "").lower()
    return any(a.replace(" ", "").lower() in n for a in rules["fonts"]["allowed"])


def font_ignored(name, rules):
    n = name.replace(" ", "").lower()
    return any(a.replace(" ", "").lower() in n for a in rules["fonts"]["ignored"])


def check_fonts(pages, rules, F):
    per_font = defaultdict(list)
    for p in pages:
        found = Counter()
        for c in p.chars:
            f = c["fontname"].split("+")[-1]
            if not font_ok(f, rules) and not font_ignored(f, rules):
                found[f] += 1
        for f, n in found.items():
            if n >= 3:
                per_font[f].append(p)
    by_family = defaultdict(set)
    for f, ps in per_font.items():
        by_family[re.sub(r"(MT|PS|-Bold|-Italic|,Bold|,Italic|-Regular).*$", "", f)].update(ps)
    want = rules["fonts"].get("name") or re.sub(r"(?<=[a-z])(?=[A-Z])", " ", rules["fonts"]["allowed"][0])
    for family, ps in by_family.items():
        ps = sorted(ps, key=lambda p: p.no)
        boxes = []
        small = True
        for p in ps:
            for l in p.lines:
                if any(k.startswith(family) for k in l.fonts):
                    boxes.append((p.no, l.bbox_pt()))
                    if l.size > rules["sizes"]["chart_text_max_pt"]:
                        small = False
        where = "chart or pasted table" if small else "text"
        F.add("fix", "Font", f"{family} is used in {where} on {pg([p.no for p in ps])} while the template uses {want} throughout.",
              ps[0].no, rules["fonts"]["source"], boxes[:40])


def heading_lines(pages):
    s = set()
    for p in pages:
        if p.heading:
            for l in p.heading["lines"]:
                s.add(id(l))
    return s


def check_body_size(pages, rules, F):
    body = rules["sizes"]["body_pt"]
    hl = heading_lines(pages)
    by_section = defaultdict(lambda: {"pages": set(), "sizes": Counter(), "boxes": []})
    table_text = defaultdict(lambda: {"pages": set(), "sizes": Counter(), "boxes": []})
    for p in pages:
        if p.kind in ("cover", "approval"):
            continue
        for l in p.lines:
            if id(l) in hl or abs(l.size - body) < 0.3 or l.size <= rules["sizes"]["superscript_max_pt"]:
                continue
            if l.size > body + 0.3 and (l.bold or l.text.isupper()):
                continue  # headings bigger than body are checked elsewhere
            sec = heading_key(p.section) if p.section else "Front pages"
            target = table_text if (l.in_table or l.size < rules["sizes"]["chart_text_max_pt"]) else by_section
            d = target[sec]
            d["pages"].add(p.no)
            d["sizes"][l.size] += 1
            d["boxes"].append((p.no, l.bbox_pt()))
    for sec, d in by_section.items():
        sizes = ", ".join(f"{s:g} pt" for s, _ in d["sizes"].most_common())
        n = sum(d["sizes"].values())
        F.add("fix", "Text size", f"{sec}: {n} {'line is' if n == 1 else 'lines are'} {sizes} on {pg(d['pages'])} while body text in the template is {body} pt.",
              min(d["pages"]), rules["sizes"]["source"], d["boxes"][:60])
    for sec, d in table_text.items():
        sizes = ", ".join(f"{s:g} pt" for s, _ in d["sizes"].most_common())
        F.add("check", "Text size", f"{sec}: text inside tables or charts is {sizes} on {pg(d['pages'])}. The template does not set a size for table text.",
              min(d["pages"]), rules["sizes"]["source"], d["boxes"][:60])


def check_margins(pages, rules, F):
    m = rules["margins"]
    tol = m["tolerance_cm"]
    right_limit = None
    issues = {"left": defaultdict(list), "right": defaultdict(list), "top": defaultdict(list), "bottom": defaultdict(list),
              "img": defaultdict(list), "img_r": defaultdict(list)}
    for p in pages:
        if p.kind == "cover" or p.width > p.height:
            continue
        right_limit = p.width - m["right_cm"]
        bottom_limit = p.height - m["bottom_cm"]
        for l in p.lines:
            if l.x0 < m["left_text_cm"] - tol:
                issues["left"][p.no].append(l)
            if l.x1 > right_limit + tol:
                issues["right"][p.no].append(l)
            if l.top < m["top_text_cm"] - tol:
                issues["top"][p.no].append(l)
            if l.bottom > bottom_limit + tol:
                issues["bottom"][p.no].append(l)
        for (x0, t, x1, b) in p.images:
            # a cropped picture reports its uncropped size, so only count what is drawn on the page
            if p.ink:
                x0, x1 = max(x0, p.ink[0]), min(x1, p.ink[1])
            if x0 < m["left_text_cm"] - tol - 0.2:
                issues["img"][p.no].append((x0, t, x1, b))
            if x1 > right_limit + tol + 0.2:
                issues["img_r"][p.no].append((x0, t, x1, b))

    def boxes(d):
        # one box per page keeps the marked PDF readable
        return [(pn, max(ls, key=lambda l: l.x1 - l.x0).bbox_pt()) for pn, ls in d.items()]

    if issues["left"]:
        allx = [(l.x0, pn) for pn, ls in issues["left"].items() for l in ls]
        common = Counter(round(x, 1) for x, _ in allx).most_common(1)[0][0]
        worst, wp = min(allx)
        extra = f", and as close as {r2(worst)} cm on page {wp}," if worst < common - 0.2 else ""
        F.add("fix", "Margins", f"Text starts {common:.2f} cm from the left edge on {pg(list(issues['left']))}{extra} while the template starts text at {m['left_text_cm']} cm.",
              min(issues["left"]), m["source_left"], boxes(issues["left"]))
    if issues["right"]:
        worst = max(l.x1 for ls in issues["right"].values() for l in ls)
        F.add("fix", "Margins", f"Text runs to {r2(worst)} cm from the left edge on {pg(list(issues['right']))}, past the {m['right_cm']} cm right margin.",
              min(issues["right"]), m["source_right"], boxes(issues["right"]))
    if issues["top"]:
        F.add("fix", "Margins", f"Text sits inside the {m['top_text_cm']} cm top margin on {pg(list(issues['top']))}.",
              min(issues["top"]), m["source_top"], boxes(issues["top"]))
    if issues["bottom"]:
        worst = max(l.bottom for ls in issues["bottom"].values() for l in ls)
        F.add("check", "Margins", f"Text goes down to {r2(pages[0].height - worst)} cm from the bottom edge on {pg(list(issues['bottom']))}. "
              f"The bottom margin is checked at {m['bottom_cm']} cm.", min(issues["bottom"]), m["source_bottom"], boxes(issues["bottom"]))
    if issues["img_r"]:
        worst = max(i[2] for ims in issues["img_r"].values() for i in ims)
        bx = [(pn, (a / PT_TO_CM, b / PT_TO_CM, c / PT_TO_CM, d / PT_TO_CM)) for pn, ims in issues["img_r"].items() for (a, b, c, d) in ims]
        F.add("fix", "Margins", f"Pictures run to {r2(worst)} cm from the left edge on {pg(list(issues['img_r']))}, past the {m['right_cm']} cm right margin.",
              min(issues["img_r"]), m["source_right"], bx)
    if issues["img"]:
        bx = [(pn, (a / PT_TO_CM, b / PT_TO_CM, c / PT_TO_CM, d / PT_TO_CM)) for pn, ims in issues["img"].items() for (a, b, c, d) in ims]
        worst = min(i[0] for ims in issues["img"].values() for i in ims)
        allx = [(i[0], pn) for pn, ims in issues["img"].items() for i in ims]
        common = Counter(round(x, 1) for x, _ in allx).most_common(1)[0][0]
        extra = f", and {r2(worst)} cm on page {min(allx)[1]}," if worst < common - 0.2 else ""
        F.add("fix", "Margins", f"Pictures start {common:.2f} cm from the left edge on {pg(list(issues['img']))}{extra} while the template starts text at {m['left_text_cm']} cm.",
              min(issues["img"]), m["source_left"], bx)


def check_page_numbers(pages, rules, F):
    r = rules["page_number"]
    numbered_early = [p.no for p in pages if p.kind in ("cover", "approval") and p.label]
    if numbered_early:
        F.add("fix", "Page numbers", f"{pg(numbered_early, True)} (cover or approval page) carry a page number while the template leaves them unnumbered.",
              numbered_early[0], r["source"])
    front = [p for p in pages if p.kind == "front"]
    body = [p for p in pages if p.kind in ("chapter", "back")]

    def run(group, style, start_name):
        if not group:
            return
        wrong_style = [p for p in group if p.label and p.label[0] != style]
        missing = [p for p in group if not p.label]
        if wrong_style:
            first = wrong_style[0]
            F.add("fix", "Page numbers", f"{pg([p.no for p in wrong_style], True)} {isare([p.no for p in wrong_style])} numbered {STYLE_NAME[wrong_style[0].label[0]]} "
                  f"while the template numbers {start_name} {STYLE_NAME[style]}.", first.no, r["source"],
                  [(p.no, p.label_line.bbox_pt()) for p in wrong_style])
        if missing:
            F.add("fix", "Page numbers", f"{pg([p.no for p in missing], True)} {'has' if len(missing) == 1 else 'have'} no page number.", missing[0].no, r["source"])
        expected = 1
        for p in group:
            if p.label:
                if p.label[1] != expected:
                    F.add("fix", "Page numbers", f"Page shows {p.shown} where {expected} comes next in sequence.", p.no, r["source"],
                          [(p.no, p.label_line.bbox_pt())])
                    expected = p.label[1]
                expected += 1
            else:
                expected += 1

    run(front, r["front_style"], "the pages from the Abstract to the lists")
    run(body, r["body_style"], "everything from Chapter 1 on")
    labelled = [p for p in pages if p.label_line and p.width < p.height]
    if r.get("align", "right") == "center":
        mid = r.get("center_cm") or text_center(rules, labelled[0] if labelled else pages[0])
        off = [p for p in labelled if abs(p.label_line.center - mid) > r["tolerance_cm"]]
        if off:
            F.add("check", "Page numbers", f"The page number on {pg([p.no for p in off])} sits {r2(off[0].label_line.center)} cm from the left edge "
                  f"while the template centres it at {r2(mid)} cm.",
                  off[0].no, r["source"], [(p.no, p.label_line.bbox_pt()) for p in off])
        return
    groups = defaultdict(list)
    for p in labelled:
        groups[round(p.label_line.x0, 1)].append(p)
    for x0, ps in groups.items():
        x1s = [p.label_line.x1 for p in ps]
        nos = [p.no for p in ps]
        if len(ps) >= 2 and max(x1s) - min(x1s) > 0.15:
            x0 = statistics.median(p.label_line.x0 for p in ps)
            F.add("check", "Page numbers", f"The page number on {pg(nos)} is aligned left at {x0:.1f} cm from the left edge while the template aligns it right, ending at {r['right_edge_cm']} cm.",
                  nos[0], r["source"], [(p.no, p.label_line.bbox_pt()) for p in ps])
        else:
            far = [p for p in ps if abs(p.label_line.x1 - r["right_edge_cm"]) > r["tolerance_cm"]]
            if far:
                F.add("check", "Page numbers", f"The page number on {pg([p.no for p in far])} ends {r2(far[0].label_line.x1)} cm from the left edge while the template ends it at {r['right_edge_cm']} cm.",
                      far[0].no, r["source"], [(p.no, p.label_line.bbox_pt()) for p in far])


def text_center(rules, page):
    return (rules["margins"]["left_text_cm"] + page.width - rules["margins"]["right_cm"]) / 2


def check_structure(pages, headings, rules, F):
    src = rules["order_source"]
    keys = [heading_key(h) for h in headings]
    present = {h["name"] for h in headings}
    if rules.get("approval_page", True) and not any(p.kind == "approval" for p in pages):
        F.add("fix", "Structure", "No approval page was found after the cover.", 2, src)
    for req in rules["required"]:
        if req == "CHAPTER":
            ok = any(h["name"] == "CHAPTER" and h["title"] not in rules["not_chapters"] for h in headings)
        elif req in ("REFERENCES", "BIBLIOGRAPHY") + BIO_NAMES:
            ok = req in present or any(h["name"] == "CHAPTER" and h["title"] == req for h in headings)
        else:
            ok = req in present
        if not ok:
            alias = next((h for h in headings for g in ALIASES if req in g and (h["name"] in g or h.get("title") in g)), None)
            if alias:
                got = alias["title"] if alias["name"] == "CHAPTER" else alias["name"]
                F.add("fix", "Structure", f"The page headed {got} is headed {req} in the template.", alias["page"].no, src,
                      [(alias["page"].no, l.bbox_pt()) for l in alias["lines"]])
            else:
                F.add("fix", "Structure", f"No {req} heading was found.", None, src)

    # back matter dressed up as a chapter
    for h in headings:
        if h["name"] == "CHAPTER":
            t = h["title"]
            if not t:
                F.add("fix", "Structure", f"CHAPTER {h['number']} has no title.", h["page"].no, src, [(h["page"].no, l.bbox_pt()) for l in h["lines"]])
            elif any(t == n or t.startswith(n) for n in rules["not_chapters"]):
                noun = {"REFERENCES": "The references are", "BIBLIOGRAPHY": "The bibliography is", "APPENDIX": "The appendix is", "APPENDICES": "The appendices are",
                        "BIOGRAPHY": "The biography is", "VITA": "The vita is"}.get(t.split()[0], f"{t} is")
                head = {"APPENDIX": "APPENDICES, then APPENDIX A, B, C"}.get(t.split()[0], t.split()[0])
                F.add("fix", "Structure", f"{noun} headed CHAPTER {h['number']} while the template heads {'it' if ' is' in noun else 'them'} {head} with no chapter number.",
                      h["page"].no, src, [(h["page"].no, l.bbox_pt()) for l in h["lines"]])

    # order
    rank = {n: i for i, n in enumerate(rules["order"])}

    def rank_of(h):
        for g in ALIASES:
            name = h["title"] if h["name"] == "CHAPTER" else h["name"]
            if name in g:
                for alt in g:
                    if alt in rank:
                        return rank[alt]
        if h["name"] == "CHAPTER":
            t = h["title"]
            for n in rules["not_chapters"]:
                if t.startswith(n):
                    return rank.get("APPENDIX" if n.startswith("APPENDI") else n, rank["CHAPTER"])
            return rank["CHAPTER"]
        return rank.get(h["name"], rank["CHAPTER"])

    def shown_name(h):
        if h["name"] == "CHAPTER":
            return h["title"] if any(h["title"].startswith(n) for n in rules["not_chapters"]) else f"CHAPTER {h['number']}"
        return heading_key(h)

    last = -1
    last_name = None
    for h in headings:
        r = rank_of(h)
        if r < last:
            F.add("fix", "Structure", f"{shown_name(h)} comes after {last_name} while the template puts {shown_name(h)} first.", h["page"].no, src)
        else:
            last, last_name = r, shown_name(h)
    # chapter numbers in sequence
    nums = [h["number"] for h in headings if h["name"] == "CHAPTER"]
    for a, b in zip(nums, nums[1:]):
        if b != a + 1:
            F.add("fix", "Structure", f"CHAPTER {b} follows CHAPTER {a}.", None, src)
    # chapter titles 1 to 5 against the template
    for h in headings:
        if h["name"] == "CHAPTER":
            want = rules["chapter_titles"].get(str(h["number"]))
            if want and h["title"] and h["title"] != want and not any(h["title"].startswith(n) for n in rules["not_chapters"]):
                F.add("check", "Structure", f"CHAPTER {h['number']} is titled {h['title']} while the template names it {want}.",
                      h["page"].no, "Template PDF chapter titles. Chapters 6 to 8 are open in the template, so other names may be accepted.")

    # how each heading looks
    hp = rules["sizes"]["heading_pt"]
    for h in headings:
        p = h["page"]
        before = [l for l in p.lines[:h["index"]] if l not in h["lines"]]
        if h["name"] not in ("ABSTRACT",) and before:
            F.add("fix", "Headings", f"{heading_key(h)} does not start at the top of a new page.", p.no, src, [(p.no, h["lines"][0].bbox_pt())])
        for l in h["lines"]:
            probs = []
            if abs(l.size - hp) > 0.3:
                probs.append(f"{l.size:g} pt")
            if not l.bold:
                probs.append("not bold")
            if abs(l.center - text_center(rules, p)) > 0.6:
                probs.append("not centred")
            if l.text.upper() != l.text:
                probs.append("not in capitals")
            if probs:
                F.add("fix", "Headings", f"Heading \"{norm(l.text)}\" is {join_and(probs)} while the template sets headings in {hp:g} pt bold capitals, centred.",
                      p.no, rules["sizes"]["source"], [(p.no, l.bbox_pt())])
        if h["name"] == "CHAPTER" and h.get("same_line"):
            F.add("fix", "Headings", f"CHAPTER {h['number']} and its title share one line while the template puts the title on the line below.",
                  p.no, src, [(p.no, h["lines"][0].bbox_pt())])


SEC_RE = re.compile(r"^(\d+)\.(\d+)((?:\.\d+)*)\.?\s+(\S.*)$")


def section_headings(pages, rules):
    out = []
    for p in pages:
        if p.kind != "chapter":
            continue
        for l in p.lines:
            if l.in_table:
                continue
            m = SEC_RE.match(norm(l.text))
            if m and l.bold and len(m.group(4)) > 1:
                level = 2 + m.group(3).count(".")
                out.append({"page": p, "line": l, "chapter": int(m.group(1)), "num": norm(l.text).split()[0].rstrip("."),
                            "level": level, "title": m.group(4)})
    return out


def check_sections(pages, rules, F):
    r = rules["section_heading"]
    secs = section_headings(pages, rules)
    bad_x = []
    bad_size = []
    for s in secs:
        p, l = s["page"], s["line"]
        ch = p.section["number"] if p.section and p.section["name"] == "CHAPTER" else None
        if ch is not None and s["chapter"] != ch:
            F.add("fix", "Headings", f"Section {s['num']} sits in CHAPTER {ch}.", p.no, r["source"], [(p.no, l.bbox_pt())])
        if abs(l.size - r["size_pt"]) > 0.3:
            bad_size.append(s)
        if s["level"] == 2 and abs(l.x0 - r["level1_x_cm"]) > r["tolerance_cm"]:
            bad_x.append(s)
    if bad_size:
        F.add("fix", "Headings", f"{len(bad_size)} section headings are not {r['size_pt']:g} pt, on {pg([s['page'].no for s in bad_size])}.",
              bad_size[0]["page"].no, r["source"], [(s["page"].no, s["line"].bbox_pt()) for s in bad_size])
    if bad_x:
        common = Counter(round(s["line"].x0, 2) for s in bad_x).most_common(1)[0][0]
        ex = next(s for s in bad_x if round(s["line"].x0, 2) == common)
        two = all(len(s["num"].split(".")[-1]) > 1 for s in bad_x)
        which = "numbered 10 and up, such as" if two else "such as"
        F.add("fix", "Headings", f"{len(bad_x)} section headings, {which} {ex['num']}, start {common:.2f} cm from the left edge while the template puts them at the "
              f"{r['level1_x_cm']} cm margin. {pg([s['page'].no for s in bad_x], True)}.",
              ex["page"].no, r["source"], [(s["page"].no, s["line"].bbox_pt()) for s in bad_x])
    # numbering runs 1, 2, 3 within each chapter
    by_parent = defaultdict(list)
    for s in secs:
        parts = s["num"].split(".")
        by_parent[".".join(parts[:-1])].append((int(parts[-1]), s))
    for parent, items in by_parent.items():
        for (a, sa), (b, sb) in zip(items, items[1:]):
            if b != a + 1:
                F.add("fix", "Headings", f"Section {sb['num']} follows {sa['num']}.", sb["page"].no, r["source"], [(sb["page"].no, sb["line"].bbox_pt())])
    return secs


def spacing_name(pitch_cm, size_pt, noun=False):
    ratio = pitch_cm / (size_pt * 1.15 * 2.54 / 72)
    for value, adj, n in ((1.0, "single", "single spacing"), (1.15, "1.15-line", "1.15 lines"), (1.5, "1.5-line", "1.5 lines"), (2.0, "double", "double spacing")):
        if abs(ratio - value) < 0.12:
            return n if noun else adj
    return f"{ratio:.1f} lines" if noun else f"{ratio:.1f}-line"


def body_lines(p, rules, hl):
    out = []
    sizes = Counter(l.size for l in p.lines if not l.bold and not l.in_table and l.size > rules["sizes"]["superscript_max_pt"])
    main = sizes.most_common(1)[0][0] if sizes else rules["sizes"]["body_pt"]
    for l in p.lines:
        if id(l) in hl or l.in_table or l.bold or abs(l.size - main) > 0.6:
            continue
        out.append(l)
    return out


def check_spacing_and_indent(pages, rules, F):
    hl = heading_lines(pages)
    ls_rule = rules["line_spacing"]
    ind_rule = rules["paragraph_indent"]
    left = rules["margins"]["left_text_cm"]
    single = {}
    no_indent = defaultdict(list)
    indented = defaultdict(int)
    for p in pages:
        if p.kind not in ("chapter",) and not (p.section and p.section["name"] in ("ABSTRACT", "ACKNOWLEDGEMENTS")):
            continue
        ls = body_lines(p, rules, hl)
        right = p.width - rules["margins"]["right_cm"]
        pitches = []
        for a, b in zip(ls, ls[1:]):
            d = b.top - a.top
            if 0.3 < d < 1.0 and a.x1 > right - 1.0:
                pitches.append(d)
        if len(pitches) >= 4:
            med = statistics.median(pitches)
            if med < ls_rule["fail_below_cm"]:
                single[p.no] = med
        # paragraph starts: a line after a short line, or after a gap
        if p.kind != "chapter":
            continue
        sec = heading_key(p.section)
        page_left = min([left] + [x.x0 for x in p.lines if not x.in_table])
        for prev, l in zip(ls, ls[1:]):
            gap = l.top - prev.top
            starts = prev.x1 < right - 1.5 or gap > 1.0
            if not starts or l.x1 < right - 1.5:
                continue
            if re.match(r"^(\(?\d+[\).]|[•·\-–]|[a-z]\))\s", l.text):
                continue  # list item
            if abs(l.x0 - page_left) < 0.35:
                no_indent[sec].append(l)
            elif l.x0 - page_left > 0.8:
                indented[sec] += 1
    if single:
        med = statistics.median(single.values())
        size = rules["sizes"]["body_pt"]
        F.add("fix", "Line spacing", f"{pg(list(single), True)} {isare(list(single))} {spacing_name(med, size)} spaced ({r2(med)} cm from line to line) "
              f"while the template uses {spacing_name(ls_rule['expected_cm'], size, noun=True)} ({ls_rule['expected_cm']} cm).", min(single), ls_rule["source"], pages=list(single))
    for sec, ls in no_indent.items():
        if len(ls) >= 2 and len(ls) >= indented[sec]:
            F.add("fix", "Paragraphs", f"{sec}: {len(ls)} paragraphs start at the margin while the template sets body text "
                  f"{ind_rule['expected_cm']} cm in from the margin. {pg([l.page for l in ls], True)}.",
                  ls[0].page, ind_rule["source"], [(l.page, l.bbox_pt()) for l in ls][:60])


def check_placeholders(pages, rules, F):
    hits = defaultdict(list)
    for p in pages:
        for l in p.lines + p.header:
            for ph in rules["placeholders"]:
                if re.search(r"(?<![A-Za-z])" + re.escape(ph) + r"(?![A-Za-z])", l.text):
                    hits[ph].append((p.no, l))
    by_pages = defaultdict(list)
    for ph, items in hits.items():
        by_pages[page_ranges([n for n, _ in items])].append((ph, items))
    for rng, group in by_pages.items():
        words = ", ".join(f"\"{ph}\"" for ph, _ in group)
        items = [it for _, its in group for it in its]
        F.add("fix", "Template text left in", f"Template text is still on {pg([n for n, _ in items])}: {words}.",
              min(n for n, _ in items), "Template placeholder text", [(n, l.bbox_pt()) for n, l in items])
    for typo, right in rules["template_typos"].items():
        found = [(p.no, l) for p in pages for l in p.lines if typo in l.text]
        if found:
            F.add("fix", "Spelling", f"\"{typo}\" on {pg([n for n, _ in found])} is a typing error copied from the template and should read \"{right}\".",
                  found[0][0], "The template itself carries this typing error", [(n, l.bbox_pt()) for n, l in found])


def block_between(lines, start_pred, stop_pred):
    out = []
    on = False
    for l in lines:
        if on and stop_pred(l):
            break
        if on:
            out.append(l)
        if not on and start_pred(l):
            on = True
    return out


def dehyphen(s):
    s = re.sub(r"-\s+(?=[a-z])", "-", s)
    return norm(s)


def check_titles(pages, rules, F):
    titles = {}
    cover = pages[0]
    by_i = [i for i, l in enumerate(cover.lines) if norm(l.text).upper() == "BY"]
    if by_i:
        tl = cover.lines[:by_i[0]]
        titles["cover page"] = (dehyphen(" ".join(l.text for l in tl)), cover.no, tl)
    appr = next((p for p in pages if p.kind == "approval"), None)
    if appr:
        tl = block_between(appr.lines, lambda l: norm(l.text).upper() == "ENTITLED", lambda l: "was approved" in l.text.lower())
        if tl:
            titles["approval page"] = (dehyphen(" ".join(l.text for l in tl)), appr.no, tl)
    abs_page = next((p for p in pages if p.section and p.section["name"] == "ABSTRACT"), None)
    if abs_page:
        txt = "\n".join(l.text for l in abs_page.lines)
        m = re.search(r"Thesis Title\s+(.*?)\n\s*Author\b", txt, re.S)
        if m:
            tl = [l for l in abs_page.lines if l.text.split("\n")[0] in m.group(0)][:6]
            titles["abstract page"] = (dehyphen(m.group(1).replace("\n", " ")), abs_page.no, tl)
    keyed = {k: v[0].lower().rstrip(".") for k, v in titles.items()}
    if len(set(keyed.values())) > 1:
        groups = defaultdict(list)
        for k, v in titles.items():
            groups[keyed[k]].append(k)
        parts = [f"{join_and(ks).capitalize()}: \"{titles[ks[0]][0]}\"" for ks in groups.values()]
        boxes = [(v[1], l.bbox_pt()) for v in titles.values() for l in v[2]]
        F.add("fix", "Title", f"The thesis title is not the same on every page. {'. '.join(parts)}.", None, "The same title appears on the cover, approval and abstract pages of the template", boxes)
    if "cover page" in titles:
        t, n, tl = titles["cover page"]
        if t != t.upper():
            F.add("check", "Cover", f"The cover title is in mixed case while the template shows it in capitals.", n,
                  "Template PDF cover page", [(n, l.bbox_pt()) for l in tl])
        for l in tl:
            if abs(l.size - rules["sizes"]["cover_title_pt"]) > 0.3 or not l.bold:
                F.add("fix", "Cover", f"The cover title is {l.size:g} pt{'' if l.bold else ' and not bold'} while the template uses "
                      f"{rules['sizes']['cover_title_pt']} pt bold.", n, rules["sizes"]["source"], [(n, l.bbox_pt())])
                break
    # the rest of the cover
    if by_i:
        rest = cover.lines[by_i[0]:]
        bad = [l for l in rest if abs(l.size - rules["sizes"]["cover_text_pt"]) > 0.3 or not l.bold]
        if bad:
            F.add("fix", "Cover", f"{len(bad)} cover lines are not {rules['sizes']['cover_text_pt']} pt bold.", cover.no, rules["sizes"]["source"],
                  [(cover.no, l.bbox_pt()) for l in bad])
        import difflib
        have = [norm(l.text).upper() for l in rest]
        joined = " ".join(have)

        def present(want):
            if want in joined:
                return True
            return any(difflib.SequenceMatcher(None, want, h[:len(want) + 4]).ratio() > 0.85 for h in have)

        for want in rules["cover_lines"]:
            if not present(want):
                F.add("fix", "Cover", f"The cover is missing the line \"{want}\".", cover.no, "Template PDF cover page")
        name = norm(rest[1].text) if len(rest) > 1 else ""
        if re.search(r"\b(M\.?D|Ph\.?D|MBA|M\.?Sc|B\.?Sc|RN|DMD)\b", name):
            F.add("check", "Cover", f"The author line \"{name}\" carries degrees while the template shows the name only.", cover.no,
                  "Template PDF cover page", [(cover.no, rest[1].bbox_pt())])


def check_abstract_fields(pages, rules, F):
    abs_page = next((p for p in pages if p.section and p.section["name"] == "ABSTRACT"), None)
    if not abs_page:
        return
    pages_abs = [p for p in pages if p.section and p.section["name"] == "ABSTRACT"]
    txt = "\n".join(p.text for p in pages_abs)
    flat = txt.replace(" ", "")
    for f in rules["abstract_fields"]:
        if f.replace(" ", "") not in flat:
            F.add("fix", "Abstract", f"The abstract page is missing the \"{f}\" line.", abs_page.no, "Template PDF abstract page")
    if rules.get("abstract_keywords", True) and not re.search(r"Keywords?\s*:", txt):
        F.add("fix", "Abstract", "The abstract has no \"Keywords:\" line.", abs_page.no, "Template PDF abstract page")
    for p in pages_abs:
        for l in p.lines:
            if re.match(r"^Faculty of [a-z]", l.text):
                fixed = re.sub(r"^Faculty of (\w)", lambda m: "Faculty of " + m.group(1).upper(), norm(l.text))
                F.add("check", "Abstract", f"\"{norm(l.text)}\" should read \"{fixed}\".", p.no, "Spelling", [(p.no, l.bbox_pt())])
    # academic year agrees with the cover
    cy = re.search(r"ACADEMIC YEAR\s+(\d{4})", pages[0].text)
    ay = re.search(r"Academic Year\s+(\d{4})", txt)
    if cy and ay and cy.group(1) != ay.group(1):
        F.add("fix", "Abstract", f"Academic year is {ay.group(1)} on the abstract page and {cy.group(1)} on the cover.", abs_page.no, "Template PDF")


def label_of(p):
    return p.label_line.text.strip() if p.label_line else None


def check_lists(pages, rules, F):
    for kind, word in (("LIST OF TABLES", "Table"), ("LIST OF FIGURES", "Figure")):
        lp = [p for p in pages if p.section and p.section["name"] == kind]
        if not lp:
            continue
        entries = []
        for p in lp:
            for l in p.lines:
                t = norm(l.text)
                m = re.match(rf"^(?:{word}s?\s*)?(\d+(?:\.\d+)?)\s*(.*?)\s*\.*\s*([ivxlc\d()]+)$", t, re.I)
                if m and not t.upper().startswith(kind) and t.lower() not in (f"{word.lower()}s page",):
                    entries.append((m.group(1), m.group(2).strip(" ."), m.group(3), p, l))
        if not entries:
            continue
        no_title = [e for e in entries if not e[1]]
        if no_title:
            msg = (f"All {len(entries)} entries in the {kind} show the {word.lower()} number with no title." if len(no_title) == len(entries)
                   else f"{len(no_title)} of {len(entries)} entries in the {kind} show the {word.lower()} number with no title.")
            F.add("fix", "Lists", msg,
                  lp[0].no, "Template PDF lists show number, title and page", [(e[3].no, e[4].bbox_pt()) for e in no_title])
        if all("." not in e[0] for e in entries):
            F.add("check", "Lists", f"{word}s are numbered 1, 2, 3 across the thesis while the template numbers them by chapter (4.1, 4.2).",
                  lp[0].no, "Template PDF lists")
        # where each caption really is
        caps = {}
        for p in pages:
            if p.kind in ("chapter", "back"):
                for l in p.lines:
                    m = re.match(rf"^{word}\s+(\d+(?:\.\d+)?)\b", norm(l.text))
                    if m and m.group(1) not in caps:
                        caps[m.group(1)] = p
        for num, title, shown, p, l in entries:
            real = caps.get(num)
            if not real:
                F.add("check", "Lists", f"{word} {num} is listed on page {shown} but no caption starting \"{word} {num}\" was found.", p.no,
                      "Captions", [(p.no, l.bbox_pt())])
            elif label_of(real) and label_of(real) != shown:
                F.add("fix", "Lists", f"{word} {num} is listed on page {shown} while its caption is on page {label_of(real)}.", p.no, "Captions",
                      [(p.no, l.bbox_pt())])


TOC_RE = re.compile(r"^(.*?)[\s.]*\.{3,}\s*([ivxlc\d()]+)\s*$|^(.*?\S)\s{1,}([ivxlc]+|\d+|\(\d+\))\s*$", re.I)


def check_toc(pages, headings, secs, rules, F):
    tp = [p for p in pages if p.section and p.section["name"] in TOC_NAMES]
    if not tp:
        return
    entries = []
    pending = ""
    for p in tp:
        for l in p.lines:
            t = norm(l.text)
            if t.upper() in TOC_NAMES + ("PAGE",):
                continue
            m = TOC_RE.match(t)
            if m:
                text = (pending + " " + (m.group(1) if m.group(1) is not None else m.group(3))).strip(" .")
                entries.append((norm(text), m.group(2) or m.group(4), p, l))
                pending = ""
            else:
                pending = (pending + " " + t).strip()
    if rules.get("toc_page_label", True) and not any(norm(l.text).upper() == "PAGE" for p in tp for l in p.lines):
        F.add("check", "Table of contents", "The table of contents has no \"Page\" label above the numbers.", tp[0].no, "Template PDF table of contents")
    where = {}
    for h in headings:
        key = heading_key(h)
        where[key] = h["page"]
        if h["name"] == "CHAPTER" and h["title"]:
            where[f"CHAPTER {h['number']} {h['title']}"] = h["page"]
    for s in secs:
        where.setdefault(s["num"], s["page"])
    for text, shown, p, l in entries:
        if re.fullmatch(r"CHAPTER\s+\d+", text.upper()):
            F.add("fix", "Table of contents", f"\"{text}\" is listed with no chapter title.", p.no, "Template PDF table of contents", [(p.no, l.bbox_pt())])
    chap_x = [(l.x0, p, l, text) for text, shown, p, l in entries if text.upper().startswith("CHAPTER")]
    if chap_x:
        common = Counter(round(x, 1) for x, *_ in chap_x).most_common(1)[0][0]
        odd = [(p, l, t) for x, p, l, t in chap_x if abs(x - common) > 0.3]
        if odd:
            F.add("check", "Table of contents", join_and([f"\"{t}\"" for _, _, t in odd]) + (" is" if len(odd) == 1 else " are")
                  + " not lined up with the other chapter lines in the table of contents.", odd[0][0].no, "Template PDF table of contents", [(p.no, l.bbox_pt()) for p, l, _ in odd])
    for text, shown, p, l in entries:
        up = text.upper()
        m = re.match(r"^(\d+(?:\.\d+)+)\s", text)
        target = None
        if m:
            target = where.get(m.group(1))
        else:
            mc = re.match(r"^CHAPTER\s+(\d+)", up)
            if mc:
                target = where.get(f"CHAPTER {mc.group(1)}")
            else:
                target = where.get(re.sub(r"\s*\(IF ANY\)", "", up))
        if target is None:
            continue
        real = label_of(target)
        if real and real.replace(" ", "") != shown.replace(" ", ""):
            F.add("fix", "Table of contents", f"\"{text}\" is listed on page {shown} while it starts on page {real}.", p.no, "Page numbers",
                  [(p.no, l.bbox_pt())])


def check_back(pages, rules, F):
    bio = [p for p in pages if p.section and (p.section["name"] in BIO_NAMES or p.section.get("title") in BIO_NAMES)]
    if bio:
        txt = norm(" ".join(p.text for p in bio))
        for f in rules["biography_fields"]:
            if not all(w in txt.split() for w in f.split()):
                F.add("fix", "Biography", f"The biography is missing the \"{f}\" line.", bio[0].no, "Template PDF biography page")
    ack = [p for p in pages if p.section and p.section["name"] == "ACKNOWLEDGEMENTS"]
    if ack and rules.get("acknowledgements_signed", True):
        last = ack[-1].lines[-1] if ack[-1].lines else None
        if last and last.x0 < ack[-1].width / 2:
            F.add("check", "Acknowledgements", "The acknowledgements do not end with the author's name on the right.", ack[-1].no, "Template PDF acknowledgements page")


# ---------------------------------------------------------------- output

LEVEL_NAME = {"fix": "Fix", "check": "Check"}


def write_report(path, src_name, pages, findings, rules):
    items = findings.numbered()
    n_fix = sum(1 for f in items if f["level"] == "fix")
    n_check = len(items) - n_fix
    shown = {p.no: p.shown for p in pages}
    rows = []
    for i, f in enumerate(items, 1):
        pg = f"{f['page']}" + (f" <span class=lab>({html.escape(shown.get(f['page'], ''))})</span>" if f["page"] else "") if f["page"] else "Whole file"
        rows.append(f"<tr class={f['level']}><td>{i}</td><td><b>{LEVEL_NAME[f['level']]}</b></td><td>{html.escape(f['area'])}</td>"
                    f"<td>{pg}</td><td>{html.escape(f['message'])}<div class=src>{html.escape(f['source'] or '')}</div></td></tr>")
    doc = f"""<!DOCTYPE html><html lang="en"><head><meta charset="utf-8">
<title>{html.escape(src_name)} format report</title>
<style>
body{{font:15px/1.5 Georgia,'Times New Roman',serif;color:#1d1d1b;max-width:1100px;margin:32px auto;padding:0 24px}}
h1{{font-size:22px;margin:0 0 4px}} .meta{{color:#555;margin-bottom:20px}}
.sum{{display:flex;gap:28px;margin:16px 0 24px;font-size:17px}} .sum b{{font-size:26px;display:block}}
table{{border-collapse:collapse;width:100%}} th,td{{text-align:left;vertical-align:top;padding:8px 10px;border-bottom:1px solid #ddd}}
th{{font-size:13px;color:#555;border-bottom:2px solid #1d1d1b}}
tr.fix td:nth-child(2){{color:#a3241c}} tr.check td:nth-child(2){{color:#8a5a00}}
.src{{color:#6b6b6b;font-size:12.5px;margin-top:3px}} .lab{{color:#6b6b6b;font-size:12.5px}}
td:nth-child(1){{color:#999;width:28px}} td:nth-child(2){{width:52px}} td:nth-child(3){{width:150px}} td:nth-child(4){{width:80px}}
p.note{{color:#444;max-width:820px}}
@media print{{body{{margin:0}} tr{{break-inside:avoid}}}}
</style></head><body>
<h1>{html.escape(src_name)}</h1>
<div class=meta>Checked against {html.escape(rules['name'])}. {len(pages)} pages.</div>
<div class=sum><div><b>{n_fix}</b>to fix</div><div><b>{n_check}</b>to check</div></div>
<p class=note>Fix means the manuscript differs from the template. Check means the template does not settle it and the graduate office decides.
Page is the PDF page, with the number printed on that page in brackets. The corrections PDF shows each page with every problem boxed and numbered to match this list.</p>
<table><tr><th>#</th><th></th><th>Area</th><th>Page</th><th>Problem</th></tr>
{''.join(rows)}
</table></body></html>"""
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(doc)


def write_corrections_pdf(src_pdf, out_pdf, src_name, pages, findings, rules):
    """Printable A4 landscape sheets: a summary list, then every page with a problem,
    shown as a picture with numbered boxes and the matching notes beside it."""
    import io
    import pypdfium2 as pdfium
    from pypdf import PdfReader, PdfWriter
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import cm as CM
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfgen import canvas
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    W, H = landscape(A4)
    RED = colors.Color(0.75, 0.1, 0.1)
    AMBER = colors.Color(0.82, 0.5, 0.0)
    items = findings.numbered()
    n_fix = sum(1 for f in items if f["level"] == "fix")
    shown = {p.no: p.shown for p in pages}
    esc = html.escape

    # summary sheets
    base = ParagraphStyle("b", fontName="Times-Roman", fontSize=10, leading=12.5)
    small = ParagraphStyle("s", parent=base, fontSize=8, leading=10, textColor=colors.Color(0.4, 0.4, 0.4))
    title = ParagraphStyle("t", parent=base, fontName="Times-Bold", fontSize=16, leading=20)
    rows = [[Paragraph("<b>#</b>", base), Paragraph("", base), Paragraph("<b>Area</b>", base), Paragraph("<b>Page</b>", base), Paragraph("<b>Problem</b>", base)]]
    for f in items:
        col = "#a3241c" if f["level"] == "fix" else "#a56400"
        pg_txt = f"{f['page']} <font size=8 color='#666666'>({esc(shown.get(f['page'], ''))})</font>" if f["page"] else "Whole file"
        rows.append([Paragraph(str(f["n"]), base), Paragraph(f"<font color='{col}'><b>{LEVEL_NAME[f['level']]}</b></font>", base),
                     Paragraph(esc(f["area"]), base), Paragraph(pg_txt, base),
                     [Paragraph(esc(f["message"]), base), Paragraph(esc(f["source"] or ""), small)]])
    tbl = Table(rows, colWidths=[0.9 * CM, 1.4 * CM, 3.6 * CM, 2.2 * CM, 19.6 * CM], repeatRows=1)
    tbl.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LINEBELOW", (0, 0), (-1, 0), 1, colors.black),
                             ("LINEBELOW", (0, 1), (-1, -1), 0.3, colors.Color(0.8, 0.8, 0.8)),
                             ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4)]))
    story = [Paragraph(esc(src_name), title), Spacer(1, 4),
             Paragraph(f"Checked against {esc(rules['name'])}. {len(pages)} pages. <b>{n_fix} to fix, {len(items) - n_fix} to check.</b>", base),
             Spacer(1, 4),
             Paragraph("Fix means the manuscript differs from the template. Check means the template does not settle it and the graduate office decides. "
                       "Page is the PDF page, with the number printed on that page in brackets. The sheets after this list show each page with its problems boxed and numbered.", base),
             Spacer(1, 10), tbl]
    summary = io.BytesIO()
    SimpleDocTemplate(summary, pagesize=(W, H), leftMargin=1.2 * CM, rightMargin=1.2 * CM, topMargin=1.2 * CM, bottomMargin=1.2 * CM,
                      title=f"{src_name} corrections").build(story)

    # which notes and boxes belong on each page
    on_page = defaultdict(list)
    for f in items:
        boxed = defaultdict(list)
        for pno, bb in f["boxes"]:
            boxed[pno].append(bb)
        for pno, bbs in boxed.items():
            on_page[pno].append((f, bbs))
        for pno in f["pages"]:
            if pno not in boxed:
                on_page[pno].append((f, []))
        if f["page"] and not f["boxes"] and f["page"] not in f["pages"]:
            on_page[f["page"]].append((f, []))

    sheets = io.BytesIO()
    c = canvas.Canvas(sheets, pagesize=(W, H))
    doc = pdfium.PdfDocument(src_pdf)
    margin = 1.0 * CM
    note_x = 15.4 * CM
    note_w = W - note_x - margin
    for pno in sorted(on_page):
        page = pages[pno - 1]
        notes = sorted(on_page[pno], key=lambda t: t[0]["n"])
        img = doc[pno - 1].render(scale=150 / 72, grayscale=True).to_pil()
        pw, ph = page.width / PT_TO_CM, page.height / PT_TO_CM  # page size in points
        avail_h = H - 2 * margin
        avail_w = note_x - 2 * margin
        scale = min(avail_h / ph, avail_w / pw)
        iw, ih = pw * scale, ph * scale
        ix, iy = margin, H - margin - ih
        buf = io.BytesIO()
        img.save(buf, "PNG", optimize=True)
        buf.seek(0)
        c.drawImage(ImageReader(buf), ix, iy, iw, ih)
        c.setStrokeColor(colors.Color(0.6, 0.6, 0.6))
        c.setLineWidth(0.5)
        c.rect(ix, iy, iw, ih)
        # boxes and number tags
        placed = []
        for f, bbs in notes:
            col = RED if f["level"] == "fix" else AMBER
            c.setStrokeColor(col)
            c.setLineWidth(1.1)
            for k, (x0, t, x1, b) in enumerate(bbs):
                rx, ry = ix + (x0 - 2) * scale, iy + ih - (b + 2) * scale
                c.rect(rx, ry, (x1 - x0 + 4) * scale, (b - t + 4) * scale, stroke=1, fill=0)
                if k == 0:
                    tx = max(ix + 6, rx - 7)
                    ty = ry + (b - t + 4) * scale / 2
                    while any(abs(tx - px) < 12 and abs(ty - py) < 12 for px, py in placed):
                        tx -= 13
                    placed.append((tx, ty))
                    c.setFillColor(col)
                    c.circle(tx, ty, 6, stroke=0, fill=1)
                    c.setFillColor(colors.white)
                    c.setFont("Helvetica-Bold", 6.5 if f["n"] < 10 else 5.5)
                    c.drawCentredString(tx, ty - 2.3, str(f["n"]))
        # notes column
        y = H - margin
        c.setFillColor(colors.black)
        c.setFont("Times-Bold", 13)
        c.drawString(note_x, y - 12, f"PDF page {pno}")
        c.setFont("Times-Roman", 10)
        sec = heading_key(page.section) if page.section else ("Cover" if page.kind == "cover" else "Approval page" if page.kind == "approval" else "")
        c.drawString(note_x, y - 26, f"Printed page number: {page.shown}" + (f"    Section: {sec}" if sec else ""))
        y -= 40
        size = 10
        while True:  # shrink the text until every note fits beside the page
            st = ParagraphStyle("n", fontName="Times-Roman", fontSize=size, leading=size * 1.25)
            paras = []
            total = 0
            for f, _ in notes:
                col = "#a3241c" if f["level"] == "fix" else "#a56400"
                para = Paragraph(f"<font color='{col}'><b>{f['n']}. {LEVEL_NAME[f['level']]}, {esc(f['area'])}.</b></font> {esc(f['message'])}", st)
                _, h = para.wrap(note_w - 0.6 * CM, H)
                paras.append((f, para, h))
                total += h + size * 0.6
            if total <= y - margin or size <= 7:
                break
            size -= 0.5
        for f, para, h in paras:
            if y - h < margin:
                break
            col = RED if f["level"] == "fix" else AMBER
            c.setFillColor(col)
            c.circle(note_x + 6, y - size * 0.75, 6, stroke=0, fill=1)
            c.setFillColor(colors.white)
            c.setFont("Helvetica-Bold", 6.5 if f["n"] < 10 else 5.5)
            c.drawCentredString(note_x + 6, y - size * 0.75 - 2.3, str(f["n"]))
            para.drawOn(c, note_x + 0.6 * CM, y - h)
            y -= h + size * 0.6
        c.setFillColor(colors.Color(0.45, 0.45, 0.45))
        c.setFont("Times-Roman", 8)
        c.drawRightString(W - margin, margin / 2, f"{src_name}  |  PDF page {pno}")
        c.showPage()
    c.save()

    w = PdfWriter()
    for part in (summary, sheets):
        part.seek(0)
        for pg_ in PdfReader(part).pages:
            w.add_page(pg_)
    with open(out_pdf, "wb") as fh:
        w.write(fh)


# ---------------------------------------------------------------- Word files

WORD_TIMEOUT = 240          # seconds Word gets for one file before it is shut down
CANCEL = threading.Event()  # set by the Stop button
NO_WINDOW = 0x08000000 if os.name == "nt" else 0   # keep helper programs from flashing a console


class Cancelled(Exception):
    pass


def word_job(kind, src, out):
    """Runs in its own process: Word opens src and saves it as PDF (kind "pdf") or .docx (kind "docx").
    Macros, link updates and spelling are switched off so nothing can stop Word with a hidden question."""
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Word.Application")
    try:
        app.Visible = False
        app.DisplayAlerts = 0
        for setting, value in (("AutomationSecurity", 3),):          # 3 = macros off
            try:
                setattr(app, setting, value)
            except Exception:
                pass
        for option, value in (("ConfirmConversions", False), ("UpdateLinksAtOpen", False), ("CheckSpellingAsYouType", False),
                              ("CheckGrammarAsYouType", False), ("UpdateFieldsAtPrint", False), ("UpdateLinksAtPrint", False),
                              ("Pagination", True)):
            try:
                setattr(app.Options, option, value)
            except Exception:
                pass
        doc = app.Documents.Open(FileName=os.path.abspath(src), ConfirmConversions=False, ReadOnly=True, AddToRecentFiles=False,
                                 Revert=False, Visible=False, OpenAndRepair=False, NoEncodingDialog=True)
        try:
            if kind == "pdf":
                try:  # show the final text, without tracked changes or comment balloons
                    view = doc.ActiveWindow.View
                    view.RevisionsFilter.Markup = 0
                    view.RevisionsFilter.View = 0
                    view.ShowRevisionsAndComments = False
                except Exception:
                    pass
                for _ in range(2):  # only the contents and lists are refreshed, never fields that call other programs
                    for toc in doc.TablesOfContents:
                        toc.Update()
                    for tof in doc.TablesOfFigures:
                        tof.Update()
                doc.ExportAsFixedFormat(OutputFileName=os.path.abspath(out), ExportFormat=17, OpenAfterExport=False,
                                        OptimizeFor=0, Range=0, Item=0, IncludeDocProps=False, CreateBookmarks=0)
            else:
                doc.SaveAs2(os.path.abspath(out), 16)
        finally:
            doc.Close(0)
    finally:
        try:
            app.Quit(0)
        except Exception:
            pass


def self_command():
    """How to start this program again, for running Word in a separate process."""
    if getattr(sys, "frozen", False):
        return [sys.executable]
    return [sys.executable, os.path.join(os.path.dirname(HERE), "launcher.py")]


def word_pids():
    if os.name != "nt":
        return set()
    try:
        out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE", "/FO", "CSV", "/NH"], capture_output=True, text=True,
                             timeout=20, creationflags=NO_WINDOW).stdout
    except Exception:
        return set()
    pids = set()
    for line in out.splitlines():
        parts = [x.strip('"') for x in line.split('","')]
        if len(parts) > 1 and parts[1].isdigit():
            pids.add(int(parts[1]))
    return pids


def kill_pids(pids):
    for pid in pids:
        try:
            subprocess.run(["taskkill", "/PID", str(pid), "/T", "/F"], capture_output=True, timeout=20, creationflags=NO_WINDOW)
        except Exception:
            pass


def run_word(kind, src, out, timeout=WORD_TIMEOUT):
    """Has Word convert one file in a separate process, which is shut down if it hangs or Stop is pressed."""
    name = os.path.basename(src)
    before = word_pids()
    err = out + ".err"
    proc = subprocess.Popen(self_command() + ["--word", kind, src, out], creationflags=NO_WINDOW)
    started = time.time()
    while proc.poll() is None:
        if CANCEL.is_set() or time.time() - started > timeout:
            proc.kill()
            kill_pids(word_pids() - before)
            if CANCEL.is_set():
                raise Cancelled()
            raise ConversionError(f"Word stopped responding while opening {name} and was closed after {timeout // 60} minutes. "
                                  "Open the file in Word, close any message it shows, save it, and check it again. "
                                  "Saving it as PDF from Word and checking the PDF also works.")
        time.sleep(0.5)
    if proc.returncode != 0 or not os.path.exists(out):
        detail = ""
        if os.path.exists(err):
            with open(err, encoding="utf-8", errors="replace") as fh:
                lines = fh.read().strip().splitlines()
            detail = lines[-1][:200] if lines else ""
        raise ConversionError(f"Word could not open {name}{f' ({detail})' if detail else ''}. Save it as PDF from Word and check the PDF.")


def word_available():
    try:
        import win32com.client  # noqa: F401
        import winreg
        winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Word.Application")
        return True
    except Exception:
        return False


def find_soffice():
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    for guess in (r"C:\Program Files\LibreOffice\program\soffice.exe", r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"):
        if not soffice and os.path.exists(guess):
            soffice = guess
    return soffice


def docx_to_pdf(path, outdir):
    """Lay the Word file out as a PDF. Microsoft Word gives the same pages the student sees.
    LibreOffice is the fallback and can move page breaks or swap fonts."""
    out = os.path.join(outdir, os.path.splitext(os.path.basename(path))[0] + ".pdf")
    if os.name == "nt" and word_available():
        run_word("pdf", path, out)
        return out, "Microsoft Word"
    soffice = find_soffice()
    if soffice:
        # LibreOffice's own python can refresh the table of contents before saving
        lo_py = os.path.join(os.path.dirname(soffice), "python.exe" if os.name == "nt" else "python")
        candidates = [lo_py] if os.path.exists(lo_py) else []
        candidates.append(sys.executable)
        for py in candidates:
            try:
                subprocess.run([py, os.path.join(BASE, "lo_convert.py"), soffice, path, out], check=True, timeout=600,
                               stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                if os.path.exists(out):
                    return out, "LibreOffice"
            except Exception:
                continue
        subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", outdir, path],
                       check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=300)
        if os.path.exists(out):
            return out, "LibreOffice"
    raise ConversionError(f"{os.path.basename(path)} could not be opened since Microsoft Word was not found. Save it as PDF from Word and check the PDF.")


# ---------------------------------------------------------------- main

def measure_ink(pdf_path, pages, rules):
    try:
        import pypdfium2 as pdfium
        from PIL import ImageOps
    except Exception:
        return
    doc = pdfium.PdfDocument(pdf_path)
    zone = rules["page_number"].get("zone_bottom_cm", 2.95) + 0.2
    for p in pages:
        img = doc[p.no - 1].render(scale=72 / 72, grayscale=True).to_pil()
        px_per_cm = img.width / p.width
        body = img.crop((0, int(zone * px_per_cm), img.width, img.height))
        box = ImageOps.invert(body).point(lambda v: 255 if v > 60 else 0).getbbox()
        if box:
            p.ink = (box[0] / px_per_cm, box[2] / px_per_cm)


def run(path, rules_path=DEFAULT_RULES, outdir=None):
    with open(rules_path, encoding="utf-8") as fh:
        rules = json.load(fh)
    global FRONT, BACK
    for key in ("order", "required", "front_headings"):
        if key in rules:
            rules[key] = [h.upper().replace("ACKNOWLEDGMENTS", "ACKNOWLEDGEMENTS") for h in rules[key]]
    FRONT = list(dict.fromkeys(DEFAULT_FRONT + [h.upper() for h in rules.get("front_headings", [])]))
    BACK = list(dict.fromkeys(DEFAULT_BACK + [h.upper() for h in rules.get("back_headings", [])]))
    src_name = os.path.basename(path)
    outdir = outdir or os.path.dirname(os.path.abspath(path))
    os.makedirs(outdir, exist_ok=True)
    tmp = None
    pdf_path = path
    F = Findings()
    if path.lower().endswith((".docx", ".doc")):
        tmp = tempfile.mkdtemp()
        pdf_path, converter = docx_to_pdf(path, tmp)
        if converter != "Microsoft Word":
            F.add("check", "Word file", f"{src_name} was laid out by {converter} because Microsoft Word was not found. {converter} can move page breaks "
                  "and swap fonts, so page numbers and fonts in this report may differ from Word. Check the PDF the student saved from Word instead.",
                  None, "Conversion")
    try:
        with pdfplumber.open(pdf_path) as pdf:
            pages = [Page(i, pg, rules) for i, pg in enumerate(pdf.pages, 1)]
    except Exception as e:
        detail = (type(e).__name__ + repr(e) + repr(getattr(e, "args", ""))).lower()
        if "password" in detail or "encrypt" in detail:
            raise ConversionError(f"{src_name} is password protected. Ask for a copy without a password.")
        raise ConversionError(f"{src_name} could not be opened as a PDF ({type(e).__name__}).")
    if sum(len(p.chars) for p in pages) < 40 * max(1, len(pages)):
        raise ConversionError(f"{src_name} has almost no text in it, so it is probably scanned. Ask for the PDF saved from Word.")
    measure_ink(pdf_path, pages, rules)
    headings = classify(pages, rules, F)
    check_page_size(pages, rules, F)
    check_fonts(pages, rules, F)
    check_body_size(pages, rules, F)
    check_margins(pages, rules, F)
    check_page_numbers(pages, rules, F)
    check_structure(pages, headings, rules, F)
    secs = check_sections(pages, rules, F)
    check_spacing_and_indent(pages, rules, F)
    check_placeholders(pages, rules, F)
    check_titles(pages, rules, F)
    check_abstract_fields(pages, rules, F)
    check_lists(pages, rules, F)
    check_toc(pages, headings, secs, rules, F)
    check_back(pages, rules, F)

    base = os.path.splitext(src_name)[0]
    report = os.path.join(outdir, f"{base} - format report.html")
    marked = os.path.join(outdir, f"{base} - corrections.pdf")
    write_report(report, src_name, pages, F, rules)
    write_corrections_pdf(pdf_path, marked, src_name, pages, F, rules)
    if tmp:
        shutil.rmtree(tmp, ignore_errors=True)
    return report, marked, F, pages, headings


def pick_rules(given=None):
    """The rules file to use: the one given, the only one in rules/, or the one picked from a list."""
    if given:
        return given
    files = rules_files()
    if len(files) <= 1 or not sys.stdin.isatty():
        return DEFAULT_RULES if DEFAULT_RULES in files or not files else files[0]
    names = []
    for f in files:
        try:
            names.append(json.load(open(f, encoding="utf-8")).get("name", os.path.basename(f)))
        except Exception:
            names.append(f)
    print("Which format?")
    for i, n in enumerate(names, 1):
        print(f"  {i}. {n}")
    while True:
        a = input("Type the number and press Enter: ").strip()
        if a.isdigit() and 1 <= int(a) <= len(files):
            return files[int(a) - 1]


def main():
    ap = argparse.ArgumentParser(description="Check a thesis against the school template.")
    ap.add_argument("file", help="thesis as .pdf or .docx")
    ap.add_argument("--rules", default=None)
    ap.add_argument("--out", default=None, help="folder for the report (default: next to the thesis)")
    a = ap.parse_args()
    try:
        report, marked, F, pages, _ = run(a.file, pick_rules(a.rules), a.out)
    except ConversionError as e:
        sys.exit(str(e))
    n_fix = sum(1 for f in F.items if f["level"] == "fix")
    print(f"{n_fix} to fix, {len(F.items) - n_fix} to check.")
    print(report)
    print(marked)


if __name__ == "__main__":
    main()
