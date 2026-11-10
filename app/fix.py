"""Manuscript auto-fix.

Makes a corrected copy of a thesis Word file and leaves the original untouched:
  <name> - fixed.docx         the corrected copy
  <name> - fixed changes.txt  what was changed
then checks the corrected copy and writes its report and corrections sheets.

Usage:
  python fix.py "My Thesis.docx"
"""
import argparse
import copy
import json
import os
import re
import sys
import zipfile
from collections import Counter

from lxml import etree

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import check  # noqa: E402
DEFAULT_RULES = check.DEFAULT_RULES

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
W = "{%s}" % W_NS


def whole(v, default=0):
    """A whole number from a Word setting. Files converted from PDF write some as 400.0."""
    if v is None:
        return default
    try:
        return int(round(float(v)))
    except ValueError:
        return default
R = "{%s}" % R_NS
NS = {"w": W_NS, "r": R_NS, "wp": WP_NS, "a": A_NS}

TWIP_PER_CM = 1440 / 2.54
EMU_PER_TWIP = 635
TNR = "Times New Roman"
KEEP_FONTS = ("symbol", "wingdings", "cambria math", "courier new", "segoe ui symbol", "mt extra")

FRONT = ["ABSTRACT", "ACKNOWLEDGEMENTS", "ACKNOWLEDGMENTS", "TABLE OF CONTENTS", "LIST OF TABLES", "LIST OF FIGURES", "LIST OF ABBREVIATIONS"]
BACK = ["REFERENCES", "BIBLIOGRAPHY", "APPENDIX", "APPENDICES", "BIOGRAPHY"]


def tw(cm):
    return int(round(cm * TWIP_PER_CM))


def el(tag, **attrs):
    e = etree.Element(W + tag)
    for k, v in attrs.items():
        e.set(W + k, str(v))
    return e


def ptext(p):
    out = []
    for e in p.iter(W + "t", W + "br"):
        out.append(" " if e.tag == W + "br" else (e.text or ""))
    return "".join(out)


def get_ppr(p):
    ppr = p.find(W + "pPr")
    if ppr is None:
        ppr = el("pPr")
        p.insert(0, ppr)
    return ppr


# children of pPr must keep schema order or Word refuses the file
PPR_ORDER = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr", "suppressLineNumbers",
             "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE",
             "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
             "suppressOverlap", "jc", "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle",
             "rPr", "sectPr", "pPrChange"]
RPR_ORDER = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow", "emboss",
             "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs",
             "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout",
             "specVanish", "oMath"]


def set_child(parent, tag, order, **attrs):
    """Replace or insert parent/<tag> keeping schema order."""
    old = parent.find(W + tag)
    new = el(tag, **attrs)
    if old is not None:
        parent.replace(old, new)
        return new
    idx = order.index(tag)
    pos = len(parent)
    for i, ch in enumerate(parent):
        name = etree.QName(ch).localname
        if name in order and order.index(name) > idx:
            pos = i
            break
    parent.insert(pos, new)
    return new


def get_rpr(r):
    rpr = r.find(W + "rPr")
    if rpr is None:
        rpr = el("rPr")
        r.insert(0, rpr)
    return rpr


class Docx:
    def __init__(self, path):
        self.zin = zipfile.ZipFile(path)
        self.parts = {}
        for name in ("word/document.xml", "word/styles.xml", "word/numbering.xml", "word/settings.xml", "word/_rels/document.xml.rels"):
            if name in self.zin.namelist():
                self.parts[name] = etree.fromstring(self.zin.read(name))
        self.doc = self.parts["word/document.xml"]
        self.body = self.doc.find(W + "body")
        self.styles = self.parts.get("word/styles.xml")
        self.numbering = self.parts.get("word/numbering.xml")
        rels = self.parts["word/_rels/document.xml.rels"]
        self.rels = {r.get("Id"): r.get("Target") for r in rels}
        self.extra = {}  # header parts we rewrite

    def part(self, name):
        if name not in self.parts:
            self.parts[name] = etree.fromstring(self.zin.read(name))
        return self.parts[name]

    def save(self, path):
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in self.zin.infolist():
                if item.filename in self.parts:
                    data = etree.tostring(self.parts[item.filename], xml_declaration=True, encoding="UTF-8", standalone=True)
                    zout.writestr(item, data)
                else:
                    zout.writestr(item, self.zin.read(item.filename))

    # ---- style and numbering lookups
    def style(self, sid):
        if self.styles is None or not sid:
            return None
        for s in self.styles.findall(W + "style"):
            if s.get(W + "styleId") == sid:
                return s
        return None

    def style_ind(self, sid, depth=0):
        s = self.style(sid)
        if s is None or depth > 8:
            return {}
        ind = s.find(f"{W}pPr/{W}ind")
        base = s.find(W + "basedOn")
        out = self.style_ind(base.get(W + "val"), depth + 1) if base is not None else {}
        if ind is not None:
            out.update({etree.QName(k).localname: whole(v) for k, v in ind.attrib.items() if re.fullmatch(r"-?\d+(\.\d+)?", v)})
        return out

    def num_ind(self, num_id, ilvl):
        if self.numbering is None:
            return {}
        absid = None
        lvl_override = None
        for n in self.numbering.findall(W + "num"):
            if n.get(W + "numId") == num_id:
                a = n.find(W + "abstractNumId")
                absid = a.get(W + "val") if a is not None else None
                for o in n.findall(W + "lvlOverride"):
                    if o.get(W + "ilvl") == ilvl and o.find(W + "lvl") is not None:
                        lvl_override = o.find(W + "lvl")
        lvl = lvl_override
        if lvl is None:
            for a in self.numbering.findall(W + "abstractNum"):
                if a.get(W + "abstractNumId") == absid:
                    for l in a.findall(W + "lvl"):
                        if l.get(W + "ilvl") == ilvl:
                            lvl = l
        if lvl is None:
            return {}
        ind = lvl.find(f"{W}pPr/{W}ind")
        if ind is None:
            return {}
        return {etree.QName(k).localname: whole(v) for k, v in ind.attrib.items() if re.fullmatch(r"-?\d+(\.\d+)?", v)}

    def eff_ind(self, p):
        ppr = p.find(W + "pPr")
        sid = ppr.find(W + "pStyle").get(W + "val") if ppr is not None and ppr.find(W + "pStyle") is not None else "Normal"
        out = dict(self.style_ind(sid))
        num = numpr(p)
        if num:
            out.update(self.num_ind(*num))
        ind = ppr.find(W + "ind") if ppr is not None else None
        if ind is not None:
            out.update({etree.QName(k).localname: whole(v) for k, v in ind.attrib.items() if re.fullmatch(r"-?\d+(\.\d+)?", v)})
        out.setdefault("left", out.get("start", 0))
        out.setdefault("right", out.get("end", 0))
        return out


def numpr(p):
    n = p.find(f"{W}pPr/{W}numPr")
    if n is None or n.find(W + "numId") is None:
        return None
    ilvl = n.find(W + "ilvl")
    return n.find(W + "numId").get(W + "val"), (ilvl.get(W + "val") if ilvl is not None else "0")


def style_of(p):
    s = p.find(f"{W}pPr/{W}pStyle")
    return s.get(W + "val") if s is not None else ""


def in_table(e):
    return any(etree.QName(a).localname == "tbl" for a in e.iterancestors())


def in_textbox(e):
    return any(etree.QName(a).localname in ("txbxContent", "drawing", "pict") for a in e.iterancestors())


def write_ind(p, ind):
    ppr = get_ppr(p)
    attrs = {"left": max(0, ind.get("left", 0))}
    if ind.get("right"):
        attrs["right"] = ind["right"]
    if ind.get("hanging"):
        attrs["hanging"] = ind["hanging"]
    elif ind.get("firstLine"):
        attrs["firstLine"] = ind["firstLine"]
    else:
        attrs["firstLine"] = 0
    set_child(ppr, "ind", PPR_ORDER, **attrs)


# ---------------------------------------------------------------- the fixes

class Fixer:
    def __init__(self, path, rules):
        self.d = Docx(path)
        self.rules = rules
        self.log = []
        m = rules["margins"]
        self.left_tw = tw(m["left_text_cm"])
        self.right_tw = tw(m["right_cm"])
        self.bottom_tw = tw(m["bottom_cm"])
        self.indent_tw = tw(rules["paragraph_indent"]["expected_cm"])
        sz = rules["sizes"]
        self.body_hp = int(round(sz["body_pt"] * 2))        # Word sizes are in half points
        self.head_hp = int(round(sz["heading_pt"] * 2))
        ratio = rules["line_spacing"]["expected_cm"] / (sz["body_pt"] * 1.15 * 2.54 / 72)
        self.line = str(int(round(ratio * 4) / 4 * 240))     # 240 = single, 360 = 1.5, 480 = double
        allowed = rules["fonts"].get("name") or rules["fonts"]["allowed"][0]
        self.font = {"timesnewroman": "Times New Roman", "arialmt": "Arial", "arial": "Arial", "liberationsans": "Arial",
                     "liberationserif": "Times New Roman", "calibri": "Calibri", "cambria": "Cambria"}.get(allowed.replace(" ", "").lower(), allowed)
        self.order = rules.get("order", [])

    def note(self, msg):
        self.log.append(msg)

    def body_items(self):
        return list(self.d.body)

    def sections(self):
        """[(sectPr, [body elements in that section])]"""
        out = []
        cur = []
        for e in self.body_items():
            if etree.QName(e).localname == "sectPr":
                out.append((e, cur))
                cur = []
                continue
            cur.append(e)
            sp = e.find(f"{W}pPr/{W}sectPr") if etree.QName(e).localname == "p" else None
            if sp is not None:
                out.append((sp, cur))
                cur = []
        return out

    def front_range(self):
        secs = [sp for sp, _ in self.sections()]

        def fmt(sp):
            pn = sp.find(W + "pgNumType")
            return pn.get(W + "fmt") if pn is not None else None

        def start(sp):
            pn = sp.find(W + "pgNumType")
            return pn.get(W + "start") if pn is not None else None

        first_front = next((i for i, sp in enumerate(secs) if fmt(sp) in ("lowerRoman", "upperRoman")), None)
        first_body = next((i for i, sp in enumerate(secs) if first_front is not None and i > first_front
                           and fmt(sp) not in ("lowerRoman", "upperRoman") and start(sp) is not None), None)
        return first_front, first_body

    def heading_kind(self, p):
        t = re.sub(r"\s+", " ", ptext(p)).strip().upper()
        if not t or is_contents_line(p):
            return None
        if t in FRONT or t in BACK or re.fullmatch(r"APPENDIX [A-Z]", t):
            return t
        if re.match(r"^CHAPTER\s+\d+", t):
            return "CHAPTER"
        if style_of(p) in ("Heading1", "Heading2", "FrontHeading") and t.isupper():
            return "TITLE"
        return None

    # 1. back matter that is labelled as a chapter, and references before the appendix
    def fix_back_matter(self):
        items = self.body_items()
        removed = []
        for i, e in enumerate(items):
            if etree.QName(e).localname != "p":
                continue
            t = re.sub(r"\s+", " ", ptext(e)).strip().upper()
            m = re.fullmatch(r"CHAPTER (\d+)", t)
            if not m:
                continue
            nxt = next((x for x in items[i + 1:] if etree.QName(x).localname == "p" and ptext(x).strip()), None)
            nt = re.sub(r"\s+", " ", ptext(nxt)).strip().upper() if nxt is not None else ""
            if any(nt.startswith(b) for b in BACK):
                e.getparent().remove(e)
                removed.append(f"CHAPTER {m.group(1)} above {nt}")
        if removed:
            self.note("Removed the chapter label from " + ", ".join(removed) + ".")

        items = self.body_items()

        def find(name):
            for i, e in enumerate(items):
                if etree.QName(e).localname == "p" and re.sub(r"\s+", " ", ptext(e)).strip().upper() == name:
                    return i
            return None

        ref, app, bio = find("REFERENCES"), find("APPENDIX"), find("BIOGRAPHY")
        if app is None:
            app = find("APPENDICES")
        order = [h.upper() for h in self.order]
        refs_first = not order or ("REFERENCES" in order and "APPENDIX" in order and order.index("REFERENCES") < order.index("APPENDIX"))
        if ref is not None and app is not None and ref > app and refs_first:
            end = bio if bio is not None and bio > ref else len(items) - 1
            # stop before page-break paragraphs that lead into the next heading
            while end - 1 > ref and is_break_only(items[end - 1]):
                end -= 1
            block = items[ref:end]
            # a page-break paragraph left behind just before REFERENCES would make a blank page
            prev = items[ref - 1]
            if is_break_only(prev):
                prev.getparent().remove(prev)
            anchor = items[app]
            # move the page-break paragraphs that sat before the appendix heading along with it
            k = app
            while k - 1 >= 0 and is_break_only(items[k - 1]):
                k -= 1
            anchor = items[k]
            for e in block:
                anchor.addprevious(e)
            set_child(get_ppr(items[app]), "pageBreakBefore", PPR_ORDER)
            self.note("Moved REFERENCES ahead of the appendix.")

        # appendix: APPENDICES heading, then APPENDIX A, B, C from its numbered sections
        items = self.body_items()
        app = find_idx(items, "APPENDIX")
        if app is not None:
            head = items[app]
            replace_text(head, "APPENDICES")
            sub = []
            for e in items[app + 1:]:
                if etree.QName(e).localname != "p":
                    continue
                t = re.sub(r"\s+", " ", ptext(e)).strip().upper()
                if t in BACK or t.startswith("CHAPTER"):
                    break
                if numpr(e) and style_of(e) == "Heading3" and ptext(e).strip():
                    sub.append(e)
            for k, e in enumerate(sub):
                letter = chr(ord("A") + k)
                title = ptext(e).strip().upper()
                ppr = get_ppr(e)
                n = ppr.find(W + "numPr")
                if n is not None:
                    ppr.remove(n)
                set_child(ppr, "pStyle", PPR_ORDER, val="Heading2")
                set_child(ppr, "pageBreakBefore", PPR_ORDER)
                for r in e.findall(W + "r"):
                    e.remove(r)
                r = el("r")
                t1 = el("t")
                t1.text = f"APPENDIX {letter}"
                r.append(t1)
                r.append(el("br"))
                t2 = el("t")
                t2.text = title
                r.append(t2)
                e.append(r)
            # the empty numbered paragraphs that only carried the chapter number
            for e in items[app + 1:]:
                if etree.QName(e).localname == "p" and numpr(e) and not ptext(e).strip() and style_of(e) == "ListParagraph":
                    e.getparent().remove(e)
            if sub:
                self.note(f"Renamed the appendix heading to APPENDICES and its sections to APPENDIX A to APPENDIX {chr(ord('A') + len(sub) - 1)}, each on a new page.")

    # 2a. CHAPTER n on its own line with the title below, starting a new page
    def fix_chapter_headings(self):
        items = self.body_items()
        n = 0
        for i, p in enumerate(items):
            if etree.QName(p).localname != "p":
                continue
            t = re.sub(r"\s+", " ", ptext(p)).strip()
            m = re.match(r"^CHAPTER\s*(\d+)\s*(.*)$", t, re.I)
            if not m or is_contents_line(p) or (not style_of(p).startswith("Heading") and not t.isupper()):
                continue
            num, title = m.group(1), m.group(2).strip()
            if True:
                lines = [f"CHAPTER {num}"]
                if not title:
                    nxt = next((x for x in items[i + 1:] if etree.QName(x).localname == "p" and ptext(x).strip()), None)
                    if nxt is not None and ptext(nxt).strip().isupper() and style_of(nxt).startswith("Heading"):
                        title = re.sub(r"\s+", " ", ptext(nxt)).strip()
                        for bm in nxt.findall(W + "bookmarkStart") + nxt.findall(W + "bookmarkEnd"):
                            p.append(bm)
                        nxt.getparent().remove(nxt)
                if title:
                    lines.append(title.upper())
            if lines:
                set_lines(p, lines)
            ppr = get_ppr(p)
            set_child(ppr, "pStyle", PPR_ORDER, val="Heading2")
            prev = p.getprevious()
            starts_page = prev is None or is_break_only(prev) or prev.find(f"{W}pPr/{W}sectPr") is not None
            if not starts_page:
                set_child(ppr, "pageBreakBefore", PPR_ORDER)
            n += 1
        if n:
            self.note(f"Split {n} chapter headings into two lines, such as CHAPTER 1 and then INTRODUCTION, each starting a new page.")

    # 2b. front headings out of the outline, lists of tables and figures built from the captions
    def fix_contents(self):
        self.ensure_style("FrontHeading", "Front Heading", "Heading2", outline=9)
        self.ensure_style("TableCaption", "Table Caption", "BodyText")
        self.ensure_style("FigureCaption", "Figure Caption", "BodyText")
        text_w = self.text_width()
        # captions
        caps = {"Table": 0, "Figure": 0}
        for p in self.d.body.iter(W + "p"):
            if in_table(p) or in_textbox(p):
                continue
            t = ptext(p).strip()
            m = re.match(r"^(Table|Figure)\s+\d+(?:\.\d+)*\.?\s", t)
            if m and len(t) < 250:
                set_child(get_ppr(p), "pStyle", PPR_ORDER, val=m.group(1) + "Caption")
                caps[m.group(1)] += 1
        # lists of tables and figures held in a table: replace with a heading and a list Word fills in
        for tbl in list(self.d.body.findall(W + "tbl")):
            first = tbl.find(W + "tr")
            head = re.sub(r"\s+", " ", ptext(first) if first is not None else "").strip().upper()
            for name, word, style in (("LIST OF TABLES", "Tables", "Table Caption"), ("LIST OF FIGURES", "Figures", "Figure Caption")):
                if head == name:
                    h = el("p")
                    h.append(el("pPr"))
                    set_lines(h, [name])
                    sub = el("p")
                    sppr = el("pPr")
                    tabs = el("tabs")
                    tabs.append(el("tab", val="right", pos=text_w))
                    sppr.append(tabs)
                    sppr.append(el("spacing", before=240, line=self.line, lineRule="auto"))
                    sub.append(sppr)
                    sub.append(run_text(word))
                    r = el("r")
                    r.append(el("tab"))
                    sub.append(r)
                    sub.append(run_text("Page"))
                    lst = field_par(f' TOC \\h \\z \\t "{style},1" ', "Right-click and choose Update Field to fill in this list.")
                    tbl.addprevious(h)
                    tbl.addprevious(sub)
                    tbl.addprevious(lst)
                    tbl.getparent().remove(tbl)
                    self.note(f"Rebuilt the {name} from the {caps[word[:-1]]} {word.lower()[:-1]} captions. It lists each title now and Word keeps the page numbers up to date.")
        # front headings: bookmarked, kept out of the table of contents field, listed with (n) page numbers instead
        front = []
        for p in self.body_items():
            if etree.QName(p).localname != "p":
                continue
            t = re.sub(r"\s+", " ", ptext(p)).strip().upper()
            if t in FRONT:
                name = f"_Front{len(front) + 1}"
                set_child(get_ppr(p), "pStyle", PPR_ORDER, val="FrontHeading")
                bs = el("bookmarkStart", id=900 + len(front), name=name)
                be = el("bookmarkEnd", id=900 + len(front))
                ppr = p.find(W + "pPr")
                ppr.addnext(bs)
                p.append(be)
                front.append((t.replace("ACKNOWLEDGMENTS", "ACKNOWLEDGEMENTS"), name))
        toc = None
        for sdt in self.d.body.iter(W + "sdt"):
            if any("TOC" in (x.text or "") for x in sdt.iter(W + "instrText")):
                toc = sdt
                break
        if toc is not None and front and self.rules["page_number"].get("front_style") == "paren":
            page = el("p")
            pp = el("pPr")
            pp.append(el("jc", val="right"))
            page.append(pp)
            page.append(run_text("Page"))
            toc.addprevious(page)
            for title, bm in front:
                line = el("p")
                lp = el("pPr")
                lp.append(el("pStyle", val="TOC2"))
                tabs = el("tabs")
                tabs.append(el("tab", val="right", leader="dot", pos=text_w))
                lp.append(tabs)
                line.append(lp)
                line.append(run_text(title))
                r = el("r")
                r.append(el("tab"))
                line.append(r)
                line.append(run_text("("))
                for part in field_runs(f" PAGEREF {bm} \\h ", "1"):
                    line.append(part)
                line.append(run_text(")"))
                toc.addprevious(line)
            self.note("Listed the front pages in the table of contents as (1), (2), (3), and added the Page label above the numbers.")
        # table of contents styles: entries at the margin, page numbers at the right margin
        for sid, left in (("TOC1", 0), ("TOC2", 0), ("TOC3", 720)):
            st = self.d.style(sid)
            if st is None:
                continue
            ppr = st.find(W + "pPr")
            if ppr is None:
                ppr = el("pPr")
                st.insert(0, ppr)
            tabs = el("tabs")
            tabs.append(el("tab", val="right", leader="dot", pos=text_w))
            old = ppr.find(W + "tabs")
            if old is not None:
                ppr.remove(old)
            set_child(ppr, "tabs", PPR_ORDER)
            ppr.replace(ppr.find(W + "tabs"), tabs)
            set_child(ppr, "spacing", PPR_ORDER, before=240 if sid == "TOC2" else 0, line=self.line, lineRule="auto")
            set_child(ppr, "ind", PPR_ORDER, left=left, right=0)
            set_child(ppr, "jc", PPR_ORDER, val="left")

    def text_width(self):
        for sp, _ in self.sections():
            pgsz = sp.find(W + "pgSz")
            if pgsz is not None and whole(pgsz.get(W + "w", 0)) < whole(pgsz.get(W + "h", 1)):
                return whole(pgsz.get(W + "w")) - self.left_tw - self.right_tw
        return 8470

    def ensure_style(self, sid, name, based, outline=None):
        if self.d.style(sid) is not None:
            return
        st = el("style", type="paragraph", styleId=sid)
        st.append(el("name", val=name))
        st.append(el("basedOn", val=based))
        st.append(el("qFormat"))
        if outline is not None:
            ppr = el("pPr")
            ppr.append(el("outlineLvl", val=outline))
            st.append(ppr)
        self.d.styles.append(st)

    # 2. headings: 14 pt bold capitals, centred, no indent
    def fix_headings(self):
        n = 0
        for p in self.d.body.iter(W + "p"):
            if in_textbox(p):
                continue
            kind = self.heading_kind(p)
            if not kind:
                continue
            if in_table(p) and kind not in ("LIST OF TABLES", "LIST OF FIGURES", "LIST OF ABBREVIATIONS"):
                continue
            ppr = get_ppr(p)
            set_child(ppr, "ind", PPR_ORDER, left=0, right=0, firstLine=0)
            set_child(ppr, "jc", PPR_ORDER, val="center")
            set_child(ppr, "keepNext", PPR_ORDER)
            for r in p.iter(W + "r"):
                rpr = get_rpr(r)
                set_child(rpr, "b", RPR_ORDER)
                set_child(rpr, "sz", RPR_ORDER, val=self.head_hp)
                set_child(rpr, "szCs", RPR_ORDER, val=self.head_hp)
                for t in r.findall(W + "t"):
                    if t.text and t.text != t.text.upper():
                        t.text = t.text.upper()
            n += 1
        for sid in ("Heading1", "Heading2"):
            s = self.d.style(sid)
            if s is None:
                continue
            ppr = s.find(W + "pPr")
            if ppr is not None and ppr.find(W + "ind") is not None:
                ppr.remove(ppr.find(W + "ind"))
            rpr = s.find(W + "rPr")
            if rpr is not None:
                set_child(rpr, "sz", RPR_ORDER, val=self.head_hp)
                set_child(rpr, "szCs", RPR_ORDER, val=self.head_hp)
        self.note(f"Set {n} main headings, such as ABSTRACT, the chapter headings and REFERENCES, in {self.head_hp / 2:g} pt bold capitals, centred.")

    # 3. numbered section headings (2.1, 2.2) start at the margin
    def fix_section_heading_indents(self):
        changed = 0
        for p in self.d.body.iter(W + "p"):
            num = numpr(p)
            if not num or style_of(p) not in ("Heading3", "Heading2") or not ptext(p).strip() or in_table(p):
                continue
            ind = self.d.eff_ind(p)
            left = ind.get("left", 0)
            hang = ind.get("hanging", 0)
            numpos = left - hang if hang else left + ind.get("firstLine", 0)
            if numpos != 0:
                gap = hang if hang else 720
                ind = {"left": gap, "hanging": gap}
                write_ind(p, ind)
                changed += 1
        if changed:
            self.note(f"Moved {changed} section headings so the number starts at the margin.")

    # 3b. each chapter's section numbers come from its own list, starting at n.1
    def fix_section_numbering(self):
        if self.d.numbering is None:
            return
        chapters = []
        cur = None
        for p in self.d.body.iter(W + "p"):
            t = re.sub(r"\s+", " ", ptext(p)).strip().upper()
            m = re.match(r"^CHAPTER (\d+)", t)
            if m and style_of(p).startswith("Heading"):
                cur = (int(m.group(1)), [])
                chapters.append(cur)
                continue
            if t in BACK or t == "APPENDICES":
                cur = None
            if cur is not None and numpr(p):
                cur[1].append(p)
        seen = {}
        fixed = []
        nums = self.d.numbering
        next_id = max([whole(n.get(W + "numId")) for n in nums.findall(W + "num")] + [0]) + 1
        for chap, ps in chapters:
            heads = [p for p in ps if style_of(p) == "Heading3"]
            ids = {numpr(p)[0] for p in heads}
            for nid in ids:
                if nid in seen and seen[nid] != chap:
                    src = next(n for n in nums.findall(W + "num") if n.get(W + "numId") == nid)
                    abs_id = src.find(W + "abstractNumId").get(W + "val")
                    absn = next(a for a in nums.findall(W + "abstractNum") if a.get(W + "abstractNumId") == abs_id)
                    # a copy of the list definition with its own identity, so Word and LibreOffice both restart it
                    copy_abs = copy.deepcopy(absn)
                    new_abs = max(whole(a.get(W + "abstractNumId")) for a in nums.findall(W + "abstractNum")) + 1
                    copy_abs.set(W + "abstractNumId", str(new_abs))
                    nsid = copy_abs.find(W + "nsid")
                    if nsid is not None:
                        nsid.set(W + "val", f"{(0x5A000000 + new_abs * 7919) & 0xFFFFFFFF:08X}")
                    for lvl in copy_abs.findall(W + "lvl"):
                        st = lvl.find(W + "start")
                        if st is not None:
                            st.set(W + "val", str(chap if lvl.get(W + "ilvl") == "0" else 1))
                    nums.findall(W + "abstractNum")[-1].addnext(copy_abs)
                    new = el("num", numId=next_id)
                    new.append(el("abstractNumId", val=new_abs))
                    nums.append(new)
                    for p in ps:
                        if numpr(p)[0] == nid:
                            p.find(f"{W}pPr/{W}numPr/{W}numId").set(W + "val", str(next_id))
                    fixed.append(chap)
                    next_id += 1
                else:
                    seen.setdefault(nid, chap)
        if fixed:
            self.note("Gave " + ", ".join(f"Chapter {c}" for c in fixed) + " its own section numbering so its headings run "
                      + ", ".join(f"{c}.1, {c}.2" for c in fixed) + " instead of carrying on from the chapter before.")

    # 4. margins, keeping every paragraph where it sits relative to the text
    def fix_margins(self):
        pg_w = None
        changed_sections = 0
        moved = 0
        first_front, _ = self.front_range()
        for si, (sp, elems) in enumerate(self.sections()):
            pgsz = sp.find(W + "pgSz")
            pgmar = sp.find(W + "pgMar")
            if pgmar is None or pgsz is None:
                continue
            landscape = pgsz.get(W + "orient") == "landscape" or whole(pgsz.get(W + "w", 0)) > whole(pgsz.get(W + "h", 0))
            if landscape:
                continue
            pg_w = whole(pgsz.get(W + "w"))
            old_left = whole(pgmar.get(W + "left", 0))
            delta = self.left_tw - old_left
            touched = False
            if old_left != self.left_tw:
                pgmar.set(W + "left", str(self.left_tw))
                touched = True
            if whole(pgmar.get(W + "right", 0)) < self.right_tw:
                pgmar.set(W + "right", str(self.right_tw))
                touched = True
            title_pages = first_front is not None and si < first_front
            if whole(pgmar.get(W + "bottom", 0)) < self.bottom_tw and not title_pages:
                pgmar.set(W + "bottom", str(self.bottom_tw))
                touched = True
            changed_sections += touched
            text_w = pg_w - self.left_tw - whole(pgmar.get(W + "right"))
            for e in elems:
                kind = etree.QName(e).localname
                paras = [e] if kind == "p" else (list(e.iter(W + "p")) if kind == "sdt" else [])
                for p in paras:
                    if in_table(p) or in_textbox(p):
                        continue
                    if delta:
                        ind = self.d.eff_ind(p)
                        left = ind.get("left", 0)
                        hang = ind.get("hanging", 0)
                        first = ind.get("firstLine", 0)
                        start = min(left, left - hang, left + (first if first < 0 else 0))
                        new_start = max(0, start - delta)
                        if new_start != start + delta:
                            moved += 1
                        ind["left"] = left + (new_start - start)
                        write_ind(p, ind)
                    self.fit_images(p, text_w)
                    self.fit_tabs(p, text_w)
                if kind == "tbl":
                    self.fit_table(e, delta, text_w)
        if self.d.styles is not None and pg_w:
            text_w = pg_w - self.left_tw - self.right_tw
            for tab in self.d.styles.iter(W + "tab"):
                if tab.get(W + "pos") and whole(tab.get(W + "pos")) > text_w:
                    tab.set(W + "pos", str(text_w))
        self.note(f"Set the left margin to {self.rules['margins']['left_text_cm']} cm, the right margin to at least {self.rules['margins']['right_cm']} cm "
                  f"and the bottom margin to at least {self.rules['margins']['bottom_cm']} cm in {changed_sections} sections, and shifted "
                  f"{moved} indented paragraphs to match.")

    def fit_tabs(self, p, text_w):
        ind = self.d.eff_ind(p)
        limit = text_w - max(0, ind.get("left", 0)) - max(0, ind.get("right", 0)) + max(0, ind.get("left", 0))
        for tab in p.iter(W + "tab"):
            if tab.getparent().tag == W + "tabs" and tab.get(W + "pos") and whole(tab.get(W + "pos")) > limit:
                tab.set(W + "pos", str(limit))
                self.tabs_moved = getattr(self, "tabs_moved", 0) + 1

    def fit_images(self, p, text_w):
        ind = self.d.eff_ind(p)
        avail = (text_w - max(0, ind.get("left", 0)) - max(0, ind.get("right", 0))) * EMU_PER_TWIP
        for inl in list(p.iter("{%s}inline" % WP_NS)) + list(p.iter("{%s}anchor" % WP_NS)):
            ext = inl.find("{%s}extent" % WP_NS)
            cx, cy = whole(ext.get("cx")), whole(ext.get("cy"))
            if cx > avail:
                k = avail / cx
                ext.set("cx", str(int(cx * k)))
                ext.set("cy", str(int(cy * k)))
                for a in inl.iter("{%s}ext" % A_NS):
                    if a.get("cx"):
                        a.set("cx", str(int(whole(a.get("cx")) * k)))
                        a.set("cy", str(int(whole(a.get("cy")) * k)))
                self.images_scaled = getattr(self, "images_scaled", 0) + 1

    def fit_table(self, tbl, delta, text_w):
        tpr = tbl.find(W + "tblPr")
        grid = tbl.find(W + "tblGrid")
        if tpr is None or grid is None:
            return
        ind_el = tpr.find(W + "tblInd")
        ind = whole(ind_el.get(W + "w", 0)) if ind_el is not None and ind_el.get(W + "type", "dxa") == "dxa" else 0
        new_ind = max(0, ind - delta) if delta else ind
        if ind_el is not None:
            ind_el.set(W + "w", str(new_ind))
        cols = grid.findall(W + "gridCol")
        total = sum(whole(c.get(W + "w", 0)) for c in cols)
        avail = text_w - new_ind
        if total > avail and total:
            k = avail / total
            for c in cols:
                c.set(W + "w", str(int(whole(c.get(W + "w")) * k)))
            for tcw in tbl.iter(W + "tcW"):
                if tcw.get(W + "type", "dxa") == "dxa":
                    tcw.set(W + "w", str(int(whole(tcw.get(W + "w", 0)) * k)))
            tw_el = tpr.find(W + "tblW")
            if tw_el is not None and tw_el.get(W + "type") == "dxa":
                tw_el.set(W + "w", str(int(avail)))
            self.tables_scaled = getattr(self, "tables_scaled", 0) + 1

    # 5. body paragraphs that start at the margin get the indent the rest of the chapters use
    def fix_paragraph_indent(self):
        body = []
        in_chapter = False
        for e in self.body_items():
            if etree.QName(e).localname != "p":
                continue
            kind = self.heading_kind(e)
            if kind == "CHAPTER":
                in_chapter = True
            elif kind:
                in_chapter = kind not in BACK and kind != "APPENDICES" and in_chapter
            if not in_chapter or kind or numpr(e) or e.find(".//" + W + "drawing") is not None:
                continue
            t = ptext(e).strip()
            jc = e.find(f"{W}pPr/{W}jc")
            if len(t) < 150 or (jc is not None and jc.get(W + "val") in ("center", "right")) or re.match(r"^(Table|Figure)\s+\d", t):
                continue
            body.append(e)
        shapes = Counter()
        for e in body:
            ind = self.d.eff_ind(e)
            if ind.get("left", 0) >= 170 or ind.get("firstLine", 0) >= 170:
                shapes[(ind.get("left", 0), ind.get("firstLine", 0))] += 1
        left, first = shapes.most_common(1)[0][0] if shapes else (self.indent_tw, 0)
        n = 0
        for e in body:
            ind = self.d.eff_ind(e)
            if ind.get("left", 0) < 170 and ind.get("firstLine", 0) < 170 and not ind.get("hanging"):
                ind["left"], ind["firstLine"] = left, first
                write_ind(e, ind)
                n += 1
        if n:
            how = f"{left / TWIP_PER_CM:.2f} cm from the margin" + (f" with the first line {first / TWIP_PER_CM:.2f} cm further in" if first else "")
            self.note(f"Indented {n} body paragraphs {how} to match the rest of the chapters.")

    # 6. line spacing 1.5 from the abstract on
    def fix_line_spacing(self):
        started = False
        n = 0
        for e in self.body_items():
            names = [e] if etree.QName(e).localname == "p" else list(e.iter(W + "p")) if etree.QName(e).localname == "sdt" else []
            for p in names:
                if not started and re.sub(r"\s+", " ", ptext(p)).strip().upper() == "ABSTRACT":
                    started = True
                if not started or in_table(p) or in_textbox(p) or p.find(".//" + W + "drawing") is not None or is_break_only(p):
                    continue
                ppr = get_ppr(p)
                sp = ppr.find(W + "spacing")
                attrs = {etree.QName(k).localname: v for k, v in sp.attrib.items()} if sp is not None else {}
                if attrs.get("line") == self.line and attrs.get("lineRule", "auto") == "auto":
                    continue
                attrs["line"] = self.line
                attrs["lineRule"] = "auto"
                set_child(ppr, "spacing", PPR_ORDER, **attrs)
                n += 1
        name = {"240": "single", "276": "1.15", "360": "1.5", "480": "double"}.get(self.line, f"{int(self.line) / 240:g}")
        self.note(f"Set {name} line spacing on {n} paragraphs from the abstract to the end.")

    # 7. body text 12 pt Times New Roman
    def fix_text(self):
        started = False
        resized = 0
        fonts = 0
        abbrev = False
        for e in self.body_items():
            for p in ([e] if etree.QName(e).localname == "p" else list(e.iter(W + "p"))):
                t = re.sub(r"\s+", " ", ptext(p)).strip().upper()
                if t == "ABSTRACT":
                    started = True
                if t == "LIST OF ABBREVIATIONS":
                    abbrev = True
                elif self.heading_kind(p) and t != "LIST OF ABBREVIATIONS":
                    abbrev = False
                heading = bool(self.heading_kind(p))
                table = in_table(p)
                for r in p.iter(W + "r"):
                    rpr = r.find(W + "rPr")
                    # fonts everywhere, tables included
                    if rpr is not None and rpr.find(W + "rFonts") is not None:
                        f = rpr.find(W + "rFonts")
                        name = (f.get(W + "ascii") or f.get(W + "hAnsi") or "").lower()
                        if not any(k in name for k in KEEP_FONTS) and (name not in ("", self.font.lower()) or f.get(W + "asciiTheme") or f.get(W + "hAnsiTheme")):
                            for a in ("asciiTheme", "hAnsiTheme"):
                                if f.get(W + a):
                                    del f.attrib[W + a]
                            f.set(W + "ascii", self.font)
                            f.set(W + "hAnsi", self.font)
                            fonts += 1
                    if not started or heading or in_textbox(p) or (table and not abbrev):
                        continue
                    if rpr is not None and rpr.find(W + "vertAlign") is not None:
                        continue
                    sz = rpr.find(W + "sz") if rpr is not None else None
                    val = whole(sz.get(W + "val")) if sz is not None else None
                    if val is None or (self.body_hp - 6 <= val <= self.body_hp + 2 and val != self.body_hp):
                        rpr = get_rpr(r)
                        set_child(rpr, "sz", RPR_ORDER, val=self.body_hp)
                        set_child(rpr, "szCs", RPR_ORDER, val=self.body_hp)
                        if val is not None:
                            resized += 1
        if self.d.styles is not None:
            for f in self.d.styles.iter(W + "rFonts"):
                name = (f.get(W + "ascii") or "").lower()
                if any(k in name for k in KEEP_FONTS):
                    continue
                for a in ("asciiTheme", "hAnsiTheme"):
                    if f.get(W + a):
                        del f.attrib[W + a]
                if name != self.font.lower():
                    f.set(W + "ascii", self.font)
                    f.set(W + "hAnsi", self.font)
                    fonts += 1
        if self.d.numbering is not None:
            for f in self.d.numbering.iter(W + "rFonts"):
                name = (f.get(W + "ascii") or "").lower()
                if name and not any(k in name for k in KEEP_FONTS) and name != self.font.lower():
                    f.set(W + "ascii", self.font)
                    f.set(W + "hAnsi", self.font)
                    fonts += 1
        if resized:
            self.note(f"Set body text to {self.body_hp / 2:g} pt in {resized} {'place' if resized == 1 else 'places'}.")
        if fonts:
            self.note(f"Changed text in other fonts to {self.font} in {fonts} {'place' if fonts == 1 else 'places'}, heading numbers included.")

    # 8. page numbers: (1), (2), (3) on the front pages, 1, 2, 3 after, at the top right
    def fix_page_numbers(self):
        pn = self.rules["page_number"]
        if pn.get("position", "top") != "top" or pn.get("align", "right") != "right":
            self.note(f"Page numbers were not changed since this format puts them at the {pn.get('position')} {pn.get('align')}. Check them in the report.")
            return
        secs = [sp for sp, _ in self.sections()]

        def fmt(sp):
            pn = sp.find(W + "pgNumType")
            return pn.get(W + "fmt") if pn is not None else None

        def start(sp):
            pn = sp.find(W + "pgNumType")
            return pn.get(W + "start") if pn is not None else None

        first_front, first_body = self.front_range()
        front_rids = set()
        body_rids = set()
        current = None
        for i, sp in enumerate(secs):
            ref = sp.find(W + "headerReference[@" + W + "type='default']")
            if ref is not None:
                current = ref.get(R + "id")
            front = first_front is not None and i >= first_front and (first_body is None or i < first_body)
            pn = sp.find(W + "pgNumType")
            want_roman = front and self.rules["page_number"].get("front_style") == "roman"
            if pn is not None and pn.get(W + "fmt") and not want_roman:
                del pn.attrib[W + "fmt"]
            if current and has_page_field(self.d.part("word/" + self.d.rels[current])):
                (front_rids if front else body_rids).add(current)
        shared = front_rids & body_rids
        if shared:
            self.note("A header is shared by the front pages and the chapters, so its page number was left without brackets.")
        for rid in front_rids | body_rids:
            name = "word/" + self.d.rels[rid]
            hdr = self.d.part(name)
            for child in list(hdr):
                hdr.remove(child)
            p = el("p")
            ppr = el("pPr")
            ppr.append(el("jc", val="right"))
            p.append(ppr)
            paren = rid in front_rids and rid not in shared and self.rules["page_number"].get("front_style") == "paren"

            def run(text=None, fld=None, instr=None):
                r = el("r")
                rpr = el("rPr")
                rpr.append(el("rFonts", ascii=self.font, hAnsi=self.font, cs=self.font))
                rpr.append(el("sz", val=self.body_hp))
                r.append(rpr)
                if fld:
                    r.append(el("fldChar", fldCharType=fld))
                elif instr:
                    t = el("instrText")
                    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    t.text = instr
                    r.append(t)
                else:
                    t = el("t")
                    t.text = text
                    r.append(t)
                return r
            if paren:
                p.append(run("("))
            p.append(run(fld="begin"))
            p.append(run(instr=" PAGE "))
            p.append(run(fld="separate"))
            p.append(run("1"))
            p.append(run(fld="end"))
            if paren:
                p.append(run(")"))
            hdr.append(p)
        fs = self.rules["page_number"].get("front_style")
        self.note(f"Numbered the front pages {'(1), (2), (3)' if fs == 'paren' else 'i, ii, iii' if fs == 'roman' else '1, 2, 3'} and the chapters 1, 2, 3, with the number at the top right.")

    # 9. spelling carried over from the template, and fields to refresh
    def fix_leftovers(self):
        # text boxes holding only a page number, left behind from the template
        n = 0
        for p in list(self.d.body.findall(W + "p")):
            boxes = list(p.iter(W + "txbxContent"))
            if boxes and all(re.fullmatch(r"\d*", ptext(b).strip()) for b in boxes) and not ptext(p).strip().strip("0123456789"):
                for r in list(p.findall(W + "r")):
                    if r.find(".//" + W + "txbxContent") is not None:
                        p.remove(r)
                        n += 1
        if n:
            self.note(f"Removed {n} stray page-number {'box' if n == 1 else 'boxes'} left in the text from the template.")
        # labels the template expects
        for tbl in self.d.body.findall(W + "tbl"):
            for tc in tbl.iter(W + "tc"):
                if ptext(tc).strip() == "Education":
                    ts = list(tc.iter(W + "t"))
                    ts[0].text = "Educational Attainment"
                    for t in ts[1:]:
                        t.text = ""
                    self.note("Renamed the biography label Education to Educational Attainment.")

    def fix_words(self):
        fixes = dict(self.rules.get("template_typos", {}))
        fixes["Faculty of medicine"] = "Faculty of Medicine"
        done = []
        for t in self.d.body.iter(W + "t"):
            for a, b in fixes.items():
                if t.text and a in t.text:
                    t.text = t.text.replace(a, b)
                    done.append(f"{a} to {b}")
        if done:
            self.note("Corrected " + ", ".join(sorted(set(done))) + ".")
        st = self.d.parts.get("word/settings.xml")
        if st is not None and st.find(W + "updateFields") is None:
            after = ("hdrShapeDefaults", "footnotePr", "endnotePr", "compat", "docVars", "rsids", "mathPr", "attachedSchema",
                     "themeFontLang", "clrSchemeMapping", "doNotIncludeSubdocsInStats", "doNotAutoCompressPictures", "forceUpgrade",
                     "captions", "readModeInkLockDown", "smartTagType", "schemaLibrary", "shapeDefaults", "doNotEmbedSmartTags",
                     "decimalSymbol", "listSeparator")
            nxt = next((c for c in st if etree.QName(c).localname in after), None)
            uf = el("updateFields", val="true")
            nxt.addprevious(uf) if nxt is not None else st.append(uf)

    def run(self):
        self.fix_back_matter()
        self.fix_chapter_headings()
        self.fix_margins()
        self.fix_section_heading_indents()
        self.fix_section_numbering()
        self.fix_contents()
        self.fix_headings()
        self.fix_paragraph_indent()
        self.fix_line_spacing()
        self.fix_text()
        self.fix_page_numbers()
        self.fix_words()
        self.fix_leftovers()
        if getattr(self, "images_scaled", 0):
            self.note(f"Shrank {self.images_scaled} pictures to fit inside the new margins.")
        if getattr(self, "tables_scaled", 0):
            self.note(f"Narrowed {self.tables_scaled} tables to fit inside the new margins.")


def run_text(text):
    r = el("r")
    t = el("t")
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    r.append(t)
    return r


def field_runs(instr, shown):
    out = []
    for kind in ("begin", "instr", "separate", "shown", "end"):
        r = el("r")
        if kind == "instr":
            t = el("instrText")
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            t.text = instr
            r.append(t)
        elif kind == "shown":
            t = el("t")
            t.text = shown
            r.append(t)
        else:
            r.append(el("fldChar", fldCharType=kind))
        out.append(r)
    return out


def field_par(instr, shown):
    p = el("p")
    for r in field_runs(instr, shown):
        p.append(r)
    return p


def set_lines(p, lines):
    """Replace the visible text of a paragraph with lines joined by line breaks. Bookmarks stay."""
    for child in list(p):
        if etree.QName(child).localname in ("r", "hyperlink", "smartTag", "fldSimple", "ins"):
            p.remove(child)
    r = el("r")
    for k, line in enumerate(lines):
        if k:
            r.append(el("br"))
        t = el("t")
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = line
        r.append(t)
    p.append(r)


def has_page_field(part):
    for t in part.iter(W + "instrText"):
        if (t.text or "").strip().upper().startswith("PAGE"):
            return True
    return part.find(".//" + W + "fldSimple") is not None


def is_contents_line(p):
    """A line in a contents list: it has a tab or ends in a page number, or uses a TOC style."""
    if style_of(p).upper().startswith("TOC"):
        return True
    if any(r.find(W + "tab") is not None for r in p.iter(W + "r")):
        return True
    return bool(re.search(r"\s(\d+|[ivxlc]+|\(\d+\))\s*$", ptext(p).strip(), re.I)) and not ptext(p).strip().upper().startswith("CHAPTER " + ptext(p).strip().split()[-1])


def is_break_only(p):
    return (etree.QName(p).localname == "p" and not ptext(p).strip()
            and p.find(".//" + W + "br[@" + W + "type='page']") is not None and p.find(".//" + W + "drawing") is None)


def find_idx(items, name):
    for i, e in enumerate(items):
        if etree.QName(e).localname == "p" and re.sub(r"\s+", " ", ptext(e)).strip().upper() == name:
            return i
    return None


def replace_text(p, text):
    ts = list(p.iter(W + "t"))
    if not ts:
        return
    ts[0].text = text
    for t in ts[1:]:
        t.text = ""


NOT_FIXED = [
    "Wording, such as a title that differs between pages and leftover template text.",
    "Page numbers in the contents and lists. Open the fixed file in Word and answer Yes when it asks to update fields, then save it as PDF and check the PDF.",
    "Check items in the report, which the graduate office decides.",
]


def fix(path, rules_path=DEFAULT_RULES, outdir=None, check_after=True):
    with open(rules_path, encoding="utf-8") as fh:
        rules = json.load(fh)
    outdir = outdir or os.path.dirname(os.path.abspath(path))
    base = os.path.splitext(os.path.basename(path))[0]
    if path.lower().endswith(".doc"):
        # an old Word file is saved as .docx by Word first
        if not ((os.name == "nt" or check.MAC) and check.word_available()):
            raise check.ConversionError(f"{os.path.basename(path)} is an old .doc file and needs Microsoft Word to open.")
        import tempfile
        newer = os.path.join(tempfile.mkdtemp(), base + ".docx")
        check.run_word("docx", path, newer)
        path = newer
    elif not path.lower().endswith(".docx"):
        raise check.ConversionError("The corrected copy needs the Word file (.docx or .doc).")
    os.makedirs(outdir, exist_ok=True)
    out = os.path.join(outdir, f"{base} - fixed.docx")
    f = Fixer(path, rules)
    f.run()
    f.d.save(out)
    log = os.path.join(outdir, f"{base} - fixed changes.txt")
    with open(log, "w", encoding="utf-8") as fh:
        fh.write(f"Changes made in {os.path.basename(out)}:\n\n")
        for line in f.log:
            fh.write(f"- {line}\n")
        fh.write("\nLeft for the student:\n\n")
        for line in NOT_FIXED:
            fh.write(f"- {line}\n")
    results = None
    if check_after:
        results = check.run(out, rules_path, outdir)
    return out, log, f.log, results


def main():
    ap = argparse.ArgumentParser(description="Make a corrected copy of a thesis Word file.")
    ap.add_argument("file", help="thesis as .docx")
    ap.add_argument("--rules", default=None)
    ap.add_argument("--out", default=None)
    ap.add_argument("--no-check", action="store_true", help="skip checking the corrected copy")
    a = ap.parse_args()
    try:
        out, log, lines, results = fix(a.file, check.pick_rules(a.rules), a.out, not a.no_check)
    except check.ConversionError as e:
        sys.exit(str(e))
    print(out)
    print(log)
    if results:
        report, corr, F, _, _ = results
        n_fix = sum(1 for x in F.items if x["level"] == "fix")
        print(f"After fixing: {n_fix} to fix, {len(F.items) - n_fix} to check.")
        print(report)
        print(corr)


if __name__ == "__main__":
    main()
