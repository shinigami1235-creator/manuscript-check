"""Learn a thesis format from a school's template.

Measures a template PDF (and its Word file, when given) and writes a rules file
that check.py and fix.py use. Look over the numbers in the rules file once before
using it, since a template full of short placeholder lines can hide the right margin.

Usage:
  python learn.py "Template.pdf"
  python learn.py "Template.pdf" --docx "Template.docx" --name "Mahidol MSc thesis 2025"
"""
import argparse
import json
import os
import re
import statistics
import sys
from collections import Counter

import pdfplumber

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import check  # noqa: E402

PLACEHOLDER_WORDS = ["Insert", "Firstname", "Surname", "Academic Title", "20xx", "20XX", "XXXX", "Degree Title", "(If any)",
                     "Approval date", "Table xx", "Figure xx", "Choose one of the following", "Main heading", "Subheading",
                     "CHAPTER TITLE", "THESIS TITLE", "Thesis Title Here", "Your Name", "Name Surname", "Lorem ipsum"]
OPTIONAL = ("LIST OF", "APPENDIX", "APPENDICES", "GLOSSARY", "DEDICATION", "PREFACE")


def mode(values, default=None):
    values = [v for v in values if v is not None]
    return Counter(values).most_common(1)[0][0] if values else default


def font_family(name):
    name = name.split("+")[-1]
    name = re.sub(r"[-,](Bold|Italic|BoldItalic|Regular|Roman|Oblique)+(MT|PS|PSMT)?$", "", name)
    name = re.sub(r"(PSMT|MT|PS)$", "", name)
    return name.rstrip(",-")


def is_placeholder(text):
    return any(w.lower() in text.lower() for w in PLACEHOLDER_WORDS) or bool(re.search(r"[xX]{2,}", text))


def learn(pdf_path, docx_path=None, name=None):
    base_rules = json.load(open(check.DEFAULT_RULES, encoding="utf-8"))
    notes = []
    with pdfplumber.open(pdf_path) as pdf:
        # first pass with a loose page-number zone, both top and bottom
        probe = json.loads(json.dumps(base_rules))
        probe["page_number"]["zone_bottom_cm"] = 3.3
        pages = [check.Page(i, pg, probe) for i, pg in enumerate(pdf.pages, 1)]
        raw = [check.read_lines(pg, i) for i, pg in enumerate(pdf.pages, 1)]
    W = mode(round(p.width, 1) for p in pages)
    H = mode(round(p.height, 1) for p in pages)

    # page numbers: top or bottom, which edge, which style
    top_labels, bottom_labels = [], []
    for p, ls in zip(pages, raw):
        for l in ls:
            if check.parse_label(l.text):
                if l.top < 3.3:
                    top_labels.append((p, l))
                elif l.top > p.height - 3.3:
                    bottom_labels.append((p, l))
    position = "top" if len(top_labels) >= len(bottom_labels) else "bottom"
    labels = top_labels if position == "top" else bottom_labels
    x1s = [l.x1 for _, l in labels]
    centers = [l.center for _, l in labels]
    align = "center" if labels and statistics.pstdev(centers) < statistics.pstdev(x1s) and abs(statistics.median(centers) - W / 2) < 1.5 else "right"
    right_edge = round(statistics.median(x1s), 2) if labels else round(W - 2.4, 2)
    if position == "top":
        zone = round(max((l.top for _, l in labels), default=2.6) + 0.35, 2)
    else:
        zone = round(H - min((l.top for _, l in labels), default=H - 2.5) + 0.3, 2)
    probe["page_number"].update({"position": position, "zone_bottom_cm": zone if position == "top" else 2.95, "zone_from_bottom_cm": zone})
    with pdfplumber.open(pdf_path) as pdf:
        pages = [check.Page(i, pg, probe) for i, pg in enumerate(pdf.pages, 1)]

    # body text size and font
    body_lines = [l for p in pages[1:] for l in p.lines if not l.bold and len(l.text.split()) >= 2]
    body_pt = mode(l.size for l in body_lines) or 12
    body_pt = int(body_pt) if float(body_pt).is_integer() else body_pt
    fam = Counter()
    for p in pages:
        for c in p.chars:
            fam[font_family(c["fontname"])] += 1
    family = fam.most_common(1)[0][0] if fam else "TimesNewRoman"
    allowed = [family]
    if "times" in family.lower():
        allowed = list(dict.fromkeys(base_rules["fonts"]["allowed"] + [family]))

    # headings: bold capitals larger than body text, or bold centred capitals
    def centred(l, p):
        return abs(l.center - p.width / 2) < 2.0

    cands = [(p, l) for p in pages[1:] for l in p.lines if l.bold and l.text.isupper() and centred(l, p) and len(l.text) > 3]
    heading_pt = mode(l.size for _, l in cands if l.size > body_pt + 0.5) or mode(l.size for _, l in cands) or body_pt + 2

    # cover
    cover = pages[0]
    cover_title_pt = max((l.size for l in cover.lines), default=heading_pt)
    cover_text_pt = mode(l.size for l in cover.lines if l.size < cover_title_pt) or heading_pt
    cover_lines = []
    for l in cover.lines:
        t = check.norm(l.text).upper()
        if l.size >= cover_title_pt - 0.2:
            continue
        t = re.sub(r"\s*\b(20|25)[0-9X]{2}\b.*$", "", t).strip()   # ACADEMIC YEAR 20XX -> ACADEMIC YEAR
        if t and not is_placeholder(t) and t not in cover_lines:
            cover_lines.append(t)

    # margins: text edges on pages after the cover
    lefts = [round(min(l.x0 for l in p.lines), 2) for p in pages[1:] if p.lines and p.width < p.height]
    left_text = mode(lefts, 3.81)
    tops = [round(p.lines[0].top, 2) for p in pages[1:] if p.lines]
    rights = [l.x1 for p in pages[1:] for l in p.lines if p.width < p.height]
    right_cm = round(W - max(rights), 2) if rights else 2.54
    top_cm = round(min(tops) - 0.1, 2) if tops else 2.54
    bottom_cm = 2.54
    sources = {"left": "Template PDF: where text starts on most pages",
               "right": "Template PDF: the widest line",
               "top": "Template PDF: highest text below the page number",
               "bottom": "A template PDF does not show the bottom margin."}
    if docx_path:
        try:
            import docx
            d = docx.Document(docx_path)
            secs = [s for s in d.sections if s.page_width and s.page_width < s.page_height]
            cmv = lambda v: round(v.cm, 2)
            r_docx = mode(cmv(s.right_margin) for s in secs)
            t_docx = mode(cmv(s.top_margin) for s in secs)
            b_docx = mode(cmv(s.bottom_margin) for s in secs)
            if r_docx:
                right_cm, sources["right"] = r_docx, "Template Word file: right margin"
            if t_docx:
                top_cm, sources["top"] = t_docx, "Template Word file: top margin"
            if b_docx and b_docx >= 1.5:
                bottom_cm, sources["bottom"] = b_docx, "Template Word file: bottom margin"
            elif b_docx:
                notes.append(f"The template Word file sets the bottom margin to {b_docx} cm, which is too small to print, and 2.54 cm is used.")
        except Exception as e:
            notes.append(f"The template Word file could not be read ({e}) and the margins come from the PDF.")

    # line spacing: distance between consecutive body lines
    pitches = []
    for p in pages[1:]:
        ls = [l for l in p.lines if not l.bold and abs(l.size - body_pt) < 0.4]
        for a, b in zip(ls, ls[1:]):
            d = b.top - a.top
            if 0.3 < d < 1.3:
                pitches.append(round(d, 2))
    pitch = mode(pitches, 0.74)
    spacing_name = {1: "single", 1.5: "1.5 lines", 2: "double"}.get(round(pitch / (body_pt * 1.15 * 2.54 / 72) * 2) / 2, f"{pitch} cm")

    # paragraph indent: how far body lines sit in from the margin
    offs = [round(l.x0 - left_text, 2) for l in body_lines if l.x0 - left_text > 0.4 and not centred(l, pages[l.page - 1])]
    indent = mode(offs, 1.27)

    # numbered section headings
    secs = [l for p in pages for l in p.lines if l.bold and check.SEC_RE.match(check.norm(l.text)) and check.norm(l.text).split()[0].count(".") == 1]
    sec_x = mode(round(l.x0, 2) for l in secs) or left_text
    sec_pt = mode(l.size for l in secs) or body_pt

    # headings in page order
    order, chapter_titles, front, back = [], {}, [], []
    seen_chapter = False
    approval = next((p.no for p in pages[:6] if "was approved" in p.text.lower()), 1)
    for p in pages[approval:]:
        for i, l in enumerate(p.lines):
            t = check.norm(l.text).upper()
            if not (l.bold and abs(l.size - heading_pt) < 0.6 and centred(l, p)):
                continue
            m = re.match(r"^CHAPTER\s+(\d+)\s*(.*)$", t)
            if m:
                seen_chapter = True
                title = m.group(2).strip()
                if not title and i + 1 < len(p.lines):
                    title = check.norm(p.lines[i + 1].text).upper()
                if title and not is_placeholder(title):
                    chapter_titles[m.group(1)] = title
                if "CHAPTER" not in order:
                    order.append("CHAPTER")
                break
            if re.match(r"^APPENDIX\s+[A-Z0-9]+$", t):
                t = "APPENDIX"
            if t in order or is_placeholder(t) or len(t) > 40 or len(t) < 4:
                break
            order.append(t)
            (back if seen_chapter else front).append(t)
            break
    required = [h for h in order if not h.startswith(OPTIONAL)]

    # labelled fields on the abstract and biography pages: the words left of the value column
    def labels_on(page_name):
        page = next((p for p in pages if any(check.norm(l.text).upper() == page_name for l in p.lines)), None)
        if page is None:
            return []
        with pdfplumber.open(pdf_path) as pdf:
            words = pdf.pages[page.no - 1].extract_words(keep_blank_chars=False)
        rows = {}
        for w in words:
            rows.setdefault(round(w["top"] / 3), []).append(w)
        firsts = Counter(round(check.cm(sorted(ws, key=lambda w: w["x0"])[0]["x0"]), 1) for ws in rows.values())
        label_x = firsts.most_common(1)[0][0] if firsts else None
        out = []
        for ws in rows.values():
            ws = sorted(ws, key=lambda w: w["x0"])
            if round(check.cm(ws[0]["x0"]), 1) != label_x:
                continue
            label = [ws[0]["text"]]
            for a, b in zip(ws, ws[1:]):
                if check.cm(b["x0"] - a["x1"]) > 0.6:
                    break
                label.append(b["text"])
            raw_label = " ".join(label)
            if re.search(r"\((if any|optional|if applicable)\)", " ".join(w["text"] for w in ws), re.I):
                continue
            text = re.sub(r"\(.*?\)", "", raw_label).strip(" :")
            if 2 < len(text) < 40 and text[0].isupper() and not text.isupper() and len(label) < len(ws) and text not in out:
                out.append(text)
        return out

    abstract_fields = labels_on(front[0]) if front else []
    bio = next((h for h in back if h in check.BIO_NAMES), None)
    biography_fields = labels_on(bio) if bio else []

    # placeholder text found in the template
    placeholders = list(base_rules["placeholders"])
    for p in pages:
        for l in p.lines:
            t = check.norm(l.text)
            m = re.match(r"^(Insert [A-Za-z ]+?)(?:\s*\d+)?$", t)
            if m and m.group(1) not in placeholders:
                placeholders.append(m.group(1))

    # page number styles: before and after the first chapter
    first_ch = next((p.no for p in pages if p.lines and re.match(r"^CHAPTER\s+\d", check.norm(p.lines[0].text).upper())), None)
    front_styles = Counter(p.label[0] for p in pages if p.label and (first_ch is None or p.no < first_ch))
    body_styles = Counter(p.label[0] for p in pages if p.label and first_ch and p.no >= first_ch)
    front_style = front_styles.most_common(1)[0][0] if front_styles else "roman"
    body_style = body_styles.most_common(1)[0][0] if body_styles else "arabic"
    front_first = next((p for p in pages if p.label and p.label[0] == front_style), None)
    h = check.find_heading(front_first, probe) if front_first else None
    starts_at = check.heading_key(h) if h else None

    alltext = "\n".join(p.text for p in pages)
    toc_page = next((p for p in pages if any(check.norm(l.text).upper() in check.TOC_NAMES for l in p.lines)), None)
    ack_page = next((p for p in pages if any(check.norm(l.text).upper() in ("ACKNOWLEDGEMENTS", "ACKNOWLEDGMENTS") for l in p.lines)), None)
    rules = {
        "name": name or os.path.splitext(os.path.basename(pdf_path))[0],
        "approval_page": "was approved" in alltext.lower(),
        "abstract_keywords": bool(re.search(r"Keywords?\s*:", alltext)),
        "toc_page_label": bool(toc_page and any(check.norm(l.text).upper() == "PAGE" for l in toc_page.lines)),
        "acknowledgements_signed": bool(ack_page and ack_page.lines and ack_page.lines[-1].x0 > ack_page.width / 2),
        "template_file": os.path.basename(pdf_path),
        "learned_from": os.path.basename(pdf_path) + (f" and {os.path.basename(docx_path)}" if docx_path else ""),
        "page": {"width_cm": W, "height_cm": H, "tolerance_cm": 0.2, "source": f"Template PDF: pages are {W} x {H} cm"},
        "fonts": {"allowed": allowed, "ignored": base_rules["fonts"]["ignored"], "name": re.sub(r"(?<=[a-z])(?=[A-Z])", " ", family),
                  "source": f"Template PDF: text is set in {family}"},
        "sizes": {"body_pt": body_pt, "heading_pt": heading_pt, "cover_title_pt": cover_title_pt, "cover_text_pt": cover_text_pt,
                  "superscript_max_pt": round(body_pt * 0.78, 1), "chart_text_max_pt": round(body_pt * 0.84, 1),
                  "source": f"Template PDF: body {body_pt:g} pt, headings {heading_pt:g} pt bold, cover title {cover_title_pt:g} pt, other cover lines {cover_text_pt:g} pt"},
        "margins": {"left_text_cm": left_text, "right_cm": right_cm, "top_text_cm": top_cm, "bottom_cm": bottom_cm, "tolerance_cm": 0.12,
                    "source_left": sources["left"], "source_right": sources["right"], "source_top": sources["top"], "source_bottom": sources["bottom"]},
        "page_number": {"position": position, "align": align, "zone_bottom_cm": zone if position == "top" else 2.95, "zone_from_bottom_cm": zone,
                        "right_edge_cm": right_edge, "tolerance_cm": 0.35,
                        "center_cm": round(statistics.median(centers), 2) if labels else None,
                        "front_style": front_style, "front_starts_at": starts_at, "body_style": body_style,
                        "source": f"Template PDF: number at the {position} {align}, {check.STYLE_NAME.get(front_style, front_style)} on the front pages, "
                                  f"then {check.STYLE_NAME.get(body_style, body_style)} from Chapter 1"},
        "line_spacing": {"expected_cm": pitch, "fail_below_cm": round(pitch * 0.84, 2),
                         "source": f"Template PDF: {body_pt:g} pt lines sit {pitch} cm apart ({spacing_name})"},
        "paragraph_indent": {"expected_cm": indent, "tolerance_cm": 0.6,
                             "source": f"Template PDF: body text under a heading starts {indent} cm in from the margin"},
        "section_heading": {"bold": True, "size_pt": sec_pt, "level1_x_cm": sec_x, "tolerance_cm": 0.25,
                            "source": f"Template PDF: 1.1 headings are {sec_pt:g} pt bold, {sec_x} cm from the left edge"},
        "order": order, "required": required, "not_chapters": base_rules["not_chapters"],
        "order_source": "Template PDF page order",
        "front_headings": front, "back_headings": back,
        "chapter_titles": chapter_titles,
        "cover_lines": cover_lines,
        "abstract_fields": abstract_fields,
        "biography_fields": biography_fields,
        "placeholders": placeholders,
        "template_typos": {},
    }
    return rules, notes


def describe(rules, notes):
    m, s, pn = rules["margins"], rules["sizes"], rules["page_number"]
    lines = [
        f"Learned from {rules['learned_from']}.",
        "",
        f"Page: {rules['page']['width_cm']} x {rules['page']['height_cm']} cm",
        f"Font: {rules['fonts'].get('name') or rules['fonts']['allowed'][0]}",
        f"Text sizes: body {s['body_pt']:g} pt, headings {s['heading_pt']:g} pt, cover title {s['cover_title_pt']:g} pt",
        f"Margins: text starts {m['left_text_cm']} cm from the left, right {m['right_cm']} cm, top {m['top_text_cm']} cm, bottom {m['bottom_cm']} cm",
        f"Line spacing: {rules['line_spacing']['source'].split(': ', 1)[1]}",
        f"Body text indent: {rules['paragraph_indent']['expected_cm']} cm",
        f"Page numbers: {pn['source'].split(': ', 1)[1]}",
        f"Section order: {', '.join(rules['order'])}",
        f"Required: {', '.join(rules['required'])}",
        f"Chapter titles: {', '.join(f'{k} {v}' for k, v in rules['chapter_titles'].items()) or 'none named in the template'}",
        f"Cover lines: {' / '.join(rules['cover_lines'])}",
        f"Abstract page fields: {', '.join(rules['abstract_fields']) or 'none found'}",
        f"Biography fields: {', '.join(rules['biography_fields']) or 'none found'}",
    ]
    if notes:
        lines += [""] + notes
    lines += ["", "Compare these numbers with the school's format guide and correct the rules file where they differ."]
    return "\n".join(lines)


def learn_and_save(pdf, docx_path=None, name=None):
    """Learn a format and save it with the formats learned on this computer. Returns (rules file, summary)."""
    if not docx_path:
        guess = os.path.splitext(pdf)[0] + ".docx"
        docx_path = guess if os.path.exists(guess) else None
    rules, notes = learn(pdf, docx_path, name)
    slug = re.sub(r"[^a-z0-9]+", "-", rules["name"].lower()).strip("-")[:60] or "format"
    os.makedirs(check.USER_RULES, exist_ok=True)
    out = os.path.join(check.USER_RULES, f"{slug}.json")
    n = 2
    while os.path.exists(out):
        out = os.path.join(check.USER_RULES, f"{slug}-{n}.json")
        n += 1
    with open(out, "w", encoding="utf-8") as fh:
        json.dump(rules, fh, indent=2, ensure_ascii=False)
    return out, describe(rules, notes)


def main():
    ap = argparse.ArgumentParser(description="Learn a thesis format from a template.")
    ap.add_argument("pdf", help="template as PDF")
    ap.add_argument("--docx", help="the same template as a Word file, for its margins")
    ap.add_argument("--name", help="name for this format, such as 'Mahidol MSc thesis 2025'")
    ap.add_argument("--out", help="rules file to write (default: rules/<template name>.json)")
    a = ap.parse_args()
    docx_path = a.docx
    if not docx_path:
        guess = os.path.splitext(a.pdf)[0] + ".docx"
        docx_path = guess if os.path.exists(guess) else None
    rules, notes = learn(a.pdf, docx_path, a.name)
    slug = re.sub(r"[^a-z0-9]+", "-", rules["name"].lower()).strip("-")[:60]
    os.makedirs(check.USER_RULES, exist_ok=True)
    out = a.out or os.path.join(check.USER_RULES, f"{slug}.json")
    if os.path.exists(out) and not a.out:
        out = out[:-5] + "-new.json"
    with open(out, "w", encoding="utf-8") as fh:
        json.dump(rules, fh, indent=2, ensure_ascii=False)
    summary = describe(rules, notes)
    note_path = os.path.join(os.path.dirname(os.path.abspath(a.pdf)), f"{rules['name']} - format learned.txt")
    with open(note_path, "w", encoding="utf-8") as fh:
        fh.write(summary + f"\n\nRules file: {out}\n")
    print(summary)
    print()
    print(f"Saved as {out}")


if __name__ == "__main__":
    main()
