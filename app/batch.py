"""Checks a list of theses one after another and keeps the results for the app window."""
import datetime
import json
import os
import re
import threading
import traceback

import check
import fix

SETTINGS = os.path.join(check.USER_DIR, "settings.json")


def documents_folder():
    if os.name == "nt":
        try:
            import ctypes
            from ctypes import wintypes
            buf = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
            ctypes.windll.shell32.SHGetFolderPathW(None, 5, None, 0, buf)  # 5 = My Documents
            if buf.value:
                return buf.value
        except Exception:
            pass
    return os.path.join(os.path.expanduser("~"), "Documents")


def load_settings():
    s = {"results": os.path.join(documents_folder(), "Manuscript Check results"), "format": check.DEFAULT_RULES, "make_fixed": True}
    try:
        with open(SETTINGS, encoding="utf-8") as fh:
            s.update(json.load(fh))
    except Exception:
        pass
    if not os.path.exists(s["format"]):
        s["format"] = check.DEFAULT_RULES
    return s


def save_settings(s):
    os.makedirs(check.USER_DIR, exist_ok=True)
    with open(SETTINGS, "w", encoding="utf-8") as fh:
        json.dump(s, fh, indent=2)


def formats():
    out = []
    for f in check.rules_files():
        try:
            name = json.load(open(f, encoding="utf-8")).get("name", os.path.basename(f))
        except Exception:
            continue
        out.append({"file": f, "name": name, "learned": os.path.dirname(f) == check.USER_RULES})
    return out


KINDS = {".pdf": "PDF", ".docx": "Word", ".doc": "Word"}


class Item:
    _next = 1

    def __init__(self, path):
        self.id = Item._next
        Item._next += 1
        self.path = path
        self.name = os.path.basename(path)
        self.kind = KINDS[os.path.splitext(path)[1].lower()]
        self.reset()

    def reset(self):
        self.status = "waiting"      # waiting, checking, fixing, done, failed
        self.message = ""
        self.fix_count = self.check_count = self.pages = None
        self.report = self.corrections = self.fixed = self.changes = None
        self.findings = []

    def to_dict(self):
        return {k: getattr(self, k) for k in ("id", "path", "name", "kind", "status", "message", "fix_count", "check_count", "pages",
                                               "report", "corrections", "fixed", "changes")}


def safe_name(name):
    return re.sub(r'[<>:"/\\|?*]+', "_", name).strip(" .") or "thesis"


class Batch:
    def __init__(self, items, rules_path, results_root, make_fixed):
        self.items = items
        self.rules_path = rules_path
        self.results_root = results_root
        self.make_fixed = make_fixed
        self.folder = None
        self.running = False
        self.done = False
        self.thread = None

    def start(self):
        self.running = True
        self.thread = threading.Thread(target=self.run, daemon=True)
        self.thread.start()

    def run(self):
        stamp = datetime.datetime.now().strftime("%Y-%m-%d %H.%M")
        self.folder = os.path.join(self.results_root, f"Checked {stamp}")
        os.makedirs(self.folder, exist_ok=True)
        uses_word = any(i.kind == "Word" for i in self.items)
        if uses_word and os.name == "nt":
            check.WordSession.start()
        try:
            for item in self.items:
                if item.status == "done":
                    continue
                self.run_one(item)
        finally:
            check.WordSession.stop()
            self.running = False
            self.done = True

    def run_one(self, item):
        item.reset()
        base = os.path.splitext(item.name)[0]
        outdir = os.path.join(self.folder, safe_name(base))
        n = 2
        while os.path.exists(outdir):
            outdir = os.path.join(self.folder, f"{safe_name(base)} ({n})")
            n += 1
        os.makedirs(outdir)
        try:
            item.status = "checking"
            report, corrections, F, pages, _ = check.run(item.path, self.rules_path, outdir)
            item.report, item.corrections = report, corrections
            item.pages = len(pages)
            items = F.numbered()
            item.findings = items
            item.fix_count = sum(1 for f in items if f["level"] == "fix")
            item.check_count = len(items) - item.fix_count
            if item.kind == "Word" and self.make_fixed:
                item.status = "fixing"
                fixed, changes, _, _ = fix.fix(item.path, self.rules_path, outdir, check_after=False)
                item.fixed, item.changes = fixed, changes
            item.status = "done"
        except check.ConversionError as e:
            item.status, item.message = "failed", str(e)
        except Exception as e:
            item.status = "failed"
            item.message = f"{item.name} could not be checked ({type(e).__name__}: {e})."
            try:
                with open(os.path.join(outdir, "error.txt"), "w", encoding="utf-8") as fh:
                    fh.write(traceback.format_exc())
            except Exception:
                pass


def export_summary(batch, rules_name):
    """One spreadsheet for the whole batch: a row per thesis, then every problem."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    wb = Workbook()
    ws = wb.active
    ws.title = "Theses"
    head = ["File", "Type", "Pages", "To fix", "To check", "Main problems", "Report", "Marked pages", "Corrected copy", "Note"]
    ws.append([f"Checked against {rules_name} on {datetime.datetime.now():%d %B %Y}"])
    ws.append(head)
    for i in batch.items:
        main = "; ".join(f["area"] for f in i.findings if f["level"] == "fix")
        main = ", ".join(dict.fromkeys(main.split("; "))) if main else ""
        ws.append([i.name, i.kind, i.pages, i.fix_count, i.check_count, main, "Open" if i.report else "", "Open" if i.corrections else "",
                   "Open" if i.fixed else "", i.message])
        r = ws.max_row
        for col, path in ((7, i.report), (8, i.corrections), (9, i.fixed)):
            if path:
                ws.cell(r, col).hyperlink = path
                ws.cell(r, col).style = "Hyperlink"
    ws2 = wb.create_sheet("All problems")
    ws2.append(["File", "#", "Fix or Check", "Area", "PDF page", "Problem", "Template rule"])
    for i in batch.items:
        for f in i.findings:
            ws2.append([i.name, f["n"], "Fix" if f["level"] == "fix" else "Check", f["area"], f["page"], f["message"], f["source"] or ""])
    bold = Font(bold=True)
    fill = PatternFill("solid", fgColor="E9E7E1")
    for sheet, widths, hrow in ((ws, [42, 8, 7, 8, 9, 60, 9, 13, 14, 50], 2), (ws2, [42, 5, 12, 20, 9, 110, 50], 1)):
        for c in sheet[hrow]:
            c.font, c.fill = bold, fill
        for k, w in enumerate(widths):
            sheet.column_dimensions[chr(65 + k)].width = w
        sheet.freeze_panes = f"A{hrow + 1}"
        for row in sheet.iter_rows(min_row=hrow + 1):
            for c in row:
                c.alignment = Alignment(vertical="top", wrap_text=c.column in (6, 10) if sheet is ws else c.column in (6, 7))
    ws["A1"].font = Font(bold=True, size=12)
    out = os.path.join(batch.folder, "Summary.xlsx")
    wb.save(out)
    return out
