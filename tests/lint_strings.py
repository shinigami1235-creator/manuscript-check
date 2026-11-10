"""Fails on wording Gid has rejected, in every string a person reads.

    python tests/lint_strings.py           strings in the code, page, rules, installer and Mac read-me
    python tests/lint_strings.py --run     also checks the sample thesis and fixes the sample Word file,
                                           then lints the report messages and the change list

Exit code 1 when anything is flagged.
"""
import ast
import glob
import html
import json
import os
import re
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "app"))

RULES = [
    (r"^(That|This|These|Those) (is|are|was|were)\b", "a sentence that only labels the one before it"),
    (r"\bnot only\b", "not only X but Y"),
    (r", not (the |a |an )?\w+", "a tailing 'not' contrast"),
    (r"\b(is|are|was|does|did) not\.$", "a split contrast ending in 'is not'"),
    (r"\b(actually|simply|merely|truly|essentially|fundamentally|arguably|basically)\b", "filler adverb"),
    (r"[—–]", "dash"),
    (r", so (it|you|the|nothing|this|that|they)\b", "a ', so' clause that may only reassure"),
    (r"\bnothing (is|was|gets|will be) (cut|lost|changed|broken|cropped)\b|\bloses nothing\b|\bno matter\b", "reassurance"),
    (r"\bworks too\b|\bone or many\b|\bas well\.$", "padding"),
    (r"\b(crucial|vital|pivotal|robust|seamless|leverage|delve|underscore|tapestry|testament|showcase|holistic|nuanced|myriad|realm|"
     r"elevate|vibrant|boasts|nestled|stands as|serves as)\b", "AI vocabulary"),
    (r"\b(Think of it|In other words|The whole point|At its core|which is (what|why)|The real question)\b", "AI tic"),
    (r", (ensuring|allowing|highlighting|reflecting|meaning|making sure)\b", "decorative -ing tail"),
    (r"\b(uploaded|sent anywhere|leaves your|stays on your|never stored|stays private|was not changed|untouched)\b", "a promise about itself"),
    (r"\band nothing\b", "a state with nothing to do"),
    (r"\b(Please note|Note that|feel free|don't worry|no need to|Rest assured)\b", "padding"),
    (r"!", "exclamation"),
    (r"^Add (it|them)\.$", "a fragment split from the sentence it belongs to"),
]

STOP = set("a an the and or of to in on at is are was were be it its this that these with for from by as so then than into his her "
           "their your you i we not no all any each every one".split())


def words(s):
    return [w for w in re.findall(r"[a-z0-9]+", s.lower()) if w not in STOP]


def sentences(text):
    text = re.sub(r"\s+", " ", text).strip()
    return [s for s in re.split(r"(?<=[.!?])\s+(?=[A-Z\"(])", text) if s]


def lint_text(where, text, out):
    sents = sentences(text)
    for i, s in enumerate(sents):
        for pat, why in RULES:
            if re.search(pat, s):
                out.append((where, why, s))
        if i:
            prev, cur = set(words(sents[i - 1])), words(s)
            if len(cur) >= 3 and sum(w in prev for w in cur) / len(cur) >= 0.75:
                out.append((where, "repeats the sentence before it", s))


# ---------------------------------------------------------------- collecting strings

def py_strings(path):
    tree = ast.parse(open(path, encoding="utf-8").read())
    docs = set()
    for node in ast.walk(tree):
        if isinstance(node, (ast.Module, ast.FunctionDef, ast.ClassDef, ast.AsyncFunctionDef)) and node.body:
            first = node.body[0]
            if isinstance(first, ast.Expr) and isinstance(first.value, ast.Constant):
                docs.add(id(first.value))
    for node in ast.walk(tree):
        if isinstance(node, ast.JoinedStr):
            text = "".join(v.value if isinstance(v, ast.Constant) else "{}" for v in node.values)
            for v in node.values:
                if isinstance(v, ast.Constant):
                    docs.add(id(v))
        elif isinstance(node, ast.Constant) and isinstance(node.value, str) and id(node) not in docs:
            text = node.value
        else:
            continue
        if prose(text):
            yield node.lineno, text


def prose(text):
    """A string a person reads: at least three words, starting with a capital or a placeholder, no code characters."""
    t = text.strip()
    if len(t.split()) < 3 or not re.match(r"[A-Z{\"(]", t):
        return False
    return not re.search(r"[\\<>=;|]|\bw:|https?:|\.py\b|\.json\b|^\{\}\s*\{\}", t)


def html_strings(path):
    src = open(path, encoding="utf-8").read()
    body = re.sub(r"<style.*?</style>", "", src, flags=re.S)
    script = "\n".join(re.findall(r"<script>(.*?)</script>", body, flags=re.S))
    markup = re.sub(r"<script.*?</script>", "", body, flags=re.S)
    for m in re.finditer(r"(placeholder|title)=\"([^\"]+)\"", markup):
        yield "attribute", m.group(2)
    for chunk in re.split(r"<[^>]+>", markup):
        t = html.unescape(chunk).strip()
        if t:
            yield "text", t
    for m in re.finditer(r"`([^`]*)`|\"([^\"\n]*)\"|'([^'\n]*)'", script):
        t = next(g for g in m.groups() if g is not None)
        t = re.sub(r"<[^>]+>", " ", re.sub(r"\$\{[^}]*\}", "{}", t)).strip()
        if prose(t):
            yield "script", t


def collect():
    out = []
    for path in sorted(glob.glob(os.path.join(ROOT, "app", "*.py")) + [os.path.join(ROOT, "launcher.py")]):
        rel = os.path.relpath(path, ROOT)
        for line, text in py_strings(path):
            lint_text(f"{rel}:{line}", text, out)
    for path in glob.glob(os.path.join(ROOT, "app", "ui", "*.html")):
        for kind, text in html_strings(path):
            lint_text(f"{os.path.relpath(path, ROOT)} {kind}", text, out)
    for path in glob.glob(os.path.join(ROOT, "app", "rules", "*.json")):
        def walk(o, key=""):
            if isinstance(o, dict):
                for k, v in o.items():
                    walk(v, k)
            elif isinstance(o, list):
                for v in o:
                    walk(v, key)
            elif isinstance(o, str) and key.startswith("source"):
                lint_text(f"{os.path.relpath(path, ROOT)} {key}", o, out)
        walk(json.load(open(path, encoding="utf-8")))
    for path in glob.glob(os.path.join(ROOT, "packaging", "*.txt")):
        for para in open(path, encoding="utf-8").read().split("\n\n"):
            lint_text(os.path.relpath(path, ROOT), para, out)
    for line in open(os.path.join(ROOT, "packaging", "installer.iss"), encoding="utf-8"):
        m = re.match(r"\s*(WelcomeLabel2|Name: .*Description)\s*[=:]\s*\"?(.*?)\"?\s*$", line)
        if m:
            lint_text("packaging/installer.iss", m.group(2).replace("%n", " "), out)
    return out


def run_samples(out):
    import check
    import fix
    tmp = tempfile.mkdtemp()
    pdfs = [os.path.join(ROOT, "tests", "sample-thesis.pdf")] + sys.argv[2:]
    for pdf in pdfs:
        F = check.run(pdf, outdir=tmp)[2]
        for f in F.numbered():
            lint_text(f"report {os.path.basename(pdf)} #{f['n']}", f["message"], out)
            if f["level"] == "check" and f.get("source"):
                lint_text(f"report {os.path.basename(pdf)} #{f['n']} rule", f["source"], out)
    fix.fix(os.path.join(ROOT, "tests", "sample-thesis.docx"), outdir=tmp, check_after=False)
    for log in glob.glob(os.path.join(tmp, "*changes.txt")):
        for line in open(log, encoding="utf-8"):
            lint_text("change list", line.lstrip("- "), out)


if __name__ == "__main__":
    found = collect()
    if "--run" in sys.argv:
        run_samples(found)
    for where, why, s in found:
        print(f"{where}: {why}\n    {s}")
    print(f"{len(found)} flagged")
    sys.exit(1 if found else 0)
