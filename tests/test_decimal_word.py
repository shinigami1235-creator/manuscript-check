"""Word files converted from PDF write settings such as margins as 400.0. The corrected copy must still be made."""
import os
import re
import sys
import tempfile
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "app"))
import fix  # noqa: E402

tmp = tempfile.mkdtemp()
src = os.path.join(ROOT, "tests", "sample-thesis.docx")
dst = os.path.join(tmp, "decimal settings.docx")
pat = r'(w:(?:w|h|left|right|top|bottom|header|footer|gutter|hanging|firstLine|start|end|pos|line|before|after|val)=")(\d+)"'
with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
    for info in zin.infolist():
        data = zin.read(info.filename)
        if info.filename.startswith("word/") and info.filename.endswith(".xml") and "numbering" not in info.filename:
            data = re.sub(pat, r'\g<1>\g<2>.0"', data.decode("utf-8")).encode("utf-8")
        zout.writestr(info, data)
fixed = fix.fix(dst, outdir=os.path.join(tmp, "out"), check_after=False)[0]
assert os.path.exists(fixed), fixed
print("OK", fixed)
