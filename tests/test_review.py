"""A school stamp in the margin is left out of the checks, and problems ignored in the review
window are left out of the report and the marked pages."""
import html
import os
import sys
import tempfile
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "app"))
sys.path.insert(0, os.path.join(ROOT, "tests"))
tmp = tempfile.mkdtemp()
os.environ["APPDATA"] = tmp
import check  # noqa: E402
import gui  # noqa: E402
from make_stamped import stamp  # noqa: E402

plain = os.path.join(ROOT, "tests", "sample-thesis.pdf")
stamped = os.path.join(tmp, "stamped.pdf")
stamp(plain, stamped)


def messages(path):
    return {f["message"] for f in check.run(path, outdir=os.path.join(tmp, "out"))[2].numbered()}


want = messages(plain)
assert messages(stamped) == want, messages(stamped) ^ want
keep = check.drop_stamps
check.drop_stamps = lambda pages, rules: None
assert messages(stamped) != want, "the stamp made no difference, so this test proves nothing"
check.drop_stamps = keep

api = gui.Api()
api._settings["results"] = os.path.join(tmp, "results")
api._add([stamped])
api.start()
while api._batch.running:
    time.sleep(0.5)
item = api._items[0]
assert item.status == "done", item.message
review = api.review(item.id)
assert len(review["findings"]) == item.fix_count + item.check_count
assert len(review["pages"]) == item.pages
assert api.review_page(item.id, 1).startswith("data:image/jpeg;base64,")
first, second = review["findings"][0], review["findings"][1]
before = open(item.report, encoding="utf-8").read()
marked_before = os.path.getsize(item.corrections)
api.set_ignored(item.id, first["id"], True)
api.set_ignored(item.id, second["id"], True)
api.set_ignored(item.id, second["id"], False)
api.finish_review(item.id)
after = open(item.report, encoding="utf-8").read()
assert item.ignored == 1
assert html.escape(first["message"]) in before and html.escape(first["message"]) not in after
assert html.escape(second["message"]) in after
assert after.count("<tr") == before.count("<tr") - 1, (before.count("<tr"), after.count("<tr"))
assert os.path.getsize(item.corrections) != marked_before
print("OK")
