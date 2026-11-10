"""Downloads newer program code from the release page and unpacks it for the next start."""
import hashlib
import json
import os
import sys
import shutil
import tempfile
import urllib.request
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
# The release page the updates come from. latest.json there names the newest version.
FEED = "https://github.com/shinigami1235-creator/manuscript-check/releases/latest/download/latest.json"
if os.name == "nt":
    UPDATES = os.path.join(os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"), "Manuscript Check", "updates")
elif sys.platform == "darwin":
    UPDATES = os.path.join(os.path.expanduser("~"), "Library", "Application Support", "Manuscript Check", "updates")
else:
    UPDATES = os.path.join(os.path.expanduser("~"), ".local", "share", "Manuscript Check", "updates")


def installer_url(info):
    """The installer for this kind of computer."""
    if sys.platform == "darwin":
        import platform
        key = "installer_url_mac" if platform.machine() == "arm64" else "installer_url_mac_intel"
        return info.get(key) or info.get("installer_url_mac") or ""
    return info.get("installer_url", "")


def vtuple(v):
    try:
        return tuple(int(x) for x in str(v).split("."))
    except ValueError:
        return (0,)


def current():
    with open(os.path.join(HERE, "version.json"), encoding="utf-8") as fh:
        return json.load(fh)


def fetch(url, timeout=15):
    req = urllib.request.Request(url, headers={"User-Agent": "ManuscriptCheck-updater"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read()


def newest_installed():
    best = vtuple(current()["version"])
    if os.path.isdir(UPDATES):
        for name in os.listdir(UPDATES):
            d = os.path.join(UPDATES, name)
            if os.path.exists(os.path.join(d, "version.json")) and not os.path.exists(os.path.join(d, "bad")):
                best = max(best, vtuple(name))
    return best


def check_and_download(shell, feed=None):
    """Returns one of:
    {"status": "current"}                       nothing newer
    {"status": "ready", "version", "notes"}     newer code unpacked, used from the next start
    {"status": "installer", "version", "url"}   newer version needs a new installer
    {"status": "offline"}                       the release page could not be reached
    """
    feed = feed or current().get("feed") or FEED
    try:
        info = json.loads(fetch(feed, timeout=8))
    except Exception:
        return {"status": "offline"}
    version = info["version"]
    if vtuple(version) <= newest_installed():
        here = vtuple(current()["version"])
        if vtuple(version) > here:
            return {"status": "ready", "version": version, "notes": info.get("notes", "")}
        return {"status": "current"}
    if os.path.exists(os.path.join(UPDATES, version, "bad")):
        return {"status": "current"}   # this version failed to start here before
    if int(info.get("shell", 1)) > shell:
        return {"status": "installer", "version": version, "url": installer_url(info)}
    data = fetch(info["code_url"], timeout=120)
    if hashlib.sha256(data).hexdigest() != info["sha256"]:
        return {"status": "offline"}
    os.makedirs(UPDATES, exist_ok=True)
    work = tempfile.mkdtemp(dir=UPDATES, prefix=".incoming-")
    try:
        zpath = os.path.join(work, "code.zip")
        with open(zpath, "wb") as fh:
            fh.write(data)
        unpacked = os.path.join(work, "code")
        with zipfile.ZipFile(zpath) as z:
            for m in z.namelist():
                target = os.path.realpath(os.path.join(unpacked, m))
                if not target.startswith(os.path.realpath(unpacked) + os.sep):
                    raise ValueError("unsafe path in update")
            z.extractall(unpacked)
        if not os.path.exists(os.path.join(unpacked, "gui.py")):
            raise ValueError("update has no program code")
        final = os.path.join(UPDATES, version)
        if os.path.exists(final):
            shutil.rmtree(final, ignore_errors=True)
        os.replace(unpacked, final)
    finally:
        shutil.rmtree(work, ignore_errors=True)
    tidy(keep=2)
    return {"status": "ready", "version": version, "notes": info.get("notes", "")}


def tidy(keep=2):
    """Keep the newest few downloaded versions and the one running now."""
    if not os.path.isdir(UPDATES):
        return
    running = os.path.realpath(HERE)
    names = sorted((n for n in os.listdir(UPDATES) if not n.startswith(".")), key=vtuple, reverse=True)
    for n in names[keep:]:
        d = os.path.join(UPDATES, n)
        if os.path.realpath(d) != running:
            shutil.rmtree(d, ignore_errors=True)
