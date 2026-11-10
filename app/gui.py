"""Manuscript Check: the window staff use."""
import json
import os
import subprocess
import sys
import threading

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import batch  # noqa: E402
import check  # noqa: E402

def _version():
    try:
        with open(os.path.join(HERE, "version.json"), encoding="utf-8") as fh:
            return json.load(fh)["version"]
    except Exception:
        return "1.0.0"


VERSION = _version()


def open_path(path):
    if not path or not os.path.exists(path):
        return False
    if os.name == "nt":
        os.startfile(path)
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])
    return True


class Api:
    def __init__(self):
        self._window = None
        self._settings = batch.load_settings()
        self._items = []
        self._batch = None
        self._word = None
        self._shell = 1
        self._update = {"status": "checking"}

    # ---- state for the page
    def state(self):
        if self._word is None:
            self._word = check.word_available()
        return {"formats": batch.formats(), "format": self._settings["format"], "results": self._settings["results"],
                "make_fixed": self._settings["make_fixed"], "word": self._word, "version": VERSION,
                "items": [i.to_dict() for i in self._items], "running": bool(self._batch and self._batch.running),
                "folder": self._batch.folder if self._batch else None}

    # ---- files
    def _add(self, paths):
        known = {os.path.normcase(i.path) for i in self._items}
        added = []
        for p in paths or []:
            if not p:
                continue
            if os.path.isdir(p):
                inner = [os.path.join(p, f) for f in sorted(os.listdir(p))]
                added += self._add(inner)
                continue
            ext = os.path.splitext(p)[1].lower()
            if ext not in batch.KINDS or os.path.basename(p).startswith("~$") or os.path.normcase(p) in known:
                continue
            item = batch.Item(p)
            self._items.append(item)
            known.add(os.path.normcase(p))
            added.append(item.to_dict())
        return added

    def add_files(self):
        import webview
        paths = self._window.create_file_dialog(webview.OPEN_DIALOG, allow_multiple=True,
                                               file_types=("Theses (*.pdf;*.docx;*.doc)", "All files (*.*)"))
        self._add(paths)
        return self.state()

    def add_folder(self):
        import webview
        paths = self._window.create_file_dialog(webview.FOLDER_DIALOG)
        self._add(paths)
        return self.state()

    def remove(self, item_id):
        if not (self._batch and self._batch.running):
            self._items = [i for i in self._items if i.id != item_id]
        return self.state()

    def clear(self):
        if not (self._batch and self._batch.running):
            self._items = []
            self._batch = None
        return self.state()

    # ---- settings
    def set_format(self, path):
        self._settings["format"] = path
        batch.save_settings(self._settings)
        return self.state()

    def set_make_fixed(self, value):
        self._settings["make_fixed"] = bool(value)
        batch.save_settings(self._settings)
        return self.state()

    def change_folder(self):
        import webview
        paths = self._window.create_file_dialog(webview.FOLDER_DIALOG, directory=self._settings["results"] if os.path.isdir(self._settings["results"]) else "")
        if paths:
            self._settings["results"] = paths[0]
            batch.save_settings(self._settings)
        return self.state()

    # ---- running
    def start(self):
        todo = [i for i in self._items if i.status != "done"]
        if not todo or (self._batch and self._batch.running):
            return self.state()
        for i in todo:
            i.reset()
        self._batch = batch.Batch(todo, self._settings["format"], self._settings["results"], self._settings["make_fixed"])
        self._batch.start()
        return self.state()

    def stop(self):
        if self._batch and self._batch.running:
            self._batch.stop()
        return self.state()

    def recheck_all(self):
        for i in self._items:
            i.reset()
        return self.start()

    # ---- results
    def open(self, path):
        return open_path(path)

    def open_results(self):
        folder = self._batch.folder if self._batch and self._batch.folder else self._settings["results"]
        os.makedirs(folder, exist_ok=True)
        return open_path(folder)

    def save_summary(self):
        if not self._batch or not self._batch.folder:
            return None
        name = next((f["name"] for f in batch.formats() if f["file"] == self._settings["format"]), "the template")
        out = batch.export_summary(self._batch, name)
        open_path(out)
        return out

    # ---- updates
    def _check_updates(self):
        try:
            import updater
            self._update = updater.check_and_download(self._shell)
        except Exception as e:
            self._update = {"status": "offline", "detail": str(e)}

    def update_status(self):
        return dict(self._update, current=VERSION)

    def restart(self):
        if self._batch and self._batch.running:
            return False
        if getattr(sys, "frozen", False):
            subprocess.Popen([sys.executable])
        else:
            subprocess.Popen([sys.executable, os.path.join(os.path.dirname(HERE), "launcher.py")])
        self._window.destroy()
        return True

    def open_url(self, url):
        import webbrowser
        return webbrowser.open(url)

    # ---- formats
    def learn_format(self, name):
        import webview
        paths = self._window.create_file_dialog(webview.OPEN_DIALOG, file_types=("Template PDF (*.pdf)",))
        if not paths:
            return {"cancelled": True}
        import learn
        try:
            out, summary = learn.learn_and_save(paths[0], None, name.strip() or None)
        except Exception as e:
            return {"error": f"{os.path.basename(paths[0])} could not be read as a template ({type(e).__name__}: {e})."}
        self._settings["format"] = out
        batch.save_settings(self._settings)
        return {"summary": summary, "file": out, "state": self.state()}

    def open_format_file(self, path):
        if os.name == "nt" and os.path.exists(path):
            subprocess.Popen(["notepad.exe", path])
            return True
        return open_path(path)

    def open_formats_folder(self):
        os.makedirs(check.USER_RULES, exist_ok=True)
        return open_path(check.USER_RULES)


def on_drop(api, e):
    files = (e.get("dataTransfer") or {}).get("files") or []
    paths = [f.get("pywebviewFullPath") for f in files]
    api._add(paths)
    api._window.evaluate_js(f"window.refresh({json.dumps(api.state())})")


def bind_drop(api):
    from webview.dom import DOMEventHandler
    doc = api._window.dom.document
    doc.events.dragover += DOMEventHandler(lambda e: None, True, True)
    doc.events.drop += DOMEventHandler(lambda e: on_drop(api, e), True, True)


def selftest(pdf, outdir):
    """Checks one PDF with no window, for testing the installed program."""
    log = os.path.join(outdir, "selftest.log")
    os.makedirs(outdir, exist_ok=True)
    try:
        item = batch.Item(pdf)
        b = batch.Batch([item], check.DEFAULT_RULES, outdir, False)
        b.run()
        summary = batch.export_summary(b, "self test")
        ok = item.status == "done" and os.path.exists(item.report) and os.path.exists(item.corrections) and os.path.exists(summary)
        with open(log, "w", encoding="utf-8") as fh:
            fh.write(json.dumps(item.to_dict(), indent=2) + "\n" + ("OK" if ok else "FAILED " + item.message) + "\n")
        return 0 if ok else 1
    except Exception:
        import traceback
        with open(log, "w", encoding="utf-8") as fh:
            fh.write(traceback.format_exc())
        return 1


def uitest(out):
    """Opens the real window, waits for the page to load and talk to the program, then closes it."""
    import time
    import webview
    os.makedirs(out, exist_ok=True)
    log = os.path.join(out, "uitest.log")
    result = {"ok": False}
    api = Api()
    page = os.path.join(check.BASE, "ui", "index.html")
    api._window = webview.create_window("Manuscript Check", page, js_api=api, width=1080, height=760)

    def probe(api):
        try:
            bind_drop(api)
            n = 0
            for _ in range(60):
                time.sleep(0.5)
                try:
                    n = api._window.evaluate_js("document.querySelectorAll('#format option').length")
                except Exception:
                    n = 0
                if n:
                    break
            label = api._window.evaluate_js("document.getElementById('start').textContent")
            result.update(ok=bool(n), formats=n, start_button=label)
        except Exception as e:
            result["error"] = repr(e)
        finally:
            with open(log, "w", encoding="utf-8") as fh:
                fh.write(json.dumps(result) + "\n")
            api._window.destroy()

    webview.start(probe, (api,), gui="edgechromium" if os.name == "nt" else None)
    return 0 if result["ok"] else 1


def main(shell=1, code_dir=None):
    if len(sys.argv) >= 5 and sys.argv[1] == "--word":
        kind, src, out = sys.argv[2:5]
        try:
            check.word_job(kind, src, out)
            sys.exit(0)
        except Exception:
            import traceback
            with open(out + ".err", "w", encoding="utf-8") as fh:
                fh.write(traceback.format_exc())
            sys.exit(2)
    if len(sys.argv) >= 3 and sys.argv[1] == "--uitest":
        sys.exit(uitest(sys.argv[2]))
    if len(sys.argv) >= 3 and sys.argv[1] in ("--updatetest", "--whichversion"):
        os.makedirs(sys.argv[2], exist_ok=True)
        if sys.argv[1] == "--whichversion":
            with open(os.path.join(sys.argv[2], "version.txt"), "w") as fh:
                fh.write(VERSION + "\n" + (code_dir or HERE) + "\n")
            sys.exit(0)
        import updater
        r = updater.check_and_download(shell)
        with open(os.path.join(sys.argv[2], "update.log"), "w") as fh:
            fh.write(json.dumps(r) + "\n")
        sys.exit(0 if r.get("status") in ("ready", "current") else 1)
    if len(sys.argv) >= 4 and sys.argv[1] == "--selftest":
        sys.exit(selftest(sys.argv[2], sys.argv[3]))
    if len(sys.argv) >= 2 and sys.argv[1] == "--version":
        print(VERSION)
        return
    import webview
    api = Api()
    api._shell = shell
    threading.Thread(target=api._check_updates, daemon=True).start()
    page = os.path.join(check.BASE, "ui", "index.html")
    api._window = webview.create_window("Manuscript Check", page, js_api=api, width=1080, height=760, min_size=(860, 600))
    webview.start(bind_drop, (api,), gui="edgechromium" if os.name == "nt" else None)


if __name__ == "__main__":
    main()
