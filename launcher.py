"""Starts Manuscript Check from the newest program code on this computer.

The installer carries Python and the libraries (the shell) plus a copy of the program code.
Updates replace only the program code: they are unpacked into
%LOCALAPPDATA%\\Manuscript Check\\updates\\<version>\\ on Windows, or
~/Library/Application Support/Manuscript Check/updates/<version>/ on a Mac, and used from the next start.
An update that fails to start is marked bad and the previous code is used instead.
"""
import json
import os
import sys

SHELL = 1   # raise this when the installer itself must change, such as a new library


def _deps():
    # never called: lists what the program code imports so the installer packs it
    import argparse, copy, csv, ctypes, datetime, difflib, hashlib, html, logging, re, shutil, statistics  # noqa
    import subprocess, tempfile, threading, traceback, urllib.request, uuid, zipfile, io, glob, string  # noqa
    import pdfplumber, pdfminer, pypdf, pypdf.annotations, pypdf.generic, pypdfium2  # noqa
    import reportlab, reportlab.pdfgen.canvas, reportlab.platypus, reportlab.lib.styles  # noqa
    import openpyxl, openpyxl.styles, lxml.etree, PIL.Image, PIL.ImageOps  # noqa
    import webview, webview.dom  # noqa
    if os.name == "nt":
        import winreg, pythoncom, win32com.client  # noqa
    if sys.platform == "darwin":
        import platform, uuid, webview.platforms.cocoa  # noqa


def base_dir():
    return getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))


def bundled_code():
    d = os.path.join(base_dir(), "appcode")
    return d if os.path.isdir(d) else os.path.join(os.path.dirname(os.path.abspath(__file__)), "app")


def updates_root():
    if os.name == "nt":
        root = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    elif sys.platform == "darwin":
        root = os.path.join(os.path.expanduser("~"), "Library", "Application Support")
    else:
        root = os.path.join(os.path.expanduser("~"), ".local", "share")
    return os.path.join(root, "Manuscript Check", "updates")


def vtuple(v):
    try:
        return tuple(int(x) for x in str(v).split("."))
    except ValueError:
        return (0,)


def read_version(d):
    try:
        with open(os.path.join(d, "version.json"), encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return None


def candidates():
    """Program code folders this shell can run, newest first."""
    out = []
    b = read_version(bundled_code())
    out.append((vtuple(b["version"] if b else "0"), bundled_code(), True))
    root = updates_root()
    if os.path.isdir(root):
        for name in os.listdir(root):
            d = os.path.join(root, name)
            v = read_version(d)
            if v and not os.path.exists(os.path.join(d, "bad")) and int(v.get("shell", 1)) <= SHELL and os.path.exists(os.path.join(d, "gui.py")):
                out.append((vtuple(v["version"]), d, False))
    return sorted(out, reverse=True)


def main():
    if getattr(sys, "frozen", False) and sys.platform == "darwin":
        sys.dont_write_bytecode = True   # writing .pyc files inside the app would break its signature
    for _, code, bundled in candidates():
        sys.path.insert(0, code)
        try:
            import gui
            return gui.main(shell=SHELL, code_dir=code)
        except SystemExit:
            raise
        except Exception:
            sys.path.remove(code)
            for m in ("gui", "batch", "check", "fix", "learn", "updater"):
                sys.modules.pop(m, None)
            if not bundled:
                with open(os.path.join(code, "bad"), "w") as fh:
                    fh.write("did not start\n")
                continue
            raise


if __name__ == "__main__":
    main()
