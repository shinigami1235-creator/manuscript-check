"""Lay out a Word file with LibreOffice, refresh the table of contents and lists, and save it as PDF.
Run with a Python that can import uno (LibreOffice's own python on Windows).
Usage: python lo_convert.py <soffice> <in.docx> <out.pdf>"""
import os
import shutil
import subprocess
import tempfile
import sys
import time
import uuid

import uno
from com.sun.star.beans import PropertyValue


def prop(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


def main(soffice, src, out):
    pipe = "mc" + uuid.uuid4().hex[:8]
    profile_dir = tempfile.mkdtemp(prefix="lo-profile-")
    profile = uno.systemPathToFileUrl(profile_dir)
    proc = subprocess.Popen([soffice, "--headless", "--invisible", "--nologo", "--norestore", f"-env:UserInstallation={profile}",
                             f"--accept=pipe,name={pipe};urp;"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
        ctx = None
        for _ in range(120):
            try:
                ctx = resolver.resolve(f"uno:pipe,name={pipe};urp;StarOffice.ComponentContext")
                break
            except Exception:
                time.sleep(0.5)
        if ctx is None:
            sys.exit("LibreOffice did not start.")
        desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(os.path.abspath(src)), "_blank", 0, (prop("Hidden", True),))
        idx = doc.getDocumentIndexes()
        for _ in range(2):  # twice, since refreshing the contents can move the pages it lists
            for i in range(idx.getCount()):
                idx.getByIndex(i).update()
            doc.getTextFields().refresh()
        doc.storeToURL(uno.systemPathToFileUrl(os.path.abspath(out)), (prop("FilterName", "writer_pdf_Export"),))
        doc.close(True)
        try:
            desktop.terminate()
        except Exception:
            pass
    finally:
        time.sleep(1)
        proc.kill()
        shutil.rmtree(profile_dir, ignore_errors=True)


if __name__ == "__main__":
    main(*sys.argv[1:4])
