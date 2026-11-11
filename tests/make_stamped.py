"""Writes a copy of a PDF with a school stamp in the bottom margin of every page, the way the
graduate office adds 'Ref. code: 25666529040468VQK'."""
import io
import sys

from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas


def stamp(src, out, text="Ref. code: 25666529040468VQK"):
    r = PdfReader(src)
    w = PdfWriter()
    for page in r.pages:
        pw, ph = float(page.mediabox.width), float(page.mediabox.height)
        buf = io.BytesIO()
        c = canvas.Canvas(buf, pagesize=(pw, ph))
        c.setFont("Helvetica", 6.5)
        c.drawString(pw - 155, 16, text)
        c.save()
        buf.seek(0)
        page.merge_page(PdfReader(buf).pages[0])
        w.add_page(page)
    with open(out, "wb") as fh:
        w.write(fh)


if __name__ == "__main__":
    stamp(sys.argv[1], sys.argv[2])
