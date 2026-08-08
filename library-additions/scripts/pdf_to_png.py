#!/usr/bin/env python3
"""Render every page of a PDF to PNG so a human/agent can LOOK at the layout.

Why: verify_exports.py only checks that comp names appear in the PDF's text
layer. ReportLab's fixed-width/fixed-height Table neither wraps nor clips a
plain string -- it OVERPRINTS the neighbouring cell -- and that failure is
completely invisible to a text-layer check. Render and eyeball on every run.
(sales-comps-from-an-address-only-8-2026.md 5, pipeline-hardening 5)

Usage: python3 pdf_to_png.py <in.pdf> [outdir] [--dpi 150]
Outputs: <outdir>/<pdfstem>-p1.png, -p2.png, ...
"""
import sys, os
import pypdfium2 as pdfium

def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    dpi = 150
    for a in sys.argv[1:]:
        if a.startswith("--dpi"):
            dpi = int(a.split("=", 1)[1]) if "=" in a else 150
    src = args[0]
    outdir = args[1] if len(args) > 1 else os.path.dirname(os.path.abspath(src))
    os.makedirs(outdir, exist_ok=True)
    stem = os.path.splitext(os.path.basename(src))[0]
    pdf = pdfium.PdfDocument(src)
    for i, page in enumerate(pdf, 1):
        out = os.path.join(outdir, f"{stem}-p{i}.png")
        page.render(scale=dpi / 72).to_pil().save(out)
        print(out)

if __name__ == "__main__":
    main()
