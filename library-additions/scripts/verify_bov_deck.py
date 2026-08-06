#!/usr/bin/env python
"""
verify_bov_deck.py -- post-render verification harness for `bov-deck` PDFs.

WHAT IT DOES
    Runs the three checks the bov-deck SKILL.md requires before delivery:
      1. Page count and page geometry (must be 1700x1080 px = 1275x810 pt).
      2. Substring check of EVERY key figure against the extracted text,
         after normalizing en/em dashes, curly quotes, arrows and whitespace
         (line wraps otherwise break naive matching).
      3. Owner-facing-language check: flags buyer-pitch phrases ("offered at",
         "offering price"), leftover template strings from the worked example
         (Aden Crest / Fort Worth), and any "Placeholder" label still visible
         after real photos were dropped in.
    It also writes one PNG per page so the pages can be eyeballed for text
    collisions and content running off the bottom -- the single most common
    defect when the template's prose is replaced with longer deal copy.

INPUTS
    pdf          path to the rendered <Property>_BOV.pdf
    figures      path to a UTF-8 text file, one expected figure/name per line
                 (blank lines and lines starting with '#' are ignored)
    --pngdir     directory for the page PNGs (default: <pdf dir>/pages)
    --dpi        PNG render dpi (default 72 -- enough to spot collisions)
    --extra-banned  comma-separated additional phrases that must NOT appear

OUTPUT
    Prints a pass/fail line per check and every missing figure / banned
    phrase. Exit code 0 if all figures matched and no banned phrase appeared,
    1 otherwise -- so it can gate a build.

USAGE
    python verify_bov_deck.py Landing_BOV.pdf figures.txt
    python verify_bov_deck.py Landing_BOV.pdf figures.txt --pngdir out/pages

NOTES
    - A "missing" figure is very often a rendering artifact, not a real error:
      an HTML entity such as &ndash; can extract as a separate run, giving
      "- 2.1%" where the check expects "-2.1%". Always grep the extracted text
      for the surrounding context before editing the deck.
    - Requires PyMuPDF (`pip install pymupdf`).
"""
import argparse
import os
import re
import sys

import fitz

# Phrases that must never appear in an owner-facing BOV.
DEFAULT_BANNED = [
    "offered at",         # buyer-pitch language; the BOV advises the owner
    "offering price",     # the offering price is unknown pre-engagement
    "asking price",
    "Aden Crest",         # bov_template.html worked-example leftovers
    "Fort Worth",
    "Placeholder",
    "placeholder",
]

EXPECTED_PT = (1275, 810)  # 1700x1080 css px at 0.75 pt/px


def normalize(text):
    """Fold the typographic characters that break naive substring matching."""
    for src, dst in (
        ("–", "-"), ("—", "-"),   # en / em dash
        ("’", "'"), ("“", '"'), ("”", '"'),
        ("→", "->"), ("†", "+"),  # right arrow, dagger
        (" ", " "),
    ):
        text = text.replace(src, dst)
    return re.sub(r"\s+", " ", text)


def load_figures(path):
    figures = []
    # utf-8-sig: PowerShell's Out-File/Set-Content write a BOM, which would
    # otherwise be glued to the first line and never match.
    with open(path, encoding="utf-8-sig") as fh:
        for line in fh:
            line = line.strip()
            if line and not line.startswith("#"):
                figures.append(line)
    return figures


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pdf")
    ap.add_argument("figures")
    ap.add_argument("--pngdir")
    ap.add_argument("--dpi", type=int, default=72)
    ap.add_argument("--extra-banned", default="")
    args = ap.parse_args()

    doc = fitz.open(args.pdf)
    w, h = doc[0].rect.width, doc[0].rect.height
    geom_ok = (round(w), round(h)) == EXPECTED_PT
    print(f"pages={len(doc)}  page size={w:.0f}x{h:.0f} pt "
          f"({'OK' if geom_ok else 'UNEXPECTED - expected 1275x810'})")

    raw = "".join(page.get_text("text") for page in doc)
    norm = normalize(raw)

    figures = load_figures(args.figures)
    missing = [f for f in figures if normalize(f) not in norm]
    print(f"figure checks: {len(figures) - len(missing)}/{len(figures)} pass")
    for item in missing:
        print("  MISSING:", item)

    banned = DEFAULT_BANNED + [b.strip() for b in args.extra_banned.split(",") if b.strip()]
    hits = [b for b in banned if b in raw]
    for b in hits:
        print("  BANNED PHRASE PRESENT:", b)
    if not hits:
        print("owner-facing language: OK (no banned phrases)")

    pngdir = args.pngdir or os.path.join(os.path.dirname(os.path.abspath(args.pdf)), "pages")
    os.makedirs(pngdir, exist_ok=True)
    for i, page in enumerate(doc):
        page.get_pixmap(dpi=args.dpi).save(os.path.join(pngdir, f"p{i + 1:02d}.png"))
    print(f"wrote {len(doc)} page PNGs to {pngdir} -- inspect every one for text "
          f"collisions and content cut off at the bottom")

    sys.exit(1 if (missing or hits) else 0)


if __name__ == "__main__":
    main()
