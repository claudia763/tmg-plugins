#!/usr/bin/env python3
"""Strip every externalLink part from an .xlsx so LibreOffice can open it headless.

WHY: the TMG underwriting model template carries 8 `xl/externalLinks/externalLink*.xml`
parts pointing at a copy of the template that does not exist on this machine.
`lo_model.py prep` (per underwriting/references/technical-notes.md SS1-2) removes the
external *defined names* and rewrites external-link *formulas* to their cached values,
but it does NOT remove the link parts themselves. LibreOffice still tries to resolve
them on load and blocks -- headless, with no dialog to answer, soffice.bin sits at
~0% CPU forever and the driving macro never runs. (Observed on Shady Oaks 8/7/2026:
14 s of CPU in 10 minutes, then killed.) Excel does not hang because it silently
falls back to the cached values.

Run this AFTER `lo_model.py prep` and before any LibreOffice recalc/export.

INPUT   an .xlsx (normally the prep() output)
OUTPUT  the same workbook with all externalLink parts, their relationships,
        their [Content_Types].xml overrides and the workbook.xml
        <externalReferences> element removed; every other part copied byte for byte.

Usage:  python3 strip_external_links.py <in.xlsx> <out.xlsx>
"""
import re
import shutil
import sys
import zipfile


def strip(src, dst):
    ext_re = re.compile(r"^xl/externalLinks/")
    with zipfile.ZipFile(src) as zin:
        names = zin.namelist()
        ext_parts = [n for n in names if ext_re.match(n)]
        if not ext_parts:
            print("no externalLink parts found; copying unchanged")
            shutil.copyfile(src, dst)
            return

        # relationship ids in xl/_rels/workbook.xml.rels that point at them
        # Attribute order varies between writers, so match the whole element on
        # its Target and count what actually came out rather than assuming.
        rels = zin.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        rel_re = re.compile(r'<Relationship\b[^>]*externalLinks/[^>]*/>')
        n_rels = len(rel_re.findall(rels))
        rels_out = rel_re.sub("", rels)
        assert "externalLinks/" not in rels_out, "externalLink relationship survived"

        wb = zin.read("xl/workbook.xml").decode("utf-8")
        wb_out = re.sub(r"<externalReferences>.*?</externalReferences>", "", wb, flags=re.S)

        ct = zin.read("[Content_Types].xml").decode("utf-8")
        ct_out = re.sub(r'<Override[^>]*PartName="/xl/externalLinks/[^"]*"[^>]*/>', "", ct)

        with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                n = item.filename
                if ext_re.match(n) or n.startswith("xl/externalLinks/_rels/"):
                    continue
                if n == "xl/_rels/workbook.xml.rels":
                    zout.writestr(item, rels_out)
                elif n == "xl/workbook.xml":
                    zout.writestr(item, wb_out)
                elif n == "[Content_Types].xml":
                    zout.writestr(item, ct_out)
                else:
                    zout.writestr(item, zin.read(n))

    print(f"removed {len(ext_parts)} externalLink part(s), "
          f"{n_rels} relationship(s); wrote {dst}")


if __name__ == "__main__":
    strip(sys.argv[1], sys.argv[2])
