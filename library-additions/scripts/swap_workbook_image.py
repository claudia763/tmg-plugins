#!/usr/bin/env python3
"""Replace one embedded image inside an .xlsx, byte for byte, keeping its placement.

WHY NOT openpyxl: loading and re-saving a 19 MB, 67-sheet model just to move a picture
costs minutes and quietly drops things openpyxl does not round-trip. And the TMG model's
map sits in a `twoCellAnchor` (from col 8 / row 120 to col 13 / row 146) -- the frame,
not the pixel size, decides how big it draws. So the safe edit is to swap the bytes of
`xl/media/imageN.png` and leave every XML part, anchor and relationship untouched: the
new picture lands in exactly the same frame at exactly the same size.

WHY NOT the COM recipe in `comp-map-generation.md` s.3: that is
`ws.Shapes.AddPicture(...)` via Excel COM, which does not exist on the Linux agent box.
This is the Linux equivalent and it is strictly less invasive.

FINDING THE RIGHT imageN.png -- do not guess by pixel size, two pictures can match.
Trace sheet -> drawing -> rels:

    xl/workbook.xml                     <sheet name="..." r:id="rIdN">
    xl/_rels/workbook.xml.rels          rIdN -> worksheets/sheetN.xml
    xl/worksheets/_rels/sheetN.xml.rels -> ../drawings/drawingN.xml
    xl/drawings/_rels/drawingN.xml.rels rIdM -> ../media/imageM.png
    xl/drawings/drawingN.xml            which anchor carries r:embed="rIdM"

On the 8/2026 model, sheet `PDF Output - F&C` -> drawing5: rId4 = image4.png is the TMG
letterhead at col 2 / row 0, and rId5 = **image5.png is the sale-comp map** at col 8 /
row 120. The map ships loaded with the previous deal's geography and there is no formula
behind it, so it never updates on recalc -- it has to be replaced explicitly or the page
prints the wrong city under a correct comp grid.

Run this AFTER the final recalc: swapping an image does not change a single formula, so
it costs nothing and needs no recalc behind it.

Usage: python3 swap_workbook_image.py <in.xlsx> <out.xlsx> <xl/media/imageN.png> <new.png>
"""
import shutil
import sys
import zipfile


def main(src, dst, member, newpng):
    with open(newpng, "rb") as fh:
        blob = fh.read()

    zin = zipfile.ZipFile(src)
    if member not in zin.namelist():
        raise SystemExit(f"{member!r} not in {src}; have "
                         f"{[n for n in zin.namelist() if 'media/' in n]}")
    old = len(zin.read(member))

    tmp = dst + ".part"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = blob if item.filename == member else zin.read(item.filename)
            # preserve each part's own compression so nothing else is rewritten
            zi = zipfile.ZipInfo(item.filename, date_time=item.date_time)
            zi.compress_type = item.compress_type
            zi.external_attr = item.external_attr
            zout.writestr(zi, data)
    zin.close()
    shutil.move(tmp, dst)
    print(f"  {member}: {old:,} -> {len(blob):,} bytes")
    print(f"  wrote {dst}")


if __name__ == "__main__":
    main(*sys.argv[1:5])
