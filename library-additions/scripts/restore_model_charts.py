"""Restore the pristine TMG-model chart XML parts after openpyxl edits, then
apply the house chart-color fix.

Why: openpyxl re-serializes every chart, flipping roundedCorners, dropping the
transparent chart-area spPr and dash styles (grey rounded boxes, lost gallery
style). Additionally the template itself ships the "yellow dots error": the
Trended Revenue/Unit scatters carry hardcoded FFFF00 markers on both series
(house fix per Dmytro 8/6/2026: comp circles -> 4F81BD Style-1 blue, Subject
diamond -> FDB714 TMG gold). The printing scatters are chart2 (PDF Output -
F&C) and chart8 (PDF Output - Assume Loan); chart5/11/14/15 are mirrors.

Usage: python restore_model_charts.py <edited.xlsx> <pristine_template.xlsx> <out.xlsx>
Series ranges are identical between template and deal file; Excel refreshes
cached chart values on the next full recalc.
"""
import shutil, sys, zipfile

SRC, PRISTINE, DST = sys.argv[1], sys.argv[2], sys.argv[3]

CIRCLE_OLD = '<c:symbol val="circle"/><c:size val="5"/><c:spPr><a:solidFill><a:srgbClr val="FFFF00"/>'
CIRCLE_NEW = '<c:symbol val="circle"/><c:size val="5"/><c:spPr><a:solidFill><a:srgbClr val="4F81BD"/>'
DIAMOND_OLD = '<c:symbol val="diamond"/><c:size val="10"/><c:spPr><a:solidFill><a:srgbClr val="FFFF00"/>'
DIAMOND_NEW = '<c:symbol val="diamond"/><c:size val="10"/><c:spPr><a:solidFill><a:srgbClr val="FDB714"/>'

zp = zipfile.ZipFile(PRISTINE)
pristine_charts = {n: zp.read(n) for n in zp.namelist() if n.startswith("xl/charts/chart") and n.endswith(".xml")}
zp.close()

zin = zipfile.ZipFile(SRC)
with zipfile.ZipFile(DST, "w", zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = pristine_charts.get(item.filename)
        if data is not None:
            x = data.decode("utf-8")
            if CIRCLE_OLD in x or DIAMOND_OLD in x:
                x = x.replace(CIRCLE_OLD, CIRCLE_NEW).replace(DIAMOND_OLD, DIAMOND_NEW)
                print("restored + recolored", item.filename)
            else:
                print("restored", item.filename)
            data = x.encode("utf-8")
        else:
            data = zin.read(item.filename)
        zout.writestr(item, data)
zin.close()
print("saved", DST)
