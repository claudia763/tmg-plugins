"""Restore the pristine template's chart XML parts (openpyxl re-serialization
flipped roundedCorners, dropped the transparent chart-area spPr and dash
styles on all 16 charts). Series ranges are identical; Excel refreshes the
cached values on the next full recalc."""
import shutil, zipfile

SRC = "CK St Nicholas Place 8-6-2026.xlsx"
PRISTINE = "model.xlsx"
DST = "model_v9.xlsx"

shutil.copyfile(SRC, DST)
zp = zipfile.ZipFile(PRISTINE)
pristine_charts = {n: zp.read(n) for n in zp.namelist() if n.startswith("xl/charts/chart") and n.endswith(".xml")}
zp.close()

zin = zipfile.ZipFile(SRC)
with zipfile.ZipFile(DST, "w", zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = pristine_charts.get(item.filename, None)
        if data is None:
            data = zin.read(item.filename)
        else:
            print("restored", item.filename)
        zout.writestr(item, data)
zin.close()
print("saved", DST)
