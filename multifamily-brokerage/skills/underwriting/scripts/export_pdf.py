#!/usr/bin/env python3
"""Export a single sheet range of an .xlsx to PDF via a LibreOffice Basic macro.

Usage:
    python3 export_pdf.py <workbook.xlsx> <output.pdf> [sheet] [range]

Defaults: sheet="PDF Output - F&C", range="B1:O333".

Why: `soffice --convert-to pdf` exports every sheet (even hidden ones).
This installs a Basic macro into a throwaway LibreOffice profile and calls
storeToURL with FilterData Selection = the given sheet range, which exports
only those pages. Requires the xlsx skill's office/soffice.py helper on
sys.path (or plain `soffice` on PATH in an unsandboxed environment).
"""
import sys, tempfile, glob, os
from pathlib import Path
from xml.sax.saxutils import escape

# locate the xlsx skill's soffice helper (handles sandboxed environments)
for cand in glob.glob(os.path.expanduser('~/.claude/skills/xlsx/scripts')) + \
            glob.glob('/root/.claude/skills/xlsx/scripts'):
    sys.path.insert(0, cand)
try:
    from office.soffice import run_soffice
except ImportError:
    import subprocess
    def run_soffice(args, **kw):
        return subprocess.run(['soffice'] + list(args), **kw)

book = Path(sys.argv[1]).resolve()
out = Path(sys.argv[2]).resolve()
sheet = sys.argv[3] if len(sys.argv) > 3 else 'PDF Output - F&C'
rng = sys.argv[4] if len(sys.argv) > 4 else 'B1:O333'

MACRO = f'''<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">
    Sub ExportRangePDF()
      Dim oDoc, oSheet, oRange, aFD(0) As New com.sun.star.beans.PropertyValue
      Dim aArgs(1) As New com.sun.star.beans.PropertyValue
      oDoc = ThisComponent
      oSheet = oDoc.Sheets.getByName("{escape(sheet)}")
      oDoc.CurrentController.setActiveSheet(oSheet)
      oRange = oSheet.getCellRangeByName("{rng}")
      aFD(0).Name = "Selection"
      aFD(0).Value = oRange
      aArgs(0).Name = "FilterName"
      aArgs(0).Value = "calc_pdf_Export"
      aArgs(1).Name = "FilterData"
      aArgs(1).Value = aFD()
      oDoc.storeToURL("{out.as_uri()}", aArgs())
      oDoc.close(False)
    End Sub
</script:module>'''

profile = Path(tempfile.mkdtemp(prefix='lo_pdf_profile_'))
url = profile.as_uri()
run_soffice(['--headless', '--terminate_after_init', f'-env:UserInstallation={url}'],
            capture_output=True, timeout=120)
macro_dir = profile / 'user' / 'basic' / 'Standard'
assert macro_dir.exists(), 'LibreOffice profile bootstrap failed'
(macro_dir / 'Module1.xba').write_text(MACRO)

r = run_soffice([
    '--headless', '--norestore', f'-env:UserInstallation={url}', str(book),
    'vnd.sun.star.script:Standard.Module1.ExportRangePDF?language=Basic&location=application',
], capture_output=True, timeout=480)
print('rc', r.returncode)
if out.exists():
    print('wrote', out, out.stat().st_size, 'bytes')
else:
    sys.exit(f'export failed: {r.stderr!r}')
