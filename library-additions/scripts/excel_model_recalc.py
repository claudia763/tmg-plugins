"""Excel COM helpers for the TMG underwriting model on Windows (no LibreOffice).

What it does: recalculates the model with real desktop Excel, reads the key
Assumptions/Master output cells, and exports the 'PDF Output - F&C' tab as a
single-sheet PDF — avoiding the M365 stale-value strikethrough trap (see
instructions/excel-com-recalc-windows.md: calculation is ALWAYS returned to
Automatic before any export or save, otherwise every recalculated cell prints
with strikethrough in the PDF and the stale state survives SaveAs).

Inputs:  a TMG model .xlsx (openpyxl-edited is fine)
Outputs: recalculated .xlsx and/or a single-sheet PDF of PDF Output - F&C

Usage:
  python excel_model_recalc.py recalc <in.xlsx> <out.xlsx>   rebuild + save (clean)
  python excel_model_recalc.py read   <file.xlsx>            print key output cells
  python excel_model_recalc.py pdf    <in.xlsx> <out.pdf> [<out.xlsx>]
         one session: rebuild -> automatic -> export PDF -> optional SaveAs

Requires: pywin32 (pip install pywin32), desktop Excel.
"""
import os
import sys
import time

import win32com.client as win32
from win32com.client import constants as C

KEY_CELLS = [
    ("Assumptions", "F5",  "Project IRR"),
    ("Assumptions", "F7",  "Avg Cash-on-Cash"),
    ("Assumptions", "I8",  "T-3 DSCR"),
    ("Assumptions", "G48", "Target IRR"),
    ("Assumptions", "F50", "Price (Low)"),
    ("Assumptions", "G50", "Price (STRIKE)"),
    ("Assumptions", "H50", "Price (High)"),
    ("Assumptions", "G58", "Terminal Cap"),
    ("Assumptions", "G61", "LTV"),
    ("Assumptions", "G62", "Interest Rate"),
    ("Master",      "G33", "Total Units"),
    ("Master",      "H33", "Total SF"),
    ("Master",      "B22", "Occupancy"),
    ("Master",      "B24", "Total Tax Rate"),
]


def app():
    xl = win32.gencache.EnsureDispatch("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AskToUpdateLinks = False
    xl.ScreenUpdating = False
    return xl


def _open(xl, path):
    return xl.Workbooks.Open(os.path.abspath(path), UpdateLinks=0, ReadOnly=False)


def _rebuild_then_automatic(xl, passes=2):
    """Full rebuild under Manual, then RETURN TO AUTOMATIC (clears the M365
    stale-value strikethrough marks). Never export or save before this ran."""
    xl.Calculation = C.xlCalculationManual
    for i in range(passes):
        t0 = time.time()
        xl.CalculateFullRebuild()
        while xl.CalculationState != 0:      # 0 = xlDone
            time.sleep(0.5)
        print(f"  pass {i+1}: {time.time()-t0:.1f}s")
    xl.Calculation = C.xlCalculationAutomatic
    while xl.CalculationState != 0:
        time.sleep(0.5)
    time.sleep(1)


def _export_pdf(wb, dst, sheet="PDF Output - F&C", area="$B$1:$O$333"):
    ws = wb.Worksheets(sheet)
    ws.PageSetup.PrintArea = area
    ws.PageSetup.Orientation = 2             # landscape
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = False
    ws.ExportAsFixedFormat(0, os.path.abspath(dst))
    print("wrote", dst)


def recalc(src, dst):
    xl = app()
    try:
        wb = _open(xl, src)
        _rebuild_then_automatic(xl)
        wb.SaveAs(os.path.abspath(dst), FileFormat=51)
        wb.Close(SaveChanges=False)
        print("saved", dst)
    finally:
        xl.Quit()


def read(path):
    xl = app()
    try:
        wb = _open(xl, path)
        for sheet, cell, label in KEY_CELLS:
            try:
                v = wb.Worksheets(sheet).Range(cell).Value
            except Exception as e:
                v = f"ERR {e}"
            print(f"  {label:22s} {sheet}!{cell:5s} = {v}")
        wb.Close(SaveChanges=False)
    finally:
        xl.Quit()


def pdf(src, dst, save_xlsx=None):
    xl = app()
    try:
        wb = _open(xl, src)
        _rebuild_then_automatic(xl)
        _export_pdf(wb, dst)
        if save_xlsx:
            wb.SaveAs(os.path.abspath(save_xlsx), FileFormat=51)
            print("saved", save_xlsx)
        wb.Close(SaveChanges=False)
    finally:
        xl.Quit()


if __name__ == "__main__":
    cmd = sys.argv[1]
    if cmd == "recalc":
        recalc(sys.argv[2], sys.argv[3])
    elif cmd == "read":
        read(sys.argv[2])
    elif cmd == "pdf":
        pdf(sys.argv[2], sys.argv[3], sys.argv[4] if len(sys.argv) > 4 else None)
