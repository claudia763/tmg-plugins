#!/usr/bin/env python3
"""Sweep Assumptions!G50 through the TMG model and report the green tests.

Each step: write G50 with openpyxl, `soffice --convert-to xlsx` (the workbook
carries fullCalcOnLoad="1", so a plain convert is a full recalc), then read
F5 (Project IRR) / F7 (avg cash-on-cash) / I8 (T-3 DSCR) with data_only=True.

Green (underwriting SKILL + aggressive-pricing-house-rule-8-2026.md):
  F5 >= G48 target (20% on non-recourse agency debt, 25% recourse)
  F7 >= 0.10
  I8 >= 1.25
House rule: take the MAXIMUM price where all three hold, snapped to a $10,000
boundary (the PDF ROUNDs the printed Sales Range to -4, so an off-boundary
strike breaks the cost stack on page 1).

Usage: python3 sweep_price.py <model.xlsx> <price> [<price> ...]
Prints one TSV line per price.
"""
import shutil
import subprocess
import sys
import warnings
from pathlib import Path

warnings.filterwarnings("ignore")
import openpyxl

SCRATCH = Path(__file__).parent / "sweep"


def evaluate(src, price):
    SCRATCH.mkdir(exist_ok=True)
    work = SCRATCH / f"p{price}.xlsx"
    wb = openpyxl.load_workbook(src)
    wb["Assumptions"]["G50"] = price
    # Arkansas does NOT reassess to sale price on transfer (Amendment 79 has no
    # change-of-ownership reset), so taxes must be driven by the COUNTY's own
    # appraised value, not the strike. The model computes
    # taxes = G39 x total rate x price, so G39 has to track the price:
    #   G39 = CAD full market value / price  ->  taxes = rate x CAD value.
    cad = wb["Master"]["B9"].value
    if cad:
        wb["Assumptions"]["G39"] = round(cad / price, 6)
    wb.save(work)
    out = SCRATCH / f"o{price}"
    shutil.rmtree(out, ignore_errors=True)
    out.mkdir(parents=True)
    subprocess.run(
        ["soffice", "--headless", "--norestore",
         f"-env:UserInstallation=file:///var/tmp/losweep{price}",
         "--convert-to", "xlsx", "--outdir", str(out), str(work)],
        capture_output=True, timeout=900)
    got = out / work.name
    if not got.exists():
        return None
    v = openpyxl.load_workbook(got, data_only=True)
    a, f = v["Assumptions"], v["UW - F&C"]
    r = dict(price=price, irr=a["F5"].value, coc=a["F7"].value,
             dscr=a["I8"].value, target=a["G48"].value,
             noi=f["AM38"].value, loan=a["I5"].value, equity=a["I6"].value,
             cap=a["C10"].value, termcap=a["G58"].value, mult=a["F6"].value)
    v.close()
    shutil.rmtree(out, ignore_errors=True)
    work.unlink(missing_ok=True)
    return r


def main():
    src = sys.argv[1]
    print("price\tIRR\tCoC\tDSCR\ttarget\tNOI\tPFcap\tloan\tequity\tgreen")
    for p in [int(x) for x in sys.argv[2:]]:
        r = evaluate(src, p)
        if not r:
            print(f"{p}\tCONVERT FAILED")
            continue
        ok = (r["irr"] is not None and r["irr"] >= r["target"]
              and r["coc"] is not None and r["coc"] >= 0.10
              and r["dscr"] is not None and r["dscr"] >= 1.25)
        print(f"{r['price']:,}\t{r['irr']:.4f}\t{r['coc']:.4f}\t{r['dscr']:.4f}\t"
              f"{r['target']:.2f}\t{r['noi']:,.0f}\t{r['cap']:.4f}\t"
              f"{r['loan']:,.0f}\t{r['equity']:,.0f}\t{'GREEN' if ok else 'red'}",
              flush=True)


if __name__ == "__main__":
    main()
