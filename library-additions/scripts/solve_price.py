#!/usr/bin/env python3
"""Solve the strike price inside the TMG model itself (Linux, LibreOffice).

Linux counterpart of library-additions/scripts/model_price_solver.py (Excel COM).
Sets Assumptions!G50, recalculates via `soffice --convert-to` (the workbook carries
fullCalcOnLoad="1", so a plain convert recalcs everything -- ~20 s, and unlike the
macro/throwaway-profile route it does not deadlock on this box), then reads
F5 / F7 / I8 with data_only=True.

Green test (underwriting SKILL + aggressive-pricing house rule 8/6/2026):
  F5 Project IRR >= G48 target, F7 avg cash-on-cash >= 10%, I8 T-3 DSCR >= 1.25.
The house rule is to land the strike at the MAXIMUM price where all three hold.

Usage: python3 solve_price.py <model.xlsx> <lo|hi|step>  e.g. 1000000 1200000 25000
"""
import subprocess
import sys
import shutil
import warnings
from pathlib import Path

warnings.filterwarnings("ignore")
import openpyxl

HERE = Path(__file__).parent
SCRATCH = HERE / "solve"


def evaluate(src, price, tag):
    SCRATCH.mkdir(exist_ok=True)
    work = SCRATCH / f"p{tag}.xlsx"
    wb = openpyxl.load_workbook(src)
    wb["Assumptions"]["G50"] = price
    wb.save(work)
    out = SCRATCH / f"out{tag}"
    shutil.rmtree(out, ignore_errors=True)
    out.mkdir(parents=True)
    subprocess.run(
        ["soffice", "--headless", "--norestore",
         f"-env:UserInstallation=file:///var/tmp/losolve{tag}",
         "--convert-to", "xlsx", "--outdir", str(out), str(work)],
        capture_output=True, timeout=900,
    )
    got = out / work.name
    if not got.exists():
        return None
    v = openpyxl.load_workbook(got, data_only=True)
    a = v["Assumptions"]
    r = {
        "price": price,
        "irr": a["F5"].value,
        "coc": a["F7"].value,
        "dscr": a["I8"].value,
        "target": a["G48"].value,
        "mult": a["F6"].value,
        "cap": a["G58"].value,
        "noi": v["UW - F&C"]["AK38"].value,
        "loan": a["I5"].value,
        "equity": a["I6"].value,
    }
    shutil.rmtree(out, ignore_errors=True)
    work.unlink(missing_ok=True)
    return r


def green(r):
    return (r["irr"] is not None and r["irr"] >= r["target"]
            and r["coc"] >= 0.10 and r["dscr"] >= 1.25)


def main():
    src = sys.argv[1]
    lo, hi, step = int(sys.argv[2]), int(sys.argv[3]), int(sys.argv[4])
    print(f"{'PRICE':>11} {'$/unit':>9} {'IRR':>8} {'CoC':>8} {'DSCR':>7} "
          f"{'ExitCap':>8} {'Yr1 NOI':>10} {'PFcap':>7} GREEN")
    best = None
    for i, p in enumerate(range(lo, hi + 1, step)):
        r = evaluate(src, p, i)
        if r is None:
            print(f"{p:>11,}  <recalc failed>")
            continue
        g = green(r)
        if g:
            best = r
        print(f"{p:>11,} {p/19:>9,.0f} {r['irr']*100:>7.2f}% {r['coc']*100:>7.2f}% "
              f"{r['dscr']:>7.3f} {r['cap']*100:>7.2f}% {r['noi']:>10,.0f} "
              f"{r['noi']/p*100:>6.2f}% {'YES' if g else 'no'}")
    if best:
        print(f"\nMAX GREEN: ${best['price']:,} = ${best['price']/19:,.0f}/unit  "
              f"IRR {best['irr']*100:.2f}% (target {best['target']*100:.0f}%)  "
              f"CoC {best['coc']*100:.2f}%  DSCR {best['dscr']:.3f}  "
              f"EqMult {best['mult']:.2f}x")


if __name__ == "__main__":
    main()
