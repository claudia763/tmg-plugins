"""Solve the TMG underwriting model's strike price AND leverage together against
the green rules, using Excel COM.

WHY: the model sizes the loan purely off Max LTV from its `Loan Terms` tab. On a
DSCR-constrained deal that produces a loan no lender would actually fund, and the
green test then fails on DSCR -- so a price-only search crushes the price to make
an unfinanceable loan work. The underwriting SKILL's house rule is the opposite:
drop leverage until DSCR is green. This sweeps LTV, finds the maximum green price
at each, and reports the LTV that supports the highest price. The curve is not
monotonic: below the peak the equity cheque throttles IRR/CoC, above it DSCR binds.

GREEN RULES (read from the workbook, not hardcoded -- the current template makes
the IRR target depend on recourse: Assumptions!G48 = 20% non-recourse, 25% recourse):
    Assumptions!F5 (Project IRR)   >= Assumptions!G48 (target)
    Assumptions!F7 (Avg CoC)       >= 0.10
    Assumptions!I8 (T-3 DSCR)      >= dscr-min (default 1.25)

DISTRESS WAIVER: if T-3 economic loss ('UW - F&C'!AC8+AC9+AC10) exceeds 30% the
deal is assumed to go bridge and the DSCR test is waived entirely -- pass
--no-dscr in that case (the script prints the T-3 economic loss so you can check).

INPUTS   --model <working model .xlsx>   (edited in place; run on a COPY)
         --ltvs 0.55,0.60,...            LTV grid to sweep
         --lo/--hi/--coarse/--fine       price search bounds and step sizes
OUTPUT   a table of max-green price per LTV; the winning combination is written
         back into Assumptions!G63 (LTV) and G50 (price) and saved.

NOTE Assumptions!G63 is a formula in the template (INDEX into `Loan Terms`).
Writing a literal here is a deliberate deal-specific override -- say so in the
delivery notes, and state the program's own Max LTV alongside it.

Requires: pywin32 + desktop Excel. ~1 s per recalc, so a 6 x 20 sweep is ~2 min.
"""
import argparse
import time
import warnings

warnings.filterwarnings("ignore")
import win32com.client as win32
from win32com.client import constants as C


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--model", required=True)
    ap.add_argument("--ltvs", default="0.55,0.60,0.65,0.70,0.75,0.80")
    ap.add_argument("--lo", type=int, default=1_000_000)
    ap.add_argument("--hi", type=int, default=50_000_000)
    ap.add_argument("--coarse", type=int, default=100_000)
    ap.add_argument("--fine", type=int, default=10_000)
    ap.add_argument("--coc-min", type=float, default=0.10)
    ap.add_argument("--dscr-min", type=float, default=1.25)
    ap.add_argument("--no-dscr", action="store_true",
                    help="waive the DSCR test (distressed deal, assumed bridge)")
    a = ap.parse_args()
    ltvs = [float(x) for x in a.ltvs.split(",")]

    xl = win32.gencache.EnsureDispatch("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AskToUpdateLinks = False
    try:
        wb = xl.Workbooks.Open(a.model, UpdateLinks=0)
        xl.Calculation = C.xlCalculationAutomatic
        A = wb.Worksheets("Assumptions")
        U = wb.Worksheets("UW - F&C")

        def calc():
            xl.CalculateFullRebuild()
            while xl.CalculationState != 0:
                time.sleep(0.2)

        def num(ref, ws=A):
            v = ws.Range(ref).Value
            return v if isinstance(v, (int, float)) else None

        calc()
        loss = [num(f"AC{r}", U) for r in (8, 9, 10)]
        if all(x is not None for x in loss):
            print(f"T-3 economic loss = {sum(loss):.2%} "
                  f"(vacancy {loss[0]:.2%} + concessions {loss[1]:.2%} + bad debt {loss[2]:.2%})"
                  + ("  -> >30%, consider --no-dscr" if sum(loss) > 0.30 else ""))

        def green():
            irr, tgt = num("F5"), num("G48")
            coc, dscr = num("F7"), num("I8")
            ok = (irr is not None and tgt is not None and irr >= tgt
                  and coc is not None and coc >= a.coc_min)
            if not a.no_dscr:
                ok = ok and dscr is not None and dscr >= a.dscr_min
            return ok, irr, tgt, coc, dscr

        results = []
        for ltv in ltvs:
            A.Range("G63").Value = ltv
            best, price = None, a.lo
            while price <= a.hi:
                A.Range("G50").Value = price
                calc()
                if green()[0]:
                    best = price
                elif best is not None:
                    break
                price += a.coarse
            if best is not None:
                while True:
                    trial = best + a.fine
                    A.Range("G50").Value = trial
                    calc()
                    if not green()[0]:
                        break
                    best = trial
                A.Range("G50").Value = best
                calc()
            ok, irr, tgt, coc, dscr = green()
            results.append((ltv, best, irr, coc, dscr))
            fm = lambda x, s: format(x, s) if isinstance(x, (int, float)) else "n/a"
            print(f"  LTV {ltv:>6.2%}  max green price "
                  f"{format(best, ',') if best else 'none':>12}  "
                  f"IRR {fm(irr, '.2%')}  CoC {fm(coc, '.2%')}  DSCR {fm(dscr, '.3f')}")

        good = [r for r in results if r[1]]
        if not good:
            raise SystemExit("no green price at any LTV -- revisit the Year-1 "
                             "assumptions or the value-add selections before repricing")
        win = max(good, key=lambda r: r[1])
        print(f"\nBEST: LTV {win[0]:.0%} at ${win[1]:,} "
              f"(IRR {win[2]:.2%}, CoC {win[3]:.2%}, DSCR {win[4]:.3f})")
        A.Range("G63").Value = win[0]
        A.Range("G50").Value = win[1]
        calc()
        wb.Save()
        wb.Close(SaveChanges=False)
    finally:
        xl.Quit()


if __name__ == "__main__":
    main()
