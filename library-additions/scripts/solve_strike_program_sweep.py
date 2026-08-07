#!/usr/bin/env python3
"""Solve the TMG strike price by sweeping LOAN PROGRAMME x LTV, in pure Python.

WHEN TO USE
    Linux / headless, where scripts/model_price_solver.py (Excel COM) cannot run,
    and you already have a per-deal port of the underwriting model as a module
    exposing run_model(a=...) and ASSUMPTIONS -- i.e. the underwriting skill's
    scripts/"tmg valuation.py" copied into the job folder and refilled.
    ~0.02 s per evaluation, so a full 10-programme x 20-LTV sweep is seconds,
    versus ~1 s per recalc through Excel or ~20 s through LibreOffice.

WHY PROGRAMME **AND** LTV
    The template sizes the loan off the selected programme's Max LTV
    (Assumptions!G64 = INDEX('Loan Terms'!F,...)), so a price-only search is
    really a search at one fixed leverage. On a DSCR-constrained deal that is
    the wrong leverage: coverage binds, and the only way to raise the price is
    to drop leverage. Sweeping both is what finds the real ceiling.
    The curve is not monotonic in LTV -- below the peak the equity cheque
    throttles IRR and cash-on-cash, above it DSCR binds.

THE THREE TRAPS THIS SCRIPT EXISTS TO AVOID
  1. TARGET IRR IS NOT A CONSTANT. Assumptions!G48 is
         =IF($G$63="Non-Recourse",$K$48,$M$48)      K48=0.20  M48=0.25
     so it is 20% on agency/HUD/LifeCo/CMBS and 25% only on bank/credit-union
     recourse paper. references/model-map.md's flat "F48:H48 = 0.25" describes
     the Low/High scenario columns, NOT the strike column the green logic reads.
     Hardcoding 25% on an agency deal silently crushes the price (~$500k on a
     $6.4M deal).
  2. G63 IS RECOURSE, G64 IS MAX LTV. Writing leverage into G63 both leaves
     leverage unchanged AND destroys the text G48 keys off, flipping the target
     from 20% to 25%. Two silent errors compounding in the same direction.
  3. EACH PROGRAMME CARRIES ITS OWN MINIMUM DSCR ('Loan Terms' col G): HUD
     1.176x, LifeCo 1.35x, everything else 1.25x. Size to the STRICTER of that
     and TMG's 1.25 house rule or you quote a loan the lender would not make.
     And model HUD ALL-IN -- note rate plus the ~0.60% annual MIP. On the bare
     note rate HUD looks like the cheapest execution in the book; it is not.

THE STRIKE RULE  (instructions/aggressive-pricing-house-rule-8-2026.md, which
supersedes the underwriting SKILL's "2-4 point cushion"):
    the MAXIMUM price at which  IRR >= target, avg CoC >= 10%, DSCR >= floor.
    No cushion. And no excess DSCR -- if coverage sits above the floor at the
    ceiling price, leverage is too low, not the price.
DISTRESS WAIVER: if T-3 economic loss (vacancy + concessions + bad debt) > 30%
the deal is assumed to go bridge and the DSCR test is waived entirely.

CREDIBILITY IS STILL A HUMAN CALL. The top solve is not automatically the
answer. A life company will not quote 96 doors of 1963 Class C product however
well it scores -- its own Loan Terms note reads "stabilized Class A/B". Pass
`credible=` to mark which programmes may win, and report the excluded ones
alongside so the choice is visible rather than buried.

USAGE
    import tmg_valuation as T
    from solve_strike_program_sweep import sweep, PROGRAMS_AUG2026
    rows = sweep(T, PROGRAMS_AUG2026, credible={"Fannie Mae - Conventional", ...})
"""
from __future__ import annotations

# ('Loan Terms' rows 4-16 as quoted Aug 2026. all_in_rate already includes HUD MIP.)
# name, all_in_rate, max_ltv, min_dscr, io_yrs, term_mo, amort_mo, recourse
PROGRAMS_AUG2026 = [
    ("Fannie Mae - Conventional",  0.0624, 0.75, 1.25,  3, 120, 360, "Non-Recourse"),
    ("Freddie Mac - Conventional", 0.0619, 0.75, 1.25,  3, 120, 360, "Non-Recourse"),
    ("Fannie Mae - Small Balance", 0.0674, 0.80, 1.25,  1, 120, 360, "Non-Recourse"),
    ("Freddie Mac - SBL",          0.0669, 0.80, 1.25,  1, 120, 360, "Non-Recourse"),
    ("HUD/FHA - 223(f) all-in",    0.0639, 0.85, 1.176, 0, 420, 420, "Non-Recourse"),
    ("Life Company (LifeCo)",      0.0604, 0.65, 1.35,  2, 120, 360, "Non-Recourse"),
    ("CMBS / Conduit",             0.0719, 0.75, 1.25,  5, 120, 360, "Non-Recourse"),
    ("Bank - Balance Sheet",       0.0665, 0.70, 1.25,  1,  60, 300, "Recourse"),
    ("Credit Union",               0.0675, 0.75, 1.25,  0,  60, 360, "Recourse"),
]
# Bridge and mezzanine rows are deliberately absent: they are floating-rate,
# full-term-IO (amort 0) instruments the model's pmt() cannot amortise, and they
# only apply when the G62 occupancy fallback (<75%) actually fires.

TARGET_NON_RECOURSE, TARGET_RECOURSE = 0.20, 0.25
COC_MIN, DSCR_HOUSE = 0.10, 1.25


def _target(prog):
    return TARGET_NON_RECOURSE if prog[7] == "Non-Recourse" else TARGET_RECOURSE


def _dscr_floor(prog, waive):
    return 0.0 if waive else max(DSCR_HOUSE, prog[3])


def sweep(T, programs=PROGRAMS_AUG2026, credible=None, ltv_lo=0.55,
          ltv_step=0.01, lo=1_000_000, hi=50_000_000, step=10_000,
          waive_dscr=False, units=None):
    """Return rows sorted by strike, descending.

    T        the per-deal model module (needs run_model + ASSUMPTIONS)
    credible set of programme names allowed to win; None = all
    waive_dscr  set True when T-3 economic loss > 30% (deal goes bridge)
    """
    units = units or T.PROPERTY["units"]
    out = []

    def run(price, prog, ltv):
        a = dict(T.ASSUMPTIONS)
        a.update(purchase_price=price, interest_rate=prog[1], ltv=ltv,
                 io_years=prog[4], loan_term_months=prog[5],
                 amortization_months=prog[6])
        return T.run_model(a=a)

    def green(r, prog):
        return (r["project_irr"] >= _target(prog)
                and r["avg_cash_on_cash"] >= COC_MIN
                and r["t3_dscr"] >= _dscr_floor(prog, waive_dscr))

    for prog in programs:
        best = None
        ltv = ltv_lo
        while ltv <= prog[2] + 1e-9:
            if green(run(lo, prog, ltv), prog):
                a_lo, a_hi = lo, hi
                while a_hi - a_lo > step:
                    mid = (a_lo + a_hi) / 2
                    if green(run(mid, prog, ltv), prog): a_lo = mid
                    else: a_hi = mid
                p = int(a_lo // step * step)
                if best is None or p > best[0]:
                    best = (p, ltv, run(p, prog, ltv))
            ltv = round(ltv + ltv_step, 6)
        if best:
            p, ltv, r = best
            out.append({
                "programme": prog[0], "strike": p, "per_unit": p / units,
                "ltv": ltv, "rate": prog[1], "irr": r["project_irr"],
                "target": _target(prog), "coc": r["avg_cash_on_cash"],
                "dscr": r["t3_dscr"], "dscr_floor": _dscr_floor(prog, waive_dscr),
                "equity_multiple": r["equity_multiple"],
                "loan": r["loan_amount"], "equity": r["equity_required"],
                "y1_noi": r["noi"][1], "going_in_cap": r["noi"][1] / p,
                "terminal_cap": r["terminal_cap"],
                "credible": credible is None or prog[0] in credible,
                "result": r,
            })
    out.sort(key=lambda d: -d["strike"])
    return out


def report(rows):
    print(f"{'Programme':<28}{'LTV':>5}{'Rate':>7}{'tgt':>5}{'Strike':>12}{'$/unit':>9}"
          f"{'IRR':>7}{'CoC':>7}{'DSCR':>6}{'EM':>6}{'Cap':>7}  ")
    print("-" * 104)
    for d in rows:
        print(f"{d['programme']:<28}{d['ltv']:>5.0%}{d['rate']:>7.2%}{d['target']:>5.0%}"
              f"{d['strike']:>12,}{d['per_unit']:>9,.0f}{d['irr']:>7.2%}{d['coc']:>7.2%}"
              f"{d['dscr']:>6.2f}{d['equity_multiple']:>6.2f}{d['going_in_cap']:>7.2%}"
              f"{'' if d['credible'] else '   [not credible - excluded]'}")
    win = next((d for d in rows if d["credible"]), None)
    if win:
        print(f"\nSTRIKE ${win['strike']:,} (${win['per_unit']:,.0f}/unit) — "
              f"{win['programme']} @ {win['ltv']:.0%} LTV, {win['rate']:.3%}")
        print(f"  IRR {win['irr']:.2%} (target {win['target']:.0%}) · "
              f"CoC {win['coc']:.2%} · DSCR {win['dscr']:.3f} "
              f"(floor {win['dscr_floor']:.3f}) · EM {win['equity_multiple']:.2f}x")
        print(f"  loan ${win['loan']:,.0f} · equity ${win['equity']:,.0f} · "
              f"Y1 NOI ${win['y1_noi']:,.0f} ({win['going_in_cap']:.2%}) · "
              f"exit cap {win['terminal_cap']:.2%}")
        spread = [d["strike"] for d in rows if d["credible"]]
        if len(spread) > 1:
            print(f"  credible executions span ${min(spread):,}-${max(spread):,} — "
                  f"a tight spread means the strike is not an artefact of programme choice")
