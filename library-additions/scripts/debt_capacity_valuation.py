#!/usr/bin/env python3
"""
Debt-capacity / market-cap-rate valuation tests for a stabilized multifamily asset.

WHAT IT DOES
    Runs the three pricing tests TMG uses when a deal has no meaningful value-add
    and must "trade on a market cap rate basis and align with free and clear debt
    sizing" (Dmytro's framing, Westlake 8/2026):

      1. DEBT-CAPACITY TEST  - for each income basis (T-12 as reported, T-12 with
         normalized expenses, T-3, Year-1 underwritten): max debt service = NOI
         / DSCR, supportable loan = that / loan constant, supported price =
         loan / LTV. Also reports the implied LTV at a given asking price, which
         is what tells you whether DSCR or LTV is the binding constraint.
      2. RATE SENSITIVITY   - holds NOI/DSCR/LTV, varies the coupon, and reports
         the change in supportable price plus the DSCR at the asking price. Use
         this to find how many bps of cushion sit between full proceeds and a
         DSCR-constrained loan.
      3. CAP-RATE MATRIX    - value at a range of cap rates for each NOI basis,
         and the cap rates implied by a range of candidate prices.

    Also computes the loan constant from first principles and checks it against
    the model's printed monthly payment, and (optionally) reconciles an NOI
    bridge so the line items provably sum to the Year-1 NOI.

WHY THE LOAN CONSTANT MATTERS
    constant = (i/(1-(1+i)^-n)) * 12 on a monthly basis. Always verify it
    reproduces the underwriting model's printed monthly payment before trusting
    any downstream number -- at Westlake, 6.13%/30yr gave 7.2952% and $52,434/mo
    on $8,625,000, matching the model exactly, which validated the whole test.

NEGATIVE LEVERAGE
    The script prints cap rate minus loan constant in bps for each basis. A deal
    whose in-place cap sits below the constant is negative-leverage from day one,
    which is the defensible answer to "why can't buyers pay a lower cap rate."

INPUTS
    Edit the CONFIG block, or import and call the functions. No external deps.

OUTPUTS
    Formatted tables to stdout. Nothing is written to disk.

VALIDATED
    Westlake Apartments, Lubbock TX, 174 units, 8/6/2026. Reproduced the model's
    printed 7.85% / 7.15% / 8.56% cap rates and $52,434 payment exactly, and the
    NOI bridge tied to $1,045,473 / $984,573 to the dollar.
"""

# ---------------------------------------------------------------- CONFIG ----
UNITS       = 174
ASK         = 11_500_000        # price being tested
COUPON      = 0.0613            # underwritten coupon
AM_YEARS    = 30
DSCR_MIN    = 1.25
LTV_MAX     = 0.75
RESERVES    = 60_900            # annual replacement reserves (already in NOI below)

# NOI bases, AFTER reserves (TMG convention: cap rates are struck on NOI after reserves)
NOI_BASES = {
    "T-12 as reported":      902_980,
    "T-12 normalized opex":  827_918,
    "T-3 normalized opex":   822_774,
    "Year-1 underwritten":   984_573,
}

RATE_SCENARIOS = [0.0563, 0.0613, 0.0663, 0.0713]
CAP_SCENARIOS  = [0.0650, 0.0666, 0.0700, 0.0720, 0.0785, 0.0856]
PRICE_SCENARIOS = [10_690_719, 11_200_000, 11_500_000, 12_400_000]

# Optional NOI bridge: (label, amount). Must sum from the first entry to the target.
NOI_BRIDGE = [
    ("T-12 NOI (before reserves)",        963_880),
    ("GPR growth / loss-to-lease burnoff", 124_268),
    ("Economic loss normalization",        -32_341),
    ("Other income program",                64_728),
    ("Operating expense normalization",    -75_062),
]
NOI_BRIDGE_TARGET = 1_045_473


# ------------------------------------------------------------- FUNCTIONS ----
def loan_constant(rate: float, years: int = 30) -> float:
    """Annual loan constant for a fully amortizing monthly-pay loan."""
    i = rate / 12
    n = years * 12
    return (i / (1 - (1 + i) ** -n)) * 12


def monthly_payment(principal: float, rate: float, years: int = 30) -> float:
    return principal * loan_constant(rate, years) / 12


def debt_capacity(noi: float, rate: float, dscr: float = 1.25,
                  ltv: float = 0.75, years: int = 30) -> dict:
    """Price a financeable buyer can pay for a given NOI."""
    k = loan_constant(rate, years)
    max_ds = noi / dscr
    loan = max_ds / k
    return {"constant": k, "max_debt_service": max_ds,
            "loan": loan, "price": loan / ltv}


def main() -> None:
    k = loan_constant(COUPON, AM_YEARS)
    loan_at_ask = ASK * LTV_MAX
    print(f"Loan constant @ {COUPON:.2%}, {AM_YEARS}yr am : {k:.4%}")
    print(f"Monthly payment on ${loan_at_ask:,.0f}      : "
          f"${monthly_payment(loan_at_ask, COUPON, AM_YEARS):,.0f}"
          "   <- must match the model's printed payment\n")

    if NOI_BRIDGE:
        print("=== NOI BRIDGE ===")
        run = 0
        for label, amt in NOI_BRIDGE:
            run = amt if run == 0 else run + amt
            print(f"  {label:<38} {amt:>+12,}   running {run:>12,}")
        ok = run == NOI_BRIDGE_TARGET
        print(f"  {'TARGET':<38} {NOI_BRIDGE_TARGET:>12,}   ties={ok}")
        if not ok:
            print(f"  *** OFF BY {run - NOI_BRIDGE_TARGET:+,} - fix before publishing ***")
        print(f"  after ${RESERVES:,} reserves: {run - RESERVES:,}\n")

    print(f"=== DEBT CAPACITY  ({DSCR_MIN}x DSCR, {COUPON:.2%}, {LTV_MAX:.0%} LTV) ===")
    print(f"  {'Basis':<22}{'NOI':>11}{'Max DS':>11}{'Loan':>13}"
          f"{'Price':>14}{'$/Unit':>10}{'LTV@ask':>9}")
    for label, noi in NOI_BASES.items():
        r = debt_capacity(noi, COUPON, DSCR_MIN, LTV_MAX, AM_YEARS)
        print(f"  {label:<22}{noi:>11,.0f}{r['max_debt_service']:>11,.0f}"
              f"{r['loan']:>13,.0f}{r['price']:>14,.0f}"
              f"{r['price']/UNITS:>10,.0f}{r['loan']/ASK:>8.1%}")
    print("  (implied LTV above the max => DSCR has headroom and LTV is binding)\n")

    ds_amort = loan_at_ask * k
    ds_io = loan_at_ask * COUPON
    print(f"=== DSCR AT ${ASK:,} / {LTV_MAX:.0%} LTV "
          f"(amortizing ${ds_amort:,.0f} | IO ${ds_io:,.0f}) ===")
    for label, noi in NOI_BASES.items():
        print(f"  {label:<22} amortizing {noi/ds_amort:>6.3f} | IO {noi/ds_io:>6.3f}")
    print()

    print("=== NEGATIVE LEVERAGE CHECK (cap rate less loan constant) ===")
    for label, noi in NOI_BASES.items():
        cap = noi / ASK
        print(f"  {label:<22} cap {cap:>7.2%}  vs constant {k:.2%}  "
              f"= {(cap-k)*10000:>+7.0f} bps")
    print()

    print(f"=== RATE SENSITIVITY (NOI held, {DSCR_MIN}x, {LTV_MAX:.0%} LTV) ===")
    basis_label, basis_noi = list(NOI_BASES.items())[1]
    print(f"  basis: {basis_label} = ${basis_noi:,}")
    prev = None
    for r in RATE_SCENARIOS:
        d = debt_capacity(basis_noi, r, DSCR_MIN, LTV_MAX, AM_YEARS)
        delta = "" if prev is None else f"  change ${d['price']-prev:>+12,.0f}"
        print(f"  {r:>7.2%}  const {d['constant']:.4%}  loan ${d['loan']:>11,.0f}"
              f"  price ${d['price']:>12,.0f}  DSCR@ask {basis_noi/(loan_at_ask*d['constant']):.3f}{delta}")
        prev = d["price"]
    print()

    print("=== CAP RATE -> VALUE ===")
    print(f"  {'Cap':>7}" + "".join(f"{lbl[:16]:>18}" for lbl in NOI_BASES))
    for c in CAP_SCENARIOS:
        row = "".join(f"{noi/c:>18,.0f}" for noi in NOI_BASES.values())
        print(f"  {c:>7.2%}{row}")
    print()

    print("=== PRICE -> IMPLIED CAP RATES ===")
    print(f"  {'Price':>13}{'$/Unit':>10}" + "".join(f"{lbl[:16]:>18}" for lbl in NOI_BASES))
    for p in PRICE_SCENARIOS:
        row = "".join(f"{noi/p:>18.2%}" for noi in NOI_BASES.values())
        print(f"  {p:>13,}{p/UNITS:>10,.0f}{row}")


if __name__ == "__main__":
    main()
