# SSI410 "Rent Roll Report" — parser notes and caption variants

Covers the SSI410 rent-roll PDF format handled by `SSI410Parser` in
`process_rent_roll.py`: which of the toolkit's built-in reconciliation checks
can silently degrade to informational on some SSI installs, how to get a real
tie-out anyway, and the layout quirks worth knowing before you judge a number.
Read this when you are processing an SSI410 roll — especially if the run's
reconciliation block prints `~ ... (no report total to check)` lines.

Written from Harvest Moon (Oak Leaf Management, 75 doors, as-of 6/1/2026).

## The problem: caption drift makes real checks go informational

`SSI410Parser._parse_checks()` harvests the report's own totals by caption
regex. SSI installs print the same blocks under different captions, so a
perfectly good parse can produce a reconciliation block with only two hard
checks:

```
OK  Unit count: parsed 75.00 vs report 75.00
OK  Total sq ft: parsed 69,382.00 vs report 69,382.00
~ Occupied units: 72.00 (no report total to check)
~ Current/On-Notice lease charges: 103,438.00 (no report total to check)
```

Nothing is mis-parsed — the totals simply weren't captured. The two variants
seen so far:

1. **Occupancy block captioned "Unit Analysis"**, not "Occupancy Status", and
   it prints `Description | Units | Percent` with **no sqft column**. The
   parser wants three groups (`Occupied\s+(\d+)\s+([\d,]+)\s+([\d,.]+)`) so it
   cannot match a two-number row. The block is also laid out in newspaper
   columns beside the income-code legend, so `extract_text` glues them:
   `MISCI AMENITY FEE L Leased OC Occupied Occupied 72 96.00`. Match it with
   an **end-of-line anchor** (`Occupied\s+(\d+)\s+[\d.]+\s*$`), not a
   start anchor.
2. **Charge summary captioned "(Current, On-Notice, Transfer Out residents
   only)"**, not "Current/On-Notice".

The iron rule is that every run shows a reconciliation block tying out. Do not
ship a roll on two checks. Either extend `_parse_checks()` with a registered
caption alternative, or run the verifier below.

## The verifier

`scripts/verify_ssi410_rentroll.py` (this library) reads the PDF's printed
grand totals directly — it does not reuse the parser, so it is a genuinely
independent second path rather than the parse checking itself:

```
python verify_ssi410_rentroll.py "<source.pdf>" "<RR - Property - M-D-YYYY.xlsx>"
```

Ten checks: unit count (twice, header and Unit Analysis), total Net Sf,
occupied, vacant, total market rent, contractual rent vs the BASE RENT code,
other income vs the non-RENT codes, total lease charges, and vacant-at-market.
Exits non-zero on any failure; validated by fault injection. Tolerance is $0.01
on money and exact on counts — never widen them.

Where to find the report's own numbers, all on the last two pages:

- Header band, every page: `75 Apts, 69,382 Sq. Ft.`
- `Unit Analysis`: Occupied / Vacant / Down / Total Units, plus Employee,
  Model, Construction, Other Use, Total Special Use.
- `Grand Total :` strip — four money columns in this order:
  **Market Rent · Actual Lease Rent · Gross Possible · Potential Charges**,
  then Security / Other / Total Deposits and Ending Balance.
- `Grand Summary of Actual Charges by Income Code` — per-code totals. `RENT`
  (BASE RENT) is the contractual-rent tie-out target; every other code sums to
  Other Income.

Useful identities that hold internally and are worth checking:

- Actual Lease Rent = BASE RENT + all other income codes.
- Gross Possible − Potential Charges = the vacant units priced at market rent.
  (Potential Charges is occupied-only and includes any rent-classified charge
  beyond BASE RENT — see the MODEL note below.)

## Layout and judgment notes

- **Vacant-leased (`VL`) and on-notice (`NL`) units print TWICE.** A `VL` unit
  prints a pre-leased `L` resident whose charges are all starred (`*` = not
  included in the summary of lease charges), then a separate `VACANCY` line.
  An `NL` unit prints the in-place `N` resident, then the incoming `L` lease.
  Resident priority `C > N > L > P` resolves both correctly: `VL` → Vacant with
  $0 rent, `NL` → Occupied with Vac. Notice = Yes. Cross-check the unit-header
  line count against the door count (Harvest Moon: 78 header lines = 75 units +
  3 duplicated).
- **A starred charge is never money.** Harvest Moon's only `MTM 200.00` in the
  whole file sat on a starred future lease on a vacant unit. That is not an
  in-place month-to-month resident, and MTM is never inferred — leave the MTM
  column empty (see `house-rules.md`).
- **`MODEL` and other odd income codes are Other Income, not rent.** The
  registered SSI410 rule is RENT → Contractual Rent, everything else → Other
  Income. SSI's own Potential Charges column disagrees — it folds a $5 MODEL
  charge into the unit's potential — which is exactly why Potential Charges
  ($93,263) exceeded BASE RENT ($93,258) at Harvest Moon. Follow the parser
  rule and tie to BASE RENT; flag the difference for the analyst rather than
  reclassifying.
- **Watch for a unit with two RENT charges.** Harvest Moon 01-115 printed
  `RENT 1,262.50` + `RENT 1,281.00` = $2,543.50, and the report's own Gross
  Possible / Total for that unit agreed. It is the report's number, not a parse
  error (a mid-term rent change or a double-billing artifact in SSI), but it
  inflates the plan average and the roll total by about one month's rent.
  Carry it through so the workbook ties, and raise it as an owner question.
- **SSI410 prints M/I Date and Lease Expires but no lease START date.** Lease
  Start Date stays blank for every unit, so the Floor Plan Summary's
  "# Leases by Lease Start Date" columns (Recent 2 / last 90 / 60 / 30 days)
  evaluate empty. That is correct, not a defect — never back-fill it from the
  move-in date. It also means `--estimate-market`'s recency filter would have
  to fall back to lease END dates on this format.
- Sq.Ft. and the Apt Type code (`Eff`, `1/1`, `2/2`, `3/2`) are both printed,
  and the type code DOES encode bed/bath here — no `--sqft` / `--bedbath` is
  needed, and `Eff` reads as 0 bed / 1 bath.
