# McNeil Star Apartments — 3707 McNeil St, Dallas, TX 75227 — processed 8/4/2026

Source: two ResMan PDFs from Touchstone Property Management, both printed
8/3/2026 — "Twelve Month Profit and Loss" (Sep 2025–Aug 2026, Accrual) and
"Rent Roll Summary" (as of 8/3/2026). 32 units, 2 stories, built 1985.

## T-12 (ResMan parser, native)

- Reconciles clean: Revenue 205,322.17 / OpEx 99,549.88 / NOI 105,772.29.
- **Sep–Dec 2025 print as 0.00 in every row — operations begin Jan 2026, so
  the statement is effectively 8 months of operations on a 12-month axis.**
  Noted in the Trailing Financials header line; do not annualize casually.
- Mapping ruling: income-side "Pest Control" under "Other Income -
  Contract/Billback" → `ro` (grouped with Utility Billback-Water → rw and
  Utility Billback-Trash → rt; the corpus is split oi/ro on pest-control
  billbacks and its top hit is the expense-side `cs`, cross-ledger-rejected).
  Applied via run-local corpus override; shared corpus untouched.
- Two flags reviewed and kept as coded: "App/Admin Fees Concession" (−50) →
  nr (exact hit, conflicts with Other Income section but is a concession);
  "Property Tax Refund" (+28.98, Fixed Expenses) → tx (corpus said oi —
  wrong side, rejected).
- Below the line → Capex & Misc workbook, 87,129.58 total: Mortgage Interest
  29,923.33, Other Non-Operating 5,565.82 (professional fees / tax &
  accounting / start-up), Capital Improvements 51,640.43 (heavy rehab:
  HVAC 14.5k, unit rehab labor 9.3k, paint 6k, appliances 5.8k, flooring
  5.8k…). Net Income after non-operating: 18,642.71.

## Rent roll — NEW PARSER ADDED: ResManSummaryParser ("Rent Roll Summary")

The ResMan "Rent Roll Summary" is a one-row-per-unit condensation (no
per-charge blocks): Unit / Type / SqFt / Residents / Status / Market Rent /
Rent / Other Charges / Credits / Total / Move In / Start / Lease End / Move
Out / Surety / Deposits / Balance. `process_rent_roll.py` now has
`ResManSummaryParser`, registered BEFORE `ResManParser` (whose detect —
"ResMan" + "Rent Roll" — also matches a summary). Validated here with 25/25
reconciliation checks against the report's grand strip, Property Occupancy
and Unit Type Occupancy tables.

Key facts / conventions:

- 32 units: A1 498sf ×24, A2 597sf ×4, B1 791sf ×4. 27 occupied / 5 vacant
  (101, 103, 203, 207, 211). Market rent 31,780; contract rent 26,595;
  lease charges 27,372.99; deposits 5,310; balances 13,360.44.
- Bed/bath are NOT in the Type codes and the code digit lies (A2 at 597sf
  is a 1-bed). Web-sourced (apartmenthomeliving.com listing, sqft matched
  exactly): A1 = 1/1, A2 = 1/1, B1 = 2/2 → passed via --bedbath as sourced.
- Status "UE" = under eviction → occupied door, resident status C, FLAGged
  (unit 213 here, balance 4,390.99). "Other Charges" lump → Other Income;
  Credits (none here) → negative CREDIT charge through classify_charge.
- Cross-check: RR market rent 31,780 == T-12 Aug 2026 Gross Potential Rent.

Deliverables: `RR - McNeil Star - 8-3-2026.xlsx`, `T-12 - McNeil Star -
August 2026.xlsx`, `Capex & Misc - August 2026.xlsx`.
