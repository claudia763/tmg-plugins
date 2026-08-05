# The Haylie — processed 8/5/2026

Source: two ResMan PDFs from Touchstone Property Management, both printed
8/4/2026 — "Twelve Month Profit and Loss" (Sep 2025–Aug 2026, Accrual,
Accounting Book: Default, 4 pages) and "Rent Roll Summary" (as of 8/4/2026,
2 pages). 32 units in two buildings (unit numbers 724-1xx/2xx and
732-1xx/2xx); no street address, year built or story count is printed on
either report. Same PM and same two report types as McNeil Star — both
formats parsed natively, no code changes.

## T-12 (ResMan PDF parser, native)

- Reconciles clean: Revenue 469,421.43 / OpEx 191,235.98 / NOI 278,185.45.
  RawData sum-checks 0.00 on all three rows.
- **All twelve months carry operations** (unlike McNeil Star, whose
  Sep–Dec 2025 printed zeros). But **Aug 2026 is the in-progress month** —
  the statement was printed 8/4/2026, income is fully accrued (GPR 43,840,
  Total Revenue 43,317.46) while most expenses have not posted: water,
  trash, electric, real estate taxes, management fee, R&M and make-ready
  all print 0.00, so Aug OpEx is 2,880.46 vs a ~17,200 monthly average and
  Aug NOI 40,437.00 is roughly double a normal month. Called out in the
  Trailing Financials A2 header line — do not use Aug 2026 as a run-rate
  month, and expect the trailing NOI to fall ~14k once August posts.
- Revenue detail: GPR 520,905.00, Gain/Loss to Lease (13,427.31),
  Concessions (25,320.54), Vacancy Loss (43,205.08), Bad Debt (16,017.72)
  → Net Rental Income 422,934.35 (81.2% economic occupancy on GPR);
  Other Income 46,487.08, of which Building & Facilities Fee 26,481.32 and
  Late Fees 9,128.97 are the bulk.
- OpEx 191,235.98 = 5,976/unit. Heaviest lines: Real Estate Taxes 67,917.67,
  Water & Sanitation 28,911.63, Property Insurance 25,552.28, Salaries &
  Wages 22,587.30, Management Fee 14,897.69.
- Mapping: 60 exact / 3 fuzzy / 5 section / 1 keyword, 49 below-the-line
  lines excluded. **Only one REVIEW flag**, and it is the McNeil Star line:
  "App/Admin Fees Concession" (−250, Other Income - Fees) → `nr`
  (exact corpus hit, conflicts with its section but it is a concession) —
  kept as coded per the McNeil Star 8/4/2026 ruling.
- **No run-local corpus override was needed.** The two other McNeil rulings
  had nothing to bite on here: there is no income-side "Pest Control"
  billback (the Other Income - Contract/Billback section holds only
  Appliance Fee - W/D 1,085.00 and Liability Waiver 84.94, both `oi`, and
  Utility Billback-Water 428.67 sits in its own section → `rw`), and there
  is no "Property Tax Refund" line. Shared corpus untouched.
- Below the line → Capex & Misc workbook, 187,590.55 total, tying to the
  statement's TOTAL NON-OPERATING EXPENSE: Mortgage Interest 97,446.70,
  Other Non-Operating 9,093.35 (tax & accounting 4,293.35, legal 4,000,
  other 800), Capital Improvements 81,050.50 — a heavy in-place rehab
  (unit rehab labor 15,446, paint 6,666, appliances 6,894, plumbing 6,589,
  HVAC 3,777, resurface 3,605, carpets 3,186, hard-surface flooring 3,166,
  floor installation 3,139). Net Income after non-operating: 90,594.90.

## Rent roll (ResManSummaryParser, native — parser added for McNeil Star)

24/24 reconciliation checks pass against the report's grand strip, Property
Occupancy and Unit Type Occupancy tables (no credit-total check exists; the
per-unit printed Totals all tie to Rent + Other − Credits).

- 32 units, 25,600 sf: **1x1 700sf ×16, 2x1 900sf ×8, 2x1.5 900sf ×8**.
  31 occupied / 1 vacant (724-201, a 1x1) = 96.9% by unit, 97.3% by sqft,
  97.1% by market rent.
- Market rent 43,840 (1,370/unit, $1.71/sf); contract rent 40,720
  (1,313.55/occupied unit); lease charges 42,868.96; other charges 3,343.96;
  credits 1,195.00; deposits 4,050; resident balances 15,291.10; no surety
  bonds.
- **Unlike McNeil Star's A1/A2/B1, the Type codes here DO encode bed/bath**
  (1x1, 2x1, 2x1.5). Filled via `--bedbath "1x1=1/1,2x1=2/1,2x1.5=2/1.5"`,
  source cited as the ResMan unit-type code itself; sanity-checked against
  the printed sqft (700sf 1-bed, 900sf 2-bed). Sqft is in the roll and was
  never overridden.
- Status "UE" = under eviction → occupied door, resident status C, FLAGged
  (unit 732-202, LaDarius Stockland, balance 3,765.07 — the second-largest
  balance on the property).
- The "Credits" column is a lump with no charge code, so it routes through
  `classify_charge`'s unknown-negative rule: six credits of $50–200 became
  Recurring Concessions (col K, −960 total, inside NER) and one of $235
  (unit 724-107) exceeded the $200 threshold and became an Upfront
  Concession (col N, excluded from NER by house rule). All 1,195 is
  captured either way; only 724-107's NER differs (1,285 vs 1,050). The
  $235 looks like the same recurring monthly concession as the others
  (1,285 − 235 = 1,050) — worth Dmytro's call.
- Page 2 prints a Collections block (former resident "732-208 - Dennis
  Harris", 870.98) that is not part of the roll; it is correctly excluded
  and the balance total ties without it.
- Cross-check: **RR market rent 43,840.00 == T-12 Aug 2026 Gross Potential
  Rent 43,840.00** — the same tie that held at McNeil Star.

Deliverables: `RR - The Haylie - 8-4-2026.xlsx`, `T-12 - The Haylie -
August 2026.xlsx`, `Capex & Misc - August 2026.xlsx`.
