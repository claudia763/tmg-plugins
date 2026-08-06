# Magnolia Place — 301 & 303 Goessler St, Brenham, TX 77833 — processed 8/6/2026

Source: two owner-made XLSX workbooks, no PM system anywhere in either.
T-12 = "Magnolia Place Financials (1).xlsx", an operating statement kept as
**one tab per calendar year** (`2025`, `2026`) — neither tab is a trailing
twelve on its own. Rent roll = "Magnolia Place July 2026 RR.xlsx", where each
unit occupies a **block of 5-7 rows** and the charge codes under the rent are a
standing empty menu. 20 units, 14,900 sf, 1.102 acres, built 1971 (CAD
confirmation pending — one source said 1989). Both formats were NEW; a
registered parser was added to each script.

Windows: RR as of **7/31/2026** (assigned — the sheet prints no date), T-12
**Jul 2025 – Jun 2026** (stitched).

## T-12 — NEW PARSER ADDED: `parse_t12_owner_calendar_year_tabs`

`process_t12.py` sniffs the dialect with `_is_owner_calendar_year_tabs`
(registered FIRST in `T12_XLSX_PARSERS`): two or more sheets named as 4-digit
years, `A1 == 'Category'`, `B1..M1` = January..December, `N1` an annual total.
Helpers `_owner_cy_year_sheets`, `_owner_cy_parse_sheet`, `_owner_cy_key`,
`_owner_cy_num`, `_GRAND_PATS_OWNER_CY`. Full write-up in
`library-additions/instructions/owner-calendar-year-tabs-t12.md`.

Layout gotchas for future sessions:

- **The annual-total caption lies.** The *2026* tab captions column N
  `Total 2025` — a copy-paste from the prior year. It is really the 2026
  total. Detection and windowing key off the **tab name plus the
  January–December header row**, never the caption; the parser prints
  `CAPTION DOES NOT MATCH THE TAB YEAR` when it sees this. Same family of bug
  as Yardi titling a six-month export "Statement (12 months)".
- **Find the window end from OPERATING data only.** The 2026 tab is pre-filled
  with hard zeros for months not yet posted, but the below-the-line `Mortgage`
  row is pre-filled for all twelve. "Last month with any value" would claim
  December and silently pad six empty months into the T-12. Only rows above
  the printed NOI are tested → window ends **Jun 2026**, correctly.
- **ALL-CAPS is not a section marker.** `MANAGEMENT FEES`, `UTILITIES` and
  `PAYROLL` are ALL-CAPS single-line *accounts* carrying values, sitting
  alongside genuine ALL-CAPS headers (`RENTAL INCOME`, `CONTRACT SERVICES`,
  `REPAIRS & MAINTENANCE`, `TAXES & INSURANCE`…). Classify structurally — does
  deeper detail follow? — never by capitalisation.
- **Accounts drift between tabs.** Merged (and reported, never silent):
  `'Registered Agent Fees '` == `'Registered Agent Fees'` (whitespace only)
  and `'Restripping Parking Lot'` == `'Restriping Parking Lot'` (0.978, CapEx).
  3 accounts absent from the 2025 tab and 12 from the 2026 tab get
  **genuinely blank** months on the other tab's span — never zeros.
- **Not auto-merged, deliberately:** `Maintenance` (1,800.00, Aug–Dec 2025)
  and `Grounds Cleaning` (1,250.00, Jan–Jun 2026) read like the same recurring
  vendor renamed at the year turn, but the statement prints two names, so they
  stay two lines and the run notes say so.

### The stitch, and how it is proved

The trailing twelve crosses calendar years — **Jul–Dec 2025 from tab `2025`,
Jan–Jun 2026 from tab `2026`** — so there is no single printed annual total to
tie to. Four verification layers, all of which passed before anything was
written:

- **[a] Row check** — every row's twelve monthly cells vs its printed col-N
  annual total, over each *full calendar year*, before any windowing:
  tab 2025 **96 rows tie**, tab 2026 **87 rows tie**.
- **[b] Printed roll-ups vs their own detail**, per tab, all twelve months:
  TOTAL RENTAL INCOME, TOTAL INCOME, TOTAL EXPENSES, NET OPERATING INCOME —
  **12/12 each, both tabs**.
- **[c] `Total CapEx` vs capex detail** — 12/12 months and annual, both tabs
  (2025: 30,144.66 over 35 lines; 2026: 13,382.53 over 31 lines).
- **[d]** the same four identities re-run on the **stitched window** (12/12
  each — this is what actually proves the stitch), then the windowed grand
  rows vs the sum of the statement's own printed grand rows over those same
  twelve months.

- **Reconciles clean: Revenue 231,194.34 / OpEx 80,468.29 / NOI 150,726.05**
  (variance +0.00 on all three, from both the [d] check and `reconcile()`).
  34.8% expense ratio; **$11,560 / $4,023 / $7,536 per door**.
- Below the line → `Capex & Misc - June 2026.xlsx`, 96,619.14: Total CapEx
  30,555.54 (Unit Renovations 10,471.09, Plumbing Repairs 5,916.77, Appliances
  3,822.03, Int. paint 2,218.28, Landscaping 1,600, Exterior LED 1,542.37,
  Electrical Panel 1,485, Concrete Walkway 1,200, Pressure Washing 1,000,
  Restriping 500, Moisture Remediation 500, locks 300) plus Mortgage 66,063.60.

### `Units Rehabbed` — a memo row, excluded and reported

The 2026 tab carries a `Units Rehabbed` row that repeats the CapEx line
`Unit Renovations` **cell for cell** (Jan-2026 5,181.00 on both, zero in every
other month of both tabs). It is a memo, not a second cost; counting it would
double January's renovation spend. Removed with `--exclude-account` +
`--exclude-reason` so the removal, the account, the side, every monthly value
and the reason are printed **and written to the workbook**.

### Charge codes — five overrides, through a run-local corpus copy

Applied with `--mappings work/t12_mappings_magnolia.csv`; the shared
`t12_mappings.csv` was **not** touched:

| Account | → | Why |
| --- | --- | --- |
| Landscaping Maintenance | `cs` | recurring $200-400/mo inside the statement's own CONTRACT SERVICES section; house `landscap` keyword rule agrees. Overrides a 2-vote corpus `rm`. |
| Maintenance | `cs` | recurring **$300/mo** vendor line inside CONTRACT SERVICES (Aug/Oct/Dec-25 at 300, one 900 catch-up in Sep-25; 1,800 in the window). Overrides corpus-majority `pr` (16-13). |
| Trash Hauling | `tr` | a hauling contract is Trash; the section rule would otherwise code it `cs` silently. Zero in the window. |
| Evictions | `bd` | sits in RENTAL INCOME as an eviction-loss adjustment, not fee income. Zero in the window. |
| Resident Referalls *(owner's spelling)* | `nr` | a referral CREDIT against rental income — concession, not marketing expense. Zero in the window. |

**The bare "Maintenance" call went the OPPOSITE way from Royal Oaks 8/2026,
on purpose.** Royal Oaks' `Maintenance` was lumpy, job-sized, with wages
broken out separately → `rm`. Magnolia's is a **flat $300 every month, sitting
inside the statement's own CONTRACT SERVICES section, with a separate REPAIRS &
MAINTENANCE section carrying all the job-sized work (Make Ready, Plumbing,
A/C, Appliances, Rekey…) and a separate $0 PAYROLL line** → `cs`. **Section placement and recurrence pattern decide this
line, not the label** — the label is identical in both deals. Both rulings are
in the run-local corpus copies with their reasoning; neither is in the shared
corpus, and the exact corpus hit stays silent (no REVIEW flag), so this line
must be inspected by hand on every owner-books statement.

Mapping outcome: 36 exact, 15 section, 2 fuzzy, 2 keyword. Four REVIEW flags,
all reviewed and kept: "Overpayment Feturn" [RENTAL INCOME] → `oi` (fuzzy
0.82, owner's typo for Return); "Eviction Fees" [OTHER INCOME] → `oi` (corpus
said `ad` — wrong side, rejected); "Leasing Commissions" [G&A] → `pr`;
"Trash Hauling" [CONTRACT SERVICES] → `tr` (the override above, flagged
because it conflicts with its section).

### Statement quirks worth knowing before underwriting

- **Taxes and insurance are already accrued evenly month by month**, so
  `--prorate-bulk` was NOT applied. The owner accrues $2,320/mo of tax
  ($27,840/yr) — at Brenham's 1.6638 composite that implies an assessed value
  near **$1.67M**, far above the last published $517,000 (flat for 2019-2021).
  Do not underwrite taxes off $517,000; a sale likely triggers reassessment.
- Utilities are only **$3,249/yr ($162/door)** and there is no RUBS or
  reimbursement income at all — consistent with tenants paying all utilities.
- Laundry income (2,694 over the twelve) **starts in Sep-2025** — machines look
  newly installed or newly metered. Do not annualize the T-12 figure.
- Three unlabelled cells belonging to no account row were ignored and reported
  (2025!P81 = 46,580; 2025!N134 = 47,229.73; "Unit 9" memo text at 2025!M130,
  2026!B117 and 2026!M118), plus two free-text memos beside the 2025 Repairs &
  Maintenance rows describing amounts pushed into January. **No amounts were
  moved.**

## Rent roll — NEW PARSER ADDED: `OwnerBlockRentRollXlsxParser`

Registered FIRST in `XLSX_PARSERS`. Header (not on row 1 — a "July 2026 Rent
Roll" title sits above it): `Unit Type | Unit # | Sq Ft | Occupancy | Resident
| Market Rent | Charge Code | Charges | Move-in Date | Lease Start | Lease End
| Deposit | Deposit Notes | Renewal | NTV`. The block's first row carries the
unit, sqft, occupancy, first resident, market rent, the `Rent` charge, the
dates and the deposit; every following row carries only a charge code and
sometimes a second name, a phone number, an extra Unit-Type word or an extra
deposit. Full write-up in
`library-additions/instructions/owner-block-rent-roll-xlsx.md`.

Layout gotchas for future sessions:

- **The charge menu is not charges.** Every block prints `Pet` ×20, `MTM` ×20,
  `NEW CREDIT` ×20, `MISC. FEE` ×19 (plus one typo'd `MISC, FEE`), `Late Fee`
  ×18 and `Patio` ×2 with **empty amounts** — a template the owner fills only
  when something is actually billed, and one `Late Fee` printed with an
  explicit `0.00` (still a menu row). Only `Rent` ever carries money. Booking
  these as $0 charges would invent charges the sheet does not make: Other
  Income, concessions and discounts here are genuinely **nil**, not zeroed. A
  reconciliation check counts non-rent charges booked (must be 0) so this
  cannot regress silently.
- **The Unit Type column holds more than the unit type.** It also carries an
  **UNLABELLED DATE** on each block's first row — 2025-01-01 on every 500 sf
  unit, 2025-02-01 on every 850 sf unit, i.e. correlated with *size*, not with
  the unit. It is not written to Floor Plan and is FLAGged for the requester
  (market-rent effective date? renovation date?) rather than guessed at.
- **The Resident column sometimes holds a PHONE NUMBER** — units 2, 8, 12 and
  15 (four of twenty). Detected, excluded from the names, and reported by unit.
- **A deposit can be split across two rows** — unit 12 is $500 + $200 = $700,
  summed, and that is what makes the printed 13,650 tie. Conversely occupied
  **unit 5 has a blank Deposit cell and stays BLANK, not zero**; the sheet's
  own total excludes it, which is the proof.
- **A literal word in the Lease End cell** — unit 16 prints `Monthly`. That is
  the source *stating* month-to-month, so MTM = Yes and Lease Expiration is
  left blank. The amount-less `MTM` menu row was **not** used as an MTM signal
  (the house rule bars *inferring* MTM). FLAGged either way.
- **No as-of date and no property name anywhere** — `asof_found = False`, so
  `--asof` is mandatory and the run hard-exits rather than guessing.
  `--property "Magnolia Place" --asof 2026-07-31` were both passed.
- RENEWAL and NTV are printed but empty for all 20 units — no unit on notice.

Reconciliation — **14 checks, all tie**: unit count 20, occupied 20, total
sq ft 14,900, market rent 20,900.00, contract rent 20,145.00, deposits
13,650.00, current lease charges 20,145.00; rent charges booked 20 (one per
occupied unit); **non-rent charges booked 0** and **Other Income booked 0**.
The sheet's totals row prints only market rent, charges and deposits — no unit
count and no sqft — so following the `OwnerSheetPdfParser` precedent, unit
count, occupied count and sqft come from an **independent second pass that
reads the worksheet XML straight out of the zip container**, sharing no code
path with the block parser, and are filed as "vs re-extract" checks. Fault
injection (+$1 on a market rent, inventing unit 5's missing deposit, booking
the $0 menu row) makes the block fail as it should.

**100% occupied, 20/20.** In place 20,145 vs market 20,900 — **3.6% below
market**, unusually tight for a small owner-run asset.

### Bed/bath — web-sourced, and how

The source carries **no bed/bath at all**, and "Unit Type" is a renovation
tier (PREMIUM / CLASSIC / PARTIAL), not a floor plan. "Magnolia Place" is not
independently indexed; the property was found by **floor-plan match** —
searching for a Brenham TX complex with exactly a 500 sf and an 850 sf plan at
the roll's rent levels — which led to 301 & 303 Goessler St, where an
apartments.com listing then names the community explicitly.

- **500 sf = 1 bed / 1 bath.** Zumper building p142643 (directly fetched)
  lists "303 Goessler Street – 1" as 1bd/1ba/500 sf. Corroborated on PadMapper.
- **850 sf = 2 bed / 1 BATH — not 2/2.** The same Zumper page lists units 11,
  18 and 19 as 2bd/1ba/850 sf; corroborated by ApartmentGuide and the
  apartments.com 301 Goessler listing. **2/1 rather than 2/2 matters for exit
  rent and comp selection.**
- Independent corroboration: the county **improvement area 14,792 sf** vs the
  **14,900 sf** rentable implied by 6×500 + 14×850 is **+0.73%**, which
  confirms both the door count and the 6/14 size split. Corroboration only —
  the workbook keeps the rent roll's own sqft.

Passed via `--bedbath`; the cells are black, not red — this is sourced data,
not an estimate.

### Floor plans — tier in the name AND in Renovation Status

Plans are named `{bed}x{bath} {Tier}`, so each spans exactly one sqft (asserted
in verification) and the tier rent spread reads straight off the Floor Plan
tab:

| Plan | # | Sq Ft | Market | In place | Units |
| --- | --- | --- | --- | --- | --- |
| 1x1 Premium | 3 | 500 | 800–900 | 900–950 | 1, 3, 9 |
| 1x1 Classic | 2 | 500 | 800 | 825–900 | 2, 8 |
| 1x1 Partial | 1 | 500 | 900 | 975 | 7 |
| 2x1 Premium | 8 | 850 | 1,200 | 1,075–1,100 | 4, 5, 6, 10, 11, 13, 17, 19 |
| 2x1 Classic | 4 | 850 | 1,000 | 875–970 | 12, 15, 16, 20 |
| 2x1 Partial | 2 | 850 | 1,100 | 1,100–1,115 | 14, 18 |

The tier is ALSO written to the per-unit **Renovation Status** column (Rent
Roll col G, named range `RenovationString`) — leaving it only inside the plan
name hides it from the rediQ import and from anyone filtering the Rent Roll
tab. Because every unit in a plan shares its tier, Floor Plan Summary col C
("Renovated") populates for every plan and flows to the Floor Plan tab. Both
places; plans are **not** collapsed to bare `1x1` / `2x1`.

**Patio units folded into their base plans.** Units 13 and 15 carry a further
`PATIO` word on a lower row (enclosed patios), but they are at the **same
market rent as their tier-mates** — the amenity carries no premium. Giving it
its own plan would create one-unit plans whose averages are meaningless, so it
is folded in and kept as a run FLAG that reaches the delivery notes. The call
is a class-level constant `PATIO_AS_PLAN = False` — a one-line switch if a
future roll prices patios separately.

## Toolkit changes made

- `process_t12.py`: `parse_t12_owner_calendar_year_tabs` /
  `_is_owner_calendar_year_tabs` (+ `_owner_cy_year_sheets`,
  `_owner_cy_parse_sheet`, `_owner_cy_key`, `_owner_cy_num`,
  `_GRAND_PATS_OWNER_CY`), registered first in `T12_XLSX_PARSERS`.
- `process_t12.py`: `SECTION_ALLOWED` gains a combined **`Taxes & Insurance`**
  rule → `{i, tx}`, tested BEFORE the plain `^tax|taxes` rule (same shape as
  the existing `FIXED ADMINISTRATIVE` precedent). Without it, property
  insurance filed under a combined `TAXES & INSURANCE` head fires a spurious
  REVIEW.
- `process_t12.py`: **`write_capex()` now honours `Line.empty`**, so a
  genuinely blank capex month stays blank instead of printing 0.00
  (`write_workbook` already did this).
- `process_t12.py` **bug fix:** `--exclude-account` removal reports were
  printed to console but **never written into the workbook**. They now land on
  the **Comments** tab, as the 8/6/2026 no-red-text house rule requires. New
  repeatable **`--note`** flag writes methodology / data-quality notes to the
  same tab.
- `process_rent_roll.py`: `OwnerBlockRentRollXlsxParser`, registered first in
  `XLSX_PARSERS`; plus per-unit Renovation Status output (Rent Roll col G and
  Floor Plan Summary col C) and two provenance `key=` additions in
  `reconcile()` (`total_sqft`, `total_deposits`).
- `supported-formats.md`: new sections documenting both dialects.

**Regression gate:** `parser_detection_regression.py` over 41 XLSX files
against all 9 registered detectors (4 rent-roll, 5 T-12) — every file claimed
by at most one detector, `OwnerBlockRentRollXlsxParser` and
`_is_owner_calendar_year_tabs` each claim their own file and nothing else, and
no existing detector claims either Magnolia file. All checks pass.

## Cross-checks between the two documents

- **RR market rent 20,900/mo × 12 = 250,800 == T-12 Gross Potential Rent
  250,800 to the dollar.** Two independently parsed documents in two different
  new dialects landing on the same GPR is the strongest cross-check on the
  pair, and it independently confirms the 20-door count and the market-rent
  column.
- RR in-place 20,145/mo × 12 = 241,740 vs T-12 net rental income 227,077.77
  (GPR 250,800 less loss-to-lease 12,887.55 and vacancy loss 10,834.68). The
  roll is the current run-rate at 100% occupancy; the T-12 is a trailing year
  that included turns and a wider loss-to-lease.
- The roll books **no other income at all** (empty charge menu) while the T-12
  carries 4,116.57 of it — laundry 2,694 (from Sep-2025), transfer fees 465,
  deposit forfeit 757.57, late fees 200. Those are property-level events, not
  standing resident charges, so this is not a discrepancy.
- 0 vacant doors at 7/31/2026 vs 10,834.68 (4.3% of GPR) of vacancy loss over
  the trailing twelve.

Deliverables: `RR - Magnolia Place - 7-31-2026.xlsx`,
`T-12 - Magnolia Place - June 2026.xlsx`, `Capex & Misc - June 2026.xlsx`,
plus `Magnolia Place - Property Research Notes.md`.

## Open for Dmytro

1. **The unlabelled date in the Unit Type column** (2025-01-01 on the 500 sf
   units, 2025-02-01 on the 850 sf units). Correlates with size, not unit.
   Market-rent effective date? Ask ownership.
2. **Property name.** 303 Goessler is also marketed publicly as **"Heritage
   Place"** (Yelp, Facebook, and Zumper serves the same building ID under both
   slugs). Best read: Heritage Place is the legacy/operating name, Magnolia
   Place the ownership-side name for the same 20 doors — but no document states
   a rename. Confirm before this goes on marketing material.
3. **Washington CAD, (979) 277-3740, account R19947 / 0124-000-00810.** One
   call resolves current assessed value (the $517,000 on record is flat
   2019-2021 and stale against a $27,840/yr accrual), year built (1971 vs
   1989), building count, stories and owner of record.
4. **Unit 5's blank deposit** — genuinely not collected, or a gap in the
   sheet? It is left blank, and the printed deposit total agrees with blank.
