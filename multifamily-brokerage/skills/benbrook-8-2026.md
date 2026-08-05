# Benbrook Apartments — 4639 Williams Rd, Benbrook, TX 76116 — processed 8/5/2026

Source: two PDFs. T-12 = QuickBooks Online "4639 Williams LLC / Profit and
Loss / July, 2025-June, 2026" (Cash Basis, printed 08/05/2026, 3 pages
landscape). Rent roll = Buildium-style "Rent Roll / Prepared By: Olivos
Management / As of 8/5/2026, 4639 Williams Road, Current leases" (3 pages
landscape). 40 units, 1979 build, 4.2 acres. Both formats were NEW — a
parser was added to each script.

## T-12 — NEW PARSER ADDED: `parse_t12_qbo_pdf` (QuickBooks Online P&L PDF)

`process_t12.py` now sniffs QuickBooks with `is_qbo_pdf()` ("Profit and
Loss" title + a 12× "Mon YYYY" header row + QuickBooks' "Total for X"
subtotals) and dispatches from both `main()` and `parse_t12_pdf()`.

Layout gotchas for future sessions:

- **Rows are SPARSE and values MUST be positional.** Most accounts print in
  only some months and QuickBooks prints *nothing* (not 0.00) in the others.
  Zipping tokens onto months shifts every gap month's money left, silently.
  Every numeric token is snapped to the month column whose printed RIGHT
  EDGE it matches (~45pt pitch, tol 12pt; the header token's right edge sits
  2-3pt left of the data's). A token that snaps nowhere **aborts** the parse.
  In-period blanks are read as $0.00 (a QBO blank = no transactions posted)
  and counted: 207 cells across 28 rows here.
- **The section tree is indentation and "Total for X" closes a section.**
  Top-level ~26pt, section ~32pt, account ~35pt — but the depth varies (the
  below-NOI Other Income block sits a level up), so an indent STACK is used,
  not fixed columns. Sections carry across page breaks (page 2 opens mid-way
  through Expenses with "Income Fees" — which is an EXPENSE section:
  merchant/bill-pay fees, not revenue).
- **Cost of Goods Sold sits between Income and Expenses**, so QuickBooks'
  printed "Total for Expenses" (190,871.68) EXCLUDES it. Canonical
  Total Operating Expenses = printed Total for Expenses + printed Total for
  Cost of Goods Sold = **192,586.68**, and that sum is *proved* by the
  printed NOI (466,687.15 − 192,586.68 = 274,100.47 = printed NOI). It is
  not a derived/self-referential number, so `reconcile()` accepts it.
- Everything after the printed Net Operating Income (QuickBooks'
  Other Income / Other Expenses block) is below-the-line → Capex & Misc.

Verification (all pass, nothing widened): every one of the 51 printed row
Totals ties to its own monthly cells; all 12 "Total for X" section subtotals
tie to their account detail per month and annually; Total for Income, Total
Operating Expenses, Gross Profit, NOI, Net Other Income and Net Income all
tie to the parsed detail.

- **Reconciles clean: Revenue 466,687.15 / OpEx 192,586.68 / NOI 274,100.47.**
- Below the line → `Capex & Misc - June 2026.xlsx`: Other Expenses
  131,149.44 (Construction Labor 67,351.70, Construction Supplies &
  Materials 55,878.80, Construction Miscellaneous 4,200, Construction
  Disposal 1,955.43, Construction Consultant Fees 1,763.51) less CapEx
  Income 9,800.00 (Jan 2026) → Net Other Income −121,349.44; Net Income
  152,751.03. This is a heavy in-place rehab, ~$3.3k/door of construction.
- Mapping — everything mapped cleanly against the shared corpus; **no
  mapping override was needed and `t12_mappings.csv` was NOT touched.**
  Rental Income→r, Convenience/Late/NSF/Other Income→oi, Utilities
  Reimbursement→ro (RUBS), Commissions (COGS)→pr, Contract Labor + Lawn
  Maintenance→cs, Onsite Maintenance + Supplies and Materials→rm, Admin /
  Legal & Professional / Melio + Merchant fees / QBO Subscription→ad,
  Google→m, Management Fees→mf, Insurance→i, Property Tax→tx,
  Electricity→e, Trash→tr, Water & Sewer→w.
- One REVIEW flag, reviewed and kept as coded: **"Merchant Services"
  [Income Fees] → `ad` (fuzzy 0.86)** — credit-card processing fees, admin
  is right. Worth a corpus harvest.
- Judgement to confirm: **"Commissions" (Cost of Goods Sold, 1,715.00 in
  Sep/Oct/Feb) → `pr`** on an exact corpus hit, which is silent (no REVIEW).
  Reads as leasing commissions; `pr` is the corpus majority. Watch this line
  on other QuickBooks books.

Statement quirks worth knowing before underwriting:

- **Fee Revenue and Utilities Reimbursement only exist from Jan 2026** —
  Jul-Dec 2025 print blank for Convenience/Late/NSF Fee and Utilities
  Reimbursement. The T-12 therefore books only 18,277.50 of utility billback
  where the current rent roll bills 3,252/mo (~39.0k/yr run-rate) and
  23,188.65 of total fee revenue on six months of billing.
- Dec 2025 carries the whole year's Insurance (20,544.80) and Property Tax
  (34,604.10) in one month (cash basis) → that month's NOI is −18,762.68.
- Jul 2025 Management Fees are 12,117.34 vs a ~2,600-3,300/mo run-rate
  elsewhere — looks like an onboarding/catch-up fee.

## Rent roll — NEW PARSER ADDED: `BuildiumRentRollParser`

Registered FIRST in `PARSERS`; `detect()` requires "Rent Roll" plus
Lease Start + Lease End + Bed/Bath + Prepayments + Balance Due in the page-1
header band (specific enough that it cannot steal ResMan/OneSite rolls —
regression-checked). Columns: Lease Start | Lease End | Bed/Bath | Rent
Cycle | Rent Start | Rent | Recurring Charges | Recurring Credits | Total |
Deposits Held | Prepayments | Balance Due.

Layout gotchas for future sessions:

- **There is NO unit identifier column** — the detail rows are anonymous.
  Verified against `page.chars`: no white/clipped text, nothing left of the
  Lease Start window; the only unit-level identity in the report is the
  group header. Units are numbered in REPORT ORDER (`Unit 01`..`Unit 40`)
  and that substitution is FLAGged. **These are not the property's real
  unit numbers** (public listings show real ones like 105/107/113/118).
- **Fake-bold rows print every glyph twice** — the group header
  "44663399 WWiilllliiaammss RRooaadd" and every totals row
  ("$$3366,,006655..0000"). `_undouble()` / `_row_undouble()` fix it; the
  test is deliberately WHOLE-ROW, because an isolated "4400" is itself a
  valid pair-doubled string and must never be halved to "40".
- **The Rent column is LEFT-aligned (x0 ≈ 336); every other numeric column
  is right-aligned.** Snapping Rent by right edge fails — its right edge
  moves 362→368 with the digit count and 385 on the bold totals row. Rent is
  claimed by an x0 window, the rest by right edge (445/504/564/623/683/758).
  A vacant unit prints "--" in that window.
- **"Rent Cycle: Monthly" is the BILLING cycle, not a month-to-month
  lease** — MTM is never inferred from it. **"Rent Start" is a DATE** (when
  the current rent schedule began), not a scheduled/market rent figure; it
  was checked explicitly. There is no status, tenant-name, move-in or
  move-out column.
- **Recurring Charges INCLUDES the Rent.** Other Income = Recurring Charges
  − Rent, booked as one lump charge (the report carries no per-charge
  detail): 3,252.00/mo across the 35 occupied units, ~74-126/unit.
- **Buildium ROUNDS its occupancy percentages half-up** (19/24 = 79.1666 →
  79.17). Do NOT reuse `reconcile()`'s `trunc2` (Yardi/OneSite truncate) —
  the property-level and per-plan % checks are filed as `extra_checks` with
  round-half-up.

Reconciliation — **33 checks, all tie**: unit count 40, occupied 35, vacant
5, 87.50% occupied, total sq ft 32,800, contract rent 36,065.00, lease
charges 39,317.00, deposits held 8,993.00, prepayments 2.00, balance due
8,649.60; the whole "Total for 4639 Williams Road" strip column by column;
per plan units/vacant/occupied/sq ft/% for 1/1 and 2/2; plus an independent
second pass (`extract_text` line splitting, a different code path from the
positional word parser) on unit count, occupied count and rent sum.

Unit mix and rents: **1/1 750 sf ×24** (19 occupied, 5 vacant — the only
vacancy in the property) and **2/2 925 sf ×16** (100% occupied). In-place
1/1 rents 800-1,030; 2/2 rents 900-1,365. Recurring credits are zero and no
concessions are printed anywhere, so NER = contract rent.

### Square footage — from the rent roll itself, not the web

The detail rows carry no sq ft, but the report's own last-page "Summary by
bed/bath" block does: 1/1 = 18,000 sf / 24 units, 2/2 = 14,800 sf / 16
units. **Both divide EXACTLY** (750 and 925), so the plans are uniform and
per-unit Net Sf is allocated from that block — rent-roll data, black text,
not a red-flagged estimate; the 32,800 total ties.

Independently confirmed on apartments.com listings for this address:
unit 105 = 1 bd / 1 ba / **750 sf**, unit 118 = 2 bd / 2 ba / **925 sf**.
Zillow's "total interior area 55,712 sf" (supplied by Dmytro) is a
gross-building figure — 32,800 sf is net rentable, and the roll wins. The
55,712 ÷ 40 = 1,393 sf/unit average was NOT used.

### Market rent — estimated (house rule 8/2026), and FLAGGED

The detail has no Market Rent column. The report's grand totals do print
**Market rent $13,665.00**, but that is an aggregate over only the units
that have one set in Buildium — 1/1: 7,700.00 total / 962.50 avg = 8 of 24
units; 2/2: 5,965.00 total / 1,193.00 avg = 5 of 16 units, 13 of 40 in all.
It cannot be allocated per unit, so it is **not** used as a tie-out target
and no market-rent reconciliation check is filed.

`--estimate-market` (max stated contractual rents, per plan, 6-month
lease-start window, highest rent repeating 3+ times) gave:

- **1/1 → $925.00** — 5 leases started on/after 02/05/2026 {925×2, 895×2,
  870×1}; nothing hit 3, so the fallback took the most-repeated (highest on
  the tie). **THIN — flagged.**
- **2/2 → $1,215.00** — only 2 leases in the window {1,215×1, 1,195×1} →
  highest. **THIN — flagged.**

Total estimated market rent 41,640/mo (24×925 + 16×1,215). Cross-checks:
the roll's own subset averages are 962.50 (1/1) and 1,193.00 (2/2); current
apartments.com asking rents at this address are **$925 for a 750 sf 1/1**
(exactly the estimate) and **$1,295 for a 925 sf 2/2** (the 2/2 estimate is
conservative by ~$80). Every estimated cell is red (FFC7CE / 9C0006) with
the note "highlighted market rent values are best estimates, not provided
by ownership".

## Toolkit changes made

- `process_t12.py`: `is_qbo_pdf()`, `parse_t12_qbo_pdf()`, `_qbo_rows()`,
  `_qbo_f()`, `_undouble`-free positional snapping; dispatched from `main()`
  and from `parse_t12_pdf()`. Additive — no existing parser touched
  (ResMan T-12 regression re-run clean).
- `process_rent_roll.py`: `BuildiumRentRollParser` (+ `_undouble`,
  `_is_doubled_row`, `_row_undouble`), registered first in `PARSERS`
  (ResMan Summary regression re-run clean).
- `process_rent_roll.py`: **`--estimate-market` / `estimate_market_rents()`
  and `_est_note_text()` were documented in CLAUDE.md (8/2026 house rule)
  but did NOT exist in the archived script — they are now implemented**,
  with `EST_COLS["market"] = 9` so estimated market-rent cells get the same
  red treatment, and the note under the data naming exactly the fields
  estimated in the run.

## Cross-checks between the two documents

- RR contract rent 36,065/mo × 12 = 432,780 vs T-12 Rental Income
  443,423.50 (cash basis, trailing months collected at then-current rents).
- RR other charges 3,252/mo (~39.0k/yr) vs T-12 Utilities Reimbursement
  18,277.50 — the T-12 only bills utilities from Jan 2026 (see above).
- 5 vacant doors (all 1/1) = 12.5% physical vacancy at 8/5/2026.

Deliverables: `RR - Benbrook - 8-5-2026.xlsx`, `T-12 - Benbrook -
June 2026.xlsx`, `Capex & Misc - June 2026.xlsx`.

## Open for Dmytro

1. **Unit numbers.** The roll is anonymous; the workbook says Unit 01-40 in
   report order. If Olivos can re-run the roll with the unit column on, the
   mapping should be redone before this goes in the model.
2. **Market rent.** Both plan estimates are THIN (5 and 2 data points).
   Alternatives: the roll's own subset averages (962.50 / 1,193.00) or
   current asking rents (925 / 1,295). Say the word and it re-runs.
3. **"Commissions" → `pr`** (silent exact corpus hit) — confirm it is
   leasing commissions and not a disposition/brokerage item.
