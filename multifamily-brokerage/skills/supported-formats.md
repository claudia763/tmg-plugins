# Supported source formats

Per-format parser documentation, extracted near-verbatim from TMG's toolkit
instructions plus the 8/2026 deal notes. All formats are auto-detected:
rent rolls via `PARSERS` (PDF) / `XLSX_PARSERS` in `process_rent_roll.py`,
T-12s via PDF sniffers and `T12_XLSX_PARSERS` in `process_t12.py`. New
format? Add a registered parser — never one-off processing.

## Rent rolls

### SSI410 "Rent Roll Report" PDF (`SSI410Parser`)

Positional column parser. Resident priority C > N > L > P; NA/notice ->
Vac. Notice "Yes"; charge mapping: RENT -> Contractual Rent,
concession/discount codes -> cols K/N/O, all other charges -> Other Income.
NER = rent + recurring conc + discounts; upfront concessions are reported
but excluded from NER (house rule). Unknown negative charges: |x| > $200 ->
upfront, else recurring.

### AppFolio XLSX export (`AppFolioXlsxParser`)

Header row Unit/BD/BA/Tenant/Status... No sqft in this format; 1970-01-01
dates are placeholders -> blank; lease start dates populate col Q. MTM is
NOT inferred from expired leases. Multi-property exports (repeated header +
per-property group rows) are supported; pass `--property` for the combined
name. "Include Advertised Rentals" placeholder rows (marketing names like
`2Bed/2.5Bath Luxury...`, no tenant) are auto-excluded and backed out of the
report totals — this is why Clark/Pecan shows 154 doors vs the export's
printed 158. `--sqft "1/1.5=1100,2/2.5=1600"` fills Net Sf by floor plan
when sourced externally (cite the source in the delivery summary). Keys that
aren't floor plans match the unit name instead (exact, then substring) and
win over the floor-plan value — for scattered-site portfolios where one
floor plan spans differently-sized properties (e.g. `Holly=1636`).
Single-unit properties whose detail row has an empty Unit cell (identity
carried by the group header, e.g. Holly Ln / Wren Rowe duplex halves) are
named from the group short name — do not let them drop silently (84-door
M White portfolio, 7/2026).

### Yardi/ResMan "Rent Roll" XLSX (`YardiRentRollXlsxParser`)

Validated on The Meadows/tmtx. Two-row column header joined per column;
`Current/Notice/Vacant Residents` IS the door count, and rows under `Future
Residents/Applicants` are applicants for units already listed above —
attached as status-`L` residents, never counted as extra doors (a vacant
unit with an applicant becomes VR). Floor plan = the unit-type code
(`tm1x1a`); the codes are NOT interchangeable (sqft differs per code) so
they are kept verbatim, and `UnitRecord.bed_bath` reads the NxN inside the
code. Actual Rent maps to one RENT charge; there is no charge detail, so
Other Income/concessions stay blank, as do Lease Start and MTM (expired
dates are never read as MTM). The Summary Groups block drives a 19-line
reconciliation (units, sqft, market/contract rent, deposits, future
residents, and % unit / % sqft occupancy — reports print those
**truncated**, 81.5789 -> 81.57).

### Yardi "Rent Roll with Lease Charges" XLSX (same parser, variant)

Auto-detected off the header: the single `Actual Rent` column is replaced by
`Charge Code` + `Amount`, each unit's charges continue on unnamed rows below
it and close with a per-unit `Total` row. Every printed per-unit Total is
checked against that unit's charge detail and a single mismatch **aborts
before writing** (validated by fault injection). The report's "Summary of
Charges by Charge Code" block is tied out code by code, and Balance columns
tie out for residents and applicants separately. The Summary Groups block
sits one column further right in this variant, so its columns are located
from its own two-row header (`Lease Charges` = ALL charges, which ties to
total lease charges, NOT to Contractual Rent). Section banners are anchored
(`^current...`) because the charge block prints a "(Current/Notice Residents
Only)" qualifier line that an unanchored pattern reads as a new detail
section.

### RealPage OneSite Rents v3.0 "RENT ROLL DETAIL" PDF (`OneSiteRentsParser`)

Report id mgt-521-003, "Details + Summary"; validated on Synott Square,
Houston TX, 108 doors, 6/29/2026. Rows cluster on their vertical CENTRE
(tol 3.5pt) — OneSite baselines differ up to 2.5pt within a line. Text
columns use fixed x-windows; the six right-aligned numeric columns match on
their printed right edge, and the `*` "not included in detail totals" glyph
is found past that edge. Sub Journal + Trans Code print GLUED
(`RESIDENTBLDFAC`, `RESIDENTPESTCONTROLREIMB0.00`): numeric tail split off
by right edge, sub-journal prefix stripped via the report's own sub-journal
summary, every code asserted against its transaction-code summary
(unresolved -> FLAG, never a guess). A bare status row ("Pending"/"Pending
renewal"/"Applicant") opens an ADDITIONAL lease block on the same unit
(status L, never an extra door, matching the `*` exclusion). Move-Out prints
one row below Move-In; resident names re-wrap on every charge row.
`ADMINUNIT`/`EMPLCRED` are recurring credits -> Emp./Other Discounts (col O,
inside NER). MTM only from the explicit MTOM fee. Bed/bath is NOT in a
OneSite roll and A1/B1 codes don't encode it (use `--bedbath` with cited
source); sqft IS printed and ties to the report. "Details + Summary" means
everything the summary prints is reconciled — 93 checks at Synott
(floor-plan table with round-half-up averages, occupancy/rents buckets where
potential rent = contract for occupied and market for vacant, sub-journal
and trans-code summaries, every lease's printed Total Billing). Generic
hook: `checks["extra_checks"] = [(label, fn(units)->value, printed, tol)]`,
recomputed from the UnitRecord list so checks tie output to report.

### ResMan rent roll PDF (`ResManParser`)

Lofts at Taft. Positional column windows on the 792pt landscape page; a long
floor-plan name wraps around the unit row (token above + remainder below).
The Description column is hard-truncated at ~16 chars, so charge names are
resolved against the report's own Total Charges/Total Credits summary —
exact, then unambiguous prefix, then **residual match** (an unresolved group
whose total equals exactly one summary line's unallocated remainder).
Anything still unresolved is left verbatim and the by-description
reconciliation fails loudly. Surety bonds are captured
(`Resident.surety_bond`) and tied to the printed total. MTM comes only from
an explicit "Month to Month Fee" charge. `RENT_CODES` also matches
`RESIDENT RENT`. Bed count comes from the plan code (F1/G1 = 1BR); **bath is
not in a ResMan rent roll and is left blank** — fill it with `--bedbath`
from a cited source, never guess.

### ResMan "Rent Roll Summary" PDF (`ResManSummaryParser`)

Added for McNeil Star 8/4/2026; registered BEFORE `ResManParser` (whose
detect — "ResMan" + "Rent Roll" — also matches a summary). A one-row-per-unit
condensation with no per-charge blocks: Unit / Type / SqFt / Residents /
Status / Market Rent / Rent / Other Charges / Credits / Total / Move In /
Start / Lease End / Move Out / Surety / Deposits / Balance. Reconciles
against the report's grand strip, Property Occupancy and Unit Type Occupancy
tables (25/25 checks at McNeil Star, 24/24 at The Haylie). Conventions:

- Status "UE" = under eviction -> occupied door, resident status C, FLAGged.
- "Other Charges" lump -> Other Income; "Credits" is a lump with no charge
  code, routed through `classify_charge`'s unknown-negative rule (recurring
  vs upfront by the $200 threshold).
- Type codes may or may not encode bed/bath: McNeil Star's A1/A2/B1 do NOT
  (and the digit lies — A2 at 597sf is a 1-bed; web-source and pass
  `--bedbath`), while The Haylie's 1x1/2x1/2x1.5 DO (cite the type code
  itself as the source).
- A page-2 Collections block (former residents) is not part of the roll;
  excluded, and the balance total ties without it.
- Cross-check available: RR market rent should equal the T-12's
  current-month Gross Potential Rent (held at both McNeil Star and Haylie).

### Buildium rent roll PDF (`BuildiumRentRollParser`)

Added for Benbrook 8/5/2026; registered FIRST in `PARSERS`. **Detection:**
"Rent Roll" plus Lease Start + Lease End + Bed/Bath + Prepayments + Balance
Due in the page-1 header band (specific enough that it cannot steal
ResMan/OneSite rolls — regression-checked). Columns: Lease Start | Lease
End | Bed/Bath | Rent Cycle | Rent Start | Rent | Recurring Charges |
Recurring Credits | Total | Deposits Held | Prepayments | Balance Due.

Layout gotchas:

- **No unit identifier column** — detail rows are anonymous; the only
  unit-level identity is the group header. Units are numbered in REPORT
  ORDER (`Unit 01`..`Unit N`) and that substitution is FLAGged — these are
  not the property's real unit numbers.
- **Fake-bold rows print every glyph doubled** — group headers
  ("44663399 WWiilllliiaammss RRooaadd") and totals rows
  ("$$3366,,006655..0000"). `_undouble()` / `_row_undouble()` fix it; the
  test is deliberately WHOLE-ROW, because an isolated "4400" is itself a
  valid pair-doubled string and must never be halved to "40".
- **The Rent column is LEFT-aligned; every other numeric column is
  right-aligned.** Snapping Rent by right edge fails (its right edge moves
  with the digit count and on bold totals rows); Rent is claimed by an x0
  window, the rest by right-edge positional snapping. A vacant unit prints
  "--" in the Rent window.
- **"Rent Cycle: Monthly" is the billing cycle, not MTM**, and **"Rent
  Start" is a DATE**, not a scheduled/market rent. There is no status,
  tenant-name, move-in or move-out column.
- **Recurring Charges INCLUDES the Rent** — Other Income = Recurring
  Charges − Rent, booked as one lump charge (no per-charge detail).
- Sqft may come from the report's own last-page "Summary by bed/bath" block:
  if total sf ÷ units divides EXACTLY per plan, allocate per-unit Net Sf
  from it as rent-roll data (black text, not an estimate). Aggregate
  "Market rent" in the grand totals may cover only the subset of units with
  one set in Buildium — it cannot be allocated per unit, is NOT a tie-out
  target, and missing market rents follow the `--estimate-market` house
  rule.
- **Buildium ROUNDS its occupancy percentages half-up** (19/24 -> 79.17).
  Do NOT reuse `reconcile()`'s `trunc2` (Yardi/OneSite truncate) — the
  property-level and per-plan % checks are filed as `extra_checks` with
  round-half-up. An independent second pass (`extract_text` line splitting)
  re-derives unit count, occupied count and rent sum.

### Owner-made rent roll spreadsheet printed to PDF (`OwnerSheetPdfParser`)

Gardens Apartments. Unit / Tenant / Rent / Unit Type only: no market rent,
no lease dates, no charge codes, no sqft — those columns stay blank rather
than being back-filled from the contract rent. Occupancy is read from the
tenant cell (blank or Vacant/VAC/Empty = vacant). Because owner sheets print
no totals, `parse` **always** re-derives unit count / occupied count / rent
sum through a second independent path (pdfplumber `extract_tables`) and
files them as checks labelled "re-extract", so the reconciliation block
still has something real to tie to. Per-check provenance prints in that
block ("vs report" / "vs re-extract" / "vs printed totals row").

### Rent roll flags that interact with formats

- `--asof YYYY-MM-DD` — sets/overrides the as-of date. A parser that
  explicitly reports "this source prints no date" (`asof_found = False`)
  makes it **mandatory**: the run hard-exits rather than guessing.
- `--bedbath "F1=1/1,F2=2/1.5"` — fills Bed/Bath by floor plan for sources
  that don't carry them (bare `2` sets beds only). Same rules as `--sqft`:
  cited source, stated in the delivery summary. `--sqft-est`/`--bedbath-est`
  and `--estimate-market` are the red-flagged estimate variants (see
  house-rules.md).

## T-12s

### Generic owner/PM-prepared income statement XLSX (`parse_t12_xlsx`)

Validated on Clark Duplexes & Pecan Townhomes 5/2026. Months header like
`June-2025`/`Sept-2025` (may be dotted, `Apr. 2025`), TOTALS column (`TOTAL`
or `TOTALS`), ALL-CAPS section headers, single-line INSURANCE/TAXES
accounts. Every printed annual row total is checked against its monthly sum
(abort on mismatch; `--trust-monthly` with the user's OK). SUBTOTAL-ROW and
GRAND-ROW mismatches are flagged (broken =SUM ranges, double-counted
subtotals, hardcoded cells — all seen in the 84-door M White statement,
whose printed NOI overstated the monthly-detail NOI by $5,170). An ALL-CAPS
zero-filled row with a later `Total X` row is a section header; side flips
to expense at the printed revenue grand row (handles sections like GENERAL &
ADMINISTRATIVE with no "expense" keyword). Slashes in property names are
sanitized to `-` in filenames only. This is also the recommended
intermediate layout for image-only/scanned statements: transcribe into this
layout so all row/subtotal/grand validations run against the transcription
(Royal Oaks 8/2026).

### Yardi "Statement (N months)" XLSX (`parse_t12_yardi_xlsx`)

Dispatched by `_is_yardi_xlsx`; The Meadows. GL number in col A, indented
name in col B, monthly columns only — **there is no TOTALS column**, so
instead of annual row checks every aggregate is verified column-by-column
against the detail it rolls up (`_yardi_verify`, longest-run greedy
consumer, so a zero-summing group like TOTAL RENTS can't under-consume).
Beware caption lies: Yardi titled a six-month export "Statement
(12 months)" — the month COLUMNS are the truth (12-month guard /
`--allow-partial` / `--pad-to-12`).

### RealPage/OneSite "Twelve Month Trailing Income Statement" XLSX (`parse_t12_onesite_xlsx`)

Via `_is_onesite_xlsx`; Synott Square 04-356, 6/2026. Month columns are
period-END dates (`07/31/2025`...) in B..M plus a Twelve Month/Total col N.
The statement tree lives in col A as leading-space indentation (2/level); a
section header's month cells hold the same whitespace padding instead of
numbers. Accounts are `"<GL> - <Name>"`. Roll-ups are `Total <section>` OR
the header's own name repeated (`Net Rental Income`, `Non-Operating
Expenses`) — any valued non-account row is a subtotal; `Net Operating Income
(Loss)` and `Total Current Net Income` are DIFFERENCES (first child less the
rest), retried as such. Every roll-up checks against exactly the
not-yet-consumed deeper-indented rows, and every row checks against its
printed Twelve Month total; both layers print a positive "all tie" line and
abort on failure unless `--trust-monthly`. Below printed NOI: Owner Expense,
Debt Services, Capital General, Other Non-Operating -> Capex & Misc. The
OneSite grand-total pattern set is ANCHORED because `Total Rental Income`/
`Total Other Income` nest above the real grand rows and an unanchored
`total.*income` grabs the wrong one.

### AppFolio "Cash Flow - N Month" / "Income Statement" XLSX (`parse_t12_appfolio_xlsx`)

Via `_is_appfolio_xlsx`; validated on Vista Lago, Tyler TX, Jul 2025-Jun
2026. A parameter preamble (Exported On / Properties / Period Range /
Accounting Basis / Level of Detail) sits above the header row; col A is
literally "Account Name", months start at col B as "Mon YYYY" text, totals
column captioned "Total". The statement tree is leading-space indentation
(4/level) with genuinely EMPTY section-header month cells; leaves and
section headers share indent levels, so classification is structural (a
valued row is a roll-up only if deeper-indented rows are still unconsumed
beneath it). Two AppFolio traps handled explicitly: **restatement rows**
below NOI ("Total Income"/"Total Expense"/"Net Income"/"Cash Flow" repeating
the operating totals) are verified month-by-month against their operating
counterparts and then dropped — left in, they'd land in Capex & Misc
spanning both sides of the ledger; **cash-balance rows** ("Beginning Cash",
"Actual Ending Cash") are point-in-time balances, not flows — excluded
entirely (reported, never silent) with the roll-forward checked for
information. Multi-property exports need `--property`.

### Owner-made "Rental schedule" operating statement (.xls/.xlsx) (`parse_t12_owner_rental_schedule`)

Via `_is_owner_rental_schedule`; validated on Heritage Ridge, OKC, 8/2026
(BIFF8 .xls). Free-form header block ("Rental Property Statement for ...",
"Operating Statement as of:" + a real date cell); ONE row of real date cells
one per month, NO totals column, possibly 4+ fiscal years side by side — a
trailing window is cut out of it (`window` / `window_end`). No property name
anywhere — the parser falls back to the address; pass `--property`. Body:
one row per account, no sections/GL/indentation; a revenue grand row
("Total Gross Rent") on top, everything below is expense. Hand-maintained,
so the parser assumes nothing: body rows with a DATE cell in a month column
are structurally corrupt — dropped and reported verbatim, never guessed at;
`**`-footnoted text amounts ('**23610.82') are parsed and listed in the
notes; when the revenue grand row has no surviving detail beneath it, it is
PROMOTED to an account (so the money is in the T-12), the Total Revenue
subtotal becomes derived, and the LUMP_INCOME rule REVIEW-flags it.

### ResMan "Twelve Month Profit and Loss" PDF (`parse_t12_pdf_resman`)

Sniffed inside `parse_t12_pdf`; Lofts at Taft (also McNeil Star, The
Haylie). 14 numeric columns = 12 months + Adjusted Total + budget Variance
(ignored). ResMan emits every word as its own text op with no spaces, so
extract_text glues labels to numbers; tokens are rebuilt from `page.chars`
in stream order. Sections come from a label-indent stack; each row's printed
Adjusted Total is checked against its monthly sum. `_resman_below` stops
`BELOW_PAT` from swallowing the *operating* section "7100 Interest,
Insurance & Taxes" on the word "interest". Watch for: months that print
0.00 across every row (operations not yet begun — McNeil Star) and an
in-progress final month with accrued income but unposted expenses (The
Haylie) — call both out in the header line, never annualize casually.

### QuickBooks Online P&L PDF (`parse_t12_qbo_pdf`)

Added for Benbrook 8/5/2026. **Detection** (`is_qbo_pdf()`): "Profit and
Loss" title + a 12× "Mon YYYY" header row + QuickBooks' "Total for X"
subtotals; dispatched from both `main()` and `parse_t12_pdf()`.

Layout gotchas:

- **Rows are SPARSE and values MUST be positional.** Most accounts print in
  only some months and QuickBooks prints *nothing* (not 0.00) in the others;
  zipping tokens onto months shifts every gap month's money left, silently.
  Every numeric token is snapped to the month column whose printed RIGHT
  EDGE it matches (~45pt pitch, tol 12pt; the header token's right edge sits
  2-3pt left of the data's). A token that snaps nowhere **aborts** the
  parse. In-period blanks are read as $0.00 (a QBO blank = no transactions
  posted) and counted.
- **The section tree is indentation and "Total for X" closes a section.**
  Depth varies (the below-NOI Other Income block sits a level up), so an
  indent STACK is used, not fixed columns. Sections carry across page
  breaks — and beware "Income Fees" opening a page: it is an EXPENSE section
  (merchant/bill-pay fees), not revenue.
- **Cost of Goods Sold sits between Income and Expenses**, so QuickBooks'
  printed "Total for Expenses" EXCLUDES it. Canonical Total Operating
  Expenses = printed Total for Expenses + printed Total for Cost of Goods
  Sold, and that sum is *proved* by the printed NOI (Total for Income −
  the sum = printed NOI), so `reconcile()` accepts it — it is not a
  derived/self-referential number.
- Everything after the printed Net Operating Income (QuickBooks' Other
  Income / Other Expenses block) is below-the-line -> Capex & Misc.
- Verification layers: every printed row Total vs its own monthly cells;
  every "Total for X" section subtotal vs its account detail per month and
  annually; Total for Income, Total Operating Expenses, Gross Profit, NOI,
  Net Other Income and Net Income vs the parsed detail.
- Cash-basis quirks to surface (not "fix"): whole-year insurance/tax posted
  in one month; fee/billback lines that only begin partway through the
  trailing period; onboarding-sized first-month management fees.

### Google-Sheets-printed income statement PDF (`parse_t12_gsheet_pdf`)

Via `is_gsheet_pdf`; The Gardens. Partial-period capable and **exempt from
the 12-month guard** because it reads an explicit reporting period out of
the subtitle and stamps PARTIAL PERIOD on every output. Values are assigned
positionally by right edge, so a gap month doesn't shift the row left;
out-of-period columns are dropped only after asserting they are empty/zero;
a stub final month (07/01-07/25) is called out. Printed subtotals are
checked per month AND for the period total, monthly detail winning (house
rule).

### T-12 output geometry (all formats)

`write_workbook`/`write_capex` take the real month count; unused template
columns stay blank (never zero-filled) and the Total column stays pinned to
its template position (Trailing col N, RawData/Final T-12 col O) so
downstream formulas and the model import still line up. Total then means
"sum of the months actually reported".
