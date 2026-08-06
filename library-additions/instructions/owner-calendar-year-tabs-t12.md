# Owner-made CALENDAR-YEAR TABS statement -> stitching a real T-12

Covers the owner/PM-prepared Excel operating statement that keeps **one tab
per calendar year** instead of one trailing-twelve sheet, and how to build a
genuine T-12 out of it. Read this when a T-12 source workbook has sheets named
`2025`, `2026`, ... and neither one is a trailing twelve on its own.
Added 8/6/2026 (Magnolia Place, Brenham TX, 20 units).

## Recognising the dialect

- Two or more worksheets whose names are 4-digit years.
- `A1 == 'Category'`; `B1..M1` = `January`..`December`; `N1` = an annual total
  column.
- ALL-CAPS section headers with no values (`RENTAL INCOME`, `OTHER INCOME`,
  `EXPENSES`, `GENERAL & ADMINISTRATIVE`, `CONTRACT SERVICES`,
  `REPAIRS & MAINTENANCE`, `TAXES & INSURANCE`, `MARKETING`) mixed with
  ALL-CAPS rows that ARE single-line accounts carrying values
  (`MANAGEMENT FEES`, `UTILITIES`, `PAYROLL`). A header can also print a
  zero-filled row before its detail — classify structurally (does deeper
  detail follow?), never by capitalisation alone.
- Printed roll-ups: `TOTAL RENTAL INCOME`, `TOTAL INCOME`, `TOTAL EXPENSES`,
  `NET OPERATING INCOME`.
- Below the printed NOI: a `CAPEX EXPENSES` block with **its own repeated
  month-header row**, a `Total CapEx` row, a `Mortgage` row, and memo rows.

Parser: `parse_t12_owner_calendar_year_tabs` / `_is_owner_calendar_year_tabs`
in `process_t12.py`, registered first in `T12_XLSX_PARSERS`.

## The three traps

1. **The annual-total caption lies.** Magnolia's *2026* tab captions column N
   `Total 2025` — a copy-paste from the prior year's tab. It is really the
   2026 total. Key detection and windowing off the **tab name plus the
   January–December header row**, never the caption. (Same family of bug as
   Yardi titling a six-month export "Statement (12 months)".)
2. **Find the window end from OPERATING data only.** The current-year tab is
   pre-filled with hard zeros for months not yet posted, but a below-the-line
   `Mortgage` row is often pre-filled for all twelve. Testing "last month with
   any value" then claims December and silently pads six empty months into the
   T-12. Test only rows above the printed NOI.
3. **Accounts drift between tabs.** Owners rename, add and drop lines at the
   year turn. Match by section + whitespace-normalised name; merge only
   whitespace-identical or >=0.90 whole-string matches inside the same
   section, and **report every merge**. Accounts present on only one tab get
   **genuinely blank** months on the other tab's span — never zeros.
   Magnolia: `'Registered Agent Fees '` == `'Registered Agent Fees'`;
   `'Restripping Parking Lot'` == `'Restriping Parking Lot'` (0.978).
   Do NOT auto-merge lines that merely look like the same vendor renamed
   (Magnolia's `Maintenance` ending Dec-25 vs `Grounds Cleaning` starting
   Jan-26) — keep them separate and say so in the notes.

## Four verification layers (all must tie before writing)

a. Per row, per tab: the twelve monthly cells vs the printed col-N annual
   total — run this over each **full calendar year**, before any windowing.
b. Per month, per tab: `TOTAL RENTAL INCOME` vs its rental detail;
   `TOTAL INCOME` vs rental + other income; `TOTAL EXPENSES` vs expense
   detail; `NET OPERATING INCOME` vs income less expense.
c. `Total CapEx` vs the capex detail, per month and per year.
d. Re-run (b) on the **stitched window** — that is what actually proves the
   stitch — and check the windowed Total Revenue / Total OpEx / NOI against
   the sum of the statement's own printed grand rows over those same months.

The trailing window crosses calendar years, so there is no single printed
annual total to tie to. (a) and (d) together are the honest substitute: every
month is proved against a printed number on its own tab, and the window is
proved against the sum of printed grand rows.

## Related fixes landed alongside this parser

- `SECTION_ALLOWED` gains a combined `Taxes & Insurance` rule -> `{i, tx}`,
  tested before the plain `^tax|taxes` rule (same shape as the existing
  `FIXED ADMINISTRATIVE` precedent). Without it, property insurance filed
  under a combined `TAXES & INSURANCE` head fires a spurious REVIEW.
- `write_capex()` now honours `Line.empty`, so a genuinely blank capex month
  stays blank instead of printing 0.00 (`write_workbook` already did this).
- **Bug fix:** `--exclude-account` removal reports were printed to console but
  never written into the workbook. They now land on the **Comments** tab, as
  the 8/6/2026 no-red-text house rule requires. New repeatable `--note` flag
  writes methodology/data-quality notes to the same tab.

## Judgment calls worth reusing

- A memo row that repeats a capex line cell for cell (Magnolia's
  `Units Rehabbed` == `Unit Renovations`) is **not** a second cost. Remove it
  with `--exclude-account` and `--exclude-reason` so the removal is reported,
  never silent.
- `Landscaping Maintenance` and a bare `Maintenance` sitting inside the
  statement's own `CONTRACT SERVICES` section, recurring at a flat monthly
  amount, with a separate `REPAIRS & MAINTENANCE` section carrying the
  job-sized work and a separate `$0 PAYROLL` line -> **`cs`**. This is
  deliberately the opposite of the Royal Oaks 8/2026 ruling (bare, lumpy
  "Maintenance" -> `rm`): the section placement and the flat recurrence are
  what decide it, not the label.
- Apply deal-specific overrides through a **run-local corpus copy**
  (`--mappings work/t12_mappings_<deal>.csv`). Never edit the shared
  `t12_mappings.csv` inside a job.
