# St Nicholas Place (4639 Williams Rd, Benbrook TX) — underwriting worked notes, 8/6/2026

Covers: the full-pipeline underwriting of this 40-unit deal (rent roll/T-12 ->
comps -> loan terms -> model -> writeup -> BOV). Read alongside the main
library's `benbrook-8-2026.md` (rent-roll/T-12 parsing notes) before touching
this deal again or reusing the model workflow.

## Final model (TMG template from Git, built on Windows Excel COM)

- Deliverables: `CK - St Nicholas Place - 8-6-2026.xlsx` (naming convention
  "CK - Property Name - Date", Dmytro 8/6) + single-sheet PDF of
  `PDF Output - F&C`.
- **Round 2 reprice under the 8/6 aggressive house rule** (see
  `aggressive-pricing-house-rule-8-2026.md`): strike **$2,830,000**
  ($70,750/unit), Light + Premium Interior Renovations marked (C18+C20,
  $130k capital, +$48k/yr program income, all 40 doors covered); IRR 25.13%
  (floor binds), CoC 13.3%, T-3 DSCR 1.43 at 65% LTV; Year-1 UW NOI $207,111
  (7.32% pro forma cap); supported band $2.3M–$3.3M. The writeup/BOV
  delivered earlier on 8/6 still reflect the prior $2.45M pricing.
- Round 1 (superseded): strike $2,450,000, no renovations, IRR 27.2%,
  CoC 14.9%, DSCR 1.73x.
- T-3 economic-loss distress test = 0% (cash-basis QuickBooks T-12 has no
  vacancy/concession/bad-debt lines -> UW - F&C AC8:AC10 all zero), so the
  DSCR>1.25 rule applied.
- Factors marked: **Low Unit Count (50 bps) + Old Vintage 1979 (100 bps)**;
  the template's custom rows 24–25 carried Aden Crest's "Lease-Up Risk" and
  "MUD District (Wilbarger Creek)" — cleared (no MUD on this parcel; lease-up
  captured via 10% Year-1 vacancy instead). Terminal cap = 5.74% DFW agency
  base + 150 bps -> 7.25%.
- Value-Add kept from template where defensible: pet fees, cable/internet,
  package lockers, water RUBS (+$31,680 NOI total). UNMARKED "reduce opex to
  comp averages" (subject self-manages BELOW benchmarks — the row printed a
  confusing negative delta) and "reduce insurance" (actual $514/unit is under
  the $800 benchmark).
- **Greg's water/sewer ruling (8/5/2026 email): leaks fixed, W&S returns to
  ~$2.4k/month once the city reassesses.** Applied as `Assumptions!G34 =
  28800` (the manual W&S override feeding pro-forma AK28). Note: the T-3 DSCR
  formula swaps in T-3 ACTUAL utilities, so the normalization moves pro-forma
  NOI only. The T-12 deliverable itself keeps actuals.
- Rent comps marked (Rent Comparison col AK — names only appear in AM after
  a recalc, so mark AFTER the first COM recalc): Ridgmar Townhomes, West Wind,
  7700 Chapin Road - 1, Monterrey (Fort West and Oaktree excluded as low/high
  outliers).
- Sale-comp page refreshed via the `Comparable Grid` helper block N3:AE52
  (literals + x/xx/... marks; drift bps in E25:I25, adj in E26:I26) with the
  5 comps from the sales-comps skill run: Ridgmar Townhomes, Waverly Park,
  Vista Del Sol, Elizabeth Gardens, Azle Creek -> indicated $102,567/unit
  ($4.10M). The ~40% gap to the $2.45M strike is the writeup's operational-
  discount narrative (no payroll on owner books, tax reset at sale, SBL-scale
  debt sizing, 12.5% vacancy).
- `Master!D3` (zip) must be an INTEGER — a string zip broke the
  QueryRegion lookup and printed #N/A on the PDF (see
  excel-com-recalc-windows.md, which also documents the stale-value
  strikethrough trap that cost this session several export cycles).
- Known accepted artifact: `UW - F&C!AO10` #DIV/0! (a "% vs T-12" helper
  dividing by zero T-12 bad debt; off-PDF, referenced by nothing).

## Round 4 (8/6 ~2:15 AM)

- **Comments-tab house rule landed** (repo commit 0c78e86): NO red text on
  the T-12 Trailing Financials export. All notes (proration, exclusions,
  pad-to-12, data quality) go to a separate "Comments" tab, plain black, one
  note per row, created only when notes exist. `write_workbook` `note_line`
  now takes a string or list. Rent-roll red ESTIMATE cells unaffected.

## Round 3 (8/6 ~2:00 AM)

- **`--prorate-bulk` house rule landed in the repo** (commit e3d2ea1: new
  flag in `process_t12.py` + section in `house-rules.md`): bulk i/tx
  payments concentrated in <= 3 months get annual/12 respread (cents in the
  last month), lines renamed " (prorated)", applied after parse validation,
  proration checks added to the reconciliation block, harvest strips the
  suffix. T-12 deliverable regenerated; no-flag run verified cell-identical
  to the pre-rule output.
- Model's Final_T_12 rows i/tx respread the same way (labels unchanged
  there — the UW display names are SUMIF-driven); annual totals, metrics and
  strike unchanged; the Dec-25 NOI spike disappears from the page-6 charts.
- Map refit to **I121:O148** (Dmytro's spec; render the PNG at that frame's
  aspect first). "Yellow dots error": the template's scatter markers are
  hardcoded FFFF00 on BOTH series — the chart that PRINTS is chart2 on the
  PDF sheet (chart8 Assume-Loan variant; 5/11/14/15 mirrors). House colors:
  comps 4F81BD, subject diamond FDB714 (see excel-com-recalc-windows.md).

## Other pipeline outputs (same job)

- Sales comps: subject geocoded 32.70512, -97.45191; 94 comps scored, top-10
  trimmed Ridgmar Oaks (high) and Rose Garden (low) at 1.0 SD.
- Loan terms: workbook refreshed with 8/5/2026 yields (UST10 4.63 / UST5 4.34
  / SOFR30 3.62). 40-unit deal -> Freddie SBL 6.63% / Fannie SB 6.68% are the
  natural quotes.
- 2026 TAD assessment $5,471,650 (parcel 06978037) vs 2025 taxes paid $34,604
  — assessment gap flagged to ownership in the writeup (protest evidence).
