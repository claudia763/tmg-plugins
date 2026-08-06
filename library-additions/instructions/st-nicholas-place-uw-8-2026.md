# St Nicholas Place (4639 Williams Rd, Benbrook TX) — underwriting worked notes, 8/6/2026

Covers: the full-pipeline underwriting of this 40-unit deal (rent roll/T-12 ->
comps -> loan terms -> model -> writeup -> BOV). Read alongside the main
library's `benbrook-8-2026.md` (rent-roll/T-12 parsing notes) before touching
this deal again or reusing the model workflow.

## Final model (TMG template from Git, built on Windows Excel COM)

- Deliverables: `CK St Nicholas Place 8-6-2026.xlsx` + single-sheet PDF of
  `PDF Output - F&C`. Strike **$2,450,000** ($61,250/unit) — upper half of the
  model's supported band $1.9M–$2.7M; IRR 27.2% (target 25%), avg CoC 14.9%,
  T-3 DSCR 1.73x at 65% LTV / 6.58% avg agency rate (G62 left blank).
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

## Other pipeline outputs (same job)

- Sales comps: subject geocoded 32.70512, -97.45191; 94 comps scored, top-10
  trimmed Ridgmar Oaks (high) and Rose Garden (low) at 1.0 SD.
- Loan terms: workbook refreshed with 8/5/2026 yields (UST10 4.63 / UST5 4.34
  / SOFR30 3.62). 40-unit deal -> Freddie SBL 6.63% / Fannie SB 6.68% are the
  natural quotes.
- 2026 TAD assessment $5,471,650 (parcel 06978037) vs 2025 taxes paid $34,604
  — assessment gap flagged to ownership in the writeup (protest evidence).
