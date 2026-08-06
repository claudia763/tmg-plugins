# Relinking spec — Sale Comparables Workbook → Automatic CMA Analysis

The template (`assets/Sale Comparables Workbook.xlsx`, one sheet: `Comparable
Grid`) was born linked to a full 63-tab underwriting model (`[1]` = "DG -
Benbrook Apartments…xlsx"). `build_output.py` rewires it so its only external
dependency is the Automatic CMA Analysis file sitting **in the same folder**.
All edits are raw-XML so styles, the map drawing, data validation, and the web
extension survive.

## Sheet anatomy (Comparable Grid)

- **C5:I35** — the visible 5-comp table. Comp columns E–I pull from the helper
  grid via `INDEX($col$3:$col$52, MATCH("x"/"xx"/…, $AE$3:$AE$52))`; the
  `AE` column accumulates one "x" per included row, so the Nth marked row in
  L3:L52 becomes Comp N. These formulas are left untouched.
- **L2:AE52** — helper grid, rows 3–52 ↔ data rows 1–50 of the CMA's `Output
  Analysis Data` tab (which select_comps writes already sorted by TotalPoints,
  so row 3 = best-scoring comp).

## What gets rewritten

| Target | Before (underwriting link) | After (CMA link) |
|---|---|---|
| externalLink1 rels (all 3 path forms) | `…DG - Benbrook Apartments….xlsx` | `Automatic CMA Analysis.xlsx` (relative) |
| externalLink1 cache | 63 underwriting sheets | CMA sheet list + cached values for `Inputs` B2:B9 and `Output Analysis Data` rows 1–51, so the grid displays correctly even with links not yet updated |
| Helper N,O,Q–Z rows 3–52 | `IF($M$1="Manual",…,INDEX([1]!Output_Analysis_Data__2[Col],ROW()-2))` | direct ref `'[1]Output Analysis Data'!<col><row-1>` (N→M, O→A, Q→B, R→C, S→D, T→E, U→F, V→R, W→L, X→G, Y→H, Z→J) |
| `rediq_dealname/address1/city/state/zip` defined names | `[1]Master!$B$1…` | `[1]Inputs!$B$2…$B$6` (fixes C5, D8, D10–D13) |
| D17 / D21 / D23 (subject units, year, avg size) | `[1]Master!G33/B18/H34` | `'[1]Inputs'!B9 / B7 / B8` |
| L3:L52 markers | template's leftover x's | exactly the 5 selected comps (`x` at row rank+2) |
| E25:I25 Cap Rate Drift | formula into `'[1]Agency Loan-Sale Comps'` | computed bps values (from Fannie/Freddie analysis) |
| E26:I26 drift adjustment | ref to `'[1]Agency Loan-Sale Comps'!$Z$40` | `=IFERROR(IF($AI$1="",0,(((E15*($AI$1-E25/10000))/($AI$1))/E15-1)),"")` — same algebra, cap rate now in AI1 |
| AH1/AI1 (new) | — | label + current avg cap rate (decimal) |
| workbook.xml | — | `fullCalcOnLoad="1"`; calcChain.xml deleted (Excel rebuilds) |
| cached `<v>` on formula cells | stale Benbrook values | stripped, so nothing shows old-deal data |

## Deliberately left alone

- `P` ( $/unit), `M` (comment), `AD` (lookup hook), `AE` (x accumulator) — local
  formulas.
- `AA:AC` (agency/CoStar/Yardi price probes) — reference tabs that don't exist
  in the CMA; they are IFERROR-wrapped and only feed Manual mode, so they
  quietly return "".
- **Manual mode is inert after relinking**: the `CoStarSale*` / `YardiSale*`
  defined names still point at `'[1]Sale Comps'`/`'[1]Raw Rents'`, which the
  CMA doesn't have. M1 stays "Automatic". If TMG ever needs Manual mode back,
  those names would have to be remapped to the CMA's `CoStar Sale Data` /
  `Yardi Sale Data` tabs (different column layout — not a find/replace job).

## Why the CMA copy is written too

The helper grid reads `Output Analysis Data` rows 1–50 **by position**, and the
shipped CMA's tab is stale (computed for whatever deal it last refreshed, via a
now-dead Bing geocoder). So build_output.py writes a per-deal CMA copy whose
tab contains the freshly scored, sorted comps — the workbook pair is
self-consistent. Keep both files in the same folder; Excel resolves the
relative link on open ("Update Links" → values refresh; "Don't Update" → the
cached values we wrote show the same numbers).

If the user later refreshes Power Query in the CMA copy, the `Output Analysis
Data` query will fail at the geocode step (Bing Maps API retired) — the
values written by this skill remain in place unless that specific query is
refreshed. Refreshing the Yardi/CoStar/Agency source queries is fine.
