# Building the TMG underwriting model end-to-end with Excel COM on Windows (8/2026)

Covers: the full populate → tune → export → deliver loop for the TMG underwriting
model using desktop Excel via COM only, plus the places where the CURRENT Git
template (`references/Template Model.xlsx`, as of 8/2026) differs from the
`underwriting` skill's `references/model-map.md`. Read this alongside
`excel-com-recalc-windows.md` before running the `underwriting` skill on a
Windows job machine. Validated on Werner Creek Apartments (36 units, Houston,
8/6/2026) — see `werner-creek-uw-8-2026.md`.

## 1. On Windows you do NOT need the zip-surgery deliverable path

`technical-notes.md` §8 says any save strips `customXml/item*.xml` (Power Query),
`xl/connections.xml` and `xl/queryTables/*`, so the deliverable must be built by
zip-level surgery on a pristine copy. **That is true of openpyxl and
LibreOffice. It is not true of desktop Excel.** A `Workbook.SaveAs(...,
FileFormat=51)` through COM preserves all of it.

Verified on the Werner build by counting parts in the template vs the
deliverable — identical on every one:

| part | template | deliverable |
|---|---|---|
| `customXml/item*` | 8 | 8 |
| `xl/connections.xml` | present | present |
| `xl/queryTables/*` | 11 | 11 |
| `xl/charts/chart*.xml` | 16 | 16 |
| pivotCache parts | 3 | 3 |

So: do **every** edit through COM (`Range.Value = ...`, 2-D tuple assignment for
blocks), recalc with COM, export the PDF from the same session, and `SaveAs` the
deliverable. One file, one path, no surgery, and no openpyxl chart mangling —
which also makes `restore_model_charts.py` unnecessary on this route.

Use openpyxl only for **read-only** inspection (`data_only=True`) and for the
error-diff gate, which is still the reliable cross-file check.

## 2. Assumptions tab — the loan block moved (this WILL bite you)

`model-map.md` documents `G61 LTV / G62 Rate / G63 DSCR / G64-66 IO,term,am`.
The current template's loan block is **program-driven** and one row lower:

| Cell | Current template | Note |
|---|---|---|
| G61 | Loan Type | `=IF('UW - F&C'!AB13<$K$61,"Bridge — Debt Fund",$M$61)`; K61 = 0.75 occupancy threshold |
| **M61** | **the loan program name — SHIPS BLANK** | **must be set**, else G61 = 0 and the entire debt block is empty (LTV/rate/DSCR/term all blank, DSCR reads as an error) |
| G62 | Recourse Type | INDEX from `Loan Terms`!L |
| G63 | LTV | INDEX from `Loan Terms`!F (Max LTV) |
| G64 | Interest Rate | INDEX from `Loan Terms`!E |
| G65 | Min DSCR (reference) | INDEX from `Loan Terms`!G |
| G66 / G67 | IO years / Loan term (months) | INDEX from `Loan Terms` H / I |
| G68 | Amortization (months) | literal, 360 |

Valid `M61` values are the `Loan Terms`!A4:A16 program names, e.g.
`Fannie Mae — Conventional`, `Freddie Mac — SBL`, `Bridge — Debt Fund` (note the
em dash — copy it, do not retype).

**Target IRR is no longer a flat 25%.** `G48 = IF($G$62="Non-Recourse",$K$48,$M$48)`
with K48 = 0.20 and M48 = 0.25 — i.e. **non-recourse debt targets a 20% IRR,
recourse targets 25%.** F48/H48 stay 0.25 for the low/high scenarios. Read G48 at
runtime; do not hardcode 0.25 into the green test.

The model also now carries its own `Treasury Yields` and `Loan Terms` tabs
(FRED `=WEBSERVICE()` pulls with yellow manual-override cells + as-of dates), so
the `loan-terms-lookup` skill's yields should be written into
`Treasury Yields`!F4:F8 with the as-of date in H4:H8 rather than typed into the
Assumptions tab.

## 3. Expense benchmarks: clear the "x" before typing an override

`F28:F32` and `F38` carry `x` = "use the agency benchmark for this line". While
the `x` is present the model uses the benchmark and **ignores whatever you type
in column G**. To underwrite an honest number, `ClearContents` the F cell first,
then write G. (Confirmed: with F28:F32 cleared and G28:G32 typed, `UW - F&C`
AM21:AM25 pick up the typed values exactly.)

This matters on small, owner-managed assets. Fannie's payroll benchmark is
$1,300/unit — an institutional number that assumes on-site staff. A 36-door
property run by a third-party PM with a $750/month contractor will never spend
it, and leaving the benchmark in place cost ~$76k of Year-1 NOI at Werner.

## 4. Comp tables are ranges, not tables; and FinalRR BA:BZ are literals

In the current template `TableRecentLeases`, `TablePropertyData (2)` and
`Final_RR_Floor Plan` carry **no ListObject** — they are plain ranges. Nothing to
resize; just clear generously and write. Likewise `FinalRR!BA9:BZ…` already holds
**literal values** (not the legacy `=Final_RR_Floor_Plan[col]` CSE arrays), so
the array-spill workaround in `technical-notes.md` §3 does not apply.

The ListObjects that DO exist and may need `.Resize()`: `Final_RR`,
`Census_Tract`, `RR_asof_date`, `Final_T_12`, `LoanRates`, `Table002__Page_1`,
`CapRates`, `Output_Analysis_Data__2`, `ZipCodeComps`, `Region_Comps`,
`Append2_2`, `Table6`.

`Value-Add` column C contains merged cells — `Range("C16:C52").ClearContents()`
raises *"We can't do that to a merged cell."* Clear cell-by-cell in a try/except.

## 5. The sale-comp page: feed `Auto Sales`, do not hand-fill the grid

`excel-com-recalc-windows.md` describes overwriting the `Comparable Grid` helper
block `N3:AE52` with literals. That is the **Manual** data-source mode. The
template ships with `Comparable Grid`!M1 = `Automatic`, and in that mode the
whole page is formula-driven off the `Output_Analysis_Data__2` query table on the
hidden `Auto Sales` sheet. The clean path:

1. Run the `sales-comps` skill. Its per-deal `Automatic CMA Analysis.xlsx` copy
   has an `Output Analysis Data` sheet with the top 50 scored comps in
   TotalPoints order and a 24-column schema.
2. Paste those rows into `Auto Sales`!A2:Y…, **inserting a blank at column L** —
   the model's sheet has an extra `Column1` filler there, so source cols 1-11 →
   A-K, blank → L, source cols 12-24 → M-Y. Resize `Output_Analysis_Data__2`.
3. Mark the selected comps with `x` in `Comparable Grid`!L3:L52 (grid row =
   output row + 2). The helper chain in AE turns those into x/xx/xxx/xxxx/xxxxx
   and columns E-I of the grid fill themselves.
4. Overwrite the cap-rate drift with the CMA's own figures: bps literals in
   `E25:I25`, and `E26:I26 = (cap - bps/10000)/cap - 1` where `cap` is the CMA
   AgencyDrift current trailing-12 average.

Done this way the model's `Comparable Grid`!D34/D35 reproduce the Sale
Comparables workbook's indicated value **to the cent** (Werner: $78,751.47/unit,
$2,835,052.94) — which is the check that the two deliverables agree.

## 6. `Agency Region` / `Agency-Data` ship with the PREVIOUS deal's extract

Both sheets are query tables whose Power Query source is not available on the job
machine, so they never refresh — the template ships loaded with Dallas-Fort Worth
rows. Left alone they silently drive:

- `'Agency Loan-Sale Comps'!Z40`, the **base of the terminal cap rate**
  (`Assumptions!G58 = MROUND(Z40 + Factors!N17/10000 + SUMIFS(Factors!I16:I47,…"x")/10000, 0.0025)`), and
- the PDF's "Agency Loan-Sale Comparables (Geographic Region)" page, which would
  otherwise print another metro's properties under this deal's search criteria.

Fix: repopulate both from the CMA workbook's `AgencyDrift` tab, which carries the
**same 26-column schema** (`Loan Type, ID, State, Vintage, TMG Region, Zip,
Origination, Property Name, Address, City, Built, Units, Value at contribution,
Cap Rate, Value/Unit, Month Income/Unit, Expenses/Unit, Reserves/Unit, GRM, LTV,
Revenue, Operating Expenses, NOI, NCF, Loan Balance, Investor Lookup ID`) plus two
extra columns you drop. Filter to the subject's TMG Region, the sheet's own
year-built / unit-count / origination criteria, and sort by origination
descending — only the first 22 rows feed the MAX/MIN/AVERAGE block at rows 38-40.
`scripts/refresh_agency_region.py` in this folder does it.

At Werner this moved Z40 from 5.738% (DFW) to 6.093% (Houston, n=22) — and the
6.093% independently matched the CMA AgencyDrift TX/1961-81 average of 6.0965%,
which is a good cross-check that the filter is right.

`YardiProjections` has the same problem (it ships with a Little Rock forecast) and
feeds `Factors!N16` (avg 5-yr rent growth) and `N17` (market occupancy at
reversion). Replace rows 5/6/7 columns B..O with the subject submarket's series
from the Yardi Forecast Trends PDF; leave quarters the report does not publish
genuinely blank rather than inventing them.

## 7. `PDF Output - F&C` has TWO pictures — do not delete both

- **`Picture 1`, anchored row 1 col 3, ~561 × 77 pt — the TMG letterhead. KEEP IT.**
- `Picture 2`, anchored ~row 121 col 9, ~529 × 404 pt — the previous deal's comp
  map. Delete and replace.

A blanket "delete every msoPicture on this sheet" strips the logo and the export
comes out with an empty band across the top of page 1. Filter on
`sh.TopLeftCell.Row > 100`.

Generate the replacement map per `comp-map-generation.md` and insert it at the
`I121:O147` frame. Render the PNG at that frame's aspect first — 529/404 ≈ 1.309,
so 1530 × 1168 px, not the template html's 1530 × 1000.

Other presentation notes:
- The "black box" conditional format is on **`B50:J80`** in the current template,
  not the `B52:J77` in `model-map.md`. On the Werner build it did not actually
  render a black box (empty floor-plan rows printed white), so check the exported
  PDF before surgery rather than deleting rules blind.
- `Master!B5` (CAD account) must be written as **text** — set
  `NumberFormat = "@"` before assigning, or `0650980000002` stores as the float
  `650980000002` and the leading zero is lost. Same class of bug as the
  `Master!D3` zip-must-be-an-integer trap, opposite direction.
- The scatter marker recolour ("yellow dots") can be done through COM on the live
  workbook — iterate `ChartObjects` → `SeriesCollection` and set
  `MarkerBackgroundColor` / `MarkerForegroundColor` (BGR: comps `0xBD814F` =
  4F81BD, subject `0x14B7FD` = FDB714). No ZIP patching needed on this route.

## 8. Tuning: solve LTV and price together, not price alone

The template sizes the loan purely off `Loan Terms` Max LTV. On a
DSCR-constrained deal that produces a loan no lender would fund, and then the
green rule fails on DSCR — so a naive price-only search crushes the price to make
an unfinanceable loan work.

The house rule (underwriting SKILL: "usually forces lower leverage (45–55% LTV)
on weak trailing NOI") means the correct lever is leverage. Sweep LTV and, for
each, find the maximum price where **F5 ≥ G48, F7 ≥ 0.10 and I8 ≥ 1.25** all
hold; take the LTV that supports the highest price.
`scripts/model_price_solver.py` in this folder does it. The curve is genuinely
non-monotonic — at Werner:

| LTV | max green price | binding constraint |
|---|---|---|
| 55% | $2,300,000 | IRR |
| 65% | $2,400,000 | IRR |
| **70%** | **$2,460,000** | **IRR (DSCR 1.27, slack)** |
| 75% | $2,360,000 | DSCR |
| 80% | $2,250,000 | DSCR |

Below the peak the equity cheque throttles IRR/CoC; above it DSCR binds. The peak
is worth ~$210k of price versus running the template's stock 80% LTV.

Excel rebuilds this model in ~1 s, so a 10-step-per-LTV sweep costs a couple of
minutes — cheap enough to run every time rather than guessing.

## 9. Known-acceptable artifacts on this template

The error-diff gate (openpyxl `data_only=True`, deliverable vs pristine template)
came back with **10 new error cells** on the Werner build, all on hidden utility
sheets and none inside the `B1:O333` print area:
`Agency-Data!C1` (the "Filepath:" helper string), `MasterDealList!A1/F4/F6/BH4/BH6`,
`Validation!F2/F5/F6/F8`. `technical-notes.md` already lists Validation and the
Master helper columns as pre-existing error territory; these are the same family
re-triggered by new inputs. Note them, don't chase them.
