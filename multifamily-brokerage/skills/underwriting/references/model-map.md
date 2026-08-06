# TMG Underwriting Model — Cell Map

Sheet names below are exact. The model is ~65 tabs; only these need edits.

## Master

| Cell | Content |
|---|---|
| B1 | Property name |
| B2 | Street address |
| B3 / C3 / D3 | City / State / Zip (C3 and D3 feed the PDF address line) |
| B4 | Link to property CAD page |
| B5 | CAD account / PID |
| B6 | County |
| B7 | ISD |
| B8 | Valuation year (formula `=YEAR(TODAY())` — leave) |
| B9 | Full market value (CAD net appraised, current year) |
| B10 | Assessor ratio (1) |
| B11–B17 | Tax rates per $100: city, ISD, county, college, hospital, "County URD", spare ("-"). Map every CAD taxing unit somewhere in these seven rows (MUD/ESD go in B16/B17 with a cell comment); B24 must reproduce the CAD total rate. |
| B18 | Year built |
| B19 | Land size (SF) |
| B20 | Date of rent roll (the "as of" date in the file) |
| B29–B38 | Assumable/current loan details — leave blank if none provided |
| G33 / H33 | Calculated total units / total SF — verify after paste |

## FinalRR

- C3: literal string `Rent Roll as of:  <Month D, YYYY>` (two spaces after the colon — C2 SUBSTITUTEs on that exact prefix).
- C9:Y(8+n): rent roll rows, columns in source order: Unit No., Floor Plan, Net Sf, Bed, Bath, Lease Type, Renovation Status, Occupancy Status, Market Rent, Contractual Rent, Recurring Concessions, Net Effective Rent, Supplemental Rent, Upfront Concessions, Emp./Other Discounts, Other Income, Lease Start, Lease Expiration, Lease Term, MTM, Move In, Move Out, Vac. Notice. Skip trailing junk rows (unit "0"). Clear C9:Y658 first.
- BA9:BZ(8+k): floor-plan summary rows (values — replace the legacy array formulas that point at the `Final_RR_Floor_Plan` table). Also write the same data into the `Final_RR_Floor Plan` sheet rows 2+ and resize its table ref.
- Named ranges rrUnitFloorPlan/SF/Bed/Bath, rrMarketRent, rrInPlaceRent → FinalRR D:L rows 9:2000.
- The `Final_RR` sheet (underscore, queryTable) is unreferenced — do not touch.
- `RR_asof_date`!A2: set to the rent roll date.

## Final_T_12

- Row 2 C:N = 12 month-start dates (`mmm-yy`); O2 label "Adjusted Total".
- Row map by code: r→4, ll→5, v→6, nr→7, bd→8, rw→9, ro→10 (add `rt` here), oi→11, cs→14, rm→15, ad→16, m→17, pr→18, w→19, tr→20, e→21, o→22, mf→23, i→24, tx→25. Column O = 12-month total.

## Comp tables

- `TableRecentLeases` (sheet + table, A1:V…): cols A–T = Unit-Level Data B–U
  (Property Name … Detected Unit Amenities), U = Units_Leased (1), V =
  FixedName (= property name). Comps only — no subject rows. Resize table ref.
- `TablePropertyData (2)` sheet, table name `TablePropertyData` (A1:AQ…):
  row 2 subject, row 3 "Comp Average" (street "--", rest blank), rows 4+ comps.
  A Property, B Street, C " City" (leading space), D State, E Zip (int),
  F Similarity, G Dist, H Quality, I Yr Built, J # Units, K Stories,
  L Avg Sqft, M Leased %, N Exposure %, O–Z asking rent/PSF by BR,
  AA/AB avg asking, AC–AN effective by BR, AO/AP avg effective, AQ FixedName.
- **Raw Rents mirror (value-fill; the array formulas there have fixed stale
  refs — overwrite with literals):** lease-level rows 4+: W name, X beds,
  Y baths, Z sqft, AA units_leased, AE asking rent, AF asking PSF,
  AG effective rent, AH effective PSF (clear through row 999). Property-level
  rows 4+ (subject, Comp Average, comps): AL street, AM name, AN units,
  AT quality, AW city, AX state, AY zip, AZ avg sqft, BB avg asking rent,
  BL leased %, BN year built (clear rows 4:16 of AL:BS first). Keep the
  per-row formulas in U and AJ.
- **Hard cap: 996 lease rows** (named ranges stop at row 999). Filter to last
  12 months + active listings, then cap per property so the total fits.

## Rent Comparison

- Column AK rows 4+: "x" to include a comp (comp names appear in AM after
  recalc). Clear all old x's first. The Rent Summary/PDF table shows only the
  first four selected (helper chain "x","xx","xxx","xxxx"); a 5th feeds
  averages only.
- Columns C:AF rows 4+ are optional per-comp amenity adjustments (blank = 0).

## Value-Add (column C "x", max 3 per bucket)

Renovations C18:C23 · Amenity upgrades C27:C40 · Opex/RUBS C44:C52.
Row 45 (lease-up) is dead — its include cell feeds nothing; capture lease-up
via Year-1 vacancy instead. Selection guidance lives in column A of each row.

## Factors

- Cap Rate Adjustments: G (label) / I (bps) / J ("x") rows 16–47. Rows with
  formulas in J are auto-triggered — leave them. Manual rows: J16 tertiary,
  J17 demographics, J18 low unit count, J23 old vintage. **Rows 24–25 are the
  custom slots**: overwrite G24/G25 labels, set I24/I25 bps, mark J24/J25.
- Feeds terminal cap: `Agency Loan-Sale Comps'!Z40 + Factors!N17/10000 +
  SUMIFS(I16:I47 where J="x")/10000`, MROUND to 25 bps.

## Assumptions (input column G unless noted)

| Cell | Input | House default |
|---|---|---|
| G17 | Year-1 rent growth | 0–2% (check UW - F&C AM4 comp-supported growth; often negative) |
| G18 | Loss to lease | ~5% |
| G19 | Vacancy | ~10% on a lease-up story |
| G20 | Concessions | ~3% |
| G21 | Bad debt | ~2.5% |
| F28:F32, F38 | "x" = agency benchmarks (template default — keep) |
| G27 | Expense comp set | Fannie |
| G37 | Mgmt fee | 3% |
| G39 / G40 | Tax assessment factor F&C / Assumption | 1.00 / 0.80 |
| G41 | CapEx reserve $/unit/yr | 350 |
| F43/G43/H43 | GPR growth Low/STRIKE/High | 2.0% / 2.5% |
| G46 | Tax growth | 2.5% |
| F48:H48 | Target IRR | **0.25** |
| G50 | Purchase price STRIKE (the main lever) | solve |
| G52 | Assumption-scenario price | blank/0 unless assumable loan |
| G56 | Origination | 3% |
| G57 | Hold | 5 yrs |
| G58 | Terminal cap | leave formula (auto = sale comps + factors) |
| G61 | LTV | 65–75% distressed/bridge; 45–55% if DSCR must pass |
| G62 | Rate | blank = avg agency quote |
| G64/G65/G66 | IO / term / amortization | 3 yrs / 120 mo / 360 mo |
| G63 | DSCR requirement (threshold for I8 green) | 1.25 |

**Green logic (conditional formats):** F5 green > G48 · F7 green > 0.10 ·
I8 green > G63. Read values after recalc with `data_only=True`.

## PDF Output - F&C

- Print area B1:O333, landscape, fit-to-width. ~8 pages.
- Delete the CF rule on B52:J77 (empty → black fill) before export.
- Known template typo (fixed in TMG's current file, check anyway):
  'UW - F&C'!G69 must be `=Final_T_12[date2]` as a CSE array formula —
  the template shipped with `[date3]`, duplicating April in the T-12 header.
