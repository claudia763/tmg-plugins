# TMG Underwriting Model Skill

## What this skill produces

1. The completed underwriting model (.xlsx) — named
   **"<Initials> \- <Property Name> \- <Date> \- Final.xlsx"** — populated
   directly with openpyxl and recalculated/verified in LibreOffice. There is
   **no separate "surgery" build anymore**\: we run the deliverable in
   LibreOffice, which cannot refresh Power Query in any case, so the
   openpyxl\-saved workbook IS the deliverable. A worked example of the
   finished product is bundled at
   `references/Initials - Property Name - Date - Final.xlsx`.
2. A PDF of the **PDF Output \- F&C** tab only (landscape, fit\-to\-width),
   exported with the bundled `scripts/export_pdf.py`.

Read `references/model-map.md` for the exact cell map before touching the
workbook, and `references/technical-notes.md` for the openpyxl/LibreOffice
gotchas — several are load\-bearing (external defined names, array formulas,
row\-capacity limits).

## LibreOffice\-first ground rules

The model is authored, recalculated, and consumed in **LibreOffice** — not
Excel. Keep the workbook inside LibreOffice's comfort zone:

- Do not rely on, add, or try to preserve Power Query / DataMashup,
  workbook connections, query\-table refresh, or pivot\-cache refresh. An
  openpyxl or LibreOffice save strips them; nothing downstream needs them.
- No Excel\-only or "clever" functionality in anything you write into the
  workbook: no dynamic arrays / spill functions (XLOOKUP, FILTER, LET,
  LAMBDA), no volatile whole\-column tricks, no new external links, no new
  structured\-table formulas. Plain cell references and the functions the
  template already uses are the ceiling.
- Values beat formulas: where a block only reports data (comp tables,
  Raw Rents mirrors, anything LibreOffice mangles per technical\-notes
  §1–3), write literal values.
- Recalc and PDF export happen through LibreOffice headless (xlsx skill
  `recalc.py`, `scripts/export_pdf.py`); expect the template's \~380
  pre\-existing error cells and require zero NEW errors.

## Inputs to collect

- The model workbook. If the user attaches one ("Initials  Property Name
  Date …xlsx"), use it — and if two copies are attached, use the one
  containing a `FinalRR` sheet. If none is attached, start from the bundled
  blank template at `references/Template Model.xlsx`.
- Rent roll workbook (redIQ export: `Rent Roll` \+ `Floor Plan Summary` sheets).
- T\-12 workbook (`Trailing 12 Month Statement` sheet with code column: r, ll,
  v, nr, bd, rw, rt, ro, oi, cs, rm, ad, m, pr, w, tr, e, o, mf, i, tx).
- Rent Data market workbook (`Rent Comps`, `Unit-Level Data`, `Reviews Analysis`, `Specials`, `Fees & Amenities` sheets).
- County appraisal district card (property values, taxing units, year built,
  land SF). Verify any tax bill actually matches the subject property —
  billing numbers can coincide with unrelated parcels; the CAD card is the
  authority for tax rates and assessed value.

## Workflow

**Phase A — value the deal in Python (`scripts/tmg_valuation.py`).**
All assumption tuning happens here; iteration is instant instead of a \~2\-min
LibreOffice recalc per pass.

1. Copy `scripts/tmg_valuation.py` into the job folder and fill its config
   blocks from the source documents — they mirror the model tab\-for\-tab:
   - `PROPERTY` — units, rentable SF, market rent/unit, total tax rate from
     the CAD card (Master\-tab facts).
   - `T12_MONTHLY` — monthly income and expense series aggregated from the
     T\-12 statement by code (same lines as Final\_T\_12), oldest → newest.
   - `ASSUMPTIONS` — the Assumptions\-tab blue cells (Year\-1 operating
     inputs, expense overrides / "agency" benchmark flags, growth rates,
     price, debt terms, syndication terms).
   - `VALUE_ADD_ITEMS` — include\-flags for the Value\-Add programs (top 3
     per category count, exactly like the tab). A 2018\+ build gets no
     interior renovation program.
   - `FACTORS` — sale\-comp cap rate plus cap\-rate risk adjustments in bps
     (keep custom adjustments 25–75 bps and defensible).
2. Research the property (reviews, specials, amenities; Rent Data
   `Reviews Analysis` sentiment vs the T\-12) to justify Value\-Add
   selections and risk factors before flipping flags.
3. Run the script and tune purchase price / LTV / assumptions against the
   house rules below. The script prints the full proforma, Project IRR,
   Avg Cash\-on\-Cash, T\-3 DSCR, and a validation table;
   `solve_price_for_irr()` goal\-seeks the pricing band.

**Phase B — populate the deliverable model.**
Work on the actual deliverable copy from the start; no disposable working
copy and no zip\-surgery rebuild afterward.

4. **Master tab** — property info in B1:B19 (\+ C3 state, D3 zip). Tax rates
   B11:B17 from the CAD card's taxing\-units table; total must reproduce the
   card's aggregate rate. Full market value \= current\-year net appraised.
   Date of rent roll \= the "as of" date inside the rent roll file, not its
   filename.
5. **FinalRR \+ Final\_T\_12** — paste rent roll rows into FinalRR C9:Y…, floor
   plan summary into FinalRR BA9:BZ… and the `Final_RR_Floor_Plan` table;
   aggregate the T\-12 statement by code into Final\_T\_12 (fold `rt` into the
   `ro` row; set the month dates in row 2; column O \= row totals). Cross\-check
   rev/exp/NOI against the statement's own `rev`/`exp`/`noi` rows — these must
   be the same numbers already loaded into `T12_MONTHLY` in Phase A.
6. **Comp tables** — TableRecentLeases from Unit\-Level Data (comps only,
   exclude the subject; filter to leases in the last 12 months \+ active
   listings; cap total rows at 996 — see technical notes); TablePropertyData
   from the Rent Comps sheet (subject row, "Comp Average" row, then comps,
   with address split into street/city/state/zip). Then value\-fill the
   Raw Rents mirror columns (see model map).
7. **Rent comp selection** — recalc, then "x" in Rent Comparison column AK for
   the best 4–5 comps: highest\-renting cluster, dropping low\-rent outliers and
   any single comp priced far above the pack. The PDF summary table only shows
   the first four selected; a fifth still feeds the averages.
8. **Transcribe the settled valuation** — copy the final Phase\-A inputs into
   the model verbatim: Assumptions\-tab blue cells (including strike price and
   loan terms), Value\-Add column\-C "x" flags, and Factors column\-J flags plus
   any custom rows 24–25. Do not re\-tune in the spreadsheet; the Python run
   is the source of truth.
9. **Verify \+ deliver** — LibreOffice recalc, then read F5 (IRR), F7 (CoC),
   I8 (T\-3 DSCR) with `data_only=True` and confirm they match the Python
   outputs within ±2% (they normally match to rounding; a bigger gap means a
   transcription error — find it, don't shrug). Zero new formula errors.
   Export the PDF Output \- F&C tab with `scripts/export_pdf.py`, verify page 1
   visually, name the workbook
   "<Initials> \- <Property Name> \- <Date> \- Final.xlsx", and send both files.

## House rules (green thresholds)

- **Target IRR: 25%** (set Assumptions F48:H48 \= 0.25). Green requires
  Project IRR (F5) \> target — land it with a 2–4 point cushion.
- **Avg cash\-on\-cash (F7) \> 10%.**
- **T\-3 DSCR (I8) \> 1.25 — UNLESS the deal is distressed.** Distress test:
  T\-3 economic loss \= T\-3 vacancy % \+ concessions % \+ bad debt % (UW \- F&C
  AC8\+AC9\+AC10). If **\> 30%**, assume the deal goes to bridge financing and
  **ignore the DSCR test entirely** — tune only IRR and CoC, and leverage may
  go to 65–75% LTV. If ≤ 30%, DSCR must be green, which usually forces lower
  leverage (45–55% LTV) on weak trailing NOI.
- Price aggressively for the listing pitch: push the strike toward the upper
  half of the model's own supported band (Assumptions F50/H50) while keeping
  every required metric green with cushion.
- RE tax assessment factor: 100% of purchase price (F&C), tax rate from the
  CAD card. Leave interest rate blank to use the average agency quote unless
  the deal story says otherwise.

## Presentation rules

- Remove the "black box": the PDF Output \- F&C sheet has a conditional format
  on the floor\-plan block (≈ B52:J77) that fills empty rows black. Delete that
  rule before exporting.
- Zero *new* formula errors after recalc (compare against the template's
  pre\-existing error set; the template ships with known \#REF\!/\#VALUE\!
  artifacts on Validation, Tax Assessment, Yields, and unused Value\-Add rows).
- Flag anything you could not refresh (sale\-comp pages, agency datasets) in
  the summary to the user — never present stale prior\-deal comps silently.
