---
name: underwriting
description: >
  Use this skill to underwrite a multifamily property in The Multifamily Group's
  Excel underwriting model (the "Initials  Property Name  Date" / Aden Crest
  template). The workflow populates the Master tab, FinalRR rent roll,
  Final_T_12 trailing financials, and rent-comp tables from attached documents
  (rent roll .xlsx, T-12 .xlsx, market Rent Data workbook, CAD appraisal card),
  selects rent comps and value-add programs, sets cap-rate risk factors,
  tunes the Assumptions tab until the return metrics hit TMG's green
  thresholds, and prints the "PDF Output - F&C" tab as the deliverable.
  Trigger whenever the user asks to "underwrite" a deal or model, "populate
  the model," "run the underwriting," or attaches a rent roll + T-12 alongside
  the TMG underwriting model workbook.
---

# TMG Underwriting Model Skill

## What this skill produces

1. The completed underwriting model (.xlsx) — named
   **"<Initials> - <Property Name> - <Date> - Final.xlsx"** — built via the
   Power Query-preserving surgery path (technical-notes §8 /
   `scripts/surgery_example.py`), with all inputs populated, return metrics
   green, and the template's Power Query / connections / query tables intact.
   A worked example of the finished product is bundled at
   `references/Initials - Property Name - Date - Final.xlsx`.
2. A PDF of the **PDF Output - F&C** tab only (landscape, fit-to-width),
   exported with the bundled `scripts/export_pdf.py` from the disposable
   working copy.

Read `references/model-map.md` for the exact cell map before touching the
workbook, and `references/technical-notes.md` for the openpyxl/LibreOffice
gotchas — several are load-bearing (external defined names, array formulas,
row-capacity limits).

## Inputs to collect

- The model workbook. If the user attaches one ("Initials  Property Name
  Date …xlsx"), use it — and if two copies are attached, use the one
  containing a `FinalRR` sheet. If none is attached, start from the bundled
  blank template at `references/Template Model.xlsx`.
- Rent roll workbook (redIQ export: `Rent Roll` + `Floor Plan Summary` sheets).
- T-12 workbook (`Trailing 12 Month Statement` sheet with code column: r, ll,
  v, nr, bd, rw, rt, ro, oi, cs, rm, ad, m, pr, w, tr, e, o, mf, i, tx).
- Rent Data market workbook (`Rent Comps`, `Unit-Level Data`, `Reviews
  Analysis`, `Specials`, `Fees & Amenities` sheets).
- County appraisal district card (property values, taxing units, year built,
  land SF). Verify any tax bill actually matches the subject property —
  billing numbers can coincide with unrelated parcels; the CAD card is the
  authority for tax rates and assessed value.

## Workflow

1. **Master tab** — property info in B1:B19 (+ C3 state, D3 zip). Tax rates
   B11:B17 from the CAD card's taxing-units table; total must reproduce the
   card's aggregate rate. Full market value = current-year net appraised.
   Date of rent roll = the "as of" date inside the rent roll file, not its
   filename.
2. **FinalRR + Final_T_12** — paste rent roll rows into FinalRR C9:Y…, floor
   plan summary into FinalRR BA9:BZ… and the `Final_RR_Floor_Plan` table;
   aggregate the T-12 statement by code into Final_T_12 (fold `rt` into the
   `ro` row; set the month dates in row 2; column O = row totals). Cross-check
   rev/exp/NOI against the statement's own `rev`/`exp`/`noi` rows.
3. **Comp tables** — TableRecentLeases from Unit-Level Data (comps only,
   exclude the subject; filter to leases in the last 12 months + active
   listings; cap total rows at 996 — see technical notes); TablePropertyData
   from the Rent Comps sheet (subject row, "Comp Average" row, then comps,
   with address split into street/city/state/zip). Then value-fill the
   Raw Rents mirror columns (see model map).
4. **Rent comp selection** — recalc, then "x" in Rent Comparison column AK for
   the best 4–5 comps: highest-renting cluster, dropping low-rent outliers and
   any single comp priced far above the pack. The PDF summary table only shows
   the first four selected; a fifth still feeds the averages.
5. **Research + Value-Add + Factors** — search the property online (reviews,
   specials, amenities) and compare with the Rent Data `Reviews Analysis`
   sentiment and the T-12. Select up to 3 per Value-Add bucket (column C "x").
   A 2018+ build gets no interior renovation program. On the Factors tab mark
   applicable cap-rate adjustments in column J and use rows 24–25 for custom
   adjustments (e.g., lease-up risk when occupancy trails comps, MUD tax
   burden) — keep each 25–75 bps and defensible.
6. **Assumptions tuning** — set honest Year-1 inputs (see model map for the
   house defaults), then solve purchase price and LTV against the green rules
   below. Iterate: write inputs → recalc → read F5/F7/I8 → adjust.
7. **Deliver** — export the PDF Output - F&C tab with
   `scripts/export_pdf.py` from the working copy and verify page 1 visually.
   Then build the deliverable workbook by zip-level surgery on a pristine
   copy of the template (never a round-tripped save — see technical-notes §8
   and `scripts/surgery_example.py`), name it
   "<Initials> - <Property Name> - <Date> - Final.xlsx", and send both files.

## House rules (green thresholds)

- **Target IRR: 25%** (set Assumptions F48:H48 = 0.25). Green requires
  Project IRR (F5) > target — land it with a 2–4 point cushion.
- **Avg cash-on-cash (F7) > 10%.**
- **T-3 DSCR (I8) > 1.25 — UNLESS the deal is distressed.** Distress test:
  T-3 economic loss = T-3 vacancy % + concessions % + bad debt % (UW - F&C
  AC8+AC9+AC10). If **> 30%**, assume the deal goes to bridge financing and
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

- Remove the "black box": the PDF Output - F&C sheet has a conditional format
  on the floor-plan block (≈ B52:J77) that fills empty rows black. Delete that
  rule before exporting.
- Zero *new* formula errors after recalc (compare against the template's
  pre-existing error set; the template ships with known #REF!/#VALUE!
  artifacts on Validation, Tax Assessment, Yields, and unused Value-Add rows).
- Flag anything you could not refresh (sale-comp pages, agency datasets) in
  the summary to the user — never present stale prior-deal comps silently.
