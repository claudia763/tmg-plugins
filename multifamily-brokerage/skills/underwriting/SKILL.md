---
name: underwriting
description: >
  Use this skill to underwrite a multifamily property in The Multifamily Group's
  Excel underwriting model (the "Initials  Property Name  Date" / Aden Crest
  template). The valuation itself is run first in the bundled Python engine
  (scripts/tmg_valuation.py — a faithful port of the model's hidden
  'UW - F&C' engine) to tune assumptions and price against TMG's green
  thresholds, then the rent roll, trailing financials, and the settled
  Assumptions / Value-Add / Factors inputs are copied into the deliverable
  Excel model and verified with a LibreOffice recalc. Trigger whenever the
  user asks to "underwrite" a deal or model, "populate the model," "run the
  underwriting," or attaches a rent roll + T-12 alongside the TMG
  underwriting model workbook.
---

# TMG Underwriting Model Skill

## What this skill produces

1. The completed underwriting model (.xlsx) — named
   **"<Initials> - <Property Name> - <Date> - Final.xlsx"** — populated
   directly with openpyxl and recalculated/verified in LibreOffice. There is
   **no separate "surgery" build anymore**: we run the deliverable in
   LibreOffice, which cannot refresh Power Query in any case, so the
   openpyxl-saved workbook IS the deliverable. A worked example of the
   finished product is bundled at
   `references/Initials - Property Name - Date - Final.xlsx`.
2. A PDF of the **PDF Output - F&C** tab only (landscape, fit-to-width),
   exported with the bundled `scripts/export_pdf.py`.

Read `references/model-map.md` for the exact cell map before touching the
workbook, and `references/technical-notes.md` for the openpyxl/LibreOffice
gotchas — several are load-bearing (external defined names, array formulas,
row-capacity limits).

## LibreOffice-first ground rules

The model is authored, recalculated, and consumed in **LibreOffice** — not
Excel. Keep the workbook inside LibreOffice's comfort zone:

- Do not rely on, add, or try to preserve Power Query / DataMashup,
  workbook connections, query-table refresh, or pivot-cache refresh. An
  openpyxl or LibreOffice save strips them; nothing downstream needs them.
- No Excel-only or "clever" functionality in anything you write into the
  workbook: no dynamic arrays / spill functions (XLOOKUP, FILTER, LET,
  LAMBDA), no volatile whole-column tricks, no new external links, no new
  structured-table formulas. Plain cell references and the functions the
  template already uses are the ceiling.
- Values beat formulas: where a block only reports data (comp tables,
  Raw Rents mirrors, anything LibreOffice mangles per technical-notes
  §1–3), write literal values.
- Recalc and PDF export happen through LibreOffice headless (xlsx skill
  `recalc.py`, `scripts/export_pdf.py`); expect the template's ~380
  pre-existing error cells and require zero NEW errors.

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

**Phase A — value the deal in Python (`scripts/tmg_valuation.py`).**
All assumption tuning happens here; iteration is instant instead of a ~2-min
LibreOffice recalc per pass.

1. Copy `scripts/tmg_valuation.py` into the job folder and fill its config
   blocks from the source documents — they mirror the model tab-for-tab:
   - `PROPERTY` — units, rentable SF, market rent/unit, total tax rate from
     the CAD card (Master-tab facts).
   - `T12_MONTHLY` — monthly income and expense series aggregated from the
     T-12 statement by code (same lines as Final_T_12), oldest → newest.
   - `ASSUMPTIONS` — the Assumptions-tab blue cells (Year-1 operating
     inputs, expense overrides / "agency" benchmark flags, growth rates,
     the stabilized economic-loss input G59 (default 8.0%; drives the
     out-year vacancy walk-back), price, debt terms, syndication terms).
   - `VALUE_ADD_ITEMS` — include-flags for the Value-Add programs (top 3
     per category count, exactly like the tab). A 2018+ build gets no
     interior renovation program.
   - `FACTORS` — sale-comp cap rate plus cap-rate risk adjustments in bps
     (keep custom adjustments 25–75 bps and defensible).
2. Research the property (reviews, specials, amenities; Rent Data
   `Reviews Analysis` sentiment vs the T-12) to justify Value-Add
   selections and risk factors before flipping flags.
3. Run the script and tune purchase price / LTV / assumptions against the
   house rules below. The script prints the full proforma, Project IRR,
   Avg Cash-on-Cash, T-3 DSCR, and a validation table;
   `solve_price_for_irr()` goal-seeks the pricing band.

**Phase B — populate the deliverable model.**
Work on the actual deliverable copy from the start; no disposable working
copy and no zip-surgery rebuild afterward.

4. **Master tab** — property info in B1:B19 (+ C3 state, D3 zip). Tax rates
   B11:B17 from the CAD card's taxing-units table; total must reproduce the
   card's aggregate rate. Full market value = current-year net appraised.
   Date of rent roll = the "as of" date inside the rent roll file, not its
   filename.
5. **FinalRR + Final_T_12** — paste rent roll rows into FinalRR C9:Y…, floor
   plan summary into FinalRR BA9:BZ… and the `Final_RR_Floor_Plan` table;
   aggregate the T-12 statement by code into Final_T_12 (fold `rt` into the
   `ro` row; set the month dates in row 2; column O = row totals). Cross-check
   rev/exp/NOI against the statement's own `rev`/`exp`/`noi` rows — these must
   be the same numbers already loaded into `T12_MONTHLY` in Phase A.
6. **Comp tables** — TableRecentLeases from Unit-Level Data (comps only,
   exclude the subject; filter to leases in the last 12 months + active
   listings; cap total rows at 996 — see technical notes); TablePropertyData
   from the Rent Comps sheet (subject row, "Comp Average" row, then comps,
   with address split into street/city/state/zip). Then value-fill the
   Raw Rents mirror columns (see model map).
7. **Rent comp selection** — recalc, then "x" in Rent Comparison column AK for
   the best 4–5 comps: highest-renting cluster, dropping low-rent outliers and
   any single comp priced far above the pack. The PDF summary table only shows
   the first four selected; a fifth still feeds the averages.
8. **Transcribe the settled valuation** — copy the final Phase-A inputs into
   the model verbatim: Assumptions-tab blue cells (including strike price and
   loan terms), Value-Add column-C "x" flags, and Factors column-J flags plus
   any custom rows 24–25. Do not re-tune in the spreadsheet; the Python run
   is the source of truth.
9. **Verify + deliver** — LibreOffice recalc, then read F5 (IRR), F7 (CoC),
   I8 (T-3 DSCR) with `data_only=True` and confirm they match the Python
   outputs within ±2% (they normally match to rounding; a bigger gap means a
   transcription error — find it, don't shrug). Zero new formula errors.
   Export the PDF Output - F&C tab with `scripts/export_pdf.py`, verify page 1
   visually, name the workbook
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

## Value-Add selection workflow

The Value-Add tab now carries an **AGENT INSTRUCTIONS** column (column A) with
an INCLUDE IF trigger per row. The same triggers live in `VALUE_ADD_RULES`
inside `scripts/tmg_valuation.py` — that dict is the single source of truth
(conflicts are by driver NAME, not row number, so row inserts can't break
cross-references). Work the tab in three passes, then encode and validate.

**Pass 1 — documents (rent roll + T-12).** These triggers resolve from files
already parsed in Phase A; evaluate all of them before any web research:

- Final_RR: down/offline units (set the exact count), model/office/employee/
  storage units, renovation status split, W/D connections, physical occupancy
  vs comp average.
- T-12 per-unit screens: water > $65 (fixtures) / > $100 (plumbing), bad debt
  > 2% of GPR, payroll / insurance / refuse / common-area electric / total
  opex vs comp averages, fee income vs comp fee schedules.
- T-12 income detail — double-count guard: existing pet income, reserved
  parking, bulk internet, valet trash, renters insurance, telecom door fees,
  amenity fees. An existing line kills or shrinks the matching row.
- Owner-paid utility structure (which of water/electric/gas/trash the owner
  pays) — gates every RUBS, billback, submetering, and conservation row.

**Pass 2 — web (subject + market).** Property listings and county data:
year built (HVAC pre-1980 gate, EV 2015+ gate), unit count (the 200+ gates:
valet trash, lockers, compactor), amenity scan of subject and comp listings
(gates, pool, smart tech, carports, WiFi, garage/storage pricing), aerial for
excess land and parking layout, location drivers (university / hospital /
airport / tourist district → STR or Corporate/Furnished), and jurisdiction:
STR **prohibited vs license-required** (license-required, e.g. OKC, still
qualifies — a prohibition routes to Corporate/Furnished), submetering
legality, and HFC/PFC availability (**Texas only** — never hunt for an
Oklahoma equivalent).

**Pass 3 — comp tabs.** After the Rent Comparison is built: interior tier by
renovated-comp premium (≤$50 light / $75–125 moderate / >$125 premium —
pick exactly ONE tier), and treat a comp behavior as established when **at
least 2 of the selected comp set** show it (amenities, RUBS recovery,
billbacks, fee schedules).

**Selection rules.**

- Include **at most three rows per category** — the sheet and the engine
  count only the first three; the validator errors on more.
- `verify="site"` rows (carports, gating, hookups, laundry, compactor,
  submeter plumbing, LED fixture type, down-unit scope, solar, EV panels,
  garages, exterior amenities): may be *selected* on document evidence but
  flag them in the summary as pending physical verification; when in doubt
  model $0 and keep the narrative. `verify="flag"` rows (cell tower/billboard,
  HFC/PFC) are never modeled without an executed lease/LOI/structure docs.
- Always-include set: Pet Fees (small — $10–15 NOI-equivalent unless
  pet amenities support $30), Reserved Parking (unless already charged),
  Utility Bill Audit (modest recovery only). Each still gets the T-12
  double-count check first.
- Netting: "Reduce Opex to Comp Averages" double counts against ALL of its
  component rows — insurance, taxes, payroll, compactor, utility audit, LED,
  and both water-conservation rows. Pick the umbrella OR the components.
- Narrative rows (`effect="assumption"`): lease-up → `vacancy_pct`;
  reduce-opex/insurance → expense overrides or `"agency"` benchmarks;
  tax protest → `tax_assessment_factor`. Never wire dollars for these.
- Savings effects (`payroll_savings`, `electric_savings`, `trash_savings`,
  `water_savings`) only apply when that expense line's assumption is `None`
  (T-12-driven). If the line uses an override or `"agency"` benchmark, bake
  the saving into the override instead — the engine skips it by design to
  prevent double counting.

**Encode + validate.** Flip the include flags in `VALUE_ADD_ITEMS`, set real
`n` and `$NOI/unit/mo` on manual rows (Down-Unit, Non-Revenue, STR,
Corporate, EV — the validator errors on zeros), then run
`validate_value_add(state=<subject state>)` and clear every ERROR and
acknowledge every WARNING **before** `run_model()`. In Phase B, transcribe
the surviving selections as column-C "x" flags exactly (top-3 per category),
and give each included row a one-line justification with its evidence source
(RR / T12 / COMPS / WEB / SITE) in the summary to the user.
