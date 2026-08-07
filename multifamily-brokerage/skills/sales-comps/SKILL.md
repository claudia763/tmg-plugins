---
name: sales-comps
description: Generate a multifamily Sale Comparables package (Excel comp export + comparable sale grid PDF) from TMG's two source workbooks — "All Sales Comps" (the comps universe) and "Combined Fannie and Freddie Sales Comps" (refinance comps, used only for cap-rate drift). Use this whenever the user asks to run sales comps, sale comparables, a comp grid, comp export, a comparable sale analysis, or a CMA for an apartment/multifamily property — especially when they attach or reference either source workbook. The skill geocodes the subject, scores every comp on distance + vintage + unit count + sale date, trims price outliers, selects the best 5, and produces an Excel export with the best 5 at the top plus a one-page comparable sale grid PDF.
---

# Sales Comps (Python-native)

Produce a 5-comp Comparable Sale Analysis for a multifamily subject property.
Two inputs, two outputs — everything computed in Python, no Excel links:

**Inputs** (user attaches; ask if missing):
1. **`All Sales Comps.xlsx`** — the comps universe. One row per sale with
   Property Name/Address/City/State/ZIP, Unit Count, Year Built, Sold Price,
   Sold Price/Unit, Sale Date, Building SF, Avg Unit SF, Info Source,
   Latitude, Longitude.
2. **`Combined Fannie and Freddie Sales Comps.xlsx`** — agency loan
   originations (mostly refinance comps). Used **only** for cap-rate drift
   analysis, never as sale comps.

**Outputs**:
1. **`<Deal> - Sale Comp Export.xlsx`** — client-facing. 'Comp Export' sheet
   (Property Name | Address | City | State | ZIP | Units | Year Built |
   Avg Unit SF | Sale Date | Sold Price | $/Unit | Distance — no Comp/Rank
   columns, no scoring internals). The grid's 5 comps sit at the top
   (pale-gold rows); beyond them, comps priced more than 1 SD from the grid's
   indicated $/unit are screened out (nothing that undercuts the quoted
   number goes to a client), and if more than 20 rows survive the list is
   trimmed to the top 15. Plus a 'Comp Grid' sheet (the Subject-vs-Comp-1..5
   adjustment grid with live in-sheet formulas).
1b. **`<Deal> - Underwriting Sale Data.xlsx`** — internal, feeds the
   underwriting model. The SAME curated pool as the client export (grid 5
   first, ±1 SD screen, 15-row cap), in the exact legacy 'Output Analysis
   Data' schema (Property Name … Latitude, Longitude, Lat1_Rad, Lon1_Rad,
   Distance (mi.), DistancePoints, AgeSpread, AgePoints, DaysSinceSale,
   DatePoints, TotalPoints), with Lat1_Rad/Lon1_Rad as live `=Lat/180*PI()`
   formulas. The full unscreened list stays available in selection.json if
   ever needed.
2. **`<Deal> - Comparable Sale Grid.pdf`** — one-page portrait PDF matching
   TMG's traditional export: logo, gold property band, sectioned grid,
   gold indicated-value block, footnotes, and a map of the subject + comps.
   The map uses a real street basemap (CartoDB Positron via `contextily`)
   when the environment has open internet; in restricted sandboxes it
   automatically falls back to an offline to-scale coordinate plot with
   distance rings. Don't fight a tile failure — the fallback is expected
   behavior in cloud sessions.

## Workflow

### Step 1 — Gather subject inputs

Need: deal/property name, street address, city, state, zip, year built, unit
count, average unit size (SF). Ask only for what's genuinely missing.

### Step 2 — Geocode the subject (online search)

Find the subject's latitude/longitude with WebSearch/WebFetch (e.g. the US
Census geocoder `https://geocoding.geo.census.gov/geocoder/locations/onelineaddress?address=<url-encoded full address>&benchmark=Public_AR_Current&format=json`,
or a web search for the address's coordinates). Never fetch URLs with
bash/curl/python.

Sanity-check the result: it must fall inside the subject's state, and comps in
the same city should end up ~0–15 miles away. If the top-ranked comps are all
far away, the geocode is wrong — re-search before continuing.

Geocoders miss renamed streets (e.g. "Al Lipscomb Way" was "Grand Ave"). If
the lookup returns nothing, search the web for the address to find its former
name or a building at the same address, and geocode that instead.

### Step 3 — Score, filter, and select comps

```bash
python scripts/select_comps.py \
  --comps "All Sales Comps.xlsx" \
  --fannie-freddie "Combined Fannie and Freddie Sales Comps.xlsx" \
  --name "Vue Fitzhugh" --address "2819 N Fitzhugh Ave" --city Dallas --state TX \
  --zip 75204 --year-built 2004 --units 226 --avg-size 806 \
  --lat 32.8112 --lon -96.7768 \
  --out selection.json
```

The script (deterministic — do not re-implement this logic ad hoc):
- computes each comp's distance (spherical law of cosines, R=3959 mi), drops
  comps >175 mi, sold >3 years ago, or missing price/coords;
- scores **Distance + Age + SaleDate + UnitCount** points (brackets in
  `references/scoring.md`) and sorts by TotalPoints;
- excludes the subject's own past trade(s) (reported as `subject_self_sales` —
  always mention a recent subject trade to the user, it's pricing-relevant);
- takes the best 10, trims comps whose $/unit is more than 1.0 std dev from
  the mean of those 10, then keeps the **top 5** remaining (backfills from
  rank 11+ if the trim leaves fewer than 5);
- computes **cap-rate drift** from the Fannie/Freddie workbook: filters to the
  subject's **State and Vintage**, averages cap rates in trailing 12-month
  origination windows, and derives per-comp drift in basis points (falls back
  to wider scopes when the sample is thin — `cap_rate.scope` records which).

Read the printed summary and `selection.json`. Confirm the top comps make
sense (same metro, sane $/unit) before exporting.

### Step 4 — Build the exports

```bash
python scripts/export_comps.py --selection selection.json --outdir output/
```

### Step 5 — Verify and deliver

```bash
python scripts/verify_exports.py output/ --selection selection.json
```

This recalculates the Comp Grid through LibreOffice and confirms the sheet's
formulas reproduce the Python-computed indicated value; it must pass before
delivery. Then send both files and report a short summary: the 5 selected
comps (name, distance, units, year, sale date, $/unit), which outliers were
trimmed and why, the cap-rate drift applied, and the indicated value/unit and
total.

## Tuning

Common user asks map to CLI flags: "only comps sold in the last N months" →
`--max-days-since-sale` (default 1095 — the standing 3-year cutoff), "stay
within N miles" → `--max-distance`, "trim harder/softer" → `--outlier-sd`.
Everything else (point brackets, unit-count handling, drift fallbacks) lives
in the `CONFIG` block at the top of `select_comps.py`, documented in
`references/scoring.md` — edit constants there rather than hand-picking
comps; keep the pipeline reproducible.
