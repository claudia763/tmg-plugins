---
name: sales-comps
description: Generate a Sale Comparables workbook for a multifamily subject property from TMG's "Automatic CMA Analysis" workbook. Use this whenever the user asks to run sales comps, sale comparables, a comp grid, a comparable sale analysis, or a CMA for an apartment/multifamily property — especially when they attach or reference an "Automatic CMA Analysis" workbook or its "All Sale Comps" tab. The skill geocodes the subject, scores every comp on distance + vintage + unit count + sale date, trims price outliers, selects the best 5, and produces a "Sale Comparables Workbook" whose formulas are live-linked to the Automatic CMA Analysis file (not the underwriting model).
---

# Sales Comps (Sale Comparables Workbook)

Produce a 5-comp Comparable Sale Analysis for a multifamily subject property.
Two files come out of every run, and they must stay in the same folder:

1. **`Automatic CMA Analysis.xlsx`** (a per-deal copy of the input CMA workbook) — its
   `Inputs` tab is filled with the subject and its `Output Analysis Data` tab is
   recomputed with fresh distance/score data.
2. **`<Deal> - Sale Comparables.xlsx`** — the deliverable comp grid. Its formulas are
   external-linked **directly to the Automatic CMA Analysis file** (the original template
   linked to a full underwriting model; this skill relinks it).

The comps universe is the CMA workbook's **`All Sale Comps`** tab (each row already has
Latitude/Longitude). The selection logic replicates and extends the workbook's own
Power Query scoring (which no longer refreshes because its Bing Maps geocoder is dead).

## Workflow

### Step 1 — Gather subject inputs

Need: deal/property name, street address, city, state, zip, year built, unit count,
average unit size (SF). Pull them from the conversation, an underwriting file, or the
CMA workbook's `Inputs` tab. Ask only for what's genuinely missing.

### Step 2 — Geocode the subject (online search)

Find the subject's latitude/longitude with WebSearch/WebFetch (e.g. the US Census
geocoder `https://geocoding.geo.census.gov/geocoder/locations/onelineaddress?address=<url-encoded full address>&benchmark=Public_AR_Current&format=json`, or a web search for the address's
coordinates). Never fetch URLs with bash/curl/python.

Sanity-check the result: it must fall inside the subject's state, and comps in the same
city should end up ~0–15 miles away. If the top-ranked comps are all far away, the
geocode is wrong — re-search before continuing.

Geocoders miss renamed streets (e.g. "Al Lipscomb Way" was "Grand Ave"). If the
lookup returns nothing, search the web for the address to find its former name or a
building at the same address, and geocode that instead.

### Step 3 — Score, filter, and select comps

```bash
python scripts/select_comps.py \
  --cma "Automatic CMA Analysis.xlsx" \
  --name "Vue Fitzhugh" --address "2819 N Fitzhugh Ave" --city Dallas --state TX \
  --zip 75204 --year-built 2004 --units 226 --avg-size 806 \
  --lat 32.8112 --lon -96.7768 \
  --out selection.json
```

The script (deterministic — do not re-implement this logic ad hoc):
- computes each comp's distance (spherical law of cosines, R=3959 mi), drops comps
  >175 mi or missing price/coords;
- scores **Distance + Age + SaleDate + UnitCount** points (brackets in
  `references/scoring.md`) and sorts by TotalPoints;
- excludes the subject's own past trade(s) from the comp set (reported in the
  JSON as `subject_self_sales` — always mention a recent subject trade to the
  user, it's pricing-relevant);
- takes the best 10, trims comps whose $/unit is more than 1.0 std dev from the
  mean of those 10, then keeps the **top 5** remaining (backfills from rank 11+ if
  the trim leaves fewer than 5);
- computes **cap-rate drift** from the CMA's `AgencyDrift` tab (Combined Fannie &
  Freddie data): filters to the subject's **State and Vintage**, averages cap rates in
  trailing 12-month origination windows, and derives per-comp drift in basis points
  (falls back to all states / wider vintage when the filtered sample is thin — the
  JSON records which fallback applied).

Read the printed summary and `selection.json`. Confirm the top comps make sense
(same metro, sane $/unit) before building workbooks.

### Step 4 — Build the two workbooks

```bash
python scripts/build_output.py \
  --cma "Automatic CMA Analysis.xlsx" \
  --selection selection.json \
  --outdir output/
```

Uses the bundled template (`assets/Sale Comparables Workbook.xlsx`; override with
`--template`). Both workbooks are edited at the raw-XML level so Power Query,
formatting, and everything else survive. What gets rewritten is specified in
`references/relinking.md` — in short: the external link now targets the CMA file, the
helper grid pulls `Output Analysis Data` rows 1–50 by direct cell reference, the
Subject column pulls the CMA `Inputs` tab, the five selected comps are marked `x`,
and the Cap Rate Drift row gets the computed bps values.

### Step 5 — Verify and deliver

```bash
python scripts/verify_output.py output/
```

Fix anything it flags. Then send both files to the user (they must be kept side by
side or Excel shows #REF on refresh) and report a short summary: the 5 selected
comps (name, distance, units, year, sale date, $/unit), which outliers were trimmed
and why, the cap-rate drift applied, and the indicated value/unit ($ mean of the five
adjusted $/unit) with the indicated total.

## Tuning

Common user asks map to CLI flags: "only comps sold in the last N months" →
`--max-days-since-sale`, "stay within N miles" → `--max-distance`, "trim harder/softer"
→ `--outlier-sd`. Everything else (point brackets, unit-count handling, drift fallbacks)
lives in the `CONFIG` block at the top of `select_comps.py`, documented in
`references/scoring.md` — edit constants there rather than hand-picking comps;
keep the pipeline reproducible.

## Known limits

- The relinked workbook's "Manual" data-source mode is inert (its named ranges
  pointed at underwriting-model tabs that don't exist in the CMA file). The grid is set
  to Automatic. Details in `references/relinking.md`.
- The `AgencyDrift` tab is whatever the CMA copy last refreshed. If the user supplies
  a fresher "Combined Fannie and Freddie Sales Comps.xlsx", pass it via
  `--fannie-freddie` to use it directly.
