---
name: rent-comparable-analysis
description: Build a rent comparable analysis for a multifamily subject property from the Dwellsy Comps API — pulling unit-level lease comps, collapsing them to property level, verifying each community's name and vintage, and positioning the subject's in-place and market rents against the set — then produce the client-facing deliverables (TMG-branded Rent Comparable Analysis PDF + Rent Comparison Grid Excel) via scripts/rent_comp_export.py. Use whenever the user asks for rent comps, rent comparables, a market rent survey, a rent comp grid, "what do comps rent for", or asks to check whether a property's rents are at market. Not for SALE comps — use the sales-comps skill for those.
---

# Rent Comparable Analysis

Produce a defensible rent comp set for a multifamily subject: what comparable
units actually lease for, and where the subject's asking and in-place rents sit
against them — plus the client-facing PDF and Excel grid deliverables.

The data source is Dwellsy's Comps API via
`scripts/dwellsy_comps_lookup.py`. Its field names, quirks, and verified
data-quality traps are documented in `references/dwellsy-api.md` — **read that
file before your first pull**, not after a result looks strange.

## House rules — comp selection (8/2026)

These are TMG rules, not suggestions. `scripts/rent_comp_export.py` enforces
all of them automatically.

1. **Never use sub-5-unit properties as multifamily comps.** Houses,
   duplexes, and fourplexes are excluded from the comp set entirely, whatever
   the market.
2. **Pull a MINIMUM 3-mile radius and a 24-month lookback on every API
   request** (the API caps radius at 5), then trim here — starting with
   vintage and unit count. Never trim by shrinking the query. The deep
   lookback exists for the TRENDING page: the exporter automatically trims
   the comp set itself to the freshest window (last 6 months, falling back
   to 12, then full — urban mode only) while the trend charts and the
   prior-vs-recent movement table use the full 24 months.
3. **Urban markets only** (comp selection plentiful — in practice, more than
   ~15 qualifying listings; the exporter auto-detects, or force with
   `--urban` / `--no-urban`):
   - **Vintage: no comps 10+ years older or newer than the subject.**
     Unknown-vintage comps are kept only if dropping them would leave fewer
     than 5 comps — and then must be verified (Step 4) before delivery.
   - **No 150+ unit complexes when the subject is under 100 units** — the
     amenity packages aren't comparable.
   - **Outlier pruning by standard deviation, not fixed dollar bands:** when
     the initial search returns more than 15 comps, listings are pruned per
     bed type at ±3σ (two passes). ±3σ kills junk rents ($41/mo, $17,000/mo)
     while keeping genuinely cheap product; use `--sd-cut 2` only when you
     want an aggressive trim and have verified you aren't deleting a real
     low-rent comp.
   - Filters apply as a **trim ladder** and relax rather than leave fewer
     than 5 comp properties.
4. Sparse/small-town markets (e.g. Brenham): urban rules stay OFF; the
   fixed sane-rent band handles junk and the submarket sweep (Step 2) builds
   the set. Say plainly when comps come from other towns.

## The pull budget — 6 API pulls per request, hard stop

Every `lookup` bills a report. One analysis gets **six**. The script enforces
this itself: the 7th `lookup` exits 3 without calling the API.

```bash
python scripts/dwellsy_comps_lookup.py budget --reset   # start of EVERY analysis
python scripts/dwellsy_comps_lookup.py budget           # check spend
```

When the budget is gone you have two options and no others:

1. **Build the analysis from what you already pulled.** Partial coverage with
   stated limits beats no answer. Say which cuts are thin.
2. **Reply "Can't find comps"** — see the criteria at the bottom — and say what
   you tried.

Do not keep re-querying with nudged filters. In a market with no coverage, the
7th pull returns exactly what the 6th did. Plan the ladder below before
spending pull 1.

## Step 1 — Subject intake

Collect, from the conversation, a rent roll, a floor plan summary, or an
underwriting file:

- street address (a real one — the API rejects bare "City, ST ZIP")
- unit count and year built (year built gates the urban vintage filter —
  check prior deal notes in `library-additions/instructions/` before asking)
- **per floor plan**: bed count, net SF, market/asking rent, net effective
  (in-place) rent, unit count

Market vs net effective is the point of the exercise — the gap is the loss to
lease, and it is usually the finding. Ask only for what is genuinely missing.

## Step 2 — Spend the pulls in a deliberate ladder

Stop as soon as you have a usable set; unspent pulls are not wasted.

| Pull | Query | When |
|---|---|---|
| 1 | subject address, beds spanning the subject's mix, `--radius 3 --months 24`, `--type apartment` | always (3 mi / 24 mo minimum — house rule) |
| 2 | `--radius 5`, drop `--type` | pull 1 returned 0, or too few 5+ unit properties |
| 3 | `--beds` widened, `--months 36` | still thin |
| 4-6 | **submarket sweep**: one pull per nearby town via `--lat/--lon` | the subject's own market has no coverage |

In an urban market, pull 1 at 3 miles is usually the whole search — count the
distinct **5+ unit properties** (and, if urban rules will apply, the ones
inside the vintage window) before spending pull 2. `--sqft` bounded around the
subject's unit sizes keeps a dense-market pull manageable.

Always request CSV with `--flat` so the file opens as a normal table:

```bash
python scripts/dwellsy_comps_lookup.py lookup \
  --address "4635 Werner St, Houston, TX 77022" \
  --beds 1-2 --radius 3 --months 24 --type apartment \
  --sqft 400-1250 --format csv --flat -o out/subject.csv
```

Two constraints that shape the ladder: **`--radius` is capped at 5 miles**
(minimum 3 per house rule), and a bare city name is rejected. So the only way
to reach a neighbouring town is a separate pull centred on its coordinates —
which is why a thin market costs several pulls and why you must budget them up
front.

Verify the geocode on the first pull. The request echo at the top of the CSV
carries the `latitude`/`longitude` the API actually used; if the comp set looks
off-target, that is where to look.

## Step 3 — Analyze

```bash
python scripts/analyze_comps.py --csv "out/*.csv" \
  --subject-lat 29.833 --subject-lon -95.39697 \
  --name "Werner Creek" --units 36 \
  --plan "1BR:623:945:895:29" --plan "2BR:898:1138:1094:7"
```

`--plan` is `LABEL:SF:MARKET:EFFECTIVE:UNITS`; the bed count is read from the
label's **leading digits** ("1BR", "2x1-930", "1x1 Classic" all parse as the
right bed count). The script recomputes distances from the subject (needed
after a sweep, since each pull's `distance_miles` is measured from its own
centre), de-duplicates overlapping pulls, and prints: coverage/QC flags,
listing concentration, the property-level comp table, listing- vs
property-weighted medians, size-matched benchmarks, and the subject's position
against each.

Three things it will flag that you must carry into the writeup:

- **Rows are unit listings, not properties.** One lease-up with 59 listings can
  outvote 27 small properties. Quote the property-weighted median; report the
  listing-weighted one only if they agree.
- **Junk rents.** $41/mo and $17,000/mo rows are both real observations from
  live pulls. They are excluded automatically; medians survive them, means do
  not.
- **Blank `year_built`** (often 10-60% of rows) is dropped silently by any
  vintage filter — and in urban mode the vintage filter is mandatory, so
  blank-year comps that matter must be verified (Step 4), not assumed.

## Step 4 — Name and verify every community you will cite

**Dwellsy returns no property name at all** — `company_name` is the management
company — and its `community_unit_count` and `year_built` are unreliable on
small or recent assets. Verified failures: a property returned as "28 units,
built 1997" was actually 18 units built 2020; another as "297 units, blank
year" was 34 units built 2023. Large established properties reconciled well.

For each community in the comp table, run
`scripts/rentcast_xref.py --csv "out/*.csv" -o xref.csv` first — it pulls
county-record `yearBuilt` / `unitCount` / owner per address from the RentCast
API (`RENTCAST_API_KEY` in the repo-root `.env`; calls are metered, the script
caps at 25/run). Year built is its strong suit; unit counts are often absent
and marketing names never appear in county records, so then WebSearch the
street address to get the property name and fill the gaps. A RentCast record
coming back "Single Family" for a complex address means the lookup hit a
unit-level parcel — verify that one by hand. Report the Dwellsy value
alongside the verified one where they disagree. Never put an unverified vintage
in a deck. In urban mode a verified year can move a comp in or out of the
vintage window — verify BEFORE trimming, not after.

Watch for one community split across several addresses: Dwellsy assigns the
whole community's unit count to each address, so four consecutive addresses
all reporting "31 units" are one 31-unit property, not four. (Live examples:
Brenham Park across 2 addresses; La Casita Houston across 15.) Record the
merge in the names CSV (below) so the exporter collapses them.

## Step 5 — Deliverables: `scripts/rent_comp_export.py`

Turns the pulls + subject into the client-facing package, enforcing every
house rule above:

```bash
python scripts/rent_comp_export.py \
  --csv "out/*.csv" --subject subject.json --names names.csv --out out
```

- `subject.json`: name/address/city/state/zip, `year_built`, `units`,
  `lat`/`lon` (from the pull-1 geocode echo), and `plans`
  (`label`, `bed`, `sf`, `market`, `effective`, `units` per floor plan).
- `names.csv` (`address_1,name,units,year_built,community`): the Step-4
  verification output — verified names/units/vintages, and a shared
  `community` value to merge multi-address communities.
- Key flags: `--urban` / `--no-urban` (force the auto-detection),
  `--vintage-window 10`, `--sd-cut 3.0`, `--min-comps 5`, `--min-units 5`,
  `--sf-variance 0.25`, `--comp-months 6` (urban-only comp recency window;
  trends always use the full pulled lookback).

Outputs, named after the subject:

1. **`<Subject> - Rent Comps.xlsx`** — "Rent Comparison Grid" (Property Name |
   Address | City | State | Zip | Year Built | # of Units | Avg. Size |
   Avg. Rent/Unit | Avg. $/SF | Amenities | Interior Quality; subject row
   first, shaded), "Granular Rent Comparison" (per-floor-plan blocks at
   ±`--sf-variance` SF), and the cleaned "Comp Listings".
2. **`<Subject> - Rent Comparable Analysis.pdf`** — TMG-branded (navy/gold +
   logo; chart series use the validated colorblind-safe palette, brand colors
   are chrome only): cover with key-stat band, the grid, comps by unit type
   with property-ranking bars vs the comp median, subject positioning with
   per-plan comp-supported rent and demonstrated upside, rent-vs-size scatter
   with per-bed dashed trend lines (legend below the plot; subject = navy
   diamonds with gold rim), and a trending page over the FULL 24-month
   lookback (quarterly median rent, listings per quarter, prior-vs-recent
   movement) while the comp pages use the fresh window.

Sanity-check the PDF pages visually (render with `pdftoppm`) before delivery.

## Step 6 — Report

Lead with the answer, then the evidence:

- **the comp table** — name, city, distance, units, vintage, rent by floor plan
- **subject positioning** — market and net effective as a % of the comp median,
  per plan, plus the psf-implied rent for the subject's unit size
- **loss to lease** per plan, and any rent-roll anomaly (e.g. a 2BR carrying a
  $42 in-place premium over the 1BR against a $200 market premium)
- **the caveats that change the conclusion** — how far out the comps sit,
  whether they are a different vintage or product, how much of the set is
  active vs historical listings, all-bills-paid comps whose rents embed
  utilities, and which filters were relaxed (and why)

Say plainly when comps come from stronger submarkets than the subject's; a
psf-based indication off the closest true vintage analog is usually more
trustworthy than a raw median pulled from a bigger town.

## When to reply "Can't find comps"

Say it, rather than dressing up a set that cannot support a conclusion, when
after the budget is spent:

- no apartment-type listings exist near the subject at any bed/size, **or**
- nothing in the data is within roughly ±35% of the subject's unit sizes, **or**
- fewer than ~3 distinct properties survive, **or**
- the only comps sit in materially different submarkets and you would be
  comparing a small-town asset to metro product

Report what you pulled, what came back, and what you would need — a wider
lookback, a different data source, or the user's own market knowledge. An
empty or house-only result means Dwellsy has no coverage there; it does **not**
mean the market has no rentals, and must never be reported that way.
