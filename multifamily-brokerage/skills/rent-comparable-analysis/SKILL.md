---
name: rent-comparable-analysis
description: Build a rent comparable analysis for a multifamily subject property from the Dwellsy Comps API — pulling unit-level lease comps, collapsing them to property level, verifying each community's name and vintage, and positioning the subject's in-place and market rents against the set. Use whenever the user asks for rent comps, rent comparables, a market rent survey, a rent comp grid, "what do comps rent for", or asks to check whether a property's rents are at market. Not for SALE comps — use the sales-comps skill for those.
---

# Rent Comparable Analysis

Produce a defensible rent comp set for a multifamily subject: what comparable
units actually lease for, and where the subject's asking and in-place rents sit
against them.

The data source is Dwellsy's Comps API via
`scripts/dwellsy_comps_lookup.py`. Its field names, quirks, and verified
data-quality traps are documented in `references/dwellsy-api.md` — **read that
file before your first pull**, not after a result looks strange.

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
- unit count and year built
- **per floor plan**: bed count, net SF, market/asking rent, net effective
  (in-place) rent, unit count

Market vs net effective is the point of the exercise — the gap is the loss to
lease, and it is usually the finding. Ask only for what is genuinely missing.

## Step 2 — Spend the pulls in a deliberate ladder

Stop as soon as you have a usable set; unspent pulls are not wasted.

| Pull | Query | When |
|---|---|---|
| 1 | subject address, beds spanning the subject's mix, `--radius 2 --months 6`, `--type apartment` | always |
| 2 | `--radius 5 --months 24`, drop `--type` | pull 1 returned 0, or too few properties |
| 3 | `--beds` widened, `--months 36` | still thin |
| 4-6 | **submarket sweep**: one pull per nearby town via `--lat/--lon` | the subject's own market has no coverage |

Always request CSV with `--flat` so the file opens as a normal table:

```bash
python scripts/dwellsy_comps_lookup.py lookup \
  --address "300 State Highway 78 S, Farmersville, TX 75442" \
  --beds 1-2 --radius 5 --months 24 --format csv --flat -o out/subject.csv
```

Two constraints that shape the ladder: **`--radius` is capped at 5 miles**, and
a bare city name is rejected. So the only way to reach a neighbouring town is a
separate pull centred on its coordinates — which is why a thin market costs
several pulls and why you must budget them up front.

Verify the geocode on the first pull. The request echo at the top of the CSV
carries the `latitude`/`longitude` the API actually used; if the comp set looks
off-target, that is where to look.

## Step 3 — Analyze

```bash
python scripts/analyze_comps.py --csv "out/*.csv" \
  --subject-lat 33.16369 --subject-lon -96.37884 \
  --name "Crossroads Terrace" --units 36 \
  --plan "1BR:572:900:711:14" --plan "2BR:780:1100:753:22"
```

`--plan` is `LABEL:SF:MARKET:EFFECTIVE:UNITS`. The script recomputes distances
from the subject (needed after a sweep, since each pull's `distance_miles` is
measured from its own centre), de-duplicates overlapping pulls, and prints:
coverage/QC flags, listing concentration, the property-level comp table,
listing- vs property-weighted medians, size-matched benchmarks, and the
subject's position against each.

Three things it will flag that you must carry into the writeup:

- **Rows are unit listings, not properties.** One lease-up with 59 listings can
  outvote 27 small properties. Quote the property-weighted median; report the
  listing-weighted one only if they agree.
- **Junk rents.** $41/mo and $17,000/mo rows are both real observations from
  live pulls. They are excluded from medians automatically; medians survive
  them, means do not.
- **Blank `year_built`** (often 10-60% of rows) is dropped silently by any
  vintage filter. Decide explicitly whether to keep those properties.

## Step 4 — Name and verify every community you will cite

**Dwellsy returns no property name at all** — `company_name` is the management
company — and its `community_unit_count` and `year_built` are unreliable on
small or recent assets. Verified failures: a property returned as "28 units,
built 1997" was actually 18 units built 2020; another as "297 units, blank
year" was 34 units built 2023. Large established properties reconciled well.

For each community in the comp table, WebSearch the street address to get the
property name, then confirm unit count and vintage. Report the Dwellsy value
alongside the verified one where they disagree. Never put an unverified vintage
in a deck.

Watch for one community split across several addresses: Dwellsy assigns the
whole community's unit count to each building, so four consecutive addresses
all reporting "31 units" are one 31-unit property, not four. Collapse them
before counting comps.

## Step 5 — Report

Lead with the answer, then the evidence:

- **the comp table** — name, city, distance, units, vintage, rent by floor plan
- **subject positioning** — market and net effective as a % of the comp median,
  per plan, plus the psf-implied rent for the subject's unit size
- **loss to lease** per plan, and any rent-roll anomaly (e.g. a 2BR carrying a
  $42 in-place premium over the 1BR against a $200 market premium)
- **the caveats that change the conclusion** — how far out the comps sit,
  whether they are a different vintage or product, how much of the set is
  active vs historical listings

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
