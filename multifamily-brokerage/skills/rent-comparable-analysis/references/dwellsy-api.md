# Dwellsy Comps API — pulling rent comps

Reference notes for the `rent-comparable-analysis` skill; `SKILL.md` in the
folder above has the workflow, this file has the API's behaviour.

Use `scripts/dwellsy_comps_lookup.py` whenever a job needs rent comparables for a
subject property (BOV/OM comp pages, underwriting market rent, rent-vs-market
checks). It wraps the whole Dwellsy Comps API — spec at
<https://comps-api.dwellsy.com/doc> (raw: `/doc/openapi.json`).

## The key

`DWELLSY_API_KEY` lives in the **repo-root `.env`** (gitignored). The script finds
it automatically: environment variable first, then the nearest `.env` walking up
from the working directory, then from the script's own folder.

**Never save the key inside `library-additions/`** — every file here is pushed to
the tmg-plugins GitHub repo after each job.

## Commands

```bash
S=scripts/dwellsy_comps_lookup.py

# CSV of 1-3 bed apartment/house comps within 2 miles, last 6 months
python $S lookup --address "770 S 2780 E Street, Saint George, UT" \
    --beds 1-3 --radius 2 --months 6 --type apartment house \
    --format csv -o comps_subject.csv

# same thing by coordinates, printed to the terminal as a table
python $S lookup --lat 38.9764706 --lon -94.5718538 --beds 1-4

python $S lookup --beds 2 ... --dry-run   # show the request body, send nothing
python $S history                         # recent reports for this key
python $S report <request-id> -o rerun.csv  # re-download a past report
```

Useful flags: `--sqft 700-1200`, `--photos`, `--limit N` (preview rows), `--raw`
(dump the API JSON), `-v` (log each HTTP call), `--format s3 --s3-filename key`.

## Endpoint behavior worth knowing

- `bedrooms_min`/`bedrooms_max` are the only required fields — pass `--beds 0-3`
  when studios count. Everything else narrows the search.
- A location is required in practice: either `--address` (Dwellsy geocodes it) or
  both `--lat` and `--lon`.
- `radius` defaults to **2 miles** and is **capped at 5** — anything larger is
  rejected with `ValidationError: "radius" must be less than or equal to 5`. In
  a thin market that makes `months` your only lever for depth; 24 works and
  roughly doubled the set in Paris, TX vs 12.
- `address_type` accepts `apartment`, `house`, `mobile`. Omit it to include all;
  for multifamily BOVs `--type apartment` alone is usually the right comp set.
- The POST returns an envelope (`status`, `requestId`, `compsCount`, `url`), not
  the comps — the report itself is at `url`. The script follows it for you and
  deliberately drops the bearer token on that hop, since it is typically a
  presigned S3 link on another host.
- `--format s3` uploads to the Dwellsy-side bucket and requires `--s3-filename`.
- Rate limited (HTTP 429) and 5xx responses are retried with backoff, honoring
  `Retry-After`; HTTP 403 means the key is missing/wrong, or the request id
  belongs to a different key.
- `compsCount: 0` is a normal answer, not an error — widen `--radius`/`--months`
  or loosen `--beds` before assuming the API failed.

## What the export actually contains

Confirmed against a live 722-comp pull (Montrose at Fitzhugh, 2819 N Fitzhugh
Ave, Dallas — 0-2bd, apartment, 2mi, 6 months, Aug 2026).

**The CSV is not a flat table.** It is a three-section report:

```
Comp Request Details :     <- echo of your request, incl. the geocode used
Comp Analysis              <- comps_count, avg_bedrooms, avg_bathrooms,
                              avg_square_feet, avg_price, avg_price_per_square_feet
Comparables Details        <- the 24-column table, one row per unit
```

Opening it straight in Excel/pandas gives a one-column mess. Pass `--flat` to
write just the comparables table; the script previews it correctly either way.

The 24 columns (JSON uses the same names): `id`, `address_type`, `latitude`,
`longitude`, `address_1`, `address_2`, `address_city`, `address_state`,
`address_zip`, `zip_plus4`, `bedrooms`, `bathrooms`, `square_feet`,
`year_built`, `company_name`, `listing_amount`, `price_per_sf`,
`last_listing_creation_time`, `last_listing_deactivation_time`,
`property_listing_status`, `distance_miles`, `url`, `amenities`,
`community_unit_count`.

Notes: rent is `listing_amount` (not "rent"/"price"); there is no unit-mix or
occupancy data. `amenities` is a JSON array in JSON output and a
**semicolon-separated** string in CSV. Numbers come back as strings in JSON —
cast before doing math. CSV and JSON of the same query are row-for-row
identical, and `price_per_sf` reconciles to `listing_amount / square_feet`.

## Data-quality caveats (verified, they will bite a BOV)

1. **Rows are individual unit listings, not properties.** 722 rows covered only
   180 addresses; one lease-up contributed 59 rows. So the `Comp Analysis`
   averages are *listing-weighted* and tilt toward whichever building had the
   most turnover. For 1bd on that pull: $1,605 listing-weighted vs **$1,390
   property-weighted** — a 15% swing. Average by property before quoting a
   market rent.
2. **Junk lows and luxury highs.** That pull included three $41/mo rows (psf
   0.05) and a $17,000/mo penthouse. Filter roughly $700-$6,000, or trim by psf,
   before averaging. Medians survive this; means do not.
3. **Most rows are `inactive`** (522 of 722) — historical listings inside the
   lookback, not current availability. Medians barely moved (+0-4%) when
   restricted to active, so keep them for sample size, but quote *current*
   asking rents from `property_listing_status == "active"` only.
4. **Vintage skew.** Median `year_built` was 1984 and only 19% of rows were
   2015+. Against a modern subject the unfiltered median ($1,550 / $2.00 psf)
   badly understates: 2015+ stock ran $3,300 / $2.76 psf, and 2015+ *and* 100+
   units ran $3,550 / $3.68 psf off just 4 properties. Always filter to the
   subject's vintage and community size, then check how many properties are
   actually left.
5. `year_built` is blank on ~9% of rows and `address_2` on a few, and a vintage
   filter drops those rows silently — decide whether to keep them. On the
   Montrose pull, a +/-10 year window around the subject's 2005 vintage left 192
   of 722 rows, but 82% of them came from just three buildings, so count
   *properties* (30, or 10 once you require 100+ units) before deciding the
   sample is deep enough. `distance_miles` is reliable (matches haversine from
   the echoed geocode to 0.0005 mi).
6. The API geocodes `--address` silently — check the echoed `latitude`/
   `longitude` in the request section if the comp set looks off-target. The
   subject itself did not appear in its own comp set. A bare "City, ST ZIP" is
   rejected as `Bad address`: pass a street address, or use `--lat/--lon` to
   skip geocoding entirely (the only way to sweep a submarket, given the 5-mile
   radius cap — run one pull per town and recompute distances yourself).
7. **`community_unit_count` and `year_built` are not trustworthy on small or
   recent properties — verify before citing.** Checked against six named
   communities near Farmersville, TX: The Villages at Fate came back as "28
   units, built 1997" against an actual 18 units built 2020, and Adaline at the
   Station as "297 units, blank year" against an actual 34 units built 2023.
   Large established properties did reconcile (Southridge 163 vs 160 actual,
   1986; Villas at Wylie 314 vs 303, 2008). Dwellsy also returns no property
   name at all — `company_name` is the management company. Look the name and
   vintage up by street address before either reaches a deck.
8. **Coverage collapses in small markets.** Farmersville, TX (pop. ~3,600)
   returned *zero* apartment comps inside the 5-mile cap: 58 of 60 records were
   single-family houses, 37 of them 4-bedroom, and nothing under 900 sf. Do not
   read an empty or house-only result as "no market" — it means Dwellsy has no
   coverage there, and the comp set has to be built from surrounding towns.

## Pull budget

`lookup` bills a report, so the script caps one analysis at **6 pulls**
(`DEFAULT_PULL_BUDGET`; override with `--budget` or `DWELLSY_PULL_BUDGET`). The
7th exits 3 without calling the API. Reset at the start of each analysis with
`budget --reset`, inspect spend with `budget`. When it runs out, work with the
reports already pulled or reply "Can't find comps" — re-querying a market with
no coverage returns the same nothing.

For the full workflow built on this API — pull ladder, submarket sweeps,
property-level analysis, and the name/vintage verification step — follow
`SKILL.md` in the folder above, and run `scripts/analyze_comps.py` on the
pulls rather than hand-rolling the aggregation.

## In a comps workflow

Save the CSV into the job's work folder, then filter it down to the comp set you
actually cite (drop far-out submarkets, mismatched vintages, and unit types the
subject doesn't have) before it reaches a deck. The API returns what is nearby,
not what is comparable — that judgment is still ours. `instructions/comp-map-generation.md`
covers plotting the selected comps.
