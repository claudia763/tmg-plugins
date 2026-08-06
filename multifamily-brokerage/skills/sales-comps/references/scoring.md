# Comp scoring & selection spec

This reproduces the CMA workbook's original Power Query logic (`Output Analysis
Data` query) and extends it per TMG's instructions. Change behavior in the
`CONFIG` dict at the top of `scripts/select_comps.py`.

## 1. Distance

Subject coordinates come from an online geocode (Step 2 of the workflow); every
comp row in `All Sale Comps` already carries Latitude/Longitude. Distance uses
the spherical law of cosines with R = 3959 mi — the same math as the original
query, so results match what the workbook produced when its Bing geocoder still
worked:

```
d = acos( sin(lat1)·sin(lat2) + cos(lat1)·cos(lat2)·cos(lon2−lon1) ) · 3959
```

Comps ≥175 mi away are excluded outright (original hard filter), as are rows
missing coordinates, sale price, or unit count.

Comps missing average unit SF (after falling back to Building SF ÷ units), year
built, or sale date stay in the scored table but are **ineligible for the final
5** — the grid's adjustment formulas would produce garbage for them (e.g. a
zero avg-unit-SF comp gets a +25% size adjustment).

## 2. Blended points (distance, vintage, sale date, unit count)

| Distance (mi) | Pts | | Age spread (yrs) | Pts | | Days since sale | Pts | | Unit-count dev | Pts |
|---|---|---|---|---|---|---|---|---|---|---|
| < 1 | 100 (full) | | < 10 | 40 | | < 365 | 30 (full) | | < 25% | 30 |
| < 3 | 50 (half) | | < 20 | 30 | | < 730 | 15 (half) | | < 50% | 20 |
| else | 25 (quarter) | | < 30 | 20 | | < 1095 | 8 (quarter) | | < 100% | 10 |
| | | | < 40 | 10 | | excluded | — | | else | 0 |
| | | | else | 0 | | | | | | |

Distance and sale-date follow TMG's full/half/quarter weighting (Aug 2026).
Comps sold more than 3 years ago are **excluded outright**
(`max_days_since_sale`, default 1095; relaxable via `--max-days-since-sale`) —
that hard cutoff, not outsized points, is how staleness is handled. TMG tried
weighting recency at 100 pts and rejected it: it pulled in recent sales from
35 miles out. Distance stays the dominant criterion; don't "improve" this by
rebalancing without asking. The 175-mile hard filter still bounds the
universe. Age-spread brackets are the workbook's originals. The unit-count
column is this skill's addition (deviation = |comp units − subject units| /
subject units) so all four criteria are blended, as TMG specified. A hard
unit-count range is also available (`hard_unit_range`, e.g. `(0.4, 10.0)`),
off by default.

`TotalPoints` = sum of the four. Sort: TotalPoints desc, then distance asc,
then recency. Note the written `Output Analysis Data` tab keeps the original
24-column schema, so its `TotalPoints` column includes unit points even though
there is no separate UnitPoints column — TotalPoints may exceed
Distance+Age+Date points by up to 30.

## 3. Best 10 → outlier trim → top 5

Take the 10 highest-scoring comps. Compute the mean and population standard
deviation of their **$/unit**. Drop any comp whose $/unit deviates from the
mean by more than `outlier_sd` (default **1.0**) standard deviations — these
are the non-market trades (distressed sales, portfolio allocations, land-value
deals) that would skew the average. Keep the top 5 survivors by score; if the
trim leaves fewer than 5, backfill from rank 11+ (still applying the $/unit
band). The JSON records exactly who was trimmed and why so the user can be told.

## 4. Cap-rate drift (Combined Fannie & Freddie data)

Source: the CMA's `AgencyDrift` tab (a filtered extract of TMG's "Combined
Fannie and Freddie Sales Comps" database), or a standalone copy of that
database via `--fannie-freddie`.

Per TMG's instruction, filter on **State and Vintage** (subject state; year
built ± `drift_vintage_window`, default 10). Average the `Cap Rate` of loans
originated in each trailing 12-month window (0–1yr ago, 1–2, 2–3, 3+). A
comp's drift = (current-window avg − its sale-window avg) × 10000, rounded to
10 bps — matching the underwriting model's `ROUND(…,-1)` formula. Comps sold
within the last year get 0.

If a window has fewer than `drift_min_sample` loans, the filter widens in
order: state+vintage → state only → national+vintage → national. The
`cap_rate.scope` field in selection.json records which level was used — mention
it to the user when it isn't state+vintage.

In the output grid, row 25 holds the drift in bps and row 26 converts it to a
price adjustment against the current average cap rate (cell `AI1`):
`adj = −(drift/10000) / current_cap` — i.e. a comp that sold when cap rates
were 40 bps lower than today gets marked down by 40bps/current-cap ≈ 7% of
its $/unit.
