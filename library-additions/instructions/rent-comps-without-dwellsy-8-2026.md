# Running a rent comparable analysis when the Dwellsy API is unavailable (8/7/2026)

Covers: how to deliver the full TMG rent comp package — the branded PDF and the
Rent Comparison Grid workbook — when `dwellsy_comps_lookup.py` cannot run,
without inventing a new deliverable format and without misattributing the data.
Read alongside `rent-comparable-analysis/SKILL.md` and
`references/dwellsy-api.md`. Worked on **The Vintage, 2037 NW 26th St,
Oklahoma City, OK 73106** — 66 units, 1953.

## 1. Check for the key before you plan the pull ladder

`DWELLSY_API_KEY` is **not provisioned on every host.** On this box it is absent
from the repo-root `.env` (which carries only the mail/Anthropic/Dropbox/GitHub
keys), absent from `~/.tmg-toolchain.env`, and absent from the environment. The
`rent-comp-runs/` samples in `library-additions/` were produced on a host that
had it, so their presence proves nothing about the current machine.

Verify in one command, before spending any planning effort on the six-pull
ladder:

```bash
python scripts/dwellsy_comps_lookup.py lookup --beds 1 --address "<subject>" --dry-run
```

A missing key fails fast and prints every path it searched. `budget --reset`
succeeds regardless — it only touches the local state file — so **a clean
budget reset is not evidence that the API works.**

The Yardi Matrix MCP server is the other obvious substitute and it needs
interactive OAuth, so it is unavailable in a headless email-agent run too. Say
so in the reply rather than silently working around it.

## 2. The fallback: a point-in-time listing survey, written into Dwellsy's schema

Do **not** hand-roll a new spreadsheet. Build the comp set from public listing
sources and write it into the 24-column Dwellsy flat-CSV schema; the entire
downstream toolchain (house-rule trimming, property-level collapsing, the PDF,
the workbook) then runs unmodified and the deliverable is identical in form to
an API-sourced one.

`scripts/survey_to_flat.py` (in this folder's `scripts/`) does the conversion.
Its input is a survey CSV, one row per advertised floor plan:

```
property_name,street_address,city,state,zip,latitude,longitude,year_built,
total_units,plan_label,bedrooms,bathrooms,square_feet,asking_rent,
utilities_included,amenities,listing_status,source_url
```

It emits both the flat CSV and the `names.csv` the exporter needs — every
property in a hand survey is name-verified by definition, so Step 4 of the
SKILL is satisfied as you collect, not afterwards.

```bash
python survey_to_flat.py --survey survey.csv --out out/comps_flat.csv \
    --names-out names.csv --survey-date 2026-08-07 \
    --subject-lat 35.496423 --subject-lon -97.547025
```

Gather the survey with parallel subagents on **different angles**, not the same
angle repeated — one on Apartments.com/Zillow, one on the operator and
secondary-portal side (RentCafe, Rent.com, Zumper, the local management
companies), one on property identity and vintage via assessor/GIS records. The
third matters most: the urban vintage rule needs a real year built for every
comp, and the portals publish it inconsistently.

## 3. Two exporter flags this needs — `--source-label` and `--methodology`

`rent_comp_export.py` hardcoded "Dwellsy Comps API" into the XLSX note, the PDF
page footer, the grid footnote and the trending footnote. Shipping those on a
survey-sourced package **names a vendor that did not supply the data.** The
patched copy takes:

- `--source-label "..."` — how the source is named on every deliverable
- `--unverified-source "..."` — the label on `*`-flagged unit counts/vintages
- `--methodology <json>` — `[[heading, body], ...]`; replaces the trending page

The third exists because a survey is a **single-date snapshot**. With every row
stamped the same day, `chart_trend()` returns `None` (it needs 2+ quarters per
bed type) and the prior-vs-recent movement table renders empty — so page 5
degrades into a near-blank page carrying a false source line. `--methodology`
swaps in a "Data Sources & Methodology" page instead, which is the page a
client actually needs on a survey-sourced package.

These changes are backward-compatible: the defaults reproduce the original
Dwellsy wording and the trending page exactly. They belong upstream in the
plugin's `rent_comp_export.py`; until they land there, re-apply them in the
job-local copy.

## 4. Get the seller's own OM if one exists — it beats every portal

Before surveying, search for a prior offering memorandum on the subject. On The
Vintage the Price Edwards listing page 404'd but **the OM PDF was still live**
and reachable from a Wayback capture of the dead page. It carried, on two
pages, what no portal publishes:

- the **unit mix and in-place rents** (2 × 1x1 @ $683, 60 × 1x1 @ $684,
  4 × 2x1 @ $756 — 66 units, $688 avg in place, $733 avg market), which
  reconciled exactly to the printed $1.30/$1.39 per SF
- a **five-property lease comp set** in the same submarket with unit counts and
  rent/SF per plan

Portals publish "starting at" floors, not the unit mix, and never in-place
rents. Routes that worked: `curl` the PDF then `pymupdf` (WebFetch rejects it —
15 MB exceeds the 10 MB content limit); `curl` for web.archive.org (WebFetch is
blocked there); the **r.jina.ai text proxy** for
`docs.oklahomacounty.org/AssessorWP5/AN-R.asp?PROPERTYID=<id>`, which WAF-blocks
direct requests. The assessor detail URL takes `PROPERTYID`, not the account
number.

An OM's lease comps are a **prior-dated** set (July 2025 here). Use them as the
year-over-year cross-check, not as the delivered grid — the delivered grid
should be current. Never mix two survey dates inside one grid.

## 5. Derive comp SF from rent/SF when the OM omits it

OM lease comp pages typically print rent and rent/SF but not square feet.
`SF = rent / (rent per SF)` recovers it and is worth doing — the subject's plans
were 500-600 SF against a 628 SF comp-median 1BR, and without SF you cannot see
that the subject is being compared to materially larger product. Round to the
nearest foot and label the figure "implied".

## 6. What to say in the reply

State the substitution in one line, plainly: the API was unavailable, the set is
a point-in-time survey of published asking rents from named sources, and asking
rents are not leases. Do not present a survey as if it were lease-level data,
and do not quietly let the reader assume the usual Dwellsy pipeline ran.

## Related

- `rent-comparable-analysis/SKILL.md` — the workflow, house rules, pull budget
- `references/dwellsy-api.md` — the API's behaviour and data-quality traps
- `scripts/survey_to_flat.py` — the survey → flat-CSV converter
