# Rent comps on a small-unit, old-vintage urban asset — The Vintage, OKC (8/7/2026)

Covers: three things that a literal reading of the rent-comp house rules gets
wrong on a 1953, 66-unit, 525-SF-average walk-up, and the two
`rent_comp_export.py` patches that fix the last of them. Read alongside
`rent-comparable-analysis/SKILL.md` and `rent-comps-without-dwellsy-8-2026.md`.
Worked on **The Vintage Apartments, 2037 NW 26th St, Oklahoma City, OK 73106**
— 66 units, 1953.

## 0. The Dwellsy key is host-dependent — check, don't assume either way

`rent-comps-without-dwellsy-8-2026.md` documents this same property being run
with **no** `DWELLSY_API_KEY` on the host. On this run the key **was** present in
the repo-root `.env` and the API worked normally. So the survey fallback is not
"the way we do The Vintage" — it was a host condition. Always run the one-line
`--dry-run` check first; take the API path whenever it is available, because the
listing-level depth (1,128 rows here) is what makes the trending page and the
size-matched benchmarks possible at all.

## 1. A 3-mile pull in a dense metro is a *pull* radius, not a *comp* radius

House rule 2 requires a minimum 3-mile / 24-month pull. `rent_comp_export.py`
has **no distance filter**, so whatever the pull returns reaches the grid. In
OKC that meant the ±10-year vintage window kept 30 addresses — but almost all of
them 1.5–2.8 miles away in **Midtown / Heritage Hills**, at $1,200–$1,525 for a
1BR, while excluding every property within 0.7 miles of the subject. The
resulting "market rent" would have been indefensible.

`references/dwellsy-api.md` is explicit that this trim is ours: *"The API returns
what is nearby, not what is comparable."* Use
**`scripts/trim_comps_by_distance.py`** (in this folder) to cut the flat CSV to
the subject's own submarket before the exporter runs; it preserves the 24-column
schema so everything downstream is unmodified. Then pass `--scope-note` so the
radius is printed on the deliverable rather than being a silent decision.

```bash
python trim_comps_by_distance.py --csv "out/pull1.csv" --max-miles 1.0 \
    --subject-lat 35.49646 --subject-lon -97.54705 --out out/comps_submarket.csv
```

Sanity check before committing to a radius: the 1BR property-weighted median was
$770 at 1.0 mi, $784 at 1.5 mi and $850 at 3.0 mi. If that spread is large, the
pull is crossing a submarket boundary and the wide number is the wrong one.

## 2. On a pre-1960 subject the ±10-year vintage window inverts its own purpose

The vintage rule exists for **product parity**. On a 1953 subject it means
1943–1963 — and in an urban core that band is dominated by renovated historic
conversions, while the subject's actual competitors are the 1964–1970 walk-ups
across the street. Here, inside 1 mile, ±10 years left exactly **two** comps;
±20 years left nine, including three of the five lease comps the seller's own
broker used.

For a 1950s–60s walk-up, widen to `--vintage-window 20` and say so on the
deliverable. Both 1953 and 1969 are unrenovated pre-1975 walk-up product; the
distinction the rule is protecting against (1953 vs 2020) is still enforced.
Verify vintages **before** trimming — a verified year moved 1731 NW 32nd St from
Dwellsy's 1977 to an actual 1964, i.e. from outside the window to inside.

## 3. Extrapolating a whole-bed $/SF median onto a small floor plan understates it

Rent does not scale linearly with unit size — small units carry a much higher
$/SF. On this comp set the 1BR $/SF ran **$1.99–$2.01 at 450 SF** and **$1.06–
$1.18 at 651–876 SF**. The exporter's positioning page computed
`psf_med × plan_sf` off the whole-bed median (comp median SF 625) and applied it
to a 525 SF plan, producing a "comp-supported rent" of $651 — below the
subject's own in-place rent — and a false $0 upside.

Two patches to `rent_comp_export.py`, both backward-compatible. They belong
upstream in the plugin; until they land there, re-apply them in the job-local
copy (as `rent-comps-without-dwellsy-8-2026.md` also instructs for its own
`--source-label` / `--methodology` patches, which are still not upstream — check
whether you need both sets).

**Patch A — size-matched PSF-implied rent.** Add next to `bed_benchmarks()`:

```python
def psf_implied(rows, bed, sf, variance, min_props=2):
    """Property-weighted median $/SF of comps within +/-`variance` of `sf`,
    times `sf`. Returns None when fewer than `min_props` properties match."""
    lo, hi = sf * (1 - variance), sf * (1 + variance)
    by_prop = {}
    for r in rows:
        if (r["beds"] != bed or not r["sane"] or not r["sf"] or not r["rent"]
                or not lo <= r["sf"] <= hi):
            continue
        by_prop.setdefault(r["prop_key"], []).append(r["rent"] / r["sf"])
    if len(by_prop) < min_props:
        return None
    return med([med(v) for v in by_prop.values()]) * sf
```

then in the positioning table replace
`implied = v["psf_med"] * p["sf"] ...` with a `psf_implied(rows, b, p["sf"],
sf_variance)` call falling back to the old expression when it returns `None`,
and thread `sf_variance=args.sf_variance` through `write_pdf`. This makes page 4
agree with the "Granular Rent Comparison" tab, which already uses that band.

**Patch B — stop hardcoding the vintage window in the footnotes.** The XLSX note
and the PDF grid note both said *"vintage within 10 years of the subject"*
regardless of `--vintage-window`, so a ±20 run shipped a false statement. Add
module-level `VINTAGE_WIN = 10` / `SCOPE_NOTE = ""`, set them from the CLI at the
top of `main()`, interpolate `VINTAGE_WIN` into both notes, and add a
`--scope-note` flag that appends to both — that is where the submarket radius
from §1 gets disclosed.

## 4. Verify identities before trimming — the merge is worth more than the filter

Step 4 is not paperwork on a dense pull. Here it removed **1,452 phantom units**:
Campus Pointe was arriving as 11 separate "132-unit" properties and Copper Ridge
as 3 separate "37-unit" ones. It also caught Dwellsy reporting **33 units** at
1712 N Blackwelder Ave (actually **2**), **14 units** at University Pointe
(actually **127**), and a single-family house at 1604 NW 30th St carrying the
whole neighbouring community's 132-unit count.

The Oklahoma County Assessor is the authoritative source and is reachable only
through the r.jina.ai text proxy (direct requests are WAF-blocked):

```
https://r.jina.ai/https://docs.oklahomacounty.org/AssessorWP5/AN-R.asp?PROPERTYID=<id>
https://r.jina.ai/https://docs.oklahomacounty.org/AssessorWP5/BLDG_Detail.asp?PropertyID=<id>&BuildingSequence=1
```

Two behaviours worth knowing: on a multi-building parcel the **whole community's
`# of Res Units` sits on `BuildingSequence=1` and every other building shows 0 —
do not sum them**; and the address index carries only one address per parcel, so
a multi-address community appears under a single "Location". Confirm the other
addresses by searching each one and finding it returns *no parcel of its own*.

## 5. Verified OKC comp identities (Classen / Gatewood / Plaza District, 8/2026)

Reusable on any deal in this submarket. Units/vintage are assessor-verified.

| Property | Address(es) | Units | Built |
|---|---|---|---|
| Charleston Apartments | 2021 NW 25th St | 16 | 1969 |
| Elite at 25 Apartments | 2030 NW 25th St | 19 | 1951 |
| Flamingo Apartments | 1844 NW 23rd St | 32 | 1961 (gut-renovated 2019) |
| Campus Pointe Apartments | 1600/1602/1604/1606/1608/1610 NW 31st + 1601/1603/1607/1609/1611 NW 30th | 130 | 1969 |
| University Pointe Apartments | 1515 NW 30th St | 127 | 1968 |
| Copper Ridge Apartments | 1429 NW 24th / 1430 NW 25th / 1433 NW 24th | 36 | 1968 |
| Sienna Ridge Apartments | 1428 NW 27th St | 36 | 1970 |
| The Vic | 2000 + 2017 N Blackwelder Ave | 24 | 1964 |
| The Plaza Apartments | 1740 + 1744 NW 17th St | 16 | 1948 |
| 32nd Street Apartments | 1731 NW 32nd St | 17 | 1964 |
| The Heights OKC | 3720 N Pennsylvania Ave | 38 | 2006 |
| Kamp's Court Apartments | 1400/1402/1404 NW 25th St | 22 | 1928 |

Excluded, and why: **1712 N Blackwelder Ave** = 2 units (Dwellsy said 33);
**1604 NW 30th St** = 1920 single-family house (Dwellsy gave it Campus Pointe's
132).

## 6. Don't let the portal feed be your subject rent

Four portals (Zillow / RentCafe / ApartmentFinder / Trulia) published identical
subject asking rents — they are **one syndicated feed with four echoes**, not
four confirmations, and the feed disagreed with the manager's own leasing page by
$55–$75 a unit and contradicted itself on availability the same day. Take the
subject's rents from the rent roll or the OM unit-mix page; use portals only to
sanity-check.

## Related

- `rent-comparable-analysis/SKILL.md` — workflow, house rules, pull budget
- `rent-comps-without-dwellsy-8-2026.md` — same property, no-API path; the OM
  retrieval route and the `--source-label` / `--methodology` patches
- `scripts/trim_comps_by_distance.py` — the submarket trim from §1
- `sales-comps-from-an-address-only-8-2026.md` — the sale-comp side
