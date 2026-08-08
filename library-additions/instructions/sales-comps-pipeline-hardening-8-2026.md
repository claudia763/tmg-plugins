# Sale-comp pipeline hardening — three universe defects the scripts did not catch (8/7/2026)

Covers: three provenance/quality defects in the `All Sale Comps` universe that
reach a CLIENT-FACING sheet with every automated check green, the pipeline
changes that stop each one, a reproducible way to apply primary-source
corrections without hand-editing a grid, and one more `export_comps.py` layout
failure mode. Read **after**
`sales-comps-from-an-address-only-8-2026.md` — this extends it rather than
replacing it. Worked on **The Vintage Apartments, 2037 NW 26th St, Oklahoma
City, OK 73106** — 66 units, 1953, 529 SF avg.

## 1. The same sale, ingested twice, both rows on the client sheet

The universe is appended from several feeds, so one transaction can appear under
two source rows. **Kentucky Pines** arrived as both `5904 S Harvey Ave` (CoStar,
807 SF) and `5804 South Harvey Avenue` (Yardi, 802 SF) — same 60 units, same
2024-02-21, same $3,100,000 — and **both** reached the client-facing supporting
list. `verify_exports.py` passes: nothing it checks is about duplicates.
(`5804 S Harvey` returns no assessor record; it is a listing-site typo.)

Dedupe on the **transaction**, not the label, right before `comps.sort(...)` in
`select_comps.py`. Key on `(Sale Date, round(Sold Price), Unit Count)`; when two
rows collide keep the one whose stated `Sold Price/Unit` reconciles to
price ÷ units, else the first:

```python
def _dupe_key(c):
    return (c["Sale Date"], round(c["Sold Price"]), c["Unit Count"])

def _reconciles(c):
    return abs(c["Sold Price/Unit"] - c["Sold Price"] / c["Unit Count"]) < 1.0
```

Log what was dropped and record it in `selection.json`. Note the boundary
effect: removing one duplicate took the surviving pool from 21 to 20, and the
exporter's "cap at 20, trim to 15 when over" rule then kept 20 instead of 15.
Expect the supporting-list length to jump when you turn this on.

## 2. Portfolio allocations are labelled in plain text — screen on it

`sales-comps-from-an-address-only` §6b describes portfolio allocations as
something you catch by web-verifying. On this universe you can catch most of
them for free: **298 of 3,990 rows carry the allocation in the address string**,
e.g. `428 N Willowood Dr (Part of a 5 Property Portfolio)`. Three such rows
(all sharing sale date 2024-04-09 — one 5-property trade) reached the client
list, alongside a senior-housing property and, in the wider pool, a **mobile
home park**.

Add two CONFIG screens and apply them in the scoring loop:

```python
"exclude_portfolio_allocations": True,
"exclude_name_keywords": ["mobile home", "senior", "assisted living",
                          "student housing", "nursing"],
```

```python
if cfg.get("exclude_portfolio_allocations") and re.search(
        r"part of a.*portfolio", addr_l):
    skipped["portfolio_allocation"] = skipped.get("portfolio_allocation", 0) + 1
    continue
```

On this deal that removed 4 allocations + 1 senior property and the **grid 5 and
the indicated value were bit-identical** — the screens only cleaned the
supporting list. That is the outcome to look for; if a screen moves the grid, it
is doing more than tidying and you owe the broker both runs.

## 3. Correcting a verified defect without breaking reproducibility

§6b tells you to web-verify all five and expect half to be defective, then says
to ship an "As-Run" and a "Screened" grid. But when verification turns up a
**corrected value** rather than a reason to drop a comp, hand-editing the
exported grid destroys reproducibility, which `SKILL.md` explicitly forbids.

Use **`scripts/apply_comp_corrections.py`** (in this folder). It applies a
reviewed CSV of primary-source findings to the universe workbook and writes a
corrected copy, logging every change and refusing any row without a `source`.
Re-running the unchanged pipeline on the corrected workbook gives the second
grid, and the CSV is the audit trail:

```
property_name,sale_date_match,action,field,new_value,source
Manchester on May,2024-03-15,set,Sold Price,5500000,"Assessor deed PID 124261 …"
Highland Oaks,2025-10-02,exclude,,,"3BR townhomes, not garden walk-up …"
```

It deliberately does **not** auto-recompute `Sold Price/Unit` — the stated column
and price ÷ units genuinely disagree on ~31% of rows, so which is right has to
be a human call. Correct both fields explicitly.

## 4. Oklahoma is a DISCLOSURE state — do not reuse the Texas caveat

`sales-comps-from-an-address-only` says "Texas is a non-disclosure state, so
'price not publicly reported' is the expected answer." **That is a Texas fact,
not a general one, and copying it into an Oklahoma deliverable would be wrong.**

Oklahoma levies a documentary stamp tax of $0.75 per $500 of consideration
(68 O.S. §3201) and requires proof of purchase price with the deed, so the
county clerk records the consideration and the assessor publishes it. On this
run **5 of 5 sale prices were verified against primary records** — and two
disagreed with the database:

| Comp | Database | Recorded deed |
|---|---|---|
| Manchester on May | $5,200,000 | **$5,500,000** (5.8% understated) |
| Mason Manor | $3,170,000, 11/02/2023 | **$3,165,000, 10/30/2023** |

The Mason Manor case inverts the §6b tell. That note says an exactly-dividing
`$/unit` marks a *derived* number; here $3,165,000 ÷ 69 = $45,870 matched the
database's stated per-unit exactly, proving the per-unit was computed off the
**true** price and the *price field* was later corrupted. The non-dividing
per-unit was the trustworthy figure. Treat the tell as flagging derivation, not
error.

Also seen and worth checking for: a **$0-consideration entity transfer**
(Legacy Investment Holdings → Mason Manor Holdings, 12/30/2024) that any scraper
would book as a phantom sale.

## 5. The PDF autofit fix needs an ellipsis fallback

§5's shrink-to-fit patch floors the font at 5.0pt. Corrected labels are longer
than database ones (`1740 NW 17th St / 1712-1718 N Indiana Ave`), and at that
length **5pt still overflows and reportlab goes back to overprinting**. Extend
the patch so anything still too wide at the floor is truncated with an ellipsis:

```python
if pdfmetrics.stringWidth(val, f0, sz) > avail:
    while val and pdfmetrics.stringWidth(val + "…", f0, sz) > avail:
        val = val[:-1]
    val = (val.rstrip() + "…") if val else ""
    grid_rows[r_i][c_i] = val
```

Better still, keep grid labels short by design and put aliases (`fka Crescent
Park`, the four-address list for a multi-address parcel) in the Notes document
rather than a fixed-width cell. **Render the PDF to PNG and look at it on every
run** — this failure is invisible to `verify_exports.py`, which only checks that
the comp names appear in the text layer.

## 6. Deal record — The Vintage Apartments, OKC, 8/7/2026

66 units, 1953, 529 SF avg (implied — the OM prints no SF column),
35.496450 / −97.546994 (Nominatim; the US Census geocoder is WAF-blocked from
this host, and three independent sources agreed within 20 ft). Universe 3,990 →
42 scored (15 mi, 3-yr, 26–660 units, screens on). `hard_unit_range` (0.4, 10.0)
and `--max-distance 15` were both needed: without the unit range an 8-unit and a
16-unit infill sale scored into the top 5; without the distance cap a 77-mile
and a 103-mile comp did. Shortlist SD fell from 70.2% of mean (default) to 41.6%
with both on.

| Grid | Indicated $/unit | Total |
|---|---|---|
| As-Run (database as-is) | $43,597 | $2,877,431 |
| Verified (primary-source corrected) | **$45,551** | **$3,006,357** |

Verified grid: Briargate & Plaza (0.75 mi, 32u, 1948, $71,094), Manchester on
May (2.44, 100u, 1960, $55,000), Mason Manor (4.30, 69u, 1972, $45,870),
Kentucky Pines (6.35, 60u, 1971, $51,667), Yorktown (3.22, 92u, 1968, $50,000).
Cap-rate drift scope state-only, current avg cap **6.29%**; no comp drew a
non-zero drift. Subject's own trade 7/23/2021 at $2,900,000 ($43,939/unit) is
outside the 3-year window, so `subject_self_sales` came back empty — check the
assessor for it rather than trusting that emptiness.

**Adjusted dispersion widened on both grids** (15.9% raw → 23.7% adjusted on the
Verified grid), which §6 says is a bad sign. Here it is explained rather than
disqualifying: the subject's 529 SF units are far smaller than any available
comp (789 SF comp average even after screening), so the size adjustment does
most of the work and varies widely across comps. The divergence residual was
**+0.9%**, i.e. the subject inputs are right — judge on the residual, not the raw
divergence. Quote the range, not the point.

## 7. OKC-specific notes

- The **Oklahoma County Assessor** is the primary source and is reachable only
  through the r.jina.ai text proxy (direct requests are WAF-blocked). Its deed
  history with dollar considerations is the single best check on a stated price
  and date. Full access recipe and its two gotchas are in
  `rent-comps-the-vintage-okc-8-2026.md` §4.
- Several OKC comps sit in **separately incorporated municipalities** that
  listing databases label "Oklahoma City" — Highland Oaks is in **Warr Acres**,
  and Bethany, Del City, Midwest City, Yukon and Edmond all appear in a 15-mile
  OKC pull. Check situs before writing a city on a client sheet.
- `$/unit` is essentially **flat across property size** in OKC 1940–1980 product
  (medians $48.8k–$54.4k from the 16–35 unit bucket through 151+). So a nearby
  small property trading at a premium is a **location** signal, not a size
  effect — do not "adjust" it away.

## Related

- `sales-comps-from-an-address-only-8-2026.md` — the base playbook; read first
- `sales-comps/SKILL.md`, `references/scoring.md`
- `rent-comps-the-vintage-okc-8-2026.md` — same subject, rent side; carries the
  assessor access recipe and the verified OKC comp identities
- `scripts/apply_comp_corrections.py` — the corrections layer from §3

## 8. Follow-up: submarket-anchored grids and two more script defects

`submarket-anchored-promotional-sale-comps-8-2026.md` (8/8/2026) extends this file:

- **Placeholder property names** slip past the §2 portfolio screen. The universe
  carries unnamed rows literally called "Multi-Property Sale" at $5,591-$13,407/unit
  with no "(Part of a ... Portfolio)" marker in the address, so the regex misses
  them. Add an `exclude_placeholder_names` screen.
- **`verify_exports.py` hardcodes a five-comp grid** in three checks and fails a
  correct three- or four-comp deliverable, which reads like a real defect. Take the
  count from `selection.json`.
- **Widening the radius to reach a fifth comp can remove signal, not just add noise.**
  Going 2.5 -> 3.0 mi on this deal collapsed the shortlist mean, blew dispersion to
  83%, and moved the +/-1 sigma band so it trimmed the BEST comp instead of the worst.
  Use `--outlier-sd` on a small shortlist rather than reaching for distance.
- **Check the dispersion DIRECTION when choosing between two defensible grids.** The
  tighter-radius grid compressed dispersion under adjustment (16.7% -> 16.3%) where
  the metro grid widened it (15.9% -> 23.7%) - a better tiebreaker than distance.
