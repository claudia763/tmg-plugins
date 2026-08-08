# Building a submarket-anchored (promotional) sale comp grid — 8/8/2026

Covers: how to rebuild a sale comp grid on the subject's IMMEDIATE submarket when
a broker says "comps in the immediate area traded high, be promotional," what
happens when you widen the radius to chase a fifth comp, and two more defects in
the pipeline scripts. Read after
`sales-comps-from-an-address-only-8-2026.md` and
`sales-comps-pipeline-hardening-8-2026.md`. Worked on **The Vintage Apartments,
2037 NW 26th St, Oklahoma City** — 66 units, 1953, 529 SF average.

## 1. The setup: two legitimate grids, an order of magnitude apart in usefulness

A first pass on this deal used the standard metro configuration —
`--max-distance 15`, `hard_unit_range (0.4, 10.0)` — and produced a deed-verified
grid at **$45,551/unit ($3,006,357)**. That run also flagged, honestly, that the
trades within a mile of the subject had cleared at **$62,000–$85,600/unit** and
that this was "the argument for pushing above the grid with the right buyer."

The broker then asked for exactly that argument. Rebuilding on the immediate
Plaza District / Classen corridor gave **$60,497/unit ($3,992,820)** — a 33%
higher indication off the same universe, the same scripts and the same deed
verification.

**Both grids are defensible. They answer different questions.** The metro grid
asks "what does 1940–1980 OKC product trade for?"; the submarket grid asks "what
does product in THIS corridor trade for?" When a broker points at the corridor,
the second question is the right one — but say plainly in the deliverable which
question you answered.

## 2. The submarket grid was also the technically SOUNDER grid

This is the part worth remembering, because it turns a promotional request into a
defensible one rather than a strained one. Compare the two runs' diagnostics:

| Diagnostic | Metro grid | Submarket grid |
|---|---|---|
| Raw $/unit SD | 15.9% of mean | 16.7% |
| **Adjusted $/unit SD** | **23.7% (WIDENED)** | **16.3% (compressed)** |
| Divergence residual | +0.9% | **−0.1%** |
| Comp avg unit SF vs subject 529 | 789 SF | **675 SF** |
| Comp avg vintage vs subject 1953 | newer set | **1962.2** |

`sales-comps-from-an-address-only-8-2026.md` §6 says dispersion should COMPRESS
under adjustment and that a widening is a sign the comp set or the adjustments
are wrong. On the metro grid it widened; on the submarket grid it compressed.
The reason is unit size: the metro set averaged 789 SF against a 529 SF subject,
so the size adjustment did violent and uneven work. The corridor comps average
675 SF, so it does less.

**Check the dispersion direction before you decide which grid to lead with.** It
is a better tiebreaker than distance.

## 3. Widening the radius to reach five comps ACTIVELY destroyed the grid

At 2.5 miles only five comps scored and the ±1σ trim cut two, leaving three. The
instinct is to widen. Do not widen without looking at what arrives:

| Radius | Shortlist mean | Shortlist SD | What got added |
|---|---|---|---|
| 2.5 mi | $78,074 | $21,844 (28%) | — |
| 3.0 mi | $44,628 | $37,016 (**83%**) | Port 50 @ $19,699/u, two "Multi-Property Sale" rows @ $13,407 and $10,511/u |
| 3.5 / 4.0 mi | identical to 3.0 | identical | nothing further |

Beyond 2.5 miles this submarket contains no additional comparable product — only
distressed and unnamed rows. Worse, the collapsing mean moved the ±1σ band and
**silently trimmed the best comp (Mesta Park, $85,611) instead of the worst**, so
widening did not just add noise, it removed signal.

The right lever was the documented `--outlier-sd` flag: on a five-comp shortlist a
1.0σ trim removes 40% of an already-tiny sample. Loosening to **`--outlier-sd 1.5`**
kept a fourth comp and still excluded the genuine vintage outlier. Disclose the
setting in the deliverable.

## 4. Two more pipeline defects (both reached a client-facing artifact)

### 4a. Placeholder property names — `select_comps.py`

The universe carries unnamed rows literally called **"Multi-Property Sale"** —
four clustered on one date at $5,591–$13,407/unit. They are allocation artifacts
with no identifiable asset behind them, they **do not carry the "(Part of a …
Portfolio)" marker** the address screen from the hardening note looks for, and
they cannot be deed-verified because there is no property to look up. Two reached
the top-10 shortlist at 3.0 miles.

```python
"exclude_placeholder_names": ["multi-property sale", "multi property sale",
                              "portfolio sale", "unknown", "n/a"],
```

```python
if any(name_l.strip() == k or name_l.strip().startswith(k)
       for k in cfg.get("exclude_placeholder_names", [])):
    skipped["placeholder_name"] = skipped.get("placeholder_name", 0) + 1
    continue
```

### 4b. `verify_exports.py` hardcodes a five-comp grid

Three separate checks assume exactly five comps:

```python
top5 = [ws.cell(r, 1).value for r in range(5, 10)]      # -> [..., ..., ..., ..., None]
check("grid's 5 comps at the top of the client export", ...)
check("PDF shows all 5 comps", all(f"Comp {i}" in text for i in range(1, 6)))
check("underwriting grid-5 first ...", [...range(2, 7)] == top5 ...)
```

A thin submarket, or an outlier trim on a small shortlist, legitimately yields
three or four comps — and the harness then **fails a correct deliverable while
reading like a real defect**, which is worse than not checking at all. Take the
count from the selection:

```python
n_grid = len(sel_comps_ordered)
top_n = [ws.cell(r, 1).value for r in range(5, 5 + n_grid)]
check(f"grid's {n_grid} comps at the top of the client export",
      top_n == [c["Property Name"] for c in sel_comps_ordered], str(top_n))
if n_grid < 5:
    print(f"  [note] grid holds {n_grid} comps, not the usual 5 -- confirm "
          f"this is a genuinely thin comp set and say so in the deliverable")
```

and thread `n_grid` / `top_n` through the PDF and underwriting checks too (there
is a third reference to `top5` further down that will `NameError` if you miss it).
Both patches belong upstream; until they land, re-apply in the job-local copy.

## 5. Framing a promotional range honestly

The broker set the floor ("leave the 3.0m figure as a bottom range"). The comps
set the ceiling. What makes the range credible is that **each end has its own
independent support**, and that you say which:

- **Floor $3,000,000 ($45,455/unit).** Four independent anchors land here — the
  metro adjusted grid ($3,006,357), Year-1 NOI at 6.25% ($2,897,984), T-3 NOI at
  the 6.29% state average cap ($2,873,959), and Yardi's own submarket 5-year
  average of $44,000/unit × 66 = $2,904,000. Also the seller's 2021 basis
  ($2,900,000).
- **Ceiling ≈ $4,000,000 ($60,606/unit).** The deed-verified corridor grid
  ($3,992,820).

Then disclose the two things that constrain the top, rather than waiting for a
buyer to find them:

1. **Yield.** At the ceiling the implied cap is ~4.5% on T-3 income where the
   state average is 6.29%. The top of the range is a **basis-and-repositioning
   price, not a yield price**, and it needs an equity buyer — debt capacity on
   this asset was only ~$2.15M.
2. **Per-foot.** At $3,950,000 the subject prices at **$114.91/SF against a comp
   mean of $102.81/SF** — above every comp but one. At ~$3.6M it sits exactly at
   the comp mean, which is why that is the honest "most likely clearing" zone.

Being promotional means selecting the defensible high end and arguing it well. It
does not mean suppressing the arithmetic that a buyer will run in the first
meeting.

## 6. Deal record

Submarket grid, 2.5 mi / 3-yr / `hard_unit_range (0.24, 3.0)` / `--outlier-sd 1.5`:

| Comp | Mi | Units | Built | Sold | $/Unit | Adjusted |
|---|---|---|---|---|---|---|
| Briargate & Plaza | 0.75 | 32 | 1948 | 6/2024 | $71,094 | $65,345 |
| Charleston Apartments | 0.06 | 16 | 1969 | 9/2024 | $62,000 | $54,960 |
| Manchester on May | 2.44 | 100 | 1960 | 3/2024 | $55,000 | $47,928 |
| Mesta Park | 1.44 | 27 | 1972 | 5/2024 | $85,611 | $73,756 |

Indicated **$60,497/unit → $3,992,820**. Pennsylvania Avenue Apartments
($116,667/unit, 2006 vintage) was excluded by the pipeline's own ±1.5σ rule —
which is the correct outcome, since 2006 product is 53 years newer than the
subject and a buyer would knock it out immediately.

`hard_unit_range` was widened DOWN to 0.24 (16–198 units at a 66-unit subject) so
the 16-unit Charleston next door qualifies. The 8-, 6- and 5-unit infill trades in
the same corridor still fall out, which is correct — at that size they price like
houses ($162,500/unit at 2700 Walker Ave, $163,333/unit at 3922 N Classen).

## 7. What deed verification did to this grid — and a third hardcoded-5 defect

Verification landed AFTER the grid was built and staged, and it changed the
deliverable. Ship nothing promotional before it comes back.

**Mesta Park -> "Perle at Mesta Park", and the database unit count was wrong.**
Assessor R045007974 / PID 151697 records **28** residential units, not 27,
corroborated by MLS, Zillow and ApartmentGuide. The sale itself reconciles
exactly (Book 15754/446, 5/3/2024, $2,311,500) — but $2,311,500 / 28 =
**$82,554/unit, not the database's $85,611**, which rested on a 27-unit
denominator. The top comp was 3.6% overstated before any argument began. This is
§6b's derived-number tell firing on a bad denominator rather than a bad price.

Two further facts killed it as a graded comp: the assessor carries **Remodel Year
2015** (the subject is unrenovated), and 700-710 NW 17th St sits in the **Mesta
Park National Register historic district, zip 73103 — Midtown, 1.43 mi east
across Classen**. That is the same 1.4-2.8 mi Midtown/Heritage Hills band this
deal's RENT comp run had already rejected as indefensible and trimmed away at 1.0
mile. **Using a submarket on the sale side that you excluded on the rent side is
the inconsistency a buyer finds first, and it is in your own file.** Check your
prior runs on the same asset for geographic decisions before setting a radius.

Also resolved: a worry that "700 NW 17th St, 24u" and "710 NW 17th St, 24u" in
the rent dataset meant two properties totalling ~48 units. `700 NW 17th St`
returns **no parcel of its own** — it is one 28-unit asset double-counted across
its two street numbers, the same phantom-duplication pattern as Campus Pointe.

**Charleston's sale verified exactly** (Book 15881/1507, $992,000, 9/27/2024,
16 units, 1969) — but note two **$0 entity transfers flanking it** that a scraper
would book as phantom trades, and that at 16 units it fails the standard
`hard_unit_range`. Keeping it required widening that screen, which must be
disclosed: a buyer holding the methodology page who sees a documented screen
relaxed for the one comp that lifts the number will read it as reverse
engineering.

**Pennsylvania Avenue: exclusion confirmed correct**, and for the record it is
**38 units / $110,526per unit**, not 36 / $116,667.

### A third hardcoded-5: `select_comps.py` refuses to build a small grid

```python
if len(comps) < cfg["final_count"]:
    sys.exit("ERROR: only N usable comps after filtering — check the geocode")
```

Excluding one comp from a five-comp market tripped this and killed the run.
`final_count` is a **ceiling on the graded grid, not a minimum the market must
supply** — a submarket averaging 0.2 sales a year cannot produce five. Add a
`min_grid_comps` floor (3), refuse below it, and otherwise proceed with a printed
NOTE so the thinness gets disclosed rather than silently shipped. Together with
the `verify_exports.py` fix in §4b, that is three separate places assuming five
comps; grep for the literal 5 before running any thin-market deal.

### Final numbers

| Grid | Comps | Indicated $/unit | Total |
|---|---|---|---|
| Corridor (all unrenovated, all deed-verified) | 3 | $56,254 | **$3,712,735** |
| Corridor + Perle at Mesta Park (corrected, disclosed) | 4 | $60,029 | $3,961,921 |

Both were delivered. The corridor grid leads because it is the one defensible
under questioning; the four-comp version is the alternate for a broker who wants
the Midtown print in the story with its caveats attached.

## Related

- `sales-comps-from-an-address-only-8-2026.md` — base playbook
- `sales-comps-pipeline-hardening-8-2026.md` — dedupe, portfolio screen, corrections layer
- `rent-comps-the-vintage-okc-8-2026.md` — the assessor access recipe used to verify
- `vintage-okc-uw-writeup-8-2026.md` — the income side these prices sit against
