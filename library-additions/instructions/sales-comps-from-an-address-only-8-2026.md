# Running sale comps when the broker sends only an address (8/7/2026)

Covers: the end-to-end `sales-comps` run for a job whose entire input is a street
address and no attachments — where the source workbooks actually live, the two
selection guards you almost always need in an infill/urban market, a client-facing
PDF layout bug and its fix, and the cross-check to run before quoting the number.
Read alongside `sales-comps/SKILL.md`, `references/scoring.md`, and
`cma-subject-inputs-trap-8-2026.md`. Worked on **The Montrose (fka Vue Fitzhugh),
2819 N Fitzhugh Ave, Dallas TX 75204** — 226 units, 2004, 807 SF avg.

## 1. There is no `All Sales Comps.xlsx` — do not ask for it

`SKILL.md` says "Inputs (user attaches; ask if missing)". **Never ask.** Both inputs
are already in TMG Dropbox, and the broker has no standalone copy of the first one:

| Skill input | Where it really is |
|---|---|
| `All Sales Comps.xlsx` | the **`All Sale Comps` tab** (singular "Sale") of `Databases/- Sales Comps/Automatic CMA Analysis.xlsx` |
| `Combined Fannie and Freddie Sales Comps.xlsx` | `Databases/- Sales Comps/` — use as-is |

Extract the tab with `scripts/extract_all_sale_comps.py` (drops the junk `Column1`,
prints the row count and sale-date span so you can confirm the database is current).
As of 8/7/2026 it holds **3,990 comps spanning 2021-06-01 → 2026-06-01**.

Mount Dropbox first — see `rclone-dropbox-mount.md`.

## 2. Verify the subject yourself; the internal note may be stale

`SKILL.md`'s own worked example calls 2819 N Fitzhugh "Vue Fitzhugh" with lat/lon
32.8112 / -96.7768. Both were wrong at run time: the asset is now **The Montrose**,
and the US Census geocode is **32.815087 / -96.787260** (~0.7 mi off the example's).
A stale name on a client-facing grid is embarrassing; a stale geocode silently
re-ranks every comp. Re-geocode and re-confirm the name every time, and put the
former name in parentheses on the deliverable — `The Montrose (fka Vue Fitzhugh)` —
so the broker recognises the deal.

Derive avg unit SF yourself (`cma-subject-inputs-trap-8-2026.md`). At The Montrose
three defensible-looking numbers existed: **807 SF** (182,367 net rentable / 226 —
correct), 824.6 SF (DCAD *gross* building SF / 226, which Redfin and LoopNet both
publish as "building SF"), and 1,455 SF (DCAD *total improvements* / 226 — this adds
the property's separate 142,525 SF parking garage). The garage trap is the dangerous
one: it is a separate DCAD improvement record on the same parcel, and summing
improvements looks perfectly reasonable until the grid explodes.

## 3. Turn on `hard_unit_range` for any mid/large urban subject

`CONFIG["hard_unit_range"]` ships as `None`, so unit count contributes points only.
The smallest bracket is `|Δunits|/units < 1.00 → 10 pts`, which any property smaller
than the subject clears automatically. On a 226-unit subject the default run put a
**6-unit** ($627,000/unit) and a **19-unit 1983 walk-up** into the top 5 purely on
distance — both a third of a mile away, both nothing like the asset.

Set `"hard_unit_range": (0.4, 10.0)` in the CONFIG block of your job-local copy of
`select_comps.py` (the value `scoring.md` documents). At 226 units that admits
90–2,260 units. Effect at The Montrose: shortlist SD fell from $138,873 to $92,669
and all five comps became 149–386-unit institutional product.

It is a CONFIG constant, not a CLI flag — edit the constant, per `SKILL.md`'s
"edit constants there rather than hand-picking comps; keep the pipeline reproducible."
**Say in the reply that you turned it on**, and what it excluded.

## 4. Use `--max-distance` to bound the SUPPORTING list, not the grid

The distance brackets are 100/50/25 pts for <1 / <3 / anything beyond. Beyond 3 miles
distance stops discriminating at all, so a **90-mile Waco comp outscored a 2.3-mile
Dallas comp** (recent sale + near-identical unit count beat 25 points of distance).
The grid 5 were unaffected — all were inside 3 miles — but the client-facing sheet
carries up to 15 comps, and Waco/Northlake/Allen rows on an Uptown Dallas package
contradict the PDF's own footnote ("comparable market area/linkages").

`scoring.md` explicitly forbids rebalancing the point brackets to fix this
("don't 'improve' this by rebalancing without asking"). The sanctioned lever is the
documented CLI flag: **`--max-distance 15`** for a metro-bounded run.

Re-run selection with it and diff the grid before shipping. At The Montrose the grid
5, the outlier trim, and the indicated value were **bit-identical** — the cap only
cleaned up the supporting list (worst comp went 90.2 mi → 13.2 mi, and four
closely-comparable Dallas comps at 797–856 SF surfaced in place of suburban filler).
That is the outcome to look for. If a distance cap *does* move the grid, it is doing
more than tidying and you owe the broker an explanation of both runs.

## 5. PDF layout bug: long names and addresses overprint each other

`export_comps.py`'s `build_pdf()` passes plain strings to a reportlab `Table` with
fixed `colWidths` and fixed `rowHeights`. **Reportlab neither wraps nor clips a plain
string — it overprints the neighbouring cell.** At The Montrose "Lennox West Village"
ran into "MAA Cathedral Arts", and two addresses collided so badly that "6044 East
Lovers Lane" rendered as "3044 East Lovers Lane". The verifier passes: it checks the
comp names are *present* in the PDF text layer, which they are. **You only catch this
by rendering the PDF to PNG and looking at it.** Do that on every run.

Fix (job-local copy, just before `t.setStyle(TableStyle(style))`): resolve each cell's
effective font from the FONT entries already in `style`, then append a per-cell
override that shrinks the size until the string fits `colWidth - 12` (reportlab's
default 6pt left + 6pt right padding), floored at 5.0pt. Shrinking beats wrapping
because `rowHeights` is fixed — a wrapped second line would be clipped.

```python
PAD = 6 + 6
def _eff_font(col, row):
    f, sz = font, 7.8
    for e in style:
        if e[0] != "FONT":
            continue
        (c0, r0), (c1, r1) = e[1], e[2]
        c1 = len(grid_rows[0]) - 1 if c1 == -1 else c1
        r1 = len(grid_rows) - 1 if r1 == -1 else r1
        if c0 <= col <= c1 and r0 <= row <= r1:
            f, sz = e[3], e[4]
    return f, sz

for r_i, row in enumerate(grid_rows):
    for c_i, val in enumerate(row):
        if not isinstance(val, str) or not val.strip() or val == SP:
            continue
        avail = widths[c_i] - PAD
        f, sz = _eff_font(c_i, r_i)
        while sz > 5.0 and pdfmetrics.stringWidth(val, f, sz) > avail:
            sz -= 0.1
        if abs(sz - _eff_font(c_i, r_i)[1]) > 0.01:
            style.append(("FONT", (c_i, r_i), (c_i, r_i), f, round(sz, 1)))
```

This belongs upstream in the plugin's `export_comps.py`; until it lands there, re-apply
it in each job's copy. Everything else in the PDF was correct — the CartoDB basemap
rendered fine on this box (`pip install reportlab matplotlib contextily
--break-system-packages`; none are preinstalled and the PDF step dies without them,
*after* successfully writing both .xlsx files, so a "partial success" run is normal).

## 6. Cross-check before quoting — `scripts/comp_grid_crosscheck.py`

Run it on `selection.json` and put its numbers in the reply. It prints the grid
against three bases (adjusted / unadjusted mean $/unit / mean $/SF × subject SF),
the dispersion before and after adjustment, the trap note's divergence test, and a
subject-avg-SF sensitivity table.

The refinement it adds to the trap note: that note's "diverge by more than 1–2% →
suspect `B8`" tell **assumes a tight comp set**. The Montrose diverged **−14.4%** with
a completely correct subject input, because the comps genuinely averaged 2013.8
vintage and 1,089 SF against a 2004 / 807 SF subject. So compare the observed
divergence to the divergence the vintage and size gaps *predict* (−13.3% here);
judge on the **residual** (−1.1%), not on the raw divergence. The script does this
and prints FLAG / OK.

Two other things worth quoting to a broker, both of which this script gives you:

- **Dispersion should fall.** Raw $/unit SD 19.2% of mean → adjusted 13.3%. If the
  adjustments *widen* dispersion, the comp set or the adjustments are wrong.
- **The grid should sit between the $/SF floor and the $/unit ceiling** when the
  subject's units are smaller than the comps'. The Montrose: $51.2M ($/SF basis) <
  **$57.4M (grid)** < $67.0M (unadjusted $/unit). Give the broker the range, not just
  the point — a five-comp grid on a heterogeneous set is a range, and quoting a single
  $57,388,608 to the dollar overstates the precision.

## 6b. VERIFY THE COMPS THEMSELVES — the database is not clean

This is the most important section here, and the one I would have skipped. The
pipeline verifies *arithmetic*, not *provenance*. Every automated check passed on a
grid that turned out to rest on three bad comps. **Web-verify all five selected comps
before delivering**, and expect roughly half to have a problem.

What a check of five comps turned up on one ordinary Dallas deal:

| Comp | Defect found |
|---|---|
| Griffis Uptown | $83.64M is an **allocated slice of a 4-property, 1,421-unit off-market portfolio**. Only the Seattle leg's price was ever published. Also traded as *Uptown Trail* — "Griffis Uptown" is the buyer's post-close rebrand, a tell that the row was scraped from a current-name database rather than the deal record. DCAD carries it at $90M. |
| Eastline Residences | $116M was **never disclosed**. The acquisition financing was $87M Truist senior + $28M Macquarie pref = **$115M** — a capital stack one million below the "price," implying 99% leverage. DCAD says $140.6M. Almost certainly the stack captured as a price. |
| Perry Row | Year built is **2009, not 2010** (DCAD + five outlets). It is a **3-story townhouse community**, plans 997–2,213 SF with attached garages — the subject's *average* unit was smaller than this comp's *smallest* plan. Price never disclosed; only a $42M Mesa West acquisition-**and-renovation** loan. Pre-renovation value-add basis. |
| Lennox West Village | Price never disclosed. TRD cited DCAD at ~$39M against the database's $36.67M. |
| MAA Cathedral Arts | The **only** verified price — $106M, confirmed in MAA's own Q3-24 release and FY-24 10-K Schedule III. But **bought during lease-up**, so the price is on pro-forma NOI; do not derive a going-in cap from it. Also mislabeled "Uptown" — it is Ross & Henderson, Old East Dallas. |

**The structural tell:** every unverified $/unit is an *exact* arithmetic derivative of
its price ($375,839 × 149 = $56,000,000 exactly). A per-unit figure that divides evenly
is a derived number, not an observed one.

**Texas is a non-disclosure state**, so "price not publicly reported" is the expected
answer for most comps and is fine to say. What is *not* fine is presenting a portfolio
allocation, a financing stack, or a townhouse community as a straight $/unit comp
without saying so.

### Three data-quality defects in the universe itself

1. **31% of rows have a stated `Sold Price/Unit` that disagrees with price ÷ units**
   (1,229 of 3,974). Some are catastrophic (Mercer Park: $119,355 computed vs $385,016
   stated). **The pipeline's outlier trim reads the stated column while the client grid
   recomputes**, so the two can disagree inside one deliverable — The Montrose shipped
   $230,608 on the Comp Export tab and $230,629 in the grid and PDF for the same comp.
2. **27% of rows have `Building SF` ÷ units ≠ `Avg Unit SF`** — `Building SF` is often
   gross. MAA Cathedral Arts: 384,000 / 386 = 995 SF against a stated 837 SF. This
   drives the size adjustment, so it moves the answer (−0.9% vs −5.8% for that comp).
3. **14% of the "All Sale Comps" universe is Fannie/Freddie-sourced** (560 of 3,990) —
   agency loan originations, which the skill's own rule says to use *only* for cap-rate
   drift, never as sale comps. They are nonetheless mixed into the sale universe, and
   one (The Griffin, `Info Source = Fannie`) scored into a grid. An agency record's
   "price" may be an appraised or loan-basis value. Filter on `Info Source` if the
   grid depends on one.

### Deliver a range, and consider a screened grid

Removing the three defective comps and re-running moved The Montrose from
**$253,932/unit ($57.4M)** to **$215,383/unit ($48.7M)** — a **15% swing** with every
automated check green in both runs. Additionally excluding agency-sourced records gave
$205,500/unit ($46.4M) but with adjusted dispersion blowing out from 14.6% to 21.5%,
i.e. a *worse* grid — which is why dispersion, not just the headline, decides whether a
screen improved anything.

Ship both grids (label them "As-Run" and "Screened"), quote a range, and let the broker
choose. A single point estimate off this data is false precision.

## 7. Deal record — The Montrose, 8/7/2026

226 units, 2004, 807 SF avg, 32.815087 / −96.787260. Universe 3,990 → 138 scored
(15 mi, 3-yr, 90–2,260 units). Shortlist mean $310,159/unit, SD $92,669; trimmed as
>1 SD outliers: NOVEL Turtle Creek ($475,962), Uptown By Onni ($424,528), Trend
Design District ($194,175), Griffin, The ($191,176).

| # | Comp | Mi | Units | Built | Sold | $/Unit | Adjusted |
|---|---|---|---|---|---|---|---|
| 1 | Lennox West Village | 0.75 | 159 | 2001 | 5/31/2024 | $230,629 | $212,001 |
| 2 | MAA Cathedral Arts | 0.97 | 386 | 2024 | 10/9/2024 | $274,611 | $244,598 |
| 3 | Griffis Uptown | 1.56 | 334 | 2013 | 12/22/2025 | $250,419 | $229,376 |
| 4 | Perry Row | 2.86 | 149 | 2010 | 12/1/2025 | $375,839 | $278,056 |
| 5 | Eastline Residences | 1.84 | 330 | 2021 | 2/19/2026 | $351,515 | $305,629 |

Indicated **$253,932/unit → $57,388,608**. Cap-rate drift scope state+vintage,
current avg cap **5.68%**; only Lennox drew a non-zero drift (−10 bps → +1.8%).
No subject self-sale in the universe: the property last traded **Nov 2015** (EGW
Fitzhugh Investment LP) at an undisclosed price — Texas is a non-disclosure state —
so it is outside the 3-year window and there is no in-house trade to anchor to.
DCAD 2026 certified market value $36,750,000 ($162,611/unit) is an assessment, not a
market indication; label it as such if it comes up.

## 8. Follow-up: three more universe defects, and a state-law caveat

`sales-comps-pipeline-hardening-8-2026.md` (8/7/2026) extends this file with
three defects §6b does not cover — the same sale ingested twice from two feeds
and BOTH reaching the client sheet, portfolio allocations that are labelled in
plain text in the address field (298 of 3,990 rows) and can be screened
automatically, and a reproducible way to apply verified corrections
(`scripts/apply_comp_corrections.py`) instead of hand-editing a grid. It also
adds an ellipsis fallback to the §5 PDF patch, which still overprints at the
5pt floor once labels get longer.

**Important correction to §6b:** "Texas is a non-disclosure state" is a Texas
fact, not a general one. Oklahoma records the consideration (doc-stamp tax,
68 O.S. §3201), and on the OKC run 5 of 5 prices were verified against deeds —
two of which disagreed with the database. Check the subject state's rule before
writing "price not publicly reported" on a deliverable.
