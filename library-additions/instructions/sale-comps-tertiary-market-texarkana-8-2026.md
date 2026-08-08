# Sale comps in a TERTIARY market — when the hardened screens delete your comp set (8/8/2026)

Covers: running `sales-comps` on a small-market deal where the whole universe
inside 75 miles is **five sales**; why the metro-calibrated screens from the
earlier notes actively destroy such a grid; how to catch a rent-restricted or
age-restricted comp that no name screen can see; using TMG's **own Fannie/Freddie
workbook as a primary source to correct a sale price**; and a self-correcting
one-page fix for the PDF. Read **after**
`sales-comps-from-an-address-only-8-2026.md`,
`sales-comps-pipeline-hardening-8-2026.md` and
`submarket-anchored-promotional-sale-comps-8-2026.md` — this extends them.
Worked on **Renaissance Square, 2401 County Ave, Texarkana, AR 71854** —
65 units, 760 SF avg, geocode 33.445681 / −94.038257 (US Census, first try).

## 1. Check the market's DEPTH before you pick any setting

Do this before touching CONFIG. One pass over the universe, printing every comp
inside ~75 miles with units / vintage / date / $per unit, tells you which of the
three prior playbooks you are in:

| What the scan shows | Playbook |
|---|---|
| 100+ comps inside 15 mi | metro — `--max-distance 15`, `hard_unit_range (0.4,10)`, 1.0 SD |
| ~10–40 inside 15 mi | submarket — the promotional note's levers |
| **< 10 inside 75 mi** | **this note** |

Texarkana returned **15 comps inside 75 miles, 5 of them usable**. The next
nearest cluster was **Dallas–Fort Worth at 150–175 miles**, and because the
distance bracket flattens to a constant 25 points beyond 3 miles, a 165-mile
Uptown Dallas high-rise at $250k/unit outscores a 3-mile Texarkana walk-up.
**A distance cap is not optional in a tertiary market — without it the grid is a
DFW grid.** 75 miles was the smallest radius that reached a 5-comp Ark-La-Tex set
(Marshall, Pittsburg, Shreveport, Magnolia + Texarkana).

## 2. The metro screens delete a tertiary comp set. Check what each one ATE.

Every screen the prior notes recommend turning ON fired here, and between them
they left a Texarkana grid with **zero Texarkana comps**:

| Screen | What it removed | Right call? |
|---|---|---|
| `exclude_portfolio_allocations` | **BOTH** large Texarkana sales — Westridge and Park at Summerhill are the two legs of one local 2-property trade and both carry the "(Part of a 2 Property Portfolio)" marker | **No** — fold them into one comp instead (§4) |
| `hard_unit_range (0.4, 10.0)` | Northgate Square (16u, **0.64 mi** — the closest sale in the market) | Yes, but say so |
| `outlier_sd 1.0` then `1.5` | at 1.5 SD it trimmed **Oaks on Victory** — 60 units vs the subject's 65 and 1975 vintage, i.e. the single best size-and-vintage match in the set | **No** — see §3 |

The prior notes' rule ("if a screen moves the grid it is doing more than tidying
and you owe the broker both runs") understates it here. In a 5-comp market a
screen does not tidy the supporting list, **it is the grid**. Print what each
screen removed and eyeball it every time. To make that possible, have
`select_comps.py` record examples, not just counts:

```python
dropped_examples = {"portfolio_allocation": [], "name_keyword": [],
                    "placeholder_name": [], "unit_range": []}
# ...in each screen branch, alongside the counter:
if d < 25 and len(dropped_examples["unit_range"]) < 12:
    dropped_examples["unit_range"].append(
        f"{r.get('Property Name')} ({int(units)}u, {d:.1f} mi, ${round(price/units):,}/u)")
```
and stash it in `selection.json` under `universe.screen_examples`.

## 3. On a 5-comp shortlist, ANY sigma trim is too aggressive

`submarket-anchored…` §3 says a 1.0 SD trim removes 40% of a five-comp shortlist.
Here it removed 40% **and took the best comp with it**, because one genuinely
non-comparable row (§5) was still inflating the mean.

Sequence actually observed, subject at 65 units:

| Setting | Grid | Note |
|---|---|---|
| 1.0 SD | 3 comps, none in Texarkana | unusable |
| 1.5 SD | 4 comps — trimmed Oaks on Victory ($42,500) | **trimmed the best match** |
| **2.0 SD** | **5 comps, nothing trimmed** | shipped |

**Fix the comp set first, then set sigma.** Once the restricted-rent comp (§5) was
removed the shortlist SD fell from $9,579 to $7,869 and 2.0 SD kept all five
honestly, rather than 1.5 SD keeping four by luck. Disclose the setting on the
deliverable — a relaxed screen that happens to raise the number reads as reverse
engineering unless you name it first.

## 4. The Fannie/Freddie workbook is a PRIMARY SOURCE for prices, not just drift

`SKILL.md` restricts `Combined Fannie and Freddie Sales Comps.xlsx` to cap-rate
drift, and that rule is right for *selecting* comps. But it is also a record of
real originations with `Value at contribution`, `NOI at contribution`, `LTV` and
`Cap Rate`, keyed to property and date — **and it will catch a bad sale price in
the sale universe.** Search it by city/ZIP before you ship.

At Texarkana it produced the run's biggest finding:

| Source | Westridge (176u) | Park at Summerhill (184u) | Combined |
|---|---|---|---|
| Sale universe (CoStar / Fannie rows) | $8,399,453 · $47,724/u | $12,553,750 · $68,227/u | $20,953,203 · $58,203/u |
| **Fannie originations 3/14/2024** (IDs 1720011939 / 1720011940, both LTV 0.80) | **$13,833,325 · $78,598/u** | $12,553,750 · $68,227/u | $26,387,075 · $73,297/u |
| **Trade press (TheRealDeal, MAHB), Feb-2024** | — | — | **$23,000,000 · $63,889/u** |

Three tells fired at once:

1. **Summerhill's universe row IS the Fannie row** — its "$12,553,750" is a
   `Value at contribution` (an appraised value), not a sale price. `Info Source =
   Fannie` on a sale row means the price may never have been a price.
2. **Sequential loan IDs, same date, same LTV** — conclusive that two "separate"
   comps are one transaction.
3. The universe's Westridge price is **39% below** the agency valuation of the
   same asset in the same month.

Resolution: **one transaction is one comp.** Via `apply_comp_corrections.py`,
fold the pair into a single row at the reported $23.0M / 360 units / $63,889 per
unit, unit-weight the avg unit SF ((869×176 + 755×184)/360 = 810 SF), and
`exclude` the other leg. Do **not** invent a per-property allocation — Texas is a
non-disclosure state and none was ever published.

**Ordering trap in `apply_comp_corrections.py`:** it matches each row on
`property_name`, so a `set` of `Property Name` invalidates every later row for
that comp (they log `WARNING: no row matched`, and the run silently applies only
some corrections). **Put the rename LAST in the CSV.** Nine corrections applied
cleanly once reordered; before that, only two did.

## 5. A rent-restricted comp that no name screen can catch — check the state HFA list

**Arbor Pointe Apartments**, 600 N Oats St, 1.9 miles from the subject, 48 units,
**2006** vintage — and the **cheapest** comp in the set at **$38,542/unit**. The
newest property trading at the lowest per-unit is economically backwards, and
that inversion is the tell.

Arkansas Development Finance Authority's own county list carries it verbatim:
`Arbor Pointe Apartments / 600 N Oats St / Texarkana, AR 71854 / Elderly / 1, 2`.
It is **ADFA-financed affordable AND age-restricted**. Neither fact is in the
comps universe, and `exclude_name_keywords` (`"senior"`, `"assisted living"`)
cannot see it — the name is generic.

Dropping it moved the indication from **$47,935 to $52,933/unit (+10.4%)** and,
more tellingly, took adjusted dispersion from **24.0% (widening — the bad sign)
to ~14.5% (flat)** and the divergence residual from +1.1% to **~0.0%**.

Independent verification later confirmed all of it: Arbor Pointe is a **9% LIHTC
property with 47 of 48 units income-restricted**, **55+ age-restricted**, three
storeys with an elevator and emergency call pull-cords, waitlisted through the
Housing Authority of the City of Texarkana, and now branded *Dwell at the Arbor*.
$38,542/unit is exactly what a Year-15 restricted senior asset trades at.

**Check the state HFA list for every comp within a few miles, and for any comp
whose $/unit is inverted against its vintage.** The lists are free PDFs, one per
county, and the URL pattern is guessable:

```
https://adfa.arkansas.gov/wp-content/uploads/2024/12/<County>-County-<City>.pdf
   e.g. .../Miller-County-Texarkana.pdf, .../Columbia-County-Magnolia.pdf
```

`WebFetch` cannot read them (binary PDF), but it **saves the file locally** and
prints the path — parse that with PyMuPDF:

```python
import fitz
for p in fitz.open(saved_path):
    print(p.get_text())
```

Texas equivalent: TDHCA. Same check cleared Fox Creek Magnolia (Columbia County
list holds only Ridge at Magnolia and Preston Apartments), so it stayed in.

## 6. `export_comps.py`: make the one-page PDF self-correcting

A corrected/thin-market deal carries **more** disclosure text than a normal one —
here four footnotes about a folded portfolio, an excluded restricted comp, an
estimated vintage and a relaxed sigma. The notes block grows, and `KeepTogether`
then bounces the whole map onto page 2, failing `verify_exports.py`'s single-page
check. Estimating the remaining space by summing `flowable.wrap()` heights
**underestimates it** and still spilled.

Do not delete disclosures to fit the page. Build, count pages, shrink, repeat:

```python
import pypdfium2 as _pdfium
cap_h = None
for _ in range(12):
    build_pdf(..., map_max_h=cap_h)          # caps Image height, keeps aspect
    if len(_pdfium.PdfDocument(pdf)) == 1:
        break
    cap_h = (cap_h or 230) - 20
    if cap_h < 90:
        print("WARNING: could not fit the map — shorten the extra notes")
        break
```

with, inside `build_pdf`, `Image(map_png, width=mw, height=mh, hAlign="CENTER")`
after `if map_max_h and mh > map_max_h: mh, mw = map_max_h, map_max_h * w / h`.
Texarkana settled at 150pt on the second rebuild.

Two more `export_comps.py` items worth carrying upstream:

- **An `--extra-note` (append) flag.** Disclosures belong on the client-facing
  artifact, not only in a Notes doc. **Use `&#8226;` bullets, not `&#8308;`–`&#8311;`
  superscripts — Carlito has ¹²³ but not ⁴⁵⁶⁷, and reportlab renders the missing
  glyphs as black squares.** (Confirmed by rendering; invisible to the verifier.)
- **Strip `(Part of a N Property Portfolio)` from the PDF's Address row** and
  disclose it in a footnote. In a fixed `colWidth` cell that suffix consumes the
  whole address and ellipsizes it into noise.

The §5 autofit patch and its ellipsis fallback from the two prior notes both
remain necessary and both fired here. **Render to PNG and look at it every run** —
`scripts/pdf_to_png.py` (contributed with this note) does it via pypdfium2, since
this host has no `pdftoppm`.

## 7. When the subject's YEAR BUILT does not exist anywhere

It is in neither the rent roll nor the T-12; actDataScout, ARCountyData, LoopNet,
Apartments.com, Crexi and Redfin were all 403/CAPTCHA to automated fetching,
including through the `r.jina.ai` proxy. The only published figure — **2000**,
from the Apartments.com/CoStar record — arrives attached to a **54-unit** count
against a 65-unit rent roll, i.e. the record is demonstrably wrong on the fact we
*can* check, and 550 SF 1BR / 900 SF 2BR walk-ups spread over eight street
addresses is not 2000-vintage product. (The asset also shows a former name,
*Beechdale Apartments*, at the same address.)

Do not silently pick a number. **Estimate it, label the estimate on the
deliverable itself, and ship the sensitivity** — vintage is worth about
**$278/unit per year** on a grid like this ($0.5%/yr × mean $/unit):

| Subject year built | Indicated $/unit | Total (65 units) |
|---|---|---|
| 1960 | $50,155 | $3,260,085 |
| **1970 (used)** | **$52,933** | **$3,440,616** |
| 1980 | $55,710 | $3,621,147 |
| 2000 (the published figure) | $61,265 | $3,982,210 |

A 40-year spread is a **22% swing** — bigger than any screen argued about above.
The `cma-subject-inputs-trap` note warns that `B7` is worth 2.5% per five years;
in a market with no assessor access that is the single largest uncertainty in the
deliverable, and it deserves the top line of the caveats, not a footnote.

**Arkansas is deed-verifiable and this is resolvable in a browser.** Ark. Code
§ 26-60-105/107/110: transfer tax of **$3.30 per $1,000**, documentary stamps
affixed to the face of the deed, recorder may not record without them — so
**price = stamp ÷ 0.0033**. Miller County Circuit Clerk (870-774-4501) records
deeds; Fidlar Tapestry has remote access; actDataScout (assessor-sponsored) should
carry year built and transfer history for APN 3770030 and the adjoining Beech St
parcels. *Note the asymmetry the earlier notes flagged for Oklahoma applies here
too: **Arkansas and Louisiana record consideration; Texas does not.** Five of this
grid's comps span three states — do not write one disclosure line for all of them.*

## 8. Deal record — Renaissance Square, 8/8/2026

65 units, 760 SF avg (49,400 NRSF ÷ 65 — computed from the rent roll per
`cma-subject-inputs-trap`), vintage estimated 1970. Universe 3,990 → corrected to
3,988 → **5 scored** (75 mi, 3-yr, 26–650 units, screens on, portfolio folded).
Cap-rate drift scope fell back to **state only** (Arkansas, 5.86% current, n=13);
no state+vintage sample exists.

| # | Comp | Mi | Units | Built | Sold | $/Unit | Adjusted | Price verified? |
|---|---|---|---|---|---|---|---|---|
| 1 | Oaks on Victory (Marshall TX) | 64.7 | 60 | 1975 | 10/31/2025 | $42,500 | $39,396 | No — TX non-disclosure; CAD value within 6% |
| 2 | Pecan Estates (Pittsburg TX) | 62.6 | 40 | 1982 | 9/18/2024 | $57,500 | $54,296 | No — figure is a LoopNet **ask**; CAD value within 1.4% |
| 3 | Riverside Oaks (Shreveport LA) | 71.0 | 185 | **1972** | 7/18/2024 | $51,351 | $51,623 | **Yes** — multiple trade reports |
| 4 | Westridge / Summerhill Portfolio | 3.3 | 360 | 1984 | 2/27/2024 | $63,889 | $62,726 | Reported (5 outlets), not recorded |
| 5 | Fox Creek Magnolia (Magnolia AR) | 47.8 | 48 | 2003 | 7/1/2024 | $62,500 | $55,651 | **Yes** — recorded AR deed |

Indicated **$52,739/unit → $3,428,008**. Raw SD 14.2% → adjusted 14.4%;
divergence residual **−0.0%**.

Verification corrections that moved the grid, both from primary records:
**Riverside Oaks year built 1975 → 1972**, and **Pecan Estates avg unit SF
655 → 747** (Camp CAD property 13616: 29,862 SF of apartment buildings ÷ 40; the
database's 655 was one fourplex segment, 2,619 ÷ 4). Together they moved the
indication $52,933 → $52,739.

Three findings worth carrying into any deal in this region:

- **A "sale price" in a non-disclosure state may be an ASKING price.** Pecan
  Estates' $2,300,000 is the LoopNet ask to the dollar, and Camp CAD confirms only
  that a General Warranty Deed changed hands that day. The county appraised value
  ($2,268,861) is what makes it usable — cross-check every unverifiable price
  against the CAD/assessor value and say which you did.
- **Watch for one buyer appearing repeatedly.** Reynolds Asset Management bought
  BOTH Comp 4 and Comp 3, both brokered by John Hamilton at Marcus & Millichap —
  so two of five grid comps reflect a single firm's underwriting rather than
  independent market clearing. Disclose it.
- **Fox Creek Magnolia is a student-housing repositioning** ("Rider Ridge",
  0.4 mi from Southern Arkansas University; rents moved $575 → $875–975). Its
  prior sale was $1,895,000 in May 2021 — **+58% in three years**, which is a
  useful trend datapoint but not a stabilized workforce comp.

**The 5.86% Arkansas statewide drift cap is not Texarkana's cap rate.** The
agency workbook's own Texarkana originations run far above it — Westridge 7.77%,
Park at Summerhill 7.67%, Patriot Apartments 6.77% (all 2024); regionally,
Arkansas 70s–80s vintage 2024+ medians 6.42% and TMG's East Texas region 6.40%.
Drift is only a normalizer between sale dates so the grid is unaffected, but
**never quote 5.86% as this market's cap rate**, and hand the real numbers to the
underwriting run.

## Related

- `sales-comps-from-an-address-only-8-2026.md` — base playbook, PDF autofit, §6b verify-the-comps
- `sales-comps-pipeline-hardening-8-2026.md` — dedupe, portfolio screen, corrections layer
- `submarket-anchored-promotional-sale-comps-8-2026.md` — thin-shortlist sigma, hardcoded-5 defects
- `cma-subject-inputs-trap-8-2026.md` — why `B7`/`B8` drive everything above
- `scripts/apply_comp_corrections.py` · `scripts/comp_grid_crosscheck.py` · `scripts/pdf_to_png.py`
