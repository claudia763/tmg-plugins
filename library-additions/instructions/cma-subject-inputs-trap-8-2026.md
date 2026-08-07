# The CMA `Inputs!B8` subject-size trap — it silently inflates every sale comp grid (8/7/2026)

Covers: why the "Subject Indicated Value" on the Automatic CMA Analysis /
Sale Comparables grid can come out several points too high even when every
comparable is correctly selected and every sale price is right. Read this
before quoting a grid indication in a BOV, an underwriting writeup, or any
pricing recommendation. Found on Aldine Apartments (96 units, Houston 77039,
8/7/2026); the mechanism is generic and applies to every deal.

## The trap

`Automatic CMA Analysis.xlsx` → sheet `Inputs` has four yellow subject cells:

| Cell | Field |
|---|---|
| `B7` | Year Built |
| `B8` | **Avg. Unit Size**  ← this one |
| `B9` | Unit Count |

`Inputs!B8` gets filled from whatever third-party data sheet was at hand —
Yardi, CoStar, a broker flyer. **Those sources routinely publish a gross or
"average unit size" that is not net rentable SF ÷ unit count.** On Aldine the
data sheets said **815 SF**; the rent roll says **69,600 net rentable SF over
96 units = 725 SF**. A 90 SF, 12.4% overstatement.

Nothing in the workbook validates it, and nothing about the output looks
wrong. The grid still renders, every comp still footnotes correctly, and the
indicated value is simply too high.

## Why it moves the answer

The grid's size adjustment (delivered workbook `Comparable Grid`!E24, one row
up in the CMA's own copy) is:

```
adj_size = -(comp_avg_SF / subject_avg_SF - 1) / 4
```

Overstate the subject's size and every comp looks relatively *smaller*, so
every comp is adjusted *upward* toward the subject. The error does not
average out across comps — it pushes all of them the same direction.

Aldine, five selected comps (Scottwood, Garner Park, Long Point Plaza, Jade
Forest, Winkler):

| Subject avg size | Adjusted indication | × 96 units |
|---|---|---|
| 815 SF (as built, wrong) | $73,307 /unit | $7,037,468 |
| **725 SF (correct)** | **$71,072 /unit** | **$6,822,886** |
| — unadjusted simple mean of the five | $71,763 /unit | $6,889,248 |

**$214,582, or 3.1%, of pure input error** — on a deal whose whole pricing
argument turned on a 7% discount to the grid.

## The tell, and it is a good one

If the **adjusted** indication and the **unadjusted simple mean** of the
comps' actual $/unit diverge by more than about 1–2%, suspect `B8` before you
suspect the adjustments. Correctly specified, the two should sit close
together — a well-chosen comp set does not need large net adjustments, and a
large net adjustment in one direction is evidence of a bad subject input, not
of a subtle market insight.

On Aldine, correcting `B8` collapsed a 2.2% gap ($73,307 vs $71,763) to 1.0%
($71,072 vs $71,763). That convergence is what a healthy grid looks like.

## What to do

1. **Compute the subject average yourself** from the rent roll: total net
   rentable SF ÷ unit count. Never accept a data sheet's figure. Cross-check
   against the unit mix (Aldine: 48 × 650 + 48 × 800 = 69,600).
2. Write it into `Inputs!B8` **before** building the grid. The delivered
   `<Property> - Sale Comparables.xlsx` reads the subject size through an
   external link (`='[1]Inputs'!B8`), so this one cell drives both workbooks —
   fix it in the CMA and the comparables workbook follows.
3. Re-check `B7` (Year Built) and `B9` (Unit Count) the same way while you are
   there. `B7` feeds `adj_yb = (comp_yb - subject_yb) / -100 / 2`, so a
   five-year vintage error is worth 2.5% on every comp.
4. When you report a grid indication, **say which basis it is** — adjusted or
   unadjusted mean. If both are quoted anywhere across the deliverables, they
   will be compared in diligence, and an unexplained gap reads as sloppiness.

## Reproducing the grid outside Excel

The delivered comparables workbook pulls its whole data block from
`'[1]Output Analysis Data'` (the CMA), so on Linux — where the external link
will not resolve — `data_only=True` returns `None` for the entire grid, and
the CMA's *own* cached `Comparable Grid` values are usually **stale
prior-deal leftovers** (Aldine's shipped with five Dallas comps from an
unrelated subject). Do not read either one and believe it.

Recompute in Python instead; the whole grid is four lines:

```python
adj_yb   = (comp_yb - subj_yb) / -100 / 2
adj_size = -(comp_sf / subj_sf - 1) / 4
adj_cap  = -(drift_bps / 10000) / current_avg_cap      # Comparable Grid!AI1
adjusted = ppu * (1 + adj_yb + adj_size + adj_cap)     # ppu = sale price / units
indication_per_unit = mean(adjusted for the selected comps)
```

Take the comps and the selected ranks from `selection.json` in the job's
salescomps working directory, and `current_avg_cap` from the CMA's
`Comparable Grid`!AI1. This reproduction tied to the penny against the
workbook's own adjusted output on Aldine, which is what let the 815 SF input
be isolated as the sole cause.

## Related

- `sales-comps/SKILL.md` (main library) — comp selection and the grid itself.
- The underwriting skill's closing rule: *"Flag anything you could not refresh
  (sale-comp pages, agency datasets) in the summary to the user — never
  present stale prior-deal comps silently."* The stale cached CMA grid
  described above is exactly that hazard.
