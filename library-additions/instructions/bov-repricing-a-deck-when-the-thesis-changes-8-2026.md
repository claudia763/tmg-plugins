# Repricing a BOV deck when the underwriting thesis changes (8/8/2026)

Covers: updating an already-delivered 11-page BOV after the sale comps and the
underwriting were both rebuilt and the recommendation moved from **$2,900,000 to
$3,700,000 guidance**. The mechanical part (swap the numbers) is easy. The parts that
bite are **stale claims that contradict the new numbers but contain no numbers**, and
**a model artifact that reads as a fact**. Read after `bov-deck-layout-and-render-notes.md`
(the fixed-height overflow trap, which still governs everything). Worked on **The
Vintage Apartments, 2037 NW 26th St, Oklahoma City** — 66 units, 1953.

## 1. Three-number decks beat one-number decks when the range is wide

The v1 deck led with a single "Recommended Valuation: $2,900,000". The promotional
rebuild has a genuinely wide defensible band, and a single number could not carry it.
What worked was making **three numbers structural** — on the summary page, the
recommendation page and the back cover:

| Slot | Value | What it is |
|---|---|---|
| **Guidance** | $3,700,000 | what we launch at |
| **Marketing range** | $3,000,000 – $3,950,000 | what we quote |
| **Underwriting strike** | $3,050,000 | the highest price at which every return and lender test stays green |

The strike is what makes the floor credible: *"The floor is not a concession — the
metro-wide adjusted grid supports $3,006,357 and the income underwriting strike is
$3,050,000. The bottom of the range has support under it, not a hole."* A promotional
deck that cannot explain its own floor invites the buyer to test below it.

Corollary: when the comps and the income disagree by 22% ($3,712,735 vs $3,050,000),
**give each end its own page real estate** rather than averaging them. One "Why the
floor is $3.0M" box and one "What constrains the top" box does more than a paragraph
reconciling them.

## 2. Diff the CLAIMS, not just the figures — search for prose without numbers

The v1 financing page argued *"the price of this asset is set by the debt, not by the
real estate"* and *"leverage is negative at every income basis."* Both were correct on
v1's numbers. The rebuild moved Year-1 NOI from $181,124 to $249,400 and **inverted
both**:

| | v1 | Rebuild |
|---|---|---|
| Yield on cost | 5.55% | **8.18%** |
| Loan constant | 7.9836% | 7.9836% |
| Leverage | negative 243 bps | **positive ~20 bps** |
| T-3 DSCR | 0.93x at 75% LTV | **1.491x at 60% LTV** |

A find-and-replace on dollar figures leaves these sentences untouched and the deck
contradicts itself. **Grep for directional words, not just numbers**: `negative`,
`below`, `above`, `constrained`, `cannot`, `no positive spread`, `discount to`. Every
hit is a claim that has to be re-derived.

Same class of defect found on the rental page: a bullet claiming the smallest unit
segment showed **+2.1%** rent growth, against the authoritative figure of **−0.4%** for
one-bedrooms. It survived v1 because nothing cross-checked it.

## 3. A model artifact is not a fact — check before you print it

The deck's distress test printed *"12.26% vacancy plus **0% concessions** plus 2.19%
bad debt."* That 0% comes straight from the model and is **not true of the property**.

Concessions ran **$5,600** across the T-12 and **$800 / $800 / $1,400** in the final
three months — they are accelerating, not absent. The 0% exists only because the T-12
population folded concessions into the bad-debt row:

```python
SERIES["bd"] = [b + c for b, c in zip(BAD, CONC)]   # Final_T_12 row 8 = 10,569
                                                    #   = 5,600 conc + 4,969 bad debt
```

So the 2.19% "bad debt" already carries the concessions. Printing "0% concessions" in a
client deck states something the operating statement contradicts, and a buyer holding
both documents finds it immediately.

Two fixes, both needed:
- Restate the line as **"12.26% vacancy plus 2.19% of concessions and bad debt
  combined"** — accurate, and the 14.45% total is unchanged.
- **Restore the concessions diligence item** to the risk list: *"Concessions are still
  running and sit inside the bad-debt line. April 2026 alone carried $1,400 against zero
  reported on the rent roll — reconcile the two and prepare a leasing and collections
  summary."*

**Generalisable rule: any zero that arrives from a model, rather than from the source
document, is suspect.** Trace it before it reaches a client page. A subagent asked to
remove contradictions will "resolve" this the wrong way — deleting the true bullet
because it conflicts with the false zero. Give it the folding rule up front.

## 4. Risk lists are at capacity — swap, do not append

The Recommendation page holds **five bullets at 20.5 px** and no more (see the layout
note). Adding the concessions item meant removing one. The right one to cut is the
bullet that **duplicates content already on another page** — here, "the bridge
benchmarks only payroll and repairs & maintenance," which the Deal Optimization page
already makes at length. A live retrade lever outranks a restatement.

Never shrink the font to fit one more bullet; below ~20 px it stops reading as the OM's
type, and the overflow is silently clipped anyway.

## 5. When an excluded comp is worth showing

Mesta Park was excluded from the primary grid (1.43 mi east across Classen, 73103
historic district, remodeled 2015 against an unrenovated subject — the exact distance
band already rejected on the rent comps). It still earned a **labelled right-column
callout** giving the four-comp indication of $3,961,921 / $60,029 per unit **with the
exclusion reasons stated**.

That is better than silence in a promotional deck: it shows the work, it supports the
posture, and it pre-empts a buyer who finds the comp and assumes it was hidden. **The
rule is the label travels with the number, always.**

Separately: the broker flagged that Mesta Park is **relisted at $3,050,000** — an
unaccepted ask, not a trade — and that it must stay out of all materials. Verify by
grep before shipping, not by memory:

```
grep -ci 'relist\|unaccepted' vintage_bov.html      # must be 0
```

## 6. Two more things worth carrying forward

**Give a requested scenario its own page if the numbers need one.** Ownership asked for
a furnished-housing analysis. Compressed into a corner of the Deal Optimization page it
would have been illegible or clipped; as its own page it holds a 9-row economics ladder,
a two-comp evidence table, the regulatory position and a "Not Capitalised Into Guidance"
box. Deck went 11 → 12 pages. **Renumber the page chips.**

**Label a per-unit denominator whenever two exist.** The model computes on the rent
roll's **65** units; marketing per-unit pricing is quoted on the assessor's **66**.
Replacement reserves are struck at $350 × 65 = $22,750. Left silent this looks like an
arithmetic error; disclosed in a footnote and a risk bullet it looks like control of the
file.

## 7. Deal record

Delivered `The_Vintage_Apartments_BOV.pdf`, **12 pages**, 1700 × 1080 px, rendered from
`work4/vintage_bov.html` via Playwright/Chromium (`render.cjs`). v1 preserved at
`vintage_bov_v1_backup.html`.

Guidance **$3,700,000** ($56,061/unit, $107.64/SF) · range **$3.0M – $3.95M** · likely
clearing $3.3M – $3.6M · underwriting strike **$3,050,000** (IRR 22.02%, avg CoC 10.08%,
T-3 DSCR 1.491x, 60% LTV at 7.00%, terminal cap 6.00%) · Year-1 NOI **$249,400** at an
**8.18%** yield on cost · corridor grid **$3,712,735** ($56,254/unit).

**Every page was rendered to PNG and read.** That is the only check that catches the
fixed-height clipping, and on a repriced deck it is also how you catch a stale claim
sitting next to a fresh number.

## Related

- `bov-deck-layout-and-render-notes.md` — the vertical budget and collision rules
- `writeup-off-a-model-verify-dont-transcribe-8-2026.md` — the same verify-don't-transcribe habit on the narrative
- `vintage-okc-underwriting-8-2026.md` · `submarket-anchored-promotional-sale-comps-8-2026.md`
