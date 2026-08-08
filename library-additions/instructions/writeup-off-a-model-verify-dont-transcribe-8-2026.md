# Writing the underwriting writeup: read the MODEL, never the model-builder's notes (8/8/2026)

Covers: updating an existing Broker Valuation Summary when the underwriting has been
rebuilt, and the single habit that matters most — **every figure in the narrative must
be read back out of the delivered workbook, not copied from the handoff notes.** On
this deal that habit caught four errors and one real model defect, all of which would
otherwise have reached a client document. Worked on **The Vintage Apartments, 2037 NW
26th St, Oklahoma City** — 66 units, 1953. Read alongside
`vintage-okc-underwriting-8-2026.md` (the model build) and
`vintage-okc-uw-writeup-8-2026.md` (the original narrative).

## 1. The handoff notes were wrong in four places. The workbook was right.

Part 1 delivered the model and wrote a detailed handoff. Part 2 rebuilt the writeup
from it. Re-reading each figure out of the workbook before printing it found:

| Claim in the handoff | What the workbook actually holds |
|---|---|
| Printed sales range `$2,840,000 / $3,050,000 / $3,350,000` | **`$2,800,000 / $3,050,000 / $3,300,000`** — the handoff divided by 66; the model prints on the rent roll's **65** |
| "Value-add: operations items ON (water fixtures, RUBS, billback, pet fees, parking)" | **`Value-Add` column C has ZERO ticks.** Grand total $0; Year-1 other income identical to T-12 actual |
| Payroll benchmarked at `$1,400/unit` | **`$1,300/unit`** ($84,500 ÷ 65) — the model's own note says so |
| T-12 NOI `$160,094` | **`$160,096`** ($507,870 − $347,774) |

The value-add one is the dangerous one. Phase A's `tmg valuation.py` had the items set
to `True`; SKILL step 8 says to transcribe those into the workbook's `Value-Add`
column-C flags, and that step **was never completed**. Nothing failed. The model
recalculated clean, all three return tests went green, and the strike solved — just on
a base case with no value-add at all. **A skipped transcription step is invisible to
every automated gate.** Check the tab, don't trust the config:

```python
va = wb['Value-Add']
ticked = [va.cell(r,2).value for r in range(1,80)
          if str(va.cell(r,3).value).lower().strip() == 'x']
assert ticked, "Value-Add flags never transcribed from Phase A"
```

The recovery is honest rather than expensive: the delivered strike is a **base case**,
and the ops-only programme becomes a *stated, uncapitalised* upside ("run in the model
it lifts the supportable price to $3,120,000 — real, but a buyer's return, not a
seller's keep"). That is a better document than one quietly claiming the income.

## 2. Reconciling the bridge finds defects the model's own gates cannot

Forcing the NOI bridge to tie **line by line** to the workbook's Year-1 NOI — rather
than asserting a total — surfaced a genuine population bug:

```
Contract Services   T-12 (col V) = 2,887      Year-1 (col AK) = 44.42
```

A per-unit figure ($2,887 ÷ 65 = $44.42) was written where an **annual total** belongs.
Year-1 NOI is overstated by **$2,843** (1.1%); held at trailing it is **$246,557**, not
$249,400. Roughly $35,000 of strike at an 8% yield on cost — immaterial to the
recommendation, material to a buyer's analyst who rebuilds the expense stack.

No gate catches this. It is not a formula error, the cell is a plausible number, and
the return tests stay green. **Only a bridge that has to add up finds it.** Build the
bridge as a table of drivers with a Basis column and make the arithmetic close before
writing a word of prose; disclose anything that will not reconcile.

Related trap on the same asset: `Master!G33` takes the unit count from the **rent roll**
(65), while per-unit expense benchmarks quoted in the handoff were struck on the
**assessor's** 66. Label the denominator in every per-unit figure ("$1,632 per unit on
66 doors") rather than silently mixing them.

## 3. When the model is rebuilt, check whether the THESIS inverted

The prior writeup's financing section argued, correctly at the time, that "the price of
this asset is set by the debt, not by the real estate" and that "leverage is negative at
every income basis." The rebuild moved Year-1 NOI from $181,124 to $249,400, and that
sentence became **false**:

| | Prior writeup | Rebuilt model |
|---|---|---|
| Year-1 NOI | $181,124 | $249,400 |
| Yield on cost | 5.55% (T-3, after reserves) | **8.18%** |
| Loan constant | 7.9836% | 7.9836% |
| Leverage | **negative by 243 bps** | **positive by ~20 bps** |
| T-3 DSCR at target leverage | 0.93x at 75% | **1.491x at 60%** |

Updating numbers inside a narrative whose *argument* no longer holds produces a
document that contradicts itself. **Diff the thesis, not just the figures.** Any
sentence asserting a direction — leverage negative/positive, cap above/below the
constant, debt-constrained vs equity-constrained — has to be re-derived, not
re-numbered.

Say the size of the move and why, in the document. A reader who saw the earlier
version must be able to see why the number changed, or it reads as a silent
restatement: "Year-1 pro forma NOI is $249,400, against $181,124 in our earlier
summary. That is a large move and it is not a silent restatement; the bridge below
shows every line of it."

## 4. Carrying two legitimate valuations without picking one

The comps supported **$3,712,735**; the income underwriting topped out at
**$3,050,000** — a **$662,735 / 22%** gap. Both were correctly derived. The document's
job is to make the gap the *headline disclosure*, not to average it away:

- Lead with the marketing range and guidance ($3,000,000–$3,950,000, guidance
  $3,700,000), because that is what the broker asked for.
- State the underwriting strike as the floor's justification — "the income
  underwriting supports $3,050,000 on normal return and lender tests, and that number
  is the reason the range has a floor rather than a hole underneath it."
- Put the gap first in Key Risk: "Ownership should project offers underwritten to the
  model, not to the comps, and should decide in advance whether it will hold out for a
  basis buyer."
- Disclose both ends of any cap-rate claim when the numerator is contested: at
  $3.7M the implied cap is **4.9% on owner-reported T-3** but **5.9% on the normalised
  T-3** — give both, labelled, rather than choosing the flattering one.

## 5. Furnished-conversion analysis: separate premium from reimbursement

Reusable framing for any "should we furnish?" question. The headline premium is never
the answer — split it:

```
gross furnished premium            +$635 to +$680 /unit/mo   (~+95-100% over posted)
  less utilities + internet onto landlord   -$250
  less furniture replacement reserve         -$67   ($4,000 basis, IRS 5-yr life)
  less incremental turn cost                 -$25   (96-102 day stays = 3-4 turns/yr)
  less downtime between mid-term stays  -$10 to -$15
= NET, SELF-MANAGED                 +$278 to +$328
  less third-party management differential  -$190  (fee points PLUS a placement fee
                                                    three-four times a year, not once)
= NET, THIRD-PARTY MANAGED           +$88 to +$138
```

The management line is the one most analyses miss, and it is what decides the answer:
self-managed the strategy earns its keep, managed it does not. Two verified corridor
comps is a thin base — say "two listings is a thin base and we are saying so rather
than dressing it up as a market" rather than presenting a range as if it were a survey.

Capitalise a pilot, never the property: five units at the model's exit cap is
~$280,000–$330,000 of value self-managed. Present as buyer optionality; do not book it
in Year 1 or in the ask.

## 6. Deal record

Delivered `The_Vintage_Apartments_Broker_Valuation_Summary.docx` / `.pdf`, 8 pages,
generated from a job-local copy of the ZONE 1/2/3/4 template (`work7/build.cjs`; ZONE 1
byte-identical to the library template, a new `buildFurnishedTable()` and Furnished
Long-Term Housing section added in ZONE 4).

Guidance $3,700,000 · range $3,000,000–$3,950,000 · likely clearing $3.3M–$3.6M ·
underwriting strike $3,050,000 (IRR 22.02%, avg CoC 10.08%, T-3 DSCR 1.491, 60% LTV at
7.00%, terminal cap 6.00%) · Year-1 NOI $249,400 at an 8.18% yield on cost.

**Render every page and read it.** The furnished table added a column and needed the
DXA widths rebalanced to ~7660; nothing but looking at the render catches an overflow.

## Related

- `vintage-okc-underwriting-8-2026.md` — the model build and the value-add-lowers-the-strike finding
- `vintage-okc-uw-writeup-8-2026.md` — the original narrative and the stale-rent-roll finding
- `submarket-anchored-promotional-sale-comps-8-2026.md` — the corridor grid this prices against
- `bov-deck-layout-and-render-notes.md`
