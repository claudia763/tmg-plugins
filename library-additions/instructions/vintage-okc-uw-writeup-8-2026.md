# The Vintage Apartments (2037 NW 26th St, OKC 73106) — underwriting writeup, 8/8/2026

Deal record for a **66-unit, 1953, 529 SF-average free-and-clear Oklahoma City
asset** written up from a raw T-12 and rent roll plus two Yardi Matrix reports,
with the rent and sale comp packages already delivered earlier in the same
thread. Read alongside `valuation-summary-build-on-linux-8-2026.md` (build path),
`broker-valuation-template-calloutrow-width-bug-8-2026.md` (the `wideRow` fix),
and the main library's `broker-valuation-summary` skill.

## 1. The finding that decided the whole writeup: the rent roll was stale

The rent roll's **market rent column disagreed with the accounting system by
10%** — $47,510/mo on the roll against $43,195/mo of gross potential rent in the
T-12's April column. The roll matched the books only through **September 2025**.

The T-12's GPR steps DOWN twice inside the trailing year: $47,610/mo (May–Sep
2025) → $44,820 (Oct–Dec) → $43,195 (Jan–Apr 2026), a cumulative **9.3% cut to
the posted schedule**. Read together with the rent roll, three independent facts
reconcile to the dollar:

| | |
|---|---|
| Books' market rent per unit | $43,195 ÷ 65 = **$664.54** |
| Portals advertising the 525 SF 1BR (from the rent-comp run) | **$650–$665** |
| Roll's in-place rent per occupied unit | $703.75 |
| Implied **gain** to lease | ($703.75 − $664.54) × 64 = **$2,510/mo** |
| T-12 "Loss/Gain to Lease" line, Feb–Apr 2026 | **+$2,510/mo** — exact tie |

So the owner **traded rent for occupancy**: physical occupancy recovered to 98.5%
while the submarket fell to 83.2%, but in-place rents now sit ABOVE the current
asking schedule and 52% of leases roll within six months.

**The lesson is general.** When a rent roll's market column and the T-12's GPR
disagree, do not average them and do not assume the roll is current. Divide the
T-12's most recent GPR by the unit count and compare it to what the property is
actually advertising online. If they agree, the BOOKS are current and the roll is
stale — and any "loss to lease" on the roll is fictional. Getting this backwards
would have put a $35,640/yr loss-to-lease upside story into a document where the
truth is a gain-to-lease exposure.

This also **superseded the earlier rent-comp run's headline** in the same thread,
which had read loss-to-lease off the July 2025 offering memorandum. Say so
plainly when a later document contradicts an earlier one; do not quietly restate.

## 2. Check which constraint binds before applying the debt-capacity midpoint

`narrative-variants.md` §427 says to bracket price between in-place and Year-1
debt capacity and take the midpoint. `westlake-uw-writeup-8-2026.md` already
recorded one inversion (debt capacity landed ABOVE the grid). This deal is the
opposite extreme and the midpoint rule again does not apply:

| Basis (NOI after $300/unit reserves) | Max loan @1.25x | Price @75% LTV | $/unit |
|---|---|---|---|
| T-12 as reported ($140,294) | $1,405,817 | $1,874,422 | $28,400 |
| T-3 annualized ($160,972) | $1,613,021 | $2,150,694 | $32,586 |
| Year-1 underwritten ($161,324) | $1,616,432 | $2,155,243 | $32,655 |
| April 2026 annualized ($187,032) | $1,874,155 | $2,498,873 | $37,862 |

Debt capacity tops out ~28% BELOW the deed-verified comp grid. Taking the
midpoint would have priced the asset at ~$2.32M against comps at $3.0M and
against the seller's own 2021 basis of $2.9M.

**The reconciliation is the buyer pool, and the comps prove it.** Recent trades
within a mile cleared $62,000–$85,600/unit — far above any leveraged capacity
figure — so this corridor is bought with **equity, not maximum leverage**. State
the leverage the asset actually supports (here 55.7% amortizing / 63.6% IO) and
size the buyer's equity check; that is more useful to a broker than a
debt-derived price the market demonstrably ignores.

## 3. Convergence is the strongest exhibit available on a thin-data deal

Four independent routes landed within 1.1% of each other, which is what made a
$2,900,000 recommendation defensible in a submarket with **one recorded sale in
five years**:

| Indication | Value | $/unit |
|---|---|---|
| Deed-verified adjusted sale comp grid | $3,006,357 | $45,551 |
| Yardi submarket 5-yr avg sale price × 66 | $2,904,000 | $44,000 |
| Year-1 underwritten NOI at 6.25% | $2,897,984 | $43,909 |
| T-3 NOI at the 6.29% Oklahoma average cap | $2,873,959 | $43,545 |
| Ownership's 2021 acquisition basis | $2,900,000 | $43,939 |

Recommending exactly the seller's own 2021 basis is uncomfortable and should be
raised by us rather than by a buyer. Note also that Yardi's $44,000/unit average
rests on a **single** transaction (Liberty Station, $2.55MM, 6/2023) — quote it
as corroboration, never as standalone evidence, and disclose that Yardi had not
captured the subject's own July 2021 trade.

## 4. Yardi handling specific to this report pair

- **Occupancy that never recovers.** The 8/4/2026 forecast declines in *every*
  year of a ten-year horizon (88.4% → 82.8%), never returning above 84.3%. Do not
  soften this — underwrite to it and say buyers will too.
- **Segment-level insulation, but only as direction.** The rent decline is
  concentrated in larger units (2BR −5.2%, 3BR −8.4%, both last in the metro)
  while the smallest bucket (499 SF avg — the subject's cohort) is the only
  segment with POSITIVE growth at +2.1%. Yardi publishes **no rent level** for
  that bucket, so cite direction only.
- **The supply shock is the cause and it is spent.** All 272 recent units landed
  in a single year (2025) — a 13.47% one-year inventory jump on a 2,291-unit
  base, all Renter-by-Necessity — which lines up exactly with the subject's own
  vacancy trough (peak 27.0% in Nov-25). Zero completions in 2026 and zero units
  under construction or planned in the subject's own 73106 corridor.
- **Internal inconsistency to route around:** SUB p5's unit-type table sums to
  2,954 units (and an "Overall" row of 8,430) against a stated 2,291. The p3
  bedroom chart does sum to 2,291. Use **p5 for rents only**, never for unit
  counts or SF/unit — which also makes any rent-per-SF built on those
  denominators unreliable.
- Neither report contains **any** employment, demographic, income or cap-rate
  data. If the writeup needs a demand thesis, it has to come from elsewhere.

## 5. Reconciliation traps in the source workbooks

Both inputs were already-processed TMG deliverables (the shapes `process_t12.py`
and `process_rent_roll.py` *produce*), so no library parser applies — read the
cells directly and say so.

- **Rent roll listed 65 units against 66 of record.** One unit is absent
  entirely, not shown vacant; unit-numbering forensics put the gap in building
  2037 (which starts at -19, with lone "-17"s on 2035/2036). Neither file says
  what it is — flag it, never guess a cause.
- **Category mis-mapping that corrupts per-line comparisons:** Legal &
  Professional, Other Administrative and Management Travel are coded into the
  **Management Fee** bucket (true fee $25,652 = 5.05% of income, not the $28,468
  shown); "Other Contract Services" sits inside R&M. Reclass before quoting
  expense categories.
- **RUBS coded to generic Other Income**, so the workbook's own rw/ro RUBS
  categories read $0 while $28,531 of utility recoveries sit in Other Income.
- **Rent roll reported zero concessions in all four concession columns** while
  the T-12 booked −$1,400 in April alone. The documents disagree; make the
  reconciliation a diligence item.
- **Hidden `Sheet2` carried another deal's metadata** ("Meridian Mansions
  Apartments", 114 units, `#REF!`) — leftover rediQ scaffolding. Harmless here,
  but check it before a workbook goes to a buyer.
- The rent roll is dated in the **single best month of the year** (April vacancy
  4.1% against a 16.8% trailing average and a 27.0% peak). Annualizing it
  overstates income by 15.2% versus T-3. Always locate the as-of month on the
  T-12's monthly vacancy series before trusting a roll.

## 6. Two ZONE 4 assembly index bugs — the `.slice()` gotcha, twice

`valuation-summary-build-on-linux-8-2026.md` warns that a mis-sliced narrative
array silently renders nothing. Both failure modes appeared in this build and
**neither errored**:

1. `rentalAnalysisSection.slice(0, 9)` put the "Rental Comparable Summary" label
   *after* its own table (the label is index 9, not 8).
2. `saleCompAnalysisSection.slice(5, 9)` **dropped the "Convergence of Value
   Indications" subheader entirely** — the table rendered with no heading.

Neither is visible in the XML or in a figure-presence check. Add an explicit
ordering assertion to the verification pass — cheap and catches both:

```python
i_label = txt.find("Rental Comparable Summary")
i_table = txt.find(FIRST_COMP_NAME)
assert -1 < i_label < i_table, "section label rendered after its table"
```

and assert every subheader string appears exactly once. **Then still render the
PDF to PNG and look at it** — that is how both were actually caught.

## 7. Deliverable

`The_Vintage_Apartments_Broker_Valuation_Summary.docx` (+ PDF), 6 pages, built on
the `broker-valuation-summary` template with the standard six sections plus two
substitutions driven by the variant:

- **"Financing & Debt Capacity"** replaces the plain Financing section, carrying
  the four-basis debt-capacity table.
- **An NOI bridge** replaces the template's expense-comparison table (ties
  exactly: 160,094 + 55,309 + 7,834 − 29,525 − 4,981 − 7,607 = **181,124**).
- **The Yardi submarket transaction file** replaces the agency loan-comp table.
  No Fannie/Freddie survey was run for this asset and fabricating one would be
  indefensible; the Yardi file is real, on point, and contains the subject's own
  last recorded trade.

100 figures verified present by substring check; palette clean; one logo; zero
`w:w="undefined"` widths.

## 8. Note for the BOV run on this thread

The requester asked for the BOV to list **"Chase Davis, Job Krebbs, Paul
Yazbeck."** "Job Krebbs" is almost certainly **Jon Krebbs** — the main library's
`bov-deck/assets/` carries `hs_jon.png`, and `westlake-uw-writeup-8-2026.md`
records the lineup "Greg Miller, Jon Krebbs, Paul Yazbeck, Chase Davis." Use Jon
Krebbs and mention the correction rather than printing a misspelled advisor name
on a client deck.

## Related

- `rent-comps-the-vintage-okc-8-2026.md` — same asset, rent side; assessor access
  recipe and the verified OKC comp identities
- `sales-comps-pipeline-hardening-8-2026.md` — same asset, sale side; the
  deed-verified grid this writeup prices against
- `valuation-summary-build-on-linux-8-2026.md`, `westlake-uw-writeup-8-2026.md`
- `scripts/debt_capacity_valuation.py` — the §2 table
