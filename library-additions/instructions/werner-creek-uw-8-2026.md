# Werner Creek Apartments (4635 Werner St, Houston TX 77022) — full underwriting, 8/6/2026

Deal record for a **36-unit, 1971/2019-renovated, owner-managed, free-and-clear
Class C Houston asset** run end-to-end (rent roll + T-12 → sale comps → loan
terms → model → PDF). Read alongside `uw-model-windows-com-build-8-2026.md`
(the Windows/COM build path and the current template's cell map) before reusing
any of this.

## The deal

36 units (29 × 1BR/1BA avg 623 SF, 7 × 2BR/1BA avg 898 SF), 24,349 NRSF, avg
676 SF. 33 of 36 occupied = **91.7% physical**; 6 units explicitly month-to-month.
Market/asking rent $35,535/mo ($426,420/yr), in-place $30,968/mo ($371,616/yr) —
loss to lease 5.1% on an occupied basis. Owner Goldenwrist Investments LLC;
PM Toki Property Management LLC. HCAD acct 0650980000002, 2026 appraised
$2,156,371, land 43,183 SF, 2025 aggregate rate **2.125220 / $100**.

**Recommended strike $2,460,000 = $68,333/unit = $101/SF**, band $2.1M–$2.7M.
IRR 20.07% (target 20%), avg CoC 10.45%, T-3 DSCR 1.273, equity multiple 2.30x.
Freddie SBL 6.63%, **70% LTV** (see leverage note below), 1-yr IO, 120-mo term,
360-mo am; loan $1,722,000, equity $915,660 incl. $126,000 of value-add capital.
T-12 cap 9.08%, T-3 cap 8.72% (owner income / normalized expenses), pro forma
6.73%. Terminal cap 7.50%.

## Two ownership documents, fourteen months, and a hole in the middle

The financials arrived as two owner workbooks: a 2025 file with Jan–**Nov** 2025
(captioned "Jan 2025 to Dec 2025" — the month COLUMNS are the truth, as always)
and a 2026 file with Jan–Mar 2026. **December 2025 does not exist in either.**
There is no contiguous twelve-month window anywhere in the data.

Resolution, and the reason it is split:

- **The T-12 deliverable keeps Dec-2025 genuinely blank** on a full Apr-25→Mar-26
  axis, with a Comments-tab note naming it. Never zero, never annualized, never
  interpolated — house rule, unchanged. `--pad-to-12` had to be taught to accept
  an *interior* gap (it previously only handled a contiguous short tail); that
  change is part of the parser work landed in the repo.
- **The model's `Final_T_12` fills Dec-2025 with the straight-line average of
  Nov-25 and Jan-26 per line item**, because the model needs twelve columns or
  every annual metric — T-12 cap rate, expense $/unit reads, the trending page —
  reads ~8% light. Disclosed in the delivery notes with the exact figures used.
  Cross-check: the filled 12-month revenue of $394,924 is within 0.3% of the
  11-month actual annualized (360,986 × 12/11 = $393,802), so the interpolation
  is not moving the answer, only completing the axis.

Reconciliation on the combined statement tied to the cent against the owner's own
printed totals (Revenue 360,985.65 / OpEx 147,662.00 / NOI 213,323.65 over the 11
real months).

Two defects in the owner's books worth knowing about for the next Toki-managed
deal: their `Total Expense` and `Net Income` rows **omit** the $712/mo "Annual
Material Cost" line, so printed Net Income ≠ printed NOI; and `Market Rent` is a
memo row above the income section that the statement excludes from its own Total
Operating Income — coding it as rent would double-count. Implied economic
occupancy from the memo row: 344,898 / 393,360 = **87.7%**.

## The expense story is the whole deal

The trailing statement carries **no management fee, no admin, no marketing, no
contract services, no vacancy line, no bad-debt line, and no payroll beyond a
$750/month 1099 contractor.** Trailing opex is $4,415/unit. That is not a great
operating record; it is an incomplete one. Whoever buys this inherits every line
the seller isn't paying.

Left on the template's agency-benchmark defaults (F28:F32 + F38 = "x"), Fannie's
institutional numbers — payroll $1,300/unit, R&M $650, contract services $250,
admin $250, marketing $150 — produced a Year-1 NOI of **$91,088** and a 71.5%
expense ratio, which is not a real number for a 36-door Houston asset. Clearing
the "x" flags and typing honest inputs ($150 / $600 / $200 / $100 / $600 per unit)
moved Year-1 NOI to $125,288 before value-add and $165,567 after. The judgement:
a 36-unit property managed by a third party never carries $46,800 of payroll, but
it does carry more than $9,000 — pick the number that a buyer's own underwriting
would survive.

Utilities are the real problem and the real opportunity: the owner pays water,
sewer, gas, electric and trash with **zero reimbursement income**, and water alone
ran $44,980 (=$1,249/unit/yr, **$104/unit/month**). The monthly series is a leak
signature — $6,361 in Aug-25, $1,331 in Nov-25 — and the 2026 T-3 has settled to
~$3,110/mo ($86/unit/mo). Marking Value-Add row C44 (fixtures, $250/unit,
−$20/unit/mo) lands Year-1 water at $36,340, which is within $1,000 of the actual
2026 T-3 run rate. That coincidence is the justification, and it is worth stating
that way rather than as a promise.

**Water RUBS was deliberately NOT marked** even though it is the single largest
dollar item on the page (~$15,120/yr at the template's $35/unit/mo). The row's own
condition is "comparable product supports additional RUBS", and **all seven rent
comps advertise water as included** — four of them bundle electricity too. Carried
in the notes as unbooked buyer optionality instead. Same reasoning killed the
second water-conservation row (C46): C44 + C46 together imply $34/unit/month of
water, below anything credible for a master-metered 1971 building.

## Leverage, not price, was the lever

At the template's stock Freddie SBL 80% Max LTV the deal is DSCR-constrained and
the maximum green price is $2,250,000. Sweeping LTV against the green rules found
the peak at **70% LTV / $2,460,000** — worth **$210,000 of price**. Full curve in
`uw-model-windows-com-build-8-2026.md` §8; the solver is
`scripts/model_price_solver.py`.

At the peak the **IRR floor binds** (20.07% against a 20% target) with DSCR slack
at 1.273, which is the inverse of the St Nicholas pattern. Per the 8/6 aggressive
pricing rule the response is to add renovation programs — but the marginal ones
here were bad trades: W/D hookups + gated parking added **$108,000 of capital to
support $10,000 of price**, so they were reverted. The final package is Light
Interior (18 doors, $1,500/u, +$50) + Premium Interior (18 doors, $5,000/u, +$150)
— together covering all 36 doors at $117,000, the same capital-efficiency logic as
St Nicholas — plus pet fees, property WiFi and the water fixtures, $126,000 total
for +$58,752/yr.

Also note the current template makes the IRR target recourse-dependent
(`G48 = 20%` non-recourse / 25% recourse), so "target IRR 25%" is no longer a
constant — read G48.

## The rents are already at market — this is not a mark-to-market story

The HelloData comp set (7 comps, 572 units, all 1965–1975, 0.6–3.3 mi) puts the
subject essentially at market, not below it:

| | subject | all 7 comps | selected 4 |
|---|---|---|---|
| 1BR asking | $951 | $921 (+3.3%) | $986 (−3.6%) |
| 1BR $/SF | $1.527 | $1.421 (+7.5%) | $1.501 (+1.7%) |
| 2BR asking | $1,136 | $1,092 (+4.0%) | $1,125 (+0.9%) |

The subject is already collecting a **premium $/SF on small 1BRs**. All the
headroom is the 5.1% loss to lease and the renovation premium — there is no
"below-market rents" pitch here, and the writeup should not imply one. The model's
own comp-supported growth (`UW - F&C`!AM4) reads +1.16%; Yardi Rosslyn reads
−0.8% YoY with the forecast negative through 2027, so Year-1 rent growth was set
to **0%** and the two figures reported as the bracket.

Rent comps marked (Rent Comparison AK): **Rose Garden, Residence At Garden Oaks,
The Melrose, Victoria Manor** — the inner ring, 0.63–1.63 mi, all 1965–1973.
Excluded: Sherwood Glen (3.3 mi, only 2 unit-level listings, both still "active"
at ~1,128 DOM — stale postings, not transacted rent) and Tara Oaks (3.3 mi, no
30-day data at all, and a self-contradictory $862 1BR against a $1,363 2BR that
would depress one average and inflate the other). Donovan Village was dropped as
the low bracket (single-story, 2BR-only, running an 11.6% concession).

**Read The Melrose on the net-effective line.** It asks the highest rents in the
set ($1,171 1BR) but is 78.4% leased and stepped its concession up from 6 weeks
free to **2 months free on 7/28/2026** — net effective collapses to $976. That
$976 is the real ceiling for renovated 1BR product here, and the underwritten
post-renovation blend (~$1,051) sits above it. That is the promotional assumption
in this package and it should be named as such, not buried.

## The market workbook's subject record is fabricated — check yours

`Rent Data.xlsx` (HelloData, 8/6/2026) carried a **subject row that is not this
property**: 27×1BR@696 + 2×2BR/**2**BA@1,036 + 7×**3BR**/2BA@1,521, 875 avg SF,
and 0% leased / 100% exposure on a 91.7%-occupied asset. All 36 of its
Unit-Level Data rows are synthetic placeholders (Active Listing = TRUE, no dates,
no DOM). Six further tabs (`Data`, `Data Analysis`, `Leasing Trends`,
`Rents by Unit Type`, `Value-Add Amenity Analysis`, `30-Day Trends`) hold stale
pivot caches from a different, much higher-rent deal — `Data Analysis` reports
2,546 leases at $3,044 average against an actual 294 rows at $1,111.

Every subject figure was overridden from the processed rent roll and every comp
average recomputed from the raw tabs. **Verify the subject row in any HelloData
export against the rent roll before using it**; the vendor's own rollups are not
trustworthy on this file.

## Sale comps and the discount

The `sales-comps` run (475 scored, top-10 trimmed at 1.0 SD) selected Las Brisas,
Garner Park, The Gardens, The Oberon and 6106 Werner St → **indicated $78,751/unit
= $2,835,053**. Trimmed: Brighton Garden Oaks at $203,466/unit (327 doors, built
2024, 0.78 mi — a Class A lease-up, irrelevant to a 1971 Class C basis) and
Sunrise Villas at $48,048/unit (Killeen, 165 mi).

The $2,460,000 strike is a **13% discount to the grid**, and that discount is the
narrative: the grid values the real estate, the income does not yet support it
because the seller is not spending what an owner has to spend. Independent
cross-checks agree with the strike more than with the grid — Yardi Rosslyn's two
closest vintage analogs are Pine Arbor (1973, 114 u) at $76K/unit and Pointe Plaza
(1965, 80 u) at $92K/unit, and the refreshed Houston agency page averages
$111,632/unit at 6.09% on properties averaging 101 units, i.e. real scale.

No subject self-sale appears in the comp universe. HCAD shows Goldenwrist took
title 04/27/2022 from Heights at Werner LLC (which bought 04/29/2019); neither
trade is in the Yardi/CoStar data, so both are pricing-relevant but uncorroborated.

## Five parcels arrived; one is the subject

The CAD cards cover five Goldenwrist parcels in Pecan Gardens. Only
**0650980000002** (4635 Werner St, the improved 36-door parcel, $2,156,371) is the
subject — the T-12 is captioned "WERNER ONLY". The others: two vacant Werner
tracts ($87,900 + $425,600), a vacant Victoria tract ($59,880), and
**0650980000010, a separate 71-unit property at 300 Victoria Dr ($5,770,806)**.

The tax cross-check is what settles it: books carry $3,607/mo = $43,284/yr, which
at the 2.125220% rate implies an assessment near $2.04M — the improved parcel
alone. Adding the two vacant Werner tracts would imply $56,752. Flag the adjacent
parcels to the seller as an assemblage question rather than folding them in.

## Data / toolchain notes

- Two new registered parsers were written for this deal and landed upstream: a
  **Toki Property Management XLSX T-12** parser (month headers are real date
  cells, not text, so the generic owner parser could not sniff it) and
  **`OwnerSheetXlsxParser`** for the owner-made rent-roll sheet
  (`APT #NUMBER` / `FLOOR TYPE` / `ADV. RENT` / `CURRENT RENT`). The rent-roll
  sheet rides *behind* the T-12 sheet in the same workbook, so that parser is the
  only XLSX one that sniffs every sheet.
- The rent roll prints **no as-of date**; 2026-03-31 was supplied via `--asof`
  and flagged as an assumption (its companion T-12 runs through Mar-2026 and
  Apr–Jun 2026 expirations still show current).
- MTM came from the source's literal `"MTM"` string in the Lease Expires column —
  an explicit owner marker, so the "MTM is never inferred" rule is satisfied.
- **Unit W18** shows status Current with a move-in date but **no lease expiration
  and no current rent**. Carried as an occupied door with a blank contractual rent
  (never back-filled from the asking rent); its $0 pulls the 1x1-600 plan average
  from $923 to $852. Open item for ownership.
- Loan terms as of 8/5/2026: UST5 4.33 / UST7 4.47 / UST10 4.63 / UST30 5.17 /
  30-day SOFR 3.622. Freddie SBL 6.63%, Fannie SB 6.68%, blended agency
  conventional 6.155%. Bridge — Debt Fund prices at 6.62% floating, essentially at
  par with fixed SBL today — an unusually flat trade-off worth mentioning.
