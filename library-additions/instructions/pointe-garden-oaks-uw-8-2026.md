# Pointe at Garden Oaks (300 Victoria Dr, Houston TX 77022) — underwriting worked notes, 8/7/2026

Covers: the full-pipeline underwriting of this **68-unit, 1963-vintage, all-bills-paid
Houston Class C** deal (rent roll/T-12 → comps → loan terms → model → valuation
summary). Read alongside `werner-creek-uw-8-2026.md` (the sister Goldenwrist property,
whose water conclusion this deal **reverses**) and
`uw-model-linux-libreoffice-build-8-2026.md` (the build mechanics) before reusing any
of this.

## The deal in one paragraph

68 doors (rent roll) of which **65 produce revenue** — V59 is a SHOP and V65/V67 are
OFFICEs, all owner-used, all carried at full market rent in the owner's own GPR memo.
Occupancy 96.92%, loss to lease **1.10%**, market rent $70,100/mo = $841,200/yr,
in-place $64,506/mo, 40,104 NRSF, 590 SF average unit. T-12 Apr-2025→Mar-2026:
revenue $840,440.13 / opex $413,749.00 / **NOI $426,691.13**, all tying to the owner's
printed subtotals to the cent.

## The three findings that carried the file

### 1. This is NOT a mark-to-market story — and three sources say so
Loss to lease is 1.10% on an occupied basis ($719/mo). Occupancy 96.92% against a
Yardi Rosslyn submarket at 89.8–90.1% and Colliers' Houston Class C at 92.9%; blended
asking rent $1,030.88 vs the Colliers Class C average of $971. The asset already
outperforms its market on **both** rent and occupancy. Any pitch built on pushing
rents to market will not survive a buyer's first read of the rent roll. Say so
explicitly — it buys credibility for the parts of the story that are real.

### 2. Water is the deal, and it REVERSES the Werner Creek call
Owner pays **$119,687/yr of water** ($1,760/unit/yr, $147/unit/mo) with **zero**
reimbursement income; the property is advertised all-bills-paid. Submarket norm for a
1963 Houston garden asset is $600–900/unit/yr.

At Werner Creek all seven comps advertised water as included, so RUBS was NOT marked.
Refreshed 8/7/2026, the answer here is the opposite: water is **not** included at four
of five comps — Sherwood Glen (1970), Tara Oaks (1973) and The Melrose (1973) bill it
back; Donovan Village (1965) puts residents on direct City accounts. Only Residence at
Garden Oaks (the newest) includes it. **The Melrose publishes the price: $50/unit/month
flat** — the single most useful number in the comp set. Underwritten at **$35/door**,
i.e. the low end of a $35–50 band net of a rent give-back, not full cost recovery.

Water is also *rising*: Feb/Mar 2026 ran **34% above** the Apr–Nov 2025 average. Leak
or rate problem; both fixable, neither free. Flag it as diligence, not upside.

### 3. The expense record is incomplete and that IS the valuation gap
Trailing opex of **$6,085/unit** carries no management fee, no contract services, no
marketing, no administration beyond $299/mo of internet, no turnover/make-ready and no
reserves. Payroll is $2,500/mo of 1099 labour ($441/unit) and the only R&M line is a
flat $1,345/mo budget accrual on a 63-year-old building. Tax and insurance are level
12/12 accruals — budget figures, not bills. Year-1 normalization lands at
**$7,718/unit** including reserves. That $1,633/unit gap, capitalized, is most of the
discount to the comp grid.

## Final model

- Deliverables: `CK - Pointe at Garden Oaks - 8-7-2026.xlsx` + single-sheet PDF of
  `PDF Output - F&C`, and `Pointe at Garden Oaks - Valuation Summary - 8-7-2026.docx`.
- **Strike $4,130,000** ($60,735/unit, $103/SF) at **75% LTV** → loan $3,097,500,
  equity $1,363,425. Freddie Mac — SBL 6.63%, 1 yr IO, 120 mo, 360 am. Monthly P&I
  $19,844 after a Year-1 interest-only period at $17,114.
- IRR **20.07%** / CoC **13.25%** / T-3 DSCR **1.2516** — all green.
  **The IRR floor and the DSCR floor bind together** at essentially the same price
  (both cross at about $4,134,000), so the deal is exactly balanced between equity
  return and debt coverage. At $4,150,000 both fail simultaneously.
- Year-1 NOI **$356,783** (8.64% going-in); T-12 NOI $426,691 (10.33% on the owner's
  un-normalized books). Terminal cap **8.25%**. Reversion $5,206,606.
- Supported band solved in the workbook: **$3,835,000 (25% IRR) – $4,434,000 (15% IRR)**.
- **The strike was solved in the WORKBOOK, not in Python.** Python's engine settled
  $4,050,000; the recalculated workbook is more generous on the reversion (equity
  multiple 2.31x vs 2.20x on identical inputs) and supports $80,000 more at the same
  gates. See §6a of the Linux build note for the diagnosis procedure — all transcribed
  inputs tie exactly, so this is an engine difference and not a transcription error.
- **Watch the PDF's "Sales Range" row.** `Assumptions!F50/H50` are FORMULAS
  (`=ROUND('(--)'!E12,-5)` / `=ROUND('(++)'!E12,-5)`) that print an **unlevered DCF**
  band discounted at 25% — here $4,200,000–$4,700,000, i.e. ABOVE the levered strike.
  That is not a bug and must not be "fixed": it says the asset's unlevered cash flows
  support more than a 75%-levered buyer can pay, which is the debt-capacity story.
  Explain it rather than overwriting the formulas.
- **Target IRR is 20%, not 25%** — `G48 = IF($G$63="Non-Recourse",$K$48,$M$48)` and
  agency debt is non-recourse. Set `M62` before reading G48 or you will tune to the
  wrong target.
- T-3 economic-loss distress test = **0%** (the owner's books carry no vacancy,
  concession or bad-debt line, so `UW - F&C` AC8:AC10 all read zero), so the
  DSCR > 1.25 rule was **live and unwaivable**.

### The LTV curve has a peak — find it, don't assume
Under `aggressive-pricing-house-rule-8-2026.md` (max price at just-green, no IRR
cushion), the max-green price rises with leverage only until the DSCR floor takes
over, then **falls**:

| LTV | 55% | 65% | 70% | **75%** | 76% | 77% | 80% |
|---|---|---|---|---|---|---|---|
| Max green | $3.70M | $3.87M | $3.96M | **$4.05M** | $4.07M | $4.05M | $3.93M |
| Binding | IRR | IRR | IRR | **IRR** | IRR | DSCR | DSCR |
| T-3 DSCR | 1.964 | 1.570 | 1.416 | **1.284** | 1.259 | 1.250 | 1.251 |

(The table above is the PYTHON engine's sweep, which is what leverage selection was
based on. The final price came from the workbook — see "Final model" above.) The
Python peak was **76% / $4,070,000**, where the two floors cross; **75%** was chosen
as a standard agency leverage point 0.5% below it with more DSCR coverage, which is
cheap insurance on an SFHA asset whose expense record a lender will re-underwrite.

**Sweep LTV in 1% steps around the crossover; the naive 45–80% sweep in 5% steps
misses the peak entirely.** Then re-solve price in the workbook at the chosen LTV —
on this deal that added $80,000. The workbook's own sweep at 75% LTV:

| Price | $4,100,000 | **$4,130,000** | $4,150,000 |
|---|---|---|---|
| IRR | 20.57% | **20.07%** | 19.74% FAIL |
| CoC | 13.51% | **13.25%** | 13.09% |
| T-3 DSCR | 1.2635 | **1.2516** | 1.2438 FAIL |

### Factors — terminal cap build-up
6.0775% sale-comp cap (TX, 1963 ±10 yrs, T-12, n=24 — within 3 bps of Colliers/MSCI-RCA's
published Houston metro 6.1%) + **Old Vintage 100 bps** (this template prices it at 100,
NOT the 25 in `model-map.md`) + **Low Unit Count 50 bps** + custom row 24 **FEMA Zone AE
50 bps** + custom row 25 **narrow buyer pool 25 bps** = 225 bps → 8.3025% → MROUND
→ **8.25%**.

### Value-Add marked
Light + Premium Interior Renovations (34 doors each, $221k capital, $81,600/yr),
Pet Fees & Rent ($4,896 — the property advertises $25/mo pet rent and books **no pet
income line at all**), Reserved Parking ($4,080 — $50/mo assigned, only ~36 of 65 doors
taking it; HCAD shows just 15,360 SF of paving so supply is genuinely scarce), Water
Conservation Fixtures ($16,320), Water RUBS at $35/door ($28,560). Total capital
$238,000.

**UNMARKED "Reduce opex to comp averages"** — same call as St Nicholas Place. This
owner is already below every benchmark *because the lines are missing*, so the row
prints a confusing negative delta.

## Traps specific to this deal

- **HCAD says 71 units / 44,120 SF; the rent roll says 68 / 40,104.** Use 68 for all
  per-unit math and disclose the gap.
- **Three non-revenue doors** at full market rent in the owner's GPR memo. Per-door
  economics must be run on 65, per-unit valuation on 68.
- **Live listing showed four vacant 1BRs on 8/7/2026 against a roll showing two**, and
  one of the four is V59, the "SHOP". Live occupancy may read ~93.8%, not 96.9%. Open
  item — recommend the seller provide a current roll and leasing-activity report before
  launch.
- **The subject's own prior trade is a trap.** 300 Victoria Dr sold for **$10,000,000
  on 04/27/2022 recorded as 104 units** ($96,154/unit). That is the combined Goldenwrist
  package (68 + 36 = 104 units; recorded 64,453 SF = 40,104 + 24,349 **exactly**), not a
  Pointe-only trade. Exclude from the grid but disclose — a buyer will find it.
- **Live tax protest**: filed 5/14/2026, ARB hearing **8/27/2026**, agent O'Connor &
  Associates, status "Not Certified". Every $1M of assessed value = $21,252/yr at the
  2.125220 rate. The 2026 CAD rate column prints 0.000000 (Texas jurisdictions adopt
  Sept–Oct), so **2025 rates are the right ones to use**.
- **FEMA Zone AE / SFHA**, DFIRM 48201C, White Oak Bayou watershed. Flood insurance is
  lender-required; ~65% of Independence Heights homes were damaged in Harvey; a
  countywide re-map (HCFCD MAAPnext) is in progress. Elevation Certificate and FIRMette
  are mandatory diligence items — and it is an exit-liquidity fact, not just a diligence
  one, hence the 50 bps in Factors.
- **Houston's Apartment Inspection Ordinance** passed 5/6/2026 — public registry at 10+
  citations in six months — against a 1963 asset with zero capex on the books.
- **Freddie retired SBL on 4/15/2026**, replaced by Optigo "Conventional Small"
  ($2–10M). The loan-terms workbook row is stale on the program *name and size band*;
  the economics carry across. A $3.5–4.5M request clears the new $2M floor comfortably —
  worth checking, because it is a live problem for smaller assets (Werner Creek does not
  clear it).
- **No named source publishes a Houston Class C or 1960s cap rate.** Texas is a
  non-disclosure state. Published only: Colliers/MSCI RCA Houston metro all-class
  **6.1% Q2 2026** (+30 bps YoY, a 9-year high) and HCAD's own market study citing
  CoStar market 6.6% / transaction average 7.3%. Do not invent a Class C number.
- **Yardi self-contradicts twice**: YoY rent growth published **three ways for the same
  quarter (−0.3% / −0.8% / −1.0%)** — report the bracket, never the average — and
  inventory at 22,417 vs 26,338 units. Footnote both. Its forecast also states verbatim
  that historical rents are nominal but forecast rents are in today's dollars, so
  out-year growth (2029 +2.2% → 2036 +3.0%) is **REAL** and must not be compounded into
  a reversion.

## Rent comps (no redIQ export existed — built from primary web research)

| Property | Built | Units | 1BR ask / SF | Water |
|---|---|---|---|---|
| Residence at Garden Oaks | 2015? | 98 | $950 / 656 | **Included** |
| The Melrose, 712 Pinemont | 1973 | 88 | $999–1,250 / 712 | **Property — $50/mo** |
| Sherwood Glen, 3805 Sherwood | 1970 | 56 | $955–1,270 / 645 | Property |
| Tara Oaks, 3800 Sherwood | 1973 | 126 | $845 / 625 | Property |
| Donovan Village, 601 W Donovan | 1965 | 78 | $899 / 605 | City — resident direct |
| **SUBJECT** | 1963 | 68 | **$975 / 561** | **INCLUDED** |

Subject headline is $1.74/SF against a comp band of $1.27–$1.49 — **that is not a rent
premium, it is the utility bundle.** Net of $237/door/mo of owner-paid utilities the
subject prices at $1.32/SF and sits mid-pack. Donovan Village is the cleanest read of a
fully unbundled 1960s Class C door in this submarket and implies ~$834 of unbundled base
rent against $975 all-in — a $141 gap against $237 of utilities, i.e. the all-bills-paid
structure is currently **value-destructive at the margin**. Present that as the upper
bound of the opportunity, not the base case; a master-metered 1963 building cannot
replicate Donovan's structure without capital.

Because there was no lease-level export, **`TableRecentLeases` was cleared rather than
left carrying the previous deal's leases** — presenting stale prior-deal comps is
forbidden by the skill. Disclose which PDF elements consequently render empty.

## Sale comps and the pricing gap

Grid (5 comps, drift-adjusted) $75,841/unit = $5.16M · same comps on $/SF $61,798/unit
= $4.20M · ex the 130-mile Taylor outlier $66,002/unit = $4.49M · Yardi's only two
pre-1975 Rosslyn trades ~$84,000/unit · 2026 HCAD assessment $84,865/unit = $5.77M.
Las Brisas (68u, 1963, 2.89 mi, 10/2024, $81,029/u) is a near-twin and the standout.

The strike at $59,559/unit is a **21.5% discount to the adjusted grid, 29.8% to the
HCAD assessment, 9.8% to the ex-Taylor read**. Be honest about grid credibility:
vintage and unit count match well (comps mean 1965 / 66 units) but 40% of the grid is
Central Texas because the scoring flattens everything past 3 miles, and the comps
average 853 SF against the subject's 590 SF.

## Debt capacity is the real pricing constraint

At a $5.5M test value every program is **LTV-constrained** on trailing NOI as reported
($426,691 covers 1.25x with room to spare) but flips to **DSCR-constrained** on a
normalized NOI of ~$260,000, and agency proceeds collapse from $4.4M to about $2.7M.
**Both answers must be stated together** — the $1.7M spread between them is the whole
financing conversation, and the buyers who bid off the first column are the ones who
have not yet added a management fee.
