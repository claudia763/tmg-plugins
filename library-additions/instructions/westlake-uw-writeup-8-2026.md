# Westlake Apartments (5202 Bangor Ave, Lubbock TX 79414) — valuation summary, 8/6/2026

Deal record for a **174-unit, two-phase (1973 west / 1976 east), stabilized,
free-and-clear Lubbock asset** written up from a client-supplied underwriting
summary PDF rather than from a raw rent roll and T-12. Read alongside
`valuation-summary-build-on-linux-8-2026.md` (the build path) and the main
library's `broker-valuation-summary` skill.

## The request and the archetype

Dmytro's framing verbatim: *"pricing is limited by a lack of major upside other
than other income increases and slight payroll reduction. Since the property is
already very well stabilized, its valuation must effectively trade on a market
capitalization rate basis and align with free and clear debt sizing."*

That maps to the **debt-capacity-constrained pricing** variant in
`narrative-variants.md` (§427), specifically sub-sections 1 (named
debt-capacity section with its own table), 2 (rate sensitivity as risk #1),
5 (zero-value-add → NOI bridge) and 6 (Yardi handling). **But the variant's
default conclusion inverted here** — see below.

## Inputs

- `Westlake Apartments - UW Summary.pdf` (9pp) — the model's PDF Output - F&C
  tab. This is the authoritative source for units (174), NOI, cap rates, loan
  terms, rent comps, sale comps and the agency loan-comp table.
- `Westlake - E-Brochure.pdf` (35pp) — Yardi Matrix data sheet. Says **173**
  units and a different plan-level mix; the model/rent roll wins.
- `Yardi Submarket Report.pdf` (9pp) — Lubbock–Southwest, June 2026.
- `Yardi Rental and Occupancy Forecast Trends.pdf` (4pp) — same submarket,
  forecast revised 8/4/2026. **Chart-only; must be read as rendered images.**

## Numbers that tie (verified, reusable as regression checks)

- Loan constant at 6.13% / 30-yr = **7.2952%**, reproducing the model's printed
  **$52,434/mo** on $8,625,000 exactly. Always run this check first — it
  validates every downstream debt number.
- Cap rates at $11.5M reproduce the model exactly: **7.85%** on T-12 NOI after
  reserves ($902,980), **7.15%** on T-3 with normalized expenses ($822,774),
  **8.56%** on Year-1 ($984,573). The model's "UW NOI" of $827,918 =
  T-12 income $1,831,017 − Year-1 normalized expenses $942,199 − $60,900
  reserves, giving a **7.20%** normalized in-place cap.
- **NOI bridge ties to the dollar:** 963,880 + 124,268 − 32,341 + 64,728
  − 75,062 = **1,045,473**; less $60,900 reserves = **984,573**.
- The +$124,268 decomposes as **$64,727 loss-to-lease burnoff + $59,541 of
  3.0% growth on market rent** ($1,984,692 × 1.03 − $1,919,965).
- All per-unit expense figures tie on **174 units** ($4,983/unit T-12,
  $5,415/unit Year-1), confirming 174 as the model's basis.

## The three findings that carried the writeup

### 1. Operations net NEGATIVE; the pro forma is a rent-growth assumption
Stripping the 3.0% growth, the operational actions sum to **−$42,675**
(+$64,728 other income − $75,062 expense normalization − $32,341 vacancy
normalization). Meanwhile Yardi's own 8/4/2026 forecast for this submarket puts
rent growth at **0.6%** through YE-2026 and below 2.0% until late 2027 — i.e.
the model underwrites roughly **3x the forecast**. At 1.0% growth, Year-1 NOI
after reserves falls $984,573 → **$947,261** and the pro forma cap 8.56% →
**8.24%**. The safe answer is that the valuation does not depend on it: at a
7.20% normalized in-place cap the price stands with zero rent growth.

### 2. Occupancy cannot be reconciled — the dominant risk
Yardi says **83.2%** physical (USPS-derived), the model's own comp page says
**82%**, but the T-12 books only **$90,313** of vacancy = **4.70%** of GPR
(95.3% economic), and T-1 annualized is still only 5.9%. If economic occupancy
were truly 83.2%, T-12 NOI drops $902,980 → **$670,739** and value at 7.20%
drops to ~**$9.3M ($53,539/unit)** — a ~$2.2M swing. The submarket held
90.6–91.1% all year and is not forecast below 89.0% for a decade, so **there is
no market-level explanation**; if the number is right it is property-specific.
Also note: bad debt and concessions are **$0 in all twelve months**, which is
not credible for a 54%-AMI qualified-census-tract asset.

### 3. Agency "value/unit" data blends refinances with sales — do not price off it
The agency table's own search criteria read *"All sales **and refinances**."*
**Parkside** proves the distortion: 171 units, 1972, zip 79414 — it **sold** for
$10,000,000 (**$58,480/unit**) in 7/2024 (independently confirmed by the Yardi
submarket transaction table) and carries an agency-reported value of
$14,430,000 (**$84,386/unit**) in 1/2026. Same asset, +44% in 18 months.
Treat agency $/unit as an appraisal-informed **upper bound**; anchor on the
sale grid.

**Also flagged:** the agency table contains **"Westlake Apartments East"** —
59 units, 1976, 5128 Aberdeen Ave, 79414, $5,200,000 ($88,136/unit) at 6.95%,
originated 6/30/2025. The subject's floor-plan codes split **W- = 115 units /
E- = 59 units**, and the brochure confirms two assessor parcels
(R138162, R39044) and a 1976 phase. The unit count matches exactly, so this
is very likely the subject's own east phase — which would mean **existing
agency debt on part of an asset being underwritten free and clear**. Raised as
a title/payoff confirmation item, not asserted as fact.

## Where the variant's default conclusion inverted

The variant says to bracket price between in-place and Year-1 debt capacity and
**take the midpoint as the recommendation**. That would have produced ~$13.25M
here — far above the sale grid. On this deal debt capacity is **not** the
binding constraint:

| Basis | NOI | Max loan @1.25x | Price @75% LTV |
|---|---|---|---|
| T-12 as reported | $902,980 | $9,902,173 | $13,202,897 |
| T-12 normalized opex | $827,918 | $9,079,035 | $12,105,380 |
| T-3 normalized opex | $822,774 | $9,022,625 | $12,030,167 |
| Year-1 underwritten | $984,573 | $10,796,930 | $14,395,907 |

At $11.5M a lender sizing to normalized trailing income supports $9,079,035 =
**78.9% LTV**, so **LTV binds, not DSCR** — the favorable case, and a genuine
marketing point (full 75% agency proceeds on in-place income, which little
1970s product achieves). The pricing constraint is the **sale comp grid**
($61,441/unit adjusted → $10,690,719); the recommendation of $11.5M is a
**+7.6% premium** to it, defended on unit size (980 SF vs 593–899 SF across the
grid; the size adjustment exceeds the vintage adjustment in 4 of 5 comps).

**Lesson: run the debt-capacity test, but check which constraint actually
binds before applying the variant's midpoint rule.** When debt capacity lands
above the comp grid, say plainly that the real estate sets the price and the
debt merely has to keep up — that *is* the finding.

## Rate sensitivity (the risk-#1 table)

Normalized in-place NOI $827,918, 1.25x, 75% LTV:

| Coupon | Constant | Max loan | Price | DSCR @$11.5M/75% |
|---|---|---|---|---|
| 5.63% | 6.9117% | $9,582,848 | $12,777,131 | 1.389x |
| 6.13% | 7.2952% | $9,079,035 | $12,105,380 | 1.316x |
| 6.63% | 7.6877% | $8,615,509 | $11,487,345 | **1.249x** |
| 7.13% | 8.0887% | $8,188,419 | $10,917,892 | 1.187x |

~$620–670K of price per 50 bps. **50 bps is the entire cushion** between full
75% proceeds and a DSCR-constrained loan — mitigant is procedural (term sheets
from 2–3 lenders pre-launch). Positive leverage: 7.85% T-12 cap and 8.56%
Year-1 cap both clear the 7.2952% constant (+55 / +126 bps); normalized in-place
7.20% is 10 bps inside it, i.e. neutral — which is the answer to "why can't
buyers pay a lower cap rate."

Script: `scripts/debt_capacity_valuation.py` runs all of the above.

## Yardi handling (per variant §6) and document contradictions

The submarket documents contradict each other; **footnote both as published
rather than reconciling**:
- Rent growth Q2-2026: **0.4%** (Submarket Report) vs **1.3%** (Forecast doc).
  Do not blend these in one exhibit.
- Rent level June/Q2 2026: **$985** vs **$981**.
- Unit counts: **8,951** (cover/narrative/charts) vs **10,167** (unit-type
  table, which is also the source of the 893 SF average). Don't present
  "$985 on 8,951 units" and "$985 ÷ 893 SF" as the same population.
- Monterey Villas rent-up: table says **04/2026**, narrative says **April
  2027**. Avoid depending on it; say the 279 units complete 12/2026 and are
  counted in 2027 inventory growth.
- **Real-vs-nominal trap:** the forecast doc's footnote reads *"Historical
  rents are in nominal terms, forecasted rents are in today's dollars."* Taken
  literally the out-year 2.4–3.6% growth is **real**. Do not compound the
  10-year $1,331 figure into a reversion without confirming the basis.
- Yardi's "$85K/unit" 5-year submarket sale price is an **unweighted mean of
  annual averages**; unit-weighted it is $94.8K. Recent same-vintage prints are
  the relevant evidence: Parkside $58K (7/2024), Aspen Village $51K (4/2025),
  Farrar $57K (4/2026) — all **below** the $66,092/unit recommendation, which
  is why the grid adjustments have to be defensible line by line.

## External corroboration (web research, 8/6/2026)

Worth having on file for the next Lubbock deal:

- **75% loan-to-cost is the achieved debt quantum in this submarket.** Dividing
  Yardi's loan column by its price column: Parkside 75.0% (7/2024), Aspen
  Village 75.0% (4/2025), Courtyards 75.0% (4/2023), Lakeway 75.1% (9/2022),
  **Farrar 75.0% (4/2026)**. The lone exception, **Lubbock Square at 85.2%**
  (3/2022) into a 2024 maturity, is now marketed by Marcus & Millichap at 79%
  occupancy — the cautionary comp for over-leverage.
- **Agency window is favorable:** FHFA 2026 caps $88B each (+~20%), and
  **workforce-housing loans are excluded from the caps entirely**; both agencies
  ran below cap pace through H1 2026. **Freddie retired SBL on 4/15/2026**
  (replaced by "Conventional Small," $2–10M), so a 174-unit / $8.6M request is
  standard Fannie DUS or Freddie Conventional. **Fannie Sponsor-Dedicated
  Workforce** gives a pricing benefit for restricting 20%+ of units at ≤80% AMI
  — a live lever on a 54%-AMI qualified-census-tract asset.
- Fannie Tier 2 (1.25x / 80%) 10-yr all-in was quoted **5.78–6.18%** in early
  August 2026, which brackets the model's 6.13% — the rate assumption is market.
- **Buyer pool is small private / 1031 / regional, with no institutional bid
  anywhere in the data.** Bentwood (280 units, 1/2026) traded to a California
  group via **1031 plus assumption of existing Fannie debt**, one of eight West
  Texas buys that sponsor made in 2025.
- **TMG brokered Birchwood** (236 units, 1977, 6402 Albany Ave, REO, closed
  ~11/2025) — by vintage and location the best comp in the market, and its price
  is **not public**. Check our own file before building any Lubbock comp grid.
- **Vendor occupancy conflict:** Yardi ~90.7% vs **HUD PD&R CHMA (CoStar data,
  as of 12/1/2024) 13.0% vacancy = ~87.0% occupancy** metro-wide, with student
  housing at 17.4% vacancy and HUD stating demand for new rental units of
  "No Units" through 12/2027. ~370 bps apart on different universes. Pick one
  vendor and disclose it; do not blend.
- **Texas is a non-disclosure state** — no named institutional source publishes
  a Lubbock cap rate, and CBRE's Cap Rate Survey covers **Class A only**, no
  Class B/C and no Lubbock. Any "CBRE Lubbock Class B/C cap rate" circulating
  online is fabricated. The Yardi submarket transaction table is the best comp
  evidence that exists for this market.
- Two claims to reject on sight: a "+14,200 jobs / +8.89%" Lubbock job-growth
  figure (payrolls are ~176K; actual growth is ~1%), and Maxey Park at
  $248K/unit for a 1964 C/C asset (portfolio allocation or data error — exclude).
- BLS LAUS: **Lubbock MSA unemployment 4.5% (June 2026) vs 3.8% (June 2025)**.
  TTU enrollment **42,272** fall 2025 (12th class day), a record and the fourth
  straight increase; note THECB's certified figure is 40,869 on a different
  basis — cite one and state it.

## Deliverable

`Westlake_Apartments_Broker_Valuation_Summary.docx` (+ PDF), 7 pages, built on
the `broker-valuation-summary` template with two sections added beyond the
six-section default: a named **"Valuation Basis: Market Cap Rate and
Free-and-Clear Debt Capacity"** section (debt-capacity + rate-sensitivity
tables) and an **NOI bridge** table replacing the template's
expense-comparison table. 63 figures verified present by substring check.

## Model artifacts worth reporting back to the analyst

- Page 2's floor-plan block prints **`#DIV/0!`** in the Pro Forma and Pro Forma
  $/SF columns for six of ten plans — cosmetic, but it must be cleared before
  anything derived from this tab goes to a buyer.
- Printed **$/SF on the sales range is inconsistent**: $64 / $67 / $71 for
  $11.2M / $11.5M / $12.4M. Only the midpoint ties to the 170,451 SF implied by
  the floor-plan table; the low and high ends imply ~174,800 SF. Per-unit
  figures are unaffected.
- Page 3 of the UW summary is **blank** (no text, no images).
