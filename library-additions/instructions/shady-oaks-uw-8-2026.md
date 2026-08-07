# Shady Oaks Apartments (8210 Shady Dr, Houston TX 77016) — full underwriting, 8/7/2026

Deal record for a **19-unit, 1970, owner-managed, free-and-clear Class C Houston asset**
run end-to-end on the LINUX agent server (rent roll + T-12 → sale comps → loan terms →
model → PDF → writeup). Read alongside
`uw-model-linux-libreoffice-build-8-2026.md` (the LibreOffice build path and the current
template's cell map) before reusing any of this. This is the smallest asset TMG has run
through the model, and several house defaults break at 19 doors.

## The deal

19 units (10 × 2BR/1BA @ 873 SF, 8 × 1BR/1BA @ 656 SF, 1 × 1BR/1BA @ 462 SF),
14,440 NRSF, avg 760 SF, land 28,811 SF. Owner **SHADY MC LLC**, took title 4/27/2022
from DTB Enterprises LLC; no resale since. HCAD acct **0730710110228**, 2026 noticed
value $1,474,730 (under protest, ARB 8/14/2026), aggregate rate **2.1252200 / $100**.

19/19 physically occupied; **17 of 19 pay rent** (S04 and S10 print "Preleased" with a
blank rent). Rent roll as of 3/31/2026: market $16,781/mo, contractual $14,787/mo.

**Recommended valuation $1,025,000 = $53,947/unit = $71/SF.** Year-1 UW NOI $83,888
(8.18% pro forma cap), T-12 cap 11.26%, T-3 cap 11.56%. Fannie Mae Small Balance,
80% LTV = $820,000 @ 6.74%, 1-yr IO. IRR 20.13% (target 20%), avg CoC 12.65%,
T-3 DSCR 1.344, equity multiple 2.21x, terminal cap 8.50%.

## The whole deal is the expense normalization

Trailing NOI is **$115,372 on revenue of $190,962.90** — a 39.6% expense ratio and
$3,978/unit of opex, which reads like a fantastic buy at an 11.3% trailing cap. It is
not a lean operation; it is an **incomplete record**. The books carry:

- **no management fee at all**
- **$0 of W-2 payroll** (one 1099 contractor, and it fell from $400/mo to $50/mo in 2026)
- **$1 of administrative expense** for the year (an internet charge)
- **no marketing**, no vacancy line, no bad-debt line
- **no utility reimbursement income of any kind** — the owner pays water, sewer, gas,
  electric and trash and bills back nothing

Loading an honest third-party-managed structure — the Werner Creek per-unit numbers,
$150 contract services / $600 R&M / $200 admin / $100 marketing / $600 payroll, plus a
3% management fee and a $350/unit capex reserve — takes Year-1 NOI from $115,372 to
**$83,888**. That $31,484 gap IS the valuation argument, and the writeup leads with it.

Note the direction of one line: **contract services normalizes DOWN**, from $3,750
actual to $2,850. Do not write "the seller pays nothing for contract services."

## Small-asset traps the house defaults get wrong

- **Agency expense benchmarks are absurd at 19 doors.** Fannie's set totals $2,700/unit
  ($51,300); the typed overrides total $1,650/unit ($31,350). Clear `F28:F32` and `F38`
  BEFORE typing column G or the benchmark silently wins.
- **The 3% management fee (G37) is too low and we kept it anyway.** 3% of EGI on this
  deal is $5,619 = $25/unit/month; no third-party manager takes 19 doors for $468/month
  — the real number is 5-6% or a monthly minimum. It was left at the house default for
  comparability with prior TMG deals and **disclosed as a buyer-side risk** rather than
  quietly changed. If Dmytro wants it changed, it is one cell.
- **Renovations are a NEGATIVE trade here.** Marking Light + Premium interior
  ($60,000 of capital across 19 doors) *lowered* the max green price from $1,040,000 to
  $990,000 — the capital outlay throttles IRR faster than the rent premium lifts it.
  The comp evidence agrees: Dodson Place's own renovated-vs-classic 1BR pairing shows a
  **$5/mo** premium. The 8/6 aggressive-pricing rule says to add renovations when IRR
  binds before DSCR; on a sub-20-unit asset, test it rather than assume it.
- **Water RUBS is not defensible** even though the owner pays all utilities and bills
  back nothing (~$47/unit/mo water). **7 of 10 rent comps and 3 of the selected 5
  advertise water as included**; four bundle electricity. Same ruling as Werner. Carried
  as unbooked buyer optionality. The Value-Add row's own condition ("if the water bill is
  over $65/unit/mo") also fails at $47.
- **Only Pet Fees & Rent was marked** (+$1,368/yr): three comps charge $15-30/mo and the
  subject charges nothing. That is the entire supportable value-add list.

## Leverage: unlike Werner, higher is better all the way up

Werner's price/LTV curve peaked mid-range because DSCR bound. Here **IRR binds at every
LTV** and DSCR keeps 8-14 points of slack, so the max green price rises monotonically to
the program's 80% ceiling. The final solve was run in the workbook itself (20 s/step):

| Price | $/unit | IRR | CoC | T-3 DSCR | PF cap |
|---|---|---|---|---|---|
| $1,000,000 | $52,632 | 22.55% | 13.89% | 1.386 | 8.44% |
| **$1,025,000** | **$53,947** | **20.13%** | **12.65%** | **1.344** | **8.18%** |
| $1,050,000 | $55,263 | 17.70% | 11.47% | 1.304 | 7.94% |
| $1,075,000 | $56,579 | 15.27% | 10.34% | 1.265 | 7.70% |
| $1,100,000 | $57,895 | 12.82% | 9.27% | 1.229 | 7.48% |

IRR binds at $1,025,000; **DSCR fails above about $1,075,000**, which is the agency
financeability ceiling and a better thing to quote to a seller than a return threshold.

T-3 economic loss (`UW - F&C` AC8+AC9+AC10) = **7.72%**, far under 30%, so the DSCR test
applied and the deal is not a bridge story. `AB13` occupancy 92.3% clears the 75% bridge
trigger comfortably.

## Cap-rate factors: 250 bps, and the third one is a judgement call

Marked **Low Unit Count (50) + Old Vintage (100) + Poor demographics/low AMI (100)** →
terminal cap **8.50%** on a 6.0932% Houston agency base. Note row 23 "Old Vintage" is
**100 bps in the current template, not the 25 bps in `model-map.md`**.

The demographics factor was initially left unmarked for lack of a quantified AMI figure,
then marked once the research came back: **FY2026 SAFMR for 77016 is only 84% of metro
FMR** and the area grades **F / 4th percentile for violent crime**. Marking it moved the
supported price from ~$1,040,000 to $1,025,000. It is the single most arguable input in
the build — flag it for the broker rather than burying it.

## The sale-comp grid says $1.65M and the income says $1.03M. Say so.

The `sales-comps` run indicated **$87,070.94/unit = $1,654,347.91**, but the five comps
span **$34,237 to $162,189** adjusted $/unit — a 4.7× spread, **coefficient of variation
56%**. That is not a valuation, it is the midpoint of a cloud; the honest presentation is
a range of roughly **$1.3M-$1.8M** and the income approach leading.

Two comps are institutional scale (462 and 135 units) and one is a **5-unit** building at
$167,000/unit that trades on a residential basis and drags the mean up ~$19k/unit on its
own. Haverstock Hill was trimmed at $345,714/unit — almost certainly a mis-keyed
$242,000,000 for $24.2M. The most relevant comp on size+vintage+proximity is
**4914 Kashmere St** (15 units, 1968, 2.19 mi) at $86,667/unit, but it sold **Aug-2023**,
pre-rate-move, and sits 15 days from the 3-year staleness cutoff.

At the grid's $1,654,348 the model's T-3 DSCR is **0.70** — unfinanceable on agency debt
at any leverage. That single number is the cleanest way to explain the gap to a seller.
For a re-run, set `hard_unit_range` to about `(0.3, 8.0)` to exclude the 5-unit and
462-unit records on size.

## Rents are BELOW the comps but AT the property's own asking rents

Loss to lease on the 17 rent-paying doors is **0.65%** — in-place rents are at the
property's own asking schedule, so there is no mark-to-market story against the rent roll.
Against the **comp set** there is one: 1BR $849 vs $941.56 (−9.8%), 2BR $929 vs $987.85
(−6.0%). Underwrite the dollar gap against the full 10-comp set, not the $/SF gap against
the selected five — the subject's 873 SF 2BR is the largest in the set, so its −26% $/SF
gap is a size artefact.

Rent comps marked: **Dodson View, Royal North, 5716 Pickfair, Manus** (+ Pickfair
Apartments 5 as the fifth). **Read Royal North on the net-effective line** — it asks the
set's highest 1BR at $1,020 while running half-month specials, so the real ceiling is
$977.50. HelloData did not net-effective it, and the `Tables` tab shows no specials while
two were live to the report date: **read the `Specials` tab, not `Tables`**.

Year-1 rent growth set to **0.0%**: Yardi Mount Houston forecasts four consecutive
negative/near-zero quarters (−0.6, −0.8, −0.2, +0.3) and 8 of 10 comps show identical
30/60/90-day rents. Yardi's forecast rents are in **today's dollars**, so `Factors!N16`
(2.25%) is a REAL rate — do not reuse it as nominal Year-1 growth.

## Two owner workbooks, one clean trailing twelve

Financials arrived as a 2025 file (Jan-Dec) and a 2026 file (Jan-Mar), same Toki-style
layout as Werner. Unlike Werner there is **no missing month**: Apr-2025 → Mar-2026 is
contiguous, so `combine_partial_year_t12.py` stitched it with no padding. Revenue
$190,962.90 / OpEx $75,590.85 / NOI $115,372.05, tying to the cent against both sources'
printed totals.

Reusable findings:

- **`Market Rent` is a memo row** above the income section, excluded from the statement's
  own Total Operating Income — same as Werner. Coding it as rent double-counts.
- **Werner's "Annual Material Cost omitted from Total Expense" defect is NOT present
  here** — Shady's grand rows are correct. Probe for it, don't assume it.
- **Property tax was silently miscoding to R&M.** The blank-but-labelled row
  `Other annual maintenance / month` reads as a section head, so the bare `Tax` line below
  it inherited the maintenance section rule and landed on `rm` — a $29,210 miscode that
  reconciliation cannot catch because totals tie either way. Fixed with a **run-local
  corpus override**, never the shared corpus.
- Rent-roll contractual $14,787 **ties to the cent** against the Mar-2026 rent income
  line — the strongest available confirmation of an as-of date the rent roll never prints.

## Tax findings that matter more than the model

1. **2025 property taxes are DELINQUENT: $23,516.53 as of 8/6/2026** ($19,761.79 base +
   $3,754.74 P&I), after three prior years paid on time. A tax lien primes any mortgage
   and must clear at closing. This is the most urgent item in the whole file.
2. **The books' tax line is the 2024 bill frozen.** $2,434.17 × 12 = $29,210.04 against a
   Tax Office receipt of $29,210.50 paid 1/27/2025 — a $0.46 annual variance. The actual
   2025 levy was **$24,808.81**, so the statement **overstates tax by ~$4,401/yr** and
   trailing NOI is understated by that much.
3. The owner protests every year through an agent and wins (2024 −22.5%, 2025 −11.2%).
   The 2025 bill was computed on ~$1,167,353 against a certified $1,275,000 — a $107,647
   gap most likely from a post-certification settlement.
4. Model taxes use the house rule (100% of purchase price × 2.125220%) = $21,783 at the
   strike, i.e. **below** the current levy, because the strike is below the assessment.

## Other disclosures carried into the writeup

- HCAD records **18 units / 10,976 SF** against the rent roll's 19 / 14,440 — reconcile
  before a lender's appraisal.
- HCAD **"Cooling Type: None"** on both buildings — no central A/C.
- Rent roll market rents ($16,781/mo) exceed the statement memo ($16,681/mo) by $100/mo.
  **The workbook underwrites off the rent roll** (Master → GPR $201,372); the T-12
  deliverable keeps the statement's $200,172. Disclose the difference; see
  `uw-model-linux-libreoffice-build-8-2026.md` §6 for why the two engines diverge.
- **The Opportunity Zone claim is FALSE** — tract 48201230300 is not on the CDFI Fund
  final list (three adjacent tracts are). Do not repeat it in marketing.
- Marketed Aug-Sep 2025 by Partners Real Estate, withdrawn, never priced.
- **Zero online reviews exist anywhere** — a genuine absence of evidence, not evidence of
  good operations. Renovation and opex assumptions cannot be justified from sentiment here.
- Flood Zone X today; Harris County's MAAPnext remapping expands the SFHA 33%.

## Deliverables

`RR - Shady Oaks - 3-31-2026.xlsx` · `T-12 - Shady Oaks - March 2026.xlsx` (no Capex &
Misc — the books have nothing below the printed NOI) · `CK - Shady Oaks - 8-7-2026.xlsx`
+ single-sheet PDF of `PDF Output - F&C` · `Shady Oaks Apartments - Broker Valuation
Summary.docx`. Sale comps were delivered in the prior round and re-verified this round —
the emailed pair is correct and was NOT re-sent.
