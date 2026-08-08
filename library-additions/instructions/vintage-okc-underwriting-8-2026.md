# The Vintage Apartments (OKC) underwriting — furnished-conversion research and three model traps (8/8/2026)

Covers: a full TMG model build on the Linux box for a **66-unit, 1953, 529 SF-average
OKC walk-up**; what a properly-researched **furnished long-term conversion** actually
does to a workforce asset (it is NOI-negative); and three model behaviours that will
mislead you if you have not seen them. Read after
`uw-model-linux-libreoffice-build-8-2026.md` (the mechanics) and
`aggressive-pricing-house-rule-8-2026.md` (the tuning rule).

## 1. Value-add capital can LOWER the supportable strike — check which test binds

The house rule says "if IRR runs out before DSCR does, add renovation programs to
boost IRR." On this deal that remedy was **actively counterproductive**, and the
reason generalises.

Max-green strike by value-add configuration (2.5-mi corridor comps, 60% LTV, 7.00%):

| Configuration | Value-add capital | Max-green strike |
|---|---|---|
| Operations only (water fixtures, RUBS step-up, billback) | $16,500 | **$3,120,000** |
| + Light interior reno, 20 units | $46,500 | $3,090,000 |
| + Premium interior reno, 20 units | $116,500 | $3,010,000 |
| Light + premium reno | $146,500 | $2,970,000 |

**Every dollar of renovation capital is equity that has to clear the 10% average
cash-on-cash test immediately, while renovation NOI ramps over years 1-3.** When
**cash-on-cash is the binding test rather than IRR**, adding programmes reduces the
price the model supports. Identify the binding constraint *before* reaching for the
value-add tab:

- IRR binding → add programmes (the house rule's case)
- **CoC binding → capital-light is worth more strike than capital-heavy**
- DSCR binding → lower LTV, never edit G66

Related: sweeping LTV found a genuine optimum. Below 60% CoC binds; above 60% DSCR
binds; **60% maximised the strike**. Positive leverage made this possible — the
going-in cap (8.4%) exceeded the 7.98% loan constant.

## 2. Master G33 is blank until FinalRR is populated — and it fails silently

Symptom: `'UW - F&C'!AC42` (T-3 DSCR), `F5`, `F7` and every `PDF Output` return
metric come back **`#VALUE!`** after a clean recalc, with no obvious cause.

Root cause chain, which takes four cell reads to find:

```
AC42 = AB41/AL42  ->  AB41 = AB19-AM37+AK31-AB31  ->  AM37 = AM35+AM36
AM36 = AK36 = AI49*AL53                       <- AI49 is Master unit count
F36  = ($AL$53*$AI49)/12                      <- "Recurring Capital Expenditures"
```

`Master!G33` ("calculated total units") is **derived from FinalRR**, not typed. Populate
Master and Final_T_12 but skip FinalRR and G33 stays `None`, so the capex-reserve row
multiplies by nothing and the error propagates into every return metric. **Populate
FinalRR before you try to read a single return.** The model map lists G33 under "verify
after paste" — treat that as load-bearing, not cosmetic.

Note the consequence: **G33 takes the RENT ROLL count**, so this asset modelled on
**65 units** (the roll) against **66 of record** (assessor). Per-unit expense
benchmarks and the printed unit count follow the roll. Reconcile the two in the notes
rather than forcing either.

## 3. The Python engine and the workbook disagree in BOTH directions

`uw-model-linux-libreoffice-build-8-2026.md` §6a/6b records the workbook being *more*
generous than `tmg valuation.py` on three consecutive deals. **Here it was less
generous**, so do not assume the direction:

| | Python | Workbook |
|---|---|---|
| Strike tested | $3,170,000 | $3,170,000 |
| Project IRR | 21.03% | **19.94%** |
| Avg cash-on-cash | 10.07% | **9.17%** |

Most of the gap is the 66-vs-65 unit basis (§2). The rule that still holds is §6b's:
**re-solve the strike in the workbook and quote the workbook's metrics**, because the
workbook is what the client receives. Python's $3,170,000 was red in the deliverable;
the workbook's max-green strike is **$3,050,000** ($3,100,000 fails CoC at 9.70%).

## 4. Refresh BOTH stale market tabs — the script only fixes one automatically

`refresh_market_tabs.py --region "Oklahoma City"` correctly rebuilt `Agency Region`
(11 OKC comps, 1940-1980, 20-250 units → Z40 sale-comp cap **6.069%**). But
**`YardiProjections` is hardcoded to a Mount Houston series inside the script** and
will silently write Houston data for any deal. It drives `Factors!N16/N17` and
therefore the terminal cap.

Houston left in place: N16 2.25% rent growth, N17 **91.00%** occupancy.
OKC-Central actual: N16 **2.50%**, N17 **83.25%** — a **7.75-point** occupancy error
straight into the exit cap.

The shipped period structure (4 history Q, 4 forecast Q, 5 annual Q4, 1 ten-year) is
**identical to the printed table on page 1 of the Yardi "Rental and Occupancy Forecast
Trends" PDF**, so substituting is a 14-value swap of `YARDI_RENT` / `YARDI_YOY` /
`YARDI_OCC` plus `YARDI_LABEL`. Do it every time the region changes.

Likewise `refresh_sale_comps.py --cma` loads the **CMA's** default rows, not this
deal's. On this build it wrote five unrelated properties. Write your own
`Underwriting Sale Data.xlsx` into `Auto Sales` instead — the sale-comps skill emits
exactly the 25-column legacy schema **including the `Column1` placeholder at L**, so it
copies column-for-column with no offset shear.

## 5. Furnished long-term conversion — the research answer, for reuse

A broker asked to underwrite "transition of units to furnished long term housing as a
value-add." Three independent research streams converged: **for an unrenovated
workforce walk-up, it is NOI-negative.** The arithmetic is reusable.

**The regulatory picture is genuinely clean** (this part is good news, and it is
primary-sourced from the OKC Municipal Code):
- OKC Home Sharing is defined as stays **"less than 30 consecutive days"** (§ 13-500(1),
  § 59-9350.38.1.A). A 30+ day lease is outside the regime — no $120/unit licence, no
  $1,100 Board of Adjustment special exception, and critically **no 10%-of-block
  density cap** (§ 59-9350.38.1.G) which would otherwise make a 66-unit conversion
  impossible in R-4.
- **Hotel tax 9.25%** (§§ 52-63.1 + 52-63.2 — *not* the 5.5% still printed on the
  City's own FAQ) does not apply: a "permanent resident" is exempt retroactively to day
  one (§ 52-62.1(9), § 52-64.1(a)(1)). Must be the **same individual** for 30
  consecutive days; any break resets it.
- **Sales tax escape is NOT a day count.** OAC 710:65-19-143(a): lodging receipts are
  taxable "without regard to the length of guest stay," and the exemption is that the
  agreement is **governed by the Oklahoma Residential Landlord and Tenant Act**. Paper
  it as a residential lease or you owe 8.625% regardless of stay length.
- **IBC occupancy is also a 30-day line.** Transient = "not more than 30 days" → Group
  R-1; a conventional apartment is R-2. Going sub-30-day is a **change of occupancy**
  under the IEBC on a 1953 structure (sprinklers, alarm, egress). Use **31+ day**
  written leases to stay clear of all three thresholds at once.

**The economics are the problem.** Against a $704 in-place rent:
- Headline furnished premium is ~**+$250/mo gross**, but ~**$200/mo is reimbursement**
  for utilities and internet moved onto the landlord ledger (electric ~$100, water/
  sewer/trash ~$0-105 depending on master metering, internet ~$60).
- Add furniture reserve (~$800/yr on a $4,000 budget-tier basis, IRS 5-year life),
  incremental turns, and a management differential of +5-10 points plus **recurring
  placement fees at 3-4 turns/year instead of one**.
- **Total opex delta $3,460/yr self-managed, $5,740/yr third-party managed.**
- Breakeven gross rent is **$1,042-$1,149/mo (+48% to +63%)**. Net NOI per converted
  unit: **-$2,972 conservative / -$1,048 base / +$478 upside.**
- **Furniture rental is dead on arrival**: CORT's OKC 1-bedroom Starter Package is
  $240/mo before the 12% damage waiver, delivery and separately-priced housewares —
  against a $250/mo gross premium. Buy at budget tier or don't do it.

**The one channel that might work is the FAA Academy**, and it is worth knowing about
for any OKC deal: the Mike Monroney Aeronautical Center runs **~22,000 resident
students a year, ~1,100 on campus daily**, on 2-5 month stays with **federally set
lodging rates** ($69.60/night long-term). The FAA publishes a live opt-in provider list
(`academy.faa.gov/sserv/api/v1/housing`, 102 approved providers) and a working operator
publishes **$49/night all-in ≈ $1,490/mo** — 2.24× the subject's in-place rent. But the
FAA criteria require **24/7 on-site management, licensed commercial lodging, ADA
compliance, on-property laundry and internet, and all utilities included**, and there is
**no FAA shuttle stop in 73106**. A single unrenovated walk-up does not clear those bars.

**Conclusion to reuse:** furnished conversion on Class C workforce product is a
*renovation + furnishing + operations business*, not a repositioning line item. Present
it as buyer optionality with the breakeven stated, pilot it on 4-6 units on natural
turnover, and **do not capitalise it into a seller's asking price**. The cheaper
adjacent lever on the same dollar is straight renovation — this deal's own rent-comp
file shows The Vic (1964, renovated, 450 SF) at $2.01/SF against the subject's $1.39.

## 6. Deal record

Universe/inputs: T-12 May-25→Apr-26 (GPR $545,290, OpEx $347,774, NOI $160,094);
rent roll 4/30/2026 (65 units listed, 64 occupied); CAD R052357400, 2026 market value
$3,470,487; effective tax rate **1.33935%** of market value (T-12 $46,482 ÷ CAD value).

Year-1 assumptions: 0% rent growth (Yardi forecasts -1.4%), 7% vacancy, 2%
concessions, 1% bad debt, mgmt fee **5%** (the property's actual), agency benchmarks
applied **only** to R&M and payroll — the two lines running hot — with everything else
held at T-12 actual rather than inflated to a higher benchmark.

**Final: strike $3,050,000 · IRR 22.02% · avg CoC 10.08% · T-3 DSCR 1.491 · terminal
cap 6.00% · 60% LTV at 7.00%, Fannie Mae — Small Balance (non-recourse, so `G48`
target reads 20%).** Sales range printed **$2,800,000 / $3,050,000 / $3,300,000**
(on the roll's 65 units — an earlier draft of this note divided by 66 and printed
$2,840,000 / $3,350,000; the workbook is the authority).

**Two corrections found when the writeup was built off this model** — see
`writeup-off-a-model-verify-dont-transcribe-8-2026.md`:

1. **The `Value-Add` column-C flags were never transcribed from Phase A.** Every tick
   is absent, the tab totals $0, and Year-1 other income equals the T-12 actual. The
   delivered strike is therefore a **base case with no value-add**. Nothing failed —
   the model recalculated clean and all three tests went green — because a skipped
   SKILL step 8 transcription is invisible to every automated gate. Assert on it.
2. **`Contract Services` Year-1 holds `44.42` where an annual total belongs** — the
   per-unit figure ($2,887 ÷ 65) was written into a total cell. Year-1 NOI is
   overstated by **$2,843**; held at trailing it is **$246,557**, not $249,400.
   Also note payroll benchmarks at **$1,300/unit** ($84,500 ÷ 65), not the $1,400
   this note originally recorded.

**Known gap on this build:** `TableRecentLeases` / `TablePropertyData` were not
populated, so the PDF's Rent Comparable Summary prints empty with `#DIV/0!`. The income
model does not read those tables, so every return metric is unaffected — but the page
is client-visible and must be filled before the model goes to a buyer. Same class of
defect as §5a in the Linux note: the gates stay green because sale and rent comps do not
feed the income engine. **Render all eight pages and look at them.**

## Related

- `uw-model-linux-libreoffice-build-8-2026.md` · `aggressive-pricing-house-rule-8-2026.md`
- `submarket-anchored-promotional-sale-comps-8-2026.md` — the comp grid this prices against
- `vintage-okc-uw-writeup-8-2026.md` — the stale-rent-roll finding that sets Year-1 GPR
