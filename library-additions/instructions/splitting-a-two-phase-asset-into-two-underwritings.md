# Splitting a two-phase asset into two independent underwritings (8/7/2026)

Covers: what actually has to change when a property already underwritten as ONE asset has
to be re-underwritten as TWO standalone assets — and which inputs are safe to reuse
verbatim. Written from Westlake East (59u, 1976) / Westlake West (115u, 1973), Lubbock TX,
where the broker asked for "separate underwriting for each… you may reuse the rent comps
and the sale comps already done."

Read alongside `uw-model-linux-libreoffice-build-8-2026.md` (build mechanics),
`combining-two-property-t12s.md` (the inverse operation), `westlake-uw-writeup-8-2026.md`
and `westlake-east-loan-assumption-8-2026.md` (this deal's specifics).

## Why this comes up

The listing question is usually "which piece actually sells." On Westlake the managing
director's words were: *"The bigger one for sure has to be listed but there is a small
chance where he holds on to the smaller one."* Once that is live, the combined model is no
longer the deliverable — two models are, and each has to stand on its own in front of a
different buyer.

## The trap: what is NOT reusable

The instinct is that only the unit count and the financials change. Four things change that
are easy to miss, and three of the four move value.

### 1. The sale-comp grid re-adjusts. This is the big one.

TMG's comparable sale grid adjusts each comp's $/unit **for the subject's vintage and
average unit size**:

```
year_adj = (subject_year_built - comp_year_built) * 0.005
size_adj = (subject_avg_sf - comp_avg_sf) / subject_avg_sf * 0.25     # NOTE: subject SF is the denominator
adjusted_$/unit = comp_$/unit * (1 + year_adj + size_adj)
indicated $/unit = simple average across comps (equal weighting)
```

Both terms key off the subject, so **the same comp set gives a different indicated value
for each phase.** On Westlake, with identical comps:

| Subject | Vintage | Avg SF | Indicated $/unit | Indicated value |
|---|---|---|---|---|
| Combined (as printed by the client, 8/6/2026) | 1973 | 980 | $61,441 | $10,690,719 |
| East | 1976 | 1,122 | **$63,798** | $3,764,066 |
| West | 1973 | 907 | **$60,501** | $6,957,567 |

East re-adjusts **+3.8%/unit above the combined figure** purely because its units average
1,122 SF against the combined 980. Reusing the combined $61,441 for both phases would have
underpriced East by ~$139k and overpriced West by ~$108k.

`scripts/sale_comp_grid.py` does this. Its `--self-test` reproduces the client's printed
combined grid to the cent (all five adjusted $/unit, the dollar adjustment, and both
totals), so the formula is verified rather than assumed — run it before trusting a
re-adjusted grid.

**The easy mistake:** dividing the size adjustment by the *comp's* average SF instead of
the subject's. It looks right and is wrong; on a small-unit comp it produces 16.3% where
the model prints 9.87%. Check any hand-built grid against the self-test.

**Recombination check:** East + West should land within ~0.5% of the combined grid
(Westlake: $10,721,633 vs $10,690,719, +0.29%). A bigger gap means an input is wrong.

### 2. Cap-rate Factors are per-phase

`Factors` row 18 **"Low Unit Count" (50 bps) fires on one phase and not the other** — East
at 59 units qualifies, West at 115 does not. That is 50 bps of terminal cap, i.e. real
money, and it is the correct answer: a 59-unit asset genuinely has a thinner buyer pool
than a 115-unit one. Vintage, location and demographics factors normally apply to both.

Apply the *same* answer to both models for any factor that is a property-level judgement
(demographics/AMI, tertiary location) so the two are comparable; let only the genuinely
size- or vintage-driven ones differ.

### 3. Rent comps are reusable at the property level, but the SUBJECT row is not

The comp rows in `TablePropertyData` carry over unchanged. The subject row does not — units,
occupancy, average SF and year built are all per-phase. On Westlake the subject's average
unit size moved 907 → 1,122 SF between phases, which flips whether the subject reads as
above or below the 944 SF comp average.

Watch the $/SF trap: East's larger units make its $/SF gap to the comps look worse than
West's while its $/unit rent is higher. Underwrite the dollar gap, not the $/SF gap
(same ruling as Shady Oaks).

### 4. Tax rates are shared; assessed values are not

Both parcels sat in the same five taxing units at the same 1.769426 aggregate rate, but
carried different assessed values, different per-unit values, **different CAD valuation
methods** (East "Income", West "Override") and opposite five-year trajectories. Do not
share a tax assumption between the two models — only the rate.

## What IS safe to reuse verbatim

- The comp *set* itself (which properties, their addresses, vintages, sizes, prices, dates).
- The Year-1 expense normalization **per unit** ($/unit contract services, R&M, admin,
  marketing, payroll, insurance, capex reserve) and the management-fee percentage. Reusing
  the per-unit basis is what keeps the two models comparable with each other and with the
  combined work already delivered. Convert to dollars at each phase's own unit count.
- Growth rates, hold period, origination cost, sales expense, target IRR.
- The Value-Add programs, **pro-rated by unit count**.
- The market tabs (`Agency Region`, `YardiProjections`) — same market, same submarket.

## Utilities and any metered line: never blend

Westlake East ran $503/unit/yr water and $524/unit/yr electric against West's $330 and
$210. Both phases are separately metered with separate GL accounts under separate ownership
entities. Leave each phase on its own T-12 actual. A blended utility assumption would have
overstated West's expenses by ~$55k/yr and understated East's.

The general rule: **if the two phases have separate ledgers, every line that is actually
metered or billed separately stays separate.** Only lines you are deliberately normalizing
to a benchmark (payroll, admin, marketing, insurance, reserves) should share a per-unit
basis.

## When one phase has a shorter operating history

Westlake East's ownership entity opened its books 10/1/2025, so its "T-12" is a T-10.
The model's trailing-twelve column is `'UW - F&C'!V = SUM(F:Q)` across the twelve monthly
columns — **not** the Final_T_12 "Adjusted Total" column O. Ten populated months therefore
understate trailing revenue *and* expenses by ~17% and corrupt every T-12 metric.

Fill the missing months in the **model's** `Final_T_12` grid with the actual-months average
so the trailing-twelve column is a true annual run rate, and say so on the tab. The T-3
column (`AB = SUM(O:Q)*4`) is untouched and stays pure actuals — which matters, because
T-3 DSCR is the green test, so the gross-up cannot flatter the binding metric.

**Keep the delivered T-12 workbook itself genuinely short** (missing months blank, not
zero, not grossed up) per the house rule in `combining-two-property-t12s.md`. Only the
model's internal grid is annualized, and it is disclosed.

## Check the ownership structure before assuming "two phases, one seller"

On Westlake the two parcels turned out to be owned by **different LLCs** (V Westlake LLC,
deed 9/12/2025; Valor West LLC, deed 4/28/2022) with different debt, different chart-of-
accounts numbering and different books-open dates. That is not a formatting detail — it
means two PSAs, two sets of title and estoppels, and it explains the short operating
history. It also means one phase can carry assumable agency debt while the other does not.

Pull the CAD sales history for **both** parcels before writing the deal structure.

## Shared physical plant is a real diligence item

Westlake's single 1,645 SF stand-alone leasing office sits on the **West** parcel and serves
both phases. If the phases sell separately, the office, the on-site staff and any shared
amenity need an easement or a management agreement. Look for it — it will not be in the
rent roll or the T-12.

## Tie-out that proves the split

Reconcile the two phases back to the combined deliverable on every headline figure before
shipping. Westlake: 59 + 115 = 174 units, 66,201 + 104,250 = 170,451 SF, 49 + 94 = 143
occupied, $46,673 + $84,708 = $131,381/mo contract rent, $11,738.85 + $13,472.05 =
$25,210.90/mo other income. All exact. If the phase splits do not re-sum, the phase filter
is wrong.
