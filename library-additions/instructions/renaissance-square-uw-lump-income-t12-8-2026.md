# Underwriting a LUMP_INCOME T-12, and four model traps that silently move the answer (8/8/2026)

Covers: what to do when the owner's T-12 is a single "Income" line with **no GPR,
vacancy, concessions, bad debt or RUBS** and an expense side missing payroll,
utilities, management, admin and marketing; how to handle a rent roll whose
Market Rent column sits BELOW in-place rent; **Arkansas property tax, which does
not reassess on sale**; and four template behaviours that produce a plausible
wrong answer with every gate green. Read after
`uw-model-linux-libreoffice-build-8-2026.md` (mechanics) and
`aggressive-pricing-house-rule-8-2026.md` (tuning). Worked on **Renaissance
Square, 2401 County Ave, Texarkana, AR 71854** — 65 units, 49,400 SF, est. 1970.

## 1. A LUMP_INCOME T-12 cannot drive expenses — say so and use the benchmarks

The statement ties perfectly (rev $509,418 · exp $77,145 · noi $432,273, all six
cross-checks exact) and is still unusable. Codes `ll v nr bd rw ro` exist as
**label rows with twelve blank months**; `pr ad m mf w tr e o` are absent
entirely. The result is a **15.1% expense ratio** and a printed NOI of $432,273
that must never be quoted.

What to do:

- Turn on the agency benchmark flags `F28:F32` and `F38` (Contract Services,
  R&M, Admin, Marketing, Payroll, Insurance). On this deal they produced
  $250 / $650 / $250 / $150 / $1,300 / $800 per unit.
- **Utilities have NO agency flag** (rows 33–36 are "blank = T-12"), so a T-12
  with no utility lines leaves them at **$0** — on a 1970 master-metered
  walk-up. You must input them. See §3 for the units trap.
- Sanity-check the total against agency records for the SAME market rather than
  a national benchmark. TMG's Fannie/Freddie workbook had five Texarkana
  originations: total expenses **$3,730 / $4,172 / $4,507 / $5,542 / $6,132 per
  unit**. This build landed at **$4,934/unit** (64% ratio) — inside the range,
  toward the conservative end, and defensible on that basis.
- Say in the notes that the expense build is a **construction, not a
  transcription**, and that the seller should be asked for the missing lines.

## 2. When the rent roll's Market Rent is BELOW in-place rent, restate it

The owner's Market Rent field was set at each unit's last lease event and never
refreshed: **26 identical 550 SF 1/1s carried 14 distinct values spanning
$375–$800**, and 28 of 58 occupied units billed above it. Fed to the model
as-is it produced **negative loss-to-lease and a Year-1 revenue 14% BELOW the
trailing actual** — the model deducts vacancy from a GPR that is already net.

Restate the Market Rent column to the **current posted schedule** and leave
Contractual / Net Effective untouched. Corroborate it three ways before you do
(all three agreed here): the **vacant units' own** market rents, the market
rents on **leases signed in the last 12 months**, and the `Missing MARKET RENT`
house-rule highest-repeatable calculation. Texarkana: 1/1 $675 · 2/1 $725 ·
2/1.5 $825 · 2/2 $800 → GPR **$559,500** against the stale column's $499,440
(+12.0%), and loss-to-lease falls out at a sane **2.68%**.

Keep the restatement in the populate script as data (a `market_rent_override`
map), not as a hand edit, and disclose it — it is a change to the owner's
document and a buyer will diff it.

## 3. `Assumptions` G33:G36 say "($/mo)" and feed an ANNUAL total

```
'UW - F&C'!AM27 = IF(Assumptions!$G$33<>"",Assumptions!$G$33,V27)
```

The input drops **straight into the annual expense row**. The "($/mo)" in the
label describes the *reference* columns `C33`/`J33`, which are $/unit/month.
Enter `55` meaning "$55/unit/month" and the model books **$55 of water for the
year**. It is the same class of defect as the Contract Services `44.42` bug in
`vintage-okc-underwriting-8-2026.md` §7, and it is invisible: no error, and the
line is too small to notice next to a $52,000 insurance row.

Write annual dollars: `$/unit/mo x units x 12`. Here $12/$45/$18 per unit per
month → **9,360 / 35,100 / 14,040**, total $58,500 = $900/unit/yr.

**Do NOT "fix" `AM32` the same way.** Management fee reads
`=N(Assumptions!$G$37)`, so a 4% input displays as ~0 in that cell — but the fee
IS applied correctly at 4% of EGI elsewhere in the subtotal. Verify before
touching it: `AM35 - (controllables + utilities + insurance + taxes)` should
equal `G37 x AM19`. It did, to the dollar.

## 4. The `UW - F&C` T-12 block drops a month — write it explicitly

`'UW - F&C'` rows **71–92** are the model's own copy of the trailing statement,
keyed by redIQ code in column A (`r`=71 … `tx`=92), columns **F:Q** = the twelve
months. The template stores them as literals and something in the LibreOffice
round-trip repopulates them — **with one month blank**. Here May-25 came back
empty and the printed T-12 page totalled **$457,048 against the statement's
$497,048**, a clean $40,000 short.

It is not a pivot (the workbook's only pivot is at `AP9:AY12`) and `Final_T_12`
was correct throughout, so the usual `G69 = Final_T_12[date2]` header fix does
not address it. Write the block yourself from the same series that feeds
`Final_T_12`:

```python
UW_ROWS = {"r":71,"ll":72,"v":73,"nr":74,"bd":75,"rw":76,"ro":77,"oi":78,
           "cs":81,"rm":82,"ad":83,"m":84,"pr":85,"w":86,"tr":87,"e":88,
           "o":89,"mf":90,"i":91,"tx":92}
for code, row in UW_ROWS.items():
    for i, v in enumerate(series_for(code)):
        uw.cell(row, 6 + i).value = float(v or 0.0)     # F..Q
```

**Impact check before you panic:** the T-3 window (last 3 months) was unaffected,
so DSCR and the returns did not move — only the printed T-12 page and the T-12
cap rate. Assert on the row total against the statement every run.

## 5. Arkansas does NOT reassess to sale price — and the factor must track the price

This is the single largest expense assumption on the deal and the model's default
gets it wrong by design (`G39 = 1.00`, i.e. assess at 100% of purchase price).

**Amendment 79 caps assessment increases at 10%/yr on non-homestead property and
contains no change-of-ownership reset.** A buyer's taxes do not step up at
closing. The first genuine repricing event is the next countywide reappraisal
(Miller County: reappraised 2025, next **2029**). Substantial *improvements* are
assessed outside the cap, so a heavy rehab does get picked up.

Model it as taxes = millage x the **county's** appraised value:

```
G39 = CAD full market value / purchase price      ->  taxes = rate x CAD value
```

**`G39` therefore has to be re-set at every price you test.** Left fixed while
sweeping, it silently understates taxes as the price falls — on this deal a
factor set at a $3.4M draft price understated taxes by $1,858/yr at the $3.06M
answer and overstated the supportable strike by $30,000. `uw_sweep_price.py`
(contributed with this note) now sets it per step; the tell that it is working
is that **NOI stops varying with price**.

Arkansas millage arithmetic, worth keeping: mills are levied on the **20%**
assessed ratio, so `effective rate on market value = mills/1000 x 0.20`.
Texarkana AR = 38.9 (school) + 10.5 (city) + 8.8 (county) = **58.2 mills =
1.164%**. Enter it in `Master!B11:B17` as **rate per $100 of market value**
(0.778 / 0.210 / 0.176) with `B10 = 1`.

Also: the assessor's own numbers are reachable without actDataScout (which was
down for this whole job) via the **Arkansas GIS Office statewide CAMA parcel
service**, an ArcGIS REST endpoint fed from county CAMA systems. It returned all
five parcels of this complex and their appraised/assessed splits.

## 6. `Agency Region`: rebuild it from the FULL agency workbook when AgencyDrift is thin

`refresh_market_tabs.py --region` filters the CMA's `AgencyDrift` sheet (3,556
rows). For Texarkana that yields **five East-Texas rows averaging 5.834%** — a
terminal cap below what the market actually finances at, in the direction that
overstates value.

The full `Combined Fannie and Freddie Sales Comps.xlsx` (48,138 rows) carries the
same first 26 columns. Build a filtered, CMA-shaped workbook from it and feed
that instead. Arkansas + TMG "East Texas", 1950–2000, 20–500 units, 3 years →
53 rows, **Z40 = 6.1995%**.

Then set the risk factors from **observed local caps, not vibes**. Texarkana's
own 2024 agency originations priced at **7.77%** (Westridge) and **7.67%** (Park
at Summerhill) against that 6.20% regional average — a ~150bp tertiary spread,
which is exactly `Factors!J16` (Tertiary Location, 150bp). Add `J18` (Low Unit
Count, 50bp) at 65 units → **terminal cap 8.25%**. Deliberately skip `J23` (Old
Vintage, 100bp): those two observed caps were on 1984–85 product and already
embed the market's quality discount, so stacking it double-counts.

`YardiProjections` remains hardcoded to Mount Houston inside the script
(`vintage-okc` §4) and **there is no Texarkana file** in
`Databases/- Sales Comps/Yardi/` (18 metros, none of them Texarkana). Relabel it
and set N16/N17 from local evidence; the `N17/10000` term is ~0.9bp on the
terminal cap, so the number is immaterial and only the printed label matters.

## 7. Empty rent comps print `#DIV/0!` on page 1 — wrap, do not fabricate

With `TableRecentLeases` / `TablePropertyData` unpopulated, `'Rent Summary'!H7:M7`
are `AVERAGE()` over an empty range and **`PDF Output - F&C`!H44:M44 print
`#DIV/0!` on the first client-facing page**, next to a correct subject row.
`PDF Output - F&C`!**J44 is different** — it is `=AVERAGE(J40:J43)` over the page
itself, so wrapping only `Rent Summary` leaves one error behind.

Wrap both in `IFERROR(...,"")`. The row then prints blank and **still populates
itself the moment comps are loaded** — which is the right trade against either
shipping an error or inventing a comp set. `uw_finalize_model.py` does this.

Related: the black-box conditional format on `B52:J77` that
`underwriting/SKILL.md` says to delete **is not present in the 8/2026 template**
(the script reports "removed 0"). Empty floor-plan rows rendered pale blue, not
black. Check rather than assume.

## 8. Deal record — Renaissance Square, 8/8/2026

Inputs: rent roll 7/28/2026 (65 units, 58 occupied, 89.23%, 49,400 SF, 760 SF
avg); T-12 Apr-25→Mar-26 (LUMP_INCOME, rental income $497,048, other income
$12,370). Five parcels, combined CAD appraised **$1,588,650**, 58.2 mills.
No subsidy (ADFA Miller County list checked — the property is NOT on it).

Assumptions: GPR $559,500 (restated posted schedule) · LTL 2.7% · vacancy 9% ·
concessions 1% · bad debt 3% · mgmt 4% · agency benchmarks on the six flagged
lines · utilities $900/unit/yr · capex reserve $350/unit · taxes pinned to CAD
value · terminal cap 8.25% · Fannie Mae — Small Balance, 80% LTV, 6.68%, 1 yr IO.

Value-add: 7 rows, capital **$48,750**, NOI **+$86,580/yr** — Light Interior
(33 units); Pet Fees, Fee Income Optimization, Renters Insurance; Water RUBS,
Misc Billback, Utility Bill Audit. Note the **Utility Bill Audit contributed $0**
and that is correct: its expense line uses an override, so the engine skips the
saving by design (SKILL "Savings effects"). Excluded: `C62` reduce-opex-to-comp
(umbrella, and it read **-$159,210**, i.e. the subject is BELOW comp averages),
`C63`/`C64` (negative), `C69` payroll optimization (payroll on an agency
benchmark), and **`C71` HFC/PFC — Texas only, the subject is Arkansas**.

| | |
|---|---|
| **Strike** | **$3,030,000** ($46,615/unit, $61.34/SF) |
| Project IRR / avg CoC / T-3 DSCR | **24.45% / 13.04% / 1.2525** |
| Year-1 NOI / PF cap / YoC | $225,084 / 7.43% / 7.31% |
| Loan / equity | $2,424,000 @ 6.68%, 80% LTV / $727,470 |
| Printed sales range | $2,600,000 / $3,030,000 / $3,200,000 |

**DSCR binds, and the causal chain is worth stating that way**: the loan is
LTV-sized, so debt service scales with price while in-place NOI does not
($3,040,000 → 1.2484, red). IRR had 4.5 points of headroom at the answer.

**The income approach came in BELOW the sale comps and that is the finding.**
The same deal's sale-comp grid (part 1 of this job) indicated **$3,428,008**, and
the model's own internal grid **$3,270,307** — the two differ only in the
cap-rate-drift normalisation basis. Underwriting supports $3,030,000. Quote the
spread rather than hiding it: it is the difference between what the asset earns
and what comparable assets fetched, and on a property with a broken expense
statement that gap is the seller's burden of proof.

## Related

- `uw-model-linux-libreoffice-build-8-2026.md` · `aggressive-pricing-house-rule-8-2026.md`
- `vintage-okc-underwriting-8-2026.md` — the Contract Services units bug, same family as §3
- `sale-comps-tertiary-market-texarkana-8-2026.md` — the comp grid this prices against
- `scripts/uw_sweep_price.py` · `scripts/uw_finalize_model.py`
