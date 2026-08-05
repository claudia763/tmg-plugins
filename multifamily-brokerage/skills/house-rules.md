# House rules and judgment protocols

Extracted near-verbatim from TMG's toolkit instructions (CLAUDE.md, built
July 2026 with Dmytro; validated on Harvest Moon and Renaissance Square).
Read this before making any judgment call.

## NER and concessions

- NER = rent + recurring concessions + discounts; **upfront concessions are
  reported but excluded from NER** (house rule).
- Unknown negative charges: |x| > $200 -> upfront concession, else recurring.
- `ADMINUNIT`/`EMPLCRED` (OneSite) are recurring credits -> Emp./Other
  Discounts (col O, inside NER); `^ADMIN\s*UNIT$` is in DISCOUNT_CODES so
  the unknown-negative fallback can't park them in Upfront Concessions.

## HAP / voucher house ruling (Dmytro, 7/31/2026)

`SUBSIDY_CODES` (`subsidy`, `hap`, `section 8`, `voucher`, ...) classify as
**rent**: Contractual Rent and NER are the FULL contract rent, tenant portion
plus subsidy — a voucher unit is not a discounted unit, the rent just has two
payors. `Resident.tenant_rent` / `.subsidy_charge` keep the split for
reporting. Units carrying a subsidy charge are labelled **"Section 8
Voucher"** in `UnitRecord.lease_type` -> Lease Type (col F), which flows to
the Floor Plan tab; the floor-plan code (`tm1x1a`) is deliberately left
intact so plan rollups and sqft still work. Floor Plan Summary col B only
shows a lease type when every unit in the plan shares it. (The Meadows
7/31/2026: 45 of 124 occupied units, tenant $15,828 + HAP $44,866.)

## MTM is never inferred

MTM is NOT inferred from expired leases (open decision — ask Dmytro if
needed). MTM comes only from an explicit month-to-month fee/charge (OneSite
MTOM fee; ResMan "Month to Month Fee"). Buildium's "Rent Cycle: Monthly" is
the BILLING cycle, not a month-to-month lease. Expired lease dates are never
read as MTM.

## Missing sqft or bed/bath (protocol; step 3 updated 7/2026 per Dmytro)

When the source rent roll lacks unit square footages and/or bed/bath counts:

1. **Search public sources first** (WebSearch/WebFetch): the property's own
   website / floor-plan page, apartments.com / Zillow / Apartment List
   listings, and the county appraisal district (CAD) record for the parcel.
   Older/small assets often only have CAD data — building total sqft ÷ units
   is an acceptable per-unit average if floor-plan-level numbers can't be
   found; note it as such.
2. **If found:** fill the values via `--sqft` / `--bedbath` (map floor
   plan -> sqft / bed / bath), state the source and numbers plainly in the
   summary so Dmytro can correct them, and mark them as sourced-from-web,
   not from the rent roll.
3. **If not found: do NOT ship blanks (7/2026 rule).** After searching
   public sources, fill the BEST ESTIMATE and mark it — `--sqft-est` /
   `--bedbath-est` take the same syntax as `--sqft`/`--bedbath` (keys may
   be floor plans or literal unit names) but highlight every cell they fill
   in red (fill FFC7CE, font 9C0006) and add one red note under the data:
   "highlighted SF, bed, bath counts are best estimates, not provided by
   ownership". `--bedbath-est` marks per component, so a parser-derived bed
   count stays black while the bath count the source could not provide is
   flagged (Lofts at Taft), and a unit with no floor plan at all gets one
   synthesised from the estimate so it joins the Floor Plan rollups instead
   of dropping out (Gardens 5-101). Plain `--sqft`/`--bedbath` remain the
   flags for values that DO have a citable source. In an attended session
   with low confidence, ask the user for the inputs instead.
4. Never invent numbers without a source, and never let a web-sourced
   figure silently override a value that IS present in the rent roll.

Bed/bath is NOT in a OneSite roll and A1/B1 codes don't encode it (the
generic fallback guesses A2=2/2, B1=1/1 — both wrong at Synott); bath is not
in a ResMan rent roll (bed comes from the plan code). Use `--bedbath` with a
cited source.

## Missing MARKET RENT (house rule added 8/2026, per Dmytro)

Some sources carry no Market Rent column at all (AppFolio exports built
without it — Vista Lago; owner sheets; Buildium). Leaving col I blank ships
a market-rent column of blanks and `#DIV/0!` rollups. The house rule is to
**estimate market rents from the maximum stated contractual rents**, opt-in
via **`--estimate-market`** (`estimate_market_rents()`), computed **per
floor plan**:

1. Data points are that plan's **occupied** units with a real contractual
   rent. A $0 rent (employee/model unit) is not a market signal and is
   excluded.
2. **Recency filter**, in order of preference: lease START dates within the
   most recent 6 months of the as-of date; failing that, lease END dates in
   that same window; failing that, all occupied units of the plan.
3. The estimate is the **HIGHEST contractual rent that occurs at least 3
   times** in that set — the highest rent the property can actually repeat,
   not the single best outlier.
4. **Graceful fallback, never silently empty:** if no value reaches 3 data
   points, take the value with the most occurrences (ties -> the highest of
   them) and FLAG it (`FLAG: market-rent estimate thin - ...`). A plan whose
   occupied units carry no rent at all is reported and left blank.

Every filled cell gets the standard red treatment (FFC7CE / 9C0006), and the
note under the data names exactly the fields estimated in that run
(`_est_note_text()`). A market rent the source DID provide is never
overridden — the flag is inert on sources that carry the column — and the
estimate deliberately files **no market-rent reconciliation check** (there
is no report total to tie to; that would be checking the parse against
itself). The full derivation prints per plan: the basis used, the candidate
counts, and the chosen value.

**Vista Lago 7/2026 validated: $1,675.** 40 doors, one 3/2 plan; 4 leases
started inside the 6-month window ($1,675 x2, $1,550 x1, $1,475 x1), no
value reached 3 points, so the fallback took the highest most-repeated
rent — $1,675, matching Dmytro's own figure. (All-occupied would have given
$1,650; the recency window is what makes the number right.)

## Monthly detail wins (`--trust-monthly`)

Every printed annual row total is checked against its monthly sum. If a
hardcoded printed total disagrees with the monthly detail, the run aborts;
with Dmytro's OK use `--trust-monthly` — monthly detail wins, the variance
is printed, and grand totals are adjusted by the known variance for
reconciliation (Clark/Pecan: Application Fee Income printed $3,830 vs
monthly sum $3,732; Dmytro chose monthly detail, 7/2026). SUBTOTAL-ROW and
GRAND-ROW mismatches (broken =SUM ranges, double-counted subtotals,
hardcoded cells) are also flagged; under `--trust-monthly` the grand totals
always equal the monthly detail.

## `--exclude-account` is never silent

For lines OWNERSHIP HAS CONFIRMED are not real costs — typically a
hand-keyed duplicate of another line — the honest fix is to remove them and
say so, loudly, rather than leave a known-bad number in the T-12 or quietly
patch a cell. `--exclude-account` matches account names exactly, **aborts if
a name matches nothing** (refusing to silently exclude nothing), and reports
every removal: account, side, every non-zero monthly value removed, the
removed total, and the stated `--exclude-reason` — printed to the console,
added to the delivery notes, and written as a red note on the Trailing
Financials tab.

## Charge-code house rules (already in corpus + keyword rules)

- Uniforms, telephone/internet/answering service -> `ad`.
- Anything named "management fee" -> `mf`.
- Property insurance -> `i` BUT employee-benefit insurance
  (health/dental/vision/workers comp/ERISA) stays `pr`, and cell-phone
  *allowances* stay `pr`.
- The 7/2026 corpus cleanup standardized ~330 codes; `Vacancy Loss -> v`
  and `Model Units -> nr` (the original Harvest Moon template had these
  swapped — corpus majority is correct).
- Mapping is layered: exact corpus match -> section rules -> keyword
  rules -> fuzzy (>=0.90 auto) -> REVIEW flag (never silent).
- Section rules: `FIXED ADMINISTRATIVE` allows {ad, mf, i, tx} (tested
  before the plain `admin` rule, which would otherwise bury the house-rule
  mf/i/tx codes); the maintenance rule matches `maint`, not just
  `maintenance`, so Yardi's "IN-HOUSE GENERAL MAINT/SUPPLY" is recognised.
- Subtotal -> rev/exp/noi control codes are matched by regex, so "Total
  Expenses" / "Total Operating Expenses" / "Total Income" all land.

## LUMP_INCOME

A revenue account named just Income / Revenue / Gross Income / Collections
is the whole undifferentiated revenue side with no rent/other-income split.
Parked in Rental Income (`r`) and **always** REVIEW-flagged — the split is a
judgement call, never silent.

## Cross-ledger corpus guard

An exact corpus hit whose code sits on the wrong side of the ledger is
rejected (the line falls through to the rules layers) and flagged, whichever
layer then resolves it. Same label, opposite side: "LEGAL FEES" as an
expense vs Yardi's resident-charged legal-fee income; "WASHER/DRYER RENTAL"
as laundry income vs the equipment lease. Honouring it would move money
across the revenue/expense line.

## Fuzzy subset-artifact guard (important)

`_sim` is rapidfuzz `token_set_ratio`, which returns **1.00 whenever one
label's token set is a subset of the other's**. A bare "Income" therefore
scored a perfect match against 'fee income' (oi), 'gas income' (ro) and
'hap income' (r) alike and would have auto-coded The Gardens' entire revenue
line. Two guards: ties are broken by whole-string `SequenceMatcher`
(`_whole`), and a hit is rejected -> default + REVIEW when *every*
top-scoring candidate is a strict superset of the account name **and** those
candidates disagree on the code. Names that are the corpus label *plus
qualifiers* ("Hap Contract Rent" -> `r`) or whose qualified variants all
agree ("Billing Fee" -> `ro`) still map silently.

## Hard-won gotchas (do not regress these)

1. **Output .xlsx must be normalized** — openpyxl writes inline strings with
   no sharedStrings.xml and leaves query-table markers from the template;
   both crash Excel/JS loaders ("getElementsByTagName of null" / "Removed
   Part: External data range"). `process_t12.py`'s `_normalize_xlsx()` +
   `_purge_broken_names()` handle this; any new openpyxl-written deliverable
   derived from these templates needs the same treatment. This now applies
   to rent rolls too: both writers save through `_save_normalized()`, which
   also strips the rent-roll template Table1's `calculatedColumnFormula`
   entries pointing at sheets that don't exist in the deliverable.
2. Templates carry the exact client formatting — never rebuild styling in
   code; fill the template. If formatting preferences change, edit the
   template file, not the script.
3. `Vacancy Loss -> v` and `Model Units -> nr` (see charge-code rules above).
4. Every run must show its reconciliation block tying out. If a check
   mismatches, the mapping or parse is wrong — fix that, never widen
   tolerances.

## Other standing rules

- A T-12 must be twelve months. Report captions lie — Yardi titled a
  six-month export "Statement (12 months)"; the month COLUMNS are the truth.
  Short statements abort with the count and span; `--allow-partial` opts
  into a clearly-labelled partial workbook; `--pad-to-12` shows a contiguous
  short tail on a full trailing-12 axis with genuinely blank missing months
  (never zeros, never annualized) and a red note naming them. The
  Google-Sheets parser is exempt (it reads an explicit reporting period and
  stamps PARTIAL PERIOD on every output).
- Statements that print no Total Revenue / no NOI get them computed and
  marked derived; `reconcile()` **refuses derived rows as tie-out targets**
  (checking the parse against itself) and prints them with `~`. Only the
  printed totals that actually exist tie out — that is the correct, honest
  result, not a failure.
- Below-the-line sections (Debt Service, CapEx, Cash Flow Adjustments —
  anything after printed NOI) are excluded from operations and go to the
  Capex & Misc workbook.
- When a new PM-system format appears, extend the relevant script with a new
  parser rather than one-off processing, and validate against the source's
  own printed totals before delivering.
- Trailing Financials formatting house rules (7/2026): bold category heads
  with blank charge cells; the Total column is always bold; category heads
  with no accounts under them are trimmed. RawData is deleted unless
  `--keep-raw` or a sum-check fails (on failure it is retained for
  inspection).
- Validation gates (script exits non-zero otherwise): parsed totals vs the
  statement's printed Total Revenue / Total Operating Expense / NOI
  (±$0.05); RawData sum-check rows recomputed from written values (±10 per
  row).
