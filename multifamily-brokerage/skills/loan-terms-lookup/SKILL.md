---
name: loan-terms-lookup
description: >
  Look up current estimated multifamily debt terms by lender/program — Fannie
  Mae, Freddie Mac (conventional and SBL/Small Balance), HUD/FHA 223(f) and
  221(d)(4), LifeCo, CMBS, bank, credit union, bridge (debt fund or bank), and
  mezz/pref equity — from The Multifamily Group's Estimated Loan Terms
  workbook, with the rate built as current index (UST or SOFR) + spread. Use
  this skill whenever the user asks about current loan terms, debt pricing,
  interest rates, proceeds/LTV, DSCR, interest-only, term, or amortization for
  a multifamily loan — e.g. "what's the current loan terms on bridge debt?",
  "where is agency pricing today?", "what would Freddie SBL quote?", "current
  LifeCo rates", "loan assumptions for new debt" — even if they don't mention
  the workbook. Also use it to refresh the workbook's index-rate overrides or
  when adding/editing lender programs in the workbook.
---

# Loan Terms Lookup

Answers "what are current terms on X debt?" for multifamily financing by
running the **Estimated Loan Terms workbook** bundled at
`assets/Estimated Loan Terms - Multifamily Debt.xlsx`.

The workbook is the source of truth. Its `Loan Terms` tab holds one row per
loan program with the broker-maintained assumptions: index, spread (bps), max
LTV, min DSCR, IO years, term, amortization, rate type, recourse, notes. The
rate is always **index yield + spread** — never a hardcoded number — so the
answer is only as fresh as the index yields you feed it.

If the user supplies a newer copy of the workbook (attached to the chat or in
a connected folder), use that copy instead of the bundled one — the broker
edits spreads there as the market moves.

## Workflow

### 1. Get fresh index yields (best effort, ~30 seconds)

Try the web first; the goal is today's 5/7/10/30-Yr UST and 30-Day Avg SOFR.

- **Treasuries**: WebFetch
  `https://home.treasury.gov/resource-center/data-chart-center/interest-rates/TextView?type=daily_treasury_yield_curve&field_tdr_date_value_month=YYYYMM`
  (substitute current year+month) and take the most recent row's 5, 7, 10,
  and 30 Yr par yields.
- **SOFR**: WebFetch `https://www.bluegamma.io/compounded-rates/sofr` (or the
  NY Fed reference-rates page) for the latest 30-day compounded average SOFR.
- Direct FRED CSV pulls (`fred.stlouisfed.org/graph/fredgraph.csv?id=DGS10`)
  are often blocked in the sandbox — don't burn time retrying them; the
  treasury.gov page above is the reliable path.

If the web is unavailable, skip `--yields` entirely: the script falls back to
the workbook's cached/override rates and reports their as-of date, which you
must then state in your answer.

### 2. Run the query script

```bash
python scripts/query_terms.py \
  --workbook "assets/Estimated Loan Terms - Multifamily Debt.xlsx" \
  --program "bridge" \
  --yields '{"UST 5 Yr": 4.33, "UST 7 Yr": 4.47, "UST 10 Yr": 4.63, "UST 30 Yr": 5.17, "30-Day Avg SOFR": 3.62}' \
  --as-of "08/05/2026"
```

- `--program` is a case-insensitive substring match ("bridge" matches both
  Bridge — Debt Fund and Bridge — Bank; omit it to list every program).
  "agency" won't match anything — use "fannie"/"freddie", or run all and
  present the agency rows.
- Yields are in **percent** (4.63 = 4.63%), keyed exactly as on the
  workbook's `Treasury Yields` tab.
- Add `--set-overrides` when you fetched fresh yields — it writes them into
  the workbook's yellow Manual Override cells so the file stays current for
  Excel use. Add `--json` for structured output.

### 3. Answer the user

Keep it tight — a broker wants the quote card, not a lecture. Lead with the
rate build-up, then the structure, then caveats:

> **Bridge — Debt Fund** (est., indicative)
> Rate: **~6.62% floating** = 30-Day SOFR 3.62% + 300 bps (yields as of 8/5/26)
> Leverage: up to 70% LTV (70–80% LTC) · DSCR ref: 1.00x going-in
> Structure: 36-mo term, full-term IO, 1+1 extensions, non-recourse w/ carve-outs, rate cap required
> *Spreads are placeholder estimates from the terms workbook — confirm with lenders before quoting.*

Always state the as-of date of the yields and that spreads/terms are
broker-maintained estimates, not lender quotes. If the user asks for several
programs or "all options," present a compact comparison table.

## Maintaining the workbook

- **Spreads or structure changed?** Edit the blue cells on the `Loan Terms`
  tab (col D spread, cols F–L terms) — the script and Excel both read them.
- **New lender program?** Insert a row inside the table on `Loan Terms`, copy
  the Index Rate (col C) and Interest Rate (col E) formulas down, and extend
  the `LenderList` named range. The script picks up new rows automatically.
- The workbook also self-updates in Excel for Windows via `=WEBSERVICE()`
  pulls from FRED with manual-override fallback — that path is independent of
  this skill and needs no maintenance here.

## Caveats

- Everything is indicative screening data — never present output as a rate
  lock, quote, or commitment.
- FRED and treasury.gov post yields with up to a 1-business-day lag; SOFR
  averages lag similarly. Say "as of <date>" rather than "right now".
- Mezz/pref row uses a fixed placeholder coupon (no index) — flag that if it
  comes up.
