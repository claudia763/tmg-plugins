---
name: loan-terms-lookup
description: >
  Look up current estimated multifamily debt terms by lender/program — Fannie
  Mae, Freddie Mac (conventional and SBL/Small Balance), HUD/FHA 223(f) and
  221(d)(4), LifeCo, CMBS, bank, credit union, bridge (debt fund or bank), and
  mezz/pref equity — with the rate built as current index (UST or SOFR) +
  spread. Use this skill whenever the user asks about current loan terms, debt
  pricing, interest rates, proceeds/LTV, DSCR, interest-only, term, or
  amortization for a multifamily loan — e.g. "what's the current loan terms on
  bridge debt?", "where is agency pricing today?", "what would Freddie SBL
  quote?", "current LifeCo rates", "loan assumptions for new debt" — even if
  they don't mention the workbook. Also use it to regenerate the Estimated
  Loan Terms Excel workbook / Loan Terms PDF, or when lender spreads or
  programs need updating.
---

# Loan Terms Lookup

Answers "what are current terms on X debt?" and produces the **Estimated Loan
Terms** deliverables (Excel workbook + one-page PDF) by running one
self-contained Python script:

```bash
python scripts/loan_terms.py --outdir <job folder> [--program "bridge"]
```

Everything is pure Python (`requests` + `openpyxl` + `reportlab`). **Do not
use LibreOffice, `soffice`, or Excel WEBSERVICE anywhere in this workflow** —
the script runs headless on a Linux server, computes all values itself, and
the workbook's formulas recalculate on their own when someone opens the file
in Excel.

## What one run does

1. **Fetches index rates** (UST 5/7/10/30-Yr, 30-Day Avg SOFR) by walking a
   source tree until every index is filled. The server sits on a German IP,
   so EU-reachable mirrors back up the US sources:

   | Priority | UST curve | 30-Day SOFR |
   |---|---|---|
   | 1 | treasury.gov XML feed (official) | NY Fed markets API (official) |
   | 2 | FRED `fredgraph.csv` | FRED `SOFR30DAYAVG` |
   | 3 | Stooq (EU) CSV — no 7-Yr | — |
   | 4 | Yahoo Finance chart API | — |
   | 5 | finanzen.net (DE) scrape, last resort | — |
   | 6 | config fallback (last good fetch) | config fallback |

   A missing 7-Yr is interpolated from the 5/10-Yr. Every value passes a
   sanity check (0.05–15%, with a 10x fix for CBOE-style quotes), and each
   fetched rate is logged with its source. Successful fetches are written
   back into the config as the new fallbacks, so even a fully offline run
   uses the last known-good curve. (Dukascopy is deliberately not in the
   tree: its free feeds are bond *futures prices*, not yields.)

2. **Builds the Excel workbook from scratch** (`Estimated Loan Terms -
   Multifamily Debt.xlsx`): `Loan Assumptions` (dropdown → LTV, Interest
   Rate, DSCR Requirement, IO Years, Loan Term, Amortization via
   INDEX/MATCH), `Loan Terms` (full lookup matrix, rate = index + spread as
   live formulas), `Treasury Yields` (the fetched values with source +
   as-of). Named ranges `LenderList` / `LoanTermsTable` are available for
   other models.

3. **Renders `Loan Terms.pdf`** — a one-page landscape terms sheet of the
   whole matrix in TMG navy/gold, with the rate build-up, as-of line, and
   disclaimer. This is the attachment for emails; the xlsx is for models.

## Answering a terms question

Run with `--program "<query>"` (case-insensitive; all words must appear in
the program name or notes — "bridge", "freddie sbl", "credit union"). Add
`--no-files` when the user only wants an answer, `--json` for structured
output. Then reply with a tight quote card — a broker wants the numbers, not
a lecture:

> **Bridge — Debt Fund** (est., indicative)
> Rate: **~6.62% floating** = 30-Day SOFR 3.62% + 300 bps (NY Fed, 8/5/26)
> Leverage: up to 70% LTV · DSCR ref: 1.00x going-in
> Structure: 36-mo term, full-term IO, non-recourse w/ carve-outs, rate cap required
> *Spreads are broker estimates — confirm with lenders before quoting.*

Always state the as-of date and source of the index, and that spreads/terms
are estimates, not lender quotes. If every index logged as "config fallback",
say so — the answer is only as fresh as the last successful fetch. For "all
options" questions, attach the PDF instead of pasting thirteen cards.

## Maintaining assumptions

`scripts/loan_terms_config.json` is the source of truth:

- **Spreads or terms moved?** Edit the program's entry (spread_bps, max_ltv,
  min_dscr, io_years, term_months, amort_months, notes) and re-run.
- **New lender program?** Add an object to `programs`; the workbook, PDF, and
  lookups pick it up automatically. `amort_months: 0` means full-term IO.
  Programs priced without an index use `"index": "Fixed (no index)"` with the
  whole coupon in `spread_bps`.
- The `indices` block is auto-maintained by the script (`--no-save` to
  disable); only touch its fallbacks when seeding a fresh install.

## Flags

`--offline` (no web), `--no-files` (query only), `--no-save` (don't update
config fallbacks), `--outdir`, `--basename`, `--config`, `--json`.

## Caveats

- Output is indicative screening data — never present it as a rate lock,
  quote, or commitment. HUD rates exclude MIP (~0.60%/0.25% green).
- Official sources post with up to a 1-business-day lag; say "as of <date>".
- The finanzen.net scraper is layout-sensitive by design; when it breaks it
  logs and falls through rather than returning a bad number. Don't "fix" a
  failed fetch by hand-typing rates into the script — put them in the config
  fallbacks with an as-of date instead.
