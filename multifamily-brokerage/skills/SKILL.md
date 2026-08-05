---
name: rent-roll-t12-processing
description: "Converts multifamily rent rolls and T-12 / trailing-twelve operating statements into TMG's standardized underwriting Excel workbooks. Use whenever the user provides a rent roll or T-12/P&L (PDF or XLSX, from any property-management system — ResMan, Yardi, RealPage OneSite, AppFolio, Buildium, QuickBooks, SSI, Google Sheets, owner-made) and wants it processed, standardized, converted, or reconciled. Triggers: 'process this rent roll', 'process this T-12', 'run the T-12', 'standardize these financials', 'convert this P&L', 'underwriting workbook'."
---

# Rent Roll & T-12 Processing (TMG)

This skill packages The Multifamily Group's production toolkit for turning
raw property-management reports into standardized underwriting deliverables.
Run the scripts — do not hand-build workbooks or re-implement parsing.

## What the toolkit produces

- **Rent roll workbook** — `RR - {Property} - {M-D-YYYY}.xlsx` (as-of date,
  no zero padding). Tabs: Floor Plan, Rent Roll, Floor Plan Summary — all
  summary tabs are live formulas off the Rent Roll tab; rediQ named ranges
  defined for compatibility.
- **T-12 workbook** — `T-12 - {Property} - {Month YYYY}.xlsx` (Month = last
  statement month; corporate suffixes like "Apartments"/"LLC" stripped).
  Tabs: Trailing Financials (visible, client-formatted) and Final T-12
  (hidden Excel table used for model import — never delete it).
- **Capex & Misc workbook** — `Capex & Misc - {Month YYYY}.xlsx`: everything
  below the statement's printed NOI (debt service, capex, non-operating).

## Setup

1. Dependencies: `pip install pdfplumber openpyxl rapidfuzz` (add
   `--break-system-packages` if pip refuses).
2. Copy the toolkit next to your working files before running — the
   templates and `t12_mappings.csv` MUST stay in the same directory as the
   scripts (they auto-discover them):

   ```bash
   cp -r "<this skill's directory>/toolkit" ./toolkit
   ```

3. Work only on the copy. Runs that harvest or extend the mapping corpus
   must NEVER modify the copy inside the installed plugin — the shared
   corpus changes only through a deliberate, user-approved harvest.

## Usage

Rent rolls (source format is auto-detected):

```bash
python3 toolkit/process_rent_roll.py "<rent roll.pdf|.xlsx>" \
    [--property "Name"] [--asof YYYY-MM-DD] \
    [--sqft "F1=750,F2=925"] [--bedbath "F1=1/1,F2=2/2"] \
    [--sqft-est ...] [--bedbath-est ...] [--estimate-market]
```

- `--asof` overrides the as-of date; mandatory when the source prints none
  (the run hard-exits rather than guess).
- `--sqft` / `--bedbath` fill values with a **citable source** (state it in
  the delivery summary). `--sqft-est` / `--bedbath-est` are the estimate
  variants — same syntax, but every filled cell is red-flagged. Keys may be
  floor plans or literal unit names (unit names win).
- `--estimate-market` estimates missing market rents from max stated
  contractual rents (opt-in; cells red-flagged).

T-12s:

```bash
python3 toolkit/process_t12.py "<t12.pdf|.xlsx>" \
    [--template toolkit/t12_processor_template.xlsx] \
    [--trust-monthly] [--allow-partial] [--pad-to-12] \
    [--exclude-account "Name"] [--mappings <corpus.csv>] [--keep-raw]
```

- `--trust-monthly`: when a printed row total disagrees with its monthly
  detail, use only with the user's OK — monthly detail wins and the variance
  is printed.
- `--allow-partial`: opt-in for statements shorter than 12 months (otherwise
  the run aborts). `--pad-to-12` shows a short statement on a full
  trailing-12 axis with genuinely blank missing months (never zeros, never
  annualized) and a red note naming them.
- `--exclude-account`: exclusions are always reported, never silent.
- `--mappings`: a run-local corpus copy for deal-specific mapping overrides
  (the shared `t12_mappings.csv` stays untouched).
- Corpus maintenance: after correcting REVIEW-flagged lines in a finished
  workbook, `python3 toolkit/process_t12.py harvest "<file>.xlsx"`; for bulk
  archives, `python3 toolkit/harvest_t12_corpus.py <directories.xlsx>`
  (column A = deal folder paths; majority vote wins, conflicts reported).

## Iron rules

1. **Every run must show its reconciliation block tying out** against the
   report's own printed totals. If a check mismatches, the parse or mapping
   is wrong — fix that. Never widen tolerances.
2. **Monthly detail wins over printed totals** (with the user's OK via
   `--trust-monthly`), but variances are always surfaced.
3. **Never invent numbers.** Missing sqft, bed/bath, or market rent follow
   the estimate protocol in `references/house-rules.md`: public sources
   first, then estimate flags that red-highlight every filled cell and add a
   note — never silent, never blanks.
4. **New source formats get a new registered parser in the script** —
   subclass/register in `process_rent_roll.py` (`PARSERS`/`XLSX_PARSERS`) or
   `process_t12.py` (`T12_XLSX_PARSERS`/PDF sniffers). Never one-off
   processing; validate against the source's own printed totals before
   delivering.
5. **Deliverable .xlsx must go through the scripts' normalized save paths**
   (`_save_normalized()` → sharedStrings fix + broken-name purge). openpyxl
   output that skips this crashes Excel/JS loaders.
6. **A T-12 must be 12 months** unless the user opts into a partial
   (`--allow-partial` / `--pad-to-12`). The month columns are the truth, not
   the report caption.

## Pointers

- Read `references/house-rules.md` **before making any judgment call**
  (NER, concessions, HAP/vouchers, MTM, estimates, charge-code rulings,
  mapping guards, gotchas).
- Read `references/supported-formats.md` when identifying a source format or
  writing a new parser.
- `references/deals/` holds worked precedents — including specific mapping
  rulings (Pest Control billback, bare "Maintenance", App/Admin Fees
  Concession, Commissions) — check them before re-litigating a call.
