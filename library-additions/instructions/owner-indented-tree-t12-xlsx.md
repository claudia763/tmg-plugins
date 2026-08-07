# Owner INDENTED-ACCOUNT-TREE T-12 XLSX (the Goldenwrist / Toki dialect)

Covers the owner/PM operating statement whose body is a leading-space indented
account tree under an `Account Name` header row with REAL DATE month cells.
Read this when a T-12 .xlsx looks like:

    row 1   <management company>
    row 2   Properties: <name> - <address>
    row 3   Period Range: <Mon YYYY> to <Mon YYYY>
    row 5   Account Name | 2025-04-01 | ... | 2026-03-01 | Total
    body    "    Income" / "        FEES" / "            Late Fee" ...

Parser: `_is_toki_xlsx` / `parse_t12_toki_xlsx` in `process_t12.py`.
Reference copy: `scripts/owner_indented_tree_t12_parser.py`. First written
8/2026 for Werner Creek (Goldenwrist Investments LLC / Toki Property
Management LLC); **extended 8/2026 for Aldine Apartments (Goldenwrist Capital
LLC), which is the same dialect with a different internal shape.**

## Recognising it — and why there is only ONE parser for it

Detection is narrow on purpose: an `Account Name` header row whose month cells
are real DATE values starting at col B, a Total column, a `Properties:` line
above it, and at least five leading-space-indented body rows. AppFolio also
captions col A "Account Name" and also prints a `Properties:` preamble, but
its month headers are TEXT (`Jul 2025`), so it cannot collide; OneSite uses
date headers but labels accounts `<GL> - <Name>`, which this dialect never
does.

**Do not add a second parser for a new variant of this layout.** The two
variants below print the same header block, the same `Period Range:` caption
and the same indented tree — any predicate separating them would have to key
on body structure, which is exactly the thing that varies. One registered
parser that PROVES which shape it is looking at, month by month, is both
safer and shorter. That is what the 8/2026 extension does.

## The two structural variants seen so far

|  | Werner Creek | Aldine |
|---|---|---|
| memo/GPR row | `Market Rent`, printed ABOVE the first section header | `MARKET RENT TOTAL`, printed INSIDE the Income section as the PARENT of `Rent Income` |
| expense blocks | Tax / Insurance / Maintenance printed BELOW `Total Operating Expense` | all printed ABOVE it |
| printed `Total Operating Expense` | a PARTIAL roll-up (UTILITY + Wages only) | the true grand total |
| outer-indent accounts | none | two real maintenance accounts at indent 0 |
| final restatement row | `Net Income` | just `Net` |

## The four traps (all four are what the 8/2026 extension added)

1. **A MEMO ROW CAN BE A PARENT, NOT JUST A ROW ABOVE THE LEDGER.**
   Aldine prints

       Income
           MARKET RENT TOTAL        104,160  ...  (memo / gross potential rent)
               Rent Income           90,410  ...  (the actual ledger line)
           Total RENTS               90,410  ...

   `MARKET RENT TOTAL` is $1,140,960 over the trailing period and is
   **deliberately excluded from the owner's own Total Operating Income**
   (994,550 RENTS + 52,313.71 FEES = 1,046,863.71 uses Rent Income). Map it as
   revenue and the statement does not tie, by exactly the market-rent amount.

   The rule: a valued account row whose IMMEDIATELY FOLLOWING row is a valued,
   deeper-indented, non-`Total` account is a PARENT — a memo candidate, on
   equal footing with the Werner-style row printed above the first section
   header. It is only a CANDIDATE: the revenue grand row is retried with the
   candidates excluded and then included, and the statement itself decides.
   If the owner's total does include them, they are put back into the output
   in their printed position (never dropped on a guess), and the run says so.

2. **A GRAND EXPENSE ROW IS NOT RELIABLY AN INDENT SUBTREE.** Aldine prints
   two real maintenance accounts at the **outer** indent level:

       Annual Material Cost / month          1,900/mo   (20,900 total)
       Other annual maintenance  / month     (entirely empty)
           Total Operating Expense          41,002.17

   "Every unconsumed row indented deeper than me" structurally cannot see an
   indent-0 row, so it under-counts Total Operating Expense by the $20,900 —
   and at Werner the same rule is the *correct* one, because the printed row
   there is a partial roll-up with Tax/Insurance below it. Both readings are
   now offered and the first that ties **month by month** wins:

   1. the indent subtree (Werner);
   2. every unconsumed expense-side row printed above it (Aldine).

   Neither ties ⇒ the run aborts. Any account admitted by reading (2) from an
   indent shallower than the roll-up is named in the console output, so the
   judgement is visible.

   The empty `Other annual maintenance / month` row has no values, so it is
   classified as a section head with nothing under it and **trimmed** by
   `write_workbook` — it is not carried as a row of zeros.

3. **BELOW-NOI RESTATEMENTS ARE NOT ALWAYS CAPTIONED "Total".** Aldine's last
   row is captioned simply `Net` (with a trailing space). The old code only
   examined `Total`-captioned rows after NOI, so `Net` fell through to the
   account branch and was booked as a **$563,086.18 expense** — which then
   made coded expenses equal coded revenue and NOI zero. Everything printed
   after the statement's own NOI row is now matched as a restatement first,
   against `Total (Operating) Income|Revenue`, `Total (Operating) Expense` and
   `Net|Net Income|Net Profit|Net Loss`, verified month by month against its
   operating counterpart and then dropped. An unrecognised valued row after
   NOI is a hard failure, not a silent expense.

4. **DON'T HARD-CODE ONE VARIANT'S STORY IN THE NOTES.** The Comments-tab note
   about `Total Operating Expense` is now chosen by comparing the printed row
   with the canonical total (printed Total Operating Income less printed NOI),
   instead of always asserting Werner's "rolls up UTILITY + Wages only".

## Empty month columns

A dated month column whose every data cell is blank was not reported at all.
It is dropped from the parse — never read as zero — named loudly, and
`--pad-to-12` puts it back on a full trailing-twelve axis as a genuinely blank
column with a Comments-tab note. At Aldine this is Dec-2025, which exists in
neither of the owner's two statements; see
`instructions/combining-two-property-t12s.md` for the overlapping-period
discipline that produces such a gap, and always add a `--header-note` saying
WHY, so nobody has to guess whether it is missing data.

## Mapping-engine fix that shipped with this

    (re.compile(r"payroll|benefits|wages?\b", re.I), {"pr"}),

Owner books head the payroll block **"Wages"**, not "Payroll". Without
`wages` in `SECTION_ALLOWED` the section carries no constraint at all, and
`Temp. Contractor (1099)` lands in **Repair & Maintenance** — because the
`a/?c` alternative in `KEYWORD_RULES` matches the "ac" inside "Contr**ac**tor".
Same shape as the existing `FIXED ADMINISTRATIVE` and `Taxes & Insurance`
precedents.

## Mapping calls this dialect produces (Aldine 8/2026)

- **`Insurance Services` appears on BOTH sides of the ledger** — as FEES
  income (resident-billed insurance programme, $15,658.16) and as an expense
  (the property's own policy, $75,944.00 level at $6,904/mo). This is the
  textbook **cross-ledger corpus guard** case: the corpus holds
  `Insurance Services -> oi` (29 occurrences), which is right on the income
  side and wrong on the expense side. The guard rejects the exact hit on the
  expense line, drops it to the keyword layer (`insurance -> i`) and REVIEW-
  flags it. Confirm in the run output that the guard actually fired — the
  method string reads `keyword [corpus said 'oi' - wrong side of ledger,
  rejected - REVIEW]`. If it silently mapped, $75,944 of expense became
  income and NOI is $151,888 too high.
- `Internet/Phone` under a UTILITY section -> `ad` (house rule: telephone /
  internet / answering service are Administrative). Expect utilities to read
  low by that amount versus the owner's own UTILITY subtotal.
- `Trash-Out Income` fuzzy-matches `trash income` -> `rt` **silently**, via
  the keyword layer, with no REVIEW flag. At Aldine the line posts twice in
  eleven months ($350, $550), which is a move-out billback like its sibling
  `Move-Out Cleaning Income` (-> `oi`), not a monthly utility reimbursement.
  Check the posting PATTERN before accepting it: a RUBS reimbursement recurs
  every month. Note that `rt` lines render under the `ro` category head on the
  Trailing Financials tab, so the money is not lost either way — only the
  revenue classification moves.

## Verification layers (all abort unless `--trust-monthly`)

Every printed row Total vs its own monthly cells; every roll-up vs the detail
it consumes, per month; the revenue grand vs all unconsumed income detail
(with and without memo candidates); the expense grand under both readings;
NOI vs revenue less every expense account; the canonical Total Operating
Expenses vs the sum of every expense account; and every below-NOI restatement
vs its operating counterpart. Aldine: 38 printed row totals + 10 structural
checks, all tying.
