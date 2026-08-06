# Combining two properties into one master T-12

How TMG builds a single combined operating statement out of two properties'
P&Ls: which months the combined statement is allowed to cover, how accounts
from two different charts get merged, and how the result is proved. Written
8/6/2026 on Westlake East + Westlake West (59 + 115 doors, two QuickBooks
Online entities). Companion script: `scripts/build_master_t12.py`. Companion
format note: `instructions/quickbooks-online-xlsx-t12-parser.md`.

## 1. Build a combined SOURCE, then run the toolkit on it

Do not hand-assemble a deliverable workbook. Write the combined statement in
the **same source layout the toolkit already parses** (here: the QuickBooks
Online xlsx layout) and then run `process_t12.py` on it exactly like a native
export. That is iron rule #4/#5 doing real work:

- every row total, every section subtotal and every control row of the combined
  statement gets validated by the registered parser, not by the ad-hoc script;
- the deliverable goes through `_save_normalized()` like every other T-12;
- the master carries the same Trailing Financials / hidden Final T-12 /
  Comments geometry as the single-property workbooks, so the model import and
  the client formatting are identical.

The combined source workbook is itself a deliverable (buyers ask how the
combination was made), so run it through `_normalize_xlsx()` too — raw openpyxl
output crashes Excel/JS loaders.

## 2. THE OVERLAPPING-PERIOD RULE

**A combined statement may only cover the months BOTH properties' books
cover.** Everything else is an invented number.

Westlake: West has 12 months (Aug-25 – Jul-26). East's ownership books
(V Westlake LLC) begin 10/1/2025, so East has 10 (Oct-25 – Jul-26). A combined
Aug-25 or Sep-25 column would be West-only — the "12-month total" would then
NOT be a combined run rate, it would be 10 months of the asset plus 2 months of
half the asset, and every T-3/T-1 and per-unit read off it would be wrong.

So the master covers the **10 shared months, Oct 2025 – Jul 2026**, and:

- East is **not** zero-filled (a zero is a claim that nothing was spent);
- West's Aug-25/Sep-25 are **not** carried (they are not combined months);
- nothing is annualized or grossed up.

The master is then processed with `--allow-partial --pad-to-12`, which is the
standing house rule for a short contiguous statement: it displays on the full
trailing-twelve axis ending at the last real month, the missing months get
dated headers but **genuinely empty** cells, the Total column stays the sum of
the real months, and a Comments-tab note names the gap. Add your own
`--header-note` explaining WHY those two months are blank — the reader must not
be left to guess whether it is missing data or a shorter ownership period.

Alternatives, and why they lose:

| Option | Why not |
|---|---|
| Zero-fill East for Aug/Sep-25 | Asserts $0 revenue and $0 expense for a real property. Depresses combined NOI and every per-unit metric. |
| Carry West alone in Aug/Sep-25 | The 12-month total stops being a combined run rate; nobody reading the Total column knows two of its months are half the portfolio. |
| Annualize East's 10 months to 12 | Inventing numbers. Also wrong: East's early months carry onboarding-shaped costs. |
| Deliver only the 10-month master | Correct data, but the model and the client's eye both expect a trailing-12 axis. `--pad-to-12` gives both. |

## 3. MERGE BY ACCOUNT NAME, NEVER BY GL NUMBER

Two entities are two independent charts of accounts and they reuse numbers.
Westlake: `50400` is Software at East and Bank Fees at West; `50610` is
Electricity at East and HVAC Repair at West; `50640` is Trash Removal at East
and General Repairs at West. Merging on the number cross-posts electricity into
HVAC repair and trash into general repairs, and every total still ties, so the
error is invisible.

Method:

1. Run `process_t12.py` on each source alone first. Take the account list it
   prints (and the code each line mapped to) as the input to the merge map.
2. Declare every merge explicitly, by name, in a table: `master line name <-
   [(source, source account name), ...]`. Copy the names verbatim, typos
   included — `50930 Quicbkooks` is a real account name.
3. **Refuse to run if any source account is unaccounted for.** The script
   asserts that every account in either source is either merged, carried
   through, or on an explicit DROP list, and aborts otherwise. This is the
   single check that stops money going missing in a rename.
4. Zero-only parent rows go on the DROP list, and the script asserts they are
   actually zero before dropping them.

### When the two charts disagree about granularity, keep both lines

East reports one `50650 Landscaping & Pest Control` line; West splits
`50550 Pest Control` and `50750 Landscaping`. There is no honest way to merge
them — you would have to split East's line, which is an invented number.
Carrying all three lines preserves every source figure exactly; they all map to
Contract Services (`cs`) anyway, so the category rollup on the Trailing
Financials tab is correct even though the line detail is not perfectly
symmetric. Say so in the delivery notes.

Same principle for merged names: use a clean, GL-free name for a line that
exists at both properties (`Payroll Expenses`, `Water & Sewer`, `Waste
Removal`), and carry a property-specific line through under its ORIGINAL name
(`50400 Bank Fees`, `50650 Landscaping & Pest Control`) so a reader can find it
in the source P&L.

## 4. Do not clean anything up on the way through

A combining script is a tempting place to "fix" a suspect line. Don't.

- A line that looks like a duplicated booking (Westlake East carries both
  `50800 Insurance` at $2,107/mo AND a second `Insurance` at $4,050/mo) stays
  in, in every workbook, until ownership confirms — `--exclude-account` is only
  for lines OWNERSHIP HAS CONFIRMED are not real costs. Merge both into the
  combined Insurance line, and write a Comments note with both amounts, the
  per-unit comparison against the sister property, and an explicit request to
  confirm. Report the NOI sensitivity to the broker in the delivery summary,
  not in the workbook.
- An owner's margin annotation in the source ("double billing in a month" next
  to a trash line) is data-quality information, not an instruction. Carry the
  line unchanged and surface the note.

## 5. Prove it before writing anything

The combining script asserts, and prints, all of:

1. coverage — every source account merged / carried / dropped exactly once;
2. every merged line equals the sum of its named components, month by month
   and in total;
3. the master's `Total for Income` / `Gross Profit` / `Total for Expenses` /
   `Net Operating Income`, computed from the merged detail, equal East + West
   over the shared months, month by month and annually;
4. the master's control rows compared to each source's own PRINTED control
   rows, with any difference reported in full.

Check 4 is the one that earns its keep: Westlake East's printed
`Total for Expenses` omits the `Make Ready Cleaning` line it prints two rows
above it ($1,185.00). The master is built from the DETAIL (house rule: monthly
detail wins), so the master's expenses are $1,185.00 higher than East's printed
subtotal plus West's — and that difference is stated, not buried. East's own
workbook needs `--trust-monthly` for the same reason.

Then run the master through `process_t12.py` and confirm its reconciliation
block ties. Finally re-open every delivered workbook with openpyxl and check
the month headers, the genuinely blank padded columns, the hidden Final T-12
tab, the Comments tab, and Total = sum of months on every line.

Note on that last check: the writer stores each month as `round(v, 2)` and the
Total as `round(sum of the exact values, 2)`, so a line the source posts in
fractional cents (West's `50325 - Software` is 760.51/12 = 63.37583333/mo)
shows a few cents of pure display rounding. The correct bound is half a cent
per written month cell plus half a cent for the Total — not a hand-picked
tolerance. Anything past that is a real error.

## Westlake result (8/6/2026)

| | months | Revenue | Op Ex | NOI | OpEx % |
|---|---|---|---|---|---|
| Master (174 doors) | 10 (Oct-25–Jul-26) | 1,522,409.99 | 727,055.29 | 795,354.70 | 47.76% |
| West (115 doors) | 12 (Aug-25–Jul-26) | 1,165,590.10 | 529,080.33 | 636,509.77 | 45.39% |
| East (59 doors) | 10 (Oct-25–Jul-26) | 552,255.84 | 298,464.95 | 253,790.89 | 54.04% |

All three reconciliation blocks tie to their own statement's printed totals.
