# QuickBooks Online P&L exported to XLSX — registered T-12 parser

What the format is, the traps that make a naive parse silently wrong, and the
full source of the registered parser so a future agent can re-apply it to
`process_t12.py`. Added 8/6/2026 (Westlake East / Westlake West, TMG).

## When you are looking at one

QuickBooks Online's on-screen Profit and Loss sent straight to Excel. Single
sheet (`Sheet1`), no parameter preamble:

```
A1   V Westlake LLC                     <- entity / PROPERTY NAME comes from here
A2   Profit and Loss
A3   October 1, 2025-August 5, 2026     <- a CAPTION. Never read it as the period.
A4   (blank)
A5   ""    Oct 2025  Nov 2025  ...  Jul 2026   Total
A6   Income
A7   Total Net Collections   51228.41  53139.10  ...        552015.84
A8   42000 Application fees   (blank)   (blank)  ...            240.00
A9   Total for Income        ...
A10  Gross Profit            ...
A11  Expenses                                                     0
A12  50100 Payroll Labor     ...
     ...
A29  Total for Expenses      ...
A30  Net Operating Income    ...
A38  Cash Basis Wednesday, August 05, 2026 03:39 PM GMT-05:00
```

Detection (`_is_quickbooks_xlsx`) demands all three of: a col-A cell that is
exactly "Profit and Loss" in the first six rows; a `<Mon> <YYYY>` header row
with a `Total` column to the right of the months; at least one `Total for X`
row below the header. Regression-checked against ResMan trailing P&L, AppFolio,
OneSite, Yardi and the generic owner layout — none of them satisfies all three,
and none of them claims a QBO file. Re-run
`library-additions/scripts/parser_detection_regression.py` after any change.

## TRAP 1 — "Total Net Collections" is an ACCOUNT; "Total for Income" is the subtotal

This is the whole reason the format needs its own parser. QuickBooks' net
collections revenue line is literally called **Total Net Collections**, and the
row that closes the section is **Total for Income**. A parser that keys
subtotals off a bare `^total` gets it exactly inverted: it discards the real
revenue line and promotes the section subtotal to an account, so coded revenue
double-counts everything else in the section.

Westlake West before the fix: coded revenue 1,166,586.98 vs printed
1,165,590.10 — the $996.88 of application fees counted twice, NOI wrong by the
same amount, RawData sum-checks failing by ~2x. The generic owner-sheet parser
also read A3 as the property name ("August, 2025-July, 2026").

The rule: **only `Total for X` closes a section.** `Gross Profit`,
`Net Operating Income`, `Net Income`, `Net Other Income` and `Net Revenue` are
control rows. Everything else carrying money is an account.

Mapping-side companion fix: `LUMP_INCOME` was extended with an optional `net`
so "Total Net Collections" is caught by the house rule for an undifferentiated
lump revenue line (parked in Rental Income `r`, ALWAYS REVIEW-flagged, never
silent). Without it the `bad debt|...|collection|...` keyword rule codes an
entire revenue side as Bad Debt, silently.

## TRAP 2 — GL numbers are not portable between two QuickBooks files

Two QBO entities are two independent charts of accounts and they reuse the same
numbers for different things. Westlake:

| GL    | East                       | West                    |
|-------|----------------------------|-------------------------|
| 50400 | Software                   | Bank Fees               |
| 50610 | Electricity                | HVAC Repair             |
| 50620 | Water and sewer            | (unused; water is 50720)|
| 50640 | Trash Removal              | General Repairs         |
| 50650 | Landscaping + Pest Control | Shipping Container      |

Never merge, compare or map across QBO entities by number — match on the
account NAME. (The toolkit corpus is already safe: `by_gl` is keyed on
`(gl, normalised-name)`, never the number alone.)

## TRAP 3 — empty parent rows, stray marginal cells, trailing footer

- A parent account with no postings of its own prints as a bare label with a
  blank or `0` Total (East: `50600 Utilities`, `50620 Water`; the `Expenses`
  section head itself carries Total=0). Structure, not money — dropped and
  reported. Do NOT let them open a section: East's account list is flat, so a
  "50600 Utilities" section would still be open when
  "50650 Landscaping and Pest Control" arrives and would constrain that line to
  utility codes. A parent that IS closed by its own `Total for X` row does open
  a section; that is decided by a lookahead over the `Total for X` labels,
  never by indentation (this export carries none).
- Owners scribble in the margins. The West file has a bare `7500` in K40 and
  `Implemented AI` in P12; the East file has `double billing in a month` in
  M20. Reading ONLY the month columns and the Total column, and skipping any
  row whose col A is blank, keeps every one of them out of the numbers.
- `Cash Basis <timestamp>` / `Accrual Basis ...` ends the statement block.

## TRAP 4 — the printed section subtotal can itself be broken

Westlake East's own printed `Total for Expenses` omits the `Make Ready
Cleaning` line it prints two rows above it ($985.00 Nov-25 + $200.00 Dec-25 =
$1,185.00). Every individual row ties to its own printed Total; only the
subtotal breaks — and the printed `Net Operating Income` is consistent with the
broken subtotal, so the control-row check passes too. The parser catches it in
the section check and aborts. House-rule resolution: `--trust-monthly` (monthly
detail wins; the variance is printed in full and belongs in the Comments tab).

## Cash-basis quirks to SURFACE, not fix

Same list as the PDF parser: a whole year of insurance or tax posted in one
month (use `--prorate-bulk` only when the concentration test actually fires —
at Westlake both properties post insurance and tax evenly every month, so it
does not), fee/billback lines that begin partway through the trailing period,
onboarding-sized first-month management fees.

## Registration

```python
T12_XLSX_PARSERS = [
    (_is_quickbooks_xlsx, parse_t12_quickbooks_xlsx),
    (_is_resman_trailing_xlsx, parse_t12_resman_trailing_xlsx),
    (_is_appfolio_xlsx, parse_t12_appfolio_xlsx),
    (_is_onesite_xlsx, parse_t12_onesite_xlsx),
    (_is_yardi_xlsx, parse_t12_yardi_xlsx),
]
```

Three supporting edits landed with it:

1. `LUMP_INCOME` gained an optional `net`:

   ```python
   LUMP_INCOME = re.compile(r"^(total\s+)?(net\s+)?(income|revenues?|gross "
                            r"(income|revenues?)|collections?)$", re.I)
   ```

2. `_xlsx_grand_finalize` no longer hardcodes 12 months when it builds the
   canonical Total Revenue / Total Operating Expenses / Total Net Operating
   Income rows. A partial statement shown on a `--pad-to-12` axis has fewer
   display slots than 12 and used to raise `IndexError` in `write_workbook`:

   ```python
   n_mo = next((len(l.values) for l in lines_out
                if l.kind == "account" and l.values), 12)
   for nm, a in canon:
       lines_out.append(Line("subtotal", nm, [a] + [0.0] * (n_mo - 1), "", ""))
   ```

3. `--header-note` became repeatable (`action="append"`) and now feeds the
   plain-black **Comments** tab instead of overwriting the Trailing Financials
   A2 caption — house rule 8/6/2026, no commentary on the send-out tab.

## Validated

- Westlake West, Aug-25..Jul-26, 12 months: Revenue 1,165,590.10 / OpEx
  529,080.33 / NOI 636,509.77 — all three tie to the statement's printed rows.
- Westlake East, Oct-25..Jul-26, 10 months (`--allow-partial --pad-to-12
  --trust-monthly`): Revenue 552,255.84 / OpEx 298,464.95 / NOI 253,790.89.
- A combined master written in this same layout and read back by this parser:
  Revenue 1,522,409.99 / OpEx 727,055.29 / NOI 795,354.70. See
  `instructions/combining-two-property-t12s.md` and
  `scripts/build_master_t12.py`.

## Full source — drop in above `T12_XLSX_PARSERS` in process_t12.py

```python
# ---------------------------------------------------------------------------
# QuickBooks Online "Profit and Loss" exported to XLSX
# ---------------------------------------------------------------------------
# Added for Westlake East / Westlake West 8/6/2026. This is QuickBooks Online's
# on-screen P&L sent straight to Excel - the same report `parse_t12_qbo_pdf`
# handles in PDF form, so the SEMANTICS below are mirrored from that parser;
# only the geometry differs (real cells instead of right-edge-snapped text
# tokens, so no positional snapping is needed here).
#
# Layout:
#   A1        entity / property name        <- the property name comes from HERE
#   A2        "Profit and Loss"
#   A3        period caption ("August, 2025-July, 2026")  <- a CAPTION, not data
#   row 5     month header: "Aug 2025" ... "Jul 2026" + a "Total" column
#   body      section head ("Income"), account rows, "Total for Income",
#             "Gross Profit", "Expenses", account rows, "Total for Expenses",
#             "Net Operating Income"
#   footer    "Cash Basis Wednesday, August 05, 2026 03:39 PM GMT-05:00"
#             ...and, below/right of that, whatever the owner scribbled in the
#             margins (a bare 7500 in K40, "Implemented AI" in P12, "double
#             billing in a month" in M20 of the East file).
#
# TRAPS - both of these cost real money if missed:
#
#  1. "Total Net Collections" is an ACCOUNT, not a subtotal. It is QuickBooks'
#     net-collections revenue line; the row that closes the section is
#     "Total for Income". A parser that keys subtotals off a bare /^total/
#     gets this exactly inverted - it discards the real revenue line and
#     promotes the section subtotal to an account, so the coded revenue
#     double-counts everything else in the section (Westlake West: 1,166,586.98
#     coded vs 1,165,590.10 printed - the $996.88 of application fees counted
#     twice). ONLY "Total for X" closes a section here; "Gross Profit",
#     "Total for Expenses" and "Net Operating Income" are the control rows.
#
#  2. GL NUMBERS ARE NOT PORTABLE between two QuickBooks files. Westlake East
#     calls 50400 "Software" while Westlake West calls 50400 "Bank Fees";
#     50610 is "Electricity" at East and "HVAC Repair" at West. Never merge,
#     compare or map accounts by GL number across QBO entities - match on the
#     account NAME. (The corpus is safe here because `by_gl` is keyed on
#     (gl, normalised-name), never on the number alone.)
#
#  3. Parent accounts with no postings print as a bare label with a 0 or blank
#     Total ("50600 Utilities", "50620 Water" at East, and the "Expenses"
#     section head itself carries Total=0). They are structure, not money:
#     dropped and reported. A parent that IS closed by its own "Total for X"
#     row opens a real section instead - that is decided with a lookahead over
#     the "Total for X" labels, never by indentation (this export carries none).

QBOX_MONTH = re.compile(r"^([A-Za-z]{3,9})\.?[-/. ]+(\d{4})$")
# ledger-side / QuickBooks-standard section heads (a valueless row with one of
# these names opens a section rather than being dropped as an empty parent)
QBOX_SECTION_HEAD = re.compile(
    r"^(income|revenues?|expenses?|cost of goods sold|other income"
    r"|other expenses?|other income and expenses?)$", re.I)
QBOX_FOOTER = re.compile(r"^(cash|accrual)\s+basis\b|^page \d+\b", re.I)

_GRAND_PATS_QBO_XLSX = {
    "rev": r"^total for income$",
    # canonical opex row this parser always emits = printed Total for Expenses
    # (+ printed Total for Cost of Goods Sold when QuickBooks prints one, which
    # its own "Total for Expenses" excludes - see parse_t12_qbo_pdf)
    "exp": r"^total operating expenses$",
    "noi": r"^net operating income$",
    "drop": r"^total for (income|expenses|cost of goods sold)$"
            r"|^total operating expenses$|^gross profit$"
            r"|^net operating income$",
}


def _qbox_txt(v):
    return re.sub(r"\s+", " ", str(v if v is not None else "")).strip()


def _qbox_num(v):
    """Cell -> float, or None when the cell is genuinely blank."""
    if v is None:
        return None
    if isinstance(v, str):
        s = v.strip().replace(",", "").replace("$", "")
        if not s:
            return None
        neg = s.startswith("(") and s.endswith(")")
        try:
            f = float(s.strip("()"))
        except ValueError:
            return None
        return -f if neg else f
    if isinstance(v, datetime):
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _qbox_month(v):
    """'Aug 2025' / 'August 2025' / a real date cell -> month-start datetime."""
    if isinstance(v, datetime):
        return datetime(v.year, v.month, 1)
    m = QBOX_MONTH.match(_qbox_txt(v))
    if not m:
        return None
    for cand in (m.group(1), m.group(1)[:3]):
        for fmt in ("%B", "%b"):
            try:
                return datetime.strptime(
                    f"{cand.title()} {m.group(2)}", f"{fmt} %Y")
            except ValueError:
                pass
    return None


def _qbox_header(rows):
    """-> (header_row_index, [month cols], [months], total col) or Nones.

    The month COLUMNS are the truth (iron rule): the A3 period caption is a
    report parameter and is never read for the period.
    """
    for i, r in enumerate(rows[:15]):
        if not r:
            continue
        got = [(j, _qbox_month(v)) for j, v in enumerate(r) if j >= 1]
        got = [(j, d) for j, d in got if d]
        if len(got) < 2:
            continue
        mcols = [j for j, _ in got]
        tot = None
        for j, v in enumerate(r):
            if j > max(mcols) and _qbox_txt(v).lower() == "total":
                tot = j
                break
        return i, mcols, [d for _, d in got], tot
    return None, [], [], None


def _is_quickbooks_xlsx(path):
    """True for a QuickBooks Online 'Profit and Loss' xlsx export.

    Keyed on the exact report title cell, a '<Mon> <YYYY>' month header row
    with a 'Total' column right of it, and at least one QuickBooks
    'Total for X' subtotal row. Narrow enough that it cannot collide with
    ResMan's trailing P&L ('Trailing Profit And Loss Detail' title, '<Mon>
    <YYYY> Actual' captions, 'Adjusted Total'), AppFolio ('Account Name' in
    col A behind a parameter preamble, no 'Total for X' rows), OneSite
    (MM/DD/YYYY period-end headers), Yardi (no header caption row at all) or
    the generic owner layout (TOTALS column, ALL-CAPS sections, no 'Total for
    X').
    """
    from openpyxl import load_workbook as _lw
    try:
        wb = _lw(path, read_only=True, data_only=True)
    except Exception:
        return False
    try:
        ws = wb[wb.sheetnames[0]]
        rows = [list(r) for r in ws.iter_rows(values_only=True, max_row=120)]
    except Exception:
        return False
    finally:
        try:
            wb.close()
        except Exception:
            pass
    if not any(re.fullmatch(r"profit\s+(and|&)\s+loss",
                            _qbox_txt(r[0] if r else None), re.I)
               for r in rows[:6]):
        return False
    hdr_i, _mc, months, tot = _qbox_header(rows)
    if hdr_i is None or tot is None or len(months) < 2:
        return False
    return any(QBO_TOTAL_FOR.match(_qbox_txt(r[0] if r else None))
               for r in rows[hdr_i + 1:])


def parse_t12_quickbooks_xlsx(path, trust_monthly=False, allow_partial=False):
    """QuickBooks Online P&L xlsx -> (property, months, [Line]).

    Property name comes from A1 (`--property` overrides it). The month columns
    of the header row are the period - the A3 caption is never read as data.
    """
    from openpyxl import load_workbook as _lw

    wb = _lw(path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    try:
        wb.close()
    except Exception:
        pass

    hdr_i, month_cols, months, tot_col = _qbox_header(rows)
    if hdr_i is None:
        sys.exit("ERROR: no '<Mon> <YYYY>' month header row found in the "
                 "QuickBooks Online P&L xlsx.")
    if tot_col is None:
        sys.exit("ERROR: the QuickBooks P&L xlsx has no 'Total' column - "
                 "every printed row total would be unverifiable.")
    for a, b in zip(months, months[1:]):
        if (b.year - a.year) * 12 + (b.month - a.month) != 1:
            sys.exit(f"ERROR: the month columns are not contiguous "
                     f"({a:%b %Y} -> {b:%b %Y}) - refusing to guess the "
                     f"period.")

    prop = _qbox_txt(rows[0][0]) if rows and rows[0] else ""
    caption = _qbox_txt(rows[2][0]) if len(rows) > 2 and rows[2] else ""

    # ---- footer (basis / timestamp) marks the end of the statement block ----
    basis = ""
    for r in rows[hdr_i + 1:]:
        t = _qbox_txt(r[0] if r else None)
        m = re.match(r"^(cash|accrual)\s+basis\b", t, re.I)
        if m:
            basis = m.group(1).title()
            break
    print(f"  QuickBooks Online P&L (xlsx): entity {prop!r} from A1"
          + (f" | {basis} basis" if basis else "")
          + f" | {len(months)} month column(s) {months[0]:%b %Y}-"
            f"{months[-1]:%b %Y}"
          + (f" | A3 caption {caption!r} IGNORED (the month columns are the "
             f"period)" if caption else ""))

    # ---- lookahead: which labels are closed by their own "Total for X" ------
    closed = set()
    for r in rows[hdr_i + 1:]:
        t = _qbox_txt(r[0] if r else None)
        if QBOX_FOOTER.match(t):
            break
        m = QBO_TOTAL_FOR.match(t)
        if m:
            closed.add(m.group(1).strip().lower())

    # ---- body ---------------------------------------------------------------
    lines_out, row_fails, sub_fails, grand_fails = [], [], [], []
    dropped_parents, blank_cells = [], []
    acct_by_section, printed_secs, grand = {}, {}, {}
    stack, side, below, n_checked = [], "inc", False, 0

    for r in rows[hdr_i + 1:]:
        name = _qbox_txt(r[0] if r else None)
        if not name:
            continue                      # blank label = not a statement row
        if QBOX_FOOTER.match(name):
            break                         # basis/timestamp line: end of block
        # ONLY the month columns and the Total column are read, so stray cells
        # outside the statement block (M20 "double billing in a month",
        # P12 "Implemented AI", K40 7500) can never enter the numbers.
        raw = [_qbox_num(r[j]) if j < len(r) else None for j in month_cols]
        annual = _qbox_num(r[tot_col]) if tot_col < len(r) else None
        mvals = [0.0 if v is None else v for v in raw]
        has_vals = any(v is not None for v in raw)
        low = name.lower()

        tf = QBO_TOTAL_FOR.match(name)
        ctrl = QBO_COMPUTED.match(low)
        # a row with no monthly cells and a blank/zero Total is structure
        empty_row = (not has_vals) and (annual is None or abs(annual) <= 0.005)

        if not tf and not ctrl and empty_row:
            if QBOX_SECTION_HEAD.match(low) or low in closed:
                stack.append(name)
                if re.fullmatch(r"income|revenues?", low):
                    side = "inc"
                elif re.fullmatch(r"expenses?|cost of goods sold", low):
                    side = "exp"
                ln = Line("section", name, None, name, side)
                ln.below = below
                lines_out.append(ln)
            else:
                dropped_parents.append(name)
            continue

        kind = ("subtotal" if (tf or ctrl) else "account")
        if annual is not None:
            n_checked += 1
            if abs(sum(mvals) - annual) > 0.05:
                row_fails.append((name, sum(mvals), annual, kind, side))

        # ---- "Total for X": the ONLY row type that closes a section ---------
        if tf:
            sect = tf.group(1).strip()
            printed_secs[sect.lower()] = (mvals, annual)
            while stack and stack[-1].strip().lower() != sect.lower():
                stack.pop()
            parent = stack[-1] if stack else sect
            if stack:
                stack.pop()
            ln = Line("subtotal", name, mvals, parent, side)
            ln.values_annual = annual
            ln.below = below
            lines_out.append(ln)
            if low == "total for income":
                grand["rev"] = (mvals, annual)
                side = "exp"          # everything after income is expense
            elif low == "total for expenses":
                grand["exp"] = (mvals, annual)
            elif low == "total for cost of goods sold":
                grand["cogs"] = (mvals, annual)
            continue

        # ---- Gross Profit / Net Operating Income / Net Income control rows --
        if ctrl:
            ln = Line("subtotal", name, mvals, "", side)
            ln.values_annual = annual
            ln.below = below
            lines_out.append(ln)
            grand[low.replace(" ", "_")] = (mvals, annual)
            if low == "net operating income":
                below = True              # QBO's Other Income/Expenses block
                stack = []
            continue

        # ---- account --------------------------------------------------------
        sect = stack[-1] if stack else ""
        ln = Line("account", name, mvals, sect, side)
        ln.values_annual = annual
        ln.below = below or bool(BELOW_PAT.search(sect or ""))
        lines_out.append(ln)
        if not ln.below:
            acct_by_section.setdefault(sect.lower(), []).append((name, mvals))
        for j, v in enumerate(raw):
            if v is None:
                blank_cells.append(name)

    # ---- canonical Total Operating Expenses = Expenses (+ COGS) -------------
    if "exp" in grand:
        opex = list(grand["exp"][0])
        opex_annual = grand["exp"][1]
        if "cogs" in grand:
            opex = [a + b for a, b in zip(opex, grand["cogs"][0])]
            if opex_annual is not None and grand["cogs"][1] is not None:
                opex_annual += grand["cogs"][1]
            print("  NOTE: QuickBooks prints Cost of Goods Sold above Gross "
                  f"Profit, so its 'Total for Expenses' "
                  f"({sum(grand['exp'][0]):,.2f}) EXCLUDES COGS "
                  f"({sum(grand['cogs'][0]):,.2f}). Canonical Total Operating "
                  f"Expenses = the sum of those two printed rows, "
                  f"{sum(opex):,.2f}.")
        idx = next((i for i, l in enumerate(lines_out)
                    if l.kind == "subtotal"
                    and l.name.strip().lower() == "total for expenses"),
                   len(lines_out) - 1)
        ln = Line("subtotal", "Total Operating Expenses", opex, "", "exp")
        ln.values_annual = (opex_annual if opex_annual is not None
                            else sum(opex))
        lines_out.insert(idx + 1, ln)
        grand["opex"] = (opex, ln.values_annual)

    # ---- verification -------------------------------------------------------
    # (a) every "Total for X" vs the accounts inside section X, month by month
    for sect_low, (svals, _sa) in printed_secs.items():
        kids = acct_by_section.get(sect_low, [])
        if not kids:
            continue
        got = [sum(k[1][x] for k in kids) for x in range(len(months))]
        if not all(abs(a - b) <= 0.05 for a, b in zip(got, svals)):
            sub_fails.append((sect_low, sum(got), sum(svals),
                              [k[0] for k in kids]))

    # (b) the control rows vs the section subtotals they are computed from
    def _mcmp(label, got, want):
        if not all(abs(a - b) <= 0.05 for a, b in zip(got, want)):
            grand_fails.append((label, sum(got), sum(want)))

    if "gross_profit" in grand and "rev" in grand:
        want = list(grand["rev"][0])
        if "cogs" in grand:
            want = [a - b for a, b in zip(want, grand["cogs"][0])]
        _mcmp("Gross Profit = Total for Income"
              + (" - Total for Cost of Goods Sold" if "cogs" in grand else ""),
              grand["gross_profit"][0], want)
    if "noi" not in grand and "net_operating_income" in grand:
        grand["noi"] = grand["net_operating_income"]
    if "noi" in grand and "rev" in grand and "opex" in grand:
        _mcmp("Net Operating Income = Total for Income - Total Operating "
              "Expenses", grand["noi"][0],
              [a - b for a, b in zip(grand["rev"][0], grand["opex"][0])])

    # ---- report -------------------------------------------------------------
    print(f"  QBO row check: {n_checked} row(s) vs the printed 'Total' column"
          + (" - all tie." if not row_fails
             else f" - {len(row_fails)} MISMATCH."))
    print(f"  QBO section check: {len(printed_secs)} 'Total for X' subtotal "
          f"row(s) vs their account detail, month by month"
          + (" - all tie." if not sub_fails
             else f" - {len(sub_fails)} MISMATCH."))
    for nm, got, want, kids in sub_fails:
        print(f"  SECTION-SUBTOTAL MISMATCH Total for {nm}: detail sums to "
              f"{got:,.2f} vs printed {want:,.2f} (variance {got - want:+,.2f})"
              f"; children: {', '.join(kids)}")
    print(f"  QBO control check: Gross Profit and Net Operating Income vs the "
          f"printed section subtotals, month by month"
          + (" - all tie." if not grand_fails
             else f" - {len(grand_fails)} MISMATCH."))
    for lbl, got, want in grand_fails:
        print(f"  CONTROL-ROW MISMATCH {lbl}: printed {got:,.2f} vs computed "
              f"{want:,.2f} (variance {got - want:+,.2f})")
    if dropped_parents:
        print(f"  Dropped {len(dropped_parents)} parent/header row(s) whose "
              f"only value is 0 or blank (structure, not money): "
              + "; ".join(dropped_parents))
    if blank_cells:
        agg = {}
        for nm in blank_cells:
            agg[nm] = agg.get(nm, 0) + 1
        print(f"  {len(blank_cells)} month cell(s) across {len(agg)} row(s) "
              f"print blank in QuickBooks (no transactions posted) and are "
              f"read as $0.00; every affected row's printed Total ties to its "
              f"monthly cells: "
              + "; ".join(f"{nm} [{c} mo]" for nm, c in agg.items()))

    if sub_fails or grand_fails:
        if not trust_monthly:
            sys.exit("ERROR: QuickBooks subtotal/control rows do not tie to "
                     "the detail beneath them - aborting. (re-run with "
                     "--trust-monthly to treat the monthly detail as source "
                     "of truth)")
        print("  --trust-monthly: monthly detail wins over the printed "
              "subtotal/control rows.")

    return prop, months, _xlsx_grand_finalize(
        lines_out, row_fails, trust_monthly, _GRAND_PATS_QBO_XLSX)
```
