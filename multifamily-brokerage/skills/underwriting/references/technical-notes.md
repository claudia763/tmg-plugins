# Technical Notes — openpyxl + LibreOffice on the TMG model

Work on a copy (`model.xlsx`). Two loads are always needed: default load for
formulas/structure, `data_only=True` for values (only valid after a recalc).

## 1. External defined names (do this FIRST, before any recalc)

The template carries ~68 **sheet-scoped** defined names (mostly scoped to
'Comparable Assets', plus 'periods' on 'UW - Assume Loan') that point at an
external copy of the template (`'[1]Raw Rents'!…`, `'[1]redIQ RR'!…`).
Clean workbook-level equivalents exist for nearly all of them. If left in
place, the recalc script refuses to run (it flags thousands of cells at risk).

Fix: delete every sheet-scoped name whose text matches the external-ref
pattern `\[\d+\]` — formulas then resolve to the workbook-level names:

```python
EXT = re.compile(r"""(?<![\w"\[])'?\[\d+\][^!"\[\]]*'?!""")
for ws in wb.worksheets:
    for n in list(ws.defined_names):
        dn = ws.defined_names[n]
        if isinstance(dn.value, str) and EXT.search(dn.value):
            del ws.defined_names[n]
```

'Comparable Assets' has zero formulas, so this is safe.

## 2. Cells with direct external-link formulas

'Comparable Grid' (~15 cells), 'MFQA' (~5), 'UW - F&C'!C69,
'Agency Loan-Sale Comps'!AP2 contain `[n]Sheet!…` formulas. Before the first
save, replace each with its cached value read from the ORIGINAL upload with
`data_only=True` (pre-existing `#REF!` caches → 0).

## 3. Legacy array formulas with fixed spill refs

`Raw Rents` W4/X4/…/AH4 (`=TableRecentLeases[col]`, refs like W4:W460) and
`FinalRR` BA9:BZ9 (`=Final_RR_Floor_Plan[col]`) are CSE arrays sized to the
PREVIOUS deal's row count. LibreOffice will not resize them. Replace the
whole spill columns with literal values and clear the stale tail.

When a single-cell formula must reference a one-row table column (e.g.
`=Final_T_12[date2]`), write it as `ArrayFormula('G69', '=Final_T_12[date2]')`
— as a plain formula LibreOffice returns #VALUE!.

## 4. Recalculation loop

Use the xlsx skill's `scripts/recalc.py` with a long timeout (~570 s; the
model has ~106k formulas and takes ~2 min per pass). `errors_found` is normal:
the template ships with ~380 pre-existing error cells (Validation, Tax
Assessment, Yields, unused Value-Add IRR rows, Master helper columns). After
the first recalc, diff the error list against the original file's cached
values and require **zero new** errors. Tuning loop = write Assumptions
inputs → recalc → read F5/F7/I8 with `data_only=True` → adjust price/LTV.

## 5. Table resizing

`ws.tables['Name'].ref = 'A1:V861'` after writing rows; clear all old data
rows first (write `None`, don't delete rows — formulas reference fixed rows).

## 6. Single-sheet PDF export

`soffice --convert-to pdf` exports EVERY sheet, even hidden ones. Use the
bundled `scripts/export_pdf.py`, which bootstraps a LibreOffice profile,
installs a Basic macro, and calls `storeToURL` with `FilterData:Selection =
sheet range B1:O333` — that exports only the PDF Output - F&C pages.
Escape `&` as `&amp;` inside the macro XML ("PDF Output - F&amp;C").

## 7. Misc

- `pip` needs `--break-system-packages`; openpyxl/pandas preinstalled.
- Warnings about Data Validation / Conditional Formatting extensions on load
  are harmless — suppress them.
- Number stored as fraction: percentages are decimals (0.25 = 25%).
- The subject occupancy cell on the PDF ('Rent Summary'!J8*100 with `0#\%`
  format) intentionally multiplies by 100 — not a bug.
- openpyxl re-save shrinks the file several MB (drops unused caches) — fine.
- Black-box conditional format: rebuild `ws.conditional_formatting` without
  the `B52:J77` range (openpyxl has no delete API):

```python
from openpyxl.formatting.formatting import ConditionalFormattingList
new = ConditionalFormattingList()
for rng in ws.conditional_formatting:
    if str(rng.sqref) == 'B52:J77':
        continue
    for rule in rng.rules:
        new.add(str(rng.sqref), rule)
ws.conditional_formatting = new
```

## 8. Preserving Power Query / connections / query tables ("patch, don't round-trip")

Any openpyxl or LibreOffice **save** silently strips `customXml/item1.xml`
(the DataMashup part holding all Power Query M code), `xl/connections.xml`,
and every `xl/queryTables/*` part. If the deliverable must keep refresh
functionality, split the work:

1. **Working copy** — full openpyxl + recalc loop exactly as above. Use it
   for tuning, error checks, and the PDF export. Disposable.
2. **Deliverable** — zip-level surgery on a pristine copy of the original:
   edit only the sheet XML parts that need new input values (lxml), update
   the table part refs, and re-zip copying every other part byte-for-byte.
   Never open-and-save the result with openpyxl or LibreOffice.

Mechanics that matter:
- Write strings as inline strings (`t="inlineStr"`) — avoids touching
  sharedStrings.xml.
- Preserve each cell's existing `s` style attribute; for NEW date cells
  append cellXfs entries to styles.xml (numFmtId 14 built-in; custom
  formats get a fresh numFmtId ≥ 164) and reference the new xf indices.
- Keep `<row>`/`<c>` elements sorted; array formulas need
  `<f t="array" ref="...">`.
- Set `fullCalcOnLoad="1"` on `<calcPr>` in workbook.xml — the pristine
  copy's cached values are stale, so Excel must recalc on first open.
- Set `refreshOnLoad="1"` on the pivotCacheDefinition so the FinalRR pivot
  block refreshes off the new rent roll (its output cells are static
  otherwise and nothing else recomputes them; no formulas consume it).
- Leave the sheet-scoped external defined names and external-link formula
  cells ALONE in the preserved copy — with externalLinks intact they behave
  exactly as the original did in Excel.
- Verify by static parity diff (literal input cells preserved-vs-working
  copy must match; the only acceptable diffs are empty-string spacer cells
  LibreOffice deletes and the pivot output block). Do NOT LibreOffice-recalc
  the preserved file even as a test — LO chokes on the query-table parts
  (observed: a 2 GB temp file) and a save would strip them anyway.

`scripts/surgery_example.py` in this skill is the worked implementation from
the Flats at Shadowglen build.
