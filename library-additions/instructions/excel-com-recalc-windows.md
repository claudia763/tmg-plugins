# Recalculating the TMG underwriting model with real Excel (Windows COM)

Covers: how to recalc/tune/export the ~106k-formula TMG model on a Windows
machine with desktop Excel (no LibreOffice), and the stale-value strikethrough
trap. Read this before running the `underwriting` skill on Windows — the
skill's technical-notes assume LibreOffice.

## The one trap that will ruin your PDF: stale-value strikethrough

Modern Excel (M365) has **stale value formatting**: under
`Calculation = xlCalculationManual`, any cell whose dependencies changed since
its last calc is rendered WITH STRIKETHROUGH — on screen, in
`ExportAsFixedFormat` PDFs, and the stale state SURVIVES SaveAs, so every
export from that saved file is struck too. Volatile functions (`TODAY()` in
the model's Agency Loan-Sale Comps sheet) re-dirty the chain immediately after
every calculation, so in manual mode there are ALWAYS stale cells.

Neither the file styles, nor conditional formats, nor `Range.Font.Strikethrough`
show anything — the lines exist only in rendering. Diagnosis symptom: exported
PDF shows every (or every recalculated) cell struck through.

**Fix: always switch back to Automatic before exporting or saving:**

```python
xl.Calculation = C.xlCalculationManual
for _ in range(2):                      # full rebuild passes (fast in Excel, ~1s)
    xl.CalculateFullRebuild()
    while xl.CalculationState != 0:     # 0 = xlDone
        time.sleep(0.5)
xl.Calculation = C.xlCalculationAutomatic   # <-- clears stale marks
while xl.CalculationState != 0:
    time.sleep(0.5)
# ... ExportAsFixedFormat / SaveAs only after this point
```

Do the recalc, key-cell reads, PDF export and final SaveAs in ONE COM session
(see `scripts/excel_model_recalc.py`). A file saved while manual/stale needs a
re-open + automatic recalc + re-save to be cleaned.

## Other Windows/COM notes for the TMG model

- Excel recalcs the whole model in ~1s per pass (vs ~2 min in LibreOffice);
  the openpyxl-edit -> COM-recalc -> read `data_only=True` loop from the skill
  works unchanged otherwise.
- Single-sheet PDF: `ws.PageSetup.PrintArea = "$B$1:$O$333"`, Orientation=2,
  `Zoom=False`, `FitToPagesWide=1`, `FitToPagesTall=False`, then
  `ws.ExportAsFixedFormat(0, path)`. No LibreOffice macro needed.
- `Range.Address(False, False)` fails with win32com EnsureDispatch
  (`'str' object is not callable`) — `Address` binds as a property; use
  `rng.GetAddress(False, False)` or just `rng.Address`.
- Error-diff check (zero-new-errors gate): scanning cached values with
  openpyxl `data_only=True` for `#REF!/#VALUE!/#DIV/0!/#N/A/...` strings is a
  reliable cross-file diff and avoids the COM SpecialCells quirks.
- **openpyxl mangles every chart in the model** (16 charts): it flips
  `<c:roundedCorners>` to 1, drops the chart-area `<c:spPr>` (transparent
  fill / no border) and injects solid dash props — charts render with a grey
  rounded box and lose their gallery style in Excel. Fix: after all openpyxl
  edits, ZIP-replace `xl/charts/chartN.xml` with the pristine template's
  parts (series ranges are identical; Excel refreshes cached values on the
  next full recalc). See `restore_charts.py` pattern in the St Nicholas notes.
- The PDF sheet's comp-map frame is `I121:O147` — insert the generated map
  picture at exactly that range's Left/Top/Width/Height (LockAspectRatio=0)
  so it fills the frame snugly; render the PNG at the frame's aspect ratio
  first (~1530x1182 px for 527x407 pt).
- Zip codes: write `Master!D3` (and comp-table zips) as INTEGERS. A string zip
  breaks `Agency Loan-Sale Comps` `QueryRegion` (INDEX/MATCH against a numeric
  zip->region table) and prints `#N/A` in the PDF's search-criteria block.
- The model's `Comparable Grid` (PDF "Sale Comparable Summary" page) pulls
  comps from helper columns `N3:AE52` on the same sheet (`AE` marks
  x/xx/xxx/xxxx/xxxxx select comps 1-5). To refresh it with a new deal's sale
  comps, overwrite that block with literal values (source, name, $/unit,
  street, city, state, zip int, units, distance, avg SF, year, price,
  date) + marks, and set drift bps literals in `E25:I25` with adjustment
  values in `E26:I26` (= (cap - bps/10000)/cap - 1, cap = CMA current avg).
  The stale prior-deal map picture on `PDF Output - F&C` (anchored ~row 120,
  col 8) should be deleted and can be replaced with a generated map — see
  `instructions/comp-map-generation.md`.
