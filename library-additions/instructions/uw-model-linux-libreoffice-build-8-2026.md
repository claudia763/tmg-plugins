# Building the TMG underwriting model on Linux with LibreOffice (8/7/2026)

Covers: how to populate, recalculate, solve and export the TMG underwriting model on
the email-cowork server, where there is **no desktop Excel and no COM**. Read this
INSTEAD of `uw-model-windows-com-build-8-2026.md` on the Linux agent box (that file's
cell map and tuning logic still apply — only the mechanics differ). Validated on
Shady Oaks Apartments (19 units, Houston 77016, 8/7/2026).

**Headline: a plain `soffice --convert-to xlsx` recalculates the whole 106k-formula
model in ~20 seconds. The macro-driven route hangs forever.** If you take one thing
from this file, take that.

## 1. Three blockers, all fatal, all silent

The model will not recalculate on this box until all three are cleared. Each one
presents as "LibreOffice is running but nothing happens" — `soffice.bin` sits at
**0% CPU with `wchan = futex_do_wait`**, and because it is headless there is no
dialog to answer. Diagnose with:

```bash
ps -eo pid,etime,time,pcpu,args | grep '[s]office.bin'   # CPU TIME must keep climbing
cat /proc/$(pgrep -x soffice.bin | head -1)/wchan        # futex_do_wait => blocked
```

A run that burns ~14 s of CPU and then flatlines is hung, not slow. Kill it by PID.

### Blocker 1 — leftover `externalLink` parts

`lo_model.py prep` (technical-notes §1–2) deletes the external *defined names* and
rewrites external-link *formulas* to their cached values, but the workbook still
contains **8 `xl/externalLinks/externalLink*.xml` parts** plus their rels and an
`<externalReferences>` element in `workbook.xml`. LibreOffice tries to resolve them
against a template copy that does not exist here and blocks. Excel never hangs on
this because it silently falls back to the cached values.

Fix: `scripts/strip_external_links.py` (zip-level; removes the parts, the
relationships, the `[Content_Types].xml` overrides and the `<externalReferences>`
element, copying every other part byte for byte).

### Blocker 2 — `=WEBSERVICE()` FRED pulls

`'Treasury Yields'!D4:D8` are `=IFERROR(_xludf.webservice(C4),"")` live pulls from
fred.stlouisfed.org. LibreOffice evaluates them during a full recalc and blocks on
the socket.

Fix: `scripts/fix_treasury_yields.py`. Write literal `""` into D4:D8 — this is the
tab's **own documented fallback**, since `G4:G8 = IF(ISNUMBER(E),E,F)` then takes the
yellow manual-override column F. Set F from the `loan-terms-lookup` run and stamp the
as-of date in H. You are using the model as designed, not defeating it.

### Blocker 3 — the macro/throwaway-profile route deadlocks

`lo_model.py`'s `recalc` / `read` / `pdf` subcommands install a Basic macro into a
throwaway `-env:UserInstallation` profile and drive the document from it. **On this
box that deadlocks**, and it deadlocks *reliably* when a second LibreOffice instance
is alive — two agents doing LO work at once will both hang. Observed independently by
two agents on the same day; it is very likely what the OOM-killed 8/6 session actually
hit, before the memory pressure.

Fix: don't use it for recalc. See §2.

## 2. The recipe that works

The template ships `<calcPr fullCalcOnLoad="1"/>`, and openpyxl preserves it — so
simply **converting the file recalculates it**. No macro, no profile, no UNO bridge.

```bash
python3 lo_model.py prep      template.xlsx prepped.xlsx     # external names/formulas
python3 fix_treasury_yields.py  prepped.xlsx  t1.xlsx        # blocker 2
python3 refresh_market_tabs.py  t1.xlsx       t2.xlsx --region <Metro>   # §5
python3 strip_external_links.py t2.xlsx       base.xlsx      # blocker 1 — MUST BE LAST
python3 populate_model.py       base.xlsx    populated.xlsx  # openpyxl writes
soffice --headless --norestore -env:UserInstallation=file:///var/tmp/loconv \
        --convert-to xlsx --outdir out populated.xlsx        # ~20 s, full recalc
python3 read_model.py out/populated.xlsx template.xlsx       # values + error diff
```

**Ordering correction (Pointe at Garden Oaks, 8/7/2026) — this supersedes the order
printed in earlier revisions of this file.** `strip_external_links.py` must run
**after every openpyxl step**, not before them. openpyxl *preserves and rewrites* the
`xl/externalLinks/*` parts on save, so stripping first and then running any
openpyxl-based script puts them straight back and Blocker 1 is live again. The
symptom is a base workbook that looks prepped but still carries the parts; verify
rather than assume:

```python
import zipfile
z = zipfile.ZipFile("base.xlsx"); wb = z.read("xl/workbook.xml").decode()
assert not [n for n in z.namelist() if "externalLink" in n]
assert "<externalReferences" not in wb
```

If you populate with openpyxl after stripping (the `populate_model.py` line above
does exactly that), re-run the strip on the populated file too, or simply re-verify
with the snippet and strip again if the assert trips.

Note also that §1's "8 externalLink parts" is really **4 parts plus their 4 `_rels`
files** = 8 zip entries. The strip script reports it the latter way.

Timing on the 8/2026 template (67 sheets, ~106k formulas, 19 MB): **~20 s wall,
~27 s user (it threads), ~870 MB RSS.** Budget one recalc per tuning step — that is
fast enough to solve price *in the spreadsheet*, which the Windows notes assume you
cannot do on Linux.

Give each concurrent run its **own** `-env:UserInstallation` directory, and prefer to
serialise LO work anyway.

**`pkill -f soffice.bin` will kill your own shell** (its command line contains the
pattern). Use `pkill -x soffice.bin` or kill by PID.

## 3. Scripts added alongside this note

| Script | Replaces (Windows) | What it does |
|---|---|---|
| `strip_external_links.py` | — | Blocker 1 |
| `fix_treasury_yields.py` | — | Blocker 2 |
| `refresh_market_tabs.py` | `refresh_agency_region.py` | Repopulates `Agency Region` from a CMA `AgencyDrift` sheet **and** `YardiProjections` from the submarket forecast |
| `solve_price.py` | `model_price_solver.py` | Sets `G50`, converts, reads F5/F7/I8; prints the max-green price |

## 4. Current-template cell map corrections (beyond the Windows note)

The Windows note says the loan block is at `G61`/`M61`. In the 8/2026 template it is
**one row lower again**:

| Cell | Content |
|---|---|
| `G62` | Loan Type — `=IF('UW - F&C'!AB13<$K$62,"Bridge — Debt Fund",$M$62)`, K62 = 0.75 occupancy |
| **`M62`** | **the loan program name — SHIPS BLANK, must be set**, else the whole debt block is empty |
| `G63` | Recourse Type · `G64` LTV · `G65` Interest Rate · `G66` DSCR Requirement · `G67` IO yrs · `G68` term months · `G69` amortization (literal 360) |
| `G48` | `=IF($G$63="Non-Recourse",$K$48,$M$48)` — **note G63, not G62**. K48 = 0.20, M48 = 0.25, so agency debt targets a **20%** IRR. Read it at runtime. |

Confirmed again on Pointe at Garden Oaks (8/7/2026), with three refinements:

- `G66` is labelled **"DSCR Requirement (reference)"** and `M66` reads *"Info only —
  sizing is LTV-driven."* The model does **not** size the loan to DSCR; it sizes to
  LTV and reports the resulting DSCR. So you hit a DSCR target by lowering `G64`
  (LTV), never by editing `G66`.
- `model-map.md` says the `I8` green test is `I8 > G63`. That is **stale** — G63 is
  the Recourse Type *string*. Since the whole block shifted down a row, the DSCR
  threshold is `G66`. Read the conditional-format rule on I8 at runtime rather than
  trusting either document.
- `M62` genuinely has **no `<c>` element at all** in the shipped template (not an
  empty string). Until you set it, `G62` recalcs to `#DIV/0!` and the entire debt
  block is empty — which also means `G48` sees a blank `G63` and resolves to the
  **25%** recourse target. Set `M62` *first*, then re-read `G48`: selecting a
  non-recourse agency program flips the IRR target to **20%**.

`'Loan Terms'!A4:A16` holds 13 programs (verbatim, em dash U+2014): Fannie Mae —
Conventional · Freddie Mac — Conventional · Fannie Mae — Small Balance · Freddie Mac
— SBL · HUD/FHA — 223(f) · HUD/FHA — 221(d)(4) · Life Company (LifeCo) · CMBS /
Conduit · Bank — Balance Sheet · Credit Union · Bridge — Debt Fund · Bridge — Bank ·
Mezzanine / Pref Equity.

Green-test cell provenance (all three addresses in `model-map.md` are correct):
`F5` Project IRR → `'PDF Output - F&C'!E34` · `F7` Avg Cash-on-Cash → `'PDF Output -
F&C'!E33` · `I8` T-3 DSCR → `'UW - F&C'!AC42`.

**The error-diff gate in §7 is meaningless on an unpopulated base.** With no deal
data the `#DIV/0!` cascade stays lit, so a prepped-but-empty workbook diffs at ~9,600
errors against the template's ~9,614 (Pointe: 9,612, with 998 "new" and 1,000
"resolved" — the near-symmetry is the tell that these are the same cells changing
error *string*, ~254 of them `#NAME?` from LibreOffice not implementing `_xlfn.*` /
`_xludf.*`). Only run the gate **after** population, and only care about the
`PDF Output - F&C` print area.

Valid `M62` values are the `Loan Terms`!A4:A16 program names — copy the string out of
the sheet, don't retype it (em dash). LTV and rate are INDEXed from `Loan Terms`, so
picking the program is usually enough; overwrite `G64` with a literal only if you
genuinely need an LTV the program doesn't offer.

Also: `Factors` row 23 "Old Vintage (1980s or older)" is **100 bps** in this template,
not the 25 bps in `model-map.md`. `Factors!N16/N17` are decimals, so the `N17/10000`
term in the `G58` build-up is correctly negligible (~0.9 bp).

## 5. The two stale tabs still bite, exactly as on Windows

`Agency Region` ships with a **Dallas-Fort Worth** extract and `YardiProjections` with
a **Little Rock** forecast. Left alone on the Shady Oaks build they produced
`'Agency Loan-Sale Comps'!Z40 = 5.738%` (DFW) instead of 6.093% (Houston) — a 25 bp
error in the terminal cap that flowed straight into a **2.3-point IRR overstatement**,
and the PDF would have printed Dallas properties under Houston search criteria.

`refresh_market_tabs.py --region Houston` fixed both. Cross-check that landed:
Houston n=22 → **6.0932%**, against Werner Creek's independent 8/6 Houston refresh of
**6.093%** and the CMA's own TX/1961-81 average of 6.0965%. Three sources, same
number — that is the check worth running.

## 6. Reconciling Phase-A Python against the recalculated model

The skill wants F5/F7/I8 within ±2% of `tmg valuation.py`. On Shady Oaks the DSCR
matched to 4 decimals and every expense line to the cent, but **Year-1 NOI differed by
$1,043 and IRR by 2.3 points**. The cause is not a transcription error and is worth
knowing before you go hunting:

> **`tmg valuation.py` takes Year-1 GPR from its `T12_MONTHLY["market_rent"]` series.
> The workbook takes it from `Master` — i.e. from the RENT ROLL.** When the rent
> roll's market-rent column and the T-12's market-rent memo row disagree, the two
> engines disagree by exactly that amount, grossed up.

Shady Oaks: rent roll $16,781/mo vs statement memo $16,681/mo — $1,200/yr, which after
loss-to-lease/vacancy/bad-debt is the $1,043 of NOI. Feed the **rent-roll** figure into
`T12_MONTHLY["market_rent"]` if you want the two to agree, and reconcile the owner's
two documents in the delivery notes either way. Solve the final strike in the workbook
(§2) — it is the deliverable, and at 20 s a step there is no reason not to.

## 7. Error-diff gate

`read_model.py <recalced> <pristine template>` diffs cached error cells. Expect the
recalculated deliverable to have **fewer** total errors than the template (Shady Oaks:
3,785 vs 9,614) with a few hundred "new" ones concentrated on hidden utility sheets —
`(int-assum)`, `(internal)`, `Financials T-12`, `FinalRR!C5`, `Agency Loan-Sale
Comps!AP2`. These are LibreOffice not implementing `_xlfn.*` / `_xludf.*` functions
Excel does, plus `#DIV/0!` on unused scenario rows. **None are inside the
`PDF Output - F&C` print area — that is the check that matters.** Verify the exported
PDF visually rather than chasing the hidden-sheet count to zero.
