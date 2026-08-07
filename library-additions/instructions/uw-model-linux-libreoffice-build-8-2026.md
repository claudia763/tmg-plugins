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
| `refresh_sale_comps.py` | — | Loads this deal's comps into `Auto Sales` / `Output_Analysis_Data__2` and sets the include marks — s.5a |
| `swap_workbook_image.py` | COM `Shapes.AddPicture` | Swaps one `xl/media/imageN.png` at zip level, keeping its anchor and frame — s.5a |

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

## 5. The stale tabs still bite, exactly as on Windows — and there are FOUR, not two

`Agency Region` ships with a **Dallas-Fort Worth** extract and `YardiProjections` with
a **Little Rock** forecast. Left alone on the Shady Oaks build they produced
`'Agency Loan-Sale Comps'!Z40 = 5.738%` (DFW) instead of 6.093% (Houston) — a 25 bp
error in the terminal cap that flowed straight into a **2.3-point IRR overstatement**,
and the PDF would have printed Dallas properties under Houston search criteria.

`refresh_market_tabs.py --region Houston` fixed both. Cross-check that landed:
Houston n=22 → **6.0932%**, against Werner Creek's independent 8/6 Houston refresh of
**6.093%** and the CMA's own TX/1961-81 average of 6.0965%. Three sources, same
number — that is the check worth running.

### 5a. The sale-comp grid and its map are ALSO stale (Shady Oaks, 8/7/2026)

Two more, found only by rendering the PDF and reading page 4:

1. **`Auto Sales` — the sale comparables.** `PDF Output - F&C` rows 118-147 mirror
   `Comparable Grid`!C7:I35, whose N3:Z52 INDEX into the table
   **`Output_Analysis_Data__2`** on the `Auto Sales` sheet (A1:Y286) whenever
   `Comparable Grid`!M1 = "Automatic" — how it ships. That table arrives **preloaded
   with a previous deal's CMA export** (285 Dallas/Irving rows tagged `7.xlsx`). Shady
   Oaks printed five Irving comps and *"Subject Indicated Total Value $2,030,830"*
   against a real income value of $1,020,000.
2. **The map picture.** `xl/media/image5.png`, anchored col 8 / row 120 (the I121
   panel) on drawing5. It is a **static image with no formula behind it**, so fixing
   the grid does not touch it — the page then shows correct Houston comps beside a
   Dallas map. Regenerate per `comp-map-generation.md` and swap it with
   `swap_workbook_image.py` (zip-level; keeps the anchor, so it lands in the same frame
   at the same size — the COM `AddPicture` recipe in that note has no Linux equivalent).

Fix the data with `refresh_sale_comps.py --cma "<this deal's Automatic CMA Analysis.xlsx>"`.
Selection is `Comparable Grid`!L3:L52 ("Include (x)"); **grid row r reads sheet row r-1**,
and only the first five marks print.

**The offset that will bite you:** the model's table has an extra empty column
`Column1` at L that the CMA export lacks. CMA A..K → model A..K, CMA L..X → model
**M..Y**. Copy column-for-column and Avg Unit SF lands in the Column1 slot and shears
everything after it — no error, just quietly wrong comps. The script asserts the headers
line up before writing.

> **Why none of this is caught by the gates.** There is no error cell, no `#REF!`, no
> blank. s.7's print-area error gate returns **0 errors**, `finalize.py` reports **ALL
> GREEN**, every return metric is right — and the page is still about the wrong city.
> Sale comps do not feed the income model, so nothing downstream moves. **Render the PDF
> and read all eight pages every time.** The gate proves the model computes; only your
> eyes prove it is about this deal.

### 5b. `Agency Loan-Sale Comps`!C16:H17 keeps two Irving rows — harmless, leave them

After `refresh_market_tabs.py`, two stale records (Towne Oaks Townhomes, Grove at
Irving) survive at C16:H17 under a header at C15. **Page 5 does not read them** — the
printed block sources `Agency Loan-Sale Comps`!R15:AD… (columns R onward), which the
refresh populates correctly; verified 22 Houston rows on the rendered page. Likewise
`Cap Rates` and `Recent Sales Experience` legitimately contain Irving properties (a
national reference list and TMG's own track record). Don't chase these to zero — check
what the print area actually references before deleting anything.

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

### 6a. A SECOND, unrelated divergence: the reversion (Pointe at Garden Oaks, 8/7/2026)

On Pointe the market-rent trap above did **not** apply — rent roll and T-12 memo row
both read $841,200 — and the two engines still disagreed:

| | Python | Workbook | |
|---|---|---|---|
| T-3 DSCR | 1.2836139878631 | 1.2836139878631 | identical to 7 dp |
| Loan / value-add capital / Year-1 NOI | $3,037,500 / $238,000 / $358,483 | same | identical |
| Project IRR | 20.06% | **21.41%** | +6.7% |
| Avg cash-on-cash | 13.36% | **13.93%** | +4.3% |
| Equity multiple | 2.20x | **2.31x** | ~$143k more distributions |

**Diagnose it this way, and in this order** — it takes three cell reads and rules out a
transcription error in about a minute:

1. `'UW - F&C'!AL63` (loan) and `AL50` (value-add capital) — if these tie, the debt and
   capital transcription is right.
2. `'UW - F&C'!AC42` (T-3 DSCR) — a 7-decimal match proves the T-3 block and the whole
   debt stack are identical.
3. `'PDF Output - F&C'!E35` (Yield on Cost) — this is Year-1 NOI ÷ (price + capital), so
   an exact match proves **Year-1 NOI ties** without hunting through the proforma.

If all three tie and only IRR / CoC / equity multiple move, the difference is in
**years 2–5 growth and the reversion**, not in anything you transcribed. Do not go
looking for a typo, and do not "fix" the Python engine to match. The skill's ±2%
reconciliation gate is about catching transcription errors; verified-identical inputs
satisfy it in substance.

**Then re-solve the strike in the workbook, because the workbook is more generous and
the house rule forbids leaving that on the table.** Under
`aggressive-pricing-house-rule-8-2026.md` the target is the maximum price at just-green,
so a workbook IRR of 21.41% against a 20% target is excess to be priced away. On Pointe
this moved the strike from Python's $4,050,000 to the workbook's answer and **flipped
the binding constraint from the IRR floor to the T-3 DSCR floor** — precisely because
DSCR ties across engines while IRR does not, so the DSCR-binding price is unchanged
while the IRR-binding price rises above it.

Practical shortcut: the DSCR-binding price from the Python `max_green_price` /
`solve_price_for_metric("t3_dscr", 1.25)` run is **directly reusable in the workbook**.
Start the workbook binary search just under it rather than sweeping blind.

### 6b. Third confirmation, and skip the binary search entirely (Aldine, 8/7/2026)

Aldine Apartments (96 units, Houston 77039) reproduced §6a exactly — the third deal in
two days, so treat this as the normal case, not an anomaly:

| | Python | Workbook | |
|---|---|---|---|
| Loan / value-add capital | $4,217,400 / $374,400 | same | identical |
| T-3 DSCR (`'UW - F&C'!AC42`) | 1.2594243853574 | 1.25942438535739 | **13 dp** |
| Yield on Cost (`'PDF Output - F&C'!E35`) | 0.0739253524782 | 0.0739253524782196 | identical |
| Project IRR | 20.08% | **24.10%** | +4.0 pts |
| Avg cash-on-cash | 12.26% | **13.85%** | +1.6 pts |

All three §6a diagnostic reads tied, so the divergence was again confined to years 2–5
and the reversion. The Python-solved strike of $6,390,000 (IRR-binding at exactly 20%)
left **4.1 points of IRR** on the table in the deliverable.

**Because DSCR ties to 13 decimals, you do not need a workbook binary search at all.**
Solve the DSCR floor in Python and the answer transfers directly:

```python
lo, hi = python_strike, python_strike + 250_000
while hi - lo > 1000:                      # bisect on the DSCR floor alone
    mid = (lo + hi) // 2
    if run(mid)["t3_dscr"] >= 1.25: lo = mid
    else: hi = mid
```

Aldine: crossing at **$6,426,093**, so **$6,425,000** on $5k steps ($6,430,000 goes red
at 1.24908). One confirming workbook recalc at that price is all you need — down from a
~10-step search at ~20 s each.

**Watch the sign trap when you write this up.** DSCR *falls* as price rises, because the
loan is LTV-sized (66% of price) so debt service scales with price while in-place NOI does
not. That is why the constraint flips: the IRR floor moves *up* on the workbook's more
generous engine, past the DSCR floor, which does not move at all.

It also gives the writeup a better argument than "the IRR target caps the price." The
honest framing is **"LTV sizes the loan, and that is exactly why DSCR caps the price"** —
one causal chain instead of two competing claims, resting on the single metric both engines
agree on to 13 decimals and on a hard lender minimum rather than on a modelled return.
Prefer the DSCR-bound number in client-facing prose for that reason.

**Quote the WORKBOOK's metrics in every client-facing document**, not Python's. At the
re-solved Aldine strike, Python reports an IRR *below* the 20% target while the workbook
reports comfortably above it. The workbook is what the buyer receives; a writeup quoting
the Python figures would contradict the model attached to the same email.

## 7. Error-diff gate

`read_model.py <recalced> <pristine template>` diffs cached error cells. Expect the
recalculated deliverable to have **fewer** total errors than the template (Shady Oaks:
3,785 vs 9,614) with a few hundred "new" ones concentrated on hidden utility sheets —
`(int-assum)`, `(internal)`, `Financials T-12`, `FinalRR!C5`, `Agency Loan-Sale
Comps!AP2`. These are LibreOffice not implementing `_xlfn.*` / `_xludf.*` functions
Excel does, plus `#DIV/0!` on unused scenario rows. **None are inside the
`PDF Output - F&C` print area — that is the check that matters.** Verify the exported
PDF visually rather than chasing the hidden-sheet count to zero.
