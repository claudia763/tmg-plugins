# ResMan "Rent Roll (Itemized)" XLSX rent roll — format notes and parser

What it does: documents the ResMan **"Rent Roll (Itemized)"** Excel export and
the `ResManItemizedXlsxParser` written for it (Westlake, 174 doors, as of
7/31/2026). Read this when a rent roll arrives as an .xlsx whose A1 reads
"Rent Roll (Itemized)", or when a rent roll has PHASE groups, colliding
floor-plan codes, or more than 18 distinct floor plans.

**Status: the parser class below is NOT yet in the shipped toolkit.** It must
be merged into `multifamily-brokerage/skills/toolkit/process_rent_roll.py`
(GitHub `claudia763/tmg-plugins`) — see
`instructions/merging-toolkit-changes-to-github.md`. The full source is in this
file so it can be re-applied verbatim.

---

## 1. What the source looks like

Single sheet (`Sheet1`), report-parameter preamble in column A, one header row,
then the detail:

```
A1  Rent Roll (Itemized)
A2  Exported On: 08/05/2026 03:48 PM
A4  Portfolios: Westlake
A5  Units: Active
A6  Include Advertised Rent: Yes
A8  As of: 07/31/2026            <- the as-of date
A11 header row:
    Unit | Unit Type | BD/BA | Status | Sqft | Total | Rent Income |
    Loss/Gain to Market | Concessions | Pet Fee-Non Refundable |
    Electricity Reimbursement | Water Reimbursement Fee | Gas Reimbursement |
    Miscellaneous Income | Liability Fee | Pest Control | Electricity |
    Water | Sewer | Garbage and Recycling
A12 East Phase                   <- phase banner (col A only, rest blank)
A13..A71  59 unit rows
A72 "59 Units" | | | 0.8305084745762712 | 66201 | 57598.10 | 46673 | ...
A73 West Phase
A74..A188 115 unit rows
A189 "115 Units" | | | 0.8173913043478261 | 104250 | 95653.05 | 84708 | ...
```

Everything to the right of `Concessions` is an itemized charge column. The
parser does NOT hard-code that list: any header column outside the fixed set
(`unit, unit type, bd/ba, status, sqft, total, rent income,
loss/gain to market, concessions`) becomes a charge whose CODE is the column
caption verbatim, routed through `classify_charge()`. A portfolio billing a
different set of charges parses unchanged and still ties out column by column.

## 2. Quirks (each one cost a decision)

### Phase groups with colliding floor-plan codes

The detail is split into phases by a bare banner row (column A only, every
other cell empty), each closed by its own `"<N> Units"` subtotal row —
occupancy fraction in the *Status* column, then the column totals.

The two phases **reuse the same plan codes at different square footages**:
`A1 L` is 750 sf in East and 720 sf in West; `B1.5 RWD` is 1,100 sf East and
1,080 sf West (7 codes collide at Westlake). Merging them would create
mixed-sqft plans and a meaningless Floor Plan rollup, so when more than one
phase is present each plan code is namespaced with the phase's initial:
`E-A1 L`, `W-B1.5 RWD`. **Unit numbers are left exactly as printed** (East is
numeric — 102, 703; West is alphanumeric — A102, K304 — they never collide).
With a single phase (or none) nothing is prefixed.

### Column F "Total" INCLUDES column H "Loss/Gain to Market"

Verified row by row and on both subtotal rows: `F = sum(G..T)`, and G..T
includes H. But *Loss/Gain to Market* is a computed rent-vs-market variance,
not a billed item, and in this export it is **not maintained** — populated on
2 of 174 units (E203 and F204, both −49.00) while other units on the same plan
with the identical $749 rent print 0.

Rulings:

* H is **not** booked as a charge (it would fabricate a −$49 "concession" on
  two units and drag their NER).
* H is **not** read as a market rent. The source has no usable market rent at
  all, so market rents come from the `--estimate-market` house rule (8/2026).
* Because ResMan's own Total includes it, the printed Total is tied out as
  **"Total (col F) less Loss/Gain to Market (col H)"** — both numbers taken
  from the report's own subtotal row, so it is still a real tie-out, and the
  whole thing is FLAGged in the run output.
* Every unit's printed Total is checked against its own charge detail + H, and
  a single mismatch **aborts the run before anything is written**.

### No market rent, no tenant names, no lease dates

Nothing to infer from: Lease Start/End, Move-In/Out, Lease Term, MTM, Lease
Type and Vac. Notice all stay blank (MTM is never inferred — house rule).
Because there are no lease dates, `estimate_market_rents()` falls through its
recency filters to the "all occupied units" basis on every plan; expect a lot
of `[THIN]` flags on small plans, which is the honest result.

### Concessions and the $200 split

Column I is the only credit column. Applying the NER house rule
(|x| > $200 → upfront, else recurring) needs care: `RECURRING_CONC_CODES`
matches the substring `CONC` **anywhere** and is tested first in
`classify_charge()`, so a code named "UPFRONT CONCESSION" lands in *recurring*.
The upfront code must therefore contain no "conc" at all — the parser books
`"UPFRONT"` / `"CONCESSION"`. (Westlake: 34 concessions, one of them −393.75 on
unit 803 → Upfront Concessions, col N, outside NER.)

### Vacant rows are genuinely empty

A vacant unit prints Unit / Unit Type / BD/BA / Status / Sqft and then nothing
— blank Total, blank every charge. No $0 charge is invented. `Vacant-Rented`
counts as VACANT in the report's own occupancy fraction (Westlake West:
94/115 = 0.8173913… with the one Vacant-Rented door on the vacant side), so it
maps to apt status `VR`.

### Bed/bath is printed, plan codes lie

`BD/BA` gives "1/1.00" / "2/1.50" → bed/bath verbatim, marked
`bed_bath_explicit` (the generic `^[A-Za-z]*(\d)` fallback would read
"B1.5 RWD" as a 1-bed). Sqft is printed per unit. Neither needs an estimate
flag.

### An occupied unit can carry $0 rent

Westlake unit 104 is `Current` with Rent Income $0 and $925 of Miscellaneous
Income. It stays an occupied door with a $0 contractual rent (and $0 NER); the
market-rent estimator already excludes $0 rents as a market signal.

## 3. Detection and registration

`detect_xlsx()` demands **all** of: the "Rent Roll (Itemized)" title in the
first 15 column-A cells, a `Portfolios:`/`Properties:` line, an `As of:` line,
and the full fixed header set. It cannot claim an AppFolio or Yardi export.

The parser is registered **FIRST** in `XLSX_PARSERS`:

```python
XLSX_PARSERS = [ResManItemizedXlsxParser, AppFolioUnitTypeXlsxParser,
                AppFolioXlsxParser, YardiRentRollXlsxParser]
```

Order alone is not enough. `AppFolioXlsxParser.detect_xlsx()` asked only for
Unit + BD/BA + Status, which this export satisfies — that is why running the
unmodified toolkit on it mis-detected AppFolio and died with
`TypeError: list indices must be integers or slices, not NoneType` (it is
`r[col("tenant")]` with `col("tenant") is None`). The AppFolio detector is
therefore tightened to also require a `tenant` column, which its own `parse()`
reads unconditionally:

```python
                if "unit" in vals and "bd/ba" in vals and "status" in vals \
                        and "tenant" in vals:
                    return True
```

`scripts/parser_detection_regression.py` passes with both changes: 15 files ×
9 detectors, every file claimed by at most one detector on its side.

## 4. Reconciliation (36 checks at Westlake, all tie)

Base checks from `reconcile()`: unit count 174, total sq ft 170,451, occupied
units 143 (from the printed occupancy fractions × the printed unit counts),
total contract rent 131,381, current/on-notice lease charges 153,349.15.

`extra_checks` adds, for **East / West / All phases**: units, sq ft, the
printed occupancy fraction (exact float, tol 1e-9), Rent Income (col G),
Concessions (col I), itemized charges → Other Income, and Total (col F) less
Loss/Gain to Market (col H) — then one check per itemized charge column against
the combined printed subtotals. All are recomputed from the reconstructed
`UnitRecord` list, and the phase selector is the plan-code prefix, so the
checks tie the OUTPUT to the report rather than the parser to itself.

Westlake numbers: East 59 units / 66,201 sf / 0.8305 occ / $46,673 rent /
−$813.75 conc / $57,598.10 total; West 115 / 104,250 / 0.8174 / $84,708 /
−$2,429 / $95,751.05 (= printed 95,653.05 + 98); combined 174 / 170,451 /
$131,381 / −$3,242.75 / $153,349.15.

## 5. Writer changes this format forced (same file, outside the parser)

### More than 18 floor plans

Namespacing the plan codes gives Westlake **26** plans, but the template ships
18 slots (`Floor Plan Summary` rows 3–20, `Floor Plan` rows 6–23 with
Total/Average at 24/25). The old writer silently dropped plans 19+ and the
Floor Plan Total row then counted fewer than the real doors. `_expand_fp_slots()`
clones the last template plan row down (formulas translated with
`openpyxl.formula.translate.Translator`, styles copied), re-lays the
Total/Average rows at the new bottom, and keeps the K-column conditional format
and the print area in step; `_rewrite_fp_totals()` then re-points the
SUBTOTAL/SUMPRODUCT ranges. The `n < 18` trim path is unchanged.

### A floor plan with no OCCUPIED units → #DIV/0! everywhere

Westlake plan `E-C2.5 RWD` is a single vacant 2,275 sf unit. The Floor Plan
Summary's `AVERAGEIFS` columns (O "All Units" market rent, P/Q/R "Currently
Occupied Units") have no row to average and return `#DIV/0!`; the Floor Plan
tab's `SUMPRODUCT` Total/Average rows inherit it, so the **entire** market-rent
and NER rollup printed `#DIV/0!` (11 error cells, confirmed by a real-Excel
recalc). `_harden_fps_formulas()` wraps those four cells in `IFERROR(...,0)` on
the plan rows only — identical values wherever a real average exists. Verified
against the delivered Eclipse of White Rock workbook: re-running that deal
through the modified script produces a byte-for-byte identical workbook except
those 16 formula strings (4 plans × 4 cells), with every cached value unchanged.

Note that a plan with no occupied unit ALSO has no market-rent estimate under
the house rule (`estimate_market_rents()` reports it and leaves column I
blank), so such a plan legitimately shows $0 market rent / $0 in-place rent on
the summary tabs. Say so in the delivery notes — it is one door of understated
GPR, and the fix is a cited market rent from ownership, not a guess.

## 6. Run command

```powershell
python toolkit\process_rent_roll.py "inbox\Westlake RR Jul 2026 MFG.xlsx" `
    --property "Westlake" --asof 2026-07-31 --estimate-market `
    -o "outbox\RR - Westlake - 7-31-2026.xlsx"
```

`--asof` is optional here (the export prints one) but pinning it is cheap
insurance. `--estimate-market` is mandatory for this format — without it the
deliverable ships a blank market-rent column.

## 7. The parser class (merge this into `process_rent_roll.py`)

Place it immediately after `AppFolioUnitTypeXlsxParser`, before the
`PARSERS`/`XLSX_PARSERS` registry.

```python
class ResManItemizedXlsxParser(RentRollParser):
    """ResMan "Rent Roll (Itemized)" XLSX export (Westlake, 7/31/2026).

    Report-parameter preamble in column A ("Rent Roll (Itemized)" title,
    "Exported On:", "Portfolios:", "Units:", "Include Advertised Rent:",
    "As of:") above a single header row:

        Unit | Unit Type | BD/BA | Status | Sqft | Total | Rent Income |
        Loss/Gain to Market | Concessions | <one column per itemized charge>

    Everything to the right of Concessions is an itemized charge column
    (Pet Fee-Non Refundable, Electricity/Water/Gas Reimbursement,
    Miscellaneous Income, Liability Fee, Pest Control, Electricity, Water,
    Sewer, Garbage and Recycling, ...) taken verbatim as the charge code and
    routed through `classify_charge` - so the column set is not hard-coded and
    a portfolio billing different codes still parses.

    Layout gotchas:

    * **PHASE GROUPS.** The detail is broken into phases by a bare label row
      ("East Phase" / "West Phase" - column A only, every other cell empty),
      each closed by its own "<N> Units" subtotal row (occupancy fraction in
      the Status column, then column totals). The floor-plan CODES COLLIDE
      across phases with DIFFERENT square footages ("B1.5 RWD" is 1,100 sf in
      East and 1,080 sf in West), so when more than one phase is present the
      plan code is namespaced with the phase's initial ("E-B1.5 RWD" /
      "W-B1.5 RWD"). Unit numbers are kept exactly as printed (East is
      numeric, West alphanumeric - they do not collide).
    * **Column F "Total" INCLUDES column H "Loss/Gain to Market"** (verified
      row by row and on both subtotal rows). Loss/Gain to Market is a
      computed rent-vs-market variance field, not a billed item, and in this
      export it is not maintained (populated on 2 of 174 units, with
      identical rents on the same plan printing 0 and -49). It is therefore
      NOT booked as a charge and NOT read as a market rent; the printed Total
      is tied out as "Total less Loss/Gain to Market", both numbers coming
      from the report's own subtotal row. Every unit's printed Total is
      checked against its own charge detail and a single mismatch aborts the
      run before anything is written.
    * **No market rent, no tenant names, no lease dates** anywhere in the
      export. Market rent follows the `--estimate-market` house rule
      (8/2026); Lease Start/End, Move-In/Out and MTM stay blank (MTM is never
      inferred).
    * Bed/bath comes verbatim from the BD/BA column ("2/1.50" -> 2 / 1.5) and
      is marked explicit - the plan codes ("A1 L", "B1.5 RWD FP", "C2.5 RWD")
      would otherwise be misread by the generic code fallback. Sqft is
      printed per unit.
    * Vacant units print the status and sqft and then nothing at all (blank
      Total and blank charge cells) - handled without inventing a $0 charge.
      "Vacant-Rented" counts as VACANT in the report's own occupancy
      fraction, so it maps to apt status VR.
    """
    name = "ResMan (Itemized)-xlsx"
    asof_found = True

    # Header columns that are NOT itemized charges; everything else on the
    # header row is a charge column.
    FIXED_COLS = ("unit", "unit type", "bd/ba", "status", "sqft", "total",
                  "rent income", "loss/gain to market", "concessions")
    # Detection demands the whole fixed set (the AppFolio exports carry no
    # "rent income"/"loss/gain to market"/"concessions" columns at all).
    HEADER_KEYS = FIXED_COLS

    STATUS_MAP = {            # ResMan status -> (apt_status, resident status)
        "current": ("OC", "C"),
        "evict": ("OC", "C"),
        "under eviction": ("OC", "C"),
        "notice": ("NA", "N"),
        "notice-rented": ("NA", "N"),
        "notice-unrented": ("NA", "N"),
        "vacant": ("VU", ""),
        "vacant-unrented": ("VU", ""),
        "vacant-rented": ("VR", ""),
    }

    def __init__(self):
        self.flags = []
        self.ltm_units = []          # [(unit, loss/gain to market)]

    # -- helpers -------------------------------------------------------------

    @staticmethod
    def _norm(v):
        return re.sub(r"\s+", " ", str(v if v is not None else "")).strip()

    @classmethod
    def _key(cls, v):
        return cls._norm(v).lower()

    @staticmethod
    def _num(v):
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
        try:
            return float(v)
        except (TypeError, ValueError):
            return None

    @staticmethod
    def _bb(v):
        """'1/1.00' -> (1, 1); '2/1.50' -> (2, 1.5)."""
        m = re.match(r"^\s*(\d+(?:\.\d+)?)\s*/\s*(\d+(?:\.\d+)?)",
                     str(v or ""))
        if not m:
            return None, None

        def trim(x):
            f = float(x)
            return int(f) if f == int(f) else f
        return trim(m.group(1)), trim(m.group(2))

    @classmethod
    def _find_header(cls, rows):
        for i, r in enumerate(rows[:30]):
            vals = [cls._key(v) for v in r]
            if all(k in vals for k in cls.HEADER_KEYS):
                return i, {v: j for j, v in enumerate(vals) if v}
        return None, {}

    @staticmethod
    def _phase_codes(labels):
        """Phase label -> short unique prefix ('East Phase' -> 'E')."""
        codes, used = {}, set()
        for i, lbl in enumerate(labels):
            letters = re.sub(r"[^A-Za-z]", "", re.sub(r"phase", "", lbl,
                                                      flags=re.I)) or "P"
            code = None
            for n in range(1, len(letters) + 1):
                cand = letters[:n].upper()
                if cand not in used:
                    code = cand
                    break
            if code is None:
                code = f"P{i + 1}"
            used.add(code)
            codes[lbl] = code
        return codes

    @staticmethod
    def detect_xlsx(path):
        cls = ResManItemizedXlsxParser
        try:
            from openpyxl import load_workbook as _lw
            wb = _lw(path, read_only=True, data_only=True)
            ws = wb[wb.sheetnames[0]]
            rows = [list(r) for r in ws.iter_rows(values_only=True,
                                                  max_row=30)]
        except Exception:
            return False
        pre = [cls._key(r[0]) for r in rows[:15] if r]
        if not any(p.startswith("rent roll (itemized)") for p in pre):
            return False
        if not any(re.match(r"(portfolios?|properties):", p) for p in pre):
            return False
        if not any(p.startswith("as of:") for p in pre):
            return False
        i, _ = cls._find_header(rows)
        return i is not None

    # -- parse ---------------------------------------------------------------

    def parse(self, path):
        from openpyxl import load_workbook as _lw
        wb = _lw(path, read_only=True, data_only=True)
        ws = wb[wb.sheetnames[0]]
        rows = [list(r) for r in ws.iter_rows(values_only=True)]

        hdr_i, cols = self._find_header(rows)
        if hdr_i is None:
            sys.exit("ERROR: ResMan itemized header row not found.")
        raw_hdr = [self._norm(v) for v in rows[hdr_i]]
        # every non-fixed header column, in report order, is a charge code
        charge_cols = [(raw_hdr[j], j)
                       for j in sorted(cols.values())
                       if self._key(raw_hdr[j]) not in self.FIXED_COLS
                       and raw_hdr[j]]

        def cell(r, key):
            j = cols.get(key)
            return r[j] if j is not None and j < len(r) else None

        asof, prop = None, ""
        for r in rows[:hdr_i]:
            a = self._norm(r[0]) if r else ""
            m = re.match(r"As of:\s*(\d{1,2}/\d{1,2}/\d{4})", a, re.I)
            if m:
                asof = datetime.strptime(m.group(1), "%m/%d/%Y").date()
            m = re.match(r"(?:Portfolios?|Properties):\s*(.+)$", a, re.I)
            if m:
                prop = m.group(1).strip()

        pending, phases, printed = [], [], {}
        bad_totals = []
        cur = None                       # current phase label (None = no phases)
        for r in rows[hdr_i + 1:]:
            unit = self._norm(cell(r, "unit"))
            if not unit:
                continue
            # phase banner: column A only, every other cell empty
            if all(v in (None, "") for v in r[1:]):
                cur = unit
                if cur not in phases:
                    phases.append(cur)
                continue
            m = re.match(r"^(?:total\s+)?([\d,]+)\s+units?$", unit, re.I)
            if m:                        # the phase's own subtotal row
                p = {"units": int(m.group(1).replace(",", "")),
                     "occ": self._num(cell(r, "status")),
                     "sqft": self._num(cell(r, "sqft")) or 0.0,
                     "total": self._num(cell(r, "total")) or 0.0,
                     "rent": self._num(cell(r, "rent income")) or 0.0,
                     "ltm": self._num(cell(r, "loss/gain to market")) or 0.0,
                     "conc": self._num(cell(r, "concessions")) or 0.0}
                for code, j in charge_cols:
                    p[code] = self._num(r[j] if j < len(r) else None) or 0.0
                p["other"] = sum(p[code] for code, _ in charge_cols)
                printed[cur] = p
                continue
            if self._key(unit) == "unit":            # repeated header row
                continue

            status = self._norm(cell(r, "status"))
            mapped = self.STATUS_MAP.get(status.lower())
            if mapped is None:
                mapped = ("VU", "") if status.lower().startswith("vacant") \
                    else ("OC", "C")
                self.flags.append(
                    f"unit {unit}: unrecognised Status '{status}' - treated "
                    f"as {mapped[0]}")
            apt_st, res_st = mapped

            bed, bath = self._bb(cell(r, "bd/ba"))
            u = UnitRecord(unit=unit,
                           floor_plan=self._norm(cell(r, "unit type")),
                           sqft=self._num(cell(r, "sqft")),
                           apt_status=apt_st,
                           market_rent=None,      # the export carries none
                           bed_explicit=bed, bath_explicit=bath,
                           bed_bath_explicit=True)

            rent = self._num(cell(r, "rent income"))
            ltm = self._num(cell(r, "loss/gain to market")) or 0.0
            conc = self._num(cell(r, "concessions"))
            charges = []
            if rent is not None:
                charges.append(("RENT", rent, False))
            if conc:
                # House rule: a concession over $200 in magnitude is a
                # one-time (upfront) concession - reported in col N but
                # EXCLUDED from NER; smaller ones are recurring and sit
                # inside NER. The codes are chosen so classify_charge()
                # routes them there: RECURRING_CONC_CODES matches "CONC"
                # ANYWHERE and is tested first, so the upfront code must not
                # contain the substring "conc" at all.
                charges.append(("UPFRONT" if abs(conc) > 200
                                else "CONCESSION", conc, False))
            for code, j in charge_cols:
                v = self._num(r[j] if j < len(r) else None)
                if v:
                    charges.append((code, v, False))
            if ltm:
                self.ltm_units.append((unit, ltm))

            # the printed Total column = the charge detail PLUS Loss/Gain to
            # Market; a single row that disagrees aborts the run
            printed_total = self._num(cell(r, "total"))
            if printed_total is not None or charges:
                got = sum(a for _, a, _ in charges) + ltm
                if abs(got - (printed_total or 0.0)) > 0.005:
                    bad_totals.append(
                        f"{unit}: printed Total {printed_total or 0:,.2f} vs "
                        f"charge detail + Loss/Gain to Market {got:,.2f}")

            if charges:
                if not res_st:
                    self.flags.append(
                        f"unit {unit}: status '{status}' is vacant but the "
                        f"row carries charges - booked as a future/leased "
                        f"resident so the money is not lost")
                u.residents.append(Resident(name="", status=res_st or "L",
                                            charges=charges))
            elif res_st:
                u.residents.append(Resident(name="", status=res_st,
                                            charges=[]))
                self.flags.append(
                    f"unit {unit}: status '{status}' but the row carries no "
                    f"charges at all - occupied with no rent")
            pending.append((cur, u))

        if bad_totals:
            sys.exit("ERROR: the printed per-unit Total does not equal the "
                     "row's own charge detail (+ Loss/Gain to Market) for "
                     f"{len(bad_totals)} unit(s):\n  "
                     + "\n  ".join(bad_totals))
        if not printed:
            sys.exit("ERROR: ResMan itemized export has no '<N> Units' "
                     "subtotal row to reconcile against.")

        # ---- phase namespacing ------------------------------------------
        codes = self._phase_codes(phases)
        multi = len(phases) > 1
        for ph, u in pending:
            if multi and ph and u.floor_plan:
                u.floor_plan = f"{codes[ph]}-{u.floor_plan}"
        units = [u for _, u in pending]
        if multi:
            # name the codes that actually collide, and at what sizes
            seen = {}
            for ph, u in pending:
                bare = u.floor_plan[len(codes[ph]) + 1:] if ph else u.floor_plan
                seen.setdefault(bare, {}).setdefault(ph, set()).add(u.sqft)
            collide = [(bare, d) for bare, d in seen.items()
                       if len(d) > 1
                       and len({sf for s in d.values() for sf in s}) > 1]
            eg = ""
            if collide:
                bare, d = sorted(collide)[0]
                eg = (" (e.g. '" + bare + "' is "
                      + ", ".join(f"{int(min(s)):,} sf in {ph}"
                                  for ph, s in d.items()) + ")")
            self.flags.append(
                f"the export is split into {len(phases)} phases ("
                + ", ".join(phases) + f"); {len(collide)} floor-plan code(s) "
                "are reused across phases at DIFFERENT square footages" + eg
                + ", so every plan code is namespaced with its phase initial ("
                + ", ".join(f"{ph} -> '{c}-'" for ph, c in codes.items())
                + "). Unit numbers are carried exactly as printed.")
        if self.ltm_units:
            self.flags.append(
                "the report's 'Loss/Gain to Market' column (col H) is "
                f"populated on only {len(self.ltm_units)} of {len(units)} "
                "unit(s) - "
                + ", ".join(f"{n} {v:,.2f}" for n, v in self.ltm_units)
                + " - and is inconsistent with the rents it prints, so the "
                "field is not maintained. It is a computed rent-vs-market "
                "variance, not a billed charge: it is excluded from the unit "
                "charges and NOT used as a market rent (market rents follow "
                "the --estimate-market house rule). ResMan's own Total "
                "column DOES include it, so the Total tie-out below is "
                "'Total less Loss/Gain to Market', both figures taken from "
                "the report's printed subtotal rows.")

        # ---- reconciliation ---------------------------------------------
        combined = {"units": sum(p["units"] for p in printed.values()),
                    "occ": None,
                    "sqft": sum(p["sqft"] for p in printed.values()),
                    "total": sum(p["total"] for p in printed.values()),
                    "rent": sum(p["rent"] for p in printed.values()),
                    "ltm": sum(p["ltm"] for p in printed.values()),
                    "conc": sum(p["conc"] for p in printed.values()),
                    "other": sum(p["other"] for p in printed.values())}
        for code, _ in charge_cols:
            combined[code] = sum(p[code] for p in printed.values())

        checks = {
            "unit_count": combined["units"],
            "total_sqft": combined["sqft"],
            "total_contract_rent": combined["rent"],
            "current_lease_charges": combined["total"] - combined["ltm"],
        }
        occ_known = [p for p in printed.values() if p["occ"] is not None]
        if len(occ_known) == len(printed):
            checks["occupied_count"] = sum(
                int(round(p["occ"] * p["units"])) for p in printed.values())

        def _res(us):
            return [r for u in us for r in u.residents]

        def _rent(us):
            return sum(r.rent_charge or 0 for r in _res(us))

        def _conc(us):
            return sum((r.recurring_concessions or 0)
                       + (r.upfront_concessions or 0) for r in _res(us))

        def _other(us):
            return sum(r.other_income or 0 for r in _res(us))

        def _all(us):
            return sum(r.total_charges for r in _res(us))

        def _sel(pfx):
            if pfx is None:
                return lambda us: list(us)
            return lambda us: [u for u in us if u.floor_plan.startswith(pfx)]

        groups = []
        for ph in phases or [None]:
            if ph in printed:
                groups.append((ph, f"{codes[ph]}-" if (multi and ph) else None,
                               printed[ph]))
        if None in printed and not phases:
            groups = [("Report", None, printed[None])]
        if len(groups) > 1:
            groups.append(("All phases", None, combined))

        extra = []
        for label, pfx, p in groups:
            sel = _sel(pfx)
            extra.append((f"{label}: units",
                          lambda us, s=sel: len(s(us)), p["units"], 0.5))
            extra.append((f"{label}: sq ft",
                          lambda us, s=sel: sum(u.sqft or 0 for u in s(us)),
                          p["sqft"], 0.01))
            if p["occ"] is not None:
                extra.append((
                    f"{label}: occupancy (printed fraction)",
                    lambda us, s=sel: (
                        sum(1 for u in s(us) if not u.is_vacant) / len(s(us))
                        if s(us) else 0.0),
                    p["occ"], 1e-9))
            extra.append((f"{label}: Rent Income (col G)",
                          lambda us, s=sel: _rent(s(us)), p["rent"], 0.01))
            extra.append((f"{label}: Concessions (col I)",
                          lambda us, s=sel: _conc(s(us)), p["conc"], 0.01))
            extra.append((f"{label}: itemized charges -> Other Income",
                          lambda us, s=sel: _other(s(us)), p["other"], 0.01))
            extra.append((
                f"{label}: Total (col F) less Loss/Gain to Market (col H)",
                lambda us, s=sel: _all(s(us)), p["total"] - p["ltm"], 0.01))

        def _code_total(us, code):
            return sum(a for r in _res(us) for c, a, _ in r.charges
                       if c == code)

        for code, _ in charge_cols:
            extra.append((f"  charge column '{code}'",
                          lambda us, c=code: _code_total(us, c),
                          combined[code], 0.01))
        checks["extra_checks"] = extra
        return prop, asof, units, checks

    def source_note(self, asof):
        return (f"Generated from {self.name} rent roll {self.source_kind} "
                f"({asof.strftime('%m/%d/%Y') if asof else 'unknown date'}). "
                "Contractual Rent = the 'Rent Income' column; Recurring / "
                "Upfront Concessions = the 'Concessions' column (split at the "
                "$200 house-rule threshold); Other Income = the sum of the "
                "report's itemized charge columns (pet fee, utility "
                "reimbursements, misc. income, liability fee, pest control, "
                "electricity, water, sewer, garbage). Floor-plan codes are "
                "prefixed with their phase (E- / W-) because the phases reuse "
                "the same codes at different square footages. The export "
                "carries no market rent, no tenant names and no lease dates, "
                "so those columns are estimated (market rent) or "
                "intentionally blank.")
```

## 8. The writer helpers (same file, before `write_workbook_from_template`)

Plus these three call-site edits inside `write_workbook_from_template()`:

```python
    fps_ws = wb["Floor Plan Summary"]
    plans = _floor_plans(units)
    if len(plans) > FP_TEMPLATE_SLOTS:                     # NEW
        _expand_fp_slots(wb, len(plans))                   # NEW
    for i, r in enumerate(range(3, 3 + max(len(plans), FP_TEMPLATE_SLOTS))):
        ...
    n = len(plans)
    if n != FP_TEMPLATE_SLOTS:                             # was: if n < 18
        fp_ws = wb["Floor Plan"]
        if n < FP_TEMPLATE_SLOTS:
            fp_ws.delete_rows(6 + n, FP_TEMPLATE_SLOTS - n)
        _rewrite_fp_totals(fp_ws, n)
    _harden_fps_formulas(fps_ws, n)                        # NEW
```

```python
# The shipped template carries this many floor-plan rows: 'Floor Plan
# Summary' rows 3..20 and 'Floor Plan' rows 6..23 (Total/Average at 24/25).
FP_TEMPLATE_SLOTS = 18


def _clone_row(ws, src, dst, max_col):
    """Copy one template row (values, formulas translated to the new row, and
    the exact cell styling) to another row on the same sheet."""
    from copy import copy as _copy
    from openpyxl.formula.translate import Translator
    for c in range(1, max_col + 1):
        s = ws.cell(row=src, column=c)
        d = ws.cell(row=dst, column=c)
        v = s.value
        if isinstance(v, ArrayFormula):
            d.value = ArrayFormula(
                d.coordinate,
                Translator(v.text, origin=s.coordinate)
                .translate_formula(d.coordinate))
        elif isinstance(v, str) and v.startswith("="):
            d.value = Translator(v, origin=s.coordinate) \
                .translate_formula(d.coordinate)
        else:
            d.value = v
        d._style = _copy(s._style)
    if ws.row_dimensions[src].height is not None:
        ws.row_dimensions[dst].height = ws.row_dimensions[src].height


def _expand_fp_slots(wb, n):
    """Grow the template's floor-plan blocks to hold `n` plans.

    The template has FP_TEMPLATE_SLOTS (18) plan rows on each summary tab.
    A property with more distinct plans than that - e.g. a two-phase asset
    whose colliding plan codes are namespaced apart (Westlake: 13 + 13 = 26) -
    would otherwise have its extra plans silently dropped from the Floor Plan
    and Floor Plan Summary tabs, and the Total row would count fewer than the
    real door count. The last template plan row is cloned down (formulas
    translated, styles preserved); the caller re-points the Total/Average
    formulas with _rewrite_fp_totals(). """
    from copy import copy as _copy
    if n <= FP_TEMPLATE_SLOTS:
        return
    extra = n - FP_TEMPLATE_SLOTS

    # ---- Floor Plan Summary: plan rows 3..20, nothing below -------------
    fps = wb["Floor Plan Summary"]
    last = 2 + FP_TEMPLATE_SLOTS
    for k in range(1, extra + 1):
        _clone_row(fps, last, last + k, 26)

    # ---- Floor Plan: plan rows 6..23, then Total (24) / Average (25) ----
    fp = wb["Floor Plan"]
    last_fp = 5 + FP_TEMPLATE_SLOTS
    keep = {}                       # style of the Total/Average rows
    for off in (1, 2):
        keep[off] = [_copy(fp.cell(row=last_fp + off, column=c)._style)
                     for c in range(1, 15)]
    for k in range(1, extra + 1):   # overwrites the old Total/Average rows
        _clone_row(fp, last_fp, last_fp + k, 14)
    for off in (1, 2):              # ... which are re-laid at the new bottom
        for c in range(1, 15):
            cell = fp.cell(row=5 + n + off, column=c)
            cell.value = None
            cell._style = keep[off][c - 1]
        fp.row_dimensions[5 + n + off].height = \
            fp.row_dimensions[last_fp + off].height
    # keep the K-column conditional format and the print area in step
    try:
        from openpyxl.formatting.formatting import ConditionalFormattingList
        old = list(fp.conditional_formatting)
        fp.conditional_formatting = ConditionalFormattingList()
        for cf in old:
            sq = (f"K6:K{5 + n}" if str(cf.sqref) == f"K6:K{last_fp}"
                  else str(cf.sqref))
            for rule in cf.rules:
                fp.conditional_formatting.add(sq, rule)
    except Exception:
        pass
    if fp.print_area:
        fp.print_area = f"$A$1:$N${7 + n}"


# Floor Plan Summary columns that average over a SUBSET of a plan's units
# (All Units market rent, and the three "Currently Occupied Units" columns).
_FPS_GUARD_COLS = ("O", "P", "Q", "R")


def _harden_fps_formulas(fps_ws, n):
    """Make the Floor Plan Summary averages survive a plan with no occupied
    units.

    `AVERAGEIFS` returns #DIV/0! when no row matches, so a floor plan whose
    every door is vacant (Westlake 7/31/2026: plan E-C2.5 RWD is a single
    vacant unit) poisons columns O-R, and the Floor Plan tab's SUMPRODUCT
    Total/Average rows inherit the error - the whole deliverable prints
    #DIV/0!. Each of those four cells is wrapped in IFERROR(...,0) on the
    plan rows: identical values wherever a real average exists, 0 where the
    average is genuinely undefined."""
    pat = re.compile(r'^=IF\((\$?A\d+)="","",(.+)\)$', re.S)
    for r in range(3, 3 + n):
        for col in _FPS_GUARD_COLS:
            c = fps_ws[f"{col}{r}"]
            v = c.value
            if not isinstance(v, str) or "IFERROR" in v:
                continue
            m = pat.match(v.strip())
            if m:
                c.value = f'=IF({m.group(1)}="","",IFERROR({m.group(2)},0))'
```
