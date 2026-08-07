# Owner-made ONE-ROW-PER-UNIT rent roll XLSX (blank-tenant / facility-label)

Covers the owner rent-roll spreadsheet with one row per unit, no status
column, a Tenants column that is BLANK on every leased door, and a facility
label ("Office", "Shop", "Storage") on the doors that are not apartments.
Read this when a rent roll .xlsx has a
`COMPLEX | APT #NUMBER | FLOOR TYPE | Sq. FEET | Tenants | Move In |
Lease Expires | ADV. RENT | CURRENT RENT` header and the Tenants cells are
mostly empty. Added 8/2026 (Aldine Apartments, Houston TX, 96 doors, owner
Goldenwrist Capital LLC).

Parser: `OwnerSimpleRowRentRollXlsxParser` in `process_rent_roll.py`,
registered immediately BEFORE `OwnerSheetXlsxParser` in `XLSX_PARSERS`.
Reference copy: `scripts/owner_simple_row_rent_roll_parser.py`.

## THE IMPORTANT PART: this is NOT the Werner dialect, and it looks identical

`OwnerSheetXlsxParser` (Werner Creek, 8/2026, same owner group) reads a sheet
with **exactly the same nine header captions**. It is a different dialect, and
running the wrong one produces a plausible workbook with every occupancy
answer wrong:

| | Werner Creek (`OwnerSheetXlsxParser`) | Aldine (`OwnerSimpleRow...`) |
|---|---|---|
| Tenants column | a STATUS word: `Current`, `Vacant-Unrented`, `Notice` | blank on leased doors; a facility USE on the rest |
| occupancy comes from | the status word | the presence of CURRENT RENT |
| non-revenue doors | none | 7 of 96 |

What the Werner parser does to an Aldine sheet: its openpyxl pass reads a
blank status as *occupied* while its own independent re-extraction reads the
same blank as *vacant* (so the reconciliation cannot even be internally
consistent); it reads `Office` / `Shop` / `Storage` as resident names on
occupied doors; and it has no non-revenue concept at all.

**The header cannot tell them apart. The Tenants column can, exhaustively.**
Do not settle for registration order:

- the new parser requires that **no** data row's Tenants cell is a Werner
  status word, and that at least three rows pair a **blank** Tenants cell
  with a numeric CURRENT RENT;
- `OwnerSheetXlsxParser.detect_xlsx` was **tightened** to require at least one
  Werner status word in the Tenants column (it previously returned True on
  the header alone).

Both directions are then structurally impossible, not merely masked. Prove it
before shipping:

    python scripts/parser_detection_regression.py --toolkit ./toolkit \
        --files "./sources/*.xlsx" "./toolkit/*.xlsx" \
        --expect OwnerSimpleRowRentRollXlsxParser="./sources/new roll.xlsx" \
        --expect OwnerSheetXlsxParser="./sources/werner style roll.xlsx"

If you do not have a real file of the other dialect to hand, synthesise a
six-row one — the point of the run is the claim matrix, not the data.

## The five traps

1. **Occupancy is CURRENT RENT, never the dates.** The sheet prints no
   status. A door with a contractual rent is occupied; a door without one is
   vacant. Aldine A94 carries Move In 6/6/2025 **and** Lease Expires
   5/31/2026 **and no CURRENT RENT** — a vacated/notice door whose dates were
   left on the row. Any rule shaped like "has dates ⇒ occupied" mis-states it,
   and it is the single most expensive error available in this format
   (it moves a door from vacant to occupied and invents in-place rent for it
   if you also back-fill from ADV. RENT). Carry the dates through to the
   deliverable verbatim; just do not read occupancy off them. FLAG the door.

2. **A facility label is a NON-REVENUE door, not a vacancy and not a
   resident.** Aldine: A01 = Office, A03/A05 = Shop, A73/A75/A77/A79 =
   Storage. These are doors ownership has taken out of the rentable pool.
   - They belong in **neither** the occupied nor the vacant bucket. Counting
     them vacant understates occupancy by 7 doors (88.5% instead of 95.5% of
     the revenue doors); counting them occupied overstates it and invents
     $0-rent leases.
   - `UnitRecord.non_revenue` drives it. `is_vacant` and `on_notice` both
     return False for such a unit and the writers stamp Occupancy Status
     **`Non-Rev`** — the literal `rentroll_template.xlsx` already counts in
     Floor Plan Summary col K and shows on the Floor Plan tab. No template
     change is required; the column was always there waiting.
   - **Carry their ADV. RENT.** The source's own market-rent total, and the
     companion statement's MARKET RENT TOTAL, both include them. Dropping it
     breaks the tie-out.
   - The use goes in Lease Type as `Non-Revenue (Office)` so it survives into
     the Floor Plan tabs and the rediQ import.
   - A facility row that DOES print a current rent is carried as a normal
     occupied door and FLAGged — a door collecting rent is not non-revenue.
   - File the identity **occupied + vacant + non-revenue = unit count** as a
     reconciliation check. It is the one check that catches a door quietly
     changing buckets.

3. **`FLOOR TYPE` is bed/bath ("1/1.00"), not a plan name.** The sheet names
   no plans. Synthesise `{bed}x{bath}` and append a size suffix
   (`1x1-650`) **only when one bed/bath count spans more than one Sq. FEET
   value in that roll** — Aldine's plans each span exactly one size, so they
   stay clean `1x1` / `2x1`, while Werner's 2/1.00 doors (829/840/900/930 sf)
   would still get split. Assert one-sqft-per-plan in the pre-send gate.
   Bed, bath and sqft all come from the source: nothing here is an estimate,
   so **do not** reach for `--sqft-est` / `--bedbath-est`.

4. **MTM markers vary and an unrecognised one is silent.** Aldine writes
   `M2M` in the Lease Expires cell (Werner writes `MTM`). An unmatched marker
   just becomes "no lease expiration", which looks like sloppy data rather
   than a bug. The regex accepts `MTM` / `MTOM` / `M-T-M` / `M2M` /
   `month to month`. A literal marker IS the source stating a month-to-month
   term, so MTM = Yes and Lease Expiration blank — consistent with the house
   rule that MTM is never *inferred* from a missing or expired date.

5. **No as-of date, no totals row.** `asof_found = False` makes `--asof`
   mandatory (the run hard-exits rather than guess), and the as-of date must
   be reported as **INFERRED**, on the Comments tab and in the delivery
   summary. Because there is no totals row of any kind, `parse` ALWAYS
   re-derives the door count, all three occupancy buckets, total sqft, total
   market rent, total contract rent and the per-plan sub-totals through an
   independent second pass over the raw worksheet XML in the .xlsx zip
   (`_reextract`), sharing no code with the openpyxl pass — the same
   discipline as `OwnerSheetPdfParser` and `OwnerBlockRentRollXlsxParser`.

## Bracketing an as-of date from the roll's own lease dates

The honest method, and the one to write down in the notes:

- the LATEST move-in shown **with** a current rent sets the floor (Aldine:
  A54/A55 move in 4/4/2026 and pay rent ⇒ as-of ≥ 4/4/2026);
- the EARLIEST lease expiry still shown **with** a current rent sets the
  ceiling (Aldine: A40/A56/A78 expire 4/30/2026 and still pay ⇒ as-of ≤
  4/30/2026);
- take month-end inside that window (TMG convention) ⇒ **4/30/2026**.

State the bracket, not just the answer, so the broker can overrule it in one
line. It is an inference, and it drives the filename, cell B2 and every
"as of" the buyer reads.

## Cross-check against the companion statement in the same workbook

This dialect ships the T-12 and the rent roll as two sheets of ONE workbook.
That makes a real, free cross-check available: the roll's total ADV. RENT
should equal the statement's `MARKET RENT TOTAL` memo line. The parser scans
the sibling sheets for that row and either files a reconciliation check naming
the matching months, or FLAGs with the statement's values and the difference —
never a silent pass.

At Aldine the two disagree, and that is worth knowing: the roll totals
**104,160**, which ties exactly to the owner's 2025 statement for Jun-2025
through Nov-2025, while the 2026 statement (the sheet that ships with the
roll) prints **103,200** for Jan–Mar 2026 — $960 lower, i.e. $20/door on the
48 two-bedrooms. The rent roll is the authority for the rent-roll
deliverable; say so and move on.

## Also worth knowing

- A marketing brochure's square footages are a survey estimate and are NOT a
  correction to ownership's rent roll. Aldine's Yardi brochure claims
  780 / 850 sf against the roll's 650 / 800 sf. Use the roll, and report the
  discrepancy — a 17% sqft error would move every $/sf metric in the deck.
- Quote occupancy **both ways** whenever a property has non-revenue doors:
  Aldine is 88.5% of all 96 doors and 95.5% of the 89 revenue doors. Only the
  second is an occupancy statistic; the first is a door count.

## Before you send

    python scripts/verify_deliverable_workbooks.py outbox/ --excel

Note the gate itself was fixed alongside this parser: its "Contractual Rent
populated on all N units" check is now scoped to rows whose Occupancy Status
is `Occupied` (it previously failed on every rent roll that had a vacancy),
and it prints the occupancy-status census so a mislabelled bucket is visible.
