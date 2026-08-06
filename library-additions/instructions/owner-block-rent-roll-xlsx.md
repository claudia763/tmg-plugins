# Owner-made MULTI-ROW-BLOCK rent roll XLSX

Covers the owner/small-PM rent roll spreadsheet where each unit occupies a
BLOCK of 5–7 rows rather than one row, and the charge codes below the rent are
a standing menu rather than actual charges. Read this when a rent roll .xlsx
has a `Unit Type | Unit # | Sq Ft | Occupancy | Resident | Market Rent |
Charge Code | Charges | ...` header and blank-looking rows between units.
Added 8/6/2026 (Magnolia Place, Brenham TX, 20 doors).

Parser: `OwnerBlockRentRollXlsxParser` in `process_rent_roll.py`, registered
first in `XLSX_PARSERS`.

## Layout

Header row (not row 1 — a title like "July 2026 Rent Roll" sits above it):

    Unit Type | Unit # | Sq Ft | Occupancy | Resident | Market Rent |
    Charge Code | Charges | Move-in Date | Lease Start | Lease End |
    Deposit | Deposit Notes | Renewal | NTV

The block's FIRST row carries unit number, sqft, occupancy, the first
resident, market rent, the `Rent` charge and its amount, and the dates and
deposit. Every following row in the block carries only a charge code, and
sometimes a second occupant name, a phone number, an extra Unit-Type word or
an extra deposit amount. A totals row closes the sheet.

## The six traps

1. **The charge menu is not charges.** Every block prints `Pet`, `Patio`,
   `NEW CREDIT`, `MISC. FEE`, `MTM`, `Late Fee` with EMPTY amounts — a
   template the owner fills in only when something is actually billed. Only
   `Rent` ever carries money. Booking these as $0 charges invents charges the
   sheet does not make; Other Income, concessions and discounts are genuinely
   nil, not zeroed. Expect typo variants (`MISC, FEE`, trailing spaces) and at
   least one menu row printed with an explicit `0.00` — still a menu row.
   File a reconciliation check that counts non-rent charges booked (must be 0)
   so this cannot regress silently.
2. **The Unit Type column holds more than the unit type.** At Magnolia it
   carried an UNLABELLED DATE on each block's first row (2025-01-01 on every
   500 sf unit, 2025-02-01 on every 850 sf unit — correlated with size, not
   with the unit), the tier word, and for two units a further `PATIO` word on
   a lower row. Take the words, leave the date out of the Floor Plan field,
   and FLAG the date for the requester rather than guessing what it means.
3. **The Resident column sometimes holds a PHONE NUMBER**, not an occupant
   (four of twenty units at Magnolia). Detect and exclude them from the names;
   report which units. Genuine co-occupant names on continuation rows join
   onto the one lease so charges and deposits still tie.
4. **A deposit can be split across two rows of the block** (Magnolia unit 12:
   $500 + $200). Sum them — that is what makes the printed deposit total tie.
   Conversely an occupied unit with a blank Deposit cell stays BLANK, not
   zero; the sheet's own total excludes it, which is the proof.
5. **A literal word in the Lease End cell.** Magnolia unit 16 prints
   `Monthly`. That is the source STATING a month-to-month term, not an
   inference from an expired date, so MTM = Yes and Lease Expiration is left
   blank — consistent with the house rule (which bars *inferring* MTM,
   including from the amount-less `MTM` menu row). FLAG it either way.
6. **No as-of date and no property name anywhere.** Set `asof_found = False`
   so `--asof` is mandatory and the run hard-exits rather than guessing; pass
   `--property` too.

## Reconciliation when the sheet prints only three totals

The totals row gives market rent, charges and deposits only — no unit count
and no total sqft. Follow the `OwnerSheetPdfParser` precedent: re-derive unit
count, occupied count and sqft through an **independent second pass** and file
them as checks labelled "vs re-extract". For an .xlsx that second pass should
read the worksheet XML straight out of the zip container, so it shares no code
path with the block parser. Magnolia lands 14 checks; fault injection (+$1 on
a market rent, inventing the missing deposit, booking the $0 menu row) makes
the block fail as it should.

## Floor plans on a source with a renovation tier

Magnolia's "Unit Type" is a renovation tier (PREMIUM / CLASSIC / PARTIAL), not
a floor plan, and bed/bath is absent from the source entirely. What worked:

- Web-source the bed/bath (see below) and name plans `{bed}x{bath} {Tier}` —
  `1x1 Premium`, `2x1 Classic`, ... Every plan then spans exactly one sqft, so
  the rollups are valid, and the tier rent spread is readable straight off the
  Floor Plan tab. Assert one-sqft-per-plan in verification.
- ALSO populate the per-unit **Renovation Status** column (Rent Roll col G,
  named range `RenovationString`). Leaving the tier only inside the plan-name
  string hides it from the rediQ import and from anyone filtering the Rent
  Roll tab. Because every unit in a plan shares its tier, Floor Plan Summary
  col C ("Renovated") then populates for every plan and flows to the Floor
  Plan tab. Put the tier in BOTH places; do not collapse plans to bare
  `1x1`/`2x1`.
- An amenity that carries no rent premium (Magnolia's enclosed patios on units
  13 and 15 — same market rent as their tier-mates) should NOT get its own
  floor plan. It creates one-unit plans whose averages are meaningless. Fold
  it into the base plan and keep it as a run FLAG so it reaches the delivery
  notes. Make it a class-level constant (`PATIO_AS_PLAN`) so the call is a
  one-line switch.

## Sourcing bed/bath for a small unnamed asset (worked example)

The property name may not be indexed anywhere. Magnolia Place returned nothing
on its own; it was found by **floor-plan match** — searching for a Brenham TX
complex with exactly a 500 sf and an 850 sf plan at the roll's rent levels.
That led to the street address, where an apartments.com listing then named the
property explicitly. Syndication sites (Zumper, PadMapper, ApartmentGuide)
publish per-unit bd/ba/sqft and are directly fetchable when
apartments.com/Zillow/LoopNet return 403.

Cross-check the answer against the county improvement area: 14,792 sf recorded
vs 14,900 sf implied by 6×500 + 14×850 is +0.73%, which independently confirms
both the door count and the size split. Note that as corroboration only — the
workbook keeps the rent roll's own sqft.
