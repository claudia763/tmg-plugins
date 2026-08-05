# Royal Oaks Apartments — processed 8/4/2026 (scanned QuickBooks P&L + handwritten rent roll)

Source: one 11-page scan. Pages 1–4 = QuickBooks "Profit & Loss, August 2025
through July 2026" (Cash Basis, printed 08/03/26) split across sheets: page 2
is the lower half of page 1 (Aug 25–Apr 26), page 4 the lower half of page 3
(May 26–Jul 26 + TOTAL). Pages 5–11 = 7 handwritten "Month Ending July 2026"
collection sheets (sheets 1-7 … 7-0). Both are image scans — no text layer —
so everything was transcribed manually from 200-DPI page renders and validated
against the documents' own printed/handwritten totals.

## T-12 (manual → generic owner-statement XLSX → process_t12.py)

- Transcription was written into the generic `parse_t12_xlsx` layout (months
  header + TOTALS column + ALL-CAPS section headers) so all of the script's
  row/subtotal/grand validations ran against the transcription. This is the
  recommended path for any future image-only statement: build the
  intermediate XLSX, name grand rows "Total Operating Income" / "Total
  Operating Expense" / "Net Operating Income" so `_GRAND_PATS_DEFAULT`
  matches, and put the below-NOI QuickBooks Other Income/Expense block after
  the NOI row (it lands in Capex & Misc automatically).
- QuickBooks subaccounts "Property" appear twice (under Insurance and under
  Taxes); disambiguated as "Insurance - Property" / "Taxes - Property" etc.
- Every monthly subtotal column (Total Insurance/Taxes/Utilities/Wages,
  Total Expense, Net Ordinary Income, Net Income) cross-foots; all three
  reconciliations tie: Revenue 1,884,853.61 / OpEx 904,285.15 / NOI
  980,568.46 (QuickBooks "Net Ordinary Income" = NOI; "Net Income"
  980,068.46 after the 500.00 Nov-25 Other Expense, which went to Capex &
  Misc).
- **Mapping ruling (Dmytro 8/4/2026): this statement's bare "Maintenance"
  line (17,031, lumpy job-sized amounts, wages broken out separately) is
  `rm`, overriding the corpus-majority exact hit `pr` (the known 16-vs-13
  conflict).** Done with a run-local corpus copy prepending
  `,Maintenance,rm,…`; the shared t12_mappings.csv was left untouched.
  Watch this line on any owner-books statement — the exact corpus hit is
  silent (no REVIEW flag).
- Everything else mapped cleanly: Rental→r, Cleaning/roofing/Supplies→rm,
  Bank Service Charges/Licenses and Fees/Telephone→ad, Bonus/Wages-Other→pr,
  Electric→e, Gas→o, Trash→tr, Water→w, Insurance→i, Taxes→tx.

## Rent roll (handwritten → driver script → toolkit writer)

119 units (102–496, evens above 129), 115 occupied, 4 vacant (222, 402,
448, 490). Every page's Base Rent / Other / Late / Total Paid column was
tied to the sheet's handwritten TOTALS row, and the grand base rent ties to
Dmytro's verified 153,735 (26,525 / 24,175 / 22,900 / 25,350 / 24,150 /
25,425 / 5,210).

House decisions (Dmytro, 8/4/2026) — reuse for future owner collection sheets:

- "Other" column → Other Income (250 each on units 123, 125, 420; 750
  total). Late charges (1,200 collected) and Total Paid are collections
  activity, excluded from the RR.
- Tenant-named units with Ø base rent → the owner's margin figure as
  Contractual Rent, red-flagged: 123 Yesenia Flores 1,325; 420 Angelica
  Rojo 1,450; 214 Carmen Ortiz had no margin figure → uniform 1Bd rate
  1,325, red-flagged. Flagged rents total 4,100 (so the RR's contract-rent
  sum is 157,835 = 153,735 verified + 4,100 flagged).
- Vacant units' margin figures → Market Rent, red-flagged (1,325 / 1,450 /
  1,325 / 1,450 on 222 / 402 / 448 / 490).
- Bed/Bath/SqFt are NOT on the sheets; supplied by Dmytro: 1Bd = 780 sf,
  1/1.0; 2Bd = 900 sf, 2/2.0 (uniform across all units).
- As-of 7/31/2026 assigned ("Month Ending July 2026"; no printed date).

Rents: 1Bd 1,325 everywhere; 2Bd 1,450 except the 102–110 block at 1,500
(105 is 1,450). Two units carry sub-rate base rents with margin notes that
sum to the full rate — 125 José Maravilla 650 (+ "$675" margin) and 492
Maria C. Beltran 985 (+ "$465" margin): kept at the written base (they're in
the verified 153,735); the margins look like balances owed, not rent.

Low-confidence handwriting reads (verify if they matter): "Esmeralda
Juarez" (109), "Kenia Marciel" (129), "Santiago Croc" (354), "Jaran
Sevilla" (430), "Ruis Sustainta" (442).

Deliverables: `RR - Royal Oaks - 7-31-2026.xlsx`, `T-12 - Royal Oaks -
July 2026.xlsx`, `Capex & Misc - July 2026.xlsx` (500.00 Other Expense).
