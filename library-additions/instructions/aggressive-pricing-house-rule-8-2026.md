# Aggressive pricing house rule (Dmytro, 8/6/2026) — TMG underwriting model

Covers: how to set the strike price and Value-Add renovation picks in the TMG
underwriting model. Supersedes the "2–4 point IRR cushion" guidance in the
underwriting skill — read this whenever tuning Assumptions.

## The rule

1. **No excess DSCR.** If the T-3 DSCR (Assumptions I8) is above the 1.25
   requirement, RAISE the price until it isn't. Properties very rarely trade
   with excess DSCR; leaving it on the table underprices the listing.
2. **If IRR runs out before DSCR does, add renovation programs** (Value-Add
   rows C18–C23, up to 3) to boost IRR — even when the renovation premium is
   not strictly supportable from the rent roll. The broker's job at the
   pitch stage is promotional; the buyer will underwrite their own plan.
3. Land the strike at the **maximum price where Project IRR (F5) ≥ target
   (G48, 25%), avg CoC (F7) ≥ 10%, and DSCR (I8) ≥ 1.25** — with renovations
   marked as needed. No cushion above the IRR target — just green.

## Practical notes (validated on St Nicholas Place, 40u Benbrook)

- Binary-search G50 in $10k steps with Calculation=Automatic via Excel COM
  (~2 s/step). Test renovation combos separately — each changes both the
  income ramp and the closing capital (Value-Add H74 -> 'UW - F&C' AL50).
- Template renovation programs and their preset scopes: Light Interior
  ($1,500/unit, 20 units, +$50/mo), Premium Interior ($5,000/unit, 20 units,
  +$150/mo), Moderate Interior ($5,000/unit, all 40, +$100/mo), W/D Hookups
  ($5,000/unit, 20, +$75/mo), Exterior/Deferred ($3,000/unit, cost only —
  narrative prerequisite). Light+Premium together cover all 40 doors and beat
  Moderate on capital efficiency (St Nicholas: supported $2.83M vs $2.74M).
- Sanity-check the premium against comp $/SF so the promoted rents stay
  inside the comp range on the PDF's rent-comp page.
- Result pattern: the IRR floor usually binds well before DSCR reaches 1.25;
  residual DSCR excess is acceptable once renovations are maxed sensibly.

## Model deliverable naming (same email)

The model workbook is named **"CK - Property Name - M-D-YYYY.xlsx"**
(hyphen-separated, e.g. "CK - St Nicholas Place - 8-6-2026.xlsx").
