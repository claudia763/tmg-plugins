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
   (G48), avg CoC (F7) ≥ 10%, and DSCR (I8) ≥ 1.25** — with renovations
   marked as needed. No cushion above the IRR target — just green.

## The IRR target is 20% on agency debt, NOT 25% (corrected 8/7/2026, Aldine)

**Never hardcode 25%.** `G48` is a formula, not an input:

```
G48 = IF($G$63="Non-Recourse", $K$48, $M$48)      K48 = 0.20   M48 = 0.25
G63 = INDEX('Loan Terms'!$L$4:$L$16, MATCH($G$62,...))   -> col L = Recourse
```

So the target is **20% for non-recourse** (all agency, HUD, LifeCo, CMBS) and
**25% only for recourse** (bank balance sheet, credit union). Since almost every
deal is quoted on agency debt, 20% is the normal case.

`F48`/`H48` do ship as 0.25 and `references/model-map.md` records them as "Target
IRR — 0.25", but those are the Low/High scenario columns; `G48` is the STRIKE
column and the only one the green conditional format reads. Tuning to 25% on an
agency deal silently crushes the strike — on Aldine it cost roughly $500k of
supportable price (\$5.82M vs \$6.39M).

**Set `M62` (the program dropdown) before reading `G48`.** G62 → G63 → G48 is a
lookup chain off the program name; with M62 blank the chain returns "" and G48
falls through to the recourse target. Confirmed independently on Pointe at Garden
Oaks and Aldine.

## Related trap: `G63` is Recourse, `G64` is Max LTV

`G64 = INDEX('Loan Terms'!$F$4:$F$16, ...)` is the LTV cell. Writing leverage into
`G63` (as `scripts/model_price_solver.py` did until 8/7/2026) both fails to change
leverage AND destroys the recourse text G48 depends on, flipping the target from
20% to 25%. Two silent errors pushing the same direction — see that script's
header. Each program also carries its **own** minimum DSCR in `'Loan Terms'` col G
(HUD 1.176x, LifeCo 1.35x, everything else 1.25x); size to the stricter of that
and the 1.25 house rule, or you quote a loan the lender would not make.

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
