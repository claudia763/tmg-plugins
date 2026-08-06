# Multifamily Brokerage Plugin

Disposition toolkit for The Multifamily Group–branded seller advisory work.
Five skills, fully self-contained (brand colors, logo artwork, advisor
headshots, document templates, and all accumulated deal-narrative playbooks
are bundled inside — no project memory or external folder required).

## Skills

### `underwriting`
Runs a deal through the TMG Excel underwriting model (the
"Initials  Property Name  Date" template): populates Master / FinalRR /
Final_T_12 / comp tables from a rent roll, T-12, Rent Data workbook, and CAD
card; selects rent comps, value-add programs, and cap-rate factors; tunes the
Assumptions tab to TMG's green thresholds (target IRR 25%, CoC > 10%,
T-3 DSCR > 1.25 — DSCR waived when T-3 economic loss exceeds 30%, since the
deal is assumed to go bridge); and prints the "PDF Output - F&C" deliverable.
- `references/model-map.md` — tab-by-tab cell map of the model.
- `references/technical-notes.md` — openpyxl/LibreOffice gotchas (external
  defined names, array-formula spills, recalc loop, 996-row comp cap).
- `scripts/export_pdf.py` — single-sheet PDF exporter (LibreOffice macro).

Triggers: "underwrite this model/deal," "populate the model," rent roll +
T-12 attached alongside the model workbook.

### `sales-comps`
Generates TMG's Sale Comparables workbook from an "Automatic CMA Analysis"
workbook (comps universe = its All Sale Comps tab): geocodes the subject,
scores every comp on distance/vintage/unit count/sale date (agreed Aug-2026
settings: >3-yr sales hard-excluded; distance 100/50/25 pts at <1/<3 mi;
cap-rate drift from the AgencyDrift tab), trims $/unit outliers, selects the
best 5, and writes a per-deal CMA copy plus a relinked
"<Deal> - Sale Comparables.xlsx" (both files must stay in one folder).
- `assets/Sale Comparables Workbook.xlsx` — the output template.
- `scripts/select_comps.py` / `build_output.py` / `verify_output.py`.
- `references/scoring.md` and `references/relinking.md` — the selection
  settings and the hard-won Excel external-link surgery notes.

Triggers: "run sales comps," "comp grid," "CMA for <property>."

### `loan-terms-lookup`
Quotes current estimated multifamily debt terms (Fannie, Freddie incl. SBL,
HUD 223(f)/221(d)(4), LifeCo, CMBS, bank, credit union, bridge, mezz/pref)
from TMG's Estimated Loan Terms workbook, building every rate as live index
(UST/SOFR) + spread. **Note: this copy carries SKILL.md only — add
`scripts/query_terms.py` and `assets/Estimated Loan Terms - Multifamily
Debt.xlsx` from the original build before relying on it.**

Triggers: "current loan terms on bridge," "where is agency pricing today."

### `broker-valuation-summary`
Creates the underwriting writeup: a polished .docx broker valuation summary
(seller advisory / disposition analysis) in the TMG navy/gold brand.
- `scripts/template.js` — the docx build template (three zones: brand
  infrastructure / data / narrative), current brand palette with embedded logo.
- `references/narrative-variants.md` — the deal-archetype playbook (assumable
  loan variants, free-and-clear, underwater, rate-constrained,
  debt-capacity-constrained, tertiary, etc.) distilled from ~25 prior builds.
- `references/build-notes.md` — brand palette spec, environment workarounds,
  and the post-build verification checklist.

Triggers: "broker valuation summary," "valuation summary for <property>,"
"turn this underwriting into a writeup," etc.

### `bov-deck`
Creates a Broker Opinion of Value as a designed, landscape, OM-style PDF deck
(1700×1080/page) from a completed valuation summary or underwriting data.
- `assets/bov_template.html` — complete worked example (11 pages) with the
  full CSS design system.
- `assets/` — navy + white logo variants, advisor headshots.
- `scripts/render.js` — Playwright/Chromium HTML→PDF renderer.

Triggers: "make a BOV," "broker opinion of value," "valuation deck/presentation."

## Setup on a new account

1. Install this plugin (open the `.plugin` file in the Claude desktop app).
2. Optional but recommended: copy your "Multifamily Valuation" folder (sample
   documents, prior valuations, sample OM PDF) to the new machine and connect
   it to a project.
3. Recommended project instructions for that project:

   > Reports are used in a brokerage capacity; the target audience is the
   > owners of the multifamily complex being valued. Use terms like "valuation"
   > instead of "being offered" — the offering price is often unknown since we
   > haven't won the business of the seller.

Both skills already carry this rule internally; the project instruction is a
belt-and-suspenders reinforcement.

## Notes

- The bundled `template.js` matches the current brand palette (navy #1B3E6F,
  gold #FDB714, embedded wordmark logo). Verify with the staleness check in
  `references/build-notes.md` if builds ever look off-brand.
- Advisor headshots/contacts reflect the team as of Aug 2026 (Land, Krebbs,
  Yazbeck, Davis) — confirm the lineup per deal.
