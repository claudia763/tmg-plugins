# Multifamily Brokerage Plugin

Disposition toolkit for The Multifamily Group–branded seller advisory work.
Two skills, fully self-contained (brand colors, logo artwork, advisor
headshots, document templates, and all accumulated deal-narrative playbooks
are bundled inside — no project memory or external folder required).

## Skills

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
