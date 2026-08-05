---
name: bov-deck
description: >
  Use this skill to create a Broker Opinion of Value (BOV) as a designed,
  landscape, OM-style PDF deck for a multifamily property, in The Multifamily
  Group navy-and-gold brand. Trigger whenever the user asks for a "BOV",
  "broker opinion of value", "valuation deck", "valuation presentation",
  "pitch deck for the owner", "OM-style valuation", or asks to turn a broker
  valuation summary / underwriting into a presentation or PDF deck. The input
  is typically a completed broker valuation summary (.docx) or underwriting
  data; the output is a multi-page 1700x1080 landscape PDF.
---

# BOV Deck Skill

## What this skill produces

A ~11-page landscape PDF (1700×1080 px per page) styled like The Multifamily
Group's offering memorandums — navy/gold brand, gold right-edge blade, branded
headers, data tables, gold highlight bands — but written for the **owner of the
property** (a valuation pitch), not for buyers. Use "valuation" / "recommended
valuation" language, never "offered at" / "offering price."

## Inputs

Usually a completed broker valuation summary (built with the
`broker-valuation-summary` skill) or the same underlying data: recommended
valuation and range, key metrics, rent comps, sale comps, agency loan comps,
expense normalization / value-add table, financing assumptions, return metrics,
market context (Yardi), and the recommendation narrative. If a valuation
summary docx exists, extract its numbers rather than re-deriving them.

Ask the user up front (AskUserQuestion) about: output format if unclear
(PDF vs. editable PPTX) and how to handle photos (real photos supplied vs.
branded placeholder frames — web image downloads are typically blocked in the
sandbox, so default to placeholder frames the client can swap later).

## Page structure (adapt as the deal requires)

1. **Cover** — navy, hero photo (or placeholder), angled navy panel with white
   logo, "Broker Opinion of Value" kicker in gold, property name, address,
   units, year built, "Confidential | Prepared for Ownership | <date>"
2. **Investment Advisors** — headshots + contact info (assets provided),
   navy angled bands, photo placeholder
3. **Valuation Summary** — left navy card (property snapshot table) + right
   white card (Recommended Valuation figure, range, key metrics table)
4. **Valuation Strategy / Highlights** — executive summary prose on the left,
   large navy "Valuation Highlights" headline with gold highlight bands on
   the right (like the OM's "Investment Highlights" page)
5. **Rental Analysis** — market-context bullets (Yardi) + rent comp table with
   AVERAGE row (gold) and SUBJECT row (pale gold), plus a gold-left-border
   callout box for the top disclosure/retrading item
6. **Deal Optimization** — expense normalization / value-add table, navy circle
   badge with pro forma NOI + cap rate, "buyer-identifiable levers" bullets
7. **Sale Comparable Analysis** — adjusted comps table + discount/premium
   narrative + navy card framing the spread (operational vs. structural)
8. **Agency Loan Comparables** — Fannie/Freddie survey table + takeaway bullets
9. **Financing** — constraint-or-advantage narrative + indicative loan terms
   table + three navy stat tiles (IRR / equity multiple / cash-on-cash)
10. **Recommendation** — range rationale, three marketing pillars (pale gold
    cards), risk factors, navy callout stack (range / returns / target buyer)
11. **Back cover** — logo, confidentiality note

Adapt content per the deal's narrative variant (see the
`broker-valuation-summary` skill's `references/narrative-variants.md`) — e.g.,
an assumable-loan deal makes Financing an advantage page, an underwater deal
reframes the Recommendation page around hold/negotiate paths.

## Build workflow

1. Create a working directory; copy `assets/` from this skill into it (logos,
   headshots, `bov_template.html`).
2. Copy `bov_template.html` to `<property>_bov.html` and rewrite the CONTENT of
   each page for the subject property, keeping the CSS, page scaffolding, and
   class system intact. The template is a complete worked example (Aden Crest,
   Fort Worth) — every page shows the intended layout and tone.
3. Render with `scripts/render.js`:
   ```bash
   npm ls -g playwright || npm install -g playwright   # chromium usually pre-installed
   NODE_PATH=<global node_modules> node render.js <property>_bov.html <Property>_BOV.pdf
   ```
4. Verify (below), deliver the PDF, and save a copy to the user's project
   folder if one is connected.

## Design rules (must hold on every page)

- Colors: navy `#1B3E6F` (dark variant `#16345E`), gold `#FDB714`, pale gold
  bands `#FDEFD2`, body ink `#333B45`. NAVY text on gold fills, never white.
- Font: Carlito (Calibri metric-compatible, present in the sandbox); Calibri/
  Arial fallback. Headings extra-bold navy.
- Every content page: gold blade on the right edge (`.blade`), gold page-number
  chip bottom-right (`.pgnum`), header with title + gray pentagon logo box +
  "multifamilygrp.com / <Property> | <City>, <ST>" at top right.
- **The header logo box must sit at `left:740px` or further right** — closer
  collides with long titles like "Sale Comparable Analysis."
- Tables: class `t` — navy header row, zebra rows, gold `hl` row for
  averages/totals, pale-gold `subj` row for the subject property.
- Photos: if no property photos are provided, use the built-in placeholder
  frames (`.ph` — navy gradient, dashed gold border, building silhouette SVG,
  "Property Photo Placeholder" label). Do NOT let placeholder labels peek out
  from behind overlapping cards — remove the label divs when a placeholder is
  mostly covered.
- Advisor assets included: `hs_yonnic.png` (Yonnic Land, Sr. Managing
  Director), `hs_jon.png` (Jon Krebbs, Managing Partner), `hs_paul.png`
  (Paul Yazbeck, Managing Partner), `hs_chase.png` (Chase Davis, COO), plus
  `tmg_logo_navy.png` / `tmg_logo_white.png`. Confirm the advisor lineup with
  the user if the deal team may differ.

## Verification (before delivering)

1. Extract all text with PyMuPDF; normalize en-dashes to hyphens; substring-
   check EVERY key figure from the source data (prices, caps, comps, returns).
   Line wraps break naive matching — check numbers and names separately.
2. Render each page to PNG (PyMuPDF `get_pixmap`) and visually inspect a
   contact-sheet grid: no text collisions, no placeholder labels peeking from
   behind cards, headers aligned, tables not overflowing.
3. Confirm owner-facing language throughout — no "offered at," no buyer-pitch
   framing.

## Output filename

`<Property_Name>_BOV.pdf` (underscores for spaces).
