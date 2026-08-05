---
name: broker-valuation-summary
description: >
  Use this skill to create a professionally formatted broker valuation summary
  (also called a seller advisory, disposition analysis, or listing valuation memo)
  for a multifamily property. Produces a polished .docx file in a navy blue and
  gold color scheme with section headers, data tables, bullet points, and a
  callout summary box. Trigger whenever the user asks for a broker valuation
  summary, seller valuation, property valuation memo, disposition summary,
  listing analysis, underwriting writeup, or any document that advises an
  owner/seller on how to price and market a multifamily asset for sale. Also
  trigger when the user says something like "turn this underwriting into a
  valuation summary" or "make a writeup for this deal" and attaches financial data.
---

# Broker Valuation Summary Skill

## What this skill produces

A single-file .docx broker valuation summary written from the perspective of the
**listing broker advising the seller** — not a buyer pitch. The target audience
is the owners of the multifamily complex being valued; the business has usually
not been won yet, so the offering price is unknown. The document includes:

1. **Introduction** — Property overview and recommended pricing rationale
2. **Financing** — Loan structure and how to use it (or its constraints) in marketing
3. **Rental Analysis** — Market context, rent comps table, property performance,
   disclosure recommendations
4. **Deal Optimization / Value-Add** — Operational upside and expense comparison table
5. **Sale Comparable Analysis** — Traditional sale comps + agency loan comps with
   pricing narrative
6. **Conclusion** — Disposition recommendation, marketing pillars, and seller
   risk considerations

## Step 1: Gather the data

Before writing, collect the following from the user's attachments, messages, or
uploaded files. Ask for anything missing that's critical to the document.

**Property basics**
- Property name, address, city, state, zip
- Unit count, year built, average unit size (SF)
- Current occupancy %

**Offering / pricing**
- Recommended valuation (and price/unit, price/SF if available)
- Pricing range (low / mid / high), if applicable

**Loan (if assumable or notable)**
- Loan balance, interest rate, LTV, maturity date
- Monthly payment (P&I), IO period remaining if any
- Equity requirement at the recommended valuation

**Financial performance**
- T-12 Total Income, NOI, Total Expenses
- T-3 or Pro Forma NOI (normalized)
- T-12, T-3, and Pro Forma cap rates
- Key income line items: gross potential rent, loss to lease %, vacancy %,
  bad debt %, RUBS income, other income
- Key expense line items: payroll ($/unit), contract services, R&M, admin,
  marketing, utilities, management fee, insurance, taxes

**Return metrics** (if underwritten)
- IRR, equity multiple, avg cash-on-cash, yield on cost

**Rent comparables** (typically 3–5 properties)
- Name, address, units, year built, occupancy %, avg SF, avg rent/unit, avg $/SF

**Sale comparables** (typically 3–5 properties)
- Name, address, units, year built, sale price, sale date, adjustments, adjusted $/unit

**Agency loan comparables** (if available — Fannie/Freddie MBS data)
- Property name, units, year built, value/unit, cap rate, address

**Deal context** (narrative inputs)
- Primary distress factors or challenges (e.g., bad debt, falling rents,
  elevated payroll, deferred maintenance)
- Value-add opportunities (opex reductions, RUBS expansion, etc.)
- Whether renovation upside is supportable (check if renovated units command
  a meaningful premium over unrenovated — if not, do NOT lead with renovation)

## Step 2: Pick the narrative variant

Read `references/narrative-variants.md` and identify which deal archetype this
property matches — the pricing story, section structure, table substitutions,
and risk framing all change with the variant. Archetypes covered: assumable-loan
hooks (below-market rate / outsized LTV / loan sizing / LTV-as-price-floor),
free-and-clear with an owner ask-gap, underwater / equity-wipeout (HUD and
bridge-debt sub-variants), rate-constrained premium- and discount-to-grid,
tertiary bank-debt-only deals, debt-capacity-constrained pricing, and the
"assumable loan adds no value" dismissal. If none matches cleanly, use the
default six-section structure and note the closest analogues.

## Step 3: Research market conditions

Use web search to find current conditions in the property's market:
- Apartment vacancy rate and rent growth/decline trend (Yardi, CoStar, Northmarq, etc.)
- New supply delivered in the last 1–2 years and pipeline outlook
- Employment / unemployment rate and wage growth
- Any submarket-specific factors (new development concentrations, etc.)

This context goes into the Rental Analysis section and supports the pricing narrative.

## Step 4: Build the document

### Setup (run once per session in the outputs directory)
```bash
npm install docx
```

Copy `scripts/template.js` (from this skill directory) to the outputs directory,
rename it for the property (e.g., `oak_creek_valuation.js`), then fill in the
data and narrative content as described in the template.

Run it:
```bash
node oak_creek_valuation.js
```

The output file is saved as `<PropertyName>_Broker_Valuation_Summary.docx`.

### Template structure

The template has three clearly marked zones:

**ZONE 1 — Infrastructure** (never modify):
TMG brand colors (navy #1B3E6F, mid-navy #345279, pale navy #DCE6F2, gold
#FDB714), the embedded "the multifamily group." wordmark logo (LOGO_BASE64,
rendered centered at the top of the title block), border presets, and helper
functions. These produce consistent formatting across all documents.

A build-safe pattern: copy ZONE 1 verbatim with `head -265 template.js > build.js`,
then append ZONE 2/3 via a quoted heredoc. Any NEW helpers (e.g. richBullet,
smallDataTable, smallCalloutRow) go in the assembly zone below ZONE 1 so ZONE 1
stays byte-identical and diffable.

Formatting notes for the assembly/content zones:
- The brand gold #FDB714 is bright — use NAVY text on gold-filled table or
  callout cells, never white (gold-on-navy boxes are fine as-is).
- For long tables (e.g., the agency loan-comp survey), add `cantSplit: true`
  to row properties so rows break between, not across, pages.

**ZONE 2 — Data** (fill in all property-specific numbers):
JavaScript objects for property info, loan, returns, rent comps, sale comps,
agency comps, and expense comparison rows. Replace every placeholder value.

**ZONE 3 — Narrative** (write the actual text):
Six arrays of `para()` and `bullet()` calls, one per document section. This is
where the property's story gets told. Write from the seller/broker's perspective:
recommendations to ownership, not a pitch to buyers.

## Step 5: Verify the build

Read `references/build-notes.md` for the full checklist and environment
workarounds. Minimum checks on every finished .docx:
- Legacy hex counts `1F3864` / `C9A84C` / `D6E4F0` in document.xml = ZERO
  (healthy full-length builds run roughly 1B3E6F ×180, 345279 ×32,
  DCE6F2 ×31, FDB714 ×68).
- Exactly one `word/media/*.png` byte-identical to LOGO_BASE64, with one
  `<w:drawing>` reference.
- Every key dollar figure and percentage from the source data appears in the
  document text (use quoted-heredoc Python substring checks, not shell regex —
  `$` breaks bash grep).
- Cross-check any "newest / largest / only" superlative claims against every
  comp row, and reconcile derived figures (cap rates to NOI tables, callout
  sums to model deltas) exactly.

## Writing guidelines

**Tone**: The document advises the seller and listing broker. Every section should
frame findings as recommendations — what to disclose, how to position the asset,
what buyers will focus on, what protects pricing.

**Pricing language**: Use "valuation" / "recommended valuation," never
"offered at" or "offering price" — the offering price is usually unknown
pre-engagement (the business hasn't been won yet).

**Bad debt / collections**: If bad debt is elevated, proactively recommend
ownership prepare a collections action plan for the due diligence package.
Name this as a retrading risk explicitly.

**Renovation narrative**: Only support it if renovated units command a clear,
meaningful premium (at least $75–100+/unit above unrenovated effective rents),
OR if the comps themselves prove the target rents (see the tertiary variants in
`references/narrative-variants.md`). If the premium is thin, say so clearly and
recommend against leading with renovation in marketing.

**Assumable loans**: Lead with the loan as a marketing advantage when the rate
is below-market — but first check the variant guide: the hook may actually be
LTV, loan sizing, or an LTV price floor, and sub-50%-LTV loans may deserve the
"adds no value" dismissal instead.

**Per-unit basis**: When pricing involves a distress discount to comps, quantify
the spread explicitly (e.g., "40% discount to the adjusted comp average") and
explain that the discount is operational, not structural — this protects
the seller's pricing from being challenged on real estate quality grounds.

**Market context**: Use the Yardi / market research to explain *why* rents are
declining or soft. Sophisticated buyers already know; getting ahead of it with
a clear explanation builds credibility and prevents buyers from using it as
a retrading lever.

## Output filename
Save as: `<PropertyName>_Broker_Valuation_Summary.docx`
Use underscores for spaces. Save to the outputs directory so the user can
download it.
