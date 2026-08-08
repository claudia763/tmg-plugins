# The Vintage Apartments (OKC) BOV deck — font fallback, competitor-OM photos, page substitutions (8/8/2026)

Deal record for an 11-page BOV built for a **66-unit, 1953, 529 SF-average
Oklahoma City asset** off a finished broker valuation summary. Read alongside
`bov-deck-layout-and-render-notes.md` (the vertical budget and render path) and
the `bov-deck` skill. This file adds three things that were not covered there.

## 1. Carlito is NOT installed on this server — the budget arithmetic shifts

`bov-deck-layout-and-render-notes.md` §1 derives characters-per-line from
`F × 0.47`, which is **Carlito/Calibri** metrics. On `ubuntu-8gb-nbg1-2` as of
8/8/2026:

```bash
fc-list | grep -ci carlito     # 0
fc-match Carlito               # DejaVuSans.ttf  <-- NOT metric-compatible
fc-match Arial                 # LiberationSans-Regular.ttf
```

The template's stack is `'Carlito', Arial, sans-serif`. With Carlito absent the
browser falls through to **Arial → Liberation Sans**, which is roughly **10%
wider** than Calibri. Every character-count anchor in the layout notes is
therefore optimistic on this box, and text that "should" fit overflows.

Two mitigations, both cheap:

1. Put `'Liberation Sans'` explicitly in the stack after `'Carlito'` so the
   fallback is deterministic rather than depending on an Arial alias existing:
   `font-family:'Carlito','Liberation Sans', Arial, sans-serif;`
2. **Budget about 10% fewer characters per line than the notes predict**, and
   treat their per-page ceilings (6 bullets at 22.5px, 5 risk bullets at 21px)
   as upper bounds rather than targets.

On this build the base type was dropped a notch across the board — `table.t`
24→21px, `ul.blt li` 24→21.5px, `p.body` 24.5→22px, `.gband` 29→25px — and every
page still read as OM-grade at 1700 px. Do that first rather than cutting
content.

## 2. The overflow trap fired anyway — and only the PNGs caught it

Page 7's fourth callout box was **clipped mid-sentence** ("...all point to the
same number in a") with the remainder cut off by `overflow:hidden`. Nothing
errored, and critically:

- `verify_bov_deck.py` reported **147/147 figures pass** on the broken render.
  The clipped text contained no figures, so the harness was blind to it.
- The page-geometry and banned-phrase checks were also green.

**The figure harness does not detect overflow. Only looking at the page images
does.** Render every page and read all of them, every time — this is the second
build in a row where that check was the only thing standing between a truncated
sentence and a client.

Fix applied: shortened the three callout bodies and tightened their
`margin-bottom` from 20 to 16px. Prefer trimming copy over shrinking type once
you are already at ~21px.

## 3. Sourcing photos from a COMPETITOR's offering memorandum

The layout notes' §4c covers pulling photos from a *Yardi e-brochure*. Here the
only imagery was a rival brokerage's OM for the same asset — a different problem
with three specific traps. Of 46 unique rasters across 26 pages, 12 were subject
photos and 32 were rejected (map tiles, **comparable-property photos of other
assets**, the competitor's own broker headshots and logos).

- **Comp photos are the dangerous rejects.** An OM's comparables pages carry
  photographs of *other properties*. They sit in the same PDF, at the same
  resolution, and `get_images()` gives you no way to tell them apart. Nine were
  present here. **Look at every candidate before slotting it** — shipping a comp
  building as the subject's hero is an unrecoverable error.
- **Check for markup baked into pixels.** The p1 cover aerial carried a yellow
  parcel outline and a callout box burned into the image — competitor OM markup
  that cannot be cleaned. Discard rather than crop; there is usually another
  unmarked aerial deeper in the document.
- **Third-party attribution is a judgment call, not just a rule.** One exterior
  carried a CoStar/Apartments.com watermark. §4c's rule (preserve attribution,
  do not crop it off) is right — but preserving it means a competitor-adjacent
  watermark renders on a TMG advisors page. The better answer is to **re-crop a
  different, unmarked source** rather than ship the mark or violate the rule.
  Here the p7 low-oblique aerial replaced it and was the stronger image anyway.

Unlike the e-brochure case, an OM's photos are typically **2000–2500 px natives**,
so every deliverable is a *downscale*. Use the lighter cover scrim
(`rgba(10,26,48,.72) → .22 → .78` was still fine) — there is no softness to mask,
and a heavy scrim needlessly dulls a good photo.

Two real gaps worth flagging to the client rather than papering over: this OM had
**no unit kitchen and no bedroom photograph** anywhere, and the remaining
interiors (dated pink-tile bathroom, laundry room with visible wall damage) are
not marketing-grade. Recommend ownership commission a shoot before launch.

## 4. Page substitutions driven by the deal, and honest stat tiles

The stock page plan assumes an underwriting model exists. This deal's writeup was
built without one (no IRR, equity multiple or cash-on-cash), so two pages changed:

- **Page 8 "Agency Loan Comparables" → "Submarket Transactions."** No
  Fannie/Freddie survey was run for this asset. Rather than fabricate one, the
  page carries the Yardi Matrix transaction file for the submarket — which is
  real, on point, and contains the subject's own last recorded trade. Same
  substitution as the valuation summary, so the two documents agree.
- **Page 9's three stat tiles carry `IRR / equity multiple / cash-on-cash` in the
  template.** With no model, inventing them would be indefensible. They were
  replaced with figures the analysis actually produced: **T-3 cap rate at the
  recommendation, supported leverage, and the buyer's equity requirement.**
  A tile that says something true is worth more than a tile that matches the
  template.

## 5. Advisor lineup — confirm the spelling, not just the names

The request read *"Chase Davis, **Job Krebbs**, Paul Yazbeck."* That is **Jon
Krebbs** — the skill's `assets/` carries `hs_jon.png` and
`westlake-uw-writeup-8-2026.md` records the same lineup. Corrected silently on
the deck and flagged in the reply; printing a misspelled advisor name on a client
deliverable is worse than asking.

Contact details came from `templates/bov-advisor-assets/README.md`, which is
current — **not** from `bov_template.html`, whose Aden Crest example carries
stale phone numbers for Jon and Paul.

Banned-phrase list used as a build gate (all passed):

```
--extra-banned "Aden Crest,Fort Worth,Yonnic,Dmytro,Gladchenko,Greg Miller,
                Price Edwards,Job Krebbs,offered at,offering price"
```

`Price Edwards` is worth banning on any deck built from a competitor's OM, and
the misspelled `Job Krebbs` guards against the typo propagating from the request.

## 6. Deliverable

`The_Vintage_Apartments_BOV.pdf` — 11 pages, 1700×1080, 2.1 MB, five real
photographs, no placeholder frames. 147/147 figures verified, page geometry
correct, no banned phrases, all 11 pages visually inspected.

Deck numbers (all traced to the valuation summary): $2,900,000 / $43,939 per unit
/ $84.36 per SF; T-12 NOI $160,094 (5.52%), T-3 $180,772 (6.23%), Year-1 $181,124
(6.25%); 98.5% occupancy against an 83.2% submarket; adjusted comp grid
$3,006,357; debt capacity $2,150,694 on T-3 at 1.25x / 75%.

## Related

- `bov-deck-layout-and-render-notes.md` — vertical budget, header collisions,
  Linux render path (§4b), photo sourcing (§4c)
- `vintage-okc-uw-writeup-8-2026.md` — the valuation summary this deck renders
- `rent-comps-the-vintage-okc-8-2026.md`, `sales-comps-pipeline-hardening-8-2026.md`
- `templates/bov-advisor-assets/README.md` — current advisor contact details
