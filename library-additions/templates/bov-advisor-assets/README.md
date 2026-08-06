# BOV / OM Investment Advisor assets — headshots and current contact details

Covers: the advisor headshots and contact block used on the **Investment
Advisors** page of a BOV deck (`bov-deck` skill, page 2) or an OM. Use these
instead of re-extracting from a PDF each job. Captured 8/6/2026 from the
current TMG OM advisors page.

## Why this exists

The main library's `bov-deck/assets/` ships only four headshots — **Yonnic Land,
Jon Krebbs, Paul Yazbeck, Chase Davis** — and has **no Greg Miller**, who is on
active deal teams. It also carries phone numbers for Jon and Paul that no longer
match the current OM page. This folder is the corrected, extended set.

## Contact details (verbatim from the 8/6/2026 OM advisors page)

| Name | Title | Email | Phone |
|---|---|---|---|
| Greg Miller | Managing Director | greg.miller@multifamilygrp.com | c. 210.901.0254 |
| Jon Krebbs | Managing Partner | jon.krebbs@multifamilygrp.com | o. 972.379.9843 |
| Paul Yazbeck | Managing Partner | paul.yazbeck@multifamilygrp.com | o. 972.379.9844 |
| Chase Davis | Chief Operating Officer | chase.davis@multifamilygrp.com | o. 972.465.9533 |
| Dmytro Gladchenko | Chief Analyst | dmytro.gladchenko@multifamilygrp.com | c. 469.789.6805 |

**Conflicts with the main library's `bov_template.html`** (which is the Aden
Crest worked example, not current): it shows Jon at `o. 972.379.9862` and Paul
at `c. 972.310.1032`. Prefer the table above.

**Dmytro is listed here for completeness only.** On Westlake (8/6/2026) he asked
explicitly to be left off the deck — *"Leave me off it, I don't like getting
spam called."* No headshot for him is stored. **Always confirm the lineup with
the requester rather than defaulting to the template's four**, and add the
excluded names to `verify_bov_deck.py --extra-banned` so a stray mention fails
the build.

## Files

`hs_greg.png`, `hs_jon.png`, `hs_paul.png`, `hs_chase.png` — 357×358 px, drop
straight into the deck's `assets/` folder. The template renders them at
`width:120px; height:124px; object-fit:cover;`.

## Re-extracting from a PDF if the lineup changes

An OM advisors page is usually image-only, so text extraction returns nothing
and `get_images()` returns the headshots in arbitrary order. **Map each image
xref to its position on the page and sort top-to-bottom** — the visual order is
the only thing that ties a face to a name:

```python
import fitz
d = fitz.open("advisors.pdf"); p = d[0]
rows = []
for im in p.get_images(full=True):
    xref = im[0]
    for r in p.get_image_rects(xref):
        if r.width < 100:                      # headshots, not the hero photo
            rows.append((r.y0, xref))
for y, xref in sorted(rows):                   # order matches the printed list
    px = fitz.Pixmap(d, xref)
    if px.n - px.alpha > 3: px = fitz.Pixmap(fitz.csRGB, px)
    px.save(f"hs_{y:.0f}.png")
```

Then render the page at 150 dpi and read it to confirm each face against its
printed name before wiring them into the deck. Do not trust `get_images()` order.
