# BOV deck — vertical space budget, collision rules, and Windows render notes

Read this alongside the `bov-deck` skill BEFORE writing page content. It covers
the two things that actually go wrong when the template's worked example is
rewritten for a new deal: **content overflowing the fixed 1080 px page height**
(every `.page` is `overflow:hidden`, so overflow is silently *cut off*, not
flagged), and the **Windows render toolchain**. Built 8/2026 on the 8111 Landing
Apartments BOV (San Antonio, 12 units).

## 1. The overflow trap

`bov_template.html` pages are fixed 1700 × 1080 px with `overflow:hidden`. Deal
copy is almost always longer than the Aden Crest example's copy, so the first
render will quietly truncate the last bullet, the total row of a table, or the
bottom half of a callout box. **Nothing errors. You will only catch it by
looking at the page PNGs.** Render every page and inspect all of them — see
`scripts/verify_bov_deck.py` in this folder.

### Vertical budget per content page

| Band | y-range | Notes |
|---|---|---|
| Header (`.hdr`) | 0 – 150 | rule at 132; content must start ≥ 165 |
| Content | ~170 – 1040 | **≈ 870 px of usable height** |
| Page-number chip | 1040 – 1080 | keep clear |

### Line-height arithmetic that predicts the fit

For Carlito/Calibri at font-size `F` px:

- Rendered line height ≈ `F × 1.45` for `ul.blt li`, `F × 1.5` for `p.body`.
- Characters per line ≈ `(column_width − padding) / (F × 0.47)`.
  Handy anchors measured on real renders: at F=24 in an 840 px column ≈ 62
  chars/line; at F=22.5 in the same column ≈ 66; at F=21 in a 764 px column
  ≈ 72; at F=22.5 in a 1080 px column ≈ 96.
- `ul.blt li` adds a 12 px bottom margin per bullet.
- `table.t` row height ≈ `F × 1.2 + 22` (11 px padding top and bottom); the
  header row is ~2 px taller. **A wrapped cell adds a full line to the row.**

Estimate before you render: sum the blocks, and if the total lands above ~840
px of the 870 available, cut text rather than shrinking the font below ~20 px
(smaller than that stops reading as the OM's type).

### Killing table wraps is the cheapest height saving

Two wrapped rows cost roughly one extra row of height each. Prefer shortening
the *content* over shrinking the font:

- Abbreviate long header cells (`Origination` → `Orig.`, `Value/Unit` → `$/Unit`,
  `Monthly Income/Unit` → `Inc./Unit`, `Multiple` → `Mult.`).
- Split a two-fact cell into two rows (`T-12 NOI $17,221 ($13,021 after
  reserves)` → a `T-12 NOI` row and a `T-12 NOI After Reserves` row). Two clean
  rows are shorter than one wrapped row and read better.
- Estimate a table's natural width as `Σ(longest cell chars) × F × 0.47 +
  32 × columns`. If that exceeds the column width, something wraps.

### Per-page layouts that fit (validated)

- **Valuation Summary (p3)** — left navy card at font 22 / line-height 1.72
  holds **12 label-value rows plus a 3-line Profile** if no value wraps.
- **Rental Analysis (p5)** — 6 bullets at font 22.5 in an 840 px column
  ≈ 23 lines; that is the ceiling.
- **Deal Optimization (p6)** — a 7-row bridge table at font 21 (top 178,
  ends ~495), a 4-line paragraph, then **two side-by-side bottom blocks
  starting at y≈664**: a 7-row table at font 18.5 on the left (700 px) and
  4 bullets at font 20.5 on the right (790 px). If you also want the circle
  badge, park it at `left:1300px; top:180px` — clear of both the header's
  `.site` text and the 72 px gold blade.
- **Financing (p9)** — a 3-line intro plus 5 bullets at font 22.5 in an
  850 px column is the maximum; the right column holds a 9-row table, a
  three-tile stat row, and one 3-line callout.
- **Recommendation (p10)** — a 3-line intro, the three pale-gold pillar cards
  at font 20, and **5 risk bullets at font 21** (≈11 lines total). Anything
  longer loses the last two bullets.

## 2. Header collision rules

- The skill's rule is `.hdr .logobox` at `left:740px` or further right. On this
  build **770 px** was used and is safer.
- The real constraint is the **subtitle** (`.hdr .sub`), not the `h1`. At font
  28 the subtitle must stay under ~48 characters or it runs under the logo box.
  Two that failed and had to be cut:
  - "Fannie Mae / Freddie Mac Disclosed MBS Data — San Antonio Region" → drop
    the region (it belongs in the source footnote anyway).
  - "Clean Delivery, But the Debt Sizes to the Income — Not to 70%" →
    "Clean Delivery — But the Debt Sizes to the Income".
- Rough test: subtitle chars × 13 px must be < 760.

## 3. Real photos instead of placeholder frames

When ownership supplies photos (they often arrive as inline images on the email,
not as a separate deliverable), **replace the `.ph` divs entirely** rather than
layering images over them — a leftover `.lab` / `.lab2` label peeking out from
behind a card is the defect the skill warns about. Use
`scripts/prep_deck_photos.py` in this folder to strip the MLS copyright strip
and downsize first.

Placements that worked with three photos (one aerial, two interiors):

- Cover: aerial full-bleed, plus a 90° gradient scrim
  (`rgba(10,26,48,.55) → .18 → .74`) so the gold kicker and white property name
  stay legible over a bright daytime shot. Raise the left navy panel to
  `rgba(22,52,94,.93)` — at the template's `.68` the white logo washes out
  against a bright photo.
- Advisors page: aerial at `left:0; top:150px; width:1000px; height:780px`.
- Valuation Summary: aerial full-bleed under a flat `rgba(16,38,68,.35)` scrim.
- Agency comps page: two interiors stacked at `left:1190px`, 392 × 395 each.
- Back cover: aerial under `rgba(22,52,94,.86)`.

## 4. Rendering on Windows

Three gotchas, all one-time:

1. **`render.js` fails with `require is not defined in ES module scope`** when a
   parent directory's `package.json` declares `"type": "module"`. Copy it into
   the job folder as **`render.cjs`** and call `node render.cjs ...`. Do not
   edit the copy in the read-only library.
2. **Chromium is not bundled** with the npm package: run
   `npx --yes playwright install chromium` once per machine.
3. Set `NODE_PATH` to the global modules directory before invoking node:
   `$env:NODE_PATH="$env:APPDATA\npm\node_modules"`.

Full sequence:

```powershell
npm install -g playwright                      # once
npx --yes playwright install chromium          # once
Copy-Item <skill>\scripts\render.js .\render.cjs
$env:NODE_PATH="$env:APPDATA\npm\node_modules"
node render.cjs 8111_landing_bov.html 8111_Landing_Apartments_BOV.pdf
python verify_bov_deck.py 8111_Landing_Apartments_BOV.pdf figures.txt
```

A rendered 11-page deck with three full-bleed photos lands around 2.5 MB, which
is fine as an email attachment. If it exceeds ~8 MB, the photos were not
downsized.

## 4b. Rendering on the Linux agent server (added 8/6/2026, Westlake)

Same ESM gotcha as Windows, different paths. On `ubuntu-8gb-nbg1-2`:

```bash
cp "<skill>/scripts/render.js" ./render.cjs          # .cjs — parent package.json is "type":"module"
ln -sfn /home/claudia/email-cowork-server/node_modules ./node_modules   # for the HTML's asset resolution
NODE_PATH=/home/claudia/.npm-global/lib/node_modules node render.cjs westlake_bov.html Westlake_Apartments_BOV.pdf
python3 <library-additions>/scripts/verify_bov_deck.py Westlake_Apartments_BOV.pdf figures.txt \
        --extra-banned "Dmytro,Gladchenko,Yonnic"
```

Playwright and Chromium are already installed here — `npm install` is not
needed; confirm with `node -e "console.log(require.resolve('playwright'))"`
(resolves to `/home/claudia/.npm-global/lib/node_modules/playwright`).
An 11-page deck with four photos renders in a few seconds at ~2.5 MB.

**Use `--extra-banned` for people who must NOT appear.** On Westlake the
requester asked to be left off the advisors page; passing his name as a banned
phrase turns "did I actually remove him everywhere" into a build gate instead of
a manual scan. Same trick for a dropped comp or a superseded price.

## 4c. Sourcing real photos when ownership supplies none

The Yardi e-brochure usually carries 2–4 usable subject photos even when no
photo package exists — extract them from the brochure PDF rather than shipping
placeholder frames:

```python
import fitz
b = fitz.open("Property - E-Brochure.pdf")
for pno in (0, 1):                       # cover + the data page carry the photos
    p = b[pno]
    for im in p.get_images(full=True):
        px = fitz.Pixmap(b, im[0])
        if px.n - px.alpha > 3: px = fitz.Pixmap(fitz.csRGB, px)
        if px.width >= 300: px.save(f"p{pno+1}_{im[0]}.png")
```

They come out small (typically 400×320 or 640×512), so **upscale with LANCZOS
plus a mild unsharp mask before use** — browser bilinear scaling to a 1700 px
page looks visibly mushy, and a full-bleed cover is a 2.6× upscale:

```python
im = im.resize((tw, th), Image.LANCZOS).filter(
        ImageFilter.UnsharpMask(radius=2, percent=65, threshold=3))
```

Crop to the target aspect ratio *before* resizing, and lean on a heavier cover
scrim (`rgba(10,26,48,.72) → .22 → .78`) to mask residual softness.

**Do not crop off the "Image courtesy of Yardi Matrix" watermark.** It is an
attribution, not an MLS junk strip — `prep_deck_photos.py`'s copyright-strip
removal is for MLS photos and should not be pointed at Yardi images. Instead,
flag the photography gap in the Recommendation page and recommend ownership
commission a shoot before launch.

## 5. Verification gotcha

`&ndash;`, `&rarr;` and `&dagger;` extract as separate text runs, so PyMuPDF can
return `"- 2.1%"` where your check expects `"-2.1%"`. Before "fixing" a missing
figure, grep the extracted text for its surrounding context — the number is
usually there and correct.
