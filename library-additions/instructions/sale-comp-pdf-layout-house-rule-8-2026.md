# HOUSE RULE — the sale-comp PDF page 1 is FIXED; extra notes go on a "Comments" page (Dmytro, 8/8/2026)

Covers: the layout rule for `<Deal> - Comparable Sale Grid.pdf` produced by the
`sales-comps` skill's `scripts/export_comps.py`. **This is a direct instruction
from Dmytro and it supersedes earlier guidance** — specifically §6 of
`sale-comps-tertiary-market-texarkana-8-2026.md`, which told you to shrink the
map so extra notes would fit on one page. Do not do that any more.

## The rule, verbatim in effect

> "The notes in the pdf sale comp report are fixed and should not be changed as
> this messes up the map size. Any extra notes are to be added to a second page
> on the pdf as 'Comments'."

Three consequences, all load-bearing:

1. **The page-1 Notes block is FIXED.** It carries the standard selection
   sentence, the equal-weighting sentence and footnotes ¹ ² ³ — and nothing
   else. Never append to it.
2. **The map is never scaled to make room for text.** It renders at full page
   width at its designed aspect. Any mechanism that shrinks the map to reclaim
   vertical space is wrong, however clever.
3. **Every extra disclosure goes on page 2 under a "Comments" band.** There is
   no limit on how much goes there; that is the point of moving it.

## Why the old approach was wrong

The self-correcting shrink loop in the superseded §6 built the PDF, counted
pages, and shrank the map by 20pt at a time until it fit on one page. On
Renaissance Square that drove the map down to **150pt** — roughly half its
designed height — because the deal carried four disclosures. The grid was
correct and every automated check passed, and the client-facing map was still
squashed to a strip. **Page count was being protected at the expense of the
artifact.** Page 1 is a fixed report format, not a container to be optimised.

The tell is the same class as the stale-map defect in
`uw-model-linux-libreoffice-build-8-2026.md` §5a: no error, no failing gate,
just a worse document. Render and look.

## The implementation

Three edits to `export_comps.py`. Signature: drop `map_max_h` entirely.

**1. Stop appending extras to the page-1 notes block.** Where the old code had
`for extra in (extra_notes or []): notes.append(...)`, delete it and leave the
comment so nobody re-adds it:

```python
    # HOUSE RULE (Dmytro, 8/8/2026): the page-1 Notes block is FIXED. Nothing
    # is appended to it — growing it squeezes the map on the same page and
    # wrecks the map size. Extra notes go to a "Comments" page instead.
```

**2. Map at native size, always:**

```python
        story.append(Spacer(1, 8))
        story.append(KeepTogether([
            band("Map", navy_dark, colors.white, font_b, 8.5, 15),
            Spacer(1, 2),
            Image(map_png, width=PW, height=PW * h / w, hAlign="CENTER"),
        ]))
```

**3. The Comments page**, appended just before `doc.build(story)` (import
`PageBreak` from `reportlab.platypus` alongside `Image`, `KeepTogether`, …):

```python
    if extra_notes:
        story.append(PageBreak())
        story.append(band("Comments", navy_dark, colors.white, font_b, 8.5, 15))
        story.append(Spacer(1, 8))
        for extra in extra_notes:
            story.append(Paragraph(extra, note))
            story.append(Spacer(1, 5))
```

**4. Page count becomes deterministic**, so `main()` builds once instead of
looping — and a wrong count now means page 1 genuinely overflowed, which is a
real defect worth surfacing:

```python
    build_pdf(sel, rows, ind_ppu, ind_total, cap, pdf, extra_notes=args.extra_note)
    import pypdfium2 as _pdfium
    n_pages = len(_pdfium.PdfDocument(pdf))
    expect = 2 if args.extra_note else 1
    if n_pages != expect:
        print(f"WARNING: PDF is {n_pages} pages, expected {expect} — page 1 has "
              f"overflowed its fixed layout; check the grid, not the notes")
```

**5. `verify_exports.py` must stop asserting a single page.** Replace the
`"PDF is a single page"` check:

```python
        n_pages = len(rd.pages)
        p2 = (rd.pages[1].extract_text() or "") if n_pages > 1 else ""
        has_comments = "Comments" in p2
        ok = n_pages == 1 or (n_pages == 2 and has_comments)
        check("PDF is 1 page, or 2 with a Comments page (page 1 layout is fixed)",
              ok, f"{n_pages} pages")
```

Note the check still bites: 2 pages where page 2 is **not** the Comments page
means the grid overflowed, and 3+ pages always does.

## Writing the Comments themselves

- `--extra-note` is repeatable; one call per comment. Each becomes its own
  paragraph on page 2.
- **Lead each with a bold subject** — `<b>Comp 4 is ONE transaction, not two.</b>`
  — reportlab's `Paragraph` takes that inline markup and it makes the page
  scannable. This is the one formatting habit worth keeping.
- **`&#8226;` bullets and `&#8594;` arrows render; `&#8308;`–`&#8311;` (⁴–⁷) do NOT.**
  Carlito ships ¹ ² ³ but not the higher superscripts, and reportlab draws the
  missing glyph as a black square. That trap is unchanged and still invisible to
  the verifier — see `sale-comps-tertiary-market-texarkana-8-2026.md` §6.
- Escape `$` in shell heredocs (`\$23,000,000`), or the shell eats it.

## Still true, and unaffected by this change

The autofit / ellipsis patch for the grid's fixed-width cells
(`sales-comps-from-an-address-only-8-2026.md` §5, extended in
`sales-comps-pipeline-hardening-8-2026.md` §5) applies to the **grid table on
page 1** and is still required. So is stripping
`(Part of a N Property Portfolio)` out of the Address row — with the disclosure
now living on the Comments page rather than in a page-1 footnote.

**Render every page to PNG and look at it on every run.** The verifier proves
the numbers; only your eyes prove the layout. `scripts/pdf_to_png.py`.

## Applied on

Renaissance Square (2401 County Ave, Texarkana AR) — reissued 8/8/2026 with the
same five comps, the same $52,739/unit indication, page 1 restored to its fixed
layout with a full-size map, and four disclosures moved to a Comments page.
Grid and both workbooks were unchanged and cell-identical to the original issue.

## Related

- `sales-comps/SKILL.md` — Step 4 builds the exports
- `sale-comps-tertiary-market-texarkana-8-2026.md` — **§6 SUPERSEDED by this file**
- `sales-comps-from-an-address-only-8-2026.md` · `sales-comps-pipeline-hardening-8-2026.md`
- `scripts/pdf_to_png.py`
