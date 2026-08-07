# ZONE 1 `calloutRow()` mis-sizes every cell after a spanned one (8/7/2026)

Applies to `broker-valuation-summary/scripts/template.js`, ZONE 1, the
`calloutRow(cells, widths)` helper. Found on Aldine Apartments (96 units,
Houston TX, 8/7/2026) while building a sale-comp callout that needed **two**
value columns after a wide label. Read alongside
`valuation-summary-build-on-linux-8-2026.md`.

## The defect

`calloutRow` looks a cell's width up by its **position in the `cells` array**,
not by the column it actually occupies:

```js
cells.map((c, i) => new TableCell({
  columnSpan: c.span || 1,
  width: {
    size: c.span ? widths.slice(i, i + c.span).reduce((a, b) => a + b, 0)
                 : widths[i],          // <-- i is the CELLS index
    type: WidthType.DXA
  },
  ...
}))
```

As soon as one cell carries a `span`, every following cell is off by
`span - 1` columns. The spanned cell itself is correct; nothing after it is.

**Worked example** — the template's own `buildSaleCompTable()`, with
`saleCompWidths = [1700, 700, 800, 1400, 1000, 1500]`:

| cells index | cell | width used | width it should use |
|---|---|---|---|
| 0 | label, `span: 5` | `slice(0,5)` = 5600 | 5600 — correct |
| 1 | the value | `widths[1]` = **700** | `widths[5]` = **1500** |

So the stock template ships a callout value cell declared at 700 DXA in a
1500 DXA column. It renders acceptably in LibreOffice and Word because both
resolve the final column geometry from the table's `columnWidths` grid and
treat `<w:tcW>` as advisory — which is exactly why this has survived: **it
never errors and it usually does not look wrong.** It bites when the declared
width is small enough to force an early wrap, or when a row has two or more
cells after the spanned one and the accumulated offset walks off the end of
the `widths` array (`widths[i]` becomes `undefined` -> `size: undefined`).

## Do NOT fix it in ZONE 1

ZONE 1 stays byte-identical to the library template — that is the entire point
of the zone discipline, and `diff <(head -265 template.js) zone1.js` must stay
silent. Fix it in ZONE 4, where table builders live.

## The workaround: `wideRow()` in ZONE 4

Take an explicit width per cell instead of inferring it from position:

```js
function wideRow(cells) {
  return new TableRow({
    children: cells.map(c => new TableCell({
      borders,
      columnSpan: c.span || 1,
      width: { size: c.width, type: WidthType.DXA },
      shading: { fill: c.fill || LIGHT_BLUE, type: ShadingType.CLEAR },
      margins: { top: 60, bottom: 60, left: 80, right: 80 },
      children: [new Paragraph({
        alignment: c.align || AlignmentType.LEFT,
        children: [new TextRun({
          text: c.text, bold: c.bold !== false, size: 16, font: "Arial",
          color: c.color || NAVY,
        })],
      })],
    })),
  });
}

// and, for the common "one cell per column, no span" callout:
function plainWideRow(values, widths, fill, color) {
  return wideRow(values.map((v, i) => ({
    text: v, width: widths[i], fill, color,
    align: i === 0 ? AlignmentType.LEFT : AlignmentType.CENTER,
  })));
}
```

Call site — note the label width is summed explicitly, so it cannot drift:

```js
const w = saleCompWidths;
const labelWidth = w[0] + w[1] + w[2] + w[3];
table.root.push(wideRow([
  { text: label,  span: 4, width: labelWidth, fill: LIGHT_BLUE, color: NAVY },
  { text: value,  width: w[4], fill: LIGHT_BLUE, color: NAVY, align: AlignmentType.CENTER },
  { text: perUnit, width: w[5], fill: LIGHT_BLUE, color: NAVY, align: AlignmentType.CENTER },
]));
```

On Aldine this was used for every callout row in all four tables (rent comp
average + subject, sale comp indication + subject, NOI bridge net impact +
result, debt capacity grid + recommendation). Eleven pages rendered with no
wrapped headers and no mis-sized cells.

## Related house rule you will trip over in the same function

The template's `buildSaleCompTable()` / `buildAgencyTable()` / `buildExpenseTable()`
all set GOLD-filled cells to `color: WHITE`. **Gold cells must carry NAVY
text** (the gold-contrast rule) — white on `FDB714` is unreadable in print.
Override it wherever you copy a stock builder; it is not a ZONE 1 edit, the
colour is passed in at the call site.

## Cheap check

After any build with a spanned callout row, grep the generated XML for a
missing width, which is the loud failure mode of the same bug:

```bash
python3 -c "
import zipfile
x = zipfile.ZipFile('X.docx').read('word/document.xml').decode()
print('undefined widths:', x.count('w:w=\"undefined\"'))"
```

Then rasterize and look at the row — the silent failure mode has no signature
in the XML at all.
