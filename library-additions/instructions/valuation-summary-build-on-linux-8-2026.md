# Building the broker-valuation-summary .docx on the Linux agent server (8/6/2026)

Covers: the environment fixes and zone-splitting workflow needed to run the
main library's `broker-valuation-summary/scripts/template.js` on the
email-cowork server. Read this **before** the skill's own
`references/build-notes.md`, which assumes a Windows/device-bridge setup and
does not mention the two blockers below. Validated on Westlake Apartments
(174 units, Lubbock TX, 8/6/2026).

## The two blockers (both fatal, both one-line fixes)

### 1. `require is not defined in ES module scope`

`/home/claudia/email-cowork-server/package.json` contains `"type": "module"`,
so **every `.js` file under that tree is treated as ESM** and the template's
CommonJS `require()` calls fail immediately.

**Fix: name the build script `.cjs`, not `.js`.**

```bash
cat zone1.js zone2.js zone3.js zone4.js > build.cjs
node build.cjs
```

Keep the *fragments* as `.js` (they are never executed directly, only
concatenated); only the concatenated, executed file needs `.cjs`.

### 2. `Cannot find module './node_modules/docx'`

ZONE 1 of the template hardcodes `require('./node_modules/docx')` — a
**relative** path. `docx` is installed at
`/home/claudia/email-cowork-server/node_modules`, not in the job's build dir.

**Fix: symlink it, so ZONE 1 stays byte-identical.**

```bash
ln -sfn /home/claudia/email-cowork-server/node_modules ./node_modules
```

Do NOT edit the require line in ZONE 1 — the whole point of the zone
discipline is that ZONE 1 stays diffable against the library template. A
symlink satisfies the relative path without touching the file. There is also
no need to `npm install docx` per job; it is already installed at the server
root (confirm with `node -e "console.log(require.resolve('docx'))"`).

## Zone-splitting workflow (recommended)

The skill suggests `head -265 template.js > build.js` then appending. Splitting
into four files instead makes iteration far cheaper, because you rebuild by
re-concatenating rather than regenerating one large file:

```bash
cp "<skill>/scripts/template.js" .
head -265 template.js > zone1.js     # ZONE 1 infrastructure — never edit
#   zone2.js  -> data (propertyInfo, keyMetrics, comp rows, callouts)
#   zone3.js  -> narrative (intro/financing/rental/optimization/comps/conclusion)
#   zone4.js  -> table builders + document assembly
cat zone1.js zone2.js zone3.js zone4.js > build.cjs && node build.cjs
```

`head -265` is correct for the current template — line 265 is the last line of
`headerFooter()`, the final ZONE 1 function. **Re-check that boundary** if the
library template changes: `grep -n "ZONE 2" template.js` and take the line
before the comment banner.

Editing zone files with the Write/Edit tools works fine on this server (the
EPERM/device-bridge warning in the skill's build notes does not apply here),
even for files originally created by a bash heredoc.

## Adding sections beyond the six-section default

Extra sections are wired in `zone4.js`, not ZONE 1. Two things are needed:

1. A builder function alongside `buildRentCompTable()` etc. — reuse the ZONE 1
   helpers `dataTable(headers, widths, rows, {headerFill})` and
   `calloutRow(cells, widths)` rather than constructing `Table` by hand.
2. An entry in the `children:` array, e.g.
   `sectionHeader("..."), ...mySection.slice(0,2), buildMyTable(), ...`.

Splitting a narrative array with `.slice()` around a table is the template's
idiom for interleaving prose and exhibits — follow it.

**Gotcha:** if you add a `...somethingPost` narrative array, make sure the
assembly actually references it. A trailing `...section.slice(5)` on a
five-element array silently renders nothing and the section just ends early.

## Column widths

Column widths are DXA units in per-table arrays and must be tuned by eye —
too narrow and a header wraps mid-word ("Occupan / cy"). Keep each table's
width array summing to the same total when rebalancing, e.g. the rent-comp
table's `[1600, 900, 900, 1150, 850, 1050, 1210]` sums to 7660 as the original
`[1600, 1100, 900, 900, 900, 1100, 1160]` did.

## Verification (run every build)

Beyond the skill's checklist, this pair catches the most failures:

```bash
soffice --headless --convert-to pdf X.docx      # then render pages to PNG and LOOK at them
```

```python
# substring checks, NOT regex — bash/regex mangle "$"
import zipfile, re, base64
z = zipfile.ZipFile("X.docx"); xml = z.read("word/document.xml").decode()
txt = re.sub(r'<[^>]+>', '', xml)
assert all(xml.count(h) == 0 for h in ("1F3864", "C9A84C", "D6E4F0"))   # legacy palette
raw = base64.b64decode(re.search(r'const LOGO_BASE64 = "([^"]+)"', open("zone1.js").read()).group(1))
assert all(z.read(n) == raw for n in z.namelist() if n.endswith(".png"))
assert xml.count("<w:drawing>") == 1
missing = [f for f in EVERY_KEY_FIGURE if f not in txt]                 # must be empty
```

Build a list of **every** dollar figure and percentage the document asserts and
check them all — on Westlake that was 63 values, and the check is what proves
no exhibit drifted from the source model during narrative edits.

Note the healthy-count guidance in the skill's build notes (`1B3E6F` x180 etc.)
describes a longer build; a 7-page summary lands nearer `1B3E6F` x73,
`345279` x22, `DCE6F2` x13, `FDB714` x33. Treat those as a smell test for
"did the palette get applied at all," not as a pass/fail threshold.

## LibreOffice rendering notes

- `soffice --headless --convert-to pdf` is available at `/usr/bin/soffice`.
- LibreOffice repeats a table's header row when the table breaks across a page,
  which reads fine; add `cantSplit: true` to row properties only if a single
  row is splitting across the break.
- Always render and actually look at page 1 plus every page containing a new
  table before shipping.
