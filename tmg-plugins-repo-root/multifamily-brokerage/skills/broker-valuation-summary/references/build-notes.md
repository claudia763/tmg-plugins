# Build Notes & Verification Checklist

Recurring snags and required checks when generating broker valuation summaries
with `scripts/template.js`. Following these avoids re-debugging.

## Audience & language (always)

- Reports are used in a brokerage capacity; the target audience is the OWNERS of
  the multifamily complex being valued.
- Use "valuation" / "recommended valuation," never "being offered" / "offered at" —
  the offering price is often unknown since the seller's business hasn't been won yet.

## Environment workarounds

- **EPERM "rename operation not permitted" from Write/Edit tools.** When working
  in a folder mounted from the user's machine (device bridge), file tools can
  fail overwriting a file that was created or is held by the Linux shell (e.g. a
  template copied via `cp`, or the generated `.docx`).
  **Workaround:** do the whole build inside the shell. Write the Node build
  script with a bash heredoc (`cat > build.js << 'EOF' ... EOF`) and run it with
  `node`. Don't mix: if bash created the file, keep editing it from bash.
- **Generate, then copy.** Build the `.docx` in the outputs/working directory,
  then `cp` (or commit via the device bridge) into the user's project folder.
- **docx dependency.** Run `npm install docx` once per session in the outputs
  dir before building; the template requires `./node_modules/docx`.
- **Verifying content — escape `$` in shell.** When grepping extracted
  document.xml from `python3 -c "..."`, bash treats `$` as a variable and `\$`
  as a regex anchor, giving false "missing" results for dollar figures.
  **Workaround:** use a quoted heredoc (`python3 << 'PYEOF'`) and plain
  substring (`in`) checks instead of regex for `$`-prefixed values.

## Brand palette (ZONE 1 of the template)

Matches The Multifamily Group's OM brand identity:

- Brand navy: `#1B3E6F` (NAVY — title, section headers, table headers)
- Coordinated mid-navy: `#345279` (DARK_BLUE — agency table header, middle callout box)
- Pale navy tint: `#DCE6F2` (LIGHT_BLUE — subject/callout shading)
- Brand gold: `#FDB714` (GOLD — accent rules, bullets, price line, highlight cells)
- Logo: "the multifamily group." navy wordmark (700×107 PNG), embedded as base64
  (`LOGO_BASE64`), rendered centered at the top of the title block.
- **Gold-contrast rule:** `#FDB714` is bright — use NAVY text (not white) on
  gold-filled table/callout cells; gold-on-navy boxes are fine as-is.

Legacy palette (must NEVER appear): navy `1F3864`, gold `C9A84C`, light blue `D6E4F0`.
Before building, cheap staleness check on whatever template copy you're using:

```bash
grep -oE '1F3864|C9A84C|D6E4F0|1B3E6F|FDB714|LOGO_BASE64' template.js | sort | uniq -c
```

Legacy hexes must count zero and `LOGO_BASE64` must be present.

## Verification checklist (every finished .docx)

1. Legacy hex counts `1F3864` / `C9A84C` / `D6E4F0` in `word/document.xml` = ZERO.
   Healthy counts on a full-length build: roughly `1B3E6F` ×180, `345279` ×32,
   `DCE6F2` ×31, `FDB714` ×68.
2. Exactly one `word/media/*.png` byte-identical to `LOGO_BASE64`, with exactly
   one `<w:drawing>` reference in the document.
3. Every key dollar figure and percentage from the source underwriting appears
   in the document (substring checks; normalize en-dashes before comparing).
4. Data hygiene (lessons from past builds):
   - Verify "newest / largest / only" superlative claims against every comp row.
   - Reconcile quoted cap rates to the NOI tables they reference, exactly.
   - Callout/bridge table rows must sum exactly to the model delta they explain.
   - Omit source-sheet ranges that don't reconcile with computed values; quote
     the reconciling midpoint only.
   - Internally inconsistent loan maturity dates → use "N-year term" language.
   - Audit the T-12 for missing/lumpy months (annualize the complete months and
     disclose the delta); check insurance and utility lines especially.
