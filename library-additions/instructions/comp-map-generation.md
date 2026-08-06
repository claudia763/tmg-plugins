# Generating a sale-comp location map (subject + comps) as a PNG

Covers: producing a clean, branded comp map image when a workbook's original
map is stale (the template Bing geocoder is dead) or missing. Use for the TMG
model's "Sale Comparable Summary" page, BOV decks, or any deliverable needing
a subject-plus-comps map. Requires internet (OSM tiles), Node + playwright
with chromium.

## Method

1. Copy `scripts/comp_map_template.html`, fill in the `pts` array
   (lat/lon/label; `subject: true` marks the gold subject pin) and the legend
   rows. Viewport in the template is 1530x1000 px; change `#map` CSS and the
   screenshot viewport together if another size is needed.
2. Screenshot with `scripts/shot_map.cjs`:
   ```powershell
   $env:NODE_PATH = "$env:APPDATA\npm\node_modules"   # global playwright
   node shot_map.cjs                                   # writes comp_map.png
   ```
   (Keep the `.cjs` extension — a parent package.json with "type": "module"
   breaks `require()` in `.js` files.)
3. Insert into Excel via COM sized to the target frame (points, not pixels):
   ```python
   frame = ws.Range("I121:O140")           # the map panel range
   w = frame.Width - 4; h = w / 1.53       # keep the 1530:1000 aspect
   if h > frame.Height - 4: h = frame.Height - 4; w = h * 1.53
   ws.Shapes.AddPicture(png_path, 0, -1, frame.Left + 2, frame.Top + 2, w, h)
   ```
   Do NOT reuse the old image's stored pixel extents as points — 1530 px
   is 1147 pt and will massively overflow the frame.

Style: numbered navy (#1B3E6F) circle pins with white borders, gold (#FDB714)
"S" pin for the subject, white legend box bottom-left — matches TMG brand.
Wait for `networkidle` + ~6s so all OSM tiles land before the screenshot.
