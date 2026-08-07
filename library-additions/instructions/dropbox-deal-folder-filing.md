# Filing a job's materials into the Dropbox deal folder (staging `dropbox/`)

Covers: what goes where when a job asks you to maintain a `dropbox/` staging
folder alongside `outbox/`. Read at the start of any job whose instructions
mention Dropbox filing. The server mirrors `dropbox/` into the company deal
folder for that email thread after the job. For the infrastructure side —
installing, running and linking the Dropbox client on the server so that mirror
actually works — see `dropbox-headless-linux-setup.md`.

## The five categories (create only the ones you use)

| Folder | What belongs there |
|---|---|
| `- Info for Buyers/` | Finished, buyer-facing deliverables: processed T-12, rent roll, Capex & Misc, BOV deck, OM |
| `Brochures/` | Yardi brochures / market reports |
| `Rent and Sales Comps/` | Rent comps and sale comps — both what the broker submitted and what we produced (CMA workbook, Sale Comparables workbook, Rent Data) |
| `Photos/` | Property photos |
| `- Underwriting/` | Underwriting models, writeups, valuation summaries, loan-terms workbook, CAD card |

Anything valuable that fits no category goes at the `dropbox/` root.

## Rules learned in practice

- **Rename on the way in.** Inbound attachments often arrive with server
  dedupe suffixes (`Rent Data-2.xlsx`, `image-2-3-4-5.png`). File them under
  clean, descriptive names — `Rent Data.xlsx`, `Exterior - Street View.png`.
  Photos especially: name them by what they show, since the folder is what a
  future OM build will shop from.
- **Keep linked workbook pairs together.** The Sale Comparables workbook has
  live external links to `Automatic CMA Analysis.xlsx` — both go in
  `Rent and Sales Comps/` so the links still resolve.
- **Overwrite on revisions** so Dropbox always holds the latest version; do
  not accumulate `-v2` copies. `Copy-Item` overwrites by default.
- **Distinguish our output from the broker's.** When the requester sends
  their own version of a deliverable (e.g. their own underwriting summary),
  file it with an attributing suffix — `... (Dmytro 8-6-2026).pdf` — next to
  ours rather than overwriting ours.
- Raw email attachments are archived by the server automatically; the agent's
  job is only the categorized filing, so do not bulk-copy `inbox/`.
- **"Just save it, don't process it" requests are filing-only jobs.** Honor
  them literally: identify the document enough to pick the right category
  (open it, read the header), copy it in under a clean name, and stop. Do not
  build a workbook, run a parser, or produce an analysis nobody asked for.
  Verify the staged copy matches the source (`md5sum` both) and say so in the
  reply, since the file itself is the deliverable. Source documents sent this
  way still go in the category folder that matches what they are — an
  unprocessed owner T-12 PDF belongs in `- Info for Buyers/` — and keep the
  original filename if it is already clean and descriptive, so a later
  processed version (usually `.xlsx`) sits beside it rather than colliding.
- **A document about a DIFFERENT property goes at the `dropbox/` root, never
  in a category folder.** The staging folder mirrors into one deal's folder,
  and `- Info for Buyers/` is buyer-facing: a second property's operating
  statement sitting in there reads as though it belongs to the deal, which is
  worse than misfiling — it is misleading to a buyer. Tell-tales that a file
  is reference material rather than deal material: a property name that does
  not match the thread, a period years staler than the deal's own documents,
  and wording like "file this for reference." Keep the original filename when
  it already names the property and period, so the root copy is
  self-identifying, and say plainly in the reply why it went to the root.
- **Do not copy a real property's financials into `library-additions/`.** That
  folder auto-commits to the tmg-plugins GitHub repo. Even when an inbound
  workbook is a perfect exemplar of a house format, file the workbook in
  Dropbox and contribute only a written format note — never the live data.
- `dropbox/` is separate from `outbox/`: only `outbox/` gets emailed, only
  `dropbox/` gets mirrored. A file usually belongs in both.
