# Filing a job's materials into the Dropbox deal folder (staging `dropbox/`)

Covers: what goes where when a job asks you to maintain a `dropbox/` staging
folder alongside `outbox/`. Read at the start of any job whose instructions
mention Dropbox filing. The server mirrors `dropbox/` into the company deal
folder for that email thread after the job.

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
- `dropbox/` is separate from `outbox/`: only `outbox/` gets emailed, only
  `dropbox/` gets mirrored. A file usually belongs in both.
