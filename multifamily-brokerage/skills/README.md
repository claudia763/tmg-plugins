# Rent Roll & T-12 Processor (TMG)

A Claude plugin that packages The Multifamily Group's production toolkit for
processing multifamily property financials. Give Claude a rent roll or a
T-12 / trailing-twelve operating statement — PDF or Excel, from ResMan,
Yardi, RealPage OneSite, AppFolio, Buildium, QuickBooks (including
QuickBooks Online P&L PDFs), SSI, Google Sheets, or owner-made
spreadsheets — and it converts the file into TMG's standardized underwriting
workbooks:

- **RR - Property - M-D-YYYY.xlsx** — standardized rent roll with Floor
  Plan, Rent Roll, and Floor Plan Summary tabs (live formulas, rediQ-
  compatible named ranges).
- **T-12 - Property - Month YYYY.xlsx** — trailing financials with every
  account auto-mapped to TMG's charge codes (an ~11,000-mapping corpus built
  from 367 prior deals) plus a hidden model-import tab.
- **Capex & Misc - Month YYYY.xlsx** — everything below the line (debt
  service, capex, non-operating items).

Every run reconciles the output against the source report's own printed
totals and refuses to deliver numbers that don't tie out. Missing data
(square footage, bed/bath, market rent) is never invented — the toolkit
follows TMG's house rules: cite a public source, or fill a clearly
red-flagged best estimate.

## What's inside

```
rent-roll-t12-processor/
├── .claude-plugin/plugin.json          plugin manifest
├── README.md                           this file
└── skills/rent-roll-t12-processing/
    ├── SKILL.md                        instructions Claude follows
    ├── toolkit/                        the actual scripts + templates
    │   ├── process_rent_roll.py
    │   ├── process_t12.py
    │   ├── harvest_t12_corpus.py       corpus builder (bulk archives)
    │   ├── rentroll_template.xlsx
    │   ├── t12_processor_template.xlsx
    │   └── t12_mappings.csv            charge-code mapping corpus
    └── references/
        ├── house-rules.md              TMG judgment protocols
        ├── supported-formats.md        per-format parser documentation
        └── deals/                      worked precedents (4 deal notes)
```

## Dependencies

The scripts need Python 3 with three packages:

```
pip install pdfplumber openpyxl rapidfuzz
```

(Add `--break-system-packages` if pip refuses on a managed environment.
`rapidfuzz` is optional but speeds up fuzzy charge-code matching.)

## Install — Claude Cowork / claude.ai desktop

Install the packaged `rent-roll-t12-processor.plugin` file through the
Cowork plugin manager (Settings > Plugins, or drag the .plugin file in).

## Install — Claude Code

Two options:

**Option 1 — local plugin marketplace (recommended).** Unzip
`rent-roll-t12-processor-claude-code.zip` somewhere permanent (the install
references it). The zip root is a marketplace directory containing
`.claude-plugin/marketplace.json` and the plugin. Then, inside Claude Code:

```
/plugin marketplace add /path/to/unzipped/folder
/plugin install rent-roll-t12-processor@tmg-tools
```

(Or from a terminal: `claude plugin marketplace add /path/to/unzipped/folder`
then `claude plugin install rent-roll-t12-processor@tmg-tools`.) If the
install summary says `Run /reload-plugins to activate.`, run that. The skill
is then available as `/rent-roll-t12-processor:rent-roll-t12-processing`
and triggers automatically when you ask Claude to process a rent roll or
T-12.

**Option 2 — plain skill copy.** Copy the skill folder into your personal
skills directory, where Claude Code auto-discovers it (no marketplace, no
namespacing):

```
cp -r rent-roll-t12-processor/skills/rent-roll-t12-processing ~/.claude/skills/
```

## Usage

Attach or point Claude at a rent roll or T-12 and say, for example,
"process this rent roll" or "run the T-12". Claude copies the toolkit into
the working directory, auto-detects the source format, runs the right
script, shows the reconciliation block tying out against the report's own
printed totals, and delivers the standardized workbook(s). For sources
missing square footage, bed/bath, or market rent it follows the house
rules in `references/house-rules.md` — web-sourced values are cited,
estimates are red-flagged, and nothing is ever silently invented.
