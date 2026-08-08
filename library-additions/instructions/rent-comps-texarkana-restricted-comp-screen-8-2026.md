# Rent comps: screening rent-restricted comps, and three Dwellsy data traps (8/8/2026)

Covers: the screen that keeps a **tax-credit property out of a market-rate rent
comp set** (the rent-side twin of the Arbor Pointe trap on the sale side), how
Dwellsy's multi-address explosion inflates a small market, when to widen
`--comp-months` instead of relaxing a house rule, and a `rent_comp_export.py`
footnote that misrepresents verified work. Read alongside
`rent-comparable-analysis/SKILL.md` and `references/dwellsy-api.md`. Worked on
**Renaissance Square, 2401 County Ave, Texarkana, AR 71854** — 65 units, est.
1970, 550 SF 1BR / 900 SF 2BR, master-metered.

## 1. The state HFA list is a MANDATORY screen on the rent side too

A LIHTC property's asking rents are capped by AMI limits, not set by the market,
so it cannot price a conventional subject — and **nothing in Dwellsy flags it**.
Neither does the name: the offender here advertises as "Aspen Grove Apartments",
which no keyword screen catches. Its ADFA entity name is *Broadmoore Apartments
dba Aspen Grove*.

This is the same failure mode that put **Arbor Pointe** (an ADFA-financed,
age-restricted property) into the SALE comp set for this same deal — see
`sale-comps-tertiary-market-texarkana-8-2026.md` §5. Run the check on both sides
of every deal:

- **Arkansas** — ADFA county lists,
  `https://adfa.arkansas.gov/wp-content/uploads/2024/12/<County>-County-<City>.pdf`
  **WebFetch cannot parse these PDFs; the `r.jina.ai` text proxy can.**
- **Texas** — TDHCA inventory, plus the regional ADRC "Affordable Housing
  Inventory" lists (e.g. East Texas ADRC for Bowie County).
- Cross-check `affordablehousingonline.com` and `lowincomehousing.us`.

**Two cautions learned here, both of which change the answer:**

1. **The ADFA list has no program column.** Name, phone, address,
   Elderly/Disabled, and bedroom counts — that is all. Being on it proves ADFA
   financing touched the property, not that rents are restricted *today*. Aspen
   Grove is on the list, absent from every LIHTC roll, and advertises "no
   Section 8, vouchers, or rental assistance." Genuinely contradictory. **Drop
   it anyway** — an unresolvable restriction question has no place on a client
   deliverable — and say why.
2. **"Accepts vouchers" is NOT a rent restriction.** Two Texas comps here were
   tagged "mixed-income" on the Bowie County ADRC list purely because they take
   Housing Choice Vouchers; both market open rents with no income limits. Do not
   drop a good comp on that basis. The ADRC list is a social-services referral
   directory, not a regulatory one.

`scripts/rentcomp_filter_restricted.py` (contributed with this note) drops named
addresses from the flat CSVs **before** the exporter runs, because
`rent_comp_export.py` has no exclude switch and forcing exclusion by editing the
unit count in `names.csv` would misreport the property. It logs every exclusion
with its reason — paste that log into the delivery notes.

## 2. Dwellsy's multi-address explosion — the biggest single distortion

Dwellsy assigns a community's **whole unit count to each of its street numbers**.
On this pull, one new townhome development appeared as **24 separate "24-unit
properties" on Lionel Ave** — a nominal 576 units, all at $1,475, all with
identical coordinates. Left alone it outvoted the entire real comp set: my raw
quarterly cut showed a 2BR median leaping to **$1,475 in 2026Q1** purely from
that one community's listings.

**Diagnose it with coordinates and the manager field, not by eye:**

```
425 Westlawn Dr   u=37 yb=1969  Narrow Path Property Management
501 Westlawn Dr   u=37 yb=1969  Narrow Path Property Management   -> 217 ft apart
2400 Brookridge   u=32 yb=1960  Unicorn Residential
2401 Brookridge   u= 9 yb=1960  Raffaelli Property Management     -> 329 ft, SEPARATE
5201/5205 Lionel  u=24 yb=?     PH Property Management            ->   0 ft apart
```

Identical unit count + same manager + a few hundred feet = one community; merge
via the `community` column in `names.csv`. **Different unit counts and different
managers = two assets, even at 329 feet across the same street.** Verified
independently: Apartments.com markets Westlawn as literally "425-501 Westlawn
Dr", and its true size is **39 units total — not 37+37=74.**

## 3. Verify identities BEFORE trimming — verified vintages grow the comp set

The urban vintage filter drops unknown-vintage comps silently, so Step 4 is not
optional housekeeping — it directly determines how many comps you get.

| Stage | Comp properties |
|---|---|
| Unknown vintages left blank | **5** (one of them the LIHTC property) |
| After verification filled 4 missing years | **7** |

Verification recovered **1970 for Magnolia Garden** (a perfect vintage match
0.7 mi away that the filter had been discarding), 1980 for River Run, 1978 for
Pineview, and 1959 for Linden Court. Dwellsy's unit counts disagreed with the
verified value on **9 of 17 addresses**; its years on 3, and 4 were blank.

**TMG's own Fannie/Freddie workbook beat Dwellsy every time they conflicted** —
with one instructive exception. The Freddie record labelled `1410 Richmond Rd`
as *"The Pines at Richmond"*. That is the **wrong address**: 1410 is **Ridgewood
Apartments**; Pines at Richmond is a different property at **1915** Richmond Rd.
Loan records are authoritative on units/vintage/income and are still capable of
carrying a mis-keyed address — check the marketing name separately.

Note `rentcast_xref.py` needs `RENTCAST_API_KEY`, which is **not provisioned on
this host** (same class of gap as `DWELLSY_API_KEY` in
`rent-comps-without-dwellsy-8-2026.md` §1). Verification here was web + the
agency workbook.

## 4. In a thin market widen `--comp-months`, do not relax a house rule

The 6-month default left **31 listings / 5 properties**; 12 months gave
**50 / 7**. Widening the recency window costs only currency of the asking rents
and is fully disclosable. Compare that with the alternatives:

- relaxing the **150+ unit cap** would have re-admitted Beacon Point (184 u) and
  Summerhill Woods (175 u) — the two properties whose amenity packages the rule
  exists to exclude;
- relaxing the **vintage window** would have admitted 2022 and 2026 townhomes at
  $1,449–$1,595.

Widen the window first. Report the larger excluded properties as named market
context in the notes instead of forcing them into the grid — on this deal Beacon
Point is separately the best *income* analogue and belongs in the narrative, not
the comp table.

## 5. `rent_comp_export.py` prints "not independently verified" even when you did

Two places hardcode it — the XLSX grid note and the PDF grid footnote:

```
"* Unit count / year built per Dwellsy — not independently verified."
```

Ship that on a set where you completed SKILL Step 4 and the deliverable
**understates the work and invites exactly the wrong caveat** from a buyer. The
patched job-local copy takes `--verified-by "<sources>"` and swaps in a truthful
sentence naming them; the default is unchanged, so this is backward-compatible.
It belongs upstream alongside the `--source-label` / `--methodology` flags added
in `rent-comps-without-dwellsy-8-2026.md` §3.

Watch the signature when you patch: `main` is **`def main(argv=None):`**, and the
parse line is **`args = ap.parse_args(argv)`** — patching against `def main():`
or `a = ap.parse_args()` fails silently and leaves a `NameError` at run time.

## 6. Deal record — Renaissance Square, 8/8/2026

2 Dwellsy pulls of the 6-pull budget (3 mi/24 mo apartments, then 5 mi/24 mo all
types) → 462 listings, 123 addresses. Screened → 447. Urban mode ON (198
qualifying listings), vintage window 1960-1980, `--comp-months 12`.
**Comp set: 50 listings → 7 properties, all 0.7-2.8 mi, all 1960-1980.**

| Comp | Built | Units | Avg SF | Avg rent | $/SF |
|---|---|---|---|---|---|
| Magnolia Garden | 1970 | 39 | 712 | $713 | 1.02 |
| River Run (fka Quill Creek) | 1980 | 112 | 483 | $695 | 1.44 |
| Chateau Apartments | 1960 | 9 | 612 | $658 | 1.11 |
| Brookridge Apartments | 1960 | 32 | 1,007 | $875 | 0.87 |
| Patriot Apartments (utilities bundled) | 1960 | 26 | 588 | $765 | 1.33 |
| Ridgewood Apartments | 1972 | 60 | 782 | $785 | 1.02 |
| Westlawn Apartments | 1969 | 39 | 625 | $595 | 0.95 |

Subject $717 market / $698 in-place, 760 SF, $0.94/SF. Comp medians **1BR $695,
2BR $841**.

**The finding is plan-specific and it matters:** the 1BR is at market (in-place
$674 = 97% of the comp median, upside $0 once the PSF-implied $631 floor is
respected), while every 2BR sits **11-17% below** — 2/1 in-place $700 vs an $841
median. **Total demonstrated upside $4,983/mo = $59,796/yr**, essentially all of
it on the 39 two-bedroom units. Trend on the comp set: 1BR +5.0%, 2BR +14.0%
(prior 6-24 mo vs last 6 mo).

Excluded and why: **Aspen Grove** (ADFA list, §1); **511 Champion Pl** (Dwellsy's
record belongs to a different parcel, 615 Champion Pl); **Parkway Townhomes**
(2022, $1,449+) and **Townhomes on University** (2026, $1,475, student-oriented,
individually metered) on vintage; **Beacon Point** and **Summerhill Woods** on
the 150+ unit rule.

**Utilities are the open comparability question.** The subject is master-metered
and markets as utilities-included; only **Patriot Apartments** is confirmed
bundled, Westlawn includes trash, and the rest are unconfirmed. Portals do not
publish this reliably — it needs a leasing-office call round, and until then the
subject's rents are being compared against a mix of bundled and unbundled
product.

## Related

- `rent-comparable-analysis/SKILL.md` · `references/dwellsy-api.md`
- `rent-comps-without-dwellsy-8-2026.md` — the survey fallback and the exporter's other hardcoded source labels
- `sale-comps-tertiary-market-texarkana-8-2026.md` — the same restricted-comp trap on the sale side, same deal
- `renaissance-square-uw-lump-income-t12-8-2026.md` — the underwriting these rents feed
- `scripts/rentcomp_filter_restricted.py`
