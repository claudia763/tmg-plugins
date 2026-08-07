#!/usr/bin/env python3
"""Cross-check a sale-comp grid indication against three independent bases, and
run the subject-average-SF sensitivity that `cma-subject-inputs-trap-8-2026.md`
tells you to run before quoting any grid number.

WHAT / WHEN
    Run this on every `selection.json` produced by the `sales-comps` skill, BEFORE
    quoting the indicated value in a BOV, an underwriting writeup, or an email.
    It answers the two questions a broker will actually be asked in diligence:

      1. "Why is your adjusted number so different from what the comps actually
         traded at?"  -> the adjusted-vs-unadjusted divergence, decomposed into the
         vintage and unit-size gaps that caused it.
      2. "What if the subject's average unit size is wrong?"  -> the sensitivity
         table, which is the trap note's whole point.

THE TRAP NOTE'S TELL, AND ITS LIMIT
    `cma-subject-inputs-trap-8-2026.md` says: if the adjusted indication and the
    unadjusted simple mean diverge by more than ~1-2%, suspect the subject's avg-SF
    input (`Inputs!B8`) before you suspect the adjustments.

    That tell assumes a TIGHT comp set. It is necessary but not sufficient: a large
    divergence is equally consistent with a comp set that is genuinely newer and/or
    larger-unit than the subject. So this script does not just print the divergence
    — it prints the expected divergence implied by the vintage and size gaps. If the
    observed divergence is close to the expected one, the adjustments are doing their
    job on a heterogeneous comp set and the subject input is fine. If they disagree,
    THAT is when to go re-derive avg SF from the rent roll.

    Validated on The Montrose (2819 N Fitzhugh Ave, Dallas, 8/7/2026): observed
    divergence -14.4%, expected from the gaps -13.6%. The subject's 807 SF was
    independently confirmed (182,367 net rentable SF / 226 units), and the 807-vs-
    824.6 SF sensitivity was worth only 0.9% of value — so the wide divergence was
    real market heterogeneity, not an input error.

THE THIRD BASIS
    A $/SF indication (mean comp $/SF x subject avg SF) is included because it is the
    natural check on a size adjustment: it applies the size difference at 100% rather
    than the grid's 25%. On a subject whose units are much smaller than the comps',
    the $/SF basis is the conservative floor and the unadjusted $/unit mean is the
    aggressive ceiling; the grid should land between them. If it does not, something
    is wrong.

INPUT   selection.json from `sales-comps/scripts/select_comps.py`
OUTPUT  a report on stdout; JSON with --json

    The grid formulas replicated here are the delivered workbook's, per the trap note:
        adj_yb   = (comp_yb - subj_yb) / -100 / 2
        adj_size = -(comp_sf / subj_sf - 1) / 4
        adj_cap  = -(drift_bps / 10000) / current_avg_cap
        adjusted = ppu * (1 + adj_yb + adj_size + adj_cap)

Usage:
    python3 comp_grid_crosscheck.py selection.json
    python3 comp_grid_crosscheck.py selection.json --sensitivity 807 824.6 850
    python3 comp_grid_crosscheck.py selection.json --json crosscheck.json
"""
import argparse
import json
import statistics


def grid_rows(sel, subj_sf=None):
    """Recompute the grid for the selected comps. subj_sf overrides the subject's
    average unit size (that is the whole point of the sensitivity run)."""
    s = sel["subject"]
    subj_sf = subj_sf if subj_sf is not None else s["avg_size"]
    cap = sel["cap_rate"]["current_avg_cap"]
    out = []
    for c in sel["comps"]:
        if c["Rank"] not in sel["selected_ranks"]:
            continue
        ppu = c["Sold Price"] / c["Unit Count"]
        sf = c["Avg Unit SF"]
        adj_yb = (c["Year Built"] - s["year_built"]) / -100 / 2
        adj_size = -(sf / subj_sf - 1) / 4
        adj_cap = -(c.get("CapRateDriftBps", 0) / 10000) / cap if cap else 0.0
        out.append({
            "name": c["Property Name"], "units": c["Unit Count"],
            "year_built": c["Year Built"], "avg_sf": sf, "ppu": ppu,
            "psf": ppu / sf, "adj_yb": adj_yb, "adj_size": adj_size,
            "adj_cap": adj_cap, "adjusted": ppu * (1 + adj_yb + adj_size + adj_cap),
        })
    if not out:
        raise SystemExit("ERROR: selection.json has no selected comps")
    return out


def analyze(sel, sensitivities=None):
    s = sel["subject"]
    subj_sf, units = s["avg_size"], s["units"]
    rows = grid_rows(sel)

    adjusted = statistics.mean(r["adjusted"] for r in rows)
    unadj = statistics.mean(r["ppu"] for r in rows)
    median_ppu = statistics.median(r["ppu"] for r in rows)
    mean_psf = statistics.mean(r["psf"] for r in rows)
    median_psf = statistics.median(r["psf"] for r in rows)

    comp_yb = statistics.mean(r["year_built"] for r in rows)
    comp_sf = statistics.mean(r["avg_sf"] for r in rows)
    # what the vintage + size gaps alone predict the divergence should be
    expected = ((comp_yb - s["year_built"]) / -100 / 2) + (-(comp_sf / subj_sf - 1) / 4) \
        + statistics.mean(r["adj_cap"] for r in rows)
    observed = adjusted / unadj - 1

    bases = {
        "adjusted_grid": adjusted,
        "unadjusted_mean_ppu": unadj,
        "median_ppu": median_ppu,
        "mean_psf_x_subject_sf": mean_psf * subj_sf,
        "median_psf_x_subject_sf": median_psf * subj_sf,
    }
    sens = {}
    for sf in (sensitivities or []):
        sens[sf] = statistics.mean(r["adjusted"] for r in grid_rows(sel, sf))

    return {
        "subject": s, "rows": rows, "bases": bases,
        "totals": {k: v * units for k, v in bases.items()},
        "dispersion": {
            "raw_sd": statistics.pstdev([r["ppu"] for r in rows]),
            "raw_cv": statistics.pstdev([r["ppu"] for r in rows]) / unadj,
            "adj_sd": statistics.pstdev([r["adjusted"] for r in rows]),
            "adj_cv": statistics.pstdev([r["adjusted"] for r in rows]) / adjusted,
        },
        "divergence": {"observed": observed, "expected_from_gaps": expected,
                       "unexplained": observed - expected,
                       "comp_avg_year_built": comp_yb, "comp_avg_sf": comp_sf},
        "sensitivity": sens,
        "cap_rate": sel.get("cap_rate", {}),
    }


def render(a):
    s, d, dis = a["subject"], a["divergence"], a["dispersion"]
    L = [f"SUBJECT: {s.get('name','(subject)')} — {s['units']} units, built "
         f"{s['year_built']}, avg {s['avg_size']:,.0f} SF", ""]
    L.append(f"{'Comp':<26}{'Units':>6}{'Built':>7}{'Avg SF':>8}{'$/Unit':>11}"
             f"{'$/SF':>9}{'Yr adj':>9}{'Size adj':>10}{'Cap adj':>9}{'Adjusted':>11}")
    for r in a["rows"]:
        L.append(f"{r['name'][:25]:<26}{r['units']:>6}{r['year_built']:>7}"
                 f"{r['avg_sf']:>8,.0f}{r['ppu']:>11,.0f}{r['psf']:>9,.2f}"
                 f"{r['adj_yb']:>+8.1%}{r['adj_size']:>+10.1%}{r['adj_cap']:>+9.1%}"
                 f"{r['adjusted']:>11,.0f}")
    L += ["", "VALUE INDICATIONS (per unit / total)"]
    label = {"adjusted_grid": "Adjusted grid indication",
             "unadjusted_mean_ppu": "Unadjusted mean $/unit",
             "median_ppu": "Median $/unit",
             "mean_psf_x_subject_sf": "Mean $/SF x subject SF",
             "median_psf_x_subject_sf": "Median $/SF x subject SF"}
    for k, v in a["bases"].items():
        L.append(f"  {label[k]:<32}${v:>10,.0f}    ${a['totals'][k]:>14,.0f}")
    L += ["", "DISPERSION (adjustments should COMPRESS this — if they widen it, the "
          "comp set or the adjustments are wrong)",
          f"  raw $/unit      SD ${dis['raw_sd']:>9,.0f}  ({dis['raw_cv']:.1%} of mean)",
          f"  adjusted $/unit SD ${dis['adj_sd']:>9,.0f}  ({dis['adj_cv']:.1%} of mean)",
          "", "DIVERGENCE CHECK (cma-subject-inputs-trap-8-2026.md)",
          f"  comps average {d['comp_avg_year_built']:.1f} vintage vs subject "
          f"{s['year_built']}, and {d['comp_avg_sf']:,.0f} SF vs subject "
          f"{s['avg_size']:,.0f} SF",
          f"  observed divergence (adjusted vs unadjusted) : {d['observed']:+.1%}",
          f"  expected from the vintage + size gaps alone  : {d['expected_from_gaps']:+.1%}",
          f"  unexplained residual                         : {d['unexplained']:+.1%}"]
    if abs(d["unexplained"]) > 0.02:
        L.append("  >> FLAG: the divergence is NOT explained by the comp/subject gaps. "
                 "Re-derive the subject's avg SF from the rent roll before quoting.")
    elif abs(d["observed"]) > 0.02:
        L.append("  >> OK: the divergence is large but fully explained by a comp set that "
                 "is newer and/or larger-unit than the subject. Say so when you quote it.")
    else:
        L.append("  >> OK: adjusted and unadjusted agree closely — a tight comp set.")
    if a["sensitivity"]:
        L += ["", "SUBJECT AVG-SF SENSITIVITY"]
        base = a["bases"]["adjusted_grid"]
        for sf, v in sorted(a["sensitivity"].items()):
            L.append(f"  {sf:>7,.1f} SF -> ${v:>9,.0f}/unit   ${v*s['units']:>14,.0f}"
                     f"   ({v/base-1:+.1%} vs base)")
    return "\n".join(L)


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("selection", help="selection.json from select_comps.py")
    ap.add_argument("--sensitivity", type=float, nargs="*", default=None,
                    help="subject avg-SF values to test (default: base, ±2.5%%, ±5%%)")
    ap.add_argument("--json", help="also write the full analysis here")
    args = ap.parse_args()
    sel = json.load(open(args.selection))
    base = sel["subject"]["avg_size"]
    sens = args.sensitivity if args.sensitivity is not None else \
        [round(base * m, 1) for m in (0.95, 0.975, 1.0, 1.025, 1.05)]
    a = analyze(sel, sens)
    print(render(a))
    if args.json:
        json.dump(a, open(args.json, "w"), indent=1, default=str)
        print(f"\nWrote {args.json}")


if __name__ == "__main__":
    main()
