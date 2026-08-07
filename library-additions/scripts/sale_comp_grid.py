#!/usr/bin/env python3
"""Re-adjust a TMG sale-comparable grid for a different subject property.

WHAT / WHEN
    TMG's underwriting model prints a "Comparable Sale Grid" that adjusts each comp's
    $/unit for the SUBJECT's vintage and average unit size. Because both adjustments key
    off the subject, the same comp set produces a DIFFERENT indicated value for a
    different subject -- which is exactly what happens when a two-phase asset previously
    underwritten as one property has to be split into two independent underwritings
    (Westlake East/West, Lubbock, 8/2026).

    Use this to reuse an existing, already-approved comp set against a new subject
    without rebuilding the grid by hand, and to prove the re-adjusted grid still sums
    back to the original when the phases are recombined.

THE FORMULA (reverse-engineered from TMG's printed output and confirmed to the cent
against the Westlake Apartments 8/6/2026 UW Summary -- see the self-test at the bottom):

    year_adj  = (subject_year_built - comp_year_built) * 0.005
    size_adj  = (subject_avg_sf - comp_avg_sf) / subject_avg_sf * 0.25
    drift_adj = 0.0                      # printed in the model but zeroed in practice
    adjusted_price_per_unit = comp_price_per_unit * (1 + year_adj + size_adj + drift_adj)
    indicated_value_per_unit = simple average of the adjusted $/unit (comps equally weighted)

    Note the size adjustment divides by the SUBJECT's average SF, not the comp's. Dividing
    by the comp's SF reproduces small comps badly (Farrar: 16.3% instead of the printed
    9.87%) and is the easy mistake to make here.

INPUT   a JSON file (or the COMPS constant below) shaped as:
            {"subject": {"name":..., "year_built":1976, "avg_sf":1122, "units":59},
             "comps": [{"name":..., "sale_price":3560000, "units":70,
                        "year_built":1965, "avg_sf":899, "sale_date":"4/4/2025"}, ...]}
OUTPUT  the adjusted grid as a table on stdout, plus JSON with --json

Usage:
    python3 sale_comp_grid.py grid.json
    python3 sale_comp_grid.py grid.json --json out.json
    python3 sale_comp_grid.py --self-test
"""
import argparse
import json
import sys


def adjust(subject, comps, drift_adj=0.0):
    """Return the adjusted grid rows plus the indicated value for `subject`."""
    rows = []
    for c in comps:
        ppu = c["sale_price"] / c["units"]
        year_adj = (subject["year_built"] - c["year_built"]) * 0.005
        size_adj = (subject["avg_sf"] - c["avg_sf"]) / subject["avg_sf"] * 0.25
        total_adj = year_adj + size_adj + drift_adj
        rows.append(
            {
                "name": c["name"],
                "sale_date": c.get("sale_date", ""),
                "units": c["units"],
                "sale_price": c["sale_price"],
                "price_per_unit": ppu,
                "year_adj": year_adj,
                "size_adj": size_adj,
                "total_adj": total_adj,
                "dollar_adj": ppu * total_adj,
                "adjusted_price_per_unit": ppu * (1 + total_adj),
            }
        )
    indicated_ppu = sum(r["adjusted_price_per_unit"] for r in rows) / len(rows)
    return {
        "subject": subject,
        "rows": rows,
        "indicated_price_per_unit": indicated_ppu,
        "indicated_total_value": indicated_ppu * subject["units"],
    }


def render(result):
    s = result["subject"]
    out = [
        f"SUBJECT: {s.get('name','(subject)')} — {s['units']} units, built {s['year_built']}, "
        f"avg {s['avg_sf']:,.0f} SF",
        "",
        f"{'Comparable':<18}{'Sold':>11}{'Units':>7}{'$/Unit':>11}"
        f"{'Yr adj':>9}{'Size adj':>10}{'Total':>9}{'Adj $/Unit':>13}",
        "-" * 88,
    ]
    for r in result["rows"]:
        out.append(
            f"{r['name']:<18}{r['sale_date']:>11}{r['units']:>7}{r['price_per_unit']:>11,.0f}"
            f"{r['year_adj']:>8.1%}{r['size_adj']:>10.2%}{r['total_adj']:>9.2%}"
            f"{r['adjusted_price_per_unit']:>13,.0f}"
        )
    out += [
        "-" * 88,
        f"Indicated value / unit : ${result['indicated_price_per_unit']:,.0f}",
        f"Indicated total value  : ${result['indicated_total_value']:,.0f}",
    ]
    return "\n".join(out)


# The Westlake comp set, used as the regression fixture. Source: "Westlake Apartments -
# UW Summary.pdf" p.5, TMG's own combined model dated 8/6/2026.
_WESTLAKE_COMPS = [
    {"name": "Aspen Village", "sale_price": 3_560_000, "units": 70, "year_built": 1965,
     "avg_sf": 899, "sale_date": "4/4/2025"},
    {"name": "Farrar", "sale_price": 7_720_000, "units": 135, "year_built": 1981,
     "avg_sf": 593, "sale_date": "4/30/2026"},
    {"name": "Parkside", "sale_price": 10_000_000, "units": 171, "year_built": 1972,
     "avg_sf": 871, "sale_date": "7/24/2024"},
    {"name": "Raiders Walk", "sale_price": 12_500_000, "units": 196, "year_built": 1975,
     "avg_sf": 770, "sale_date": "10/24/2025"},
    {"name": "Casa Orlando", "sale_price": 4_370_000, "units": 70, "year_built": 1972,
     "avg_sf": 790, "sale_date": "12/11/2024"},
]


def self_test():
    """Reproduce TMG's printed combined grid, then split it into the two phases."""
    combined = adjust(
        {"name": "Westlake Apartments", "year_built": 1973, "avg_sf": 980, "units": 174},
        _WESTLAKE_COMPS,
    )
    # Printed on p.5 of the 8/6/2026 UW Summary.
    expected_adjusted = [53_942, 60_543, 60_398, 66_554, 65_767]
    expected_dollar_adj_aspen = 3_085.16
    expected_ppu, expected_total = 61_441, 10_690_719

    print(render(combined))
    print()
    ok = True
    for row, exp in zip(combined["rows"], expected_adjusted):
        got = round(row["adjusted_price_per_unit"])
        flag = "OK " if abs(got - exp) <= 1 else "FAIL"
        ok &= flag == "OK "
        print(f"  {flag} {row['name']:<16} adjusted $/unit {got:>8,} (printed {exp:,})")
    aspen = combined["rows"][0]["dollar_adj"]
    flag = "OK " if abs(aspen - expected_dollar_adj_aspen) < 0.01 else "FAIL"
    ok &= flag == "OK "
    print(f"  {flag} Aspen Village $/unit adjustment {aspen:,.2f} (printed {expected_dollar_adj_aspen:,.2f})")
    for label, got, exp in (
        ("indicated $/unit", round(combined["indicated_price_per_unit"]), expected_ppu),
        ("indicated total", round(combined["indicated_total_value"]), expected_total),
    ):
        flag = "OK " if abs(got - exp) <= 25 else "FAIL"
        ok &= flag == "OK "
        print(f"  {flag} {label:<16} {got:>12,} (printed {exp:,})")

    print("\n--- Re-adjusted for the two phases ---\n")
    east = adjust({"name": "Westlake East", "year_built": 1976, "avg_sf": 1122, "units": 59},
                  _WESTLAKE_COMPS)
    west = adjust({"name": "Westlake West", "year_built": 1973, "avg_sf": 907, "units": 115},
                  _WESTLAKE_COMPS)
    for r in (east, west):
        print(render(r), "\n")
    recombined = east["indicated_total_value"] + west["indicated_total_value"]
    print(f"East + West = ${recombined:,.0f} vs combined-subject grid "
          f"${combined['indicated_total_value']:,.0f} "
          f"({recombined / combined['indicated_total_value'] - 1:+.2%})")
    print("\nSELF-TEST:", "PASS" if ok else "FAIL")
    return 0 if ok else 1


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("grid", nargs="?", help="JSON file with {subject, comps}")
    ap.add_argument("--json", help="write the result to this JSON path")
    ap.add_argument("--drift", type=float, default=0.0,
                    help="cap-rate drift adjustment as a decimal (TMG prints it but zeroes it)")
    ap.add_argument("--self-test", action="store_true",
                    help="reproduce the Westlake 8/6/2026 grid as a regression check")
    args = ap.parse_args()

    if args.self_test:
        return self_test()
    if not args.grid:
        ap.error("give a grid JSON path, or --self-test")

    with open(args.grid) as fh:
        spec = json.load(fh)
    result = adjust(spec["subject"], spec["comps"], drift_adj=args.drift)
    print(render(result))
    if args.json:
        with open(args.json, "w") as fh:
            json.dump(result, fh, indent=2)
        print(f"\nwrote {args.json}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
