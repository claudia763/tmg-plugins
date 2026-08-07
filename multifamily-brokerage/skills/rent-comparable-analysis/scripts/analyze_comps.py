#!/usr/bin/env python3
"""Turn Dwellsy comp pulls into a property-level rent comparable analysis.

INPUT   one or more CSV reports from dwellsy_comps_lookup.py (either the raw
        three-section export or a --flat comparables table), plus the subject's
        floor plans.
OUTPUT  a printed analysis: coverage/QC flags, a property-level comp table
        (collapsed from unit-level listings), listing- vs property-weighted
        rents, subject positioning per floor plan, and an explicit list of
        communities whose name/vintage must be looked up before they are cited.

Why property-level: Dwellsy returns one row per unit listing, so a single
lease-up with 59 listings otherwise outvotes 27 small properties. Every headline
number here is computed both ways and the divergence is reported.

    python analyze_comps.py --csv "out/*.csv" \
        --subject-lat 33.16369 --subject-lon -96.37884 \
        --name "Crossroads Terrace" --units 36 \
        --plan "1BR:572:900:711:14" --plan "2BR:780:1100:753:22"

--plan is LABEL:SF:MARKET_RENT:EFFECTIVE_RENT:UNIT_COUNT (last two optional).
"""
import argparse
import csv
import glob
import math
import re
import statistics as st
import sys
from collections import Counter

EARTH_MI = 3958.7613
# Rents outside this band are data errors, not comps (a $41/mo and a $17,000/mo
# row both turned up in live pulls).
SANE_RENT = (400.0, 6000.0)


# --------------------------------------------------------------------------
# loading
# --------------------------------------------------------------------------

def load_csv(path):
    """Read a Dwellsy export, sectioned or flat, into dicts."""
    with open(path, encoding="utf-8-sig", newline="") as handle:
        rows = list(csv.reader(handle))
    start = next((i for i, r in enumerate(rows) if r and r[0].strip() == "id"), None)
    if start is None:
        return []
    header = [c.strip() for c in rows[start]]
    out = []
    for row in rows[start + 1:]:
        if not any(c.strip() for c in row) or len(row) != len(header):
            continue
        out.append(dict(zip(header, row)))
    return out


def as_float(value, default=0.0):
    try:
        return float(str(value).strip())
    except (TypeError, ValueError):
        return default


def as_int(value):
    text = str(value).strip()
    return int(text) if text.isdigit() else None


def haversine(lat1, lon1, lat2, lon2):
    rad = math.radians
    dlat, dlon = rad(lat2 - lat1), rad(lon2 - lon1)
    a = (math.sin(dlat / 2) ** 2
         + math.cos(rad(lat1)) * math.cos(rad(lat2)) * math.sin(dlon / 2) ** 2)
    return 2 * EARTH_MI * math.asin(math.sqrt(a))


def enrich(rows, subject):
    """Normalise types and attach distance-from-subject + a property key."""
    out, seen = [], set()
    for row in rows:
        if row.get("id") in seen:          # town sweeps overlap
            continue
        seen.add(row.get("id"))
        row["rent"] = as_float(row.get("listing_amount"))
        row["sf"] = as_float(row.get("square_feet"))
        row["psf"] = as_float(row.get("price_per_sf"))
        row["beds"] = str(row.get("bedrooms", "")).strip()
        row["units"] = as_int(row.get("community_unit_count"))
        row["year"] = as_int(row.get("year_built"))
        row["active"] = row.get("property_listing_status") == "active"
        if subject:
            row["mi"] = haversine(subject[0], subject[1],
                                  as_float(row.get("latitude")),
                                  as_float(row.get("longitude")))
        else:
            row["mi"] = as_float(row.get("distance_miles"))
        row["prop"] = (row.get("address_1", "").strip(),
                       row.get("address_city", "").strip())
        out.append(row)
    return out


def by_property(rows):
    groups = {}
    for row in rows:
        groups.setdefault(row["prop"], []).append(row)
    return groups


def med(values):
    return st.median(values) if values else None


# --------------------------------------------------------------------------
# sections
# --------------------------------------------------------------------------

def section_coverage(rows):
    print("=" * 78)
    print("COVERAGE AND DATA QUALITY")
    print("=" * 78)
    props = by_property(rows)
    active = sum(1 for r in rows if r["active"])
    print("  listings            : %d" % len(rows))
    print("  distinct addresses  : %d" % len(props))
    print("  active listings     : %d (%.0f%%)"
          % (active, 100.0 * active / len(rows)))
    print("  address types       : %s"
          % dict(Counter(r.get("address_type") for r in rows)))
    print("  bedrooms            : %s"
          % dict(sorted(Counter(r["beds"] for r in rows).items())))
    print("  distance from subj. : %.1f - %.1f mi"
          % (min(r["mi"] for r in rows), max(r["mi"] for r in rows)))

    no_year = [r for r in rows if r["year"] is None]
    if no_year:
        print("  ! blank year_built  : %d listings / %d properties -- a vintage "
              "filter drops these SILENTLY"
              % (len(no_year), len({r["prop"] for r in no_year})))
    odd = [r for r in rows if not (SANE_RENT[0] <= r["rent"] <= SANE_RENT[1])]
    if odd:
        print("  ! implausible rents : %d listings outside $%.0f-$%.0f "
              "(excluded from medians below)"
              % (len(odd), SANE_RENT[0], SANE_RENT[1]))
        for r in sorted(odd, key=lambda x: x["rent"])[:5]:
            print("      $%-8.0f %-3sbd %-6.0f sf  %s"
                  % (r["rent"], r["beds"], r["sf"], r["address_1"]))
    return props


def section_concentration(rows, props):
    print("\n" + "=" * 78)
    print("CONCENTRATION  (one row = one unit listing, not one property)")
    print("=" * 78)
    ranked = sorted(props.items(), key=lambda kv: -len(kv[1]))
    top3 = sum(len(v) for _, v in ranked[:3])
    print("  top 3 addresses hold %d of %d listings (%.0f%%)"
          % (top3, len(rows), 100.0 * top3 / len(rows)))
    for (addr, city), items in ranked[:5]:
        print("    %-28s %-14s %3d listings" % (addr[:28], city[:14], len(items)))
    if top3 / len(rows) > 0.5:
        print("  ! Listing-weighted averages are dominated by a few buildings. "
              "Quote the property-weighted figure.")


def section_comps(rows, props, min_units, beds_of_interest):
    print("\n" + "=" * 78)
    print("COMP TABLE  (communities of %d+ units, apartment type)" % min_units)
    print("=" * 78)
    comms = {k: v for k, v in props.items()
             if (v[0]["units"] or 0) >= min_units
             and v[0].get("address_type") == "apartment"}
    if not comms:
        print("  NONE. No apartment community of %d+ units in the data." % min_units)
        return comms
    head = "  %-26s %-13s %-5s %-6s %-6s" % ("address", "city", "units", "built", "mi")
    for bed in beds_of_interest:
        head += " %-9s" % ("%sbd rent" % bed)
    print(head)
    print("  " + "-" * (len(head) - 2))
    for key, items in sorted(comms.items(), key=lambda kv: kv[1][0]["mi"]):
        first = items[0]
        line = ("  %-26s %-13s %-5s %-6s %-6.1f"
                % (first["address_1"][:26], first["address_city"][:13],
                   first["units"], first["year"] or "?", first["mi"]))
        for bed in beds_of_interest:
            sub = [r["rent"] for r in items
                   if r["beds"] == bed and SANE_RENT[0] <= r["rent"] <= SANE_RENT[1]]
            line += " %-9s" % (("$%.0f" % med(sub)) if sub else "-")
        print(line)

    missing = [v[0] for v in comms.values() if v[0]["year"] is None]
    print("\n  LOOK UP BEFORE CITING (Dwellsy gives no property name, and its "
          "unit counts/vintages are unreliable on small or recent assets):")
    for key, items in sorted(comms.items(), key=lambda kv: kv[1][0]["mi"]):
        first = items[0]
        flag = "  <-- vintage MISSING" if first["year"] is None else ""
        print("    %-26s %-13s (Dwellsy: %s units, %s)%s"
              % (first["address_1"][:26], first["address_city"][:13],
                 first["units"], first["year"] or "blank", flag))
    if missing:
        print("    %d of %d need a vintage lookup." % (len(missing), len(comms)))
    return comms


def section_weighting(rows, label, beds_of_interest):
    print("\n" + "=" * 78)
    print("RENT BENCHMARKS  (%s)" % label)
    print("=" * 78)
    clean = [r for r in rows if SANE_RENT[0] <= r["rent"] <= SANE_RENT[1]]
    if not clean:
        print("  no usable listings")
        return {}
    result = {}
    for bed in beds_of_interest:
        sub = [r for r in clean if r["beds"] == bed]
        if not sub:
            continue
        listing_med = med([r["rent"] for r in sub])
        per_prop = [med([r["rent"] for r in items])
                    for items in by_property(sub).values()]
        prop_med = med(per_prop)
        result[bed] = {"listing": listing_med, "property": prop_med,
                       "sf": med([r["sf"] for r in sub if r["sf"]]),
                       "psf": med([r["psf"] for r in sub if r["psf"]]),
                       "n": len(sub), "props": len(per_prop)}
        gap = (100.0 * (listing_med - prop_med) / prop_med) if prop_med else 0.0
        print("  %sbd  n=%-4d props=%-3d  listing-wtd $%-7.0f  property-wtd "
              "$%-7.0f  (%+.0f%%)  med sf %-6.0f psf %.2f"
              % (bed, len(sub), len(per_prop), listing_med, prop_med, gap,
                 result[bed]["sf"] or 0, result[bed]["psf"] or 0))
    return result


def section_subject(plans, benchmarks, size_benchmarks):
    if not plans:
        return
    print("\n" + "=" * 78)
    print("SUBJECT POSITIONING")
    print("=" * 78)
    print("  %-8s %-7s %-9s %-9s %-9s %-9s"
          % ("plan", "sf", "market", "net eff", "mkt psf", "eff psf"))
    for plan in plans:
        mkt = plan["market"]
        eff = plan["effective"]
        print("  %-8s %-7.0f %-9s %-9s %-9s %-9s"
              % (plan["label"], plan["sf"],
                 "$%.0f" % mkt if mkt else "-",
                 "$%.0f" % eff if eff else "-",
                 "$%.2f" % (mkt / plan["sf"]) if mkt and plan["sf"] else "-",
                 "$%.2f" % (eff / plan["sf"]) if eff and plan["sf"] else "-"))
        if mkt and eff and mkt > eff:
            print("           loss to lease: %.0f%%" % (100.0 * (mkt - eff) / mkt))

    for source, title in ((benchmarks, "vs all comps"),
                          (size_benchmarks, "vs size-matched comps")):
        if not source:
            continue
        print("\n  %s (property-weighted medians):" % title)
        for plan in plans:
            stats = source.get(plan["beds"])
            if not stats or not stats["property"]:
                print("    %-8s no comp sample" % plan["label"])
                continue
            base = stats["property"]
            bits = []
            if plan["market"]:
                bits.append("market %.0f%%" % (100.0 * plan["market"] / base))
            if plan["effective"]:
                bits.append("net effective %.0f%%" % (100.0 * plan["effective"] / base))
            print("    %-8s comp median $%-7.0f -> subject %s"
                  % (plan["label"], base, ", ".join(bits) or "n/a"))
            if stats["psf"] and plan["sf"]:
                implied = stats["psf"] * plan["sf"]
                print("             comp psf $%.2f x %.0f sf implies $%.0f/mo"
                      % (stats["psf"], plan["sf"], implied))


# --------------------------------------------------------------------------
# cli
# --------------------------------------------------------------------------

def parse_plan(text):
    parts = text.split(":")
    if len(parts) < 2:
        raise argparse.ArgumentTypeError(
            "--plan needs LABEL:SF[:MARKET[:EFFECTIVE[:COUNT]]]")
    label = parts[0].strip()
    m = re.match(r"\d+", label)          # leading digits only: "1x1 Classic"
    beds = m.group(0) if m else "0"       # parses as 1 bed, not 11
    return {"label": label, "beds": beds, "sf": float(parts[1]),
            "market": float(parts[2]) if len(parts) > 2 and parts[2] else None,
            "effective": float(parts[3]) if len(parts) > 3 and parts[3] else None,
            "count": int(parts[4]) if len(parts) > 4 and parts[4] else None}


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--csv", action="append", required=True,
                    help="CSV export (repeatable; globs allowed)")
    ap.add_argument("--subject-lat", type=float)
    ap.add_argument("--subject-lon", type=float)
    ap.add_argument("--name", default="subject")
    ap.add_argument("--units", type=int, help="subject unit count")
    ap.add_argument("--plan", action="append", type=parse_plan, default=[],
                    help="LABEL:SF[:MARKET[:EFFECTIVE[:COUNT]]] (repeatable)")
    ap.add_argument("--min-units", type=int, default=20,
                    help="floor for 'a real community' in the comp table (default 20)")
    ap.add_argument("--size-tolerance", type=float, default=0.35,
                    help="size-matched band around each plan's SF (default 0.35)")
    args = ap.parse_args(argv)

    paths = []
    for pattern in args.csv:
        hits = sorted(glob.glob(pattern))
        paths.extend(hits or [pattern])
    rows = []
    for path in paths:
        loaded = load_csv(path)
        print("  loaded %-52s %4d listings" % (path.split("\\")[-1], len(loaded)))
        rows.extend(loaded)
    if not rows:
        print("\nNo comparables in any input file -> report \"Can't find comps\".")
        return 1

    subject = None
    if args.subject_lat is not None and args.subject_lon is not None:
        subject = (args.subject_lat, args.subject_lon)
    rows = enrich(rows, subject)
    print("\nSubject: %s%s" % (args.name,
                               " (%d units)" % args.units if args.units else ""))

    beds = [p["beds"] for p in args.plan] or \
        sorted({r["beds"] for r in rows if r["beds"].isdigit()}, key=int)[:3]

    props = section_coverage(rows)
    section_concentration(rows, props)
    comms = section_comps(rows, props, args.min_units, beds)

    comm_rows = [r for items in comms.values() for r in items]
    benchmarks = section_weighting(comm_rows or rows,
                                   "communities of %d+ units" % args.min_units
                                   if comm_rows else "ALL listings (no qualifying "
                                                     "community -- weak basis)",
                                   beds)
    size_benchmarks = {}
    if args.plan:
        matched = []
        for row in rows:
            for plan in args.plan:
                if row["beds"] != plan["beds"] or not row["sf"]:
                    continue
                lo = plan["sf"] * (1 - args.size_tolerance)
                hi = plan["sf"] * (1 + args.size_tolerance)
                if lo <= row["sf"] <= hi:
                    matched.append(row)
                    break
        if matched:
            size_benchmarks = section_weighting(
                matched, "size-matched to the subject's floor plans "
                         "(+/-%.0f%% SF)" % (args.size_tolerance * 100), beds)
        else:
            print("\n  ! No comp unit falls within +/-%.0f%% of the subject's "
                  "unit sizes -- the subject's product does not exist in this data."
                  % (args.size_tolerance * 100))

    section_subject(args.plan, benchmarks, size_benchmarks)

    print("\n" + "=" * 78)
    print("Before this reaches a deck: look up each community's NAME and confirm "
          "its vintage\nand unit count from a listing site or county record. "
          "Dwellsy supplies neither\nreliably -- company_name is the management "
          "company, not the property.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
