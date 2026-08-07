#!/usr/bin/env python3
"""
trim_comps_by_distance.py -- submarket-trim a Dwellsy flat comps CSV before it
reaches rent_comp_export.py.

WHEN TO USE
    The rent-comparable-analysis house rule requires every API pull to be a
    MINIMUM 3-mile radius, and rent_comp_export.py has no distance filter -- so
    in a dense metro the exported grid can reach across a submarket boundary
    into materially stronger product. references/dwellsy-api.md is explicit
    that this trim is ours to make: "The API returns what is nearby, not what
    is comparable."

    Use this to cut the pull down to the subject's own submarket while leaving
    every other house rule (min-units, vintage window, SD pruning, recency
    ladder) to the exporter. ALWAYS state the radius used in the deliverable.

INPUT   one or more Dwellsy flat CSVs (the --flat output of
        dwellsy_comps_lookup.py). Rows must carry `distance_miles`, or
        `latitude`/`longitude` plus --subject-lat/--subject-lon.

OUTPUT  --out  a flat CSV in the identical 24-column schema, so the whole
        downstream toolchain runs unmodified.

    python trim_comps_by_distance.py --csv "out/pull*.csv" --max-miles 1.0 \
        --out out/comps_submarket.csv --subject-lat 35.49646 --subject-lon -97.54705

NOTE
    Distances in a multi-pull submarket sweep are measured from EACH pull's own
    centre, so pass --subject-lat/--subject-lon to recompute them from the
    subject; without it the file's own distance_miles is trusted as-is.
"""
import argparse
import csv
import glob
import math
import sys

EARTH_MI = 3958.8


def haversine(a, b, c, d):
    r = math.radians
    dlat, dlon = r(c - a), r(d - b)
    h = (math.sin(dlat / 2) ** 2
         + math.cos(r(a)) * math.cos(r(c)) * math.sin(dlon / 2) ** 2)
    return 2 * EARTH_MI * math.asin(math.sqrt(h))


def num(v):
    try:
        return float(str(v).strip())
    except (TypeError, ValueError):
        return None


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[1])
    ap.add_argument("--csv", action="append", required=True,
                    help="glob of Dwellsy flat CSVs (repeatable)")
    ap.add_argument("--out", required=True)
    ap.add_argument("--max-miles", type=float, required=True)
    ap.add_argument("--subject-lat", type=float)
    ap.add_argument("--subject-lon", type=float)
    args = ap.parse_args()

    paths = sorted({p for g in args.csv for p in glob.glob(g)})
    if not paths:
        sys.exit("ERROR: no CSVs matched")

    rows, header, seen = [], None, set()
    for p in paths:
        with open(p, newline="") as f:
            rd = csv.DictReader(f)
            header = header or rd.fieldnames
            for r in rd:
                if r.get("id") in seen:          # de-dup overlapping pulls
                    continue
                seen.add(r.get("id"))
                rows.append(r)

    kept, dropped = [], 0
    for r in rows:
        d = num(r.get("distance_miles"))
        if args.subject_lat is not None and args.subject_lon is not None:
            la, lo = num(r.get("latitude")), num(r.get("longitude"))
            if la is not None and lo is not None:
                d = haversine(args.subject_lat, args.subject_lon, la, lo)
                r["distance_miles"] = round(d, 4)
        if d is None or d > args.max_miles:
            dropped += 1
            continue
        kept.append(r)

    with open(args.out, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=header)
        w.writeheader()
        w.writerows(kept)

    addrs = len({r["address_1"] for r in kept})
    print(f"read {len(rows)} listings from {len(paths)} file(s); "
          f"kept {len(kept)} within {args.max_miles} mi "
          f"({addrs} addresses), dropped {dropped}")
    print(f"wrote {args.out}")


if __name__ == "__main__":
    main()
