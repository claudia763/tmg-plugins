#!/usr/bin/env python3
"""
RentCast cross-reference: verify comp property unit counts / vintages
=====================================================================
Quick verification helper for the rent-comparable-analysis workflow (Step 4).
Takes comp addresses — from a Dwellsy flat CSV, a names CSV, or the command
line — and pulls each property's county-record data from the RentCast API
(https://developers.rentcast.io), then writes a cross-reference table
comparing Dwellsy's community_unit_count / year_built against RentCast's
features.unitCount / yearBuilt, with owner and property type as identity
hints.

What RentCast can and cannot verify (tested live, 8/2026):
  - yearBuilt      often present for apartment parcels (818 Pinemont -> 1972)
  - unitCount      present only when the assessor recorded it — expect gaps
  - propertyType   "Apartment" vs "Single Family" is itself a screen: a
                   comp address returning Single Family (e.g. a unit-level
                   record) means the parcel lookup missed the complex —
                   verify that one by hand
  - marketing name NOT in county records; owner.names + subdivision are the
                   closest hints. The community's leasing name still comes
                   from listing sites (skill Step 4).

Usage
-----
  # cross-reference every unique comp address in a Dwellsy pull
  python rentcast_xref.py --csv "out/*.csv" -o xref.csv

  # or a names CSV (address_1,name,units,year_built,community + city/state/zip cols optional)
  python rentcast_xref.py --names names.csv --city Houston --state TX -o xref.csv

  # or ad-hoc
  python rentcast_xref.py -a "818 Pinemont Dr, Houston, TX, 77018" -a "4300 Rosslyn Rd, Houston, TX, 77018"

Key: --key, else RENTCAST_API_KEY env var, else the nearest .env walking up
from the working directory. Calls are metered per key — the script caps at
--limit (default 25) addresses per run and sleeps 0.6s between calls.
"""

import argparse
import csv
import glob
import json
import os
import sys
import time
import urllib.parse
import urllib.request

API = "https://api.rentcast.io/v1/properties"


def find_key(cli_key):
    if cli_key:
        return cli_key
    if os.environ.get("RENTCAST_API_KEY"):
        return os.environ["RENTCAST_API_KEY"]
    d = os.getcwd()
    while True:
        p = os.path.join(d, ".env")
        if os.path.exists(p):
            for line in open(p):
                line = line.strip()
                if line.startswith("RENTCAST_API_KEY="):
                    return line.split("=", 1)[1].strip().strip('"')
        parent = os.path.dirname(d)
        if parent == d:
            break
        d = parent
    sys.exit("ERROR: no RentCast key (--key, RENTCAST_API_KEY, or .env)")


def fetch(key, address, verbose=False):
    url = API + "?" + urllib.parse.urlencode({"address": address})
    req = urllib.request.Request(url, headers={"X-Api-Key": key})
    if verbose:
        print(f"  GET {address}", file=sys.stderr)
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            data = json.loads(resp.read().decode())
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return None          # no record for this address
        if e.code == 401:
            sys.exit("ERROR: RentCast rejected the API key (401)")
        if e.code == 429:
            sys.exit("ERROR: RentCast rate/usage limit hit (429) — stop, "
                     "don't burn the key's remaining quota")
        raise
    if isinstance(data, list):
        return data[0] if data else None
    return data or None


def gather_addresses(args):
    """Return ordered unique [(address_1, city, state, zip, dw_units, dw_year)]."""
    seen, out = set(), []

    def add(a1, city, st, zc, units="", year=""):
        k = (a1.strip().lower(), (zc or "").strip())
        if a1.strip() and k not in seen:
            seen.add(k)
            out.append((a1.strip(), city, st, zc, units, year))

    for g in args.csv or []:
        for path in sorted(glob.glob(g)):
            with open(path, newline="") as f:
                for r in csv.DictReader(f):
                    add(r.get("address_1", ""), r.get("address_city", ""),
                        r.get("address_state", "TX"), r.get("address_zip", ""),
                        r.get("community_unit_count", ""),
                        r.get("year_built", ""))
    if args.names:
        with open(args.names, newline="") as f:
            for r in csv.DictReader(f):
                add(r.get("address_1", ""), r.get("city", args.city),
                    r.get("state", args.state), r.get("zip", args.zip),
                    r.get("units", ""), r.get("year_built", ""))
    for a in args.address or []:
        parts = [p.strip() for p in a.split(",")]
        add(parts[0], parts[1] if len(parts) > 1 else args.city,
            parts[2] if len(parts) > 2 else args.state,
            parts[3] if len(parts) > 3 else args.zip)
    return out


def flag(dw, rc):
    if not dw or not rc:
        return "?"
    try:
        return "Y" if int(float(dw)) == int(float(rc)) else "N"
    except ValueError:
        return "?"


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[1])
    ap.add_argument("--csv", action="append",
                    help="Dwellsy flat CSV glob(s) — unique comp addresses")
    ap.add_argument("--names", help="names CSV (address_1,... )")
    ap.add_argument("-a", "--address", action="append",
                    help='ad-hoc "Street, City, ST, Zip" (repeatable)')
    ap.add_argument("--city", default="")
    ap.add_argument("--state", default="TX")
    ap.add_argument("--zip", default="")
    ap.add_argument("--key", help="RentCast API key")
    ap.add_argument("--limit", type=int, default=25,
                    help="max addresses to query this run (default 25 — "
                         "RentCast keys are metered)")
    ap.add_argument("-o", "--out", default="rentcast_xref.csv")
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args(argv)

    key = find_key(args.key)
    addrs = gather_addresses(args)
    if not addrs:
        sys.exit("ERROR: no addresses (use --csv, --names, or -a)")
    if len(addrs) > args.limit:
        print(f"NOTE: {len(addrs)} addresses, querying first {args.limit} "
              f"(--limit)", file=sys.stderr)
        addrs = addrs[:args.limit]

    cols = ["address", "city", "zip", "dwellsy_units", "dwellsy_year",
            "rc_property_type", "rc_unit_count", "rc_year_built",
            "rc_owner", "rc_subdivision", "units_match", "year_match",
            "note"]
    rows = []
    for i, (a1, city, st, zc, dw_u, dw_y) in enumerate(addrs):
        full = ", ".join(x for x in (a1, city, st, zc) if x)
        rec = fetch(key, full, args.verbose)
        if rec is None:
            rows.append(dict(zip(cols, [a1, city, zc, dw_u, dw_y,
                                        "", "", "", "", "", "?", "?",
                                        "no RentCast record"])))
        else:
            feats = rec.get("features") or {}
            rc_u = feats.get("unitCount", "")
            rc_y = rec.get("yearBuilt", "")
            owner = "; ".join((rec.get("owner") or {}).get("names") or [])
            ptype = rec.get("propertyType", "")
            note = ""
            if ptype and ptype not in ("Apartment", "Multi-Family"):
                note = (f"record is {ptype} — likely a unit-level parcel, "
                        f"verify by hand")
            rows.append(dict(zip(cols, [a1, city, zc, dw_u, dw_y, ptype,
                                        rc_u, rc_y, owner,
                                        rec.get("subdivision", ""),
                                        flag(dw_u, rc_u), flag(dw_y, rc_y),
                                        note])))
        if i < len(addrs) - 1:
            time.sleep(0.6)

    with open(args.out, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        w.writerows(rows)

    # console summary
    wid = max(len(r["address"]) for r in rows) + 2
    print(f"{'ADDRESS':<{wid}}{'DW U':>6}{'RC U':>6} {'U?':<3}"
          f"{'DW YR':>7}{'RC YR':>7} {'Y?':<3} TYPE / NOTE")
    for r in rows:
        print(f"{r['address']:<{wid}}{r['dwellsy_units']:>6}"
              f"{str(r['rc_unit_count']):>6} {r['units_match']:<3}"
              f"{r['dwellsy_year']:>7}{str(r['rc_year_built']):>7} "
              f"{r['year_match']:<3}"
              f"{r['rc_property_type']}{' — ' + r['note'] if r['note'] else ''}")
    print(f"\nWrote {args.out} ({len(rows)} addresses)")


if __name__ == "__main__":
    main()
