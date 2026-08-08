#!/usr/bin/env python3
"""Drop rent-restricted (LIHTC / project-based / age-restricted) properties from
Dwellsy flat CSVs before the rent-comp exporter runs.

WHY: a tax-credit property's asking rents are capped by AMI limits, not set by
the market, so it cannot price a conventional market-rate subject. Dwellsy has
no subsidy flag and the exporter has no exclude switch, so the screen has to
happen on the input. The comparable trap on the SALE side of this same deal was
Arbor Pointe (600 N Oats St) — an ADFA-financed, age-restricted property whose
2006 vintage carried the LOWEST $/unit in the set. The rent side has the same
hazard, and a generic name ("Aspen Grove") defeats any keyword screen.

The authority is the STATE housing finance agency's own property list, not a
listing portal:
  Arkansas — ADFA county lists,
    https://adfa.arkansas.gov/wp-content/uploads/2024/12/<County>-County-<City>.pdf
  Texas    — TDHCA property inventory

Usage:
  python3 filter_restricted.py --csv "out/*.csv" --outdir out_screened \\
      --exclude "1919 E 18th St=Aspen Grove Apartments (ADFA/LIHTC)"

Inputs : Dwellsy flat CSVs (24-column schema)
Outputs: the same CSVs with the excluded addresses removed, plus a printed log
         of exactly what came out and why (paste it into the delivery notes).
"""
import argparse
import csv
import glob
import os


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--csv", required=True, action="append",
                    help="glob of Dwellsy flat CSVs (repeatable)")
    ap.add_argument("--outdir", required=True)
    ap.add_argument("--exclude", action="append", default=[],
                    help='"<address_1>=<reason>" (repeatable); address matched '
                         'case-insensitively on the address_1 column')
    a = ap.parse_args()

    reasons = {}
    for e in a.exclude:
        addr, _, why = e.partition("=")
        reasons[addr.strip().lower()] = why.strip() or "rent-restricted"

    os.makedirs(a.outdir, exist_ok=True)
    files = sorted({p for g in a.csv for p in glob.glob(g)})
    if not files:
        raise SystemExit(f"no CSVs matched {a.csv}")

    total_in = total_out = 0
    dropped = {}
    for path in files:
        with open(path, newline="") as fh:
            rd = csv.DictReader(fh)
            rows = list(rd)
            hdr = rd.fieldnames
        keep = []
        for r in rows:
            key = (r.get("address_1") or "").strip().lower()
            if key in reasons:
                dropped.setdefault(r.get("address_1"), [reasons[key], 0])
                dropped[r.get("address_1")][1] += 1
                continue
            keep.append(r)
        out = os.path.join(a.outdir, os.path.basename(path))
        with open(out, "w", newline="") as fh:
            w = csv.DictWriter(fh, fieldnames=hdr)
            w.writeheader()
            w.writerows(keep)
        total_in += len(rows)
        total_out += len(keep)
        print(f"  {os.path.basename(path)}: {len(rows)} -> {len(keep)} rows")

    print(f"\nscreened {total_in} listings -> {total_out} "
          f"({total_in - total_out} removed)")
    for addr, (why, n) in sorted(dropped.items()):
        print(f"  EXCLUDED {addr} — {n} listings — {why}")
    for addr in reasons:
        if not any(k.lower() == addr for k in dropped):
            print(f"  [note] no listings matched '{addr}' — check the spelling")


if __name__ == "__main__":
    main()
