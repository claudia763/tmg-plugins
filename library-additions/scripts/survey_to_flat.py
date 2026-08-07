#!/usr/bin/env python3
"""
survey_to_flat.py -- convert a hand-built rent-comp listing survey into the
Dwellsy "flat" CSV schema that rent_comp_export.py / analyze_comps.py consume.

WHEN TO USE
    When the Dwellsy Comps API is unavailable (no DWELLSY_API_KEY on the host,
    outage, or a market Dwellsy does not cover) and the rent comp set has to be
    assembled from public listing sources instead. Writing the survey into the
    Dwellsy schema means the whole downstream TMG toolchain -- house-rule
    filtering, property-level collapsing, the branded PDF and the Excel grid --
    runs unmodified, so the deliverable is identical in form to an API-sourced
    one.

INPUT   a survey CSV, one row per advertised floor plan, with columns:
            property_name, street_address, city, state, zip, latitude,
            longitude, year_built, total_units, plan_label, bedrooms,
            bathrooms, square_feet, asking_rent, utilities_included,
            amenities, listing_status, source_url
        Extra columns are ignored. Blank square_feet / year_built / total_units
        are passed through as blank -- never invent them.

OUTPUT  --out       flat CSV in the 24-column Dwellsy comparables schema
        --names-out names CSV (address_1,name,units,year_built,community) --
                    the Step-4 "verified identities" file the exporter merges
                    multi-address communities with. Every property surveyed by
                    hand is verified by definition, so this is written for you.

    python survey_to_flat.py --survey survey.csv --out out/comps_flat.csv \
        --names-out names.csv --survey-date 2026-08-07

NOTE ON DATES
    Every row is stamped with --survey-date and property_listing_status is
    taken from the survey ("active" unless the survey says otherwise). A survey
    is a point-in-time snapshot, so the export has NO usable time dimension:
    run rent_comp_export.py with --methodology to replace the trending page,
    and with --source-label so the deliverable names the real source.
"""
import argparse
import csv
import hashlib
import math

FLAT_COLS = [
    "id", "address_type", "latitude", "longitude", "address_1", "address_2",
    "address_city", "address_state", "address_zip", "zip_plus4", "bedrooms",
    "bathrooms", "square_feet", "year_built", "company_name", "listing_amount",
    "price_per_sf", "last_listing_creation_time",
    "last_listing_deactivation_time", "property_listing_status",
    "distance_miles", "url", "amenities", "community_unit_count",
]
EARTH_MI = 3958.8


def haversine(a, b, c, d):
    r = math.radians
    dlat, dlon = r(c - a), r(d - b)
    h = (math.sin(dlat / 2) ** 2
         + math.cos(r(a)) * math.cos(r(c)) * math.sin(dlon / 2) ** 2)
    return 2 * EARTH_MI * math.asin(math.sqrt(h))


def num(v):
    try:
        return float(str(v).replace("$", "").replace(",", "").strip())
    except (TypeError, ValueError):
        return None


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[1])
    ap.add_argument("--survey", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--names-out")
    ap.add_argument("--survey-date", required=True,
                    help="YYYY-MM-DD the listings were observed")
    ap.add_argument("--subject-lat", type=float)
    ap.add_argument("--subject-lon", type=float)
    args = ap.parse_args()

    rows = list(csv.DictReader(open(args.survey, newline="")))
    out, names, seen_addr = [], {}, set()
    for i, s in enumerate(rows):
        addr = (s.get("street_address") or "").strip()
        if not addr:
            continue
        rent, sf = num(s.get("asking_rent")), num(s.get("square_feet"))
        lat, lon = num(s.get("latitude")), num(s.get("longitude"))
        mi = ""
        if lat and lon and args.subject_lat is not None:
            mi = round(haversine(args.subject_lat, args.subject_lon, lat, lon), 4)
        # stable synthetic id so re-runs de-duplicate identically
        rid = hashlib.sha1(
            f'{addr}|{s.get("plan_label")}|{s.get("bedrooms")}|{sf}|{rent}'
            .encode()).hexdigest()[:16]
        out.append({
            "id": rid,
            "address_type": "apartment",
            "latitude": lat or "", "longitude": lon or "",
            "address_1": addr, "address_2": "",
            "address_city": (s.get("city") or "").strip(),
            "address_state": (s.get("state") or "").strip(),
            "address_zip": (s.get("zip") or "").strip(),
            "zip_plus4": "",
            "bedrooms": (s.get("bedrooms") or "").strip(),
            "bathrooms": (s.get("bathrooms") or "").strip(),
            "square_feet": int(sf) if sf else "",
            "year_built": (s.get("year_built") or "").strip(),
            "company_name": "",
            "listing_amount": int(rent) if rent else "",
            "price_per_sf": round(rent / sf, 4) if (rent and sf) else "",
            "last_listing_creation_time": args.survey_date,
            "last_listing_deactivation_time": "",
            "property_listing_status": (s.get("listing_status")
                                        or "active").strip() or "active",
            "distance_miles": mi,
            "url": (s.get("source_url") or "").strip(),
            "amenities": (s.get("amenities") or "").strip(),
            "community_unit_count": (s.get("total_units") or "").strip(),
        })
        if addr not in seen_addr:
            seen_addr.add(addr)
            names[addr] = {
                "address_1": addr,
                "name": (s.get("property_name") or "").strip(),
                "units": (s.get("total_units") or "").strip(),
                "year_built": (s.get("year_built") or "").strip(),
                "community": (s.get("property_name") or "").strip(),
            }

    with open(args.out, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=FLAT_COLS)
        w.writeheader()
        w.writerows(out)
    print(f"wrote {len(out)} listing rows -> {args.out}")

    if args.names_out:
        with open(args.names_out, "w", newline="") as f:
            w = csv.DictWriter(
                f, fieldnames=["address_1", "name", "units", "year_built",
                               "community"])
            w.writeheader()
            w.writerows(names.values())
        print(f"wrote {len(names)} property identities -> {args.names_out}")


if __name__ == "__main__":
    main()
