"""Dwellsy Comps API lookup -- pull rent comps for a subject property.

WHY: rent comps for BOVs/OMs/underwriting otherwise get gathered by hand. The
Dwellsy Comps API (https://comps-api.dwellsy.com/doc) returns closed/listed
rental comps around a point, filtered by bed count, unit size, radius and
lookback window, as JSON or CSV. This wraps the three endpoints in that spec:

  POST /comps          generate a report          (Bearer auth)
  GET  /comps/history  recent reports for the key (Bearer auth)
  GET  /comps/{id}     re-download a past report

USAGE
  # comps by address, CSV to disk (--flat = spreadsheet-ready, see below)
  python dwellsy_comps_lookup.py lookup \
      --address "770 S 2780 E Street, Saint George, UT" \
      --beds 1-3 --radius 2 --months 6 --type apartment house \
      --format csv --flat -o comps_saint_george.csv

  # comps by lat/long, printed as a table
  python dwellsy_comps_lookup.py lookup --lat 38.9764706 --lon -94.5718538 \
      --beds 1-4

  # straight to S3 (bucket side is configured on the Dwellsy account)
  python dwellsy_comps_lookup.py lookup --lat 37.433695 --lon -122.135618 \
      --beds 2 --type house --format s3 --s3-filename comps/palo_alto_2bd.csv

  python dwellsy_comps_lookup.py history
  python dwellsy_comps_lookup.py report <request-id> -o rerun.csv

PULL BUDGET: every `lookup` bills a report, so one analysis may spend at most
6 (override with --budget or DWELLSY_PULL_BUDGET). The count persists in a
temp-file counter and the 7th lookup exits 3 without calling the API. Start
each new analysis with `budget --reset`; check spend with `budget`. When the
budget is gone, build the analysis from the reports already pulled or reply
"Can't find comps" -- do not keep re-querying.

NOTE ON THE CSV: the export is a three-section report (request echo, summary
block, then the comparables table), so it does not open as a plain spreadsheet.
`--flat` writes only the comparables table. One row = one unit listing, not one
property, and rent is the `listing_amount` column. See
instructions/dwellsy-comps-api.md for the field list and the weighting/outlier
caveats before averaging anything.

AUTH: the API key (a bearer token) is read, in order, from
  1. the DWELLSY_API_KEY environment variable,
  2. --env-file <path>, or
  3. the nearest .env walking up from the working directory / this script.
On the cowork server that means the repo-root .env, which is gitignored. Never
put the key in this folder -- everything under library-additions/ is pushed to
the public tmg-plugins repo. See instructions/dwellsy-comps-api.md.

Requires: Python 3.8+, standard library only.
"""
import argparse
import csv
import http.client
import io
import json
import os
import socket
import ssl
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

DEFAULT_BASE_URL = "https://comps-api.dwellsy.com"
API_KEY_VAR = "DWELLSY_API_KEY"
BASE_URL_VAR = "DWELLSY_API_BASE"

ADDRESS_TYPES = ("apartment", "house", "mobile")
RESPONSE_TYPES = ("json", "csv", "s3")
RETRY_STATUSES = (429, 500, 502, 503, 504)
USER_AGENT = "tmg-comps-lookup/1.0"

# Each `lookup` bills a report, and a thin market tempts an agent into endless
# re-tries on slightly different filters. Six pulls is the ceiling for one
# analysis: after that, work with what came back or say the comps aren't there.
DEFAULT_PULL_BUDGET = 6
PULL_BUDGET_VAR = "DWELLSY_PULL_BUDGET"
BUDGET_FILE_VAR = "DWELLSY_BUDGET_FILE"
# A counter nobody reset must not brick the next job, so it ages out.
BUDGET_STALE_HOURS = 12

# Preferred column order when a report is printed as a table. These are the
# real Dwellsy field names (confirmed against a live 722-comp pull); any that a
# payload actually carries are shown first, the rest fill in after.
PREFERRED_COLUMNS = (
    "address_1", "address_2", "bedrooms", "bathrooms", "square_feet",
    "listing_amount", "price_per_sf", "distance_miles",
    "property_listing_status", "year_built", "community_unit_count",
    "company_name", "last_listing_creation_time", "address_city",
)

# The CSV export is a three-part report, not a flat table: a request echo, a
# summary block, then the comparables. Section titles sit alone on a line.
SECTION_TITLES = ("comp request details", "comp analysis", "comparables details")


class ApiError(RuntimeError):
    def __init__(self, message, status=None, body=None):
        super().__init__(message)
        self.status = status
        self.body = body


class BudgetExhausted(RuntimeError):
    """Raised when one analysis has already spent its allowance of pulls."""


# --------------------------------------------------------------------------
# pull budget -- see DEFAULT_PULL_BUDGET above
# --------------------------------------------------------------------------

def budget_path():
    override = os.environ.get(BUDGET_FILE_VAR)
    if override:
        return Path(override)
    import tempfile
    return Path(tempfile.gettempdir()) / "dwellsy_pull_budget.json"


def budget_limit(args=None):
    explicit = getattr(args, "budget", None)
    if explicit is not None:
        return max(0, int(explicit))
    env = os.environ.get(PULL_BUDGET_VAR)
    if env and env.strip().lstrip("-").isdigit():
        return max(0, int(env.strip()))
    return DEFAULT_PULL_BUDGET


def read_budget():
    path = budget_path()
    try:
        state = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {"count": 0, "started": time.time(), "pulls": []}
    if not isinstance(state, dict):
        return {"count": 0, "started": time.time(), "pulls": []}
    started = state.get("started")
    if not isinstance(started, (int, float)):
        started = time.time()
    if time.time() - started > BUDGET_STALE_HOURS * 3600:
        return {"count": 0, "started": time.time(), "pulls": []}
    return {"count": int(state.get("count") or 0), "started": started,
            "pulls": list(state.get("pulls") or [])}


def write_budget(state):
    try:
        budget_path().write_text(json.dumps(state, indent=2), encoding="utf-8")
    except OSError:
        pass                      # never fail a lookup over bookkeeping


def reset_budget():
    state = {"count": 0, "started": time.time(), "pulls": []}
    write_budget(state)
    return state


def check_budget(args):
    """Raise BudgetExhausted if this analysis has no pulls left."""
    limit = budget_limit(args)
    if limit <= 0:
        return read_budget(), limit
    state = read_budget()
    if state["count"] >= limit:
        raise BudgetExhausted(
            "pull budget exhausted: {0} of {0} API pulls already used for this "
            "analysis.\n"
            "Do NOT run another lookup. Either:\n"
            "  * build the analysis from the reports already pulled "
            "(list them with: history), or\n"
            "  * reply \"Can't find comps\" if what came back cannot support "
            "one.\n"
            "Pulls used: {1}\n"
            "Starting a genuinely new analysis? Run: budget --reset"
            .format(limit, ", ".join(state["pulls"]) or "(not recorded)"))
    return state, limit


def record_pull(label):
    state = read_budget()
    state["count"] += 1
    state["pulls"].append(label)
    write_budget(state)
    return state


# --------------------------------------------------------------------------
# configuration
# --------------------------------------------------------------------------

def parse_env_file(path):
    """Minimal KEY=VALUE .env reader (no dependency on python-dotenv)."""
    values = {}
    try:
        text = Path(path).read_text(encoding="utf-8-sig")
    except (OSError, UnicodeDecodeError):
        return values
    for line in text.splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        if line[:7].lower() == "export ":
            line = line[7:].lstrip()
        key, sep, val = line.partition("=")
        if not sep:
            continue
        key, val = key.strip(), val.strip()
        if len(val) >= 2 and val[0] == val[-1] and val[0] in "\"'":
            val = val[1:-1]
        if key:
            values[key] = val
    return values


def candidate_env_files():
    """.env files to try: cwd and script dir, then each parent of both."""
    seen, out = set(), []
    for start in (Path.cwd(), Path(__file__).resolve().parent):
        try:
            start = start.resolve()
        except OSError:
            continue
        for folder in (start, *start.parents):
            candidate = folder / ".env"
            if candidate not in seen:
                seen.add(candidate)
                out.append(candidate)
    return out


def load_config(env_file=None):
    """Return (api_key, base_url, source_description)."""
    key = os.environ.get(API_KEY_VAR, "").strip()
    base = os.environ.get(BASE_URL_VAR, "").strip()
    if key:
        return key, base or DEFAULT_BASE_URL, "environment"

    if env_file:
        paths = [Path(env_file)]
        if not paths[0].is_file():
            raise ApiError("env file not found: {}".format(env_file))
    else:
        paths = candidate_env_files()

    for path in paths:
        if not path.is_file():
            continue
        values = parse_env_file(path)
        key = (values.get(API_KEY_VAR) or "").strip()
        if key:
            base = base or (values.get(BASE_URL_VAR) or "").strip()
            return key, base or DEFAULT_BASE_URL, str(path)

    searched = "\n  ".join(str(p) for p in paths[:8])
    raise ApiError(
        "no {var} found.\nSet it in your .env file, e.g.\n  {var}=dw_live_xxxxx\n"
        "Looked in:\n  {paths}".format(var=API_KEY_VAR, paths=searched)
    )


# --------------------------------------------------------------------------
# API client
# --------------------------------------------------------------------------

class DwellsyClient:
    def __init__(self, api_key, base_url=DEFAULT_BASE_URL, timeout=180,
                 retries=3, verbose=False):
        self.api_key = api_key
        self.base_url = base_url.rstrip("/")
        self.timeout = timeout
        self.retries = max(0, retries)
        self.verbose = verbose

    def _log(self, message):
        if self.verbose:
            print("[dwellsy] {}".format(message), file=sys.stderr)

    def request(self, method, path, payload=None, authorize=True):
        """Perform a request with retry/backoff. Returns (body_bytes, headers)."""
        url = path if path.startswith("http") else self.base_url + path
        data = None
        headers = {"User-Agent": USER_AGENT, "Accept": "application/json"}
        if payload is not None:
            data = json.dumps(payload).encode("utf-8")
            headers["Content-Type"] = "application/json"
        # Only ever send the token to the API host -- report URLs can be
        # presigned S3/CDN links where an Authorization header breaks the
        # signature (and would leak the key to a third party).
        if authorize and self._same_host(url):
            headers["Authorization"] = "Bearer " + self.api_key

        attempt = 0
        while True:
            attempt += 1
            self._log("{} {} (attempt {})".format(method, url, attempt))
            req = urllib.request.Request(url, data=data, headers=headers,
                                         method=method)
            try:
                with urllib.request.urlopen(req, timeout=self.timeout) as resp:
                    return resp.read(), dict(resp.headers)
            except urllib.error.HTTPError as exc:
                body = exc.read()
                if exc.code in RETRY_STATUSES and attempt <= self.retries:
                    delay = self._retry_delay(exc.headers, attempt)
                    self._log("HTTP {} -- retrying in {}s".format(exc.code, delay))
                    time.sleep(delay)
                    continue
                raise ApiError(self._describe(exc.code, body), status=exc.code,
                               body=body)
            except ssl.SSLCertVerificationError as exc:
                raise ApiError("TLS certificate check failed for {}: {}".format(
                    url, exc))
            except (urllib.error.URLError, http.client.HTTPException,
                    socket.timeout, OSError) as exc:
                # Covers DNS/refused/timeout plus mid-flight drops such as
                # http.client.RemoteDisconnected, which is not a URLError.
                reason = getattr(exc, "reason", exc)
                if attempt <= self.retries:
                    delay = 2 ** attempt
                    self._log("network error ({}) -- retrying in {}s".format(
                        reason, delay))
                    time.sleep(delay)
                    continue
                raise ApiError("network error contacting {}: {}".format(
                    url, reason))

    def _same_host(self, url):
        return urllib.parse.urlsplit(url).netloc == \
            urllib.parse.urlsplit(self.base_url).netloc

    @staticmethod
    def _retry_delay(headers, attempt):
        raw = (headers or {}).get("Retry-After")
        if raw:
            try:
                return max(1, min(120, int(float(raw))))
            except ValueError:
                pass
        return min(60, 2 ** attempt)

    @staticmethod
    def _describe(status, body):
        detail = ""
        try:
            parsed = json.loads(body.decode("utf-8"))
            if isinstance(parsed, dict):
                detail = str(parsed.get("error") or parsed.get("message") or "")
        except Exception:
            detail = (body or b"").decode("utf-8", "replace").strip()[:300]
        hints = {
            400: "bad request -- check bedroom range, address and address_type",
            403: "not authorized -- {} is missing, wrong, or the request id "
                 "does not belong to this key".format(API_KEY_VAR),
            429: "rate limited by Dwellsy -- slow down or retry later",
        }
        parts = ["HTTP {}".format(status)]
        if status in hints:
            parts.append(hints[status])
        if detail:
            parts.append(detail)
        return ": ".join(parts)

    def request_json(self, method, path, payload=None):
        body, _ = self.request(method, path, payload)
        try:
            return json.loads(body.decode("utf-8"))
        except (ValueError, UnicodeDecodeError):
            raise ApiError("expected JSON from {} but got: {}".format(
                path, body[:200].decode("utf-8", "replace")))

    def generate_report(self, payload):
        return self.request_json("POST", "/comps", payload)

    def history(self):
        return self.request_json("GET", "/comps/history")

    def fetch_report(self, request_id):
        return self.request("GET", "/comps/" + urllib.parse.quote(request_id))

    def download(self, url):
        return self.request("GET", url)


# --------------------------------------------------------------------------
# request building
# --------------------------------------------------------------------------

def parse_range(text, label):
    """Accept '2' or '1-3' (also '1..3'). Returns (min, max)."""
    raw = str(text).strip().replace("..", "-")
    try:
        if "-" in raw[1:]:                       # keep a leading '-' out of it
            low, high = raw[1:].split("-", 1)
            low = raw[0] + low
        else:
            low = high = raw
        low, high = int(low), int(high)
    except ValueError:
        raise ApiError("could not read {} range from '{}' "
                       "(use '2' or '1-3')".format(label, text))
    if low > high:
        raise ApiError("{} minimum ({}) is above the maximum ({})".format(
            label, low, high))
    return low, high


def build_payload(args):
    payload = {}

    if args.address:
        payload["address"] = args.address
    if args.lat is not None or args.lon is not None:
        if args.lat is None or args.lon is None:
            raise ApiError("--lat and --lon must be given together")
        payload["latitude"] = args.lat
        payload["longitude"] = args.lon
    if not payload:
        raise ApiError("give a location: --address, or --lat and --lon")

    beds_min, beds_max = parse_range(args.beds, "bedroom")
    payload["bedrooms_min"] = beds_min
    payload["bedrooms_max"] = beds_max

    if args.sqft:
        sqft_min, sqft_max = parse_range(args.sqft, "square foot")
        payload["square_feet_min"] = sqft_min
        payload["square_feet_max"] = sqft_max

    if args.radius is not None:
        payload["radius"] = args.radius
    if args.months is not None:
        payload["months"] = args.months
    if args.type:
        payload["address_type"] = list(dict.fromkeys(args.type))

    payload["response_type"] = args.format
    if args.photos:
        payload["photos"] = True
    if args.format == "s3":
        if not args.s3_filename:
            raise ApiError("--format s3 needs --s3-filename, "
                           "e.g. comps/subject_2bd.csv")
        payload["s3_filename"] = args.s3_filename
    elif args.s3_filename:
        payload["s3_filename"] = args.s3_filename
    return payload


# --------------------------------------------------------------------------
# output helpers
# --------------------------------------------------------------------------

def column_width(name):
    """URLs and addresses need more room than the default column."""
    name = str(name).lower()
    if "url" in name or "link" in name:
        return 64
    if "address" in name or "name" in name:
        return 38
    return 24


def cell(value, width=24):
    """Render one value on a single line, ASCII-truncated to `width`."""
    if value is None:
        text = ""
    elif isinstance(value, bool):
        text = "yes" if value else "no"
    elif isinstance(value, float):
        # No thousands separators: these columns mix money with years and zips.
        text = ("{:.2f}".format(value).rstrip("0").rstrip(".")) or "0"
    elif isinstance(value, (dict, list)):
        text = json.dumps(value, separators=(",", ":"))
    else:
        text = str(value)
    text = " ".join(text.split())
    return text[:width - 3] + "..." if len(text) > width else text


def print_table(rows, columns, limit=20):
    """rows: list of dicts. Prints a fixed-width table to stdout."""
    if not rows:
        print("  (no rows)")
        return
    shown = rows[:limit]
    caps = [column_width(c) for c in columns]
    widths = [min(cap, max(len(str(col)),
                           *(len(cell(r.get(col), cap)) for r in shown)))
              for col, cap in zip(columns, caps)]
    header = "  ".join(str(c)[:w].ljust(w) for c, w in zip(columns, widths))
    print("  " + header.rstrip())
    print("  " + "  ".join("-" * w for w in widths))
    for row in shown:
        print("  " + "  ".join(cell(row.get(c), w).ljust(w)
                               for c, w in zip(columns, widths)).rstrip())
    if len(rows) > limit:
        print("  ... {} more rows (raise --limit to see them)".format(
            len(rows) - limit))


def pick_columns(rows, max_columns=8):
    keys = []
    for row in rows[:50]:
        for key in row:
            if key not in keys:
                keys.append(key)
    ordered = [k for k in PREFERRED_COLUMNS if k in keys]
    ordered += [k for k in keys if k not in ordered]
    return ordered[:max_columns]


def extract_records(doc):
    """Find the list of comp records inside whatever shape the report has."""
    if isinstance(doc, list):
        return [r for r in doc if isinstance(r, dict)]
    if isinstance(doc, dict):
        for key in ("comps", "results", "data", "records", "listings", "rentals"):
            value = doc.get(key)
            if isinstance(value, list):
                return [r for r in value if isinstance(r, dict)]
        for value in doc.values():
            if isinstance(value, list) and value and isinstance(value[0], dict):
                return value
    return []


def parse_sections(data):
    """Split a Dwellsy CSV report into [{title, header, rows}, ...].

    Parsing it as one flat table is wrong -- the request echo and summary block
    sit above the comparables, so row 1 is a single-cell title and every column
    count downstream is off.
    """
    text = data.decode("utf-8-sig", "replace")
    sections, current = [], None
    for row in csv.reader(io.StringIO(text)):
        cells = [c.strip() for c in row]
        if not any(cells):                       # padding / separator rows
            continue
        title = cells[0].rstrip(" :")
        if title.lower() in SECTION_TITLES and not any(cells[1:]):
            current = {"title": title, "header": None, "rows": []}
            sections.append(current)
            continue
        if current is None:                      # a plain CSV, no sections
            current = {"title": "", "header": None, "rows": []}
            sections.append(current)
        if current["header"] is None:
            current["header"] = row
        else:
            current["rows"].append(row)
    return [s for s in sections if s["header"]]


def comparables_section(sections):
    for section in sections:
        if section["title"].lower().startswith("comparables"):
            return section
    # A titled report with no comparables block means zero comps -- do NOT fall
    # back to the summary section, or an empty result reads as "1 comp".
    if any(s["title"] for s in sections):
        return None
    return sections[-1] if sections else None


def flatten_csv(data):
    """Strip the report down to just the comparables table (one header row)."""
    comps = comparables_section(parse_sections(data))
    if not comps or not comps["rows"]:
        return data
    buf = io.StringIO(newline="")
    writer = csv.writer(buf, lineterminator="\r\n")
    writer.writerow(comps["header"])
    writer.writerows(comps["rows"])
    return buf.getvalue().encode("utf-8")


def preview_csv(data, limit=10):
    sections = parse_sections(data)
    comps = comparables_section(sections)
    if not comps:
        print("  (no comparables in this report)")
        return
    for section in sections:
        if section["title"].lower() == "comp analysis" and section["rows"]:
            summary = "  ".join(
                "{}={}".format(k, v)
                for k, v in zip(section["header"], section["rows"][0]) if v)
            print("  summary: " + summary)
    print("  {} comps, {} columns".format(len(comps["rows"]), len(comps["header"])))
    dict_rows = [dict(zip(comps["header"], r)) for r in comps["rows"]]
    print_table(dict_rows, pick_columns(dict_rows), limit=limit)


def write_out(path, data):
    path = Path(path)
    if path.parent and not path.parent.exists():
        path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(data)
    print("Saved {} ({:,} bytes)".format(path, len(data)))


def default_out_name(request_id, response_type):
    suffix = "csv" if response_type == "csv" else "json"
    return "comps_{}.{}".format((request_id or "report")[:8], suffix)


# --------------------------------------------------------------------------
# commands
# --------------------------------------------------------------------------

def cmd_lookup(args, client_factory):
    payload = build_payload(args)
    if args.dry_run:
        print(json.dumps(payload, indent=2))
        return 0

    state, limit = check_budget(args)          # raises before spending a pull

    client = client_factory()
    result = client.generate_report(payload)
    label = payload.get("address") or "{},{}".format(
        payload.get("latitude"), payload.get("longitude"))
    if limit > 0:
        state = record_pull(str(label))
        print("[pull {}/{}]".format(state["count"], limit))
    if args.raw:
        print(json.dumps(result, indent=2))

    request_id = result.get("requestId")
    count = result.get("compsCount")
    url = result.get("url")
    print("status     : {}".format(result.get("status", "?")))
    print("request id : {}".format(request_id))
    print("comps found: {}".format(count if count is not None else "?"))
    if url:
        print("report url : {}".format(url))

    if not count:
        left = max(0, limit - state["count"]) if limit > 0 else None
        print("\nNo comps matched. Try more --months, a broader --beds range, "
              "or fewer --type filters (--radius is capped at 5).")
        if left is not None:
            print("Pulls left in this analysis: {}.".format(left)
                  + ("" if left else " Report \"Can't find comps\" instead of "
                                     "retrying."))
        return 0

    if args.format == "s3":
        print("\nUploaded to S3 as: {}".format(payload.get("s3_filename")))
        return 0

    data = None
    if url:
        data, _ = client.download(url)
    elif args.format == "json":
        records = extract_records(result)
        if records:
            data = json.dumps(records, indent=2).encode("utf-8")

    if data is None:
        print("\nNo report body returned; fetch it later with:\n"
              "  python {} report {}".format(Path(__file__).name, request_id))
        return 0

    out_path = args.out or (default_out_name(request_id, args.format)
                            if args.format == "csv" else None)

    print("")
    if args.format == "csv":
        preview_csv(data, limit=args.limit)
        if args.flat:
            data = flatten_csv(data)
            print("  (--flat: saving the comparables table only, "
                  "no request/summary sections)")
    else:
        try:
            doc = json.loads(data.decode("utf-8"))
        except (ValueError, UnicodeDecodeError):
            doc = None
        records = extract_records(doc) if doc is not None else []
        if records:
            print("  {} comps".format(len(records)))
            print_table(records, pick_columns(records), limit=args.limit)
        else:
            print(data.decode("utf-8", "replace")[:2000])

    if out_path:
        write_out(out_path, data)
    return 0


def cmd_budget(args, client_factory):
    if args.reset:
        reset_budget()
        print("Pull budget reset. {} pulls available for this analysis."
              .format(budget_limit(args)))
        return 0
    state = read_budget()
    limit = budget_limit(args)
    print("Pulls used : {} of {}".format(state["count"], limit))
    print("Remaining  : {}".format(max(0, limit - state["count"])))
    print("State file : {}".format(budget_path()))
    if state["pulls"]:
        print("Spent on   :")
        for i, label in enumerate(state["pulls"], 1):
            print("  {}. {}".format(i, label))
    return 0


def cmd_history(args, client_factory):
    client = client_factory()
    result = client.history()
    if args.raw:
        print(json.dumps(result, indent=2))
        return 0
    reports = result.get("reports") or []
    if not reports:
        print("No reports on record for this API key.")
        return 0
    print("{} recent report(s)".format(len(reports)))
    print_table(reports,
                ["requestTime", "requestId", "responseType", "compsCount", "url"],
                limit=args.limit)
    return 0


def cmd_report(args, client_factory):
    client = client_factory()
    data, headers = client.fetch_report(args.request_id)
    content_type = (headers.get("Content-Type") or "").lower()

    # /comps/{id} may hand back the report itself or a JSON envelope with a URL.
    if "json" in content_type:
        try:
            doc = json.loads(data.decode("utf-8"))
        except (ValueError, UnicodeDecodeError):
            doc = None
        if isinstance(doc, dict) and doc.get("url"):
            print("report url : {}".format(doc["url"]))
            data, headers = client.download(doc["url"])
            content_type = (headers.get("Content-Type") or "").lower()
        elif doc is not None:
            records = extract_records(doc)
            if records:
                print("{} comps".format(len(records)))
                print_table(records, pick_columns(records), limit=args.limit)

    is_csv = "csv" in content_type or (args.out or "").lower().endswith(".csv")
    if not is_csv and data[:64].lstrip().lower().startswith(b"comp request"):
        is_csv = True                      # served as text/plain but is the CSV
    if is_csv:
        preview_csv(data, limit=args.limit)
        if args.flat:
            data = flatten_csv(data)
            print("  (--flat: saving the comparables table only)")

    out_path = args.out or default_out_name(
        args.request_id, "csv" if is_csv else "json")
    write_out(out_path, data)
    return 0


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------

# Applied after parsing, because the shared flags below use SUPPRESS so that
# they work either before or after the subcommand name.
GLOBAL_DEFAULTS = {
    "env_file": None, "base_url": None, "timeout": 180,
    "retries": 3, "limit": 20, "raw": False, "verbose": False,
}


def add_common_options(parser):
    group = parser.add_argument_group("common options")
    group.add_argument("--env-file", default=argparse.SUPPRESS,
                       help="explicit .env file holding " + API_KEY_VAR)
    group.add_argument("--base-url", default=argparse.SUPPRESS,
                       help="override the API base URL (default: {})".format(
                           DEFAULT_BASE_URL))
    group.add_argument("--timeout", type=int, default=argparse.SUPPRESS,
                       help="per-request timeout in seconds (default: 180)")
    group.add_argument("--retries", type=int, default=argparse.SUPPRESS,
                       help="retries on 429/5xx/network errors (default: 3)")
    group.add_argument("--limit", type=int, default=argparse.SUPPRESS,
                       help="rows to print in previews (default: 20)")
    group.add_argument("--raw", action="store_true", default=argparse.SUPPRESS,
                       help="also dump the raw API JSON")
    group.add_argument("-v", "--verbose", action="store_true",
                       default=argparse.SUPPRESS,
                       help="log each HTTP call to stderr")
    return parser


def build_parser():
    parser = argparse.ArgumentParser(
        prog="dwellsy_comps_lookup.py",
        description="Look up rent comps through the Dwellsy Comps API.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="The API key is read from {} (environment or .env).".format(
            API_KEY_VAR))
    add_common_options(parser)

    subparsers = parser.add_subparsers(dest="command", required=True)

    look = add_common_options(
        subparsers.add_parser("lookup", help="generate a comps report"))
    look.add_argument("--address", help="subject address, geocoded by Dwellsy")
    look.add_argument("--lat", type=float, help="subject latitude")
    look.add_argument("--lon", "--lng", dest="lon", type=float,
                      help="subject longitude")
    look.add_argument("--beds", required=True,
                      help="bedroom count or range, e.g. 2 or 1-3 (studios = 0)")
    look.add_argument("--sqft", help="unit size range, e.g. 700-1200")
    look.add_argument("--radius", type=int,
                      help="search radius in miles, 5 is the API maximum "
                           "(default: 2)")
    look.add_argument("--months", type=int,
                      help="lookback window in months, e.g. 6")
    look.add_argument("--type", nargs="+", choices=ADDRESS_TYPES,
                      help="property types to include (default: all)")
    look.add_argument("--format", choices=RESPONSE_TYPES, default="json",
                      help="response format (default: json)")
    look.add_argument("--photos", action="store_true",
                      help="include listing photos")
    look.add_argument("--s3-filename",
                      help="destination key when --format s3")
    look.add_argument("-o", "--out", help="write the report to this file")
    look.add_argument("--flat", action="store_true",
                      help="CSV only: save just the comparables table, "
                           "dropping the request/summary sections so the file "
                           "opens as a normal spreadsheet")
    look.add_argument("--budget", type=int, default=None,
                      help="max API pulls allowed for this analysis "
                           "(default: {}; 0 disables the check)"
                           .format(DEFAULT_PULL_BUDGET))
    look.add_argument("--dry-run", action="store_true",
                      help="print the request body and exit")
    look.set_defaults(func=cmd_lookup)

    bud = add_common_options(
        subparsers.add_parser("budget",
                              help="show or reset the pull budget for an analysis"))
    bud.add_argument("--reset", action="store_true",
                     help="start a new analysis with a full allowance")
    bud.add_argument("--budget", type=int, default=None,
                     help="allowance to report against (default: {})"
                          .format(DEFAULT_PULL_BUDGET))
    bud.set_defaults(func=cmd_budget)

    hist = add_common_options(
        subparsers.add_parser("history", help="list recent comps reports"))
    hist.set_defaults(func=cmd_history)

    rep = add_common_options(
        subparsers.add_parser("report", help="download a past report by id"))
    rep.add_argument("request_id", help="request id from lookup or history")
    rep.add_argument("-o", "--out", help="write the report to this file")
    rep.add_argument("--flat", action="store_true",
                     help="CSV only: save just the comparables table")
    rep.set_defaults(func=cmd_report)

    return parser


def main(argv=None):
    # Comp addresses carry accents; a legacy Windows console cannot encode them
    # and would otherwise raise UnicodeEncodeError mid-table. Keep the console's
    # own encoding (re-encoding to UTF-8 just produces mojibake there).
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")
        except (AttributeError, ValueError):
            pass

    args = build_parser().parse_args(argv)
    for name, value in GLOBAL_DEFAULTS.items():
        if not hasattr(args, name):
            setattr(args, name, value)

    try:
        def client_factory():
            api_key, base_url, source = load_config(args.env_file)
            if args.verbose:
                print("[dwellsy] key from {}".format(source), file=sys.stderr)
            return DwellsyClient(
                api_key,
                base_url=args.base_url or base_url,
                timeout=args.timeout,
                retries=args.retries,
                verbose=args.verbose,
            )

        return args.func(args, client_factory)
    except BudgetExhausted as exc:
        print("STOP: {}".format(exc), file=sys.stderr)
        return 3
    except ApiError as exc:
        print("Error: {}".format(exc), file=sys.stderr)
        return 2
    except KeyboardInterrupt:
        print("Interrupted.", file=sys.stderr)
        return 130


if __name__ == "__main__":
    sys.exit(main())
