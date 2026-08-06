#!/usr/bin/env python3
"""Regression-check that a newly added rent-roll / T-12 XLSX parser cannot
steal another format's file (and that no existing parser steals the new one).

Run this before merging any new registered parser into process_rent_roll.py or
process_t12.py. The toolkit dispatches XLSX sources by asking each registered
detector, in order, "is this yours?" — so the failure mode that matters is a
detector that is too loose and claims a file belonging to another format. That
bug is silent: the wrong parser runs, produces plausible numbers, and the
reconciliation block ties against the wrong report structure.

What it does: offers every XLSX you point it at to EVERY registered detector on
both sides (rent roll `XLSX_PARSERS`, T-12 `T12_XLSX_PARSERS`), prints the full
claim matrix, and asserts:

  1. No file is claimed by more than one detector on the same side. (A second
     claimant is only masked by registration order — reorder the list, or add
     a format whose file both match, and it becomes a live bug.)
  2. Every --expect DETECTOR=FILE pairing holds.
  3. Each --expect'ed file is claimed by that detector and no other.
  4. Each --expect'ed detector claims nothing except its own file.

Exit status is 0 only if every assertion passes.

Usage
-----
    python parser_detection_regression.py --toolkit ./toolkit \
        --files "./sources/*.xlsx" "./toolkit/*.xlsx" "./out/*.xlsx" \
        --expect AppFolioUnitTypeXlsxParser="./sources/new roll.xlsx" \
        --expect _is_resman_trailing_xlsx="./sources/new t12.xlsx"

Point --files at everything you have: the new source, the toolkit templates,
previously delivered RR/T-12 workbooks, and any other format's sources you can
get hold of. Coverage is only as good as the file set — the script reports how
many files it tested so a thin run is visible rather than falsely reassuring.

Detector calling convention (as of 8/2026): both sides take a PATH, not an
open workbook. Rent roll detectors are `Parser.detect_xlsx(path)` (a static or
class method — call it on the CLASS, not an instance). T-12 detectors are the
first element of each `(detect_fn, parse_fn)` tuple in `T12_XLSX_PARSERS`.
"""
import argparse
import glob
import os
import sys


def load_toolkit(toolkit_dir):
    """Import the two toolkit scripts from an arbitrary directory."""
    toolkit_dir = os.path.abspath(toolkit_dir)
    if not os.path.isdir(toolkit_dir):
        sys.exit(f"ERROR: --toolkit {toolkit_dir} is not a directory")
    sys.path.insert(0, toolkit_dir)
    mods = {}
    for mod_name, attr in (("process_rent_roll", "XLSX_PARSERS"),
                           ("process_t12", "T12_XLSX_PARSERS")):
        path = os.path.join(toolkit_dir, mod_name + ".py")
        if not os.path.isfile(path):
            print(f"  note: {mod_name}.py not in the toolkit dir - "
                  f"skipping that side")
            mods[mod_name] = None
            continue
        mod = __import__(mod_name)
        if not hasattr(mod, attr):
            sys.exit(f"ERROR: {mod_name} has no {attr} - has the registration "
                     f"convention changed?")
        mods[mod_name] = mod
    return mods["process_rent_roll"], mods["process_t12"]


def expand(patterns):
    """Expand globs/paths into a de-duplicated, order-stable file list."""
    out, seen = [], set()
    for pat in patterns:
        hits = sorted(glob.glob(pat)) or ([pat] if os.path.isfile(pat) else [])
        if not hits:
            print(f"  note: --files pattern matched nothing: {pat}")
        for h in hits:
            key = os.path.normcase(os.path.abspath(h))
            if key not in seen:
                seen.add(key)
                out.append(h)
    return out


def detectors(prr, pt):
    """[(side, name, callable)] for every registered XLSX detector."""
    found = []
    if prr is not None:
        for P in prr.XLSX_PARSERS:
            # detect_xlsx is a static/classmethod: bind to the CLASS.
            found.append(("rent roll", P.__name__, P.detect_xlsx))
    if pt is not None:
        for det, _parse in pt.T12_XLSX_PARSERS:
            found.append(("T-12", det.__name__, det))
    return found


def claims_for(path, dets, side):
    """Names of the detectors on `side` that claim `path`."""
    out = []
    for d_side, name, fn in dets:
        if d_side != side:
            continue
        try:
            if fn(path):
                out.append(name)
        except Exception as e:                       # a detector must never
            out.append(f"{name}!ERR:{type(e).__name__}")   # raise on a
    return out                                       # foreign file


def main(argv=None):
    ap = argparse.ArgumentParser(
        description="Cross-detection regression for toolkit XLSX parsers.")
    ap.add_argument("--toolkit", default="toolkit",
                    help="directory holding process_rent_roll.py / "
                         "process_t12.py (default: ./toolkit)")
    ap.add_argument("--files", nargs="+", required=True,
                    help="xlsx paths or globs to offer to every detector")
    ap.add_argument("--expect", action="append", default=[],
                    metavar="DETECTOR=FILE",
                    help="assert DETECTOR claims FILE (and only FILE); "
                         "repeatable")
    args = ap.parse_args(argv)

    prr, pt = load_toolkit(args.toolkit)
    dets = detectors(prr, pt)
    files = expand(args.files)
    if not files:
        sys.exit("ERROR: no files to test")

    expects = []
    for spec in args.expect:
        if "=" not in spec:
            sys.exit(f"ERROR: --expect needs DETECTOR=FILE, got {spec!r}")
        name, path = spec.split("=", 1)
        name, path = name.strip(), path.strip().strip('"')
        if not os.path.isfile(path):
            sys.exit(f"ERROR: --expect file does not exist: {path}")
        if name not in [n for _s, n, _f in dets]:
            sys.exit(f"ERROR: no registered detector named {name!r}. "
                     f"Registered: {[n for _s, n, _f in dets]}")
        expects.append((name, path))

    same = lambda a, b: os.path.normcase(os.path.abspath(a)) == \
        os.path.normcase(os.path.abspath(b))

    matrix = {}
    for side in ("rent roll", "T-12"):
        names = [n for s, n, _f in dets if s == side]
        if not names:
            continue
        print(f"=== {side} detectors, in registration order ===")
        print(f"    {names}")
        for f in files:
            got = claims_for(f, dets, side)
            matrix[(side, os.path.abspath(f))] = got
            print(f"  {os.path.basename(f)[:56]:58s} -> "
                  f"{got or ['(none)']}")
        print()

    fails = []

    def check(label, ok):
        print(("  PASS  " if ok else "  FAIL  ") + label)
        if not ok:
            fails.append(label)

    print("=== assertions ===")

    # 1. no file may be claimed twice on the same side
    for (side, path), got in sorted(matrix.items()):
        real = [g for g in got if "!ERR:" not in g]
        errs = [g for g in got if "!ERR:" in g]
        if len(real) > 1:
            check(f"[{side}] {os.path.basename(path)[:40]} claimed by exactly "
                  f"one detector (got {real})", False)
        for e in errs:
            check(f"[{side}] detector raised on "
                  f"{os.path.basename(path)[:40]}: {e}", False)

    # 2-4. the --expect pairings
    for name, path in expects:
        side = next(s for s, n, _f in dets if n == name)
        got = matrix[(side, os.path.abspath(path))]
        check(f"{name} claims {os.path.basename(path)[:44]}", name in got)
        check(f"{name} is the ONLY claimant of "
              f"{os.path.basename(path)[:36]}", got == [name])
        others = [f for f in files if not same(f, path)
                  and name in matrix[(side, os.path.abspath(f))]]
        check(f"{name} claims nothing else "
              f"({len(files) - 1} other file(s) tested)", not others)
        if others:
            for o in others:
                print(f"          also claimed: {os.path.basename(o)}")

    print()
    print(f"tested {len(files)} file(s) against {len(dets)} detector(s)")
    if fails:
        print(f"RESULT: {len(fails)} FAILURE(S)")
        return 1
    print("RESULT: all checks pass")
    return 0


if __name__ == "__main__":
    sys.exit(main())
