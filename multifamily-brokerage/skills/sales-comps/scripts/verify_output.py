#!/usr/bin/env python3
"""Sanity-check the two output workbooks before delivery.

Usage: python verify_output.py <outdir> [--cma-name "Automatic CMA Analysis.xlsx"]
Exits non-zero if any check fails.
"""
import argparse
import glob
import os
import re
import sys
import zipfile
from xml.etree import ElementTree

FAILED = []


def check(label, ok, detail=""):
    print(f"  [{'ok' if ok else 'FAIL'}] {label}" + (f" — {detail}" if detail and not ok else ""))
    if not ok:
        FAILED.append(label)


def parts(path):
    out = {}
    with zipfile.ZipFile(path) as z:
        for n in z.namelist():
            out[n] = z.read(n)
    return out


def xml_wellformed(p):
    bad = []
    for n, b in p.items():
        if n.endswith((".xml", ".rels")):
            try:
                ElementTree.fromstring(b)
            except ElementTree.ParseError as e:
                bad.append(f"{n}: {e}")
    return bad


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("outdir")
    ap.add_argument("--cma-name", default="Automatic CMA Analysis.xlsx")
    args = ap.parse_args()

    cma_path = os.path.join(args.outdir, args.cma_name)
    grids = [f for f in glob.glob(os.path.join(args.outdir, "*.xlsx"))
             if os.path.basename(f) != args.cma_name]
    check("both output files exist", os.path.exists(cma_path) and len(grids) == 1,
          f"cma={os.path.exists(cma_path)}, grids={grids}")
    if FAILED:
        sys.exit(1)
    grid_path = grids[0]

    print(f"\n{args.cma_name}:")
    p = parts(cma_path)
    bad = xml_wellformed(p)
    check("all XML parts well-formed", not bad, "; ".join(bad[:3]))
    check("Power Query connections preserved", "xl/connections.xml" in p)
    check("DataMashup (queries) preserved", "customXml/item1.xml" in p)
    check("calcChain removed (Excel rebuilds it)", "xl/calcChain.xml" not in p)
    tbl = next((b for n, b in p.items() if n.startswith("xl/tables/")
                and b"Output_Analysis_Data" in b), b"")
    m = re.search(rb'ref="A1:X(\d+)"', tbl)
    check("Output_Analysis_Data table ref updated", bool(m and int(m.group(1)) > 1))

    print(f"\n{os.path.basename(grid_path)}:")
    p = parts(grid_path)
    bad = xml_wellformed(p)
    check("all XML parts well-formed", not bad, "; ".join(bad[:3]))
    rels = p.get("xl/externalLinks/_rels/externalLink1.xml.rels", b"").decode()
    from urllib.parse import quote
    check("external link targets the CMA file (URI-encoded — literal spaces make Excel 'repair' the part)",
          all(t == quote(args.cma_name) for t in re.findall(r'Target="([^"]+)"', rels)),
          rels[:200])
    link_xml = p["xl/externalLinks/externalLink1.xml"].decode()
    n_names = len(re.findall(r"<sheetName ", link_xml))
    ids = re.findall(r'<sheetData sheetId="(\d+)"', link_xml)
    check("cached data covers every external sheet (Excel repairs sparse sheetDataSets)",
          ids == [str(i) for i in range(n_names)] and "refreshError" not in link_xml,
          f"{len(ids)} sheetData vs {n_names} sheetNames")
    check("xxl21 alternateUrls removed (a relative path in absoluteUrl breaks Excel)",
          "alternateUrls" not in link_xml
          and len(re.findall(r"<Relationship ", rels)) == 1, rels[:200])
    wb = p["xl/workbook.xml"].decode()
    rediq = dict(re.findall(r'<definedName name="(rediq_[^"]+)"[^>]*>([^<]*)</definedName>', wb))
    check("rediq names remapped to Inputs",
          all(v.startswith("[1]Inputs!") for v in rediq.values()) and len(rediq) == 5,
          str(rediq))
    check("fullCalcOnLoad set", 'fullCalcOnLoad="1"' in wb)
    sheet = p["xl/worksheets/sheet1.xml"].decode()
    check("helper grid relinked (N3)", "'[1]Output Analysis Data'!M2</f>" in sheet)
    check("helper grid relinked (V52 distance)", "'[1]Output Analysis Data'!R51</f>" in sheet)
    check("no formulas still reference the underwriting table",
          "Output_Analysis_Data__2" not in sheet)
    xs = re.findall(r'<c r="L(\d+)"[^>]*t="inlineStr"><is><t[^>]*>x</t>', sheet)
    check("exactly 5 comps marked x", len(xs) == 5, f"marked rows: {xs}")
    link = p["xl/externalLinks/externalLink1.xml"].decode()
    check("cached link data includes Output Analysis Data",
          'val="Output Analysis Data"' in link and "<sheetDataSet>" in link)
    valid_sheets = {s.replace("&amp;", "&")
                    for s in re.findall(r'<sheetName val="([^"]+)"', link)}
    used = set()
    for src in (wb, sheet):
        for s in re.findall(r"'?\[\d+\]([^'!\]]+)'?!", src.replace("&apos;", "'")):
            used.add(s)
    dangling = used - valid_sheets
    check("every external sheet reference resolves in the link's sheet list "
          "(dangling refs make Excel repair the part)", not dangling, str(dangling))
    import html
    decoded = [html.unescape(s) for s in re.findall(r'<sheetName val="([^"]+)"', link)]
    check("cached sheet names: no double-escaping, all ≤31 chars",
          "&amp;amp;" not in link and all(len(s) <= 31 for s in decoded),
          str([s for s in decoded if len(s) > 31]))
    has_meta_refs = bool(re.search(r'<c [^>]*(?:cm|vm)="', sheet))
    check("no orphaned metadata part (unreferenced xl/metadata.xml gets 'repaired')",
          ("xl/metadata.xml" in p) == has_meta_refs,
          f"metadata part present={'xl/metadata.xml' in p}, cells referencing it={has_meta_refs}")

    # optional: LibreOffice round-trip (catches anything Excel would 'repair')
    from shutil import which
    if which("soffice"):
        import subprocess, tempfile
        with tempfile.TemporaryDirectory() as td:
            r = subprocess.run(["soffice", "--headless", "--convert-to", "xlsx",
                                "--outdir", td, grid_path],
                               capture_output=True, timeout=180)
            check("LibreOffice opens the grid without error",
                  r.returncode == 0 and glob.glob(os.path.join(td, "*.xlsx")) != [],
                  r.stderr.decode()[:200])
    else:
        print("  [skip] LibreOffice not installed — skipping open test")

    print()
    if FAILED:
        print(f"{len(FAILED)} check(s) FAILED — fix before delivering.")
        sys.exit(1)
    print("All checks passed.")


if __name__ == "__main__":
    main()
