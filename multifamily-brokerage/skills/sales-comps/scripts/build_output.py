#!/usr/bin/env python3
"""Build the two output workbooks from a selection.json:

1. A per-deal copy of the Automatic CMA Analysis workbook with the Inputs tab
   filled and the 'Output Analysis Data' tab rewritten with the recomputed,
   sorted top-50 comps.
2. A Sale Comparables Workbook (from the bundled template) whose external link
   is retargeted to the CMA file: the helper grid pulls 'Output Analysis Data'
   by direct cell reference, the Subject column pulls the Inputs tab, the five
   selected comps are marked, and the Cap Rate Drift row carries computed bps.

All edits happen at the raw-XML zip level so Power Query connections, styles,
and everything else in both workbooks survive untouched.
See references/relinking.md for the exact cell/formula map.
"""
import argparse
import json
import os
import re
import shutil
import sys
import zipfile
from datetime import date

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_TEMPLATE = os.path.join(HERE, "..", "assets", "Sale Comparables Workbook.xlsx")
EXCEL_EPOCH = date(1899, 12, 30)

# helper-grid column -> Output Analysis Data column (both on the linked CMA file)
HELPER_MAP = {"N": "M", "O": "A", "Q": "B", "R": "C", "S": "D", "T": "E",
              "U": "F", "V": "R", "W": "L", "X": "G", "Y": "H", "Z": "J"}
# Subject cells that pulled from the underwriting model's Master tab -> Inputs
SUBJECT_MAP = {"D17": "'[1]Inputs'!B9", "D21": "'[1]Inputs'!B7", "D23": "'[1]Inputs'!B8"}
REDIQ_MAP = {"rediq_dealname": "[1]Inputs!$B$2", "rediq_address1": "[1]Inputs!$B$3",
             "rediq_city": "[1]Inputs!$B$4", "rediq_state": "[1]Inputs!$B$5",
             "rediq_zip": "[1]Inputs!$B$6"}
HELPER_ROWS = range(3, 53)          # grid rows 3..52 <-> table data rows 1..50
OUT_COLS = ["Property Name", "Property Address", "City", "State", "ZIP", "Unit Count",
            "Year Built", "Sold Price", "Sold Price/Unit", "Sale Date", "Building SF",
            "Avg Unit SF", "Info Source", "Latitude", "Longitude", "Lat1_Rad",
            "Lon1_Rad", "Distance (mi.)", "DistancePoints", "AgeSpread", "AgePoints",
            "DaysSinceSale", "DatePoints", "TotalPoints"]
STR_COLS = {"Property Name", "Property Address", "City", "State", "ZIP", "Info Source"}


def esc(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            .replace('"', "&quot;"))


def col_letter(i):  # 1-based
    s = ""
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


class Zip:
    """In-memory xlsx editor operating on raw part XML."""

    def __init__(self, path):
        self.parts, self.order = {}, []
        with zipfile.ZipFile(path) as z:
            for n in z.namelist():
                self.parts[n] = z.read(n)
                self.order.append(n)

    def text(self, name):
        return self.parts[name].decode("utf-8")

    def set(self, name, xml):
        self.parts[name] = xml.encode("utf-8")

    def drop_part(self, part):
        """Remove a workbook-level part plus its content-type and relationship."""
        if part not in self.parts:
            return
        del self.parts[part]
        self.order.remove(part)
        base = part.split("/")[-1]
        ct = self.text("[Content_Types].xml")
        self.set("[Content_Types].xml",
                 re.sub(r'<Override PartName="/%s"[^>]*/>' % re.escape(part), "", ct))
        rels = self.text("xl/_rels/workbook.xml.rels")
        self.set("xl/_rels/workbook.xml.rels",
                 re.sub(r'<Relationship [^>]*Target="%s"[^>]*/>' % re.escape(base), "", rels))

    def drop_calc_chain(self):
        self.drop_part("xl/calcChain.xml")

    def sheet_part(self, title):
        wbxml = self.text("xl/workbook.xml")
        m = re.search(r'<sheet name="%s"[^>]*r:id="(rId\d+)"' % re.escape(esc(title)), wbxml)
        if not m:
            sys.exit(f"ERROR: sheet '{title}' not in workbook.xml")
        rels = self.text("xl/_rels/workbook.xml.rels")
        m2 = re.search(r'<Relationship Id="%s"[^>]*Target="([^"]+)"' % m.group(1), rels)
        t = m2.group(1)
        return "xl/" + t if not t.startswith("/") else t.lstrip("/")

    def save(self, path):
        os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            for n in self.order:
                z.writestr(n, self.parts[n])


def cell_xml(ref, value, style=None, kind="auto"):
    s = f' s="{style}"' if style else ""
    if value is None or value == "":
        return f'<c r="{ref}"{s}/>'
    if kind == "str" or (kind == "auto" and isinstance(value, str)):
        return f'<c r="{ref}"{s} t="inlineStr"><is><t xml:space="preserve">{esc(value)}</t></is></c>'
    return f'<c r="{ref}"{s}><v>{value}</v></c>'


def replace_cell(sheet_xml, ref, new_cell):
    """Replace (or insert, keeping column order) a <c> element in a row."""
    pat = re.compile(r'<c r="%s"(?: [^>]*)?(?:/>|>.*?</c>)' % ref, re.S)
    if pat.search(sheet_xml):
        return pat.sub(lambda _: new_cell, sheet_xml, count=1)
    # insert into the row
    rownum = re.match(r"[A-Z]+(\d+)", ref).group(1)
    colnum = col_num(re.match(r"([A-Z]+)", ref).group(1))
    rowpat = re.compile(r'(<row r="%s"[^>]*>)(.*?)(</row>)' % rownum, re.S)
    m = rowpat.search(sheet_xml)
    if not m:
        sys.exit(f"ERROR: row {rownum} not found for insert of {ref}")
    cells = re.findall(r'<c r="([A-Z]+)%s"' % rownum, m.group(2))
    body = m.group(2)
    inserted = False
    for c in cells:
        if col_num(c) > colnum:
            body = re.sub(r'(<c r="%s%s")' % (c, rownum), new_cell + r"\1", body, count=1)
            inserted = True
            break
    if not inserted:
        body += new_cell
    return sheet_xml[: m.start()] + m.group(1) + body + m.group(3) + sheet_xml[m.end():]


def set_formula(sheet_xml, ref, formula, cached=None):
    formula = formula[1:] if formula.startswith("=") else formula  # <f> stores no '='
    pat = re.compile(r'<c r="%s"( [^>]*)?(?:/>|>.*?</c>)' % ref, re.S)
    m = pat.search(sheet_xml)
    attrs = m.group(1) or "" if m else ""
    style = re.search(r's="(\d+)"', attrs)
    s = f' s="{style.group(1)}"' if style else ""
    v = f"<v>{cached}</v>" if cached is not None else ""
    new = f'<c r="{ref}"{s}><f>{esc(formula)}</f>{v}</c>'
    if m:
        return pat.sub(lambda _: new, sheet_xml, count=1)
    return replace_cell(sheet_xml, ref, new)


def col_num(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def date_serial(iso):
    y, m, d = map(int, iso[:10].split("-"))
    return (date(y, m, d) - EXCEL_EPOCH).days


def row_styles(sheet_xml, rownum):
    """Sample per-column style ids from an existing row."""
    m = re.search(r'<row r="%d"[^>]*>(.*?)</row>' % rownum, sheet_xml, re.S)
    styles = {}
    if m:
        for ref, attrs in re.findall(r'<c r="([A-Z]+)%d"((?: [^>]*)?)[/>]' % rownum, m.group(1)):
            sm = re.search(r's="(\d+)"', attrs)
            if sm:
                styles[ref] = sm.group(1)
    return styles


# ---------------------------------------------------------------------------
# Part A: the CMA copy
# ---------------------------------------------------------------------------
def build_cma(cma_path, sel, out_path):
    z = Zip(cma_path)
    subj = sel["subject"]

    # Inputs tab: B2..B9
    part = z.sheet_part("Inputs")
    xml = z.text(part)
    for ref, val in [("B2", subj["name"]), ("B3", subj["address"]), ("B4", subj["city"]),
                     ("B5", subj["state"]), ("B6", subj["zip"]), ("B7", subj["year_built"]),
                     ("B8", subj["avg_size"]), ("B9", subj["units"])]:
        styles = row_styles(xml, int(ref[1:]))
        xml = replace_cell(xml, ref, cell_xml(ref, val, styles.get(ref[0])))
    z.set(part, xml)

    # Output Analysis Data tab: rewrite with recomputed top-50
    part = z.sheet_part("Output Analysis Data")
    xml = z.text(part)
    styles = row_styles(xml, 2)
    comps = sel["comps"]
    rows = []
    m = re.search(r'<row r="1"[^>]*>.*?</row>', xml, re.S)
    rows.append(m.group(0))  # keep the header row verbatim
    for i, c in enumerate(comps):
        r = i + 2
        cells = []
        for j, colname in enumerate(OUT_COLS):
            L = col_letter(j + 1)
            v = c.get(colname)
            if colname == "Sale Date" and v:
                cells.append(cell_xml(f"{L}{r}", date_serial(v), styles.get(L), kind="num"))
            elif colname in STR_COLS:
                cells.append(cell_xml(f"{L}{r}", "" if v is None else str(v), styles.get(L), kind="str"))
            else:
                cells.append(cell_xml(f"{L}{r}", "" if v is None else v, styles.get(L), kind="num"))
        rows.append(f'<row r="{r}" spans="1:24">' + "".join(cells) + "</row>")
    last = len(comps) + 1
    body = "<sheetData>" + "".join(rows) + "</sheetData>"
    xml = re.sub(r"<sheetData>.*</sheetData>", lambda _: body, xml, count=1, flags=re.S)
    xml = re.sub(r'<dimension ref="[^"]*"/>', f'<dimension ref="A1:X{last}"/>', xml, count=1)
    z.set(part, xml)

    # table + autofilter refs
    for name in list(z.parts):
        if name.startswith("xl/tables/") and b"Output_Analysis_Data" in z.parts[name]:
            t = z.text(name)
            t = re.sub(r'ref="A1:X\d+"', f'ref="A1:X{last}"', t)
            z.set(name, t)

    z.drop_calc_chain()
    z.save(out_path)
    return last


# ---------------------------------------------------------------------------
# Part B: the relinked Sale Comparables Workbook
# ---------------------------------------------------------------------------
def cached_external_data(cma_zip, sel, n_rows):
    """Cached values for the external link so the grid shows data even before
    Excel refreshes the link. sheetId indexes into the CMA's sheet order."""
    wbxml = cma_zip.text("xl/workbook.xml")
    # names captured from raw XML are ALREADY xml-escaped ("US&amp;admi") —
    # escaping them again produces "&amp;amp;", which decodes to a 35-char
    # sheet name, breaches Excel's 31-char limit, and triggers a repair
    sheets_escaped = re.findall(r'<sheet name="([^"]+)"', wbxml)
    import html
    sheets = [html.unescape(s) for s in sheets_escaped]
    subj = sel["subject"]

    def cell(ref, v, string=False):
        if v is None or v == "":
            return ""
        if string:
            return f'<cell r="{ref}" t="str"><v>{esc(v)}</v></cell>'
        return f'<cell r="{ref}"><v>{v}</v></cell>'

    inputs_rows = "".join(
        f'<row r="{r}">' + cell(f"B{r}", v, isinstance(v, str)) + "</row>"
        for r, v in [(2, subj["name"]), (3, subj["address"]), (4, subj["city"]),
                     (5, subj["state"]), (6, subj["zip"]), (7, subj["year_built"]),
                     (8, subj["avg_size"]), (9, subj["units"])])
    oad_rows = []
    for i, c in enumerate(sel["comps"]):
        r = i + 2
        cells = []
        for j, colname in enumerate(OUT_COLS):
            L = col_letter(j + 1)
            v = c.get(colname)
            if colname == "Sale Date" and v:
                cells.append(cell(f"{L}{r}", date_serial(v)))
            elif colname in STR_COLS:
                cells.append(cell(f"{L}{r}", v, string=True))
            else:
                cells.append(cell(f"{L}{r}", v))
        oad_rows.append(f'<row r="{r}">' + "".join(cells) + "</row>")
    names = "".join(f'<sheetName val="{esc(s)}"/>' for s in sheets)
    # Excel writes a sheetData element for EVERY sheet of the external book
    # (empty ones self-closed) — omitting them triggers a repair of the part
    filled = {sheets.index("Inputs"): inputs_rows,
              sheets.index("Output Analysis Data"): "".join(oad_rows)}
    data = "".join(
        f'<sheetData sheetId="{i}">{filled[i]}</sheetData>' if i in filled
        else f'<sheetData sheetId="{i}"/>'
        for i in range(len(sheets)))
    return names, data


def build_grid(template, cma_zip, sel, cma_name, out_path):
    z = Zip(template)
    selected = [c for c in sel["comps"] if c["Rank"] in sel["selected_ranks"]]
    selected.sort(key=lambda c: sel["selected_ranks"].index(c["Rank"]))
    positions = {c["Rank"] + 2 for c in selected}          # helper-grid rows to mark
    if any(c["Rank"] > 50 for c in selected):
        sys.exit("ERROR: a selected comp ranks below 50 — raise the top-50 window.")

    # 1. retarget the external link (targets are URIs — spaces must be %-encoded
    # or Excel "repairs" the external link part)
    from urllib.parse import quote
    link = "xl/externalLinks/externalLink1.xml"
    xml0 = z.text(link)
    # The template carries Excel-2021 alternate URLs (xxl21): rId2 must be an
    # ABSOLUTE url and rId3 a relative one. We can't know the user's absolute
    # path, so drop the whole mechanism and keep a single relative link —
    # a relative path in <xxl21:absoluteUrl> makes Excel repair the part.
    xml0 = re.sub(r"<xxl21:alternateUrls>.*?</xxl21:alternateUrls>", "", xml0, flags=re.S)
    z.set(link, xml0)
    book_rid = re.search(r'<externalBook[^>]*r:id="(rId\d+)"', xml0).group(1)
    rels = "xl/externalLinks/_rels/externalLink1.xml.rels"
    rx = z.text(rels)
    rx = re.sub(r"<Relationship [^>]*/>",
                lambda m: m.group(0) if f'Id="{book_rid}"' in m.group(0) else "", rx)
    rx = re.sub(r'Target="[^"]*"', f'Target="{quote(cma_name)}"', rx)
    z.set(rels, rx)

    # 2. rebuild cached external book (sheet names + values from the new CMA)
    xml = z.text(link)
    names, data = cached_external_data(cma_zip, sel, len(sel["comps"]))
    xml = re.sub(r"<sheetNames>.*?</sheetNames>", lambda _: "<sheetNames>" + names + "</sheetNames>",
                 xml, count=1, flags=re.S)
    xml = re.sub(r"<sheetDataSet>.*</sheetDataSet>", lambda _: "<sheetDataSet>" + data + "</sheetDataSet>",
                 xml, count=1, flags=re.S)
    z.set(link, xml)

    # 3. defined names + force full recalc
    wb = z.text("xl/workbook.xml")
    for nm, ref in REDIQ_MAP.items():
        wb = re.sub(r'(<definedName name="%s"[^>]*>)[^<]+(</definedName>)' % nm,
                    lambda m: m.group(1) + esc(ref) + m.group(2), wb, count=1)
    # Every reference to a sheet that is NOT in the new external book's sheet
    # list is unresolvable — Excel treats the whole external-link part as
    # corrupt and "repairs" it. Drop the ~75 leftover underwriting-model names
    # (CoStarSale*, YardiSale*, rr*, Region*, [1]Master leftovers, ...).
    cma_sheets = set(re.findall(r'<sheet name="([^"]+)"', cma_zip.text("xl/workbook.xml")))

    def refs_missing_sheet(text):
        for sheet in re.findall(r"'?\[\d+\]([^'!\]]+)'?!", text.replace("&apos;", "'")):
            if sheet not in cma_sheets:
                return True
        return False

    wb = re.sub(r"<definedName [^>]*>[^<]*</definedName>",
                lambda m: "" if refs_missing_sheet(m.group(0)) else m.group(0), wb)
    if "<calcPr" in wb:
        wb = re.sub(r"<calcPr ", '<calcPr fullCalcOnLoad="1" ', wb, count=1)
    else:
        wb = wb.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>')
    z.set("xl/workbook.xml", wb)

    # 4. rewrite the Comparable Grid sheet
    part = z.sheet_part("Comparable Grid")
    xml = z.text(part)
    for r in HELPER_ROWS:
        for gc, oc in HELPER_MAP.items():
            xml = set_formula(xml, f"{gc}{r}", f"='[1]Output Analysis Data'!{oc}{r-1}")
        mark = "x" if r in positions else ""
        styles = row_styles(xml, r)
        xml = replace_cell(xml, f"L{r}", cell_xml(f"L{r}", mark, styles.get("L"), kind="str"))
    for ref, f in SUBJECT_MAP.items():
        xml = set_formula(xml, ref, "=" + f)

    # blank the Manual-mode agency/CoStar/Yardi probe columns — they reference
    # underwriting-model sheets/names that don't exist in the CMA link, and any
    # unresolvable external reference makes Excel repair the link part
    for r in HELPER_ROWS:
        for col in ("AA", "AB", "AC"):
            styles = row_styles(xml, r)
            xml = replace_cell(xml, f"{col}{r}", cell_xml(f"{col}{r}", "", styles.get(col)))

    # cap-rate drift row + current cap rate hook
    cap = sel["cap_rate"].get("current_avg_cap")
    xml = replace_cell(xml, "AH1", cell_xml("AH1", "Current Avg Cap Rate:", kind="str"))
    xml = replace_cell(xml, "AI1", cell_xml("AI1", round(cap, 6) if cap else "", kind="num"))
    for i, col in enumerate(["E", "F", "G", "H", "I"]):
        bps = selected[i]["CapRateDriftBps"] if i < len(selected) else 0
        styles = row_styles(xml, 25)
        xml = replace_cell(xml, f"{col}25", cell_xml(f"{col}25", bps, styles.get(col), kind="num"))
        adj = (f'=IFERROR(IF($AI$1="",0,((({col}15*($AI$1-{col}25/10000))/($AI$1))/{col}15-1)),"")')
        xml = set_formula(xml, f"{col}26", adj)

    # strip stale cached values from every remaining formula cell (fullCalcOnLoad recalcs)
    xml = re.sub(r"(<f[ >][^<]*(?:</f>|/>))<v>[^<]*</v>", r"\1", xml)
    xml = re.sub(r"(<f>[^<]*</f>)<v>[^<]*</v>", r"\1", xml)
    # the template's array formulas carried dynamic-array cell metadata
    # (cm=/vm= attrs -> xl/metadata.xml); our rewrites removed the formulas, so
    # strip any remaining attrs and the now-orphaned metadata part — Excel
    # treats the orphan as corruption and repairs
    xml = re.sub(r'(<c [^>]*?) (?:cm|vm)="\d+"', r"\1", xml)
    z.set(part, xml)

    if not re.search(r'<c [^>]*(?:cm|vm)="', xml):
        z.drop_part("xl/metadata.xml")
    z.drop_calc_chain()
    z.save(out_path)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--cma", required=True, help="source Automatic CMA Analysis .xlsx")
    ap.add_argument("--selection", required=True)
    ap.add_argument("--template", default=DEFAULT_TEMPLATE)
    ap.add_argument("--outdir", default="output")
    ap.add_argument("--cma-out-name", default="Automatic CMA Analysis.xlsx",
                    help="file name for the CMA copy (also the link target)")
    args = ap.parse_args()
    with open(args.selection) as f:
        sel = json.load(f)

    os.makedirs(args.outdir, exist_ok=True)
    cma_out = os.path.join(args.outdir, args.cma_out_name)
    n = build_cma(args.cma, sel, cma_out)
    grid_out = os.path.join(args.outdir, f"{sel['subject']['name']} - Sale Comparables.xlsx")
    build_grid(args.template, Zip(cma_out), sel, args.cma_out_name, grid_out)
    print(f"Wrote {cma_out} (Output Analysis Data: {n-1} comps)")
    print(f"Wrote {grid_out} (linked to '{args.cma_out_name}' — keep both files together)")


if __name__ == "__main__":
    main()
