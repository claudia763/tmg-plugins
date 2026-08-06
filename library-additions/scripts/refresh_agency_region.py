"""Refresh the TMG underwriting model's agency loan-sale pages for the subject's
own metro, replacing the previous deal's extract that ships in the template.

WHY: 'Agency Region' (table Region_Comps) and 'Agency-Data' (table ZipCodeComps)
are Power Query tables whose source is not available on a job machine, so they
never refresh -- the Git template ships loaded with whatever metro the last deal
used. Left alone they silently drive (a) 'Agency Loan-Sale Comps'!Z40, which is
the BASE of the terminal cap rate in Assumptions!G58, and (b) the PDF's
"Agency Loan-Sale Comparables (Geographic Region)" page, which would print
another metro's properties under this deal's search criteria.

The `sales-comps` skill's "Automatic CMA Analysis.xlsx" carries an `AgencyDrift`
sheet with the SAME 26-column schema (plus two trailing columns we drop), so it
is a drop-in source.

INPUTS
  --model    the working model .xlsx (edited in place through Excel COM)
  --cma      an Automatic CMA Analysis workbook containing an `AgencyDrift` sheet
  --region   substring matched against AgencyDrift's "TMG Region" column
             (e.g. "Houston", "Dallas-Fort Worth")
  --zips     comma-separated zips for the zip-cluster page (Agency-Data)
  --year-min/--year-max/--units-min/--units-max/--years-back
             the same criteria the sheet prints in its own search-criteria block
             (template defaults: 1950-2000, 0-500 units, 3 years)

OUTPUT
  The model is saved with both sheets repopulated and their tables resized, and
  the script prints the resulting regional average cap rate so you can sanity
  check it against the CMA's own AgencyDrift average for the same state/vintage.

NOTE: only the first 22 rows of 'Agency Region' feed the MAX/MIN/AVERAGE block at
rows 38-40 of 'Agency Loan-Sale Comps', so rows are sorted newest-origination
first.

Requires: pywin32 + desktop Excel, openpyxl.
"""
import argparse
import datetime
import warnings

warnings.filterwarnings("ignore")
import openpyxl
import win32com.client as win32
from win32com.client import constants as C

SCHEMA_COLS = 26          # A..Z, shared by AgencyDrift and both model sheets
COL_REGION, COL_ZIP, COL_ORIGIN = 4, 5, 6
COL_BUILT, COL_UNITS, COL_CAP = 10, 11, 13


def load_rows(cma_path, region, zips, year_min, year_max,
              units_min, units_max, years_back):
    wb = openpyxl.load_workbook(cma_path, data_only=True, read_only=True)
    ws = wb["AgencyDrift"]
    cutoff = datetime.datetime.now() - datetime.timedelta(days=365 * years_back)
    region_rows, zip_rows = [], []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] in (None, ""):
            continue
        rec = list(row[:SCHEMA_COLS])
        built, units, orig = rec[COL_BUILT], rec[COL_UNITS], rec[COL_ORIGIN]
        if not isinstance(built, (int, float)) or not isinstance(units, (int, float)):
            continue
        if not isinstance(orig, datetime.datetime):
            continue
        if not (year_min <= built <= year_max):
            continue
        if not (units_min <= units <= units_max):
            continue
        if orig < cutoff:
            continue
        rec[COL_REGION] = str(rec[COL_REGION] or "")
        try:
            zi = int(str(rec[COL_ZIP]).strip())
        except (TypeError, ValueError):
            zi = None
        if zi is not None:
            rec[COL_ZIP] = zi
        if region.lower() in rec[COL_REGION].lower():
            region_rows.append(rec)
            if zi in zips:
                zip_rows.append(rec)
    wb.close()
    region_rows.sort(key=lambda r: r[COL_ORIGIN], reverse=True)
    zip_rows.sort(key=lambda r: r[COL_ORIGIN], reverse=True)
    return region_rows, zip_rows


def write(model_path, region_rows, zip_rows):
    xl = win32.gencache.EnsureDispatch("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AskToUpdateLinks = False
    try:
        wb = xl.Workbooks.Open(model_path, UpdateLinks=0)
        xl.Calculation = C.xlCalculationManual

        AR = wb.Worksheets("Agency Region")            # header row 1, data row 2+
        AR.Range("A2:Z400").ClearContents()
        rows = region_rows[:325]
        if rows:
            AR.Range(AR.Cells(2, 1), AR.Cells(1 + len(rows), SCHEMA_COLS)).Value = \
                tuple(tuple(r) for r in rows)
            _resize(AR, "Region_Comps", f"A1:Z{1 + len(rows)}")

        AD = wb.Worksheets("Agency-Data")               # header row 2, data row 3+
        AD.Range("A3:Z400").ClearContents()
        zrows = zip_rows[:200]
        if zrows:
            AD.Range(AD.Cells(3, 1), AD.Cells(2 + len(zrows), SCHEMA_COLS)).Value = \
                tuple(tuple(r) for r in zrows)
            _resize(AD, "ZipCodeComps", f"A2:Z{2 + len(zrows)}")

        xl.Calculation = C.xlCalculationAutomatic
        xl.CalculateFullRebuild()
        z40 = wb.Worksheets("Agency Loan-Sale Comps").Range("Z40").Value
        wb.Save()
        wb.Close(SaveChanges=False)
        return z40
    finally:
        xl.Quit()


def _resize(ws, name, ref):
    try:
        ws.ListObjects(name).Resize(ws.Range(ref))
    except Exception as exc:                            # plain range in some builds
        print(f"   note: could not resize {name}: {exc}")


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--model", required=True)
    ap.add_argument("--cma", required=True)
    ap.add_argument("--region", required=True)
    ap.add_argument("--zips", default="")
    ap.add_argument("--year-min", type=int, default=1950)
    ap.add_argument("--year-max", type=int, default=2000)
    ap.add_argument("--units-min", type=int, default=0)
    ap.add_argument("--units-max", type=int, default=500)
    ap.add_argument("--years-back", type=int, default=3)
    a = ap.parse_args()

    zips = {int(z) for z in a.zips.replace(" ", "").split(",") if z}
    region_rows, zip_rows = load_rows(a.cma, a.region, zips, a.year_min, a.year_max,
                                      a.units_min, a.units_max, a.years_back)
    caps = [r[COL_CAP] for r in region_rows[:22] if isinstance(r[COL_CAP], float)]
    print(f"region rows {len(region_rows)}, zip-cluster rows {len(zip_rows)}")
    if caps:
        print(f"top-22 cap rates: n={len(caps)} avg={sum(caps)/len(caps):.4%} "
              f"min={min(caps):.4%} max={max(caps):.4%}")
    if not region_rows:
        raise SystemExit("no rows matched -- check --region against AgencyDrift's "
                         "'TMG Region' values before writing anything")
    z40 = write(a.model, region_rows, zip_rows)
    print(f"'Agency Loan-Sale Comps'!Z40 (terminal-cap base) is now {z40:.4%}")


if __name__ == "__main__":
    main()
