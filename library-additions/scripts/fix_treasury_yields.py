#!/usr/bin/env python3
"""Neutralise the model's =WEBSERVICE() FRED pulls so LibreOffice can recalc headless.

WHY: 'Treasury Yields'!D4:D8 are `=IFERROR(_xludf.webservice(C4),"")` live pulls from
fred.stlouisfed.org. LibreOffice evaluates WEBSERVICE during calculateAll and blocks on
the socket -- soffice.bin sits at 0% CPU indefinitely and the driving macro never runs.
(Shady Oaks 8/7/2026: two runs killed after 10+ minutes each, 14 s CPU total.)

The tab is DESIGNED for this: E4:E8 parse the CSV response, and
G4:G8 = IF(ISNUMBER(E),E,F) fall back to the yellow manual-override column F whenever the
pull yields a non-number. Writing "" into D4:D8 therefore takes the documented fallback
path rather than defeating the model -- the rates used become exactly column F, which we
set from the loan-terms lookup and stamp with an as-of date in H.

INPUT   a TMG model .xlsx
OUTPUT  the same workbook with D4:D8 literal "" and F/H refreshed
Usage:  python3 fix_treasury_yields.py <in.xlsx> <out.xlsx>
"""
import sys
import warnings

warnings.filterwarnings("ignore")
import openpyxl

# Live indices from the 8/6/2026 loan-terms lookup (work/terms_8-6-2026.json).
# UST 7 Yr and UST 30 Yr were not re-pulled then; the template's own 8/5/2026
# overrides are kept for those two and labelled with their own as-of date.
OVERRIDES = {
    4: ("UST 5 Yr", 0.0440, "08/06/2026"),
    5: ("UST 7 Yr", None, "08/05/2026"),      # None = keep template value
    6: ("UST 10 Yr", 0.0469, "08/06/2026"),
    7: ("UST 30 Yr", None, "08/05/2026"),
    8: ("30-Day Avg SOFR", 0.03622, "08/06/2026"),
}


def main(src, dst):
    wb = openpyxl.load_workbook(src)
    ws = wb["Treasury Yields"]
    killed = 0
    for row, (label, rate, asof) in OVERRIDES.items():
        assert ws.cell(row, 1).value == label, \
            f"row {row} is {ws.cell(row,1).value!r}, expected {label!r}"
        ws.cell(row, 4).value = ""            # D: kill the WEBSERVICE call
        killed += 1
        if rate is not None:
            ws.cell(row, 6).value = rate      # F: manual override
            ws.cell(row, 8).value = asof      # H: as-of
        print(f"  {label:16s} rate used -> {ws.cell(row,6).value}  as-of {ws.cell(row,8).value}")

    left = sum(
        1 for w in wb.worksheets for r in w.iter_rows() for c in r
        if isinstance(c.value, str) and "webservice" in c.value.lower()
        and c.value.startswith("=")
    )
    wb.save(dst)
    print(f"cleared {killed} WEBSERVICE formula cells; {left} WEBSERVICE formulas remain")
    print(f"wrote {dst}")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
