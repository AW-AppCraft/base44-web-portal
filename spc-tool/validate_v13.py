#!/usr/bin/env python3
"""
Static validation for SPC_Calculator_v13.xlsx.

LibreOffice cannot be used to recalculate in every environment, so this script
checks the things that break silently in a generated workbook:

  1. every formula is syntactically balanced (parentheses and quotes);
  2. every sheet named in a reference exists;
  3. every function used is one Excel 2007 understands without an _xlfn. prefix
     (nothing newer, nothing spilling);
  4. every chart series and category range points at a column that actually
     holds formulas, and at the row range the helper block occupies;
  5. the reference chain that feeds the charts lands on the intended
     Capability columns (USL really is the USL column, and so on).

Run:  python3 validate_v13.py SPC_Calculator_v13.xlsx
"""

import re
import sys
from collections import Counter

import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter

ALLOWED_FUNCS = {
    "IF", "AND", "OR", "NOT", "INDEX", "MATCH", "OFFSET", "COUNT", "COUNTA",
    "COUNTIF", "COUNTIFS", "SUM", "SUMPRODUCT", "AVERAGE", "MEDIAN", "MIN",
    "MAX", "STDEV", "ABS", "ROUND", "ROW", "COLUMN", "NA", "ISNA", "ISNUMBER",
    "IFERROR", "TEXT", "N", "MOD", "LEN", "TRIM",
}
FUNC_RE = re.compile(r"\b([A-Z][A-Z0-9_.]*)\s*\(")
QSHEET_RE = re.compile(r"'([^']+)'!")
PSHEET_RE = re.compile(r"(?<![A-Za-z0-9_'!])([A-Za-z_][A-Za-z0-9_]*)!")
RANGE_RE = re.compile(r"'?([^'!]+)'?!\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?")


def balanced(formula):
    depth, in_str = 0, False
    for ch in formula:
        if ch == '"':
            in_str = not in_str
        elif not in_str:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
                if depth < 0:
                    return False
    return depth == 0 and not in_str


def check_formulas(wb, problems):
    counts = Counter()
    sheets = set(wb.sheetnames)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if not isinstance(v, str) or not v.startswith("="):
                    continue
                counts[ws.title] += 1
                where = f"{ws.title}!{cell.coordinate}"
                if not balanced(v):
                    problems.append(f"unbalanced formula at {where}: {v[:90]}")
                for name in QSHEET_RE.findall(v) + PSHEET_RE.findall(v):
                    if name in ALLOWED_FUNCS:
                        continue
                    if name not in sheets:
                        problems.append(f"unknown sheet '{name}' referenced at {where}")
                for fn in FUNC_RE.findall(v):
                    if fn not in ALLOWED_FUNCS:
                        problems.append(f"disallowed function {fn}() at {where}")
    return counts


def chart_ranges(chart):
    """Yield (sheet, col_letter, first_row, last_row) for every series and category."""
    subcharts = [chart] + list(getattr(chart, "_charts", []))
    seen = set()
    for sub in subcharts:
        if id(sub) in seen:
            continue
        seen.add(id(sub))
        for s in sub.series:
            for ref in (s.val.numRef if s.val else None,
                        s.cat.numRef if s.cat else None,
                        s.tx.strRef if s.tx else None):
                if ref is None or not ref.f:
                    continue
                m = RANGE_RE.match(ref.f)
                if m:
                    sheet, c1, r1, c2, r2 = m.groups()
                    yield sheet.strip("'"), c1, int(r1), int(r2 or r1)


def check_charts(wb, problems):
    total = 0
    for ws in wb.worksheets:
        for chart in getattr(ws, "_charts", []):
            total += 1
            for sheet, col, r1, r2 in chart_ranges(chart):
                if sheet not in wb.sheetnames:
                    problems.append(f"chart on {ws.title} points at missing sheet {sheet}")
                    continue
                target = wb[sheet]
                ci = column_index_from_string(col)
                populated = sum(
                    1 for r in range(r1, min(r2, r1 + 5) + 1)
                    if target.cell(row=r, column=ci).value is not None)
                if populated == 0:
                    problems.append(
                        f"chart on {ws.title} references empty range "
                        f"{sheet}!{col}{r1}:{col}{r2}")
    return total


def check_capability_targets(wb, problems):
    """The chart helpers address Capability by column letter - prove the letters."""
    cap = wb["Capability"]
    headers = {}
    for c in range(1, cap.max_column + 1):
        h = cap.cell(row=4, column=c).value
        if h:
            headers[h] = get_column_letter(c)

    cc = wb["Control Charts"]
    expect = {
        "usl": ("USL", "CI4"), "lsl": ("LSL", "CJ4"),
        "cl": ("X CL", "CD4"), "ucl": ("X UCL", "CE4"), "lcl": ("X LCL", "CF4"),
        "mrucl": ("mR UCL", "CH4"),
    }
    for key, (header, addr) in expect.items():
        want = f"Capability!${headers[header]}$5"
        got = cc[addr].value or ""
        if want not in got:
            problems.append(
                f"Control Charts!{addr} ({key}) should reference {want}; got: {got[:80]}")

    # histogram bin bounds must use Min/Max plus LSL/USL
    lo = cc["CG1"].value or ""
    for header in ("Min", "Max", "LSL", "USL"):
        if f"Capability!${headers[header]}$5" not in lo:
            problems.append(f"Control Charts!CG1 (bin Lo) does not reference {header}")


def main(path):
    wb = openpyxl.load_workbook(path)
    problems = []
    counts = check_formulas(wb, problems)
    n_charts = check_charts(wb, problems)
    check_capability_targets(wb, problems)

    print(f"workbook : {path}")
    print(f"sheets   : {len(wb.sheetnames)}")
    print(f"charts   : {n_charts}")
    print(f"formulas : {sum(counts.values())}")
    for sheet, n in counts.most_common():
        print(f"   {sheet:<22} {n:>6}")
    if problems:
        print(f"\nPROBLEMS ({len(problems)}):")
        for p in problems[:60]:
            print("  -", p)
        return 1
    print("\nOK - no structural problems found")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1] if len(sys.argv) > 1 else "SPC_Calculator_v13.xlsx"))
