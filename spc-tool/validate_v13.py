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
     Capability columns (USL really is the USL column, and so on);
  6. no formula compares a cell against NA(). `IF(A1=NA(), ...)` looks
     reasonable and is always wrong: comparing anything with the #N/A error
     yields #N/A, so the whole formula returns #N/A whatever A1 holds. This
     silently emptied the entire BinCnt column of the v12.1 Visual SPC
     histogram, so its bars never drew. The correct test is ISNA(A1).

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
# IF(A1=NA(), ...) - always #N/A regardless of A1. Use ISNA(A1).
NA_COMPARE_RE = re.compile(r"[A-Z]{1,3}\$?\d+\s*(?:=|<>|<|>)\s*NA\(\)")
# A formula that lost its leading "=" is stored as inert text: it displays as
# the formula source, evaluates to nothing, and silently kills everything
# downstream of it.
LOOKS_LIKE_FORMULA_RE = re.compile(
    r"^\s*(?:IF|INDEX|COUNT|COUNTIF|COUNTIFS|SUM|SUMPRODUCT|MAX|MIN|AVERAGE|"
    r"MEDIAN|MATCH|STDEV|ABS|ROUND|IFERROR|NA|ISNA|TEXT)\s*\(")


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
                if not isinstance(v, str):
                    continue
                if not v.startswith("="):
                    if LOOKS_LIKE_FORMULA_RE.match(v):
                        problems.append(
                            f"formula stored as text (missing leading '=') at "
                            f"{ws.title}!{cell.coordinate}: {v[:90]}")
                    continue
                counts[ws.title] += 1
                where = f"{ws.title}!{cell.coordinate}"
                if not balanced(v):
                    problems.append(f"unbalanced formula at {where}: {v[:90]}")
                if NA_COMPARE_RE.search(v):
                    problems.append(
                        f"compares a cell against NA() at {where} "
                        f"(always #N/A - use ISNA()): {v[:90]}")
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


def hidden_columns(ws):
    """Column letters hidden outright or by a collapsed outline group."""
    hidden = set()
    for key, dim in ws.column_dimensions.items():
        if not dim.hidden:
            continue
        if dim.min and dim.max:
            for i in range(dim.min, dim.max + 1):
                hidden.add(get_column_letter(i))
        else:
            hidden.add(key)
    return hidden


def check_charts(wb, problems):
    total = 0
    for ws in wb.worksheets:
        hidden_by_sheet = {s: hidden_columns(wb[s]) for s in wb.sheetnames}
        for chart in getattr(ws, "_charts", []):
            total += 1
            # Excel plots only visible cells unless told otherwise. Every helper
            # block in this workbook lives in hidden columns, so a chart left at
            # the default draws nothing at all - the failure looks exactly like
            # "my formulas are broken".
            if getattr(chart, "visible_cells_only", True):
                srcs = {f"{s}!{c}" for s, c, _, _ in chart_ranges(chart)
                        if c in hidden_by_sheet.get(s, ())}
                if srcs:
                    problems.append(
                        f"chart on {ws.title} reads hidden columns "
                        f"({', '.join(sorted(srcs)[:3])}) but visible_cells_only "
                        f"is True - it will render empty")
            # Limit spikes overlaid on a histogram have one non-#N/A point
            # each. A line series with a single point and no marker renders as
            # nothing at all, so the limits become invisible.
            # (openpyxl puts the chart itself in _charts, so skip it: a full
            # 200-point control-limit line needs no marker.)
            for sub in getattr(chart, "_charts", []):
                if sub is chart or type(chart).__name__ != "BarChart":
                    continue
                if type(sub).__name__ != "LineChart":
                    continue
                for s in sub.series:
                    sym = s.marker.symbol if s.marker else None
                    if sym in (None, "none"):
                        ref = s.val.numRef.f if s.val and s.val.numRef else "?"
                        problems.append(
                            f"overlay series {ref} on {ws.title} has no marker - "
                            f"a single-point series will render invisible")

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

    # Raw histogram bounds must span the data and every limit, so that the
    # spec spikes and the control spikes both land inside the plotted bins.
    for addr, name, needed in (("CO1", "RawLo", ("Min", "LSL", "X LCL")),
                               ("CQ1", "RawHi", ("Max", "USL", "X UCL"))):
        formula = cc[addr].value or ""
        for header in needed:
            if f"Capability!${headers[header]}$5" not in formula:
                problems.append(
                    f"Control Charts!{addr} ({name}) does not reference {header} "
                    f"- that limit could fall outside the histogram bins")
    # and the padded Lo/Hi must derive from those raw bounds
    for addr, raw in (("CG1", "CO1"), ("CM1", "CQ1")):
        if raw.rstrip("1") not in (cc[addr].value or ""):
            problems.append(f"Control Charts!{addr} does not derive from {raw}")


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
