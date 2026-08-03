#!/usr/bin/env python3
"""
SPC Calculator v13 - All-in-One build script
=============================================
Regenerates SPC_Calculator_v13.xlsx from scratch.

Upgrades over v12.1:
  * Control Charts / Monitor: a live statistics panel (incl. data range) sits to the
    LEFT of every graph.
  * X charts carry USL / LSL lines alongside UCL / CL / LCL.
  * Histogram chart per feature, with LSL/USL spikes overlaid.
  * Rolling chart window - graphs follow the last N points instead of the first 200.
  * "Data Collection Form" - a print-ready sheet for recording readings off-PC.
  * "Visual SPC" analysis sheet - date/shift/operator/tool filters drive every plot.
  * Monitor sheet carries folder-watch configuration + an all-feature alarm table
    for the companion VBA module (SPC_V13_Monitor.bas).

Run:  python3 build_spc_v13.py [output.xlsx]
"""

import sys
import datetime as _dt

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference, Series
from openpyxl.chart.marker import Marker
from openpyxl.chart.data_source import NumDataSource, NumRef
from openpyxl.drawing.line import LineProperties
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as CL
from openpyxl.worksheet.datavalidation import DataValidation

# --------------------------------------------------------------------------
# Configuration
# --------------------------------------------------------------------------
VERSION = "v13"
DOC_ID = "RG-SPC-013"
REVISION = "R01"
COMPANY = "R&G Precision Engineering Ltd"
OWNER = "Andrij Medwid, Senior Quality Engineer"
ISSUE_DATE = "2026-08-03"

MAX_FEATURES = 50          # feature definition rows on Settings (17..66)
CHART_FEATURES = 10        # features that get a full chart block
DATA_FIRST = 5             # first measurement row on Data Entry
DATA_ROWS = 10000
DATA_LAST = DATA_FIRST + DATA_ROWS - 1          # 10004
FEAT_COL_FIRST = 8         # column H
FEAT_COL_LAST = FEAT_COL_FIRST + MAX_FEATURES - 1   # column BE
SET_ROW_FIRST = 17         # Settings feature rows
WINDOW = 200               # points plotted per chart
NBINS = 15                 # histogram bins
FORM_ROWS = 25             # blank rows on the printable collection form

# Statistics block on Data Entry
ST = {
    "count": 10009, "mean": 10010, "min": 10011, "max": 10012, "range": 10013,
    "stdev": 10014, "mrbar": 10015, "sigma": 10016, "ucl": 10017, "cl": 10018,
    "lcl": 10019, "mrucl": 10020, "median": 10021, "lastrow": 10022,
}
STAT_ORDER = [
    ("count", "Count (n)"), ("mean", "Mean (X-bar)"), ("min", "Min"), ("max", "Max"),
    ("range", "Range (Max-Min)"), ("stdev", "Std Dev (overall)"), ("mrbar", "mR-bar"),
    ("sigma", "Sigma (within)"), ("ucl", "UCL (X)"), ("cl", "CL (X)"),
    ("lcl", "LCL (X)"), ("mrucl", "UCL (mR)"), ("median", "Median"),
    ("lastrow", "Last data row"),
]

# --------------------------------------------------------------------------
# Styling
# --------------------------------------------------------------------------
FONT = "Arial"
NAVY = "1F3864"
BLUE = "2E74B5"
LIGHT = "DCE6F1"
GREY = "F2F2F2"
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")     # yellow = user input
CALC_FILL = PatternFill("solid", fgColor="F2F2F2")
HEAD_FILL = PatternFill("solid", fgColor=NAVY)
SUB_FILL = PatternFill("solid", fgColor=LIGHT)
BAND_FILL = PatternFill("solid", fgColor="E2EFDA")

F_TITLE = Font(name=FONT, size=16, bold=True, color="FFFFFF")
F_SUB = Font(name=FONT, size=10, italic=True, color="404040")
F_HEAD = Font(name=FONT, size=10, bold=True, color="FFFFFF")
F_SECTION = Font(name=FONT, size=11, bold=True, color=NAVY)
F_LABEL = Font(name=FONT, size=10, bold=True)
F_BODY = Font(name=FONT, size=10)
F_INPUT = Font(name=FONT, size=10, color="0000FF")
F_SMALL = Font(name=FONT, size=9)

THIN = Side(style="thin", color="A6A6A6")
MED = Side(style="medium", color="404040")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BOX_MED = Border(left=MED, right=MED, top=MED, bottom=MED)

NUM4 = "0.0000"
NUM3 = "0.000"
NUM2 = "0.00"
NUM0 = "0"

# Series colours
C_DATA = "2E74B5"
C_CL = "00B050"
C_CTRL = "C00000"
C_SPEC = "7030A0"
C_BAR = "9DC3E6"


def title_block(ws, title, subtitle, width=12):
    """Standard banner across the top of a sheet."""
    ws["A1"] = title
    ws["A1"].font = F_TITLE
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=width)
    for c in range(1, width + 1):
        ws.cell(row=1, column=c).fill = HEAD_FILL
    ws.row_dimensions[1].height = 24
    ws["A2"] = subtitle
    ws["A2"].font = F_SUB


def label(ws, row, col, text, font=F_LABEL, fill=None, border=None, align=None):
    c = ws.cell(row=row, column=col, value=text)
    c.font = font
    if fill:
        c.fill = fill
    if border:
        c.border = border
    if align:
        c.alignment = align
    return c


def col_of(feature):
    """Data Entry column index for feature number (1-based)."""
    return FEAT_COL_FIRST + feature - 1


def de_letter(feature):
    return CL(col_of(feature))


def cap_row(feature):
    """Capability sheet row for a feature."""
    return 4 + feature


def styled_line(series, colour, width=18000, dash=None, smooth=False, marker=False):
    series.graphicalProperties.line = LineProperties(solidFill=colour, w=width,
                                                     prstDash=dash)
    series.smooth = smooth
    if marker:
        series.marker = Marker(symbol="circle", size=5)
    else:
        series.marker = Marker(symbol="none")
    return series


# ==========================================================================
# 1. SETTINGS
# ==========================================================================
def build_settings(wb, demo_specs):
    ws = wb.create_sheet("Settings")
    title_block(ws, f"SPC CALCULATOR {VERSION} - SETTINGS",
                "Yellow cells are yours to edit. Everything else is calculated.", 7)

    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 34
    for c in "DE":
        ws.column_dimensions[c].width = 14
    ws.column_dimensions["F"].width = 24
    ws.column_dimensions["G"].width = 12

    # -- study parameters -------------------------------------------------
    for col, txt in ((1, "PARAMETER"), (2, "VALUE"), (3, "NOTES")):
        label(ws, 4, col, txt, F_HEAD, HEAD_FILL, BOX)

    params = [
        ("Number of Features", CHART_FEATURES, "Features with full chart blocks (1-%d)." % CHART_FEATURES),
        ("Data Rows Available", DATA_ROWS, "Measurement rows on Data Entry."),
        ("Feature Start Column", "H", "First measurement column on Data Entry."),
        ("Company", COMPANY, ""),
        ("Document ID", DOC_ID, ""),
        ("Revision", REVISION, ""),
        ("Issue Date", ISSUE_DATE, ""),
        ("Owner", OWNER, ""),
        ("Status", "DRAFT", "DRAFT / ISSUED / SUPERSEDED"),
    ]
    for i, (name, val, note) in enumerate(params):
        r = 5 + i
        label(ws, r, 1, name, F_LABEL, None, BOX)
        c = ws.cell(row=r, column=2, value=val)
        c.font = F_INPUT
        c.fill = INPUT_FILL
        c.border = BOX
        label(ws, r, 3, note, F_SMALL, None, BOX)

    # -- display & monitor options ---------------------------------------
    label(ws, 4, 5, "DISPLAY & MONITOR OPTIONS", F_HEAD, HEAD_FILL, BOX)
    ws.merge_cells(start_row=4, start_column=5, end_row=4, end_column=6)
    ws.cell(row=4, column=6).fill = HEAD_FILL

    opts = [
        ("Chart Window (points)", WINDOW),
        ("Histogram Bins", NBINS),
        ("Show Spec Limits", "Y"),
        ("Watch Folder", ""),
        ("File Pattern", "*.csv"),
        ("Auto-Refresh (sec)", 60),
        ("Archive Imported Files", "Y"),
        ("Archive Subfolder", "archive"),
        ("Kiosk Cycle (sec)", 20),
    ]
    for i, (name, val) in enumerate(opts):
        r = 5 + i
        label(ws, r, 5, name, F_LABEL, None, BOX)
        c = ws.cell(row=r, column=6, value=val)
        c.font = F_INPUT
        c.fill = INPUT_FILL
        c.border = BOX
    ws.cell(row=14, column=5, value="Charts plot the most recent 'Chart Window' points.")
    ws.cell(row=14, column=5).font = F_SMALL

    # -- feature definitions ---------------------------------------------
    label(ws, 15, 1, "FEATURE DEFINITIONS", F_SECTION)
    heads = ["#", "Feature Name", "Nominal", "LSL", "USL", "Unit", "Active"]
    for i, h in enumerate(heads):
        label(ws, 16, 1 + i, h, F_HEAD, HEAD_FILL, BOX,
              Alignment(horizontal="center"))

    for f in range(1, MAX_FEATURES + 1):
        r = SET_ROW_FIRST + f - 1
        active = "Y" if f <= CHART_FEATURES else "N"
        ws.cell(row=r, column=1, value=f).font = F_BODY
        vals = [f"Feature {f}", None, None, None, "mm", active]
        if f in demo_specs:
            nom, lsl, usl = demo_specs[f]
            vals = [f"Feature {f}", nom, lsl, usl, "mm", active]
        for i, v in enumerate(vals):
            c = ws.cell(row=r, column=2 + i, value=v)
            c.font = F_INPUT
            c.fill = INPUT_FILL
            c.border = BOX
            if i in (1, 2, 3):
                c.number_format = NUM4
        ws.cell(row=r, column=1).border = BOX
        ws.cell(row=r, column=1).alignment = Alignment(horizontal="center")

    dv = DataValidation(type="list", formula1='"Y,N"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add(f"G{SET_ROW_FIRST}:G{SET_ROW_FIRST + MAX_FEATURES - 1}")

    ws.freeze_panes = "A17"
    return ws


# ==========================================================================
# 2. DATA ENTRY
# ==========================================================================
def build_data_entry(wb, demo_rows):
    ws = wb.create_sheet("Data Entry")
    title_block(ws, f"SPC CALCULATOR {VERSION} - DATA ENTRY",
                "Type measurements from column H onward. Every chart in the workbook "
                "follows this sheet.", 12)

    widths = {"A": 10, "B": 10, "C": 12, "D": 8, "E": 11, "F": 10, "G": 10}
    for k, v in widths.items():
        ws.column_dimensions[k].width = v
    for f in range(1, MAX_FEATURES + 1):
        ws.column_dimensions[de_letter(f)].width = 12

    heads = ["Sample #", "Subgroup", "Date", "Shift", "Operator", "Tool ID", "Cluster"]
    for i, h in enumerate(heads):
        label(ws, 4, 1 + i, h, F_HEAD, HEAD_FILL, BOX,
              Alignment(horizontal="center", wrap_text=True))
    for f in range(1, MAX_FEATURES + 1):
        sr = SET_ROW_FIRST + f - 1
        c = ws.cell(row=4, column=col_of(f),
                    value=f'=IF(Settings!G{sr}="Y",Settings!B{sr},"")')
        c.font = F_HEAD
        c.fill = HEAD_FILL
        c.border = BOX
        c.alignment = Alignment(horizontal="center", wrap_text=True)

    for r in range(DATA_FIRST, DATA_LAST + 1):
        ws.cell(row=r, column=1, value=f'=IF(H{r}="","",ROW()-4)').font = F_SMALL
        ws.cell(row=r, column=2, value=f'=IF(H{r}="","",ROW()-4)').font = F_SMALL

    # demo data
    for i, row in enumerate(demo_rows):
        r = DATA_FIRST + i
        for j, v in enumerate(row):
            c = ws.cell(row=r, column=3 + j, value=v)
            c.font = F_BODY
            if j >= 5:
                c.number_format = NUM4

    ws.freeze_panes = "H5"
    ws.auto_filter.ref = f"A4:{CL(FEAT_COL_LAST)}{DATA_LAST}"

    # ---- statistics block ----------------------------------------------
    label(ws, ST["count"] - 2, 1, "STATISTICS (auto-calculated)", F_SECTION)
    label(ws, ST["count"] - 1, 1, "Statistic", F_HEAD, HEAD_FILL, BOX)
    for f in range(1, MAX_FEATURES + 1):
        c = ws.cell(row=ST["count"] - 1, column=col_of(f),
                    value=f"={de_letter(f)}4")
        c.font = F_HEAD
        c.fill = HEAD_FILL
        c.border = BOX

    for key, text in STAT_ORDER:
        label(ws, ST[key], 1, text, F_LABEL, SUB_FILL, BOX)

    for f in range(1, MAX_FEATURES + 1):
        L = de_letter(f)
        rng = f"'Data Entry'!${L}${DATA_FIRST}:${L}${DATA_LAST}"
        prev = f"'Data Entry'!${L}${DATA_FIRST}:${L}${DATA_LAST - 1}"
        nxt = f"'Data Entry'!${L}${DATA_FIRST + 1}:${L}${DATA_LAST}"
        # NB: the leading "=" belongs here - every formula built from `guard`
        # is a complete cell formula, not a fragment.
        guard = f'=IF(COUNT({rng})=0,""'
        formulas = {
            "count": f"=COUNT({rng})",
            "mean": f"{guard},AVERAGE({rng}))",
            "min": f"{guard},MIN({rng}))",
            "max": f"{guard},MAX({rng}))",
            "range": f"{guard},MAX({rng})-MIN({rng}))",
            "median": f"{guard},MEDIAN({rng}))",
            "stdev": f'=IF(COUNT({rng})<2,"",STDEV({rng}))',
            "mrbar": (f'=IF(COUNT({rng})<2,"",SUMPRODUCT(({nxt}<>"")*({prev}<>"")'
                      f"*ABS({nxt}-{prev}))/MAX(1,SUMPRODUCT(({nxt}<>\"\")*({prev}<>\"\")*1)))"),
            "sigma": f'=IF({L}{ST["mrbar"]}="","",{L}{ST["mrbar"]}/1.128)',
            "ucl": f'=IF({L}{ST["mrbar"]}="","",{L}{ST["mean"]}+2.66*{L}{ST["mrbar"]})',
            "cl": f'=IF({L}{ST["mean"]}="","",{L}{ST["mean"]})',
            "lcl": f'=IF({L}{ST["mrbar"]}="","",{L}{ST["mean"]}-2.66*{L}{ST["mrbar"]})',
            "mrucl": f'=IF({L}{ST["mrbar"]}="","",3.267*{L}{ST["mrbar"]})',
            "lastrow": f"=IFERROR(MATCH(9.99999999999999E+307,{rng})+{DATA_FIRST - 1},{DATA_FIRST - 1})",
        }
        for key, _ in STAT_ORDER:
            c = ws.cell(row=ST[key], column=col_of(f), value=formulas[key])
            c.font = F_BODY
            c.fill = CALC_FILL
            c.border = BOX
            c.number_format = NUM0 if key in ("count", "lastrow") else NUM4
    return ws


# ==========================================================================
# 3. CAPABILITY
# ==========================================================================
CAP_COLS = [
    ("Feature", 20), ("n", 7), ("Mean", 11), ("Median", 11), ("Min", 11),
    ("Max", 11), ("Range", 11), ("Sigma (within)", 13), ("Sigma (overall)", 13),
    ("X UCL", 11), ("X CL", 11), ("X LCL", 11), ("mR-bar", 11), ("mR UCL", 11),
    ("Nominal", 11), ("LSL", 11), ("USL", 11), ("Tol Width", 11),
    ("Cp", 9), ("Cpk", 9), ("Pp", 9), ("Ppk", 9),
    ("In Control?", 12), ("In Spec?", 11), ("Verdict", 16),
]
# column letters by header name
CAPC = {h: CL(i + 1) for i, (h, _) in enumerate(CAP_COLS)}


def build_capability(wb):
    ws = wb.create_sheet("Capability")
    title_block(ws, "PROCESS CAPABILITY - LIVE FROM DATA",
                "Cpk >= 1.33 green, 1.00-1.33 amber, < 1.00 red. "
                "Blank cells mean the feature has no spec limits yet.",
                len(CAP_COLS))

    for i, (h, w) in enumerate(CAP_COLS):
        ws.column_dimensions[CL(i + 1)].width = w
        label(ws, 4, i + 1, h, F_HEAD, HEAD_FILL, BOX,
              Alignment(horizontal="center", wrap_text=True))
    ws.row_dimensions[4].height = 28

    for f in range(1, MAX_FEATURES + 1):
        r = cap_row(f)
        sr = SET_ROW_FIRST + f - 1
        L = de_letter(f)
        rng = f"'Data Entry'!${L}${DATA_FIRST}:${L}${DATA_LAST}"
        S = lambda k: f"'Data Entry'!{L}{ST[k]}"  # noqa: E731

        vals = {
            "Feature": f'=IF(Settings!G{sr}="Y",Settings!B{sr},"")',
            "n": f"={S('count')}",
            "Mean": f"={S('mean')}",
            "Median": f"={S('median')}",
            "Min": f"={S('min')}",
            "Max": f"={S('max')}",
            "Range": f"={S('range')}",
            "Sigma (within)": f"={S('sigma')}",
            "Sigma (overall)": f"={S('stdev')}",
            "X UCL": f"={S('ucl')}",
            "X CL": f"={S('cl')}",
            "X LCL": f"={S('lcl')}",
            "mR-bar": f"={S('mrbar')}",
            "mR UCL": f"={S('mrucl')}",
            "Nominal": f'=IF(Settings!C{sr}="","",Settings!C{sr})',
            "LSL": f'=IF(Settings!D{sr}="","",Settings!D{sr})',
            "USL": f'=IF(Settings!E{sr}="","",Settings!E{sr})',
        }
        c_usl, c_lsl = f"{CAPC['USL']}{r}", f"{CAPC['LSL']}{r}"
        c_mean, c_sw = f"{CAPC['Mean']}{r}", f"{CAPC['Sigma (within)']}{r}"
        c_so = f"{CAPC['Sigma (overall)']}{r}"
        nospec = f'OR({c_usl}="",{c_lsl}="")'
        vals["Tol Width"] = f'=IF({nospec},"",{c_usl}-{c_lsl})'
        vals["Cp"] = (f'=IF(OR({nospec},{c_sw}="",{c_sw}=0),"",'
                      f"({c_usl}-{c_lsl})/(6*{c_sw}))")
        vals["Cpk"] = (f'=IF(OR({nospec},{c_sw}="",{c_sw}=0),"",'
                       f"MIN({c_usl}-{c_mean},{c_mean}-{c_lsl})/(3*{c_sw}))")
        vals["Pp"] = (f'=IF(OR({nospec},{c_so}="",{c_so}=0),"",'
                      f"({c_usl}-{c_lsl})/(6*{c_so}))")
        vals["Ppk"] = (f'=IF(OR({nospec},{c_so}="",{c_so}=0),"",'
                       f"MIN({c_usl}-{c_mean},{c_mean}-{c_lsl})/(3*{c_so}))")
        vals["In Control?"] = (
            f'=IF({CAPC["mR-bar"]}{r}="","",IF(OR(MAX({rng})>{CAPC["X UCL"]}{r},'
            f'MIN({rng})<{CAPC["X LCL"]}{r}),"OUT","IN"))')
        vals["In Spec?"] = (
            f'=IF(OR({nospec},{CAPC["n"]}{r}=0),"",IF(OR(MAX({rng})>{c_usl},'
            f'MIN({rng})<{c_lsl}),"FAIL","PASS"))')
        vals["Verdict"] = (
            f'=IF({CAPC["Cpk"]}{r}="","No spec",'
            f'IF({CAPC["Cpk"]}{r}>=1.33,"Capable",'
            f'IF({CAPC["Cpk"]}{r}>=1,"Marginal","Not capable")))')

        for ci, (h, _) in enumerate(CAP_COLS):
            c = ws.cell(row=r, column=ci + 1, value=vals[h])
            c.font = F_BODY
            c.border = BOX
            if h in ("Cp", "Cpk", "Pp", "Ppk", "Tol Width"):
                c.number_format = NUM2
            elif h in ("n",):
                c.number_format = NUM0
            elif h not in ("Feature", "In Control?", "In Spec?", "Verdict"):
                c.number_format = NUM4
            else:
                c.alignment = Alignment(horizontal="center")

    last = cap_row(MAX_FEATURES)
    cpk = f"{CAPC['Cpk']}5:{CAPC['Cpk']}{last}"
    ws.conditional_formatting.add(cpk, CellIsRule(
        operator="greaterThanOrEqual", formula=["1.33"],
        fill=PatternFill("solid", fgColor="C6EFCE"), font=Font(name=FONT, color="006100")))
    ws.conditional_formatting.add(cpk, CellIsRule(
        operator="between", formula=["1", "1.33"],
        fill=PatternFill("solid", fgColor="FFEB9C"), font=Font(name=FONT, color="9C6500")))
    ws.conditional_formatting.add(cpk, CellIsRule(
        operator="lessThan", formula=["1"],
        fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(name=FONT, color="9C0006")))
    for colname in ("In Control?", "In Spec?"):
        rng = f"{CAPC[colname]}5:{CAPC[colname]}{last}"
        ws.conditional_formatting.add(rng, CellIsRule(
            operator="equal", formula=['"OUT"'] if colname == "In Control?" else ['"FAIL"'],
            fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(name=FONT, color="9C0006")))
        ws.conditional_formatting.add(rng, CellIsRule(
            operator="equal", formula=['"IN"'] if colname == "In Control?" else ['"PASS"'],
            fill=PatternFill("solid", fgColor="C6EFCE"), font=Font(name=FONT, color="006100")))
    ws.freeze_panes = "B5"
    return ws


# ==========================================================================
# Shared: helper block + stats panel + charts
# ==========================================================================
HELPER_OFFSETS = {
    "idx": 0, "x": 1, "cl": 2, "ucl": 3, "lcl": 4, "mr": 5, "mrucl": 6,
    "usl": 7, "lsl": 8, "bin": 9, "cnt": 10, "lslmark": 11, "uslmark": 12,
}
HELPER_WIDTH = 14
HELPER_HEAD_ROW = 3
HELPER_FIRST = 4
HELPER_LAST = HELPER_FIRST + WINDOW - 1


def bin_bounds(min_ref, max_ref, lsl_ref, usl_ref):
    """Histogram bin limits: data min/max widened to cover LSL/USL, padded 5%.

    Returns (lo_formula, hi_formula). Both give "" when there is no data, or
    when every reading is identical (a zero-width range has no useful bins).
    """
    base_lo = f'IF({lsl_ref}="",{min_ref},MIN({min_ref},{lsl_ref}))'
    base_hi = f'IF({usl_ref}="",{max_ref},MAX({max_ref},{usl_ref}))'
    span = f"(({base_hi})-({base_lo}))"
    guard = f'IF(OR({min_ref}="",{max_ref}="",{span}<=0),""'
    return (f"={guard},({base_lo})-{span}*0.05)",
            f"={guard},({base_hi})+{span}*0.05)")


def write_helper_block(ws, base_col, refs, window_ref, hist_range, hbin_ref):
    """Write one 14-column helper block.

    refs: dict of formula fragments -> lastrow, x_value(k), cl, ucl, lcl,
          mrucl, usl, lsl  (all as *formula text without leading '=')
    """
    B = lambda k: CL(base_col + HELPER_OFFSETS[k])  # noqa: E731

    # scalar row 1: LastRow / Start / Lo / Width / MaxCnt / Hi
    lo_c = CL(base_col + 5)
    hi_c = CL(base_col + 11)
    wd_c = CL(base_col + 7)
    mx_c = CL(base_col + 9)
    lr_c = CL(base_col + 1)
    st_c = CL(base_col + 3)

    lo, hi = bin_bounds(refs["min"], refs["max"], refs["lsl"], refs["usl"])
    scal_vals = {
        "LastRow": f'={refs["lastrow"]}',
        "Start": f"=MAX({DATA_FIRST},{lr_c}1-{window_ref}+1)",
        "Lo": lo,
        "Hi": hi,
        "Width": (f'=IF(OR({lo_c}1="",{hi_c}1="",{hi_c}1<={lo_c}1),"",'
                  f"({hi_c}1-{lo_c}1)/{hbin_ref})"),
        "MaxCnt": f"=MAX({B('cnt')}{HELPER_FIRST}:{B('cnt')}{HELPER_FIRST+NBINS-1})",
    }
    order = ["LastRow", "Start", "Lo", "Width", "MaxCnt", "Hi"]
    for i, name in enumerate(order):
        ws.cell(row=1, column=base_col + i * 2, value=name).font = F_SMALL
        c = ws.cell(row=1, column=base_col + i * 2 + 1, value=scal_vals[name])
        c.font = F_SMALL
        c.number_format = NUM4

    heads = ["idx", "X", "CL", "UCL", "LCL", "mR", "mR_UCL", "USL", "LSL",
             "Bin", "Count", "LSLmark", "USLmark", ""]
    for i, h in enumerate(heads):
        ws.cell(row=HELPER_HEAD_ROW, column=base_col + i, value=h).font = F_SMALL

    for k in range(WINDOW):
        r = HELPER_FIRST + k
        pos = f"({st_c}$1+{k})"                       # source row on Data Entry
        live = f"IF({pos}>{lr_c}$1,NA()"
        ws.cell(row=r, column=base_col + HELPER_OFFSETS["idx"],
                value=f"={live},{pos}-{DATA_FIRST - 1})")
        ws.cell(row=r, column=base_col + HELPER_OFFSETS["x"],
                value=f"={live},{refs['x'](pos)})")
        na_guard = f"IF(ISNA({B('idx')}{r}),NA()"
        for key, src in (("cl", "cl"), ("ucl", "ucl"), ("lcl", "lcl"),
                         ("mrucl", "mrucl")):
            ws.cell(row=r, column=base_col + HELPER_OFFSETS[key],
                    value=f"={na_guard},{refs[src]})")
        for key in ("usl", "lsl"):
            ws.cell(row=r, column=base_col + HELPER_OFFSETS[key],
                    value=(f"={na_guard},IF(OR({refs[key]}=\"\","
                           f"Settings!$F$7<>\"Y\"),NA(),{refs[key]}))"))
        if k == 0:
            ws.cell(row=r, column=base_col + HELPER_OFFSETS["mr"], value="=NA()")
        else:
            xc, xp = f"{B('x')}{r}", f"{B('x')}{r-1}"
            ws.cell(row=r, column=base_col + HELPER_OFFSETS["mr"],
                    value=f"=IF(OR(ISNA({xc}),ISNA({xp})),NA(),ABS({xc}-{xp}))")

    # histogram rows
    for j in range(NBINS):
        r = HELPER_FIRST + j
        lo_e = f"({lo_c}$1+{wd_c}$1*{j})"
        hi_e = f"({lo_c}$1+{wd_c}$1*{j + 1})"
        ws.cell(row=r, column=base_col + HELPER_OFFSETS["bin"],
                value=f'=IF({wd_c}$1="","",ROUND({lo_c}$1+{wd_c}$1*{j + 0.5},4))')
        op = "<=" if j == NBINS - 1 else "<"
        ws.cell(row=r, column=base_col + HELPER_OFFSETS["cnt"],
                value=(f'=IF({wd_c}$1="","",COUNTIFS({hist_range},">="&{lo_e},'
                       f'{hist_range},"{op}"&{hi_e}))'))
        for key, ref in (("lslmark", refs["lsl"]), ("uslmark", refs["usl"])):
            ws.cell(row=r, column=base_col + HELPER_OFFSETS[key],
                    value=(f'=IF(OR({ref}="",{wd_c}$1=""),NA(),'
                           f"IF(AND({ref}>={lo_e},{ref}<{hi_e}),"
                           f"{mx_c}$1*1.15,NA()))"))
    return B


def add_x_chart(ws, base_col, anchor, title, width=22, height=8.5):
    B = lambda k: base_col + HELPER_OFFSETS[k]  # noqa: E731
    ch = LineChart()
    ch.title = title
    ch.style = 2
    ch.width, ch.height = width, height
    ch.y_axis.title = "Measurement"
    ch.x_axis.title = "Sample #"
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    cats = Reference(ws, min_col=B("idx"), min_row=HELPER_FIRST, max_row=HELPER_LAST)
    spec = [("x", C_DATA, None, True), ("cl", C_CL, "dash", False),
            ("ucl", C_CTRL, None, False), ("lcl", C_CTRL, None, False),
            ("usl", C_SPEC, "lgDash", False), ("lsl", C_SPEC, "lgDash", False)]
    for key, colour, dash, marker in spec:
        ref = Reference(ws, min_col=B(key), min_row=HELPER_HEAD_ROW, max_row=HELPER_LAST)
        ch.add_data(ref, titles_from_data=True)
        styled_line(ch.series[-1], colour, 22000 if key == "x" else 14000, dash, False, marker)
    ch.set_categories(cats)
    ch.display_blanks = "gap"
    # helper columns are hidden; without this Excel plots nothing
    ch.visible_cells_only = False
    ws.add_chart(ch, anchor)
    return ch


def add_mr_chart(ws, base_col, anchor, title, width=22, height=6.5):
    B = lambda k: base_col + HELPER_OFFSETS[k]  # noqa: E731
    ch = LineChart()
    ch.title = title
    ch.style = 2
    ch.width, ch.height = width, height
    ch.y_axis.title = "Moving range"
    ch.x_axis.title = "Sample #"
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    for key, colour, marker in (("mr", C_DATA, True), ("mrucl", C_CTRL, False)):
        ref = Reference(ws, min_col=B(key), min_row=HELPER_HEAD_ROW, max_row=HELPER_LAST)
        ch.add_data(ref, titles_from_data=True)
        styled_line(ch.series[-1], colour, 20000 if key == "mr" else 14000,
                    None, False, marker)
    ch.set_categories(Reference(ws, min_col=B("idx"), min_row=HELPER_FIRST,
                                max_row=HELPER_LAST))
    ch.display_blanks = "gap"
    # helper columns are hidden; without this Excel plots nothing
    ch.visible_cells_only = False
    ws.add_chart(ch, anchor)
    return ch


def add_hist_chart(ws, base_col, anchor, title, width=13, height=8.5):
    B = lambda k: base_col + HELPER_OFFSETS[k]  # noqa: E731
    last = HELPER_FIRST + NBINS - 1
    bar = BarChart()
    bar.type = "col"
    bar.title = title
    bar.style = 10
    bar.width, bar.height = width, height
    bar.gapWidth = 8
    bar.y_axis.title = "Frequency"
    bar.x_axis.title = "Value"
    bar.x_axis.delete = False
    bar.y_axis.delete = False
    bar.add_data(Reference(ws, min_col=B("cnt"), min_row=HELPER_HEAD_ROW, max_row=last),
                 titles_from_data=True)
    bar.series[0].graphicalProperties.solidFill = C_BAR
    bar.series[0].graphicalProperties.line.solidFill = BLUE
    bar.set_categories(Reference(ws, min_col=B("bin"), min_row=HELPER_FIRST, max_row=last))

    line = LineChart()
    for key, colour in (("lslmark", C_SPEC), ("uslmark", C_SPEC)):
        line.add_data(Reference(ws, min_col=B(key), min_row=HELPER_HEAD_ROW,
                                max_row=last), titles_from_data=True)
        styled_line(line.series[-1], colour, 22000, "lgDash")
    bar += line
    bar.display_blanks = "gap"
    # helper columns are hidden; without this Excel plots nothing
    bar.visible_cells_only = False
    line.visible_cells_only = False
    ws.add_chart(bar, anchor)
    return bar


PANEL_ROWS = [
    ("Feature", "Feature"), ("Unit", None), ("Samples (n)", "n"),
    ("--- DATA RANGE ---", None),
    ("Min", "Min"), ("Max", "Max"), ("Range", "Range"),
    ("Mean", "Mean"), ("Median", "Median"),
    ("Std Dev (overall)", "Sigma (overall)"), ("Sigma (within)", "Sigma (within)"),
    ("--- CONTROL ---", None),
    ("UCL", "X UCL"), ("CL", "X CL"), ("LCL", "X LCL"),
    ("mR-bar", "mR-bar"), ("mR UCL", "mR UCL"),
    ("--- SPECIFICATION ---", None),
    ("USL", "USL"), ("Nominal", "Nominal"), ("LSL", "LSL"),
    ("--- CAPABILITY ---", None),
    ("Cp", "Cp"), ("Cpk", "Cpk"), ("Pp", "Pp"), ("Ppk", "Ppk"),
    ("In control?", "In Control?"), ("In spec?", "In Spec?"), ("Verdict", "Verdict"),
]


def write_stats_panel(ws, top_row, feature_ref, unit_ref, direct=True):
    """Statistics panel in columns A/B. feature_ref -> Capability row (as text)."""
    r = top_row
    for text, capcol in PANEL_ROWS:
        if capcol is None and text.startswith("---"):
            c = label(ws, r, 1, text.replace("---", "").strip(), F_LABEL, SUB_FILL, BOX)
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
            ws.cell(row=r, column=2).fill = SUB_FILL
            ws.cell(row=r, column=2).border = BOX
            c.alignment = Alignment(horizontal="center")
            r += 1
            continue
        label(ws, r, 1, text, F_BODY, None, BOX)
        if capcol is None:                    # Unit
            v = unit_ref
        elif direct:
            v = f"=Capability!{CAPC[capcol]}{feature_ref}"
        else:
            v = (f"=INDEX(Capability!${CAPC[capcol]}$5:${CAPC[capcol]}"
                 f"${cap_row(MAX_FEATURES)},{feature_ref})")
        c = ws.cell(row=r, column=2, value=v)
        c.font = F_BODY
        c.border = BOX
        c.fill = CALC_FILL
        if capcol in ("Cp", "Cpk", "Pp", "Ppk"):
            c.number_format = NUM2
        elif capcol in ("n",):
            c.number_format = NUM0
        elif capcol not in (None, "Feature", "In Control?", "In Spec?", "Verdict"):
            c.number_format = NUM4
        r += 1
    return r


# ==========================================================================
# 4. CONTROL CHARTS
# ==========================================================================
BAND_HEIGHT = 40
CC_HELPER_START = 80        # column CB


def build_control_charts(wb):
    ws = wb.create_sheet("Control Charts")
    title_block(ws, "CONTROL CHARTS - INDIVIDUALS (X), MOVING RANGE (mR) & HISTOGRAM",
                "Live statistics sit to the left of each graph. X charts show UCL/CL/LCL "
                "plus USL/LSL. Charts follow the most recent points "
                "(Settings > Chart Window).", 20)
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 15
    ws.column_dimensions["C"].width = 3

    for f in range(1, CHART_FEATURES + 1):
        top = 4 + (f - 1) * BAND_HEIGHT
        base = CC_HELPER_START + (f - 1) * HELPER_WIDTH
        crow = cap_row(f)
        srow = SET_ROW_FIRST + f - 1
        L = de_letter(f)

        head = label(ws, top, 1,
                     f"=\"FEATURE {f}: \"&IF(Settings!B{srow}=\"\",\"(unnamed)\","
                     f"Settings!B{srow})", F_HEAD, HEAD_FILL, BOX)
        ws.merge_cells(start_row=top, start_column=1, end_row=top, end_column=2)
        ws.cell(row=top, column=2).fill = HEAD_FILL
        head.alignment = Alignment(horizontal="left")

        write_stats_panel(ws, top + 1, crow, f"=Settings!F{srow}", direct=True)

        refs = {
            "lastrow": f"'Data Entry'!{L}{ST['lastrow']}",
            "min": f"Capability!${CAPC['Min']}${crow}",
            "max": f"Capability!${CAPC['Max']}${crow}",
            "cl": f"Capability!${CAPC['X CL']}${crow}",
            "ucl": f"Capability!${CAPC['X UCL']}${crow}",
            "lcl": f"Capability!${CAPC['X LCL']}${crow}",
            "mrucl": f"Capability!${CAPC['mR UCL']}${crow}",
            "usl": f"Capability!${CAPC['USL']}${crow}",
            "lsl": f"Capability!${CAPC['LSL']}${crow}",
            "x": lambda pos, L=L: (f"INDEX('Data Entry'!${L}${DATA_FIRST}:${L}"
                                   f"${DATA_LAST},{pos}-{DATA_FIRST - 1})"),
        }
        write_helper_block(
            ws, base, refs, "Settings!$F$5",
            f"'Data Entry'!${L}${DATA_FIRST}:${L}${DATA_LAST}", "Settings!$F$6")

        name = f'Settings!B{srow}'
        add_x_chart(ws, base, f"D{top}", f"Feature {f} - Individuals (X)")
        add_mr_chart(ws, base, f"D{top + 18}", f"Feature {f} - Moving Range (mR)")
        add_hist_chart(ws, base, f"S{top}", f"Feature {f} - Histogram")
        _ = name

    hide_from = CL(CC_HELPER_START)
    hide_to = CL(CC_HELPER_START + CHART_FEATURES * HELPER_WIDTH)
    ws.column_dimensions.group(hide_from, hide_to, hidden=True)
    return ws


# ==========================================================================
# 5. MONITOR
# ==========================================================================
MON_SLOTS = 4
MON_HELPER_START = 80
MON_BAND = 32
MON_FIRST_BAND = 34


def build_monitor(wb):
    ws = wb.create_sheet("Monitor")
    title_block(ws, f"SPC MONITOR {VERSION} - LIVE PROCESS CONTROL",
                "Folder watch + auto-refresh (requires SPC_V13_Monitor.bas). "
                "Statistics panel left, charts right.", 20)
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 3

    # ---- config block ---------------------------------------------------
    label(ws, 4, 1, "MONITOR CONFIGURATION", F_HEAD, HEAD_FILL, BOX)
    ws.merge_cells("A4:B4")
    ws.cell(row=4, column=2).fill = HEAD_FILL

    cfg = [
        ("Watch Folder", "=Settings!F8", False),
        ("File Pattern", "=Settings!F9", False),
        ("Refresh Interval (s)", "=Settings!F10", False),
        ("Archive After Import", "=Settings!F11", False),
        ("Status", "IDLE", True),
        ("Last Scan", "", True),
        ("Files Imported", 0, True),
        ("Rows Imported", 0, True),
        ("Last File", "", True),
        ("Last Message", "", True),
    ]
    for i, (name, val, is_out) in enumerate(cfg):
        r = 5 + i
        label(ws, r, 1, name, F_LABEL, None, BOX)
        c = ws.cell(row=r, column=2, value=val)
        c.font = F_BODY
        c.border = BOX
        c.fill = CALC_FILL if is_out else SUB_FILL
    ws.cell(row=17, column=1,
            value="Edit paths on Settings (F8:F11). Run SPC_StartMonitor to begin.").font = F_SMALL

    label(ws, 19, 1, "CHART SLOTS", F_HEAD, HEAD_FILL, BOX)
    ws.merge_cells("A19:B19")
    ws.cell(row=19, column=2).fill = HEAD_FILL
    for i in range(MON_SLOTS):
        r = 20 + i
        label(ws, r, 1, f"Slot {i + 1} - Feature #", F_LABEL, None, BOX)
        c = ws.cell(row=r, column=2, value=i + 1)
        c.font = F_INPUT
        c.fill = INPUT_FILL
        c.border = BOX
    dv = DataValidation(type="whole", operator="between", formula1=1,
                        formula2=MAX_FEATURES, allow_blank=False)
    ws.add_data_validation(dv)
    dv.add(f"B20:B{20 + MON_SLOTS - 1}")

    # ---- alarm table ----------------------------------------------------
    ar = 26
    label(ws, ar, 1, "PROCESS STATUS - ALL ACTIVE FEATURES", F_SECTION)
    heads = ["Feature", "n", "Last value", "Mean", "Range", "Cpk",
             "In control?", "In spec?", "Verdict"]
    for i, h in enumerate(heads):
        label(ws, ar + 1, 1 + i, h, F_HEAD, HEAD_FILL, BOX,
              Alignment(horizontal="center", wrap_text=True))
    for f in range(1, CHART_FEATURES + 1):
        r = ar + 1 + f
        crow = cap_row(f)
        L = de_letter(f)
        vals = [
            f"=Capability!{CAPC['Feature']}{crow}",
            f"=Capability!{CAPC['n']}{crow}",
            (f"=IF('Data Entry'!{L}{ST['lastrow']}<{DATA_FIRST},\"\","
             f"INDEX('Data Entry'!${L}${DATA_FIRST}:${L}${DATA_LAST},"
             f"'Data Entry'!{L}{ST['lastrow']}-{DATA_FIRST - 1}))"),
            f"=Capability!{CAPC['Mean']}{crow}",
            f"=Capability!{CAPC['Range']}{crow}",
            f"=Capability!{CAPC['Cpk']}{crow}",
            f"=Capability!{CAPC['In Control?']}{crow}",
            f"=Capability!{CAPC['In Spec?']}{crow}",
            f"=Capability!{CAPC['Verdict']}{crow}",
        ]
        for i, v in enumerate(vals):
            c = ws.cell(row=r, column=1 + i, value=v)
            c.font = F_BODY
            c.border = BOX
            c.fill = CALC_FILL
            if i in (2, 3, 4):
                c.number_format = NUM4
            elif i == 5:
                c.number_format = NUM2
            elif i == 1:
                c.number_format = NUM0
    last_alarm = ar + 1 + CHART_FEATURES
    for colletter, bad, good in (("G", "OUT", "IN"), ("H", "FAIL", "PASS")):
        rng = f"{colletter}{ar + 2}:{colletter}{last_alarm}"
        ws.conditional_formatting.add(rng, CellIsRule(
            operator="equal", formula=[f'"{bad}"'],
            fill=PatternFill("solid", fgColor="FFC7CE"),
            font=Font(name=FONT, color="9C0006", bold=True)))
        ws.conditional_formatting.add(rng, CellIsRule(
            operator="equal", formula=[f'"{good}"'],
            fill=PatternFill("solid", fgColor="C6EFCE"),
            font=Font(name=FONT, color="006100")))

    # ---- slot bands -----------------------------------------------------
    capmax = cap_row(MAX_FEATURES)
    for slot in range(MON_SLOTS):
        top = MON_FIRST_BAND + slot * MON_BAND + 8
        base = MON_HELPER_START + slot * HELPER_WIDTH
        sel = f"$B${20 + slot}"

        head = label(ws, top, 1,
                     f'="SLOT {slot + 1}: "&IF(INDEX(Settings!$B$17:$B$66,{sel})="",'
                     f'"(unnamed)",INDEX(Settings!$B$17:$B$66,{sel}))',
                     F_HEAD, HEAD_FILL, BOX)
        ws.merge_cells(start_row=top, start_column=1, end_row=top, end_column=2)
        ws.cell(row=top, column=2).fill = HEAD_FILL
        head.alignment = Alignment(horizontal="left")

        write_stats_panel(ws, top + 1, sel,
                          f"=INDEX(Settings!$F$17:$F$66,{sel})", direct=False)

        idx = lambda col: f"INDEX(Capability!${col}$5:${col}${capmax},{sel})"  # noqa: E731
        refs = {
            "lastrow": (f"INDEX('Data Entry'!$H${ST['lastrow']}:"
                        f"${CL(FEAT_COL_LAST)}${ST['lastrow']},1,{sel})"),
            "min": idx(CAPC["Min"]), "max": idx(CAPC["Max"]),
            "cl": idx(CAPC["X CL"]), "ucl": idx(CAPC["X UCL"]),
            "lcl": idx(CAPC["X LCL"]), "mrucl": idx(CAPC["mR UCL"]),
            "usl": idx(CAPC["USL"]), "lsl": idx(CAPC["LSL"]),
            "x": lambda pos, sel=sel: (
                f"INDEX('Data Entry'!$H${DATA_FIRST}:${CL(FEAT_COL_LAST)}"
                f"${DATA_LAST},{pos}-{DATA_FIRST - 1},{sel})"),
        }
        # INDEX(range,0,n) yields the whole n-th column as a reference and,
        # unlike OFFSET, is not volatile - the histogram COUNTIFS then only
        # recalculate when their inputs actually change.
        hist = (f"INDEX('Data Entry'!$H${DATA_FIRST}:${CL(FEAT_COL_LAST)}"
                f"${DATA_LAST},0,{sel})")
        write_helper_block(ws, base, refs, "Settings!$F$5", hist, "Settings!$F$6")

        add_x_chart(ws, base, f"D{top}", f"Slot {slot + 1} - Individuals (X)",
                    width=20, height=8)
        add_hist_chart(ws, base, f"R{top}", f"Slot {slot + 1} - Histogram",
                       width=12, height=8)
        add_mr_chart(ws, base, f"D{top + 17}", f"Slot {slot + 1} - Moving Range",
                     width=20, height=5.5)

    hide_from = CL(MON_HELPER_START)
    hide_to = CL(MON_HELPER_START + MON_SLOTS * HELPER_WIDTH)
    ws.column_dimensions.group(hide_from, hide_to, hidden=True)
    return ws


# ==========================================================================
# 6. _Calc  (filter engine for Visual SPC)
# ==========================================================================
def build_calc(wb):
    ws = wb.create_sheet("_Calc")
    ws["A1"] = "Filter engine for Visual SPC. Do not edit."
    ws["A1"].font = F_SUB
    for i, h in enumerate(["Src row", "Match", "Rank", "Value", "Prev value", "mR"]):
        ws.cell(row=4, column=1 + i, value=h).font = F_SMALL
    ws["C4"] = 0
    ws["E4"] = ""

    VS = "'Visual SPC'!"
    for r in range(DATA_FIRST, DATA_LAST + 1):
        ws.cell(row=r, column=1, value=r)
        val = (f"INDEX('Data Entry'!$H${DATA_FIRST}:${CL(FEAT_COL_LAST)}"
               f"${DATA_LAST},{r}-{DATA_FIRST - 1},{VS}$B$3)")
        ws.cell(row=r, column=2, value=(
            f'=IF({val}="",0,IF(AND('
            f'OR({VS}$C$4="",\'Data Entry\'!$C{r}>={VS}$C$4),'
            f'OR({VS}$E$4="",\'Data Entry\'!$C{r}<={VS}$E$4),'
            f'OR({VS}$G$4="",\'Data Entry\'!$D{r}={VS}$G$4),'
            f'OR({VS}$I$4="",\'Data Entry\'!$E{r}={VS}$I$4),'
            f'OR({VS}$K$4="",\'Data Entry\'!$F{r}={VS}$K$4)),1,0))'))
        ws.cell(row=r, column=3, value=f"=C{r - 1}+B{r}")
        ws.cell(row=r, column=4, value=f'=IF(B{r}=1,{val},"")')
        ws.cell(row=r, column=5, value=f'=IF(B{r}=1,D{r},E{r - 1})')
        ws.cell(row=r, column=6, value=(
            f'=IF(AND(B{r}=1,C{r}>1,E{r - 1}<>""),ABS(D{r}-E{r - 1}),"")'))

    ws.sheet_state = "hidden"
    return ws


# ==========================================================================
# 7. VISUAL SPC (filters + plots)
# ==========================================================================
VS_HELPER_START = 60
VS_STAT_ROW = 8


def build_visual_spc(wb):
    ws = wb.create_sheet("Visual SPC")
    title_block(ws, "VISUAL SPC - FILTERED ANALYSIS",
                "Pick a feature and filters; every plot and statistic below follows "
                "the filtered subset. Leave a filter blank to include everything.", 20)
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 15
    ws.column_dimensions["C"].width = 3

    label(ws, 3, 1, "Feature #:", F_LABEL)
    c = ws.cell(row=3, column=2, value=1)
    c.font, c.fill, c.border = F_INPUT, INPUT_FILL, BOX
    ws.cell(row=3, column=3,
            value=f'=IF(AND(B3>=1,B3<={MAX_FEATURES}),'
                  f'INDEX(Settings!$B$17:$B$66,B3),"?")').font = F_SECTION

    filters = [("Date from:", "C"), ("Date to:", "E"), ("Shift:", "G"),
               ("Operator:", "I"), ("Tool:", "K")]
    label(ws, 4, 1, "Filters:", F_LABEL)
    for text, colletter in filters:
        prev = CL(ws[f"{colletter}4"].column - 1)
        label(ws, 4, ws[f"{prev}4"].column, text, F_SMALL)
        cell = ws[f"{colletter}4"]
        cell.font, cell.fill, cell.border = F_INPUT, INPUT_FILL, BOX
    label(ws, 4, 12, "(blank = all;  dates as YYYY-MM-DD)", F_SMALL)

    base = VS_HELPER_START
    capmax = cap_row(MAX_FEATURES)
    sel = "$B$3"
    rank_col = "_Calc!$C$5:$C$" + str(DATA_LAST)
    val_col = "_Calc!$D$5:$D$" + str(DATA_LAST)
    mr_col = "_Calc!$F$5:$F$" + str(DATA_LAST)

    # ---- filtered statistics (columns A/B) ------------------------------
    stats = [
        ("Feature", f'=IF(AND(B3>=1,B3<={MAX_FEATURES}),INDEX(Settings!$B$17:$B$66,B3),"")'),
        ("Unit", f"=INDEX(Settings!$F$17:$F$66,{sel})"),
        ("--- FILTERED DATA RANGE ---", None),
        ("Samples (n)", f"=COUNT({val_col})"),
        ("Min", f'=IF(COUNT({val_col})=0,"",MIN({val_col}))'),
        ("Max", f'=IF(COUNT({val_col})=0,"",MAX({val_col}))'),
        ("Range", f'=IF(COUNT({val_col})=0,"",MAX({val_col})-MIN({val_col}))'),
        ("Mean", f'=IF(COUNT({val_col})=0,"",AVERAGE({val_col}))'),
        ("Median", f'=IF(COUNT({val_col})=0,"",MEDIAN({val_col}))'),
        ("Std Dev (overall)", f'=IF(COUNT({val_col})<2,"",STDEV({val_col}))'),
        ("mR-bar", f'=IF(COUNT({mr_col})=0,"",AVERAGE({mr_col}))'),
        ("Sigma (within)", None),
        ("--- CONTROL (filtered) ---", None),
        ("UCL", None), ("CL", None), ("LCL", None), ("mR UCL", None),
        ("--- SPECIFICATION ---", None),
        ("USL", f"=INDEX(Capability!${CAPC['USL']}$5:${CAPC['USL']}${capmax},{sel})"),
        ("Nominal", f"=INDEX(Capability!${CAPC['Nominal']}$5:${CAPC['Nominal']}${capmax},{sel})"),
        ("LSL", f"=INDEX(Capability!${CAPC['LSL']}$5:${CAPC['LSL']}${capmax},{sel})"),
        ("--- CAPABILITY (filtered) ---", None),
        ("Cp", None), ("Cpk", None), ("Pp", None), ("Ppk", None),
    ]
    row_of = {}
    r = VS_STAT_ROW
    for name, formula in stats:
        row_of[name] = r
        r += 1
    R = row_of
    mean, sw, so = f"$B${R['Mean']}", f"$B${R['Sigma (within)']}", f"$B${R['Std Dev (overall)']}"
    mrb = f"$B${R['mR-bar']}"
    usl, lsl = f"$B${R['USL']}", f"$B${R['LSL']}"
    nospec = f'OR({usl}="",{lsl}="")'
    dyn = {
        "Sigma (within)": f'=IF({mrb}="","",{mrb}/1.128)',
        "UCL": f'=IF({mrb}="","",{mean}+2.66*{mrb})',
        "CL": f'=IF({mean}="","",{mean})',
        "LCL": f'=IF({mrb}="","",{mean}-2.66*{mrb})',
        "mR UCL": f'=IF({mrb}="","",3.267*{mrb})',
        "Cp": f'=IF(OR({nospec},{sw}="",{sw}=0),"",({usl}-{lsl})/(6*{sw}))',
        "Cpk": f'=IF(OR({nospec},{sw}="",{sw}=0),"",MIN({usl}-{mean},{mean}-{lsl})/(3*{sw}))',
        "Pp": f'=IF(OR({nospec},{so}="",{so}=0),"",({usl}-{lsl})/(6*{so}))',
        "Ppk": f'=IF(OR({nospec},{so}="",{so}=0),"",MIN({usl}-{mean},{mean}-{lsl})/(3*{so}))',
    }
    for name, formula in stats:
        rr = R[name]
        if formula is None and name.startswith("---"):
            cc = label(ws, rr, 1, name.replace("---", "").strip(), F_LABEL, SUB_FILL, BOX)
            ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=2)
            ws.cell(row=rr, column=2).fill = SUB_FILL
            ws.cell(row=rr, column=2).border = BOX
            cc.alignment = Alignment(horizontal="center")
            continue
        label(ws, rr, 1, name, F_BODY, None, BOX)
        cc = ws.cell(row=rr, column=2, value=formula if formula else dyn[name])
        cc.font, cc.border, cc.fill = F_BODY, BOX, CALC_FILL
        if name in ("Cp", "Cpk", "Pp", "Ppk"):
            cc.number_format = NUM2
        elif name == "Samples (n)":
            cc.number_format = NUM0
        elif name not in ("Feature", "Unit"):
            cc.number_format = NUM4

    # ---- helper block (filtered series) ---------------------------------
    B = lambda k: CL(base + HELPER_OFFSETS[k])  # noqa: E731
    tot_c, st_c = CL(base + 1), CL(base + 3)
    lo_c, wd_c, mx_c, hi_c = (CL(base + 5), CL(base + 7), CL(base + 9), CL(base + 11))
    srow_c = CL(base + 13)     # spare column holds the matched source row

    vs_lo, vs_hi = bin_bounds(f'$B${R["Min"]}', f'$B${R["Max"]}', lsl, usl)
    scal = [
        ("Total", f"=MAX({rank_col})"),
        ("Start", f"=MAX(1,{tot_c}1-Settings!$F$5+1)"),
        ("Lo", vs_lo),
        ("Width", f'=IF(OR({lo_c}1="",{hi_c}1="",{hi_c}1<={lo_c}1),"",'
                  f"({hi_c}1-{lo_c}1)/Settings!$F$6)"),
        ("MaxCnt", f"=MAX({B('cnt')}{HELPER_FIRST}:{B('cnt')}{HELPER_FIRST + NBINS - 1})"),
        ("Hi", vs_hi),
    ]
    for i, (name, formula) in enumerate(scal):
        ws.cell(row=1, column=base + i * 2, value=name).font = F_SMALL
        cc = ws.cell(row=1, column=base + i * 2 + 1, value=formula)
        cc.font, cc.number_format = F_SMALL, NUM4

    heads = ["idx", "X", "CL", "UCL", "LCL", "mR", "mR_UCL", "USL", "LSL",
             "Bin", "Count", "LSLmark", "USLmark", "SrcRow"]
    for i, h in enumerate(heads):
        ws.cell(row=HELPER_HEAD_ROW, column=base + i, value=h).font = F_SMALL

    for k in range(WINDOW):
        rr = HELPER_FIRST + k
        target = f"({st_c}$1+{k})"
        ws.cell(row=rr, column=base + 13, value=(
            f"=IF({target}>{tot_c}$1,NA(),MATCH({target},{rank_col},0))"))
        na = f"IF(ISNA({srow_c}{rr}),NA()"
        ws.cell(row=rr, column=base + HELPER_OFFSETS["idx"],
                value=f"={na},{target})")
        ws.cell(row=rr, column=base + HELPER_OFFSETS["x"],
                value=f"={na},INDEX({val_col},{srow_c}{rr}))")
        ws.cell(row=rr, column=base + HELPER_OFFSETS["mr"], value=(
            f'={na},IF(INDEX({mr_col},{srow_c}{rr})="",NA(),'
            f"INDEX({mr_col},{srow_c}{rr})))"))
        for key, ref in (("cl", f"$B${R['CL']}"), ("ucl", f"$B${R['UCL']}"),
                         ("lcl", f"$B${R['LCL']}"), ("mrucl", f"$B${R['mR UCL']}"),
                         ("usl", usl), ("lsl", lsl)):
            ws.cell(row=rr, column=base + HELPER_OFFSETS[key],
                    value=f'={na},IF({ref}="",NA(),{ref}))')

    for j in range(NBINS):
        rr = HELPER_FIRST + j
        lo_e = f"({lo_c}$1+{wd_c}$1*{j})"
        hi_e = f"({lo_c}$1+{wd_c}$1*{j + 1})"
        ws.cell(row=rr, column=base + HELPER_OFFSETS["bin"],
                value=f'=IF({wd_c}$1="","",ROUND({lo_c}$1+{wd_c}$1*{j + 0.5},4))')
        op = "<=" if j == NBINS - 1 else "<"
        ws.cell(row=rr, column=base + HELPER_OFFSETS["cnt"],
                value=(f'=IF({wd_c}$1="","",COUNTIFS({val_col},">="&{lo_e},'
                       f'{val_col},"{op}"&{hi_e}))'))
        for key, ref in (("lslmark", lsl), ("uslmark", usl)):
            ws.cell(row=rr, column=base + HELPER_OFFSETS[key],
                    value=(f'=IF(OR({ref}="",{wd_c}$1=""),NA(),'
                           f"IF(AND({ref}>={lo_e},{ref}<{hi_e}),"
                           f"{mx_c}$1*1.15,NA()))"))

    add_x_chart(ws, base, "D3", "Filtered Individuals (X)", width=21, height=8.5)
    add_hist_chart(ws, base, "S3", "Filtered Histogram", width=13, height=8.5)
    add_mr_chart(ws, base, "D21", "Filtered Moving Range (mR)", width=21, height=6.5)

    ws.column_dimensions.group(CL(base), CL(base + HELPER_WIDTH), hidden=True)
    return ws


# ==========================================================================
# 8. DATA COLLECTION FORM (printable)
# ==========================================================================
def build_form(wb):
    ws = wb.create_sheet("Data Collection Form")
    ws.sheet_view.showGridLines = False

    ncols = 4 + CHART_FEATURES        # Sample, Time, Part/Serial, ... features, Notes
    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 9
    ws.column_dimensions["C"].width = 12
    for f in range(CHART_FEATURES):
        ws.column_dimensions[CL(4 + f)].width = 11
    ws.column_dimensions[CL(4 + CHART_FEATURES)].width = 22

    # header
    ws["A1"] = "SPC DATA COLLECTION SHEET"
    ws["A1"].font = Font(name=FONT, size=15, bold=True)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols - 3)
    ws.cell(row=1, column=ncols - 2, value=f'="Doc: "&Settings!B9&"  Rev "&Settings!B10')
    ws.cell(row=1, column=ncols - 2).font = F_SMALL
    ws.merge_cells(start_row=1, start_column=ncols - 2, end_row=1, end_column=ncols)
    ws["A2"] = "=Settings!B8"
    ws["A2"].font = F_SUB
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)

    fields = [("Part No.", "Operation"), ("Machine / Cell", "Gauge ID"),
              ("Operator", "Shift"), ("Date", "Sample frequency")]
    for i, (l1, l2) in enumerate(fields):
        r = 4 + i
        label(ws, r, 1, l1 + ":", F_LABEL)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        for c in range(3, 7):
            ws.cell(row=r, column=c).border = Border(bottom=THIN)
        label(ws, r, 8, l2 + ":", F_LABEL)
        for c in range(10, min(ncols, 14) + 1):
            ws.cell(row=r, column=c).border = Border(bottom=THIN)

    # characteristics reference
    cr = 9
    label(ws, cr, 1, "CHARACTERISTICS", F_HEAD, HEAD_FILL, BOX)
    ws.merge_cells(start_row=cr, start_column=1, end_row=cr, end_column=ncols)
    for c in range(1, ncols + 1):
        ws.cell(row=cr, column=c).fill = HEAD_FILL

    spec_rows = [("Characteristic", "B"), ("Nominal", "C"), ("LSL", "D"), ("USL", "E"),
                 ("Unit", "F")]
    for i, (name, setcol) in enumerate(spec_rows):
        r = cr + 1 + i
        label(ws, r, 1, name, F_LABEL, SUB_FILL, BOX)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        for c in (2, 3):
            ws.cell(row=r, column=c).fill = SUB_FILL
            ws.cell(row=r, column=c).border = BOX
        for f in range(1, CHART_FEATURES + 1):
            sr = SET_ROW_FIRST + f - 1
            cc = ws.cell(row=r, column=3 + f,
                         value=f'=IF(Settings!{setcol}{sr}="","-",Settings!{setcol}{sr})')
            cc.font = F_SMALL
            cc.border = BOX
            cc.alignment = Alignment(horizontal="center", wrap_text=True)
            if setcol in "CDE":
                cc.number_format = NUM4
        ws.cell(row=r, column=3 + CHART_FEATURES + 1).border = BOX

    # reading grid
    gr = cr + 1 + len(spec_rows) + 1
    heads = ["Sample #", "Time", "Serial / Batch"] + \
            [f"Char {i}" for i in range(1, CHART_FEATURES + 1)] + ["Notes"]
    for i, h in enumerate(heads):
        label(ws, gr, 1 + i, h, F_HEAD, HEAD_FILL, BOX,
              Alignment(horizontal="center", wrap_text=True))
    ws.row_dimensions[gr].height = 24

    for k in range(FORM_ROWS):
        r = gr + 1 + k
        ws.row_dimensions[r].height = 20
        cc = ws.cell(row=r, column=1, value=k + 1)
        cc.font, cc.border = F_BODY, BOX
        cc.alignment = Alignment(horizontal="center")
        for c in range(2, ncols + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = BOX
            if k % 5 == 4:
                cell.border = Border(left=THIN, right=THIN, top=THIN, bottom=MED)

    # footer
    fr = gr + FORM_ROWS + 2
    label(ws, fr, 1, "Notes / out-of-spec actions:", F_LABEL)
    for r in range(fr + 1, fr + 4):
        for c in range(1, ncols + 1):
            ws.cell(row=r, column=c).border = Border(bottom=THIN)
    sr = fr + 5
    for i, txt in enumerate(["Recorded by:", "Signature:", "Checked by:", "Date:"]):
        label(ws, sr, 1 + i * 3, txt, F_LABEL)
        for c in range(2 + i * 3, 4 + i * 3):
            ws.cell(row=sr, column=c).border = Border(bottom=THIN)
    label(ws, sr + 2, 1,
          "Transfer readings into the Data Entry sheet (Char 1 -> column H, "
          "Char 2 -> column I, ...). Keep this sheet with the job pack.", F_SMALL)

    # print setup
    ws.print_area = f"A1:{CL(ncols)}{sr + 2}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.print_title_rows = f"1:{gr}"
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.5
    ws.oddFooter.left.text = "&F - &A"
    ws.oddFooter.right.text = "Page &P of &N"
    return ws


# ==========================================================================
# 9. DASHBOARD / DOC CONTROL / HELP
# ==========================================================================
def build_dashboard(wb):
    ws = wb.create_sheet("Dashboard")
    title_block(ws, "SPC DASHBOARD", "Single-feature summary. Change the feature "
                "number in B4.", 6)
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 20
    label(ws, 4, 1, "Select Feature #:", F_LABEL)
    c = ws.cell(row=4, column=2, value=1)
    c.font, c.fill, c.border = F_INPUT, INPUT_FILL, BOX

    capmax = cap_row(MAX_FEATURES)
    r = 6
    for header, _ in CAP_COLS:
        label(ws, r, 1, header, F_BODY, None, BOX)
        cc = ws.cell(row=r, column=2, value=(
            f"=INDEX(Capability!${CAPC[header]}$5:${CAPC[header]}${capmax},$B$4)"))
        cc.font, cc.border, cc.fill = F_BODY, BOX, CALC_FILL
        if header in ("Cp", "Cpk", "Pp", "Ppk", "Tol Width"):
            cc.number_format = NUM2
        elif header == "n":
            cc.number_format = NUM0
        elif header not in ("Feature", "In Control?", "In Spec?", "Verdict"):
            cc.number_format = NUM4
        r += 1
    return ws


def build_doc_sheets(wb):
    ws = wb.create_sheet("Document Control")
    title_block(ws, "DOCUMENT CONTROL", "", 4)
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 42
    rows = [("Document ID", DOC_ID), ("Title", f"SPC Calculator {VERSION}"),
            ("Revision", REVISION), ("Status", "DRAFT"), ("Owner", OWNER),
            ("Company", COMPANY), ("Created", ISSUE_DATE),
            ("Classification", "Internal")]
    label(ws, 4, 1, "Field", F_HEAD, HEAD_FILL, BOX)
    label(ws, 4, 2, "Value", F_HEAD, HEAD_FILL, BOX)
    for i, (k, v) in enumerate(rows):
        label(ws, 5 + i, 1, k, F_LABEL, None, BOX)
        label(ws, 5 + i, 2, v, F_BODY, None, BOX)

    ws2 = wb.create_sheet("Revision History")
    title_block(ws2, "REVISION HISTORY", "", 4)
    for i, w in enumerate((8, 14, 10, 90)):
        ws2.column_dimensions[CL(1 + i)].width = w
    for i, h in enumerate(["Rev", "Date", "Author", "Description"]):
        label(ws2, 4, 1 + i, h, F_HEAD, HEAD_FILL, BOX)
    hist = [
        ("R01", "2026-07-30", "AM", "v11: Monitor sheet, VBA auto-import, chart toggles"),
        ("R02", "2026-07-30", "AM", "v11: NA() formulas for auto-shrink charts"),
        ("R03", "2026-07-30", "AM", "v12: Visual SPC dashboard"),
        ("R01", ISSUE_DATE, "AM",
         "v13 all-in-one: stats/data-range panel left of every chart; USL & LSL on X "
         "charts; histogram per feature with spec spikes; rolling chart window; "
         "printable Data Collection Form; filtered Visual SPC plots; folder-watch "
         "monitor config + all-feature alarm table; VBA module SPC_V13_Monitor.bas"),
    ]
    for i, row in enumerate(hist):
        for j, v in enumerate(row):
            label(ws2, 5 + i, 1 + j, v, F_BODY, None, BOX,
                  Alignment(wrap_text=True, vertical="top"))

    ws4 = wb.create_sheet("File Locations")
    title_block(ws4, "FILE LOCATIONS", "", 3)
    ws4.column_dimensions["A"].width = 24
    ws4.column_dimensions["B"].width = 70
    label(ws4, 4, 1, "Item", F_HEAD, HEAD_FILL, BOX)
    label(ws4, 4, 2, "Path", F_HEAD, HEAD_FILL, BOX)
    items = [("Build script", "spc-tool/build_spc_v13.py"),
             ("VBA module", "spc-tool/SPC_V13_Monitor.bas"),
             ("VBA module (txt)", "spc-tool/SPC_V13_Monitor_CODE.txt"),
             ("Validator", "spc-tool/validate_v13.py"),
             ("Read me", "spc-tool/README.md"),
             ("Watch folder", "=Settings!F8"),
             ("Archive folder", "=Settings!F12")]
    for i, (k, v) in enumerate(items):
        label(ws4, 5 + i, 1, k, F_LABEL, None, BOX)
        label(ws4, 5 + i, 2, v, F_BODY, None, BOX)


def build_how_to_use(wb):
    """First sheet: step-by-step instructions + the area the macro buttons use."""
    ws = wb.create_sheet("How To Use")
    title_block(ws, f"HOW TO USE - SPC CALCULATOR {VERSION}",
                "Read this once. The charts need no macros; only folder "
                "monitoring does.", 6)
    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 78
    ws.column_dimensions["D"].width = 3
    for c in "EFG":
        ws.column_dimensions[c].width = 22

    r = [4]

    def section(text):
        label(ws, r[0], 1, text, F_HEAD, HEAD_FILL, BOX)
        ws.merge_cells(start_row=r[0], start_column=1, end_row=r[0], end_column=3)
        for c in (2, 3):
            ws.cell(row=r[0], column=c).fill = HEAD_FILL
            ws.cell(row=r[0], column=c).border = BOX
        r[0] += 1

    def step(num, head, body):
        if num is not None:
            c = ws.cell(row=r[0], column=1, value=num)
            c.font = Font(name=FONT, size=11, bold=True, color=BLUE)
            c.alignment = Alignment(horizontal="center", vertical="top")
        label(ws, r[0], 2, head, F_LABEL, None, None,
              Alignment(vertical="top", wrap_text=True))
        label(ws, r[0], 3, body, F_BODY, None, None,
              Alignment(vertical="top", wrap_text=True))
        r[0] += 1

    def gap():
        r[0] += 1

    section("QUICK START - NO MACROS NEEDED")
    step(1, "Set up your features",
         "Settings sheet. For each feature type the name, Nominal, LSL, USL and "
         "unit, and set Active = Y. Yellow cells are the ones you edit.")
    step(2, "Enter measurements",
         "Data Entry sheet. Column H is feature 1, I is feature 2, and so on. "
         "Date / Shift / Operator / Tool ID in columns C-F are optional but they "
         "are what the Visual SPC filters use.")
    step(3, "Look at the charts",
         "Control Charts updates by itself - there is nothing to run. Each "
         "feature shows a statistics panel on the left, then the X chart, the "
         "moving-range chart and a histogram.")
    step(4, "Check capability",
         "Capability sheet gives Cp, Cpk, Pp, Ppk with traffic lights, plus an "
         "in-control and an in-spec verdict per feature.")
    step(5, "Filter and explore",
         "Visual SPC. Pick a feature in B3, then filter by date range, shift, "
         "operator or tool. Every statistic and all three plots recompute from "
         "the filtered rows only. Blank filter = include everything.")
    step(6, "Recording readings by hand",
         "Data Collection Form is already set up to print: A4 landscape, one "
         "page, spec limits pulled from Settings. Ctrl+P.")
    gap()

    section("IF THE CHARTS LOOK EMPTY")
    step(None, "Press Ctrl+Alt+F9",
         "That forces a full recalculation. The workbook is built to recalculate "
         "when it opens, but this settles it if a chart looks stale.")
    step(None, "Check Active = Y",
         "A feature with Active = N on Settings has no name in Data Entry row 4 "
         "and is excluded from Capability.")
    step(None, "Check the readings are numbers",
         "Values pasted as text are ignored by every statistic. Excel shows them "
         "left-aligned with a green corner marker.")
    step(None, "#N/A in the hidden columns is correct",
         "That is how the charts shrink and grow with the data - Excel skips "
         "#N/A points. Do not 'fix' those cells.")
    gap()

    section("CHART COLOUR KEY")
    step(None, "Blue line + markers", "Individual measurements")
    step(None, "Green dashed", "CL - the process mean")
    step(None, "Red", "UCL / LCL - 3-sigma control limits calculated from mR-bar")
    step(None, "Purple long-dash", "USL / LSL - your specification limits")
    step(None, "Histogram bars",
         "Distribution of all readings for that feature. The purple spikes mark "
         "the bins holding LSL and USL, so you can see the distribution against "
         "the tolerance rather than against itself.")
    gap()

    section("WHAT EACH SHEET IS FOR")
    for name, purpose in [
        ("Settings", "Feature definitions and display / monitor options"),
        ("Data Entry", "All measurements. Column H onward"),
        ("Data Collection Form", "Printable sheet for recording readings off-PC"),
        ("Control Charts", "Per feature: stats panel, X chart, mR chart, histogram"),
        ("Monitor", "Live screen: 4 selectable slots + all-feature status table"),
        ("Visual SPC", "Filtered analysis and plots"),
        ("Capability", "Cp / Cpk / Pp / Ppk and verdicts"),
        ("Dashboard", "One-feature summary"),
        ("Document Control", "Doc ID, revision, owner"),
    ]:
        step(None, name, purpose)
    gap()

    section("DISPLAY OPTIONS (Settings, column F)")
    for cell, meaning in [
        ("F5  Chart Window", "How many of the most recent points the charts plot "
                             "(default 200). Charts always follow the newest data."),
        ("F6  Histogram Bins", "Number of histogram bars (default 15)."),
        ("F7  Show Spec Limits", "Y shows the USL/LSL lines, N hides them without "
                                 "deleting the limits."),
    ]:
        step(None, cell, meaning)
    gap()

    section("FOLDER MONITORING - THE ONLY PART THAT NEEDS MACROS")
    step(1, "Save as .xlsm",
         "File > Save As > Excel Macro-Enabled Workbook. Macros cannot be stored "
         "in a .xlsx.")
    step(2, "Import the module",
         "Alt+F11 > File > Import File > SPC_V13_Monitor.bas. (If you were sent "
         "the .txt version, rename it to .bas first, or paste its contents into "
         "a new module and drop the first Attribute line.)")
    step(3, "Buttons appear",
         "Save, close and reopen the workbook. The buttons on the right of this "
         "sheet and on the Monitor sheet are created automatically. If they are "
         "missing, run SPC_BuildButtons once from Alt+F8.")
    step(4, "Point it at a folder",
         "Settings F8 = folder to watch, F9 = pattern such as *.csv, "
         "F10 = seconds between scans, F11 = Y to move imported files into the "
         "archive subfolder named in F12.")
    step(5, "Press Start Monitoring",
         "New files are appended to Data Entry, all charts refresh, and the "
         "Monitor sheet logs status, last scan, and files / rows imported. Each "
         "file is imported once - a hidden log prevents double counting.")
    step(None, "Accepted file layouts",
         "Date, Shift, Operator, ToolID, Cluster, F1, F2, ... Fn   -or-   "
         "F1, F2, ... Fn on its own. A header row is skipped automatically.")
    gap()

    section("BUTTONS")
    step(None, "Right of this sheet",
         "Actions and navigation. They are drawn by the macro module, so they "
         "only appear once macros are enabled. Everything in the Quick Start "
         "above works without them.")

    ws.sheet_view.showGridLines = False
    return ws


# ==========================================================================
# Demo data
# ==========================================================================
def make_demo():
    """Reproducible demo data: 100 rows x 10 features."""
    import random
    random.seed(20260803)
    centres = [25.0, 12.5, 8.0, 50.0, 3.17, 6.32, 10.02, 15.0, 20.0, 30.0]
    sigmas = [0.06, 0.07, 0.05, 0.20, 0.012, 0.028, 0.020, 0.070, 0.060, 0.110]
    shifts = ["A", "B", "C"]
    rows = []
    start = _dt.date(2026, 7, 2)
    for i in range(100):
        d = start + _dt.timedelta(days=i)
        s = i % 3
        meta = [d.isoformat(), shifts[s], f"OP{s + 1}", f"T0{s + 1}", None]
        vals = [round(random.gauss(c, sg), 4) for c, sg in zip(centres, sigmas)]
        rows.append(meta + vals)

    # spec limits: nominal = centre, tolerance = +/- 4 sigma (documented assumption)
    specs = {}
    for f, (c, sg) in enumerate(zip(centres, sigmas), start=1):
        tol = round(4 * sg, 4)
        specs[f] = (c, round(c - tol, 4), round(c + tol, 4))
    return rows, specs


# ==========================================================================
def main(out="SPC_Calculator_v13.xlsx"):
    demo_rows, demo_specs = make_demo()
    wb = Workbook()
    wb.remove(wb.active)

    build_how_to_use(wb)
    build_settings(wb, demo_specs)
    build_data_entry(wb, demo_rows)
    build_form(wb)
    build_control_charts(wb)
    build_monitor(wb)
    build_visual_spc(wb)
    build_capability(wb)
    build_dashboard(wb)
    build_calc(wb)
    build_doc_sheets(wb)

    wb.active = wb.sheetnames.index("How To Use")
    wb.save(out)
    print(f"written: {out}")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "SPC_Calculator_v13.xlsx")
