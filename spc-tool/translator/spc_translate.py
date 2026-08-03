#!/usr/bin/env python3
"""
SPC Translator - turn arbitrary CSV / Excel exports into SPC Calculator format.

Reads every file in an INPUT folder one at a time, works out which columns hold
readings and which hold traceability, fills in whatever is missing with sensible
defaults, and writes a clean file to the OUTPUT folder. The SPC Calculator then
watches the OUTPUT folder only, so it never sees a malformed file.

Output format (always this, always with a header):

    Date,Shift,Operator,Tool ID,Cluster,F1,F2,...,Fn
    2026-08-04,A,Operator 1,Tool 1,,25.0343,12.5482

Defaults when a field is missing from the source:

    Date      the date the file was translated
    Shift     A
    Operator  Operator 1
    Tool ID   Tool 1
    Cluster   blank

Files that yield no usable readings are moved to the REJECTS folder with the
reason recorded, so they never reach the SPC Calculator.

Usage
-----
    python3 spc_translate.py --in IN --out OUT [options]

    --in       folder to read from                       (required)
    --out      folder the SPC Calculator watches         (required)
    --archive  move successfully translated sources here (default: IN/archive)
    --rejects  move unusable sources here                (default: IN/rejects)
    --features comma-separated feature names, in SPC feature order. When given,
               source columns are matched to these names and anything that does
               not match is dropped - this is how you guarantee that feature 3
               is always feature 3, whatever order the source columns arrive in.
    --max-features  cap on readings per row (default 50)
    --keep     do not move source files, just translate
    --log      folder for translation_log.csv     (default: IN)
    --dry-run  report what would happen, write nothing
    --once     translate a single pass and exit (default)
    --watch N  keep running, rescanning every N seconds

Every run appends to translation_log.csv in the INPUT folder. The OUTPUT
folder is kept pristine so the SPC Calculator only ever sees good data files.
"""

import argparse
import csv
import datetime as dt
import os
import re
import shutil
import sys
import time

SPC_HEADER = ["Date", "Shift", "Operator", "Tool ID", "Cluster"]
DATA_EXT = {".csv", ".txt", ".prn", ".tsv"}
EXCEL_EXT = {".xlsx", ".xlsm", ".xls"}

DEFAULT_SHIFT = "A"
DEFAULT_OPERATOR = "Operator 1"
DEFAULT_TOOL = "Tool 1"

# Header names we recognise for the traceability columns. Matching is
# case-insensitive and ignores punctuation, so "Tool_ID" and "tool id" both hit.
FIELD_ALIASES = {
    "date": ["date", "measured", "measurementdate", "day", "datetime",
             "timestamp", "inspectiondate", "recorded"],
    "shift": ["shift", "turn", "team"],
    "operator": ["operator", "op", "inspector", "user", "name", "operatorid",
                 "employee"],
    "tool": ["tool", "toolid", "machine", "machineid", "equipment", "gauge",
             "gage", "gaugeid", "device", "cell", "station"],
    "cluster": ["cluster", "batch", "lot", "serial", "part", "partno",
                "serialno", "job", "order", "cavity"],
    "time": ["time", "clock", "hour"],
}
# Columns that are numeric but are never measurements
NON_MEASURE = ["no", "num", "number", "index", "id", "seq", "sequence",
               "sample", "sampleno", "item", "row", "count", "n"]


def norm(s):
    """Lower-case, strip everything that is not a letter or digit."""
    return re.sub(r"[^a-z0-9]", "", str(s).strip().lower())


def is_number(v):
    if v is None:
        return False
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return True
    s = str(v).strip().replace(" ", "")
    if not s:
        return False
    # tolerate a decimal comma when there is no decimal point
    if "," in s and "." not in s:
        s = s.replace(",", ".")
    try:
        float(s)
        return True
    except ValueError:
        return False


def to_number(v):
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    s = str(v).strip().replace(" ", "")
    if "," in s and "." not in s:
        s = s.replace(",", ".")
    return float(s)


def to_iso_date(v, fallback):
    """Best-effort date parse. Returns YYYY-MM-DD, or the fallback."""
    if v is None or str(v).strip() == "":
        return fallback
    if isinstance(v, dt.datetime):
        return v.date().isoformat()
    if isinstance(v, dt.date):
        return v.isoformat()
    s = str(v).strip()
    # drop a trailing time component
    s = re.split(r"[T ]", s)[0]
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d.%m.%Y", "%Y/%m/%d",
                "%d-%m-%Y", "%Y%m%d", "%d %b %Y", "%d %B %Y"):
        try:
            return dt.datetime.strptime(s, fmt).date().isoformat()
        except ValueError:
            continue
    return fallback


# --------------------------------------------------------------------------
# Reading source files
# --------------------------------------------------------------------------
def read_delimited(path):
    """Return a list of rows (lists of strings)."""
    with open(path, "r", encoding="utf-8-sig", errors="replace") as fh:
        sample = fh.read(8192)
        fh.seek(0)
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
            delim = dialect.delimiter
        except csv.Error:
            counts = {d: sample.count(d) for d in ",;\t|"}
            delim = max(counts, key=counts.get) if max(counts.values()) else ","
        return [row for row in csv.reader(fh, delimiter=delim)]


def read_excel(path):
    try:
        import openpyxl
    except ImportError:
        raise RuntimeError("openpyxl is needed to read Excel files "
                           "(pip install openpyxl)")
    if path.lower().endswith(".xls"):
        raise RuntimeError("old .xls format not supported - re-save as .xlsx")
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    return rows


def read_any(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in EXCEL_EXT:
        return read_excel(path)
    return read_delimited(path)


# --------------------------------------------------------------------------
# Working out the shape of the data
# --------------------------------------------------------------------------
def find_header(rows, scan=25):
    """Index of the header row, or None if the file has no header.

    The header is the FIRST essentially-text row that is directly followed by a
    more numeric row. Taking the first such row matters: a data row whose
    leading fields are a date, an operator and a tool is itself mostly text, so
    a looser rule swallows it as a header and silently loses a reading.
    """
    for i, row in enumerate(rows[:scan]):
        cells = [c for c in row if c is not None and str(c).strip() != ""]
        if len(cells) < 2:
            continue
        share_here = sum(1 for c in cells if is_number(c)) / len(cells)
        if share_here > 0.2:
            continue                      # a header row is essentially all text
        for j in range(i + 1, min(i + 4, len(rows))):
            nxt = [c for c in rows[j] if c is not None and str(c).strip() != ""]
            if not nxt:
                continue
            share_next = sum(1 for c in nxt if is_number(c)) / len(nxt)
            # the row below must be measurably more numeric, otherwise this is
            # just another line of text and the real header is further down
            if share_next > 0 and share_next > share_here:
                return i
            break
    return None


def first_data_row(rows):
    for i, row in enumerate(rows):
        cells = [c for c in row if c is not None and str(c).strip() != ""]
        if cells and any(is_number(c) for c in cells):
            return i
    return None


def classify_columns(rows, header_idx, wanted_features, max_features):
    """Decide what each column is.

    Returns (fields, measure_cols, feature_names, notes) where `fields` maps
    date/shift/operator/tool/cluster to a column index (or None), and
    `measure_cols` is the ordered list of columns holding readings.
    """
    notes = []
    header = rows[header_idx] if header_idx is not None else []
    ncols = max(len(r) for r in rows) if rows else 0
    data_rows = rows[(header_idx + 1) if header_idx is not None
                     else first_data_row(rows) or 0:]

    fields = {k: None for k in ("date", "shift", "operator", "tool", "cluster")}
    named = {}
    if header_idx is not None:
        for c in range(min(len(header), ncols)):
            key = norm(header[c])
            if not key:
                continue
            named[c] = str(header[c]).strip()
            for field, aliases in FIELD_ALIASES.items():
                if field == "time":
                    continue
                if any(key == a or key.startswith(a) for a in aliases):
                    if fields.get(field) is None:
                        fields[field] = c
                    break

    # numeric-ness per column, over the data rows
    numeric_share = {}
    for c in range(ncols):
        vals = [r[c] for r in data_rows
                if c < len(r) and r[c] is not None and str(r[c]).strip() != ""]
        if not vals:
            numeric_share[c] = 0.0
            continue
        numeric_share[c] = sum(1 for v in vals if is_number(v)) / len(vals)

    taken = {v for v in fields.values() if v is not None}
    measure = []
    for c in range(ncols):
        if c in taken:
            continue
        if numeric_share.get(c, 0) < 0.8:
            continue
        key = norm(named.get(c, ""))
        if key and key in NON_MEASURE:
            notes.append(f"column '{named[c]}' looks like a counter, not a reading")
            continue
        # a column of consecutive integers 1..n is a sample counter
        vals = [to_number(r[c]) for r in data_rows
                if c < len(r) and is_number(r[c])]
        if len(vals) >= 5 and all(float(v).is_integer() for v in vals):
            if vals == list(range(int(vals[0]), int(vals[0]) + len(vals))):
                notes.append(f"column {c + 1} is a 1,2,3... counter, ignored")
                continue
        measure.append(c)

    feature_names = [named.get(c, f"Column {c + 1}") for c in measure]

    if wanted_features:
        wanted_norm = [norm(w) for w in wanted_features]
        remapped, remapped_names = [], []
        for w, wn in zip(wanted_features, wanted_norm):
            hit = None
            for c in measure:
                if norm(named.get(c, "")) == wn:
                    hit = c
                    break
            remapped.append(hit)
            remapped_names.append(w)
        if all(h is None for h in remapped):
            notes.append("no source column matched --features; using column order")
        else:
            measure, feature_names = remapped, remapped_names

    if len(measure) > max_features:
        notes.append(f"{len(measure)} reading columns found, keeping first {max_features}")
        measure = measure[:max_features]
        feature_names = feature_names[:max_features]

    return fields, measure, feature_names, notes


# --------------------------------------------------------------------------
# Translation
# --------------------------------------------------------------------------
def translate(path, args, today):
    """Returns (out_rows, feature_names, notes, reason_if_rejected)."""
    try:
        rows = read_any(path)
    except Exception as exc:                                  # noqa: BLE001
        return [], [], [], f"could not read file: {exc}"

    rows = [r for r in rows if any(
        c is not None and str(c).strip() != "" for c in r)]
    if not rows:
        return [], [], [], "file is empty"

    header_idx = find_header(rows)
    fields, measure, names, notes = classify_columns(
        rows, header_idx, args.features, args.max_features)

    if not measure:
        return [], [], notes, "no column of readings found"

    start = (header_idx + 1) if header_idx is not None else (first_data_row(rows) or 0)
    out, skipped = [], 0

    for r in rows[start:]:
        readings = []
        for c in measure:
            if c is not None and c < len(r) and is_number(r[c]):
                readings.append(f"{to_number(r[c]):.4f}")
            else:
                readings.append("")
        if not any(readings):
            skipped += 1
            continue

        def pick(field, default):
            c = fields.get(field)
            if c is None or c >= len(r) or r[c] is None:
                return default
            v = str(r[c]).strip()
            return v if v else default

        out.append([
            to_iso_date(r[fields["date"]] if fields["date"] is not None
                        and fields["date"] < len(r) else None, today),
            pick("shift", DEFAULT_SHIFT),
            pick("operator", DEFAULT_OPERATOR),
            pick("tool", DEFAULT_TOOL),
            pick("cluster", ""),
        ] + readings)

    if skipped:
        notes.append(f"{skipped} row(s) had no readings and were dropped")
    for field, default in (("date", today), ("shift", DEFAULT_SHIFT),
                           ("operator", DEFAULT_OPERATOR), ("tool", DEFAULT_TOOL)):
        if fields.get(field) is None:
            notes.append(f"{field} not in source, set to '{default}'")
    if any(norm(str(h)) in FIELD_ALIASES["time"]
           for h in (rows[header_idx] if header_idx is not None else [])):
        notes.append("a time column was present; SPC Data Entry has no time "
                     "field, so it was not carried over")

    if not out:
        return [], names, notes, "no data rows with readings"
    return out, names, notes, None


def write_output(out_path, rows, feature_names):
    header = SPC_HEADER + [n if n else f"F{i + 1}"
                           for i, n in enumerate(feature_names)]
    with open(out_path, "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(header)
        w.writerows(rows)


def move_to(path, folder, dry_run):
    if dry_run:
        return
    os.makedirs(folder, exist_ok=True)
    target = os.path.join(folder, os.path.basename(path))
    stem, ext = os.path.splitext(target)
    n = 1
    while os.path.exists(target):
        target = f"{stem}_{n}{ext}"
        n += 1
    shutil.move(path, target)


def log(log_dir, row, dry_run):
    if dry_run:
        return
    os.makedirs(log_dir, exist_ok=True)
    path = os.path.join(log_dir, "translation_log.csv")
    new = not os.path.exists(path)
    with open(path, "a", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        if new:
            w.writerow(["Translated at", "Source file", "Result", "Rows out",
                        "Features", "Output file", "Notes"])
        w.writerow(row)


def run_once(args):
    today = dt.date.today().isoformat()
    os.makedirs(args.out, exist_ok=True)

    candidates = sorted(
        f for f in os.listdir(args.in_dir)
        if os.path.isfile(os.path.join(args.in_dir, f))
        and os.path.splitext(f)[1].lower() in (DATA_EXT | EXCEL_EXT))

    if not candidates:
        print("nothing to translate")
        return 0

    ok = bad = total_rows = 0
    for name in candidates:                      # one file at a time
        src = os.path.join(args.in_dir, name)
        rows, names, notes, reason = translate(src, args, today)
        stamp = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        if reason:
            bad += 1
            print(f"REJECT  {name:<40} {reason}")
            log(args.log, [stamp, name, "REJECTED", 0, "", "",
                           "; ".join([reason] + notes)], args.dry_run)
            if not args.keep:
                move_to(src, args.rejects, args.dry_run)
            continue

        stem = os.path.splitext(name)[0]
        out_name = f"{stem}_spc.csv"
        out_path = os.path.join(args.out, out_name)
        n = 1
        while os.path.exists(out_path):
            out_name = f"{stem}_spc_{n}.csv"
            out_path = os.path.join(args.out, out_name)
            n += 1

        if not args.dry_run:
            write_output(out_path, rows, names)
        ok += 1
        total_rows += len(rows)
        print(f"OK      {name:<40} {len(rows):>4} rows, "
              f"{len(names)} feature(s) -> {out_name}")
        for note in notes:
            print(f"        note: {note}")
        log(args.log, [stamp, name, "OK", len(rows), " | ".join(names),
                       out_name, "; ".join(notes)], args.dry_run)
        if not args.keep:
            move_to(src, args.archive, args.dry_run)

    print(f"\n{ok} translated ({total_rows} rows), {bad} rejected")
    if args.dry_run:
        print("dry run - nothing was written or moved")
    return 0 if bad == 0 else 1


def main(argv=None):
    p = argparse.ArgumentParser(
        description="Translate CSV/Excel exports into SPC Calculator format.")
    p.add_argument("--in", dest="in_dir", required=True, help="input folder")
    p.add_argument("--out", required=True,
                   help="output folder (the one SPC Calculator watches)")
    p.add_argument("--archive", help="move translated sources here")
    p.add_argument("--rejects", help="move unusable sources here")
    p.add_argument("--features", help="comma-separated feature names in SPC order")
    p.add_argument("--max-features", type=int, default=50)
    p.add_argument("--keep", action="store_true", help="do not move sources")
    p.add_argument("--log", help="folder for translation_log.csv (default: IN)")
    p.add_argument("--dry-run", action="store_true")
    p.add_argument("--watch", type=int, metavar="SECONDS",
                   help="keep running, rescanning every N seconds")
    args = p.parse_args(argv)

    if not os.path.isdir(args.in_dir):
        p.error(f"input folder not found: {args.in_dir}")
    args.archive = args.archive or os.path.join(args.in_dir, "archive")
    args.rejects = args.rejects or os.path.join(args.in_dir, "rejects")
    args.log = args.log or args.in_dir
    args.features = ([f.strip() for f in args.features.split(",") if f.strip()]
                     if args.features else None)

    if not args.watch:
        return run_once(args)

    print(f"watching {args.in_dir} every {args.watch}s - Ctrl+C to stop")
    try:
        while True:
            run_once(args)
            time.sleep(max(5, args.watch))
    except KeyboardInterrupt:
        print("\nstopped")
        return 0


if __name__ == "__main__":
    sys.exit(main())
