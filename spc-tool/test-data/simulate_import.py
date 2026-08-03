#!/usr/bin/env python3
"""
Simulate the VBA importer (ImportDelimited + WriteRow) against the test files.

The VBA cannot be run here, so this mirrors its logic line for line and reports
what would land on Data Entry. If this says 35 rows, so should Excel.

Run: python3 simulate_import.py
"""

import glob
import os


def is_numeric(s):
    """VBA IsNumeric, near enough for these files."""
    try:
        float(s)
        return True
    except ValueError:
        return False


def clean_field(s):
    t = s.strip()
    if len(t) >= 2 and t.startswith('"') and t.endswith('"'):
        t = t[1:-1]
    return t.strip()


def write_row(parts):
    """Returns (meta, readings) or None if the line is not a data row."""
    if not parts:
        return None
    parts = [clean_field(p) for p in parts]

    if is_numeric(parts[0]):
        first_feat, meta = 0, [""] * 5           # layout (b)
    else:
        if len(parts) < 6:
            return None                           # too short to be layout (a)
        if not is_numeric(parts[5]):
            return None                           # header row
        meta, first_feat = parts[0:5], 5

    readings, written = [], 0
    for i in range(first_feat, len(parts)):
        if len(readings) >= 50:
            break
        if parts[i] and is_numeric(parts[i]):
            readings.append(float(parts[i]))
            written += 1
        else:
            readings.append(None)
    return (meta, readings) if written else None


def import_file(path):
    rows = []
    with open(path, encoding="utf-8") as fh:
        for line in fh:
            line = line.strip()
            if not line:
                continue
            line = line.replace("\t", ",")
            if "," not in line and ";" in line:
                line = line.replace(";", ",")
            row = write_row(line.split(","))
            if row:
                rows.append(row)
    return rows


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    total = 0
    print(f"{'file':<26} {'rows':>5}  {'F1 (col H)':>22}  {'F2 (col I)':>22}  meta")
    print("-" * 100)
    for path in sorted(glob.glob(os.path.join(here, "*.csv"))):
        rows = import_file(path)
        total += len(rows)
        f1 = [r[1][0] for r in rows if r[1] and r[1][0] is not None]
        f2 = [r[1][1] for r in rows if len(r[1]) > 1 and r[1][1] is not None]
        has_meta = any(r[0][0] for r in rows)
        print(f"{os.path.basename(path):<26} {len(rows):>5}  "
              f"{min(f1):.4f}-{max(f1):.4f} (n={len(f1):>2})  "
              f"{min(f2):.4f}-{max(f2):.4f} (n={len(f2):>2})  "
              f"{'yes' if has_meta else 'blank (readings-only file)'}")

    print("-" * 100)
    print(f"{'TOTAL':<26} {total:>5}")

    problems = []
    if total != 35:
        problems.append(f"expected 35 data rows, parser produced {total}")
    for path in sorted(glob.glob(os.path.join(here, "*.csv"))):
        for meta, readings in import_file(path):
            if len(readings) != 2:
                problems.append(f"{os.path.basename(path)}: {len(readings)} "
                                f"readings on a row, expected 2")
                break
    print()
    if problems:
        for p in problems:
            print("PROBLEM:", p)
        return 1
    print("OK - all five files parse to 35 rows, 2 readings each, headers skipped")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
