#!/usr/bin/env python3
"""
Generate sample import files for testing SPC v13 folder monitoring.

5 files, 2 features, 35 rows in total. Between them they exercise every layout
the importer accepts, so if all five land correctly the reader is working.

  batch_01_full.csv      header + 5 metadata fields + 2 readings   (8 rows)
  batch_02_full.csv      same layout, no header row                (7 rows)
  batch_03_readings.csv  readings only, no metadata, no header     (7 rows)
  batch_04_semicolon.csv semicolon-separated, header, quoted text  (7 rows)
  batch_05_drift.csv     full layout - contains the excursion      (6 rows)

Feature 1 runs at 25.000 (spec 24.760 / 25.240 in the shipped Settings).
Feature 2 runs at 12.500 (spec 12.220 / 12.780).
batch_05 walks feature 1 upward past the USL so the control chart shows an
out-of-control run and the histogram spikes have something to sit against.

Run: python3 make_test_files.py
"""

import datetime as dt
import random
import pathlib

random.seed(4711)

HERE = pathlib.Path(__file__).parent
HEADER = "Date,Shift,Operator,Tool ID,Cluster,Feature 1,Feature 2"
SHIFTS = ["A", "B", "C"]
START = dt.date(2026, 8, 4)


def meta(i):
    d = START + dt.timedelta(days=i // 3)
    s = i % 3
    return [d.isoformat(), SHIFTS[s], f"OP{s + 1}", f"T0{s + 1}", ""]


def reading(centre, sigma):
    return f"{random.gauss(centre, sigma):.4f}"


def full_rows(start_i, n, f1=(25.000, 0.045), f2=(12.500, 0.055)):
    return [",".join(meta(start_i + k) + [reading(*f1), reading(*f2)])
            for k in range(n)]


def write(name, lines):
    path = HERE / name
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"{name:<26} {len(lines)} lines")


# 1 - the documented layout, with a header row
write("batch_01_full.csv", [HEADER] + full_rows(0, 8))

# 2 - same layout, header omitted (the importer does not require one)
write("batch_02_full.csv", full_rows(8, 7))

# 3 - readings only: no date/shift/operator/tool, no header
write("batch_03_readings.csv",
      [f"{reading(25.000, 0.045)},{reading(12.500, 0.055)}" for _ in range(7)])

# 4 - semicolon separated, quoted text fields (common European export)
rows = []
for k in range(7):
    m = meta(22 + k)
    rows.append(";".join([f'"{m[0]}"', f'"{m[1]}"', f'"{m[2]}"',
                          f'"{m[3]}"', '""',
                          reading(25.000, 0.045), reading(12.500, 0.055)]))
write("batch_04_semicolon.csv",
      ['"Date";"Shift";"Operator";"Tool ID";"Cluster";"Feature 1";"Feature 2"']
      + rows)

# 5 - a real excursion: feature 1 drifts up through the 25.240 USL
drift = []
for k in range(6):
    centre = 25.09 + k * 0.045          # ends around 25.31, past the USL
    drift.append(",".join(meta(29 + k)
                          + [reading(centre, 0.020), reading(12.500, 0.055)]))
write("batch_05_drift.csv", [HEADER] + drift)

total = 8 + 7 + 7 + 7 + 6
print(f"\ntotal data rows across all five files: {total}")
