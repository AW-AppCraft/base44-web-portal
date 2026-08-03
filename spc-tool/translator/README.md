# SPC Translator

Takes messy CSV / Excel exports from any gauge, CMM or hand-typed sheet and
turns them into clean SPC Calculator files. It sits *in front of* the folder
monitor, so the SPC Calculator only ever watches a folder of known-good data.

```
  gauge exports          translator            SPC Calculator
  ─────────────          ──────────            ──────────────
  C:\SPC\inbox    ──►   spc_translate.py  ──►  C:\SPC\spc-ready
                              │                (Settings F8 points here)
                              ├──►  inbox\archive    (sources, kept)
                              ├──►  C:\SPC\rejects   (unusable files)
                              └──►  inbox\translation_log.csv
```

The output folder holds nothing but data files — the log deliberately stays in
the input folder so the monitor never tries to import it.

---

## Running it

```bash
python3 spc_translate.py --in C:\SPC\inbox --out C:\SPC\spc-ready
```

On Windows, edit the three paths at the top of **`translate.bat`** and
double-click it instead. Needs Python 3 (python.org — tick *Add to PATH*).

| Option | Meaning |
|---|---|
| `--in` | folder to read from (required) |
| `--out` | folder the SPC Calculator watches (required) |
| `--archive` | where translated sources go (default `IN\archive`) |
| `--rejects` | where unusable files go (default `IN\rejects`) |
| `--features` | comma-separated feature names, in SPC feature order |
| `--max-features` | cap on readings per row (default 50) |
| `--keep` | translate but do not move anything |
| `--log` | folder for `translation_log.csv` (default `IN`) |
| `--dry-run` | report only, write nothing |
| `--watch N` | keep running, rescan every N seconds |

Start with `--dry-run` on a folder of real exports. It prints exactly what it
would produce, including which column it thinks is which, without touching a
file.

---

## What it works out for itself

**Header row.** Found automatically, including files with a title and blank
lines above the table. Files with no header at all are handled too.

**Separator.** Comma, semicolon, tab or pipe — detected per file.

**Decimal comma.** `25,0312` is read as `25.0312` when the file uses another
separator.

**Which columns are readings.** A column is treated as a measurement when it is
at least 80% numeric and is not one of the traceability fields. Sample counters
are excluded: both by name (`No`, `Index`, `Sample`…) and by shape — a column of
consecutive integers `1,2,3…` is a counter, not a reading.

**Which columns are traceability.** Matched on header name, case and
punctuation insensitive:

| Field | Recognised as |
|---|---|
| Date | date, measured, timestamp, datetime, inspection date, recorded, day |
| Shift | shift, turn, team |
| Operator | operator, op, inspector, user, name, employee |
| Tool ID | tool, tool id, machine, equipment, gauge, gage, device, cell, station |
| Cluster | cluster, batch, lot, serial, part, job, order, cavity |

**Dates** are normalised to `YYYY-MM-DD` from `04/08/2026`, `04.08.2026`,
`2026/08/04`, `20260804`, `4 Aug 2026` and Excel date cells. A trailing time is
dropped.

---

## Defaults for missing fields

Exactly as requested — if the source doesn't carry it, it gets filled in:

| Field | Default |
|---|---|
| Date | **the date the file was translated** |
| Shift | `A` |
| Operator | `Operator 1` |
| Tool ID | `Tool 1` |
| Cluster | left blank |

So a bare file of numbers still produces complete, filterable rows. Every
default applied is listed in the run output and in `translation_log.csv`, so
you always know which traceability is real and which was invented.

**Time is dropped.** SPC Data Entry has no time column. If the source has one,
the log says so rather than silently discarding it.

---

## Keeping features in the right order

By default, reading columns are taken **in the order they appear**. That is fine
when every file comes off the same machine, but if one export puts Length before
Bore, its readings land on the wrong features.

Pin the order:

```bash
python3 spc_translate.py --in inbox --out spc-ready --features "Bore,Length,Runout"
```

Now `Bore` always lands on feature 1, `Length` on feature 2, `Runout` on
feature 3, whatever order the source uses, and columns that match nothing are
dropped. Use the same names as Settings column B in the workbook.

---

## Rejected files

A file is rejected — moved to the rejects folder, never sent onward — when it
cannot be read, is empty, has no column of readings, or has no data rows. The
reason is written to `translation_log.csv`. Nothing malformed reaches the SPC
Calculator.

---

## Log

`translation_log.csv` in the input folder, appended every run:

```
Translated at, Source file, Result, Rows out, Features, Output file, Notes
```

Notes carry the useful detail: which defaults were applied, which columns were
ignored as counters, how many rows had no readings.

---

## Two things it will not do

* **Commas inside quoted values** in a comma-separated file. Splitting happens
  before quotes are considered, so `"25.03, nominal"` breaks the row.
* **Old `.xls`.** Re-save as `.xlsx`. The file is rejected with that reason
  rather than being read wrongly.

---

## Checking the result

The translated files go through the same parser the workbook's VBA uses. To see
what the SPC Calculator would make of them without opening Excel, copy
`../test-data/simulate_import.py` into the output folder and run it — it prints
the row count and value range each file would produce.
