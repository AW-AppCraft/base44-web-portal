# Import file format — SPC Calculator v13

What the folder monitor accepts, and five files to prove it works.

---

## The short version

One row per part. Readings in **feature order**: the first reading goes to
Data Entry column **H** (feature 1), the second to **I** (feature 2), and so on
up to 50. Nothing else matters.

```
Date,Shift,Operator,Tool ID,Cluster,Feature 1,Feature 2
2026-08-04,A,OP1,T01,,25.0343,12.5482
```

---

## Accepted layouts

### Layout A — metadata then readings (recommended)

```
Date, Shift, Operator, Tool ID, Cluster, F1, F2, ... Fn
```

The first five fields land in Data Entry columns C–F (Cluster in G) and are what
the **Visual SPC** filters use. Leave any of them empty if you don't collect
them — write the comma, keep the position:

```
2026-08-04,A,OP1,T01,,25.0343,12.5482
                    ^^ Cluster empty, position preserved
```

### Layout B — readings only

```
F1, F2, ... Fn
```

```
25.0825,12.4772
```

Used when the gauge exports nothing but numbers. Date/shift/operator stay blank,
so those rows are invisible to the Visual SPC filters but appear on every chart.

### Which layout is used

Decided per row, by the first field: **numeric → layout B, otherwise layout A**.
You can mix both in one folder; you should not mix them inside one file.

---

## Rules

| Rule | Detail |
|---|---|
| **Header row** | Optional. A row with text where the first reading belongs is skipped, so `Date,Shift,...,Feature 1,Feature 2` is discarded automatically. |
| **Separators** | Comma, tab, or semicolon. Semicolons are only treated as separators when the line contains no comma, so a genuine comma file is never mangled. |
| **Quotes** | `"2026-08-04"` is fine — surrounding double quotes are stripped from every field. |
| **Decimal mark** | Must be a full stop: `25.0343`, not `25,0343`. A comma decimal would be read as a field separator. |
| **Commas inside values** | Not supported. The parser splits on separators without honouring quoted commas. |
| **Blank readings** | A blank between commas leaves that feature empty for that row. |
| **Non-numeric reading** | Ignored for that cell; the rest of the row still imports. |
| **Rows with no valid reading** | Skipped entirely — they never reach Data Entry. |
| **Line endings** | CRLF or LF are both fine from v13 onward — the importer reads the whole file and splits on any ending. Older builds used `Line Input #`, which ends a line on CR/CRLF only, so an LF-only file was read as one enormous line and its readings were scattered across the sheet. |
| **Date format** | `YYYY-MM-DD`. The Visual SPC date filters compare text, so anything else will not filter correctly. |
| **File types** | `*.csv`, `*.txt`, `*.prn`, `*.xlsx`, `*.xlsm`, `*.xls`. For workbooks the first sheet is read. |
| **Duplicate imports** | Each file is recorded on the hidden `_ImportLog` by name and modified time. Re-dropping an unchanged file imports nothing; a modified file is re-imported. |

---

## The five test files

35 data rows, 2 features, one file per layout variation.

| File | Rows | What it tests |
|---|---|---|
| `batch_01_full.csv` | 8 | Layout A with a header row |
| `batch_02_full.csv` | 7 | Layout A with no header row |
| `batch_03_readings.csv` | 7 | Layout B, readings only |
| `batch_04_semicolon.csv` | 7 | Semicolon separated, every field quoted |
| `batch_05_drift.csv` | 6 | Layout A, and a genuine process excursion |

`batch_05_drift.csv` walks feature 1 from 25.13 up to 25.32, through the 25.240
USL in the shipped Settings. Import it last and the control chart should show a
rising run breaking the upper limit, `In spec?` should flip to `FAIL`, and the
histogram's purple USL triangle should end up inside the right-hand bars rather
than clear of them. If that doesn't happen, something is wrong.

Regenerate them at any time:

```bash
python3 make_test_files.py
```

Check the layout from inside Excel before importing anything — the
**Preview Import** button (`SPC_ImportPreview`) reports rows and columns per
file and writes nothing. If a file shows far more columns than the features you
measure, its layout is being misread.

Check what the importer would make of them, without Excel:

```bash
python3 simulate_import.py
```

That mirrors the VBA `ImportDelimited` / `WriteRow` logic line for line and
prints the rows and value ranges each file would produce.

---

## Test run

1. Open the workbook, save as `.xlsm`, import `SPC_V13_Monitor.bas`.
2. On **Settings**: `F8` = the folder holding these files, `F9` = `*.csv`,
   `F10` = `15`, `F11` = `N` while testing (so files stay put and you can
   re-run).
3. Optional but clearer: set Active = `N` on Settings for features 3–10, so only
   the two features with data are charted.
4. Run `SPC_ClearData` to drop the shipped demo rows.
5. Press **Start Monitoring**. Watch the Monitor sheet: Status `RUNNING`,
   Files Imported `5`, Rows Imported `35`.
6. Data Entry rows 5–39 should now hold the data, with columns C–F populated for
   everything except the seven rows from `batch_03_readings.csv`.

Expected on Feature 1 after all five files: n = 35, mean near 25.03, and the
last six points climbing away from centre.

To re-run from scratch: `SPC_ClearData`, then clear the hidden `_ImportLog`
sheet (unhide via right-click on a tab → Unhide) or the files will be recognised
as already imported.
