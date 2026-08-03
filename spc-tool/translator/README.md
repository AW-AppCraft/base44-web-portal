# SPC Translator — Excel edition

Takes messy CSV / Excel exports from any gauge, CMM or hand-typed sheet and
turns them into clean SPC Calculator files. It sits *in front of* the folder
monitor, so the SPC Calculator only ever watches a folder of known-good data.

```
  gauge exports         SPC_Translator.xlsm        SPC Calculator
  ─────────────         ───────────────────        ──────────────
  C:\SPC\inbox    ──►   Preview / Translate   ──►  C:\SPC\spc-ready
                              │                    (Settings F8 points here)
                              ├──►  inbox\archive   translated sources
                              ├──►  inbox\rejects   unusable files
                              └──►  Log sheet       one row per file
```

## Files

| File | What it is |
|---|---|
| `SPC_Translator.xlsx` | The workbook — Config, Log and How To Use sheets |
| `SPC_Translator_CODE.txt` | The macro module to import (rename to `.bas`) |
| `build_translator.py` | Rebuilds the workbook from scratch |
| `spc_translate.py` | Optional command-line version, same logic, for scheduled/headless use |
| `translate.bat` | Launcher for the command-line version |

---

## Setup — once

1. Open `SPC_Translator.xlsx` and **save it as `.xlsm`** (Excel Macro-Enabled
   Workbook). Macros cannot be stored in a `.xlsx`.
2. `Alt+F11` → *File* → *Import File* → `SPC_Translator.bas`. If you were sent
   the `.txt`, rename it to `.bas` first, or paste the contents into a new
   module and drop the first `Attribute` line.
3. Save, close, reopen. The buttons appear on the **How To Use** sheet
   automatically. If they are ever missing, run `TR_BuildButtons` from `Alt+F8`.
4. On the **Config** sheet set `B4` (raw exports arrive here) and `B5` (clean
   files go here). Point **Settings F8** of the SPC Calculator at that same
   `B5` folder.

## Every day

| Button | Does |
|---|---|
| **Preview (writes nothing)** | Reports what each file would produce — check the feature count matches what you actually measure |
| **Translate Now** | One pass over the input folder |
| **Start / Stop Watching** | Keeps checking every `B9` seconds |
| **Open Input / Output Folder** | Opens them in Explorer |
| **Clear Log** | Empties the Log sheet |

---

## Config sheet

| Cell | Setting |
|---|---|
| `B4` | Input folder |
| `B5` | Output folder — the one the SPC Calculator watches |
| `B6` | Archive folder (blank = `input\archive`) |
| `B7` | Rejects folder (blank = `input\rejects`) |
| `B8` | File pattern — `*.csv`, `*.xlsx`, or `*.*` |
| `B9` | Watch interval, seconds (minimum 10) |
| `B10` | Move source files after translating, Y/N |
| `B11` | Max features per row |
| `B12`–`B15` | Defaults: Shift, Operator, Tool ID, Date |
| `B17`–`B21` | Status, last run, counts — written by the macro |
| `B25:B74` | Optional feature order map |

---

## What it works out by itself

**Header row** — found automatically, including files with a title and blank
lines above the real table. Files with no header work too.

**Separator** — comma, semicolon, tab or pipe, detected per file. Quoted fields
are honoured, so a comma inside `"25.03, nominal"` does not break the row.

**Decimal comma** — `25,0312` is read as `25.0312` when the file uses another
separator. Output is always written with a full stop regardless of the PC's
regional settings.

**Which columns are readings** — a column counts as a measurement when at least
80% of its values are numeric and it is not one of the traceability fields.
Sample counters are excluded both by name (`No`, `Index`, `Sample`…) and by
shape: a column of consecutive integers `1,2,3…` is a counter, not a reading.

**Which columns are traceability** — matched on heading, ignoring case and
punctuation:

| Field | Recognised as |
|---|---|
| Date | date, measured, timestamp, datetime, inspection date, recorded, day |
| Shift | shift, turn, team |
| Operator | operator, op, inspector, user, employee |
| Tool ID | tool, tool id, machine, equipment, gauge, gage, device, cell, station |
| Cluster | cluster, batch, lot, serial, part, job, order, cavity |

**Dates** — normalised to `YYYY-MM-DD`, including real Excel date cells. A
trailing time is dropped.

---

## Missing fields get defaults

| Field | Default | Change on Config |
|---|---|---|
| Date | **the date the file is translated** | `B15` |
| Shift | `A` | `B12` |
| Operator | `Operator 1` | `B13` |
| Tool ID | `Tool 1` | `B14` |
| Cluster | blank | — |

So a bare list of numbers still produces complete, filterable rows. **Every
default applied is named in the Log**, so you always know which traceability is
real and which was invented.

**Time is dropped** — SPC Data Entry has no time column. If the source has one
the Log says so rather than discarding it silently.

---

## Keeping features in the right order

By default, reading columns are taken **in the order they appear**. That is fine
when every file comes off the same machine, but if one export puts Length before
Bore, its readings land on the wrong features.

Fill in `B25:B74` on the Config sheet with your feature names in SPC order:

```
   B25   Bore
   B26   Length
   B27   Runout
```

Now `Bore` always lands on feature 1 whatever order the source uses, and columns
matching nothing are dropped. Use the same names as Settings column B in the
SPC Calculator.

---

## Rejected files

A file is rejected — moved to the rejects folder, never sent onward — when it
cannot be read, is empty, has no column of readings, or has no data rows. The
reason goes in the Log. Nothing malformed reaches the SPC Calculator.

---

## Known limits

* **Old `.xls`** — re-save as `.xlsx`. The file is rejected with that reason
  rather than read wrongly.
* **One table per file** — the first sheet of a workbook, and one table in it.
  Multi-block reports need splitting first.
* **Ambiguous dates** — `04/08/2026` is read using the PC's regional settings,
  so it means 4 August in the UK and 8 April in the US. Sources that emit
  `YYYY-MM-DD` avoid the question entirely.

---

## Command-line version

`spc_translate.py` does the same job from a command prompt or a scheduled task,
for anyone who wants it running without Excel open:

```bash
python3 spc_translate.py --in C:\SPC\inbox --out C:\SPC\spc-ready --watch 30
```

Same detection rules, same defaults, same reject behaviour. It needs Python 3;
the Excel edition does not.
