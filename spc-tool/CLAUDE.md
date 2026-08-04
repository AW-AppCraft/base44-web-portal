# Working rules

## Working folder

**All Claude sessions for this work use:**

```
C:\Hermes-Ollama\projects\Claude local sessions
```

Create anything new under that folder. Do not scatter work into unrelated
repositories.

This rule only binds sessions that can actually see that path — i.e. Claude Code
running **locally** on the Windows machine. A cloud/web session runs in a Linux
container with no access to `C:\`, so if a session cannot reach the folder it
must say so up front rather than silently working somewhere else, and hand the
result back as downloadable files.

To make this rule global for every local session, this file also belongs at:

```
C:\Users\<you>\.claude\CLAUDE.md
```

---

# SPC Tool — project notes

An all-in-one SPC package for R&G Precision Engineering. Two Excel tools plus
the Python that generates them.

| Path | What it is |
|---|---|
| `SPC_Calculator_v13.xlsx` | The calculator — 13 sheets, 45 charts |
| `SPC_V13_Monitor_CODE.txt` / `.bas` | Its macros: folder watch, import, auto-scale, kiosk |
| `build_spc_v13.py` | **Source of truth** — regenerates the workbook |
| `validate_v13.py` | Static checks; run after every build |
| `translator/SPC_Translator.xlsx` | Converts arbitrary exports into SPC format |
| `translator/SPC_Translator_CODE.txt` | Its macros |
| `translator/build_translator.py` | Regenerates the translator workbook |
| `test-data/` | Five sample import files + format spec |

## Rules for changing it

1. **Never hand-edit the .xlsx.** Change `build_spc_v13.py` and rebuild, or the
   next rebuild silently discards the edit.
2. **Always run `validate_v13.py` after building.** It must print
   `OK - no structural problems found`.
3. **Excel formulas must be Excel-2007-safe.** No `XLOOKUP`, `FILTER`,
   `TEXTJOIN`, `UNIQUE`, `SEQUENCE`.
4. **Chart helper columns are hidden**, so every chart needs
   `visible_cells_only = False` or it renders empty.
5. **VBA cannot be compiled in a cloud session.** Structural checks only —
   procedure balance, argument counts, no single-line `Sub`/`End Sub`. Say so
   rather than implying it was tested.

## Bugs already fixed — do not reintroduce

| Bug | Guard |
|---|---|
| Formula built without a leading `=` (stored as text, killed every chart) | `validate_v13.py` rejects text that starts with a function call |
| Charts reading hidden columns with `plotVisOnly` left at default | validator checks every chart |
| Single-point overlay series with no marker (invisible spec limits) | validator requires markers on histogram overlays |
| `IF(A1=NA(), …)` — always `#N/A` whatever `A1` holds; use `ISNA()` | validator rejects the pattern |
| `NextFreeRow` via `End(xlUp)` on column H — overwrote collected data | uses `MATCH` across all 50 feature columns, then walks past occupied rows |
| `Line Input #` on LF-only files — read whole file as one line, scattered readings | whole-file read, split on CRLF/CR/LF |
| `Dir()` called inside a `Dir()` enumeration loop | file list collected before processing |
| Blank spec limit read as `0`, giving nonsense Cp/Cpk | Capability returns `""` for blank limits |
| `Format$` / `.Text` on a comma-decimal locale | `NumStr()` and `.Value` |

## Conventions

- Yellow fill + blue text = user input. Grey = calculated.
- Arial throughout.
- Data Entry: column H = feature 1 … BE = feature 50; rows 5–10004.
- Import files: CRLF, `YYYY-MM-DD` dates, full-stop decimal.
