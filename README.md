# Base44 Web Portal

## Project Description
Base44 Web Portal is a robust application designed to streamline the management and deployment of web applications with ease and efficiency.

## Features
- User-friendly interface for managing web applications
- Integration with various APIs for enhanced functionality
- Secure subscription management system
- Deployment capabilities for GitHub Pages

## Setup Instructions
1. Clone the repository:
   ```bash
   git clone https://github.com/AW-AppCraft/base44-web-portal.git
   ```
2. Navigate to the project directory:
   ```bash
   cd base44-web-portal
   ```
3. Install the necessary dependencies:
   ```bash
   npm install
   ```

## How to Add Apps
To add a new application, follow these steps:
1. Navigate to the `apps` directory.
2. Create a new folder for your app and add the necessary files.
3. Update the configuration file to include your app details.

## Subscription Management
Manage user subscriptions through the admin panel, where you can:
- View current subscriptions
- Edit subscription plans
- Add new users and assign their subscription levels.

## Deployment Guide for GitHub Pages
1. Ensure your project is ready for production.
2. Run the build command:
   ```bash
   npm run build
   ```
3. Deploy to GitHub Pages:
   ```bash
   npm run deploy
   ```
4. Visit your GitHub Pages link to see your app in action!

---

## Drawing Ballooning & FAIR/ISIR Automation (`ballooning.html`)

A browser-based tool that turns an engineering drawing PDF into a **bible drawing**
(ballooned print + numbered characteristic list) and exports the paperwork that
ISIR and FAIR submissions need.

Open `ballooning.html` — no build step, no server, no upload. Everything runs
client-side, so drawings never leave the machine (which matters for ITAR/EAR and
customer-proprietary prints).

### What it extracts

| Category | Recognized |
|---|---|
| Dimensions | linear, diameter, radius, angular, chamfer, bilateral (`25.4 ±0.05`), unequal bilateral (`+0.03/-0.01`), limit (`12.50/12.45`), reference `(38.10)`, quantity prefixes (`4X ⌀5.2`) |
| GD&T | all 14 characteristic symbols, tolerance value, diametral zone, MMC/LMC/projected modifiers, datum references |
| Threads | metric (`M8x1.25-6H`), unified (`1/4-20 UNC-2B`), NPT |
| Surface finish | `Ra`, `Rz`, µin/RMS |
| Material | AMS, ASTM, MIL, QQ, UNS, AISI, SAE, alloy/temper designations, superalloys, PH and austenitic stainless, EN/DIN, polymers |
| Coating / finish | anodize (MIL-A-8625), chem film (MIL-DTL-5541), zinc (B633), electroless nickel (B733/AMS 2404), passivation (AMS 2700), cadmium, black oxide, aerospace paint systems, ISO plating specs, powder/e-coat, dry film lube |
| Process | heat treat, hardness, shot peen, penetrant inspection, edge condition, welding specs |

### Workflow

1. Load the drawing PDF.
2. Set units, the title-block default tolerances (`.X` / `.XX` / `.XXX` / angular)
   and the border zone grid.
3. The tool highlights every requirement, colors it by category, drops a numbered
   balloon with a leader line, and fills the characteristic table.
4. Review — drag balloons, edit any field, add balloons the parser missed,
   delete false positives, renumber.
5. Export.

Numbering follows inspector reading order: left-to-right within horizontal bands,
top-to-bottom, sheet by sheet.

### Exports

- **AS9102 Rev C Form 3** (CSV) — characteristic accountability, with measurement
  columns left blank for inspection to fill.
- **PPAP / ISIR dimensional results** (CSV) — nominal, limits, suggested
  inspection method, and N sample columns.
- **Ballooned sheet** (PNG) — the bible drawing itself, balloons and leaders
  burned in.
- **Project** (JSON) — round-trippable state.

Each characteristic also carries a suggested inspection method derived from its
type and tolerance band (CMM, air gage, micrometer, thread gage, profilometer…),
as a starting point for the inspection plan.

### Known limits — read before submitting a FAIR

These are properties of the problem, not gaps to be patched later:

- **Extraction is an accelerator, not an authority.** AS9102 puts characteristic
  accountability on the supplier. Every balloon needs a human check.
- **Basic dimensions cannot be detected from text alone.** The rectangle around a
  basic dimension is drawing *geometry*, not text. Untoleranced dimensions are
  flagged for review rather than guessed at.
- **Legacy CAD symbol fonts.** `gdt.shx` and its relatives map GD&T glyphs onto
  ASCII letters, so a position callout extracts as `j 0.25 m A B C`. The tool
  remaps them best-effort via `SYMBOL_FONT_MAP` in `ballooning.js` and always
  flags the result — calibrate the map against a known drawing before trusting it.
- **Scanned drawings yield nothing from the text layer.** No text layer, no
  parsing — use the vision pass below, or balloon manually.
- **Association is positional, not semantic.** The tool knows a callout's location
  on the sheet, not which feature it dimensions. Extracting that requires the CAD
  model, not the PDF.

### Vision pass — local Ollama (optional)

Covers what the text-layer parser structurally cannot: **scanned drawings,
outlined text, and legacy symbol fonts.** The page calls Ollama running on the
same machine, so the drawing is never uploaded anywhere — no cloud API, no
vendor, nothing leaves the shop.

**Setup**

```bash
ollama pull gemma4:12b        # multimodal. The larger 26b/31b variants read small text better.
```

Any vision-capable model works — the model list is populated from your Ollama and
labelled with **real capabilities** reported by `/api/show`, not by matching model
names against a hardcoded list. Text-only models are labelled `(no vision)` and
refused if selected, rather than failing with an unhelpful error mid-pass.

Ollama blocks cross-origin browser requests by default, so allow this page's
origin and restart it:

```bash
# macOS / Linux
OLLAMA_ORIGINS="http://localhost:8080" ollama serve

# Windows PowerShell — then quit Ollama from the tray and reopen it
setx OLLAMA_ORIGINS "http://localhost:8080"
```

Serve the folder over HTTP rather than opening the file from disk — a `file://`
page has a null origin, which cannot be allowed:

```bash
npx http-server -p 8080
```

Then: **Connect** → pick the model → **Run vision pass on this sheet**. The
"CORS help" button prints the exact command for whatever origin you are on.

**How it works, and why it is built this way**

- **Tiled, not whole-sheet.** A D-size drawing squeezed into one image renders
  dimension text a few pixels tall and unreadable. The sheet is split into an
  overlapping grid (default 2×2, raise it for dense drawings) and each region is
  read separately at high DPI. Overlap stops a callout on a seam being cut in
  half; duplicates are removed afterwards.
- **The model does perception only.** It transcribes glyphs and returns rough
  boxes. Every tolerance, limit, and inspection method is then computed by the
  same deterministic parser used for the text layer — the model is never asked
  to do arithmetic, so its errors stay confined to transcription.
- **The prompt is closed-ended.** Open-ended prompts make small vision models
  narrate geometry and invent plausible-looking dimensions. Asking only for
  glyphs actually present keeps the failure mode at "missed a callout" rather
  than "fabricated a callout".
- **Text layer always wins.** Where both sources see the same callout, the exact
  text-layer value is kept and the transcription is discarded. The number
  suppressed is reported, so a genuinely repeated callout is not lost silently.
- **Everything is flagged.** Vision results get purple balloons and a permanent
  review flag. They are a starting point for the operator, not a result.
- **Thinking is switched off** on models that support it. Transcription is not a
  reasoning task, and deliberating over every tile costs minutes across a tiled
  sheet for no accuracy gain.

**Expectations.** A 12B model reading a dense D-size drawing will miss callouts
and misread digits — small vision models are weakest at exactly what matters
here: tiny text, GD&T glyphs, and stacked tolerances. Treat it as a way to avoid
typing a scanned drawing from scratch, not as an inspector. Read every balloon.

### Where this goes next

The vision pass above closes the scanned-drawing gap. The remaining step is
reading the native CAD file, where dimensions, tolerances and
their owning features are already structured data — which removes the guesswork
entirely.

### Files

| File | Role |
|---|---|
| `ballooning.html` | UI |
| `ballooning.js` | extraction engine — PDF text → classified characteristics → exporters |
| `vision.js` | optional vision pass — tiling, Ollama client, dedup |
| `ballooning-ui.js` | viewer, balloon overlay, editable table, persistence |
| `ballooning.css` | styles (scoped `.bl-*`) |

---

*For further assistance, please refer to the documentation or contact support.*