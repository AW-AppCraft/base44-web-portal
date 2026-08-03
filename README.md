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
- **Scanned drawings yield nothing.** No text layer, no extraction. Balloon
  manually, or pre-process the file with OCR or a vision model first.
- **Association is positional, not semantic.** The tool knows a callout's location
  on the sheet, not which feature it dimensions. Extracting that requires the CAD
  model, not the PDF.

### Where this goes next

The natural upgrade for scanned drawings and for feature association is a vision
model pass (e.g. the Claude API) over each sheet region, returning structured
characteristics that feed the same table and exporters. The extraction stage in
`ballooning.js` is deliberately separated from numbering, rendering and export so
a second extractor can be added without touching the rest of the pipeline. The
step after that is reading the native CAD file, where dimensions, tolerances and
their owning features are already structured data — which removes the guesswork
entirely.

### Files

| File | Role |
|---|---|
| `ballooning.html` | UI |
| `ballooning.js` | extraction engine — PDF text → classified characteristics → exporters |
| `ballooning-ui.js` | viewer, balloon overlay, editable table, persistence |
| `ballooning.css` | styles (scoped `.bl-*`) |

---

*For further assistance, please refer to the documentation or contact support.*