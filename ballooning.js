/* ballooning.js
 *
 * Drawing Ballooning Engine — reads an engineering drawing PDF, extracts every
 * dimension / GD&T feature control frame / note / material + coating spec,
 * highlights each one, assigns a sequential characteristic number, and exports
 * an AS9102 Form 3 (FAIR) or PPAP dimensional-results sheet (ISIR).
 *
 * Pipeline:
 *   PdfLoader   -> pages[] with positioned text items
 *   LineBuilder -> merges text items into logical annotation strings
 *   Classifier  -> tags each string as DIM / GDT / THREAD / MATERIAL / COATING ...
 *   Numberer    -> assigns characteristic numbers in drawing reading order
 *   Renderer    -> canvas + balloon/highlight overlay
 *   Exporters   -> AS9102 Form 3 CSV, PPAP CSV, project JSON, ballooned PNG
 *
 * Coordinates are stored normalized (0..1 of page width/height) so zoom and
 * export never invalidate a balloon position.
 */

(function (global) {
  'use strict';

  // ---------------------------------------------------------------------------
  // Symbol tables
  // ---------------------------------------------------------------------------

  // Unicode GD&T characteristic symbols, as emitted by modern CAD PDF writers.
  var GDT_SYMBOLS = {
    '⌖': 'Position',
    '◎': 'Concentricity',
    '⌯': 'Symmetry',
    '⏤': 'Straightness',
    '⏥': 'Flatness',
    '○': 'Circularity',
    '⌭': 'Cylindricity',
    '⌒': 'Profile of a Line',
    '⌓': 'Profile of a Surface',
    '∠': 'Angularity',
    '⊥': 'Perpendicularity',
    '∥': 'Parallelism',
    '↗': 'Circular Runout',
    '⌰': 'Total Runout'
  };

  var MODIFIERS = {
    'Ⓜ': 'MMC',   // circled M
    'Ⓛ': 'LMC',   // circled L
    'Ⓟ': 'Projected Tolerance Zone',
    'Ⓕ': 'Free State',
    'Ⓣ': 'Tangent Plane',
    'Ⓢ': 'Statistical Tolerance',
    'Ⓤ': 'Unequally Disposed Profile'
  };

  var FEATURE_SYMBOLS = {
    '⌀': 'Diameter',      // ⌀
    'Ø': 'Diameter',      // Ø
    '⌴': 'Counterbore',   // ⌴
    '⌵': 'Countersink',   // ⌵
    '↧': 'Depth',         // ↧
    '□': 'Square',
    '⌲': 'Conical Taper',
    '⌳': 'Slope'
  };

  /* Legacy CAD symbol fonts (gdt.shx, AMGDT, ISOCPEUR variants) do NOT emit the
   * Unicode above — they map GD&T glyphs onto plain ASCII letters, so a position
   * callout extracts as garbage like "j 0.25 m A B C". This map is applied ONLY
   * when the text item's font name matches SYMBOL_FONT_RE, and it MUST be
   * calibrated against a known-good drawing before you trust it — vendors ship
   * different mappings under the same font name. The UI exposes it for editing;
   * an uncalibrated mapping is why every such frame is flagged REVIEW. */
  var SYMBOL_FONT_RE = /(gdt|amgdt|isocp|romant|simplex|txt)\b/i;
  var SYMBOL_FONT_MAP = {
    a: '∠', b: '⊥', c: '⏥', d: '⌓', e: '⌒',
    f: '∥', g: '↗', h: '⌭', i: '○', j: '⌖',
    k: '⌯', l: '⏤', m: 'Ⓜ', n: '⌰', u: '⌀',
    r: 'Ⓛ', p: 'Ⓟ'
  };

  // ---------------------------------------------------------------------------
  // Spec pattern tables — material and coating/finish callouts
  // ---------------------------------------------------------------------------

  var MATERIAL_PATTERNS = [
    { re: /\bAMS\s?-?\s?(\d{4}[A-Z]?)\b/i,                     label: 'AMS material spec' },
    { re: /\bASTM\s+([A-Z]\s?\d{1,4}[A-Z]?(?:\/[A-Z]\s?\d+[A-Z]?)?)\b/i, label: 'ASTM material spec' },
    { re: /\bMIL-(?:DTL|S|T|A|P|C)-\d{3,5}[A-Z]?\b/i,          label: 'MIL material spec' },
    { re: /\bQQ-[APS]-\d{2,4}[A-Z]?\b/i,                       label: 'Federal material spec' },
    { re: /\bUNS\s?[A-Z]\d{5}\b/i,                             label: 'UNS designation' },
    { re: /\bAISI\s?\d{3,4}[A-Z]?\b/i,                         label: 'AISI grade' },
    { re: /\bSAE\s?\d{3,4}\b/i,                                label: 'SAE grade' },
    { re: /\b(?:2024|6061|7075|5052|6063|2219|7050)\s?-\s?[TO]\d{1,3}\b/i, label: 'Aluminum alloy/temper' },
    { re: /\bTi\s?-\s?6Al\s?-\s?4V(?:\s?ELI)?\b/i,             label: 'Titanium alloy' },
    { re: /\b(?:INCONEL|HASTELLOY|MONEL|WASPALOY|NITRONIC)\s?-?\s?\w*\b/i, label: 'Superalloy' },
    { re: /\b1[57]-[47]\s?PH\b/i,                              label: 'PH stainless' },
    { re: /\b(?:304|316|321|347|410|416|420|440C|430)\s?L?\s?(?:SST|SS|STAINLESS)?\b/i, label: 'Stainless grade' },
    { re: /\bEN\s?1\.\d{4}\b/i,                                label: 'EN material number' },
    { re: /\bDIN\s?\d{3,5}\b/i,                                label: 'DIN material spec' },
    { re: /\bPEEK|ULTEM|DELRIN|ACETAL|NYLON\s?6\/?6?|PTFE|TEFLON\b/i, label: 'Engineering polymer' },
    { re: /\bMATERIAL\s*[:.\-]/i,                              label: 'Material note' }
  ];

  var COATING_PATTERNS = [
    { re: /\bMIL-A-8625\b.*?(?:TYPE\s?[IVX]+)?(?:.*?CLASS\s?\d)?/i, label: 'Anodize (MIL-A-8625)' },
    { re: /\bMIL-DTL-5541\b.*?(?:TYPE\s?[IVX]+)?(?:.*?CLASS\s?\d[A-Z]?)?/i, label: 'Chem film (MIL-DTL-5541)' },
    { re: /\bALODINE|IRIDITE|CHEM\s?FILM|CONVERSION\s?COAT/i,   label: 'Chemical conversion coating' },
    { re: /\bASTM\s?B633\b.*?(?:TYPE\s?[IVX]+)?/i,             label: 'Zinc plating (ASTM B633)' },
    { re: /\bASTM\s?B733\b|\bAMS\s?2404\b|\bELECTROLESS\s?NICKEL\b/i, label: 'Electroless nickel' },
    { re: /\bAMS\s?2403\b|\bELECTROLYTIC\s?NICKEL\b/i,         label: 'Nickel plating' },
    { re: /\bAMS\s?2700\b|\bASTM\s?A967\b|\bQQ-P-35\b|\bPASSIVAT\w+/i, label: 'Passivation' },
    { re: /\bAMS-QQ-P-416\b|\bCADMIUM\s?PLAT\w+/i,             label: 'Cadmium plating' },
    { re: /\bMIL-DTL-13924\b|\bBLACK\s?OXIDE\b/i,              label: 'Black oxide' },
    { re: /\bMIL-PRF-23377\b|\bMIL-PRF-85285\b|\bMIL-PRF-22750\b/i, label: 'Aerospace paint system' },
    { re: /\bMIL-STD-171\b/i,                                  label: 'Finish standard' },
    { re: /\bISO\s?2081\b|\bISO\s?4042\b|\bISO\s?10683\b/i,    label: 'ISO plating spec' },
    { re: /\bPOWDER\s?COAT\w*|\bE-?COAT\w*|\bELECTRO\s?DEPOSIT\w*/i, label: 'Organic coating' },
    { re: /\bHARD\s?ANODIZE|\bTYPE\s?III\s?ANODIZE/i,          label: 'Hardcoat anodize' },
    { re: /\bDRY\s?FILM\s?LUBE|\bMIL-PRF-46010\b|\bMOLY\s?COAT/i, label: 'Dry film lubricant' },
    { re: /\bFINISH\s*[:.\-]|\bCOATING\s*[:.\-]|\bPLATE\s*[:.\-]/i, label: 'Finish note' }
  ];

  var PROCESS_PATTERNS = [
    { re: /\bHEAT\s?TREAT\w*|\bAMS\s?2759\b|\bAMS-H-6875\b/i,  label: 'Heat treatment' },
    { re: /\bHRC\s?\d{2}(?:\s?-\s?\d{2})?|\bROCKWELL\b/i,      label: 'Hardness requirement' },
    { re: /\bSHOT\s?PEEN\w*|\bAMS\s?2430\b/i,                  label: 'Shot peening' },
    { re: /\bFPI\b|\bPENETRANT\s?INSPECT\w*|\bASTM\s?E1417\b/i, label: 'Penetrant inspection' },
    { re: /\bDEBURR\w*|\bBREAK\s?(?:ALL\s?)?SHARP\s?EDGES?\b/i, label: 'Edge condition' },
    { re: /\bWELD\w*\s?PER\b|\bAWS\s?D\d+\.\d+\b|\bAMS-STD-1595\b/i, label: 'Welding spec' }
  ];

  // ---------------------------------------------------------------------------
  // Dimension patterns
  // ---------------------------------------------------------------------------

  var NUM = '[-+]?\\d+(?:\\.\\d+)?';
  var UNUM = '\\d+(?:\\.\\d+)?';

  /* Order matters and is load-bearing. The tolerance forms must be tested before
   * the bare-value forms, otherwise "⌀12.7 +0.03/-0.01" is torn apart by the
   * limit-dimension pattern into a meaningless 0.03/-0.01 pair. Likewise CHAMFER
   * precedes ANGLE so "1.5 X 45°" is not filed as a plain angular dimension. */
  var DIM_PATTERNS = [
    { key: 'UNILATERAL', re: new RegExp('(' + NUM + ')\\s*\\+\\s*(' + UNUM + ')\\s*[/\\s]\\s*-\\s*(' + UNUM + ')'), desc: 'Unequal bilateral tolerance' },
    { key: 'BILATERAL',  re: new RegExp('(' + NUM + ')\\s*(?:\\u00B1|\\+/-|\\+-)\\s*(' + UNUM + ')'), desc: 'Bilateral tolerance' },
    { key: 'LIMIT',      re: new RegExp('(' + UNUM + ')\\s*[/\\u2013-]\\s*(' + UNUM + ')\\s*$'), desc: 'Limit dimension' },
    { key: 'REFERENCE',  re: new RegExp('^\\(\\s*(' + NUM + ')\\s*\\)$'), desc: 'Reference dimension' },
    { key: 'CHAMFER',    re: new RegExp('(' + UNUM + ')\\s*[Xx\\u00D7]\\s*(' + UNUM + ')\\s*\\u00B0'), desc: 'Chamfer' },
    { key: 'ANGLE',      re: new RegExp('(' + NUM + ')\\s*\\u00B0'), desc: 'Angular dimension' },
    { key: 'RADIUS',     re: new RegExp('\\bR\\s?(' + UNUM + ')'), desc: 'Radius' },
    { key: 'DIAMETER',   re: new RegExp('[\\u2300\\u00D8]\\s?(' + UNUM + ')'), desc: 'Diameter' },
    { key: 'LINEAR',     re: new RegExp('^\\s*(' + NUM + ')\\s*$'), desc: 'Linear dimension' }
  ];

  var THREAD_RE = /\b(?:M\d+(?:\.\d+)?(?:\s?[xX]\s?\d+(?:\.\d+)?)?(?:\s?-\s?\d[gGhH]\d?)?|\d+(?:\/\d+)?\s?-\s?\d+\s?UN[CFEJ]?[FS]?(?:\s?-\s?\d[AB])?|#\d+\s?-\s?\d+\s?UN[CF]|\d+\/\d+\s?-\s?\d+\s?NPT[F]?)\b/;
  var SURFACE_RE = /(?:\bRa\s?\d+(?:\.\d+)?|\b\d{1,3}\s?(?:µin|uin|RMS|AA)\b|\bRz\s?\d+(?:\.\d+)?|∇)/i;
  var QTY_RE = /^\s*(\d+)\s*[Xx×]\s+/;
  var CRITICAL_RE = /\b(?:CRITICAL|KPC|KEY\s?CHARACTERISTIC|SAFETY|FLIGHT\s?SAFETY|CC\b|SC\b)\b|[◆◇]/i;
  var BASIC_HINT_RE = /\bBASIC\b|\bBSC\b/i;
  var GENERAL_NOTE_RE = /UNLESS\s?OTHERWISE\s?(?:SPECIFIED|NOTED)|\bU\.?O\.?S\.?\b|\bNOTES?\s*:|\bTOLERANCES?\s?(?:ARE|:)/i;
  var TITLEBLOCK_RE = /\b(?:DRAWN|CHECKED|APPROVED|SCALE|SHEET|REV(?:ISION)?|DWG\s?NO|PART\s?(?:NO|NUMBER)|TITLE|DATE|SIZE|CAGE)\b/i;

  // ---------------------------------------------------------------------------
  // Utility
  // ---------------------------------------------------------------------------

  function clamp(v, lo, hi) { return v < lo ? lo : (v > hi ? hi : v); }
  function round(v, d) { var f = Math.pow(10, d == null ? 4 : d); return Math.round(v * f) / f; }
  function uid() { return 'c' + Math.random().toString(36).slice(2, 10); }

  function normalizeText(s) {
    return String(s || '')
      .replace(/ /g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  // Translate ASCII-mapped legacy symbol-font glyphs into real GD&T characters.
  function applySymbolFont(str, fontName, map) {
    if (!fontName || !SYMBOL_FONT_RE.test(fontName)) return { text: str, remapped: false };
    var out = '', hit = false;
    for (var i = 0; i < str.length; i++) {
      var ch = str[i], low = ch.toLowerCase();
      if (Object.prototype.hasOwnProperty.call(map, low)) { out += map[low]; hit = true; }
      else out += ch;
    }
    return { text: out, remapped: hit };
  }

  // ---------------------------------------------------------------------------
  // PdfLoader — render pages and pull positioned text
  // ---------------------------------------------------------------------------

  function PdfLoader(pdfjsLib) {
    this.pdfjsLib = pdfjsLib;
    this.doc = null;
    this.pages = [];
  }

  PdfLoader.prototype.load = function (arrayBuffer) {
    var self = this;
    return this.pdfjsLib.getDocument({ data: arrayBuffer }).promise.then(function (doc) {
      self.doc = doc;
      var jobs = [];
      for (var p = 1; p <= doc.numPages; p++) jobs.push(self._readPage(p));
      return Promise.all(jobs);
    }).then(function (pages) {
      self.pages = pages;
      return pages;
    });
  };

  PdfLoader.prototype._readPage = function (pageNumber) {
    var self = this;
    return this.doc.getPage(pageNumber).then(function (page) {
      var viewport = page.getViewport({ scale: 1 });
      return page.getTextContent().then(function (content) {
        var items = content.items.map(function (it) {
          var tx = self.pdfjsLib.Util.transform(viewport.transform, it.transform);
          var height = Math.hypot(tx[2], tx[3]) || it.height || 1;
          var width = it.width || (it.str.length * height * 0.5);
          var angle = Math.atan2(tx[1], tx[0]) * 180 / Math.PI;
          return {
            raw: it.str,
            fontName: it.fontName,
            x: tx[4],
            y: tx[5] - height,   // top-left of the glyph box
            w: width,
            h: height,
            angle: Math.round(angle / 15) * 15
          };
        }).filter(function (it) { return normalizeText(it.raw).length > 0; });

        return {
          pageNumber: pageNumber,
          width: viewport.width,
          height: viewport.height,
          page: page,
          items: items
        };
      });
    });
  };

  PdfLoader.prototype.renderPage = function (pageIndex, canvas, scale) {
    var pageData = this.pages[pageIndex];
    var viewport = pageData.page.getViewport({ scale: scale });
    canvas.width = Math.floor(viewport.width);
    canvas.height = Math.floor(viewport.height);
    return pageData.page.render({ canvasContext: canvas.getContext('2d'), viewport: viewport }).promise;
  };

  // ---------------------------------------------------------------------------
  // LineBuilder — merge glyph runs into logical annotations
  // ---------------------------------------------------------------------------

  function buildLines(pageData, symbolMap) {
    var buckets = {};

    pageData.items.forEach(function (it) {
      var conv = applySymbolFont(it.raw, it.fontName, symbolMap);
      var rad = it.angle * Math.PI / 180;
      // Project onto the axis perpendicular to the text direction so rotated
      // callouts (very common on drawings) still group into one line.
      var perp = -it.x * Math.sin(rad) + it.y * Math.cos(rad);
      var along = it.x * Math.cos(rad) + it.y * Math.sin(rad);
      var key = it.angle + '|' + Math.round(perp / Math.max(it.h * 0.7, 1));
      (buckets[key] = buckets[key] || []).push({
        item: it, text: conv.text, remapped: conv.remapped, along: along
      });
    });

    var lines = [];
    Object.keys(buckets).forEach(function (key) {
      var group = buckets[key].sort(function (a, b) { return a.along - b.along; });
      var run = null;

      group.forEach(function (entry) {
        var it = entry.item;
        var gap = run ? entry.along - (run.alongEnd) : 0;
        // ~1.8x text height tolerates the loose word spacing CAD writers emit in
        // note blocks ("ANODIZE PER MIL-A-8625") without bridging two separate
        // callouts, which on a drawing sit far further apart than that.
        if (run && gap <= it.h * 1.8) {
          run.text += (gap > it.h * 0.28 ? ' ' : '') + entry.text;
          run.alongEnd = entry.along + it.w;
          run.x0 = Math.min(run.x0, it.x);
          run.y0 = Math.min(run.y0, it.y);
          run.x1 = Math.max(run.x1, it.x + it.w);
          run.y1 = Math.max(run.y1, it.y + it.h);
          run.remapped = run.remapped || entry.remapped;
        } else {
          if (run) lines.push(run);
          run = {
            text: entry.text,
            angle: it.angle,
            alongEnd: entry.along + it.w,
            x0: it.x, y0: it.y, x1: it.x + it.w, y1: it.y + it.h,
            remapped: entry.remapped,
            fontName: it.fontName
          };
        }
      });
      if (run) lines.push(run);
    });

    return lines.map(function (l) {
      l.text = normalizeText(l.text);
      return l;
    }).filter(function (l) { return l.text.length > 0; });
  }

  // ---------------------------------------------------------------------------
  // Classifier
  // ---------------------------------------------------------------------------

  function parseFeatureControlFrame(text) {
    var symbol = null, symbolChar = null;
    for (var ch in GDT_SYMBOLS) {
      if (text.indexOf(ch) !== -1) { symbol = GDT_SYMBOLS[ch]; symbolChar = ch; break; }
    }
    if (!symbol) return null;

    var after = text.slice(text.indexOf(symbolChar) + 1);
    var tolMatch = after.match(new RegExp('[\\u2300\\u00D8]?\\s*(' + NUM + ')'));
    var mods = [];
    for (var m in MODIFIERS) if (text.indexOf(m) !== -1) mods.push(MODIFIERS[m]);

    // Datum references: trailing single/double capital letters, optionally with
    // a material-condition modifier, separated by spaces or pipes.
    var datums = [];
    var datumScan = after.replace(new RegExp('[\\u2300\\u00D8]?\\s*' + NUM), ' ');
    var dm = datumScan.match(/\b([A-Z]{1,2})\b/g);
    if (dm) datums = dm.filter(function (d) { return d !== 'X'; });

    return {
      symbol: symbol,
      tolerance: tolMatch ? parseFloat(tolMatch[1]) : null,
      diametral: /[⌀Ø]/.test(after.slice(0, (tolMatch ? after.indexOf(tolMatch[0]) : 0) + 2)),
      modifiers: mods,
      datums: datums
    };
  }

  function parseDimension(text) {
    var qty = 1;
    var qm = text.match(QTY_RE);
    var body = text;
    if (qm) { qty = parseInt(qm[1], 10); body = text.replace(QTY_RE, ''); }

    for (var i = 0; i < DIM_PATTERNS.length; i++) {
      var p = DIM_PATTERNS[i];
      var m = body.match(p.re);
      if (!m) continue;

      var nominal = null, upper = null, lower = null;
      if (p.key === 'LIMIT') {
        var a = parseFloat(m[1]), b = parseFloat(m[2]);
        var hi = Math.max(a, b), lo = Math.min(a, b);
        nominal = round((hi + lo) / 2, 5); upper = round(hi - nominal, 5); lower = round(lo - nominal, 5);
      } else if (p.key === 'BILATERAL') {
        nominal = parseFloat(m[1]); upper = Math.abs(parseFloat(m[2])); lower = -upper;
      } else if (p.key === 'UNILATERAL') {
        nominal = parseFloat(m[1]); upper = Math.abs(parseFloat(m[2])); lower = -Math.abs(parseFloat(m[3]));
      } else {
        nominal = parseFloat(m[1]);
      }

      return {
        kind: p.key,
        description: p.desc,
        quantity: qty,
        nominal: nominal,
        upper: upper,
        lower: lower,
        toleranced: upper !== null
      };
    }
    return null;
  }

  function matchSpecTable(text, table) {
    for (var i = 0; i < table.length; i++) {
      var m = text.match(table[i].re);
      if (m) return { label: table[i].label, match: normalizeText(m[0]) };
    }
    return null;
  }

  /* Returns a characteristic object, or null when the line is not an inspectable
   * requirement (title-block furniture, view labels, sheet numbers). */
  function classify(line, opts) {
    var text = line.text;
    if (text.length < 1) return null;

    var base = {
      id: uid(),
      text: text,
      remapped: !!line.remapped,
      flags: [],
      review: false,
      critical: CRITICAL_RE.test(text)
    };
    if (base.critical) base.flags.push('Key characteristic keyword');
    if (base.remapped) {
      base.review = true;
      base.flags.push('Legacy symbol font remapped — verify against print');
    }

    // 1. Feature control frame (highest value, check before plain numbers).
    var fcf = parseFeatureControlFrame(text);
    if (fcf) {
      base.type = 'GD&T';
      base.subtype = fcf.symbol;
      base.nominal = 0;
      base.upper = fcf.tolerance;
      base.lower = 0;
      base.unit = opts.unit;
      base.requirement = text;
      base.datums = fcf.datums.join(', ');
      base.modifiers = fcf.modifiers.join(', ');
      base.method = fcf.tolerance != null && fcf.tolerance < 0.05 ? 'CMM' : 'CMM / functional gage';
      if (fcf.tolerance == null) { base.review = true; base.flags.push('Tolerance value not resolved from frame'); }
      return base;
    }

    /* 2. Coating / process / material specs — checked before dimensions so a spec
     *    number is never mistaken for a size. Coating and process come first
     *    because they are the more specific match: "ANODIZE PER MIL-A-8625" and
     *    "HEAT TREAT PER AMS 2759" both also satisfy the generic MIL-/AMS-
     *    material patterns, and filing them as material would put the wrong
     *    verification method on the FAIR. */
    var coat = matchSpecTable(text, COATING_PATTERNS);
    if (coat) {
      base.type = 'Coating';
      base.subtype = coat.label;
      base.requirement = text;
      base.method = 'C of C / plating cert, thickness check per spec';
      return base;
    }
    var proc = matchSpecTable(text, PROCESS_PATTERNS);
    if (proc) {
      base.type = 'Process';
      base.subtype = proc.label;
      base.requirement = text;
      base.method = 'Process cert / visual';
      return base;
    }
    var mat = matchSpecTable(text, MATERIAL_PATTERNS);
    if (mat) {
      base.type = 'Material';
      base.subtype = mat.label;
      base.requirement = text;
      base.method = 'Certificate of conformance / material cert review';
      return base;
    }

    // 3. Threads.
    var thr = text.match(THREAD_RE);
    if (thr) {
      base.type = 'Thread';
      base.subtype = normalizeText(thr[0]);
      base.requirement = text;
      base.method = 'Go/No-Go thread gage';
      return base;
    }

    // 4. Surface finish.
    if (SURFACE_RE.test(text)) {
      base.type = 'Surface Finish';
      base.subtype = 'Roughness';
      base.requirement = text;
      base.method = 'Profilometer / comparator';
      return base;
    }

    // 5. General notes and tolerance blocks — recorded, not measured.
    if (GENERAL_NOTE_RE.test(text)) {
      base.type = 'Note';
      base.subtype = 'General note';
      base.requirement = text;
      base.method = 'Review';
      return base;
    }

    // 6. Dimensions.
    var dim = parseDimension(text);
    if (dim) {
      // A bare integer inside title-block furniture is a sheet number, not a size.
      if (dim.kind === 'LINEAR' && TITLEBLOCK_RE.test(text)) return null;

      base.type = 'Dimension';
      // A toleranced diameter matches BILATERAL/LIMIT before DIAMETER, so record
      // the feature symbol separately rather than losing it to the tolerance form.
      base.diametral = /[⌀Ø]/.test(text);
      base.subtype = (base.diametral && dim.kind !== 'DIAMETER' ? 'Diameter, ' : '') + dim.description;
      base.quantity = dim.quantity;
      base.nominal = dim.nominal;
      base.unit = opts.unit;
      base.requirement = text;

      if (dim.toleranced) {
        base.upper = dim.upper;
        base.lower = dim.lower;
        base.toleranceSource = 'Stated on drawing';
      } else if (dim.kind === 'REFERENCE') {
        base.upper = null; base.lower = null;
        base.toleranceSource = 'Reference — not inspected';
        base.method = 'Reference only';
      } else {
        var def = defaultTolerance(dim.nominal, dim.kind, opts.generalTol);
        base.upper = def.upper;
        base.lower = def.lower;
        base.toleranceSource = 'Applied from general tolerance block';
        base.review = true;
        base.flags.push('Tolerance inherited from title block — confirm, and confirm the dimension is not BASIC');
      }

      if (!base.method) base.method = suggestMethod(dim, base.upper);
      if (BASIC_HINT_RE.test(text)) {
        base.toleranceSource = 'Basic dimension';
        base.upper = null; base.lower = null;
        base.method = 'Controlled by associated feature control frame';
      }
      return base;
    }

    return null;
  }

  function defaultTolerance(nominal, kind, gt) {
    if (kind === 'ANGLE') return { upper: gt.angular, lower: -gt.angular };
    var s = String(nominal);
    var dot = s.indexOf('.');
    var places = dot === -1 ? 0 : s.length - dot - 1;
    var t = places >= 3 ? gt.three : places === 2 ? gt.two : gt.one;
    return { upper: t, lower: -t };
  }

  function suggestMethod(dim, upper) {
    var band = upper == null ? null : Math.abs(upper) * 2;
    if (dim.kind === 'ANGLE' || dim.kind === 'CHAMFER') return 'Optical comparator / protractor';
    if (dim.kind === 'RADIUS') return 'Radius gage / comparator';
    if (dim.kind === 'DIAMETER') {
      if (band != null && band <= 0.02) return 'Air gage / bore gage';
      return 'Micrometer / pin gage';
    }
    if (band != null && band <= 0.02) return 'CMM';
    if (band != null && band <= 0.1) return 'Micrometer';
    return 'Caliper';
  }

  // ---------------------------------------------------------------------------
  // Numberer — characteristic numbers in drawing reading order
  // ---------------------------------------------------------------------------

  /* AS9102 requires each characteristic to carry a unique number that is stable
   * across the ballooned print and Form 3. Order is left-to-right within
   * horizontal bands, top to bottom, page by page — matching how an inspector
   * reads the sheet. Band height is a fraction of page height so nearly-aligned
   * callouts do not shuffle. */
  function assignNumbers(characteristics, bandFraction) {
    var band = bandFraction || 0.04;
    characteristics.sort(function (a, b) {
      if (a.pageIndex !== b.pageIndex) return a.pageIndex - b.pageIndex;
      var ba = Math.floor(a.ny / band), bb = Math.floor(b.ny / band);
      if (ba !== bb) return ba - bb;
      return a.nx - b.nx;
    });
    characteristics.forEach(function (c, i) { c.number = i + 1; });
    return characteristics;
  }

  /* Drawing zone (e.g. "C3") from the standard border grid: letters bottom-to-top
   * on the vertical edge, numbers right-to-left on the horizontal edge. */
  function computeZone(nx, ny, rows, cols) {
    if (!rows || !cols) return '';
    var letters = 'ABCDEFGHJKLMNPRSTUV';
    var r = clamp(Math.floor((1 - ny) * rows), 0, rows - 1);
    var c = clamp(Math.floor((1 - nx) * cols), 0, cols - 1);
    return letters[r] + (c + 1);
  }

  // ---------------------------------------------------------------------------
  // Extractor — pipeline entry point
  // ---------------------------------------------------------------------------

  function extract(pages, opts) {
    opts = opts || {};
    var options = {
      unit: opts.unit || 'mm',
      generalTol: opts.generalTol || { one: 0.5, two: 0.25, three: 0.1, angular: 0.5 },
      symbolMap: opts.symbolMap || SYMBOL_FONT_MAP,
      zoneRows: opts.zoneRows || 0,
      zoneCols: opts.zoneCols || 0,
      includeNotes: opts.includeNotes !== false
    };

    var out = [];
    pages.forEach(function (pageData, pageIndex) {
      var lines = buildLines(pageData, options.symbolMap);
      lines.forEach(function (line) {
        var c = classify(line, options);
        if (!c) return;
        if (!options.includeNotes && c.type === 'Note') return;

        var cx = (line.x0 + line.x1) / 2;
        var cy = (line.y0 + line.y1) / 2;
        c.pageIndex = pageIndex;
        c.pageNumber = pageData.pageNumber;
        c.nx = clamp(cx / pageData.width, 0, 1);
        c.ny = clamp(cy / pageData.height, 0, 1);
        c.box = {
          x: line.x0 / pageData.width,
          y: line.y0 / pageData.height,
          w: (line.x1 - line.x0) / pageData.width,
          h: (line.y1 - line.y0) / pageData.height
        };
        // Balloon parks up-left of the callout with a leader back to it; the
        // operator can drag it anywhere and the leader follows.
        c.balloon = {
          x: clamp(c.box.x - 0.022, 0.008, 0.99),
          y: clamp(c.box.y - 0.018, 0.008, 0.99)
        };
        c.zone = computeZone(c.nx, c.ny, options.zoneRows, options.zoneCols);
        c.source = 'auto';
        out.push(c);
      });
    });

    return assignNumbers(out);
  }

  // ---------------------------------------------------------------------------
  // Exporters
  // ---------------------------------------------------------------------------

  function csvCell(v) {
    var s = v == null ? '' : String(v);
    return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  }

  function toCsv(rows) {
    return rows.map(function (r) { return r.map(csvCell).join(','); }).join('\r\n');
  }

  function requirementString(c) {
    if (c.type === 'Material' || c.type === 'Coating' || c.type === 'Process' || c.type === 'Note') return c.requirement;
    if (c.type === 'GD&T') {
      var s = c.subtype + ' ' + (c.upper != null ? c.upper : '?');
      if (c.modifiers) s += ' (' + c.modifiers + ')';
      if (c.datums) s += ' | ' + c.datums;
      return s;
    }
    if (c.upper == null) return c.requirement;
    var lo = round(c.nominal + c.lower, 5), hi = round(c.nominal + c.upper, 5);
    return c.nominal + ' (' + lo + ' / ' + hi + ') ' + (c.unit || '');
  }

  /* AS9102 Rev C Form 3 — Characteristic Accountability, Verification and
   * Compatibility Evaluation. Measurement columns are intentionally blank: they
   * are filled by inspection, not by extraction. */
  function exportAS9102Form3(chars, meta) {
    var rows = [
      ['AS9102 Rev C - Form 3: Characteristic Accountability, Verification and Compatibility Evaluation'],
      ['Part Number', meta.partNumber || '', 'Part Name', meta.partName || '', 'Serial Number', meta.serialNumber || ''],
      ['Drawing Number', meta.drawingNumber || '', 'Drawing Revision', meta.revision || '', 'FAIR Number', meta.fairNumber || ''],
      [],
      ['Char. No.', 'Reference Location', 'Characteristic Designator', 'Requirement',
       'Results', 'Designed Tooling', 'Non-Conformance Number', 'Notes']
    ];
    chars.forEach(function (c) {
      rows.push([
        c.number,
        'Sheet ' + c.pageNumber + (c.zone ? ', Zone ' + c.zone : ''),
        c.critical ? 'KEY' : '',
        requirementString(c),
        '',
        '',
        '',
        [c.type + ' / ' + (c.subtype || ''), c.method ? 'Method: ' + c.method : '',
         c.toleranceSource || '', c.flags.join('; ')].filter(Boolean).join(' | ')
      ]);
    });
    return toCsv(rows);
  }

  /* PPAP / ISIR dimensional results sheet — five measurement columns is the
   * common default for a 5-piece capability sample. */
  function exportPPAP(chars, meta, sampleSize) {
    var n = sampleSize || 5;
    var header = ['Item No.', 'Sheet', 'Zone', 'Characteristic', 'Specification',
                  'Nominal', 'Lower Limit', 'Upper Limit', 'Unit', 'Inspection Method'];
    for (var i = 1; i <= n; i++) header.push('Sample ' + i);
    header.push('Result (OK/NG)', 'Comments');

    var rows = [
      ['PPAP / ISIR Dimensional Results'],
      ['Part Number', meta.partNumber || '', 'Drawing Rev', meta.revision || '', 'Supplier', meta.supplier || ''],
      [],
      header
    ];

    chars.forEach(function (c) {
      var lo = (c.upper == null || c.nominal == null) ? '' : round(c.nominal + c.lower, 5);
      var hi = (c.upper == null || c.nominal == null) ? '' : round(c.nominal + c.upper, 5);
      var row = [
        c.number, c.pageNumber, c.zone || '',
        c.type + (c.subtype ? ' - ' + c.subtype : ''),
        requirementString(c),
        c.nominal == null ? '' : c.nominal,
        lo, hi, c.unit || '', c.method || ''
      ];
      for (var j = 0; j < n; j++) row.push('');
      row.push('', c.flags.join('; '));
      rows.push(row);
    });
    return toCsv(rows);
  }

  function exportJson(chars, meta) {
    return JSON.stringify({ version: 1, meta: meta, characteristics: chars }, null, 2);
  }

  // ---------------------------------------------------------------------------
  // Summary for the coverage panel
  // ---------------------------------------------------------------------------

  function summarize(chars) {
    var byType = {}, review = 0, critical = 0;
    chars.forEach(function (c) {
      byType[c.type] = (byType[c.type] || 0) + 1;
      if (c.review) review++;
      if (c.critical) critical++;
    });
    return { total: chars.length, byType: byType, review: review, critical: critical };
  }

  global.Ballooning = {
    PdfLoader: PdfLoader,
    buildLines: buildLines,
    classify: classify,
    extract: extract,
    assignNumbers: assignNumbers,
    computeZone: computeZone,
    requirementString: requirementString,
    exportAS9102Form3: exportAS9102Form3,
    exportPPAP: exportPPAP,
    exportJson: exportJson,
    summarize: summarize,
    uid: uid,
    round: round,
    GDT_SYMBOLS: GDT_SYMBOLS,
    MODIFIERS: MODIFIERS,
    FEATURE_SYMBOLS: FEATURE_SYMBOLS,
    SYMBOL_FONT_MAP: SYMBOL_FONT_MAP
  };

})(typeof window !== 'undefined' ? window : this);
