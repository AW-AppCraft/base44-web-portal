/* vision.js — optional vision-model extraction pass via a local Ollama server.
 *
 * Fills the gap the text-layer parser cannot: scanned drawings, outlined text,
 * and legacy symbol fonts. The page talks to Ollama directly from the browser,
 * so the drawing never leaves the operator's machine.
 *
 * Division of labour, deliberately:
 *   the model  -> perception only. Reads glyphs, returns strings + rough boxes.
 *   ballooning.js classify() -> all semantics. Parses tolerances, picks methods.
 *
 * The model is never asked to compute a limit, decide a tolerance or judge
 * criticality. Its numbers are transcription, not arithmetic. Everything it
 * produces is flagged for review, without exception.
 */

(function (global) {
  'use strict';

  var B = global.Ballooning;

  // ---------------------------------------------------------------------------
  // Ollama client
  // ---------------------------------------------------------------------------

  function OllamaClient(endpoint) {
    this.endpoint = (endpoint || 'http://localhost:11434').replace(/\/+$/, '');
  }

  /* Returns [{ name, capabilities, vision, thinking, capsKnown }].
   *
   * Capability comes from Ollama itself (/api/show reports a capabilities array),
   * not from matching model names. Name matching cannot work: any hardcoded list
   * is wrong the moment a new model ships, and it silently mislabels the new one
   * as text-only — exactly the failure this replaced. The name pattern survives
   * only as a fallback for Ollama versions predating the capabilities field. */
  OllamaClient.prototype.listModels = function () {
    var self = this;
    return fetch(this.endpoint + '/api/tags')
      .then(function (r) {
        if (!r.ok) throw new Error('Ollama returned HTTP ' + r.status);
        return r.json();
      })
      .then(function (j) {
        var names = (j.models || []).map(function (m) { return m.name; });
        return Promise.all(names.map(function (n) {
          return self.showModel(n).then(function (caps) {
            var known = caps.length > 0;
            return {
              name: n,
              capabilities: caps,
              capsKnown: known,
              vision: known ? caps.indexOf('vision') !== -1 : VISION_NAME_RE.test(n),
              thinking: caps.indexOf('thinking') !== -1
            };
          });
        }));
      })
      .catch(function (err) {
        // A browser CORS rejection surfaces as an opaque TypeError; say what it
        // actually means rather than leaking "Failed to fetch" to the operator.
        throw new Error(describeFetchError(err));
      });
  };

  OllamaClient.prototype.showModel = function (name) {
    return fetch(this.endpoint + '/api/show', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ model: name })
    })
      .then(function (r) { return r.ok ? r.json() : null; })
      .then(function (j) { return (j && Array.isArray(j.capabilities)) ? j.capabilities : []; })
      // An older Ollama has no /api/show capabilities; degrade to the name test
      // rather than failing the whole listing.
      .catch(function () { return []; });
  };

  // Fallback only — used when Ollama does not report capabilities.
  var VISION_NAME_RE = /gemma[3-9]\d*|llava|llama3\.2-vision|minicpm-v|qwen2?\.?5?-?vl|moondream|bakllava|pixtral|granite3\.2-vision|mistral-small3/i;

  OllamaClient.prototype.vision = function (model, prompt, imageBase64, signal, thinking) {
    var body = {
      model: model,
      prompt: prompt,
      images: [imageBase64],
      stream: false,
      format: 'json',
      options: { temperature: 0, num_predict: 2048 }
    };

    /* Transcription is not a reasoning task, and a thinking model deliberating
     * over every tile turns a dozen regions into a long wait for no accuracy
     * gain. Only sent when the model actually reports the capability — older
     * Ollama builds reject the field outright on models that lack it. */
    if (thinking) body.think = false;

    return fetch(this.endpoint + '/api/generate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      signal: signal,
      body: JSON.stringify(body)
    }).then(function (r) {
      if (!r.ok) {
        return r.text().then(function (t) {
          throw new Error('Ollama HTTP ' + r.status + ': ' + t.slice(0, 300));
        });
      }
      return r.json();
    }).then(function (j) {
      return j.response || '';
    }).catch(function (err) {
      if (err.name === 'AbortError') throw err;
      throw new Error(describeFetchError(err));
    });
  };

  function describeFetchError(err) {
    if (err && err.message && /Failed to fetch|NetworkError|Load failed/i.test(err.message)) {
      return 'Cannot reach Ollama. Check that it is running (ollama serve), and that this ' +
             'page\'s origin is allowed — Ollama blocks cross-origin browser requests by ' +
             'default. Start it with OLLAMA_ORIGINS set to this page\'s origin (or "*").';
    }
    return err && err.message ? err.message : String(err);
  }

  // ---------------------------------------------------------------------------
  // Prompt
  // ---------------------------------------------------------------------------

  /* Kept blunt and closed-ended. Open-ended "describe this drawing" prompts make
   * small vision models narrate geometry and invent plausible-looking dimensions;
   * asking only for glyphs actually present keeps the failure mode at "missed a
   * callout" rather than "fabricated a callout". */
  var PROMPT = [
    'You are reading a region of a mechanical engineering drawing.',
    '',
    'List every ANNOTATION you can actually read in this image: dimensions,',
    'tolerances, GD&T feature control frames, thread callouts, surface finish',
    'symbols, material specifications, coating and plating specifications,',
    'heat treat callouts, and general notes.',
    '',
    'Rules:',
    '- Transcribe the text EXACTLY as printed, character for character.',
    '- Use the real symbols where present: diameter as ⌀, plus/minus as ±,',
    '  degrees as °, and GD&T symbols such as ⌖ ⏥ ⏤ ⊥ ∥ ∠ ◎ ○ ⌭ ⌒ ⌓ ⌯ ↗.',
    '- Do NOT calculate anything. Do NOT expand a tolerance into limits.',
    '- Do NOT guess at text that is blurred or cut off. Omit it instead.',
    '- Ignore the title block, border zone letters/numbers, revision tables,',
    '  view labels such as "SECTION A-A", and the company name.',
    '- bbox is [x0,y0,x1,y1] as fractions of THIS image, 0 to 1, top-left origin.',
    '',
    'Return ONLY JSON in exactly this form:',
    '{"characteristics":[{"text":"25.40 ±0.05","bbox":[0.10,0.22,0.18,0.25]}]}',
    '',
    'If you can read no annotations, return {"characteristics":[]}.'
  ].join('\n');

  // ---------------------------------------------------------------------------
  // Tiling
  // ---------------------------------------------------------------------------

  /* A D-size drawing scaled to fit a 12B model's image budget renders dimension
   * text at a handful of pixels tall — unreadable. Tiling with overlap keeps the
   * effective resolution high; the overlap stops a callout landing on a seam from
   * being cut in half, and the dedup pass downstream removes what it double-counts. */
  function tilePage(canvas, rows, cols, overlapFraction) {
    var overlap = overlapFraction == null ? 0.12 : overlapFraction;
    var tiles = [];
    var tw = canvas.width / cols;
    var th = canvas.height / rows;
    var ox = tw * overlap;
    var oy = th * overlap;

    for (var r = 0; r < rows; r++) {
      for (var c = 0; c < cols; c++) {
        var x0 = Math.max(0, c * tw - ox);
        var y0 = Math.max(0, r * th - oy);
        var x1 = Math.min(canvas.width, (c + 1) * tw + ox);
        var y1 = Math.min(canvas.height, (r + 1) * th + oy);

        var t = document.createElement('canvas');
        t.width = Math.round(x1 - x0);
        t.height = Math.round(y1 - y0);
        t.getContext('2d').drawImage(canvas, x0, y0, t.width, t.height, 0, 0, t.width, t.height);

        tiles.push({
          canvas: t,
          row: r, col: c,
          // Placement of this tile within the full page, normalized.
          fx: x0 / canvas.width,
          fy: y0 / canvas.height,
          fw: (x1 - x0) / canvas.width,
          fh: (y1 - y0) / canvas.height
        });
      }
    }
    return tiles;
  }

  function toBase64(canvas) {
    return canvas.toDataURL('image/png').split(',')[1];
  }

  // ---------------------------------------------------------------------------
  // Response parsing
  // ---------------------------------------------------------------------------

  function parseResponse(raw) {
    if (!raw) return [];
    var text = String(raw).trim();

    // format:'json' usually yields clean JSON, but models still occasionally wrap
    // it in prose or a code fence. Recover the outermost object rather than
    // discarding a tile's whole result over punctuation.
    var parsed = null;
    try {
      parsed = JSON.parse(text);
    } catch (e) {
      var start = text.indexOf('{');
      var end = text.lastIndexOf('}');
      if (start !== -1 && end > start) {
        try { parsed = JSON.parse(text.slice(start, end + 1)); } catch (e2) { return []; }
      }
    }
    if (!parsed) return [];

    var list = parsed.characteristics || parsed.annotations || parsed.items;
    if (!Array.isArray(list)) return [];

    return list.map(function (o) {
      if (!o) return null;
      var t = typeof o === 'string' ? o : (o.text || o.value || o.annotation);
      if (!t || !String(t).trim()) return null;
      var bb = Array.isArray(o.bbox) && o.bbox.length === 4 ? o.bbox.map(Number) : null;
      if (bb && bb.some(function (n) { return !isFinite(n); })) bb = null;
      return { text: String(t).trim(), bbox: bb };
    }).filter(Boolean);
  }

  // ---------------------------------------------------------------------------
  // Dedup
  // ---------------------------------------------------------------------------

  /* Collapses the ways one callout can be spelled differently by the text layer
   * and by the model: the two diameter code points, the several dash and
   * plus-minus glyphs, and OCR's classic O/zero confusion. */
  function normText(s) {
    return String(s)
      .toUpperCase()
      .replace(/[⌀Ø]/g, 'D')
      .replace(/[−–—]/g, '-')
      .replace(/\+\/-|\+-/g, '±')
      .replace(/[O]/g, '0')
      .replace(/[\s .,;:|]/g, '');
  }

  function iou(a, b) {
    if (!a || !b) return 0;
    var x0 = Math.max(a.x, b.x), y0 = Math.max(a.y, b.y);
    var x1 = Math.min(a.x + a.w, b.x + b.w), y1 = Math.min(a.y + a.h, b.y + b.h);
    if (x1 <= x0 || y1 <= y0) return 0;
    var inter = (x1 - x0) * (y1 - y0);
    return inter / (a.w * a.h + b.w * b.h - inter);
  }

  /* Two sources of duplicates: tile overlap regions, and re-running vision over a
   * page whose text layer already parsed. Text-layer characteristics always win —
   * they are exact, the model's are transcription. */
  function isDuplicate(candidate, existing) {
    for (var i = 0; i < existing.length; i++) {
      var e = existing[i];
      if (e.pageIndex !== candidate.pageIndex) continue;
      var sameText = normText(e.text) === normText(candidate.text);
      var overlap = iou(e.box, candidate.box);

      /* Identical text on the same sheet counts as a duplicate regardless of
       * position. Requiring the boxes to agree does not work: a small vision
       * model's bbox is often far enough off that a true duplicate scores zero
       * overlap, and duplicates then leak onto every FAIR. A genuinely repeated
       * callout does exist on real drawings, so the suppressed count is reported
       * back to the operator rather than dropped silently. */
      if (sameText) return true;
      if (overlap > 0.45) return true;
    }
    return false;
  }

  function dist(a, b) {
    return Math.hypot((a.nx || 0) - (b.nx || 0), (a.ny || 0) - (b.ny || 0));
  }

  // ---------------------------------------------------------------------------
  // Pass runner
  // ---------------------------------------------------------------------------

  /* opts: { endpoint, model, rows, cols, scale, overlap, unit, generalTol,
   *         zoneRows, zoneCols, signal, onProgress(done,total,msg) } */
  function runVisionPass(loader, pageIndex, existingChars, opts) {
    var client = new OllamaClient(opts.endpoint);
    var rows = opts.rows || 2;
    var cols = opts.cols || 2;
    var report = opts.onProgress || function () {};

    var work = document.createElement('canvas');

    return loader.renderPage(pageIndex, work, opts.scale || 2.5).then(function () {
      var tiles = tilePage(work, rows, cols, opts.overlap);
      var found = [];
      var errors = [];
      var chain = Promise.resolve();

      tiles.forEach(function (tile, i) {
        chain = chain.then(function () {
          if (opts.signal && opts.signal.aborted) throw new DOMException('Aborted', 'AbortError');
          report(i, tiles.length, 'Reading region ' + (i + 1) + ' of ' + tiles.length + ' …');

          return client.vision(opts.model, PROMPT, toBase64(tile.canvas), opts.signal, opts.thinking)
            .then(function (raw) {
              parseResponse(raw).forEach(function (item) {
                found.push(mapToPage(item, tile, pageIndex, loader, opts));
              });
            })
            .catch(function (err) {
              if (err.name === 'AbortError') throw err;
              // One bad tile must not lose the other seven.
              errors.push('Region ' + (i + 1) + ': ' + err.message);
            });
        });
      });

      return chain.then(function () {
        report(tiles.length, tiles.length, 'Merging results …');

        var accepted = [];
        var suppressed = 0;
        found.forEach(function (c) {
          if (!c) return;
          if (isDuplicate(c, existingChars) || isDuplicate(c, accepted)) { suppressed++; return; }
          accepted.push(c);
        });

        return {
          characteristics: accepted,
          suppressed: suppressed,
          errors: errors,
          tiles: tiles.length
        };
      });
    });
  }

  function mapToPage(item, tile, pageIndex, loader, opts) {
    var page = loader.pages[pageIndex];

    // Run the model's string through the deterministic classifier so a vision
    // characteristic and a text-layer characteristic are interpreted identically.
    var c = B.classify({ text: item.text, remapped: false }, {
      unit: opts.unit || 'mm',
      generalTol: opts.generalTol || { one: 0.5, two: 0.25, three: 0.1, angular: 0.5 }
    });

    if (!c) {
      // Readable, but not a requirement the parser recognizes. Keep it as an
      // unclassified note rather than silently dropping something on the print.
      c = {
        id: B.uid(), text: item.text, requirement: item.text,
        type: 'Note', subtype: 'Unclassified (vision)',
        nominal: null, upper: null, lower: null,
        method: 'Review', flags: [], critical: false
      };
    }

    var bb = item.bbox;
    if (bb) {
      var x0 = Math.min(bb[0], bb[2]), x1 = Math.max(bb[0], bb[2]);
      var y0 = Math.min(bb[1], bb[3]), y1 = Math.max(bb[1], bb[3]);
      c.box = {
        x: tile.fx + Math.max(0, Math.min(1, x0)) * tile.fw,
        y: tile.fy + Math.max(0, Math.min(1, y0)) * tile.fh,
        w: Math.max(0.004, Math.min(1, x1 - x0)) * tile.fw,
        h: Math.max(0.004, Math.min(1, y1 - y0)) * tile.fh
      };
    } else {
      // No usable box: park it at the tile centre so the operator can drag it
      // into place rather than lose the characteristic.
      c.box = { x: tile.fx + tile.fw / 2, y: tile.fy + tile.fh / 2, w: 0.02, h: 0.012 };
      c.flags.push('No position returned — balloon placed at region centre');
    }

    c.nx = c.box.x + c.box.w / 2;
    c.ny = c.box.y + c.box.h / 2;
    c.balloon = {
      x: Math.max(0.008, Math.min(0.99, c.box.x - 0.022)),
      y: Math.max(0.008, Math.min(0.99, c.box.y - 0.018))
    };
    c.pageIndex = pageIndex;
    c.pageNumber = page.pageNumber;
    c.zone = B.computeZone(c.nx, c.ny, opts.zoneRows, opts.zoneCols);
    c.source = 'vision';

    // Non-negotiable: a transcription is not a verified requirement.
    c.review = true;
    c.flags.push('Read by vision model — verify text and position against the print');

    return c;
  }

  global.BallooningVision = {
    OllamaClient: OllamaClient,
    tilePage: tilePage,
    parseResponse: parseResponse,
    runVisionPass: runVisionPass,
    isDuplicate: isDuplicate,
    iou: iou,
    VISION_NAME_RE: VISION_NAME_RE,
    PROMPT: PROMPT
  };

})(typeof window !== 'undefined' ? window : this);
