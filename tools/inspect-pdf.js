#!/usr/bin/env node
/**
 * inspect-pdf.js — Phase 1 calibration tool for PDF Template Builder
 *
 * Reads a J&T / generic shipping-label PDF and produces a normalized layout.json
 * describing where every text run and image lives on the first page.
 *
 * Adds heuristic classification of LOCKED / SHRINKABLE / CUSTOMIZABLE zones
 * so the Phase 2 WYSIWYG editor can pin carrier-required elements.
 *
 * Usage:
 *   node tools/inspect-pdf.js <pdf-path> <out-json-path>
 *
 * Dependencies:
 *   pdf2json   (preferred — `npm install pdf2json`, or `npx --yes pdf2json ...`)
 *
 * Graceful degradation:
 *   - If pdf2json is missing, prints install hint and falls back to a minimal
 *     raw-stream text extractor (limited: no image bbox, best-effort only).
 *
 * IMPORTANT calibration note:
 *   Our own example.pdf was printed THROUGH this extension so it already has
 *   overlays baked in (worker name @ top-right, alias watermark @ bottom).
 *   Text found in those zones is tagged `extensionOverlay: true` and NOT
 *   classified as LOCKED carrier content.
 */

'use strict';

const fs = require('fs');
const path = require('path');

// ── Constants ─────────────────────────────────────────────────────────────
const POINTS_PER_PDFUNIT = 1; // pdf2json already emits PDF points via `x/y` scaled units
const JNT_FINGERPRINTS = [/\bJ&T\b/i, /\bJNT\b/i, /jtexpress/i, /เจแอนด์ที/];
const FLASH_FINGERPRINTS = [/flash[ \-]?express/i, /แฟลช/];
const KERRY_FINGERPRINTS = [/\bkerry\b/i, /เคอรี่/];
const SPX_FINGERPRINTS = [/\bSPX\b/i, /shopee ?express/i];
const THAIPOST_FINGERPRINTS = [/thailand ?post/i, /ไปรษณีย์ไทย/];

// Extension overlays we know exist on already-processed labels
const EXTENSION_OVERLAY_BOTTOM_PT = 36; // alias baseline y≈6, top of text up to ~28pt; pad to 36 for safety
const EXTENSION_OVERLAY_TOPRIGHT_FRACTION = 0.12; // top 12% × right 40%

// ── CLI entry ─────────────────────────────────────────────────────────────
async function main() {
  const [,, pdfPath, outPath] = process.argv;
  if (!pdfPath || !outPath) {
    console.error('Usage: node tools/inspect-pdf.js <pdf-path> <out-json-path>');
    process.exit(2);
  }
  if (!fs.existsSync(pdfPath)) {
    console.error(`ERR: PDF not found: ${pdfPath}`);
    process.exit(2);
  }
  const absPdf = path.resolve(pdfPath);
  const absOut = path.resolve(outPath);

  let layout;
  try {
    layout = await extractViaPdf2json(absPdf);
  } catch (e) {
    console.error('[warn] pdf2json path failed:', e.message);
    console.error('       install hint: npm install pdf2json   (or: npx --yes pdf2json ...)');
    console.error('       falling back to raw-stream text parser (no image bboxes)');
    layout = await extractViaRawFallback(absPdf);
    layout.warnings = layout.warnings || [];
    layout.warnings.push('pdf2json unavailable — image bboxes skipped, text only');
  }

  const classified = classifyRegions(layout);
  fs.mkdirSync(path.dirname(absOut), { recursive: true });
  fs.writeFileSync(absOut, JSON.stringify(classified, null, 2), 'utf8');
  console.log(`[ok] layout written: ${absOut}`);
  console.log(`[ok] carrier=${classified.carrier} page=${classified.pageSize.w}×${classified.pageSize.h}pt`);
  console.log(`[ok] texts=${classified.texts.length} images=${classified.images.length}`);
  console.log(`[ok] locked=${Object.keys(classified.hints.locked).length} shrinkable=${Object.keys(classified.hints.shrinkable).length}`);
}

// ── pdf2json extractor ────────────────────────────────────────────────────
async function extractViaPdf2json(pdfPath) {
  let PDFParser;
  try {
    // eslint-disable-next-line global-require
    PDFParser = require('pdf2json');
  } catch (e) {
    throw new Error('module pdf2json not installed');
  }

  return new Promise((resolve, reject) => {
    const parser = new PDFParser(null, 1); // verbosity low
    parser.on('pdfParser_dataError', (err) => reject(new Error(err?.parserError || String(err))));
    parser.on('pdfParser_dataReady', (data) => {
      try {
        resolve(normalizePdf2jsonData(data));
      } catch (e) { reject(e); }
    });
    parser.loadPDF(pdfPath);
  });
}

/**
 * pdf2json emits coordinates in its own "units" where 1 unit ≈ 16pt horizontally
 * / 16pt vertically (depends on page size). We convert back to PDF points using
 * the ratio between reported Width/Height (in units) and actual page dimensions.
 */
function normalizePdf2jsonData(data) {
  const page = (data.Pages && data.Pages[0]) || {};
  // pdf2json page.Width/Height are in its internal unit scale where unit=16pt by default.
  // Actual page dimensions in points:
  const wPt = (page.Width || 0) * 16;
  const hPt = (page.Height || 0) * 16;
  const unitToPt = 16;

  const rawTexts = (page.Texts || []).map((t, i) => {
    const runs = (t.R || []).map((r) => ({
      text: safeDecodeURI(r.T || ''),
      size: (r.TS && r.TS[1]) ? Number(r.TS[1]) : null,
      bold: !!(r.TS && r.TS[2]),
      italic: !!(r.TS && r.TS[3]),
    }));
    const full = runs.map((r) => r.text).join('');
    return {
      id: `t_${i}`,
      text: full,
      x: +(t.x * unitToPt).toFixed(2),
      // pdf2json uses top-left origin; convert to PDF bottom-left origin
      y: +(hPt - (t.y * unitToPt)).toFixed(2),
      w: +((t.w || 0) * unitToPt / 4).toFixed(2),
      size: runs[0]?.size || null,
      bold: runs[0]?.bold || false,
      raw: runs,
    };
  }).filter((t) => t.text.length > 0);

  const texts = coalesceGlyphs(rawTexts);

  const images = (page.Fills || []).length >= 0 ? [] : []; // pdf2json drops raster images by default
  // pdf2json exposes bitmap / picture fills under page.Fills with oc flag sometimes; scanner
  // heuristics will instead rely on text density + gaps, so we leave images as []
  // But if page.VLines / HLines define framing boxes, treat big ones as candidate image regions.
  const hLines = page.HLines || [];
  const vLines = page.VLines || [];
  const rects = inferRectanglesFromLines(hLines, vLines, unitToPt, hPt);

  return {
    pageSize: { w: +wPt.toFixed(2), h: +hPt.toFixed(2) },
    texts,
    images,      // empty — pdf2json can't reliably give raster bboxes
    rects,       // line-framed boxes (often wrap barcodes, QR, images)
    meta: {
      extractor: 'pdf2json',
      pageCount: (data.Pages || []).length,
    },
  };
}

function safeDecodeURI(s) {
  try { return decodeURIComponent(s); } catch { return s; }
}

/**
 * pdf2json emits one "Text" per glyph for Thai/CJK content. Coalesce glyphs
 * sharing the same row (y within ±1.5pt) and same font-size into a single
 * logical text run. Preserves x of the first glyph and accumulates width.
 */
function coalesceGlyphs(raws) {
  if (!raws.length) return [];
  // Sort by y (desc — PDF bottom-left origin, higher y = higher on page) then x
  const sorted = [...raws].sort((a, b) => (b.y - a.y) || (a.x - b.x));
  const rows = [];
  const TOL_Y = 2.0;
  for (const g of sorted) {
    const row = rows.find((r) => Math.abs(r.y - g.y) < TOL_Y && Math.abs((r.size || 0) - (g.size || 0)) < 0.5);
    if (row) row.items.push(g);
    else rows.push({ y: g.y, size: g.size, items: [g] });
  }
  const out = [];
  let id = 0;
  for (const row of rows) {
    row.items.sort((a, b) => a.x - b.x);
    let current = null;
    for (const g of row.items) {
      if (!current) {
        current = { ...g, text: g.text, id: `t_${id++}`, endX: g.x + (g.w || 0) };
        continue;
      }
      const gap = g.x - current.endX;
      // 6pt gap tolerance — same logical string
      if (gap <= 6) {
        current.text += g.text;
        current.endX = Math.max(current.endX, g.x + (g.w || 0));
        current.bold = current.bold || g.bold;
      } else {
        out.push({ id: current.id, text: current.text, x: current.x, y: current.y,
                   w: +(current.endX - current.x).toFixed(2), size: current.size,
                   bold: current.bold, raw: current.raw });
        current = { ...g, text: g.text, id: `t_${id++}`, endX: g.x + (g.w || 0) };
      }
    }
    if (current) {
      out.push({ id: current.id, text: current.text, x: current.x, y: current.y,
                 w: +(current.endX - current.x).toFixed(2), size: current.size,
                 bold: current.bold, raw: current.raw });
    }
  }
  return out.filter((t) => t.text.trim().length > 0);
}

/** Build candidate "rect" boxes from HLines + VLines (pdf2json framing). */
function inferRectanglesFromLines(hLines, vLines, unitToPt, hPt) {
  const out = [];
  // Pair up horizontal lines by y; simple greedy rect candidates
  for (let i = 0; i < hLines.length; i += 1) {
    const a = hLines[i];
    for (let j = i + 1; j < hLines.length; j += 1) {
      const b = hLines[j];
      if (Math.abs(a.x - b.x) < 0.5 && Math.abs(a.l - b.l) < 0.5 && Math.abs(a.y - b.y) > 0.3) {
        out.push({
          x: +(a.x * unitToPt).toFixed(2),
          y: +(hPt - (Math.max(a.y, b.y) * unitToPt)).toFixed(2),
          w: +(a.l * unitToPt).toFixed(2),
          h: +(Math.abs(a.y - b.y) * unitToPt).toFixed(2),
        });
      }
    }
  }
  return out.slice(0, 60); // cap noise
}

// ── Raw-stream fallback (no deps) ─────────────────────────────────────────
async function extractViaRawFallback(pdfPath) {
  const buf = fs.readFileSync(pdfPath);
  const str = buf.toString('latin1');
  // Page size via /MediaBox
  const mb = str.match(/\/MediaBox\s*\[\s*([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s*\]/);
  const w = mb ? Number(mb[3]) - Number(mb[1]) : 288;
  const h = mb ? Number(mb[4]) - Number(mb[2]) : 432;
  // Extract text fragments (crude: BT ... Tj ... ET)
  const texts = [];
  const re = /BT[\s\S]*?ET/g;
  let m; let idx = 0;
  while ((m = re.exec(str))) {
    const block = m[0];
    const tm = block.match(/([-\d.]+)\s+([-\d.]+)\s+Td/);
    const pieces = [...block.matchAll(/\((.*?)\)\s*Tj/g)].map((x) => x[1]);
    if (pieces.length && tm) {
      texts.push({
        id: `t_${idx++}`,
        text: pieces.join(''),
        x: Number(tm[1]),
        y: Number(tm[2]),
        size: null, bold: false, raw: [],
      });
    }
  }
  return {
    pageSize: { w, h },
    texts, images: [], rects: [],
    meta: { extractor: 'raw-fallback', pageCount: 1 },
  };
}

// ── Classification ────────────────────────────────────────────────────────
function classifyRegions(layout) {
  const { pageSize, texts, images, rects } = layout;
  const carrier = detectCarrier(texts);

  // Flag extension overlays first so they're filtered out of LOCKED scan
  const topRightYCutoff = pageSize.h * (1 - EXTENSION_OVERLAY_TOPRIGHT_FRACTION);
  const topRightXCutoff = pageSize.w * 0.6;
  const taggedTexts = texts.map((t) => {
    const isThai = /[\u0E00-\u0E7F]/.test(t.text);
    const inTopRight = t.y >= topRightYCutoff && t.x >= topRightXCutoff && isThai;
    const inBottomBand = t.y <= EXTENSION_OVERLAY_BOTTOM_PT;
    const extensionOverlay = inTopRight || inBottomBand;
    return { ...t, extensionOverlay };
  });

  const hints = { locked: {}, shrinkable: {}, customizable: {}, unclassified: [] };

  // 1) Order ID — long digit string (14+ digits) anywhere on page
  const orderIdTxt = taggedTexts
    .filter((t) => !t.extensionOverlay)
    .find((t) => /\d{14,}/.test(t.text));
  if (orderIdTxt) {
    const digits = orderIdTxt.text.match(/\d{14,}/)[0];
    hints.locked.orderId = {
      value: redactDigits(digits),
      x: orderIdTxt.x, y: orderIdTxt.y,
      size: orderIdTxt.size, textId: orderIdTxt.id,
      reason: 'digit run ≥14 chars — order/tracking ID',
    };
  }

  // 2) Tracking / waybill — 10–13 digit isolated string (J&T uses 12)
  const trackTxt = taggedTexts
    .filter((t) => !t.extensionOverlay)
    .find((t) => /^\s*\d{10,13}\s*$/.test(t.text) && (!orderIdTxt || t.id !== orderIdTxt.id));
  if (trackTxt) {
    hints.locked.trackingNumber = {
      value: redactDigits(trackTxt.text.trim()),
      x: trackTxt.x, y: trackTxt.y,
      size: trackTxt.size, textId: trackTxt.id,
      reason: 'isolated 10–13 digit string — carrier tracking number',
    };
  }

  // 3) Sort code — short alphanumeric (1–6 chars) in large font near top-right
  const bigTexts = [...taggedTexts]
    .filter((t) => t.size && t.size >= 18 && !t.extensionOverlay)
    .sort((a, b) => (b.size || 0) - (a.size || 0));
  const sortCodeCand = bigTexts.find((t) => /^[A-Z0-9\-]{1,8}$/.test(t.text.trim()));
  if (sortCodeCand) {
    hints.locked.sortCode = {
      value: sortCodeCand.text.trim(),
      x: sortCodeCand.x, y: sortCodeCand.y,
      size: sortCodeCand.size, textId: sortCodeCand.id,
      reason: 'large-font short alphanum — carrier sort code',
    };
  }

  // 3b) Route / zone code — very large (80pt+) mixed alphanum (J&T "H1 F04-33")
  const routeCand = [...taggedTexts]
    .filter((t) => !t.extensionOverlay && t.size && t.size >= 80 && /[A-Z]/.test(t.text) && /\d/.test(t.text))
    .sort((a, b) => (b.size || 0) - (a.size || 0))[0];
  if (routeCand && (!sortCodeCand || routeCand.id !== sortCodeCand.id)) {
    hints.locked.routeCode = {
      value: routeCand.text.trim(),
      x: routeCand.x, y: routeCand.y,
      size: routeCand.size, textId: routeCand.id,
      reason: 'very large alphanum string — carrier route/zone code',
    };
  }

  // 3c) Sub-zone code — medium alphanum (50–80pt) alphanum like "004A"
  const subZoneCand = [...taggedTexts]
    .filter((t) => !t.extensionOverlay && t.size && t.size >= 50 && t.size < 80
      && /^[A-Z0-9]{2,6}$/.test(t.text.trim())
      && (!sortCodeCand || t.id !== sortCodeCand.id)
      && (!routeCand || t.id !== routeCand.id))
    .sort((a, b) => (b.size || 0) - (a.size || 0))[0];
  if (subZoneCand) {
    hints.locked.subZoneCode = {
      value: subZoneCand.text.trim(),
      x: subZoneCand.x, y: subZoneCand.y,
      size: subZoneCand.size, textId: subZoneCand.id,
      reason: 'medium alphanum — carrier sub-zone code',
    };
  }

  // 3d) Service type badge — "DROP-OFF" / "PICK-UP" / "COD"
  const serviceCand = taggedTexts.find((t) => /^\s*(DROP-?OFF|PICK-?UP|COD|CODE)\s*$/i.test(t.text.trim()));
  if (serviceCand) {
    hints.locked.serviceType = {
      value: serviceCand.text.trim(),
      x: serviceCand.x, y: serviceCand.y,
      size: serviceCand.size, textId: serviceCand.id,
      reason: 'service type badge — carrier-defined, must stay readable',
    };
  }
  const codCand = taggedTexts.find((t) => /\bC ?OD\b/.test(t.text) && (t.size || 0) >= 80);
  if (codCand) {
    hints.locked.codLabel = {
      value: codCand.text.trim(),
      x: codCand.x, y: codCand.y,
      size: codCand.size, textId: codCand.id,
      reason: 'large COD label — cash-on-delivery flag (carrier-required)',
    };
  }

  // 4) Barcode main — widest rect in top 40% of page
  const top40 = pageSize.h * 0.6;
  const barcodeCand = [...rects]
    .filter((r) => r.y >= top40 && r.w >= pageSize.w * 0.5)
    .sort((a, b) => b.w * b.h - a.w * a.h)[0];
  if (barcodeCand) {
    hints.locked.barcodeMain = {
      ...barcodeCand,
      reason: 'largest framed rect in top 40% — main barcode region',
    };
  }

  // 5) QR code — square-ish rect mid-right
  const qrCand = [...rects]
    .filter((r) => Math.abs(r.w - r.h) / Math.max(r.w, r.h, 1) < 0.2 && r.w >= 60 && r.w <= 180)
    .sort((a, b) => b.w - a.w)[0];
  if (qrCand) {
    hints.locked.qrCode = { ...qrCand, reason: 'square ~60–180pt rect — carrier QR code' };
  }

  // 6) Side digit columns — J&T prints rotated tracking digits on both edges
  //    (these ENCODE the barcode for visual redundancy — classify as LOCKED)
  const leftCol = taggedTexts.filter((t) => t.x < 15 && /^\d$/.test(t.text.trim()));
  const rightCol = taggedTexts.filter((t) => t.x > (pageSize.w - 25) && /^\d$/.test(t.text.trim()));
  if (leftCol.length >= 6) {
    const ys = leftCol.map((t) => t.y);
    hints.locked.barcodeLeft = {
      x: Math.min(...leftCol.map((t) => t.x)),
      y: Math.min(...ys),
      w: 14,
      h: Math.max(...ys) - Math.min(...ys) + 10,
      glyphCount: leftCol.length,
      reason: 'vertical digit column on left edge — J&T side barcode mirror',
    };
  }
  if (rightCol.length >= 6) {
    const ys = rightCol.map((t) => t.y);
    hints.locked.barcodeRight = {
      x: Math.min(...rightCol.map((t) => t.x)),
      y: Math.min(...ys),
      w: 14,
      h: Math.max(...ys) - Math.min(...ys) + 10,
      glyphCount: rightCol.length,
      reason: 'vertical digit column on right edge — J&T side barcode mirror',
    };
  }

  // 7) Shrinkable — SKU table area (detected by "Product Name" / "SKU" / "Qty" header)
  const skuHeader = taggedTexts.find((t) => /Product ?Name|SKU|Qty/i.test(t.text));
  if (skuHeader) {
    // Table spans from header down to Qty Total row
    const qtyTotal = taggedTexts.find((t) => /Qty ?Total/i.test(t.text));
    const yTop = skuHeader.y;
    const yBot = qtyTotal ? qtyTotal.y : Math.max(0, skuHeader.y - 60);
    hints.shrinkable.skuTable = {
      x: 0, y: yBot,
      w: pageSize.w,
      h: yTop - yBot + (skuHeader.size ? skuHeader.size * 0.2 : 6),
      reason: 'SKU table — detected via "Product Name/SKU/Qty" header',
    };
  }

  // 8) Shrinkable — address block (Thai text cluster between tracking and order ID)
  //    Keep x=0..pageSize.w, y roughly pageSize.h*0.35 .. pageSize.h*0.65
  hints.shrinkable.addressBlock = {
    x: 0, y: pageSize.h * 0.35, w: pageSize.w, h: pageSize.h * 0.3,
    reason: 'approximate — Thai sender/recipient address lines between tracking area and SKU table',
    needsVerification: true,
  };

  // (Platform logo band detection removed — the "top 5%" band on J&T holds the
  //  sort code and carrier-required elements, not a platform logo. TikTok logo
  //  lives inline somewhere else and needs another sample to calibrate.)

  // 7) Extension overlays (ours) — NOT LOCKED, shrinkable/customizable
  const overlayBottom = taggedTexts.filter((t) => t.extensionOverlay && t.y <= EXTENSION_OVERLAY_BOTTOM_PT);
  if (overlayBottom.length) {
    hints.customizable.aliasWatermark = {
      count: overlayBottom.length,
      sampleY: overlayBottom[0].y,
      extensionOverlay: true,
      reason: 'added by QF extension overlayAliasOnPdf() — not a LOCKED carrier element',
    };
  }
  const overlayTopRight = taggedTexts.filter((t) => t.extensionOverlay && t.y > EXTENSION_OVERLAY_BOTTOM_PT);
  if (overlayTopRight.length) {
    hints.customizable.workerNameBadge = {
      count: overlayTopRight.length,
      sampleY: overlayTopRight[0].y,
      extensionOverlay: true,
      reason: 'added by QF extension overlayAliasOnPdf() — worker name badge',
    };
  }

  // 10) Anything else → unclassified (for user review)
  const claimedIds = new Set(
    Object.values(hints.locked).concat(Object.values(hints.shrinkable))
      .map((r) => r.textId).filter(Boolean),
  );
  // Also treat side digit columns (barcodeLeft/Right) as claimed
  const isSideDigit = (t) => (
    (t.x < 15 || t.x > pageSize.w - 25) && /^\d$/.test(t.text.trim())
  );
  // Claim order/tracking neighbor labels by proximity
  const claimByProximity = (t) => {
    for (const key of ['orderId', 'trackingNumber', 'sortCode']) {
      const r = hints.locked[key]; if (!r) continue;
      if (Math.abs(t.y - r.y) < 6 && Math.abs(t.x - r.x) < 140) return true;
    }
    return false;
  };
  hints.unclassified = taggedTexts
    .filter((t) => !t.extensionOverlay && !claimedIds.has(t.id) && !isSideDigit(t) && !claimByProximity(t))
    .slice(0, 40)
    .map((t) => ({ id: t.id, preview: redactText(t.text).slice(0, 40), x: t.x, y: t.y, size: t.size }));

  return {
    carrier,
    pageSize,
    texts: taggedTexts.map((t) => ({
      id: t.id, text: redactText(t.text), x: t.x, y: t.y, w: t.w, size: t.size, bold: t.bold,
      extensionOverlay: t.extensionOverlay,
    })),
    images,
    rects,
    hints,
    meta: layout.meta,
    notes: [
      'Phase 1 calibration output. SAMPLE PDF was already post-processed by this extension,',
      'so bottom-band Thai text and top-right Thai text are extension overlays (not carrier LOCKED).',
      'For 100% accurate LOCKED coords please re-run on a RAW J&T PDF from the generate API.',
    ],
  };
}

function detectCarrier(texts) {
  const joined = texts.map((t) => t.text).join(' ');
  if (JNT_FINGERPRINTS.some((r) => r.test(joined))) return 'jnt';
  if (FLASH_FINGERPRINTS.some((r) => r.test(joined))) return 'flash';
  if (KERRY_FINGERPRINTS.some((r) => r.test(joined))) return 'kerry';
  if (SPX_FINGERPRINTS.some((r) => r.test(joined))) return 'spx';
  if (THAIPOST_FINGERPRINTS.some((r) => r.test(joined))) return 'thaipost';

  // Structural heuristics — J&T signatures:
  // - vertical digit columns on both left+right edges (x<15 AND x>270)
  // - 12-digit tracking number
  // - "DROP-OFF" or "IN TRANSIT BY:" phrases
  const leftCol = texts.filter((t) => t.x < 15 && /^\d$/.test(t.text.trim()));
  const rightCol = texts.filter((t) => t.x > 270 && /^\d$/.test(t.text.trim()));
  const has12Digit = texts.some((t) => /^\d{12}$/.test(t.text.trim()) || /^\d{4}[•\-]{2}\d{2}$/.test(t.text.trim()));
  const dropoff = /DROP-OFF/i.test(joined);
  if (leftCol.length >= 8 && rightCol.length >= 8 && (has12Digit || dropoff)) return 'jnt';

  return 'unknown';
}

// Redact obvious PII before writing to JSON — we still record coords.
function redactText(s) {
  if (!s) return s;
  let out = s;
  // Phone (TH 0[0-9]{8,9} or +66...)
  out = out.replace(/(\+?66|0)\d{8,9}/g, '{phone_redacted}');
  // Long digit runs (order / tracking) — keep first 4 + last 2 for reference
  out = out.replace(/\d{10,}/g, (m) => `${m.slice(0, 4)}••${m.slice(-2)}`);
  return out;
}
function redactDigits(d) {
  if (!d) return d;
  return d.length > 6 ? `${d.slice(0, 4)}••${d.slice(-2)}` : d;
}

// ── go ─────────────────────────────────────────────────────────────────────
main().catch((e) => {
  console.error('FATAL:', e.stack || e.message);
  process.exit(1);
});
