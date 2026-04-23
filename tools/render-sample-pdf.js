#!/usr/bin/env node
/*
 * tools/render-sample-pdf.js — Render a J&T sample PDF to a PNG preview
 *
 * Generates vendor/jnt_sample_preview.png used as the background image in
 * the PDF Template Builder editor. The sample gives users visual context
 * for where logos/text will land relative to LOCKED carrier zones.
 *
 * Pipeline:
 *   1. Load the PDF via pdfjs-dist (legacy build — Node compatible)
 *   2. Render page N onto a node-canvas at 2× A6 (596×840px)
 *   3. Convert to grayscale (the shipped asset must not leak brand colour)
 *   4. Overlay white rectangles on PHI (phone / recipient name / address)
 *   5. Write PNG to vendor/jnt_sample_preview.png
 *
 * Usage:
 *   node tools/render-sample-pdf.js [pdfPath] [pageNumber]
 *     pdfPath     default .claude/example.pdf
 *     pageNumber  default 1 (1-indexed)
 *
 * Fallback (if pdfjs + canvas both fail): produces a placeholder PNG
 * — a white A6 canvas with the text "J&T Sample (placeholder)". The
 * extension still loads, the editor shows a plain preview.
 *
 * Privacy: ALL runs of this script apply PHI masking before saving.
 * Coordinates come from .claude/samples/jnt_layout.json (addressBlock
 * SHRINKABLE zone + unclassified Thai/phone text runs t_56, t_63,
 * t_65—t_92). If you re-render from a different sample, update the
 * PHI_RECTS_PDF array below so the new sample is also sanitized.
 */

'use strict';

const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const DEFAULT_PDF = path.join(ROOT, '.claude', 'example.pdf');
const OUT_PNG = path.join(ROOT, 'vendor', 'jnt_sample_preview.png');

// A6 page size in PDF points (from .claude/samples/jnt_layout.json)
const PAGE_W_PT = 298;
const PAGE_H_PT = 420;
// 2× scale — matches PDF_TEMPLATE_SCALE in content.js
const SCALE = 2;
const OUT_W = PAGE_W_PT * SCALE; // 596
const OUT_H = PAGE_H_PT * SCALE; // 840

// PHI rectangles in PDF points (bottom-left origin, matching jnt_layout.json).
// Derived from:
//   - SHRINKABLE.addressBlock: x:0 y:147 w:298 h:126  → whole address + phone block
//   - t_56 "praaorawee" (recipient name): y=285.04 size=71   (inside addressBlock)
//   - t_63 "(+66)90*****49" (phone):     y=273.87 size=34   (inside addressBlock)
// The addressBlock covers everything sensitive — we just mask that.
// We also mask the SKU table area where product names might reveal merchant identity
// if the sample had personal info, AND the sender line (t_32/t_33 y≈316) which
// shows a shop address. Shopping info (orderId, tracking, barcode, QR) stays visible
// so users can see the LOCKED zones.
const PHI_RECTS_PDF = [
  // Full recipient + sender block: y=147 (addressBlock top in jnt_layout.json)
  // extended up to y=335 to cover recipient name (t_56 y=285, size=71) and the
  // merchant sender lines (t_31 "จาก" y=328, t_32 "ตลาดเมืองนครราชเมื" y=317,
  // t_40 address line y=308). We keep barcode columns at x=9.62 & x=277.68
  // visible by hugging the inner content area x=30..270.
  { x: 15,  y: 147, w: 268, h: 200, note: 'recipient + sender block' },
];

function pdfRectToPx(r) {
  // pdf-points bottom-left → canvas top-left pixel coords
  const x = r.x * SCALE;
  const y = (PAGE_H_PT - r.y - r.h) * SCALE;
  const w = r.w * SCALE;
  const h = r.h * SCALE;
  return { x, y, w, h };
}

async function renderWithPdfjs(pdfPath, pageNumber) {
  // pdfjs-dist ships only ESM — use dynamic import from legacy build.
  const pdfjsLib = await import('pdfjs-dist/legacy/build/pdf.mjs');
  // pdfjs-dist 5.x requires its optional peer `@napi-rs/canvas` for Node
  // rendering. It's auto-installed via optionalDependencies.
  const { createCanvas } = require('@napi-rs/canvas');

  const data = new Uint8Array(fs.readFileSync(pdfPath));
  const loadingTask = pdfjsLib.getDocument({
    data,
    // Disable worker in Node.
    disableWorker: true,
    // Don't try to load system fonts (avoids warnings on headless Linux).
    useSystemFonts: false,
    // Silence font warnings — we convert to grayscale anyway and PHI is masked.
    verbosity: 0,
  });
  const pdf = await loadingTask.promise;
  if (pageNumber > pdf.numPages) {
    throw new Error(`PDF has ${pdf.numPages} pages — requested page ${pageNumber}`);
  }
  const page = await pdf.getPage(pageNumber);
  const viewport = page.getViewport({ scale: SCALE });
  const w = Math.round(viewport.width);
  const h = Math.round(viewport.height);

  const canvas = createCanvas(w, h);
  const ctx = canvas.getContext('2d');
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, w, h);

  // pdfjs 5.x: pass `canvas` directly; renderer uses its own canvasFactory.
  await page.render({ canvas, viewport }).promise;

  return { canvas, ctx, w, h };
}

function applyGrayscaleAndMaskPhi({ canvas, ctx, w, h }) {
  // Grayscale: convert pixel-by-pixel (node-canvas supports ctx.filter but
  // reliability varies across versions — explicit loop is deterministic).
  const img = ctx.getImageData(0, 0, w, h);
  const d = img.data;
  for (let i = 0; i < d.length; i += 4) {
    // ITU-R BT.709 luma
    const gray = Math.round(0.2126 * d[i] + 0.7152 * d[i + 1] + 0.0722 * d[i + 2]);
    d[i] = gray;
    d[i + 1] = gray;
    d[i + 2] = gray;
  }
  ctx.putImageData(img, 0, 0);

  // Overlay white rectangles on PHI. Use solid white with a thin border
  // so users still see the zone is intentionally masked.
  let maskedPixels = 0;
  for (const r of PHI_RECTS_PDF) {
    const px = pdfRectToPx(r);
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(px.x, px.y, px.w, px.h);
    ctx.strokeStyle = '#cbd5e1';
    ctx.lineWidth = 1;
    ctx.strokeRect(px.x + 0.5, px.y + 0.5, px.w - 1, px.h - 1);
    // Center label
    ctx.fillStyle = '#94a3b8';
    ctx.font = 'italic 12px sans-serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    // ASCII-only label — @napi-rs/canvas has no Thai font, non-ASCII → tofu
    ctx.fillText('[ address hidden ]', px.x + px.w / 2, px.y + px.h / 2);
    maskedPixels += Math.round(px.w * px.h);
  }
  ctx.textAlign = 'start';
  ctx.textBaseline = 'alphabetic';
  return maskedPixels;
}

function writePlaceholder() {
  // Prefer @napi-rs/canvas (no native build step); fall back to `canvas`.
  let createCanvas;
  try { ({ createCanvas } = require('@napi-rs/canvas')); }
  catch { ({ createCanvas } = require('canvas')); }
  const canvas = createCanvas(OUT_W, OUT_H);
  const ctx = canvas.getContext('2d');
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, OUT_W, OUT_H);
  ctx.strokeStyle = '#cbd5e1';
  ctx.lineWidth = 2;
  ctx.strokeRect(1, 1, OUT_W - 2, OUT_H - 2);
  ctx.fillStyle = '#64748b';
  ctx.font = 'bold 32px sans-serif';
  ctx.textAlign = 'center';
  ctx.fillText('J&T Sample', OUT_W / 2, OUT_H / 2 - 20);
  ctx.font = '18px sans-serif';
  ctx.fillText('(placeholder — render failed)', OUT_W / 2, OUT_H / 2 + 16);
  ctx.font = '14px sans-serif';
  ctx.fillText('298 × 420 pt (A6)', OUT_W / 2, OUT_H / 2 + 44);
  return canvas;
}

async function main() {
  const pdfPath = process.argv[2] || DEFAULT_PDF;
  const pageNumber = Number(process.argv[3] || 1);

  let canvas;
  let maskedPixels = 0;
  let mode = 'pdfjs';

  try {
    if (!fs.existsSync(pdfPath)) {
      throw new Error(`PDF not found: ${pdfPath}`);
    }
    const result = await renderWithPdfjs(pdfPath, pageNumber);
    maskedPixels = applyGrayscaleAndMaskPhi(result);
    canvas = result.canvas;
  } catch (err) {
    console.error('[render-sample-pdf] PDF render failed:', err.message);
    console.error('[render-sample-pdf] Falling back to placeholder PNG.');
    canvas = writePlaceholder();
    mode = 'placeholder';
  }

  const outDir = path.dirname(OUT_PNG);
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });
  const buf = canvas.toBuffer('image/png');
  fs.writeFileSync(OUT_PNG, buf);

  console.log(`[render-sample-pdf] mode=${mode}`);
  console.log(`[render-sample-pdf] output=${OUT_PNG}`);
  console.log(`[render-sample-pdf] size=${buf.length} bytes (${canvas.width}×${canvas.height})`);
  console.log(`[render-sample-pdf] maskedPixels=${maskedPixels}`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
