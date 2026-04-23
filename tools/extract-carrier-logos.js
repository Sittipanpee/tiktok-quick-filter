#!/usr/bin/env node
/*
 * tools/extract-carrier-logos.js — Extract carrier + TikTok logos as PNG assets
 *
 * Renders the top band of a J&T sample PDF at ultra-high resolution, then
 * crops two bounding boxes (TikTok Shop mark, J&T Express mark) into shipped
 * grayscale PNGs. The logos are brand marks only — no PHI concern.
 *
 * Output:
 *   vendor/tiktok_shop_logo.png   (~120×60 px, grayscale, PNG)
 *   vendor/jnt_express_logo.png   (~160×60 px, grayscale, PNG)
 *
 * Pipeline (mirrors tools/render-sample-pdf.js):
 *   1. pdfjs-dist (legacy) loads the PDF
 *   2. @napi-rs/canvas renders page 1 at SCALE=4  (ultra-sharp)
 *   3. Crop two logo rects (defined in PDF points, bottom-left origin)
 *   4. Convert to grayscale (ITU-R BT.709 luma)
 *   5. Add 2px transparent padding around the crop
 *   6. Write PNG files
 *
 * Usage:
 *   node tools/extract-carrier-logos.js [pdfPath] [pageNumber]
 *     pdfPath     default .claude/example.pdf
 *     pageNumber  default 1
 *
 * Privacy: operates on the top header band only (y ≥ 395 pt on A6) — never
 * touches address/phone/recipient text. Logos are brand marks, not PHI.
 */

'use strict';

const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const DEFAULT_PDF = path.join(ROOT, '.claude', 'example.pdf');
const OUT_TIKTOK = path.join(ROOT, 'vendor', 'tiktok_shop_logo.png');
const OUT_JNT = path.join(ROOT, 'vendor', 'jnt_express_logo.png');

// A6 page size in PDF points (from .claude/samples/jnt_layout.json)
const PAGE_W_PT = 298;
const PAGE_H_PT = 420;
// 4× scale for sharp crops
const SCALE = 4;

// Logo bounding boxes in PDF points (bottom-left origin).
// Both logos live in the top header band, above the tracking number barcode.
// Approximated from the sample — tweak PAD_PT if a new sample shifts.
const PAD_PT = 2;
const LOGOS = [
  {
    name: 'tiktok',
    out: OUT_TIKTOK,
    // Top-left header: small "TikTok Shop" mark above the barcode area.
    // PDF points: x 0–90, y 395–418 (23pt tall).
    rectPt: { x: 0, y: 395, w: 92, h: 24 },
  },
  {
    name: 'jnt',
    out: OUT_JNT,
    // Next to TikTok logo: "J&T Express" mark.
    // Range conservative; extension can't control exact placement but
    // cropping slack pixels is cheap.
    rectPt: { x: 92, y: 395, w: 96, h: 24 },
  },
];

function padRect(r, pad) {
  return {
    x: Math.max(0, r.x - pad),
    y: Math.max(0, r.y - pad),
    w: r.w + pad * 2,
    h: r.h + pad * 2,
  };
}

function pdfRectToPx(r, pageHPt, scale) {
  // pdf-points bottom-left → canvas top-left pixel coords
  const x = Math.round(r.x * scale);
  const y = Math.round((pageHPt - r.y - r.h) * scale);
  const w = Math.round(r.w * scale);
  const h = Math.round(r.h * scale);
  return { x, y, w, h };
}

async function renderPage(pdfPath, pageNumber, scale) {
  const pdfjsLib = await import('pdfjs-dist/legacy/build/pdf.mjs');
  const { createCanvas } = require('@napi-rs/canvas');

  const data = new Uint8Array(fs.readFileSync(pdfPath));
  const loadingTask = pdfjsLib.getDocument({
    data,
    disableWorker: true,
    useSystemFonts: false,
    verbosity: 0,
  });
  const pdf = await loadingTask.promise;
  if (pageNumber > pdf.numPages) {
    throw new Error(`PDF has ${pdf.numPages} pages — requested ${pageNumber}`);
  }
  const page = await pdf.getPage(pageNumber);
  const viewport = page.getViewport({ scale });
  const w = Math.round(viewport.width);
  const h = Math.round(viewport.height);

  const canvas = createCanvas(w, h);
  const ctx = canvas.getContext('2d');
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, w, h);
  await page.render({ canvas, viewport }).promise;
  return { canvas, ctx, w, h, pageWPt: viewport.width / scale, pageHPt: viewport.height / scale };
}

function cropAndGrayscale(srcCtx, rectPx) {
  const { createCanvas } = require('@napi-rs/canvas');
  const out = createCanvas(rectPx.w, rectPx.h);
  const outCtx = out.getContext('2d');
  outCtx.fillStyle = '#ffffff';
  outCtx.fillRect(0, 0, rectPx.w, rectPx.h);
  // Pull the source region
  const srcImg = srcCtx.getImageData(rectPx.x, rectPx.y, rectPx.w, rectPx.h);
  const d = srcImg.data;
  for (let i = 0; i < d.length; i += 4) {
    const gray = Math.round(0.2126 * d[i] + 0.7152 * d[i + 1] + 0.0722 * d[i + 2]);
    d[i] = gray;
    d[i + 1] = gray;
    d[i + 2] = gray;
    // keep alpha as-is
  }
  outCtx.putImageData(srcImg, 0, 0);
  return out;
}

async function main() {
  const pdfPath = process.argv[2] || DEFAULT_PDF;
  const pageNumber = Number(process.argv[3] || 1);

  if (!fs.existsSync(pdfPath)) {
    console.error(`[extract-carrier-logos] PDF not found: ${pdfPath}`);
    process.exit(1);
  }

  const { ctx, pageHPt } = await renderPage(pdfPath, pageNumber, SCALE);
  const logWidth = Math.round(PAGE_W_PT * SCALE);
  const logHeight = Math.round(PAGE_H_PT * SCALE);
  console.log(`[extract-carrier-logos] rendered ${logWidth}×${logHeight}px @ scale=${SCALE}`);

  for (const logo of LOGOS) {
    const padded = padRect(logo.rectPt, PAD_PT);
    const px = pdfRectToPx(padded, pageHPt || PAGE_H_PT, SCALE);
    const canvas = cropAndGrayscale(ctx, px);
    const buf = canvas.toBuffer('image/png');
    fs.writeFileSync(logo.out, buf);
    console.log(`[extract-carrier-logos] ${logo.name}: ${px.w}×${px.h}px  ${buf.length}B  → ${logo.out}`);
  }
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
