#!/usr/bin/env node
/*
 * tools/test-mask-tiktok-logo.js — Verify white-fill mask placement for TikTok Shop top logo
 *
 * Pipeline:
 *   1. Load .claude/example.pdf
 *   2. Draw white rectangle over TikTok Shop top logo zone
 *   3. Save to /tmp/masked.pdf
 *   4. Render page 1 to PNG at 2x via pdfjs → vendor/jnt_masked_preview.png
 *
 * Iterate coords inside MASK_RECT until the preview PNG shows:
 *   - TikTok Shop text + music-note glyph fully covered (no remnants)
 *   - "V" stray glyph covered
 *   - J&T Express logo untouched (starts ~x_pdf=100)
 *   - No spill into "EZ" sortCode or "698" or "เกม" at top-right
 */
'use strict';

const fs = require('fs');
const path = require('path');
const { PDFDocument, rgb, degrees } = require('pdf-lib');
const fontkit = require('@pdf-lib/fontkit');

const ROOT = path.resolve(__dirname, '..');
const SRC_PDF = path.join(ROOT, '.claude', 'example.pdf');
const FONT_PATH = path.join(ROOT, 'vendor', 'Sarabun-Bold.ttf');
const OUT_PDF = path.join('/tmp', 'jnt_masked.pdf');
const OUT_PNG = path.join(ROOT, 'vendor', 'jnt_masked_preview.png');
const OUT_JPG = path.join(ROOT, 'vendor', 'jnt_masked_preview.jpg');

// PDF-points, bottom-left origin. Page is 298×420pt (A6).
// Masks (bottom only per user spec):
//   - TikTok Shop FOOTER logo + "Order ID: 5836..." + horizontal separator line
//     Visually ~y_px 700..755 (in 840px image) → y_pdf ≈ 42..70
//     Horizontal line sits just above footer → include it → top of rect ≈ y_pdf 74
// Iteration 2: shift mask UP to actually hit the footer row.
//   Image-space landmarks (596×840, SCALE=2):
//     - Horizontal separator line   ≈ y_px 695  →  y_pdf (840-695)/2 = 72.5
//     - "TikTok Shop" footer logo   ≈ y_px 710-730 → y_pdf 55-65
//     - "Order ID: 5836..." text    ≈ y_pdf 63.47 (t_116)
//     - "Qty Total: 1" text above   ≈ y_pdf 78.94 (t_114) — DO NOT cover
//   Rect: y = 53, h = 22 → covers y_pdf 53..75 (footer + line, leaves Qty Total safe)
// Iteration 3: pdf2json y in jnt_layout.json != PDF-native y (offset ~10-12pt).
// Calibrated from iter 2 render: TikTok footer sits at image y 735-760 on 596×840,
//   → y_pdf = (840-760)/2 = 40  to  (840-735)/2 = 52.5
// Horizontal separator line just above footer ≈ y_pdf 55-62
// Qty Total row bottom ≈ y_pdf 80 (must NOT cover)
// Rect: y=38, h=30 → covers y_pdf 38..68 (footer + line, clear of Qty Total)
// Iteration 4 — pixel-calibrated via PIXEL_SCAN on unmasked baseline.
// Landmarks (596×840 render, SCALE=2, PAGE_H=420pt):
//   y_img=708  Qty Total last dark row     (y_pdf 66.0)
//   y_img=714  horizontal line #1          (y_pdf 63.0)
//   y_img=729-745  footer + Order ID       (y_pdf 47.5–55.5)
//   y_img=759  horizontal line #2          (y_pdf 40.5)
//   y_img=762  alias "ครีม..." first dark  (y_pdf 39.0)
// Mask band: y_img 712..760 → y_pdf 40..64, buffer 2px above (Qty) and 2px below (alias).
// Iteration 5 — Canva mockup spec (pixel-scanned multi-rect):
//   Match: TikTok TOP logo + barcode + Order ID footer retained;
//   Remove: ALL 10 horizontal lines, vertical OCR digit columns, TikTok Shop FOOTER logo.
// Landmarks (pixel-scanned, 596×840 @ SCALE=2):
//   horizontal lines (y_img → y_pdf):
//     54→393, 190→325, 329→255.5, 462→189, 507→166.5, 549→145.5,
//     628→106, 687-688→76, 714→63, 759→40.5
//   vertical digit columns: x_pdf 6-14 (L), 281-289 (R). Card borders at 16.5/278.5 preserved.
//   footer row y_img 720-758: TikTok Shop logo x_pdf 7-77; Order ID x_pdf 208-292.
//   Top barcode starts at y_img=65 (y_pdf=387.5) — safe from hline-393 mask.
// Iteration 6 — user: "ลบแค่ hr line zone ล่าง ไม่ใช่ทั้งหมด และเลขแนวตั้ง
//   ลบแค่ 2/3 ตามตัวอย่าง ในแต่ละ col"
//   → keep 7 upper hr separators (card structure); remove only 3 bottom-zone hr lines
//   → vertical digit columns: mask bottom 2/3 only (220pt of 330pt), keep top 1/3
//     near barcode (y_pdf 282..392). Overlay text lands in the masked 2/3.
const MASK_RECTS = process.env.NO_MASK ? [] : [
  // Vertical OCR digit columns — bottom 2/3 masked (220pt), top 1/3 (110pt) revealed.
  // Full height=330pt (y_pdf 62..392). Mask y_pdf 62..282, visible y_pdf 282..392.
  { x: 0,   y: 62,    w: 15,  h: 220, label: 'vcol-L-2/3' },
  { x: 285, y: 62,    w: 13,  h: 220, label: 'vcol-R-2/3' },
  // TikTok Shop FOOTER logo (left portion; preserve Order ID right of x_pdf=95).
  { x: 0,   y: 41.5,  w: 95,  h: 20,  label: 'tt-footer' },
  // Bottom-zone horizontal lines only (3 of 10): above Qty Total, above/below footer.
  { x: 0,   y: 75,    w: 298, h: 3,   label: 'h-76'  },
  { x: 0,   y: 62,    w: 298, h: 2,   label: 'h-63'  },
  { x: 0,   y: 39.5,  w: 298, h: 2,   label: 'h-41'  },
];

// Marketing text overlay config.
// MARKETING_TEXT: vertical text drawn rotated 90° inside each masked column.
// FOOTER_CUSTOM_TEXT: horizontal text in the masked TikTok Shop footer logo zone.
const MARKETING_TEXT   = process.env.MARKETING_TEXT   || 'กรุณาถ่ายรูปก่อนเปิดกล่องพัสดุ';
const FOOTER_CUSTOM    = process.env.FOOTER_CUSTOM     || 'shopname.th';
const ORANGE = rgb(1, 0.40, 0);   // #ff6600
const COL_FONT_SIZE    = 7;
const FOOTER_FONT_SIZE = 7;

async function maskPdf() {
  const pdfBytes = fs.readFileSync(SRC_PDF);
  const pdf = await PDFDocument.load(pdfBytes);
  pdf.registerFontkit(fontkit);
  const fontBytes = fs.readFileSync(FONT_PATH);
  const font = await pdf.embedFont(fontBytes, { subset: true });
  const page = pdf.getPage(0);

  // 1. Draw white mask rects
  for (const r of MASK_RECTS) {
    page.drawRectangle({
      x: r.x, y: r.y, width: r.w, height: r.h,
      color: rgb(1, 1, 1), borderWidth: 0,
    });
  }

  if (!process.env.NO_OVERLAY) {
    // 2. Vertical marketing text — left column (rotated 90°, centered in x=0..15)
    // Column center x = 7.5. Text baseline runs upward (rotation=90°).
    // Place at bottom of masked zone (y=62), text extends upward.
    const colCxL = 7.5;
    const colCxR = 285 + 6.5; // center of right column (285..298)
    const colYStart = 65;     // just above the h-63 line

    page.drawText(MARKETING_TEXT, {
      x: colCxL,
      y: colYStart,
      size: COL_FONT_SIZE,
      font,
      color: ORANGE,
      rotate: degrees(90),
      opacity: 0.85,
    });
    page.drawText(MARKETING_TEXT, {
      x: colCxR,
      y: colYStart,
      size: COL_FONT_SIZE,
      font,
      color: ORANGE,
      rotate: degrees(90),
      opacity: 0.85,
    });

    // 3. Footer custom text — centered in masked footer logo zone (x=0..95, y=41.5..61.5)
    const footerY = 44;
    page.drawText(FOOTER_CUSTOM, {
      x: 4,
      y: footerY,
      size: FOOTER_FONT_SIZE,
      font,
      color: ORANGE,
      opacity: 0.9,
    });
  }

  const out = await pdf.save();
  fs.writeFileSync(OUT_PDF, out);
  console.log(`[mask] ${OUT_PDF} (${out.length} bytes)`);
}

async function renderToPng() {
  const pdfjsLib = await import('pdfjs-dist/legacy/build/pdf.mjs');
  const { createCanvas } = require('@napi-rs/canvas');
  const data = new Uint8Array(fs.readFileSync(OUT_PDF));
  const loadingTask = pdfjsLib.getDocument({
    data, disableWorker: true, useSystemFonts: false, verbosity: 0,
  });
  const doc = await loadingTask.promise;
  const page = await doc.getPage(1);
  const viewport = page.getViewport({ scale: 2 });
  const w = Math.round(viewport.width);
  const h = Math.round(viewport.height);
  const canvas = createCanvas(w, h);
  const ctx = canvas.getContext('2d');
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, w, h);
  await page.render({ canvas, viewport }).promise;

  // Grayscale (matches jnt_sample_preview.png pipeline so visual compares align)
  const img = ctx.getImageData(0, 0, w, h);
  const d = img.data;
  for (let i = 0; i < d.length; i += 4) {
    const g = Math.round(0.2126 * d[i] + 0.7152 * d[i + 1] + 0.0722 * d[i + 2]);
    d[i] = g; d[i + 1] = g; d[i + 2] = g;
  }
  ctx.putImageData(img, 0, 0);

  // Debug overlay: show each mask rect with translucent red fill
  const SCALE = 2, PAGE_H = 420;
  const DRAW_DEBUG = !process.env.NO_DEBUG_OVERLAY;
  for (const r of (DRAW_DEBUG ? MASK_RECTS : [])) {
    const rx = r.x * SCALE;
    const ry = (PAGE_H - r.y - r.h) * SCALE;
    const rw = r.w * SCALE;
    const rh = r.h * SCALE;
    ctx.fillStyle = 'rgba(255, 0, 0, 0.25)';
    ctx.fillRect(rx, ry, rw, rh);
    ctx.strokeStyle = '#ff0000';
    ctx.lineWidth = 2;
    ctx.strokeRect(rx + 1, ry + 1, rw - 2, rh - 2);
    ctx.fillStyle = '#ff0000';
    ctx.font = 'bold 11px sans-serif';
    ctx.fillText(`MASK:${r.label || ''}`, rx + 6, ry + 14);
  }

  if (process.env.PIXEL_SCAN) {
    const rawCtx = canvas.getContext('2d');
    const scan = rawCtx.getImageData(0, 0, w, h).data;
    const MODE = process.env.SCAN_MODE || 'horizontal-lines';

    if (MODE === 'horizontal-lines') {
      // Find full-width-ish dark horizontal lines (dark_px > threshold, spans most of width)
      const yStart = parseInt(process.env.Y_START || '80', 10);
      const yEnd = parseInt(process.env.Y_END || '820', 10);
      const widthThreshold = parseInt(process.env.W_THRESHOLD || '400', 10); // px
      console.log(`[hlines] y_img  dark_px  y_pdf  (threshold=${widthThreshold}px, width=${w})`);
      for (let yy = yStart; yy <= yEnd; yy++) {
        let dark = 0;
        for (let xx = 0; xx < w; xx++) {
          const off = (yy * w + xx) * 4;
          if (scan[off] < 128) dark++;
        }
        if (dark >= widthThreshold) {
          const yPdf = ((h - yy) / 2).toFixed(2);
          console.log(`  y_img=${yy}  dark=${String(dark).padStart(4)}  y_pdf=${yPdf}`);
        }
      }
    } else if (MODE === 'vertical-columns') {
      // Find dark columns near left/right edges — the vertical digit strips
      const xStart = parseInt(process.env.X_START || '0', 10);
      const xEnd = parseInt(process.env.X_END || String(w), 10);
      const yTop = parseInt(process.env.Y_TOP || '80', 10);
      const yBot = parseInt(process.env.Y_BOT || '700', 10);
      console.log(`[vcols] x_img  dark_px  x_pdf  (y range ${yTop}..${yBot})`);
      for (let xx = xStart; xx < xEnd; xx++) {
        let dark = 0;
        for (let yy = yTop; yy <= yBot; yy++) {
          const off = (yy * w + xx) * 4;
          if (scan[off] < 128) dark++;
        }
        if (dark > 0) {
          const xPdf = (xx / 2).toFixed(2);
          console.log(`  x_img=${xx}  dark=${String(dark).padStart(4)}  x_pdf=${xPdf}`);
        }
      }
    } else if (MODE === 'row') {
      // Per-row report (legacy)
      const yStart = parseInt(process.env.Y_START || '640', 10);
      const yEnd = parseInt(process.env.Y_END || '820', 10);
      console.log('[row] y_img  dark_px  y_pdf');
      for (let yy = yStart; yy <= yEnd; yy++) {
        let dark = 0;
        for (let xx = 20; xx < w - 20; xx++) {
          const off = (yy * w + xx) * 4;
          if (scan[off] < 128) dark++;
        }
        const yPdf = ((h - yy) / 2).toFixed(2);
        if (dark > 0 || yy % 4 === 0) {
          console.log(`  y_img=${yy}  dark=${String(dark).padStart(4)}  y_pdf=${yPdf}`);
        }
      }
    } else if (MODE === 'footer-row-xscan') {
      // For a given y_img range, show dark pixels per x — to locate TikTok Shop footer vs Order ID split.
      const yTop = parseInt(process.env.Y_TOP || '729', 10);
      const yBot = parseInt(process.env.Y_BOT || '745', 10);
      console.log(`[xscan] x_img  dark_px  x_pdf  (y ${yTop}..${yBot})`);
      for (let xx = 0; xx < w; xx++) {
        let dark = 0;
        for (let yy = yTop; yy <= yBot; yy++) {
          const off = (yy * w + xx) * 4;
          if (scan[off] < 128) dark++;
        }
        if (dark > 0) {
          const xPdf = (xx / 2).toFixed(2);
          console.log(`  x_img=${xx}  dark=${String(dark).padStart(4)}  x_pdf=${xPdf}`);
        }
      }
    }
  }
  fs.writeFileSync(OUT_PNG, canvas.toBuffer('image/png'));
  console.log(`[png] ${OUT_PNG} (${canvas.width}×${canvas.height})`);
  fs.writeFileSync(OUT_JPG, await canvas.encode('jpeg', 92));
  console.log(`[jpg] ${OUT_JPG} (${canvas.width}×${canvas.height}, q=92)`);
}

(async () => {
  await maskPdf();
  await renderToPng();
})().catch(e => { console.error(e); process.exit(1); });
