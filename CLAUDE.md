# CLAUDE.md — Quick Filter for TikTok Seller

> **LAZY-LOAD GUIDE:** อ่าน Zone 1 ก่อนเสมอ (40 บรรทัด) → เพียงพอสำหรับงานส่วนใหญ่
> Zone 2+ อ่านเฉพาะเมื่อต้องการรายละเอียดเฉพาะฟีเจอร์

---

## 🔴 ZONE 1 — Orientation (อ่านทุกครั้ง)

### โปรเจคคืออะไร

Chrome Extension (Manifest V3) ฝัง floating widget ในหน้า TikTok Seller Center และ Shopee Seller Center
ให้พิมพ์ฉลาก 1000+ ใบได้โดยไม่ติดปัญหา pagination/clear-selection ของ TikTok + แปะ alias overlay บน PDF

### ไฟล์หลัก (อัปเดตล่าสุด)

| ไฟล์ | หน้าที่ | บรรทัด |
|------|--------|--------|
| [content.js](content.js) | Logic ทั้งหมด: fetch hook, scan, print, UI, planning | ~10,500 |
| [content.css](content.css) | Style: widget, modal, plan panel, cards | ~3,900 |
| [manifest.json](manifest.json) | Manifest V3, content scripts, permissions | ~30 |
| [asset-bridge.js](asset-bridge.js) | ส่ง font/asset URL จาก ISOLATED → MAIN world | ~10 |
| vendor/ | pdf-lib, fontkit, Sarabun-Bold.ttf (ไม่แก้) | — |

### กฎสำคัญ (ห้ามละเมิด)

- **อย่ายิง print API โดยไม่มี confirmation** — TikTok mark labels ทันที + เสีย quota
- **อย่า bypass `applyCarrierFilter`** — ต้องเคารพ filter ที่ user toggle ไว้
- **อย่า hardcode token/session params** — ต้อง capture สดทุก session
- **อย่า mutate state โดยตรง** — ใช้ immutable pattern เสมอ (`{ ...obj, field: value }`)

### กฎการอัปเดต CLAUDE.md (บังคับ)

> **ทุกครั้งที่เพิ่ม ลบ หรือแก้ไขฟีเจอร์ ต้องอัปเดต CLAUDE.md ด้วยเสมอ**
>
> อัปเดตสิ่งต่อไปนี้ให้ครบ:
> 1. Feature Inventory (Zone 2) — เพิ่ม/ลบ/แก้รายการ
> 2. บรรทัดไฟล์ (Zone 1 ตาราง) — ถ้า diff > 200 บรรทัด
> 3. localStorage keys — ถ้าเพิ่ม/เปลี่ยน key
> 4. Architecture notes — ถ้าเปลี่ยน pattern/flow

### Reload workflow (ทุกครั้งหลังแก้ไฟล์)

1. `chrome://extensions/` → กด **Reload** ที่ extension
2. หน้า TikTok → **hard refresh** `Cmd+Shift+R`

⚠️ refresh page อย่างเดียวไม่พอ — Chrome cache extension files

---

## 🟡 ZONE 2 — Feature Inventory (อ่านเมื่อต้องการรู้ว่ามีอะไร)

### ฟีเจอร์ที่ใช้งานได้แล้ว (TikTok)

| ฟีเจอร์ | Entry point (content.js) | หมายเหตุ |
|---------|--------------------------|---------|
| Fetch hook | `const _origFetch = window.fetch` (บรรทัดต้นๆ) | ดักจับ API ทุก call |
| Scan labels page | `scanLabelsPage()` | API + DOM fallback |
| Scan order page | `scanOrderPage()` / `awaitBodyTemplate()` | capture body template |
| Print labels | `printIds()` → `runChunkedExport()` | chunk + PDF overlay |
| PDF alias overlay | `overlayAliasOnPdf()` | Sarabun, opacity 0.4 |
| Product alias | `getAlias()` / `setAlias()` | localStorage, platform-prefixed |
| Variant alias | `getVariantInfo()` / `setVariantInfo()` | same key pattern |
| Carrier filter | `applyCarrierFilter()` | toggle per carrier |
| Pre-order filter | `passesPreOrder()` | |
| Label status filter | `labelStatusMatches()` | not_printed / printed / all |
| Print history | `openHistoryModal()` | localStorage, cap 200, 30 days |
| Planning mode | `openPlanningPanel()` | drag-drop, multi-worker columns |
| Plan: print-all parallel | print-all handler → `Promise.all` | single modal, all workers parallel |
| Plan: auto-split | `applyAutoSplit()` | even or by-SKU |
| Plan: session persist | `debouncedSavePlan()` | `qf_plan_snapshots_v1` |
| Chunk plan modal | `showChunkPlanModal()` | title param, แบ่งไฟล์ |
| Result modal | `showChunkedResult()` | minimize-bubble + guard |
| PDF Template Builder | `openPdfTemplateEditor()` | Phase 3+4a done |
| Layout presets | `LAYOUT_PRESETS` | 5 presets, verifyPresets() |
| System element overrides | `applyPdfTemplate()` | barcode/QR untouchable |
| Picking List PDF | `buildPickingListPdf()` | optional, per chunk |
| CSV export | `downloadCsv()` | history + daily summary + order history |
| Order claim lookup | `queryOrderHistory(orderId)` | search box ใน history modal |

### ฟีเจอร์ที่ทำได้บางส่วน (Shopee)

| ฟีเจอร์ | สถานะ |
|---------|-------|
| Scan + group + alias + variant | ✅ ใช้งานได้ |
| Print waybill | ❌ disabled (toast แนะนำปุ่ม Shopee) |
| Print history | ❌ ยังไม่รองรับ |
| Planning mode | ❌ disabled (toast) |

### localStorage Keys ทั้งหมด

| Key | ข้อมูล | ขนาด/Cleanup |
|-----|-------|-------------|
| `qf_product_aliases_v1` | `{tk:productId → alias}` | ไม่มี |
| `qf_variant_aliases_v1` | `{tk:productId:skuId → {alias, replace}}` | ไม่มี |
| `qf_workers_v1` | worker list `[{id, name, icon}]` | ไม่มี |
| `qf_teams_v1` | team list | ไม่มี |
| `qf_print_history_v1` | print history array | cap 200, >30d drop |
| `qf_plan_snapshots_v1` | plan snapshots รายวัน | trim >30 วัน |
| `qf_last_picking_list_v1` | checkbox state 'true'/'false' | 5 bytes |
| `qf_pdf_templates_v1` | PDF template configs | per-template |
| `qf_pdf_templates_seeded_v1` | flag: default templates seeded | bool |
| `qf_label_template_v1` | (planned Phase 2) | — |
| `qf_label_template_lastcarrier_v1` | (planned Phase 2) | — |

### IndexedDB

| DB | Store | Schema | Cleanup |
|----|-------|--------|---------|
| `qf_order_history` | `orders` | `{orderId, fulfillUnitId, ts, date, carrier, assigneeKind, teamId, teamName, teamSnapshot, workerId, workerName}` | auto-drop >90 วัน ตอน save |

---

## 🟢 ZONE 3 — Architecture (อ่านเมื่อ implement/debug)

### Two-world setup (CRITICAL)

| World | ไฟล์ | เข้าถึงอะไรได้ | ห้าม |
|-------|------|--------------|------|
| MAIN | content.js, vendor/*.js | `window.fetch`, DOM | `chrome.runtime.*` |
| ISOLATED | asset-bridge.js | `chrome.runtime.getURL()` | `window.fetch` hook |

Communication: ISOLATED → MAIN ผ่าน `window.postMessage` (ใช้แค่ font URL)

`run_at: "document_start"` — content.js ต้องรันก่อน TikTok wrap fetch

### Fetch hook pattern

```javascript
const _origFetch = window.fetch;        // บันทึกก่อน TikTok wrap
window.fetch = async function(...args) { /* hook */ };
```

- ใช้ `_origFetch` เมื่อยิง API เอง → ไม่วนซ้ำ + ไม่โดน TikTok sign
- TikTok wrap fetch อีกชั้นแต่ chain ยังเรียก hook ของเราอยู่

### Print flow (Labels page)

```
Scan → state.records (Map<fulfillUnitId, record>)
     → คลิก card / print-all
     → confirmInline / showChunkPlanModal (ครั้งเดียว)
     → printIds() / printPlanColumn()
     → POST /api/v1/fulfillment/shipping_doc/generate
     → รับ doc_url → fetch PDF → overlayAliasOnPdf() → blob URL
     → window.open() / runChunkedExport() → showChunkedResult()
```

### Planning mode flow

```
openPlanningPanel(session)
  → newPlanningSession()           // filter IDs ตาม labelStatusFilter
  → drag cards → assign to columns
  → 🖨 print-all button:
      → detect multiSku across ALL columns
      → showChunkPlanModal (ONE modal, title: "พิมพ์ทุกคน · กาย, เก่, เกม")
      → Promise.all(columns.map → printPlanColumn(session, cid, ids, renderFn, { plan }))
      → แต่ละ renderFn merge เฉพาะ column ตัวเองเข้า sharedSession
```

**`newPlanningSession()` — labelStatus filter logic:**
- `labelStatusFilter === 'not_printed'` → records มีแต่ status 30 อยู่แล้ว → ไม่ filter เพิ่ม
- `labelStatusFilter === 'all'` → กรอง status 50 ออก (ไม่ให้แอบ reprint)
- `labelStatusFilter === 'printed'` → เอาทุก record (user ตั้งใจ reprint)

**`printPlanColumn(session, workerId, ids, renderFn, opts)`:**
- `opts.plan` → ข้าม modal, ใช้ plan ที่ส่งมาโดยตรง (สำหรับ print-all)
- `opts.reprint` → ไม่ filter printedIds

### confirmInline — danger vs normal

```javascript
confirmInline(title, label, isDanger = false)
```

- `isDanger = true` → ปุ่มใช้ `.qf-btn-danger` (dark red) — ลบ, ล้าง, ละทิ้ง
- `isDanger = false` (default) → ปุ่มใช้ `.qf-btn-confirm` (TikTok brand red) — พิมพ์, ยืนยัน

### label_status values

| ค่า | ความหมาย | constant |
|-----|---------|---------|
| 30 | ยังไม่พิมพ์ | `LABEL_STATUS_NOT_PRINTED` |
| 50 | พิมพ์แล้ว | `LABEL_STATUS_PRINTED` |

### Platform detection

```javascript
isTikTok()    // seller-th.tiktok.com
isShopee()    // seller.shopee.co.th
isLabelsPage()  // TikTok /shipment/labels OR any Shopee /portal/sale
isOrderPage()   // TikTok /order OR Shopee /portal/sale/order
```

Theme: Shopee ใช้ `html[data-qf-theme="shopee"]` → CSS override เป็น orange (#ee4d2d)

---

## 🔵 ZONE 4 — API Reference (อ่านเมื่อแก้ API/scan)

### TikTok Endpoints (fragile — เปลี่ยนได้ทุกเมื่อ)

| Endpoint | Method | ใช้กับ | Key |
|----------|--------|--------|-----|
| `/order/list` (URL random) | POST | Order scan | capture สดผ่าน hook |
| `/api/fulfillment/package/list` | POST | Labels scan | `seller_packages_list` |
| `/api/v1/fulfillment/shipping_doc/generate` | POST | สร้าง PDF | `fulfill_unit_id_list[]` |
| `/api/v1/fulfillment/doc/print_status/verify` | POST | (optional) | `fulfill_unit_id[]` |

### Shopee Endpoints (XHR, not fetch)

| Endpoint | Method | ใช้กับ |
|----------|--------|--------|
| `/api/v3/order/search_order_list_index` | POST | order_id list |
| `/api/v3/order/get_order_list_card_list` | POST | รายละเอียด + SKU |
| `/api/v3/logistics/can_print_waybill` | POST | (TODO) |

### API Record Structure (Labels normalize)

```
API (snake_case) → normalizeApiRecord() → processLabelRecord()
{
  fulfill_unit_id,
  label_module: { label_status, batch_id },
  sku_module: [{ product_id, sku_id, product_name, quantity, product_image }],
  delivery_module: { shipment_provider_info: { id, name, icon_url } }
}
```

---

## ⚪ ZONE 5 — Feature Deep Dives (อ่านเฉพาะฟีเจอร์ที่กำลังทำ)

### Print Result Modal (showChunkedResult)

- แสดงทีละ modal ต่อ `runChunkedExport` call
- ปิดด้วย ×/ปิด/ย่อ เท่านั้น (ไม่มี backdrop click — blob URL หายง่าย)
- `downloaded` flag per chunk → ถ้าปิดโดยยังไม่ download → confirm dialog
- "ย่อ" → `.qf-chunked-bubble` (bottom-right pill) → กด re-expand

### Print History

- `localStorage.qf_print_history_v1` — newest first, cap 200, auto-drop >30 days
- เก็บ fulfillUnitId arrays ไม่ใช่ PDF bytes — re-download replay ผ่าน generate API
- `historyMeta = null` → skip save (สำหรับ re-download ที่มาจาก history เอง)
- เปิดจาก `⏱` button ใน header (TikTok labels page only)

### PDF Template Builder (Phase status)

| Phase | สถานะ | สรุป |
|-------|-------|------|
| 1 — Calibration | ✅ DONE | `tools/inspect-pdf.js`, `jnt_layout.json` |
| 2 — Canvas editor | 🔲 ยังไม่เริ่ม | draggable overlays |
| 3 — System overrides | 🟡 IN PROGRESS | resize/move carrier elements |
| 4a — Presets | ✅ DONE | 5 presets, picker modal |
| 4b — Multi-carrier | 🔲 ยังไม่เริ่ม | Flash, Kerry, SPX, ไปรษณีย์ |

**Zone model 3 ชั้น:**
- 🔒 LOCKED — barcode, QR, sortCode, orderId, trackingNumber, serviceType, codLabel — ห้ามแตะ
- 🔄 SHRINKABLE — TikTok logo, carrier logo, skuTable, addressBlock
- 🎨 CUSTOMIZABLE — shop logo, thank-you text, LINE QR, alias watermark

**Safety guard:**
```javascript
NEVER_OVERRIDE_KEYS = { barcodeMain, barcodeLeft, barcodeRight, qrCode, qr }
// hard-blocked ใน validateOverrides() + applyPdfTemplate()
```

**5 Layout presets:** `original`, `slim-header`, `review-promo`, `compact`, `branded`
→ constant `LAYOUT_PRESETS`, verified by `verifyPresets()` ตอน init

### Cross-platform Alias Isolation

Keys prefix: `tk:productId` / `sp:productId` — ป้องกัน collision ระหว่าง platform
TikTok fallback อ่าน unprefixed key เดิมด้วย (backward compat)

### Shopee Record Shape (3 รูปแบบ)

1. `card.package_card` — tab 300 (ที่ต้องจัดส่ง)
2. `card.order_card` — บาง detail page
3. `card.package_level_order_card` — tab 100 (ทั้งหมด, multi-package)

Field: `inner_item_ext_info.model_id` → skuId (null ถ้าไม่มี, ห้าม fallback เป็น item_id)

---

## ⚫ ZONE 6 — Debugging & Maintenance (อ่านเมื่อ debug)

### Quick state inspection

```javascript
const s = window.__qfState
console.table([...s.products.values()].map(p => ({
  name: p.productName.slice(0,30),
  single: p.fulfillUnitIdsSingle.size,
  multi: p.fulfillUnitIdsMulti.size,
})))
console.log('records:', s.records.size, 'carriers:', s.carriers.size)
console.log('labelStatusFilter:', s.labelStatusFilter)
// planning session:
window.__qfPlanningSession?.()
```

### Test print without API

```javascript
const _f = window.fetch, _o = window.open
window.fetch = async (url, opts) => { console.log('FETCH:', url, opts?.body); return new Response('{}'); }
window.open = (url) => { console.log('OPEN:', url); return null; }
// คืนค่า:
window.fetch = _f; window.open = _o
```

### agent-browser (Claude debug)

```bash
agent-browser --cdp 9222 eval 'expression'
agent-browser --cdp 9222 snapshot -i
```

### Maintenance Scenarios

**A. Order scan → "API ตอบกลับ error" / "NO TEMPLATE"**
1. Body ไม่ถูกจับ → ดู `awaitBodyTemplate()`, trigger ผ่าน pagination click
2. API path เปลี่ยน → DevTools Network → update regex ใน hook
3. Required field เพิ่ม → เปรียบเทียบ body ที่ TikTok ส่งกับของเรา

**B. Labels page → DOM fallback ตลอด**
1. API path เปลี่ยน → หา `_labelsApiUrl` ใน content.js → update regex
2. Body schema เปลี่ยน → capture body ใหม่จาก Network tab
3. Response key เปลี่ยน → เพิ่ม fallback ใน `getApiList()`

**C. Print API → code != 0**
1. IDs หมดอายุ (order ถูกลบ) → scan ใหม่
2. Required field เปลี่ยน (เช่น `print_source`) → copy body จาก Network ขณะกดพิมพ์บน TikTok UI

**D. PDF overlay ไม่ขึ้น / เป็นกล่องสี่เหลี่ยม**
1. `window.fontkit` undefined → เช็ค manifest.json ลำดับ vendor scripts
2. Font URL หาไม่เจอ → เช็ค `state.fontUrl`, `web_accessible_resources`
3. ข้อความ parameters: `size = Math.min(height * 0.05, 22)`, `y = 6`, `opacity: 0.4`

**E. Carrier logo ไม่แสดง**
→ แก้ field path ใน `processLabelRecord()`: `rec.shippingProviderInfo || rec.deliveryInfo?.shippingProvider`

**F. Weird combo grouping ผิด**
→ signature key คือ `productId:quantity` — ถ้าต้องการ group ไม่สนใจ qty ให้แก้เป็น `productId` only

**G. Alias หาย**
→ เช็ค `qf_product_aliases_v1` ใน localStorage, ตรวจ platform prefix `tk:`/`sp:`

**H. Printed card หายไปแทนที่จะเป็นสีเทา (qf-done visual)**
- ภายใต้ filter "ยังไม่พิมพ์" (default): หลังพิมพ์แล้ว `rec.labelStatus = 50` → `passesLabelStatus` ตก → `carrierFilteredSize` ได้ 0 → ถูก filter ออกจาก grid
- ✅ Fix: ใช้ `carrierFilteredSizeIgnoreLabel()` เป็น `_totalVisible` → render predicate: `_count > 0 || (_totalVisible > 0 && isLabelsDone(...))` → card ยังแสดงเป็นสีเทา
- Sort: done cards ถูก push ไปท้ายลิสต์ (`a._count === 0` → bottom)
- ทั้ง products, variant badges, qty chips, combos ใช้ pattern เดียวกัน

**Done state = session-local เท่านั้น** (สำคัญสำหรับ reprint):
- `isLabelsDone(idSet)` ดูแค่ `state.printedUnitIds.has(id)` — **ไม่** ดู `rec.labelStatus === 50` จาก server
- qty chip `allDone` ก็ใช้ `totalIds.every(id => state.printedUnitIds.has(id))` เช่นเดียวกัน
- เหตุผล: user อาจ scan สินค้าที่พิมพ์ไปแล้ว (ฉลากหาย/เสีย) เพื่อจะ reprint → card ต้องคลิกพิมพ์ได้ทันที ไม่ต้องผ่าน unmark modal
- scan ใหม่ → `printedUnitIds.clear()` (line 1298) → done state reset → ทุกการ์ด enabled
- พิมพ์ผ่าน extension → `printedUnitIds.add(id)` → disabled (greyed ✓) ใน session นั้น

### Conventions

- ภาษา: UI strings + comments = ไทย, code identifiers = English
- Naming: state = camelCase, API fields = snake_case
- No build step: แก้ไฟล์ → reload extension → ใช้งานได้เลย
- No npm/package.json: vendor libraries pin version แล้ว
- Testing: manual ผ่าน DevTools + agent-browser

### Run PDF Calibration ใหม่

```bash
npm install pdf2json  # ครั้งแรก
node tools/inspect-pdf.js path/to/label.pdf .claude/samples/<carrier>_layout.json
```
