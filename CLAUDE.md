# CLAUDE.md

คู่มือสำหรับ Claude Code ในการทำงานกับ Quick Filter for TikTok Seller extension

## ภาพรวมโปรเจค

Chrome Extension (Manifest V3) ที่ฝัง widget ลอยในหน้า TikTok Seller Center รองรับ 2 หน้า:

1. **Order page** (`/order`) — กรอง + auto-select ออเดอร์ตามสินค้าก่อนนัดหมายจัดส่ง/พิมพ์
2. **Labels page** (`/shipment/labels`) — สแกนฉลาก แล้วยิง print API ตรง ๆ พร้อมแปะลายน้ำ alias สีดำ opacity 40%

ปัญหาที่แก้:
- TikTok limit ทีละหน้า + clear selection เมื่อเปลี่ยนหน้า → พิมพ์ฉลาก 1000+ ใบลำบาก
- พนักงานต้องตรวจชื่อสินค้าจาก label ลำบาก → แปะ alias เช่น "แดง1" บน PDF

## Architecture

### Two-world setup

Manifest V3 บังคับแยก execution context. โปรเจคนี้ใช้ทั้งสอง:

- **MAIN world** ([content.js](content.js), [vendor/pdf-lib.min.js](vendor/pdf-lib.min.js), [vendor/fontkit.umd.min.js](vendor/fontkit.umd.min.js))
  - เข้าถึง `window.fetch` ของ TikTok ได้ → hook ดักจับ API ได้
  - **ไม่มี** `chrome.runtime.*` API
  - `run_at: "document_start"` — ต้องรันก่อน TikTok wrap fetch
- **ISOLATED world** ([asset-bridge.js](asset-bridge.js))
  - เข้า `chrome.runtime.getURL()` ได้ → หา URL ของไฟล์ extension
  - ส่งไป MAIN world ผ่าน `window.postMessage`
  - ใช้แค่กับ font URL (Sarabun-Bold.ttf)

### Fetch hook pattern (CRITICAL)

[content.js](content.js) บันทึก `window.fetch` ไว้ตั้งแต่ `document_start`:

```javascript
const _origFetch = window.fetch;
window.fetch = async function(...args) { ... };
```

`_origFetch` ใช้สำหรับยิง API เองโดยไม่โดนวน hook ของเราซ้ำ และไม่โดน TikTok's wrapper ที่อาจ sign request

**TikTok wrap fetch ของเราอีกชั้น** (เห็นจาก `window.fetch.toString()` แสดง wrapper ของ TikTok ไม่ใช่ของเรา) — แต่ chain ยังเรียก hook ของเราอยู่

### Print flow (Approach B)

แทนที่จะ tick checkbox + คลิก "พิมพ์เอกสาร" บน UI:

1. Scan ดึง `fulfillUnitId` ทุกฉลาก
2. คลิก card → confirm modal → ยิง `POST /api/v1/fulfillment/shipping_doc/generate` พร้อม `fulfill_unit_id_list[]` (สูงสุด 500/batch)
3. รับ `data.doc_url` (signed URL)
4. ถ้ามี alias → fetch PDF → pdf-lib overlay → blob URL
5. `window.open(url)` → tab ใหม่ พร้อมลายน้ำ

### Print result modal: minimize-bubble + undownloaded-guard

`showChunkedResult` (content.js) shows one modal per `runChunkedExport` call.
Background-click dismiss was removed (ข้อมูล blob URL หายง่าย) so users close
explicitly via × / ปิด / ย่อ. A `downloaded` flag is set per chunk when the
user clicks เปิด / ดาวน์โหลด / ดาวน์โหลดทั้งหมด; closing while any completed
chunk is still `downloaded=false` triggers a confirm dialog pointing to the
history icon. "ย่อ" hides the overlay and spawns `.qf-chunked-bubble`
(fixed, bottom-right, pill-shaped) that shows `กำลังพิมพ์ N/M` or
`พร้อมโหลด N` and re-expands on click.

### Print history (TikTok only)

Stored in `localStorage.qf_print_history_v1` (JSON array, newest first,
cap 200, auto-drop > 30 days old, trim to 50 on quota-exceeded).

Entry schema:
```js
{
  id, timestamp, title, baseFilename, totalLabels,
  chunks: [{label, filename, ids: [fulfillUnitId, ...]}],
  platform: 'tiktok',
}
```

Critical: history stores **fulfillUnitId arrays**, NOT PDF bytes — re-download
replays the chunks through `runChunkedExport(chunks, title, null)` (null meta
to suppress duplicate entries) which re-calls TikTok's generate API. Free
quota: TikTok already accepted the IDs once. If TikTok has since deleted
an order, the chunk's retry button handles the error.

Header shows `⏱` button (TikTok labels page only) with a `.qf-history-badge`
bubble showing count of entries from the last 24h. Click opens
`openHistoryModal()` — day-grouped list with `ดาวน์โหลดใหม่` (primary) and
`🗑` (delete) per entry, plus `ล้างทั้งหมด` footer.

Helpers: `loadHistory`, `saveHistory`, `addHistoryEntry`, `deleteHistoryEntry`,
`clearAllHistory`, `historyRecentCount`, `renderHistoryBadge`. Called from
`runChunkedExport` after `result.allDone()` only if `historyMeta` is truthy
and platform is TikTok.

## ไฟล์สำคัญ

| ไฟล์ | หน้าที่ | บรรทัดประมาณ |
|------|--------|---------|
| [manifest.json](manifest.json) | Manifest V3, content scripts, web_accessible_resources | ~30 |
| [content.js](content.js) | Logic ทั้งหมด — fetch hook, scan, print, UI | ~1300 |
| [content.css](content.css) | Style widget + modal + combo card | ~400 |
| [asset-bridge.js](asset-bridge.js) | ส่ง font URL จาก ISOLATED → MAIN | ~10 |
| [vendor/pdf-lib.min.js](vendor/pdf-lib.min.js) | PDF manipulation (513KB) | UMD |
| [vendor/fontkit.umd.min.js](vendor/fontkit.umd.min.js) | Custom font embedding (741KB) | UMD |
| [vendor/Sarabun-Bold.ttf](vendor/Sarabun-Bold.ttf) | Thai font สำหรับ overlay (88KB) | binary |

## Multi-platform support

Extension runs on both **TikTok Seller Centre** (`seller-th.tiktok.com`)
and **Shopee Seller Centre** (`seller.shopee.co.th`). Detection via:

```javascript
isTikTok() // hostname check
isShopee() // hostname check
isLabelsPage() // true on TikTok labels OR any Shopee /portal/sale
isOrderPage()  // true on TikTok /order OR Shopee /portal/sale/order
```

Shopee uses the **same widget UI**, but theme swaps to orange (#ee4d2d /
#f76b1c gradient) via `html[data-qf-theme="shopee"]` set in `buildWidget`.
All `.qf-*` color classes have `html[data-qf-theme=shopee]` overrides in
content.css.

**Shopee print is currently disabled** — printIds returns early with a
toast pointing user to Shopee's native button. The print waybill API
hasn't been wired yet (next iteration). Scan + group + alias + variant
override + multi-select selection all work identically to TikTok.

### Cross-platform alias isolation

`localStorage.qf_product_aliases_v1` and `qf_variant_aliases_v1` are
keyed with a platform prefix (`tk:` or `sp:`) so the same numeric
productId/skuId on TikTok and Shopee doesn't collide. Read sites on
TikTok fall back to the legacy unprefixed key shape, so aliases saved
before this change continue to work (they stay on the TikTok side).
All writes go through `setAlias()` / `setVariantInfo()` which write
prefixed keys and clear the legacy shadow on TikTok.

### Shopee record → internal shape

- `pushShopeeRecord()` now accepts `pkgExt` alongside `ext` and
  `fulfilment`, and reads `logistics_status` from any of the three to
  drive `labelStatus` (≥ 3 → PRINTED, < 3 → NOT_PRINTED).
- Pre-order detection inspects `order_ext_info.is_pre_order`,
  `package_ext_info.is_pre_order`, `fulfilment_info.is_pre_order`, and
  item-level `inner_item_ext_info.is_pre_order`. If all are absent,
  `pushShopeeRecord` dumps the first record's key list to the console
  once per scan (`_shopeePreOrderProbeLogged`) so the real field path
  can be identified — **needs live verification** to pin down which
  field Shopee actually returns.
- `buildShopeeSkuList()` returns `skuId = null` when `model_id` is
  absent instead of reusing `item_id`; render paths filter
  `v.skuId != null` so the synthetic no-variant entry never shows a
  "null" badge.

### Shopee scan fallback (tab safety)

When the XHR hook hasn't captured the URL/body yet (e.g. single-page
list, no pagination clicked), `scanShopeePage` calls
`detectShopeeActiveTab()` which reads the active tab's Thai label.
Only when the label matches ที่ต้องจัดส่ง/รอจัดส่ง/etc. do we apply
the hardcoded to-ship default body; otherwise scan bails with a
"กรุณาคลิกเปลี่ยนหน้า/แท็บหนึ่งครั้ง" message so the user is never
silently served wrong-tab data.

## Critical TikTok endpoints

ทุกตัวเป็น **fragile** — TikTok เปลี่ยนได้ทุกเมื่อ ดูจุด maintenance ด้านล่าง

| Endpoint | Method | ใช้กับ | Body key สำคัญ |
|----------|--------|------|---------------|
| `/order/list` (URL random) | POST | Order page scan | URL + body จับสดผ่าน hook |
| `/api/fulfillment/package/list` | POST | Labels page scan | `seller_packages_list` (response key) |
| `/api/v1/fulfillment/shipping_doc/generate` | POST | สร้าง PDF | `fulfill_unit_id_list[]`, `content_type_list:[1,2]` |
| `/api/v1/fulfillment/doc/print_status/verify` | POST | (optional) ตรวจสถานะก่อนพิมพ์ | `fulfill_unit_id[]` |

## Critical Shopee endpoints (XHR, not fetch)

| Endpoint | Method | ใช้กับ | หมายเหตุ |
|----------|--------|------|---------|
| `/api/v3/order/search_order_list_index` | POST | ดึง order_id list | response: `data.index_list[]`, total: `data.pagination.total` |
| `/api/v3/order/get_order_list_card_list` | POST | ดึงรายละเอียด+SKU | body uses `order_param_list` หรือ `package_param_list` ตาม tab |
| `/api/v3/logistics/can_print_waybill` | POST | (TODO) ตรวจก่อนพิมพ์ | ยังไม่ใช้ |

### Shopee response shapes (3 รูปแบบ)

`processShopeeRecord` รองรับทั้ง 3:
1. `card.package_card` — tab 300 (ที่ต้องจัดส่ง) — flat single package
2. `card.order_card` — บางหน้า detail
3. `card.package_level_order_card` — tab 100 (ทั้งหมด) — order มี multi-package

Field mapping (snake_case → internal camelCase):
- `inner_item_ext_info.item_id` → `productId`
- `inner_item_ext_info.model_id` → `skuId` (null when absent — don't fall back to item_id)
- `name` → `productName`
- `image` (URI) → `productImageURL` (expand to `https://down-th.img.susercontent.com/file/{uri}_tn`)
- `amount` → `quantity`
- `order_ext_info.logistics_status` ≥ 3 → `LABEL_STATUS_PRINTED` (wired in `pushShopeeRecord`)
- `*.is_pre_order` → `state.preOrderOf` — exact field TBD; see probe log (needs live verification)
- `fulfilment_info.fulfilment_channel_name` → carrier name (ไม่มี icon URL)

### API record structure (Labels)

API คืนค่ามาในรูป snake_case + module structure:
```
{
  fulfill_unit_id, label_module: {label_status, batch_id},
  sku_module: [{product_id, sku_id, product_name, quantity, product_image: {url_list}}],
  delivery_module: {shipment_provider_info: {id, name, icon_url}}
}
```

แปลงเป็น camelCase + flat structure ผ่าน `normalizeApiRecord()` ก่อน feed เข้า `processLabelRecord()` (เพื่อให้ใช้ logic เดียวกับ DOM scan ที่อ่านจาก React fiber)

**`label_status` ค่าที่เห็น:**
- `30` = ยังไม่พิมพ์
- `50` = พิมพ์แล้ว

## Maintenance Guide

### A. Order page scan ขึ้น "API ตอบกลับ error"

อาการ: คลิก Scan แล้วได้ `code=X msg="..."` หรือ `NO TEMPLATE`

**สาเหตุที่เป็นไปได้:**
1. **Body template ไม่ถูกจับ** — TikTok เปลี่ยนวิธีส่ง body (เช่น Request object stream อ่านไม่ทัน)
   - แก้: ดู `awaitBodyTemplate()` รอนานขึ้น หรือ trigger ยิง API ใหม่ผ่าน pagination/tab click
2. **TikTok เปลี่ยน path API** — เดิม `/order/list`
   - แก้: หาผ่าน DevTools Network (filter `Fetch/XHR`, search "order") → update regex ใน hook
3. **TikTok เพิ่ม required field** — body ที่ capture ขาดฟิลด์ใหม่
   - แก้: เห็น msg ใน error → เปรียบเทียบกับ body ที่ TikTok ยิงเอง → เพิ่ม field

### B. Labels page เข้า DOM fallback ตลอด

อาการ: status แสดง "⚠️ ใช้ DOM fallback" หรือ "สแกน หน้า X/Y"

**Debug:**
```javascript
// ใน DevTools Console (หน้า labels)
window.__qfState  // ดู records.size, products.size
```

**สาเหตุ:**
1. **API path เปลี่ยน** — เดิมจับ `/api/fulfillment/package/list`
   - แก้: ใน [content.js](content.js) หา `_labelsApiUrl` → update regex
2. **Body schema เปลี่ยน** — TikTok เพิ่ม required field
   - แก้: capture body ใหม่ผ่าน Network tab → เปรียบเทียบกับที่ extension ส่ง
3. **Response key เปลี่ยน** — เดิมใช้ `seller_packages_list`
   - แก้: เพิ่ม fallback ใน `getApiList()`

### C. Print API ส่ง code != 0

**สาเหตุ:**
1. **`fulfill_unit_id_list` มี ID ที่หมดอายุ** — order ถูกลบ/ยกเลิกหลัง scan
   - แก้: ลบ ID ที่ fail ออกแล้วยิงใหม่ (ปัจจุบันยังไม่ retry — ต้อง scan ใหม่)
2. **Required field เปลี่ยน** เช่น `print_source: 201` → 202
   - แก้: เปิด Network tab ขณะกด "พิมพ์" บน TikTok UI → copy body มาแก้ `printIds()`

### D. PDF overlay ไม่ขึ้น text หรือเป็นกล่องสี่เหลี่ยม

**สาเหตุ:**
1. **fontkit ไม่ load** — `window.fontkit` undefined
   - แก้: เช็ค `manifest.json` มี `vendor/fontkit.umd.min.js` ใน `js[]` ก่อน `content.js`
2. **Font ไฟล์เสีย/หา URL ไม่เจอ**
   - แก้: เช็ค `state.fontUrl` มีค่า, เช็ค `web_accessible_resources` ใน manifest
3. **อักษรไทยขึ้นเป็นช่อง** — ใช้ standard font แทน Sarabun
   - แก้: ตรวจ `pdfDoc.embedFont(fontBytes, {subset: true})` ใช้ fontBytes จริง

ตำแหน่ง/สี/ขนาด text ปรับใน `overlayAliasOnPdf()`:
- `size = Math.min(height * 0.05, 22)` — ขนาดสูงสุด 22pt
- `y = 6` — ห่างจากขอบล่าง 6pt
- `color: rgb(0, 0, 0)` — ดำ
- `opacity: 0.4` — ลายน้ำ 40%

### E. Carrier filter ไม่แสดง logo

**สาเหตุ:** API field name เปลี่ยน

แก้ใน `processLabelRecord()`:
```javascript
const sp = rec.shippingProviderInfo || rec.deliveryInfo?.shippingProvider;
```
เพิ่ม path เพิ่มเติมที่ TikTok อาจใช้

### F. Weird combo grouping ผิด

อาการ: combo เดียวกันแต่แสดงเป็น 2 cards แยก

**สาเหตุ:** signature key ใช้ `productId:quantity` sort ตาม productId
- ถ้า quantity ของ product ใน 2 records ต่างกัน → คนละ combo (ตามสเปค)
- ถ้าต้องการ group ที่ "ชนิดสินค้าเหมือนกัน ไม่สนใจ qty" → แก้ sigKey เป็น `sorted.map(i => i.productId).join('|')`

### G. Alias หาย/ไม่ persist

เก็บใน `localStorage[qf_product_aliases_v1]` keyed by `productId`

ถ้า persist หาย:
- เปลี่ยน LocalStorage key (ผู้ใช้ clear cookies/site data)
- เปลี่ยน productId (TikTok rename product → ID เดิม → alias ยังอยู่ ถูกต้อง)

## Debugging tips

### Quick state inspection (DevTools Console)

```javascript
const s = window.__qfState
console.table([...s.products.values()].map(p => ({
  name: p.productName.slice(0,30),
  single: p.fulfillUnitIdsSingle.size,
  multi: p.fulfillUnitIdsMulti.size,
})))
console.table([...s.carriers.values()])
console.log('combos:', s.weirdCombos.size, 'records:', s.records.size)
```

### Verify API hook captured

```javascript
// ใน console จะมี closure variables — เช็คผ่าน trigger:
document.getElementById('qf-scan-btn').click()
// ถ้า scan ขึ้น "API ไม่จับ" → hook ไม่จับ → ดู Network tab ว่า TikTok ยิงอะไร
```

### Test print without firing API

```javascript
// Stub fetch + window.open → see what would be sent
const _f = window.fetch, _o = window.open
window.fetch = async (url, opts) => { console.log('FETCH:', url, opts?.body); return new Response('{}'); }
window.open = (url) => { console.log('OPEN:', url); return null; }
// คลิก card → จะเห็น log แทนการพิมพ์จริง
// คืนค่าเดิม:
window.fetch = _f; window.open = _o
```

### agent-browser (สำหรับ Claude)

ถ้า user มี Brave เปิด debugging port 9222:
```bash
agent-browser --cdp 9222 eval 'expression'
agent-browser --cdp 9222 snapshot -i
```
ใช้ debug live state, ทดสอบ API call โดยไม่ต้องให้ user ทำเอง

## Conventions

- **ภาษา:** UI strings + comments เป็นภาษาไทย, code identifiers ภาษาอังกฤษ
- **Naming:** state field ใช้ camelCase, API field ใช้ snake_case (matched กับที่ TikTok ส่ง)
- **No build step:** plain JS — แก้ไฟล์ → reload extension → ใช้งานได้ทันที
- **No package.json:** ไม่มี npm dependencies, vendor libraries pin version แล้ว
- **No tests:** test แบบ manual ผ่าน DevTools + agent-browser

## Reload workflow

หลังแก้ไฟล์ทุกครั้ง:
1. ไป `chrome://extensions/` (หรือ `brave://extensions/`)
2. กดปุ่ม **Reload** ที่ extension
3. กลับไปที่หน้า TikTok → **hard refresh** (`Cmd+Shift+R`)

⚠️ การแก้แค่ content.js แล้ว refresh page เฉย ๆ จะยังใช้ code เก่า เพราะ Chrome cache extension files

## ห้ามทำ

- **อย่ายิง print API โดยไม่มี confirmation** — TikTok mark labels เป็น "พิมพ์แล้ว" ทันที + เปลือง quota
- **อย่ายิง API จาก ID ที่ scan ผิด scenario** — เช่น คลิก "1 ชิ้น" ต้องไม่ติด weird order มาด้วย
- **อย่า bypass `applyCarrierFilter`** — ถ้า user toggle carrier ไว้ ต้องเคารพ filter
- **อย่า hardcode `fp` token หรือ session params** — capture สดทุก session

## localStorage keys เพิ่มเติม (Phase 2 Lane D)

| Key | Purpose | ขนาดสูงสุด | Cleanup |
|-----|---------|-----------|---------|
| `qf_last_picking_list_v1` | ค่า checkbox "แนบใบ Picking List" ล่าสุด ('true'/'false') | 5 bytes | ไม่มี |
| `qf_plan_snapshots_v1` | Snapshot แผนงานรายวัน (30 วัน) สำหรับ plannedQty ใน CSV | ~50 KB | trim >30 วัน เมื่อบันทึก |

`state.sellerEmail` — อ่านจาก cookie `passport_user_email=` ตอน init สำหรับ Picking List header (masked)

## PDF Template Builder (in progress)

ฟีเจอร์ WYSIWYG editor ให้เจ้าของร้านปรับแต่งฉลาก (ใส่โลโก้, ข้อความขอบคุณ, QR LINE ฯลฯ)
โดยไม่แตะ element ที่ **LOCKED** ของ carrier (barcode, QR, sort code, order ID)

### Zone model 3 ชั้น
- 🔒 **LOCKED** — `barcode main/left/right`, `qrCode`, `sortCode/routeCode/subZoneCode`,
  `orderId`, `trackingNumber`, `serviceType`, `codLabel` — render @ original coords เสมอ
- 🔄 **SHRINKABLE** — TikTok logo, carrier logo, `skuTable`, `addressBlock` — ย่อ/เลื่อนได้มี constraint
- 🎨 **CUSTOMIZABLE** — shop logo, thank-you text, LINE QR, `aliasWatermark`, `workerNameBadge` — ของ user

### Phase 1 status — **DONE** (calibration)

| Path | Purpose |
|---|---|
| [tools/inspect-pdf.js](tools/inspect-pdf.js) | Node script อ่าน PDF → dump layout JSON + auto-classify LOCKED/SHRINKABLE zones |
| [.claude/samples/jnt_layout.json](.claude/samples/jnt_layout.json) | Baseline J&T layout จาก `.claude/example.pdf` |
| [docs/pdf-builder-phase1.md](docs/pdf-builder-phase1.md) | รายงาน Phase 1 ภาษาไทย + proposed `J_AND_T_LAYOUT` constant |

Calibration ตรวจพบ: carrier=jnt, page=298×420pt (A6),
9 LOCKED regions, 2 SHRINKABLE, 2 CUSTOMIZABLE (overlays ของ extension เอง)

**Caveat:** `example.pdf` เป็นไฟล์ที่ผ่าน `overlayAliasOnPdf()` แล้ว — tool ทำ auto-filter
overlay ของเรา (`extensionOverlay:true`) ก่อนจัดหมวด LOCKED.
สำหรับ baseline 100% ต้อง re-run กับ **raw J&T PDF** ที่ไม่เคยผ่าน extension (ดู Phase 1 report §9)

### Phase 2 preview (ยังไม่เริ่ม)
- Canvas-based editor UI — consume `<carrier>_layout.json` → render draggable/resizable overlays
- Save user config ลง `localStorage.qf_label_template_v1` (planned key)
- Render pipeline ใหม่ใน `printIds()` → อ่าน template + overlay ด้วย pdf-lib แทน `overlayAliasOnPdf()`

### Phase 3 preview (ยังไม่เริ่ม)
- Multi-carrier: Flash, Kerry, SPX, Thailand Post — calibrate แต่ละราย carrier → `<carrier>_layout.json`
- Preset library — thank-you templates, QR positions, brand-logo sizing

### Run calibration ใหม่

```bash
# (ครั้งแรก) ติดตั้ง pdf2json ใน project root
npm install pdf2json

# รันบน PDF ใหม่
node tools/inspect-pdf.js path/to/label.pdf .claude/samples/<carrier>_layout.json
```

Tool จะ:
1. Extract text + coalesce glyphs (pdf2json แยก Thai เป็นรายตัว — ต้อง group ก่อน)
2. Auto-detect carrier ด้วย fingerprint (J&T, Flash, Kerry, SPX, ไปรษณีย์)
3. Classify LOCKED ด้วย heuristics: digit-length, font-size, position, service keywords
4. Tag extension overlays (top-right Thai text + bottom ≤36pt band) เป็น `extensionOverlay:true`
5. Redact PII (phone, long digit runs) ก่อนเขียน JSON

### localStorage keys (Phase 2 — planned)

| Key | Purpose |
|---|---|
| `qf_label_template_v1` | user-saved template config per carrier (position overrides, toggles, custom elements) |
| `qf_label_template_lastcarrier_v1` | carrier ล่าสุดที่ user เปิด editor
