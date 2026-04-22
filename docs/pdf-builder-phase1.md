# PDF Template Builder — Phase 1 Report (Calibration)

สรุปผลการ calibrate coordinate system ของฉลาก J&T จาก `.claude/example.pdf`
เพื่อเตรียม Phase 2 WYSIWYG editor

> **ข้อควรระวังสำคัญ** — sample PDF ที่ใช้ถูกพิมพ์ผ่าน extension นี้แล้ว
> ดังนั้นมี overlay 2 ชุด (worker-name มุมขวาบน, alias watermark ล่างสุด)
> ซึ่ง **ไม่ใช่ LOCKED** ของ carrier. Tool ทำ auto-filter ให้แล้ว (`extensionOverlay:true`).
> สำหรับ baseline 100% ที่แม่นยำ ต้อง re-calibrate ด้วย raw PDF ที่ยังไม่เคยผ่าน extension.

## 1. ภาพรวมผลลัพธ์

| รายการ | ค่า |
|---|---|
| Carrier ที่ตรวจพบ | **`jnt`** (ผ่าน structural heuristics — 2 side digit columns + 12-digit tracking + "DROP-OFF") |
| Page size | **298 × 420 pt** (≈ 4.14 × 5.83 นิ้ว — ใกล้ A6) |
| Text runs (coalesced) | 117 |
| Rects (framing boxes จาก HLines/VLines) | 0 — pdf2json ไม่เจอ (ต้อง render เป็น raster หรือใช้ pdfjs-dist ภายหลัง) |
| LOCKED regions classified | **9** |
| SHRINKABLE regions classified | **2** |
| CUSTOMIZABLE (extension overlays) | **2** |
| Unclassified (ต้อง user review) | 32 |

Tool script: [`tools/inspect-pdf.js`](../tools/inspect-pdf.js)
Output JSON: [`.claude/samples/jnt_layout.json`](../.claude/samples/jnt_layout.json)

### วิธีใช้
```bash
node tools/inspect-pdf.js <pdf-path> <out-json-path>
# ตัวอย่าง:
node tools/inspect-pdf.js .claude/example.pdf .claude/samples/jnt_layout.json
```

## 2. LOCKED zone (carrier-required — ห้ามเลื่อน/ย่อ)

| Key | coord (x, y) | size | ที่มา |
|---|---|---|---|
| `orderId` | (42.86, 162.83) | 41pt | digit run ≥14 chars — order ID (ฟิลด์ตัวเลข TikTok ID) |
| `trackingNumber` | (105.86, 341.5) | 65pt | 12-digit เด่นกลางบน — waybill ของ J&T |
| `sortCode` | (188.18, 412.38) | 78pt | `"EZ"` — sort code มุมขวาบน |
| `routeCode` | (164.38, 312.38) | **124pt** | `"H1 F04-33"` — zone/route สำคัญของ J&T |
| `subZoneCode` | (204.54, 290.38) | 66pt | `"004A"` — sub-zone |
| `serviceType` | (215.62, 186.42) | 56pt | `"DROP-OFF"` — บริการที่เลือก |
| `codLabel` | (53.28, 183.15) | 107pt | `"C OD"` — ป้าย COD (ถ้ามี) |
| `barcodeLeft` | x=9.62, y=187.68, w=14, h=209.95 | — | คอลัมน์เลขแนวตั้งขอบซ้าย (J&T mirror ของ tracking number) |
| `barcodeRight` | x=277.68, y=78.94, w=14, h=321.51 | — | คอลัมน์เลขแนวตั้งขอบขวา |

**หมายเหตุ:** barcode bar image และ QR code ของ J&T เป็น raster ที่ pdf2json มองไม่เห็น
ต้อง Phase 2 เติม extractor อื่น (pdfjs-dist หรือ pdftoppm + image analysis) เพื่อจับ bbox ของ barcode bar ภาพจริง
ตอนนี้ใช้ "proxy LOCKED" ผ่าน digit columns และ tracking number ซึ่งเพียงพอสำหรับ editor ที่จะ mask พื้นที่นี้

## 3. SHRINKABLE zone (ย่อ/เลื่อนได้ — editor allow)

| Key | bbox | หมายเหตุ |
|---|---|---|
| `skuTable` | x=0, y=78.94, w=298, h=51.57 | ตาราง "Product Name / SKU / Qty" ด้านล่าง; ย่อได้ถ้า shop มี SKU น้อย |
| `addressBlock` | x=0, y=147, w=298, h=126 | block ที่อยู่ผู้ส่ง/ผู้รับ (approx — `needsVerification:true`) |

## 4. CUSTOMIZABLE (extension overlays — ของเราเอง)

| Key | จุด | หมายเหตุ |
|---|---|---|
| `aliasWatermark` | y≈16–31 bottom band | ลายน้ำ alias เช่น `{alias_redacted}` ที่ `overlayAliasOnPdf()` ใส่ |
| `workerNameBadge` | y≈409 top-right, x>180 | ชื่อคนแพ็กเช่น `{worker_redacted}` |

ทั้ง 2 รายการนี้คือ **editor target** — Phase 2 ให้ user toggle on/off หรือเปลี่ยน position/style ได้

## 5. Unclassified — รายการที่ต้องให้ user ช่วย label

32 text runs ที่ยังไม่ match heuristic ใด — ส่วนใหญ่คือ:
- `"698"` (y=424, x=275.8), `"V"` (y=420, x=4.6) — ตัวอักษรเดี่ยว/ตัวเลขเล็กๆ ที่อาจเป็น template version marker (LOCKED แต่ optional)
- `"จาก"` / `"ถึง"` — Thai labels "FROM" / "TO" ของ address block (SHRINKABLE ใน address block อยู่แล้ว)
- บรรทัด address ภาษาไทย/ตัวเลข postal (อยู่ใน addressBlock bbox แล้ว)
- `"Shipping Date:" / "Estimated Date:" / "In transit by:"` — labels ของวันที่ (LOCKED — วันที่สำคัญ)
- `"Product NameSKUSeller SKUQty"` — header ของตาราง SKU (LOCKED header)
- `"người mua không cần phải trả chuyển phát"` — ภาษาเวียดนามในฉลาก (?? unusual — อาจเป็น template leak)

## 6. Proposed `J_AND_T_LAYOUT` constant

พร้อม paste เข้า `content.js` เมื่อ Phase 2 render pipeline ต้องใช้:

```javascript
// Phase 1 calibration baseline — coordinates in PDF points (bottom-left origin)
// Source: .claude/samples/jnt_layout.json from .claude/example.pdf
// Status: DRAFT — re-verify with 2+ raw J&T samples before shipping to users
const J_AND_T_LAYOUT = {
  carrier: 'jnt',
  pageSize: { w: 298, h: 420 },
  locked: {
    orderId:         { x: 42.86,  y: 162.83, size: 41,  kind: 'text' },
    trackingNumber:  { x: 105.86, y: 341.5,  size: 65,  kind: 'text' },
    sortCode:        { x: 188.18, y: 412.38, size: 78,  kind: 'text' },
    routeCode:       { x: 164.38, y: 312.38, size: 124, kind: 'text' },
    subZoneCode:     { x: 204.54, y: 290.38, size: 66,  kind: 'text' },
    serviceType:     { x: 215.62, y: 186.42, size: 56,  kind: 'text' },
    codLabel:        { x: 53.28,  y: 183.15, size: 107, kind: 'text' },
    barcodeLeft:     { x: 9.62,   y: 187.68, w: 14, h: 209.95, kind: 'region' },
    barcodeRight:    { x: 277.68, y: 78.94,  w: 14, h: 321.51, kind: 'region' },
    // TODO(phase2): barcodeMain bar image bbox (requires pdfjs-dist or image analysis)
    // TODO(phase2): qrCode bbox
  },
  shrinkable: {
    skuTable:     { x: 0, y: 78.94, w: 298, h: 51.57 },
    addressBlock: { x: 0, y: 147,   w: 298, h: 126, needsVerification: true },
  },
  customizable: {
    // Extension-owned overlays — editor can move/toggle
    aliasWatermark:   { defaultY: 6,   size: 22, color: '#000', opacity: 0.4 },
    workerNameBadge:  { defaultY: 409, defaultX: 261, size: 20.8 },
  },
};
```

## 7. ความไม่แน่นอน (uncertainties)

1. **Barcode bar image bbox ยังไม่รู้** — pdf2json ไม่ export raster image bbox
   ต้อง Phase 2 เติม pdfjs-dist หรือใช้ pdftoppm -r 300 render เป็น image แล้ว detect bar region ด้วย OpenCV/Canvas
2. **QR code bbox ยังไม่รู้** — เหตุผลเดียวกัน
3. **TikTok platform logo position** — ตอนนี้ยังไม่ detect (sample ฉลาก J&T นี้อาจไม่มีโลโก้ TikTok — ต้อง confirm ด้วย sample ที่สอง)
4. **addressBlock bbox ประมาณการเอง** — ยังไม่มี rect framing; ต้องยืนยันด้วย sample ที่สอง
5. **สถานะ serviceType/codLabel ว่า LOCKED จริงหรือไม่** — COD badge บาง carrier ก็ดึงออกได้ถ้าไม่ใช่ COD order
   ควรเช็คตรรกะ render ที่ TikTok ฝั่งเซิร์ฟเวอร์
6. **Side digit columns (`barcodeLeft/Right`) LOCKED ระดับใด** — เป็นเพียง mirror ของ tracking number
   ถ้า editor ต้องการย่อให้ใส่เนื้อหาอื่นเพิ่ม อาจอนุญาตให้ซ่อนได้ (แต่ default ควร keep)
7. **Page size 298 × 420pt ไม่ตรง standard** — A6 = 298 × 420 pt (actually matches A6 exactly)
   ยืนยันได้ว่าไม่เกิดจาก crop ผิดพลาด

## 8. Next steps (Phase 2 preview)

Phase 2 editor จะ consume `jnt_layout.json` ดังนี้:

1. **Canvas renderer** อ่าน `pageSize` → scale to DOM
2. **LOCKED overlay** — วาด mask สี/เส้นประรอบ bbox แต่ละ locked region (user เลื่อน/ลบไม่ได้)
3. **SHRINKABLE handles** — drag/resize handles on skuTable, addressBlock; min/max constraint
4. **CUSTOMIZABLE toolbox** — add shop logo / thank-you text / LINE QR / custom alias position
5. **Export** — layout config เป็น JSON แล้วส่งเข้า render pipeline ที่ใช้ pdf-lib overlay
6. **Multi-carrier support** — Flash/Kerry/SPX/Thaipost: เก็บ `<carrier>_layout.json` แยกไฟล์ย่อย key ลง `samples/`

## 9. Action items สำหรับ user

- [ ] อัปโหลด raw J&T PDF (ที่ไม่เคยผ่าน extension) 2–3 ใบ เพื่อ re-verify LOCKED coords stable
- [ ] อัปโหลด J&T PDF ที่มี:
  - [ ] COD flag (เทียบ codLabel position)
  - [ ] Multi-SKU (เทียบ skuTable expansion)
  - [ ] ที่อยู่ยาว/สั้น (เทียบ addressBlock ยืดหยุ่น)
- [ ] Review `.claude/samples/jnt_layout.json` → approve/reject หรือ hand-label `hints.unclassified` entries
- [ ] ถ้า approve → Phase 2 เริ่ม editor UI และ wire `J_AND_T_LAYOUT` constant เข้า `content.js`

## 10. Privacy notes

ไฟล์ output JSON มีการ redact แล้ว:
- Phone: `(+66)XXX` → `{phone_redacted}`
- Long digit runs (order/tracking): `5836XXXXX26` → `5836••26`
- Thai address lines: **ไม่ redacted** เพราะ coord สำคัญกว่าค่า (ควรดูผ่าน coord เท่านั้น)

หากต้อง commit JSON เข้า git ควร double-check ไม่มีเลขโทรศัพท์เต็ม / ชื่อจริงเหลืออยู่
