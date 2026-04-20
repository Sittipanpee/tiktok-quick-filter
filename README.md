# ⚡ Quick Filter for TikTok Seller

Chrome Extension แบบ floating widget สำหรับ**หน้า "ที่จะจัดส่ง"** ของ TikTok Seller Center (TH)
ช่วยให้พนักงานกรองออเดอร์ตามสินค้า + ประเภทออเดอร์ได้ด้วยคลิกเดียว และเลือก (select) ออเดอร์ทั้งหมดที่ตรงเงื่อนไขอัตโนมัติ

## ✨ คุณสมบัติ

- **Scan อัตโนมัติ** — อ่านออเดอร์ทุกหน้า (pagination) เก็บเป็นลิสต์สินค้าแบบ unique (รูป + ชื่อ + จำนวน)
- **3 ประเภทฟิลเตอร์** ตรงกับตัวเลือก "เนื้อหาคำสั่งซื้อ" ของ TikTok
  - **1 ชิ้น** = ออเดอร์ที่สั่งสินค้านั้นชิ้นเดียว (รายการเดียว)
  - **หลายชิ้น** = ออเดอร์ที่สั่งสินค้านั้นหลายชิ้น แต่เป็น SKU เดียวกัน (SKU เดียว)
  - **ออเดอร์แปลก** = ออเดอร์ที่มีหลายสินค้าใน 1 คำสั่งซื้อ (SKU หลายรายการ)
- **คลิกเดียวกรอง** — คลิกภาพสินค้าในแต่ละแท็บ widget จะจัดการ:
  1. ใส่ product ID ลงช่องค้นหากลาง
  2. สลับกลับแท็บ "ที่จะจัดส่ง"
  3. เลือกฟิลเตอร์ "เนื้อหาคำสั่งซื้อ" ตามประเภท
  4. **Auto select** ออเดอร์ทั้งหมดที่ตรงเงื่อนไข (รวมข้ามหน้า) พร้อมให้กดปุ่ม "นัดหมายการจัดส่งเป็นชุด" / "พิมพ์เอกสาร" ต่อทันที

## 📦 ติดตั้ง

1. ดาวน์โหลด/โคลน repo นี้ลงเครื่อง
   ```bash
   git clone https://github.com/Sittipanpee/tiktok-quick-filter.git
   ```
2. เปิด Chrome → ไปที่ `chrome://extensions`
3. เปิด **Developer mode** (มุมขวาบน)
4. กด **Load unpacked** → เลือกโฟลเดอร์ `tiktok-quick-filter`
5. เปิดหน้า `https://seller-th.tiktok.com/order?tab=to_ship` จะเห็น widget **"⚡ Quick Filter"** ที่มุมขวาล่าง

## 🚀 วิธีใช้

1. เปิดแท็บ **"ที่จะจัดส่ง"** ของหน้า Order Management
2. กดปุ่ม **"🔍 Scan ทั้งหมด"** บน widget → รอ ~10–20 วิ (วนทุกหน้าอัตโนมัติ)
3. สลับแท็บ **1 ชิ้น / หลายชิ้น / ออเดอร์แปลก** บน widget
4. **คลิกรูปสินค้า** → widget จะกรอง + auto-select ออเดอร์ทั้งหมด
5. กด **"นัดหมายการจัดส่ง"** / **"พิมพ์เอกสาร"** บน TikTok ได้ต่อทันที

### ตัวเลือกเพิ่มเติม

- Toggle **"เลือกออเดอร์ทั้งหมดอัตโนมัติหลังกรอง"** — ปิดได้หากอยากกรองอย่างเดียวโดยไม่ auto-select
- ปุ่ม **↺** บน header — รีเซ็ต search + filter กลับสถานะเริ่มต้น
- ปุ่ม **−** — ย่อ widget

## 🛠️ โครงสร้างไฟล์

```
tiktok-quick-filter/
├── manifest.json       # Chrome Extension Manifest V3
├── content.js          # Logic หลัก (scan + filter + auto-select)
├── content.css         # Style widget
└── README.md
```

## 📝 หมายเหตุทางเทคนิค

- **ไม่เรียก API ภายนอก** — อ่านข้อมูลออเดอร์ผ่าน React Fiber ของตาราง (field `record.skuList`)
- ฟิลเตอร์ทำผ่านการจำลองคลิกแบบ **full event sequence** (`pointerdown → mousedown → pointerup → mouseup → click`) เพื่อให้ Pulse component ของ TikTok รับรู้
- รองรับ search ด้วย **productId** แทนการพิมพ์ชื่อ → match แม่นยำ 100%
- Auto-select รองรับทั้ง ≤50 รายการ (ใน 1 หน้า) และ >50 รายการ (ข้ามหน้า)

## ⚠️ ข้อจำกัด

- ใช้ได้เฉพาะ **TikTok Seller Center TH** (`seller-th.tiktok.com`)
- UI ของ TikTok เปลี่ยนบ่อย หาก selector เปลี่ยน อาจต้องแก้ `content.js`
- ก่อน scan ควรกดปุ่ม ↺ (reset) เพื่อล้างฟิลเตอร์เดิม จะได้ข้อมูล unique ครบทุกสินค้า

## 📄 License

Private use only.
# tiktok-quick-filter
