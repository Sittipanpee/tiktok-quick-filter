(() => {
  'use strict';

  // ==================== CONFIG ====================
  const WAIT_AFTER_PAGE_CLICK = 2500;
  const WAIT_AFTER_FILTER_CLICK = 2500;
  const WAIT_AFTER_SEARCH = 2500;
  const DONE_TIMEOUT_MS = 30 * 60 * 1000;

  // ==================== STATE ====================
  const state = {
    scanning: false,
    products: new Map(),
    weirdOrders: [],
    currentTab: 'single',
    autoSelectAll: true,
    doneItems: new Map(), // key: "productId" or "productId:skuId" → timestamp
  };

  // ==================== UTIL ====================
  const sleep = (ms) => new Promise(r => setTimeout(r, ms));

  function doneKey(productId, skuId) {
    return skuId ? `${productId}:${skuId}` : productId;
  }

  function markDone(productId, skuId) {
    state.doneItems.set(doneKey(productId, skuId), Date.now());
  }

  function isDone(productId, skuId) {
    const key = doneKey(productId, skuId);
    const ts = state.doneItems.get(key);
    if (!ts) return false;
    if (Date.now() - ts > DONE_TIMEOUT_MS) { state.doneItems.delete(key); return false; }
    return true;
  }

  function simulateClick(el) {
    if (!el) return;
    const rect = el.getBoundingClientRect();
    const x = rect.left + rect.width / 2;
    const y = rect.top + rect.height / 2;
    ['pointerdown', 'mousedown', 'pointerup', 'mouseup', 'click'].forEach(type => {
      el.dispatchEvent(new MouseEvent(type, {
        bubbles: true, cancelable: true, view: window,
        clientX: x, clientY: y, button: 0,
      }));
    });
  }

  function getOrderRecords() {
    const trs = [...document.querySelectorAll('tr')].filter(tr =>
      /\d{15,}/.test((tr.textContent || '').trim())
    );
    const records = [];
    for (const tr of trs) {
      const fk = Object.keys(tr).find(k => k.startsWith('__reactFiber'));
      if (!fk) continue;
      // ไต่ fiber ขึ้นไปหา record (เพราะ TikTok อาจ wrap หลายชั้น)
      let fiber = tr[fk];
      let rec = null;
      let depth = 0;
      while (fiber && depth < 10) {
        if (fiber.memoizedProps?.record) { rec = fiber.memoizedProps.record; break; }
        if (fiber.memoizedProps?.rowData) { rec = fiber.memoizedProps.rowData; break; }
        fiber = fiber.return;
        depth++;
      }
      if (rec && rec.skuList) records.push(rec);
    }
    return records;
  }

  // ---------- Pagination (รองรับหลาย selector + หน้าเดียว) ----------
  function getPageItems() {
    // ลอง selector ต่างๆ (p-pagination, arco-pagination, หรือ [aria-label^="Page"])
    const candidates = [
      '.p-pagination-item',
      '.arco-pagination-item',
      '[class*="pagination"] [aria-label^="Page"]',
      '[class*="Pagination"] [aria-label^="Page"]',
    ];
    for (const sel of candidates) {
      const items = [...document.querySelectorAll(sel)].filter(li =>
        /^\d+$/.test((li.textContent || '').trim())
      );
      if (items.length) return items;
    }
    return [];
  }

  function getCurrentPage() {
    const activeCandidates = [
      '.p-pagination-item-active',
      '.arco-pagination-item-active',
      '[class*="pagination-item-active"]',
      '[class*="Pagination"] [aria-current="page"]',
    ];
    for (const sel of activeCandidates) {
      const el = document.querySelector(sel);
      if (el) {
        const label = el.getAttribute('aria-label') || el.textContent.trim();
        const m = label.match(/\d+/);
        if (m) return m[0];
      }
    }
    // ถ้ามี pagination items แต่หา active ไม่เจอ → สมมติหน้า 1
    if (getPageItems().length) return '1';
    // ไม่มี pagination เลย = หน้าเดียว
    return null;
  }

  function getTotalPages() {
    const items = getPageItems();
    if (!items.length) return 1; // หน้าเดียว
    return Math.max(...items.map(li => parseInt(li.textContent.trim())));
  }

  async function waitForRecordsChange(prevFirstId, maxWait = 5000) {
    const start = Date.now();
    while (Date.now() - start < maxWait) {
      await sleep(200);
      const recs = getOrderRecords();
      if (recs.length > 0 && recs[0].mainOrderId !== prevFirstId) return true;
    }
    return false;
  }

  async function goToPage(pageNum) {
    const items = getPageItems();
    if (!items.length) return false;
    const target = items.find(li => {
      const aria = li.getAttribute('aria-label');
      if (aria && aria === `Page ${pageNum}`) return true;
      return (li.textContent || '').trim() === String(pageNum);
    });
    if (!target) return false;
    const before = getOrderRecords()[0]?.mainOrderId;
    simulateClick(target);
    await waitForRecordsChange(before, WAIT_AFTER_PAGE_CLICK + 2000);
    await sleep(300);
    return true;
  }

  function showToast(msg, duration = 2500) {
    document.querySelectorAll('.qf-toast').forEach(e => e.remove());
    const t = document.createElement('div');
    t.className = 'qf-toast';
    t.textContent = msg;
    document.body.appendChild(t);
    setTimeout(() => t.remove(), duration);
  }

  // ==================== SCANNING ====================
  async function scanAllPages() {
    if (state.scanning) return;
    state.scanning = true;
    state.products.clear();
    state.weirdOrders = [];
    state.doneItems.clear();

    const statusEl = document.getElementById('qf-scan-status');
    const btn = document.getElementById('qf-scan-btn');
    btn.disabled = true;

    try {
      const totalPages = getTotalPages();
      const currentPage = getCurrentPage();

      // กลับหน้า 1 เฉพาะเมื่อรู้แน่ว่าอยู่หน้าอื่น
      if (totalPages > 1 && currentPage && currentPage !== '1') {
        statusEl.textContent = 'กำลังกลับหน้า 1...';
        await goToPage(1);
      }

      // scan หน้าปัจจุบันเสมอ (แม้ไม่มี pagination)
      if (totalPages === 1) {
        statusEl.textContent = `กำลังสแกน...`;
        processRecordsOnPage();
      } else {
        for (let p = 1; p <= totalPages; p++) {
          statusEl.textContent = `กำลังสแกนหน้า ${p}/${totalPages}...`;
          const cur = getCurrentPage();
          if (cur !== null && cur !== String(p)) {
            const ok = await goToPage(p);
            if (!ok) {
              console.warn('[QF] goToPage failed at', p);
              break;
            }
          }
          processRecordsOnPage();
        }
      }

      statusEl.textContent = `✓ พบ ${state.products.size} สินค้า, ${state.weirdOrders.length} ออเดอร์แปลก`;
      renderAll();
    } catch (err) {
      console.error('[QF] scan error', err);
      statusEl.textContent = 'ผิดพลาด: ' + err.message;
    } finally {
      state.scanning = false;
      btn.disabled = false;
    }
  }

  function processRecordsOnPage() {
    const records = getOrderRecords();
    for (const rec of records) {
      const skus = rec.skuList || [];
      for (const s of skus) {
        if (!state.products.has(s.productId)) {
          state.products.set(s.productId, {
            productId: s.productId,
            productName: s.productName || '(ไม่มีชื่อ)',
            productImageURL: s.productImageURL || '',
            variants: new Map(),
            orderCountSingle: 0,
            orderCountMulti: 0,
          });
        }
        const product = state.products.get(s.productId);
        if (!product.variants.has(s.skuId)) {
          product.variants.set(s.skuId, {
            skuId: s.skuId,
            skuName: s.skuName || '',
            sellerSkuName: s.sellerSkuName || '',
            orderCountSingle: 0,
            orderCountMulti: 0,
          });
        }
      }
      if (skus.length === 1) {
        const product = state.products.get(skus[0].productId);
        product.orderCountSingle++;
        product.variants.get(skus[0].skuId).orderCountSingle++;
      } else if (skus.length > 1) {
        state.weirdOrders.push({
          orderId: rec.mainOrderId,
          skus: skus.map(s => ({
            productId: s.productId,
            productName: s.productName || '',
            productImageURL: s.productImageURL || '',
            quantity: s.quantity,
          })),
        });
        for (const s of skus) {
          const product = state.products.get(s.productId);
          product.orderCountMulti++;
          product.variants.get(s.skuId).orderCountMulti++;
        }
      }
    }
  }

  // ==================== FILTER ACTIONS ====================
  async function applyFilterCountType(type) {
    const labels = ['เนื้อหาคำสั่งซื้อ', 'รายการเดียว', 'SKU เดียว', 'SKU หลายรายการ'];
    let trigger = null;
    for (const label of labels) {
      trigger = [...document.querySelectorAll('div')].find(el => {
        const t = (el.textContent || '').trim();
        return t === label && typeof el.onclick === 'function' && el.closest('.p-space-item');
      });
      if (trigger) break;
    }
    if (!trigger) throw new Error('ไม่เจอตัวกรอง "เนื้อหาคำสั่งซื้อ"');
    simulateClick(trigger);
    await sleep(500);
    const optionText = {
      single_item: 'รายการเดียว',
      single_sku: 'SKU เดียว',
      multi_sku: 'SKU หลายรายการ',
    }[type];
    const option = [...document.querySelectorAll('.p-dropdown-menu-item')]
      .find(el => el.textContent.trim() === optionText);
    if (!option) throw new Error('ไม่เจอตัวเลือก ' + optionText);
    simulateClick(option);
    await sleep(WAIT_AFTER_FILTER_CLICK);
  }

  async function clearFilterCountType() {
    const items = [...document.querySelectorAll('.p-space-item')];
    for (const item of items) {
      const txt = (item.textContent || '').trim();
      if (/รายการเดียว|SKU เดียว|SKU หลายรายการ/.test(txt) && !/เนื้อหาคำสั่งซื้อ/.test(txt)) {
        const closeBtn = item.querySelector('svg, [class*="close"], [class*="Close"]');
        if (closeBtn) {
          simulateClick(closeBtn.closest('svg') || closeBtn);
          await sleep(800);
          return true;
        }
      }
    }
    return false;
  }

  async function setSearchBox(value) {
    const input = document.querySelector('input[placeholder*="หมายเลขคำสั่งซื้อ"]');
    if (!input) throw new Error('ไม่เจอช่องค้นหา');
    const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
    setter.call(input, value);
    input.dispatchEvent(new Event('input', { bubbles: true }));
    input.focus();
    await sleep(300);
    input.dispatchEvent(new KeyboardEvent('keydown', {
      bubbles: true, key: 'Enter', code: 'Enter', keyCode: 13, which: 13,
    }));
    await sleep(WAIT_AFTER_SEARCH);
  }

  async function clickToShipTab() {
    const tab = [...document.querySelectorAll('.p-tabs-header-title')]
      .find(el => /^ที่จะจัดส่ง/.test(el.textContent.trim()));
    if (!tab) throw new Error('ไม่เจอแท็บ "ที่จะจัดส่ง"');
    simulateClick(tab);
    await sleep(WAIT_AFTER_PAGE_CLICK);
  }

  // ==================== AUTO SELECT ALL ====================
  function findTrForOrder(orderId) {
    return [...document.querySelectorAll('tr')].find(tr => {
      const fk = Object.keys(tr).find(k => k.startsWith('__reactFiber'));
      if (!fk) return false;
      let fiber = tr[fk], depth = 0;
      while (fiber && depth < 10) {
        if (fiber.memoizedProps?.record?.mainOrderId === orderId) return true;
        if (fiber.memoizedProps?.rowData?.mainOrderId === orderId) return true;
        fiber = fiber.return; depth++;
      }
      return false;
    });
  }

  async function waitForStable(maxWait = 8000) {
    const start = Date.now();
    let prevCount = null, stableFor = 0;
    while (Date.now() - start < maxWait) {
      const countEl = [...document.querySelectorAll('*')]
        .find(el => /^พบค่าคำสั่งซื้อ/.test((el.textContent || '').trim()));
      const m = countEl?.textContent?.match(/พบค่าคำสั่งซื้อ (\d+) รายการ/);
      const now = m ? parseInt(m[1]) : null;
      if (now !== null && now === prevCount) {
        stableFor += 200;
        if (stableFor >= 600) break;
      } else { stableFor = 0; prevCount = now; }
      await sleep(200);
    }
  }

  async function selectAllOrders(filterSkuId = null, maxWait = 8000) {
    await waitForStable(maxWait);

    if (filterSkuId) {
      const totalPages = getTotalPages();
      let count = 0;
      for (let p = 1; p <= totalPages; p++) {
        if (p > 1) {
          const ok = await goToPage(p);
          if (!ok) break;
          await waitForStable(4000);
        }
        for (const rec of getOrderRecords()) {
          if (!rec.skuList?.some(s => s.skuId === filterSkuId)) continue;
          const tr = findTrForOrder(rec.mainOrderId);
          const checkbox = tr?.querySelector('label.p-checkbox');
          if (checkbox) { simulateClick(checkbox); count++; await sleep(50); }
        }
      }
      return count;
    }

    const headerLabel = document.querySelector(
      'th[data-log_click_for="select_all_items_in_page"] label.p-checkbox'
    );
    if (!headerLabel) throw new Error('ไม่เจอ header checkbox');
    simulateClick(headerLabel);
    await sleep(800);
    let selectAllLink = null;
    const t2 = Date.now();
    while (Date.now() - t2 < 2000) {
      selectAllLink = [...document.querySelectorAll('span')].find(el => {
        const t = (el.textContent || '').trim();
        if (!/^เลือกคำสั่งซื้อทั้งหมด \d+ รายการ$/.test(t)) return false;
        return getComputedStyle(el).cursor === 'pointer';
      });
      if (selectAllLink) break;
      await sleep(200);
    }
    if (selectAllLink) {
      simulateClick(selectAllLink);
      await sleep(1000);
      const banner = [...document.querySelectorAll('th[colspan]')]
        .find(el => (el.textContent || '').includes('เลือกคำสั่งซื้อทั้งหมด'));
      const m = banner?.textContent?.match(/เลือกคำสั่งซื้อทั้งหมด (\d+) รายการ/);
      return m ? parseInt(m[1]) : null;
    }
    return document.querySelectorAll('tr td.col-checkbox label.p-checkbox-checked').length;
  }

  // ==================== HIGH-LEVEL ACTIONS ====================
  async function applyProductFilter(productId, skuId, type) {
    try {
      const label = skuId ? 'กำลังกรอง variant...' : 'กำลังกรอง...';
      showToast(label, 10000);
      await setSearchBox(productId);
      await clickToShipTab();
      await applyFilterCountType(type);
      if (state.autoSelectAll) {
        showToast('กำลังเลือกทั้งหมด...', 5000);
        await sleep(500);
        const total = await selectAllOrders(skuId);
        showToast(`✓ กรอง + เลือก ${total ?? ''} ออเดอร์`, 2500);
      } else {
        showToast('✓ กรองเรียบร้อย', 2000);
      }
      markDone(productId, skuId);
      renderAll();
    } catch (e) {
      console.error('[QF]', e);
      showToast('ผิดพลาด: ' + e.message, 3000);
    }
  }

  async function applyWeirdFilter() {
    try {
      showToast('กำลังกรองออเดอร์แปลก...', 10000);
      await setSearchBox('');
      await clickToShipTab();
      await applyFilterCountType('multi_sku');
      if (state.autoSelectAll) {
        showToast('กำลังเลือกทั้งหมด...', 5000);
        await sleep(500);
        const total = await selectAllOrders();
        showToast(`✓ เลือกออเดอร์แปลก ${total ?? ''} รายการ`, 2500);
      } else {
        showToast('✓ แสดงออเดอร์แปลกทั้งหมด', 2000);
      }
    } catch (e) {
      console.error('[QF]', e);
      showToast('ผิดพลาด: ' + e.message, 3000);
    }
  }

  async function resetFilters() {
    try {
      showToast('กำลังรีเซ็ต...', 5000);
      await setSearchBox('');
      await clickToShipTab();
      await clearFilterCountType();
      showToast('✓ รีเซ็ตแล้ว', 1500);
    } catch (e) {
      console.error('[QF]', e);
    }
  }

  // ==================== UI ====================
  function buildWidget() {
    if (document.getElementById('qf-widget')) return;
    const w = document.createElement('div');
    w.id = 'qf-widget';
    w.innerHTML = `
      <div id="qf-header">
        <span>⚡ Quick Filter</span>
        <div id="qf-header-actions">
          <button id="qf-reset-btn" title="รีเซ็ตฟิลเตอร์">↺</button>
          <button id="qf-toggle-btn" title="ย่อ/ขยาย">−</button>
        </div>
      </div>
      <div id="qf-body">
        <div id="qf-scan-row">
          <button id="qf-scan-btn">🔍 Scan ทั้งหมด</button>
          <span id="qf-scan-status">กดปุ่มเพื่อเริ่มสแกน</span>
        </div>
        <div id="qf-options-row">
          <label>
            <input type="checkbox" id="qf-auto-select" checked />
            เลือกออเดอร์ทั้งหมดอัตโนมัติหลังกรอง
          </label>
        </div>
        <div id="qf-tabs">
          <div class="qf-tab active" data-tab="single">1 ชิ้น <span class="qf-tab-count" id="qf-count-single">0</span></div>
          <div class="qf-tab" data-tab="multi">หลายชิ้น <span class="qf-tab-count" id="qf-count-multi">0</span></div>
          <div class="qf-tab" data-tab="weird">ออเดอร์แปลก <span class="qf-tab-count" id="qf-count-weird">0</span></div>
        </div>
        <div id="qf-content">
          <div class="qf-empty">ยังไม่ได้สแกน</div>
        </div>
      </div>
    `;
    document.body.appendChild(w);
    document.getElementById('qf-header').addEventListener('click', (e) => {
      if (e.target.closest('#qf-header-actions')) return;
      w.classList.toggle('qf-collapsed');
      document.getElementById('qf-toggle-btn').textContent =
        w.classList.contains('qf-collapsed') ? '+' : '−';
    });
    document.getElementById('qf-toggle-btn').addEventListener('click', (e) => {
      e.stopPropagation();
      w.classList.toggle('qf-collapsed');
      e.target.textContent = w.classList.contains('qf-collapsed') ? '+' : '−';
    });
    document.getElementById('qf-reset-btn').addEventListener('click', (e) => {
      e.stopPropagation();
      resetFilters();
    });
    document.getElementById('qf-scan-btn').addEventListener('click', scanAllPages);
    document.getElementById('qf-auto-select').addEventListener('change', (e) => {
      state.autoSelectAll = e.target.checked;
    });
    document.querySelectorAll('.qf-tab').forEach(tab => {
      tab.addEventListener('click', () => {
        document.querySelectorAll('.qf-tab').forEach(t => t.classList.remove('active'));
        tab.classList.add('active');
        state.currentTab = tab.dataset.tab;
        renderContent();
      });
    });
    makeDraggable(w, document.getElementById('qf-header'));
  }

  function makeDraggable(el, handle) {
    let sx = 0, sy = 0, sl = 0, st = 0, dragging = false;
    handle.addEventListener('mousedown', (e) => {
      if (e.target.closest('button')) return;
      dragging = true;
      sx = e.clientX; sy = e.clientY;
      const r = el.getBoundingClientRect();
      sl = r.left; st = r.top;
      el.style.right = 'auto'; el.style.bottom = 'auto';
      el.style.left = r.left + 'px'; el.style.top = r.top + 'px';
      e.preventDefault();
    });
    window.addEventListener('mousemove', (e) => {
      if (!dragging) return;
      el.style.left = (sl + e.clientX - sx) + 'px';
      el.style.top = (st + e.clientY - sy) + 'px';
    });
    window.addEventListener('mouseup', () => dragging = false);
  }

  function renderAll() {
    const singleCount = [...state.products.values()].filter(p => p.orderCountSingle > 0).length;
    const multiCount = [...state.products.values()].filter(p => p.orderCountMulti > 0).length;
    document.getElementById('qf-count-single').textContent = singleCount;
    document.getElementById('qf-count-multi').textContent = multiCount;
    document.getElementById('qf-count-weird').textContent = state.weirdOrders.length;
    renderContent();
  }

  function renderContent() {
    const wrap = document.getElementById('qf-content');
    wrap.innerHTML = '';
    if (state.products.size === 0 && state.weirdOrders.length === 0) {
      wrap.innerHTML = '<div class="qf-empty">ยังไม่ได้สแกน</div>';
      return;
    }
    if (state.currentTab === 'single' || state.currentTab === 'multi') {
      const key = state.currentTab === 'single' ? 'orderCountSingle' : 'orderCountMulti';
      const type = state.currentTab === 'single' ? 'single_item' : 'single_sku';
      const products = [...state.products.values()]
        .filter(p => p[key] > 0)
        .sort((a, b) => b[key] - a[key]);
      if (!products.length) {
        wrap.innerHTML = '<div class="qf-empty">ไม่พบสินค้าในประเภทนี้</div>';
        return;
      }
      const grid = document.createElement('div');
      grid.className = 'qf-product-grid';
      for (const p of products) {
        const card = document.createElement('div');
        const cardDone = isDone(p.productId, null);
        card.className = 'qf-product-card' + (cardDone ? ' qf-done' : '');
        card.title = p.productName;
        const variants = [...p.variants.values()].filter(v => v[key] > 0);
        const hasBadges = variants.length > 1;
        card.innerHTML = `
          <img src="${p.productImageURL}" alt="" referrerpolicy="no-referrer"/>
          <div class="qf-product-name">${escapeHtml(p.productName)}</div>
          <div class="qf-product-count">${p[key]} ออเดอร์</div>
          ${hasBadges ? `<div class="qf-variant-badges"></div>` : ''}
        `;
        if (hasBadges) {
          const badgesEl = card.querySelector('.qf-variant-badges');
          for (const v of variants) {
            const badgeDone = isDone(p.productId, v.skuId);
            const badge = document.createElement('span');
            badge.className = 'qf-variant-badge' + (badgeDone ? ' qf-badge-done' : '');
            badge.dataset.skuId = v.skuId;
            badge.textContent = `${v.skuName || v.sellerSkuName || v.skuId} (${v[key]})`;
            badge.addEventListener('click', (e) => {
              e.stopPropagation();
              if (isDone(p.productId, v.skuId)) {
                state.doneItems.delete(doneKey(p.productId, v.skuId));
                renderAll();
              } else {
                applyProductFilter(p.productId, v.skuId, type);
              }
            });
            badgesEl.appendChild(badge);
          }
        }
        card.addEventListener('click', () => {
          if (cardDone) {
            state.doneItems.delete(doneKey(p.productId, null));
            renderAll();
          } else {
            applyProductFilter(p.productId, null, type);
          }
        });
        grid.appendChild(card);
      }
      wrap.appendChild(grid);
    } else if (state.currentTab === 'weird') {
      const applyBtn = document.createElement('button');
      applyBtn.id = 'qf-weird-apply-btn';
      applyBtn.textContent = `📋 แสดงออเดอร์แปลกทั้งหมด (${state.weirdOrders.length})`;
      applyBtn.addEventListener('click', applyWeirdFilter);
      wrap.appendChild(applyBtn);
      if (!state.weirdOrders.length) {
        const empty = document.createElement('div');
        empty.className = 'qf-empty';
        empty.textContent = 'ไม่พบออเดอร์แปลก';
        wrap.appendChild(empty);
        return;
      }
      const list = document.createElement('div');
      list.id = 'qf-weird-list';
      for (const o of state.weirdOrders) {
        const card = document.createElement('div');
        card.className = 'qf-weird-order';
        const items = o.skus.map(s => `
          <div class="qf-weird-item">
            <img src="${s.productImageURL}" referrerpolicy="no-referrer"/>
            <span>${escapeHtml((s.productName || '').substring(0, 40))} ×${s.quantity}</span>
          </div>
        `).join('');
        card.innerHTML = `
          <div class="qf-weird-order-id">#${o.orderId}</div>
          <div class="qf-weird-items">${items}</div>
        `;
        list.appendChild(card);
      }
      wrap.appendChild(list);
    }
  }

  function escapeHtml(s) {
    return (s || '').replace(/[&<>"']/g, c => ({
      '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
    }[c]));
  }

  // ==================== INIT ====================
  function init() {
    if (!/\/order(\?|$|\/)/.test(location.pathname + location.search)) return;
    buildWidget();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
