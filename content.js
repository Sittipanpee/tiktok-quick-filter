(() => {
  'use strict';

  // ==================== FETCH HOOK (runs at document_start, before TikTok saves window.fetch) ====================
  let _apiListUrl = null;
  let _apiListBodyTemplate = null;
  let _labelsApiUrl = null;
  let _labelsApiBodyTemplate = null;
  // Shopee captured URLs/bodies
  let _shopeeIndexUrl = null;
  let _shopeeIndexBody = null;
  let _shopeeCardUrl = null;
  let _shopeeCardBody = null;
  const _origFetch = window.fetch;
  window.fetch = async function(...args) {
    // Extract URL + body, robust to both URL-string and Request-object call shapes
    let url = '';
    let rawBodyStr = null;
    let bodyPromise = null; // for Request objects (read async)

    if (args[0] instanceof Request) {
      url = args[0].url;
      try { bodyPromise = args[0].clone().text(); } catch (e) {}
    } else {
      url = typeof args[0] === 'string' ? args[0] : (args[0]?.url || '');
      const b = args[1]?.body;
      if (typeof b === 'string') rawBodyStr = b;
      else if (b instanceof Blob) bodyPromise = b.text();
    }

    const captureBody = (sink) => {
      if (rawBodyStr) {
        try { sink(JSON.parse(rawBodyStr)); } catch (e) {}
      } else if (bodyPromise) {
        bodyPromise.then(t => { try { sink(JSON.parse(t)); } catch (e) {} }).catch(() => {});
      }
    };

    // Order list — capture URL + body template
    if (!_apiListUrl && url.includes('/order/list')) {
      _apiListUrl = url;
      captureBody(b => { if (!_apiListBodyTemplate) _apiListBodyTemplate = b; });
    }

    // Labels API — direct URL match (fast, no response sniffing)
    if (!_labelsApiUrl && url.includes('/api/fulfillment/package/list')) {
      _labelsApiUrl = url;
      captureBody(b => { if (!_labelsApiBodyTemplate) _labelsApiBodyTemplate = b; });
    }

    // Shopee APIs (some use fetch but most XHR — both branches handled)
    if (!_shopeeIndexUrl && url.includes('/api/v3/order/search_order_list_index')) {
      _shopeeIndexUrl = url;
      captureBody(b => { if (!_shopeeIndexBody) _shopeeIndexBody = b; });
    }
    if (!_shopeeCardUrl && url.includes('/api/v3/order/get_order_list_card_list')) {
      _shopeeCardUrl = url;
      captureBody(b => { if (!_shopeeCardBody) _shopeeCardBody = b; });
    }

    return await _origFetch.apply(this, args);
  };

  // XMLHttpRequest hook (Shopee uses XHR for all order APIs)
  const _origXhrOpen = XMLHttpRequest.prototype.open;
  XMLHttpRequest.prototype.open = function(method, url, ...rest) {
    this._qfUrl = url;
    this._qfMethod = method;
    return _origXhrOpen.apply(this, [method, url, ...rest]);
  };
  const _origXhrSend = XMLHttpRequest.prototype.send;
  XMLHttpRequest.prototype.send = function(body) {
    const url = this._qfUrl || '';
    if (!_shopeeIndexUrl && url.includes('/api/v3/order/search_order_list_index')) {
      _shopeeIndexUrl = url;
      if (typeof body === 'string') { try { _shopeeIndexBody = JSON.parse(body); } catch {} }
    }
    if (!_shopeeCardUrl && url.includes('/api/v3/order/get_order_list_card_list')) {
      _shopeeCardUrl = url;
      if (typeof body === 'string') { try { _shopeeCardBody = JSON.parse(body); } catch {} }
    }
    return _origXhrSend.apply(this, [body]);
  };

  // ==================== CONFIG ====================
  const WAIT_AFTER_PAGE_CLICK = 2500;
  const WAIT_AFTER_FILTER_CLICK = 2500;
  const WAIT_AFTER_SEARCH = 2500;
  const DONE_TIMEOUT_MS = 30 * 60 * 1000;
  const LABEL_STATUS_PRINTED     = 50;
  const LABEL_STATUS_NOT_PRINTED = 30;

  // ==================== STATE ====================
  const ALIAS_STORAGE_KEY = 'qf_product_aliases_v1';
  const VARIANT_ALIAS_STORAGE_KEY = 'qf_variant_aliases_v1';
  const OVERLAY_PREF_KEY = 'qf_overlay_enabled_v1';
  function loadOverlayPref() {
    const v = localStorage.getItem(OVERLAY_PREF_KEY);
    return v === null ? true : v === 'true';
  }
  function saveOverlayPref(enabled) {
    localStorage.setItem(OVERLAY_PREF_KEY, String(!!enabled));
  }

  const DONE_STORAGE_KEY = 'qf_done_items_v1';
  function loadDoneItems() {
    try {
      const raw = localStorage.getItem(DONE_STORAGE_KEY);
      if (!raw) return new Map();
      const obj = JSON.parse(raw);
      const m = new Map();
      const cutoff = Date.now() - (30 * 60 * 1000);
      for (const [k, ts] of Object.entries(obj)) {
        if (typeof ts === 'number' && ts >= cutoff) m.set(k, ts);
      }
      return m;
    } catch { return new Map(); }
  }
  function saveDoneItems() {
    try {
      localStorage.setItem(DONE_STORAGE_KEY, JSON.stringify(Object.fromEntries(state.doneItems)));
    } catch {}
  }

  const state = {
    scanning: false,
    products: new Map(),
    weirdOrders: [],
    currentTab: 'single',
    autoSelectAll: true,
    doneItems: loadDoneItems(),    // key: "type:productId:skuId" → timestamp; persisted
    labelStatusFilter: 'not_printed', // 'all' | 'not_printed' | 'printed'
    aliases: loadAliases(),        // productId → aliasText (string)
    variantAliases: loadVariantAliases(), // `${productId}:${skuId}` → {alias: string, replace: boolean}
    fontUrl: null,                 // Sarabun-Bold.ttf URL from asset-bridge
    fontBytes: null,               // cached ArrayBuffer
    records: new Map(),            // fulfillUnitId → {skuList:[{productId,productName,quantity,...}]}
    weirdFulfillUnitIds: new Set(),
    weirdCombos: new Map(),        // sigKey → {sigKey, items:[{productId,productName,productImageURL,quantity}], fulfillUnitIds: Set, count}
    carriers: new Map(),           // carrierId → {id, name, iconUrl, count}
    carrierOf: new Map(),          // fulfillUnitId → carrierId
    carrierFilter: new Set(),      // empty = all carriers
    preOrderOf: new Map(),         // fulfillUnitId → boolean (true = pre-order)
    preOrderFilter: 'all',         // 'all' | 'preorder' | 'normal'
    selectMode: false,
    selected: new Map(),           // key → {type, productId, skuId, scenario, sigKey}
    dateFilter: { start: null, end: null, field: 'createTime' }, // field: createTime | shipByTime | autoCancelTime
    advancedOpen: false,
    calendarMonth: null, // {year, month} cursor for the visible month
    overlayEnabled: loadOverlayPref(),
  };

  function loadAliases() {
    try {
      const raw = localStorage.getItem(ALIAS_STORAGE_KEY);
      return raw ? new Map(Object.entries(JSON.parse(raw))) : new Map();
    } catch { return new Map(); }
  }

  function saveAliases() {
    const obj = Object.fromEntries(state.aliases);
    localStorage.setItem(ALIAS_STORAGE_KEY, JSON.stringify(obj));
  }

  function loadVariantAliases() {
    try {
      const raw = localStorage.getItem(VARIANT_ALIAS_STORAGE_KEY);
      return raw ? new Map(Object.entries(JSON.parse(raw))) : new Map();
    } catch { return new Map(); }
  }

  function saveVariantAliases() {
    const obj = Object.fromEntries(state.variantAliases);
    localStorage.setItem(VARIANT_ALIAS_STORAGE_KEY, JSON.stringify(obj));
  }

  function variantKey(productId, skuId) { return `${productId}:${skuId}`; }
  function getVariantInfo(productId, skuId) {
    return state.variantAliases.get(variantKey(productId, skuId)) || null;
  }
  function setVariantInfo(productId, skuId, partial) {
    const key = variantKey(productId, skuId);
    const cur = state.variantAliases.get(key) || {alias: '', replace: false};
    const next = {...cur, ...partial};
    if (!next.alias?.trim() && !next.replace) state.variantAliases.delete(key);
    else state.variantAliases.set(key, next);
    saveVariantAliases();
  }

  // Listen for messages from asset-bridge (ISOLATED world) — handles font URL,
  // local manifest version, and remote update check (CSP-safe in ISOLATED).
  const REPO_URL = 'https://github.com/Sittipanpee/tiktok-quick-filter';
  state.localVersion = null;
  state.remoteVersion = null;
  window.addEventListener('message', (e) => {
    if (e.source !== window) return;
    const d = e.data;
    if (!d?.__qfAsset) return;
    if (d.__qfAsset === 'font' && d.url) state.fontUrl = d.url;
    if (d.__qfAsset === 'manifest' && d.version) state.localVersion = d.version;
    if (d.__qfAsset === 'update' && d.remoteVersion) {
      state.remoteVersion = d.remoteVersion;
      if (d.localVersion) state.localVersion = d.localVersion;
      if (state.remoteVersion && state.localVersion && state.remoteVersion !== state.localVersion) {
        renderUpdateBadge();
      }
    }
  });
  window.postMessage({ __qfAsset: 'request_font' }, '*');

  function renderUpdateBadge() {
    const actions = document.getElementById('qf-header-actions');
    if (!actions || actions.querySelector('#qf-update-btn')) return;
    const btn = document.createElement('button');
    btn.id = 'qf-update-btn';
    btn.title = `อัพเดตใหม่ v${state.remoteVersion} (ปัจจุบัน v${state.localVersion}) — คลิกเพื่อดาวน์โหลด`;
    btn.textContent = `↑ v${state.remoteVersion}`;
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      window.open(REPO_URL + '#-ติดตั้ง', '_blank');
    });
    actions.insertBefore(btn, actions.firstChild);
  }
  // Expose state so the top-level fetch hook can write apiListUrl into it
  window.__qfState = state;

  // ==================== PAGE DETECTION ====================
  const isShopee = () => location.hostname === 'seller.shopee.co.th';
  const isTikTok = () => location.hostname === 'seller-th.tiktok.com';
  const isLabelsPage = () => isTikTok() && /\/shipment\/labels/.test(location.pathname)
                            || isShopee() && /\/portal\/sale/.test(location.pathname);
  const isOrderPage  = () => isTikTok() && /\/order/.test(location.pathname);

  // ==================== UTIL ====================
  const sleep = (ms) => new Promise(r => setTimeout(r, ms));

  async function safeJson(resp, label = 'API') {
    const text = await resp.text();
    if (!text) {
      throw new Error('เซิร์ฟเวอร์ตอบไม่ทัน ลองกดลองใหม่อีกครั้ง');
    }
    try { return JSON.parse(text); }
    catch {
      const isHtml = /^\s*<(!doctype|html)/i.test(text);
      if (isHtml) {
        throw new Error('การเชื่อมต่อสะดุด ลองกดลองใหม่ ถ้ายังไม่ได้ให้ refresh หน้านี้แล้วลองอีกครั้ง');
      }
      throw new Error(`เซิร์ฟเวอร์ตอบผิดรูปแบบ ลองกดลองใหม่ (${label})`);
    }
  }

  function doneKey(productId, skuId, type) {
    const base = skuId ? `${productId}:${skuId}` : productId;
    return `${type || ''}:${base}`;
  }

  function markDone(productId, skuId, type) {
    state.doneItems.set(doneKey(productId, skuId, type), Date.now());
    saveDoneItems();
  }

  function isDone(productId, skuId, type) {
    const key = doneKey(productId, skuId, type);
    const ts = state.doneItems.get(key);
    if (!ts) return false;
    if (Date.now() - ts > DONE_TIMEOUT_MS) { state.doneItems.delete(key); saveDoneItems(); return false; }
    return true;
  }

  function comboDoneKey(sigKey) { return `combo::${sigKey}`; }
  function markComboDone(sigKey) { state.doneItems.set(comboDoneKey(sigKey), Date.now()); saveDoneItems(); }
  function isComboDone(sigKey) {
    const ts = state.doneItems.get(comboDoneKey(sigKey));
    if (!ts) return false;
    if (Date.now() - ts > DONE_TIMEOUT_MS) { state.doneItems.delete(comboDoneKey(sigKey)); saveDoneItems(); return false; }
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

  async function ensurePageSize50() {
    const pagination = document.querySelector('.p-pagination');
    if (!pagination) return;
    const val = pagination.querySelector('.p-select-view-value')?.textContent?.trim();
    if (val === '50/Page') return;
    const select = pagination.querySelector('.p-select');
    if (!select) return;
    simulateClick(select);
    await sleep(600);
    const option = [...document.querySelectorAll('.p-dropdown-menu-item, .p-select-option, [role="option"]')]
      .find(el => /^50/.test(el.textContent.trim()));
    if (option) {
      simulateClick(option);
      await sleep(WAIT_AFTER_PAGE_CLICK);
    }
  }

  // ==================== SCANNING ====================
  const API_BATCH_SIZE = 100;   // orders per API call (TikTok accepts up to 500)
  const API_CONCURRENCY = 5;    // parallel requests

  async function awaitApiUrl(timeoutMs = 5000) {
    const deadline = Date.now() + timeoutMs;
    while (!_apiListUrl && Date.now() < deadline) await sleep(150);
    return _apiListUrl;
  }

  async function awaitBodyTemplate(ms = 1500) {
    const deadline = Date.now() + ms;
    while (!_apiListBodyTemplate && Date.now() < deadline) await sleep(100);
    return _apiListBodyTemplate;
  }

  async function ensureApiUrl() {
    if (_apiListUrl) {
      await awaitBodyTemplate();
      return _apiListUrl;
    }

    // Tab click may trigger a fetch if we weren't already on this tab
    await clickToShipTab();
    let url = await awaitApiUrl(3000);
    if (url) return url;

    // Fallback: click next page to force a fresh order/list fetch
    const nextBtn = document.querySelector('.p-pagination-item-next:not(.p-pagination-item-disabled)');
    if (nextBtn) {
      simulateClick(nextBtn);
      url = await awaitApiUrl(3000);
      if (url) {
        // Go back to page 1 so UI is consistent
        const p1 = [...document.querySelectorAll('.p-pagination-item')]
          .find(el => el.textContent.trim() === '1');
        if (p1) simulateClick(p1);
        return url;
      }
    }

    return null;
  }

  async function fetchOrderBatch(offset) {
    // Use captured body as template (correct format) — override only pagination fields
    const body = _apiListBodyTemplate
      ? { ..._apiListBodyTemplate, offset, count: API_BATCH_SIZE }
      : { offset, count: API_BATCH_SIZE };
    const resp = await _origFetch.call(window, _apiListUrl, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify(body),
    });
    return resp.json();
  }

  function extractImageUrl(img) {
    if (!img) return '';
    if (typeof img === 'string') return img;
    // API returns { url_list: [...], thumb_url_list: [...], uri, ... }
    return img.thumb_url_list?.[0] || img.url_list?.[0] || '';
  }

  function processApiOrders(orders) {
    if (!orders) return;
    for (const order of orders) {
      const skus = order.sku_module || order.skuModule || order.sku_list || order.skuList || order.items || [];
      for (const s of skus) {
        if (!state.products.has(s.product_id)) {
          state.products.set(s.product_id, {
            productId: s.product_id,
            productName: s.product_name || '(ไม่มีชื่อ)',
            productImageURL: extractImageUrl(s.product_image),
            variants: new Map(),
            orderCountSingle: 0,
            orderCountMulti: 0,
          });
        }
        const product = state.products.get(s.product_id);
        if (!product.variants.has(s.sku_id)) {
          product.variants.set(s.sku_id, {
            skuId: s.sku_id,
            skuName: s.sku_name || '',
            sellerSkuName: s.seller_sku_name || s.sku_name || '',
            orderCountSingle: 0,
            orderCountMulti: 0,
          });
        }
      }
      if (skus.length === 1) {
        const s = skus[0];
        const product = state.products.get(s.product_id);
        const variant = product.variants.get(s.sku_id);
        if (s.quantity === 1) {
          product.orderCountSingle++;
          variant.orderCountSingle++;
        } else {
          product.orderCountMulti++;
          variant.orderCountMulti++;
        }
      } else if (skus.length > 1) {
        state.weirdOrders.push({
          orderId: order.main_order_id,
          skus: skus.map(s => ({
            productId: s.product_id,
            productName: s.product_name || '',
            productImageURL: s.product_image || '',
            quantity: s.quantity,
          })),
        });
      }
    }
  }

  async function scanAllPages() {
    if (state.scanning) return;
    state.scanning = true;
    state.products.clear();
    state.weirdOrders = [];
    // Keep doneItems across scans so "ที่พิมพ์แล้ว" markers survive page refresh
    state.records.clear();
    state.weirdFulfillUnitIds.clear();
    state.weirdCombos.clear();
    state.carriers.clear();
    state.carrierOf.clear();
    state.preOrderOf.clear();

    const statusEl = document.getElementById('qf-scan-status');
    const btn = document.getElementById('qf-scan-btn');
    btn.disabled = true;

    if (isLabelsPage()) {
      try {
        statusEl.textContent = 'กำลังสแกน...';
        if (isShopee()) await scanShopeePage(statusEl);
        else await scanLabelsPage(statusEl);
        statusEl.textContent = `✓ ${state.products.size} สินค้า | ${state.weirdOrders.length} แปลก`;
        renderAll();
      } catch (err) {
        statusEl.textContent = 'ผิดพลาด: ' + err.message;
        console.error('[QF] scan failed:', err);
      } finally {
        state.scanning = false;
        btn.disabled = false;
      }
      return;
    }

    try {
      statusEl.textContent = 'กำลังเตรียม API...';
      const url = await ensureApiUrl();
      if (!url) throw new Error('ไม่สามารถจับ API URL — ลองโหลดหน้าใหม่แล้วสแกนอีกครั้ง');

      // First batch to learn total_count
      statusEl.textContent = 'กำลังดึงออเดอร์...';
      const first = await fetchOrderBatch(0);
      if (first.code !== 0) {
        const tplStatus = _apiListBodyTemplate
          ? `template OK (${Object.keys(_apiListBodyTemplate).length} fields)`
          : 'NO TEMPLATE — body fallback ขาดข้อมูล';
        console.error('[QF] order/list failed:', first, 'template:', _apiListBodyTemplate);
        throw new Error(`code=${first.code} msg="${first.message || 'empty'}" — ${tplStatus}`);
      }

      const d = first.data || {};
      const total = d.total_count ?? d.total ?? 0;
      const firstOrders = d.main_orders || d.order_list || d.orders || [];
      if (total === 0 && firstOrders.length === 0) {
        throw new Error(`API ตอบกลับว่าง (keys: ${Object.keys(d).join(',') || 'none'}) — ลองโหลดหน้าใหม่`);
      }
      processApiOrders(firstOrders);
      statusEl.textContent = `สแกน ${Math.min(API_BATCH_SIZE, total || firstOrders.length)}/${total || '?'}...`;

      const offsets = [];
      for (let off = API_BATCH_SIZE; off < total; off += API_BATCH_SIZE) offsets.push(off);

      for (let i = 0; i < offsets.length; i += API_CONCURRENCY) {
        const group = offsets.slice(i, i + API_CONCURRENCY);
        const results = await Promise.all(group.map(fetchOrderBatch));
        for (const res of results) {
          if (res.code === 0) {
            const rd = res.data || {};
            processApiOrders(rd.main_orders || rd.order_list || rd.orders || []);
          }
        }
        const scanned = Math.min((i + API_CONCURRENCY + 1) * API_BATCH_SIZE, total);
        statusEl.textContent = `สแกน ${scanned}/${total}...`;
      }

      statusEl.textContent = `✓ ${total} ออเดอร์ | ${state.products.size} สินค้า | ${state.weirdOrders.length} แปลก`;
      renderAll();
    } catch (err) {
      statusEl.textContent = 'ผิดพลาด: ' + err.message;
    } finally {
      state.scanning = false;
      btn.disabled = false;
    }
  }

  // ==================== LABELS PAGE ====================
  function getRecordFromRow(row) {
    const fk = Object.keys(row).find(k => k.startsWith('__reactFiber'));
    if (!fk) return null;
    let node = row[fk];
    for (let i = 0; i < 50 && node; i++) {
      if (node.memoizedProps?.record?.skuList) return node.memoizedProps.record;
      node = node.return;
    }
    return null;
  }

  function labelStatusMatches(labelStatus) {
    if (state.labelStatusFilter === 'all') return true;
    if (state.labelStatusFilter === 'printed')     return labelStatus === LABEL_STATUS_PRINTED;
    if (state.labelStatusFilter === 'not_printed') return labelStatus === LABEL_STATUS_NOT_PRINTED;
    return true;
  }

  function ensureProduct(s) {
    if (!state.products.has(s.productId)) {
      state.products.set(s.productId, {
        productId: s.productId,
        productName: s.productName || '(ไม่มีชื่อ)',
        productImageURL: s.productImageURL || '',
        variants: new Map(),
        orderCountSingle: 0,
        orderCountMulti: 0,
        fulfillUnitIdsSingle: new Set(),
        fulfillUnitIdsMulti: new Set(),
      });
    }
    const p = state.products.get(s.productId);
    if (!p.variants.has(s.skuId)) {
      p.variants.set(s.skuId, {
        skuId: s.skuId,
        skuName: s.skuName || '',
        sellerSkuName: s.sellerSkuName || s.skuName || '',
        orderCountSingle: 0,
        orderCountMulti: 0,
        fulfillUnitIdsSingle: new Set(),
        fulfillUnitIdsMulti: new Set(),
      });
    }
    return p;
  }

  function processLabelRecord(rec) {
    if (!rec || !labelStatusMatches(rec.labelStatus)) return;
    const fulfillUnitId = rec.fulfillUnitId;
    const skus = rec.skuList || [];
    if (fulfillUnitId) {
      state.records.set(fulfillUnitId, {
        fulfillUnitId,
        createTime: rec.createTime || null,
        shipByTime: rec.shipByTime || null,
        autoCancelTime: rec.autoCancelTime || null,
        skuList: skus.map(s => ({
          productId: s.productId,
          skuId: s.skuId,
          productName: s.productName,
          skuName: s.skuName,
          sellerSkuName: s.sellerSkuName,
          quantity: s.quantity,
        })),
      });
      state.preOrderOf.set(fulfillUnitId, !!rec.isPreOrder);
      // Track carrier
      const sp = rec.shippingProviderInfo || rec.deliveryInfo?.shippingProvider;
      const carrierId = rec.deliveryInfo?.shippingProvider?.id || sp?.name || 'unknown';
      const name = sp?.name || 'ไม่ระบุ';
      const iconUrl = sp?.iconUrl || sp?.icon_url || '';
      if (!state.carriers.has(carrierId)) {
        state.carriers.set(carrierId, { id: carrierId, name, iconUrl, count: 0 });
      }
      state.carriers.get(carrierId).count++;
      state.carrierOf.set(fulfillUnitId, carrierId);
    }
    // Always register products + variants so they appear in lists / variant badges
    for (const s of skus) ensureProduct(s);

    if (skus.length === 1) {
      const s = skus[0];
      const product = state.products.get(s.productId);
      const variant = product.variants.get(s.skuId);
      if (s.quantity === 1) {
        product.orderCountSingle++; variant.orderCountSingle++;
        if (fulfillUnitId) {
          product.fulfillUnitIdsSingle.add(fulfillUnitId);
          variant.fulfillUnitIdsSingle.add(fulfillUnitId);
        }
      } else {
        product.orderCountMulti++; variant.orderCountMulti++;
        if (fulfillUnitId) {
          product.fulfillUnitIdsMulti.add(fulfillUnitId);
          variant.fulfillUnitIdsMulti.add(fulfillUnitId);
        }
      }
    } else if (skus.length > 1) {
      const items = skus.map(s => ({
        productId: s.productId,
        productName: s.productName || '',
        productImageURL: s.productImageURL || '',
        quantity: s.quantity,
      }));
      state.weirdOrders.push({
        orderId: rec.batchId || rec.orderIds?.[0] || '',
        fulfillUnitId,
        skus: items,
      });
      if (fulfillUnitId) state.weirdFulfillUnitIds.add(fulfillUnitId);

      // Group by combination signature
      const sorted = [...items].sort((a, b) => a.productId.localeCompare(b.productId));
      const sigKey = sorted.map(i => `${i.productId}:${i.quantity}`).join('|');
      if (!state.weirdCombos.has(sigKey)) {
        state.weirdCombos.set(sigKey, {
          sigKey,
          items: sorted,
          fulfillUnitIds: new Set(),
          count: 0,
        });
      }
      const combo = state.weirdCombos.get(sigKey);
      if (fulfillUnitId) combo.fulfillUnitIds.add(fulfillUnitId);
      combo.count++;
    }
  }

  async function awaitLabelsApiReady(ms = 8000) {
    const deadline = Date.now() + ms;
    while ((!_labelsApiUrl || !_labelsApiBodyTemplate) && Date.now() < deadline) await sleep(150);
    return _labelsApiUrl && _labelsApiBodyTemplate;
  }

  async function triggerLabelsApiCapture(statusEl) {
    // Force TikTok to refetch labels by clicking the page-1 pagination item or pressing F5-like reload via filter toggle
    statusEl.textContent = 'รอจับ API จาก TikTok...';
    // Try clicking page 2 then back to 1 to trigger fresh API call
    const items = getPageItems();
    const next = items.find(it => /^2$/.test(it.textContent.trim()));
    const first = items.find(it => /^1$/.test(it.textContent.trim()));
    if (next) { simulateClick(next); await sleep(1500); }
    if (first) { simulateClick(first); await sleep(1500); }
    return await awaitLabelsApiReady(3000);
  }

  async function scanLabelsPage(statusEl) {
    let ready = await awaitLabelsApiReady(3000);
    if (!ready) {
      ready = await triggerLabelsApiCapture(statusEl);
    }
    if (ready) {
      await scanLabelsByApi(_labelsApiUrl, statusEl);
    } else {
      // DOM fallback: paginate through all pages
      statusEl.textContent = '⚠️ ใช้ DOM fallback (ช้า) — API ไม่จับ';
      const totalPages = getTotalPages();
      for (let p = 1; p <= totalPages; p++) {
        if (p > 1) { await goToPage(p); await sleep(WAIT_AFTER_PAGE_CLICK); }
        scanLabelsDom();
        statusEl.textContent = `สแกน หน้า ${p}/${totalPages} (DOM fallback)`;
      }
    }
  }

  // ==================== SHOPEE ADAPTER ====================
  async function awaitShopeeApiReady(ms = 8000) {
    const deadline = Date.now() + ms;
    while ((!_shopeeIndexUrl || !_shopeeIndexBody) && Date.now() < deadline) await sleep(150);
    return _shopeeIndexUrl && _shopeeIndexBody;
  }

  async function triggerShopeeApiCapture(statusEl) {
    statusEl.textContent = 'รอจับ API ตามแท็บที่กำลังดู...';
    // Use any DOM element with a numeric pagination text — Shopee's pagination
    // markup varies, so look for raw text-only nodes with width/height suggesting a button
    const candidates = [...document.querySelectorAll('*')]
      .filter(el => el.children.length === 0
        && /^\d+$/.test((el.textContent||'').trim())
        && el.offsetWidth > 5 && el.offsetWidth < 60
        && el.offsetHeight > 5 && el.offsetHeight < 60);
    const next = candidates.find(el => (el.textContent||'').trim() === '2');
    const first = candidates.find(el => (el.textContent||'').trim() === '1');
    if (next) { next.click(); await sleep(2000); }
    if (first) { first.click(); await sleep(1500); }
    return await awaitShopeeApiReady(3000);
  }

  function resetShopeeCapture() {
    _shopeeIndexUrl = null;
    _shopeeIndexBody = null;
    _shopeeCardUrl = null;
    _shopeeCardBody = null;
  }

  function buildShopeeSkuList(items) {
    return items.map(it => ({
      productId: String(it.inner_item_ext_info?.item_id || ''),
      skuId: String(it.inner_item_ext_info?.model_id || it.inner_item_ext_info?.item_id || ''),
      productName: it.name || '',
      skuName: it.model_name || it.variation_name || '',
      sellerSkuName: it.item_sku || '',
      productImageURL: it.image
        ? `https://down-th.img.susercontent.com/file/${it.image}_tn`
        : '',
      quantity: it.amount || 1,
    })).filter(s => s.productId);
  }

  function pushShopeeRecord({fulfillUnitId, ext, items, fulfilment, batchId}) {
    if (!fulfillUnitId) return;
    const skuList = buildShopeeSkuList(items);
    if (!skuList.length) return;
    // Shopee scan only happens on to-ship tab — every record is fair game,
    // mark all as NOT_PRINTED so the default status filter shows them.
    // TikTok-style label status doesn't map cleanly to Shopee's logistics_status.
    const labelStatus = LABEL_STATUS_NOT_PRINTED;
    const sp = fulfilment || {};
    const carrierName = sp.fulfilment_channel_name || sp.masked_channel_name || 'ไม่ระบุ';
    const carrierId = String(sp.fulfilment_channel_name || ext.masked_channel_id || 'unknown');
    processLabelRecord({
      fulfillUnitId: String(fulfillUnitId),
      batchId: String(batchId || fulfillUnitId),
      orderIds: [String(ext.order_id || '')],
      labelStatus,
      skuList,
      shippingProviderInfo: { name: carrierName, iconUrl: '' },
      deliveryInfo: { shippingProvider: { id: carrierId, name: carrierName, icon_url: '' } },
    });
  }

  function processShopeeRecord(card) {
    // Tab 300 (to-ship): package_card — single flat package
    if (card?.package_card) {
      const pc = card.package_card;
      const ext = pc.order_ext_info || {};
      const items = (pc.item_info_group?.item_info_list || []).flatMap(g => g.item_list || []);
      const pkg = pc.package_ext_info || {};
      pushShopeeRecord({
        fulfillUnitId: pkg.package_number || ext.order_id,
        batchId: pkg.package_number || ext.order_id,
        ext,
        items,
        fulfilment: pc.fulfilment_info,
      });
      return;
    }
    // Some tabs use order_card (single-package shape)
    if (card?.order_card) {
      const oc = card.order_card;
      const ext = oc.order_ext_info || {};
      const items = (oc.item_info_group?.item_info_list || []).flatMap(g => g.item_list || []);
      const pkg = oc.package_ext_info_list?.[0] || oc.package_ext_info || {};
      pushShopeeRecord({
        fulfillUnitId: pkg.package_number || ext.order_id,
        batchId: pkg.package_number || ext.order_id,
        ext,
        items,
        fulfilment: oc.fulfilment_info,
      });
      return;
    }
    // Tab 100 (all): package_level_order_card with package_list[]
    if (card?.package_level_order_card) {
      const plc = card.package_level_order_card;
      const ext = plc.order_ext_info || {};
      for (const pkg of (plc.package_list || [])) {
        const items = (pkg.item_info_group?.item_info_list || []).flatMap(g => g.item_list || []);
        const pkgInfo = pkg.package_ext_info || {};
        pushShopeeRecord({
          fulfillUnitId: pkgInfo.package_number || ext.order_id,
          batchId: pkgInfo.package_number || ext.order_id,
          ext,
          items,
          fulfilment: pkg.fulfilment_info,
        });
      }
    }
  }

  async function fetchShopeePage(pageNumber, pageSize = 40) {
    const baseBody = _shopeeIndexBody || {};
    const body = {
      ...baseBody,
      pagination: {
        ...(baseBody.pagination || {}),
        from_page_number: 1,
        page_number: pageNumber,
        page_size: pageSize,
      },
    };
    const r = await _origFetch.call(window, _shopeeIndexUrl, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify(body),
    });
    return r.json();
  }

  async function fetchShopeeCards(packageOrOrderList) {
    const baseBody = _shopeeCardBody || { order_list_tab: 100, need_count_down_desc: true };
    // Decide whether template uses package_param_list or order_param_list
    const usePackageParam = Array.isArray(baseBody.package_param_list);
    const param_list = packageOrOrderList.map(it => ({
      ...(usePackageParam
        ? { package_number: it.package_number || it.id, shop_id: it.shop_id, region_id: it.region_id || 'TH' }
        : { order_id: it.order_id || it.id, shop_id: it.shop_id, region_id: it.region_id || 'TH' }),
    }));
    const body = {
      ...baseBody,
      ...(usePackageParam
        ? { package_param_list: param_list, order_param_list: undefined }
        : { order_param_list: param_list, package_param_list: undefined }),
    };
    const r = await _origFetch.call(window, _shopeeCardUrl, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify(body),
    });
    return r.json();
  }

  function extractShopeeOrderRefs(indexData) {
    // Real key is index_list (returns [{order_id, shop_id, region_id}] or for tab 300 may include package_number)
    const list = indexData?.index_list || indexData?.list || [];
    return list.map(it => ({
      order_id: it.order_id || it.id,
      package_number: it.package_number || it.primary_package_number || it.ofg_id,
      shop_id: it.shop_id,
      region_id: it.region_id || 'TH',
    })).filter(r => r.order_id || r.package_number);
  }

  // Hardcoded fallback for to-ship tab (works without any prior API capture)
  const SHOPEE_TOSHIP_DEFAULT = {
    indexUrl: '/api/v3/order/search_order_list_index?SPC_CDS_VER=2',
    indexBody: {
      order_list_tab: 300,
      entity_type: 1,
      pagination: { from_page_number: 1, page_number: 1, page_size: 40 },
      filter: { fulfillment_type: 0, is_drop_off: 0, fulfillment_source: 0, action_filter: 0, shipping_priority: 0 },
      sort: { sort_type: 2, ascending: true },
    },
    cardUrl: '/api/v3/order/get_order_list_card_list?SPC_CDS_VER=2',
    cardBody: {
      order_list_tab: 300,
      need_count_down_desc: true,
      order_param_list: [],
    },
  };

  async function scanShopeePage(statusEl) {
    // Re-capture so body reflects the user's current Shopee tab/filter view.
    resetShopeeCapture();
    statusEl.textContent = 'รอจับ API ตามแท็บที่กำลังดู...';
    await triggerShopeeApiCapture(statusEl);

    // If still nothing captured (no pagination on the page, or single-page list),
    // fall back to a sensible default targeting the to-ship tab. This makes
    // scan work even if the user just opened the page and clicked Scan.
    if (!_shopeeIndexUrl || !_shopeeIndexBody) {
      _shopeeIndexUrl = SHOPEE_TOSHIP_DEFAULT.indexUrl;
      _shopeeIndexBody = SHOPEE_TOSHIP_DEFAULT.indexBody;
      statusEl.textContent = 'ใช้ฟิลเตอร์เริ่มต้น (ที่ต้องจัดส่ง)...';
    }
    if (!_shopeeCardUrl || !_shopeeCardBody) {
      _shopeeCardUrl = SHOPEE_TOSHIP_DEFAULT.cardUrl;
      _shopeeCardBody = SHOPEE_TOSHIP_DEFAULT.cardBody;
    }

    statusEl.textContent = 'กำลังดึงออเดอร์...';
    const PAGE_SIZE = 40;
    const first = await fetchShopeePage(1, PAGE_SIZE);
    if (first.code !== 0) throw new Error('Shopee API: ' + (first.msg || first.message || first.code));
    const fd = first.data || {};
    const total = fd.pagination?.total ?? fd.total_count ?? fd.total ?? 0;
    let firstRefs = extractShopeeOrderRefs(fd);
    if (!firstRefs.length) {
      throw new Error('Shopee ตอบว่าง (keys: ' + Object.keys(fd).join(',') + ')');
    }

    const cardResp = await fetchShopeeCards(firstRefs);
    if (cardResp.code === 0) {
      (cardResp.data?.card_list || []).forEach(c => processShopeeRecord(c));
    }
    statusEl.textContent = `สแกน ${Math.min(PAGE_SIZE, total)}/${total || '?'}...`;

    const totalPages = Math.ceil(total / PAGE_SIZE);
    const SHOPEE_CONCURRENCY = 4;
    const pagesToFetch = [];
    for (let page = 2; page <= totalPages; page++) pagesToFetch.push(page);
    for (let i = 0; i < pagesToFetch.length; i += SHOPEE_CONCURRENCY) {
      const group = pagesToFetch.slice(i, i + SHOPEE_CONCURRENCY);
      const indexResults = await Promise.all(group.map(p => fetchShopeePage(p, PAGE_SIZE)));
      const allRefs = indexResults.flatMap(r => r.code === 0 ? extractShopeeOrderRefs(r.data || {}) : []);
      if (allRefs.length) {
        // Batch card fetches in groups of 40
        const cardBatches = [];
        for (let j = 0; j < allRefs.length; j += 40) cardBatches.push(allRefs.slice(j, j + 40));
        const cardResults = await Promise.all(cardBatches.map(refs => fetchShopeeCards(refs)));
        for (const cards of cardResults) {
          if (cards.code === 0) {
            (cards.data?.card_list || []).forEach(c => processShopeeRecord(c));
          }
        }
      }
      const scanned = Math.min((i + SHOPEE_CONCURRENCY + 1) * PAGE_SIZE, total);
      statusEl.textContent = `สแกน ${scanned}/${total}...`;
    }
  }

  function getApiList(d) {
    return d.seller_packages_list || d.list || d.labels || d.orders || [];
  }

  function normalizeApiRecord(rec) {
    // Convert snake_case API record into the shape processLabelRecord expects (camelCase + skuList)
    const lm = rec.label_module || {};
    const dm = rec.delivery_module || {};
    const fm = rec.fulfillment_module || {};
    const sp = dm.shipment_provider_info || dm.shipping_provider_info || {};
    const skuList = (rec.sku_module || []).map(s => ({
      productId: s.product_id,
      skuId: s.sku_id,
      productName: s.product_name,
      skuName: s.sku_name,
      sellerSkuName: s.seller_sku_name,
      productImageURL: s.product_image?.thumb_url_list?.[0] || s.product_image?.url_list?.[0] || '',
      quantity: s.quantity || 1,
    }));
    // isPreOrder: any line marked → record is pre-order
    const olm = rec.order_label_module || [];
    const isPreOrder = olm.some(o => o.isPreOrder === 1 || o.is_pre_order === 1);
    return {
      fulfillUnitId: rec.fulfill_unit_id || lm.fulfill_unit_id,
      batchId: lm.batch_id,
      orderIds: lm.order_ids || rec.order_ids,
      labelStatus: lm.label_status,
      isPreOrder,
      // Time fields — all converted to ms timestamps (some sources are in seconds)
      // createTime: เวลาลูกค้าสร้างออเดอร์
      createTime: fm.create_time ? Number(fm.create_time)
                : (rec.trade_order_module?.[0]?.create_time ? Number(rec.trade_order_module[0].create_time) * 1000
                : (lm.purchase_time ? Number(lm.purchase_time) * 1000 : null)),
      // shipByTime: deadline ที่ต้องจัดส่งภายใน
      shipByTime: rec.trade_order_module?.[0]?.latest_tts_time
                ? Number(rec.trade_order_module[0].latest_tts_time) * 1000 : null,
      // autoCancelTime: เวลาที่จะถูกยกเลิกอัตโนมัติ
      autoCancelTime: rec.trade_order_module?.[0]?.close_sla_time
                ? Number(rec.trade_order_module[0].close_sla_time) * 1000 : null,
      skuList,
      shippingProviderInfo: {
        name: sp.name,
        iconUrl: sp.icon_url || sp.iconUrl,
      },
      deliveryInfo: {
        shippingProvider: {
          id: sp.id,
          name: sp.name,
          icon_url: sp.icon_url,
        },
      },
    };
  }

  async function scanLabelsByApi(apiUrl, statusEl) {
    const COUNT = 100;
    const makeBody = (offset) => {
      const base = _labelsApiBodyTemplate || {};
      // Inject extension filters as server-side search_condition so API pre-filters
      // (otherwise we hit the 10,000-record cap on unfiltered queries and miss records).
      const cl = {};
      if (state.labelStatusFilter === 'not_printed') cl.fulfillment_label_status = { value: ['30'] };
      else if (state.labelStatusFilter === 'printed') cl.fulfillment_label_status = { value: ['50'] };
      if (state.preOrderFilter === 'preorder') cl.fulfillment_order_label = { value: ['1'] };
      // 'normal' (fulfillment_order_label=0) is filtered client-side since the server
      // value for "not pre-order" isn't documented; client-side handles it.
      return {
        ...base,
        search_condition: { ...(base.search_condition || {}), scene: 1, condition_list: cl },
        offset,
        count: COUNT,
      };
    };

    const firstResp = await _origFetch.call(window, apiUrl, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify(makeBody(0)),
    });
    const first = await safeJson(firstResp, 'Labels list');
    if (first.code !== 0) throw new Error('Labels API error: ' + (first.message || first.code));

    const d = first.data || {};
    const total = d.total_count ?? d.total ?? 0;
    getApiList(d).forEach(r => processLabelRecord(normalizeApiRecord(r)));
    statusEl.textContent = `สแกน ${Math.min(COUNT, total)}/${total || '?'}...`;

    const offsets = [];
    for (let off = COUNT; off < total; off += COUNT) offsets.push(off);
    for (let i = 0; i < offsets.length; i += 5) {
      const results = await Promise.all(offsets.slice(i, i + 5).map(off =>
        _origFetch.call(window, apiUrl, {
          method: 'POST',
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify(makeBody(off)),
        }).then(r => safeJson(r, 'Labels list').catch(e => ({ code: -1, _err: e.message })))
      ));
      for (const res of results) {
        if (res.code === 0) {
          getApiList(res.data || {}).forEach(r => processLabelRecord(normalizeApiRecord(r)));
        }
      }
      statusEl.textContent = `สแกน ${Math.min((i + 6) * COUNT, total)}/${total}...`;
    }
  }

  function scanLabelsDom() {
    const rows = [...document.querySelectorAll('tbody tr')];
    for (const row of rows) {
      const rec = getRecordFromRow(row);
      if (rec) processLabelRecord(rec);
    }
  }

  const PRINT_BATCH_SIZE = 500; // TikTok limit per generate call

  async function ensureFontBytes() {
    if (state.fontBytes) return state.fontBytes;
    if (!state.fontUrl) {
      // Wait briefly for asset-bridge
      const deadline = Date.now() + 2000;
      while (!state.fontUrl && Date.now() < deadline) await sleep(100);
    }
    if (!state.fontUrl) throw new Error('ไม่พบ font URL (asset-bridge ยังไม่ทำงาน)');
    const r = await fetch(state.fontUrl);
    state.fontBytes = await r.arrayBuffer();
    return state.fontBytes;
  }

  function makeBaseFilename(hint) {
    const now = new Date();
    const pad = n => String(n).padStart(2, '0');
    const stamp = `${now.getFullYear()}${pad(now.getMonth()+1)}${pad(now.getDate())}${pad(now.getHours())}${pad(now.getMinutes())}`;
    const clean = (hint || 'labels')
      .replace(/[\\/:*?"<>|]/g, '')
      .replace(/\s+/g, ' ')
      .trim()
      .slice(0, 60);
    return `${clean} ${stamp}`;
  }

  function showChunkChoiceModal({total}) {
    const defaultChunks = Math.max(1, Math.ceil(total / 200));
    return new Promise(resolve => {
      const overlay = document.createElement('div');
      overlay.className = 'qf-modal-overlay qf-chunk-modal-overlay';
      const presets = [1, 2, defaultChunks, defaultChunks * 2].filter((n, i, a) => n > 0 && a.indexOf(n) === i).sort((a,b) => a-b);
      overlay.innerHTML = `
        <div class="qf-modal qf-chunk-modal" role="dialog">
          <div class="qf-modal-title">แบ่งไฟล์เพื่อความปลอดภัย</div>
          <div class="qf-modal-body">
            <div class="qf-chunk-summary">${total} ฉลาก เยอะเกิน 200 ใบ — แนะนำแบ่งไฟล์</div>
            <div class="qf-chunk-hint">แบ่งเพื่อ: ป้องกันค้างเครื่อง / ถ้ามีปัญหากลางทาง ยังได้ไฟล์ที่เสร็จแล้ว</div>
            <div class="qf-chunk-presets">
              ${presets.map(n => `<button class="qf-chunk-preset" data-n="${n}">${n} ไฟล์<span class="qf-chunk-preset-sub">~${Math.ceil(total/n)} ใบ/ไฟล์</span></button>`).join('')}
            </div>
            <div class="qf-chunk-custom">
              หรือ ตั้งเอง: <input type="number" class="qf-chunk-input" min="1" max="${total}" value="${defaultChunks}"/> ไฟล์
              <span class="qf-chunk-preview"></span>
            </div>
          </div>
          <div class="qf-modal-actions">
            <button class="qf-btn-cancel">ยกเลิก</button>
            <button class="qf-btn-confirm">เริ่มพิมพ์</button>
          </div>
        </div>
      `;
      document.body.appendChild(overlay);
      const inp = overlay.querySelector('.qf-chunk-input');
      const preview = overlay.querySelector('.qf-chunk-preview');
      const sync = () => {
        const n = Math.max(1, Math.min(total, parseInt(inp.value) || 1));
        preview.textContent = `(~${Math.ceil(total/n)} ใบ/ไฟล์)`;
      };
      sync();
      inp.addEventListener('input', sync);
      overlay.querySelectorAll('.qf-chunk-preset').forEach(b => {
        b.addEventListener('click', () => {
          inp.value = b.dataset.n;
          sync();
          overlay.querySelectorAll('.qf-chunk-preset').forEach(x => x.classList.toggle('active', x === b));
        });
      });
      const cleanup = (val) => { overlay.remove(); resolve(val); };
      overlay.querySelector('.qf-btn-cancel').onclick = () => cleanup(null);
      overlay.querySelector('.qf-btn-confirm').onclick = () => {
        const n = Math.max(1, Math.min(total, parseInt(inp.value) || 1));
        cleanup(n);
      };
      overlay.onclick = (e) => { if (e.target === overlay) cleanup(null); };
      const onKey = (e) => { if (e.key === 'Escape') { cleanup(null); document.removeEventListener('keydown', onKey); } };
      document.addEventListener('keydown', onKey);
    });
  }

  function showChunkedResult({title, totalIds, chunks}) {
    // chunks: [{count, label, filename}]
    document.querySelectorAll('.qf-progress-overlay').forEach(e => e.remove());
    const overlay = document.createElement('div');
    overlay.className = 'qf-progress-overlay';
    const chunkRows = chunks.map((c, i) => `
      <div class="qf-chunk-row" data-i="${i}">
        <div class="qf-chunk-row-top">
          <div class="qf-chunk-row-label">${escapeHtml(c.label || `ชุด ${i+1}/${chunks.length}`)} <span class="qf-chunk-row-count">${c.count} ใบ</span></div>
          <div class="qf-chunk-row-actions"></div>
        </div>
        <div class="qf-chunk-row-status">รอคิว</div>
        <div class="qf-chunk-row-bar"><div class="qf-chunk-row-fill"></div></div>
      </div>
    `).join('');
    overlay.innerHTML = `
      <div class="qf-progress-card qf-chunked-card">
        <div class="qf-progress-title">${escapeHtml(title)}</div>
        <div class="qf-chunked-sub">${totalIds} ฉลาก • ${chunks.length} ไฟล์</div>
        <div class="qf-chunked-list">${chunkRows}</div>
        <div class="qf-chunked-footer">
          <button class="qf-btn-download-all" disabled>ดาวน์โหลดทั้งหมด</button>
          <button class="qf-btn-close-result" style="display:none;">ปิด</button>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);
    const cards = overlay.querySelectorAll('.qf-chunk-row');
    const blobUrls = {}; // chunkIdx → {url, filename}
    const downloadAllBtn = overlay.querySelector('.qf-btn-download-all');
    const closeBtn = overlay.querySelector('.qf-btn-close-result');

    const cleanup = () => {
      Object.values(blobUrls).forEach(({url}) => URL.revokeObjectURL(url));
      overlay.remove();
    };
    closeBtn.onclick = cleanup;
    overlay.onclick = (e) => { if (e.target === overlay && closeBtn.style.display !== 'none') cleanup(); };

    downloadAllBtn.onclick = () => {
      Object.values(blobUrls).forEach(({url, filename}, i) => {
        setTimeout(() => {
          const a = document.createElement('a');
          a.href = url; a.download = filename;
          document.body.appendChild(a); a.click(); a.remove();
        }, i * 200);
      });
    };

    return {
      startChunk(i) {
        const row = cards[i];
        row.querySelector('.qf-chunk-row-status').textContent = 'กำลังทำ...';
        row.classList.add('qf-chunk-active');
      },
      updateChunkProgress(i, pct, label) {
        const row = cards[i];
        row.querySelector('.qf-chunk-row-fill').style.width = (pct * 100).toFixed(0) + '%';
        if (label) row.querySelector('.qf-chunk-row-status').textContent = label;
      },
      completeChunk(i, {url, pageCount}) {
        const row = cards[i];
        const filename = chunks[i].filename || `chunk-${i+1}.pdf`;
        blobUrls[i] = {url, filename};
        row.querySelector('.qf-chunk-row-fill').style.width = '100%';
        row.querySelector('.qf-chunk-row-status').textContent = `${pageCount} หน้า`;
        row.classList.remove('qf-chunk-active');
        row.classList.add('qf-chunk-done');
        const actions = row.querySelector('.qf-chunk-row-actions');
        actions.innerHTML = `
          <button class="qf-chunk-btn qf-chunk-open">เปิด</button>
          <button class="qf-chunk-btn qf-chunk-download">ดาวน์โหลด</button>
        `;
        actions.querySelector('.qf-chunk-open').onclick = () => {
          const w = window.open(url, '_blank');
          if (!w) {
            const a = document.createElement('a');
            a.href = url; a.target = '_blank'; a.rel = 'noopener';
            document.body.appendChild(a); a.click(); a.remove();
          }
        };
        actions.querySelector('.qf-chunk-download').onclick = () => {
          const a = document.createElement('a');
          a.href = url; a.download = filename;
          document.body.appendChild(a); a.click(); a.remove();
        };
        if (Object.keys(blobUrls).length > 0) {
          downloadAllBtn.disabled = false;
        }
      },
      errorChunk(i, msg) {
        const row = cards[i];
        row.querySelector('.qf-chunk-row-status').textContent = 'ล้มเหลว: ' + msg;
        row.classList.remove('qf-chunk-active');
        row.classList.add('qf-chunk-error');
        const actions = row.querySelector('.qf-chunk-row-actions');
        actions.innerHTML = `<button class="qf-chunk-btn qf-chunk-retry">ลองใหม่</button>`;
        actions.querySelector('.qf-chunk-retry').onclick = async () => {
          // External retry handler — set via setRetryHandler
          if (this._retry) await this._retry(i);
        };
      },
      setRetryHandler(fn) { this._retry = fn; },
      allDone() {
        closeBtn.style.display = '';
      },
      cleanup,
    };
  }

  function showProgress(title) {
    document.querySelectorAll('.qf-progress-overlay').forEach(e => e.remove());
    const overlay = document.createElement('div');
    overlay.className = 'qf-progress-overlay';
    overlay.innerHTML = `
      <div class="qf-progress-card">
        <div class="qf-progress-title">${escapeHtml(title)}</div>
        <div class="qf-progress-bar"><div class="qf-progress-fill"></div></div>
        <div class="qf-progress-meta">
          <span class="qf-progress-percent">0%</span>
          <span class="qf-progress-status">เริ่มต้น...</span>
        </div>
        <div class="qf-progress-result" style="display:none;"></div>
      </div>
    `;
    document.body.appendChild(overlay);
    const card = overlay.querySelector('.qf-progress-card');
    const fill = overlay.querySelector('.qf-progress-fill');
    const pct = overlay.querySelector('.qf-progress-percent');
    const status = overlay.querySelector('.qf-progress-status');
    const result = overlay.querySelector('.qf-progress-result');
    return {
      update(percent, label) {
        const p = Math.max(0, Math.min(100, percent));
        fill.style.width = p + '%';
        pct.textContent = p.toFixed(0) + '%';
        if (label) status.textContent = label;
      },
      showResult({ blobUrl, filename, total, pageCount }) {
        card.querySelector('.qf-progress-title').textContent = '✓ พิมพ์เสร็จแล้ว';
        result.style.display = 'block';
        result.innerHTML = `
          <div class="qf-result-summary">
            ${total} ฉลาก · ${pageCount} หน้า
          </div>
          <div class="qf-result-hint">เบราว์เซอร์อาจบล็อก auto-popup → กดปุ่มข้างล่าง</div>
          <div class="qf-result-actions">
            <button class="qf-result-btn qf-result-open">📄 เปิด PDF</button>
            <button class="qf-result-btn qf-result-download">💾 ดาวน์โหลด</button>
            <button class="qf-result-btn qf-result-close">ปิด</button>
          </div>
        `;
        const cleanup = () => {
          if (blobUrl) URL.revokeObjectURL(blobUrl);
          overlay.remove();
        };
        result.querySelector('.qf-result-open').onclick = () => {
          if (!blobUrl) return;
          // Fresh user click → window.open works
          const w = window.open(blobUrl, '_blank');
          if (!w) {
            const a = document.createElement('a');
            a.href = blobUrl; a.target = '_blank'; a.rel = 'noopener';
            document.body.appendChild(a); a.click(); a.remove();
          }
        };
        result.querySelector('.qf-result-download').onclick = () => {
          if (!blobUrl) return;
          const a = document.createElement('a');
          a.href = blobUrl; a.download = filename;
          document.body.appendChild(a); a.click(); a.remove();
        };
        result.querySelector('.qf-result-close').onclick = cleanup;
        overlay.onclick = (e) => { if (e.target === overlay) cleanup(); };
      },
      close() { overlay.remove(); },
    };
  }

  function openAliasModal(productId) {
    const product = state.products.get(productId);
    if (!product) { showToast('ไม่พบสินค้า', 1500); return; }
    const variants = [...product.variants.values()];

    document.querySelectorAll('.qf-alias-modal-overlay').forEach(e => e.remove());
    const overlay = document.createElement('div');
    overlay.className = 'qf-modal-overlay qf-alias-modal-overlay';
    overlay.innerHTML = `
      <div class="qf-modal qf-alias-modal" role="dialog">
        <div class="qf-alias-modal-header">
          <img src="${product.productImageURL}" referrerpolicy="no-referrer"/>
          <div class="qf-alias-modal-title">
            <div class="qf-alias-modal-name">${escapeHtml(product.productName)}</div>
            <div class="qf-alias-modal-sub">${variants.length} ตัวเลือก</div>
          </div>
          <button class="qf-alias-modal-close" aria-label="ปิด">×</button>
        </div>

        <div class="qf-alias-modal-section">
          <div class="qf-alias-modal-label">ชื่อย่อหลัก</div>
          <div class="qf-alias-modal-hint">ใช้กับทุกตัวเลือกที่ไม่ได้ตั้งชื่อแยก</div>
          <input class="qf-alias-modal-product" type="text" placeholder="เช่น ครีม, แดง1, สครับ" maxlength="20" value="${escapeHtml(state.aliases.get(productId) || '')}"/>
        </div>

        <div class="qf-alias-modal-section qf-alias-modal-section-scroll">
          <div class="qf-alias-modal-label">ตั้งชื่อแยกตามตัวเลือก</div>
          <div class="qf-alias-modal-hint">ปล่อยว่างจะใช้ชื่อหลักแทน</div>
          <div class="qf-alias-modal-variants"></div>
        </div>

        <div class="qf-alias-modal-footer">
          <button class="qf-alias-modal-done">เสร็จ</button>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);

    const productInput = overlay.querySelector('.qf-alias-modal-product');
    const variantsWrap = overlay.querySelector('.qf-alias-modal-variants');

    const updatePreviews = () => {
      variantsWrap.querySelectorAll('.qf-av-row').forEach(row => {
        const skuId = row.dataset.skuId;
        const v = product.variants.get(skuId);
        const inp = row.querySelector('.qf-av-input');
        const chk = row.querySelector('.qf-av-replace');
        const prev = row.querySelector('.qf-av-preview');
        const productAlias = productInput.value.trim() || shortName(product.productName);
        const variantAlias = inp.value.trim();
        const variantName = variantAlias || (v.skuName || v.sellerSkuName || '').trim();
        if (chk.checked && variantAlias) {
          prev.innerHTML = `บนลาเบล: <b>${escapeHtml(variantAlias)} 1</b>`;
        } else {
          prev.innerHTML = `บนลาเบล: <b>${escapeHtml(productAlias)} 1</b><br><span class="qf-av-preview-small">${escapeHtml(variantName)}</span>`;
        }
      });
    };

    productInput.addEventListener('input', () => {
      const v = productInput.value.trim();
      if (v) state.aliases.set(productId, v); else state.aliases.delete(productId);
      saveAliases();
      updatePreviews();
    });

    for (const v of variants) {
      const info = getVariantInfo(productId, v.skuId) || {alias: '', replace: false};
      const row = document.createElement('div');
      row.className = 'qf-av-row';
      row.dataset.skuId = v.skuId;
      row.innerHTML = `
        <div class="qf-av-name" title="${escapeHtml(v.skuName || v.sellerSkuName || v.skuId)}">${escapeHtml(v.skuName || v.sellerSkuName || v.skuId)}</div>
        <div class="qf-av-controls">
          <input class="qf-av-input" type="text" placeholder="ปล่อยว่าง = ใช้ชื่อหลัก" maxlength="20" value="${escapeHtml(info.alias)}"/>
          <label class="qf-av-replace-label">
            <input class="qf-av-replace" type="checkbox" ${info.replace ? 'checked' : ''}/>
            <span>แสดงชื่อนี้แทนชื่อสินค้า</span>
          </label>
        </div>
        <div class="qf-av-preview"></div>
      `;
      const inp = row.querySelector('.qf-av-input');
      const chk = row.querySelector('.qf-av-replace');
      const commit = () => {
        setVariantInfo(productId, v.skuId, {alias: inp.value, replace: chk.checked});
        updatePreviews();
      };
      inp.addEventListener('input', commit);
      chk.addEventListener('change', commit);
      variantsWrap.appendChild(row);
    }
    updatePreviews();

    const cleanup = () => { overlay.remove(); renderAll(); };
    overlay.querySelector('.qf-alias-modal-close').onclick = cleanup;
    overlay.querySelector('.qf-alias-modal-done').onclick = cleanup;
    overlay.onclick = (e) => { if (e.target === overlay) cleanup(); };
    const onKey = (e) => { if (e.key === 'Escape') { cleanup(); document.removeEventListener('keydown', onKey); } };
    document.addEventListener('keydown', onKey);
  }

  function shortName(name) {
    if (!name) return '';
    const trimmed = name.replace(/^\[[^\]]*\]\s*/, '').trim();
    return trimmed.length > 8 ? trimmed.slice(0, 8) + '…' : trimmed;
  }

  function variantDisplayName(s) {
    return (s.skuName || s.sellerSkuName || '').trim();
  }

  function buildSkuRender(s) {
    // primary = big top line (alias + qty), secondary = small bottom (variant name only)
    const v = getVariantInfo(s.productId, s.skuId);
    const productAlias = (state.aliases.get(s.productId) || '').trim() || shortName(s.productName);
    const variantOverride = (v?.alias || '').trim();
    const variantName = variantOverride || variantDisplayName(s);
    const qty = s.quantity || 1;

    if (v?.replace && variantOverride) {
      return { primary: `${variantOverride} ${qty}`, secondary: '' };
    }
    return {
      primary: `${productAlias} ${qty}`,
      secondary: variantName,
    };
  }

  function buildPageLines(record) {
    if (!record?.skuList?.length) return null;
    if (record.skuList.length === 1) {
      const r = buildSkuRender(record.skuList[0]);
      return { mode: 'single', primary: r.primary, secondary: r.secondary };
    }
    // Multi-SKU: "alias variant qty + alias variant qty"
    const parts = record.skuList.map(s => {
      const v = getVariantInfo(s.productId, s.skuId);
      const productAlias = (state.aliases.get(s.productId) || '').trim() || shortName(s.productName);
      const variantOverride = (v?.alias || '').trim();
      const variantName = variantOverride || variantDisplayName(s);
      const qty = s.quantity || 1;
      if (v?.replace && variantOverride) return `${variantOverride} ${qty}`;
      return variantName ? `${productAlias} ${variantName} ${qty}` : `${productAlias} ${qty}`;
    });
    return { mode: 'multi', text: parts.join(' + ') };
  }

  async function overlayAliasOnPdf(pdfBytes, fulfillUnitIds, onProgress) {
    if (!window.PDFLib) return pdfBytes;
    const { PDFDocument, rgb } = window.PDFLib;
    const fontBytes = await ensureFontBytes();
    const pdfDoc = await PDFDocument.load(pdfBytes);
    if (window.fontkit) pdfDoc.registerFontkit(window.fontkit);
    const font = await pdfDoc.embedFont(fontBytes, { subset: true });
    const pages = pdfDoc.getPages();
    if (!pages.length || !fulfillUnitIds?.length) return await pdfDoc.save();

    // Pages may be grouped per fulfillUnit (e.g., shipping label + packing list = 2 pages per unit)
    const pagesPerUnit = Math.max(1, Math.round(pages.length / fulfillUnitIds.length));
    const total = pages.length;

    const fitWidth = (text, baseSize, maxWidth) => {
      let size = baseSize;
      let w = font.widthOfTextAtSize(text, size);
      if (w > maxWidth) { size = size * (maxWidth / w); w = font.widthOfTextAtSize(text, size); }
      return { size, width: w };
    };
    const draw = (page, text, y, baseSize) => {
      const { width } = page.getSize();
      const { size, width: tw } = fitWidth(text, baseSize, width - 16);
      page.drawText(text, {
        x: (width - tw) / 2, y, size, font,
        color: rgb(0, 0, 0), opacity: 0.55,
      });
    };

    for (let i = 0; i < pages.length; i++) {
      const unitIdx = Math.min(fulfillUnitIds.length - 1, Math.floor(i / pagesPerUnit));
      const rec = state.records.get(fulfillUnitIds[unitIdx]);
      const lines = buildPageLines(rec);
      const page = pages[i];
      const { height } = page.getSize();
      const bigSize = Math.min(height * 0.05, 22);
      const smallSize = Math.min(height * 0.032, 13);
      if (lines?.mode === 'single') {
        // big top, small below
        if (lines.secondary) {
          draw(page, lines.primary, 4 + smallSize + 2, bigSize); // top line above small
          draw(page, lines.secondary, 4, smallSize);             // bottom line
        } else {
          draw(page, lines.primary, 6, bigSize);
        }
      } else if (lines?.mode === 'multi') {
        draw(page, lines.text, 6, bigSize);
      }
      if (onProgress && (i % 20 === 0 || i === total - 1)) {
        onProgress(i + 1, total);
        await sleep(0);
      }
    }
    return await pdfDoc.save();
  }

  function passesPreOrder(id) {
    if (state.preOrderFilter === 'all') return true;
    const isPre = !!state.preOrderOf.get(id);
    return state.preOrderFilter === 'preorder' ? isPre : !isPre;
  }
  function passesCarrier(id) {
    if (state.carrierFilter.size === 0) return true;
    return state.carrierFilter.has(state.carrierOf.get(id));
  }
  function passesDate(id) {
    const { start, end, field } = state.dateFilter;
    if (start === null && end === null) return true;
    const t = state.records.get(id)?.[field || 'createTime'];
    if (!t) return false;
    if (start !== null && t < start) return false;
    if (end !== null && t >= end) return false; // end exclusive (next-day midnight)
    return true;
  }
  function applyCarrierFilter(ids) {
    return ids.filter(id => passesCarrier(id) && passesPreOrder(id) && passesDate(id));
  }

  // ==================== MULTI-SELECT ====================
  function selectionKey(item) {
    if (item.type === 'combo') return `combo:${item.sigKey}`;
    if (item.type === 'variant') return `var:${item.productId}:${item.skuId}:${item.scenario}`;
    return `prod:${item.productId}:${item.scenario}`;
  }

  function isSelected(item) { return state.selected.has(selectionKey(item)); }

  function toggleSelection(item) {
    const key = selectionKey(item);
    if (state.selected.has(key)) state.selected.delete(key);
    else state.selected.set(key, item);
    updateSelectionBar();
  }

  function clearSelection() {
    state.selected.clear();
    renderAll();
  }

  function setSelectMode(on) {
    state.selectMode = on;
    if (!on) state.selected.clear();
    renderAll();
  }

  function resolveSelectedIds() {
    const set = new Set();
    for (const item of state.selected.values()) {
      let ids = [];
      if (item.type === 'combo') {
        const combo = state.weirdCombos.get(item.sigKey);
        if (combo) ids = applyCarrierFilter([...combo.fulfillUnitIds]);
      } else {
        ids = collectFulfillIds(item.productId, item.skuId, item.scenario);
      }
      for (const id of ids) set.add(id);
    }
    return [...set];
  }

  function describeItem(item) {
    if (item.type === 'combo') {
      const c = state.weirdCombos.get(item.sigKey);
      if (!c) return {label: 'combo', filename: 'combo'};
      const parts = c.items.map(i => (state.aliases.get(i.productId) || '').trim() || shortName(i.productName));
      return {label: parts.join(' + '), filename: parts.join('+')};
    }
    const p = state.products.get(item.productId);
    const alias = (state.aliases.get(item.productId) || '').trim();
    const baseName = alias || shortName(p?.productName);
    if (item.type === 'variant') {
      const v = p?.variants.get(item.skuId);
      const variantInfo = getVariantInfo(item.productId, item.skuId);
      const variantName = (variantInfo?.alias || '').trim() || (v?.skuName || v?.sellerSkuName || '');
      const tag = item.scenario === 'multi' ? ' (จำนวน>1)' : '';
      return {
        label: `${baseName} · ${variantName}${tag}`,
        filename: `${baseName} ${variantName}`.trim(),
      };
    }
    const tag = item.scenario === 'multi' ? ' (จำนวน>1)' : '';
    return {label: `${baseName}${tag}`, filename: baseName};
  }

  function getItemIds(item) {
    if (item.type === 'combo') {
      const c = state.weirdCombos.get(item.sigKey);
      return c ? applyCarrierFilter([...c.fulfillUnitIds]) : [];
    }
    return collectFulfillIds(item.productId, item.skuId, item.scenario);
  }

  async function printSelected() {
    const items = [...state.selected.values()];
    if (!items.length) { showToast('ยังไม่ได้เลือกอะไร', 1500); return; }

    // Build per-item chunks (1 file each = no mixing across products).
    // If a single item has >200 IDs, split it into N sub-chunks of <=200 each.
    const SUB_CHUNK_THRESHOLD = 200;
    const chunks = [];
    for (const it of items) {
      const ids = getItemIds(it);
      if (!ids.length) continue;
      const {label, filename} = describeItem(it);
      if (ids.length <= SUB_CHUNK_THRESHOLD) {
        chunks.push({item: it, ids, label, filename});
      } else {
        const subCount = Math.ceil(ids.length / SUB_CHUNK_THRESHOLD);
        const subSize = Math.ceil(ids.length / subCount);
        for (let i = 0; i < ids.length; i += subSize) {
          const slice = ids.slice(i, i + subSize);
          const idx = chunks.filter(c => c.item === it).length + 1;
          chunks.push({
            item: it,
            ids: slice,
            label: `${label} (${idx}/${subCount})`,
            filename: `${filename}-ชุด${idx}-${subCount}`,
          });
        }
      }
    }

    if (!chunks.length) { showToast('รายการที่เลือกไม่มีฉลาก', 2000); return; }

    const totalIds = chunks.reduce((a, c) => a + c.ids.length, 0);
    const sample = chunks.slice(0, 3).map(c => c.label).join(', ')
      + (chunks.length > 3 ? ` +อีก ${chunks.length - 3}` : '');

    const confirmed = await showPrintConfirm({
      title: `พิมพ์ ${chunks.length} ไฟล์ (1 ไฟล์/รายการ)`,
      summary: 'แยกไฟล์ตามสินค้า — ไม่ปนกัน',
      count: totalIds,
      sampleText: sample,
    });
    if (!confirmed) return;

    const stamp = makeBaseFilename('').trim(); // YYYYMMDDHHMM only
    const exportChunks = chunks.map(c => ({
      ids: c.ids,
      label: c.label,
      filename: `${makeBaseFilename(c.filename)}.pdf`,
    }));

    try {
      const ok = await runChunkedExport(exportChunks, `พิมพ์รวม ${chunks.length} ไฟล์`);
      if (ok) {
        for (const c of chunks) {
          const it = c.item;
          if (it.type === 'combo') markComboDone(it.sigKey);
          else {
            const type = it.scenario === 'multi' ? 'single_sku' : 'single_item';
            markDone(it.productId, it.skuId || null, type);
          }
        }
        state.selected.clear();
        state.selectMode = false;
        renderAll();
      }
    } catch (e) {
      console.error('[QF] printSelected failed:', e);
      showToast('พิมพ์ผิดพลาด: ' + e.message, 4000);
    }
  }

  function updateSelectionBar() {
    const bar = document.getElementById('qf-select-bar');
    if (!bar) return;
    const count = state.selected.size;
    if (count === 0) { bar.style.display = 'none'; return; }
    bar.style.display = 'flex';
    const ids = resolveSelectedIds();
    bar.querySelector('.qf-select-bar-count').textContent = `เลือก ${count} รายการ • ${ids.length} ฉลาก`;
  }

  function carrierFilteredSize(idSet) {
    if (!idSet) return 0;
    const dateActive = state.dateFilter.start !== null || state.dateFilter.end !== null;
    if (state.carrierFilter.size === 0 && state.preOrderFilter === 'all' && !dateActive) return idSet.size;
    let n = 0;
    for (const id of idSet) {
      if (passesCarrier(id) && passesPreOrder(id) && passesDate(id)) n++;
    }
    return n;
  }

  function collectFulfillIds(productId, skuId, scenario) {
    // scenario: 'single' (1 SKU qty=1) | 'multi' (1 SKU qty>1)
    const product = state.products.get(productId);
    if (!product) return [];
    let ids;
    if (skuId) {
      const v = product.variants.get(skuId);
      if (!v) return [];
      if (scenario === 'single')      ids = [...v.fulfillUnitIdsSingle];
      else if (scenario === 'multi')  ids = [...v.fulfillUnitIdsMulti];
      else                            ids = [...v.fulfillUnitIdsSingle, ...v.fulfillUnitIdsMulti];
    } else {
      if (scenario === 'single')      ids = [...product.fulfillUnitIdsSingle];
      else if (scenario === 'multi')  ids = [...product.fulfillUnitIdsMulti];
      else                            ids = [...product.fulfillUnitIdsSingle, ...product.fulfillUnitIdsMulti];
    }
    return applyCarrierFilter(ids);
  }

  function showPrintConfirm({ title, summary, count, sampleText }) {
    return new Promise(resolve => {
      const overlay = document.createElement('div');
      overlay.className = 'qf-modal-overlay';
      const overlayChecked = state.overlayEnabled ? 'checked' : '';
      overlay.innerHTML = `
        <div class="qf-modal" role="dialog">
          <div class="qf-modal-title">ยืนยันพิมพ์</div>
          <div class="qf-modal-body">
            <div class="qf-modal-target">${escapeHtml(title)}</div>
            ${summary ? `<div class="qf-modal-summary">${escapeHtml(summary)}</div>` : ''}
            <div class="qf-modal-count">${count} ฉลาก</div>
            ${sampleText ? `<div class="qf-modal-sample">ลายน้ำตัวอย่าง: <b>${escapeHtml(sampleText)}</b></div>` : ''}
            <label class="qf-modal-toggle">
              <input type="checkbox" class="qf-overlay-toggle" ${overlayChecked}/>
              <span class="qf-modal-toggle-label">แปะลายน้ำชื่อย่อบนใบลาเบล</span>
              <span class="qf-modal-toggle-hint">ปิดถ้าอยากได้ฉลากดิบ ไม่มีตัวอักษรใต้ฉลาก</span>
            </label>
            <div class="qf-modal-warn">ระบบจะส่งคำสั่งพิมพ์ทันที TikTok จะบันทึกว่าฉลากถูกพิมพ์</div>
          </div>
          <div class="qf-modal-actions">
            <button class="qf-btn-cancel">ยกเลิก</button>
            <button class="qf-btn-confirm">พิมพ์เลย</button>
          </div>
        </div>
      `;
      document.body.appendChild(overlay);
      const overlayCheck = overlay.querySelector('.qf-overlay-toggle');
      overlayCheck.addEventListener('change', () => {
        state.overlayEnabled = overlayCheck.checked;
        saveOverlayPref(state.overlayEnabled);
      });
      const cleanup = (val) => { overlay.remove(); resolve(val); };
      overlay.querySelector('.qf-btn-cancel').onclick = () => cleanup(false);
      overlay.querySelector('.qf-btn-confirm').onclick = () => cleanup(true);
      overlay.onclick = (e) => { if (e.target === overlay) cleanup(false); };
      const onKey = (e) => { if (e.key === 'Escape') { cleanup(false); document.removeEventListener('keydown', onKey); } };
      document.addEventListener('keydown', onKey);
    });
  }

  async function buildChunkPdf(ids, onProgress) {
    const { PDFDocument } = window.PDFLib;
    const mergedDoc = await PDFDocument.create();
    let pageCount = 0;
    const total = ids.length;

    for (let i = 0; i < ids.length; i += PRINT_BATCH_SIZE) {
      const batch = ids.slice(i, i + PRINT_BATCH_SIZE);
      const baseDone = i;
      const batchSize = batch.length;
      const update = (sub, label) => {
        const pct = (baseDone + batchSize * sub) / total;
        onProgress(pct, label);
      };

      update(0.05, 'ส่งคำขอไป TikTok...');
      const body = {
        fulfill_unit_id_list: batch,
        content_type_list: [1, 2],
        template_type: 0,
        op_scene: 2,
        file_prefix: 'Shipping label',
        request_time: Date.now(),
        print_option: {tmpl: 0, template_size: 0, layout: [0]},
        print_source: 201,
      };
      const resp = await _origFetch.call(window, '/api/v1/fulfillment/shipping_doc/generate', {
        method: 'POST',
        headers: {'content-type': 'application/json'},
        body: JSON.stringify(body),
      });
      const data = await safeJson(resp, 'TikTok print');
      if (data.code !== 0) throw new Error(`print API code=${data.code} msg="${data.message || 'empty'}"`);
      const docUrl = data.data?.doc_url;
      if (!docUrl) {
        console.warn('[QF] generate succeeded but no doc_url:', data);
        continue;
      }

      update(0.20, 'ดาวน์โหลด PDF...');
      const pdfBytes = await fetch(docUrl).then(r => r.arrayBuffer());

      let modifiedBytes;
      if (state.overlayEnabled) {
        update(0.30, 'แปะ alias...');
        try {
          modifiedBytes = await overlayAliasOnPdf(pdfBytes, batch, (cur, totPages) => {
            update(0.30 + 0.55 * (cur / totPages), `แปะ alias ${cur}/${totPages} หน้า`);
          });
        } catch (e) {
          console.warn('[QF] overlay failed, using original:', e);
          modifiedBytes = pdfBytes;
        }
      } else {
        update(0.85, 'ข้ามการแปะ alias (ปิดอยู่)');
        modifiedBytes = pdfBytes;
      }

      update(0.90, 'รวมเข้า PDF...');
      const partDoc = await PDFDocument.load(modifiedBytes);
      const indices = partDoc.getPageIndices();
      const copied = await mergedDoc.copyPages(partDoc, indices);
      copied.forEach(p => mergedDoc.addPage(p));
      pageCount += copied.length;

      update(1.0, '');
      if (i + PRINT_BATCH_SIZE < ids.length) await sleep(200);
    }

    const bytes = await mergedDoc.save();
    return { bytes, pageCount };
  }

  async function printIds(ids, displayLabel, sampleText, filenameHint) {
    if (isShopee()) {
      // Copy order_ids (extracted from records) to clipboard so user can paste into Shopee's search
      const orderIds = [];
      for (const id of ids) {
        const rec = state.records.get(id);
        const oid = rec?.orderIds?.[0];
        if (oid) orderIds.push(oid);
      }
      const unique = [...new Set(orderIds)];
      const text = unique.join('\n');
      try {
        await navigator.clipboard.writeText(text);
        showToast(`คัดลอก ${unique.length} order IDs แล้ว — paste ลงช่อง "ค้นหาคำสั่งซื้อ" ของ Shopee`, 5000);
      } catch {
        showToast(`Shopee: รวบรวม ${unique.length} IDs (clipboard ไม่ทำงาน)`, 4000);
      }
      return false;
    }
    if (!ids.length) { showToast('ไม่พบ ID สำหรับพิมพ์ — ลองสแกนใหม่', 3000); return false; }

    const confirmed = await showPrintConfirm({
      title: displayLabel || 'พิมพ์ฉลาก',
      count: ids.length,
      sampleText,
    });
    if (!confirmed) return false;

    let chunkCount = 1;
    if (ids.length > 200) {
      const choice = await showChunkChoiceModal({total: ids.length});
      if (choice === null) return false;
      chunkCount = choice;
    }

    const chunkSize = Math.ceil(ids.length / chunkCount);
    const baseFilename = makeBaseFilename(filenameHint || displayLabel);
    const chunks = [];
    for (let i = 0; i < ids.length; i += chunkSize) {
      const slice = ids.slice(i, i + chunkSize);
      const idx = chunks.length + 1;
      chunks.push({
        ids: slice,
        filename: chunkCount === 1
          ? `${baseFilename}.pdf`
          : `${baseFilename}-ชุด${idx}-${chunkCount}.pdf`,
        label: chunkCount === 1 ? 'ไฟล์เดียว' : `ชุด ${idx}/${chunkCount}`,
      });
    }
    return runChunkedExport(chunks, displayLabel || 'พิมพ์ฉลาก');
  }

  async function runChunkedExport(chunks, displayTitle) {
    const totalIds = chunks.reduce((a, c) => a + c.ids.length, 0);
    const result = showChunkedResult({
      title: displayTitle,
      totalIds,
      chunks: chunks.map(c => ({count: c.ids.length, label: c.label, filename: c.filename})),
    });

    const runChunk = async (ci) => {
      result.startChunk(ci);
      try {
        const { bytes, pageCount } = await buildChunkPdf(chunks[ci].ids, (pct, label) => {
          result.updateChunkProgress(ci, pct, label);
        });
        const blob = new Blob([bytes], {type: 'application/pdf'});
        const url = URL.createObjectURL(blob);
        result.completeChunk(ci, {url, pageCount});
        return true;
      } catch (e) {
        console.error('[QF] chunk', ci+1, 'failed:', e);
        result.errorChunk(ci, e.message);
        return false;
      }
    };

    result.setRetryHandler(runChunk);

    for (let ci = 0; ci < chunks.length; ci++) {
      await runChunk(ci); // continues even on failure
    }
    result.allDone();
    return true;
  }

  async function printProductLabels(productId, skuId, scenario) {
    const ids = collectFulfillIds(productId, skuId, scenario);
    const product = state.products.get(productId);
    const variant = skuId ? product?.variants.get(skuId) : null;
    const alias = (state.aliases.get(productId) || '').trim();
    const baseName = alias || shortName(product?.productName);
    const variantSuffix = variant ? ` · ${variant.skuName || variant.sellerSkuName || variant.skuId}` : '';
    const scenarioLabel = scenario === 'multi' ? '(qty > 1)' : '(qty = 1)';
    const title = `${product?.productName || productId}${variantSuffix} ${scenarioLabel}`;
    const sampleQty = scenario === 'multi' ? '2+' : '1';
    const filenameHint = variant
      ? `${baseName} ${variant.skuName || variant.sellerSkuName || ''}`.trim()
      : baseName;
    return printIds(ids, title, `${baseName} ${sampleQty}`, filenameHint);
  }

  async function printWeirdCombo(sigKey) {
    const combo = state.weirdCombos.get(sigKey);
    if (!combo) { showToast('ไม่พบ combo', 2000); return; }
    const ids = applyCarrierFilter([...combo.fulfillUnitIds]);
    const sample = combo.items
      .map(i => `${(state.aliases.get(i.productId) || '').trim() || shortName(i.productName)} ${i.quantity}`)
      .join(' + ');
    try {
      const filenameHint = combo.items
        .map(i => (state.aliases.get(i.productId) || '').trim() || shortName(i.productName))
        .join('+');
      const ok = await printIds(ids, `ออเดอร์แปลก: ${sample}`, sample, filenameHint);
      if (ok) { markComboDone(sigKey); renderAll(); }
      return ok;
    } catch (e) {
      console.error('[QF] printWeirdCombo failed:', e);
      showToast('พิมพ์ผิดพลาด: ' + e.message, 4000);
      return false;
    }
  }

  function selectLabelRows(productId, skuId) {
    const rows = [...document.querySelectorAll('tbody tr')];
    let count = 0;
    for (const row of rows) {
      const rec = getRecordFromRow(row);
      if (!rec?.skuList) continue;
      const matches = skuId
        ? rec.skuList.some(s => s.skuId === skuId)
        : rec.skuList.some(s => s.productId === productId);
      if (!matches) continue;
      const checkbox = row.querySelector('label.p-checkbox, .p-checkbox');
      if (checkbox) { simulateClick(checkbox); count++; }
    }
    return count;
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

  async function selectAllOrders(filterProductId, filterSkuId = null, minQty = 1, maxWait = 8000) {
    await waitForStable(maxWait);

    // Always use fiber-verified individual selection to avoid selecting wrong orders.
    // TikTok's product search does not reliably filter by productId, so we verify
    // every order via React fiber before checking its checkbox.
    const totalPages = getTotalPages();
    let count = 0;
    for (let p = 1; p <= totalPages; p++) {
      if (p > 1) {
        const ok = await goToPage(p);
        if (!ok) break;
        await waitForStable(4000);
      }
      for (const rec of getOrderRecords()) {
        const targetSku = filterSkuId
          ? rec.skuList?.find(s => s.skuId === filterSkuId)
          : rec.skuList?.find(s => s.productId === filterProductId);
        if (!targetSku || targetSku.quantity < minQty) continue;
        const tr = findTrForOrder(rec.mainOrderId);
        const checkbox = tr?.querySelector('label.p-checkbox');
        if (checkbox) { simulateClick(checkbox); count++; await sleep(50); }
      }
    }
    return count;
  }

  // ==================== HIGH-LEVEL ACTIONS ====================
  async function applyProductFilter(productId, skuId, type) {
    try {
      if (isLabelsPage()) {
        try {
          const scenario = type === 'single_sku' ? 'multi' : 'single';
          const ok = await printProductLabels(productId, skuId, scenario);
          if (ok) { markDone(productId, skuId, type); renderAll(); }
        } catch(e) {
          showToast('พิมพ์ผิดพลาด: ' + e.message, 3000);
        }
        return;
      }
      const label = skuId ? 'กำลังกรอง variant...' : 'กำลังกรอง...';
      showToast(label, 10000);
      await setSearchBox(productId);
      await clickToShipTab();
      await applyFilterCountType(type);
      if (state.autoSelectAll) {
        showToast('กำลังเลือกทั้งหมด...', 5000);
        await sleep(500);
        const minQty = type === 'single_sku' ? 2 : 1;
        const total = await selectAllOrders(productId, skuId, minQty);
        showToast(`✓ กรอง + เลือก ${total ?? ''} ออเดอร์`, 2500);
      } else {
        showToast('✓ กรองเรียบร้อย', 2000);
      }
      markDone(productId, skuId, type);
      renderAll();
    } catch (e) {
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
        const total = await selectAllOrders(null);
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
    const labels = isLabelsPage();
    if (isShopee()) document.documentElement.dataset.qfTheme = 'shopee';
    const w = document.createElement('div');
    w.id = 'qf-widget';
    w.innerHTML = `
      <div id="qf-header">
        <span>⚡ Quick Filter${isShopee() ? ' · Shopee' : (labels ? ' · Labels' : '')}</span>
        <div id="qf-header-actions">
          ${!labels ? '<button id="qf-reset-btn" title="รีเซ็ตฟิลเตอร์">↺</button>' : ''}
          <button id="qf-toggle-btn" title="ย่อ/ขยาย">−</button>
        </div>
      </div>
      <div id="qf-body">
        <div id="qf-scan-row">
          <button id="qf-scan-btn">🔍 Scan</button>
          <span id="qf-scan-status">กดปุ่มเพื่อเริ่มสแกน</span>
        </div>
        ${!labels ? `
        <div id="qf-options-row">
          <label>
            <input type="checkbox" id="qf-auto-select" checked />
            เลือกออเดอร์ทั้งหมดอัตโนมัติหลังกรอง
          </label>
        </div>` : `
        <div id="qf-options-row">
          ${isTikTok() ? `<div class="qf-filter-block">
            <div class="qf-filter-label">ประเภทออเดอร์</div>
            <div class="qf-segmented" id="qf-preorder-seg">
              <button class="qf-seg-btn ${state.preOrderFilter==='all'?'active':''}" data-val="all">ทั้งหมด</button>
              <button class="qf-seg-btn ${state.preOrderFilter==='normal'?'active':''}" data-val="normal">ปกติ</button>
              <button class="qf-seg-btn ${state.preOrderFilter==='preorder'?'active':''}" data-val="preorder">พรีออเดอร์</button>
            </div>
          </div>` : ''}
          ${isTikTok() ? `<div class="qf-filter-block">
            <div class="qf-filter-label">สถานะ</div>
            <div class="qf-segmented" id="qf-status-seg">
              <button class="qf-seg-btn ${state.labelStatusFilter==='not_printed'?'active':''}" data-val="not_printed">ยังไม่พิมพ์</button>
              <button class="qf-seg-btn ${state.labelStatusFilter==='printed'?'active':''}" data-val="printed">พิมพ์แล้ว</button>
              <button class="qf-seg-btn ${state.labelStatusFilter==='all'?'active':''}" data-val="all">ทั้งหมด</button>
            </div>
          </div>` : ''}
          <div class="qf-filter-block" id="qf-carrier-block">
            <div class="qf-filter-row-head">
              <div class="qf-filter-label">ขนส่ง <span class="qf-filter-hint">(ไม่เลือก = ทั้งหมด)</span></div>
              ${isTikTok() ? `<button id="qf-cal-icon" class="qf-cal-icon" title="กรองตามวัน">📅<span id="qf-cal-icon-dot" class="qf-cal-icon-dot" style="display:none;"></span></button>` : ''}
            </div>
            <div class="qf-carrier-chips" id="qf-carrier-chips">
              <div class="qf-carrier-empty">— สแกนเพื่อโหลดรายการขนส่ง —</div>
            </div>
          </div>
          <div class="qf-tip">คลิก card → ยืนยัน → พิมพ์ฉลาก</div>
          <button id="qf-select-toggle" class="qf-select-toggle">เลือกหลายรายการ</button>
        </div>`}
        <div id="qf-tabs">
          <div class="qf-tab active" data-tab="single">1 ชิ้น <span class="qf-tab-count" id="qf-count-single">0</span></div>
          <div class="qf-tab" data-tab="multi">หลายชิ้น <span class="qf-tab-count" id="qf-count-multi">0</span></div>
          <div class="qf-tab" data-tab="weird">ออเดอร์แปลก <span class="qf-tab-count" id="qf-count-weird">0</span></div>
        </div>
        <div id="qf-content">
          <div class="qf-empty">ยังไม่ได้สแกน</div>
        </div>
        ${labels ? `<div id="qf-select-bar" style="display:none;">
          <div class="qf-select-bar-info">
            <div class="qf-select-bar-count">เลือก 0 รายการ</div>
            <div class="qf-select-bar-hint">รวม + dedupe อัตโนมัติ</div>
          </div>
          <button class="qf-select-bar-clear">ล้าง</button>
          <button class="qf-select-bar-print">พิมพ์ที่เลือก</button>
        </div>` : ''}
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
    document.getElementById('qf-reset-btn')?.addEventListener('click', (e) => {
      e.stopPropagation();
      resetFilters();
    });
    document.getElementById('qf-scan-btn').addEventListener('click', scanAllPages);
    document.getElementById('qf-select-toggle')?.addEventListener('click', () => {
      setSelectMode(!state.selectMode);
    });
    document.getElementById('qf-cal-icon')?.addEventListener('click', () => {
      openCalendarModal();
    });
    const bar = document.getElementById('qf-select-bar');
    if (bar) {
      bar.querySelector('.qf-select-bar-clear').addEventListener('click', clearSelection);
      bar.querySelector('.qf-select-bar-print').addEventListener('click', printSelected);
    }
    document.getElementById('qf-auto-select')?.addEventListener('change', (e) => {
      state.autoSelectAll = e.target.checked;
    });
    document.querySelectorAll('#qf-status-seg .qf-seg-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const val = btn.dataset.val;
        if (state.labelStatusFilter === val) return;
        state.labelStatusFilter = val;
        document.querySelectorAll('#qf-status-seg .qf-seg-btn').forEach(b => {
          b.classList.toggle('active', b.dataset.val === val);
        });
        // Auto re-scan since status filter affects which records are processed
        if (state.products.size > 0 || state.records.size > 0) {
          showToast('กำลังสแกนใหม่ตามสถานะใหม่...', 2000);
          scanAllPages();
        }
      });
    });
    document.querySelectorAll('#qf-preorder-seg .qf-seg-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const val = btn.dataset.val;
        if (state.preOrderFilter === val) return;
        state.preOrderFilter = val;
        document.querySelectorAll('#qf-preorder-seg .qf-seg-btn').forEach(b => {
          b.classList.toggle('active', b.dataset.val === val);
        });
        // Pre-order workflow context: ฉลากมัก print ไว้ล่วงหน้าแล้วเก็บรอวันส่ง.
        // Default status="ยังไม่พิมพ์" จะตัดออเดอร์เหล่านั้นทิ้ง → user เห็นไม่ครบ.
        // Auto-relax status to "ทั้งหมด" + re-scan when switching to preorder.
        if (val === 'preorder' && state.labelStatusFilter !== 'all' && state.records.size > 0) {
          state.labelStatusFilter = 'all';
          document.querySelectorAll('#qf-status-seg .qf-seg-btn').forEach(b => {
            b.classList.toggle('active', b.dataset.val === 'all');
          });
          showToast('โหมดพรีออเดอร์: ปรับสถานะเป็น "ทั้งหมด" + กำลังสแกนใหม่', 2500);
          scanAllPages();
          return;
        }
        renderAll();
      });
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

  // ==================== ADVANCED DATE FILTER (CALENDAR) ====================
  const FIELD_LABELS = {
    createTime: 'วันที่ลูกค้าสั่ง',
    shipByTime: 'วันที่ต้องส่ง',
    autoCancelTime: 'วันที่ยกเลิกอัตโนมัติ',
  };
  const TH_MONTHS = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน',
                     'กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
  const TH_DOWS = ['อา','จ','อ','พ','พฤ','ศ','ส'];

  function dayKey(ts) {
    if (!ts) return null;
    const d = new Date(ts);
    return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
  }
  function startOfDay(ts) {
    const d = new Date(ts);
    d.setHours(0,0,0,0);
    return d.getTime();
  }
  function nextDay(ts) {
    return startOfDay(ts) + 86400000;
  }

  function getDangerZone(field, dayTs) {
    // dayTs = midnight ms of the day
    const todayStart = startOfDay(Date.now());
    const daysFromToday = Math.floor((dayTs - todayStart) / 86400000);
    if (field === 'createTime') {
      // older order = more urgent (still unshipped)
      const age = -daysFromToday;
      if (age >= 14) return 'critical';
      if (age >= 7) return 'urgent';
      if (age >= 2) return 'watch';
      return 'safe';
    }
    if (field === 'shipByTime') {
      if (daysFromToday <= 0) return 'critical'; // today or past
      if (daysFromToday === 1) return 'urgent';
      if (daysFromToday <= 3) return 'watch';
      return 'safe';
    }
    if (field === 'autoCancelTime') {
      if (daysFromToday <= 1) return 'critical'; // today/tomorrow/past
      if (daysFromToday <= 3) return 'urgent';
      if (daysFromToday <= 7) return 'watch';
      return 'safe';
    }
    return 'safe';
  }

  function summarizeZones(field) {
    const buckets = { critical: 0, urgent: 0, watch: 0, safe: 0 };
    const seenDays = new Map(); // day → zone
    for (const [id, rec] of state.records) {
      if (!passesCarrier(id) || !passesPreOrder(id)) continue;
      const t = rec[field];
      if (!t) continue;
      const k = startOfDay(t);
      let zone = seenDays.get(k);
      if (!zone) { zone = getDangerZone(field, k); seenDays.set(k, zone); }
      buckets[zone]++;
    }
    return buckets;
  }

  function buildDayCounts(field) {
    // returns Map<YYYY-MM-DD, count> counting records that match carrier+preorder filter (but NOT date filter)
    const counts = new Map();
    for (const [id, rec] of state.records) {
      if (!passesCarrier(id) || !passesPreOrder(id)) continue;
      const t = rec[field];
      if (!t) continue;
      const k = dayKey(t);
      counts.set(k, (counts.get(k) || 0) + 1);
    }
    return counts;
  }

  function getCalendarMonth() {
    if (state.calendarMonth) return state.calendarMonth;
    // Default to month of latest record
    let latest = 0;
    for (const rec of state.records.values()) {
      const t = rec[state.dateFilter.field];
      if (t && t > latest) latest = t;
    }
    const d = latest ? new Date(latest) : new Date();
    return { year: d.getFullYear(), month: d.getMonth() };
  }

  function isDateFilterActive() {
    return state.dateFilter.start !== null || state.dateFilter.end !== null;
  }

  function updateCalIconDot() {
    const dot = document.getElementById('qf-cal-icon-dot');
    if (dot) dot.style.display = isDateFilterActive() ? 'block' : 'none';
    const icon = document.getElementById('qf-cal-icon');
    if (icon) icon.classList.toggle('active', isDateFilterActive());
  }

  function openCalendarModal() {
    document.querySelectorAll('.qf-cal-modal-overlay').forEach(e => e.remove());
    const overlay = document.createElement('div');
    overlay.className = 'qf-modal-overlay qf-cal-modal-overlay';
    overlay.innerHTML = `
      <div class="qf-modal qf-cal-modal" role="dialog">
        <div class="qf-cal-modal-header">
          <div class="qf-cal-modal-title">กรองตามวัน</div>
          <button class="qf-cal-modal-close" aria-label="ปิด">×</button>
        </div>
        <div class="qf-cal-modal-body"></div>
      </div>
    `;
    document.body.appendChild(overlay);
    renderAdvanced();
    const cleanup = () => overlay.remove();
    overlay.querySelector('.qf-cal-modal-close').onclick = cleanup;
    overlay.onclick = (e) => { if (e.target === overlay) cleanup(); };
    const onKey = (e) => { if (e.key === 'Escape') { cleanup(); document.removeEventListener('keydown', onKey); } };
    document.addEventListener('keydown', onKey);
  }

  function renderAdvanced() {
    // Refresh open calendar modal if present (e.g., after carrier filter change)
    const panel = document.querySelector('.qf-cal-modal-body');
    if (!panel) { updateCalIconDot(); return; }
    const field = state.dateFilter.field;
    const counts = buildDayCounts(field);
    const cm = getCalendarMonth();
    const firstDay = new Date(cm.year, cm.month, 1);
    const daysInMonth = new Date(cm.year, cm.month + 1, 0).getDate();
    const startDow = firstDay.getDay();
    const cells = [];
    for (let i = 0; i < startDow; i++) cells.push('');
    for (let d = 1; d <= daysInMonth; d++) cells.push(d);

    const { start, end } = state.dateFilter;
    const isInRange = (day) => {
      const t = new Date(cm.year, cm.month, day).getTime();
      if (start === null && end === null) return false;
      if (start !== null && end !== null) return t >= start && t < end;
      if (start !== null) return t === start;
      return false;
    };
    const isStart = (day) => start !== null && new Date(cm.year, cm.month, day).getTime() === start;
    const isEnd = (day) => end !== null && new Date(cm.year, cm.month, day).getTime() === (end - 86400000);

    const headerSummary = (start === null && end === null)
      ? '<span class="qf-cal-empty">เลือกวันใน calendar</span>'
      : (start === end - 86400000 || end === null)
        ? `เลือก: <b>${new Date(start).toLocaleDateString('th-TH', {day:'numeric', month:'short'})}</b>`
        : `ช่วง: <b>${new Date(start).toLocaleDateString('th-TH', {day:'numeric', month:'short'})}</b> – <b>${new Date(end - 86400000).toLocaleDateString('th-TH', {day:'numeric', month:'short'})}</b>`;

    const matchedCount = (start === null && end === null) ? 0 : (() => {
      let n = 0;
      for (const [id, rec] of state.records) {
        if (!passesCarrier(id) || !passesPreOrder(id)) continue;
        if (passesDate(id)) n++;
      }
      return n;
    })();

    const zones = summarizeZones(field);
    const zoneLabels = {
      createTime: { critical: 'ค้างนานมาก', urgent: 'ค้างนาน', watch: 'รอจัดส่ง', safe: 'เพิ่งสั่ง' },
      shipByTime: { critical: 'ต้องส่งวันนี้/เลย', urgent: 'พรุ่งนี้', watch: 'ใน 3 วัน', safe: 'ปลอดภัย' },
      autoCancelTime: { critical: 'ใกล้ยกเลิก', urgent: 'รีบทำ', watch: 'ระวัง', safe: 'ปลอดภัย' },
    }[field];
    const summaryParts = ['critical','urgent','watch','safe']
      .filter(z => zones[z] > 0)
      .map(z => `<span class="qf-zone-pill qf-zone-${z}-pill"><span class="qf-zone-dot qf-zone-${z}-dot"></span>${zones[z]} ${zoneLabels[z]}</span>`);
    panel.innerHTML = `
      <div class="qf-cal-field-row">
        <select class="qf-cal-field">
          <option value="createTime" ${field==='createTime'?'selected':''}>${FIELD_LABELS.createTime}</option>
          <option value="shipByTime" ${field==='shipByTime'?'selected':''}>${FIELD_LABELS.shipByTime}</option>
          <option value="autoCancelTime" ${field==='autoCancelTime'?'selected':''}>${FIELD_LABELS.autoCancelTime}</option>
        </select>
      </div>
      ${summaryParts.length ? `<div class="qf-zone-summary">${summaryParts.join('')}</div>` : ''}
      <div class="qf-cal-header">
        <button class="qf-cal-nav qf-cal-prev" aria-label="เดือนก่อน">‹</button>
        <div class="qf-cal-title">${TH_MONTHS[cm.month]} ${cm.year + 543}</div>
        <button class="qf-cal-nav qf-cal-next" aria-label="เดือนถัดไป">›</button>
      </div>
      <div class="qf-cal-grid">
        ${TH_DOWS.map(d => `<div class="qf-cal-dow">${d}</div>`).join('')}
        ${cells.map(c => {
          if (c === '') return `<div class="qf-cal-cell qf-cal-empty"></div>`;
          const k = `${cm.year}-${String(cm.month+1).padStart(2,'0')}-${String(c).padStart(2,'0')}`;
          const cnt = counts.get(k) || 0;
          const dayTs = new Date(cm.year, cm.month, c).getTime();
          const zone = cnt > 0 ? getDangerZone(field, dayTs) : null;
          const cls = ['qf-cal-cell'];
          if (cnt > 0) cls.push('qf-cal-has', `qf-zone-${zone}`);
          if (isInRange(c)) cls.push('qf-cal-in-range');
          if (isStart(c)) cls.push('qf-cal-start');
          if (isEnd(c)) cls.push('qf-cal-end');
          return `<div class="${cls.join(' ')}" data-day="${c}">
            <div class="qf-cal-num">${c}</div>
            ${cnt > 0 ? `<div class="qf-cal-count">${cnt}</div>` : ''}
          </div>`;
        }).join('')}
      </div>
      <div class="qf-cal-footer">
        <div class="qf-cal-summary">${headerSummary}${matchedCount > 0 ? ` · <b>${matchedCount}</b> ออเดอร์` : ''}</div>
        ${(start !== null || end !== null) ? `<button class="qf-cal-clear">ล้าง</button>` : ''}
      </div>
      <div class="qf-cal-hint">คลิก = วันเดียว · Shift+คลิก = ช่วง</div>
      <div class="qf-zone-legend">
        <span><span class="qf-zone-dot qf-zone-critical-dot"></span>ใกล้/เลย</span>
        <span><span class="qf-zone-dot qf-zone-urgent-dot"></span>รีบ</span>
        <span><span class="qf-zone-dot qf-zone-watch-dot"></span>ระวัง</span>
        <span><span class="qf-zone-dot qf-zone-safe-dot"></span>ปลอดภัย</span>
      </div>
    `;

    panel.querySelector('.qf-cal-field').addEventListener('change', (e) => {
      state.dateFilter.field = e.target.value;
      state.dateFilter.start = null;
      state.dateFilter.end = null;
      renderAll();
    });
    panel.querySelector('.qf-cal-prev').addEventListener('click', () => {
      const nm = cm.month - 1;
      state.calendarMonth = nm < 0 ? {year: cm.year - 1, month: 11} : {year: cm.year, month: nm};
      renderAdvanced();
    });
    panel.querySelector('.qf-cal-next').addEventListener('click', () => {
      const nm = cm.month + 1;
      state.calendarMonth = nm > 11 ? {year: cm.year + 1, month: 0} : {year: cm.year, month: nm};
      renderAdvanced();
    });
    panel.querySelectorAll('.qf-cal-cell[data-day]').forEach(cell => {
      cell.addEventListener('click', (e) => {
        const day = parseInt(cell.dataset.day);
        const ts = new Date(cm.year, cm.month, day).getTime();
        if (e.shiftKey && state.dateFilter.start !== null) {
          // Range select: extend
          const start = Math.min(state.dateFilter.start, ts);
          const end = Math.max(state.dateFilter.start + 86400000, ts + 86400000);
          state.dateFilter.start = start;
          state.dateFilter.end = end;
        } else {
          // Single day
          state.dateFilter.start = ts;
          state.dateFilter.end = ts + 86400000;
        }
        renderAll();
      });
    });
    panel.querySelector('.qf-cal-clear')?.addEventListener('click', () => {
      state.dateFilter.start = null;
      state.dateFilter.end = null;
      renderAll();
    });
  }

  function renderCarriers() {
    const wrap = document.getElementById('qf-carrier-chips');
    if (!wrap) return;
    if (state.carriers.size === 0) {
      wrap.innerHTML = '<div class="qf-carrier-empty">— สแกนเพื่อโหลดรายการขนส่ง —</div>';
      return;
    }
    // Count per carrier respecting preOrderFilter so chip numbers match what's visible.
    // Deliberately NOT applying carrierFilter here — otherwise unselected chips collapse to 0
    // and user can't see how many each carrier would contribute.
    const counts = new Map();
    for (const [fulfillUnitId, cid] of state.carrierOf) {
      if (!passesPreOrder(fulfillUnitId)) continue;
      counts.set(cid, (counts.get(cid) || 0) + 1);
    }
    wrap.innerHTML = '';
    const carriers = [...state.carriers.values()]
      .sort((a, b) => (counts.get(b.id) || 0) - (counts.get(a.id) || 0));
    for (const c of carriers) {
      const n = counts.get(c.id) || 0;
      const chip = document.createElement('button');
      const active = state.carrierFilter.has(c.id);
      chip.className = 'qf-carrier-chip' + (active ? ' active' : '');
      chip.title = c.name;
      chip.innerHTML = `
        ${c.iconUrl ? `<img src="${c.iconUrl}" referrerpolicy="no-referrer"/>` : '<span class="qf-carrier-noicon">📦</span>'}
        <span class="qf-carrier-name">${escapeHtml(c.name)}</span>
        <span class="qf-carrier-count">${n}</span>
      `;
      chip.addEventListener('click', () => {
        if (state.carrierFilter.has(c.id)) state.carrierFilter.delete(c.id);
        else state.carrierFilter.add(c.id);
        renderAll();
      });
      wrap.appendChild(chip);
    }
  }

  function renderAll() {
    const labelsPg = isLabelsPage();
    const singleCount = labelsPg
      ? [...state.products.values()].filter(p => carrierFilteredSize(p.fulfillUnitIdsSingle) > 0).length
      : [...state.products.values()].filter(p => p.orderCountSingle > 0).length;
    const multiCount = labelsPg
      ? [...state.products.values()].filter(p => carrierFilteredSize(p.fulfillUnitIdsMulti) > 0).length
      : [...state.products.values()].filter(p => p.orderCountMulti > 0).length;
    const weirdCount = labelsPg
      ? [...state.weirdCombos.values()].filter(c => carrierFilteredSize(c.fulfillUnitIds) > 0).length
      : state.weirdOrders.length;
    document.getElementById('qf-count-single').textContent = singleCount;
    document.getElementById('qf-count-multi').textContent  = multiCount;
    document.getElementById('qf-count-weird').textContent  = weirdCount;
    renderCarriers();
    renderAdvanced();
    renderContent();
    const tog = document.getElementById('qf-select-toggle');
    if (tog) {
      tog.classList.toggle('active', state.selectMode);
      tog.textContent = state.selectMode ? 'ออกจากโหมดเลือก' : 'เลือกหลายรายการ';
    }
    updateSelectionBar();
  }

  function renderContent() {
    const wrap = document.getElementById('qf-content');
    wrap.innerHTML = '';

    if (state.products.size === 0 && state.weirdOrders.length === 0) {
      wrap.innerHTML = '<div class="qf-empty">ยังไม่ได้สแกน</div>';
      return;
    }
    if (state.currentTab === 'single' || state.currentTab === 'multi') {
      const idsKey = state.currentTab === 'single' ? 'fulfillUnitIdsSingle' : 'fulfillUnitIdsMulti';
      const type = state.currentTab === 'single' ? 'single_item' : 'single_sku';
      const labelsPg = isLabelsPage();
      const productsRaw = labelsPg
        ? [...state.products.values()].map(p => ({...p, _count: carrierFilteredSize(p[idsKey])}))
        : [...state.products.values()].map(p => ({...p, _count: state.currentTab === 'single' ? p.orderCountSingle : p.orderCountMulti}));
      const products = productsRaw
        .filter(p => p._count > 0)
        .sort((a, b) => b._count - a._count);
      if (!products.length) {
        wrap.innerHTML = '<div class="qf-empty">ไม่พบสินค้าในประเภทนี้</div>';
        return;
      }
      const grid = document.createElement('div');
      grid.className = 'qf-product-grid';
      const labels = isLabelsPage();
      const variantCount = (v) => labels
        ? carrierFilteredSize(v[idsKey])
        : (state.currentTab === 'single' ? v.orderCountSingle : v.orderCountMulti);
      const scenario = state.currentTab === 'single' ? 'single' : 'multi';
      for (const p of products) {
        const card = document.createElement('div');
        const cardDone = isDone(p.productId, null, type);
        const productItem = {type: 'product', productId: p.productId, scenario};
        const selected = state.selectMode && isSelected(productItem);
        card.className = 'qf-product-card'
          + (cardDone ? ' qf-done' : '')
          + (state.selectMode ? ' qf-select-mode' : '')
          + (selected ? ' qf-selected' : '');
        card.title = p.productName;
        const variantsRaw = [...p.variants.values()].map(v => ({v, c: variantCount(v)}));
        const variants = variantsRaw.filter(x => x.c > 0);
        const hasBadges = variants.length >= 1;
        const aliasVal = labels ? (state.aliases.get(p.productId) || '') : '';
        const showVariantToggle = labels && variants.length >= 1;
        card.innerHTML = `
          <img src="${p.productImageURL}" alt="" referrerpolicy="no-referrer"/>
          <div class="qf-product-name">${escapeHtml(p.productName)}</div>
          <div class="qf-product-count">${p._count} ออเดอร์</div>
          ${labels ? `
            <input class="qf-alias-input" type="text" placeholder="ชื่อย่อ (เช่น แดง1)" value="${escapeHtml(aliasVal)}" maxlength="20"/>
            ${variants.length > 0 ? `<button class="qf-variant-link">ปรับตัวเลือก</button>` : ''}
          ` : ''}
          ${hasBadges ? `<div class="qf-variant-badges"></div>` : ''}
        `;
        const aliasInput = card.querySelector('.qf-alias-input');
        if (aliasInput) {
          aliasInput.addEventListener('click', e => e.stopPropagation());
          aliasInput.addEventListener('change', () => {
            const v = aliasInput.value.trim();
            if (v) state.aliases.set(p.productId, v);
            else state.aliases.delete(p.productId);
            saveAliases();
          });
          aliasInput.addEventListener('keydown', e => {
            if (e.key === 'Enter') { aliasInput.blur(); }
          });
        }
        const variantLink = card.querySelector('.qf-variant-link');
        if (variantLink) {
          variantLink.addEventListener('click', e => {
            e.stopPropagation();
            openAliasModal(p.productId);
          });
        }
        if (hasBadges) {
          const badgesEl = card.querySelector('.qf-variant-badges');
          for (const {v, c} of variants) {
            const badgeDone = isDone(p.productId, v.skuId, type);
            const vInfo = labels ? getVariantInfo(p.productId, v.skuId) : null;
            const aliasOverride = (vInfo?.alias || '').trim();
            const originalName = v.skuName || v.sellerSkuName || v.skuId;
            const displayName = aliasOverride || originalName;
            const badge = document.createElement('span');
            badge.className = 'qf-variant-badge'
              + (badgeDone ? ' qf-badge-done' : '')
              + (aliasOverride ? ' qf-badge-aliased' : '');
            badge.dataset.skuId = v.skuId;
            if (aliasOverride) badge.title = `${originalName} → ${aliasOverride}`;
            badge.textContent = `${displayName} (${c})`;
            const variantItem = {type: 'variant', productId: p.productId, skuId: v.skuId, scenario};
            if (state.selectMode && isSelected(variantItem)) badge.classList.add('qf-selected');
            badge.addEventListener('click', (e) => {
              e.stopPropagation();
              if (state.selectMode) {
                if (badgeDone) return;
                toggleSelection(variantItem);
                badge.classList.toggle('qf-selected');
                return;
              }
              if (isDone(p.productId, v.skuId, type)) {
                state.doneItems.delete(doneKey(p.productId, v.skuId, type));
                saveDoneItems();
                renderAll();
              } else {
                applyProductFilter(p.productId, v.skuId, type);
              }
            });
            badgesEl.appendChild(badge);
          }
        }
        card.addEventListener('click', () => {
          if (state.selectMode) {
            if (cardDone) return; // can't select already-printed
            toggleSelection(productItem);
            card.classList.toggle('qf-selected');
            return;
          }
          if (cardDone) {
            state.doneItems.delete(doneKey(p.productId, null, type));
            saveDoneItems();
            renderAll();
          } else {
            applyProductFilter(p.productId, null, type);
          }
        });
        grid.appendChild(card);
      }
      wrap.appendChild(grid);
    } else if (state.currentTab === 'weird') {
      const labelsPage = isLabelsPage();
      if (!labelsPage) {
        const applyBtn = document.createElement('button');
        applyBtn.id = 'qf-weird-apply-btn';
        applyBtn.textContent = `📋 แสดงออเดอร์แปลกทั้งหมด (${state.weirdOrders.length})`;
        applyBtn.addEventListener('click', applyWeirdFilter);
        wrap.appendChild(applyBtn);
      }
      if (!state.weirdOrders.length) {
        const empty = document.createElement('div');
        empty.className = 'qf-empty';
        empty.textContent = 'ไม่พบออเดอร์แปลก';
        wrap.appendChild(empty);
        return;
      }

      if (labelsPage) {
        // Group by combination signature, apply carrier filter to count
        const combos = [...state.weirdCombos.values()]
          .map(c => ({...c, _count: carrierFilteredSize(c.fulfillUnitIds)}))
          .filter(c => c._count > 0)
          .sort((a, b) => b._count - a._count);
        if (!combos.length) {
          const empty = document.createElement('div');
          empty.className = 'qf-empty';
          empty.textContent = 'ไม่พบ combo (ลองยกเลิกกรองขนส่ง)';
          wrap.appendChild(empty);
          return;
        }
        const grid = document.createElement('div');
        grid.id = 'qf-weird-combo-grid';
        for (const combo of combos) {
          const card = document.createElement('div');
          const comboDone = isComboDone(combo.sigKey);
          const comboItem = {type: 'combo', sigKey: combo.sigKey};
          const selected = state.selectMode && isSelected(comboItem);
          card.className = 'qf-combo-card'
            + (comboDone ? ' qf-done' : '')
            + (state.selectMode ? ' qf-select-mode' : '')
            + (selected ? ' qf-selected' : '');
          const itemsHtml = combo.items.map((s, idx) => `
            ${idx > 0 ? '<span class="qf-combo-plus">+</span>' : ''}
            <div class="qf-combo-item" data-pid="${s.productId}">
              <img src="${s.productImageURL}" referrerpolicy="no-referrer" title="คลิกเพื่อตั้งชื่อย่อ + ตัวเลือก"/>
              <div class="qf-combo-qty">×${s.quantity}</div>
              <input class="qf-combo-alias-input" type="text" placeholder="ชื่อย่อ" value="${escapeHtml((state.aliases.get(s.productId) || '').trim())}" maxlength="20"/>
            </div>
          `).join('');
          card.innerHTML = `
            <div class="qf-combo-row">${itemsHtml}</div>
            <div class="qf-combo-count">${combo._count} ออเดอร์</div>
          `;
          // wire alias inputs + edit buttons
          card.querySelectorAll('.qf-combo-item').forEach(itemEl => {
            const pid = itemEl.dataset.pid;
            const inp = itemEl.querySelector('.qf-combo-alias-input');
            const img = itemEl.querySelector('img');
            inp.addEventListener('click', e => e.stopPropagation());
            inp.addEventListener('change', () => {
              const v = inp.value.trim();
              if (v) state.aliases.set(pid, v); else state.aliases.delete(pid);
              saveAliases();
            });
            inp.addEventListener('keydown', e => { if (e.key === 'Enter') inp.blur(); });
            img.addEventListener('click', e => {
              e.stopPropagation();
              openAliasModal(pid);
            });
          });
          card.addEventListener('click', () => {
            if (state.selectMode) {
              if (comboDone) return;
              toggleSelection(comboItem);
              card.classList.toggle('qf-selected');
              return;
            }
            if (isComboDone(combo.sigKey)) {
              state.doneItems.delete(comboDoneKey(combo.sigKey));
              saveDoneItems();
              renderAll();
            } else {
              printWeirdCombo(combo.sigKey);
            }
          });
          grid.appendChild(card);
        }
        wrap.appendChild(grid);
      } else {
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
  }

  function escapeHtml(s) {
    return (s || '').replace(/[&<>"']/g, c => ({
      '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
    }[c]));
  }

  // ==================== INIT ====================
  function init() {
    if (!isOrderPage() && !isLabelsPage()) return;
    buildWidget();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
