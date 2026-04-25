(() => {
  'use strict';
  // VERSION MARKER — bumped when testing new features in DevTools.
  // Increment this string to confirm the browser picked up the freshest
  // content.js after reloading the extension in chrome://extensions.
  window.__qfContentScriptBuild = 'shopee-discovery-v1';

  // ==================== FETCH HOOK (runs at document_start, before TikTok saves window.fetch) ====================
  let _apiListUrl = null;
  let _apiListBodyTemplate = null;
  let _labelsApiUrl = null;
  let _labelsApiBodyTemplate = null;
  // Order detail (single order) — captured to backfill address book on-demand.
  // Only populated when TikTok itself fires /order/get (user opened a detail
  // view). Until captured, the pre-print backfill is a no-op.
  let _orderGetUrl = null;
  let _orderGetBodyTemplate = null;
  // Buyer contact info — the REAL unmask endpoint fired when user clicks the
  // 👁 eye icon on the order detail page. One POST per field
  // (contact_info_type: 0=name, 1=address, 2=phone, 3=nickname). Response
  // shape is per-field: data.plain_text_name / plain_text_phone_number /
  // plain_text_nickname / plain_text_address{items,region,districts}.
  let _buyerContactUrl = null;
  let _buyerContactBodyTemplate = null;
  // Shopee captured URLs/bodies
  let _shopeeIndexUrl = null;
  let _shopeeIndexBody = null;
  let _shopeeCardUrl = null;
  let _shopeeCardBody = null;
  // One-shot probe flag: dump first Shopee record shape to console when
  // pre-order detection fails, so the user can report the real field name.
  // Reset per scan in scanAllPages so probing survives across scans.
  let _shopeePreOrderProbeLogged = false;
  let _noteProbeLogged = false;
  let _noteTemplateHintShown = false;

  // Shopee endpoint discovery probe — captures any URL touching logistics /
  // waybill / address / detail so we can map the print pipeline. Ring buffer
  // bounded to keep memory low during long sessions. Inspect via
  // window.__qfShopeeProbes() in DevTools after clicking print on Shopee.
  const _shopeeProbeRing = [];
  const SHOPEE_PROBE_RING_MAX = 20;
  // Substring patterns that flag a URL as "interesting for Shopee print/address
  // research". Add patterns here when we discover new endpoints. Negative
  // patterns excluded later (chat, image upload).
  const SHOPEE_PROBE_INTEREST = [
    'logistics',           // /api/v3/logistics/* — waybill, can_print, channel info
    'waybill',             // explicit waybill endpoints
    'shipping',            // shipping label URLs
    'print',               // any print-related path
    'order/detail',        // single-order detail
    'order/get_one',       // get_one_order
    'buyer_address',       // buyer address endpoint
    'buyer_user',          // buyer user info
    'recipient',           // recipient details
    'logistic_info',       // logistic info per order
    'order/get_logistic',  // logistic lookup
  ];
  const SHOPEE_PROBE_EXCLUDE = ['/chat/', '/image/', '/file/upload', '/track/', '/log/event'];

  function _shopeeShouldProbe(url) {
    if (!url || typeof url !== 'string') return false;
    if (!url.includes('shopee')) {
      // Same-origin requests: only probe when host is the seller domain.
      if (location.hostname !== 'seller.shopee.co.th') return false;
    }
    if (SHOPEE_PROBE_EXCLUDE.some(p => url.includes(p))) return false;
    return SHOPEE_PROBE_INTEREST.some(p => url.includes(p));
  }

  function _shopeePushProbe(entry) {
    _shopeeProbeRing.push({ ts: Date.now(), ...entry });
    if (_shopeeProbeRing.length > SHOPEE_PROBE_RING_MAX) _shopeeProbeRing.shift();
  }

  // Expose inspector helpers — copy/paste these in DevTools console after
  // performing the action you want to map (e.g. clicking Shopee's print button).
  //   __qfShopeeProbes()              — table of captured URLs + status
  //   __qfShopeeProbe(idx)            — full request+response of one entry
  //   __qfShopeeProbeClear()          — reset ring buffer
  window.__qfShopeeProbes = () => {
    if (!_shopeeProbeRing.length) {
      console.warn('[QF Shopee] no probes captured yet — perform an action (print, view order detail) on Shopee Seller');
      return [];
    }
    return _shopeeProbeRing.map((p, i) => ({
      idx: i,
      method: p.method,
      url: p.url.length > 80 ? p.url.slice(0, 80) + '…' : p.url,
      status: p.status,
      hasReqBody: !!p.requestBody,
      hasRespJson: !!p.responseJson,
      tsAgo: `${Math.round((Date.now() - p.ts) / 1000)}s ago`,
    }));
  };
  window.__qfShopeeProbe = (idx) => _shopeeProbeRing[idx] || null;
  window.__qfShopeeProbeClear = () => { _shopeeProbeRing.length = 0; console.log('[QF Shopee] probe ring cleared'); };
  // Passive API sniffer — captures the live endpoint URL and request-body template
  // the FIRST TIME TikTok/Shopee fires each order-list/labels API call, then stores
  // them for replay (bulk scan across pages). This is why content.js runs in MAIN
  // world at document_start: ISOLATED world cannot intercept window.fetch before
  // TikTok's own wrapper wraps it again, and __reactFiber DOM properties are not
  // accessible cross-world. The hook NEVER modifies requests or responses, NEVER
  // exfiltrates data, and passes every call straight through to the original fetch.
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

    // Order detail (single order) — used for pre-print address backfill.
    // URL shape seen in the wild: /api/fulfillment/order/get (TikTok may
    // rotate the prefix, so we do substring match on "fulfillment/order/get").
    const isOrderGet = /fulfillment\/order\/get\b|\/order\/get_detail\b/.test(url);
    if (isOrderGet && !_orderGetUrl) {
      _orderGetUrl = url;
      captureBody(b => { if (!_orderGetBodyTemplate) _orderGetBodyTemplate = b; });
    }

    // Buyer contact info (eye-icon decrypt) — capture URL + template body.
    // The body schema we observed: { main_order_id, contact_info_type }.
    // We also keep a resolvable request-body promise to pair with the response
    // below, because the response only contains the decrypted field value and
    // NOT the main_order_id — so we must read it from the request.
    const isBuyerContact = /\/fulfillment\/orders\/buyer_contact_info\/get\b/.test(url);
    let _buyerContactBodyPromise = null;
    if (isBuyerContact) {
      if (!_buyerContactUrl) {
        _buyerContactUrl = url;
        captureBody(b => { if (!_buyerContactBodyTemplate) _buyerContactBodyTemplate = b; });
      }
      if (rawBodyStr) {
        _buyerContactBodyPromise = Promise.resolve().then(() => {
          try { return JSON.parse(rawBodyStr); } catch (e) { return null; }
        });
      } else if (bodyPromise) {
        _buyerContactBodyPromise = bodyPromise.then(t => {
          try { return JSON.parse(t); } catch (e) { return null; }
        });
      }
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

    // Shopee print/waybill/address discovery probe — captures full req+resp for
    // any logistics/waybill/detail endpoint so we can map the print pipeline
    // without rebuilding extension between every test.
    const isShopeeProbe = _shopeeShouldProbe(url);
    let _shopeeProbeReqBody = null;
    if (isShopeeProbe) {
      if (rawBodyStr) {
        try { _shopeeProbeReqBody = JSON.parse(rawBodyStr); } catch { _shopeeProbeReqBody = rawBodyStr; }
      } else if (bodyPromise) {
        try { _shopeeProbeReqBody = await bodyPromise.then(t => { try { return JSON.parse(t); } catch { return t; } }); } catch {}
      }
    }

    const resp = await _origFetch.apply(this, args);

    if (isShopeeProbe && resp) {
      try {
        resp.clone().json()
          .then(j => {
            _shopeePushProbe({
              method: (args[1]?.method) || (args[0] instanceof Request ? args[0].method : 'GET'),
              url, status: resp.status,
              requestBody: _shopeeProbeReqBody,
              responseJson: j,
            });
            console.log(`[QF Shopee probe] ${resp.status} ${url.slice(0, 100)} — keys:`, Object.keys(j?.data || j || {}));
          })
          .catch(() => {
            _shopeePushProbe({ method: 'POST', url, status: resp.status, requestBody: _shopeeProbeReqBody, responseJson: null });
          });
      } catch {}
    }

    // Passive response sniffer for /order/get — clone and parse off-thread
    // so we never block the response path or mutate anything TikTok consumes.
    // We ONLY save address data when response is not masked (button_status!=3).
    if (isOrderGet && resp && resp.ok) {
      try {
        resp.clone().json()
          .then(j => { try { captureAddressFromOrderGet(j); } catch (e) {} })
          .catch(() => {});
      } catch (e) {}
    }

    // Passive response sniffer for buyer_contact_info/get — when the user
    // clicks an 👁 eye icon, TikTok fires this endpoint and the response body
    // contains ONE plain-text field. We pair it with the captured request body
    // (which carries main_order_id + contact_info_type) and merge the field
    // into the address book record keyed by orderId.
    if (isBuyerContact && resp && resp.ok && _buyerContactBodyPromise) {
      try {
        Promise.all([_buyerContactBodyPromise, resp.clone().json()])
          .then(([reqBody, respJson]) => {
            try { captureBuyerContactFieldResponse(reqBody, respJson); } catch (e) {}
          })
          .catch(() => {});
      } catch (e) {}
    }

    return resp;
  };

  // XMLHttpRequest hook — same purpose as the fetch hook above. Shopee routes all
  // order-list API calls through XHR rather than fetch, so both must be covered.
  // Stores URL + body on open/send, never modifies or exfiltrates requests.
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

    // Shopee discovery probe — same coverage as fetch hook. We listen for
    // load to read responseText, then push to ring buffer for inspection.
    if (_shopeeShouldProbe(url)) {
      let reqBody = body;
      if (typeof body === 'string') {
        try { reqBody = JSON.parse(body); } catch {}
      }
      const onLoad = () => {
        let respJson = null;
        try { respJson = JSON.parse(this.responseText); } catch { respJson = this.responseText; }
        _shopeePushProbe({
          method: this._qfMethod || 'POST', url, status: this.status,
          requestBody: reqBody, responseJson: respJson,
        });
        try {
          const top = respJson && typeof respJson === 'object' ? Object.keys(respJson?.data || respJson || {}) : [];
          console.log(`[QF Shopee probe] ${this.status} ${url.slice(0, 100)} — keys:`, top);
        } catch {}
        this.removeEventListener('load', onLoad);
      };
      this.addEventListener('load', onLoad);
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
  // TikTok flags orders as "การพิมพ์ไม่สำเร็จ" (print failed) after a prior
  // generate-doc call errored server-side. Treated as its own segment so it
  // doesn't overlap with 'not_printed' (30) or 'printed' (50).
  const LABEL_STATUS_PRINT_FAILED = 40;

  // ==================== STATE ====================
  const ALIAS_STORAGE_KEY = 'qf_product_aliases_v1';
  const VARIANT_ALIAS_STORAGE_KEY = 'qf_variant_aliases_v1';
  const WORKERS_STORAGE_KEY = 'qf_workers_v1';
  const OVERLAY_PREF_KEY = 'qf_overlay_enabled_v1';
  const ADDRESS_CAPTURE_PREF_KEY = 'qf_address_capture_v1';
  function loadOverlayPref() {
    const v = localStorage.getItem(OVERLAY_PREF_KEY);
    return v === null ? true : v === 'true';
  }
  function saveOverlayPref(enabled) {
    localStorage.setItem(OVERLAY_PREF_KEY, String(!!enabled));
  }

  // ==================== DIVIDER PRESET (Phase 1 of Custom Layout feature) ====================
  //
  // ฟีเจอร์ Custom PDF Layout — วางแผน 3 เฟส:
  //
  // Phase 1 (IMPLEMENTED HERE): Preset picker.
  //   - 4 presets: minimal / standard / detailed / photo-first.
  //   - UI: radio group ใน settings menu → บันทึกที่ localStorage['qf_divider_preset_v1'].
  //   - Applied โดย buildDividerPage ผ่าน getDividerPresetConfig().
  //
  // Phase 2 (SPEC ONLY — design doc below):
  //   Data model: localStorage['qf_divider_config_v1'] = JSON {
  //     fields: { alias, name, variant, carrier, image, qty, worker, footer } — each { visible: bool, size: 's'|'m'|'l' }
  //     layout: 'stacked' | 'compact'
  //   }
  //   UI mockup: modal with 2 columns — left column lists fields with checkbox + size radio;
  //              right column shows live preview (HTML approximation of PDF page at 2x scale).
  //   Integration: buildDividerPage reads config; each field block wraps in `if (cfg.fields.X.visible)` and
  //                uses SIZE_MAP[cfg.fields.X.size] to pick font size. Positions recomputed top-down so
  //                hidden fields don't leave gaps (use vertical cursor).
  //   Fallback: if config missing / malformed → use preset from Phase 1.
  //
  // Phase 3 (SPEC ONLY):
  //   Visual editor — full drag-and-drop canvas.
  //   Data model: localStorage['qf_divider_layout_v1'] = JSON {
  //     canvas: { w, h },
  //     elements: [{ id, type: 'alias'|'name'|'image'|'carrier'|'qty'|'text', x, y, w, h, size?, text? }]
  //   }
  //   UI: HTML5 canvas-backed editor, mouse-drag positions elements; snap grid 4pt; live preview.
  //   Save button serializes positions back into localStorage.
  //   Integration: if layout present, buildDividerPage ignores preset/config and renders each element at
  //                its recorded (x, y). Coords are PDF points with origin bottom-left (consistent with pdf-lib).
  //   Migration: "convert to visual layout" button in Phase 2 modal seeds the canvas from current config.
  //
  const DIVIDER_PRESET_KEY = 'qf_divider_preset_v1';
  const DIVIDER_PRESETS = ['minimal', 'standard', 'detailed', 'photo-first'];
  function loadDividerPreset() {
    const v = localStorage.getItem(DIVIDER_PRESET_KEY);
    return DIVIDER_PRESETS.includes(v) ? v : 'standard';
  }
  function saveDividerPreset(preset) {
    if (!DIVIDER_PRESETS.includes(preset)) return;
    localStorage.setItem(DIVIDER_PRESET_KEY, preset);
  }
  // Preset → field visibility + emphasis. ทุกอย่างขาวดำ.
  function getDividerPresetConfig(preset) {
    const p = DIVIDER_PRESETS.includes(preset) ? preset : 'standard';
    switch (p) {
      case 'minimal':
        // alias only — giant.
        return { preset: p, showAlias: true, aliasScale: 1.6, showName: false, showVariant: false,
          showCarrier: false, showImage: false, showQty: true, showWorker: true, showFooter: false,
          imageFirst: false };
      case 'detailed':
        return { preset: p, showAlias: true, aliasScale: 1.0, showName: true, showVariant: true,
          showCarrier: true, showImage: true, showQty: true, showWorker: true, showFooter: true,
          showMeta: true, imageFirst: false };
      case 'photo-first':
        return { preset: p, showAlias: true, aliasScale: 1.0, showName: true, showVariant: true,
          showCarrier: true, showImage: true, showQty: true, showWorker: true, showFooter: true,
          imageFirst: true };
      case 'standard':
      default:
        return { preset: p, showAlias: true, aliasScale: 1.3, showName: true, showVariant: true,
          showCarrier: true, showImage: true, showQty: true, showWorker: true, showFooter: true,
          imageFirst: false };
    }
  }

  // ==================== DIVIDER CONFIG (Phase 2 of Custom Layout) ====================
  // Fine-grained config: per-field visibility + size (s|m|l).
  // Shape: { fields: { alias, name, variant, carrier, image, qty, worker, footer }, layout: 'stacked'|'compact' }
  // Each field entry = { visible: bool, size: 's'|'m'|'l' }.
  // If absent / malformed → fall back to Phase 1 preset via presetToConfig().
  const DIVIDER_CONFIG_KEY = 'qf_divider_config_v1';
  const DIVIDER_FIELDS = ['alias', 'name', 'variant', 'carrier', 'image', 'qty', 'worker', 'footer'];
  const DIVIDER_LAYOUTS = ['stacked', 'compact'];
  const DIVIDER_SIZES = ['s', 'm', 'l'];
  const DIVIDER_SIZE_MAP = { s: 0.75, m: 1.0, l: 1.3 };

  function _validDividerField(f) {
    return f && typeof f === 'object' &&
      typeof f.visible === 'boolean' &&
      DIVIDER_SIZES.includes(f.size);
  }
  function _validDividerConfig(cfg) {
    if (!cfg || typeof cfg !== 'object') return false;
    if (!cfg.fields || typeof cfg.fields !== 'object') return false;
    for (const k of DIVIDER_FIELDS) {
      if (!_validDividerField(cfg.fields[k])) return false;
    }
    if (!DIVIDER_LAYOUTS.includes(cfg.layout)) return false;
    return true;
  }
  function loadDividerConfig() {
    try {
      const raw = localStorage.getItem(DIVIDER_CONFIG_KEY);
      if (!raw) return null;
      const parsed = JSON.parse(raw);
      return _validDividerConfig(parsed) ? parsed : null;
    } catch { return null; }
  }
  function saveDividerConfig(cfg) {
    if (!_validDividerConfig(cfg)) return false;
    try {
      localStorage.setItem(DIVIDER_CONFIG_KEY, JSON.stringify(cfg));
      return true;
    } catch { return false; }
  }
  // Convert Phase 1 preset → Phase 2 config shape (returns new object, immutable source).
  function presetToConfig(preset) {
    const p = getDividerPresetConfig(preset);
    const fld = (vis, size) => ({ visible: !!vis, size });
    // aliasScale 1.6 (minimal) / 1.3 (standard) → 'l'; 1.0 → 'm'.
    const aliasSize = (p.aliasScale && p.aliasScale >= 1.2) ? 'l' : 'm';
    return {
      fields: {
        alias:   fld(p.showAlias,   aliasSize),
        name:    fld(p.showName,    'm'),
        variant: fld(p.showVariant, 'm'),
        carrier: fld(p.showCarrier, 'm'),
        image:   fld(p.showImage,   'm'),
        qty:     fld(p.showQty,     'm'),
        worker:  fld(p.showWorker,  'm'),
        footer:  fld(p.showFooter,  's'),
      },
      layout: p.imageFirst ? 'compact' : 'stacked',
      // keep imageFirst for buildDividerPage consumption
      imageFirst: !!p.imageFirst,
    };
  }
  // Resolution: explicit config → preset fallback.
  function getEffectiveDividerConfig() {
    const cfg = loadDividerConfig();
    if (cfg) {
      // attach imageFirst flag (derived from layout) for buildDividerPage.
      return { ...cfg, imageFirst: cfg.layout === 'compact' };
    }
    return presetToConfig(loadDividerPreset());
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

  // ── Label overlay config (masks + marketing text + shop image + headers) ──
  const LABEL_OVERLAY_KEY = 'qf_label_overlay_v1';
  function loadLabelOverlay() {
    try { return Object.assign({ enabled: false, marketingText: 'กรุณาถ่ายรูปก่อนเปิดกล่องพัสดุ', shopImageDataUrl: '', headerMain: '', headerSub: '', opacity: 0.85, aliasFontSize: 0, marketingFontSize: 0, header1FontSize: 0, header2FontSize: 0 }, JSON.parse(localStorage.getItem(LABEL_OVERLAY_KEY) || 'null') || {}); }
    catch { return { enabled: false, marketingText: 'กรุณาถ่ายรูปก่อนเปิดกล่องพัสดุ', shopImageDataUrl: '', headerMain: '', headerSub: '', opacity: 0.85 }; }
  }
  function saveLabelOverlay(cfg) { try { localStorage.setItem(LABEL_OVERLAY_KEY, JSON.stringify(cfg)); } catch {} }

  // Calibrated J&T mask rects (bottom-left PDF coords, A6 298×420pt, pixel-scanned @ 2x).
  // Side columns extended to y=0 (page bottom) to fully cover vertical tracking
  // numbers that overflow below the previous y=38 floor (≈5 extra chars visible).
  // Width unchanged — widening would clip address block which starts at x≈14.
  const J_AND_T_MASK_RECTS = [
    { x: 0,   y: 0,    w: 17,  h: 340 }, // vertical OCR column L — y=0..340 (top 12-digit group at y≈345-390 left exposed)
    { x: 283, y: 0,    w: 15,  h: 340 }, // vertical OCR column R — same logic
    { x: 0,   y: 41.5, w: 95,  h: 20  }, // TikTok Shop footer logo (preserve Order ID x>190)
    { x: 0,   y: 75,   w: 298, h: 3   }, // h-line above Qty Total
    { x: 0,   y: 62,   w: 298, h: 2   }, // h-line above footer row
    { x: 0,   y: 39.5, w: 298, h: 2   }, // h-line below footer row
  ];

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
    samplePreviewUrl: null,        // vendor/jnt_sample_preview.png URL (PDF editor background)
    carrierLogoUrls: {},           // { tiktok: url, jnt: url } — phase 3 system element overrides
    carrierLogoBytes: {},          // { tiktok: Uint8Array, jnt: Uint8Array } — cached after first fetch
    records: new Map(),            // fulfillUnitId → {skuList:[{productId,productName,quantity,...}]}
    weirdFulfillUnitIds: new Set(),
    weirdCombos: new Map(),        // sigKey → {sigKey, items:[{productId,productName,productImageURL,quantity}], fulfillUnitIds: Set, count}
    carriers: new Map(),           // carrierId → {id, name, iconUrl, count}
    carrierOf: new Map(),          // fulfillUnitId → carrierId
    carrierFilter: new Set(),      // empty = all carriers
    preOrderOf: new Map(),         // fulfillUnitId → boolean (true = pre-order)
    preOrderFilter: 'all',         // 'all' | 'preorder' | 'normal'
    printedUnitIds: new Set(),     // fulfillUnitId พิมพ์แล้วใน session นี้ (labels page only, clear on scan)
    selectMode: false,
    selected: new Map(),           // key → {type, productId, skuId, scenario, sigKey}
    dateFilter: { start: null, end: null, field: 'createTime' }, // field: createTime | shipByTime | autoCancelTime
    advancedOpen: false,
    calendarMonth: null, // {year, month} cursor for the visible month
    overlayEnabled: loadOverlayPref(),
    workers: loadWorkers(),          // [{id, name, icon}] คนแพ็ค
    sellerEmail: null,               // §7.6: populated at widget boot from TikTok session cookie
    addressCaptureEnabled: loadAddressCapturePref(), // pre-print address book backfill
  };

  function loadAddressCapturePref() {
    const v = localStorage.getItem(ADDRESS_CAPTURE_PREF_KEY);
    return v === null ? true : v === 'true'; // default ON
  }
  function saveAddressCapturePref(enabled) {
    localStorage.setItem(ADDRESS_CAPTURE_PREF_KEY, String(!!enabled));
  }

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

  const WORKER_ICONS = ['★', '♥', '♦', '♣', '♠', '●', '■', '▲', '◆', '✿', '☀', '♪'];
  const WORKER_COLOR_TO_ICON = { '#2563eb': '★', '#16a34a': '♥', '#ea580c': '♦', '#dc2626': '■', '#9333ea': '▲' };

  function loadWorkers() {
    try {
      const raw = localStorage.getItem(WORKERS_STORAGE_KEY);
      if (!raw) return [];
      const arr = JSON.parse(raw);
      const migrated = arr.map(w => {
        if (w.icon) return w;
        const icon = WORKER_COLOR_TO_ICON[w.color] || '●';
        const { color: _drop, ...rest } = w;
        return { ...rest, icon };
      });
      return migrated;
    } catch { return []; }
  }

  function saveWorkers() {
    localStorage.setItem(WORKERS_STORAGE_KEY, JSON.stringify(state.workers));
  }

  // ==================== TEAMS ====================
  // §2: Team concept — groups of workers sharing a planning column.
  // Schema: Team { id, name, memberWorkerIds[], createdAt }
  // localStorage key: qf_teams_v1 (JSON: Team[])

  const TEAMS_STORAGE_KEY = 'qf_teams_v1';

  function loadTeams() {
    try {
      const raw = localStorage.getItem(TEAMS_STORAGE_KEY);
      if (!raw) return [];
      const arr = JSON.parse(raw);
      if (!Array.isArray(arr)) return [];
      return arr;
    } catch { return []; }
  }
  // Bootstrap: attach teams to state (state object is created before this module).
  state.teams = loadTeams();

  function saveTeams(teams) {
    try {
      localStorage.setItem(TEAMS_STORAGE_KEY, JSON.stringify(teams));
    } catch (e) {
      console.warn('[QF] teams save failed:', e);
    }
  }

  function createTeam({ name, memberWorkerIds }) {
    const id = 't_' + Math.random().toString(36).slice(2, 10);
    const team = { id, name: String(name || '').trim(), memberWorkerIds: memberWorkerIds || [], createdAt: Date.now() };
    state.teams = [...state.teams, team];
    saveTeams(state.teams);
    return team;
  }

  function updateTeam(id, patch) {
    const idx = state.teams.findIndex(t => t.id === id);
    if (idx < 0) return null;
    const updated = { ...state.teams[idx], ...patch };
    state.teams = state.teams.map(t => t.id === id ? updated : t);
    saveTeams(state.teams);
    return updated;
  }

  function deleteTeam(id) {
    state.teams = state.teams.filter(t => t.id !== id);
    saveTeams(state.teams);
  }

  function getTeam(id) {
    return state.teams.find(t => t.id === id);
  }

  function teamIcon() {
    // Teams always use the 👥 glyph in UI (not in PDF watermark).
    return '👥';
  }

  // Called when a worker is deleted — remove that workerId from all teams.
  function removeWorkerFromTeams(workerId) {
    const changed = state.teams.some(t => t.memberWorkerIds.includes(workerId));
    if (!changed) return;
    state.teams = state.teams.map(t => ({
      ...t,
      memberWorkerIds: t.memberWorkerIds.filter(wid => wid !== workerId),
    }));
    saveTeams(state.teams);
  }

  // ==================== PRINT HISTORY ====================
  // History stores fulfillUnitId arrays + metadata (NOT PDF bytes). Re-download
  // rebuilds the PDF by re-calling TikTok's generate API with the same IDs —
  // no extra quota cost since it's the same request TikTok already accepted.
  const HISTORY_KEY = 'qf_print_history_v1';
  const HISTORY_MAX_ENTRIES = 200;
  const HISTORY_MAX_AGE_MS = 30 * 24 * 60 * 60 * 1000; // 30 days
  const HISTORY_TRIM_ON_QUOTA = 50;

  function loadHistory() {
    try {
      const raw = localStorage.getItem(HISTORY_KEY);
      if (!raw) return [];
      const arr = JSON.parse(raw);
      if (!Array.isArray(arr)) return [];
      const cutoff = Date.now() - HISTORY_MAX_AGE_MS;
      return arr.filter(e => e && typeof e.timestamp === 'number' && e.timestamp >= cutoff);
    } catch { return []; }
  }

  function saveHistory(arr) {
    try {
      localStorage.setItem(HISTORY_KEY, JSON.stringify(arr));
    } catch (e) {
      // Quota exceeded → trim aggressively and retry
      try {
        const trimmed = arr.slice(0, HISTORY_TRIM_ON_QUOTA);
        localStorage.setItem(HISTORY_KEY, JSON.stringify(trimmed));
      } catch (e2) {
        console.warn('[QF] history save failed even after trim:', e2);
      }
    }
  }

  function addHistoryEntry(entry) {
    const arr = loadHistory();
    arr.unshift(entry);
    const capped = arr.slice(0, HISTORY_MAX_ENTRIES);
    saveHistory(capped);
  }

  function deleteHistoryEntry(id) {
    const arr = loadHistory().filter(e => e.id !== id);
    saveHistory(arr);
  }

  function clearAllHistory() {
    saveHistory([]);
  }

  // ==================== ORDER HISTORY (IndexedDB) ====================
  // Stores one record per orderId per print job — 90-day rolling window.
  // Schema: { orderId, fulfillUnitId, ts, date, carrier,
  //           assigneeKind, teamId, teamName, teamSnapshot, workerId, workerName }

  const ORDER_HISTORY_DB = 'qf_order_history';
  const ORDER_HISTORY_STORE = 'orders';
  const ORDER_HISTORY_MAX_AGE_MS = 90 * 24 * 60 * 60 * 1000;

  let _orderHistoryDb = null;

  function openOrderHistoryDb() {
    if (_orderHistoryDb) return Promise.resolve(_orderHistoryDb);
    return new Promise((resolve, reject) => {
      const req = indexedDB.open(ORDER_HISTORY_DB, 1);
      req.onupgradeneeded = (e) => {
        const db = e.target.result;
        if (!db.objectStoreNames.contains(ORDER_HISTORY_STORE)) {
          const store = db.createObjectStore(ORDER_HISTORY_STORE, { autoIncrement: true });
          store.createIndex('orderId', 'orderId', { unique: false });
          store.createIndex('ts', 'ts', { unique: false });
        }
      };
      req.onsuccess = (e) => { _orderHistoryDb = e.target.result; resolve(_orderHistoryDb); };
      req.onerror = () => reject(req.error);
    });
  }

  async function saveOrderHistoryBatch(records) {
    if (!records.length) return;
    try {
      const db = await openOrderHistoryDb();
      const cutoff = Date.now() - ORDER_HISTORY_MAX_AGE_MS;
      await new Promise((resolve, reject) => {
        const tx = db.transaction(ORDER_HISTORY_STORE, 'readwrite');
        const store = tx.objectStore(ORDER_HISTORY_STORE);
        // Purge old records via ts index
        const idx = store.index('ts');
        const range = IDBKeyRange.upperBound(cutoff);
        idx.openCursor(range).onsuccess = function(e) {
          const cursor = e.target.result;
          if (cursor) { cursor.delete(); cursor.continue(); }
        };
        for (const r of records) store.add(r);
        tx.oncomplete = resolve;
        tx.onerror = () => reject(tx.error);
      });
    } catch (e) {
      console.warn('[QF] saveOrderHistoryBatch failed:', e);
    }
  }

  async function queryOrderHistory(orderId) {
    try {
      const db = await openOrderHistoryDb();
      return await new Promise((resolve, reject) => {
        const tx = db.transaction(ORDER_HISTORY_STORE, 'readonly');
        const idx = tx.objectStore(ORDER_HISTORY_STORE).index('orderId');
        const req = idx.getAll(IDBKeyRange.only(String(orderId)));
        req.onsuccess = () => resolve(req.result || []);
        req.onerror = () => reject(req.error);
      });
    } catch (e) {
      console.warn('[QF] queryOrderHistory failed:', e);
      return [];
    }
  }

  // ==================== ADDRESS BOOK (captures unmasked recipient data BEFORE print) ====================
  //
  // Why: once a label is printed + TTS'd, TikTok masks the recipient fields
  // in every subsequent /order/get response (button_status:3, mask_infos set).
  // So we hook the call path BEFORE the first print to snapshot the unmasked
  // data into an IndexedDB address book. Two capture sources:
  //   1. Passive sniff — the fetch hook clones every /order/get response and
  //      parses it off-thread (no mutation, no extra traffic).
  //   2. Active backfill — right before buildChunkPdf calls the generate API,
  //      we fire /order/get for any IDs we haven't captured yet (throttled,
  //      only when we have the captured template so we mimic TikTok's exact
  //      shape).
  //
  // Schema: {
  //   orderId, fulfillUnitId, recipientName, recipientPhone,
  //   addressDetail, subDistrict, district, province, zipcode, countryCode,
  //   carrier, productNames, capturedAt, source('sniff'|'backfill'), isMasked
  // }
  //
  // 90-day rolling window; orderId is the primary index (updates replace).

  const ADDRESS_BOOK_DB = 'qf_address_book';
  const ADDRESS_BOOK_STORE = 'addresses';
  const ADDRESS_BOOK_MAX_AGE_MS = 90 * 24 * 60 * 60 * 1000;
  let _addressBookDb = null;

  function openAddressBookDb() {
    if (_addressBookDb) return Promise.resolve(_addressBookDb);
    return new Promise((resolve, reject) => {
      const req = indexedDB.open(ADDRESS_BOOK_DB, 1);
      req.onupgradeneeded = (e) => {
        const db = e.target.result;
        if (!db.objectStoreNames.contains(ADDRESS_BOOK_STORE)) {
          const store = db.createObjectStore(ADDRESS_BOOK_STORE, { keyPath: 'orderId' });
          store.createIndex('fulfillUnitId', 'fulfillUnitId', { unique: false });
          store.createIndex('capturedAt', 'capturedAt', { unique: false });
          store.createIndex('province', 'province', { unique: false });
        }
      };
      req.onsuccess = (e) => { _addressBookDb = e.target.result; resolve(_addressBookDb); };
      req.onerror = () => reject(req.error);
    });
  }

  // Extract the first non-null string among possible field paths. Accepts
  // either direct key access or dotted paths (for nested lookups).
  function pick(obj, ...keys) {
    if (!obj) return null;
    for (const k of keys) {
      const parts = k.split('.');
      let cur = obj;
      for (const p of parts) {
        if (cur == null) { cur = null; break; }
        cur = cur[p];
      }
      if (cur != null && cur !== '') return cur;
    }
    return null;
  }

  // Parse a single TikTok order detail object into our address book shape.
  // Returns null if the essential recipient block is missing or fully masked.
  function parseOrderDetail(od) {
    if (!od || typeof od !== 'object') return null;

    const orderId = String(pick(od, 'order_id', 'main_order_id', 'orderId') || '');
    if (!orderId) return null;

    // Possible recipient containers seen across TikTok API versions
    const r = od.recipient_address || od.recipient || od.shipping_info
           || od.receiver_info || od.buyer_info || od.delivery_address || {};

    const name = String(pick(r, 'name', 'receiver_name', 'recipient_name', 'full_name',
                             'first_name') || '').trim();
    const phone = String(pick(r, 'phone', 'phone_number', 'mobile', 'tel',
                              'receiver_phone') || '').trim();

    // Region fields: TikTok usually returns an array like
    //   region_fields: [{ level: 1, name: 'กรุงเทพ' }, { level: 2, name: 'เขต…' }, ...]
    let province = null, district = null, subDistrict = null, zipcode = null, countryCode = null;
    const regions = r.region_fields || r.regions || od.region_fields || [];
    if (Array.isArray(regions)) {
      for (const f of regions) {
        const lvl = f.level ?? f.region_level;
        const nm  = f.name || f.region_name;
        if (!nm) continue;
        if (lvl === 1 || lvl === 'province' || lvl === 'region_level_1') province = nm;
        else if (lvl === 2 || lvl === 'city' || lvl === 'district' || lvl === 'region_level_2') district = nm;
        else if (lvl === 3 || lvl === 'subdistrict' || lvl === 'region_level_3') subDistrict = nm;
      }
    }
    // Fallback flat fields
    province = province || pick(r, 'province', 'state', 'region_name_level_1', 'city_name');
    district = district || pick(r, 'district', 'city', 'region_name_level_2');
    subDistrict = subDistrict || pick(r, 'sub_district', 'subdistrict', 'region_name_level_3', 'ward');
    zipcode = pick(r, 'zipcode', 'post_code', 'postal_code', 'zip');
    countryCode = pick(r, 'country_code', 'region_code');

    const detail = pick(r, 'detail_address', 'address_detail', 'address_line_1',
                           'address_line1', 'full_address', 'address', 'street');

    // Mask detection: TikTok's mask_infos or button_status:3 signal
    const buttonStatus = pick(od, 'button_status', 'address_button_status');
    const maskInfos    = od.mask_infos || od.address_mask_infos;
    const looksMasked = buttonStatus === 3
                      || (Array.isArray(maskInfos) && maskInfos.length > 0)
                      || /\*{2,}/.test(name) || /\*{2,}/.test(detail || '');

    // Carrier + first product name for context
    const carrier = pick(od, 'shipping_provider_name', 'shipping_provider.name',
                             'delivery_module.shipping_provider_name');
    const skus = od.sku_module || od.items || od.item_list || [];
    const productNames = Array.isArray(skus)
      ? skus.map(s => s.product_name || s.productName || '').filter(Boolean).slice(0, 3).join(' | ')
      : null;

    const fulfillUnitId = String(pick(od, 'fulfill_unit_id', 'fulfillUnitId',
                                          'label_module.fulfill_unit_id') || '');

    // Reject if nothing useful was extracted
    if (!name && !phone && !detail && !province) return null;

    return {
      orderId,
      fulfillUnitId: fulfillUnitId || null,
      recipientName: name || null,
      recipientPhone: phone || null,
      addressDetail: detail || null,
      subDistrict, district, province, zipcode, countryCode,
      carrier: carrier || null,
      productNames: productNames || null,
      capturedAt: Date.now(),
      isMasked: looksMasked,
      source: 'sniff',
    };
  }

  // Entry point: response sniffer in the fetch hook calls this with the
  // parsed JSON body of /order/get. Handles both single-order and array shapes.
  function captureAddressFromOrderGet(json) {
    if (!json || json.code !== 0) return;
    const d = json.data || json.result || {};
    // Common shapes: main_order[0], order_list[], order (single)
    const orders = d.main_order || d.orders || d.order_list || d.list
                 || (d.order ? [d.order] : []) || [];
    const parsed = [];
    for (const od of Array.isArray(orders) ? orders : [orders]) {
      const rec = parseOrderDetail(od);
      if (rec) parsed.push(rec);
      // §Note capture: same response carries the buyer/seller remarks. The
      // labels list API only exposes a `has_*_note` flag (often unreliable —
      // it stays false even when a note exists), so we treat order/get as
      // the source of truth and patch state.records when either side has
      // a non-empty note. Fires non-blocking, ignores parse failures.
      try { captureNotesFromOrderDetail(od); } catch (e) {}
    }
    if (parsed.length) {
      saveAddressBatch(parsed, 'sniff')
        .then(() => { try { refreshAddressBookBadge(); } catch (e) {} })
        .catch(e => {});
    }
  }

  // Pull buyer/seller remark text out of an order/get response object and
  // patch every state.records entry whose orderIds contains this main_order_id
  // (a single order may map to multiple fulfill units when split-shipped).
  function captureNotesFromOrderDetail(od) {
    if (!od || typeof od !== 'object') return;
    const orderId = String(
      od.order_id || od.main_order_id || od.orderId
      || od.order?.order_id || od.order?.main_order_id || ''
    );
    if (!orderId) return;
    // Try every observed shape for the note text.
    const buyerSrc  = od.buyer_remark  || od.buyerRemark  || {};
    const sellerSrc = od.seller_remark || od.sellerRemark || {};
    const readNote = (src) => {
      if (!src) return '';
      if (typeof src === 'string') return src.trim();
      if (typeof src !== 'object') return '';
      const v = src.note || src.message || src.text || src.content || src.value;
      return typeof v === 'string' ? v.trim() : '';
    };
    const buyerNote  = readNote(buyerSrc)  || (typeof od.buyer_message === 'string' ? od.buyer_message.trim() : '');
    const sellerNote = readNote(sellerSrc);
    if (!buyerNote && !sellerNote) return;

    // Patch all matching records (immutable copy per CLAUDE.md rules).
    for (const [fid, rec] of state.records) {
      if (!rec.orderIds || !rec.orderIds.includes(orderId)) continue;
      const next = {
        ...rec,
        buyerNote:     buyerNote  || rec.buyerNote  || '',
        sellerNote:    sellerNote || rec.sellerNote || '',
        hasBuyerNote:  !!(buyerNote  || rec.buyerNote),
        hasSellerNote: !!(sellerNote || rec.sellerNote),
        hasNote:       !!(buyerNote || sellerNote || rec.buyerNote || rec.sellerNote),
      };
      state.records.set(fid, next);
    }
  }

  // ==================== BUYER CONTACT INFO (EYE-ICON ENDPOINT) ====================
  // /api/fulfillment/orders/buyer_contact_info/get — the REAL decrypt call
  // fired when the user clicks the 👁 eye icon next to a masked field on the
  // order detail page. One POST per field. Empirically observed shape:
  //
  //   Request:   { main_order_id: "...", contact_info_type: 0|1|2|3 }
  //   Response:  { code: 0, data: { plain_text_*: ... } }
  //
  // Response keys per type:
  //   0  →  plain_text_name           (string)
  //   1  →  plain_text_address        (object: {items[], region, districts[]})
  //   2  →  plain_text_phone_number   (string, e.g. "(+66)659549268")
  //   3  →  plain_text_nickname       (string, e.g. "toeystory.tt")
  //
  // Because each call returns only one field, we MERGE the patch into the
  // existing address-book record keyed by orderId.
  const CONTACT_TYPE_NAME     = 0;
  const CONTACT_TYPE_ADDRESS  = 1;
  const CONTACT_TYPE_PHONE    = 2;
  const CONTACT_TYPE_NICKNAME = 3;

  // Parse the per-type response into a partial patch object. Returns null if
  // the response body doesn't carry the expected field (e.g. wrong type or
  // server returned an error wrapper we don't recognize).
  function parseBuyerContactResponse(type, respData) {
    if (!respData || typeof respData !== 'object') return null;
    const patch = {};
    if (type === CONTACT_TYPE_NAME && respData.plain_text_name) {
      patch.recipientName = String(respData.plain_text_name).trim();
    } else if (type === CONTACT_TYPE_PHONE && respData.plain_text_phone_number) {
      patch.recipientPhone = String(respData.plain_text_phone_number).trim();
    } else if (type === CONTACT_TYPE_NICKNAME && respData.plain_text_nickname) {
      patch.buyerNickname = String(respData.plain_text_nickname).trim();
    } else if (type === CONTACT_TYPE_ADDRESS && respData.plain_text_address) {
      const addr = respData.plain_text_address;
      const items = {};
      if (Array.isArray(addr.items)) {
        for (const it of addr.items) {
          if (it && typeof it.key === 'string') items[it.key] = it.value;
        }
      }
      // items keys observed: zipcode, address, address_detail, house_number
      const detail = [items.house_number, items.address, items.address_detail]
        .map(v => (v == null ? '' : String(v).trim()))
        .filter(Boolean)
        .join(' ');
      patch.addressDetail = detail || null;
      patch.zipcode = items.zipcode || null;
      // districts[] ordering observed: [0]=province, [1]=district, [2]=subdistrict
      const ds = Array.isArray(addr.districts) ? addr.districts : [];
      if (ds[0]?.name) patch.province = String(ds[0].name).trim();
      if (ds[1]?.name) patch.district = String(ds[1].name).trim();
      if (ds[2]?.name) patch.subDistrict = String(ds[2].name).trim();
      if (addr.region?.name) patch.countryCode = String(addr.region.name).trim();
    } else {
      return null;
    }
    return Object.keys(patch).length ? patch : null;
  }

  // Response-side entry point. Called with the paired request body + parsed
  // response JSON from the fetch hook.
  function captureBuyerContactFieldResponse(reqBody, respJson, fulfillUnitIdHint = null) {
    if (!reqBody || !respJson) return;
    // Detect both non-zero-code limits AND soft-reject (code:0 + rejection).
    if (isRateLimitResponse(respJson)) {
      recordRateLimitHit(respJson);
      return;
    }
    if (respJson.code !== 0) return;
    const orderId = String(reqBody.main_order_id || reqBody.order_id || '');
    if (!orderId) return;
    const type = Number(reqBody.contact_info_type);
    const patch = parseBuyerContactResponse(type, respJson.data || {});
    if (!patch) return;
    patch.capturedAt = Date.now();
    patch.isMasked = false; // we have plaintext, so definitively unmasked
    patch.source = 'eye';
    mergeAddressPatch(orderId, patch, fulfillUnitIdHint)
      .then(() => { try { refreshAddressBookBadge(); } catch (e) {} })
      .catch(() => {});
  }

  // Merge a partial field patch into the address book record, preserving
  // existing unmasked fields from prior calls (since buyer_contact_info fires
  // one call per field, a full record accumulates over 4 calls).
  async function mergeAddressPatch(orderId, patch, fulfillUnitIdHint = null) {
    try {
      const db = await openAddressBookDb();
      await new Promise((resolve, reject) => {
        const tx = db.transaction(ADDRESS_BOOK_STORE, 'readwrite');
        const store = tx.objectStore(ADDRESS_BOOK_STORE);
        const getReq = store.get(String(orderId));
        getReq.onsuccess = () => {
          const existing = getReq.result || { orderId: String(orderId) };
          // Immutable merge — new patch wins per field, unmask flag sticks
          // at unmasked once any unmasked patch arrived.
          const merged = {
            ...existing,
            ...patch,
            orderId: String(orderId),
            fulfillUnitId: existing.fulfillUnitId || fulfillUnitIdHint || null,
            isMasked: patch.isMasked === false ? false : (existing.isMasked !== true ? existing.isMasked : true),
            capturedAt: patch.capturedAt || existing.capturedAt || Date.now(),
            source: patch.source || existing.source || 'sniff',
          };
          store.put(merged);
        };
        tx.oncomplete = resolve;
        tx.onerror = () => reject(tx.error);
      });
    } catch (e) {
      console.warn('[QF] mergeAddressPatch failed:', e);
    }
  }

  // ==================== RATE-LIMIT CIRCUIT BREAKER ====================
  // User-confirmed: buyer_contact_info is rate-limited. Unknown exact threshold.
  // Strategy: start conservative, learn from 429/code-error responses, back off
  // exponentially, circuit-break for the rest of the session on repeated hits.
  const RATE_LIMIT_STATE = {
    consecutiveHits: 0,
    totalHits: 0,
    lastHitAt: 0,
    backoffUntil: 0,        // epoch ms; skip all calls until this time
    tripped: false,         // permanently disabled for this session
    currentGapMs: 600,      // adaptive per-call gap (grows on each hit)
    currentConcurrency: 1,  // adaptive concurrency (shrinks on each hit)
  };

  // Detect error codes / messages that signal rate limiting. TikTok returns
  // non-zero code + message strings — we match on common substrings.
  //
  // CRITICAL: TikTok also uses a "soft reject" pattern — `code: 0` but
  // `data.contact_customer_service: true` + `rejection_message` — meaning
  // the quota is burned for PII (name/address/phone types 0/1/2). Nickname
  // (type 3) is NOT protected so that call can still succeed with code:0.
  // We treat soft-reject as a rate-limit hit too because hammering past it
  // is pointless — the server will keep rejecting until quota resets
  // (appears to be per-account per-day).
  function isRateLimitResponse(respJson) {
    if (!respJson) return false;
    // Soft reject: server returns success code but refuses to hand out PII.
    if (respJson.code === 0 && respJson.data && respJson.data.contact_customer_service === true) {
      return true;
    }
    if (respJson.code === 0) return false;
    const msg = String(respJson.message || respJson.msg || '').toLowerCase();
    // Known substrings we've seen / likely to see. Kept conservative so a
    // plain "not found" doesn't trip the breaker.
    return /rate[\s_-]?limit|too many|frequently|request limit|qps|throttle|429/.test(msg)
        || respJson.code === 429
        || respJson.code === 50002 // arbitrary TikTok "too frequent" seen in other endpoints
        || respJson.code === 50005;
  }

  function recordRateLimitHit(respJson) {
    if (!isRateLimitResponse(respJson)) return;
    RATE_LIMIT_STATE.consecutiveHits += 1;
    RATE_LIMIT_STATE.totalHits += 1;
    RATE_LIMIT_STATE.lastHitAt = Date.now();
    // Exponential backoff: 2s, 4s, 8s, ..., cap 5 min
    const backoff = Math.min(2000 * 2 ** (RATE_LIMIT_STATE.consecutiveHits - 1), 300000);
    RATE_LIMIT_STATE.backoffUntil = Date.now() + backoff;
    // Adaptive slow-down for subsequent calls
    RATE_LIMIT_STATE.currentGapMs = Math.min(RATE_LIMIT_STATE.currentGapMs * 2, 10000);
    RATE_LIMIT_STATE.currentConcurrency = 1;
    if (RATE_LIMIT_STATE.consecutiveHits >= 3) {
      // Three strikes within session → circuit-break. TikTok typically uses
      // per-account per-day quota for PII reveals — the rejection will keep
      // firing until the quota resets (appears to be ~24h).
      RATE_LIMIT_STATE.tripped = true;
      console.warn('[QF] buyer_contact_info quota exhausted. Backfill disabled until quota resets (typically ~24h).');
      try { showToast('⚠️ สมุดที่อยู่: ดึงข้อมูลเกินโควต้ารายวัน → ปิดการดึงอัตโนมัติ (reset พรุ่งนี้)', 4500); } catch (e) {}
    } else {
      console.warn(`[QF] buyer_contact_info rate-limit hit #${RATE_LIMIT_STATE.consecutiveHits}, backing off ${backoff}ms`);
    }
  }

  function recordRateLimitSuccess() {
    if (RATE_LIMIT_STATE.consecutiveHits > 0) {
      RATE_LIMIT_STATE.consecutiveHits = 0;
      // Gentle recovery: shrink gap by 25% per success, floor 400ms
      RATE_LIMIT_STATE.currentGapMs = Math.max(400, Math.floor(RATE_LIMIT_STATE.currentGapMs * 0.75));
    }
  }

  function rateLimitGateOpen() {
    if (RATE_LIMIT_STATE.tripped) return false;
    if (Date.now() < RATE_LIMIT_STATE.backoffUntil) return false;
    return true;
  }

  async function saveAddressBatch(records, source = 'sniff') {
    if (!records.length) return;
    try {
      const db = await openAddressBookDb();
      const cutoff = Date.now() - ADDRESS_BOOK_MAX_AGE_MS;
      await new Promise((resolve, reject) => {
        const tx = db.transaction(ADDRESS_BOOK_STORE, 'readwrite');
        const store = tx.objectStore(ADDRESS_BOOK_STORE);
        // Purge old records first
        const idx = store.index('capturedAt');
        idx.openCursor(IDBKeyRange.upperBound(cutoff)).onsuccess = function(e) {
          const cur = e.target.result;
          if (cur) { cur.delete(); cur.continue(); }
        };
        for (const r of records) {
          // Don't overwrite an UNMASKED record with a MASKED one — we may have
          // captured it earlier when the order was still pre-print.
          const getReq = store.get(r.orderId);
          getReq.onsuccess = () => {
            const existing = getReq.result;
            if (existing && !existing.isMasked && r.isMasked) return; // keep clean copy
            store.put({ ...r, source });
          };
        }
        tx.oncomplete = resolve;
        tx.onerror = () => reject(tx.error);
      });
    } catch (e) {
      console.warn('[QF] saveAddressBatch failed:', e);
    }
  }

  async function getAddressByOrderId(orderId) {
    try {
      const db = await openAddressBookDb();
      return await new Promise((resolve, reject) => {
        const tx = db.transaction(ADDRESS_BOOK_STORE, 'readonly');
        const req = tx.objectStore(ADDRESS_BOOK_STORE).get(String(orderId));
        req.onsuccess = () => resolve(req.result || null);
        req.onerror = () => reject(req.error);
      });
    } catch (e) { return null; }
  }

  async function getAllAddresses() {
    try {
      const db = await openAddressBookDb();
      return await new Promise((resolve, reject) => {
        const tx = db.transaction(ADDRESS_BOOK_STORE, 'readonly');
        const req = tx.objectStore(ADDRESS_BOOK_STORE).getAll();
        req.onsuccess = () => resolve(req.result || []);
        req.onerror = () => reject(req.error);
      });
    } catch (e) { return []; }
  }

  async function clearAddressBook() {
    try {
      const db = await openAddressBookDb();
      await new Promise((resolve, reject) => {
        const tx = db.transaction(ADDRESS_BOOK_STORE, 'readwrite');
        tx.objectStore(ADDRESS_BOOK_STORE).clear();
        tx.oncomplete = resolve;
        tx.onerror = () => reject(tx.error);
      });
    } catch (e) {}
  }

  // Active backfill — called from buildChunkPdf right before the generate API.
  // Prefers the eye-click endpoint (buyer_contact_info) because it's the ONLY
  // endpoint that returns definitively-unmasked data. Falls back to /order/get
  // when the eye-click template hasn't been captured yet.
  //
  // RATE LIMITING (user-confirmed hit during manual testing, exact threshold
  // unknown): we start conservative and adapt. See RATE_LIMIT_STATE.
  const ADDRESS_BACKFILL_CONCURRENCY = 1;   // start at 1 (serialize by default)
  const ADDRESS_BACKFILL_GAP_MS_INIT = 600; // 600ms per call → ~100/min max
  const ADDRESS_BACKFILL_MAX_PER_SESSION = 400; // hard cap, tune later
  let _addressBackfillCount = 0;

  async function tryBackfillAddresses(fulfillIds) {
    if (!Array.isArray(fulfillIds) || !fulfillIds.length) return;
    if (state.addressCaptureEnabled === false) return; // user-disabled
    if (!rateLimitGateOpen()) return; // circuit-breaker open / backoff active

    // Resolve orderIds from state.records (fulfillUnitId → orderIds) and
    // determine which contact_info_type slots are still missing per order.
    const tasks = [];
    // Note tasks track every (orderId, fulfillUnitId) regardless of whether
    // the address book is already full — addresses and notes are independent
    // and notes need their own backfill pass.
    const noteTasks = [];
    for (const fid of fulfillIds) {
      const rec = state.records.get(fid);
      if (!rec) continue;
      const oids = rec.orderIds || [];
      for (const oid of oids) {
        const existing = await getAddressByOrderId(oid);
        const full = existing && !existing.isMasked
                     && existing.recipientName
                     && existing.recipientPhone
                     && existing.addressDetail;
        // Always queue for note backfill if note text is missing on the
        // record, regardless of address-book state.
        if (!rec.buyerNote && !rec.sellerNote) {
          noteTasks.push({ orderId: String(oid), fulfillUnitId: String(fid) });
        }
        if (full) continue;
        const needTypes = [];
        if (!existing?.recipientName)   needTypes.push(CONTACT_TYPE_NAME);
        if (!existing?.addressDetail)   needTypes.push(CONTACT_TYPE_ADDRESS);
        if (!existing?.recipientPhone)  needTypes.push(CONTACT_TYPE_PHONE);
        tasks.push({ orderId: String(oid), fulfillUnitId: String(fid), needTypes });
      }
    }

    // Notes (buyer/seller remarks) live ONLY in /order/get's response — the
    // labels list and buyer_contact endpoints don't carry them — so we fire
    // a parallel notes backfill that derives /order/get URL from the labels
    // API URL when needed (works without a captured order_get template).
    const notesPromise = noteTasks.length ? tryBackfillNotesViaOrderGet(noteTasks) : Promise.resolve();

    if (!tasks.length) {
      // All addresses already in book — only notes remain.
      await notesPromise;
      return;
    }
    if (_buyerContactUrl && _buyerContactBodyTemplate) {
      await Promise.all([tryBackfillViaBuyerContact(tasks), notesPromise]);
      return;
    }
    if (_orderGetUrl && _orderGetBodyTemplate) {
      // No buyer_contact template yet — order/get returns address + notes in
      // one shot, so the existing flow covers both.
      await tryBackfillViaOrderGet(tasks);
      return;
    }
    // Neither full template available — notes backfill still tries via the
    // derived URL. Address won't backfill without a real template.
    await notesPromise;
  }

  // Verified empirically: TikTok's request signing (msToken / X-Bogus /
  // X-Gnarly) is PATH-SPECIFIC, so we cannot derive a working /order/get URL
  // from the labels API URL — replays return code 21001001 ("System error").
  // The extension must observe a real /order/get call (which fires when the
  // user opens any order detail page) before notes backfill can run.
  function deriveOrderGetUrl() {
    return _orderGetUrl || null;
  }

  // Notes-only backfill — one /order/get per unique orderId where the local
  // record has no buyerNote / sellerNote yet. Reuses the same throttling so
  // we don't double-bill the rate-limit budget. Skips orders we've already
  // populated from a passive sniff or an earlier print run.
  async function tryBackfillNotesViaOrderGet(tasks) {
    const orderGetUrl = deriveOrderGetUrl();
    if (!orderGetUrl || !_orderGetBodyTemplate) {
      // One-shot dev hint so the user knows why notes don't appear yet.
      if (!_noteTemplateHintShown) {
        _noteTemplateHintShown = true;
        try {
          showToast('เปิดหน้า "รายละเอียดออเดอร์" 1 ครั้ง เพื่อเปิดใช้ ★ note บนใบฉลาก', 5000);
        } catch (e) {}
        console.log('[QF] note backfill skipped — open any order detail page once to capture /order/get template');
      }
      return;
    }
    const bodyTemplate = _orderGetBodyTemplate;
    // Dedupe per orderId — multiple fulfill units can share one main order.
    const seen = new Set();
    const targets = [];
    for (const t of tasks) {
      if (seen.has(t.orderId)) continue;
      seen.add(t.orderId);
      // Skip if any record for this order already has note text — passive
      // sniff or a prior call already covered it.
      let alreadyNoted = false;
      for (const rec of state.records.values()) {
        if (rec.orderIds?.includes(t.orderId) && (rec.buyerNote || rec.sellerNote)) {
          alreadyNoted = true; break;
        }
      }
      if (!alreadyNoted) targets.push(t);
    }
    if (!targets.length) return;

    let i = 0;
    const runWorker = async () => {
      while (i < targets.length) {
        if (!rateLimitGateOpen()) break;
        const t = targets[i++];
        try {
          const body = { ...bodyTemplate, order_id: t.orderId, order_id_list: [t.orderId] };
          const r = await _origFetch.call(window, orderGetUrl, {
            method: 'POST',
            headers: { 'content-type': 'application/json' },
            body: JSON.stringify(body),
          });
          if (!r.ok) continue;
          const j = await r.json();
          if (isRateLimitResponse(j)) { recordRateLimitHit(j); break; }
          recordRateLimitSuccess();
          // Reuses the same parser — captures address (when not masked) AND
          // notes via captureNotesFromOrderDetail() inside it.
          captureAddressFromOrderGet(j);
          await sleep(RATE_LIMIT_STATE.currentGapMs);
        } catch (e) { /* best-effort */ }
      }
    };
    const workers = Array.from(
      { length: Math.max(1, RATE_LIMIT_STATE.currentConcurrency) },
      runWorker
    );
    await Promise.all(workers);
  }

  // Buyer-contact-info active backfill — fires ONE call per (order × type).
  // Self-throttled by RATE_LIMIT_STATE which grows the gap and shrinks the
  // concurrency on each 429/code-error.
  async function tryBackfillViaBuyerContact(tasks) {
    // Flatten into per-field calls so throttling is uniform.
    const calls = [];
    for (const t of tasks) {
      for (const type of t.needTypes) {
        calls.push({ orderId: t.orderId, fulfillUnitId: t.fulfillUnitId, type });
      }
    }
    if (!calls.length) return;

    // Apply session budget cap so a single print of 1000 labels can't burn
    // 4000 API calls in one go.
    const budgetRemaining = Math.max(0, ADDRESS_BACKFILL_MAX_PER_SESSION - _addressBackfillCount);
    if (budgetRemaining <= 0) {
      console.warn('[QF] buyer_contact_info session budget exhausted, skipping backfill');
      return;
    }
    const queue = calls.slice(0, budgetRemaining);

    let i = 0;
    const runWorker = async () => {
      while (i < queue.length) {
        if (!rateLimitGateOpen()) break; // trip detected mid-flight → abort
        const c = queue[i++];
        try {
          const body = {
            ...(_buyerContactBodyTemplate || {}),
            main_order_id: c.orderId,
            contact_info_type: c.type,
          };
          const r = await _origFetch.call(window, _buyerContactUrl, {
            method: 'POST',
            headers: { 'content-type': 'application/json' },
            body: JSON.stringify(body),
          });
          _addressBackfillCount += 1;
          if (!r.ok) continue;
          const j = await r.json();
          if (isRateLimitResponse(j)) {
            recordRateLimitHit(j);
            break; // drop remaining calls in this worker
          }
          recordRateLimitSuccess();
          captureBuyerContactFieldResponse(body, j, c.fulfillUnitId);
          await sleep(RATE_LIMIT_STATE.currentGapMs);
        } catch (e) { /* best-effort */ }
      }
    };
    const workers = Array.from(
      { length: Math.max(1, RATE_LIMIT_STATE.currentConcurrency) },
      runWorker
    );
    await Promise.all(workers);
  }

  // Legacy /order/get fallback — kept for when eye-click template isn't yet
  // captured. Same conservative throttle as buyer_contact_info path.
  async function tryBackfillViaOrderGet(tasks) {
    let i = 0;
    const runWorker = async () => {
      while (i < tasks.length) {
        if (!rateLimitGateOpen()) break;
        const t = tasks[i++];
        try {
          const body = { ...(_orderGetBodyTemplate || {}), order_id: t.orderId, order_id_list: [t.orderId] };
          const r = await _origFetch.call(window, _orderGetUrl, {
            method: 'POST',
            headers: { 'content-type': 'application/json' },
            body: JSON.stringify(body),
          });
          if (!r.ok) continue;
          const j = await r.json();
          if (isRateLimitResponse(j)) { recordRateLimitHit(j); break; }
          recordRateLimitSuccess();
          captureAddressFromOrderGet(j);
          await sleep(RATE_LIMIT_STATE.currentGapMs);
        } catch (e) { /* best-effort */ }
      }
    };
    const workers = Array.from(
      { length: Math.max(1, RATE_LIMIT_STATE.currentConcurrency) },
      runWorker
    );
    await Promise.all(workers);
  }
  // `sleep` helper is declared later in the file (const sleep = ...); both
  // declarations share the IIFE scope, so we reuse the existing one via
  // hoisting semantics of `const` — no redeclaration here.

  // CSV export — standardized column order
  function buildAddressBookCsv(records) {
    const cols = ['orderId', 'fulfillUnitId', 'recipientName', 'recipientPhone',
                  'addressDetail', 'subDistrict', 'district', 'province',
                  'zipcode', 'carrier', 'productNames', 'isMasked', 'capturedAt'];
    const esc = v => {
      if (v == null) return '';
      const s = String(v).replace(/"/g, '""');
      return /[",\n]/.test(s) ? `"${s}"` : s;
    };
    const header = cols.join(',');
    const rows = records.map(r => cols.map(c => {
      if (c === 'capturedAt') return r.capturedAt ? new Date(r.capturedAt).toISOString() : '';
      if (c === 'isMasked') return r.isMasked ? '1' : '0';
      return esc(r[c]);
    }).join(','));
    // BOM for Excel Thai compatibility
    return '\uFEFF' + [header, ...rows].join('\n');
  }

  function historyRecentCount(windowMs = 24 * 60 * 60 * 1000) {
    const cutoff = Date.now() - windowMs;
    return loadHistory().filter(e => e.timestamp >= cutoff).length;
  }

  function renderHistoryBadge() {
    const btn = document.getElementById('qf-history-btn');
    if (!btn) return;
    const badge = btn.querySelector('.qf-history-badge');
    if (!badge) return;
    const n = historyRecentCount();
    if (n > 0) {
      badge.textContent = n > 99 ? '99+' : String(n);
      badge.style.display = '';
    } else {
      badge.style.display = 'none';
    }
  }

  function formatHistoryDayHeader(ts) {
    const d = new Date(ts);
    const now = new Date();
    const isSameDay = (a, b) => a.getFullYear() === b.getFullYear()
      && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
    const yesterday = new Date(now);
    yesterday.setDate(now.getDate() - 1);
    if (isSameDay(d, now)) return 'วันนี้';
    if (isSameDay(d, yesterday)) return 'เมื่อวาน';
    const pad = n => String(n).padStart(2, '0');
    return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`;
  }

  function formatHistoryTime(ts) {
    const d = new Date(ts);
    const pad = n => String(n).padStart(2, '0');
    return `${pad(d.getHours())}:${pad(d.getMinutes())}`;
  }

  // Group history entries by day key (YYYY-MM-DD), preserving newest-first order.
  function groupHistoryByDay(entries) {
    const groups = [];
    let currentKey = null;
    let currentEntries = null;
    for (const e of entries) {
      const d = new Date(e.timestamp);
      const key = `${d.getFullYear()}-${d.getMonth()+1}-${d.getDate()}`;
      if (key !== currentKey) {
        currentKey = key;
        currentEntries = [];
        groups.push({ key, header: formatHistoryDayHeader(e.timestamp), entries: currentEntries });
      }
      currentEntries.push(e);
    }
    return groups;
  }

  function openHistoryModal() {
    document.querySelectorAll('.qf-history-modal-overlay').forEach(e => e.remove());
    const overlay = document.createElement('div');
    overlay.className = 'qf-modal-overlay qf-history-modal-overlay';
    overlay.innerHTML = `
      <div class="qf-modal qf-history-modal" role="dialog">
        <div class="qf-history-modal-header">
          <div class="qf-history-modal-title">ประวัติการพิมพ์</div>
          <button class="qf-history-modal-close" aria-label="ปิด">×</button>
        </div>
        <div class="qf-history-search-bar">
          <input class="qf-history-order-input" type="text" placeholder="ค้นหา Order ID (เคลมลูกค้า)…" autocomplete="off" />
          <button class="qf-history-order-search-btn">ค้นหา</button>
        </div>
        <div class="qf-history-order-results" style="display:none;"></div>
        <div class="qf-history-modal-body"></div>
        <div class="qf-history-modal-footer">
          <button class="qf-history-clear-all">ล้างทั้งหมด</button>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);

    const body = overlay.querySelector('.qf-history-modal-body');
    const clearAllBtn = overlay.querySelector('.qf-history-clear-all');

    const cleanup = () => { overlay.remove(); };

    const render = () => {
      const entries = loadHistory();
      if (entries.length === 0) {
        body.innerHTML = `<div class="qf-history-empty">ยังไม่มีประวัติ — เริ่มพิมพ์เพื่อเก็บประวัติอัตโนมัติ</div>`;
        clearAllBtn.style.display = 'none';
        return;
      }
      clearAllBtn.style.display = '';
      const groups = groupHistoryByDay(entries);
      body.innerHTML = groups.map(g => `
        <div class="qf-history-day">
          <div class="qf-history-day-header">${escapeHtml(g.header)}</div>
          ${g.entries.map(e => `
            <div class="qf-history-row" data-id="${escapeHtml(e.id)}">
              <div class="qf-history-row-info">
                <div class="qf-history-row-top">
                  <span class="qf-history-time">${formatHistoryTime(e.timestamp)}</span>
                  <span class="qf-history-title" title="${escapeHtml(e.title)}">${escapeHtml(e.title)}</span>
                </div>
                <div class="qf-history-meta">${e.chunks.length} ไฟล์ · ${e.totalLabels} ใบ</div>
                ${e.workerName ? `<div class="qf-history-packer">แพ็ค: ${escapeHtml(e.workerName)}</div>` : ''}
              </div>
              <div class="qf-history-row-actions">
                <button class="qf-history-redownload" data-id="${escapeHtml(e.id)}">ดาวน์โหลดใหม่</button>
                <button class="qf-history-delete" data-id="${escapeHtml(e.id)}" title="ลบ" aria-label="ลบ">🗑</button>
              </div>
            </div>
          `).join('')}
        </div>
      `).join('');

      body.querySelectorAll('.qf-history-redownload').forEach(btn => {
        btn.addEventListener('click', async (ev) => {
          ev.stopPropagation();
          const id = btn.dataset.id;
          const entry = loadHistory().find(x => x.id === id);
          if (!entry) { showToast('ไม่พบรายการในประวัติ', 1500); return; }
          // Close history modal; new chunked result modal replaces it.
          cleanup();
          const chunks = entry.chunks.map(c => ({
            ids: c.ids,
            label: c.label,
            filename: c.filename,
          }));
          try {
            // Pass null historyMeta to prevent duplicate entries.
            await runChunkedExport(chunks, entry.title, null);
          } catch (err) {
            console.error('[QF] re-download failed:', err);
            showErrorToast('ดาวน์โหลดซ้ำไม่สำเร็จ: ' + (err?.message || err), {
              source: 'historyRedownload',
              error: String(err && (err.stack || err.message || err)),
            });
          }
        });
      });

      body.querySelectorAll('.qf-history-delete').forEach(btn => {
        btn.addEventListener('click', (ev) => {
          ev.stopPropagation();
          deleteHistoryEntry(btn.dataset.id);
          render();
          renderHistoryBadge();
        });
      });
    };

    render();

    // Order search (IndexedDB)
    const orderInput = overlay.querySelector('.qf-history-order-input');
    const orderResultsEl = overlay.querySelector('.qf-history-order-results');
    const doOrderSearch = async () => {
      const q = orderInput.value.trim();
      if (!q) { orderResultsEl.style.display = 'none'; return; }
      orderResultsEl.style.display = '';
      orderResultsEl.innerHTML = '<div class="qf-history-order-loading">กำลังค้นหา…</div>';
      const results = await queryOrderHistory(q);
      if (!results.length) {
        orderResultsEl.innerHTML = `<div class="qf-history-order-empty">ไม่พบประวัติออเดอร์ <b>${escapeHtml(q)}</b></div>`;
        return;
      }
      const pad = n => String(n).padStart(2, '0');
      orderResultsEl.innerHTML = `
        <div class="qf-history-order-header">พบ ${results.length} รายการ สำหรับ <b>${escapeHtml(q)}</b></div>
        ${results.map(r => {
          const d = new Date(r.ts);
          const time = `${pad(d.getHours())}:${pad(d.getMinutes())}`;
          const assignee = r.assigneeKind === 'team'
            ? `ทีม: ${escapeHtml(r.teamName || '')}` + (Array.isArray(r.teamSnapshot) && r.teamSnapshot.length
                ? ` (${r.teamSnapshot.map(m => escapeHtml(m.workerName || m.workerId)).join(', ')})`
                : '')
            : r.workerName ? `คน: ${escapeHtml(r.workerName)}` : 'ไม่ระบุ';
          return `<div class="qf-history-order-row">
            <span class="qf-history-order-date">${escapeHtml(r.date)} ${time}</span>
            <span class="qf-history-order-carrier">${escapeHtml(r.carrier || '—')}</span>
            <span class="qf-history-order-assignee">${assignee}</span>
          </div>`;
        }).join('')}
      `;
    };
    overlay.querySelector('.qf-history-order-search-btn').onclick = (e) => { e.stopPropagation(); doOrderSearch(); };
    orderInput.addEventListener('keydown', (e) => { if (e.key === 'Enter') doOrderSearch(); });

    overlay.querySelector('.qf-history-modal-close').onclick = (e) => {
      e.stopPropagation();
      cleanup();
    };
    clearAllBtn.onclick = async (e) => {
      e.stopPropagation();
      const ok = await confirmInline('ล้างประวัติทั้งหมด?', 'ล้าง', true);
      if (ok) {
        clearAllHistory();
        renderHistoryBadge();
        render();
      }
    };
    overlay.onclick = (e) => { if (e.target === overlay) cleanup(); };
    const onKey = (e) => { if (e.key === 'Escape') { cleanup(); document.removeEventListener('keydown', onKey); } };
    document.addEventListener('keydown', onKey);
  }

  // ==================== ADDRESS BOOK MODAL ====================
  function renderAddressBookBadge(count) {
    const btn = document.getElementById('qf-addrbook-btn');
    if (!btn) return;
    const badge = btn.querySelector('.qf-addrbook-badge');
    if (!badge) return;
    if (count > 0) {
      badge.textContent = count > 99 ? '99+' : String(count);
      badge.style.display = '';
    } else {
      badge.style.display = 'none';
    }
  }

  async function refreshAddressBookBadge() {
    try {
      const list = await getAllAddresses();
      const clean = list.filter(r => !r.isMasked).length;
      renderAddressBookBadge(clean);
    } catch (e) {}
  }

  async function openAddressBookModal() {
    document.querySelectorAll('.qf-addrbook-modal-overlay').forEach(e => e.remove());
    const overlay = document.createElement('div');
    overlay.className = 'qf-modal-overlay qf-addrbook-modal-overlay';
    overlay.innerHTML = `
      <div class="qf-modal qf-history-modal" role="dialog">
        <div class="qf-history-modal-header">
          <div class="qf-history-modal-title">📇 สมุดที่อยู่ (จับก่อนพิมพ์)</div>
          <button class="qf-history-modal-close" aria-label="ปิด">×</button>
        </div>
        <div class="qf-addrbook-toolbar">
          <label class="qf-addrbook-toggle">
            <input type="checkbox" id="qf-addrbook-enable" ${state.addressCaptureEnabled ? 'checked' : ''} />
            <span>เปิดการจับที่อยู่อัตโนมัติก่อนพิมพ์ครั้งแรก</span>
          </label>
          <div class="qf-addrbook-stats" id="qf-addrbook-stats">กำลังโหลด…</div>
        </div>
        <div class="qf-history-search-bar">
          <input id="qf-addrbook-search" class="qf-history-order-input" type="text"
                 placeholder="ค้นหา ชื่อ / เบอร์ / Order ID / จังหวัด…" autocomplete="off" />
        </div>
        <div class="qf-history-modal-body" id="qf-addrbook-body"></div>
        <div class="qf-history-modal-footer">
          <button id="qf-addrbook-export" class="qf-btn-confirm">📥 ดาวน์โหลด CSV</button>
          <button id="qf-addrbook-clear" class="qf-history-clear-all">ล้างทั้งหมด</button>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);

    const body = overlay.querySelector('#qf-addrbook-body');
    const stats = overlay.querySelector('#qf-addrbook-stats');
    const search = overlay.querySelector('#qf-addrbook-search');
    const exportBtn = overlay.querySelector('#qf-addrbook-export');
    const clearBtn = overlay.querySelector('#qf-addrbook-clear');
    const toggle = overlay.querySelector('#qf-addrbook-enable');

    const cleanup = () => overlay.remove();

    const render = async () => {
      const all = await getAllAddresses();
      all.sort((a, b) => (b.capturedAt || 0) - (a.capturedAt || 0));
      const clean = all.filter(r => !r.isMasked).length;
      const masked = all.length - clean;
      stats.textContent = `รวม ${all.length} ราย · ไม่ถูก mask ${clean} · ถูก mask ${masked}`;

      const q = (search.value || '').trim().toLowerCase();
      const filtered = q ? all.filter(r => {
        const hay = [r.orderId, r.recipientName, r.recipientPhone,
                     r.province, r.district, r.addressDetail].filter(Boolean).join(' ').toLowerCase();
        return hay.includes(q);
      }) : all;

      if (!filtered.length) {
        body.innerHTML = `<div class="qf-history-empty">
          ${all.length === 0
            ? 'ยังไม่มีที่อยู่ที่จับได้ — ระบบจะเริ่มเก็บเมื่อคุณสแกนแล้วเปิด order detail หรือกดพิมพ์ครั้งแรก'
            : `ไม่พบรายการที่ตรงกับ "${q}"`}
        </div>`;
        return;
      }

      const esc = s => String(s ?? '').replace(/[&<>"]/g, c =>
        ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
      const rows = filtered.map(r => {
        const addr = [r.addressDetail, r.subDistrict, r.district, r.province, r.zipcode]
          .filter(Boolean).join(' · ');
        const mask = r.isMasked ? '<span class="qf-addrbook-mask">🔒 masked</span>' : '';
        return `
          <div class="qf-addrbook-row">
            <div class="qf-addrbook-row-head">
              <span class="qf-addrbook-name">${esc(r.recipientName || '—')}</span>
              <span class="qf-addrbook-phone">${esc(r.recipientPhone || '—')}</span>
              ${mask}
            </div>
            <div class="qf-addrbook-row-meta">
              <span class="qf-addrbook-oid">${esc(r.orderId)}</span>
              ${r.carrier ? `<span class="qf-addrbook-carrier">${esc(r.carrier)}</span>` : ''}
              <span class="qf-addrbook-time">${new Date(r.capturedAt).toLocaleString('th-TH')}</span>
            </div>
            <div class="qf-addrbook-row-addr">${esc(addr || '—')}</div>
          </div>`;
      }).join('');
      body.innerHTML = rows;
    };

    overlay.querySelector('.qf-history-modal-close').onclick = cleanup;
    overlay.onclick = (e) => { if (e.target === overlay) cleanup(); };
    document.addEventListener('keydown', function onKey(e) {
      if (e.key === 'Escape') { cleanup(); document.removeEventListener('keydown', onKey); }
    });

    search.addEventListener('input', () => render());
    toggle.addEventListener('change', () => {
      state.addressCaptureEnabled = toggle.checked;
      saveAddressCapturePref(toggle.checked);
      showToast(toggle.checked ? '✅ เปิดการจับที่อยู่' : '⏸ ปิดการจับที่อยู่');
    });

    exportBtn.addEventListener('click', async () => {
      const list = await getAllAddresses();
      if (!list.length) { showToast('ยังไม่มีข้อมูล'); return; }
      const csv = buildAddressBookCsv(list);
      const ts = new Date().toISOString().slice(0, 19).replace(/[:T]/g, '-');
      downloadCsv(csv, `address-book-${ts}.csv`);
      showToast(`ดาวน์โหลด ${list.length} รายการ`);
    });

    clearBtn.addEventListener('click', async () => {
      const ok = await confirmInline('ล้างสมุดที่อยู่ทั้งหมด?', 'ล้าง', true);
      if (!ok) return;
      await clearAddressBook();
      await refreshAddressBookBadge();
      render();
    });

    render();
  }

  // Small inline confirm used throughout the app. isDanger=true for destructive actions.
  function confirmInline(title, confirmLabel, isDanger = false) {
    return new Promise(resolve => {
      const overlay = document.createElement('div');
      overlay.className = 'qf-modal-overlay';
      const btnCls = isDanger ? 'qf-btn-confirm qf-btn-danger' : 'qf-btn-confirm';
      overlay.innerHTML = `
        <div class="qf-modal" role="dialog">
          <div class="qf-modal-title">${escapeHtml(title)}</div>
          <div class="qf-modal-actions">
            <button class="qf-btn-cancel">ยกเลิก</button>
            <button class="${btnCls}">${escapeHtml(confirmLabel || 'ยืนยัน')}</button>
          </div>
        </div>
      `;
      document.body.appendChild(overlay);
      const cleanup = v => { overlay.remove(); resolve(v); };
      overlay.querySelector('.qf-btn-cancel').onclick = (e) => { e.stopPropagation(); cleanup(false); };
      overlay.querySelector('.qf-btn-confirm').onclick = (e) => { e.stopPropagation(); cleanup(true); };
      overlay.onclick = e => { if (e.target === overlay) cleanup(false); };
      const onKey = e => { if (e.key === 'Escape') { cleanup(false); document.removeEventListener('keydown', onKey); } };
      document.addEventListener('keydown', onKey);
    });
  }

  // Platform-prefixed alias keys keep TikTok and Shopee aliases separate
  // when the same numeric productId/skuId collides across the two platforms.
  // Migration: existing (unprefixed) keys are read as TikTok-side legacy;
  // all new writes are platform-prefixed. Reads on TikTok check both the
  // prefixed form and the unprefixed legacy form, so old TikTok aliases
  // survive. Reads on Shopee only check the prefixed form.
  function aliasKey(productId) {
    const p = isShopee() ? 'sp' : 'tk';
    return `${p}:${productId}`;
  }
  function variantKey(productId, skuId) {
    const p = isShopee() ? 'sp' : 'tk';
    return `${p}:${productId}:${skuId}`;
  }
  function getAlias(productId) {
    // Prefixed key wins; fall back to unprefixed legacy on TikTok only.
    const prefixed = state.aliases.get(aliasKey(productId));
    if (prefixed != null) return prefixed;
    if (!isShopee()) return state.aliases.get(String(productId));
    return undefined;
  }
  function setAlias(productId, value) {
    const k = aliasKey(productId);
    if (value) state.aliases.set(k, value);
    else state.aliases.delete(k);
    // Also clear legacy unprefixed entry on TikTok so it doesn't shadow
    // the cleared prefixed one on future reads.
    if (!isShopee()) state.aliases.delete(String(productId));
    saveAliases();
  }
  function getVariantInfo(productId, skuId) {
    const prefixed = state.variantAliases.get(variantKey(productId, skuId));
    if (prefixed) return prefixed;
    if (!isShopee()) {
      // Legacy format: `${productId}:${skuId}` without platform prefix.
      return state.variantAliases.get(`${productId}:${skuId}`) || null;
    }
    return null;
  }
  function setVariantInfo(productId, skuId, partial) {
    const key = variantKey(productId, skuId);
    const cur = state.variantAliases.get(key)
              || (!isShopee() ? state.variantAliases.get(`${productId}:${skuId}`) : null)
              || {alias: '', replace: false};
    const next = {...cur, ...partial};
    if (!next.alias?.trim() && !next.replace) state.variantAliases.delete(key);
    else state.variantAliases.set(key, next);
    // Clear legacy unprefixed entry on TikTok so new prefixed write wins.
    if (!isShopee()) state.variantAliases.delete(`${productId}:${skuId}`);
    saveVariantAliases();
  }

  // Listen for messages from asset-bridge (ISOLATED world) — font URL + manifest version.
  state.localVersion = null;
  window.addEventListener('message', (e) => {
    if (e.source !== window) return;
    const d = e.data;
    if (!d?.__qfAsset) return;
    if (d.__qfAsset === 'font' && d.url) state.fontUrl = d.url;
    if (d.__qfAsset === 'samplePreview' && d.url) state.samplePreviewUrl = d.url;
    if (d.__qfAsset === 'carrierLogo' && d.carrier && d.url) {
      state.carrierLogoUrls = { ...state.carrierLogoUrls, [d.carrier]: d.url };
    }
    if (d.__qfAsset === 'manifest' && d.version) state.localVersion = d.version;
  });
  window.postMessage({ __qfAsset: 'request_font' }, '*');
  // Expose state so the top-level fetch hook can write apiListUrl into it
  window.__qfState = state;
  // Expose rate-limit state + probe helper for DevTools diagnostics.
  // Usage examples (copy into DevTools console):
  //   window.__qfRateLimitState()  // live view of backoff / hits
  //   window.__qfProbeContactRateLimit('583681301773518591', { burst: 20, gapMs: 100 })
  //     → fires N calls as fast as the gap allows, reports the first hit's
  //        timestamp so you can compute the observed threshold.
  window.__qfRateLimitState = () => ({ ...RATE_LIMIT_STATE, backfillCount: _addressBackfillCount });
  window.__qfBuyerContactTemplate = () => ({
    url: _buyerContactUrl,
    body: _buyerContactBodyTemplate ? { ..._buyerContactBodyTemplate } : null,
  });
  // Debug visibility for note backfill — exposes what we've captured so far
  // and lets us trigger a single backfill call manually for diagnostics.
  window.__qfOrderGetTemplate = () => ({
    url: _orderGetUrl,
    body: _orderGetBodyTemplate ? { ..._orderGetBodyTemplate } : null,
    derivedFromLabels: deriveOrderGetUrl(),
    labelsApiUrl: _labelsApiUrl,
  });
  window.__qfTryOrderGet = async (orderId) => {
    const url = deriveOrderGetUrl();
    if (!url) return { err: 'no labels url to derive from' };
    const body = { ...(_orderGetBodyTemplate || {}), order_id: String(orderId), order_id_list: [String(orderId)] };
    try {
      const r = await _origFetch.call(window, url, {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify(body),
      });
      const j = await r.json();
      return { code: j.code, msg: j.message, hasData: !!j.data, dataKeys: j.data ? Object.keys(j.data) : null, sample: j };
    } catch (e) { return { err: e.message }; }
  };
  window.__qfProbeParallelContact = async (orderId, { burst = 20, type = 3 } = {}) => {
    if (!_buyerContactUrl || !_buyerContactBodyTemplate) {
      console.warn('[QF] no template — click 👁 first');
      return null;
    }
    const mk = (i) => {
      const body = { ..._buyerContactBodyTemplate, main_order_id: String(orderId), contact_info_type: type };
      const t0 = Date.now();
      return _origFetch.call(window, _buyerContactUrl, {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify(body),
      })
        .then(r => r.json())
        .then(j => ({ i, t: Date.now() - t0, code: j.code, msg: (j.message || '').slice(0, 50), limited: isRateLimitResponse(j) }))
        .catch(e => ({ i, error: String(e).slice(0, 80) }));
    };
    const tStart = Date.now();
    const results = await Promise.all(Array.from({ length: burst }, (_, i) => mk(i)));
    const elapsed = Date.now() - tStart;
    return { burst, elapsed, results };
  };
  window.__qfProbeContactRateLimit = async (orderId, { burst = 20, gapMs = 100 } = {}) => {
    if (!_buyerContactUrl || !_buyerContactBodyTemplate) {
      console.warn('[QF] probe: buyer_contact_info template not captured yet — click one 👁 eye icon on any order detail page first');
      return null;
    }
    const results = [];
    const t0 = Date.now();
    for (let i = 0; i < burst; i++) {
      const callAt = Date.now() - t0;
      try {
        const body = { ...(_buyerContactBodyTemplate || {}), main_order_id: String(orderId), contact_info_type: CONTACT_TYPE_NICKNAME };
        const r = await _origFetch.call(window, _buyerContactUrl, {
          method: 'POST',
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify(body),
        });
        const j = await r.json();
        const limited = isRateLimitResponse(j);
        results.push({ i, tMs: callAt, code: j.code, msg: j.message, limited });
        if (limited) {
          console.warn(`[QF] probe: FIRST rate-limit hit at call #${i} (t=${callAt}ms, ~${Math.round(i / (callAt / 60000))}/min)`);
          break;
        }
      } catch (e) {
        results.push({ i, tMs: callAt, error: String(e) });
      }
      if (gapMs > 0) await sleep(gapMs);
    }
    console.table(results);
    return results;
  };

  // ==================== PAGE DETECTION ====================
  const isShopee = () => location.hostname === 'seller.shopee.co.th';
  const isTikTok = () => location.hostname === 'seller-th.tiktok.com';
  // Shopee: every /portal/sale* page is scan-able via the same XHR-captured
  // order/card APIs, so we treat the whole section as "labels" for routing
  // purposes (same widget UI, same scanShopeePage flow). We also surface
  // /portal/sale/order* as an order page so UI paths that branch on
  // isOrderPage() don't exclude Shopee sellers.
  const isLabelsPage = () => isTikTok() && /\/shipment\/labels/.test(location.pathname)
                            || isShopee() && /\/portal\/sale/.test(location.pathname);
  const isOrderPage  = () => isTikTok() && /\/order/.test(location.pathname)
                            || isShopee() && /\/portal\/sale\/order/.test(location.pathname);

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

  // Labels-page done: session-local only — a product/variant/combo is "done"
  // iff every visible fulfillUnitId was printed through THIS extension session
  // (tracked in state.printedUnitIds, which clears on each scan).
  //
  // Server-side labelStatus=50 is intentionally NOT treated as done here: the
  // user may re-scan previously printed labels specifically TO reprint them
  // (replacement after lost/damaged labels), so cards must remain clickable.
  // The only signal that truly means "user just printed this" is the extension
  // having added the id to printedUnitIds.
  //
  // Under "พิมพ์แล้ว" filter every visible item is already printed server-side;
  // marking them done would defeat the filter's purpose (reprint view).
  function isLabelsDone(idSet) {
    // Under 'printed' / 'failed' the user is intentionally viewing records for
    // reprint — marking them done-greyed would defeat the tab's purpose.
    if (state.labelStatusFilter === 'printed' || state.labelStatusFilter === 'failed') return false;
    const visible = [...idSet].filter(id => passesCarrier(id) && passesPreOrder(id));
    if (!visible.length) return false;
    return visible.every(id => state.printedUnitIds.has(id));
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

  // Error toast with a "คัดลอก debug" button so the user can paste diagnostic
  // JSON when reporting an issue. Stays up longer than a normal toast.
  function showErrorToast(message, details = {}, duration = 8000) {
    document.querySelectorAll('.qf-toast').forEach(e => e.remove());
    const t = document.createElement('div');
    t.className = 'qf-toast qf-toast-error';

    const msgEl = document.createElement('span');
    msgEl.className = 'qf-toast-msg';
    msgEl.textContent = message;
    t.appendChild(msgEl);

    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'qf-toast-copy';
    btn.textContent = 'คัดลอก debug';
    btn.onclick = async (e) => {
      e.stopPropagation();
      const payload = {
        message,
        details,
        url: location.href,
        userAgent: navigator.userAgent,
        ts: new Date().toISOString(),
      };
      const text = JSON.stringify(payload, null, 2);
      try {
        await navigator.clipboard.writeText(text);
        btn.textContent = 'คัดลอกแล้ว ✓';
      } catch {
        // Fallback for sites that deny clipboard permission
        const ta = document.createElement('textarea');
        ta.value = text;
        ta.style.position = 'fixed';
        ta.style.opacity = '0';
        document.body.appendChild(ta);
        ta.select();
        try { document.execCommand('copy'); btn.textContent = 'คัดลอกแล้ว ✓'; }
        catch { btn.textContent = 'คัดลอกไม่ได้'; }
        document.body.removeChild(ta);
      }
      setTimeout(() => { if (btn.isConnected) btn.textContent = 'คัดลอก debug'; }, 2000);
    };
    t.appendChild(btn);

    document.body.appendChild(t);
    setTimeout(() => { if (t.isConnected) t.remove(); }, duration);
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

  // If TikTok rejects our request because the body schema changed (missing
  // field, invalid param, etc.), a stale template will fail every subsequent
  // call forever until the user refreshes. Detect those message shapes so we
  // can reset and re-learn on the next scan.
  function isTemplateSchemaError(resp) {
    if (!resp || resp.code === 0) return false;
    const msg = String(resp.message || resp.msg || '').toLowerCase();
    if (!msg) return false;
    return /missing|required|invalid[_\s-]?param|invalid[_\s-]?argument|unknown field|unexpected field|bad request|param error|参数/.test(msg);
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
    const data = await resp.json();
    if (isTemplateSchemaError(data)) {
      // Template is stale — wipe it so the fetch hook can capture a fresh one
      // on the next user-triggered request. Show a toast so the user knows
      // exactly what to do.
      _apiListUrl = null;
      _apiListBodyTemplate = null;
      showToast('API schema เปลี่ยน — กดสแกนใหม่เพื่อจับใหม่', 4000);
    }
    return data;
  }

  function extractImageUrl(img) {
    if (!img) return '';
    if (typeof img === 'string') return img;
    // API returns { url_list: [...], thumb_url_list: [...], uri, ... }
    // Prefer url_list (full-res) so PDF dividers aren't pixelated; the grid
    // UI caches the same URL so one fetch serves both surfaces.
    return img.url_list?.[0] || img.thumb_url_list?.[0] || '';
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
    // Warn user if an active planning session exists — scan would invalidate its IDs
    if (loadPlanningSession()) {
      const ok = await confirmInline('มีแผนค้างอยู่ — ละทิ้งแผนและ scan ใหม่?', 'ละทิ้ง', true);
      if (!ok) return;
      deletePlanningSession();
      document.getElementById('qf-plan-recovery')?.remove();
    }
    state.scanning = true;
    state.products.clear();
    state.weirdOrders = [];
    // Keep non-expired doneItems across scans so "ที่พิมพ์แล้ว" markers survive
    // both page refresh and rescan. Only drop entries that have already
    // exceeded DONE_TIMEOUT_MS (30 min) — they would be dropped lazily by
    // isDone/isComboDone anyway, but a centralized sweep keeps the map small.
    {
      const nowTs = Date.now();
      let dropped = false;
      for (const [k, ts] of state.doneItems) {
        if (typeof ts !== 'number' || nowTs - ts > DONE_TIMEOUT_MS) {
          state.doneItems.delete(k);
          dropped = true;
        }
      }
      if (dropped) saveDoneItems();
    }
    state.records.clear();
    state.weirdFulfillUnitIds.clear();
    state.weirdCombos.clear();
    state.carriers.clear();
    state.carrierOf.clear();
    state.preOrderOf.clear();
    // Allow one fresh Shopee pre-order probe dump per scan
    _shopeePreOrderProbeLogged = false;
    _noteProbeLogged = false;

    const statusEl = document.getElementById('qf-scan-status');
    const btn = document.getElementById('qf-scan-btn');
    btn.disabled = true;

    if (isLabelsPage()) {
      state.printedUnitIds.clear();
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
      if (!url) {
        showErrorToast('ไม่สามารถจับ API URL — ลองโหลดหน้าใหม่แล้วสแกนอีกครั้ง', {
          source: 'scanAllPages:ensureApiUrl',
          apiListUrl: _apiListUrl,
          hasTemplate: !!_apiListBodyTemplate,
        });
        throw new Error('ไม่สามารถจับ API URL — ลองโหลดหน้าใหม่แล้วสแกนอีกครั้ง');
      }

      // First batch to learn total_count
      statusEl.textContent = 'กำลังดึงออเดอร์...';
      const hadTemplate = !!_apiListBodyTemplate;
      const first = await fetchOrderBatch(0);
      if (first.code !== 0) {
        const tplStatus = _apiListBodyTemplate
          ? `template OK (${Object.keys(_apiListBodyTemplate).length} fields)`
          : 'NO TEMPLATE — body fallback ขาดข้อมูล';
        console.error('[QF] order/list failed:', first, 'template:', _apiListBodyTemplate);
        // If we did send a captured template and it was still rejected, the
        // schema has likely changed. Drop the template so the next scan
        // re-learns from TikTok's own request. (fetchOrderBatch only resets
        // when the message matches known schema-error shapes; here we reset
        // defensively on any non-zero code when a template was in use.)
        if (hadTemplate) {
          _apiListUrl = null;
          _apiListBodyTemplate = null;
        }
        throw new Error(`code=${first.code} msg="${first.message || 'empty'}" — ${tplStatus}`);
      }

      const d = first.data || {};
      const total = d.total_count ?? d.total ?? 0;
      const firstOrders = d.main_orders || d.order_list || d.orders || [];
      if (total === 0 && firstOrders.length === 0) {
        showErrorToast('API ตอบกลับว่าง — ลองโหลดหน้าใหม่', {
          source: 'scanAllPages:emptyResponse',
          responseKeys: Object.keys(d),
          apiListUrl: _apiListUrl,
          hasTemplate: !!_apiListBodyTemplate,
        });
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
    if (state.labelStatusFilter === 'failed')      return labelStatus === LABEL_STATUS_PRINT_FAILED;
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
        // §5.1: Per-quantity buckets for multi-qty orders (qty>=2, single SKU).
        // Map<quantity: number, Set<fulfillUnitId>>
        fulfillUnitIdsByQty: new Map(),
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
        // §5.1: Per-quantity buckets (same semantics as product-level).
        fulfillUnitIdsByQty: new Map(),
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
        // §Picking list needs rec.orderIds to count unique orders — it was
        // previously dropped here so totalOrders always read 0.
        orderIds: Array.isArray(rec.orderIds) ? rec.orderIds.filter(Boolean) : [],
        labelStatus: rec.labelStatus || null,
        createTime: rec.createTime || null,
        shipByTime: rec.shipByTime || null,
        autoCancelTime: rec.autoCancelTime || null,
        hasNote: !!rec.hasNote,
        hasBuyerNote: !!(rec.hasBuyerNote || rec.buyerNote),
        hasSellerNote: !!(rec.hasSellerNote || rec.sellerNote),
        buyerNote: rec.buyerNote || '',
        sellerNote: rec.sellerNote || '',
        skuList: skus.map(s => ({
          productId: s.productId,
          skuId: s.skuId,
          productName: s.productName,
          skuName: s.skuName,
          sellerSkuName: s.sellerSkuName,
          productImageURL: s.productImageURL || '',
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
          // §5.1: Also track per-quantity bucket for qty-variant split feature.
          const q = s.quantity;
          if (!product.fulfillUnitIdsByQty.has(q)) product.fulfillUnitIdsByQty.set(q, new Set());
          product.fulfillUnitIdsByQty.get(q).add(fulfillUnitId);
          if (!variant.fulfillUnitIdsByQty.has(q)) variant.fulfillUnitIdsByQty.set(q, new Set());
          variant.fulfillUnitIdsByQty.get(q).add(fulfillUnitId);
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
    return items.map(it => {
      const modelId = it.inner_item_ext_info?.model_id;
      return {
        productId: String(it.inner_item_ext_info?.item_id || ''),
        // If Shopee item has no model_id (no variants), keep skuId null
        // rather than falling back to item_id. Reusing item_id made
        // productId === skuId and caused variant-level alias/done checks
        // to collide with product-level. Downstream sites already tolerate
        // a null skuId via doneKey()/variantKey() fallbacks.
        skuId: modelId ? String(modelId) : null,
        productName: it.name || '',
        skuName: it.model_name || it.variation_name || '',
        sellerSkuName: it.item_sku || '',
        productImageURL: it.image
          ? `https://down-th.img.susercontent.com/file/${it.image}_tn`
          : '',
        quantity: it.amount || 1,
      };
    }).filter(s => s.productId);
  }

  function pushShopeeRecord({fulfillUnitId, ext, items, fulfilment, batchId, pkgExt}) {
    if (!fulfillUnitId) return;
    const skuList = buildShopeeSkuList(items);
    if (!skuList.length) return;
    // Map Shopee's logistics_status → TikTok-style label status.
    // Per CLAUDE.md: logistics_status >= 3 means label already printed/shipped.
    // `logistics_status` lives on order_ext_info per CLAUDE.md, but fall back to
    // pkg-level + fulfilment_info in case Shopee moves it.
    const logisticsStatus = Number(
      ext?.logistics_status
      ?? pkgExt?.logistics_status
      ?? fulfilment?.logistics_status
      ?? 0
    );
    const labelStatus = logisticsStatus >= 3 ? LABEL_STATUS_PRINTED : LABEL_STATUS_NOT_PRINTED;

    // Pre-order detection. CLAUDE.md lists 3 possible locations; try each.
    const firstItem = items?.[0] || {};
    let isPreOrder = !!(
      ext?.is_pre_order
      || pkgExt?.is_pre_order
      || fulfilment?.is_pre_order
      || firstItem.is_pre_order
      || firstItem.inner_item_ext_info?.is_pre_order
      || firstItem.item_ext_info?.is_pre_order
    );

    // Carrier icon — Shopee's exact field name not yet pinned. Try the most
    // common shapes; falls back to '' so existing UI doesn't break.
    const sp = fulfilment || {};
    const carrierIconUrl = sp.icon_url
      || sp.logistics_image
      || sp.channel_logo
      || sp.image_url
      || sp.fulfilment_channel_logo
      || sp.fulfilment_channel_image
      || '';
    const carrierName = sp.fulfilment_channel_name || sp.masked_channel_name || 'ไม่ระบุ';
    const carrierId = String(sp.fulfilment_channel_id || sp.fulfilment_channel_name || ext.masked_channel_id || 'unknown');

    // Buyer note — defensive walk of common Shopee fields.
    const buyerNote = String(
      ext?.buyer_remark
      || pkgExt?.buyer_remark
      || ext?.note_to_seller
      || ext?.buyer_message
      || firstItem.buyer_remark
      || ''
    ).trim();

    // First-record discovery dump — runs once per scan when ANY of these are
    // missing (pre-order field, carrier icon, note). Reset in scanAllPages.
    if (!_shopeePreOrderProbeLogged && (ext || pkgExt || fulfilment)) {
      _shopeePreOrderProbeLogged = true;
      try {
        const probe = {
          ext_keys: Object.keys(ext || {}),
          pkgExt_keys: Object.keys(pkgExt || {}),
          fulfilment_keys: Object.keys(fulfilment || {}),
          item_keys: Object.keys(firstItem),
          inner_item_keys: Object.keys(firstItem.inner_item_ext_info || {}),
          // Detected values (or empty if missing)
          detected: { isPreOrder, carrierName, carrierIconUrl, buyerNote: buyerNote.slice(0, 50) },
          // Full objects so we can grep for the right field name visually
          fulfilment_full: fulfilment,
          ext_sample: ext,
        };
        console.log('[QF Shopee record probe] First scanned record →', probe);
        console.log('[QF Shopee record probe] To inspect more: window.__qfShopeeProbes()');
      } catch {}
    }

    processLabelRecord({
      fulfillUnitId: String(fulfillUnitId),
      batchId: String(batchId || fulfillUnitId),
      orderIds: [String(ext.order_id || '')],
      labelStatus,
      isPreOrder,
      skuList,
      shippingProviderInfo: { name: carrierName, iconUrl: carrierIconUrl },
      deliveryInfo: { shippingProvider: { id: carrierId, name: carrierName, icon_url: carrierIconUrl } },
      buyerNote,
      hasBuyerNote: !!buyerNote,
      hasNote: !!buyerNote,
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
        pkgExt: pkg,
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
        pkgExt: pkg,
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
          pkgExt: pkgInfo,
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

  // Best-effort: detect which Shopee tab the user is looking at so we don't
  // silently fall back to the to-ship body when they're on "ทั้งหมด".
  // Returns one of 'toship' | 'all' | null. Shopee's tab markup changes
  // often, so we hedge with several signals; null means "can't tell".
  function detectShopeeActiveTab() {
    // 1. Active tab element — try common shadcn-ish / antd-style selectors
    const activeSelectors = [
      '[role="tab"][aria-selected="true"]',
      '.shopee-tabs__item--active',
      '.shopee-tabs-tab--active',
      '.eds-tab--active',
      '.eds-tabs__tab--active',
      '.shopee-tab--active',
    ];
    let activeText = '';
    for (const sel of activeSelectors) {
      const el = document.querySelector(sel);
      if (el && el.textContent) { activeText = el.textContent.trim(); break; }
    }
    if (!activeText) return null;
    // 2. Match by visible label (Thai)
    if (/ทั้งหมด/.test(activeText)) return 'all';
    if (/ที่ต้องจัดส่ง|ยังไม่จัดส่ง|รอจัดส่ง|เตรียมจัดส่ง/.test(activeText)) return 'toship';
    return null;
  }

  async function scanShopeePage(statusEl) {
    // Re-capture so body reflects the user's current Shopee tab/filter view.
    resetShopeeCapture();
    statusEl.textContent = 'รอจับ API ตามแท็บที่กำลังดู...';
    await triggerShopeeApiCapture(statusEl);

    // If still nothing captured, only fall back to the hardcoded to-ship
    // body when DOM detection confirms the user is actually on that tab.
    // Otherwise we'd silently feed them wrong-tab data (e.g. user on
    // "ทั้งหมด" would get filtered to to-ship only).
    if (!_shopeeIndexUrl || !_shopeeIndexBody) {
      const tab = detectShopeeActiveTab();
      if (tab === 'toship') {
        _shopeeIndexUrl = SHOPEE_TOSHIP_DEFAULT.indexUrl;
        _shopeeIndexBody = SHOPEE_TOSHIP_DEFAULT.indexBody;
        statusEl.textContent = 'ใช้ฟิลเตอร์เริ่มต้น (ที่ต้องจัดส่ง)...';
      } else {
        // Can't confirm tab — bail with a clear message rather than risk
        // showing data from the wrong tab.
        throw new Error('ยังไม่ได้จับ API — กรุณาคลิกเปลี่ยนหน้า/แท็บหนึ่งครั้ง แล้วกด Scan ใหม่');
      }
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
      productImageURL: s.product_image?.url_list?.[0] || s.product_image?.thumb_url_list?.[0] || '',
      quantity: s.quantity || 1,
    }));
    // isPreOrder: any line marked → record is pre-order
    const olm = rec.order_label_module || [];
    const isPreOrder = olm.some(o => o.isPreOrder === 1 || o.is_pre_order === 1);
    // Note capture — separate buyer vs seller so they never collide on the
    // label or in the note-zone divider. TikTok's API uses different field
    // names across versions and sometimes tucks notes inside nested modules,
    // so we walk a wide net. First non-empty wins per side.
    const tom = rec.trade_order_module?.[0] || {};
    const oem = rec.order_extra_module?.[0] || rec.order_extra_module || {};
    const orm = rec.order_remark_module?.[0] || rec.order_remark_module || {};
    const pickStr = (...vals) => {
      for (const v of vals) {
        if (typeof v === 'string' && v.trim()) return v.trim();
      }
      return '';
    };
    // Verified by capturing the live labels API: presence of notes is signaled
    // by `rec.note_module.has_buyer_note` / `has_seller_note` (booleans). The
    // labels list response intentionally does NOT include the note text — it
    // must be fetched on-demand from the order detail endpoint (see
    // `fetchOrderNotes` below). Keep legacy field-path probes as a safety net
    // in case some shop regions still return inline text.
    const noteModule = rec.note_module || {};
    const hasBuyerFlag  = !!noteModule.has_buyer_note;
    const hasSellerFlag = !!noteModule.has_seller_note;

    let buyerNote  = pickStr(
      tom.buyer_remark?.note, tom.buyer_remark?.message, tom.buyer_remark?.text,
      rec.buyer_remark?.note, rec.buyer_remark?.message, rec.buyer_remark?.text,
      tom.buyer_message, tom.buyer_note,
      rec.buyer_message, rec.buyer_note, rec.order_note,
    );
    let sellerNote = pickStr(
      tom.seller_remark?.note, tom.seller_remark?.message, tom.seller_remark?.text,
      rec.seller_remark?.note, rec.seller_remark?.message, rec.seller_remark?.text,
      tom.seller_note, tom.seller_message, tom.seller_mark,
      rec.seller_note, rec.seller_message,
    );

    // Booleans drive the ★ badge on the label even when text isn't loaded yet.
    const hasBuyerNote  = hasBuyerFlag  || !!buyerNote;
    const hasSellerNote = hasSellerFlag || !!sellerNote;

    // §Debug: when scanning the very first record of a session that has any
    // note-shaped field, dump it so we can spot where TikTok actually puts
    // the data when the field paths above miss it. One-shot guard prevents
    // log spam. Read window.__qfNoteProbe in DevTools.
    if (!_noteProbeLogged && (buyerNote || sellerNote || tom.buyer_remark || tom.seller_remark || rec.buyer_remark || rec.seller_remark)) {
      _noteProbeLogged = true;
      try {
        window.__qfNoteProbe = {
          tom, oem, orm,
          recBuyerRemark: rec.buyer_remark, recSellerRemark: rec.seller_remark,
          tomBuyerRemark: tom.buyer_remark, tomSellerRemark: tom.seller_remark,
          recKeys: Object.keys(rec),
          tomKeys: Object.keys(tom),
        };
        console.log('[QF note-probe] tom keys:', Object.keys(tom));
        console.log('[QF note-probe] tom.buyer_remark:', tom.buyer_remark);
        console.log('[QF note-probe] tom.seller_remark:', tom.seller_remark);
        console.log('[QF note-probe] rec.buyer_remark:', rec.buyer_remark);
        console.log('[QF note-probe] rec.seller_remark:', rec.seller_remark);
        console.log('[QF note-probe] resolved:', { buyerNote, sellerNote });
      } catch {}
    }
    return {
      fulfillUnitId: rec.fulfill_unit_id || lm.fulfill_unit_id,
      batchId: lm.batch_id,
      // Order IDs live in `order_module[]` on the labels API — earlier we only
      // looked at `label_module.order_ids` / root which were always undefined,
      // so picking-list "Orders: 0" and order-claim lookup both failed.
      orderIds: (lm.order_ids && lm.order_ids.length ? lm.order_ids : null)
             || (rec.order_ids && rec.order_ids.length ? rec.order_ids : null)
             || ((rec.order_module || []).map(o => o.main_order_id).filter(Boolean)),
      labelStatus: lm.label_status,
      isPreOrder,
      // hasNote: legacy flag — any note present. Kept so old call sites
      // (filters, routing) don't break. Prefer hasBuyerNote/hasSellerNote.
      // The note_module flags from the labels API drive the booleans even
      // when text isn't loaded yet (we lazy-fetch text from order detail).
      hasNote: hasBuyerNote || hasSellerNote,
      hasBuyerNote,
      hasSellerNote,
      buyerNote,
      sellerNote,
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
      if (state.labelStatusFilter === 'not_printed') cl.fulfillment_label_status = { value: [String(LABEL_STATUS_NOT_PRINTED)] };
      else if (state.labelStatusFilter === 'printed') cl.fulfillment_label_status = { value: [String(LABEL_STATUS_PRINTED)] };
      else if (state.labelStatusFilter === 'failed')  cl.fulfillment_label_status = { value: [String(LABEL_STATUS_PRINT_FAILED)] };
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
  // §3.1: Chunk-plan modal thresholds
  const CHUNK_PROMPT_THRESHOLD = 200; // show modal when total > this
  const CHUNK_AUTO_SAFE_SIZE   = 200; // default X value in "แบ่งทุกๆ X ใบ" option

  // §Multi-page label support
  // When an order has multiple SKUs (or one with long content), TikTok renders
  // the label as 2+ pages. Mixing single-page and multi-page labels in one API
  // batch makes it impossible to know which id owns which page — breaking the
  // alias overlay's page→id mapping and the divider's id→page lookup.
  //
  // Strategy: split the id list into segments preserving original order.
  //   - multi-SKU id → its own segment of length 1 (know exact page count)
  //   - run of single-SKU ids → merged into segments up to PRINT_BATCH_SIZE
  //
  // Each segment maps to one generate-API call. The caller then derives
  // pagesPerId[] per segment and aggregates a full mapping.
  function splitIdsIntoApiSegments(ids) {
    const segments = [];
    let current = null;
    const flush = () => { if (current && current.ids.length) segments.push(current); current = null; };
    for (const id of ids) {
      const rec = state.records.get(id);
      const isMulti = (rec?.skuList?.length || 1) > 1;
      if (isMulti) {
        flush();
        segments.push({ ids: [id], isMultiSku: true });
      } else {
        if (!current) current = { ids: [], isMultiSku: false };
        current.ids.push(id);
        if (current.ids.length >= PRINT_BATCH_SIZE) flush();
      }
    }
    flush();
    return segments;
  }

  // Given a segment (from splitIdsIntoApiSegments) and the actual page count of
  // its API-generated PDF, derive pagesPerId[] — one entry per id in segment.
  //   - multi-SKU singleton: all pages belong to the one id
  //   - single-SKU segment with pageCount === ids.length: 1 page each (expected)
  //   - single-SKU segment with extras: distribute extras to leading ids
  //     (best-effort fallback — rare case where TikTok renders a single-SKU
  //     label as 2 pages; we can't know which id without per-id probing)
  function deriveSegmentPagesPerId(segment, pageCount) {
    const n = segment.ids.length;
    if (n === 0) return [];
    if (segment.isMultiSku || n === 1) return [Math.max(1, pageCount)];
    if (pageCount === n) return segment.ids.map(() => 1);
    const base = Math.floor(pageCount / n);
    let extra = pageCount - base * n;
    const out = new Array(n);
    for (let i = 0; i < n; i++) {
      out[i] = Math.max(1, base + (extra > 0 ? 1 : 0));
      if (extra > 0) extra--;
    }
    return out;
  }

  // Dedupe concurrent font fetches when multiple chunks start in parallel —
  // without this, 3 chunks would each fire a fetch for Sarabun-Bold.ttf (88KB)
  // before any wins the race to cache it in state.fontBytes.
  let _fontBytesPromise = null;
  async function ensureFontBytes() {
    if (state.fontBytes) return state.fontBytes;
    if (_fontBytesPromise) return _fontBytesPromise;
    _fontBytesPromise = (async () => {
      if (!state.fontUrl) {
        // Wait briefly for asset-bridge
        const deadline = Date.now() + 2000;
        while (!state.fontUrl && Date.now() < deadline) await sleep(100);
      }
      if (!state.fontUrl) throw new Error('ไม่พบ font URL (asset-bridge ยังไม่ทำงาน)');
      // Retry font fetch up to 3 attempts total with 500ms / 1500ms backoff.
      // A one-off 404/timeout used to permanently cache null → all PDFs
      // rendered Thai as boxes silently until the page was reloaded.
      const backoffs = [0, 500, 1500];
      let lastErr = null;
      for (let i = 0; i < backoffs.length; i++) {
        if (backoffs[i]) await sleep(backoffs[i]);
        try {
          // Use `_origFetch` to bypass TikTok's fetch wrapper — the font lives
          // on `chrome-extension://…`, and going through the wrapper adds
          // signing overhead / potential serialization for no reason.
          const r = await _origFetch.call(window, state.fontUrl);
          if (!r.ok) throw new Error(`HTTP ${r.status}`);
          const bytes = await r.arrayBuffer();
          if (!bytes || !bytes.byteLength) throw new Error('empty font response');
          state.fontBytes = bytes;
          return state.fontBytes;
        } catch (e) {
          lastErr = e;
          console.warn(`[QF] font fetch attempt ${i + 1}/${backoffs.length} failed:`, e);
        }
      }
      throw lastErr || new Error('font fetch failed');
    })();
    try {
      return await _fontBytesPromise;
    } catch (e) {
      // Reset both the in-flight promise AND state.fontBytes so the next
      // print attempt tries again from scratch instead of permanently
      // caching a failed fetch. overlayAliasOnPdf already catches and
      // falls back to the original PDF, so a dropped watermark is
      // non-fatal — we just warn the user it may be missing.
      _fontBytesPromise = null;
      state.fontBytes = null;
      try {
        showErrorToast('โหลดฟอนต์ไม่สำเร็จ — PDF อาจไม่มีลายน้ำ', {
          source: 'ensureFontBytes',
          fontUrl: state.fontUrl || null,
          error: String(e && (e.message || e)),
        });
      } catch {}
      throw e;
    }
  }

  function makeBaseFilename(hint) {
    const now = new Date();
    const pad = n => String(n).padStart(2, '0');
    const stamp = `${now.getFullYear()}${pad(now.getMonth()+1)}${pad(now.getDate())}${pad(now.getHours())}${pad(now.getMinutes())}`;
    // §8.3: Strip only truly unsafe chars; preserve [ ] (valid on Win/macOS/Linux).
    const clean = (hint || 'labels')
      .replace(/[/\\:*?"<>|]/g, '')
      .replace(/\s+/g, ' ')
      .trim()
      .slice(0, 60);
    return `${clean} ${stamp}`;
  }

  // §3: Universal pre-print chunk-plan modal.
  // Resolves to ChunkPlan { mode, n?, x?, withPickingList, combined } or null.
  const PICKING_LIST_PREF_KEY = 'qf_last_picking_list_v1';
  function loadPickingListPref() {
    try { return localStorage.getItem(PICKING_LIST_PREF_KEY) === 'true'; } catch { return false; }
  }
  function savePickingListPref(val) {
    try { localStorage.setItem(PICKING_LIST_PREF_KEY, val ? 'true' : 'false'); } catch {}
  }

  // Fast-path divider preference — remembers user's last toggle in the simple
  // print-confirm modal so repeat prints don't require re-checking every time.
  const DIVIDER_PREF_KEY = 'qf_last_divider_v1';
  function loadDividerPref() {
    try { return localStorage.getItem(DIVIDER_PREF_KEY) === 'true'; } catch { return false; }
  }
  function saveDividerPref(val) {
    try { localStorage.setItem(DIVIDER_PREF_KEY, val ? 'true' : 'false'); } catch {}
  }

  // Modal for print-all in planning mode: 3 worker-centric modes.
  function showPrintAllPlanModal({ total, nonDone, hasMultiSku, defaultPickingList = null }) {
    const lastPickingList = defaultPickingList !== null ? defaultPickingList : loadPickingListPref();
    const names = nonDone.map(([, c]) => c.teamName || c.workerName || '?');
    const titleStr = `พิมพ์ทุกคน · ${names.slice(0, 3).join(', ')}${names.length > 3 ? ` +${names.length - 3}` : ''}`;
    const maxPerPerson = Math.max(...nonDone.map(([, col]) =>
      col.fulfillUnitIds.filter(id => !col.printedIds.includes(id)).length
    ));
    const showChunkOpt = maxPerPerson > CHUNK_AUTO_SAFE_SIZE;

    const skuOpts = hasMultiSku ? `
      <label class="qf-chunk-plan-opt">
        <input type="radio" name="qf-pa-mode" value="by-person-sku">
        <span class="qf-chunk-plan-opt-body">
          <span class="qf-chunk-plan-opt-title">ไฟล์ตาม SKU</span>
          <span class="qf-chunk-plan-opt-sub">แยก PDF ต่อสินค้า-คน เช่น เก๋·สินค้า1, เกม·สินค้า1</span>
        </span>
      </label>
      <label class="qf-chunk-plan-opt">
        <input type="radio" name="qf-pa-mode" value="combined-per-person">
        <span class="qf-chunk-plan-opt-body">
          <span class="qf-chunk-plan-opt-title">รวม SKU แยกตามคน</span>
          <span class="qf-chunk-plan-opt-sub">แต่ละคนได้ PDF รวมทุก SKU พร้อมหน้าคั่น</span>
        </span>
      </label>` : '';

    const chunkOptHtml = showChunkOpt ? `
      <label class="qf-chunk-plan-extra-opt">
        <input type="checkbox" name="qf-pa-chunk" class="qf-pa-chunk-chk"/>
        แยกไฟล์เมื่อเกิน ${CHUNK_AUTO_SAFE_SIZE} ใบ/คน
      </label>` : '';

    return new Promise(resolve => {
      const overlay = document.createElement('div');
      overlay.className = 'qf-modal-overlay qf-chunk-plan-overlay';
      overlay.innerHTML = `
        <div class="qf-modal qf-chunk-plan-modal" role="dialog">
          <div class="qf-modal-title">${escapeHtml(titleStr)}</div>
          <div class="qf-modal-body">
            <div class="qf-modal-count">${total} ใบ</div>
            <div class="qf-modal-summary">${nonDone.length} คน · เลือกรูปแบบไฟล์</div>
            <div class="qf-chunk-plan-options">
              <label class="qf-chunk-plan-opt">
                <input type="radio" name="qf-pa-mode" value="by-person" checked>
                <span class="qf-chunk-plan-opt-body">
                  <span class="qf-chunk-plan-opt-title">แบ่งตามคน</span>
                  <span class="qf-chunk-plan-opt-sub">${nonDone.length} ไฟล์ · แต่ละคนได้ PDF ของตัวเอง</span>
                </span>
              </label>
              ${skuOpts}
            </div>
            <button type="button" class="qf-chunk-plan-more" aria-expanded="false">
              <span class="qf-chunk-plan-more-label">ตัวเลือกเพิ่มเติม</span>
              <span class="qf-chunk-plan-more-count"></span>
              <span class="qf-chunk-plan-chevron">›</span>
            </button>
            <div class="qf-chunk-plan-extras" hidden>
              <label class="qf-chunk-plan-extra-opt">
                <input type="checkbox" name="qf-pa-picking" class="qf-pa-picking-chk" ${lastPickingList ? 'checked' : ''}/>
                แนบใบสรุปรายการ
              </label>
              <label class="qf-chunk-plan-extra-opt">
                <input type="checkbox" name="qf-pa-divider" class="qf-pa-divider-chk" checked/>
                ใส่หน้าคั่นระหว่างสินค้าแต่ละแบบ
              </label>
              ${chunkOptHtml}
            </div>
          </div>
          <div class="qf-modal-actions">
            <button class="qf-btn-cancel">ยกเลิก</button>
            <button class="qf-btn-confirm qf-pa-submit">เริ่มพิมพ์</button>
          </div>
        </div>
      `;
      document.body.appendChild(overlay);

      const moreBtn = overlay.querySelector('.qf-chunk-plan-more');
      const extrasEl = overlay.querySelector('.qf-chunk-plan-extras');
      const moreCountEl = overlay.querySelector('.qf-chunk-plan-more-count');
      const extraChecks = () => [...overlay.querySelectorAll('.qf-chunk-plan-extras input[type="checkbox"]')];
      const updateMoreCount = () => {
        const n = extraChecks().filter(c => c.checked).length;
        moreCountEl.textContent = n > 0 ? `· เปิด ${n} รายการ` : '';
      };
      moreBtn.addEventListener('click', () => {
        const expanded = moreBtn.getAttribute('aria-expanded') === 'true';
        moreBtn.setAttribute('aria-expanded', expanded ? 'false' : 'true');
        extrasEl.hidden = expanded;
      });
      extraChecks().forEach(c => c.addEventListener('change', updateMoreCount));
      updateMoreCount();

      const cleanup = (val) => { overlay.remove(); resolve(val); };
      overlay.querySelector('.qf-btn-cancel').onclick = () => cleanup(null);
      overlay.onclick = (e) => { if (e.target === overlay) cleanup(null); };
      const onKey = (e) => { if (e.key === 'Escape') { cleanup(null); document.removeEventListener('keydown', onKey); } };
      document.addEventListener('keydown', onKey);

      overlay.querySelector('.qf-pa-submit').onclick = () => {
        const mode = (overlay.querySelector('input[name="qf-pa-mode"]:checked') || {}).value || 'by-person';
        const withPickingList = overlay.querySelector('.qf-pa-picking-chk')?.checked || false;
        const withDivider = overlay.querySelector('.qf-pa-divider-chk')?.checked || false;
        const chunkAt = overlay.querySelector('.qf-pa-chunk-chk')?.checked ? CHUNK_AUTO_SAFE_SIZE : null;
        savePickingListPref(withPickingList);
        cleanup({ mode, withPickingList, withDivider, chunkAt });
      };
    });
  }

  function showChunkPlanModal({ total, multiSku = false, defaultPickingList = null, title = null, skuCount = null }) {
    const lastPickingList = defaultPickingList !== null ? defaultPickingList : loadPickingListPref();
    const defaultN = Math.max(2, Math.ceil(total / CHUNK_AUTO_SAFE_SIZE));

    // Case A: small multi-SKU — simplified checkboxes only, no radios
    const isSmallMultiSku = multiSku && total <= CHUNK_AUTO_SAFE_SIZE;
    // Case B: large multi-SKU — radios with renamed "แยกตาม SKU"
    const isLargeMultiSku = multiSku && total > CHUNK_AUTO_SAFE_SIZE;
    // Default radio: by-sku for large multiSku, single for small single-sku, every for large single-sku
    const defaultMode = isLargeMultiSku ? 'by-sku' : (total <= CHUNK_PROMPT_THRESHOLD ? 'single' : 'every');

    // Build the options section based on case
    const buildOptionsHtml = () => {
      if (isSmallMultiSku) {
        // Case A: no radios — only checkboxes inline
        // "รวม SKU ไว้ด้วยกัน" (was "รวมสินค้าเป็นไฟล์เดียว"):
        // old wording promised "one file" but when paired with chunk splitting
        // (even/every) the result is N files, not one — the new phrase makes
        // clear what "combined" actually controls: whether SKUs are mixed
        // across files or kept each in its own file.
        return `<div class="qf-chunk-plan-options qf-chunk-plan-simple">
          <label class="qf-chunk-plan-extra-opt">
            <input type="checkbox" class="qf-chunk-combined-chk"/>
            รวม SKU ไว้ด้วยกัน
          </label>
          <label class="qf-chunk-plan-extra-opt">
            <input type="checkbox" class="qf-chunk-picking-chk" ${lastPickingList ? 'checked' : ''}/>
            แนบใบสรุปรายการ
          </label>
          <label class="qf-chunk-plan-extra-opt">
            <input type="checkbox" class="qf-chunk-divider-chk" checked/>
            ใส่หน้าคั่นระหว่างสินค้าแต่ละแบบ
          </label>
        </div>`;
      }
      // Cases B and C: radio options
      const singleOrSkuOpt = isLargeMultiSku
        ? `<label class="qf-chunk-plan-opt">
            <input type="radio" name="qf-chunk-mode" value="by-sku" ${defaultMode === 'by-sku' ? 'checked' : ''}>
            <span class="qf-chunk-plan-opt-body">
              <span class="qf-chunk-plan-opt-title">แยกตาม SKU</span>
              <span class="qf-chunk-plan-opt-sub">1 ไฟล์ต่อสินค้า</span>
            </span>
          </label>`
        : `<label class="qf-chunk-plan-opt">
            <input type="radio" name="qf-chunk-mode" value="single" ${defaultMode === 'single' ? 'checked' : ''}>
            <span class="qf-chunk-plan-opt-body">
              <span class="qf-chunk-plan-opt-title">ไฟล์เดียว</span>
              <span class="qf-chunk-plan-opt-sub">${total} ใบในไฟล์เดียว</span>
            </span>
          </label>`;
      return `<div class="qf-chunk-plan-options">
        ${singleOrSkuOpt}
        <label class="qf-chunk-plan-opt">
          <input type="radio" name="qf-chunk-mode" value="even" ${defaultMode === 'even' ? 'checked' : ''}>
          <span class="qf-chunk-plan-opt-body">
            <span class="qf-chunk-plan-opt-title">
              แบ่งเท่า ๆ กัน
              <input type="number" name="qf-chunk-N" class="qf-chunk-plan-input qf-chunk-n-input"
                min="2" max="${total}" value="${defaultN}"/>
              <span class="qf-chunk-unit">ชุด</span>
            </span>
            <span class="qf-chunk-plan-opt-sub qf-chunk-n-preview"></span>
          </span>
        </label>
        ${total > CHUNK_AUTO_SAFE_SIZE ? `<label class="qf-chunk-plan-opt">
          <input type="radio" name="qf-chunk-mode" value="every" ${defaultMode === 'every' ? 'checked' : ''}>
          <span class="qf-chunk-plan-opt-body">
            <span class="qf-chunk-plan-opt-title">
              กำหนดจำนวน
              <input type="number" name="qf-chunk-X" class="qf-chunk-plan-input qf-chunk-x-input"
                min="1" max="${total}" value="${CHUNK_AUTO_SAFE_SIZE}"/>
              <span class="qf-chunk-unit">ใบ/ไฟล์</span>
            </span>
            <span class="qf-chunk-plan-opt-sub qf-chunk-x-preview"></span>
          </span>
        </label>` : ''}
      </div>`;
    };

    // Extras section (only for Cases B and C — Case A has inline checkboxes)
    const buildExtrasHtml = () => {
      if (isSmallMultiSku) return '';
      return `<button type="button" class="qf-chunk-plan-more" aria-expanded="false">
          <span class="qf-chunk-plan-more-label">ตัวเลือกเพิ่มเติม</span>
          <span class="qf-chunk-plan-more-count"></span>
          <span class="qf-chunk-plan-chevron">›</span>
        </button>
        <div class="qf-chunk-plan-extras" hidden>
          <label class="qf-chunk-plan-extra-opt">
            <input type="checkbox" class="qf-chunk-picking-chk" ${lastPickingList ? 'checked' : ''}/>
            แนบใบสรุปรายการ
          </label>
          ${isLargeMultiSku ? `<label class="qf-chunk-plan-extra-opt">
            <input type="checkbox" class="qf-chunk-combined-chk"/>
            รวม SKU ไว้ด้วยกัน
          </label>` : ''}
          <label class="qf-chunk-plan-extra-opt">
            <input type="checkbox" class="qf-chunk-divider-chk" checked/>
            ใส่หน้าคั่นระหว่างสินค้าแต่ละแบบ
          </label>
        </div>`;
    };

    return new Promise(resolve => {
      const overlay = document.createElement('div');
      overlay.className = 'qf-modal-overlay qf-chunk-plan-overlay';
      overlay.innerHTML = `
        <div class="qf-modal qf-chunk-plan-modal" role="dialog">
          <div class="qf-modal-title">${escapeHtml(title || 'เตรียมพิมพ์ฉลาก')}</div>
          <div class="qf-modal-body">
            <div class="qf-modal-count">${total} ใบ</div>
            <div class="qf-modal-summary">${isSmallMultiSku ? 'เลือกตัวเลือกการพิมพ์' : 'เลือกวิธีแบ่งไฟล์'}</div>
            ${buildOptionsHtml()}
            ${buildExtrasHtml()}
          </div>
          <div class="qf-modal-actions">
            <button class="qf-btn-cancel">ยกเลิก</button>
            <button class="qf-btn-confirm qf-chunk-plan-submit">เริ่มพิมพ์</button>
          </div>
        </div>
      `;
      document.body.appendChild(overlay);

      const radioEls  = () => [...overlay.querySelectorAll('input[name="qf-chunk-mode"]')];
      const nInput    = overlay.querySelector('.qf-chunk-n-input');
      const xInput    = overlay.querySelector('.qf-chunk-x-input');
      const nPreview  = overlay.querySelector('.qf-chunk-n-preview');
      const xPreview  = overlay.querySelector('.qf-chunk-x-preview');
      const combinedChk = overlay.querySelector('.qf-chunk-combined-chk');
      const submitBtn = overlay.querySelector('.qf-chunk-plan-submit');

      const getMode = () => (radioEls().find(r => r.checked) || {}).value || 'single';

      // Preview adapts to the combined-SKU checkbox so the same N/X value
      // tells two very different stories:
      //   - combined ON  → N/X slices the TOTAL label set (mixed SKU per file)
      //   - combined OFF → N/X slices EACH SKU separately (×skuCount files)
      // Previously the UI showed one formula for both cases, which caused
      // users to type a number expecting global behavior but get per-SKU.
      const updatePreviews = () => {
        const combined = !!(combinedChk && combinedChk.checked);
        if (nInput && nPreview) {
          const n = parseInt(nInput.value) || defaultN;
          const sz = `~${Math.floor(total/n)}–${Math.ceil(total/n)} ใบ/ไฟล์`;
          if (combined || !multiSku) {
            nPreview.textContent = `${sz} · ${n} ไฟล์ (ผสม SKU)`;
          } else if (skuCount && skuCount > 1) {
            nPreview.textContent = `แต่ละ SKU แบ่ง ${n} ชุด · รวม ~${n * skuCount} ไฟล์`;
          } else {
            nPreview.textContent = `แต่ละ SKU แบ่ง ${n} ชุด`;
          }
        }
        if (xInput && xPreview) {
          const x = parseInt(xInput.value) || CHUNK_AUTO_SAFE_SIZE;
          if (combined || !multiSku) {
            xPreview.textContent = `${Math.ceil(total/x)} ไฟล์ (ผสม SKU · ~${x} ใบ/ไฟล์)`;
          } else {
            xPreview.textContent = `แบ่งทุก ${x} ใบต่อ SKU`;
          }
        }
      };
      updatePreviews();
      if (combinedChk) combinedChk.addEventListener('change', updatePreviews);

      [nInput, xInput].filter(Boolean).forEach(inp => {
        inp.addEventListener('focus', () => {
          const radioVal = inp.name === 'qf-chunk-N' ? 'even' : 'every';
          const radio = overlay.querySelector(`input[value="${radioVal}"]`);
          if (radio) radio.checked = true;
        });
        inp.addEventListener('input', updatePreviews);
        inp.addEventListener('blur', () => {
          const isN = inp.name === 'qf-chunk-N';
          const min = isN ? 2 : 1;
          let val = parseInt(inp.value);
          if (isNaN(val) || val < min) val = min;
          if (val > total) val = total;
          if (String(val) !== inp.value) {
            inp.value = val;
            inp.classList.add('qf-input-error');
            setTimeout(() => inp.classList.remove('qf-input-error'), 600);
          }
          updatePreviews();
        });
        inp.addEventListener('keydown', (e) => {
          if (e.key === 'Enter') { e.preventDefault(); submitBtn.click(); }
        });
      });

      radioEls().forEach(r => r.addEventListener('change', updatePreviews));

      // Collapsible extras (Cases B and C only)
      const moreBtn = overlay.querySelector('.qf-chunk-plan-more');
      const extrasEl = overlay.querySelector('.qf-chunk-plan-extras');
      const moreCountEl = overlay.querySelector('.qf-chunk-plan-more-count');
      if (moreBtn) {
        const extraChecks = () => [...overlay.querySelectorAll('.qf-chunk-plan-extras input[type="checkbox"]')];
        const updateMoreCount = () => {
          const cnt = extraChecks().filter(c => c.checked).length;
          moreCountEl.textContent = cnt > 0 ? `· เปิด ${cnt} รายการ` : '';
        };
        moreBtn.addEventListener('click', () => {
          const expanded = moreBtn.getAttribute('aria-expanded') === 'true';
          moreBtn.setAttribute('aria-expanded', expanded ? 'false' : 'true');
          extrasEl.hidden = expanded;
        });
        extraChecks().forEach(c => c.addEventListener('change', updateMoreCount));
        updateMoreCount();
      }

      const cleanup = (val) => { overlay.remove(); resolve(val); };

      overlay.querySelector('.qf-btn-cancel').onclick = () => cleanup(null);
      overlay.onclick = (e) => { if (e.target === overlay) cleanup(null); };
      const onKey = (e) => { if (e.key === 'Escape') { cleanup(null); document.removeEventListener('keydown', onKey); } };
      document.addEventListener('keydown', onKey);

      submitBtn.onclick = () => {
        const withPickingList = overlay.querySelector('.qf-chunk-picking-chk')?.checked || false;
        const combined = overlay.querySelector('.qf-chunk-combined-chk')?.checked || false;
        const withDivider = overlay.querySelector('.qf-chunk-divider-chk')?.checked !== false;
        savePickingListPref(withPickingList);

        // Case A: small multi-SKU — no radio, mode from combined checkbox
        if (isSmallMultiSku) {
          const mode = combined ? 'single' : 'by-sku';
          cleanup({ mode, n: null, x: null, withPickingList, combined, withDivider });
          return;
        }

        const mode = getMode();
        let n = null, x = null;

        if (mode === 'even') {
          n = parseInt(nInput.value);
          const clamped = Math.max(2, Math.min(total, isNaN(n) ? defaultN : n));
          if (clamped !== n) {
            nInput.value = clamped;
            nInput.classList.add('qf-input-error');
            setTimeout(() => nInput.classList.remove('qf-input-error'), 600);
            return;
          }
          n = clamped;
        } else if (mode === 'every') {
          x = parseInt(xInput.value);
          const clamped = Math.max(1, Math.min(total, isNaN(x) ? CHUNK_AUTO_SAFE_SIZE : x));
          if (clamped !== x) {
            xInput.value = clamped;
            xInput.classList.add('qf-input-error');
            setTimeout(() => xInput.classList.remove('qf-input-error'), 600);
            return;
          }
          x = clamped;
        }

        cleanup({ mode, n, x, withPickingList, combined, withDivider });
      };
    });
  }

  function showChunkedResult({title, totalIds, chunks}) {
    // chunks: [{count, label, filename}]
    document.querySelectorAll('.qf-progress-overlay').forEach(e => e.remove());
    document.querySelectorAll('.qf-chunked-bubble').forEach(e => e.remove());
    const overlay = document.createElement('div');
    overlay.className = 'qf-progress-overlay qf-chunked-overlay';
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
        <div class="qf-chunked-header">
          <div class="qf-progress-title">${escapeHtml(title)}</div>
          <div class="qf-chunked-header-actions">
            <button class="qf-btn-minimize-result" title="ย่อ">ย่อ</button>
            <button class="qf-btn-x-result" aria-label="ปิด" style="display:none;">×</button>
          </div>
        </div>
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
    // chunkIdx → {url, filename, downloaded}
    const blobUrls = {};
    const totalChunks = chunks.length;
    const downloadAllBtn = overlay.querySelector('.qf-btn-download-all');
    const closeBtn = overlay.querySelector('.qf-btn-close-result');
    const xBtn = overlay.querySelector('.qf-btn-x-result');
    const minBtn = overlay.querySelector('.qf-btn-minimize-result');
    let bubble = null;
    let allDoneFlag = false;

    const markDownloaded = (i) => {
      if (blobUrls[i]) blobUrls[i].downloaded = true;
      const row = cards[i];
      if (row) row.classList.add('qf-chunk-downloaded');
      updateBubbleLabel();
    };

    const countUndownloaded = () => {
      return Object.values(blobUrls).filter(b => b && !b.downloaded).length;
    };
    const countPending = () => {
      // Chunks that don't yet have a blob (not completed) and not errored
      let pending = 0;
      for (let i = 0; i < totalChunks; i++) {
        if (!blobUrls[i]) {
          const row = cards[i];
          if (!row.classList.contains('qf-chunk-error')) pending++;
        }
      }
      return pending;
    };

    const doCleanup = () => {
      Object.values(blobUrls).forEach(({url}) => { try { URL.revokeObjectURL(url); } catch {} });
      overlay.remove();
      if (bubble) { bubble.remove(); bubble = null; }
    };

    const tryClose = () => {
      const undownloaded = countUndownloaded();
      if (undownloaded > 0) {
        showCloseUndownloadedConfirm(undownloaded).then(ok => {
          if (ok) doCleanup();
        });
      } else {
        doCleanup();
      }
    };

    closeBtn.onclick = tryClose;
    xBtn.onclick = tryClose;

    // P0: removed background-click dismiss — users were accidentally losing
    // rendered PDFs by clicking outside the modal. Close is now explicit.

    downloadAllBtn.onclick = () => {
      const entries = Object.entries(blobUrls);
      entries.forEach(([i, {url, filename}], idx) => {
        setTimeout(() => {
          const a = document.createElement('a');
          a.href = url; a.download = filename;
          document.body.appendChild(a); a.click(); a.remove();
          markDownloaded(parseInt(i, 10));
        }, idx * 200);
      });
    };

    const updateBubbleLabel = () => {
      if (!bubble) return;
      const ready = Object.keys(blobUrls).length;
      const undownloaded = countUndownloaded();
      const pending = countPending();
      const txt = bubble.querySelector('.qf-chunked-bubble-text');
      if (!txt) return;
      if (pending > 0) {
        txt.textContent = `กำลังพิมพ์ ${ready}/${totalChunks}`;
      } else if (undownloaded > 0) {
        txt.textContent = `พร้อมโหลด ${undownloaded} ไฟล์`;
      } else {
        txt.textContent = `เสร็จแล้ว ${totalChunks} ไฟล์`;
      }
    };

    const showBubble = () => {
      if (bubble) return;
      bubble = document.createElement('div');
      bubble.className = 'qf-chunked-bubble';
      bubble.innerHTML = `
        <span class="qf-chunked-bubble-icon">📄</span>
        <span class="qf-chunked-bubble-text">${escapeHtml(title)}</span>
        <button class="qf-chunked-bubble-close" aria-label="ปิด">×</button>
      `;
      document.body.appendChild(bubble);
      // Expand on bubble click (but not on × click)
      bubble.addEventListener('click', (e) => {
        if (e.target.closest('.qf-chunked-bubble-close')) return;
        hideBubble();
        overlay.style.display = '';
      });
      bubble.querySelector('.qf-chunked-bubble-close').addEventListener('click', (e) => {
        e.stopPropagation();
        tryClose();
      });
      updateBubbleLabel();
    };

    const hideBubble = () => {
      if (!bubble) return;
      bubble.remove();
      bubble = null;
    };

    minBtn.onclick = () => {
      overlay.style.display = 'none';
      showBubble();
    };

    return {
      startChunk(i) {
        const row = cards[i];
        row.querySelector('.qf-chunk-row-status').textContent = 'กำลังทำ...';
        row.classList.add('qf-chunk-active');
        updateBubbleLabel();
      },
      updateChunkProgress(i, pct, label) {
        const row = cards[i];
        row.querySelector('.qf-chunk-row-fill').style.width = (pct * 100).toFixed(0) + '%';
        if (label) row.querySelector('.qf-chunk-row-status').textContent = label;
      },
      completeChunk(i, {url, pageCount, labelCount}) {
        const row = cards[i];
        const filename = chunks[i].filename || `chunk-${i+1}.pdf`;
        blobUrls[i] = {url, filename, downloaded: false};
        row.querySelector('.qf-chunk-row-fill').style.width = '100%';
        // Display "N ใบ · M หน้า" when dividers/picking-lists add pages beyond
        // the label count. When page count equals label count, keep the short form.
        const statusText = (labelCount != null && pageCount != null && pageCount !== labelCount)
          ? `${labelCount} ใบ · ${pageCount} หน้า`
          : (labelCount != null ? `${labelCount} ใบ` : `${pageCount} หน้า`);
        row.querySelector('.qf-chunk-row-status').textContent = statusText;
        row.classList.remove('qf-chunk-active');
        row.classList.add('qf-chunk-done');
        const actions = row.querySelector('.qf-chunk-row-actions');
        actions.innerHTML = `
          <button class="qf-chunk-btn qf-chunk-open">เปิด</button>
          <button class="qf-chunk-btn qf-chunk-download">ดาวน์โหลด</button>
        `;
        actions.querySelector('.qf-chunk-open').onclick = (e) => {
          e.stopPropagation();
          const w = window.open(url, '_blank');
          if (!w) {
            const a = document.createElement('a');
            a.href = url; a.target = '_blank'; a.rel = 'noopener';
            document.body.appendChild(a); a.click(); a.remove();
          }
          markDownloaded(i);
        };
        actions.querySelector('.qf-chunk-download').onclick = (e) => {
          e.stopPropagation();
          const a = document.createElement('a');
          a.href = url; a.download = filename;
          document.body.appendChild(a); a.click(); a.remove();
          markDownloaded(i);
        };
        if (Object.keys(blobUrls).length > 0) {
          downloadAllBtn.disabled = false;
        }
        updateBubbleLabel();
      },
      errorChunk(i, msg) {
        const row = cards[i];
        row.querySelector('.qf-chunk-row-status').textContent = 'ล้มเหลว: ' + msg;
        row.classList.remove('qf-chunk-active');
        row.classList.add('qf-chunk-error');
        const actions = row.querySelector('.qf-chunk-row-actions');
        actions.innerHTML = `<button class="qf-chunk-btn qf-chunk-retry">ลองใหม่</button>`;
        actions.querySelector('.qf-chunk-retry').onclick = async (e) => {
          e.stopPropagation();
          // External retry handler — set via setRetryHandler
          if (this._retry) await this._retry(i);
        };
        updateBubbleLabel();
      },
      setRetryHandler(fn) { this._retry = fn; },
      allDone() {
        allDoneFlag = true;
        closeBtn.style.display = '';
        xBtn.style.display = '';
        updateBubbleLabel();
      },
      cleanup: doCleanup,
    };
  }

  // Confirm dialog when closing chunked result with undownloaded files.
  // Reassures user that their history tab can recover the files.
  function showCloseUndownloadedConfirm(undownloadedCount) {
    return new Promise(resolve => {
      const overlay = document.createElement('div');
      overlay.className = 'qf-modal-overlay';
      overlay.innerHTML = `
        <div class="qf-modal" role="dialog">
          <div class="qf-modal-title">ยังมีไฟล์ที่ยังไม่ได้ดาวน์โหลด</div>
          <div class="qf-modal-body">
            <div class="qf-modal-target">มี <b>${undownloadedCount}</b> ไฟล์ที่สร้างแล้วแต่ยังไม่ได้เปิด/ดาวน์โหลด</div>
            <div class="qf-modal-summary">ทิ้งไฟล์เหล่านี้ไป? (เปิดย้อนหลังได้จากปุ่ม ⏱ ประวัติ)</div>
          </div>
          <div class="qf-modal-actions">
            <button class="qf-btn-cancel">ยกเลิก</button>
            <button class="qf-btn-confirm qf-btn-danger">ทิ้ง</button>
          </div>
        </div>
      `;
      document.body.appendChild(overlay);
      const cleanup = (v) => { overlay.remove(); resolve(v); };
      overlay.querySelector('.qf-btn-cancel').onclick = (e) => { e.stopPropagation(); cleanup(false); };
      overlay.querySelector('.qf-btn-confirm').onclick = (e) => { e.stopPropagation(); cleanup(true); };
      overlay.onclick = (e) => { if (e.target === overlay) cleanup(false); };
      const onKey = (e) => { if (e.key === 'Escape') { cleanup(false); document.removeEventListener('keydown', onKey); } };
      document.addEventListener('keydown', onKey);
    });
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
    // Exclude synthetic no-variant entries (skuId=null from Shopee items
    // without model_id) — they're not real variants.
    const variants = [...product.variants.values()].filter(v => v.skuId != null);

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

        <div class="qf-alias-modal-body">
          <div class="qf-alias-modal-section">
            <div class="qf-alias-modal-label">ชื่อย่อหลัก</div>
            <div class="qf-alias-modal-hint">ใช้กับทุกตัวเลือกที่ไม่ได้ตั้งชื่อแยก</div>
            <input class="qf-alias-modal-product" type="text" placeholder="เช่น ครีม, แดง1, สครับ" maxlength="20" value="${escapeHtml(getAlias(productId) || '')}"/>
          </div>

          <div class="qf-alias-modal-section">
            <div class="qf-alias-modal-label">ตั้งชื่อแยกตามตัวเลือก</div>
            <div class="qf-alias-modal-hint">ปล่อยว่างจะใช้ชื่อหลักแทน</div>
            <div class="qf-alias-modal-variants"></div>
          </div>
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
      setAlias(productId, v);
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

  // Resolve alias lines for a SKU with strict "user-set wins" semantics.
  //
  //   rawAlias        = explicit product alias set by user (empty string if not set)
  //   variantOverride = explicit variant alias set by user (empty if not set)
  //
  // Rule (per user request): when the user has set a product alias, we must NOT
  // fall back to the SKU default name from the API in the secondary/variant slot
  // — otherwise the user's customization is drowned out by TikTok's verbose
  // default names. Fallback to API/shortName only applies when NO product alias
  // is set (so the label is still readable for un-aliased products).
  function resolveSkuAlias(s) {
    const v = getVariantInfo(s.productId, s.skuId);
    const rawAlias = (getAlias(s.productId) || '').trim();
    const variantOverride = (v?.alias || '').trim();
    const hasProductAlias = !!rawAlias;
    const productDisplay = rawAlias || shortName(s.productName);
    // Variant slot: prefer explicit override; fall back to API name ONLY when
    // no product alias is set; otherwise blank (user has spoken).
    const variantDisplay = variantOverride
      || (hasProductAlias ? '' : variantDisplayName(s));
    return {
      productDisplay,
      variantDisplay,
      variantOverride,
      hasProductAlias,
      replace: !!(v?.replace && variantOverride),
      qty: s.quantity || 1,
    };
  }

  function buildSkuRender(s) {
    // primary = big top line (alias + qty), secondary = small bottom (variant-only).
    // Replace mode → variant alias fully stands in for product display.
    const r = resolveSkuAlias(s);
    if (r.replace) {
      return { primary: `${r.variantOverride} x${r.qty}`, secondary: '' };
    }
    return {
      primary: `${r.productDisplay} x${r.qty}`,
      secondary: r.variantDisplay,
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
      const r = resolveSkuAlias(s);
      if (r.replace) return `${r.variantOverride} x${r.qty}`;
      return r.variantDisplay
        ? `${r.productDisplay} ${r.variantDisplay} x${r.qty}`
        : `${r.productDisplay} x${r.qty}`;
    });
    return { mode: 'multi', text: parts.join(' + ') };
  }

  async function overlayAliasOnPdf(pdfBytes, fulfillUnitIds, onProgress, workerName, workerIcon, noteMap = null, pagesPerId = null) {
    if (!window.PDFLib) return pdfBytes;
    const { PDFDocument, rgb, degrees } = window.PDFLib;
    const fontBytes = await ensureFontBytes();
    const pdfDoc = await PDFDocument.load(pdfBytes);
    if (window.fontkit) pdfDoc.registerFontkit(window.fontkit);
    const font = await pdfDoc.embedFont(fontBytes, { subset: true });
    const pages = pdfDoc.getPages();
    if (!pages.length || !fulfillUnitIds?.length) return await pdfDoc.save();

    // §Multi-page labels: when caller knows exact pages per id (via segment
    // splitting), build prefix-sum array so each page maps to the correct
    // fulfillUnitId even when label page counts vary.
    //   pagesPerId[i] === number of consecutive pages belonging to ids[i]
    //   pageToUnitIdx[p] === index into fulfillUnitIds
    let pageToUnitIdx = null;
    if (Array.isArray(pagesPerId) && pagesPerId.length === fulfillUnitIds.length) {
      pageToUnitIdx = new Array(pages.length);
      let p = 0;
      for (let ui = 0; ui < pagesPerId.length; ui++) {
        const n = Math.max(1, pagesPerId[ui] | 0);
        for (let k = 0; k < n && p < pages.length; k++, p++) pageToUnitIdx[p] = ui;
      }
      // Defensive fill in case pagesPerId sum < pages.length.
      for (; p < pages.length; p++) pageToUnitIdx[p] = fulfillUnitIds.length - 1;
    }

    const lo = loadLabelOverlay();
    const OP = Math.min(lo.opacity ?? 0.85, 0.90);

    // Pre-embed shop image once for all pages
    let shopImg = null;
    if (lo.enabled && lo.shopImageDataUrl) {
      try {
        const b64 = lo.shopImageDataUrl.split(',')[1] || '';
        const raw = atob(b64);
        const u8 = new Uint8Array(raw.length);
        for (let j = 0; j < raw.length; j++) u8[j] = raw.charCodeAt(j);
        shopImg = lo.shopImageDataUrl.startsWith('data:image/png')
          ? await pdfDoc.embedPng(u8) : await pdfDoc.embedJpg(u8);
      } catch { shopImg = null; }
    }

    const pagesPerUnit = Math.max(1, Math.round(pages.length / fulfillUnitIds.length));
    const total = pages.length;

    // Fit text to maxWidth, shrink size if needed
    const fitWidth = (text, baseSize, maxWidth) => {
      let size = baseSize;
      const w = font.widthOfTextAtSize(text, size);
      if (w > maxWidth) size = size * (maxWidth / w);
      return { size, width: font.widthOfTextAtSize(text, size) };
    };

    // Draw centered text with optional white stroke (for alias/product labels)
    const draw = (page, text, y, baseSize, opts = {}) => {
      const { width } = page.getSize();
      const { size, width: tw } = fitWidth(text, baseSize, width - 16);
      const x = (width - tw) / 2;
      const op = opts.opacity ?? 0.85;
      if (opts.stroke) {
        for (const [dx, dy] of [[-1,0],[1,0],[0,-1],[0,1]]) {
          page.drawText(text, { x: x+dx, y: y+dy, size, font, color: rgb(1,1,1), opacity: Math.min(op+0.05, 0.95) });
        }
      }
      page.drawText(text, { x, y, size, font, color: rgb(0,0,0), opacity: op });
    };

    for (let i = 0; i < pages.length; i++) {
      const unitIdx = pageToUnitIdx
        ? pageToUnitIdx[i]
        : Math.min(fulfillUnitIds.length - 1, Math.floor(i / pagesPerUnit));
      const fulfillId = fulfillUnitIds[unitIdx];
      const rec = state.records.get(fulfillId);
      const lines = buildPageLines(rec);
      const page = pages[i];
      const { width: pw, height: ph } = page.getSize();
      // aliasFontSize=0 → auto formula; >0 → user-set pt value (bigSize = set, smallSize = set×0.65).
      const _afs = lo.aliasFontSize || 0;
      const bigSize = _afs > 0 ? _afs : Math.min(ph * 0.05, 22) * 0.9;
      const smallSize = _afs > 0 ? Math.max(7, _afs * 0.65) : Math.min(ph * 0.032, 13) * 0.9;

      // ── J&T carrier detection ────────────────────────────────────────────
      const carrierId = state.carrierOf.get(fulfillId) || '';
      const carrierName = (state.carriers.get(carrierId)?.name || '').toLowerCase();
      const isJnT = carrierName.includes('j&t') || carrierName.includes('j and t') || carrierName.includes('jnt') || carrierId.toLowerCase().includes('jt') || carrierId.toLowerCase().includes('j&t');
      // Only apply J&T masks when page dimensions exactly match calibrated A6 (298×420pt ±5pt).
      // Live labels from different batches may be A4 or other sizes — wrong-size masks corrupt content.
      const isA6 = Math.abs(pw - 298) <= 5 && Math.abs(ph - 420) <= 5;

      // ── Label overlay (J&T A6 only when enabled) ─────────────────────────
      if (lo.enabled && isJnT && isA6) {
        // 1. White mask rects
        for (const r of J_AND_T_MASK_RECTS) {
          page.drawRectangle({ x: r.x, y: r.y, width: r.w, height: r.h, color: rgb(1,1,1), borderWidth: 0 });
        }

        // 2. Vertical marketing text in masked columns (rotated 90°), centered in 362pt zone
        const mktText = (lo.marketingText || '').slice(0, 50);
        if (mktText) {
          // marketingFontSize=0 → auto (7–12pt based on length); >0 → user override, capped at 15pt.
          const _mfs = lo.marketingFontSize || 0;
          const colSize = _mfs > 0 ? Math.min(15, _mfs) : Math.max(7, Math.min(12, 160 / Math.max(mktText.length, 13)));
          const mktW = font.widthOfTextAtSize(mktText, colSize);
          // Mask zone: y=38..340 (302pt, reduced from 400 to leave top 12-digit OCR group uncovered).
          // Center at y=189 = (38+340)/2.
          const colY = Math.max(38, Math.min(189 - mktW / 2, 340 - mktW));
          // Baseline x shifted inward to keep ascenders inside mask at larger font sizes:
          //   left mask  x=0..17  → baseline x=11.5 (ascenders reach x≈2.5 at 12pt)
          //   right mask x=283..298 → baseline x=293.5 (ascenders reach x≈284.5 at 12pt)
          page.drawText(mktText, { x: 11.5,  y: colY, size: colSize, font, color: rgb(0,0,0), rotate: degrees(90), opacity: OP });
          page.drawText(mktText, { x: 293.5, y: colY, size: colSize, font, color: rgb(0,0,0), rotate: degrees(90), opacity: OP });
        }

        // 3. Shop image + header main/sub — only for single-SKU orders (not multi-SKU combo)
        const isSingleSku = lines?.mode !== 'multi';
        if (isSingleSku) {
          // Image slot: x=5..27, y=42..64 (22×22pt). Left padding 5pt (was 2pt) for breathing room.
          if (shopImg) {
            page.drawImage(shopImg, { x: 5, y: 42, width: 22, height: 22, opacity: OP });
          }
          // Header left padding: 30pt from edge (image+gap 22+3=25, then 5pt base = 30)
          // No-image fallback: 12pt indent for aligned visual rhythm.
          const hdrX = shopImg ? 30 : 12;
          // maxW: Order ID starts x≈190 — leave 5pt safety margin (ends at x=185).
          const hdrMaxW = shopImg ? 155 : 173;
          const h1 = (lo.headerMain || '').slice(0, 50);
          const h2 = (lo.headerSub || '').slice(0, 50);
          // h1 baseline y=54 @ 9pt → occupies y≈51.75..60.75 (safely between masks at 62 and 41.5)
          // h2 baseline y=44 @ 7pt → occupies y≈42.25..49.25 (3pt gap from h1 for designer rhythm)
          if (h1) {
            const h1Base = lo.header1FontSize > 0 ? Math.min(14, lo.header1FontSize) : 9;
            const { size: s1 } = fitWidth(h1, h1Base, hdrMaxW);
            page.drawText(h1, { x: hdrX, y: 54, size: s1, font, color: rgb(0,0,0), opacity: OP });
          }
          if (h2) {
            const h2Base = lo.header2FontSize > 0 ? Math.min(12, lo.header2FontSize) : 7;
            const { size: s2 } = fitWidth(h2, h2Base, hdrMaxW);
            page.drawText(h2, { x: hdrX, y: 44, size: s2, font, color: rgb(0,0,0), opacity: OP });
          }
        }
      }

      // ── Alias/product label — opacity reduced a further 10% (×0.9) vs global OP
      const ALIAS_OP = OP * 0.9;
      if (lines?.mode === 'single') {
        if (lines.secondary) {
          draw(page, lines.primary, 4 + smallSize + 2, bigSize, { stroke: true, opacity: ALIAS_OP });
          draw(page, lines.secondary, 4, smallSize, { stroke: true, opacity: ALIAS_OP });
        } else {
          draw(page, lines.primary, 6, bigSize, { stroke: true, opacity: ALIAS_OP });
        }
      } else if (lines?.mode === 'multi') {
        draw(page, lines.text, 6, bigSize, { stroke: true, opacity: ALIAS_OP });
      }

      // ── Worker name (top-right) ──────────────────────────────────────────
      // White stroke (4-dir offset) ช่วยให้อ่านชื่อคนแพ็คออกแม้ทับกับข้อความ
      // ฉลากที่เป็นสีเข้ม (เช่น เลข tracking, barcode labels) — ใช้ pattern
      // เดียวกับ alias overlay ข้างบนแต่ไม่ผ่าน `draw()` เพราะชื่อชิดขวา ไม่ใช่กลาง.
      if (workerName) {
        const wSize = Math.min(ph * 0.04, 18);
        const wWidth = font.widthOfTextAtSize(workerName, wSize);
        const wX = Math.max(6, pw - wWidth - 6);
        const wY = ph - wSize - 6;
        const strokeOp = Math.min(OP + 0.05, 0.95);
        for (const [dx, dy] of [[-1,0],[1,0],[0,-1],[0,1]]) {
          page.drawText(workerName, { x: wX + dx, y: wY + dy, size: wSize, font, color: rgb(1,1,1), opacity: strokeOp });
        }
        page.drawText(workerName, { x: wX, y: wY, size: wSize, font, color: rgb(0,0,0), opacity: OP });
      }

      // ── Note markers ★N ─────────────────────────────────────────────────
      // Badges use ★ (verified present in Sarabun-Bold — no .notdef square)
      // with Thai letter prefix for kind: "ล" (ลูกค้า) / "ร" (ร้านค้า).
      // Both live at TOP-LEFT in a single row (buyer first, seller after a
      // small gap). Kept away from the top-right worker-name watermark to
      // avoid collision, and drawn adjacent when both exist so packers see
      // them together as one "note cluster".
      //
      // noteMap may be either a legacy Map (buyer-only), or an object
      // { buyer?: Map, seller?: Map }. Read both forms.
      const buyerMap  = (noteMap && noteMap.buyer)  ? noteMap.buyer  : (noteMap instanceof Map ? noteMap : null);
      const sellerMap = (noteMap && noteMap.seller) ? noteMap.seller : null;
      const noteSize = Math.min(ph * 0.032, 12);
      const noteY    = ph - noteSize - 3;
      const drawNoteBadge = (text, x) => {
        for (const [dx, dy] of [[-1,0],[1,0],[0,-1],[0,1]]) {
          page.drawText(text, { x: x+dx, y: noteY+dy, size: noteSize, font, color: rgb(1,1,1), opacity: 0.8 });
        }
        page.drawText(text, { x, y: noteY, size: noteSize, font, color: rgb(0,0,0), opacity: 0.78 });
        return font.widthOfTextAtSize(text, noteSize);
      };
      let noteCursorX = 4;
      const hasBuyer  = rec?.hasBuyerNote || (rec?.hasNote && !rec?.hasSellerNote);
      if (hasBuyer) {
        const idx = buyerMap?.get(fulfillId);
        const t = idx != null ? `ล★${idx}` : 'ล★';
        noteCursorX += drawNoteBadge(t, noteCursorX) + 6; // gap before next badge
      }
      if (rec?.hasSellerNote) {
        const idx = sellerMap?.get(fulfillId);
        const t = idx != null ? `ร★${idx}` : 'ร★';
        drawNoteBadge(t, noteCursorX);
      }

      if (onProgress && (i % 20 === 0 || i === total - 1)) { onProgress(i + 1, total); await sleep(0); }
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
  // Check if record's labelStatus matches current filter — needed AFTER optimistic
  // updates flip rec.labelStatus from 30 → 50, so counts drop in "ยังไม่พิมพ์".
  // Fallback: if labelStatus is missing (record predates 1.12.1 or normalize
  // missed the field), assume it matches the server-side filter we scanned with.
  function passesLabelStatus(id) {
    if (state.labelStatusFilter === 'all') return true;
    const rec = state.records.get(id);
    if (!rec) return true;
    if (rec.labelStatus == null) return true; // unknown status → trust the scan's server filter
    if (state.labelStatusFilter === 'printed') return rec.labelStatus === LABEL_STATUS_PRINTED;
    if (state.labelStatusFilter === 'failed')  return rec.labelStatus === LABEL_STATUS_PRINT_FAILED;
    return rec.labelStatus === LABEL_STATUS_NOT_PRINTED;
  }
  function applyCarrierFilter(ids) {
    return ids.filter(id => passesCarrier(id) && passesPreOrder(id) && passesDate(id) && passesLabelStatus(id));
  }

  // ==================== MULTI-SELECT ====================
  function selectionKey(item) {
    if (item.type === 'combo') return `combo:${item.sigKey}`;
    if (item.type === 'variant') return `var:${item.productId}:${item.skuId}:${item.scenario}`;
    if (item.type === 'qty') return `qty:${item.productId}:${item.skuId ?? ''}:${item.qty}`;
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

  // Shown when user clicks a "done" card/badge/combo — choose reprint, unmark, or cancel
  function showDoneActionModal(displayLabel) {
    return new Promise(resolve => {
      const overlay = document.createElement('div');
      overlay.className = 'qf-modal-overlay';
      overlay.innerHTML = `
        <div class="qf-modal" role="dialog">
          <div class="qf-modal-title">พิมพ์ไปแล้ว</div>
          <div class="qf-modal-body">
            <div class="qf-modal-target">${escapeHtml(displayLabel || '')}</div>
            <div class="qf-modal-summary">คุณต้องการทำอะไร?</div>
          </div>
          <div class="qf-modal-actions qf-done-actions">
            <button class="qf-btn-cancel">ยกเลิก</button>
            <button class="qf-btn-unmark">ลบเครื่องหมาย ✓</button>
            <button class="qf-btn-reprint">พิมพ์ซ้ำ</button>
          </div>
        </div>
      `;
      document.body.appendChild(overlay);
      const cleanup = v => { overlay.remove(); resolve(v); };
      overlay.querySelector('.qf-btn-cancel').onclick = () => cleanup(null);
      overlay.querySelector('.qf-btn-unmark').onclick = () => cleanup('unmark');
      overlay.querySelector('.qf-btn-reprint').onclick = () => cleanup('reprint');
      overlay.onclick = e => { if (e.target === overlay) cleanup(null); };
      const onKey = e => { if (e.key === 'Escape') { cleanup(null); document.removeEventListener('keydown', onKey); } };
      document.addEventListener('keydown', onKey);
    });
  }

  // Select all visible (non-done) items in the CURRENT tab, respecting filters.
  // Used by the "เลือกทั้งหมด" button inside the select bar.
  //
  // Granularity rule: when variants are visible we insert variant-level keys so
  // the badges light up (user can then click a single badge to deselect). When
  // no variants exist, fall back to product-level. Weird tab inserts combo keys.
  //
  // Done check: labels page uses `isLabelsDone` (session-local printedUnitIds)
  // instead of `isDone` — `state.doneItems` is NEVER populated on the labels
  // page, so `isDone` always returns false and printed cards would leak into
  // the selection. See isLabelsDone comment above for the full rationale.
  function selectAllVisible() {
    const tab = state.currentTab;
    const labelsPg = isLabelsPage();
    if (tab === 'single' || tab === 'multi') {
      const idsKey = tab === 'single' ? 'fulfillUnitIdsSingle' : 'fulfillUnitIdsMulti';
      const type = tab === 'single' ? 'single_item' : 'single_sku';
      const scenario = tab === 'single' ? 'single' : 'multi';
      for (const p of state.products.values()) {
        const count = carrierFilteredSize(p[idsKey]);
        if (count === 0) continue;
        const productDone = labelsPg
          ? isLabelsDone(p[idsKey])
          : isDone(p.productId, null, type);
        if (productDone) continue;

        // Visible variants mirror the grid-render filter (skuId != null, count > 0).
        const variantList = [...(p.variants?.values() || [])].filter(v => {
          if (v.skuId == null) return false;
          return carrierFilteredSize(v[idsKey]) > 0;
        });

        if (variantList.length > 0) {
          // Insert variant keys so badges visibly reflect the selection state.
          for (const v of variantList) {
            const vDone = labelsPg
              ? isLabelsDone(v[idsKey])
              : isDone(p.productId, v.skuId, type);
            if (vDone) continue;
            const item = {type: 'variant', productId: p.productId, skuId: v.skuId, scenario};
            state.selected.set(selectionKey(item), item);
          }
        } else {
          // No variants → select at product level.
          const item = {type: 'product', productId: p.productId, scenario};
          state.selected.set(selectionKey(item), item);
        }
      }
    } else if (tab === 'weird') {
      for (const combo of state.weirdCombos.values()) {
        const count = carrierFilteredSize(combo.fulfillUnitIds);
        if (count === 0) continue;
        const comboDone = labelsPg
          ? isLabelsDone(combo.fulfillUnitIds)
          : isComboDone(combo.sigKey);
        if (comboDone) continue;
        state.selected.set(selectionKey({type:'combo', sigKey:combo.sigKey}),
                          {type:'combo', sigKey:combo.sigKey});
      }
    }
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
      } else if (item.type === 'qty') {
        const product = state.products.get(item.productId);
        if (product) {
          const bucket = product.fulfillUnitIdsByQty.get(item.qty);
          if (bucket) ids = applyCarrierFilter([...bucket]);
        }
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
      const parts = c.items.map(i => (getAlias(i.productId) || '').trim() || shortName(i.productName));
      return {label: parts.join(' + '), filename: parts.join('+')};
    }
    const p = state.products.get(item.productId);
    const alias = (getAlias(item.productId) || '').trim();
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
    const rawItems = [...state.selected.values()];
    if (!rawItems.length) { showToast('ยังไม่ได้เลือกอะไร', 1500); return; }

    // Auto-expand product-level selections into per-variant items so each
    // variant gets its own PDF (1 variant = 1 PDF rule).
    const items = [];
    for (const it of rawItems) {
      if (it.type === 'product') {
        const product = state.products.get(it.productId);
        if (!product) continue;
        const variants = [...product.variants.values()]
          .map(v => ({v, ids: collectFulfillIds(it.productId, v.skuId, it.scenario)}))
          .filter(x => x.ids.length > 0);
        if (variants.length === 0) continue;
        if (variants.length === 1) {
          items.push(it); // product has 1 variant — keep as product-level
        } else {
          // Expand into per-variant items
          for (const {v} of variants) {
            items.push({
              type: 'variant',
              productId: it.productId,
              skuId: v.skuId,
              scenario: it.scenario,
            });
          }
        }
      } else {
        items.push(it); // variant or combo — pass through
      }
    }

    if (!items.length) { showToast('รายการที่เลือกไม่มีฉลาก', 2000); return; }

    // Collect all ids and detect multi-SKU.
    // A record with >1 skus (weird combo) counts as multi-SKU on its own;
    // also count every (productId:skuId) across all records so cross-bucket mixes prompt.
    const allSelectedIds = [];
    const skuBuckets = new Set();
    let hasComboRecord = false;
    for (const it of items) {
      const ids = getItemIds(it);
      for (const id of ids) {
        allSelectedIds.push(id);
        const rec = state.records.get(id);
        if (!rec?.skuList?.length) continue;
        if (rec.skuList.length > 1) hasComboRecord = true;
        for (const s of rec.skuList) {
          skuBuckets.add(`${s.productId}:${s.skuId}`);
        }
      }
    }
    const multiSku = hasComboRecord || skuBuckets.size > 1;
    const totalIds = allSelectedIds.length;

    if (!totalIds) { showToast('รายการที่เลือกไม่มีฉลาก', 2000); return; }

    const sample = items.slice(0, 3).map(it => describeItem(it).label).join(', ')
      + (items.length > 3 ? ` +อีก ${items.length - 3}` : '');

    // Fast path (!multiSku && small batch) skips the chunk-plan modal, so we
    // offer the divider toggle inline in the confirm modal for those jobs.
    const fastPath = !multiSku && totalIds <= CHUNK_PROMPT_THRESHOLD;
    const confirm = await showPrintConfirm({
      title: multiSku ? `พิมพ์ ${skuBuckets.size} SKU` : `พิมพ์ ${items.length} ไฟล์`,
      summary: multiSku ? `${skuBuckets.size} SKU · ${totalIds} ฉลาก` : `แยก 1 ไฟล์ต่อตัวเลือก`,
      count: totalIds,
      sampleText: sample,
      offerDivider: fastPath,
    });
    if (!confirm) return;

    const { workerId, workerName, workerIcon } = confirm;

    // §3.5/§4: Use showChunkPlanModal whenever multi-SKU (so user can choose combined vs split)
    // or for large batches. Default combined=false (split per SKU).
    let plan;
    if (fastPath) {
      plan = { mode: 'single', withPickingList: loadPickingListPref(), combined: false, withDivider: !!confirm.withDivider };
    } else {
      plan = await showChunkPlanModal({
        total: totalIds,
        multiSku,
        defaultPickingList: loadPickingListPref(),
        skuCount: skuBuckets.size,  // for preview: "×skuCount ไฟล์" hint when not combined
      });
      if (!plan) return;
    }

    // §8.1: Filename using assignee bracket pattern.
    const assignee = workerName || null;

    try {
      // §4: If multi-SKU combined mode, build combined PDF.
      if (multiSku && plan.combined) {
        const groupMap = new Map();
        for (const id of allSelectedIds) {
          const rec = state.records.get(id);
          if (!rec?.skuList?.length) continue;
          const s = rec.skuList[0];
          const key = `${s.productId}:${s.skuId}`;
          if (!groupMap.has(key)) {
            const alias = (getAlias(s.productId) || '').trim();
            const variantInfo = getVariantInfo(s.productId, s.skuId);
            groupMap.set(key, {
              productId: s.productId,
              skuId: s.skuId,
              alias: alias || shortName(s.productName),
              officialName: s.productName || '',
              variantName: (variantInfo?.alias || '').trim() || (s.skuName || s.sellerSkuName || ''),
              productImageURL: s.productImageURL || null,
              ids: [],
            });
          }
          groupMap.get(key).ids.push(id);
        }
        const groups = [...groupMap.values()].sort((a, b) => a.alias.localeCompare(b.alias));
        const baseHint = `รวม ${groups.length} SKU`;
        const baseFilename = assignee
          ? makeBaseFilename(`[${baseHint}] [${assignee}]`)
          : makeBaseFilename(baseHint);

        // Slice into plan chunks; each chunk is its own combined PDF.
        // PARALLEL BUILD: each slice (combined+split=200, e.g. 5 chunks of 200)
        // produces an independent combined PDF in parallel via Promise.all.
        // The previous sequential `for` loop made 5 chunks take 5× wall time
        // — now they overlap end-to-end (generate + overlay + assemble are
        // all network/CPU-bound and benefit from concurrency). Mirrors the
        // plan-column combined path which already parallelized chunks.
        const slices = planSlice(allSelectedIds, plan);
        const chunkCount = slices.length;
        const prepProgress = showProgress(`กำลังเตรียม PDF รวม (${groups.length} SKU · ${totalIds} ฉลาก)`);
        // Shared completion counter — updates the single UI bar as parallel
        // slices land in whatever order they finish.
        let completedSlices = 0;
        const bumpProgress = () => {
          completedSlices += 1;
          const pct = (completedSlices / chunkCount) * 100;
          prepProgress.update(pct, `เสร็จ ${completedSlices}/${chunkCount}`);
        };
        let exportChunks;
        try {
          exportChunks = await Promise.all(slices.map(async (slice, i) => {
            const idx = i + 1;
            const chunkSuffix = chunkCount > 1 ? `-ชุด${idx}-${chunkCount}` : '';
            const filename = `${baseFilename}${chunkSuffix}.pdf`;
            const label = chunkCount === 1 ? 'ไฟล์เดียว' : `ชุด ${idx}/${chunkCount}`;

            // Re-group slice for this chunk's combined PDF.
            const sliceGroupMap = new Map();
            for (const id of slice) {
              const rec = state.records.get(id);
              if (!rec?.skuList?.length) continue;
              const s = rec.skuList[0];
              const key = `${s.productId}:${s.skuId}`;
              if (!sliceGroupMap.has(key)) {
                const grp = groups.find(g => g.productId === s.productId && g.skuId === s.skuId);
                if (grp) sliceGroupMap.set(key, { ...grp, ids: [] });
              }
              const grpEntry = sliceGroupMap.get(key);
              if (grpEntry) grpEntry.ids.push(id);
            }
            const sliceGroups = [...sliceGroupMap.values()].sort((a, b) => a.alias.localeCompare(b.alias));

            try {
              const { bytes } = await buildMultiSkuCombinedPdf(sliceGroups, workerName, workerIcon, plan.withPickingList, () => {}, plan.withDivider);
              bumpProgress();
              return { ids: slice, label, filename, prebuiltBytes: bytes };
            } catch (_buildErr) {
              bumpProgress();
              return { ids: slice, label, filename };
            }
          }));
          prepProgress.update(100, 'พร้อมแล้ว');
        } finally {
          // Remove immediately — runChunkedExport opens its own modal next.
          document.querySelectorAll('.qf-progress-overlay').forEach(e => e.remove());
        }

        // runChunkedExport using prebuiltBytes where available.
        const runChunks = exportChunks.map(c => ({ ids: c.ids, label: c.label, filename: c.filename, prebuiltBytes: c.prebuiltBytes }));
        const ok = await runChunkedExport(runChunks, `พิมพ์รวม ${groups.length} SKU`, {
          baseFilename,
          totalLabels: totalIds,
          workerId,
          workerName,
          workerIcon,
          assigneeKind: workerName ? 'worker' : null,
          assigneeName: workerName || null,
          withPickingList: plan.withPickingList || false,
        });
        if (ok) {
          for (const c of exportChunks) {
            for (const it of items) {
              if (it.type === 'combo') markComboDone(it.sigKey);
              else {
                const type = it.scenario === 'multi' ? 'single_sku' : 'single_item';
                markDone(it.productId, it.skuId || null, type);
              }
            }
          }
          for (const it of rawItems) {
            if (it.type === 'product') {
              const type = it.scenario === 'multi' ? 'single_sku' : 'single_item';
              markDone(it.productId, null, type);
            }
          }
          state.selected.clear();
          state.selectMode = false;
          renderAll();
        }
        return;
      }

      // Per-SKU (non-combined) or single-SKU path: 1 file per item.
      const SUB_CHUNK_THRESHOLD = CHUNK_PROMPT_THRESHOLD;
      const chunks = [];
      for (const it of items) {
        const ids = getItemIds(it);
        if (!ids.length) continue;
        const { label, filename } = describeItem(it);
        const itemAssignee = assignee;
        const baseFilename = itemAssignee
          ? makeBaseFilename(`[${filename}] [${itemAssignee}]`)
          : makeBaseFilename(filename);
        if (ids.length <= SUB_CHUNK_THRESHOLD) {
          chunks.push({ item: it, ids, label, filename: `${baseFilename}.pdf` });
        } else {
          const slices = planSlice(ids, plan);
          const subCount = slices.length;
          slices.forEach((slice, i) => {
            const idx = i + 1;
            const chunkSuffix = subCount > 1 ? `-ชุด${idx}-${subCount}` : '';
            chunks.push({
              item: it,
              ids: slice,
              label: subCount === 1 ? label : `${label} (${idx}/${subCount})`,
              filename: `${baseFilename}${chunkSuffix}.pdf`,
            });
          });
        }
      }

      if (!chunks.length) { showToast('รายการที่เลือกไม่มีฉลาก', 2000); return; }

      const exportChunks = chunks.map(c => ({ ids: c.ids, label: c.label, filename: c.filename }));
      const baseFilename = assignee
        ? makeBaseFilename(`[พิมพ์รวม-${items.length}ไฟล์] [${assignee}]`)
        : makeBaseFilename(`พิมพ์รวม-${items.length}ไฟล์`);

      const ok = await runChunkedExport(exportChunks, `พิมพ์รวม ${chunks.length} ไฟล์`, {
        baseFilename,
        totalLabels: totalIds,
        workerId,
        workerName,
        workerIcon,
        assigneeKind: workerName ? 'worker' : null,
        assigneeName: workerName || null,
        withPickingList: plan.withPickingList || false,
        withDivider: plan.withDivider || false,
      });
      if (ok) {
        for (const c of chunks) {
          const it = c.item;
          if (it.type === 'combo') markComboDone(it.sigKey);
          else {
            const type = it.scenario === 'multi' ? 'single_sku' : 'single_item';
            markDone(it.productId, it.skuId || null, type);
          }
        }
        for (const it of rawItems) {
          if (it.type === 'product') {
            const type = it.scenario === 'multi' ? 'single_sku' : 'single_item';
            markDone(it.productId, null, type);
          }
        }
        state.selected.clear();
        state.selectMode = false;
        renderAll();
      }
    } catch (e) {
      showErrorToast('พิมพ์ผิดพลาด: ' + e.message, {
        source: 'printSelected',
        error: String(e && (e.stack || e.message || e)),
        selectedCount: state.selected?.size || 0,
      });
    }
  }

  function updateSelectionBar() {
    const bar = document.getElementById('qf-select-bar');
    if (!bar) return;
    // Show bar whenever select mode is on so "เลือกทั้งหมด" is reachable
    // from a zero-selection state.
    if (!state.selectMode) { bar.style.display = 'none'; return; }
    bar.style.display = 'flex';
    const count = state.selected.size;
    const ids = count > 0 ? resolveSelectedIds() : [];
    bar.querySelector('.qf-select-bar-count').textContent =
      count > 0 ? `เลือก ${count} รายการ • ${ids.length} ฉลาก` : 'ยังไม่ได้เลือก';
    // Disable clear/print when nothing is selected
    bar.querySelector('.qf-select-bar-clear').disabled = count === 0;
    bar.querySelector('.qf-select-bar-print').disabled = count === 0;
  }

  function carrierFilteredSize(idSet) {
    if (!idSet) return 0;
    let n = 0;
    for (const id of idSet) {
      if (passesCarrier(id) && passesPreOrder(id) && passesDate(id) && passesLabelStatus(id)) n++;
    }
    return n;
  }

  // Same as carrierFilteredSize but IGNORES labelStatus filter. Used to decide
  // whether a printed card should remain visible (greyed-out) under "ยังไม่พิมพ์"
  // filter — otherwise _count drops to 0 after print and the card disappears
  // entirely instead of showing as done visual feedback.
  function carrierFilteredSizeIgnoreLabel(idSet) {
    if (!idSet) return 0;
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

  function loadLastAssignee() {
    try { return localStorage.getItem('qf_last_assignee_v1') || ''; } catch (_) { return ''; }
  }
  function saveLastAssignee(v) {
    try { localStorage.setItem('qf_last_assignee_v1', v || ''); } catch (_) {}
  }

  // showPrintConfirm — yes/no print confirm + inline options.
  //   offerDivider: when the caller will take the fast-path (single-SKU,
  //     small batch, no chunk-plan modal) set true to expose a divider
  //     checkbox. Returns confirm.withDivider for the caller's plan.
  //   Picking list is intentionally NOT offered here — it's only useful for
  //     larger multi-SKU jobs and has its own UI in the chunk-plan modal.
  function showPrintConfirm({ title, summary, count, sampleText, offerDivider = false }) {
    return new Promise(resolve => {
      const overlay = document.createElement('div');
      overlay.className = 'qf-modal-overlay';
      const overlayChecked = state.overlayEnabled ? 'checked' : '';
      const dividerChecked = offerDivider && loadDividerPref() ? 'checked' : '';
      const teams = state.teams || [];
      const hasAssignees = state.workers.length > 0 || teams.length > 0;
      const last = loadLastAssignee();
      const workerOpts = state.workers.length ? `<optgroup label="คน">${state.workers.map(w => `<option value="worker:${escapeHtml(w.id)}"${last === 'worker:' + w.id ? ' selected' : ''}>${escapeHtml(w.name)}</option>`).join('')}</optgroup>` : '';
      const teamOpts = teams.length ? `<optgroup label="ทีม">${teams.map(t => `<option value="team:${escapeHtml(t.id)}"${last === 'team:' + t.id ? ' selected' : ''}>👥  ${escapeHtml(t.name)}</option>`).join('')}</optgroup>` : '';
      const skipOpt = `<option value=""${!last ? ' selected' : ''}>— ไม่ระบุ —</option>`;
      const packerRowHtml = hasAssignees ? `
        <div class="qf-packer-row">
          <label class="qf-packer-label" for="qf-packer-select">ใครแพ็ค?</label>
          <select id="qf-packer-select" class="qf-packer-select">${workerOpts}${teamOpts}${skipOpt}</select>
        </div>
      ` : '';
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
            ${offerDivider ? `
            <label class="qf-modal-toggle">
              <input type="checkbox" class="qf-divider-toggle" ${dividerChecked}/>
              <span class="qf-modal-toggle-label">แทรกใบคั่นก่อนกลุ่มสินค้า</span>
              <span class="qf-modal-toggle-hint">ช่วยแพ็คเกอร์แยกกลุ่ม · เพิ่ม 1 หน้าต่อกลุ่ม</span>
            </label>` : ''}
            <div class="qf-modal-warn">ระบบจะส่งคำสั่งพิมพ์ทันที TikTok จะบันทึกว่าฉลากถูกพิมพ์</div>
            ${packerRowHtml}
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
      const getSelectedAssignee = () => {
        if (!hasAssignees) return { workerId: null, workerName: null, workerIcon: null };
        const select = overlay.querySelector('.qf-packer-select');
        const v = select?.value || '';
        saveLastAssignee(v);
        if (!v) return { workerId: null, workerName: null, workerIcon: null };
        const [kind, id] = v.split(':');
        if (kind === 'team') {
          const t = (state.teams || []).find(x => x.id === id);
          return { workerId: null, workerName: t?.name || null, workerIcon: '👥', teamId: t?.id || null, teamName: t?.name || null };
        }
        const w = state.workers.find(x => x.id === id);
        return { workerId: w?.id || null, workerName: w?.name || null, workerIcon: w?.icon || null };
      };
      const cleanup = (ok) => {
        overlay.remove();
        if (!ok) { resolve(false); return; }
        const a = getSelectedAssignee();
        const withDivider = offerDivider && !!overlay.querySelector('.qf-divider-toggle')?.checked;
        if (offerDivider) saveDividerPref(withDivider);
        resolve({ ok: true, workerId: a.workerId || null, workerName: a.workerName || null, workerIcon: a.workerIcon || null, teamId: a.teamId || null, teamName: a.teamName || null, withDivider });
      };
      overlay.querySelector('.qf-btn-cancel').onclick = () => cleanup(false);
      overlay.querySelector('.qf-btn-confirm').onclick = () => cleanup(true);
      overlay.onclick = (e) => { if (e.target === overlay) cleanup(false); };
      const onKey = (e) => { if (e.key === 'Escape') { cleanup(false); document.removeEventListener('keydown', onKey); } };
      document.addEventListener('keydown', onKey);
    });
  }

  async function buildChunkPdf(ids, onProgress, workerName, workerIcon) {
    const { PDFDocument } = window.PDFLib;

    // §Multi-page labels: split into segments so multi-SKU (2+ page) labels
    // land in size-1 segments where we can count their pages exactly. Single-
    // SKU runs merge into up-to-PRINT_BATCH_SIZE segments (fast path).
    const segments = splitIdsIntoApiSegments(ids);
    const segmentCount = segments.length;

    // Per-segment progress [0..1]; averaged into single bar.
    const segProgress = new Float64Array(segmentCount);
    const reportProgress = () => {
      const avg = segProgress.reduce((a, b) => a + b, 0) / segmentCount;
      onProgress(avg, segmentCount > 1 ? `ประมวลผล ${segmentCount} ชุดพร้อมกัน...` : '');
    };

    // §Note badges: build chunk-wide buyer/seller index maps so ★N numbering
    // stays consistent across all API segments within this chunk. Walk `ids`
    // in the original order so indexes match the visual sequence of pages.
    const chunkBuyerMap  = new Map();
    const chunkSellerMap = new Map();
    for (const id of ids) {
      const r = state.records.get(id);
      if (!r) continue;
      if (r.hasBuyerNote  || (r.hasNote && !r.hasSellerNote)) chunkBuyerMap.set(id, chunkBuyerMap.size + 1);
      if (r.hasSellerNote) chunkSellerMap.set(id, chunkSellerMap.size + 1);
    }
    const chunkNoteMap = { buyer: chunkBuyerMap, seller: chunkSellerMap };

    // Fire all segments in parallel — generate + download + overlay run concurrently.
    const segResults = await Promise.all(segments.map(async (segment, bi) => {
      const batch = segment.ids;
      const setP = (frac) => { segProgress[bi] = frac; reportProgress(); };
      setP(0.05);

      // §Pre-print address capture: snapshot unmasked recipient data BEFORE
      // the generate API marks these labels as TTS (which masks subsequent
      // reads forever). Best-effort, non-blocking on failure — if it fails
      // or the order-get template isn't captured yet, we proceed with print.
      try {
        await tryBackfillAddresses(batch);
      } catch (e) { /* non-fatal */ }

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
      if (resp.status === 429) {
        const err = new Error('TikTok rate limit (HTTP 429) — ลองแบ่งเป็นชุดเล็กลง หรือรอสักครู่แล้วลองใหม่');
        err.isRateLimit = true;
        throw err;
      }
      const data = await safeJson(resp, 'TikTok print');
      if (data.code !== 0) throw new Error(`print API code=${data.code} msg="${data.message || 'empty'}"`);
      const docUrl = data.data?.doc_url;

      // Optimistic update: server accepted the print for these IDs → mark them
      // as printed in local state so calendar / tab counts reflect reality
      // without waiting for a manual rescan. If a later step (overlay, merge)
      // fails, the labels are STILL printed on TikTok's side, so this is safe.
      for (const id of batch) {
        const rec = state.records.get(id);
        if (rec) rec.labelStatus = 50;
        state.printedUnitIds.add(id);
      }

      if (!docUrl) {
        console.warn('[QF] generate succeeded but no doc_url:', data);
        setP(1.0);
        return { bytes: null, pagesPerId: segment.ids.map(() => 1) };
      }

      setP(0.25);
      // IMPORTANT: use `_origFetch` (captured at document_start before TikTok's
      // own wrapper installed itself) so concurrent chunk downloads don't get
      // serialized behind TikTok's signing/queueing mutex. Without this, the
      // outer `Promise.all(chunks.map…)` + inner `Promise.all(segments.map…)`
      // effectively run one-at-a-time because TikTok's wrapper holds a lock
      // while it synchronously signs each request. `credentials: 'omit'` is
      // fine here — `docUrl` is a pre-signed TikTok CDN URL, not an API call.
      const pdfResp = await _origFetch.call(window, docUrl, { credentials: 'omit' });
      const pdfBytes = await pdfResp.arrayBuffer();
      setP(0.40);

      // Count pages now so we can derive pagesPerId for this segment AND pass
      // it to the overlay so alias text lands on the right label.
      let segPageCount = batch.length;
      try {
        const probe = await PDFDocument.load(pdfBytes);
        segPageCount = probe.getPageCount();
      } catch { /* stick with expected count */ }
      const segPagesPerId = deriveSegmentPagesPerId(segment, segPageCount);

      let modifiedBytes;
      if (state.overlayEnabled) {
        try {
          modifiedBytes = await overlayAliasOnPdf(pdfBytes, batch, (cur, totPages) => {
            segProgress[bi] = 0.40 + 0.50 * (cur / totPages);
            reportProgress();
          }, workerName, workerIcon, chunkNoteMap, segPagesPerId);
        } catch (e) {
          console.warn('[QF] overlay failed, using original:', e);
          modifiedBytes = pdfBytes;
        }
      } else {
        modifiedBytes = pdfBytes;
      }

      setP(1.0);
      return { bytes: modifiedBytes, pagesPerId: segPagesPerId };
    }));

    // Merge pages in original segment order; aggregate pagesPerId.
    const mergedDoc = await PDFDocument.create();
    let pageCount = 0;
    const pagesPerId = [];
    for (const r of segResults) {
      pagesPerId.push(...r.pagesPerId);
      if (!r.bytes) continue;
      const partDoc = await PDFDocument.load(r.bytes);
      const indices = partDoc.getPageIndices();
      const copied = await mergedDoc.copyPages(partDoc, indices);
      copied.forEach(p => mergedDoc.addPage(p));
      pageCount += copied.length;
    }

    const bytes = await mergedDoc.save();
    return { bytes, pageCount, pagesPerId };
  }

  // Sub-group ids by (productId:skuId, carrier, qty) so each SKU × carrier × qty
  // combination gets its own divider. Sorted by alias → variant → qty ASC → carrier.
  //
  // Qty dimension added so "1 SKU หลายชิ้น" (e.g. mint soap qty:2 vs qty:3) gets
  // separate dividers per qty bucket — makes packers' life easier when they need
  // to stuff N items per order.
  //
  // When all ids in a result bucket share qty=1 (or qty=null), divider renders
  // without qty suffix (backward-compat visual).
  function subGroupByCarrierAndQty(ids) {
    const bucketMap = new Map();
    for (const id of ids) {
      const rec = state.records.get(id);
      if (!rec?.skuList?.length) continue;
      const s = rec.skuList[0];
      const skuKey = `${s.productId || ''}:${s.skuId || ''}`;
      const carrierId = state.carrierOf.get(id) || 'unknown';
      const carrier = state.carriers.get(carrierId) || { name: 'ไม่ระบุ', iconUrl: '' };
      // Normalize qty: treat 0 / null / undefined as 1 (defensive).
      const qty = Math.max(1, Number(s.quantity) || 1);
      const key = `${skuKey}|${carrierId}|q${qty}`;
      if (!bucketMap.has(key)) {
        const aliasRaw = (getAlias(s.productId) || '').trim();
        const variantInfo = getVariantInfo(s.productId, s.skuId);
        bucketMap.set(key, {
          skuKey,
          carrierId,
          qty,
          alias: aliasRaw || shortName(s.productName || ''),
          officialName: s.productName || '',
          variantName: (variantInfo?.alias || '').trim() || (s.skuName || ''),
          productImageURL: s.productImageURL || null,
          carrierName: carrier.name || 'ไม่ระบุ',
          carrierIconURL: carrier.iconUrl || null,
          ids: [],
        });
      }
      bucketMap.get(key).ids.push(id);
    }
    return [...bucketMap.values()].sort((a, b) => {
      const aliasCmp = a.alias.localeCompare(b.alias, 'th');
      if (aliasCmp !== 0) return aliasCmp;
      const variantCmp = a.variantName.localeCompare(b.variantName, 'th');
      if (variantCmp !== 0) return variantCmp;
      if (a.qty !== b.qty) return a.qty - b.qty;
      return a.carrierName.localeCompare(b.carrierName, 'th');
    });
  }

  // Backward-compat alias — older callers expect the non-qty grouping.
  // Kept because callers outside the divider-insertion path (e.g. filenames)
  // may still expect flat SKU × carrier semantics.
  function subGroupByCarrier(ids) {
    // Re-bucket the qty-grouped results back to (sku, carrier) level by merging
    // qty buckets. Preserves sort order of the first qty bucket we see.
    const qtyGrouped = subGroupByCarrierAndQty(ids);
    const merged = new Map();
    for (const b of qtyGrouped) {
      const k = `${b.skuKey}|${b.carrierId}`;
      if (!merged.has(k)) {
        merged.set(k, { ...b, ids: [...b.ids] });
        delete merged.get(k).qty; // non-qty grouping drops the qty field
      } else {
        merged.get(k).ids.push(...b.ids);
      }
    }
    return [...merged.values()];
  }

  // Prepend per-subgroup divider pages to an existing label PDF. Reorders by
  // (sku, carrier) so dividers always sit directly above their labels.
  // Returns {bytes, pageCount}. idOrder must match the order ids were sent to
  // the API so page index → id can be inferred.
  async function prependDividersToChunk(labelBytes, idOrder, workerName, assigneeKind = null, pagesPerId = null) {
    const { PDFDocument } = window.PDFLib;
    // Phase 2: group by (sku, carrier, qty) — creates divider per qty bucket
    // so mixed-qty prints (e.g. soap ×2 + soap ×3) get separate headers.
    const subs = subGroupByCarrierAndQty(idOrder);
    if (subs.length === 0) return { bytes: labelBytes, pageCount: null };

    const labelDoc = await PDFDocument.load(labelBytes);
    const totalPages = labelDoc.getPageCount();
    if (totalPages === 0) return { bytes: labelBytes, pageCount: 0 };

    // Build id → [pageIndices] map. Multi-page labels (e.g. 2+ SKU orders)
    // span several pages, so each id may own a contiguous run. When caller
    // provides pagesPerId[] we use it; otherwise fall back to 1 page per id
    // (safe for historical single-SKU-only flows).
    const pageIdxById = new Map();
    {
      const perId = Array.isArray(pagesPerId) && pagesPerId.length === idOrder.length
        ? pagesPerId
        : idOrder.map(() => 1);
      let p = 0;
      for (let i = 0; i < idOrder.length && p < totalPages; i++) {
        const span = Math.max(1, perId[i] | 0);
        const indices = [];
        for (let k = 0; k < span && p < totalPages; k++, p++) indices.push(p);
        pageIdxById.set(idOrder[i], indices);
      }
    }

    const finalDoc = await PDFDocument.create();
    if (window.fontkit) finalDoc.registerFontkit(window.fontkit);
    const fontBytes = await ensureFontBytes();
    const font = await finalDoc.embedFont(fontBytes, { subset: true });

    const W = labelDoc.getPage(0).getWidth();
    const H = labelDoc.getPage(0).getHeight();

    const calendarMode = getActiveCalendarSummary();
    for (const sub of subs) {
      await buildDividerPage(finalDoc, { W, H }, {
        alias: sub.alias,
        officialName: sub.officialName,
        variantName: sub.variantName,
        productImageURL: sub.productImageURL,
        carrierName: sub.carrierName,
        carrierIconURL: sub.carrierIconURL,
        qty: sub.ids.length,
        orderQty: sub.qty || 1, // per-order item count — divider highlights when > 1
        calendarMode,
      }, font, workerName, assigneeKind);
      // Flatten every page owned by each id in this sub-group.
      const pageIdx = sub.ids.flatMap(id => pageIdxById.get(id) || []);
      if (pageIdx.length) {
        const copied = await finalDoc.copyPages(labelDoc, pageIdx);
        copied.forEach(p => finalDoc.addPage(p));
      }
    }

    const bytes = await finalDoc.save();
    return { bytes, pageCount: finalDoc.getPageCount() };
  }

  // Convert an image URL to a grayscale JPEG ArrayBuffer using an off-screen canvas.
  // Returns null on failure so callers can skip the image gracefully.
  async function toGrayscaleJpeg(url) {
    try {
      const resp = await _origFetch.call(window, url, { credentials: 'omit' });
      if (!resp.ok) return null;
      const blob = await resp.blob();
      const blobUrl = URL.createObjectURL(blob);
      return await new Promise((resolve) => {
        const img = new Image();
        img.onload = () => {
          const canvas = document.createElement('canvas');
          canvas.width = img.naturalWidth;
          canvas.height = img.naturalHeight;
          const ctx = canvas.getContext('2d');
          ctx.drawImage(img, 0, 0);
          const d = ctx.getImageData(0, 0, canvas.width, canvas.height);
          for (let i = 0; i < d.data.length; i += 4) {
            const g = Math.round(0.299 * d.data[i] + 0.587 * d.data[i + 1] + 0.114 * d.data[i + 2]);
            d.data[i] = d.data[i + 1] = d.data[i + 2] = g;
          }
          ctx.putImageData(d, 0, 0);
          canvas.toBlob(async (b) => {
            URL.revokeObjectURL(blobUrl);
            resolve(b ? await b.arrayBuffer() : null);
          }, 'image/jpeg', 0.85);
        };
        img.onerror = () => { URL.revokeObjectURL(blobUrl); resolve(null); };
        img.src = blobUrl;
      });
    } catch { return null; }
  }

  // Split text into lines that fit within maxW at the given font size (word-wrap).
  function wrapTextLines(font, text, size, maxW) {
    const words = text.split(/\s+/).filter(Boolean);
    const lines = [];
    let cur = '';
    for (const word of words) {
      const candidate = cur ? `${cur} ${word}` : word;
      if (font.widthOfTextAtSize(candidate, size) <= maxW) {
        cur = candidate;
      } else {
        if (cur) lines.push(cur);
        // If a single word is wider than maxW, let it overflow (no mid-word break).
        cur = word;
      }
    }
    if (cur) lines.push(cur);
    return lines;
  }

  // §4.5 / §4.6: Build a single divider page for the combined multi-SKU PDF.
  // payload = {alias, officialName, variantName, productImageURL, qty,
  // Split alias text into 1 or 2 lines, splitting near the visual midpoint.
  // Prefers word boundary (space); falls back to character midpoint if none found.
  function splitAliasToLines(font, text, size, maxW) {
    if (font.widthOfTextAtSize(text, size) <= maxW) return [text];
    const totalW = font.widthOfTextAtSize(text, size);
    const halfW = totalW / 2;
    let bestIdx = -1, bestDiff = Infinity;
    for (let i = 1; i < text.length; i++) {
      if (text[i] === ' ' || text[i] === '\u200b') {
        const lw = font.widthOfTextAtSize(text.slice(0, i), size);
        const diff = Math.abs(lw - halfW);
        if (diff < bestDiff) { bestDiff = diff; bestIdx = i; }
      }
    }
    if (bestIdx === -1) {
      let acc = 0;
      for (let i = 0; i < text.length; i++) {
        acc += font.widthOfTextAtSize(text[i], size);
        if (acc >= halfW) { bestIdx = i + 1; break; }
      }
      bestIdx = bestIdx || Math.floor(text.length / 2);
    }
    const l1 = text.slice(0, bestIdx).trim();
    const l2 = text.slice(bestIdx).trim();
    return l2 ? [l1, l2] : [l1];
  }

  // Render alias block (1–2 lines, dynamic size, center-aligned, no truncation).
  // Returns how many pts it consumed (updates topCursor externally via return).
  function drawAliasBlock(page, font, text, baseSize, maxW, W, topCursor, color) {
    let sz = baseSize;
    let lines = splitAliasToLines(font, text, sz, maxW);
    for (let pass = 0; pass < 4; pass++) {
      const maxLW = Math.max(...lines.map(l => font.widthOfTextAtSize(l, sz)));
      if (maxLW <= maxW) break;
      sz = sz * (maxW / maxLW) * 0.96;
      lines = splitAliasToLines(font, text, sz, maxW);
    }
    const lineGap = Math.max(2, sz * 0.14);
    const blockH = lines.length * sz + (lines.length - 1) * lineGap;
    topCursor -= blockH + 4;
    let lineY = topCursor + blockH - sz;
    for (const line of lines) {
      const lw = font.widthOfTextAtSize(line, sz);
      page.drawText(line, { x: (W - lw) / 2, y: lineY, size: sz, font, color });
      lineY -= sz + lineGap;
    }
    return topCursor - 4;
  }

  //            carrierName?, carrierIconURL?}
  // Layout redesigned (Bug #2 fix) — clean top-down stack, no overlaps.
  // ทุกอย่างขาวดำ (grayscale only).
  async function buildDividerPage(pdfDoc, { W, H }, payload, font, workerName, assigneeKind = null) {
    const { rgb } = window.PDFLib;
    const page = pdfDoc.addPage([W, H]);
    // Phase 2: use per-field config (with preset fallback). Resolve size mult per field.
    const eff = getEffectiveDividerConfig();
    const f = eff.fields;
    const mult = (k) => DIVIDER_SIZE_MAP[f[k]?.size] || 1.0;
    const vis = (k) => !!f[k]?.visible;
    // Map new config → existing cfg.* flags expected by downstream code (minimal churn).
    // name / variant / image / footer always forced on — divider must be a
    // reliable pack-check reference regardless of which preset the user chose.
    const cfg = {
      showAlias:   vis('alias'),
      showName:    true,
      showVariant: true,
      // Carrier is forced on — when groups are split by carrier the divider
      // MUST show which one this block belongs to. Regardless of preset.
      showCarrier: true,
      showImage:   true,
      showQty:     vis('qty'),
      showWorker:  vis('worker'),
      showFooter:  true,
      aliasScale:  mult('alias'),
      imageFirst:  !!eff.imageFirst,
    };
    const BLACK = rgb(0, 0, 0);
    const DARK = rgb(0.15, 0.15, 0.15);
    const MID = rgb(0.35, 0.35, 0.35);
    const LIGHT = rgb(0.55, 0.55, 0.55);

    // --- Edge margin (replaces old black banner) ---
    // Removed solid black banners (top+bottom @ 20pt each): thermal printers
    // burn through a lot of heat/ribbon on large black fills, leading to
    // dropped heads or print stalls on long jobs. Keep a small blank margin
    // so content doesn't kiss the paper edge.
    const BANNER_H = 10; // blank safety margin, no fill

    // --- Top area cursor (below top margin) ---
    // Layout: top-down stack with vertical cursor.
    const PAD_X = 12;
    const contentMaxW = W - PAD_X * 2;
    let topCursor = H - BANNER_H - 6; // y of next drawable top edge
    const imgMult = mult('image');

    // Worker/team line at top (text in grey, below banner).
    // §Team-aware: use "ทีม:" prefix when this column belongs to a team
    // so packers immediately see the label belongs to a shared workload.
    if (cfg.showWorker && workerName) {
      const wSize = 9 * mult('worker');
      topCursor -= wSize + 2;
      const prefix = assigneeKind === 'team' ? 'ทีม' : 'ผู้แพ็ค';
      page.drawText(`${prefix}: ${workerName}`, {
        x: PAD_X,
        y: topCursor,
        size: wSize,
        font,
        color: MID,
        maxWidth: contentMaxW,
      });
      topCursor -= 4;
    }

    // --- Print timestamp (always shown) ---
    {
      const now = new Date();
      const pad2 = n => String(n).padStart(2, '0');
      const tsText = `พิมพ์เมื่อ ${pad2(now.getDate())}/${pad2(now.getMonth() + 1)}/${now.getFullYear()}  ${pad2(now.getHours())}:${pad2(now.getMinutes())} น.`;
      const tsSize = 8;
      topCursor -= tsSize + 2;
      page.drawText(tsText, {
        x: PAD_X,
        y: topCursor,
        size: tsSize,
        font,
        color: LIGHT,
        maxWidth: contentMaxW,
      });
      topCursor -= 6;
    }

    // --- Calendar filter line (only when active) ---
    // Tells the packer which date-window these labels came from — matches the
    // picking-list header so both surfaces report the same source slice.
    if (payload.calendarMode) {
      const calSize = 8;
      topCursor -= calSize + 2;
      page.drawText(
        `ปฏิทิน: ${payload.calendarMode.fieldLabel} · ${payload.calendarMode.rangeText}`,
        { x: PAD_X, y: topCursor, size: calSize, font, color: MID, maxWidth: contentMaxW }
      );
      topCursor -= 6;
    }

    // --- photo-first preset renders image FIRST, at top ---
    // For other presets image is placed between carrier and qty.
    const IMG_MAX = 110 * imgMult; // max side for product image
    let topImageBottom = null;
    if (cfg.imageFirst && cfg.showImage && payload.productImageURL) {
      try {
        const grayBuf = await toGrayscaleJpeg(payload.productImageURL);
        if (grayBuf) {
          const image = await pdfDoc.embedJpg(grayBuf);
          const maxSide = 100; // photo-first: bigger
          const ratio = image.width / image.height;
          const iw = ratio >= 1 ? maxSide : maxSide * ratio;
          const ih = ratio >= 1 ? maxSide / ratio : maxSide;
          const iy = topCursor - ih - 4;
          page.drawImage(image, { x: (W - iw) / 2, y: iy, width: iw, height: ih });
          topImageBottom = iy;
          topCursor = iy - 8;
        }
      } catch { /* skip */ }
    }

    // --- Alias block: dynamic size, up to 2 lines, center-aligned, no truncation ---
    // Header now combines alias + variant as the "headline" so packers see the
    // short nickname at a glance. Official product name is deferred to the
    // image caption (rendered below image) so this top area stays clean.
    if (cfg.showAlias && payload.alias) {
      const baseSize = Math.min(H * 0.05, 20) * (cfg.aliasScale || 1);
      topCursor = drawAliasBlock(page, font, payload.alias, baseSize, contentMaxW, W, topCursor, BLACK);
    }

    // --- Variant line (13pt dark — part of the headline now, larger than before) ---
    if (cfg.showVariant && payload.variantName) {
      const varSize = 13 * mult('variant');
      const varLines = wrapTextLines(font, payload.variantName, varSize, contentMaxW).slice(0, 2);
      for (const line of varLines) {
        const lineW = font.widthOfTextAtSize(line, varSize);
        topCursor -= varSize + 2;
        page.drawText(line, {
          x: (W - lineW) / 2,
          y: topCursor,
          size: varSize,
          font,
          color: DARK,
        });
      }
      topCursor -= 4;
    }

    // --- Carrier badge — pinned to top-right corner ---
    // Was previously rendered inline in the center stack; moved out of the
    // flow so the headline (alias+variant) reads cleanly and the carrier
    // stays easy to spot at a glance. Doesn't mutate topCursor.
    if (cfg.showCarrier && payload.carrierName) {
      const carrierSize = 10 * mult('carrier');
      const carrierLabel = payload.carrierName;
      const labelW = font.widthOfTextAtSize(carrierLabel, carrierSize);
      const iconSize = 14;
      let iconImg = null;
      if (payload.carrierIconURL) {
        // Use toGrayscaleJpeg (canvas-based) — same path the product image
        // uses, which avoids CORS / mixed-content issues that broke direct
        // _origFetch + embedJpg on TikTok CDN icons.
        try {
          const iconBuf = await toGrayscaleJpeg(payload.carrierIconURL);
          if (iconBuf) iconImg = await pdfDoc.embedJpg(iconBuf);
        } catch { iconImg = null; }
      }
      const gap = iconImg ? 5 : 0;
      const iconW = iconImg ? iconSize : 0;
      const rowW = iconW + gap + labelW;
      const padX = 6;
      const boxW = rowW + padX * 2;
      const boxH = 18;
      // Anchor top-right with a small safety margin from edges.
      const boxX = W - PAD_X - boxW;
      const boxY = H - BANNER_H - boxH - 2;
      page.drawRectangle({
        x: boxX, y: boxY, width: boxW, height: boxH,
        borderColor: BLACK, borderWidth: 0.75,
        color: rgb(1, 1, 1),
      });
      let cur = boxX + padX;
      if (iconImg) {
        const r = iconImg.width / iconImg.height;
        const ih = iconSize;
        const iw = ih * r;
        page.drawImage(iconImg, { x: cur, y: boxY + (boxH - ih) / 2, width: iw, height: ih });
        cur += iw + gap;
      }
      page.drawText(carrierLabel, {
        x: cur,
        y: boxY + (boxH - carrierSize) / 2 + 1,
        size: carrierSize,
        font,
        color: BLACK,
      });
    }

    // --- Bottom-up layout: footer → qty → image above qty ---
    const FOOTER_Y = BANNER_H + 10; // above bottom margin
    // Footer is now ~75% bigger (8pt → 14pt base) so packers can read the
    // pack-check reminder from arm's length without leaning in. QTY_Y must
    // grow with footer so image column below doesn't overlap the warning.
    const footerBaseSize = 14;
    const QTY_Y = FOOTER_Y + footerBaseSize + 14;

    // Footer — "กรุณาตรวจสอบรายการสินค้าให้ถูกต้องก่อนแพ็ค"
    // Darker ink (DARK vs old LIGHT) because thermal print fades the very
    // lightest greys to nearly invisible.
    if (cfg.showFooter) {
      const footerText = 'กรุณาตรวจสอบรายการสินค้าให้ถูกต้องก่อนแพ็ค';
      const footerSize = footerBaseSize * mult('footer');
      const footerW = font.widthOfTextAtSize(footerText, footerSize);
      page.drawText(footerText, {
        x: (W - footerW) / 2,
        y: FOOTER_Y,
        size: footerSize,
        font,
        color: DARK,
      });
    }

    // Qty (always shown unless preset hides)
    // When orderQty > 1 (same SKU ordered in multi-units per order), highlight
    // the per-order qty prominently — packers must pack N pieces per waybill.
    if (cfg.showQty) {
      const orderQty = Math.max(1, Number(payload.orderQty) || 1);
      const qtyText = orderQty > 1
        ? `×${orderQty} ชิ้น/ออเดอร์ · ${payload.qty} ใบ`
        : `จำนวนทั้งหมด: ${payload.qty} ใบ`;
      const qtySize = 13 * mult('qty');
      const qtyW = font.widthOfTextAtSize(qtyText, qtySize);
      page.drawText(qtyText, {
        x: (W - qtyW) / 2,
        y: QTY_Y,
        size: qtySize,
        font,
        color: BLACK,
      });
    }

    // Product image (when NOT imageFirst) — placed between top block and qty.
    // Reserve 3 lines (≈27pt) at the bottom of the image slot for the
    // officialName caption so image + caption share the same slot cleanly.
    let imageBottomY = null;
    const captionLineH = 9;
    const captionMaxLines = 3;
    const captionReserveH = (cfg.showName && payload.officialName) ? (captionLineH * captionMaxLines + 6) : 0;
    if (!cfg.imageFirst && cfg.showImage && payload.productImageURL) {
      const availableTop = topCursor - 6;                 // top of available vertical space
      const availableBot = QTY_Y + 18 + captionReserveH;  // leave room for qty + caption
      const availH = availableTop - availableBot;
      if (availH > 30) {
        try {
          const grayBuf = await toGrayscaleJpeg(payload.productImageURL);
          if (grayBuf) {
            const image = await pdfDoc.embedJpg(grayBuf);
            const maxSide = Math.min(IMG_MAX, availH, contentMaxW);
            const ratio = image.width / image.height;
            const iw = ratio >= 1 ? maxSide : maxSide * ratio;
            const ih = ratio >= 1 ? maxSide / ratio : maxSide;
            const iy = availableBot + (availH - ih) / 2;
            page.drawImage(image, { x: (W - iw) / 2, y: iy, width: iw, height: ih });
            imageBottomY = iy;
          }
        } catch { /* skip */ }
      }
    }

    // --- Official name caption (below image, above qty) ---
    // Acts as the image's caption so the headline stays short (alias+variant)
    // while the full product name remains visible for pack-check.
    if (cfg.showName && payload.officialName) {
      const captionSize = 7.5 * mult('name');
      const captionLines = wrapTextLines(font, payload.officialName, captionSize, contentMaxW).slice(0, captionMaxLines);
      // Anchor just below the image when present; otherwise float above qty.
      const topY = imageBottomY != null
        ? imageBottomY - 4
        : QTY_Y + 18 + captionReserveH - 2;
      let y = topY;
      for (const line of captionLines) {
        y -= captionSize + 2;
        const lineW = font.widthOfTextAtSize(line, captionSize);
        page.drawText(line, {
          x: (W - lineW) / 2,
          y,
          size: captionSize,
          font,
          color: MID,
        });
      }
    }

    return page;
  }

  // §7: Build picking list pages for prepending to a label PDF.
  // Returns a temporary PDFDocument; caller copies pages into the final doc.
  async function buildPickingListPages(groups, { W, H }, headerMeta) {
    if (!window.PDFLib || !window.fontkit) return [];
    const { PDFDocument } = window.PDFLib;

    // Aggregate rows — one row per productId:skuId bucket (note orders counted together).
    const aggMap = new Map();
    for (const grp of groups) {
      const key = `${grp.productId || ''}:${grp.skuId || ''}`;
      if (!aggMap.has(key)) {
        const variantInfo = getVariantInfo(grp.productId, grp.skuId);
        aggMap.set(key, {
          alias: grp.alias || '',
          officialName: grp.officialName || '',
          variantName: grp.variantName || (variantInfo?.alias || '').trim() || '',
          qty: 0,
        });
      }
      aggMap.get(key).qty += grp.ids.length;
    }

    let no = 0;
    const rows = [...aggMap.values()]
      .sort((a, b) => a.alias.localeCompare(b.alias, 'th') || a.officialName.localeCompare(b.officialName, 'th'))
      .map(r => ({ ...r, no: ++no }));

    const tmpDoc = await PDFDocument.create();
    if (window.fontkit) tmpDoc.registerFontkit(window.fontkit);
    const fontBytes = await ensureFontBytes();
    const font = await tmpDoc.embedFont(fontBytes, { subset: true });

    // Picking-list header block grows by 14pt when calendar-mode line is shown.
    const headerHeight = 120 + (headerMeta.calendarMode ? 14 : 0);
    const colHdrHeight = 14;  // column label row
    const rowHeight    = 28;  // compact table row
    const footerHeight = 20;
    const rowsPerPage  = Math.max(1, Math.floor((H - headerHeight - colHdrHeight - footerHeight) / rowHeight));
    const pageCount    = Math.ceil(rows.length / rowsPerPage);

    for (let pi = 0; pi < pageCount; pi++) {
      const page = tmpDoc.addPage([W, H]);
      renderPickingHeader(page, headerMeta, font, W, H, pi + 1, pageCount);
      renderPickingColumnHeaders(page, W, H - headerHeight, font);

      const pageRows = rows.slice(pi * rowsPerPage, (pi + 1) * rowsPerPage);
      let curY = H - headerHeight - colHdrHeight;
      for (const row of pageRows) {
        renderPickingRowTable(page, row, curY, W, font);
        curY -= rowHeight;
      }
    }

    return tmpDoc;
  }

  function renderPickingHeader(page, headerMeta, font, W, H, pageNum, pageCount) {
    const { rgb } = window.PDFLib;
    const black = rgb(0, 0, 0);
    const grey = rgb(0.4, 0.4, 0.4);

    // Title
    page.drawText('Picking List', { x: 24, y: H - 36, size: 24, font, color: black });

    // User
    const userText = `User: ${maskEmail(headerMeta.userEmail)}`;
    page.drawText(userText, { x: 24, y: H - 56, size: 10, font, color: grey, maxWidth: W - 48 });

    // Print time
    const d = headerMeta.printedAt || new Date();
    const pad = n => String(n).padStart(2, '0');
    const dateStr = `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
    page.drawText(`Print time: ${dateStr}`, { x: 24, y: H - 70, size: 10, font, color: grey, maxWidth: W - 48 });

    // Totals
    const totalsText = `Orders: ${headerMeta.totalOrders || 0}  Products: ${headerMeta.totalProducts || 0}  Items: ${headerMeta.totalItems || 0}`;
    page.drawText(totalsText, { x: 24, y: H - 84, size: 10, font, color: grey, maxWidth: W - 48 });

    // Assignee
    let yCursor = H - 98;
    if (headerMeta.assigneeName) {
      page.drawText(`รับผิดชอบ: ${headerMeta.assigneeName}`, { x: 24, y: yCursor, size: 10, font, color: black, maxWidth: W - 48 });
      yCursor -= 14;
    }

    // Calendar-mode line: only when user has an active date filter.
    // Shows which field was used (วันที่ต้องส่ง / ลูกค้าสั่ง / ยกเลิกอัตโนมัติ)
    // plus the selected date range, so packers know why this picking list
    // contains exactly these orders.
    if (headerMeta.calendarMode) {
      const { fieldLabel, rangeText } = headerMeta.calendarMode;
      page.drawText(`ปฏิทิน: ${fieldLabel} · ${rangeText}`, {
        x: 24, y: yCursor, size: 10, font, color: black, maxWidth: W - 48,
      });
      yCursor -= 14;
    }

    // Separator line — positioned below the last drawn line.
    const sepY = Math.min(H - 108, yCursor + 4);
    page.drawLine({ start: { x: 12, y: sepY }, end: { x: W - 12, y: sepY }, thickness: 0.5, color: grey });

    // Page footer
    page.drawText(`${pageNum}/${pageCount}`, { x: W - 50, y: 18, size: 9, font, color: grey });
  }

  // Simple word-wrap: break text into up to maxLines lines that fit within maxW.
  // Breaks on space first; falls back to hard-cut for long continuous strings.
  // Last line truncates with '…' if overflowing.
  function wrapTextToLines(text, maxLines, font, size, maxW) {
    if (!text) return [];
    const words = String(text).split(/(\s+)/); // keep spaces so we can rejoin cleanly
    const lines = [];
    let buf = '';
    const flush = () => { if (buf.length) lines.push(buf); buf = ''; };
    for (const w of words) {
      const cand = buf + w;
      if (font.widthOfTextAtSize(cand, size) <= maxW) { buf = cand; continue; }
      // Doesn't fit. If buf already has content, wrap and start a new line.
      if (buf.trim().length) {
        flush();
        if (lines.length >= maxLines) break;
        buf = w.replace(/^\s+/, '');
        if (font.widthOfTextAtSize(buf, size) > maxW) {
          // Still too big → hard-cut.
          while (buf.length > 1 && font.widthOfTextAtSize(buf, size) > maxW) buf = buf.slice(0, -1);
        }
      } else {
        // Single word wider than maxW — hard-cut it.
        let cut = w;
        while (cut.length > 1 && font.widthOfTextAtSize(cut, size) > maxW) cut = cut.slice(0, -1);
        buf = cut;
      }
    }
    if (buf.length) flush();
    if (lines.length > maxLines) {
      const kept = lines.slice(0, maxLines);
      // Append ellipsis to last line, trim to fit.
      let last = kept[maxLines - 1] + '…';
      while (last.length > 1 && font.widthOfTextAtSize(last, size) > maxW) last = last.slice(0, -2) + '…';
      kept[maxLines - 1] = last;
      return kept;
    }
    return lines;
  }

  // Column header bar for picking list table.
  function renderPickingColumnHeaders(page, W, y, font) {
    const { rgb } = window.PDFLib;
    const PAD = 10;
    const rowH = 14;
    const grey = rgb(0.5, 0.5, 0.5);
    const nameW = Math.max(60, W - PAD * 2 - 16 - 70 - 30);
    page.drawRectangle({ x: 0, y: y - rowH, width: W, height: rowH, color: rgb(0.92, 0.92, 0.92) });
    page.drawText('#',         { x: PAD,                   y: y - 10, size: 7, font, color: grey });
    page.drawText('ชื่อสินค้า / ชื่อย่อ', { x: PAD + 16,              y: y - 10, size: 7, font, color: grey });
    page.drawText('ตัวเลือก',  { x: PAD + 16 + nameW,      y: y - 10, size: 7, font, color: grey, maxWidth: 66 });
    page.drawText('จำนวน',     { x: W - PAD - 28,          y: y - 10, size: 7, font, color: grey });
    page.drawLine({ start: { x: 0, y: y - rowH }, end: { x: W, y: y - rowH }, thickness: 0.5, color: grey });
  }

  // Unified compact table row for picking list (replaces tabular + compact renderers).
  function renderPickingRowTable(page, row, y, W, font) {
    const { rgb } = window.PDFLib;
    const black = rgb(0, 0, 0);
    const grey  = rgb(0.5, 0.5, 0.5);
    const rowH  = 28;
    const PAD   = 10;
    const nameW = Math.max(60, W - PAD * 2 - 16 - 70 - 30);

    if (row.no % 2 === 0) {
      page.drawRectangle({ x: 0, y: y - rowH, width: W, height: rowH, color: rgb(0.975, 0.975, 0.975) });
    }

    // Row number
    page.drawText(String(row.no), { x: PAD, y: y - 10, size: 8, font, color: rgb(0.65, 0.65, 0.65) });

    // Alias (9pt, truncate to nameW)
    const aliasRaw = row.alias || shortName(row.officialName || '');
    let aliasText = aliasRaw;
    while (aliasText.length > 1 && font.widthOfTextAtSize(aliasText, 9) > nameW - 2) aliasText = aliasText.slice(0, -1);
    page.drawText(aliasText, { x: PAD + 16, y: y - 10, size: 9, font, color: black });

    // Official name (7pt, 1 line)
    const nameLines = wrapTextToLines(row.officialName || '', 1, font, 7, nameW - 2);
    if (nameLines.length) page.drawText(nameLines[0], { x: PAD + 16, y: y - 19, size: 7, font, color: grey });

    // Variant / ตัวเลือก (7pt)
    const varText = row.variantName || '—';
    let vt = varText;
    while (vt.length > 1 && font.widthOfTextAtSize(vt, 7) > 64) vt = vt.slice(0, -1);
    page.drawText(vt, { x: PAD + 16 + nameW, y: y - 10, size: 7, font, color: grey });

    // Qty (10pt, right-aligned)
    const qStr = String(row.qty);
    const qW = font.widthOfTextAtSize(qStr, 10);
    page.drawText(qStr, { x: W - PAD - qW, y: y - 11, size: 10, font, color: black });

    page.drawLine({ start: { x: 0, y: y - rowH }, end: { x: W, y: y - rowH }, thickness: 0.25, color: rgb(0.88, 0.88, 0.88) });
  }

  // §7.3: Aggregate groups into PickingListData for headerMeta computation.
  function buildPickingHeaderMeta(groups, assigneeName) {
    const orderIds = new Set();
    const productIds = new Set();
    let totalItems = 0;
    for (const grp of groups) {
      productIds.add(`${grp.productId}:${grp.skuId || ''}`);
      for (const id of grp.ids) {
        const rec = state.records.get(id);
        const oid = rec?.orderIds?.[0];
        if (oid) orderIds.add(oid);
        for (const s of (rec?.skuList || [])) totalItems += (s.quantity || 1);
      }
    }
    return {
      userEmail: state.sellerEmail || null,
      printedAt: new Date(),
      assigneeName: assigneeName || null,
      totalOrders: orderIds.size,
      totalProducts: productIds.size,
      totalItems,
      calendarMode: getActiveCalendarSummary(),
    };
  }

  // Summarize the active calendar-filter selection for the picking list header.
  // Returns null when no date filter is active; otherwise
  // { fieldLabel, rangeText } with Thai short-date formatting.
  function getActiveCalendarSummary() {
    const df = state.dateFilter;
    if (!df || (df.start == null && df.end == null)) return null;
    const fieldLabel = FIELD_LABELS[df.field] || FIELD_LABELS.createTime;
    const fmt = (ts) => new Date(ts).toLocaleDateString('th-TH', {
      day: 'numeric', month: 'short', year: 'numeric',
    });
    let rangeText;
    if (df.start != null && df.end != null) {
      const endInclusive = df.end - 86400000;
      rangeText = (df.start === endInclusive) ? fmt(df.start) : `${fmt(df.start)} – ${fmt(endInclusive)}`;
    } else if (df.start != null) {
      rangeText = fmt(df.start);
    } else {
      rangeText = `ก่อน ${fmt(df.end - 86400000)}`;
    }
    return { fieldLabel, rangeText };
  }

  // §7.4: Insert note-zone divider page(s) before note-order labels.
  // Switches per-section to compact inline layout when notes > threshold,
  // and spawns continuation pages on overflow so nothing is truncated.
  // payload = { alias, officialName, variantName, carrierName, buyerNotes, sellerNotes }
  async function buildNoteZoneDividerPage(pdfDoc, { W, H }, payload, font) {
    const { rgb } = window.PDFLib;
    const BLACK = rgb(0, 0, 0);
    const DARK  = rgb(0.15, 0.15, 0.15);
    const MID   = rgb(0.4, 0.4, 0.4);
    const LIGHT = rgb(0.6, 0.6, 0.6);
    const PAD   = 14;
    const FOOTER_RESERVE = 18; // keep room for timestamp footer
    const COMPACT_THRESHOLD = 6; // per-section: >N → inline reflow
    const contentW = W - PAD * 2;

    // Back-compat: older callers may pass `notes` (buyer-only).
    const buyerNotes  = payload.buyerNotes  || payload.notes || [];
    const sellerNotes = payload.sellerNotes || [];

    let page = pdfDoc.addPage([W, H]);
    let y;

    const drawFooter = (pg) => {
      const now = new Date();
      const p2  = n => String(n).padStart(2, '0');
      pg.drawText(
        `พิมพ์เมื่อ ${p2(now.getDate())}/${p2(now.getMonth()+1)}/${now.getFullYear()} ${p2(now.getHours())}:${p2(now.getMinutes())} น.`,
        { x: PAD, y: 8, size: 7, font, color: LIGHT }
      );
    };

    const newPage = () => {
      drawFooter(page);
      page = pdfDoc.addPage([W, H]);
      y = H - 14;
      page.drawText('★ หมายเหตุ (ต่อ)', { x: PAD, y, size: 9, font, color: MID });
      y -= 11;
      page.drawLine({ start: { x: PAD, y }, end: { x: W - PAD, y }, thickness: 0.5, color: LIGHT });
      y -= 10;
    };

    const ensureSpace = (needed) => {
      if (y - needed < FOOTER_RESERVE) newPage();
    };

    // ── Page 1 header ─────────────────────────────────────────────────────────
    const TITLE_SIZE = 10;
    y = H - 14;
    page.drawText('★ หมายเหตุ', { x: PAD, y, size: TITLE_SIZE, font, color: MID });
    y -= TITLE_SIZE + 4;

    if (payload.alias) {
      const aliasBase = Math.min(H * 0.04, 16);
      y = drawAliasBlock(page, font, payload.alias, aliasBase, contentW, W, y, BLACK);
      y -= 2;
    }

    const productParts = [payload.officialName, payload.variantName, payload.carrierName].filter(Boolean);
    if (productParts.length) {
      page.drawText(productParts.join(' · '), { x: PAD, y, size: 7.5, font, color: DARK, maxWidth: contentW });
      y -= 10;
    }

    const noteCount = buyerNotes.length + sellerNotes.length;
    const countParts = [];
    if (buyerNotes.length)  countParts.push(`ลูกค้า ${buyerNotes.length}`);
    if (sellerNotes.length) countParts.push(`ร้านค้า ${sellerNotes.length}`);
    const countText = countParts.length ? `${noteCount} รายการ · ${countParts.join(' · ')}` : `${noteCount} ออเดอร์`;
    page.drawText(countText, { x: PAD, y, size: 8, font, color: MID });
    y -= 8;

    page.drawLine({ start: { x: PAD, y }, end: { x: W - PAD, y }, thickness: 1, color: BLACK });
    y -= 12;

    // Verbose mode (≤threshold): one note per row, marker + wrapped message.
    const drawSectionVerbose = (heading, prefix, notes) => {
      if (!notes.length) return;
      ensureSpace(11);
      page.drawText(heading, { x: PAD, y, size: 9, font, color: BLACK });
      y -= 11;
      for (const note of notes) {
        const marker = `${prefix}★${note.idx}`;
        const markerW = font.widthOfTextAtSize(marker, 10) + 6;
        const msg = note.msg ? `"${note.msg}"` : '—';
        const msgLines = wrapTextToLines(msg, 3, font, 8.5, contentW - markerW);
        const blockH = Math.max(11 * msgLines.length, 11) + 3;
        ensureSpace(blockH);
        page.drawText(marker, { x: PAD, y, size: 10, font, color: BLACK });
        let lineY = y;
        for (const line of msgLines) {
          page.drawText(line, { x: PAD + markerW, y: lineY, size: 8.5, font, color: DARK });
          lineY -= 11;
        }
        y = Math.min(lineY, y - 11) - 3;
      }
      y -= 4;
    };

    // Compact inline mode (>threshold): tokens flow with separator, wrap as paragraph.
    const drawSectionInline = (heading, prefix, notes) => {
      if (!notes.length) return;
      ensureSpace(11);
      page.drawText(heading, { x: PAD, y, size: 9, font, color: BLACK });
      y -= 11;
      const SIZE   = 8.5;
      const LINE_H = 11;
      const SEP    = ' · ';
      let curLine = '';
      const flush = () => {
        if (!curLine) return;
        ensureSpace(LINE_H);
        page.drawText(curLine, { x: PAD, y, size: SIZE, font, color: DARK });
        y -= LINE_H;
        curLine = '';
      };
      for (const note of notes) {
        const msg = (note.msg || '—').replace(/\s+/g, ' ').trim();
        const token = `${prefix}★${note.idx} ${msg}`;
        const candidate = curLine ? `${curLine}${SEP}${token}` : token;
        if (font.widthOfTextAtSize(candidate, SIZE) <= contentW) {
          curLine = candidate;
          continue;
        }
        flush();
        // Single token wider than contentW → hard-wrap across lines.
        if (font.widthOfTextAtSize(token, SIZE) > contentW) {
          const sub = wrapTextToLines(token, 99, font, SIZE, contentW);
          for (let i = 0; i < sub.length - 1; i++) {
            ensureSpace(LINE_H);
            page.drawText(sub[i], { x: PAD, y, size: SIZE, font, color: DARK });
            y -= LINE_H;
          }
          curLine = sub[sub.length - 1] || '';
        } else {
          curLine = token;
        }
      }
      flush();
      y -= 4;
    };

    const drawSection = (heading, prefix, notes) => {
      if (notes.length > COMPACT_THRESHOLD) drawSectionInline(heading, prefix, notes);
      else                                   drawSectionVerbose(heading, prefix, notes);
    };

    drawSection('หมายเหตุลูกค้า', 'ล', buyerNotes);
    drawSection('หมายเหตุร้านค้า', 'ร', sellerNotes);

    drawFooter(page);
    return page;
  }

  // §4.7: Build a combined multi-SKU PDF: [divider_1][labels_1][divider_2][labels_2]…
  // groups[] shape: [{productId, skuId, alias, officialName, variantName, productImageURL, ids: string[]}]
  //   sorted by alias ascending by caller.
  // assigneeKind ('worker'|'team'|null) lets the divider render "ทีม:" vs "ผู้แพ็ค:".
  async function buildMultiSkuCombinedPdf(groups, workerName, workerIcon, withPickingList, onProgress, withDivider = true, assigneeKind = null) {
    if (!window.PDFLib || !window.fontkit) throw new Error('PDFLib/fontkit ไม่พร้อมใช้งาน');
    const { PDFDocument } = window.PDFLib;

    const finalDoc = await PDFDocument.create();
    if (window.fontkit) finalDoc.registerFontkit(window.fontkit);
    const fontBytes = await ensureFontBytes();
    const font = await finalDoc.embedFont(fontBytes, { subset: true });

    // Probe page size from first group's labels.
    let W = 288; // A6 fallback width (pt)
    let H = 432; // A6 fallback height (pt)
    let firstGroupBytes = null;
    let firstGroupPagesPerId = null;
    try {
      const res = await callGenerateApiRaw(groups[0].ids);
      firstGroupBytes = res.bytes;
      firstGroupPagesPerId = res.pagesPerId;
      const sampleDoc = await PDFDocument.load(firstGroupBytes);
      const samplePage = sampleDoc.getPage(0);
      W = samplePage.getWidth();
      H = samplePage.getHeight();
    } catch (_probErr) {
      // Use A6 fallback — firstGroupBytes may still be valid or null.
    }

    // §7: Prepend picking list pages when requested.
    if (withPickingList) {
      try {
        const headerMeta = buildPickingHeaderMeta(groups, workerName);
        const tmpDoc = await buildPickingListPages(groups, { W, H }, headerMeta);
        if (tmpDoc) {
          const copied = await finalDoc.copyPages(tmpDoc, tmpDoc.getPageIndices());
          copied.forEach(p => finalDoc.addPage(p));
        }
      } catch (_plErr) {
        console.warn('[QF] picking list build failed, continuing without:', _plErr);
      }
    }

    for (let gi = 0; gi < groups.length; gi++) {
      const grp = groups[gi];

      // When divider is on, sub-group by (carrier, qty) so each
      // (SKU × carrier × qty) combination gets its own divider + block of labels
      // — กันแพ็คผิดเมื่อ SKU เดียวกันมีหลายขนส่ง / หลาย qty
      // (e.g. สบู่มิ้น ×2 กับ สบู่มิ้น ×3 คนละหน้าคั่น).
      const subs = withDivider ? subGroupByCarrierAndQty(grp.ids) : null;
      const parts = (subs && subs.length > 0)
        ? subs.map(s => ({
            ids: s.ids,
            alias: s.alias || grp.alias,
            officialName: s.officialName || grp.officialName,
            variantName: s.variantName || grp.variantName,
            productImageURL: s.productImageURL || grp.productImageURL,
            carrierName: s.carrierName,
            carrierIconURL: s.carrierIconURL,
            orderQty: s.qty || 1,
          }))
        : [{
            ids: grp.ids,
            alias: grp.alias,
            officialName: grp.officialName,
            variantName: grp.variantName,
            productImageURL: grp.productImageURL,
            carrierName: null,
            carrierIconURL: null,
            orderQty: 1,
          }];

      for (let pi = 0; pi < parts.length; pi++) {
        const part = parts[pi];

        // Split IDs: normal first, noted last. An id is "noted" if it has
        // EITHER a buyer note OR a seller note (or both). Each side has its
        // own independent 1..N index space so the divider page can list them
        // separately without number collisions.
        const noteIds   = part.ids.filter(id => {
          const r = state.records.get(id);
          return r?.hasBuyerNote || r?.hasSellerNote || r?.hasNote;
        });
        const normalIds = part.ids.filter(id => {
          const r = state.records.get(id);
          return !(r?.hasBuyerNote || r?.hasSellerNote || r?.hasNote);
        });
        const sortedIds = [...normalIds, ...noteIds];
        const buyerMap  = new Map();
        const sellerMap = new Map();
        for (const id of noteIds) {
          const r = state.records.get(id);
          if (r?.hasBuyerNote || (r?.hasNote && !r?.hasSellerNote)) buyerMap.set(id, buyerMap.size + 1);
          if (r?.hasSellerNote) sellerMap.set(id, sellerMap.size + 1);
        }
        const noteMap = { buyer: buyerMap, seller: sellerMap };
        const hasNotes = noteIds.length > 0;

        // Fetch label bytes (reuse first-group probe when no note re-ordering needed).
        let labelsBytes;
        let labelsPagesPerId = null;
        try {
          if (gi === 0 && pi === 0 && firstGroupBytes && part.ids.length === grp.ids.length && !hasNotes) {
            labelsBytes = firstGroupBytes;
            labelsPagesPerId = firstGroupPagesPerId;
          } else {
            const res = await callGenerateApiRaw(sortedIds);
            labelsBytes = res.bytes;
            labelsPagesPerId = res.pagesPerId;
          }
        } catch (e) {
          throw new Error(`กลุ่ม "${part.alias}" (${gi + 1}/${groups.length}) ดึง labels ไม่ได้: ${e.message}`);
        }

        // Apply alias overlay (pass noteMap so ★N appears on note-order labels).
        if (state.overlayEnabled && labelsBytes) {
          try {
            labelsBytes = await overlayAliasOnPdf(labelsBytes, sortedIds, () => {}, workerName, workerIcon, noteMap, labelsPagesPerId);
          } catch (_oErr) { /* continue unmodified */ }
        }

        // SKU divider (always before this block of labels).
        if (withDivider) {
          await buildDividerPage(finalDoc, { W, H }, {
            alias: part.alias,
            officialName: part.officialName,
            variantName: part.variantName,
            productImageURL: part.productImageURL,
            carrierName: part.carrierName,
            carrierIconURL: part.carrierIconURL,
            qty: sortedIds.length,
            orderQty: part.orderQty || 1,
            calendarMode: getActiveCalendarSummary(),
          }, font, workerName, assigneeKind);
        }

        if (labelsBytes) {
          try {
            const partDoc = await PDFDocument.load(labelsBytes);
            if (!hasNotes) {
              // No notes — copy all pages as before.
              const copied = await finalDoc.copyPages(partDoc, partDoc.getPageIndices());
              copied.forEach(p => finalDoc.addPage(p));
            } else {
              // §Multi-page labels: labels may span >1 page each, so the
              // normal/note boundary is not simply normalIds.length. Use
              // pagesPerId to compute actual page index ranges per id.
              const perId = (Array.isArray(labelsPagesPerId) && labelsPagesPerId.length === sortedIds.length)
                ? labelsPagesPerId
                : sortedIds.map(() => 1);
              const pageRangeFor = (startIdIdx, endIdIdx) => {
                let startPage = 0;
                for (let k = 0; k < startIdIdx; k++) startPage += Math.max(1, perId[k] | 0);
                let count = 0;
                for (let k = startIdIdx; k < endIdIdx; k++) count += Math.max(1, perId[k] | 0);
                return Array.from({ length: count }, (_, i) => startPage + i);
              };

              // Copy normal-order pages first (spanning 0..normalIds.length-1).
              if (normalIds.length > 0) {
                const normalIdx = pageRangeFor(0, normalIds.length);
                if (normalIdx.length > 0) {
                  const copiedNormal = await finalDoc.copyPages(partDoc, normalIdx);
                  copiedNormal.forEach(p => finalDoc.addPage(p));
                }
              }
              // Insert note-zone divider page. Buyer & seller notes are
              // listed in SEPARATE sections with their own indexes so the ★N
              // on each label maps back to exactly one message.
              const buyerNotes = [];
              const sellerNotes = [];
              for (const id of noteIds) {
                const r = state.records.get(id);
                const bIdx = buyerMap.get(id);
                if (bIdx != null && r?.buyerNote) buyerNotes.push({ idx: bIdx, msg: r.buyerNote });
                const sIdx = sellerMap.get(id);
                if (sIdx != null && r?.sellerNote) sellerNotes.push({ idx: sIdx, msg: r.sellerNote });
              }
              await buildNoteZoneDividerPage(finalDoc, { W, H }, {
                alias: part.alias,
                officialName: part.officialName,
                variantName: part.variantName,
                carrierName: part.carrierName,
                buyerNotes,
                sellerNotes,
              }, font);
              // Copy note-order pages after zone divider.
              const noteIdx = pageRangeFor(normalIds.length, sortedIds.length);
              if (noteIdx.length > 0) {
                const copiedNotes = await finalDoc.copyPages(partDoc, noteIdx);
                copiedNotes.forEach(p => finalDoc.addPage(p));
              }
            }
          } catch (_copyErr) { /* skip malformed */ }
        }
      }

      if (onProgress) onProgress((gi + 1) / groups.length);
    }

    const bytes = await finalDoc.save();
    return { bytes, pageCount: finalDoc.getPageCount() };
  }

  // Helper: call TikTok's generate API for a list of fulfillUnitIds.
  // Splits into segments that respect multi-SKU (multi-page) labels so we can
  // track the exact page count per id. Returns { bytes, pagesPerId }.
  async function callGenerateApiRaw(ids) {
    const { PDFDocument } = window.PDFLib;
    const segments = splitIdsIntoApiSegments(ids);
    const segResults = await Promise.all(segments.map(async (segment) => {
      const batch = segment.ids;
      const body = {
        fulfill_unit_id_list: batch,
        content_type_list: [1, 2],
        template_type: 0,
        op_scene: 2,
        file_prefix: 'Shipping label',
        request_time: Date.now(),
        print_option: { tmpl: 0, template_size: 0, layout: [0] },
        print_source: 201,
      };
      const resp = await _origFetch.call(window, '/api/v1/fulfillment/shipping_doc/generate', {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify(body),
      });
      if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
      const data = await safeJson(resp, 'generate API');
      if (data.code !== 0) throw new Error(`code=${data.code} msg="${data.message || ''}"`);
      const docUrl = data.data?.doc_url;
      if (!docUrl) throw new Error('ไม่มี doc_url ใน response');

      // Mark as printed in local state.
      for (const id of batch) {
        const rec = state.records.get(id);
        if (rec) rec.labelStatus = 50;
        state.printedUnitIds.add(id);
      }

      const pdfResp = await _origFetch.call(window, docUrl);
      if (!pdfResp.ok) throw new Error(`PDF fetch HTTP ${pdfResp.status}`);
      const bytes = await pdfResp.arrayBuffer();
      let pageCount = batch.length;
      try {
        const probe = await PDFDocument.load(bytes);
        pageCount = probe.getPageCount();
      } catch { /* fall through with expected count */ }
      return { bytes, pagesPerId: deriveSegmentPagesPerId(segment, pageCount) };
    }));

    const pagesPerId = segResults.flatMap(r => r.pagesPerId);
    if (segResults.length === 1) return { bytes: segResults[0].bytes, pagesPerId };
    const merged = await PDFDocument.create();
    for (const r of segResults) {
      const part = await PDFDocument.load(r.bytes);
      const copied = await merged.copyPages(part, part.getPageIndices());
      copied.forEach(p => merged.addPage(p));
    }
    return { bytes: await merged.save(), pagesPerId };
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

    // §3.5: Use showChunkPlanModal when multi-SKU OR above threshold.
    // Detect multi-SKU first so we know which path to take.
    const skuBuckets = new Set();
    let hasComboRecord = false;
    for (const id of ids) {
      const rec = state.records.get(id);
      if (!rec?.skuList?.length) continue;
      if (rec.skuList.length > 1) hasComboRecord = true;
      for (const s of rec.skuList) {
        skuBuckets.add(`${s.productId}:${s.skuId}`);
      }
    }
    const multiSku = hasComboRecord || skuBuckets.size > 1;
    const fastPath = !multiSku && ids.length <= CHUNK_PROMPT_THRESHOLD;

    const confirm = await showPrintConfirm({
      title: displayLabel || 'พิมพ์ฉลาก',
      count: ids.length,
      sampleText,
      offerDivider: fastPath,
    });
    if (!confirm) return false;
    const { workerId, workerName, workerIcon } = confirm;

    let plan;
    if (fastPath) {
      plan = { mode: 'single', withPickingList: loadPickingListPref(), combined: false, withDivider: !!confirm.withDivider };
    } else {
      plan = await showChunkPlanModal({ total: ids.length, multiSku, defaultPickingList: loadPickingListPref() });
      if (!plan) return false;
    }

    // §3.5: Map ChunkPlan → chunks[]
    const total = ids.length;
    const allIds = ids;
    let slices;
    switch (plan.mode) {
      case 'even': {
        const n = plan.n || 1;
        const sz = Math.ceil(total / n);
        slices = [];
        for (let i = 0; i < allIds.length; i += sz) slices.push(allIds.slice(i, i + sz));
        break;
      }
      case 'every': {
        const x = plan.x || CHUNK_AUTO_SAFE_SIZE;
        slices = [];
        for (let i = 0; i < allIds.length; i += x) slices.push(allIds.slice(i, i + x));
        break;
      }
      default: // 'single'
        slices = [allIds];
        break;
    }
    const chunkCount = slices.length;

    // §8.1: Build filename with assignee bracket pattern.
    const assignee = workerName || null;
    const baseHint = filenameHint || displayLabel;
    const baseFilename = assignee
      ? makeBaseFilename(`[${baseHint}] [${assignee}]`)
      : makeBaseFilename(baseHint);

    const chunks = slices.map((slice, i) => {
      const idx = i + 1;
      const chunkSuffix = chunkCount > 1 ? `-ชุด${idx}-${chunkCount}` : '';
      return {
        ids: slice,
        filename: `${baseFilename}${chunkSuffix}.pdf`,
        label: chunkCount === 1 ? 'ไฟล์เดียว' : `ชุด ${idx}/${chunkCount}`,
      };
    });

    return runChunkedExport(chunks, displayLabel || 'พิมพ์ฉลาก', {
      baseFilename,
      totalLabels: ids.length,
      workerId,
      workerName,
      workerIcon,
      assigneeKind: workerName ? 'worker' : null,
      assigneeName: workerName || null,
      withPickingList: plan.withPickingList || false,
      withDivider: plan.withDivider || false,
    });
  }

  // runChunkedExport(chunks, title, historyMeta)
  //   historyMeta = null → skip history (e.g., re-download from history itself)
  //   historyMeta = {baseFilename, totalLabels, workerId?, workerName?, withPickingList?} → save entry after allDone
  async function runChunkedExport(chunks, displayTitle, historyMeta) {
    const workerName = historyMeta?.workerName || null;
    const workerIcon = historyMeta?.workerIcon || null;
    const withPickingList = historyMeta?.withPickingList || false;
    const withDivider = historyMeta?.withDivider || false;
    const assigneeKind = historyMeta?.assigneeKind || null;
    const assigneeName = historyMeta?.assigneeName || workerName || null;

    // When divider is on, reorder each chunk's ids by (sku, carrier) so the
    // buildChunkPdf output already has carrier blocks contiguous — we can
    // then slice + prepend dividers afterwards by page index.
    if (withDivider) {
      for (const c of chunks) {
        if (c.prebuiltBytes) continue; // combined path already has dividers baked in
        const subs = subGroupByCarrier(c.ids);
        if (subs.length > 0) {
          c.ids = subs.flatMap(s => s.ids);
        }
      }
    }
    const totalIds = chunks.reduce((a, c) => a + c.ids.length, 0);
    const result = showChunkedResult({
      title: displayTitle,
      totalIds,
      chunks: chunks.map(c => ({count: c.ids.length, label: c.label, filename: c.filename})),
    });

    // §Parallelism diagnostic — emit per-chunk timing so user can verify in DevTools
    // that all chunks fire concurrently (start deltas should be <100ms between chunks).
    const runStartMs = performance.now();
    console.log(`[QF] runChunkedExport: ${chunks.length} chunks start (parallel)`);

    const runChunk = async (ci) => {
      const chunkStartMs = performance.now();
      console.log(`[QF]   chunk ${ci + 1}/${chunks.length} start @ +${(chunkStartMs - runStartMs).toFixed(0)}ms`);
      result.startChunk(ci);
      try {
        let bytes, pageCount;
        const prebuilt = chunks[ci].prebuiltBytes;
        if (prebuilt) {
          // Combined path: buildMultiSkuCombinedPdf already baked in
          // dividers + merged all SKUs — use bytes as-is.
          bytes = prebuilt;
          try {
            const { PDFDocument } = window.PDFLib;
            const d = await PDFDocument.load(bytes);
            pageCount = d.getPageCount();
          } catch { pageCount = null; }
          result.updateChunkProgress(ci, 100, 'พร้อมแล้ว');
        } else {
          // Per-chunk worker identity (by-person/by-person-sku set these per chunk).
          const chunkWorkerName = chunks[ci].workerName || workerName;
          const chunkWorkerIcon = chunks[ci].workerIcon || workerIcon;
          let chunkPagesPerId;
          ({ bytes, pageCount, pagesPerId: chunkPagesPerId } = await buildChunkPdf(chunks[ci].ids, (pct, label) => {
            result.updateChunkProgress(ci, pct, label);
          }, chunkWorkerName, chunkWorkerIcon));

          // §Divider: prepend per-subgroup dividers to each split chunk.
          // Runs BEFORE picking list so picking list stays at the very top.
          if (withDivider && window.PDFLib) {
            try {
              // Per-chunk assigneeKind overrides top-level (by-person/by-person-sku sets it per chunk).
              const chunkKind = chunks[ci].assigneeKind || assigneeKind;
              const chunkWorkerName = chunks[ci].workerName || workerName;
              const { bytes: dBytes, pageCount: dCount } = await prependDividersToChunk(bytes, chunks[ci].ids, chunkWorkerName, chunkKind, chunkPagesPerId);
              bytes = dBytes;
              if (dCount != null) pageCount = dCount;
            } catch (_dErr) {
              console.warn('[QF] divider prepend failed:', _dErr);
            }
          }

          // §7.1: Prepend picking list pages to each chunk when requested.
          // NOTE: only runs on the non-prebuilt branch — buildMultiSkuCombinedPdf
          // already bakes picking list in on the prebuilt path (Bug #1 fix: ป้องกันหน้าซ้ำ).
          if (withPickingList && window.PDFLib) {
            try {
              const { PDFDocument } = window.PDFLib;
              // Build groups for this chunk's ids.
              const chunkGroupMap = new Map();
              for (const id of chunks[ci].ids) {
                const rec = state.records.get(id);
                if (!rec?.skuList?.length) continue;
                const s = rec.skuList[0];
                const key = `${s.productId}:${s.skuId || ''}`;
                if (!chunkGroupMap.has(key)) {
                  const alias = (getAlias(s.productId) || '').trim();
                  const variantInfo = getVariantInfo(s.productId, s.skuId);
                  chunkGroupMap.set(key, {
                    productId: s.productId, skuId: s.skuId || null,
                    alias: alias || shortName(s.productName),
                    officialName: s.productName || '',
                    variantName: (variantInfo?.alias || '').trim() || (s.skuName || ''),
                    productImageURL: s.productImageURL || null,
                    sellerSku: s.sellerSkuName || '',
                    ids: [],
                  });
                }
                chunkGroupMap.get(key).ids.push(id);
              }
              const chunkGroups = [...chunkGroupMap.values()];
              // Probe page size from existing label bytes.
              let W = 288; let H = 432;
              try {
                const sampleDoc = await PDFDocument.load(bytes);
                W = sampleDoc.getPage(0).getWidth();
                H = sampleDoc.getPage(0).getHeight();
              } catch {}
              const headerMeta = buildPickingHeaderMeta(chunkGroups, assigneeName);
              const tmpDoc = await buildPickingListPages(chunkGroups, { W, H }, headerMeta);
              if (tmpDoc && tmpDoc.getPageCount() > 0) {
                const combinedDoc = await PDFDocument.create();
                if (window.fontkit) combinedDoc.registerFontkit(window.fontkit);
                const plCopied = await combinedDoc.copyPages(tmpDoc, tmpDoc.getPageIndices());
                plCopied.forEach(p => combinedDoc.addPage(p));
                const labelDoc = await PDFDocument.load(bytes);
                const lCopied = await combinedDoc.copyPages(labelDoc, labelDoc.getPageIndices());
                lCopied.forEach(p => combinedDoc.addPage(p));
                bytes = await combinedDoc.save();
                pageCount = combinedDoc.getPageCount();
              }
            } catch (_plErr) {
              console.warn('[QF] picking list prepend failed:', _plErr);
            }
          }
        }

        const blob = new Blob([bytes], {type: 'application/pdf'});
        const url = URL.createObjectURL(blob);
        const labelCount = (chunks[ci].ids || []).length;
        console.log(`[QF]   chunk ${ci + 1}/${chunks.length} done in ${(performance.now() - chunkStartMs).toFixed(0)}ms (${labelCount} ใบ → ${pageCount} หน้า)`);
        result.completeChunk(ci, {url, pageCount, labelCount});
        return true;
      } catch (e) {
        console.error('[QF] chunk', ci+1, 'failed:', e);
        if (e.isRateLimit) showToast(e.message, 6000);
        result.errorChunk(ci, e.message);
        return false;
      }
    };

    result.setRetryHandler(runChunk);

    // Run all chunks fully in parallel — user can throttle by choosing fewer
    // chunks; each chunk already fires its own batches in parallel internally.
    await Promise.all(chunks.map((_, ci) => runChunk(ci)));
    console.log(`[QF] runChunkedExport: all ${chunks.length} chunks done in ${(performance.now() - runStartMs).toFixed(0)}ms total`);
    result.allDone();

    // Save history entry for TikTok prints only. Shopee print isn't wired
    // yet, and re-downloads (historyMeta=null) must not create duplicate
    // entries.
    if (historyMeta && !isShopee()) {
      try {
        addHistoryEntry({
          id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
          timestamp: Date.now(),
          title: displayTitle,
          baseFilename: historyMeta.baseFilename || displayTitle,
          totalLabels: historyMeta.totalLabels || totalIds,
          chunks: chunks.map(c => ({
            label: c.label,
            filename: c.filename,
            ids: [...c.ids],
          })),
          platform: 'tiktok',
          workerId: historyMeta.workerId || null,
          workerName: historyMeta.workerName || null,
          // §8.2: assignee fields added for CSV export and daily planning.
          assigneeKind: historyMeta.assigneeKind || null,
          assigneeName: historyMeta.assigneeName || null,
        });
        renderHistoryBadge();

        // Save order-level history to IndexedDB (for claim lookup).
        const _pad = n => String(n).padStart(2, '0');
        const _now = Date.now();
        const _d = new Date(_now);
        const _date = `${_d.getFullYear()}-${_pad(_d.getMonth()+1)}-${_pad(_d.getDate())}`;
        const _orderRecs = [];
        for (const _chunk of chunks) {
          for (const _id of (_chunk.ids || [])) {
            const _rec = state.records.get(_id);
            if (!_rec) continue;
            const _cid = state.carrierOf.get(_id);
            const _carrier = _cid ? ([...state.carriers.values()].find(c => c.id === _cid)?.name || null) : null;
            for (const _oid of (_rec.orderIds || [])) {
              if (!_oid) continue;
              _orderRecs.push({
                orderId: String(_oid),
                fulfillUnitId: String(_id),
                ts: _now,
                date: _date,
                carrier: _carrier,
                assigneeKind: historyMeta.assigneeKind || null,
                teamId: historyMeta.teamId || null,
                teamName: historyMeta.teamName || null,
                teamSnapshot: historyMeta.teamSnapshot || null,
                workerId: historyMeta.workerId || null,
                workerName: historyMeta.workerName || null,
              });
            }
          }
        }
        saveOrderHistoryBatch(_orderRecs);
      } catch (e) {
        console.warn('[QF] failed to save history:', e);
      }
    }
    return true;
  }

  function showSplitChoiceModal({productName, alias, variants, totalCount}) {
    return new Promise(resolve => {
      const overlay = document.createElement('div');
      overlay.className = 'qf-modal-overlay qf-split-overlay';
      const defaultPick = variants.length > 1 ? 'per-variant' : 'merge';
      overlay.innerHTML = `
        <div class="qf-modal qf-split-modal" role="dialog">
          <div class="qf-modal-title">พิมพ์ "${escapeHtml(alias || productName)}"</div>
          <div class="qf-modal-body">
            <div class="qf-split-sub">${variants.length} ตัวเลือก · ${totalCount} ฉลากรวม</div>
            <div class="qf-split-options">
              <label class="qf-split-opt">
                <input type="radio" name="qf-split" value="merge"/>
                <div class="qf-split-opt-body">
                  <div class="qf-split-opt-title">รวมเป็นไฟล์เดียว</div>
                  <div class="qf-split-opt-desc">${totalCount} ฉลากในไฟล์เดียว · เหมาะกับการพิมพ์ครั้งเดียวจบ</div>
                </div>
              </label>
              <label class="qf-split-opt">
                <input type="radio" name="qf-split" value="per-variant" ${defaultPick==='per-variant'?'checked':''}/>
                <div class="qf-split-opt-body">
                  <div class="qf-split-opt-title">แยกตามตัวเลือก <span class="qf-split-badge">แนะนำ</span></div>
                  <div class="qf-split-opt-desc">${variants.length} ไฟล์ · ${variants.map(v => (v.skuName||v.sellerSkuName||v.skuId).slice(0,14)).join(', ').slice(0,80)}${variants.length > 3 ? '...' : ''}</div>
                </div>
              </label>
              <label class="qf-split-opt">
                <input type="radio" name="qf-split" value="chunked"/>
                <div class="qf-split-opt-body">
                  <div class="qf-split-opt-title">แยกตามจำนวน</div>
                  <div class="qf-split-opt-desc">เลือกจำนวนไฟล์เองในหน้าจอถัดไป</div>
                </div>
              </label>
            </div>
          </div>
          <div class="qf-modal-actions">
            <button class="qf-btn-cancel">ยกเลิก</button>
            <button class="qf-btn-confirm">พิมพ์เลย</button>
          </div>
        </div>
      `;
      document.body.appendChild(overlay);
      const cleanup = (v) => { overlay.remove(); resolve(v); };
      overlay.querySelector('.qf-btn-cancel').onclick = () => cleanup(null);
      overlay.querySelector('.qf-btn-confirm').onclick = () => {
        const picked = overlay.querySelector('input[name="qf-split"]:checked')?.value;
        cleanup(picked || defaultPick);
      };
      overlay.onclick = e => { if (e.target === overlay) cleanup(null); };
      const onKey = e => { if (e.key === 'Escape') { cleanup(null); document.removeEventListener('keydown', onKey); } };
      document.addEventListener('keydown', onKey);
    });
  }

  // §5.3: Modal for card-click when product has multiple qty buckets.
  // Resolves 'merge' | 'split' | null.
  function showQtyCombineModal({ alias, buckets }) {
    return new Promise(resolve => {
      const overlay = document.createElement('div');
      overlay.className = 'qf-modal-overlay qf-qty-combine-overlay';
      const bucketDesc = buckets.map(b => `×${b.qty} (${b.count})`).join(', ');
      const totalCount = buckets.reduce((s, b) => s + b.count, 0);
      overlay.innerHTML = `
        <div class="qf-modal qf-qty-combine-modal" role="dialog">
          <div class="qf-modal-title">พิมพ์ "${escapeHtml(alias)}"</div>
          <div class="qf-modal-body">
            <div class="qf-split-sub">${buckets.length} กลุ่มจำนวน · ${totalCount} ฉลากรวม</div>
            <div class="qf-split-options">
              <label class="qf-split-opt">
                <input type="radio" name="qf-qty-combine" value="split" checked/>
                <div class="qf-split-opt-body">
                  <div class="qf-split-opt-title">แยก PDF ต่อจำนวน (ค่าเริ่มต้น)</div>
                  <div class="qf-split-opt-desc">${buckets.length} ไฟล์ · ${escapeHtml(bucketDesc)}</div>
                </div>
              </label>
              <label class="qf-split-opt">
                <input type="radio" name="qf-qty-combine" value="merge"/>
                <div class="qf-split-opt-body">
                  <div class="qf-split-opt-title">รวมไฟล์เดียว</div>
                  <div class="qf-split-opt-desc">เรียงตามจำนวน: ${escapeHtml(bucketDesc)}</div>
                </div>
              </label>
              <label class="qf-split-opt qf-qty-combine-sub" style="display:none;">
                <input type="checkbox" name="qf-qty-divider" class="qf-qty-divider-chk"/>
                <div class="qf-split-opt-body">
                  <div class="qf-split-opt-title">แทรกหน้าคั่นระหว่างกลุ่มจำนวน</div>
                  <div class="qf-split-opt-desc">หน้าคั่นแสดงชื่อย่อ + จำนวน ก่อนฉลากแต่ละกลุ่ม</div>
                </div>
              </label>
            </div>
          </div>
          <div class="qf-modal-actions">
            <button class="qf-btn-cancel">ยกเลิก</button>
            <button class="qf-btn-confirm">พิมพ์</button>
          </div>
        </div>
      `;
      document.body.appendChild(overlay);
      const cleanup = v => { overlay.remove(); resolve(v); };
      const subRow = overlay.querySelector('.qf-qty-combine-sub');
      const divChk = overlay.querySelector('.qf-qty-divider-chk');
      overlay.querySelectorAll('input[name="qf-qty-combine"]').forEach(r => {
        r.addEventListener('change', () => {
          const mergePicked = overlay.querySelector('input[name="qf-qty-combine"]:checked')?.value === 'merge';
          subRow.style.display = mergePicked ? '' : 'none';
          if (!mergePicked) divChk.checked = false;
        });
      });
      overlay.querySelector('.qf-btn-cancel').onclick = () => cleanup(null);
      overlay.querySelector('.qf-btn-confirm').onclick = () => {
        const picked = overlay.querySelector('input[name="qf-qty-combine"]:checked')?.value || 'split';
        const withDivider = picked === 'merge' && divChk.checked;
        cleanup({ mode: picked, withDivider });
      };
      overlay.onclick = e => { if (e.target === overlay) cleanup(null); };
      const onKey = e => { if (e.key === 'Escape') { cleanup(null); document.removeEventListener('keydown', onKey); } };
      document.addEventListener('keydown', onKey);
    });
  }

  // §5.4: Print all qty buckets combined (each bucket = one divider group).
  async function printQtyBucketsCombined(productId, alias, buckets, workerName, workerIcon, plan) {
    const product = state.products.get(productId);
    const groups = buckets.map(b => ({
      productId,
      skuId: null,
      alias: alias + ' ×' + b.qty,
      officialName: product?.productName || alias,
      variantName: '×' + b.qty,
      productImageURL: product?.productImageURL || null,
      ids: [...b.ids],
    }));
    const totalIds = groups.reduce((s, g) => s + g.ids.length, 0);
    const assignee = workerName || null;
    const baseFilename = assignee
      ? makeBaseFilename('[' + alias + '] [' + assignee + ']')
      : makeBaseFilename(alias);
    try {
      const prepProgress = showProgress(`กำลังเตรียม PDF รวม (${groups.length} กลุ่ม · ${totalIds} ฉลาก)`);
      let bytes;
      try {
        ({ bytes } = await buildMultiSkuCombinedPdf(groups, workerName, workerIcon, plan.withPickingList, (pct) => {
          prepProgress.update(pct, `${pct.toFixed(0)}%`);
        }, plan.withDivider));
      } finally {
        document.querySelectorAll('.qf-progress-overlay').forEach(e => e.remove());
      }
      const allIds = groups.flatMap(g => g.ids);
      const exportChunk = { ids: allIds, label: 'ไฟล์เดียว', filename: baseFilename + '.pdf' };
      const ok = await runChunkedExport([exportChunk], alias + ' (รวมตามจำนวน)', {
        baseFilename,
        totalLabels: totalIds,
        workerName,
        workerIcon,
        assigneeKind: workerName ? 'worker' : null,
        assigneeName: workerName || null,
      });
      return ok;
    } catch (e) {
      throw new Error('รวม PDF ไม่สำเร็จ: ' + e.message);
    }
  }

  // §5.5: Print each qty bucket as a separate chunk.
  async function printQtyBucketsSplit(productId, alias, buckets, workerName, workerIcon) {
    const assignee = workerName || null;
    const chunks = buckets.map(b => {
      const hint = alias + ' x' + b.qty;
      const baseFilename = assignee
        ? makeBaseFilename('[' + hint + '] [' + assignee + ']')
        : makeBaseFilename(hint);
      return {
        ids: [...b.ids],
        label: alias + ' ×' + b.qty,
        filename: baseFilename + '.pdf',
      };
    });
    const totalIds = chunks.reduce((s, c) => s + c.ids.length, 0);
    const baseFilename = assignee
      ? makeBaseFilename('[' + alias + '] [' + assignee + ']')
      : makeBaseFilename(alias);
    return runChunkedExport(chunks, alias + ' (แยกตามจำนวน)', {
      baseFilename,
      totalLabels: totalIds,
      workerName,
      workerIcon,
      assigneeKind: workerName ? 'worker' : null,
      assigneeName: workerName || null,
    });
  }

  async function printProductByVariants(productId, scenario, variantsList) {
    const product = state.products.get(productId);
    const alias = (getAlias(productId) || '').trim();
    const baseName = alias || shortName(product?.productName);
    const SUB = 200;
    const chunks = [];
    for (const {v, ids} of variantsList) {
      const variantName = (v.skuName || v.sellerSkuName || v.skuId || '').trim();
      const label = `${baseName} · ${variantName}`;
      const fnHint = `${baseName} ${variantName}`.trim();
      if (ids.length <= SUB) {
        chunks.push({ids, label, filename: `${makeBaseFilename(fnHint)}.pdf`});
      } else {
        const subCount = Math.ceil(ids.length / SUB);
        const subSize = Math.ceil(ids.length / subCount);
        for (let i = 0; i < ids.length; i += subSize) {
          const idx = chunks.filter(c => c._v === v.skuId).length + 1;
          chunks.push({
            _v: v.skuId,
            ids: ids.slice(i, i + subSize),
            label: `${label} (${idx}/${subCount})`,
            filename: `${makeBaseFilename(fnHint)}-ชุด${idx}-${subCount}.pdf`,
          });
        }
      }
    }
    const totalIds = chunks.reduce((s, c) => s + c.ids.length, 0);
    const confirm = await showPrintConfirm({
      title: `${product?.productName || productId}`,
      summary: `แยก ${chunks.length} ไฟล์ตามตัวเลือก`,
      count: totalIds,
      sampleText: `${baseName} · ...`,
    });
    if (!confirm) return false;
    return runChunkedExport(chunks, `${alias || product?.productName} (แยกตามตัวเลือก)`, {
      baseFilename: makeBaseFilename(baseName),
      totalLabels: totalIds,
      workerId: confirm.workerId,
      workerName: confirm.workerName,
    });
  }

  async function printProductLabels(productId, skuId, scenario) {
    const product = state.products.get(productId);
    const alias = (getAlias(productId) || '').trim();
    const baseName = alias || shortName(product?.productName);

    // Card-level click (no skuId) + product has multiple variants with data → ask split choice
    if (!skuId && product) {
      const variantsList = [...product.variants.values()]
        .filter(v => v.skuId != null) // skip synthetic no-variant entry
        .map(v => ({v, ids: collectFulfillIds(productId, v.skuId, scenario)}))
        .filter(x => x.ids.length > 0);
      if (variantsList.length > 1) {
        const totalCount = variantsList.reduce((s, x) => s + x.ids.length, 0);
        const choice = await showSplitChoiceModal({
          productName: product.productName,
          alias,
          variants: variantsList.map(x => x.v),
          totalCount,
        });
        if (!choice) return false;
        if (choice === 'per-variant') {
          return printProductByVariants(productId, scenario, variantsList);
        }
        // 'merge' and 'chunked' fall through — printIds now handles chunking via showChunkPlanModal.
        const ids = collectFulfillIds(productId, null, scenario);
        const title = `${product.productName}`;
        const sampleQty = scenario === 'multi' ? '2+' : '1';
        return printIds(ids, title, `${baseName} ${sampleQty}`, baseName);
      }
    }

    // Default path (variant badge click or single-variant product)
    const ids = collectFulfillIds(productId, skuId, scenario);
    const variant = skuId ? product?.variants.get(skuId) : null;
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
    if (!ids.length) { showToast('ไม่มีฉลากใน combo นี้', 2000); return false; }

    const sample = combo.items
      .map(i => `${(getAlias(i.productId) || '').trim() || shortName(i.productName)} ${i.quantity}`)
      .join(' + ');
    const filenameHint = combo.items
      .map(i => (getAlias(i.productId) || '').trim() || shortName(i.productName))
      .join('+');
    const displayLabel = `ออเดอร์แปลก: ${sample}`;

    try {
      // §4: Weird combos are always multi-SKU by definition.
      // Show chunk-plan modal (which includes the "รวม PDF" option, default checked).
      const confirm = await showPrintConfirm({
        title: displayLabel,
        count: ids.length,
        sampleText: sample,
      });
      if (!confirm) return false;
      const { workerId, workerName, workerIcon } = confirm;

      // §4: Weird combos are always multi-SKU → always prompt so user chooses combined/split + divider.
      const plan = await showChunkPlanModal({ total: ids.length, multiSku: true, defaultPickingList: loadPickingListPref() });
      if (!plan) return false;

      const assignee = workerName || null;
      let ok;

      if (plan.combined) {
        // Build groups from combo items (each item is a distinct productId:skuId).
        const groupMap = new Map();
        for (const id of ids) {
          const rec = state.records.get(id);
          if (!rec?.skuList?.length) continue;
          const s = rec.skuList[0];
          const key = `${s.productId}:${s.skuId}`;
          if (!groupMap.has(key)) {
            const alias = (getAlias(s.productId) || '').trim();
            const variantInfo = getVariantInfo(s.productId, s.skuId);
            groupMap.set(key, {
              productId: s.productId,
              skuId: s.skuId,
              alias: alias || shortName(s.productName),
              officialName: s.productName || '',
              variantName: (variantInfo?.alias || '').trim() || (s.skuName || s.sellerSkuName || ''),
              productImageURL: s.productImageURL || null,
              ids: [],
            });
          }
          groupMap.get(key).ids.push(id);
        }
        const groups = [...groupMap.values()].sort((a, b) => a.alias.localeCompare(b.alias));
        const baseHint = `รวม ${groups.length} SKU`;
        const baseFilename = assignee
          ? makeBaseFilename(`[${baseHint}] [${assignee}]`)
          : makeBaseFilename(baseHint);

        const slices = planSlice(ids, plan);
        const chunkCount = slices.length;
        const exportChunks = slices.map((slice, i) => {
          const idx = i + 1;
          const chunkSuffix = chunkCount > 1 ? `-ชุด${idx}-${chunkCount}` : '';
          return {
            ids: slice,
            label: chunkCount === 1 ? 'ไฟล์เดียว' : `ชุด ${idx}/${chunkCount}`,
            filename: `${baseFilename}${chunkSuffix}.pdf`,
          };
        });

        ok = await runChunkedExport(exportChunks, displayLabel, {
          baseFilename,
          totalLabels: ids.length,
          workerId,
          workerName,
          workerIcon,
          assigneeKind: workerName ? 'worker' : null,
          assigneeName: workerName || null,
        });
      } else {
        // Per-SKU: delegate to printIds (no confirm re-prompt needed — already confirmed above).
        // Build flat chunks per-SKU bucket.
        const groupMap = new Map();
        for (const id of ids) {
          const rec = state.records.get(id);
          if (!rec?.skuList?.length) continue;
          const s = rec.skuList[0];
          const key = `${s.productId}:${s.skuId}`;
          if (!groupMap.has(key)) {
            const alias = (getAlias(s.productId) || '').trim();
            groupMap.set(key, { alias: alias || shortName(s.productName), ids: [] });
          }
          groupMap.get(key).ids.push(id);
        }
        const groups = [...groupMap.values()].sort((a, b) => a.alias.localeCompare(b.alias));
        const baseFilename = assignee
          ? makeBaseFilename(`[${filenameHint}] [${assignee}]`)
          : makeBaseFilename(filenameHint);
        const exportChunks = [];
        for (const grp of groups) {
          const slices = planSlice(grp.ids, plan);
          const subCount = slices.length;
          slices.forEach((slice, i) => {
            const idx = i + 1;
            const chunkSuffix = subCount > 1 ? `-ชุด${idx}-${subCount}` : '';
            exportChunks.push({
              ids: slice,
              label: subCount === 1 ? grp.alias : `${grp.alias} ชุด${idx}/${subCount}`,
              filename: assignee
                ? makeBaseFilename(`[${grp.alias}] [${assignee}]`) + `${chunkSuffix}.pdf`
                : makeBaseFilename(grp.alias) + `${chunkSuffix}.pdf`,
            });
          });
        }
        ok = await runChunkedExport(exportChunks, displayLabel, {
          baseFilename,
          totalLabels: ids.length,
          workerId,
          workerName,
          workerIcon,
          assigneeKind: workerName ? 'worker' : null,
          assigneeName: workerName || null,
        });
      }

      if (ok) { markComboDone(sigKey); renderAll(); }
      return ok;
    } catch (e) {
      showErrorToast('พิมพ์ผิดพลาด: ' + e.message, {
        source: 'printWeirdCombo',
        sigKey,
        error: String(e && (e.stack || e.message || e)),
      });
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

  // Build `mainOrderId → <tr>` map by walking each row's fiber once. Avoids
  // the O(n²) hit of calling findTrForOrder() inside selectAllOrders (which
  // re-queries + re-walks every row for every record on the page).
  function buildTrMapForPage() {
    const map = new Map();
    const trs = document.querySelectorAll('tr');
    for (const tr of trs) {
      const fk = Object.keys(tr).find(k => k.startsWith('__reactFiber'));
      if (!fk) continue;
      let fiber = tr[fk], depth = 0;
      while (fiber && depth < 10) {
        const id = fiber.memoizedProps?.record?.mainOrderId
                || fiber.memoizedProps?.rowData?.mainOrderId;
        if (id) { if (!map.has(id)) map.set(id, tr); break; }
        fiber = fiber.return; depth++;
      }
    }
    return map;
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
      // Build the trMap once per page so the inner record loop is O(n)
      // instead of O(n²).
      const trMap = buildTrMapForPage();
      for (const rec of getOrderRecords()) {
        const targetSku = filterSkuId
          ? rec.skuList?.find(s => s.skuId === filterSkuId)
          : rec.skuList?.find(s => s.productId === filterProductId);
        if (!targetSku || targetSku.quantity < minQty) continue;
        const tr = trMap.get(rec.mainOrderId);
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
          showErrorToast('พิมพ์ผิดพลาด: ' + e.message, {
            source: 'applyProductFilter:print',
            productId, skuId, type,
            error: String(e && (e.stack || e.message || e)),
          });
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

  // §7.6: Mask email for picking list header — show first 2 chars + *** + @domain.
  function maskEmail(email) {
    if (!email || typeof email !== 'string') return '(ไม่ระบุ)';
    const atIdx = email.indexOf('@');
    if (atIdx < 0) return '***';
    const local = email.slice(0, atIdx);
    const domain = email.slice(atIdx + 1);
    const keep = local.slice(0, 2);
    return keep + '***@' + domain;
  }

  // ==================== CSV EXPORT ====================
  function buildHistoryCsv(entries) {
    const q = (v) => '"' + String(v == null ? '' : v).replace(/[\r\n]+/g, ' ').replace(/"/g, '""') + '"';
    const pad = n => String(n).padStart(2, '0');
    const header = [
      q('timestamp'), q('date'), q('time'), q('platform'),
      q('title'), q('baseFilename'), q('chunkLabel'), q('chunkFilename'), q('chunkIdCount'), q('totalLabels'),
      q('assigneeKind'), q('assigneeName'),
      q('teamId'), q('teamName'),
      q('workerId'), q('workerName'),
    ].join(',');
    const rows = [];
    for (const e of entries) {
      const d = new Date(e.timestamp);
      const date = `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
      const time = `${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
      const platform = e.platform || 'tiktok';
      const title = e.title || '';
      const baseFilename = e.baseFilename || '';
      const totalLabels = e.totalLabels != null ? e.totalLabels : '';
      const assigneeKind = e.assigneeKind || (e.workerId ? 'worker' : '');
      const assigneeName = e.assigneeName || e.workerName || '';
      const teamId = e.teamId || '';
      const teamName = e.teamName || '';
      const workerId = e.workerId || '';
      const workerName = e.workerName || '';
      for (const c of (e.chunks || [])) {
        const chunkLabel = c.label || '';
        const chunkFilename = c.filename || '';
        const chunkIdCount = (c.ids || []).length;
        rows.push([
          q(e.timestamp), q(date), q(time), q(platform),
          q(title), q(baseFilename), q(chunkLabel), q(chunkFilename), q(chunkIdCount), q(totalLabels),
          q(assigneeKind), q(assigneeName),
          q(teamId), q(teamName),
          q(workerId), q(workerName),
        ].join(','));
      }
    }
    return '\ufeff' + [header, ...rows].join('\r\n');
  }

  function buildDailyCsv(entries) {
    const q = (v) => '"' + String(v == null ? '' : v).replace(/[\r\n]+/g, ' ').replace(/"/g, '""') + '"';
    const pad = n => String(n).padStart(2, '0');
    const fmtDate = (ts) => { const d = new Date(ts); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; };
    const fmtTime = (ts) => { const d = new Date(ts); return `${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`; };

    // §9.6: plannedQty from snapshots.
    const plannedMap = new Map();
    const snapshots = loadPlanSnapshots();
    const hasSnapshots = snapshots.length > 0;
    for (const snap of snapshots) {
      for (const [, col] of Object.entries(snap.columns || {})) {
        const ak = col.assigneeKind || 'unassigned';
        const an = col.assigneeName || '';
        const assigneeKey = ak === 'team' ? `team:${an}` : (ak === 'worker' ? `worker:${an}` : 'unassigned');
        for (const id of (col.ids || [])) {
          const rec = state.records.get(id);
          if (!rec?.skuList?.length) continue;
          for (const s of rec.skuList) {
            const pk = `${snap.date}:${assigneeKey}:${s.productId}:${s.skuId || ''}`;
            plannedMap.set(pk, (plannedMap.get(pk) || 0) + (s.quantity || 1));
          }
        }
      }
    }

    const dailyMap = new Map();
    for (const e of entries) {
      const day = fmtDate(e.timestamp);
      const assigneeKind = e.assigneeKind || (e.workerId ? 'worker' : 'unassigned');
      const assigneeName = e.assigneeName || e.workerName || e.teamName || '';
      const assigneeKey = assigneeKind === 'team' ? `team:${assigneeName}` : (assigneeKind === 'worker' ? `worker:${assigneeName}` : 'unassigned');
      for (const chunk of (e.chunks || [])) {
        for (const id of (chunk.ids || [])) {
          const rec = state.records.get(id);
          if (!rec?.skuList?.length) continue;
          for (const s of rec.skuList) {
            const key = `${day}:${assigneeKey}:${s.productId}:${s.skuId || ''}`;
            if (!dailyMap.has(key)) {
              const alias = (getAlias(s.productId) || '').trim();
              const variantInfo = getVariantInfo(s.productId, s.skuId);
              dailyMap.set(key, {
                date: day, platform: e.platform || 'tiktok',
                assigneeKind,
                workerName: assigneeKind === 'worker' ? assigneeName : '',
                teamName: assigneeKind === 'team' ? assigneeName : '',
                productAlias: alias || shortName(s.productName),
                officialName: s.productName || '',
                skuAlias: (variantInfo?.alias || '').trim(),
                sellerSku: s.sellerSkuName || '',
                printedQty: 0,
                startTime: e.timestamp, endTime: e.timestamp,
                _planKey: key,
              });
            }
            const row = dailyMap.get(key);
            row.printedQty += (s.quantity || 1);
            if (e.timestamp < row.startTime) row.startTime = e.timestamp;
            if (e.timestamp > row.endTime) row.endTime = e.timestamp;
          }
        }
      }
    }

    const header = [
      q('date'), q('platform'), q('assigneeKind'), q('workerName'), q('teamName'),
      q('productAlias'), q('officialName'), q('skuAlias'), q('sellerSku'),
      q('plannedQty'), q('printedQty'), q('remainingQty'),
      q('startTime'), q('endTime'),
    ].join(',');

    const preamble = hasSnapshots ? [] : [q('หมายเหตุ: ไม่พบ snapshot แผนงาน — plannedQty ใช้ printedQty แทน')];

    const rows = [...dailyMap.values()]
      .sort((a, b) => a.date.localeCompare(b.date) || a.assigneeKind.localeCompare(b.assigneeKind) || a.productAlias.localeCompare(b.productAlias))
      .map(r => {
        const planned = hasSnapshots ? (plannedMap.get(r._planKey) || 0) : r.printedQty;
        const remaining = Math.max(0, planned - r.printedQty);
        return [
          q(r.date), q(r.platform), q(r.assigneeKind), q(r.workerName), q(r.teamName),
          q(r.productAlias), q(r.officialName), q(r.skuAlias), q(r.sellerSku),
          q(planned), q(r.printedQty), q(remaining),
          q(fmtTime(r.startTime)), q(fmtTime(r.endTime)),
        ].join(',');
      });

    return '\ufeff' + [...preamble, header, ...rows].join('\r\n');
  }

  function buildOrderHistoryCsv(records) {
    const q = (v) => '"' + String(v == null ? '' : v).replace(/[\r\n]+/g, ' ').replace(/"/g, '""') + '"';
    const header = [
      q('orderId'), q('fulfillUnitId'), q('date'), q('time'), q('carrier'),
      q('assigneeKind'), q('workerName'), q('teamName'), q('teamMembers'),
    ].join(',');
    const pad = n => String(n).padStart(2, '0');
    const rows = records.map(r => {
      const d = new Date(r.ts);
      const time = `${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
      const teamMembers = Array.isArray(r.teamSnapshot)
        ? r.teamSnapshot.map(m => m.workerName || m.workerId).join(', ')
        : '';
      return [
        q(r.orderId), q(r.fulfillUnitId), q(r.date), q(time), q(r.carrier || ''),
        q(r.assigneeKind || ''), q(r.workerName || ''), q(r.teamName || ''), q(teamMembers),
      ].join(',');
    });
    return '\ufeff' + [header, ...rows].join('\r\n');
  }

  function downloadCsv(csv, filename) {
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.style.display = 'none';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  // Phase 1 Custom Layout: preset picker modal.
  function openDividerPresetModal() {
    document.querySelectorAll('.qf-divider-preset-overlay').forEach(e => e.remove());
    const current = loadDividerPreset();
    const overlay = document.createElement('div');
    overlay.className = 'qf-modal-overlay qf-divider-preset-overlay';
    const presets = [
      { id: 'minimal',     title: 'ขั้นต่ำสุด (Minimal)',    desc: 'แสดงเฉพาะชื่อย่อขนาดใหญ่กลางหน้า + จำนวน — เหมาะกับพนักงานที่จำรหัสสินค้าได้' },
      { id: 'standard',    title: 'มาตรฐาน (Standard)',      desc: 'แสดงชื่อย่อ + ชื่อสินค้า + ตัวเลือก + ขนส่ง + รูป + จำนวน (ค่าเริ่มต้น)' },
      { id: 'detailed',    title: 'ละเอียด (Detailed)',     desc: 'เหมือน Standard + แสดงรายละเอียดเพิ่ม (เหมาะกับตรวจสอบละเอียด)' },
      { id: 'photo-first', title: 'รูปก่อน (Photo-first)',  desc: 'รูปสินค้าใหญ่ด้านบน ชื่อย่อใต้รูป — เหมาะกับพนักงานใหม่จำรูป' },
    ];
    const optsHtml = presets.map(p => `
      <label class="qf-divider-preset-opt">
        <input type="radio" name="qf-divider-preset" value="${p.id}" ${p.id === current ? 'checked' : ''} />
        <div class="qf-divider-preset-text">
          <div class="qf-divider-preset-title">${escapeHtml(p.title)}</div>
          <div class="qf-divider-preset-desc">${escapeHtml(p.desc)}</div>
        </div>
      </label>
    `).join('');
    overlay.innerHTML = `
      <div class="qf-modal qf-divider-preset-modal" role="dialog">
        <div class="qf-modal-title">ใบคั่น - เลือกรูปแบบ</div>
        <div class="qf-modal-body">
          <div class="qf-modal-summary">รูปแบบนี้จะใช้ทุกครั้งที่พิมพ์ฉลากพร้อมใบคั่น (ทุกอย่างขาวดำ)</div>
          <div class="qf-divider-preset-list">${optsHtml}</div>
        </div>
        <div class="qf-modal-actions">
          <button class="qf-btn-cancel">ยกเลิก</button>
          <button class="qf-btn-confirm">บันทึก</button>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);
    const close = () => overlay.remove();
    overlay.querySelector('.qf-btn-cancel').onclick = close;
    overlay.querySelector('.qf-btn-confirm').onclick = () => {
      const chosen = overlay.querySelector('input[name="qf-divider-preset"]:checked')?.value;
      if (chosen) {
        saveDividerPreset(chosen);
        showToast(`บันทึกรูปแบบใบคั่น: ${chosen}`, 2500);
      }
      close();
    };
    overlay.onclick = (e) => { if (e.target === overlay) close(); };
    const onKey = (e) => { if (e.key === 'Escape') { close(); document.removeEventListener('keydown', onKey); } };
    document.addEventListener('keydown', onKey);
  }

  function openCsvExportModal() {
    document.querySelectorAll('.qf-csv-overlay').forEach(e => e.remove());
    const overlay = document.createElement('div');
    overlay.className = 'qf-modal-overlay qf-csv-overlay';

    const workerOptions = [
      '<option value="all">ทั้งหมด</option>',
      ...state.workers.map(w => `<option value="${escapeHtml(w.id)}">${escapeHtml(w.name)}</option>`),
      '<option value="none">ไม่ระบุ</option>',
    ].join('');

    overlay.innerHTML = `
      <div class="qf-modal qf-csv-modal" role="dialog">
        <div class="qf-workers-header">
          <div class="qf-modal-title" style="margin-bottom:0;">ดาวน์โหลด CSV</div>
          <button class="qf-csv-close qf-workers-close">×</button>
        </div>
        <div class="qf-csv-body">
          <div class="qf-csv-section-label">ประเภทรายงาน:</div>
          <label class="qf-csv-radio-label"><input type="radio" name="qf-csv-type" value="history" checked> ประวัติการพิมพ์ (per chunk audit log)</label>
          <label class="qf-csv-radio-label"><input type="radio" name="qf-csv-type" value="daily"> แผนงานรายวัน (daily planning summary)</label>
          <label class="qf-csv-radio-label"><input type="radio" name="qf-csv-type" value="orders"> ประวัติออเดอร์ (order claim lookup)</label>
          <div id="qf-csv-orders-opts" style="display:none;">
            <div class="qf-csv-section-label" style="margin-top:10px;">ช่วงเวลา (สูงสุด 90 วัน):</div>
            <label class="qf-csv-radio-label"><input type="radio" name="qf-csv-orders-range" value="7d" checked> 7 วัน</label>
            <label class="qf-csv-radio-label"><input type="radio" name="qf-csv-orders-range" value="30d"> 30 วัน</label>
            <label class="qf-csv-radio-label"><input type="radio" name="qf-csv-orders-range" value="all"> ทั้งหมด (90 วัน)</label>
          </div>
          <div id="qf-csv-history-opts">
            <div class="qf-csv-section-label" style="margin-top:10px;">ช่วงเวลา:</div>
            <label class="qf-csv-radio-label"><input type="radio" name="qf-csv-range" value="today"> วันนี้</label>
            <label class="qf-csv-radio-label"><input type="radio" name="qf-csv-range" value="7d" checked> 7 วัน</label>
            <label class="qf-csv-radio-label"><input type="radio" name="qf-csv-range" value="all"> ทั้งหมด</label>
            <div class="qf-csv-section-label" style="margin-top:10px;">คนแพ็ค:</div>
            <select class="qf-csv-worker-select">${workerOptions}</select>
          </div>
          <div id="qf-csv-daily-opts" style="display:none;">
            <div class="qf-csv-section-label" style="margin-top:10px;">ช่วงเวลา:</div>
            <label class="qf-csv-radio-label"><input type="radio" name="qf-csv-daily-range" value="today"> วันนี้</label>
            <label class="qf-csv-radio-label"><input type="radio" name="qf-csv-daily-range" value="7d" checked> 7 วัน</label>
            <label class="qf-csv-radio-label"><input type="radio" name="qf-csv-daily-range" value="all"> ทั้งหมด</label>
          </div>
        </div>
        <div class="qf-modal-actions">
          <button class="qf-btn-cancel qf-csv-cancel">ยกเลิก</button>
          <button class="qf-btn-confirm qf-csv-download">ดาวน์โหลด</button>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);

    const cleanup = () => overlay.remove();

    overlay.querySelector('.qf-csv-close').onclick = (e) => { e.stopPropagation(); cleanup(); };
    overlay.querySelector('.qf-csv-cancel').onclick = (e) => { e.stopPropagation(); cleanup(); };
    overlay.onclick = (e) => { if (e.target === overlay) cleanup(); };
    const onKey = (e) => { if (e.key === 'Escape') { cleanup(); document.removeEventListener('keydown', onKey); } };
    document.addEventListener('keydown', onKey);

    // Toggle option panels per CSV type
    overlay.querySelectorAll('input[name="qf-csv-type"]').forEach(radio => {
      radio.addEventListener('change', () => {
        const val = overlay.querySelector('input[name="qf-csv-type"]:checked').value;
        overlay.querySelector('#qf-csv-history-opts').style.display = val === 'history' ? '' : 'none';
        overlay.querySelector('#qf-csv-daily-opts').style.display = val === 'daily' ? '' : 'none';
        overlay.querySelector('#qf-csv-orders-opts').style.display = val === 'orders' ? '' : 'none';
      });
    });

    overlay.querySelector('.qf-csv-download').onclick = (e) => {
      e.stopPropagation();
      const csvType = overlay.querySelector('input[name="qf-csv-type"]:checked').value;
      const today = new Date();
      const pad = n => String(n).padStart(2, '0');
      const todayStr = `${today.getFullYear()}-${pad(today.getMonth() + 1)}-${pad(today.getDate())}`;
      const dayStart = (d) => { const x = new Date(d); x.setHours(0, 0, 0, 0); return x.getTime(); };
      const now = Date.now();

      if (csvType === 'history') {
        const range = overlay.querySelector('input[name="qf-csv-range"]:checked').value;
        const workerVal = overlay.querySelector('.qf-csv-worker-select').value;
        let cutoff = 0;
        if (range === 'today') cutoff = dayStart(now);
        else if (range === '7d') cutoff = dayStart(now) - 6 * 24 * 60 * 60 * 1000;
        let entries = loadHistory().filter(e => e.timestamp >= cutoff);
        if (workerVal === 'none') {
          entries = entries.filter(e => e.workerId == null && e.teamId == null);
        } else if (workerVal !== 'all') {
          entries = entries.filter(e => e.workerId === workerVal);
        }
        if (entries.length === 0) { showToast('ไม่มีข้อมูลในช่วงที่เลือก', 2000); return; }
        downloadCsv(buildHistoryCsv(entries), `quickfilter-history-${todayStr}.csv`);
      } else if (csvType === 'daily') {
        const range = overlay.querySelector('input[name="qf-csv-daily-range"]:checked').value;
        let cutoff = 0;
        if (range === 'today') cutoff = dayStart(now);
        else if (range === '7d') cutoff = dayStart(now) - 6 * 24 * 60 * 60 * 1000;
        const entries = loadHistory().filter(e => e.timestamp >= cutoff);
        if (entries.length === 0) { showToast('ไม่มีข้อมูลในช่วงที่เลือก', 2000); return; }
        downloadCsv(buildDailyCsv(entries), `quickfilter-daily-${todayStr}.csv`);
        cleanup();
      } else {
        // orders — async because IndexedDB
        const range = overlay.querySelector('input[name="qf-csv-orders-range"]:checked').value;
        let cutoffTs = 0;
        if (range === '7d') cutoffTs = dayStart(now) - 6 * 24 * 60 * 60 * 1000;
        else if (range === '30d') cutoffTs = dayStart(now) - 29 * 24 * 60 * 60 * 1000;
        openOrderHistoryDb().then(db => {
          const tx = db.transaction(ORDER_HISTORY_STORE, 'readonly');
          const idx = tx.objectStore(ORDER_HISTORY_STORE).index('ts');
          const range_ = cutoffTs > 0 ? IDBKeyRange.lowerBound(cutoffTs) : null;
          const req = range_ ? idx.getAll(range_) : idx.getAll();
          req.onsuccess = () => {
            const records = req.result || [];
            if (!records.length) { showToast('ไม่มีข้อมูลออเดอร์ในช่วงที่เลือก', 2000); return; }
            downloadCsv(buildOrderHistoryCsv(records), `quickfilter-orders-${todayStr}.csv`);
            cleanup();
          };
          req.onerror = () => showToast('โหลดข้อมูลไม่สำเร็จ', 2000);
        }).catch(() => showToast('เปิด IndexedDB ไม่สำเร็จ', 2000));
        return;
      }
      cleanup();
    };
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
          ${isTikTok() && labels ? '<button id="qf-history-btn" class="qf-history-btn" title="ประวัติการพิมพ์">⏱<span class="qf-history-badge" style="display:none;">0</span></button>' : ''}
          ${isTikTok() && labels ? '<button id="qf-addrbook-btn" class="qf-history-btn" title="สมุดที่อยู่ (จับก่อนพิมพ์)">📇<span class="qf-addrbook-badge" style="display:none;">0</span></button>' : ''}
          <div id="qf-settings-wrap" style="position:relative;">
            <button id="qf-settings-btn" title="ตั้งค่า">⋮</button>
            <div id="qf-settings-menu" class="qf-settings-menu" style="display:none;">
              <button id="qf-menu-csv">ดาวน์โหลดประวัติ CSV</button>
              <button id="qf-menu-plan">🎨 วางแผน</button>
              <button id="qf-menu-pdf-templates">🏷️ ปรับแต่งฉลาก</button>
            </div>
          </div>
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
              <button class="qf-seg-btn ${state.labelStatusFilter==='failed'?'active':''}" data-val="failed">พิมพ์ไม่สำเร็จ</button>
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
          <div class="qf-tip">${isShopee()
            ? 'Shopee: คลิก card → คัดลอก order IDs → paste ในช่องค้นหาของ Shopee → ใช้ปุ่มพิมพ์ของ Shopee'
            : 'คลิก card → ยืนยัน → พิมพ์ฉลาก'}</div>
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
          <button class="qf-select-bar-all">เลือกทั้งหมด</button>
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
    document.getElementById('qf-history-btn')?.addEventListener('click', (e) => {
      e.stopPropagation();
      openHistoryModal();
    });
    document.getElementById('qf-addrbook-btn')?.addEventListener('click', (e) => {
      e.stopPropagation();
      openAddressBookModal();
    });
    const settingsBtn = document.getElementById('qf-settings-btn');
    const settingsMenu = document.getElementById('qf-settings-menu');
    settingsBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      settingsMenu.style.display = settingsMenu.style.display === 'none' ? '' : 'none';
    });
    document.addEventListener('click', () => { settingsMenu.style.display = 'none'; });
    document.getElementById('qf-menu-csv').addEventListener('click', (e) => {
      e.stopPropagation();
      settingsMenu.style.display = 'none';
      openCsvExportModal();
    });
    const planBtn = document.getElementById('qf-menu-plan');
    if (planBtn) {
      planBtn.addEventListener('click', async (e) => {
        e.stopPropagation();
        settingsMenu.style.display = 'none';
        const existing = loadPlanningSession();
        openPlanningPanel(existing);
      });
    }
    const pdfTplBtn = document.getElementById('qf-menu-pdf-templates');
    if (pdfTplBtn) {
      pdfTplBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        settingsMenu.style.display = 'none';
        openLabelOverlaySettings();
      });
    }
    renderHistoryBadge();
    refreshAddressBookBadge();
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
      bar.querySelector('.qf-select-bar-all')?.addEventListener('click', selectAllVisible);
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
    // Printed orders are already handled — no urgency. Render neutral.
    if (state.labelStatusFilter === 'printed') return 'neutral';
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
    const buckets = { critical: 0, urgent: 0, watch: 0, safe: 0, neutral: 0 };
    const seenDays = new Map(); // day → zone
    for (const [id, rec] of state.records) {
      if (!passesCarrier(id) || !passesPreOrder(id) || !passesLabelStatus(id)) continue;
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
    // Counts per day that match carrier+preorder+labelStatus (but NOT date filter — this IS the date)
    const counts = new Map();
    for (const [id, rec] of state.records) {
      if (!passesCarrier(id) || !passesPreOrder(id) || !passesLabelStatus(id)) continue;
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
      ${state.labelStatusFilter === 'printed' ? `<div class="qf-zone-archive-note">พิมพ์แล้ว — โหมดข้อมูลย้อนหลัง ไม่แสดงระดับความเร่งด่วน</div>` : (summaryParts.length ? `<div class="qf-zone-summary">${summaryParts.join('')}</div>` : '')}
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
      ${state.labelStatusFilter === 'printed' ? '' : `<div class="qf-zone-legend">
        <span><span class="qf-zone-dot qf-zone-critical-dot"></span>ใกล้/เลย</span>
        <span><span class="qf-zone-dot qf-zone-urgent-dot"></span>รีบ</span>
        <span><span class="qf-zone-dot qf-zone-watch-dot"></span>ระวัง</span>
        <span><span class="qf-zone-dot qf-zone-safe-dot"></span>ปลอดภัย</span>
      </div>`}
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
      // In select mode, inline the live count so users who've scrolled far from
      // the sticky bottom select-bar still see how many items they've picked.
      const selCount = state.selected.size;
      if (state.selectMode) {
        tog.textContent = selCount > 0
          ? `ออกจากโหมดเลือก (เลือกแล้ว ${selCount} รายการ)`
          : 'ออกจากโหมดเลือก';
      } else {
        tog.textContent = 'เลือกหลายรายการ';
      }
    }
    updateSelectionBar();
    window.__qfPlanRefresh?.();
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
        ? [...state.products.values()].map(p => ({
            ...p,
            _count: carrierFilteredSize(p[idsKey]),
            // _totalVisible ignores labelStatus — lets us keep cards visible as
            // greyed-out "done" indicators under the "ยังไม่พิมพ์" filter.
            _totalVisible: carrierFilteredSizeIgnoreLabel(p[idsKey]),
          }))
        : [...state.products.values()].map(p => ({...p, _count: state.currentTab === 'single' ? p.orderCountSingle : p.orderCountMulti}));
      const products = productsRaw
        .filter(p => {
          if (p._count > 0) return true;
          // Keep printed cards visible so user gets done visual feedback.
          // isLabelsDone handles the "printed" filter edge case (returns false).
          return labelsPg && p._totalVisible > 0 && isLabelsDone(p[idsKey]);
        })
        .sort((a, b) => {
          // Push done cards to the bottom so pending ones stay prominent.
          const aDone = a._count === 0;
          const bDone = b._count === 0;
          if (aDone !== bDone) return aDone ? 1 : -1;
          return b._count - a._count;
        });
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
      const variantTotalVisible = (v) => labels
        ? carrierFilteredSizeIgnoreLabel(v[idsKey])
        : (state.currentTab === 'single' ? v.orderCountSingle : v.orderCountMulti);
      const scenario = state.currentTab === 'single' ? 'single' : 'multi';
      for (const p of products) {
        const card = document.createElement('div');
        const cardDone = labels
          ? isLabelsDone(p[idsKey])
          : isDone(p.productId, null, type);
        const productItem = {type: 'product', productId: p.productId, scenario};
        // Card highlights when EITHER the product key OR any variant/qty sub-key
        // is selected — keeps visual aligned with selectAllVisible which inserts
        // variant keys instead of product keys when variants are present.
        const hasVariantSelected = state.selectMode && [...(p.variants?.values() || [])]
          .some(v => v.skuId != null
            && isSelected({type: 'variant', productId: p.productId, skuId: v.skuId, scenario}));
        const hasQtySelected = state.selectMode && scenario === 'multi'
          && [...(p.fulfillUnitIdsByQty?.keys() || [])]
            .some(qty => isSelected({type: 'qty', productId: p.productId, skuId: null, qty}));
        const selected = state.selectMode
          && (isSelected(productItem) || hasVariantSelected || hasQtySelected);
        card.className = 'qf-product-card'
          + (cardDone ? ' qf-done' : '')
          + (state.selectMode ? ' qf-select-mode' : '')
          + (selected ? ' qf-selected' : '');
        card.title = cardDone
          ? `${p.productName}\n\nพิมพ์แล้ว — คลิกเพื่อเลือก: พิมพ์ซ้ำ หรือ ลบเครื่องหมาย ✓`
          : p.productName;
        const variantsRaw = [...p.variants.values()].map(v => ({
          v,
          c: variantCount(v),
          _totalVisible: variantTotalVisible(v),
        }));
        // Skip the synthetic "no-variant" entry (skuId=null) that Shopee
        // items without model_id produce — it's product-level, not a real
        // variant, and rendering a badge for it would show "null".
        // Keep printed variants visible (as greyed-out badges) under
        // "ยังไม่พิมพ์" filter so user still sees what was in this product.
        const variants = variantsRaw.filter(x => {
          if (x.v.skuId == null) return false;
          if (x.c > 0) return true;
          return labels && x._totalVisible > 0 && isLabelsDone(x.v[idsKey]);
        });
        const hasBadges = variants.length >= 1;
        const aliasVal = labels ? (getAlias(p.productId) || '') : '';
        const showVariantToggle = labels && variants.length >= 1;
        const countLabel = (cardDone && p._count === 0)
          ? `พิมพ์แล้ว ${p._totalVisible || ''} ออเดอร์`.replace(/\s+/g, ' ').trim()
          : `${p._count} ออเดอร์`;
        card.innerHTML = `
          <img src="${p.productImageURL}" alt="" referrerpolicy="no-referrer"/>
          <div class="qf-product-name">${escapeHtml(p.productName)}</div>
          <div class="qf-product-count">${countLabel}</div>
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
            setAlias(p.productId, v);
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
          for (const {v, c, _totalVisible: vTotal} of variants) {
            const badgeDone = labels
              ? isLabelsDone(v[idsKey])
              : isDone(p.productId, v.skuId, type);
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
            // When pending count drops to 0 but variant was printed, show
            // the pre-print count so the badge still reads sensibly.
            const badgeCount = (badgeDone && c === 0 && vTotal > 0) ? vTotal : c;
            badge.textContent = `${displayName} (${badgeCount})`;
            const variantItem = {type: 'variant', productId: p.productId, skuId: v.skuId, scenario};
            if (state.selectMode && isSelected(variantItem)) badge.classList.add('qf-selected');
            badge.addEventListener('click', async (e) => {
              e.stopPropagation();
              if (state.selectMode) {
                if (badgeDone) return;
                toggleSelection(variantItem);
                badge.classList.toggle('qf-selected');
                return;
              }
              if (badgeDone) {
                const label = `${p.productName} · ${v.skuName || v.sellerSkuName || v.skuId}`;
                const choice = await showDoneActionModal(label);
                if (choice === 'unmark') {
                  if (labels) {
                    for (const id of v[idsKey] || []) {
                      state.printedUnitIds.delete(id);
                      const rec = state.records.get(id);
                      if (rec) rec.labelStatus = LABEL_STATUS_NOT_PRINTED;
                    }
                  } else {
                    state.doneItems.delete(doneKey(p.productId, v.skuId, type));
                    saveDoneItems();
                  }
                  renderAll();
                } else if (choice === 'reprint') {
                  applyProductFilter(p.productId, v.skuId, type);
                }
                return;
              }
              applyProductFilter(p.productId, v.skuId, type);
            });
            badgesEl.appendChild(badge);
          }
        }
        // §5.2: Qty chips — in multi scenario on labels page, always show so user sees "×N (count)".
        if (labels && scenario === 'multi' && p.fulfillUnitIdsByQty.size > 0) {
          // Build buckets with BOTH the pending list (respects labelStatus filter)
          // and the total visible list (ignores labelStatus) — lets printed chips
          // remain visible as greyed-out "✓ ×N" feedback under the "ยังไม่พิมพ์" view.
          const qtyBuckets = [...p.fulfillUnitIdsByQty.entries()]
            .map(([qty, idSet]) => {
              const arr = [...idSet];
              const totalIds = arr.filter(id => passesCarrier(id) && passesPreOrder(id) && passesDate(id));
              const ids = totalIds.filter(id => passesLabelStatus(id));
              return { qty, ids, totalIds };
            })
            .filter(b => b.totalIds.length > 0)
            .sort((a, b) => a.qty - b.qty);
          if (qtyBuckets.length >= 1) {
            const chipsRow = document.createElement('div');
            chipsRow.className = 'qf-qty-chips';
            for (const b of qtyBuckets) {
              // Done only when EVERY visible id was printed through this session
              // (state.printedUnitIds). Server-side labelStatus=50 is ignored so
              // the chip stays clickable when user scans previously-printed
              // labels intending to reprint.
              const allDone = state.labelStatusFilter !== 'printed'
                && b.totalIds.length > 0
                && b.totalIds.every(id => state.printedUnitIds.has(id));
              // When chip is done and current filter hides printed items, the
              // "ids" (pending list) is empty — swap to totalIds so the reprint
              // handler has something to work with.
              if (allDone && b.ids.length === 0) b.ids = b.totalIds;
              const chip = document.createElement('span');
              chip.className = 'qf-qty-chip' + (allDone ? ' qf-qty-chip--done' : '');
              chip.dataset.qty = b.qty;
              const chipCount = allDone ? b.totalIds.length : b.ids.length;
              chip.textContent = (allDone ? '\u2713 ' : '') + '\u00d7' + b.qty + ' (' + chipCount + ')';
              const qtyItem = { type: 'qty', productId: p.productId, skuId: null, qty: b.qty };
              chip.addEventListener('click', async e => {
                e.stopPropagation();
                if (state.selectMode) {
                  if (allDone) return;
                  toggleSelection(qtyItem);
                  chip.classList.toggle('qf-selected');
                  return;
                }
                const alias = (getAlias(p.productId) || '').trim() || shortName(p.productName);
                if (allDone) {
                  const label = alias + ' \u00d7' + b.qty;
                  const choice = await showDoneActionModal(label);
                  if (choice === 'unmark') {
                    for (const id of b.ids) {
                      state.printedUnitIds.delete(id);
                      const rec = state.records.get(id);
                      if (rec) rec.labelStatus = LABEL_STATUS_NOT_PRINTED;
                    }
                    renderAll();
                  } else if (choice === 'reprint') {
                    const lbl = alias + ' \u00d7' + b.qty;
                    const hint = alias + ' x' + b.qty;
                    await printIds(b.ids, lbl, lbl, hint);
                    renderAll();
                  }
                  return;
                }
                const lbl = alias + ' \u00d7' + b.qty;
                const hint = alias + ' x' + b.qty;
                const ok = await printIds(b.ids, lbl, lbl, hint);
                if (ok) renderAll();
              });
              chipsRow.appendChild(chip);
            }
            card.appendChild(chipsRow);
          }
        }
        card.addEventListener('click', async (e) => {
          if (state.selectMode) {
            if (cardDone) return; // can't select already-printed
            toggleSelection(productItem);
            card.classList.toggle('qf-selected');
            return;
          }
          if (cardDone) {
            const choice = await showDoneActionModal(p.productName);
            if (choice === 'unmark') {
              if (labels) {
                for (const id of p[idsKey] || []) {
                  state.printedUnitIds.delete(id);
                  const rec = state.records.get(id);
                  if (rec) rec.labelStatus = LABEL_STATUS_NOT_PRINTED;
                }
              } else {
                state.doneItems.delete(doneKey(p.productId, null, type));
                saveDoneItems();
              }
              renderAll();
            } else if (choice === 'reprint') {
              applyProductFilter(p.productId, null, type);
            }
            return;
          }
          // §5.3: Multi scenario + labels page + 2+ qty buckets → show รวม/แยก modal.
          if (labels && scenario === 'multi' && p.fulfillUnitIdsByQty.size > 0) {
            const qtyBuckets = [...p.fulfillUnitIdsByQty.entries()]
              .map(([qty, idSet]) => ({ qty, ids: applyCarrierFilter([...idSet]), count: applyCarrierFilter([...idSet]).length }))
              .filter(b => b.count > 0)
              .sort((a, b) => a.qty - b.qty);
            if (qtyBuckets.length > 1) {
              const alias = (getAlias(p.productId) || '').trim() || shortName(p.productName);
              const picked = await showQtyCombineModal({ alias, buckets: qtyBuckets });
              if (!picked) return;
              const choice = picked.mode;
              const withDivider = picked.withDivider;
              const confirm = await showPrintConfirm({
                title: alias,
                summary: choice === 'merge' ? 'รวม ' + qtyBuckets.length + ' กลุ่มจำนวน' : 'แยก ' + qtyBuckets.length + ' ไฟล์',
                count: qtyBuckets.reduce((s, b) => s + b.count, 0),
                sampleText: alias,
              });
              if (!confirm) return;
              const { workerName, workerIcon } = confirm;
              let ok;
              if (choice === 'merge') {
                const totalIds = qtyBuckets.reduce((s, b) => s + b.count, 0);
                // Qty-merge is multi-SKU by definition → always prompt for chunking/picking.
                const plan = await showChunkPlanModal({ total: totalIds, multiSku: true, defaultPickingList: loadPickingListPref() });
                if (!plan) return;
                plan.withDivider = withDivider;
                try {
                  ok = await printQtyBucketsCombined(p.productId, alias, qtyBuckets, workerName, workerIcon, plan);
                } catch (err) {
                  showErrorToast('พิมพ์รวมไม่สำเร็จ: ' + err.message, { source: 'printQtyBucketsCombined', error: String(err) });
                  return;
                }
              } else {
                ok = await printQtyBucketsSplit(p.productId, alias, qtyBuckets, workerName, workerIcon);
              }
              if (ok) renderAll();
              return;
            }
          }
          applyProductFilter(p.productId, null, type);
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
          .map(c => ({
            ...c,
            _count: carrierFilteredSize(c.fulfillUnitIds),
            // _totalVisible ignores labelStatus so we can keep printed combos
            // visible as greyed-out "done" feedback under ยังไม่พิมพ์ filter.
            _totalVisible: carrierFilteredSizeIgnoreLabel(c.fulfillUnitIds),
          }))
          .filter(c => {
            if (c._count > 0) return true;
            return c._totalVisible > 0 && isLabelsDone(c.fulfillUnitIds);
          })
          .sort((a, b) => {
            const aDone = a._count === 0;
            const bDone = b._count === 0;
            if (aDone !== bDone) return aDone ? 1 : -1;
            return b._count - a._count;
          });
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
          const comboDone = isLabelsDone(combo.fulfillUnitIds);
          const comboItem = {type: 'combo', sigKey: combo.sigKey};
          const selected = state.selectMode && isSelected(comboItem);
          card.className = 'qf-combo-card'
            + (comboDone ? ' qf-done' : '')
            + (state.selectMode ? ' qf-select-mode' : '')
            + (selected ? ' qf-selected' : '');
          if (comboDone) {
            card.title = 'พิมพ์แล้ว — คลิกเพื่อเลือก: พิมพ์ซ้ำ หรือ ลบเครื่องหมาย ✓';
          }
          const itemsHtml = combo.items.map((s, idx) => `
            ${idx > 0 ? '<span class="qf-combo-plus">+</span>' : ''}
            <div class="qf-combo-item" data-pid="${s.productId}">
              <img src="${s.productImageURL}" referrerpolicy="no-referrer" title="คลิกเพื่อตั้งชื่อย่อ + ตัวเลือก"/>
              <div class="qf-combo-qty">×${s.quantity}</div>
              <input class="qf-combo-alias-input" type="text" placeholder="ชื่อย่อ" value="${escapeHtml((getAlias(s.productId) || '').trim())}" maxlength="20"/>
            </div>
          `).join('');
          const comboCountLabel = (comboDone && combo._count === 0 && combo._totalVisible > 0)
            ? `พิมพ์แล้ว ${combo._totalVisible} ออเดอร์`
            : `${combo._count} ออเดอร์`;
          card.innerHTML = `
            <div class="qf-combo-row">${itemsHtml}</div>
            <div class="qf-combo-count">${comboCountLabel}</div>
          `;
          // wire alias inputs + edit buttons
          card.querySelectorAll('.qf-combo-item').forEach(itemEl => {
            const pid = itemEl.dataset.pid;
            const inp = itemEl.querySelector('.qf-combo-alias-input');
            const img = itemEl.querySelector('img');
            inp.addEventListener('click', e => e.stopPropagation());
            inp.addEventListener('change', () => {
              const v = inp.value.trim();
              setAlias(pid, v);
            });
            inp.addEventListener('keydown', e => { if (e.key === 'Enter') inp.blur(); });
            img.addEventListener('click', e => {
              e.stopPropagation();
              openAliasModal(pid);
            });
          });
          card.addEventListener('click', async () => {
            if (state.selectMode) {
              if (comboDone) return;
              toggleSelection(comboItem);
              card.classList.toggle('qf-selected');
              return;
            }
            if (comboDone) {
              const label = combo.items.map(i => (state.aliases.get(i.productId) || '').trim() || shortName(i.productName)).join(' + ');
              const choice = await showDoneActionModal(label);
              if (choice === 'unmark') {
                for (const id of combo.fulfillUnitIds) {
                  state.printedUnitIds.delete(id);
                  const rec = state.records.get(id);
                  if (rec) rec.labelStatus = LABEL_STATUS_NOT_PRINTED;
                }
                state.doneItems.delete(comboDoneKey(combo.sigKey));
                saveDoneItems();
                renderAll();
              } else if (choice === 'reprint') {
                printWeirdCombo(combo.sigKey);
              }
              return;
            }
            printWeirdCombo(combo.sigKey);
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

  // ==================== PLANNING PANEL ====================
  const PLANNING_SESSION_KEY = 'qf_planning_session_v1';
  const PLAN_WINDOW_POS_KEY = 'qf_plan_window_pos_v1';
  const PLAN_SNAPSHOTS_KEY = 'qf_plan_snapshots_v1';
  const PLANNING_STUCK_TIMEOUT = 60 * 1000;
  let _planSaveTimer = null;

  // §9.6: Snapshot helpers — one snapshot per (sessionId, date). Retained 30 days.
  function loadPlanSnapshots() {
    try {
      const raw = localStorage.getItem(PLAN_SNAPSHOTS_KEY);
      return raw ? JSON.parse(raw) : [];
    } catch { return []; }
  }

  function savePlanSnapshots(arr) {
    try {
      const cutoff = Date.now() - 30 * 24 * 60 * 60 * 1000;
      const trimmed = arr.filter(s => new Date(s.date).getTime() >= cutoff - 24 * 60 * 60 * 1000);
      localStorage.setItem(PLAN_SNAPSHOTS_KEY, JSON.stringify(trimmed));
    } catch (e) {
      console.warn('[QF] snapshots save failed:', e);
    }
  }

  function _todayDateStr() {
    const d = new Date();
    const pad = n => String(n).padStart(2, '0');
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  }

  function pushPlanSnapshot(session) {
    if (!session || !session.sessionId) return;
    const date = _todayDateStr();
    const existing = loadPlanSnapshots();
    const already = existing.find(s => s.sessionId === session.sessionId && s.date === date);
    // §Team-aware snapshot: include teamId + memberWorkerIds so daily CSV and
    // history lookback can resolve which packers were on the team that day,
    // even if team composition changed afterwards.
    const serializeCol = (col) => ({
      assigneeKind: col.kind,
      assigneeName: col.teamName || col.workerName || null,
      workerId: col.workerId || null,
      teamId: col.kind === 'team' ? (col.teamId || null) : null,
      memberWorkerIds: col.kind === 'team' && Array.isArray(col.memberWorkerIds)
        ? [...col.memberWorkerIds]
        : null,
      ids: [...(col.fulfillUnitIds || [])],
    });
    if (already) {
      // Update in place — session may have changed columns since last save today.
      const updated = existing.map(s => {
        if (s.sessionId !== session.sessionId || s.date !== date) return s;
        const cols = {};
        for (const [cid, col] of Object.entries(session.columns || {})) {
          cols[cid] = serializeCol(col);
        }
        return { ...s, columns: cols, savedAt: Date.now() };
      });
      savePlanSnapshots(updated);
    } else {
      const cols = {};
      for (const [cid, col] of Object.entries(session.columns || {})) {
        cols[cid] = serializeCol(col);
      }
      const snap = { date, sessionId: session.sessionId, columns: cols, savedAt: Date.now() };
      savePlanSnapshots([...existing, snap]);
    }
  }

  function loadPlanningSession() {
    try {
      const raw = localStorage.getItem(PLANNING_SESSION_KEY);
      if (!raw) return null;
      const s = JSON.parse(raw);
      if (!s || (s.version !== 1 && s.version !== 2)) return null;
      if (Date.now() > s.expiresAt) { localStorage.removeItem(PLANNING_SESSION_KEY); return null; }
      const platform = isShopee() ? 'sp' : 'tk';
      if (s.platform !== platform) return null;
      // Migrate old columns: workerColor → workerIcon; add kind/columnId for v1 sessions
      const columns = {};
      for (const [wid, col] of Object.entries(s.columns || {})) {
        let migrated = col;
        // color → icon migration
        if (!migrated.workerIcon && migrated.workerColor) {
          const worker = state.workers.find(w => w.id === wid);
          const { workerColor: _drop, ...rest } = migrated;
          migrated = { ...rest, workerIcon: worker?.icon || WORKER_COLOR_TO_ICON[migrated.workerColor] || '●' };
        }
        // v1 → v2: add kind + columnId.
        // §Team-guard: if `wid` matches a known team id, restore as team column
        // (preserving teamId + memberWorkerIds); otherwise default to worker.
        // Before this guard, all v1 sessions became worker columns on reload.
        if (!migrated.kind) {
          const team = Array.isArray(state.teams) ? state.teams.find(t => t.id === wid) : null;
          if (team) {
            migrated = {
              ...migrated,
              kind: 'team',
              columnId: wid,
              teamId: team.id,
              teamName: team.name,
              memberWorkerIds: Array.isArray(team.memberWorkerIds) ? [...team.memberWorkerIds] : [],
            };
          } else {
            migrated = { ...migrated, kind: 'worker', columnId: wid, workerId: wid };
          }
        }
        columns[migrated.columnId || wid] = migrated;
      }
      return { ...s, version: 2, columns };
    } catch { return null; }
  }

  function savePlanningSession(session) {
    try { localStorage.setItem(PLANNING_SESSION_KEY, JSON.stringify(session)); } catch {}
  }

  function deletePlanningSession() {
    localStorage.removeItem(PLANNING_SESSION_KEY);
  }

  function debouncedSavePlan(session) {
    clearTimeout(_planSaveTimer);
    _planSaveTimer = setTimeout(() => {
      savePlanningSession(session);
      // §9.6: push daily snapshot so buildDailyCsv can derive plannedQty.
      pushPlanSnapshot(session);
    }, 200);
  }

  function savePanelPos(pos) {
    try { localStorage.setItem(PLAN_WINDOW_POS_KEY, JSON.stringify(pos)); } catch {}
  }

  function loadPanelPos() {
    try {
      const raw = localStorage.getItem(PLAN_WINDOW_POS_KEY);
      if (!raw) return null;
      const p = JSON.parse(raw);
      if (typeof p.x !== 'number' || typeof p.y !== 'number') return null;
      return p;
    } catch { return null; }
  }

  function newPlanningSession() {
    const platform = isShopee() ? 'sp' : 'tk';
    const sessionId = 'sess_' + Math.random().toString(36).slice(2, 10);
    const now = Date.now();
    // labelStatus rules:
    //   'all'        → both printed + not-printed in records — exclude printed to avoid accidental reprint.
    //   'printed'    → user deliberately wants to reprint — include everything.
    //   'not_printed'→ processLabelRecord already excluded printed records → include everything.
    //
    // Other filters (carrier chips, pre-order toggle, date range) must ALSO
    // apply — planning mode should mirror what the user currently sees on the
    // grid. applyCarrierFilter() handles all four dimensions in one pass.
    const rawIds = state.labelStatusFilter === 'all'
      ? [...state.records.keys()].filter(id => state.records.get(id)?.labelStatus !== LABEL_STATUS_PRINTED)
      : [...state.records.keys()];
    const allIds = applyCarrierFilter(rawIds);
    return {
      version: 2,
      sessionId,
      platform,
      scannedAt: now,
      expiresAt: now + 24 * 60 * 60 * 1000,
      unassignedIds: allIds,
      columns: {},
    };
  }

  function planTotalIds(session) {
    const assigned = Object.values(session.columns).flatMap(c => c.fulfillUnitIds);
    return session.unassignedIds.length + assigned.length;
  }

  function planPrintedIds(session) {
    return Object.values(session.columns).flatMap(c => c.printedIds);
  }

  // Build card data from session — one card per unique productId:skuId combo in the given id list
  function buildPlanCards(ids) {
    const map = new Map();
    for (const id of ids) {
      const rec = state.records.get(id);
      if (!rec?.skuList?.length) continue;

      if (rec.skuList.length > 1) {
        const sorted = [...rec.skuList].sort((a, b) =>
          String(a.productId).localeCompare(String(b.productId)));
        const sigKey = sorted.map(s => `${s.productId}:${s.quantity}`).join('|');
        const key = `combo:${sigKey}`;
        if (!map.has(key)) {
          const items = sorted.map(s => ({
            productId: s.productId,
            skuId: s.skuId,
            alias: (getAlias(s.productId) || '').trim(),
            officialName: s.productName || '',
            productImageURL: s.productImageURL || null,
            quantity: s.quantity,
          }));
          const aliasParts = items.map(it => it.alias || shortName(it.officialName));
          const comboAlias = aliasParts.join(' + ');
          const comboOfficialName = items
            .map(it => `${shortName(it.officialName) || it.alias || '?'}${it.quantity > 1 ? ` ×${it.quantity}` : ''}`)
            .join(' + ');
          map.set(key, {
            key,
            isCombo: true,
            items,
            productId: null, skuId: null,
            alias: comboAlias,
            officialName: comboOfficialName,
            variantName: '',
            productImageURL: items.find(it => it.productImageURL)?.productImageURL || null,
            name: comboAlias,
            count: 0, ids: [],
          });
        }
        map.get(key).count++;
        map.get(key).ids.push(id);
        continue;
      }

      const s = rec.skuList[0];
      const key = `${s.productId}:${s.skuId}`;
      if (!map.has(key)) {
        const alias = (getAlias(s.productId) || '').trim();
        const variantInfo = getVariantInfo(s.productId, s.skuId);
        const variantName = (variantInfo?.alias || '').trim() || (s.skuName || s.sellerSkuName || '');
        const officialName = s.productName || '';
        const productImageURL = s.productImageURL || null;
        map.set(key, {
          key, productId: s.productId, skuId: s.skuId,
          alias, officialName, variantName, productImageURL,
          name: alias || shortName(officialName),
          count: 0, ids: [],
        });
      }
      map.get(key).count++;
      map.get(key).ids.push(id);
    }
    return [...map.values()];
  }

  function showPlanChunkPopup(workerName, totalAvailable) {
    return new Promise(resolve => {
      if (totalAvailable === 0) { resolve(0); return; }

      // Quick-set buttons: 10/20/30 only when they are strictly less than the total.
      const quickVals = [10, 20, 30].filter(v => v < totalAvailable);
      const quickBtnsHtml = [
        ...quickVals.map(v => `<button class="qf-plan-chunk-quick-btn" data-val="${v}">${v}</button>`),
        `<button class="qf-plan-chunk-quick-btn" data-val="${totalAvailable}">ทั้งหมด ${totalAvailable}</button>`,
      ].join('');

      const overlay = document.createElement('div');
      overlay.className = 'qf-modal-overlay';
      overlay.innerHTML = `
        <div class="qf-modal qf-plan-chunk-popup" role="dialog">
          <div class="qf-plan-chunk-body">
            <div class="qf-plan-chunk-label">กี่ใบให้ ${escapeHtml(workerName)}?</div>
            <div class="qf-plan-chunk-val-display">
              <div class="qf-plan-chunk-val-num">${totalAvailable}</div>
              <div class="qf-plan-chunk-val-unit">ใบ</div>
            </div>
            <div class="qf-plan-chunk-slider-wrap">
              <input type="range" class="qf-plan-chunk-slider" min="1" max="${totalAvailable}" value="${totalAvailable}"/>
              <div class="qf-plan-chunk-ticks"><span>1</span><span>${totalAvailable}</span></div>
            </div>
            <div class="qf-plan-chunk-quick-btns">${quickBtnsHtml}</div>
          </div>
          <div class="qf-modal-actions">
            <button class="qf-btn-cancel">ยกเลิก</button>
            <button class="qf-btn-confirm">ตกลง</button>
          </div>
        </div>
      `;
      document.body.appendChild(overlay);

      const slider = overlay.querySelector('.qf-plan-chunk-slider');
      const valNum = overlay.querySelector('.qf-plan-chunk-val-num');

      function setVal(v) {
        v = Math.max(1, Math.min(totalAvailable, parseInt(v) || totalAvailable));
        slider.value = v;
        valNum.textContent = v;
        const pct = totalAvailable === 1 ? 100 : ((v - 1) / (totalAvailable - 1)) * 100;
        slider.style.background = `linear-gradient(to right, #fe2c55 ${pct}%, #e5e7eb ${pct}%)`;
        overlay.querySelectorAll('.qf-plan-chunk-quick-btn').forEach(b => {
          b.classList.toggle('active', parseInt(b.dataset.val) === v);
        });
      }
      setVal(totalAvailable);

      slider.addEventListener('input', () => setVal(slider.value));
      overlay.querySelectorAll('.qf-plan-chunk-quick-btn').forEach(b => {
        b.addEventListener('click', () => setVal(parseInt(b.dataset.val)));
      });

      const cleanup = v => { overlay.remove(); resolve(v); };
      overlay.querySelector('.qf-btn-cancel').onclick  = () => cleanup(null);
      overlay.querySelector('.qf-btn-confirm').onclick = () => {
        cleanup(Math.max(1, Math.min(totalAvailable, parseInt(slider.value) || totalAvailable)));
      };
      overlay.onclick = e => { if (e.target === overlay) cleanup(null); };
      const onKey = e => {
        if (e.key === 'Escape') { cleanup(null); document.removeEventListener('keydown', onKey); }
      };
      document.addEventListener('keydown', onKey);
    });
  }

  function showPlanAutoSplitModal(session) {
    return new Promise(resolve => {
      const allWorkers = state.workers;
      if (!allWorkers.length) { resolve(null); return; }
      const overlay = document.createElement('div');
      overlay.className = 'qf-modal-overlay';
      overlay.innerHTML = `
        <div class="qf-modal qf-plan-split-modal" role="dialog">
          <div class="qf-modal-title">วิธีแบ่งอัตโนมัติ</div>
          <div class="qf-plan-split-body">
            <label class="qf-plan-split-opt">
              <input type="radio" name="qf-split-mode" value="even" checked/>
              <span>เฉลี่ยทุกใบ (หาร /คน)</span>
            </label>
            <label class="qf-plan-split-opt">
              <input type="radio" name="qf-split-mode" value="sku"/>
              <span>แยกตาม SKU</span>
            </label>
            <div style="margin-top:10px;font-size:12px;font-weight:600;color:#555;">แบ่งให้:</div>
            <div class="qf-plan-workers-check">
              ${allWorkers.map(w => `
                <label class="qf-plan-worker-toggle${session.columns[w.id] ? ' active' : ''}" data-id="${escapeHtml(w.id)}">
                  <span class="qf-plan-worker-icon">${escapeHtml(w.icon)}</span>
                  ${escapeHtml(w.name)}
                  <input type="checkbox" style="display:none;" value="${escapeHtml(w.id)}" ${session.columns[w.id] ? 'checked' : ''}/>
                </label>
              `).join('')}
            </div>
          </div>
          <div class="qf-modal-actions">
            <button class="qf-btn-cancel">ยกเลิก</button>
            <button class="qf-btn-confirm">ใช้</button>
          </div>
        </div>
      `;
      document.body.appendChild(overlay);
      overlay.querySelectorAll('.qf-plan-worker-toggle').forEach(lbl => {
        lbl.addEventListener('click', () => {
          const chk = lbl.querySelector('input[type=checkbox]');
          chk.checked = !chk.checked;
          lbl.classList.toggle('active', chk.checked);
        });
      });
      const cleanup = v => { overlay.remove(); resolve(v); };
      overlay.querySelector('.qf-btn-cancel').onclick = () => cleanup(null);
      overlay.querySelector('.qf-btn-confirm').onclick = () => {
        const mode = overlay.querySelector('input[name="qf-split-mode"]:checked').value;
        const selected = [...overlay.querySelectorAll('.qf-plan-worker-toggle input[type=checkbox]:checked')]
          .map(el => el.value);
        cleanup({ mode, workerIds: selected });
      };
      overlay.onclick = e => { if (e.target === overlay) cleanup(null); };
      const onKey = e => { if (e.key === 'Escape') { cleanup(null); document.removeEventListener('keydown', onKey); } };
      document.addEventListener('keydown', onKey);
    });
  }

  function applyAutoSplit(session, mode, workerIds) {
    const workers = workerIds.map(id => state.workers.find(w => w.id === id)).filter(Boolean);
    if (!workers.length || !session.unassignedIds.length) return session;
    let newSession = { ...session, columns: { ...session.columns }, unassignedIds: [...session.unassignedIds] };

    if (mode === 'even') {
      const ids = [...newSession.unassignedIds];
      const perWorker = Math.ceil(ids.length / workers.length);
      let offset = 0;
      for (const w of workers) {
        const chunk = ids.slice(offset, offset + perWorker);
        offset += perWorker;
        if (!chunk.length) break;
        const col = newSession.columns[w.id] || { kind: 'worker', columnId: w.id, workerId: w.id, workerName: w.name, workerIcon: w.icon, fulfillUnitIds: [], printedIds: [], failedIds: [], status: 'pending', lastPrintAt: null, errorMsg: null, retryCount: 0 };
        newSession.columns[w.id] = { ...col, fulfillUnitIds: [...col.fulfillUnitIds, ...chunk] };
      }
      newSession.unassignedIds = ids.slice(offset);
    } else {
      // by SKU: group unassigned ids by productId:skuId
      const groups = new Map();
      for (const id of newSession.unassignedIds) {
        const rec = state.records.get(id);
        const s = rec?.skuList?.[0];
        const k = s ? `${s.productId}:${s.skuId}` : 'unknown';
        if (!groups.has(k)) groups.set(k, []);
        groups.get(k).push(id);
      }
      // Assign each group to worker with fewest current ids (greedy)
      const workerCounts = workers.map(w => ({ w, count: (newSession.columns[w.id]?.fulfillUnitIds || []).length }));
      for (const [, groupIds] of groups) {
        workerCounts.sort((a, b) => a.count - b.count);
        const { w } = workerCounts[0];
        const col = newSession.columns[w.id] || { kind: 'worker', columnId: w.id, workerId: w.id, workerName: w.name, workerIcon: w.icon, fulfillUnitIds: [], printedIds: [], failedIds: [], status: 'pending', lastPrintAt: null, errorMsg: null, retryCount: 0 };
        newSession.columns[w.id] = { ...col, fulfillUnitIds: [...col.fulfillUnitIds, ...groupIds] };
        workerCounts[0].count += groupIds.length;
      }
      newSession.unassignedIds = [];
    }
    return newSession;
  }

  async function printPlanColumn(session, workerId, ids, renderFn, opts = {}) {
    if (isShopee()) {
      showToast('การพิมพ์จากแผนยังไม่รองรับบน Shopee', 3000);
      return;
    }
    const col = session.columns[workerId];
    if (!col) return;
    const newCol = { ...col, status: 'printing', lastPrintAt: Date.now(), errorMsg: null };
    const newSession = { ...session, columns: { ...session.columns, [workerId]: newCol } };
    debouncedSavePlan(newSession);
    renderFn(newSession);

    try {
      const printIds_ = opts.reprint ? [...ids] : ids.filter(id => !col.printedIds.includes(id));
      if (!printIds_.length) { renderFn({ ...newSession, columns: { ...newSession.columns, [workerId]: { ...newCol, status: 'done' } } }); return; }

      // §3.6: Replace hard-coded CHUNK_SZ=200 with showChunkPlanModal for totals above threshold.
      // §4.8: Detect multi-SKU to offer combined PDF option.
      const skuBuckets = new Set();
      let hasComboRecord = false;
      for (const id of printIds_) {
        const rec = state.records.get(id);
        if (!rec?.skuList?.length) continue;
        if (rec.skuList.length > 1) hasComboRecord = true;
        for (const s of rec.skuList) {
          skuBuckets.add(`${s.productId}:${s.skuId}`);
        }
      }
      const multiSku = hasComboRecord || skuBuckets.size > 1;

      let plan;
      if (opts.plan) {
        plan = opts.plan;
      } else if (!multiSku && printIds_.length <= CHUNK_PROMPT_THRESHOLD) {
        plan = { mode: 'single', withPickingList: loadPickingListPref(), combined: false, withDivider: false };
      } else {
        plan = await showChunkPlanModal({ total: printIds_.length, multiSku, defaultPickingList: loadPickingListPref() });
        if (!plan) {
          renderFn({ ...newSession, columns: { ...newSession.columns, [workerId]: { ...col, status: 'pending' } } });
          return;
        }
      }

      // §8.1: Column name for filename: teamName (if team kind) or workerName.
      const assigneeName = col.teamName || col.workerName || null;
      const assigneeKind = col.kind === 'team' ? 'team' : (assigneeName ? 'worker' : null);

      let ok;
      if (plan.mode === 'by-sku') {
        ok = await printPlanColumnBySku(printIds_, plan, col, assigneeName, assigneeKind, newSession, newCol, workerId, renderFn);
      } else if (multiSku && plan.combined) {
        ok = await printPlanColumnCombined(printIds_, plan, col, assigneeName, assigneeKind, newSession, newCol, workerId, renderFn);
      } else {
        ok = await printPlanColumnChunked(printIds_, plan, col, assigneeName, assigneeKind, newSession, newCol, workerId, renderFn);
      }

      const updatedCol = {
        ...newCol,
        status: ok ? 'done' : 'error',
        printedIds: ok ? [...col.printedIds, ...printIds_] : col.printedIds,
        failedIds: ok ? [] : printIds_,
        lastPrintAt: Date.now(),
        errorMsg: ok ? null : 'พิมพ์ไม่สำเร็จ ลองใหม่',
        retryCount: ok ? col.retryCount : col.retryCount + 1,
      };
      if (ok) {
        for (const id of printIds_) {
          state.printedUnitIds.add(id);
          const rec = state.records.get(id);
          if (rec) rec.labelStatus = LABEL_STATUS_PRINTED;
        }
        for (const combo of state.weirdCombos.values()) {
          if ([...combo.fulfillUnitIds].every(cid => state.printedUnitIds.has(cid))) {
            markComboDone(combo.sigKey);
          }
        }
        try { if (typeof renderAll === 'function') renderAll(); } catch (_) {}
      }
      const finalSession = { ...newSession, columns: { ...newSession.columns, [workerId]: updatedCol } };
      debouncedSavePlan(finalSession);
      renderFn(finalSession);
    } catch (e) {
      const errCol = { ...newCol, status: 'error', errorMsg: e.message || 'ผิดพลาด', retryCount: col.retryCount + 1 };
      const errSession = { ...newSession, columns: { ...newSession.columns, [workerId]: errCol } };
      debouncedSavePlan(errSession);
      renderFn(errSession);
    }
  }

  // Build team snapshot from col at print time (memberWorkerIds → [{workerId, workerName, icon}]).
  function buildTeamSnapshot(col) {
    if (col.kind !== 'team' || !col.memberWorkerIds?.length) return null;
    return col.memberWorkerIds.map(wid => {
      const w = state.workers.find(x => x.id === wid);
      return { workerId: wid, workerName: w?.name || wid, icon: w?.icon || '' };
    });
  }

  // Helper: plan-column print — 1 PDF chunk per unique SKU group.
  // Returns true on success, false on failure.
  async function printPlanColumnBySku(printIds_, plan, col, assigneeName, assigneeKind, newSession, newCol, workerId, renderFn) {
    const groupMap = new Map();
    for (const id of printIds_) {
      const rec = state.records.get(id);
      if (!rec?.skuList?.length) continue;
      const s = rec.skuList[0];
      const key = `${s.productId}:${s.skuId}`;
      if (!groupMap.has(key)) {
        const alias = (getAlias(s.productId) || '').trim();
        const variantInfo = getVariantInfo(s.productId, s.skuId);
        groupMap.set(key, {
          alias: alias || shortName(s.productName),
          variantName: (variantInfo?.alias || '').trim() || (s.skuName || s.sellerSkuName || ''),
          ids: [],
        });
      }
      groupMap.get(key).ids.push(id);
    }
    const groups = [...groupMap.values()].sort((a, b) => a.alias.localeCompare(b.alias));

    const baseHint = col.teamName || col.workerName || 'plan';
    const baseFilename = assigneeName
      ? makeBaseFilename(`[${baseHint}] [${assigneeName}]`)
      : makeBaseFilename(baseHint);
    const displayTitle = `พิมพ์แยก SKU: ${col.teamName || col.workerName || 'แผน'}`;

    const chunks = groups.map((grp, i) => {
      const skuLabel = grp.variantName ? `${grp.alias} - ${grp.variantName}` : grp.alias;
      return {
        ids: grp.ids,
        label: skuLabel || `SKU ${i + 1}`,
        filename: makeBaseFilename(`[${baseHint}] [${skuLabel || `SKU ${i + 1}`}]`) + '.pdf',
      };
    });

    return runChunkedExport(chunks, displayTitle, {
      baseFilename,
      totalLabels: printIds_.length,
      workerId: col.workerId,
      workerName: col.teamName || col.workerName || null,
      workerIcon: col.workerIcon || null,
      assigneeKind,
      assigneeName,
      teamId: col.kind === 'team' ? (col.teamId || null) : null,
      teamName: col.kind === 'team' ? (col.teamName || null) : null,
      teamSnapshot: buildTeamSnapshot(col),
      withPickingList: plan.withPickingList || false,
      withDivider: false,
    });
  }

  // Helper: plan-column print using per-group combined PDF.
  // Returns true on success, false on failure.
  async function printPlanColumnCombined(printIds_, plan, col, assigneeName, assigneeKind, newSession, newCol, workerId, renderFn) {
    // Build groups sorted by alias for combined PDF.
    const groupMap = new Map();
    for (const id of printIds_) {
      const rec = state.records.get(id);
      if (!rec?.skuList?.length) continue;
      const s = rec.skuList[0];
      const key = `${s.productId}:${s.skuId}`;
      if (!groupMap.has(key)) {
        const alias = (getAlias(s.productId) || '').trim();
        const variantInfo = getVariantInfo(s.productId, s.skuId);
        groupMap.set(key, {
          productId: s.productId,
          skuId: s.skuId,
          alias: alias || shortName(s.productName),
          officialName: s.productName || '',
          variantName: (variantInfo?.alias || '').trim() || (s.skuName || s.sellerSkuName || ''),
          productImageURL: s.productImageURL || null,
          ids: [],
        });
      }
      groupMap.get(key).ids.push(id);
    }
    const groups = [...groupMap.values()].sort((a, b) => a.alias.localeCompare(b.alias));

    // Slice printIds_ into plan chunks; each chunk gets its own combined PDF.
    const slices = planSlice(printIds_, plan);
    const chunkCount = slices.length;
    const baseHint = `รวม ${groups.length} SKU`;
    const baseFilename = assigneeName
      ? makeBaseFilename(`[${baseHint}] [${assigneeName}]`)
      : makeBaseFilename(baseHint);
    const displayTitle = `พิมพ์รวม ${groups.length} SKU: ${col.teamName || col.workerName || 'แผน'}`;

    const chunks = slices.map((slice, i) => {
      const idx = i + 1;
      const chunkSuffix = chunkCount > 1 ? `-ชุด${idx}-${chunkCount}` : '';
      // Re-group the slice for combined PDF building.
      const sliceGroupMap = new Map();
      for (const id of slice) {
        const rec = state.records.get(id);
        if (!rec?.skuList?.length) continue;
        const s = rec.skuList[0];
        const key = `${s.productId}:${s.skuId}`;
        if (!sliceGroupMap.has(key)) {
          const grp = groups.find(g => g.productId === s.productId && g.skuId === s.skuId);
          sliceGroupMap.set(key, { ...grp, ids: [] });
        }
        sliceGroupMap.get(key).ids.push(id);
      }
      const sliceGroups = [...sliceGroupMap.values()].sort((a, b) => a.alias.localeCompare(b.alias));
      return {
        ids: slice,
        label: chunkCount === 1 ? 'ไฟล์เดียว' : `ชุด ${idx}/${chunkCount}`,
        filename: `${baseFilename}${chunkSuffix}.pdf`,
        sliceGroups,
        workerName: col.teamName || col.workerName || null,
        workerIcon: col.workerIcon || null,
        withPickingList: plan.withPickingList,
        withDivider: plan.withDivider,
      };
    });

    // Build combined PDFs using buildMultiSkuCombinedPdf.
    const workerNameForPdf = col.teamName || col.workerName || null;
    const workerIconForPdf = col.workerIcon || null;

    const exportChunks = await Promise.all(chunks.map(async (chunk) => {
      try {
        const { bytes } = await buildMultiSkuCombinedPdf(
          chunk.sliceGroups,
          workerNameForPdf,
          workerIconForPdf,
          chunk.withPickingList,
          () => {},
          chunk.withDivider,
          assigneeKind
        );
        return { ...chunk, prebuiltBytes: bytes };
      } catch (e) {
        return { ...chunk, prebuiltBytes: null, buildError: e };
      }
    }));

    // Fall back to flat runChunkedExport if any combined build failed.
    const allBuilt = exportChunks.every(c => c.prebuiltBytes);
    if (!allBuilt) {
      return printPlanColumnChunked(printIds_, plan, col, assigneeName, assigneeKind, newSession, newCol, workerId, renderFn);
    }

    // Use runChunkedExport with prebuilt PDF bytes (combined-per-column, per-chunk).
    // runChunkedExport reads `prebuiltBytes` at L4919 and skips buildChunkPdf entirely.
    const runChunks = exportChunks.map(c => ({
      ids: c.ids, label: c.label, filename: c.filename, prebuiltBytes: c.prebuiltBytes,
    }));
    return runChunkedExport(runChunks, displayTitle, {
      baseFilename,
      totalLabels: printIds_.length,
      workerId: col.workerId,
      workerName: workerNameForPdf,
      workerIcon: workerIconForPdf,
      assigneeKind,
      assigneeName,
      teamId: col.kind === 'team' ? (col.teamId || null) : null,
      teamName: col.kind === 'team' ? (col.teamName || null) : null,
      teamSnapshot: buildTeamSnapshot(col),
    });
  }

  // Helper: plan-column print using simple per-chunk flat export (non-combined or single-SKU).
  async function printPlanColumnChunked(printIds_, plan, col, assigneeName, assigneeKind, newSession, newCol, workerId, renderFn) {
    const slices = planSlice(printIds_, plan);
    const chunkCount = slices.length;
    const baseHint = col.teamName || col.workerName || 'plan';
    const baseFilename = assigneeName
      ? makeBaseFilename(`[${baseHint}] [${assigneeName}]`)
      : makeBaseFilename(baseHint);
    const displayTitle = `พิมพ์: ${col.teamName || col.workerName || 'แผน'}`;

    const chunks = slices.map((slice, i) => {
      const idx = i + 1;
      const chunkSuffix = chunkCount > 1 ? `-ชุด${idx}-${chunkCount}` : '';
      return {
        ids: slice,
        label: chunkCount === 1 ? 'ไฟล์เดียว' : `ชุด ${idx}/${chunkCount}`,
        filename: `${baseFilename}${chunkSuffix}.pdf`,
      };
    });

    return runChunkedExport(chunks, displayTitle, {
      baseFilename,
      totalLabels: printIds_.length,
      workerId: col.workerId,
      workerName: col.teamName || col.workerName || null,
      workerIcon: col.workerIcon || null,
      assigneeKind,
      assigneeName,
      teamId: col.kind === 'team' ? (col.teamId || null) : null,
      teamName: col.kind === 'team' ? (col.teamName || null) : null,
      teamSnapshot: buildTeamSnapshot(col),
      withPickingList: plan.withPickingList || false,
      withDivider: plan.withDivider || false,
    });
  }

  // Print all pending columns — ONE runChunkedExport call (single result modal).
  // plan.mode: 'by-person' | 'by-person-sku' | 'combined-per-person' | 'single'
  async function printAllPlanColumns(session, nonDone, plan, renderFn) {
    let updatedSession = nonDone.reduce((s, [cid, col]) => ({
      ...s, columns: { ...s.columns, [cid]: { ...col, status: 'printing' } },
    }), session);
    renderFn(updatedSession);

    const allPrintIds = nonDone.flatMap(([, col]) =>
      col.fulfillUnitIds.filter(id => !col.printedIds.includes(id))
    );

    const names = nonDone.map(([, c]) => c.teamName || c.workerName || '?');
    const baseFilename = makeBaseFilename(`พิมพ์ทุกคน-${names.slice(0, 3).join('-')}${names.length > 3 ? '+' : ''}`);
    let ok = false;
    try {
      const mode = plan.mode || 'by-person';

      if (mode === 'by-person' || mode === 'single') {
        // Each worker gets their own flat PDF (optionally chunked if chunkAt set).
        // §Team-aware: attach per-chunk {workerName, workerIcon, assigneeKind}
        // so runChunkedExport can render the correct divider ("ทีม:" vs "ผู้แพ็ค:")
        // and apply per-column overlay watermark even when chunks from multiple
        // columns are merged into one parallel Promise.all batch.
        const runChunks = [];
        for (const [, col] of nonDone) {
          const colIds = col.fulfillUnitIds.filter(id => !col.printedIds.includes(id));
          const workerName = col.teamName || col.workerName || null;
          const workerIcon = col.workerIcon || null;
          const kind = col.kind || 'worker';
          const hint = workerName || 'plan';
          const chunkMeta = { workerName, workerIcon, assigneeKind: kind };
          if (plan.chunkAt && colIds.length > plan.chunkAt) {
            const sz = plan.chunkAt;
            let ci = 0;
            for (let i = 0; i < colIds.length; i += sz) {
              ci++;
              runChunks.push({
                ...chunkMeta,
                ids: colIds.slice(i, i + sz),
                label: `${hint} ชุด ${ci}`,
                filename: makeBaseFilename(`[${hint}]-ชุด${ci}`) + '.pdf',
              });
            }
          } else {
            runChunks.push({
              ...chunkMeta,
              ids: colIds,
              label: hint,
              filename: makeBaseFilename(`[${hint}]`) + '.pdf',
            });
          }
        }
        ok = await runChunkedExport(runChunks, `พิมพ์ทุกคน ${runChunks.length} ไฟล์`, {
          baseFilename, totalLabels: allPrintIds.length,
          withPickingList: plan.withPickingList || false, withDivider: plan.withDivider || false,
        });

      } else if (mode === 'by-person-sku') {
        // Each worker × SKU = 1 chunk.
        // §Team-aware: same per-chunk identity attachment as by-person.
        const runChunks = [];
        for (const [, col] of nonDone) {
          const colIds = col.fulfillUnitIds.filter(id => !col.printedIds.includes(id));
          const workerName = col.teamName || col.workerName || null;
          const workerIcon = col.workerIcon || null;
          const kind = col.kind || 'worker';
          const hint = workerName || 'plan';
          const chunkMeta = { workerName, workerIcon, assigneeKind: kind };
          const skuMap = new Map();
          for (const id of colIds) {
            const rec = state.records.get(id);
            if (!rec?.skuList?.length) continue;
            const s = rec.skuList[0];
            const key = `${s.productId}:${s.skuId}`;
            if (!skuMap.has(key)) {
              const alias = (getAlias(s.productId) || '').trim();
              const variantInfo = getVariantInfo(s.productId, s.skuId);
              skuMap.set(key, {
                alias: alias || shortName(s.productName),
                variantName: (variantInfo?.alias || '').trim() || (s.skuName || s.sellerSkuName || ''),
                ids: [],
              });
            }
            skuMap.get(key).ids.push(id);
          }
          for (const grp of [...skuMap.values()].sort((a, b) => a.alias.localeCompare(b.alias))) {
            const skuLabel = grp.variantName ? `${grp.alias} - ${grp.variantName}` : grp.alias;
            runChunks.push({
              ...chunkMeta,
              ids: grp.ids,
              label: workerName ? `${workerName} · ${skuLabel}` : skuLabel,
              filename: makeBaseFilename(`[${hint}] [${skuLabel || 'SKU'}]`) + '.pdf',
            });
          }
        }
        ok = await runChunkedExport(runChunks, `พิมพ์แยก SKU ${runChunks.length} ไฟล์`, {
          baseFilename, totalLabels: allPrintIds.length,
          withPickingList: plan.withPickingList || false, withDivider: false,
        });

      } else if (mode === 'combined-per-person') {
        // Each worker gets a prebuilt combined-SKU PDF — run in parallel with
        // concurrency cap 3 to avoid RAM spike + main-thread overload.
        const exportChunks = new Array(nonDone.length);
        const prepProgress = showProgress(`กำลังเตรียม PDF รวม (${nonDone.length} คน · ${allPrintIds.length} ฉลาก)`);
        const workerProgress = new Float64Array(nonDone.length);
        const CONCURRENCY = 3;
        const reportAggregate = () => {
          const sum = workerProgress.reduce((a, b) => a + b, 0);
          const pct = (sum / nonDone.length) * 100;
          const active = [...workerProgress].filter(p => p > 0 && p < 1).length;
          prepProgress.update(pct, active > 0
            ? `กำลังสร้าง ${active} ไฟล์พร้อมกัน · ${pct.toFixed(0)}%`
            : 'พร้อมแล้ว');
        };
        try {
          console.time('[QF] combined-per-person parallel build');
          // Simple promise-pool: keep ≤CONCURRENCY in-flight at any time.
          let nextIdx = 0;
          const runOne = async () => {
            while (nextIdx < nonDone.length) {
              const wi = nextIdx++;
              const [, col] = nonDone[wi];
              const colIds = col.fulfillUnitIds.filter(id => !col.printedIds.includes(id));
              const workerName = col.teamName || col.workerName || null;
              const workerIcon = col.workerIcon || null;
              const kind = col.kind || 'worker';
              const hint = workerName || 'plan';
              const groupMap = new Map();
              for (const id of colIds) {
                const rec = state.records.get(id);
                if (!rec?.skuList?.length) continue;
                const s = rec.skuList[0];
                const key = `${s.productId}:${s.skuId}`;
                if (!groupMap.has(key)) {
                  const alias = (getAlias(s.productId) || '').trim();
                  const variantInfo = getVariantInfo(s.productId, s.skuId);
                  groupMap.set(key, {
                    productId: s.productId, skuId: s.skuId,
                    alias: alias || shortName(s.productName),
                    officialName: s.productName || '',
                    variantName: (variantInfo?.alias || '').trim() || (s.skuName || s.sellerSkuName || ''),
                    productImageURL: s.productImageURL || null,
                    ids: [],
                  });
                }
                groupMap.get(key).ids.push(id);
              }
              const groups = [...groupMap.values()].sort((a, b) => a.alias.localeCompare(b.alias));
              workerProgress[wi] = 0.01; // mark in-flight
              reportAggregate();
              try {
                const { bytes } = await buildMultiSkuCombinedPdf(
                  groups, workerName, workerIcon, plan.withPickingList,
                  (pct) => { workerProgress[wi] = pct / 100; reportAggregate(); },
                  plan.withDivider,
                  kind
                );
                workerProgress[wi] = 1;
                reportAggregate();
                exportChunks[wi] = {
                  ids: colIds, label: hint, filename: makeBaseFilename(`[${hint}]`) + '.pdf',
                  prebuiltBytes: bytes,
                  workerName, workerIcon, assigneeKind: kind,
                };
              } catch (_) {
                workerProgress[wi] = 1;
                reportAggregate();
                exportChunks[wi] = {
                  ids: colIds, label: hint, filename: makeBaseFilename(`[${hint}]`) + '.pdf',
                  workerName, workerIcon, assigneeKind: kind,
                };
              }
            }
          };
          await Promise.all(Array.from({ length: Math.min(CONCURRENCY, nonDone.length) }, () => runOne()));
          console.timeEnd('[QF] combined-per-person parallel build');
          prepProgress.update(100, 'พร้อมแล้ว');
        } finally {
          document.querySelectorAll('.qf-progress-overlay').forEach(e => e.remove());
        }
        ok = await runChunkedExport(
          exportChunks.map(c => ({
            ids: c.ids, label: c.label, filename: c.filename,
            prebuiltBytes: c.prebuiltBytes,
            // Per-chunk identity preserved so divider+overlay on the non-prebuilt
            // fallback path (if any chunk failed to prebuild) still uses the
            // correct column's worker/team context.
            workerName: c.workerName, workerIcon: c.workerIcon, assigneeKind: c.assigneeKind,
          })),
          `พิมพ์รวม ${nonDone.length} คน`,
          { baseFilename, totalLabels: allPrintIds.length, withPickingList: plan.withPickingList || false, withDivider: plan.withDivider || false }
        );
      }
    } catch (_) {
      ok = false;
    }

    if (ok) {
      for (const id of allPrintIds) {
        state.printedUnitIds.add(id);
        const rec = state.records.get(id);
        if (rec) rec.labelStatus = LABEL_STATUS_PRINTED;
      }
      for (const combo of state.weirdCombos.values()) {
        if ([...combo.fulfillUnitIds].every(id => state.printedUnitIds.has(id))) markComboDone(combo.sigKey);
      }
      try { if (typeof renderAll === 'function') renderAll(); } catch (_) {}
    }

    for (const [cid, col] of nonDone) {
      const colPrintIds = allPrintIds.filter(id => col.fulfillUnitIds.includes(id));
      const updatedCol = {
        ...col,
        status: ok ? 'done' : 'error',
        printedIds: ok ? [...col.printedIds, ...colPrintIds] : col.printedIds,
        failedIds: ok ? [] : colPrintIds,
        lastPrintAt: Date.now(),
        errorMsg: ok ? null : 'พิมพ์ไม่สำเร็จ ลองใหม่',
        retryCount: ok ? col.retryCount : col.retryCount + 1,
      };
      updatedSession = { ...updatedSession, columns: { ...updatedSession.columns, [cid]: updatedCol } };
    }
    debouncedSavePlan(updatedSession);
    renderFn(updatedSession);
  }

  // Utility: slice a flat array of ids per ChunkPlan into sub-arrays.
  function planSlice(ids, plan) {
    const total = ids.length;
    switch (plan.mode) {
      case 'even': {
        const n = plan.n || 1;
        const sz = Math.ceil(total / n);
        const slices = [];
        for (let i = 0; i < ids.length; i += sz) slices.push(ids.slice(i, i + sz));
        return slices;
      }
      case 'every': {
        const x = plan.x || CHUNK_AUTO_SAFE_SIZE;
        const slices = [];
        for (let i = 0; i < ids.length; i += x) slices.push(ids.slice(i, i + x));
        return slices;
      }
      default:
        return [ids];
    }
  }

  // ── Order-type helpers for planning panel ─────────────────────────────────
  // Returns 'single' | 'multi' | 'weird' for a single fulfillUnitId.
  function getOrderType(id) {
    if (state.weirdFulfillUnitIds.has(id)) return 'weird';
    const rec = state.records.get(id);
    if (!rec?.skuList?.length) return 'single';
    if (rec.skuList.length > 1) return 'weird';
    return rec.skuList[0].quantity === 1 ? 'single' : 'multi';
  }

  // Dominant type for a card (may contain many IDs). weird > multi > single.
  function getCardOrderType(card) {
    if (card.isCombo) return 'weird';
    let hasMult = false;
    for (const id of (card.ids || [])) {
      const t = getOrderType(id);
      if (t === 'weird') return 'weird';
      if (t === 'multi') hasMult = true;
    }
    return hasMult ? 'multi' : 'single';
  }

  function openPlanningPanel(initialSession) {
    // Remove any lingering window or bubble from a previous open call
    document.querySelectorAll('.qf-plan-window, .qf-plan-bubble').forEach(e => e.remove());

    if (state.records.size === 0) {
      showToast('ยังไม่ได้ scan — กด Scan ก่อน', 2500);
      return;
    }

    let session = initialSession || newPlanningSession();

    // Ensure all current workers have a column entry (kind='worker')
    for (const w of state.workers) {
      if (!session.columns[w.id]) {
        session = {
          ...session,
          columns: {
            ...session.columns,
            [w.id]: {
              kind: 'worker', columnId: w.id, workerId: w.id,
              workerName: w.name, workerIcon: w.icon,
              fulfillUnitIds: [], printedIds: [], failedIds: [],
              status: 'pending', lastPrintAt: null, errorMsg: null, retryCount: 0,
            },
          },
        };
      }
    }

    // Fix stuck printing columns on load
    for (const [cid, col] of Object.entries(session.columns)) {
      if (col.status === 'printing' && col.lastPrintAt && Date.now() - col.lastPrintAt > PLANNING_STUCK_TIMEOUT) {
        session = { ...session, columns: { ...session.columns, [cid]: { ...col, status: 'error', errorMsg: 'อาจถูกขัดจังหวะ กดพิมพ์ซ้ำ' } } };
      }
    }

    debouncedSavePlan(session);
    window.__qfPlanningSession = () => session;

    // ---- Order-type tab state (persists across render() calls) ----
    let activeOrderTab = 'all';   // 'all' | 'single' | 'multi' | 'weird'

    // ---- Floating window DOM ----
    const win = document.createElement('div');
    win.className = 'qf-plan-window';

    // Restore or compute initial position
    const savedPos = loadPanelPos();
    const defaultW = 620, defaultH = 520;
    const vw = window.innerWidth, vh = window.innerHeight;
    const initX = savedPos ? Math.min(Math.max(0, savedPos.x), vw - defaultW) : Math.round((vw - defaultW) / 2);
    const initY = savedPos ? Math.min(Math.max(0, savedPos.y), vh - 60) : Math.round((vh - defaultH) / 2);
    win.style.left = initX + 'px';
    win.style.top = initY + 'px';

    document.body.appendChild(win);

    let _dragCardKey = null;
    let _dragSourceZone = null;

    // Re-usable: wire dragstart/dragend on every .qf-plan-card[draggable] inside container.
    // Called both by attachPlanEvents (full re-render) and applyUnassignedTab (tab switch).
    const wireDragListeners = (container) => {
      container.querySelectorAll('.qf-plan-card[draggable]').forEach(card => {
        card.addEventListener('dragstart', e => {
          _dragCardKey = card.dataset.cardKey;
          _dragSourceZone = card.closest('[data-zone]')?.dataset.zone;
          card.classList.add('qf-dragging');
          e.dataTransfer.effectAllowed = 'move';
        });
        card.addEventListener('dragend', () => {
          card.classList.remove('qf-dragging');
          win.querySelectorAll('.qf-plan-cards').forEach(el => el.classList.remove('qf-drag-over'));
        });
      });
    };

    // ---- buildPlanCardNode: returns a DOM element ----
    const buildPlanCardNode = (c, allCards) => {
      // Compute split badge
      const totalAcross = allCards
        .filter(other => other.key === c.key)
        .reduce((sum, other) => sum + other.count, 0);
      const isSplit = totalAcross > c.count;

      const card = document.createElement('div');
      card.className = 'qf-plan-card' + (c.isCombo ? ' qf-plan-card--combo' : '');
      card.draggable = true;
      card.dataset.cardKey = c.key;
      card.dataset.cardCount = String(c.count);
      card.title = c.isCombo
        ? `ออเดอร์รวม: ${c.items.map(it => `${it.alias || shortName(it.officialName)} ×${it.quantity}`).join(' + ')} (${c.count} ใบ)`
        : `${c.name} (${c.count} ใบ)`;
      card.style.position = 'relative';

      if (c.isCombo && c.items?.length >= 2) {
        const stack = document.createElement('div');
        stack.className = 'qf-plan-card-img qf-plan-card-img--combo';
        const topN = c.items.slice(0, 3);
        topN.forEach((it, i) => {
          if (it.productImageURL) {
            const mini = document.createElement('img');
            mini.src = it.productImageURL;
            mini.className = 'qf-plan-card-img-mini';
            mini.style.zIndex = String(10 - i);
            mini.onerror = () => {
              const fall = document.createElement('div');
              fall.className = 'qf-plan-card-img-mini qf-plan-card-img-mini--fallback';
              fall.style.zIndex = String(10 - i);
              fall.textContent = ((it.alias || shortName(it.officialName) || '?')[0] || '?').toUpperCase();
              mini.replaceWith(fall);
            };
            stack.appendChild(mini);
          } else {
            const fall = document.createElement('div');
            fall.className = 'qf-plan-card-img-mini qf-plan-card-img-mini--fallback';
            fall.style.zIndex = String(10 - i);
            fall.textContent = ((it.alias || shortName(it.officialName) || '?')[0] || '?').toUpperCase();
            stack.appendChild(fall);
          }
        });
        if (c.items.length > 3) {
          const more = document.createElement('div');
          more.className = 'qf-plan-card-img-more';
          more.textContent = `+${c.items.length - 3}`;
          stack.appendChild(more);
        }
        card.appendChild(stack);
      } else {
        const img = document.createElement('img');
        img.className = 'qf-plan-card-img';
        const imgUrl = c.productImageURL;
        if (imgUrl) {
          img.src = imgUrl;
          img.onerror = () => {
            img.remove();
            const fallback = buildCardImgFallback(c.alias || c.name);
            card.insertBefore(fallback, card.firstChild);
          };
          card.appendChild(img);
        } else {
          const fallback = buildCardImgFallback(c.alias || c.name);
          card.appendChild(fallback);
        }
      }

      // Text area
      const textDiv = document.createElement('div');
      textDiv.className = 'qf-plan-card-text';

      const aliasEl = document.createElement('div');
      aliasEl.className = 'qf-plan-card-alias';
      aliasEl.textContent = c.alias || c.name;
      textDiv.appendChild(aliasEl);

      if (c.officialName && c.officialName !== (c.alias || c.name)) {
        const nameEl = document.createElement('div');
        nameEl.className = 'qf-plan-card-name';
        nameEl.textContent = c.officialName;
        textDiv.appendChild(nameEl);
      }

      if (c.variantName) {
        const varEl = document.createElement('div');
        varEl.className = 'qf-plan-card-variant';
        varEl.textContent = c.variantName;
        textDiv.appendChild(varEl);
      }

      card.appendChild(textDiv);

      // Badge
      const badge = document.createElement('span');
      badge.className = 'qf-plan-card-badge' + (isSplit ? ' qf-plan-card-badge--split' : '');
      badge.textContent = isSplit ? `${c.count}/${totalAcross}` : String(c.count);
      card.appendChild(badge);

      // Order-type tag pill (top-right corner)
      const orderType = getCardOrderType(c);
      const typeLabels = { single: '1 ชิ้น', multi: 'หลายชิ้น', weird: 'แปลก' };
      const typeTag = document.createElement('span');
      typeTag.className = `qf-plan-card-type-tag qf-type-${orderType}`;
      typeTag.textContent = typeLabels[orderType];
      card.appendChild(typeTag);

      return card;
    };

    // ---- Fallback image square ----
    const buildCardImgFallback = (label) => {
      const sq = document.createElement('div');
      sq.className = 'qf-plan-card-img';
      sq.style.background = '#e5e7eb';
      sq.style.display = 'flex';
      sq.style.alignItems = 'center';
      sq.style.justifyContent = 'center';
      sq.style.fontSize = '18px';
      sq.style.fontWeight = '700';
      sq.style.color = '#9ca3af';
      sq.style.flexShrink = '0';
      sq.textContent = (label || '?')[0].toUpperCase();
      return sq;
    };

    // ---- renderPlanCardsInto: fills a container with card DOM nodes ----
    const renderPlanCardsInto = (container, ids, allSessionCards) => {
      container.innerHTML = '';
      const cards = buildPlanCards(ids);
      if (!cards.length) {
        const hint = document.createElement('div');
        hint.className = 'qf-plan-empty-hint';
        hint.textContent = 'ว่าง — ลากการ์ดมาวางที่นี่';
        container.appendChild(hint);
        return;
      }
      for (const c of cards) {
        const node = buildPlanCardNode(c, allSessionCards);
        container.appendChild(node);
      }
    };

    // ---- renderColumnZone: creates a zone element for a worker or team column ----
    const renderColumnZone = (cid, col, s) => {
      const locked = col.status === 'printing' || col.status === 'done';
      const statusLabel = { pending: 'รอพิมพ์', printing: 'กำลังพิมพ์...', done: 'เสร็จ', partial: 'บางส่วน', error: 'ผิดพลาด' }[col.status] || col.status;
      const isTeam = col.kind === 'team';

      const zone = document.createElement('div');
      zone.className = 'qf-plan-zone ' + (isTeam ? 'qf-zone-team' : 'qf-zone-worker');
      zone.dataset.wzone = cid;

      const header = document.createElement('div');
      header.className = 'qf-plan-zone-header';

      // Top row: badge + title + status + edit/delete
      const headerRow = document.createElement('div');
      headerRow.className = 'qf-plan-zone-header-row';


      const titleWrap = document.createElement('div');
      titleWrap.style.flex = '1';
      titleWrap.style.minWidth = '0';
      const title = document.createElement('span');
      title.className = 'qf-plan-zone-title';
      title.textContent = (isTeam ? col.teamName : col.workerName) + ` (${col.fulfillUnitIds.length})`;
      titleWrap.appendChild(title);
      if (isTeam && col.memberWorkerIds && col.memberWorkerIds.length) {
        const memberNames = col.memberWorkerIds
          .map(wid => state.workers.find(w => w.id === wid)?.name || wid)
          .slice(0, 2);
        const extra = col.memberWorkerIds.length > 2 ? ` +${col.memberWorkerIds.length - 2}` : '';
        const sub = document.createElement('div');
        sub.style.fontSize = '10px';
        sub.style.color = '#888';
        sub.style.overflow = 'hidden';
        sub.style.whiteSpace = 'nowrap';
        sub.style.textOverflow = 'ellipsis';
        sub.textContent = memberNames.join(' · ') + extra;
        titleWrap.appendChild(sub);
      }
      headerRow.appendChild(titleWrap);

      const statusEl = document.createElement('span');
      statusEl.className = `qf-plan-zone-status ${col.status}`;
      statusEl.textContent = statusLabel;
      headerRow.appendChild(statusEl);

      const editBtn = document.createElement('button');
      editBtn.className = 'qf-plan-zone-remove';
      editBtn.dataset.action = 'editcol';
      editBtn.dataset.wid = cid;
      editBtn.title = isTeam ? 'แก้ไขทีม' : 'แก้ไขคน';
      editBtn.textContent = '✎';
      headerRow.appendChild(editBtn);

      const delBtn = document.createElement('button');
      delBtn.className = 'qf-plan-zone-remove';
      delBtn.dataset.action = 'removecol';
      delBtn.dataset.wid = cid;
      delBtn.title = isTeam ? 'ลบทีม' : 'ลบคน';
      delBtn.textContent = '🗑';
      headerRow.appendChild(delBtn);

      header.appendChild(headerRow);

      // Action row: print / reprint / reset / retry / return
      const actions = document.createElement('div');
      actions.className = 'qf-plan-zone-actions';
      const addBtn = (label, cls, action) => {
        const b = document.createElement('button');
        b.className = `qf-plan-zone-btn${cls ? ' ' + cls : ''}`;
        b.dataset.action = action;
        b.dataset.wid = cid;
        b.textContent = label;
        actions.appendChild(b);
      };

      if (col.status === 'pending') {
        addBtn('🖨 พิมพ์', 'primary', 'print');
      } else if (col.status === 'done') {
        addBtn('🖨 พิมพ์ซ้ำ', 'primary', 'reprint');
        addBtn('↻ รีเซ็ต', '', 'reset');
      } else if (col.status === 'partial') {
        addBtn(`🔁 ลองอีกครั้ง ${col.failedIds.length}`, 'primary', 'retry');
        addBtn('↩ คืน', '', 'return');
      } else if (col.status === 'error') {
        if (col.retryCount < 3) addBtn('🔁 ลองอีกครั้ง', 'primary', 'retry');
        addBtn('↩ คืน', '', 'return');
      } else if (col.status === 'printing') {
        const spin = document.createElement('span');
        spin.className = 'qf-plan-zone-btn';
        spin.style.cursor = 'default';
        spin.textContent = '⏳';
        actions.appendChild(spin);
      }

      if (actions.children.length > 0) header.appendChild(actions);

      zone.appendChild(header);

      if (col.errorMsg) {
        const errDiv = document.createElement('div');
        errDiv.style.cssText = 'font-size:10px;color:#991b1b;margin:0 14px 4px;';
        errDiv.textContent = col.errorMsg;
        zone.appendChild(errDiv);
      }

      // Cards scroll area
      const scroll = document.createElement('div');
      scroll.className = 'qf-plan-zone-scroll qf-plan-cards';
      scroll.id = `qf-plan-cards-${cid}`;
      scroll.dataset.zone = cid;
      scroll.dataset.locked = locked ? '1' : '0';
      zone.appendChild(scroll);

      return zone;
    };

    // ---- renderAddZone: inline worker/team add form ----
    const renderAddZone = (s, rerender, editingWorker = null, editingTeam = null) => {
      const zone = document.createElement('div');
      zone.className = 'qf-plan-zone qf-zone-add';

      const title = document.createElement('div');
      title.className = 'qf-zone-add-title';
      title.style.cssText = 'font-weight:700;font-size:12px;color:#555;padding:10px 10px 6px;';
      title.textContent = '+ เพิ่มคนหรือทีม';
      zone.appendChild(title);

      // Tabs
      const tabs = document.createElement('div');
      tabs.className = 'qf-zone-add-tabs';
      let activeTab = editingTeam ? 'team' : 'worker';

      const tabWorker = document.createElement('button');
      tabWorker.className = 'qf-plan-zone-btn' + (activeTab === 'worker' ? ' primary' : '');
      tabWorker.textContent = 'คนเดียว';
      const tabTeam = document.createElement('button');
      tabTeam.className = 'qf-plan-zone-btn' + (activeTab === 'team' ? ' primary' : '');
      tabTeam.textContent = 'ทีม';
      tabs.appendChild(tabWorker);
      tabs.appendChild(tabTeam);
      zone.appendChild(tabs);

      // ---- Existing team list ----
      const teamListDiv = document.createElement('div');
      teamListDiv.style.cssText = 'padding:0 10px 4px;';
      const renderTeamList = () => {
        teamListDiv.innerHTML = '';
        if (!state.teams.length) return;
        const hdr = document.createElement('div');
        hdr.style.cssText = 'font-size:10px;font-weight:600;color:#888;margin-bottom:4px;';
        hdr.textContent = 'ทีมที่มีอยู่';
        teamListDiv.appendChild(hdr);
        for (const tm of state.teams) {
          const row = document.createElement('div');
          row.style.cssText = 'display:flex;align-items:center;gap:4px;font-size:11px;margin-bottom:3px;';
          const nameSpan = document.createElement('span');
          nameSpan.style.flex = '1';
          nameSpan.textContent = '👥 ' + tm.name;
          const eBtn = document.createElement('button');
          eBtn.className = 'qf-plan-zone-btn';
          eBtn.style.padding = '2px 6px';
          eBtn.textContent = '✎';
          eBtn.onclick = () => {
            zone.remove();
            win.querySelector('.qf-plan-window-body').appendChild(
              renderAddZone(s, rerender, null, tm)
            );
          };
          const dBtn = document.createElement('button');
          dBtn.className = 'qf-plan-zone-btn';
          dBtn.style.padding = '2px 6px';
          dBtn.textContent = '🗑';
          dBtn.onclick = async () => {
            const ok = await confirmInline(`ลบทีม "${tm.name}"?`, 'ลบ', true);
            if (!ok) return;
            // Return team column ids to unassigned
            let next = s;
            if (s.columns[tm.id]) {
              const returnIds = s.columns[tm.id].fulfillUnitIds;
              const { [tm.id]: _dropped, ...restCols } = s.columns;
              next = { ...s, columns: restCols, unassignedIds: [...s.unassignedIds, ...returnIds] };
            }
            deleteTeam(tm.id);
            debouncedSavePlan(next);
            rerender(next);
          };
          row.appendChild(nameSpan);
          row.appendChild(eBtn);
          row.appendChild(dBtn);
          teamListDiv.appendChild(row);
        }
      };
      renderTeamList();

      // ---- Form area ----
      const form = document.createElement('div');
      form.className = 'qf-zone-add-form';

      const nameInput = document.createElement('input');
      nameInput.name = 'qf-new-name';
      nameInput.maxLength = 30;
      nameInput.placeholder = 'ชื่อ...';
      nameInput.style.cssText = 'width:100%;box-sizing:border-box;padding:5px 8px;border:1px solid #d1d5db;border-radius:6px;font-size:12px;';
      if (editingWorker) nameInput.value = editingWorker.name;
      if (editingTeam) nameInput.value = editingTeam.name;
      form.appendChild(nameInput);

      const selectedIcon = editingWorker ? editingWorker.icon : WORKER_ICONS[0];
      const iconPicker = document.createElement('div'); // kept for switchTab compat, not appended

      // Team members checklist
      const memberList = document.createElement('div');
      memberList.className = 'qf-zone-add-team-members';
      const existingMemberIds = editingTeam ? [...editingTeam.memberWorkerIds] : [];
      for (const w of state.workers) {
        const lbl = document.createElement('label');
        lbl.style.cssText = 'display:flex;align-items:center;gap:5px;font-size:11px;cursor:pointer;';
        const chk = document.createElement('input');
        chk.type = 'checkbox';
        chk.value = w.id;
        chk.checked = existingMemberIds.includes(w.id);
        lbl.appendChild(chk);
        lbl.appendChild(document.createTextNode(w.name));
        memberList.appendChild(lbl);
      }

      const switchTab = (tab) => {
        activeTab = tab;
        tabWorker.className = 'qf-plan-zone-btn' + (tab === 'worker' ? ' primary' : '');
        tabTeam.className = 'qf-plan-zone-btn' + (tab === 'team' ? ' primary' : '');
        form.innerHTML = '';
        form.appendChild(nameInput);
        if (tab === 'worker') {
          form.appendChild(iconPicker);
          teamListDiv.style.display = 'none';
        } else {
          form.appendChild(memberList);
          teamListDiv.style.display = '';
        }
        form.appendChild(actions);
      };

      tabWorker.onclick = () => switchTab('worker');
      tabTeam.onclick = () => switchTab('team');

      // Actions row
      const actions = document.createElement('div');
      actions.className = 'qf-zone-add-actions';
      actions.style.cssText = 'display:flex;gap:6px;';

      const cancelBtn = document.createElement('button');
      cancelBtn.className = 'qf-plan-zone-btn';
      cancelBtn.textContent = 'ยกเลิก';
      cancelBtn.onclick = () => {
        nameInput.value = '';
        switchTab('worker');
      };

      const saveBtn = document.createElement('button');
      saveBtn.className = 'qf-plan-zone-btn primary';
      saveBtn.textContent = 'บันทึก';
      saveBtn.onclick = () => {
        const name = nameInput.value.trim();
        if (!name) { showToast('กรุณากรอกชื่อ', 1800); return; }

        if (activeTab === 'worker') {
          if (editingWorker) {
            state.workers = state.workers.map(w => w.id === editingWorker.id ? { ...w, name, icon: selectedIcon } : w);
            // Update column name/icon in session
            const cid = editingWorker.id;
            if (s.columns[cid]) {
              const updatedCol = { ...s.columns[cid], workerName: name, workerIcon: selectedIcon };
              s = { ...s, columns: { ...s.columns, [cid]: updatedCol } };
            }
          } else {
            const id = Math.random().toString(36).slice(2, 10);
            const worker = { id, name, icon: selectedIcon };
            state.workers = [...state.workers, worker];
            s = {
              ...s,
              columns: {
                ...s.columns,
                [id]: {
                  kind: 'worker', columnId: id, workerId: id,
                  workerName: name, workerIcon: selectedIcon,
                  fulfillUnitIds: [], printedIds: [], failedIds: [],
                  status: 'pending', lastPrintAt: null, errorMsg: null, retryCount: 0,
                },
              },
            };
          }
          saveWorkers();
        } else {
          // Team tab
          const checkedIds = [...memberList.querySelectorAll('input[type=checkbox]:checked')].map(c => c.value);
          if (!checkedIds.length) { showToast('เลือกสมาชิกอย่างน้อย 1 คน', 1800); return; }
          if (editingTeam) {
            updateTeam(editingTeam.id, { name, memberWorkerIds: checkedIds });
            if (s.columns[editingTeam.id]) {
              s = { ...s, columns: { ...s.columns, [editingTeam.id]: { ...s.columns[editingTeam.id], teamName: name, memberWorkerIds: checkedIds } } };
            }
          } else {
            const team = createTeam({ name, memberWorkerIds: checkedIds });
            s = {
              ...s,
              columns: {
                ...s.columns,
                [team.id]: {
                  kind: 'team', columnId: team.id, teamId: team.id,
                  teamName: name, memberWorkerIds: checkedIds,
                  fulfillUnitIds: [], printedIds: [], failedIds: [],
                  status: 'pending', lastPrintAt: null, errorMsg: null, retryCount: 0,
                },
              },
            };
          }
        }
        debouncedSavePlan(s);
        nameInput.value = '';
        switchTab('worker');
        rerender(s);
      };

      actions.appendChild(cancelBtn);
      actions.appendChild(saveBtn);

      // Initial tab setup
      form.appendChild(nameInput);
      if (activeTab === 'worker') {
        form.appendChild(iconPicker);
        teamListDiv.style.display = 'none';
      } else {
        form.appendChild(memberList);
        teamListDiv.style.display = '';
      }
      form.appendChild(actions);

      zone.appendChild(teamListDiv);
      zone.appendChild(form);
      return zone;
    };

    // ---- Main render ----
    const render = (s) => {
      session = s;
      window.__qfPlanningSession = () => session;
      const total = planTotalIds(session);
      const printed = planPrintedIds(session).length;
      const anyPrinting = Object.values(session.columns).some(c => c.status === 'printing');

      // Compute all cards across all zones (for split detection)
      const allIds = [...session.unassignedIds, ...Object.values(session.columns).flatMap(c => c.fulfillUnitIds)];
      const allSessionCards = buildPlanCards(allIds);

      // Update title
      const titleEl = win.querySelector('.qf-plan-window-title');
      if (titleEl) titleEl.textContent = `🎨 แผนงานแพ็ค — ${printed}/${total} ใบ`;

      // Update print-all button state
      const printAllBtn = win.querySelector('.qf-plan-printall-btn');
      if (printAllBtn) printAllBtn.disabled = anyPrinting;

      // Rebuild body content
      const body = win.querySelector('.qf-plan-window-body');
      if (!body) return;
      body.innerHTML = '';

      // ── Unassigned zone with order-type tabs ──────────────────────────────
      const unassignedZone = document.createElement('div');
      unassignedZone.className = 'qf-plan-zone qf-zone-unassigned';

      // Count per type
      const typeCounts = { single: 0, multi: 0, weird: 0 };
      for (const id of session.unassignedIds) typeCounts[getOrderType(id)]++;

      // Header title
      const unassignedHeader = document.createElement('div');
      unassignedHeader.className = 'qf-plan-zone-header';
      const unassignedTitle = document.createElement('span');
      unassignedTitle.className = 'qf-plan-zone-title';
      unassignedTitle.textContent = `📦 ยังไม่มอบหมาย (${session.unassignedIds.length} ใบ)`;
      unassignedHeader.appendChild(unassignedTitle);
      unassignedZone.appendChild(unassignedHeader);

      // Order-type tab bar
      const tabDefs = [
        { key: 'all',    label: 'ทั้งหมด',    count: session.unassignedIds.length, typeClass: '' },
        { key: 'single', label: '1 ชิ้น',     count: typeCounts.single, typeClass: 'qf-type-single' },
        { key: 'multi',  label: 'หลายชิ้น',   count: typeCounts.multi,  typeClass: 'qf-type-multi'  },
        { key: 'weird',  label: 'ออเดอร์แปลก', count: typeCounts.weird,  typeClass: 'qf-type-weird'  },
      ];
      const tabBar = document.createElement('div');
      tabBar.className = 'qf-plan-order-tabs';
      const tabBtns = [];
      for (const td of tabDefs) {
        const btn = document.createElement('button');
        btn.className = `qf-plan-order-tab${td.typeClass ? ' ' + td.typeClass : ''}${activeOrderTab === td.key ? ' active' : ''}`;
        btn.dataset.tabKey = td.key;
        btn.disabled = td.count === 0 && td.key !== 'all';
        btn.innerHTML = `${escapeHtml(td.label)}<span class="qf-plan-order-tab-count">${td.count}</span>`;
        tabBtns.push(btn);
        tabBar.appendChild(btn);
      }
      unassignedZone.appendChild(tabBar);

      // Cards scroll area
      const unassignedScroll = document.createElement('div');
      unassignedScroll.className = 'qf-plan-zone-scroll qf-plan-cards';
      unassignedScroll.id = 'qf-plan-cards-unassigned';
      unassignedScroll.dataset.zone = 'unassigned';
      unassignedScroll.dataset.locked = '0';
      unassignedZone.appendChild(unassignedScroll);
      body.appendChild(unassignedZone);

      // Helper: filter unassigned IDs by active tab and re-render scroll
      const applyUnassignedTab = (tabKey) => {
        activeOrderTab = tabKey;
        tabBtns.forEach(b => b.classList.toggle('active', b.dataset.tabKey === tabKey));
        const filtered = tabKey === 'all'
          ? session.unassignedIds
          : session.unassignedIds.filter(id => getOrderType(id) === tabKey);
        renderPlanCardsInto(unassignedScroll, filtered, allSessionCards);
        wireDragListeners(unassignedScroll);
      };
      tabBtns.forEach(btn => {
        btn.addEventListener('click', () => applyUnassignedTab(btn.dataset.tabKey));
      });
      applyUnassignedTab(activeOrderTab);

      // Worker/Team columns
      for (const [cid, col] of Object.entries(session.columns)) {
        const colZone = renderColumnZone(cid, col, session);
        body.appendChild(colZone);
        const scroll = colZone.querySelector('.qf-plan-zone-scroll');
        if (scroll) renderPlanCardsInto(scroll, col.fulfillUnitIds, allSessionCards);
      }

      // Add zone
      body.appendChild(renderAddZone(s, render));

      // Wire events
      attachPlanEvents(win, s, render);
    };

    // ---- Wire titlebar drag ----
    const titlebar = document.createElement('div');
    titlebar.className = 'qf-plan-window-titlebar';

    const titleSpan = document.createElement('span');
    titleSpan.className = 'qf-plan-window-title';
    titleSpan.textContent = '🎨 แผนงานแพ็ค';
    titlebar.appendChild(titleSpan);

    const actions = document.createElement('div');
    actions.className = 'qf-plan-window-actions';

    const autoBtn = document.createElement('button');
    autoBtn.className = 'qf-plan-minimize-btn qf-plan-auto-btn';
    autoBtn.title = 'แบ่งงานอัตโนมัติ — เฉลี่ยหรือแยกตาม SKU';
    autoBtn.textContent = '💡 แบ่งอัตโนมัติ';
    actions.appendChild(autoBtn);

    const printAllBtn = document.createElement('button');
    printAllBtn.className = 'qf-plan-minimize-btn qf-plan-printall-btn';
    printAllBtn.title = 'พิมพ์ทั้งหมดทุกคอลัมน์';
    printAllBtn.textContent = '🖨';
    actions.appendChild(printAllBtn);

    const minimizeBtn = document.createElement('button');
    minimizeBtn.className = 'qf-plan-minimize-btn';
    minimizeBtn.title = 'ย่อ';
    minimizeBtn.textContent = '–';
    actions.appendChild(minimizeBtn);

    const closeBtn = document.createElement('button');
    closeBtn.className = 'qf-plan-close-btn';
    closeBtn.title = 'ปิด';
    closeBtn.textContent = '×';
    actions.appendChild(closeBtn);

    titlebar.appendChild(actions);
    win.appendChild(titlebar);

    const body = document.createElement('div');
    body.className = 'qf-plan-window-body';
    win.appendChild(body);

    // ---- Draggable titlebar ----
    let _dragOffX = 0, _dragOffY = 0, _isDragging = false;

    titlebar.addEventListener('mousedown', (e) => {
      if (e.target.closest('.qf-plan-window-actions')) return;
      _isDragging = true;
      const rect = win.getBoundingClientRect();
      _dragOffX = e.clientX - rect.left;
      _dragOffY = e.clientY - rect.top;
      e.preventDefault();
    });

    const onMouseMove = (e) => {
      if (!_isDragging) return;
      const vw2 = window.innerWidth, vh2 = window.innerHeight;
      const winW = win.offsetWidth, winH = win.offsetHeight;
      const nx = Math.min(Math.max(0, e.clientX - _dragOffX), vw2 - winW);
      const ny = Math.min(Math.max(0, e.clientY - _dragOffY), vh2 - winH);
      win.style.left = nx + 'px';
      win.style.top = ny + 'px';
    };

    const onMouseUp = () => {
      if (!_isDragging) return;
      _isDragging = false;
      savePanelPos({ x: parseInt(win.style.left) || 0, y: parseInt(win.style.top) || 0 });
    };

    document.addEventListener('mousemove', onMouseMove);
    document.addEventListener('mouseup', onMouseUp);

    // ---- Minimize button ----
    minimizeBtn.addEventListener('click', () => {
      win.style.display = 'none';
      renderPlanBubble(session, () => {
        win.style.display = '';
        document.querySelectorAll('.qf-plan-bubble').forEach(b => b.remove());
      }, () => {
        // Close from bubble
        const anyPrinting = Object.values(session.columns).some(c => c.status === 'printing');
        if (anyPrinting) { showToast('กำลังพิมพ์อยู่ — รอให้เสร็จก่อนปิด', 2500); return; }
        win.remove();
        document.querySelectorAll('.qf-plan-bubble').forEach(b => b.remove());
        window.removeEventListener('beforeunload', beforeUnloadHandler);
        renderRecoveryBanner();
      });
    });

    // ---- Close button ----
    closeBtn.addEventListener('click', () => {
      const anyPrinting = Object.values(session.columns).some(c => c.status === 'printing');
      if (anyPrinting) { showToast('รอให้เสร็จก่อนปิด', 2500); return; }
      win.remove();
      document.querySelectorAll('.qf-plan-bubble').forEach(b => b.remove());
      window.removeEventListener('beforeunload', beforeUnloadHandler);
      document.removeEventListener('mousemove', onMouseMove);
      document.removeEventListener('mouseup', onMouseUp);
      renderRecoveryBanner();
    });

    // ---- Auto-split button ----
    autoBtn.addEventListener('click', async () => {
      const result = await showPlanAutoSplitModal(session);
      if (!result || !result.workerIds.length) return;
      const next = applyAutoSplit(session, result.mode, result.workerIds);
      debouncedSavePlan(next);
      render(next);
    });

    // ---- Print-all button ----
    printAllBtn.addEventListener('click', async () => {
      if (isShopee()) { showToast('การพิมพ์จากแผนยังไม่รองรับบน Shopee', 3000); return; }
      const s = session;
      const nonDone = Object.entries(s.columns).filter(([, c]) => c.status !== 'done' && c.status !== 'printing' && c.fulfillUnitIds.length > 0);
      if (!nonDone.length) { showToast('ทุกคอลัมน์พิมพ์แล้วหรือว่างอยู่', 2000); return; }

      // Gather all pending IDs across columns + detect multi-SKU
      const allPrintIds = nonDone.flatMap(([, c]) => c.fulfillUnitIds.filter(id => !c.printedIds.includes(id)));
      if (!allPrintIds.length) { showToast('ไม่มีฉลากที่ยังไม่ได้พิมพ์', 2000); return; }

      const skuBuckets = new Set();
      let hasComboRecord = false;
      for (const id of allPrintIds) {
        const rec = state.records.get(id);
        if (!rec?.skuList?.length) continue;
        if (rec.skuList.length > 1) hasComboRecord = true;
        for (const sk of rec.skuList) skuBuckets.add(`${sk.productId}:${sk.skuId}`);
      }
      const multiSku = hasComboRecord || skuBuckets.size > 1;

      // §UX-fix: always show the full print-all modal so the user can see
      // divider/picking-list toggles and mode choice. Previously, small
      // single-SKU jobs bypassed the modal entirely — users with small
      // batches had no way to opt into dividers or picking lists.
      const plan = await showPrintAllPlanModal({
        total: allPrintIds.length,
        nonDone,
        hasMultiSku: multiSku,
        defaultPickingList: loadPickingListPref(),
      });
      if (!plan) return;

      await printAllPlanColumns(s, nonDone, plan, render);
    });

    // ---- attachPlanEvents: column actions + drag-drop ----
    const attachPlanEvents = (root, s, rerender) => {
      // Column action buttons
      root.querySelectorAll('[data-action]').forEach(btn => {
        const action = btn.dataset.action;
        if (!action) return;
        btn.addEventListener('click', async () => {
          const cid = btn.dataset.wid;
          const col = s.columns[cid];
          const colDisplayName = col ? (col.teamName || col.workerName || cid) : cid;

          if (action === 'print' || action === 'reprint') {
            if (!col) return;
            const ids = action === 'reprint' ? col.fulfillUnitIds : col.fulfillUnitIds.filter(id => !col.printedIds.includes(id));
            if (action === 'print') {
              // §Team-aware confirm: for team columns, show member count so the
              // manager knows how many packers share the workload before printing.
              const memberCount = col.kind === 'team' && Array.isArray(col.memberWorkerIds)
                ? col.memberWorkerIds.length : 0;
              const subject = col.kind === 'team' && memberCount > 0
                ? `ทีม ${colDisplayName} (${memberCount} คน)`
                : colDisplayName;
              const ok = await confirmInline(`พิมพ์ของ ${subject} ${ids.length} ใบ?`, 'พิมพ์');
              if (!ok) return;
            }
            printPlanColumn(s, cid, ids, ns => { rerender(ns); s = ns; }, { reprint: action === 'reprint' });
          } else if (action === 'retry') {
            if (!col) return;
            const ids = col.failedIds.length > 0 ? col.failedIds : col.fulfillUnitIds;
            printPlanColumn(s, cid, ids, ns => { rerender(ns); s = ns; });
          } else if (action === 'return') {
            if (!col) return;
            const returnIds = col.failedIds.length > 0 ? col.failedIds : col.fulfillUnitIds;
            const updatedCol = { ...col, fulfillUnitIds: col.fulfillUnitIds.filter(id => !returnIds.includes(id)), failedIds: [], status: 'pending', errorMsg: null };
            const next = { ...s, columns: { ...s.columns, [cid]: updatedCol }, unassignedIds: [...s.unassignedIds, ...returnIds] };
            debouncedSavePlan(next);
            rerender(next);
          } else if (action === 'reset') {
            if (!col) return;
            // §UX-fix: add confirmation — reset wipes printedIds/failedIds
            // for the session and was previously irreversible with no prompt.
            const okReset = await confirmInline(
              `รีเซ็ต ${colDisplayName}? ข้อมูลการพิมพ์ในรอบนี้จะถูกล้าง`,
              'รีเซ็ต',
              true
            );
            if (!okReset) return;
            const updatedCol = { ...col, status: 'pending', printedIds: [], failedIds: [], errorMsg: null, retryCount: 0 };
            const next = { ...s, columns: { ...s.columns, [cid]: updatedCol } };
            debouncedSavePlan(next);
            rerender(next);
          } else if (action === 'removecol') {
            if (!col) return;
            const ok = await confirmInline(`ลบ ${colDisplayName}? (ids จะคืนยังไม่มอบหมาย)`, 'ลบ', true);
            if (!ok) return;
            const returnIds = col.fulfillUnitIds;
            const { [cid]: _dropped, ...restCols } = s.columns;
            const next = { ...s, columns: restCols, unassignedIds: [...s.unassignedIds, ...returnIds] };
            // Also delete worker/team from state
            if (col.kind === 'worker') {
              state.workers = state.workers.filter(w => w.id !== cid);
              saveWorkers();
              removeWorkerFromTeams(cid);
            } else if (col.kind === 'team') {
              deleteTeam(cid);
            }
            debouncedSavePlan(next);
            rerender(next);
          } else if (action === 'editcol') {
            if (!col) return;
            // Remove old add-zone, rebuild with editing state
            const addZone = root.querySelector('.qf-zone-add');
            if (addZone) addZone.remove();
            const editingWorker = col.kind === 'worker' ? state.workers.find(w => w.id === cid) : null;
            const editingTeam = col.kind === 'team' ? getTeam(cid) : null;
            root.querySelector('.qf-plan-window-body').appendChild(
              renderAddZone(s, rerender, editingWorker, editingTeam)
            );
          }
        });
      });

      // Drag-drop on plan cards
      wireDragListeners(root);

      root.querySelectorAll('.qf-plan-cards').forEach(dropZone => {
        dropZone.addEventListener('dragover', e => {
          const locked = dropZone.dataset.locked === '1';
          if (locked && dropZone.dataset.zone !== _dragSourceZone) return;
          e.preventDefault();
          dropZone.classList.add('qf-drag-over');
        });
        dropZone.addEventListener('dragleave', () => dropZone.classList.remove('qf-drag-over'));
        dropZone.addEventListener('drop', async e => {
          e.preventDefault();
          dropZone.classList.remove('qf-drag-over');
          const targetZone = dropZone.dataset.zone;
          if (!_dragCardKey || targetZone === _dragSourceZone) return;

          const srcZone = _dragSourceZone;
          const cardKey = _dragCardKey;

          const srcIds = srcZone === 'unassigned' ? s.unassignedIds : (s.columns[srcZone]?.fulfillUnitIds || []);
          const cardObj = buildPlanCards(srcIds).find(c => c.key === cardKey);
          if (!cardObj) return;

          const totalAvail = cardObj.count;
          let qty = totalAvail;

          if (qty > 1 && targetZone !== 'unassigned') {
            const targetCol = s.columns[targetZone];
            const zoneName = targetCol ? (targetCol.teamName || targetCol.workerName || targetZone) : targetZone;
            const chosen = await showPlanChunkPopup(zoneName, totalAvail);
            if (chosen === null) return;
            qty = chosen;
          }

          const idsToMove = cardObj.ids.slice(0, qty);
          let next = { ...s, unassignedIds: [...s.unassignedIds], columns: { ...s.columns } };

          if (srcZone === 'unassigned') {
            next.unassignedIds = s.unassignedIds.filter(id => !idsToMove.includes(id));
          } else {
            const srcCol = s.columns[srcZone];
            next.columns[srcZone] = { ...srcCol, fulfillUnitIds: srcCol.fulfillUnitIds.filter(id => !idsToMove.includes(id)) };
          }

          if (targetZone === 'unassigned') {
            next.unassignedIds = [...next.unassignedIds, ...idsToMove];
          } else {
            const existingCol = next.columns[targetZone];
            if (!existingCol) return;
            next.columns[targetZone] = { ...existingCol, fulfillUnitIds: [...existingCol.fulfillUnitIds, ...idsToMove] };
          }

          debouncedSavePlan(next);
          rerender(next);
        });
      });
    };

    // beforeunload guard
    const beforeUnloadHandler = (e) => {
      const anyPrinting = Object.values(session.columns).some(c => c.status === 'printing');
      if (anyPrinting) { e.preventDefault(); e.returnValue = 'กำลังพิมพ์อยู่ — แน่ใจหรือว่าจะออก?'; return e.returnValue; }
    };
    window.addEventListener('beforeunload', beforeUnloadHandler);

    // Refresh hook: called by renderAll() after each scan so the plan panel stays in sync.
    window.__qfPlanRefresh = () => {
      if (!document.body.contains(win)) { window.__qfPlanRefresh = null; return; }
      const allTracked = new Set([
        ...session.unassignedIds,
        ...Object.values(session.columns).flatMap(c => c.fulfillUnitIds),
      ]);
      const filterPrinted = state.labelStatusFilter === 'all';
      const newIds = [...state.records.keys()].filter(id => {
        if (allTracked.has(id)) return false;
        if (filterPrinted && state.records.get(id)?.labelStatus === LABEL_STATUS_PRINTED) return false;
        return true;
      });
      const validId = id => state.records.has(id);
      const merged = {
        ...session,
        unassignedIds: [...session.unassignedIds.filter(validId), ...newIds],
        columns: Object.fromEntries(
          Object.entries(session.columns).map(([cid, col]) => [
            cid, { ...col, fulfillUnitIds: col.fulfillUnitIds.filter(validId) },
          ])
        ),
      };
      render(merged);
    };

    render(session);
  }

  // ---- renderPlanBubble: minimized pill shown when panel is hidden ----
  function renderPlanBubble(session, onExpand, onClose) {
    document.querySelectorAll('.qf-plan-bubble').forEach(b => b.remove());
    const total = planTotalIds(session);
    const printed = planPrintedIds(session).length;
    const anyPrinting = Object.values(session.columns).some(c => c.status === 'printing');

    const bubble = document.createElement('div');
    bubble.className = 'qf-plan-bubble';

    const txt = document.createElement('span');
    txt.className = 'qf-plan-bubble-text';
    txt.textContent = anyPrinting ? 'กำลังพิมพ์...' : `🎨 แผน ${printed}/${total} ใบ`;
    bubble.appendChild(txt);

    const closeX = document.createElement('button');
    closeX.className = 'qf-plan-bubble-close';
    closeX.title = 'ปิดแผน';
    closeX.textContent = '×';
    closeX.onclick = (e) => { e.stopPropagation(); onClose(); bubble.remove(); };
    bubble.appendChild(closeX);

    bubble.addEventListener('click', (e) => {
      if (e.target === closeX) return;
      onExpand();
      bubble.remove();
    });

    document.body.appendChild(bubble);
  }

  function renderRecoveryBanner() {
    document.getElementById('qf-plan-recovery')?.remove();
    const session = loadPlanningSession();
    if (!session) return;
    const total = planTotalIds(session);
    const colCount = Object.keys(session.columns).length;
    if (!total) return;
    const banner = document.createElement('div');
    banner.id = 'qf-plan-recovery';
    banner.className = 'qf-plan-recovery';
    banner.innerHTML = `
      <span class="qf-plan-recovery-msg">🎨 มีแผนงานค้างอยู่ (${total} ใบ, ${colCount} คน/ทีม)</span>
      <button class="qf-plan-recovery-btn qf-plan-recovery-open">เปิด</button>
      <button class="qf-plan-recovery-btn qf-plan-recovery-discard">ละทิ้ง</button>
    `;
    const body = document.getElementById('qf-body');
    if (!body) return;
    body.insertBefore(banner, body.firstChild);
    banner.querySelector('.qf-plan-recovery-open').onclick = () => {
      banner.remove();
      openPlanningPanel(session);
    };
    banner.querySelector('.qf-plan-recovery-discard').onclick = () => {
      deletePlanningSession();
      banner.remove();
    };
  }

  // ==================== END PLANNING PANEL ====================

  window.__qfPlanningSession = () => null;

  // ==================== PDF TEMPLATE BUILDER (Phase 2) ====================
  //
  // Provides a WYSIWYG editor that layers CUSTOMIZABLE elements (logo, text,
  // variables, placeholders) on top of the carrier's LOCKED zones defined by
  // J_AND_T_LAYOUT. Users save up to MAX_PDF_TEMPLATES templates; one is
  // marked "active" and automatically applied in the print pipeline AFTER
  // overlayAliasOnPdf — so the alias watermark remains visible unless the
  // template deliberately covers it.
  //
  // Data is persisted to localStorage. LOCKED regions are enforced in the
  // editor (drag/drop/resize cannot intersect) so the barcode/QR/route code
  // a carrier needs for scanning remain untouched. Phase 2 is grayscale only.

  const PDF_TEMPLATES_KEY = 'qf_pdf_templates_v1';
  const PDF_ACTIVE_TEMPLATE_KEY = 'qf_pdf_active_template_v1';
  const PDF_TEMPLATES_SEEDED_KEY = 'qf_pdf_templates_seeded_v1';
  const MAX_PDF_TEMPLATES = 3;
  const PDF_TEMPLATE_GRID_PT = 4; // snap grid in PDF points
  const PDF_TEMPLATE_SCALE = 2;   // editor canvas: 2x page points

  // Baseline carrier layout — from .claude/samples/jnt_layout.json.
  // Coordinates are PDF points with bottom-left origin. Editor converts
  // to top-left origin for display.
  const J_AND_T_LAYOUT = {
    carrier: 'jnt',
    pageSize: { w: 298, h: 420 },
    // LOCKED regions — drop/resize into these is rejected.
    locked: [
      { id: 'orderId',        label: 'Order ID',        x: 42.86,  y: 162.83, w: 150, h: 22 },
      { id: 'trackingNumber', label: 'Tracking',        x: 105.86, y: 341.5,  w: 160, h: 24 },
      { id: 'sortCode',       label: 'Sort Code',       x: 188.18, y: 402,    w: 90,  h: 28 },
      { id: 'routeCode',      label: 'Route',           x: 164.38, y: 300,    w: 125, h: 40 },
      { id: 'subZoneCode',    label: 'Sub-Zone',        x: 204.54, y: 282,    w: 80,  h: 22 },
      { id: 'serviceType',    label: 'Service',         x: 215.62, y: 178,    w: 75,  h: 20 },
      { id: 'codLabel',       label: 'COD',             x: 53.28,  y: 174,    w: 130, h: 26 },
      { id: 'barcodeLeft',    label: 'Side barcode L',  x: 4,      y: 180,    w: 24,  h: 220 },
      { id: 'barcodeRight',   label: 'Side barcode R',  x: 270,    y: 78,     w: 24,  h: 322 },
    ],
    // SHRINKABLE regions — user can toggle via zones flags.
    shrinkable: [
      { id: 'skuTable',     label: 'SKU Table',     x: 0, y: 78.94, w: 298, h: 51.57 },
      { id: 'addressBlock', label: 'Address Block', x: 0, y: 147,   w: 298, h: 126 },
    ],
    // Header zone (above LOCKED area) is where users typically drop their
    // logo + shop branding. y values are for guidance; editor enforces LOCKED.
    headerZone: { x: 0, y: 0, w: 298, h: 40 }, // top strip (in top-left editor coords)
    // PHASE 3 — System elements exposed as user-overridable overrides.
    // Coordinates below are the CARRIER-ORIGINAL rects (PDF points,
    // bottom-left origin) that the template renderer can MASK + REDRAW at
    // user-chosen coords. Barcode + QR are deliberately absent — they are
    // never exposed to the editor. Fields mirror the keys permitted in
    // template.overrides.
    //
    // kind: 'image' | 'text' | 'block'
    //   image → two-logo header marks (drawImage)
    //   text  → single-line system text (drawText)
    //   block → multi-line / tabular region (skuTable, addressBlock)
    systemElements: {
      tiktokLogo: {
        kind: 'image', label: 'TikTok Shop logo', maskable: true,
        // Left half of the top header band (above barcode frame).
        x: 0, y: 395, w: 92, h: 24,
        default: { x: 2, y: 395, w: 60, h: 18 },
      },
      jntLogo: {
        kind: 'image', label: 'J&T Express logo', maskable: true,
        x: 92, y: 395, w: 96, h: 24,
        default: { x: 94, y: 395, w: 80, h: 18 },
      },
      sortCode: {
        kind: 'text', label: 'Sort code (EZ)', maskable: true,
        x: 188.18, y: 402, w: 90, h: 28, size: 28,
        default: { x: 188.18, y: 402, w: 90, h: 28, size: 28 },
      },
      orderId: {
        kind: 'text', label: 'Order ID', maskable: true,
        x: 42.86, y: 162.83, w: 150, h: 22, size: 10,
        default: { x: 42.86, y: 162.83, w: 150, h: 22, size: 10 },
      },
      trackingNumber: {
        kind: 'text', label: 'Tracking #', maskable: true,
        x: 105.86, y: 341.5, w: 160, h: 24, size: 14,
        default: { x: 105.86, y: 341.5, w: 160, h: 24, size: 14 },
      },
      serviceType: {
        kind: 'text', label: 'Service type (DROP-OFF/EZ)', maskable: true,
        x: 215.62, y: 178, w: 75, h: 20, size: 12,
        default: { x: 215.62, y: 178, w: 75, h: 20, size: 12 },
      },
      subZone: {
        kind: 'text', label: 'Sub-zone (004A)', maskable: true,
        x: 204.54, y: 282, w: 80, h: 22, size: 12,
        default: { x: 204.54, y: 282, w: 80, h: 22, size: 12 },
      },
      codLabel: {
        kind: 'text', label: 'COD label', maskable: true,
        x: 53.28, y: 174, w: 130, h: 26, size: 14,
        default: { x: 53.28, y: 174, w: 130, h: 26, size: 14 },
      },
      skuTable: {
        kind: 'block', label: 'SKU table', maskable: true,
        x: 0, y: 78.94, w: 298, h: 51.57,
        default: { x: 4, y: 78.94, w: 290, h: 52, fontSize: 7,
          columns: ['name', 'sku', 'seller', 'qty'] },
      },
      addressBlock: {
        kind: 'block', label: 'Address block', maskable: true,
        x: 0, y: 147, w: 298, h: 126,
        default: { x: 4, y: 150, w: 290, h: 120, fontSize: 8 },
      },
    },
  };

  // PHASE 3 SAFETY — keys that are NEVER user-overridable (barcode + QR).
  // Any attempt to mask or redraw these is rejected by renderSystemOverride().
  const NEVER_OVERRIDE_KEYS = new Set([
    'barcodeMain', 'barcodeLeft', 'barcodeRight', 'qrCode', 'qr',
  ]);

  // ---- Default (seeded) templates ----
  //
  // Positions are chosen to sit in free bands of the J&T layout (not overlapping
  // any LOCKED zone). Safe bands:
  //   * Band A — y: 8-45,  x: 40-180   (above tracking/sortCode)
  //   * Band C — y: 355-400, x: 30-260 (below all barcodes & SKU table)
  // All elements are validated against intersectsLocked() at seed time; if any
  // conflict, seeding aborts gracefully and logs a warning.
  const DEFAULT_PDF_TEMPLATES = [
    {
      id: 'default_minimal',
      name: 'ขั้นต่ำ (Minimal)',
      isDefault: true,
      presetId: 'original',
      brand: { logoDataUrl: null, shopName: '', tagline: '' },
      elements: [
        { id: 'def_min_alias',   type: 'variable', variable: 'alias',
          x: 60, y: 372, w: 180, h: 22, size: 18, align: 'center', bold: true, text: '' },
        { id: 'def_min_orderid', type: 'variable', variable: 'orderId',
          x: 40, y: 10, w: 140, h: 12, size: 8, align: 'left', bold: false, text: '' },
      ],
      zones: { shrinkHeaderLogos: false, hideSkuTable: false, maskPhone: false, maskAddress: false },
      overrides: {},
    },
    {
      id: 'default_branded',
      name: 'แบรนด์ร้าน (Branded)',
      isDefault: true,
      presetId: 'branded',
      brand: { logoDataUrl: null, shopName: 'ชื่อร้านของคุณ', tagline: 'ขอบคุณที่อุดหนุน' },
      elements: [
        { id: 'def_br_logo', type: 'image', variable: '', text: '',
          x: 40, y: 10, w: 32, h: 32, size: 10, align: 'left', bold: false, dataUrl: '' },
        { id: 'def_br_shopname', type: 'variable', variable: 'shopName',
          x: 78, y: 10, w: 110, h: 16, size: 12, align: 'left', bold: true, text: '' },
        { id: 'def_br_tag', type: 'text', variable: '',
          x: 78, y: 28, w: 110, h: 12, size: 7, align: 'left', bold: false,
          text: 'ขอบคุณที่อุดหนุน' },
        { id: 'def_br_alias', type: 'variable', variable: 'alias',
          x: 60, y: 372, w: 180, h: 22, size: 18, align: 'center', bold: true, text: '' },
      ],
      zones: { shrinkHeaderLogos: false, hideSkuTable: false, maskPhone: false, maskAddress: false },
      overrides: {},
    },
    {
      id: 'default_review',
      name: 'โปรโมทรีวิว (Review Promo)',
      isDefault: true,
      presetId: 'review-promo',
      brand: { logoDataUrl: null, shopName: 'ชื่อร้านของคุณ', tagline: '' },
      elements: [
        { id: 'def_rv_shopname', type: 'variable', variable: 'shopName',
          x: 50, y: 10, w: 130, h: 16, size: 12, align: 'center', bold: true, text: '' },
        { id: 'def_rv_thanks', type: 'text', variable: '',
          x: 30, y: 355, w: 170, h: 12, size: 7, align: 'left', bold: false,
          text: 'ขอบคุณที่อุดหนุน รีวิว 5 ดาว รับคูปอง 10%' },
        { id: 'def_rv_qr', type: 'qrPlaceholder', variable: '', text: '',
          x: 220, y: 358, w: 40, h: 40, size: 10, align: 'left', bold: false, dataUrl: '' },
        { id: 'def_rv_alias', type: 'variable', variable: 'alias',
          x: 60, y: 376, w: 150, h: 22, size: 18, align: 'center', bold: true, text: '' },
      ],
      zones: { shrinkHeaderLogos: false, hideSkuTable: false, maskPhone: false, maskAddress: false },
      overrides: {},
    },
  ];

  // ---- Phase 4a — Layout presets ----
  //
  // Pre-designed coherent layouts so non-tech users pick ONE option and get
  // a usable template. All coordinates below are in TOP-LEFT editor space
  // (same coords saved in template.elements / template.overrides). They were
  // validated against J_AND_T_LAYOUT.locked rects — see verifyPresets().
  //
  // LOCKED reference (top-left editor coords, pageH=420):
  //   barcodeLeft    x:[4,28]    y:[20,240]
  //   barcodeRight   x:[270,294] y:[20,342]
  //   trackingNumber x:[105.9,265.9] y:[54.5,78.5]
  //   routeCode      x:[164.4,289.4] y:[80,120]
  //   subZoneCode    x:[204.5,284.5] y:[116,138]
  //   orderId        x:[42.9,192.9]  y:[235.2,257.2]
  //   serviceType    x:[215.6,290.6] y:[222,242]
  //   codLabel       x:[53.3,183.3]  y:[220,246]
  //   sortCode       x:[188.2,278.2] y:[0,18]
  //
  // Safe bands chosen:
  //   - Bottom strip  y:[345, 418], x:[30, 268]  — alias watermark
  //   - Top strip     y:[0, 18],    x:[0, 180]   — shop branding / logos
  //   - SKU band      y:[290, 340], x:[10, 268]  — promo text / QR
  const LAYOUT_PRESETS = [
    {
      id: 'original',
      name: 'ปกติ J&T',
      description: 'เหมือนของจริง ไม่แก้อะไร + alias ด้านล่าง',
      thumbnail: '📄',
      overrides: {},
      zones: { shrinkHeaderLogos: false, hideSkuTable: false, maskPhone: false, maskAddress: false },
      brand: { logoDataUrl: null, shopName: '', tagline: '' },
      elements: [
        { id: 'alias-1', type: 'variable', variable: 'alias',
          x: 30, y: 380, w: 235, h: 28,
          size: 22, align: 'center', bold: true, text: '' },
      ],
    },
    {
      id: 'slim-header',
      name: 'ประหยัดหัว',
      description: 'โลโก้ขนส่งเล็กลง ย้ายไปมุมขวา — ใช้พื้นที่ซ้ายใส่แบรนด์',
      thumbnail: '📦',
      // TikTok+JNT logos relocated to the very top strip (y:0-14, above
      // sortCode at y:0-18). Keep them left of sortCode (x<188).
      // Note: the editor space y:0-14 region is fine — sortCode starts at x≥188.
      overrides: {
        tiktokLogo: { enabled: true, x: 40,  y: 2, w: 60, h: 10 },
        jntLogo:    { enabled: true, x: 110, y: 2, w: 70, h: 10 },
      },
      zones: { shrinkHeaderLogos: false, hideSkuTable: false, maskPhone: false, maskAddress: false },
      brand: { logoDataUrl: null, shopName: 'ชื่อร้านของคุณ', tagline: 'ขอบคุณที่อุดหนุน' },
      elements: [
        { id: 'shop-1', type: 'variable', variable: 'shopName',
          x: 30, y: 345, w: 160, h: 14,
          size: 10, align: 'left', bold: true, text: '' },
        { id: 'tag-1', type: 'variable', variable: 'tagline',
          x: 30, y: 360, w: 160, h: 12,
          size: 8, align: 'left', bold: false, text: '' },
        { id: 'alias-1', type: 'variable', variable: 'alias',
          x: 30, y: 380, w: 235, h: 28,
          size: 22, align: 'center', bold: true, text: '' },
      ],
    },
    {
      id: 'review-promo',
      name: 'รีวิว + QR',
      description: 'ซ่อน SKU + address → ใส่ข้อความโปรโมท + ช่อง QR LINE',
      thumbnail: '⭐',
      // skuTable override compresses the table, address hidden entirely.
      // Promo text + QR sit in the freed band y:290-345, left+right split
      // around barcodeRight which is x:270+.
      overrides: {
        skuTable: { enabled: true, x: 10, y: 290, w: 180, h: 50,
                    fontSize: 6, columns: ['name', 'qty'] },
        addressBlock: { enabled: false, x: 4, y: 150, w: 290, h: 120, fontSize: 8 },
      },
      zones: { shrinkHeaderLogos: false, hideSkuTable: false, maskPhone: false, maskAddress: true },
      brand: { logoDataUrl: null, shopName: 'ชื่อร้านของคุณ', tagline: '' },
      elements: [
        { id: 'promo-1', type: 'text',
          text: '✨ ขอบคุณที่อุดหนุนค่ะ',
          x: 195, y: 295, w: 70, h: 14,
          size: 9, align: 'center', bold: true, variable: '' },
        { id: 'promo-2', type: 'text',
          text: 'รีวิว 5 ดาว รับคูปอง 10%',
          x: 195, y: 312, w: 70, h: 12,
          size: 7, align: 'center', bold: false, variable: '' },
        { id: 'qr-1', type: 'qrPlaceholder',
          x: 210, y: 345, w: 55, h: 55,
          size: 10, align: 'left', bold: false, variable: '', text: '' },
        { id: 'alias-1', type: 'variable', variable: 'alias',
          x: 30, y: 380, w: 170, h: 28,
          size: 20, align: 'center', bold: true, text: '' },
      ],
    },
    {
      id: 'compact',
      name: 'ขนาดย่อ',
      description: 'ทุกอย่างเล็กลง — เผื่อพื้นที่ใส่ข้อความเพิ่ม',
      thumbnail: '🔽',
      // Shrink logos + sku table + address. Alias stays large at bottom.
      overrides: {
        tiktokLogo:  { enabled: true, x: 40,  y: 2, w: 60, h: 10 },
        jntLogo:     { enabled: true, x: 110, y: 2, w: 70, h: 10 },
        skuTable:    { enabled: true, x: 30, y: 290, w: 235, h: 40,
                        fontSize: 6, columns: ['name', 'sku', 'qty'] },
        addressBlock:{ enabled: true, x: 30, y: 150, w: 235, h: 65,
                        fontSize: 7 },
      },
      zones: { shrinkHeaderLogos: false, hideSkuTable: false, maskPhone: false, maskAddress: false },
      brand: { logoDataUrl: null, shopName: 'ร้านของฉัน', tagline: '' },
      elements: [
        { id: 'alias-1', type: 'variable', variable: 'alias',
          x: 30, y: 375, w: 235, h: 26,
          size: 22, align: 'center', bold: true, text: '' },
        { id: 'date-1', type: 'variable', variable: 'date',
          x: 30, y: 405, w: 235, h: 10,
          size: 7, align: 'center', bold: false, text: '' },
      ],
    },
    {
      id: 'branded',
      name: 'แบรนด์ใหญ่',
      description: 'โลโก้ร้านขนาดใหญ่ตรงกลาง + tagline ชัดเจน',
      thumbnail: '🏪',
      // Carrier logos shrink to top-center band (must stay left of
      // sortCode x:188 since sortCode LOCKED spans y:0-18).
      overrides: {
        tiktokLogo: { enabled: true, x: 60,  y: 2, w: 60, h: 10 },
        jntLogo:    { enabled: true, x: 125, y: 2, w: 60, h: 10 },
      },
      zones: { shrinkHeaderLogos: false, hideSkuTable: false, maskPhone: false, maskAddress: false },
      brand: { logoDataUrl: null, shopName: 'ชื่อร้านของคุณ', tagline: 'คุณภาพส่งตรงถึงบ้าน' },
      elements: [
        { id: 'shop-big', type: 'variable', variable: 'shopName',
          x: 30, y: 345, w: 180, h: 18,
          size: 14, align: 'left', bold: true, text: '' },
        { id: 'tagline', type: 'variable', variable: 'tagline',
          x: 30, y: 363, w: 180, h: 12,
          size: 8, align: 'left', bold: false, text: '' },
        { id: 'alias-1', type: 'variable', variable: 'alias',
          x: 30, y: 380, w: 235, h: 28,
          size: 22, align: 'center', bold: true, text: '' },
      ],
    },
  ];

  // Runtime LOCKED overlap check — walks every preset and every element +
  // enabled override rect, and reports any that intersect a LOCKED zone.
  // Called once at init (after intersectsLocked is defined). If a preset
  // fails, we DON'T throw — we filter it out of the visible list so the
  // editor still works and the test plan's "no console.error" check catches
  // the regression loudly.
  let _verifiedPresets = null;
  function verifyPresets() {
    if (_verifiedPresets) return _verifiedPresets;
    if (typeof intersectsLocked !== 'function') return LAYOUT_PRESETS;
    const ok = [];
    for (const preset of LAYOUT_PRESETS) {
      const issues = [];
      for (const el of preset.elements || []) {
        const hit = intersectsLocked({ x: el.x, y: el.y, w: el.w, h: el.h });
        if (hit) issues.push(`element ${el.id} overlaps LOCKED ${hit.id}`);
      }
      for (const [key, ov] of Object.entries(preset.overrides || {})) {
        if (!ov || !ov.enabled) continue;
        if (typeof ov.x !== 'number' || typeof ov.y !== 'number') continue;
        const hit = intersectsLocked({ x: ov.x, y: ov.y, w: ov.w, h: ov.h });
        if (hit) issues.push(`override ${key} overlaps LOCKED ${hit.id}`);
      }
      if (issues.length) {
        console.error('[verifyPresets] overlap with LOCKED —', preset.id, issues);
      } else {
        ok.push(preset);
      }
    }
    _verifiedPresets = ok;
    return ok;
  }

  function getVerifiedPresets() {
    return _verifiedPresets || verifyPresets();
  }

  function getPresetById(id) {
    return LAYOUT_PRESETS.find(p => p.id === id) || null;
  }

  // Clone preset into a fresh draft (safe to mutate).
  function clonePresetAsDraft(preset) {
    if (!preset) return null;
    return {
      presetId: preset.id,
      elements: JSON.parse(JSON.stringify(preset.elements || [])),
      overrides: JSON.parse(JSON.stringify(preset.overrides || {})),
      zones: { ...(preset.zones || {}) },
      brand: { ...(preset.brand || { logoDataUrl: null, shopName: '', tagline: '' }) },
    };
  }

  // Build a mini SVG thumbnail (96×120) for a preset card.
  // Colors: LOCKED=#e5e7eb gray, text=#2563eb, image=#16a34a,
  //         variable=#7c3aed, override=#f59e0b.
  function renderPresetThumbnail(preset) {
    const W = 96, H = 120;
    const pageW = J_AND_T_LAYOUT.pageSize.w;
    const pageH = J_AND_T_LAYOUT.pageSize.h;
    const sx = W / pageW;
    const sy = H / pageH;
    const rects = [];

    // LOCKED rects (grey faded)
    for (const loc of J_AND_T_LAYOUT.locked) {
      const ty = pageH - loc.y - loc.h;
      rects.push(`<rect x="${(loc.x * sx).toFixed(2)}" y="${(ty * sy).toFixed(2)}" width="${Math.max(1, loc.w * sx).toFixed(2)}" height="${Math.max(1, loc.h * sy).toFixed(2)}" fill="#e5e7eb" />`);
    }

    // Override rects (orange)
    for (const [, ov] of Object.entries(preset.overrides || {})) {
      if (!ov?.enabled) continue;
      rects.push(`<rect x="${(ov.x * sx).toFixed(2)}" y="${(ov.y * sy).toFixed(2)}" width="${Math.max(1, ov.w * sx).toFixed(2)}" height="${Math.max(1, ov.h * sy).toFixed(2)}" fill="#f59e0b" fill-opacity="0.55" stroke="#b45309" stroke-width="0.5" />`);
    }

    // Element rects
    for (const el of preset.elements || []) {
      let color = '#2563eb'; // text
      if (el.type === 'image') color = '#16a34a';
      else if (el.type === 'qrPlaceholder') color = '#374151';
      else if (el.type === 'variable') color = '#7c3aed';
      rects.push(`<rect x="${(el.x * sx).toFixed(2)}" y="${(el.y * sy).toFixed(2)}" width="${Math.max(1, el.w * sx).toFixed(2)}" height="${Math.max(1, el.h * sy).toFixed(2)}" fill="${color}" fill-opacity="0.7" />`);
    }

    return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${W} ${H}" width="${W}" height="${H}" role="img" aria-label="${escapeHtml(preset.name)} preview"><rect x="0" y="0" width="${W}" height="${H}" fill="#ffffff" stroke="#cbd5e1" stroke-width="1" />${rects.join('')}</svg>`;
  }

  // Verify default templates programmatically — catches any position that
  // overlaps a LOCKED zone before shipping. Logs warnings once on load.
  // NOTE: intersectsLocked is defined later in this module; Node/test
  // callers should invoke verifyDefaultTemplates() after module init.
  function verifyDefaultTemplates() {
    const issues = [];
    for (const tpl of DEFAULT_PDF_TEMPLATES) {
      for (const el of tpl.elements) {
        if (typeof intersectsLocked !== 'function') break;
        const hit = intersectsLocked({ x: el.x, y: el.y, w: el.w, h: el.h });
        if (hit) {
          issues.push(`${tpl.name} / ${el.id} overlaps LOCKED ${hit.id}`);
        }
      }
    }
    if (issues.length) {
      console.warn('[qf] default PDF templates have LOCKED overlaps:', issues);
    }
    return issues;
  }

  function seedDefaultPdfTemplatesIfNeeded() {
    try {
      if (localStorage.getItem(PDF_TEMPLATES_SEEDED_KEY) === '1') return;
      const existing = loadPdfTemplatesRaw();
      // Additive: if any templates exist already, mark as seeded without adding.
      if (!Array.isArray(existing) || existing.length === 0) {
        const nowTs = Date.now();
        const seeded = DEFAULT_PDF_TEMPLATES.slice(0, MAX_PDF_TEMPLATES).map((tpl, i) => ({
          ...tpl,
          id: `tpl_default_${tpl.id}_${nowTs}_${i}`,
          createdAt: nowTs,
          updatedAt: nowTs,
        }));
        localStorage.setItem(PDF_TEMPLATES_KEY, JSON.stringify(seeded));
      }
      localStorage.setItem(PDF_TEMPLATES_SEEDED_KEY, '1');
    } catch (e) {
      console.warn('[qf] seedDefaultPdfTemplatesIfNeeded failed:', e);
    }
  }

  // Raw load (no validation) — used during seeding to check emptiness without
  // side effects from validatePdfTemplate normalization.
  function loadPdfTemplatesRaw() {
    try {
      const raw = localStorage.getItem(PDF_TEMPLATES_KEY);
      if (!raw) return [];
      const arr = JSON.parse(raw);
      return Array.isArray(arr) ? arr : [];
    } catch (_e) {
      return [];
    }
  }

  function loadPdfTemplates() {
    seedDefaultPdfTemplatesIfNeeded();
    try {
      const raw = localStorage.getItem(PDF_TEMPLATES_KEY);
      if (!raw) return [];
      const arr = JSON.parse(raw);
      if (!Array.isArray(arr)) return [];
      return arr.map(validatePdfTemplate).filter(Boolean);
    } catch (_e) {
      return [];
    }
  }

  function savePdfTemplates(arr) {
    if (!Array.isArray(arr)) throw new Error('templates must be array');
    const normalized = arr.map(validatePdfTemplate).filter(Boolean);
    if (normalized.length > MAX_PDF_TEMPLATES) {
      throw new Error(`ครบ ${MAX_PDF_TEMPLATES} เทมเพลตแล้ว`);
    }
    try {
      localStorage.setItem(PDF_TEMPLATES_KEY, JSON.stringify(normalized));
    } catch (e) {
      throw new Error('บันทึกไม่สำเร็จ (อาจเพราะ localStorage เต็ม — ลดขนาดโลโก้): ' + e.message);
    }
    return normalized;
  }

  function validatePdfTemplate(tpl) {
    if (!tpl || typeof tpl !== 'object') return null;
    const pageW = J_AND_T_LAYOUT.pageSize.w;
    const pageH = J_AND_T_LAYOUT.pageSize.h;
    const clamp = (v, lo, hi) => Math.max(lo, Math.min(hi, Number(v) || 0));
    const elements = Array.isArray(tpl.elements) ? tpl.elements : [];
    const safeElements = elements.map(el => {
      if (!el || typeof el !== 'object') return null;
      const x = clamp(el.x, 0, pageW);
      const y = clamp(el.y, 0, pageH);
      const w = clamp(el.w ?? 40, 4, pageW);
      const h = clamp(el.h ?? 16, 4, pageH);
      const type = ['text', 'image', 'qrPlaceholder', 'variable'].includes(el.type) ? el.type : 'text';
      return {
        id: String(el.id || `el_${Math.random().toString(36).slice(2, 9)}`),
        type,
        x, y, w, h,
        size: typeof el.size === 'number' ? clamp(el.size, 4, 72) : 10,
        text: typeof el.text === 'string' ? el.text.slice(0, 200) : '',
        variable: typeof el.variable === 'string' ? el.variable.slice(0, 40) : '',
        dataUrl: typeof el.dataUrl === 'string' ? el.dataUrl.slice(0, 300000) : '',
        align: ['left', 'center', 'right'].includes(el.align) ? el.align : 'left',
        bold: !!el.bold,
      };
    }).filter(Boolean);
    return {
      id: String(tpl.id || `tpl_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`),
      name: String(tpl.name || 'เทมเพลตใหม่').slice(0, 60),
      isDefault: !!tpl.isDefault,
      presetId: typeof tpl.presetId === 'string' ? tpl.presetId.slice(0, 40) : null,
      createdAt: Number(tpl.createdAt) || Date.now(),
      updatedAt: Number(tpl.updatedAt) || Date.now(),
      brand: {
        logoDataUrl: typeof tpl.brand?.logoDataUrl === 'string' ? tpl.brand.logoDataUrl : null,
        shopName: typeof tpl.brand?.shopName === 'string' ? tpl.brand.shopName.slice(0, 60) : '',
        tagline: typeof tpl.brand?.tagline === 'string' ? tpl.brand.tagline.slice(0, 100) : '',
      },
      elements: safeElements,
      zones: {
        shrinkHeaderLogos: !!tpl.zones?.shrinkHeaderLogos,
        hideSkuTable: !!tpl.zones?.hideSkuTable,
        maskPhone: !!tpl.zones?.maskPhone,
        maskAddress: !!tpl.zones?.maskAddress,
      },
      overrides: validateOverrides(tpl.overrides),
    };
  }

  // PHASE 3 — Validate user-defined system element overrides.
  // Silently drops unknown keys and NEVER_OVERRIDE_KEYS. Clamps numeric
  // values to sane ranges so a malformed localStorage entry can't crash
  // the render pipeline.
  function validateOverrides(raw) {
    if (!raw || typeof raw !== 'object') return {};
    const pageW = J_AND_T_LAYOUT.pageSize.w;
    const pageH = J_AND_T_LAYOUT.pageSize.h;
    const clamp = (v, lo, hi, fallback) => {
      const n = Number(v);
      if (!Number.isFinite(n)) return fallback;
      return Math.max(lo, Math.min(hi, n));
    };
    const out = {};
    const sys = J_AND_T_LAYOUT.systemElements || {};
    for (const key of Object.keys(raw)) {
      if (NEVER_OVERRIDE_KEYS.has(key)) continue;
      if (!sys[key]) continue; // unknown key
      const ov = raw[key];
      if (!ov || typeof ov !== 'object') continue;
      const meta = sys[key];
      const def = meta.default || meta;
      const normalized = {
        enabled: !!ov.enabled,
        x: clamp(ov.x ?? def.x, 0, pageW, def.x),
        y: clamp(ov.y ?? def.y, 0, pageH, def.y),
        w: clamp(ov.w ?? def.w, 4, pageW, def.w),
        h: clamp(ov.h ?? def.h, 4, pageH, def.h),
      };
      if (meta.kind === 'text') {
        normalized.size = clamp(ov.size ?? def.size ?? 10, 4, 72, def.size ?? 10);
      }
      if (meta.kind === 'block') {
        normalized.fontSize = clamp(ov.fontSize ?? def.fontSize ?? 7, 4, 24, def.fontSize ?? 7);
        if (key === 'skuTable') {
          const cols = Array.isArray(ov.columns) ? ov.columns : def.columns;
          normalized.columns = (cols || ['name', 'sku', 'seller', 'qty'])
            .filter(c => ['name', 'sku', 'seller', 'qty'].includes(c));
          if (normalized.columns.length === 0) normalized.columns = ['name', 'qty'];
        }
      }
      out[key] = normalized;
    }
    return out;
  }

  function getActivePdfTemplate() {
    try {
      const id = localStorage.getItem(PDF_ACTIVE_TEMPLATE_KEY);
      if (!id) return null;
      const tpls = loadPdfTemplates();
      return tpls.find(t => t.id === id) || null;
    } catch (_e) { return null; }
  }

  function setActivePdfTemplate(id) {
    try {
      if (id == null) localStorage.removeItem(PDF_ACTIVE_TEMPLATE_KEY);
      else localStorage.setItem(PDF_ACTIVE_TEMPLATE_KEY, String(id));
    } catch (_e) {}
  }

  function addPdfTemplate(tpl) {
    const tpls = loadPdfTemplates();
    if (tpls.length >= MAX_PDF_TEMPLATES) throw new Error(`ครบ ${MAX_PDF_TEMPLATES} เทมเพลตแล้ว`);
    const normalized = validatePdfTemplate(tpl);
    if (!normalized) throw new Error('เทมเพลตไม่ถูกต้อง');
    normalized.createdAt = Date.now();
    normalized.updatedAt = Date.now();
    const next = [...tpls, normalized];
    savePdfTemplates(next);
    return normalized;
  }

  function updatePdfTemplate(tpl) {
    const tpls = loadPdfTemplates();
    const idx = tpls.findIndex(t => t.id === tpl.id);
    if (idx < 0) return addPdfTemplate(tpl);
    const normalized = validatePdfTemplate({ ...tpl, createdAt: tpls[idx].createdAt });
    if (!normalized) throw new Error('เทมเพลตไม่ถูกต้อง');
    normalized.updatedAt = Date.now();
    const next = [...tpls.slice(0, idx), normalized, ...tpls.slice(idx + 1)];
    savePdfTemplates(next);
    return normalized;
  }

  function deletePdfTemplate(id) {
    const tpls = loadPdfTemplates();
    const next = tpls.filter(t => t.id !== id);
    savePdfTemplates(next);
    if (localStorage.getItem(PDF_ACTIVE_TEMPLATE_KEY) === id) {
      setActivePdfTemplate(null);
    }
  }

  function duplicatePdfTemplate(id) {
    const tpls = loadPdfTemplates();
    const src = tpls.find(t => t.id === id);
    if (!src) throw new Error('ไม่พบเทมเพลต');
    if (tpls.length >= MAX_PDF_TEMPLATES) throw new Error(`ครบ ${MAX_PDF_TEMPLATES} เทมเพลตแล้ว`);
    const copy = JSON.parse(JSON.stringify(src));
    copy.id = `tpl_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`;
    copy.name = `${src.name} (สำเนา)`.slice(0, 60);
    copy.createdAt = Date.now();
    copy.updatedAt = Date.now();
    return addPdfTemplate(copy);
  }

  // Build mock record data from a batch of fulfillUnitIds — used both at
  // render time (real printing) and in the editor preview.
  function buildTemplateRecordData(batchIds) {
    const data = {
      alias: '',
      orderId: '',
      date: formatThaiDate(new Date()),
      shopName: '',
      customerName: '',
      trackingNumber: '',
      // PHASE 3 — system element overrides need richer record data.
      sortCode: '',
      serviceType: '',
      subZone: '',
      codLabel: '',
      skuList: [],          // [{ name, sku, sellerSku, qty }]
      addressBlock: '',     // multi-line (\n separated) — user sees placeholder in preview
    };
    try {
      if (Array.isArray(batchIds) && batchIds.length) {
        const firstId = batchIds[0];
        const rec = state.records?.get?.(firstId);
        if (rec) {
          data.orderId = String(rec.orderIds?.[0] || rec.orderId || '');
          data.trackingNumber = String(rec.trackingNumber || '');
          data.customerName = String(rec.customerName || rec.recipient?.name || '');
          const firstSku = rec.skuList?.[0];
          if (firstSku && typeof getAlias === 'function') {
            try { data.alias = getAlias(firstSku.productId) || ''; } catch (_e) {}
          }
          // System element field mapping — all fields are best-effort; render
          // pipeline skips the override if the value is empty.
          data.sortCode = String(rec.sortCode || rec.routeCode || '');
          data.serviceType = String(rec.serviceType || rec.serviceName || '');
          data.subZone = String(rec.subZoneCode || rec.subZone || '');
          data.codLabel = rec.isCod ? 'COD' : '';
          data.skuList = (rec.skuList || []).map(s => ({
            name: String(s.productName || s.name || ''),
            sku: String(s.skuId || s.sku || ''),
            sellerSku: String(s.sellerSku || s.sellerSkuId || ''),
            qty: Number(s.quantity || s.qty || 1),
          }));
          const addr = rec.recipient?.address || rec.address || '';
          data.addressBlock = typeof addr === 'string' ? addr
            : Array.isArray(addr) ? addr.join('\n') : '';
        }
      }
    } catch (_e) {}
    return data;
  }

  function mockTemplateRecordData(tpl) {
    return {
      alias: 'แดง1',
      orderId: '583619604412662802',
      date: formatThaiDate(new Date()),
      shopName: tpl?.brand?.shopName || 'ร้านตัวอย่าง',
      customerName: 'คุณลูกค้า ตัวอย่าง',
      trackingNumber: '795500112243',
      sortCode: 'EZ',
      serviceType: 'DROP-OFF',
      subZone: '004A',
      codLabel: 'COD',
      skuList: [
        { name: 'สินค้าตัวอย่าง A', sku: 'SKU-001', sellerSku: 'U C-BR-501', qty: 1 },
      ],
      addressBlock: '[ตัวอย่างที่อยู่ผู้รับ]',
    };
  }

  // PHASE 3 — Fetch + cache carrier logo PNG bytes. Returns Uint8Array or null.
  // Asset-bridge posts the URL; we fetch once and keep the bytes in state.
  async function ensureCarrierLogoBytes(key) {
    if (!['tiktok', 'jnt'].includes(key)) return null;
    if (state.carrierLogoBytes?.[key]) return state.carrierLogoBytes[key];
    if (!state.carrierLogoUrls?.[key]) {
      try { window.postMessage({ __qfAsset: 'request_carrier_logos' }, '*'); } catch (_e) {}
      const deadline = Date.now() + 2000;
      while (!state.carrierLogoUrls?.[key] && Date.now() < deadline) await sleep(50);
    }
    const url = state.carrierLogoUrls?.[key];
    if (!url) return null;
    try {
      // `_origFetch` bypasses TikTok's fetch wrapper — carrier logos live on
      // extension-bundled asset URLs (chrome-extension://…) which shouldn't be
      // passing through TikTok's signing path in the first place, and the
      // wrapper's mutex would still serialize them when dividers from
      // parallel chunks all ask for the same logo simultaneously.
      const r = await _origFetch.call(window, url);
      if (!r.ok) return null;
      const buf = new Uint8Array(await r.arrayBuffer());
      state.carrierLogoBytes = { ...state.carrierLogoBytes, [key]: buf };
      return buf;
    } catch (_e) {
      return null;
    }
  }

  function formatThaiDate(d) {
    const months = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];
    return `${d.getDate()} ${months[d.getMonth()]} ${d.getFullYear() + 543}`;
  }

  function resolveTemplateVariable(varName, recordData) {
    if (!varName) return '';
    const key = varName.replace(/^\{|\}$/g, '');
    return String(recordData?.[key] ?? '');
  }

  // Check intersection of a rect with any LOCKED region. Coordinates are in
  // top-left editor space (x, y, w, h) — but LOCKED regions are stored in
  // PDF bottom-left space. This helper takes top-left coords.
  function intersectsLocked(rectTL) {
    const pageH = J_AND_T_LAYOUT.pageSize.h;
    for (const loc of J_AND_T_LAYOUT.locked) {
      // convert locked (bottom-left y) to top-left y
      const lockTLy = pageH - loc.y - loc.h;
      if (rectsOverlap(rectTL, { x: loc.x, y: lockTLy, w: loc.w, h: loc.h })) return loc;
    }
    return null;
  }

  function rectsOverlap(a, b) {
    return !(a.x + a.w <= b.x || b.x + b.w <= a.x || a.y + a.h <= b.y || b.y + b.h <= a.y);
  }

  function snapToGrid(v, grid = PDF_TEMPLATE_GRID_PT) {
    return Math.round(v / grid) * grid;
  }

  // ==================== PDF TEMPLATE RENDER PIPELINE ====================

  async function applyPdfTemplate(pdfBytes, template, recordData) {
    if (!window.PDFLib || !template) return pdfBytes;
    try {
      const { PDFDocument, rgb } = window.PDFLib;
      const fontBytes = await ensureFontBytes();
      const pdfDoc = await PDFDocument.load(pdfBytes);
      if (window.fontkit) pdfDoc.registerFontkit(window.fontkit);
      const font = await pdfDoc.embedFont(fontBytes, { subset: true });
      const pages = pdfDoc.getPages();
      if (!pages.length) return pdfBytes;

      // Embed logo once (shared across pages).
      let logoImg = null;
      if (template.brand?.logoDataUrl) {
        try {
          const dataUrl = template.brand.logoDataUrl;
          const mimeMatch = dataUrl.match(/^data:image\/(png|jpeg|jpg);base64,/i);
          if (mimeMatch) {
            const b64 = dataUrl.split(',')[1];
            const bin = atob(b64);
            const bytes = new Uint8Array(bin.length);
            for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
            if (/png/i.test(mimeMatch[1])) {
              logoImg = await pdfDoc.embedPng(bytes);
            } else {
              logoImg = await pdfDoc.embedJpg(bytes);
            }
          }
        } catch (e) {
          console.warn('[QF] template logo embed failed:', e);
        }
      }

      // Embed custom element images.
      const elementImages = new Map();
      for (const el of template.elements || []) {
        if (el.type === 'image' && el.dataUrl) {
          try {
            const mm = el.dataUrl.match(/^data:image\/(png|jpeg|jpg);base64,/i);
            if (!mm) continue;
            const b64 = el.dataUrl.split(',')[1];
            const bin = atob(b64);
            const bytes = new Uint8Array(bin.length);
            for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
            const img = /png/i.test(mm[1])
              ? await pdfDoc.embedPng(bytes)
              : await pdfDoc.embedJpg(bytes);
            elementImages.set(el.id, img);
          } catch (_e) {}
        }
      }

      // PHASE 3 — Embed carrier logos used by active overrides.
      // Only fetched if an override enables them (saves a network round-trip).
      const carrierLogoEmbeds = {};
      const overrides = template.overrides || {};
      if (overrides.tiktokLogo?.enabled) {
        const bytes = await ensureCarrierLogoBytes('tiktok');
        if (bytes) {
          try { carrierLogoEmbeds.tiktok = await pdfDoc.embedPng(bytes); }
          catch (_e) {}
        }
      }
      if (overrides.jntLogo?.enabled) {
        const bytes = await ensureCarrierLogoBytes('jnt');
        if (bytes) {
          try { carrierLogoEmbeds.jnt = await pdfDoc.embedPng(bytes); }
          catch (_e) {}
        }
      }

      for (const page of pages) {
        const { width: pw, height: ph } = page.getSize();

        // Zone transforms — cover original elements with white rectangles.
        // SAFETY: we only mask SHRINKABLE / address / sku table areas. Never
        // touch LOCKED zones (barcode/QR).
        if (template.zones?.hideSkuTable) {
          const z = J_AND_T_LAYOUT.shrinkable.find(s => s.id === 'skuTable');
          if (z) page.drawRectangle({ x: z.x, y: z.y, width: z.w, height: z.h, color: rgb(1, 1, 1) });
        }
        if (template.zones?.maskAddress) {
          const z = J_AND_T_LAYOUT.shrinkable.find(s => s.id === 'addressBlock');
          if (z) page.drawRectangle({ x: z.x, y: z.y, width: z.w, height: z.h, color: rgb(1, 1, 1), opacity: 0.85 });
        }
        if (template.zones?.maskPhone) {
          // Phone row sits just above the COD label — rough band.
          page.drawRectangle({ x: 10, y: 268, width: 140, height: 14, color: rgb(1, 1, 1) });
        }
        if (template.zones?.shrinkHeaderLogos) {
          // Blank the very top header band (TikTok/carrier logos) above all
          // LOCKED text. We avoid sort code area (x >= 180).
          page.drawRectangle({ x: 0, y: ph - 30, width: 180, height: 30, color: rgb(1, 1, 1) });
        }

        // Brand area (logo + shopName + tagline) in freed header region.
        // Position: top-left of page (PDF bottom-left origin: y near ph).
        const brand = template.brand || {};
        if (logoImg) {
          const targetH = 24;
          const ratio = logoImg.width / logoImg.height;
          const drawW = targetH * ratio;
          page.drawImage(logoImg, {
            x: 6, y: ph - targetH - 4, width: drawW, height: targetH, opacity: 0.85,
          });
        }
        if (brand.shopName) {
          page.drawText(String(brand.shopName), {
            x: logoImg ? 40 : 6, y: ph - 14, size: 9, font,
            color: rgb(0, 0, 0),
          });
        }
        if (brand.tagline) {
          page.drawText(String(brand.tagline), {
            x: logoImg ? 40 : 6, y: ph - 24, size: 7, font,
            color: rgb(0.2, 0.2, 0.2),
          });
        }

        // ========== PHASE 3 — SYSTEM ELEMENT OVERRIDES ==========
        // For each enabled override: (1) mask the original carrier rect with
        // a white rectangle, (2) redraw at user-chosen coords. NEVER mask
        // barcode/QR — those keys are not present in J_AND_T_LAYOUT.systemElements
        // and are additionally hard-guarded by NEVER_OVERRIDE_KEYS.
        const sysElements = J_AND_T_LAYOUT.systemElements || {};
        for (const [key, ov] of Object.entries(overrides)) {
          if (!ov?.enabled) continue;
          if (NEVER_OVERRIDE_KEYS.has(key)) continue; // safety guard
          const meta = sysElements[key];
          if (!meta || meta.maskable !== true) continue;

          // 1) Mask original carrier rect (bottom-left PDF coords in meta).
          try {
            page.drawRectangle({
              x: meta.x, y: meta.y, width: meta.w, height: meta.h,
              color: rgb(1, 1, 1),
            });
          } catch (_e) {}

          // 2) Redraw at new coords. Override coords stored in TOP-LEFT
          //    editor space → convert to bottom-left for pdf-lib.
          const newX = Number(ov.x) || 0;
          const newY_tl = Number(ov.y) || 0;
          const newW = Math.max(4, Number(ov.w) || 0);
          const newH = Math.max(4, Number(ov.h) || 0);
          const pdfY = ph - newY_tl - newH;

          try {
            if (meta.kind === 'image') {
              const embed = key === 'tiktokLogo' ? carrierLogoEmbeds.tiktok
                          : key === 'jntLogo'    ? carrierLogoEmbeds.jnt
                          : null;
              if (embed) {
                page.drawImage(embed, { x: newX, y: pdfY, width: newW, height: newH });
              }
            } else if (meta.kind === 'text') {
              const val = String(recordData?.[key] ?? '');
              if (!val) continue;
              const size = Math.max(4, Math.min(ov.size || meta.default?.size || 10, 48));
              page.drawText(val, {
                x: newX,
                y: pdfY + Math.max(2, newH - size - 2),
                size, font, color: rgb(0, 0, 0),
              });
            } else if (meta.kind === 'block' && key === 'skuTable') {
              const fs2 = Math.max(4, Math.min(ov.fontSize || 7, 20));
              const columns = Array.isArray(ov.columns) && ov.columns.length
                ? ov.columns : (meta.default?.columns || ['name', 'qty']);
              const rows = Array.isArray(recordData?.skuList) ? recordData.skuList : [];
              // Header row
              let cursorY = pdfY + newH - fs2 - 2;
              const colWidth = newW / columns.length;
              const headerMap = { name: 'Product', sku: 'SKU', seller: 'SellerSKU', qty: 'Qty' };
              for (let ci = 0; ci < columns.length; ci++) {
                page.drawText(headerMap[columns[ci]] || columns[ci], {
                  x: newX + ci * colWidth, y: cursorY, size: fs2, font, color: rgb(0, 0, 0),
                });
              }
              cursorY -= (fs2 + 2);
              // Data rows — clipped to block height
              const minY = pdfY + 2;
              for (const row of rows) {
                if (cursorY < minY) break;
                for (let ci = 0; ci < columns.length; ci++) {
                  const col = columns[ci];
                  const cellVal = col === 'name' ? row.name
                               : col === 'sku' ? row.sku
                               : col === 'seller' ? row.sellerSku
                               : col === 'qty' ? String(row.qty ?? '')
                               : '';
                  // Truncate to column width (approximate: 0.55 × fontSize per glyph)
                  const maxChars = Math.max(2, Math.floor(colWidth / (fs2 * 0.55)));
                  const text = (cellVal || '').slice(0, maxChars);
                  page.drawText(text, {
                    x: newX + ci * colWidth, y: cursorY, size: fs2, font, color: rgb(0, 0, 0),
                  });
                }
                cursorY -= (fs2 + 2);
              }
            } else if (meta.kind === 'block' && key === 'addressBlock') {
              const fs2 = Math.max(4, Math.min(ov.fontSize || 8, 20));
              const val = String(recordData?.addressBlock ?? '');
              if (!val) continue;
              const lines = val.split('\n');
              const lineH = fs2 + 2;
              let cursorY = pdfY + newH - fs2 - 2;
              const minY = pdfY + 2;
              for (const line of lines) {
                if (cursorY < minY) break;
                page.drawText(line, {
                  x: newX, y: cursorY, size: fs2, font, color: rgb(0, 0, 0),
                });
                cursorY -= lineH;
              }
            }
          } catch (e) {
            console.warn('[QF] override render failed for', key, e);
          }
        }
        // ========== END PHASE 3 OVERRIDES ==========

        // Custom elements. Coordinates stored in TOP-LEFT editor space —
        // convert to PDF bottom-left.
        for (const el of template.elements || []) {
          const pdfX = el.x;
          const pdfY = ph - el.y - el.h;
          if (el.type === 'text') {
            const size = Math.max(4, Math.min(el.size || 10, 48));
            const text = String(el.text || '');
            const tw = font.widthOfTextAtSize(text, size);
            let x = pdfX;
            if (el.align === 'center') x = pdfX + (el.w - tw) / 2;
            else if (el.align === 'right') x = pdfX + el.w - tw;
            page.drawText(text, {
              x, y: pdfY + Math.max(2, el.h - size - 2),
              size, font, color: rgb(0, 0, 0),
            });
          } else if (el.type === 'variable') {
            const size = Math.max(4, Math.min(el.size || 10, 48));
            const val = resolveTemplateVariable(el.variable, recordData);
            if (!val) continue;
            const tw = font.widthOfTextAtSize(val, size);
            let x = pdfX;
            if (el.align === 'center') x = pdfX + (el.w - tw) / 2;
            else if (el.align === 'right') x = pdfX + el.w - tw;
            page.drawText(val, {
              x, y: pdfY + Math.max(2, el.h - size - 2),
              size, font, color: rgb(0, 0, 0),
            });
          } else if (el.type === 'image') {
            const img = elementImages.get(el.id);
            if (img) {
              page.drawImage(img, {
                x: pdfX, y: pdfY, width: el.w, height: el.h, opacity: 0.9,
              });
            }
          } else if (el.type === 'qrPlaceholder') {
            // Draw a grayscale placeholder box (real QR generation deferred).
            page.drawRectangle({
              x: pdfX, y: pdfY, width: el.w, height: el.h,
              borderColor: rgb(0, 0, 0), borderWidth: 1, color: rgb(0.95, 0.95, 0.95),
            });
            const label = 'QR';
            page.drawText(label, {
              x: pdfX + el.w / 2 - 6, y: pdfY + el.h / 2 - 4,
              size: 8, font, color: rgb(0, 0, 0),
            });
          }
        }
      }

      return await pdfDoc.save();
    } catch (e) {
      console.warn('[QF] applyPdfTemplate error:', e);
      return pdfBytes;
    }
  }

  // ==================== LABEL OVERLAY SETTINGS ====================

  function openLabelOverlaySettings() {
    document.querySelectorAll('.qf-lo-overlay').forEach(e => e.remove());
    const cfg = loadLabelOverlay();
    const overlay = document.createElement('div');
    overlay.className = 'qf-modal-overlay qf-lo-overlay';
    const opPct = Math.round((cfg.opacity ?? 0.85) * 100);
    overlay.innerHTML = `
      <div class="qf-modal qf-lo-modal" role="dialog" aria-label="ปรับแต่งฉลาก">
        <div class="qf-modal-header">
          <span>🏷️ ปรับแต่งฉลาก (J&amp;T)</span>
          <button class="qf-modal-close">✕</button>
        </div>
        <div class="qf-lo-body">
          <div class="qf-lo-pv-col">
            <canvas class="qf-lo-canvas" width="496" height="698"></canvas>
            <div class="qf-lo-pv-caption">ตัวอย่าง J&amp;T A6</div>
          </div>
          <div class="qf-lo-fields">

            <!-- Card: toggle -->
            <div class="qf-lo-card">
              <label class="qf-lo-switch-row">
                <div class="qf-lo-switch-label">
                  <span class="qf-lo-switch-title">เปิดใช้งาน overlay</span>
                  <span class="qf-lo-switch-desc">แปะข้อความ/รูปบนฉลาก J&amp;T A6</span>
                </div>
                <div class="qf-lo-switch">
                  <input type="checkbox" id="qf-lo-enabled" ${cfg.enabled ? 'checked' : ''} />
                  <span class="qf-lo-switch-track"></span>
                </div>
              </label>
            </div>

            <div class="qf-lo-fields-wrap ${cfg.enabled ? '' : 'qf-lo-disabled'}" id="qf-lo-fields-wrap">

              <!-- Card: opacity -->
              <div class="qf-lo-card">
                <div class="qf-lo-row-between">
                  <span class="qf-lo-card-title">ความทึบ</span>
                  <span class="qf-lo-opacity-val" id="qf-lo-op-val">${opPct}%</span>
                </div>
                <input type="range" id="qf-lo-opacity" class="qf-lo-range" min="75" max="90" step="1" value="${opPct}" />
                <div class="qf-lo-hint">85% = ค่าเริ่มต้น · 90% = สูงสุด เพื่อไม่กลืนสีดำฉลาก</div>
              </div>

              <!-- Card: alias text + font size -->
              <div class="qf-lo-card">
                <div class="qf-lo-card-hd">
                  <span class="qf-lo-card-title">ชื่อ alias (ล่างฉลาก)</span>
                </div>
                <div class="qf-lo-hint">ข้อความ alias ที่แปะด้านล่างฉลาก — ตั้งขนาดได้</div>
                <div class="qf-lo-text-size-row">
                  <span class="qf-lo-field-label">ขนาด</span>
                  <input type="number" id="qf-lo-fontsize" class="qf-lo-size-input" min="0" max="28" step="1" value="${cfg.aliasFontSize || 0}" />
                  <span class="qf-lo-size-unit">pt (0 = อัตโนมัติ ~19pt)</span>
                </div>
              </div>

              <!-- Card: side strip text + font size -->
              <div class="qf-lo-card">
                <div class="qf-lo-card-hd">
                  <span class="qf-lo-card-title">↕ ข้อความแถบข้าง</span>
                  <span class="qf-lo-counter" id="qf-lo-mkt-cnt">${(cfg.marketingText||'').length}/50</span>
                </div>
                <div class="qf-lo-hint">วิ่งในแถบสีขาวซ้าย–ขวาของฉลาก</div>
                <input type="text" id="qf-lo-mkt" class="qf-lo-input" maxlength="50" value="${escapeHtml(cfg.marketingText || '')}" placeholder="กรุณาถ่ายรูปก่อนเปิดกล่องพัสดุ" />
                <div class="qf-lo-text-size-row" style="margin-top:6px;">
                  <span class="qf-lo-field-label">ขนาด</span>
                  <input type="number" id="qf-lo-mkt-size" class="qf-lo-size-input" min="0" max="15" step="1" value="${cfg.marketingFontSize || 0}" />
                  <span class="qf-lo-size-unit">pt (0 = อัตโนมัติ 7–12pt)</span>
                </div>
              </div>

              <!-- Card: shop logo + header texts (side-by-side) -->
              <div class="qf-lo-card">
                <div class="qf-lo-card-hd" style="margin-bottom:2px;">
                  <span class="qf-lo-card-title">🖼 โลโก้ร้าน + ข้อความ</span>
                </div>
                <div class="qf-lo-hint" style="margin-bottom:10px;">ออเดอร์ SKU เดียวเท่านั้น · บีบอัดอัตโนมัติ</div>
                <div class="qf-lo-shop-row">
                  <div class="qf-lo-img-col">
                    <div class="qf-lo-thumb-box" id="qf-lo-thumb-box">
                      ${cfg.shopImageDataUrl ? `<img src="${cfg.shopImageDataUrl}" class="qf-lo-thumb" />` : `<div class="qf-lo-thumb-empty"><span>ไม่มีรูป</span></div>`}
                    </div>
                    <div class="qf-lo-img-btns">
                      <button class="qf-btn-sm" id="qf-lo-upload-btn">อัปโหลด</button>
                      <button class="qf-btn-sm qf-btn-sm-danger" id="qf-lo-img-clear" ${cfg.shopImageDataUrl ? '' : 'disabled'}>ลบ</button>
                    </div>
                    <input type="file" id="qf-lo-file" accept="image/*" style="display:none" />
                  </div>
                  <div class="qf-lo-text-col">
                    <div class="qf-lo-field-row">
                      <span class="qf-lo-field-label">บรรทัดที่ 1</span>
                      <span class="qf-lo-counter" id="qf-lo-h1-cnt">${(cfg.headerMain||'').length}/50</span>
                    </div>
                    <input type="text" id="qf-lo-h1" class="qf-lo-input" maxlength="50" value="${escapeHtml(cfg.headerMain || '')}" placeholder="เพิ่มเพื่อนในไลน์ รับส่วนลดพิเศษ" />
                    <div class="qf-lo-text-size-row" style="margin-top:4px;">
                      <span class="qf-lo-field-label">ขนาด</span>
                      <input type="number" id="qf-lo-h1-size" class="qf-lo-size-input" min="0" max="14" step="1" value="${cfg.header1FontSize || 0}" />
                      <span class="qf-lo-size-unit">pt (0 = อัตโนมัติ 9pt)</span>
                    </div>
                    <div class="qf-lo-field-row" style="margin-top:8px;">
                      <span class="qf-lo-field-label">บรรทัดที่ 2</span>
                      <span class="qf-lo-counter" id="qf-lo-h2-cnt">${(cfg.headerSub||'').length}/50</span>
                    </div>
                    <input type="text" id="qf-lo-h2" class="qf-lo-input" maxlength="50" value="${escapeHtml(cfg.headerSub || '')}" placeholder="วันนี้เท่านั้น" />
                    <div class="qf-lo-text-size-row" style="margin-top:4px;">
                      <span class="qf-lo-field-label">ขนาด</span>
                      <input type="number" id="qf-lo-h2-size" class="qf-lo-size-input" min="0" max="12" step="1" value="${cfg.header2FontSize || 0}" />
                      <span class="qf-lo-size-unit">pt (0 = อัตโนมัติ 7pt)</span>
                    </div>
                  </div>
                </div>
              </div>

            </div>
          </div>
        </div>
        <div class="qf-modal-footer">
          <button id="qf-lo-cancel" class="qf-btn-cancel">ปิด</button>
          <button id="qf-lo-save" class="qf-btn-confirm">บันทึก</button>
        </div>
      </div>`;
    document.body.appendChild(overlay);

    const $  = sel => overlay.querySelector(sel);
    const close = () => overlay.remove();

    // ── Canvas live preview ───────────────────────────────────────────────
    const canvas = $('.qf-lo-canvas');
    const ctx = canvas.getContext('2d');
    const PV_W = canvas.width, PV_H = canvas.height;
    const SCALE = PV_W / 298, PDF_H = 420;
    const cx = x => x * SCALE;
    const cy = (y, h) => (PDF_H - y - h) * SCALE;
    const cw = w => w * SCALE;
    const ch = h => h * SCALE;
    const PV_MASKS = [
      {x:0,   y:0,    w:17,  h:340}, // OCR column L — y=0..340 (matches J_AND_T_MASK_RECTS)
      {x:283, y:0,    w:15,  h:340}, // OCR column R — same
      {x:0,   y:41.5, w:95,  h:20 },
      {x:0,   y:75,   w:298, h:3  },
      {x:0,   y:62,   w:298, h:2  },
      {x:0,   y:39.5, w:298, h:2  },
    ];
    let bgImg = null, shopPvImg = null;

    function renderPreview() {
      ctx.clearRect(0, 0, PV_W, PV_H);
      ctx.fillStyle = '#fff';
      ctx.fillRect(0, 0, PV_W, PV_H);
      if (bgImg && bgImg.complete && bgImg.naturalWidth) {
        ctx.globalAlpha = 0.55;
        ctx.filter = 'grayscale(1)';
        ctx.drawImage(bgImg, 0, 0, PV_W, PV_H);
        ctx.globalAlpha = 1;
        ctx.filter = 'none';
      } else {
        ctx.fillStyle = '#f0f0f0';
        ctx.fillRect(0, 0, PV_W, PV_H);
        ctx.fillStyle = '#bbb';
        ctx.font = `${cx(9.6)}px sans-serif`;
        ctx.textAlign = 'center';
        ctx.fillText('J&T label', PV_W / 2, PV_H / 2);
        ctx.textAlign = 'left';
      }
      if (!$('#qf-lo-enabled').checked) return;

      // White masks
      ctx.fillStyle = '#fff';
      for (const r of PV_MASKS) ctx.fillRect(cx(r.x), cy(r.y, r.h), cw(r.w), ch(r.h));

      // Vertical marketing text centered in mask zone
      const mktText = ($('#qf-lo-mkt').value || '').slice(0, 50);
      if (mktText) {
        ctx.save();
        const _mfsSel = Math.min(15, Math.max(0, parseInt($('#qf-lo-mkt-size').value, 10) || 0));
        const mktFontPx = _mfsSel > 0 ? cx(_mfsSel) : cx(Math.max(3.0, Math.min(6.6, 120 / Math.max(mktText.length, 13))));
        ctx.font = `bold ${mktFontPx}px sans-serif`;
        ctx.fillStyle = '#222';
        ctx.globalAlpha = 0.9;
        const tw = ctx.measureText(mktText).width;
        // Match PDF: aim center at PDF y=200 → canvas y = (420-200)*SCALE
        const canvasCenterY = (PDF_H - 200) * SCALE;
        const startY = canvasCenterY + tw / 2;
        for (const xPdf of [9, 291]) {
          ctx.save();
          ctx.translate(cx(xPdf), startY);
          ctx.rotate(-Math.PI / 2);
          ctx.fillText(mktText, 0, 0);
          ctx.restore();
        }
        ctx.restore();
      }

      // Shop image in footer zone
      if (shopPvImg && shopPvImg.complete && shopPvImg.naturalWidth) {
        ctx.drawImage(shopPvImg, cx(2), cy(42, 22), cw(22), ch(22));
      }

      // Header texts
      const h1 = ($('#qf-lo-h1').value || '').slice(0, 50);
      const h2 = ($('#qf-lo-h2').value || '').slice(0, 50);
      if (h1 || h2) {
        const hx = (shopPvImg && shopPvImg.complete) ? cx(26) : cx(2);
        ctx.fillStyle = '#000';
        ctx.globalAlpha = 0.85;
        const h1SzPdf = Math.min(14, Math.max(0, parseInt($('#qf-lo-h1-size').value, 10) || 0)) || 9;
        const h2SzPdf = Math.min(12, Math.max(0, parseInt($('#qf-lo-h2-size').value, 10) || 0)) || 7;
        if (h1) { ctx.font = `bold ${cx(h1SzPdf)}px sans-serif`; ctx.fillText(h1, hx, cy(55, h1SzPdf) + cx(h1SzPdf), cx(163)); }
        if (h2) { ctx.font = `${cx(h2SzPdf)}px sans-serif`; ctx.fillText(h2, hx, cy(44, h2SzPdf) + cx(h2SzPdf), cx(163)); }
        ctx.globalAlpha = 1;
      }

      // Alias text preview at bottom (shows font size effect)
      const _afsV = Math.min(28, Math.max(0, parseInt($('#qf-lo-fontsize').value, 10) || 0));
      const bigPt = _afsV > 0 ? _afsV : Math.min(420 * 0.05, 22) * 0.9;
      const smallPt = _afsV > 0 ? Math.max(7, bigPt * 0.65) : Math.min(420 * 0.032, 13) * 0.9;
      ctx.globalAlpha = 0.75;
      ctx.fillStyle = '#000';
      ctx.textAlign = 'center';
      ctx.font = `bold ${cx(bigPt)}px sans-serif`;
      ctx.fillText('ชื่อสินค้า alias', PV_W / 2, cy(4 + smallPt + 2, bigPt) + cx(bigPt));
      ctx.font = `${cx(smallPt)}px sans-serif`;
      ctx.fillText('ตัวเลือก/variant', PV_W / 2, cy(4, smallPt) + cx(smallPt));
      ctx.textAlign = 'left';
      ctx.globalAlpha = 1;
    }

    const sampleUrl = state.samplePreviewUrl || '';
    if (sampleUrl) {
      bgImg = new Image();
      bgImg.crossOrigin = 'anonymous';
      bgImg.onload = bgImg.onerror = renderPreview;
      bgImg.src = sampleUrl;
    }
    if (cfg.shopImageDataUrl) {
      shopPvImg = new Image();
      shopPvImg.onload = renderPreview;
      shopPvImg.src = cfg.shopImageDataUrl;
    }
    renderPreview();

    // ── Close ─────────────────────────────────────────────────────────────
    $('.qf-modal-close').addEventListener('click', close);
    $('#qf-lo-cancel').addEventListener('click', close);
    overlay.addEventListener('click', e => { if (e.target === overlay) close(); });

    // ── Toggle ────────────────────────────────────────────────────────────
    const enabledChk = $('#qf-lo-enabled');
    const fieldsWrap = $('#qf-lo-fields-wrap');
    enabledChk.addEventListener('change', () => {
      fieldsWrap.classList.toggle('qf-lo-disabled', !enabledChk.checked);
      renderPreview();
    });

    // ── Text inputs ────────────────────────────────────────────────────────
    const bindInput = (sel, cntSel) => {
      $(sel).addEventListener('input', e => { $(cntSel).textContent = `${e.target.value.length}/50`; renderPreview(); });
    };
    bindInput('#qf-lo-mkt', '#qf-lo-mkt-cnt');
    bindInput('#qf-lo-h1', '#qf-lo-h1-cnt');
    bindInput('#qf-lo-h2', '#qf-lo-h2-cnt');
    $('#qf-lo-opacity').addEventListener('input', e => { $('#qf-lo-op-val').textContent = `${e.target.value}%`; });
    ['#qf-lo-fontsize', '#qf-lo-mkt-size', '#qf-lo-h1-size', '#qf-lo-h2-size'].forEach(sel => {
      $(sel).addEventListener('input', () => renderPreview());
    });

    // ── Image upload ──────────────────────────────────────────────────────
    const fileInput = $('#qf-lo-file');
    $('#qf-lo-upload-btn').addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', async () => {
      const file = fileInput.files[0];
      if (!file) return;
      updateThumb(await compressImageToDataUrl(file, 300, 0.8));
    });

    // ── Clear image ───────────────────────────────────────────────────────
    $('#qf-lo-img-clear').addEventListener('click', () => updateThumb(''));

    function updateThumb(dataUrl) {
      overlay._shopImageDataUrl = dataUrl;
      const box = $('#qf-lo-thumb-box');
      if (dataUrl) {
        box.innerHTML = `<img src="${dataUrl}" class="qf-lo-thumb" />`;
        shopPvImg = new Image();
        shopPvImg.onload = renderPreview;
        shopPvImg.src = dataUrl;
      } else {
        box.innerHTML = `<div class="qf-lo-thumb-empty"><span>ไม่มีรูป</span></div>`;
        shopPvImg = null;
        renderPreview();
      }
      $('#qf-lo-img-clear').disabled = !dataUrl;
    }
    overlay._shopImageDataUrl = cfg.shopImageDataUrl || '';

    // ── Save ──────────────────────────────────────────────────────────────
    $('#qf-lo-save').addEventListener('click', () => {
      saveLabelOverlay({
        enabled: enabledChk.checked,
        marketingText: ($('#qf-lo-mkt').value || '').slice(0, 50),
        shopImageDataUrl: overlay._shopImageDataUrl || '',
        headerMain: ($('#qf-lo-h1').value || '').slice(0, 50),
        headerSub: ($('#qf-lo-h2').value || '').slice(0, 50),
        opacity: Math.min(parseInt($('#qf-lo-opacity').value, 10) / 100, 0.90),
        aliasFontSize: Math.min(28, Math.max(0, parseInt($('#qf-lo-fontsize').value, 10) || 0)),
        marketingFontSize: Math.min(15, Math.max(0, parseInt($('#qf-lo-mkt-size').value, 10) || 0)),
        header1FontSize: Math.min(14, Math.max(0, parseInt($('#qf-lo-h1-size').value, 10) || 0)),
        header2FontSize: Math.min(12, Math.max(0, parseInt($('#qf-lo-h2-size').value, 10) || 0)),
      });
      showToast('บันทึกการตั้งค่าฉลากแล้ว', 2000);
      close();
    });
  }

  async function compressImageToDataUrl(file, maxPx, quality) {
    return new Promise(resolve => {
      const img = new Image();
      const url = URL.createObjectURL(file);
      img.onload = () => {
        URL.revokeObjectURL(url);
        const scale = Math.min(1, maxPx / Math.max(img.width, img.height));
        const c = document.createElement('canvas');
        c.width = Math.round(img.width * scale); c.height = Math.round(img.height * scale);
        c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
        resolve(c.toDataURL('image/jpeg', quality));
      };
      img.onerror = () => { URL.revokeObjectURL(url); resolve(''); };
      img.src = url;
    });
  }

  // ==================== PDF TEMPLATE MANAGER MODAL ====================

  function openPdfTemplateManager() {
    document.querySelectorAll('.qf-pdf-tpl-mgr-overlay').forEach(e => e.remove());
    const overlay = document.createElement('div');
    overlay.className = 'qf-modal-overlay qf-pdf-tpl-mgr-overlay';

    const render = () => {
      const tpls = loadPdfTemplates();
      const activeId = localStorage.getItem(PDF_ACTIVE_TEMPLATE_KEY);
      const fmt = t => new Date(t).toLocaleString('th-TH', { dateStyle: 'short', timeStyle: 'short' });
      const rows = tpls.map(t => {
        const isActive = t.id === activeId;
        const preset = t.presetId ? getPresetById(t.presetId) : null;
        const presetTag = preset ? ` <span class="qf-pdf-tpl-preset-tag" title="สร้างจาก preset: ${escapeHtml(preset.name)}">${preset.thumbnail || '🎨'} ${escapeHtml(preset.name)}</span>` : '';
        return `
          <div class="qf-pdf-tpl-row ${isActive ? 'qf-pdf-tpl-row-active' : ''}" data-id="${escapeHtml(t.id)}">
            <div class="qf-pdf-tpl-thumb" data-tpl="${escapeHtml(t.id)}"></div>
            <div class="qf-pdf-tpl-info">
              <div class="qf-pdf-tpl-name">${escapeHtml(t.name)}${isActive ? ' <span class="qf-pdf-tpl-active-tag">ใช้งานอยู่</span>' : ''}${t.isDefault ? ' <span class="qf-pdf-tpl-default-tag">เริ่มต้น</span>' : ''}${presetTag}</div>
              <div class="qf-pdf-tpl-meta">อัพเดต ${escapeHtml(fmt(t.updatedAt))} · ${t.elements.length} elements</div>
            </div>
            <div class="qf-pdf-tpl-actions">
              <button class="qf-pdf-tpl-btn qf-pdf-tpl-edit">แก้ไข</button>
              <button class="qf-pdf-tpl-btn qf-pdf-tpl-dup">คัดลอก</button>
              <button class="qf-pdf-tpl-btn qf-pdf-tpl-activate">${isActive ? 'เลิกใช้' : 'ใช้งาน'}</button>
              <button class="qf-pdf-tpl-btn qf-pdf-tpl-export">Export JSON</button>
              <button class="qf-pdf-tpl-btn qf-pdf-tpl-del">🗑</button>
            </div>
          </div>
        `;
      }).join('');
      const full = tpls.length >= MAX_PDF_TEMPLATES;
      overlay.innerHTML = `
        <div class="qf-modal qf-pdf-tpl-mgr-modal" role="dialog">
          <div class="qf-workers-header">
            <div class="qf-modal-title" style="margin-bottom:0;">📄 เทมเพลต PDF (${tpls.length}/${MAX_PDF_TEMPLATES})</div>
            <button class="qf-workers-close qf-pdf-tpl-close">×</button>
          </div>
          <div class="qf-modal-body qf-pdf-tpl-mgr-body">
            <div class="qf-pdf-tpl-mgr-toolbar">
              <label class="qf-pdf-tpl-start-label">เริ่มจาก:
                <select class="qf-pdf-tpl-start-from" ${full ? 'disabled' : ''}>
                  <option value="__blank__">(ว่าง)</option>
                  ${DEFAULT_PDF_TEMPLATES.map(dt => `
                    <option value="default:${escapeHtml(dt.id)}" ${dt.id === 'default_branded' ? 'selected' : ''}>${escapeHtml(dt.name)}</option>
                  `).join('')}
                  ${tpls.length > 0 ? `<optgroup label="ของฉัน">${tpls.map(t => `<option value="user:${escapeHtml(t.id)}">${escapeHtml(t.name)}</option>`).join('')}</optgroup>` : ''}
                </select>
              </label>
              <button class="qf-btn-confirm qf-pdf-tpl-new" ${full ? 'disabled title="ครบ 3 แล้ว — ลบก่อน"' : ''}>+ สร้างเทมเพลตใหม่</button>
              <button class="qf-pdf-tpl-btn qf-pdf-tpl-import">📥 Import JSON</button>
            </div>
            <div class="qf-pdf-tpl-list">
              ${tpls.length === 0 ? '<div class="qf-pdf-tpl-empty">ยังไม่มีเทมเพลต — กดปุ่ม "สร้างเทมเพลตใหม่" เพื่อเริ่ม</div>' : rows}
            </div>
            <div class="qf-pdf-tpl-note">เทมเพลตที่ "ใช้งาน" จะถูกซ้อนบน PDF ฉลากหลัง alias watermark (ขาวดำเท่านั้น)</div>
          </div>
        </div>
      `;
      // Render thumbnails after DOM insertion.
      setTimeout(() => {
        overlay.querySelectorAll('.qf-pdf-tpl-thumb[data-tpl]').forEach(el => {
          const id = el.dataset.tpl;
          const tpl = tpls.find(t => t.id === id);
          if (tpl) renderTemplateThumbnail(el, tpl);
        });
      }, 0);
    };

    const cleanup = () => overlay.remove();
    overlay.onclick = (e) => { if (e.target === overlay) cleanup(); };
    const onKey = (e) => { if (e.key === 'Escape') { cleanup(); document.removeEventListener('keydown', onKey); } };
    document.addEventListener('keydown', onKey);

    overlay.addEventListener('click', (e) => {
      const row = e.target.closest('.qf-pdf-tpl-row');
      if (e.target.closest('.qf-pdf-tpl-close')) { cleanup(); return; }
      if (e.target.closest('.qf-pdf-tpl-new')) {
        if (loadPdfTemplates().length >= MAX_PDF_TEMPLATES) {
          showToast(`ครบ ${MAX_PDF_TEMPLATES} แล้ว — ลบก่อน`, 2500);
          return;
        }
        // Read the "เริ่มจาก" dropdown — if user picked a template to clone,
        // open editor pre-populated with that template's data (unsaved).
        // Otherwise (blank selection) show Phase 4a preset picker first.
        const startSelect = overlay.querySelector('.qf-pdf-tpl-start-from');
        const startVal = startSelect?.value || '';
        cleanup();
        if (startVal && startVal !== '__blank__') {
          openPdfTemplateEditorFromSeed(startVal);
        } else {
          openPresetPicker({
            mode: 'new',
            confirmReplace: false,
            onPick: (preset, isBlank) => {
              if (isBlank || !preset) {
                openPdfTemplateEditor();
                return;
              }
              try {
                const cloned = clonePresetAsDraft(preset);
                const seed = {
                  id: null,
                  name: preset.name,
                  presetId: preset.id,
                  isDefault: false,
                  elements: cloned.elements,
                  overrides: cloned.overrides,
                  zones: cloned.zones,
                  brand: cloned.brand,
                };
                openPdfTemplateEditor(null, seed);
              } catch (err) {
                console.warn('[qf] preset seed failed:', err);
                openPdfTemplateEditor();
              }
            },
          });
        }
        return;
      }
      if (e.target.closest('.qf-pdf-tpl-import')) {
        importPdfTemplateFlow(() => render());
        return;
      }
      if (!row) return;
      const id = row.dataset.id;
      if (e.target.closest('.qf-pdf-tpl-edit')) { cleanup(); openPdfTemplateEditor(id); return; }
      if (e.target.closest('.qf-pdf-tpl-dup')) {
        try { duplicatePdfTemplate(id); render(); showToast('คัดลอกแล้ว', 1500); }
        catch (err) { showToast(err.message, 2500); }
        return;
      }
      if (e.target.closest('.qf-pdf-tpl-activate')) {
        const currentActive = localStorage.getItem(PDF_ACTIVE_TEMPLATE_KEY);
        setActivePdfTemplate(currentActive === id ? null : id);
        render();
        return;
      }
      if (e.target.closest('.qf-pdf-tpl-export')) {
        const tpl = loadPdfTemplates().find(t => t.id === id);
        if (tpl) exportPdfTemplateJson(tpl);
        return;
      }
      if (e.target.closest('.qf-pdf-tpl-del')) {
        const tplToDel = loadPdfTemplates().find(t => t.id === id);
        const msg = tplToDel?.isDefault
          ? 'ลบ template เริ่มต้น? (ไม่สามารถย้อนกลับได้ — และจะไม่ seed ใหม่)'
          : 'ลบเทมเพลตนี้? (ไม่สามารถย้อนกลับได้)';
        if (!confirm(msg)) return;
        try { deletePdfTemplate(id); render(); showToast('ลบแล้ว', 1500); }
        catch (err) { showToast(err.message, 2500); }
      }
    });

    document.body.appendChild(overlay);
    render();
  }

  function exportPdfTemplateJson(tpl) {
    const blob = new Blob([JSON.stringify(tpl, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `pdf-template-${tpl.name.replace(/[^\w\u0E00-\u0E7F]+/g, '_')}.json`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function importPdfTemplateFlow(onDone) {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = 'application/json,.json';
    input.onchange = async () => {
      const file = input.files?.[0];
      if (!file) return;
      try {
        const text = await file.text();
        const parsed = JSON.parse(text);
        parsed.id = `tpl_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`;
        addPdfTemplate(parsed);
        showToast('Import สำเร็จ', 1500);
        onDone?.();
      } catch (err) {
        showToast('Import ล้มเหลว: ' + err.message, 3000);
      }
    };
    input.click();
  }

  function renderTemplateThumbnail(container, tpl) {
    // Simple canvas thumbnail — scaled-down layout preview.
    const w = 80, h = 112;
    const canvas = document.createElement('canvas');
    canvas.width = w; canvas.height = h;
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#fff';
    ctx.fillRect(0, 0, w, h);
    ctx.strokeStyle = '#d0d5dd';
    ctx.strokeRect(0, 0, w, h);
    // LOCKED regions as grey
    const sx = w / J_AND_T_LAYOUT.pageSize.w;
    const sy = h / J_AND_T_LAYOUT.pageSize.h;
    ctx.fillStyle = '#e5e7eb';
    for (const loc of J_AND_T_LAYOUT.locked) {
      const ty = h - (loc.y + loc.h) * sy;
      ctx.fillRect(loc.x * sx, ty, loc.w * sx, loc.h * sy);
    }
    // Elements
    ctx.fillStyle = '#111';
    for (const el of tpl.elements || []) {
      ctx.fillStyle = el.type === 'image' ? '#6b7280' : (el.type === 'qrPlaceholder' ? '#374151' : '#111');
      ctx.fillRect(el.x * sx, el.y * sy, Math.max(2, el.w * sx), Math.max(2, el.h * sy));
    }
    // Shop name top-left
    if (tpl.brand?.shopName) {
      ctx.fillStyle = '#111';
      ctx.font = '7px sans-serif';
      ctx.fillText(String(tpl.brand.shopName).slice(0, 12), 2, 9);
    }
    container.innerHTML = '';
    container.appendChild(canvas);
  }

  // ==================== PDF TEMPLATE EDITOR ====================

  // Open the editor pre-filled with a default or user template as a clone.
  // The draft has no id — save creates a fresh template. Useful for
  // "เริ่มจาก" dropdown.
  function openPdfTemplateEditorFromSeed(seedSpec) {
    // seedSpec format: "default:<id>" or "user:<id>"
    const [kind, id] = String(seedSpec || '').split(':');
    let source = null;
    if (kind === 'default') {
      source = DEFAULT_PDF_TEMPLATES.find(t => t.id === id) || null;
    } else if (kind === 'user') {
      source = loadPdfTemplates().find(t => t.id === id) || null;
    }
    if (!source) { openPdfTemplateEditor(); return; }
    const seed = JSON.parse(JSON.stringify(source));
    seed.id = null; // new template on save
    seed.isDefault = false;
    seed.name = `${seed.name} (สำเนา)`.slice(0, 60);
    // presetId carried over so "used preset" badge can display on clones.
    if (typeof source.presetId === 'string') seed.presetId = source.presetId;
    openPdfTemplateEditor(null, seed);
  }

  // ==================== PDF PRESET PICKER (Phase 4a) ====================
  //
  // Shown:
  //   1. Before opening a blank editor (when user clicks "+ สร้างเทมเพลตใหม่"
  //      without picking a specific "เริ่มจาก" source)
  //   2. From the editor toolbar via the "🎨 เปลี่ยน layout" button
  //
  // options:
  //   { mode: 'new' | 'replace', onPick: (preset|null, isBlank:bool) => void,
  //     confirmReplace: bool }
  function openPresetPicker(options) {
    const opts = options || {};
    const presets = getVerifiedPresets();
    document.querySelectorAll('.qf-pdf-preset-picker-overlay').forEach(e => e.remove());
    const overlay = document.createElement('div');
    overlay.className = 'qf-modal-overlay qf-pdf-preset-picker-overlay';

    const cards = presets.map(preset => {
      const svg = renderPresetThumbnail(preset);
      return `
        <button class="qf-pdf-preset-card" data-preset-id="${escapeHtml(preset.id)}"
          aria-label="${escapeHtml(preset.name)}">
          <div class="qf-pdf-preset-thumb">${svg}</div>
          <div class="qf-pdf-preset-emoji">${preset.thumbnail || ''}</div>
          <div class="qf-pdf-preset-name">${escapeHtml(preset.name)}</div>
          <div class="qf-pdf-preset-desc">${escapeHtml(preset.description || '')}</div>
        </button>`;
    }).join('');

    const blankCard = opts.mode === 'replace' ? '' : `
      <button class="qf-pdf-preset-card qf-pdf-preset-card-blank" data-preset-id="__blank__"
        aria-label="เริ่มว่างเปล่า">
        <div class="qf-pdf-preset-thumb qf-pdf-preset-thumb-blank">
          <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 96 120" width="96" height="120">
            <rect x="0" y="0" width="96" height="120" fill="#ffffff" stroke="#cbd5e1" stroke-width="1" />
            <line x1="48" y1="35" x2="48" y2="85" stroke="#94a3b8" stroke-width="2" />
            <line x1="23" y1="60" x2="73" y2="60" stroke="#94a3b8" stroke-width="2" />
          </svg>
        </div>
        <div class="qf-pdf-preset-emoji">✨</div>
        <div class="qf-pdf-preset-name">เริ่มว่างเปล่า</div>
        <div class="qf-pdf-preset-desc">สร้างจาก element ทีละตัว</div>
      </button>`;

    const title = opts.mode === 'replace' ? '🎨 เปลี่ยน layout' : '🎨 เลือก layout เริ่มต้น';
    overlay.innerHTML = `
      <div class="qf-modal qf-pdf-preset-picker-modal" role="dialog">
        <div class="qf-workers-header">
          <div class="qf-modal-title" style="margin-bottom:0;">${title}</div>
          <button class="qf-workers-close qf-pdf-preset-close">×</button>
        </div>
        <div class="qf-modal-body qf-pdf-preset-picker-body">
          <div class="qf-pdf-preset-grid">
            ${cards}
            ${blankCard}
          </div>
          <div class="qf-pdf-preset-note">
            🔒 barcode และ QR ล็อคไว้ทุก layout — ปลอดภัยสำหรับสแกน
          </div>
        </div>
      </div>
    `;

    const cleanup = () => {
      overlay.remove();
      document.removeEventListener('keydown', onKey);
    };
    const onKey = (e) => { if (e.key === 'Escape') cleanup(); };
    document.addEventListener('keydown', onKey);
    overlay.onclick = (e) => { if (e.target === overlay) cleanup(); };

    overlay.addEventListener('click', (e) => {
      if (e.target.closest('.qf-pdf-preset-close')) { cleanup(); return; }
      const card = e.target.closest('.qf-pdf-preset-card');
      if (!card) return;
      const pid = card.dataset.presetId;
      if (pid === '__blank__') {
        cleanup();
        try { opts.onPick?.(null, true); } catch (err) { console.warn('[qf] preset onPick failed:', err); }
        return;
      }
      const preset = getPresetById(pid);
      if (!preset) { cleanup(); return; }
      if (opts.confirmReplace) {
        const ok = confirm('เปลี่ยน layout จะเขียนทับ elements + overrides + brand ที่แก้ไว้ตอนนี้ ต้องการทำต่อ?');
        if (!ok) return;
      }
      cleanup();
      try { opts.onPick?.(preset, false); } catch (err) { console.warn('[qf] preset onPick failed:', err); }
    });

    document.body.appendChild(overlay);
  }

  function openPdfTemplateEditor(templateId, seedDraft) {
    document.querySelectorAll('.qf-pdf-editor-overlay').forEach(e => e.remove());
    const existing = templateId ? loadPdfTemplates().find(t => t.id === templateId) : null;
    if (templateId && !existing) {
      showToast('ไม่พบเทมเพลต', 2000);
      return;
    }

    // Working draft (mutable during edit). Priority: existing > seedDraft > blank.
    let tpl = existing ? JSON.parse(JSON.stringify(existing)) : (seedDraft ? JSON.parse(JSON.stringify(seedDraft)) : {
      id: null,
      name: 'เทมเพลตใหม่',
      presetId: null,
      brand: { logoDataUrl: null, shopName: '', tagline: '' },
      elements: [],
      zones: { shrinkHeaderLogos: false, hideSkuTable: false, maskPhone: false, maskAddress: false },
      overrides: {},
    });
    if (!tpl.overrides) tpl.overrides = {};
    let selectedId = null;
    let selectedSysKey = null;  // PHASE 3 — key of selected system override
    let showGrid = true;
    const history = [JSON.parse(JSON.stringify(tpl))];
    const maxHistory = 10;

    const pushHistory = () => {
      history.push(JSON.parse(JSON.stringify(tpl)));
      while (history.length > maxHistory) history.shift();
    };
    const undo = () => {
      if (history.length > 1) {
        history.pop();
        tpl = JSON.parse(JSON.stringify(history[history.length - 1]));
        selectedId = null;
        rerender();
      }
    };

    const overlay = document.createElement('div');
    overlay.className = 'qf-modal-overlay qf-pdf-editor-overlay';
    const canvasW = J_AND_T_LAYOUT.pageSize.w * PDF_TEMPLATE_SCALE;
    const canvasH = J_AND_T_LAYOUT.pageSize.h * PDF_TEMPLATE_SCALE;

    // Sample PDF preview as canvas background — gives users visual context.
    // URL is supplied by asset-bridge (ISOLATED world). If not yet received
    // (race on editor open), request it now and apply when it arrives.
    let showSampleBg = true;
    if (!state.samplePreviewUrl) {
      try { window.postMessage({ __qfAsset: 'request_sample_preview' }, '*'); } catch (_e) {}
    }

    overlay.innerHTML = `
      <div class="qf-modal qf-pdf-editor-modal" role="dialog">
        <div class="qf-workers-header">
          <div class="qf-modal-title" style="margin-bottom:0;">✏️ เทมเพลต PDF — <input class="qf-pdf-editor-name" value="${escapeHtml(tpl.name)}" /></div>
          <button class="qf-workers-close qf-pdf-editor-close">×</button>
        </div>
        <div class="qf-pdf-editor-grid">
          <aside class="qf-pdf-editor-left">
            <div class="qf-pdf-editor-section-title">Elements</div>
            <button class="qf-pdf-editor-add" data-add="text">+ ข้อความ</button>
            <button class="qf-pdf-editor-add" data-add="image">+ รูปภาพ</button>
            <button class="qf-pdf-editor-add" data-add="logo">+ โลโก้ร้าน</button>
            <button class="qf-pdf-editor-add" data-add="qrPlaceholder">+ QR</button>
            <div class="qf-pdf-editor-section-title">ตัวแปร</div>
            <button class="qf-pdf-editor-add" data-addvar="alias">+ {alias}</button>
            <button class="qf-pdf-editor-add" data-addvar="orderId">+ {orderId}</button>
            <button class="qf-pdf-editor-add" data-addvar="date">+ {date}</button>
            <button class="qf-pdf-editor-add" data-addvar="shopName">+ {shopName}</button>
            <button class="qf-pdf-editor-add" data-addvar="customerName">+ {customerName}</button>
            <button class="qf-pdf-editor-add" data-addvar="trackingNumber">+ {trackingNumber}</button>
            <div class="qf-pdf-editor-section-title">แบรนด์</div>
            <label class="qf-pdf-editor-field">ชื่อร้าน
              <input class="qf-pdf-editor-shopname" value="${escapeHtml(tpl.brand.shopName || '')}" />
            </label>
            <label class="qf-pdf-editor-field">Tagline
              <input class="qf-pdf-editor-tagline" value="${escapeHtml(tpl.brand.tagline || '')}" />
            </label>
            <label class="qf-pdf-editor-field">โลโก้ (PNG)
              <input class="qf-pdf-editor-logo" type="file" accept="image/png,image/jpeg" />
            </label>
            <div class="qf-pdf-editor-logo-preview"></div>
          </aside>
          <div class="qf-pdf-editor-center">
            <div class="qf-pdf-editor-toolbar">
              <label><input type="checkbox" class="qf-pdf-editor-grid-toggle" ${showGrid ? 'checked' : ''}> Grid</label>
              <label><input type="checkbox" class="qf-pdf-editor-bg-toggle" ${showSampleBg ? 'checked' : ''}> แสดงตัวอย่าง label</label>
              <button class="qf-pdf-editor-undo">↶ Undo</button>
              <button class="qf-pdf-editor-change-layout" title="เลือก preset อื่นมาแทน layout ปัจจุบัน">🎨 เปลี่ยน layout</button>
              <span class="qf-pdf-editor-hint">คลิก = เลือก · ลาก = ย้าย · Delete = ลบ · Arrow = เลื่อน</span>
            </div>
            <div class="qf-pdf-editor-canvas-wrap">
              <div class="qf-pdf-editor-canvas" style="width:${canvasW}px;height:${canvasH}px;"></div>
            </div>
          </div>
          <aside class="qf-pdf-editor-right">
            <div class="qf-pdf-editor-section-title">Properties</div>
            <div class="qf-pdf-editor-props">— เลือก element เพื่อแก้ไข —</div>
            <div class="qf-pdf-editor-section-title">Zones</div>
            <label class="qf-pdf-editor-check"><input type="checkbox" class="qf-pdf-editor-zone" data-zone="shrinkHeaderLogos" ${tpl.zones.shrinkHeaderLogos ? 'checked' : ''}> ซ่อน TikTok/carrier logos</label>
            <label class="qf-pdf-editor-check"><input type="checkbox" class="qf-pdf-editor-zone" data-zone="hideSkuTable" ${tpl.zones.hideSkuTable ? 'checked' : ''}> ซ่อน SKU table</label>
            <label class="qf-pdf-editor-check"><input type="checkbox" class="qf-pdf-editor-zone" data-zone="maskPhone" ${tpl.zones.maskPhone ? 'checked' : ''}> ปกปิดเบอร์โทร</label>
            <label class="qf-pdf-editor-check"><input type="checkbox" class="qf-pdf-editor-zone" data-zone="maskAddress" ${tpl.zones.maskAddress ? 'checked' : ''}> ปกปิดที่อยู่</label>
            <div class="qf-pdf-editor-section-title">📐 System elements</div>
            <div class="qf-pdf-editor-sys-list"></div>
            <div class="qf-pdf-editor-sys-note">🔒 Barcode / QR ขยับไม่ได้ (สแกน-critical)</div>
          </aside>
        </div>
        <div class="qf-pdf-editor-bottom">
          <button class="qf-btn-cancel qf-pdf-editor-cancel">ยกเลิก</button>
          <button class="qf-pdf-editor-saveas">บันทึกเป็นเทมเพลตใหม่</button>
          <button class="qf-btn-confirm qf-pdf-editor-save">${existing ? 'บันทึก' : 'สร้าง'}</button>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);

    const canvasEl = overlay.querySelector('.qf-pdf-editor-canvas');

    // ====== Rendering ======
    const rerender = () => {
      canvasEl.innerHTML = '';
      // Grid
      if (showGrid) {
        canvasEl.classList.add('qf-pdf-editor-canvas-grid');
      } else {
        canvasEl.classList.remove('qf-pdf-editor-canvas-grid');
      }
      // Sample preview background — inline style so grid gradient still shows
      // through (50% opacity).
      if (showSampleBg && state.samplePreviewUrl) {
        canvasEl.style.backgroundImage = `url("${state.samplePreviewUrl}")`;
        canvasEl.style.backgroundSize = '100% 100%';
        canvasEl.style.backgroundRepeat = 'no-repeat';
        canvasEl.classList.add('qf-pdf-editor-canvas-has-bg');
      } else {
        canvasEl.style.backgroundImage = '';
        canvasEl.classList.remove('qf-pdf-editor-canvas-has-bg');
      }

      // LOCKED regions — draw as red translucent boxes.
      const pageH = J_AND_T_LAYOUT.pageSize.h;
      for (const loc of J_AND_T_LAYOUT.locked) {
        const tlY = pageH - loc.y - loc.h;
        const box = document.createElement('div');
        box.className = 'qf-pdf-editor-locked';
        box.style.left = (loc.x * PDF_TEMPLATE_SCALE) + 'px';
        box.style.top = (tlY * PDF_TEMPLATE_SCALE) + 'px';
        box.style.width = (loc.w * PDF_TEMPLATE_SCALE) + 'px';
        box.style.height = (loc.h * PDF_TEMPLATE_SCALE) + 'px';
        box.title = `LOCKED: ${loc.label}`;
        box.textContent = loc.label;
        canvasEl.appendChild(box);
      }

      // Header zone hint
      const hz = J_AND_T_LAYOUT.headerZone;
      const hzEl = document.createElement('div');
      hzEl.className = 'qf-pdf-editor-headerzone';
      hzEl.style.left = (hz.x * PDF_TEMPLATE_SCALE) + 'px';
      hzEl.style.top = (hz.y * PDF_TEMPLATE_SCALE) + 'px';
      hzEl.style.width = (hz.w * PDF_TEMPLATE_SCALE) + 'px';
      hzEl.style.height = (hz.h * PDF_TEMPLATE_SCALE) + 'px';
      hzEl.textContent = 'Header zone — วางโลโก้/ข้อความแบรนด์';
      canvasEl.appendChild(hzEl);

      // PHASE 3 — render enabled system-element overrides as draggable
      // ghost boxes (blue dashed). Disabled overrides don't appear on the
      // canvas; LOCKED barcode/QR already rendered above as red boxes.
      const sysMeta = J_AND_T_LAYOUT.systemElements || {};
      for (const [key, ov] of Object.entries(tpl.overrides || {})) {
        if (!ov?.enabled) continue;
        const meta = sysMeta[key];
        if (!meta) continue;
        const node = document.createElement('div');
        node.className = 'qf-pdf-sys-el';
        if (selectedSysKey === key) node.classList.add('qf-pdf-sys-el-selected');
        node.dataset.sysKey = key;
        node.style.left = (ov.x * PDF_TEMPLATE_SCALE) + 'px';
        node.style.top = (ov.y * PDF_TEMPLATE_SCALE) + 'px';
        node.style.width = (ov.w * PDF_TEMPLATE_SCALE) + 'px';
        node.style.height = (ov.h * PDF_TEMPLATE_SCALE) + 'px';
        node.title = `${meta.label} — ลากเพื่อย้าย`;
        node.textContent = meta.label;
        // Resize handle
        const h2 = document.createElement('div');
        h2.className = 'qf-pdf-editor-resize';
        node.appendChild(h2);
        canvasEl.appendChild(node);
      }

      // Custom elements
      for (const el of tpl.elements) {
        const node = document.createElement('div');
        node.className = 'qf-pdf-editor-el qf-pdf-editor-el-' + el.type;
        if (el.id === selectedId) node.classList.add('qf-pdf-editor-el-selected');
        node.dataset.id = el.id;
        node.style.left = (el.x * PDF_TEMPLATE_SCALE) + 'px';
        node.style.top = (el.y * PDF_TEMPLATE_SCALE) + 'px';
        node.style.width = (el.w * PDF_TEMPLATE_SCALE) + 'px';
        node.style.height = (el.h * PDF_TEMPLATE_SCALE) + 'px';
        if (el.type === 'text') {
          node.textContent = el.text || '(ข้อความ)';
          node.style.fontSize = (el.size * PDF_TEMPLATE_SCALE * 0.75) + 'px';
          node.style.textAlign = el.align || 'left';
          if (el.bold) node.style.fontWeight = '700';
        } else if (el.type === 'variable') {
          node.textContent = '{' + el.variable + '}';
          node.style.fontSize = (el.size * PDF_TEMPLATE_SCALE * 0.75) + 'px';
          node.style.textAlign = el.align || 'left';
          if (el.bold) node.style.fontWeight = '700';
        } else if (el.type === 'image') {
          if (el.dataUrl) {
            const img = document.createElement('img');
            img.src = el.dataUrl;
            img.draggable = false;
            node.appendChild(img);
          } else {
            node.textContent = '[รูปภาพ]';
          }
        } else if (el.type === 'qrPlaceholder') {
          node.textContent = 'QR';
        }
        // Resize handle (bottom-right)
        const handle = document.createElement('div');
        handle.className = 'qf-pdf-editor-resize';
        node.appendChild(handle);
        canvasEl.appendChild(node);
      }

      // PHASE 3 — System elements sidebar list. Each row: checkbox + label.
      // LOCKED entries (barcode/QR) are rendered as disabled red rows at the
      // bottom with a 🔒 icon — non-draggable.
      const sysListEl = overlay.querySelector('.qf-pdf-editor-sys-list');
      if (sysListEl) {
        const lines = [];
        for (const [key, meta] of Object.entries(sysMeta)) {
          const ov = tpl.overrides?.[key];
          const checked = ov?.enabled ? 'checked' : '';
          lines.push(`
            <label class="qf-pdf-editor-sys-row" data-sys-key="${escapeHtml(key)}">
              <input type="checkbox" class="qf-pdf-editor-sys-toggle" data-sys-key="${escapeHtml(key)}" ${checked}>
              <span class="qf-pdf-editor-sys-label">${escapeHtml(meta.label)}</span>
            </label>
          `);
        }
        // LOCKED barcode/QR rows — always disabled, red 🔒
        const lockedRows = [
          { key: 'barcodeMain', label: 'Barcode (main)' },
          { key: 'qrCode',      label: 'QR code' },
          { key: 'barcodeLeft', label: 'Side barcode L' },
          { key: 'barcodeRight', label: 'Side barcode R' },
        ];
        for (const row of lockedRows) {
          lines.push(`
            <div class="qf-pdf-editor-sys-row qf-pdf-editor-sys-row-locked" title="สแกน-critical — ขยับไม่ได้">
              <span class="qf-pdf-editor-sys-lock">🔒</span>
              <span class="qf-pdf-editor-sys-label">${escapeHtml(row.label)}</span>
            </div>
          `);
        }
        sysListEl.innerHTML = lines.join('');
      }

      // Properties panel
      const propsEl = overlay.querySelector('.qf-pdf-editor-props');
      const sel = tpl.elements.find(e => e.id === selectedId);
      if (!sel) {
        propsEl.innerHTML = '— เลือก element เพื่อแก้ไข —';
      } else {
        propsEl.innerHTML = `
          <div class="qf-pdf-editor-prop-grid">
            <label>x <input type="number" class="qf-pdf-prop" data-k="x" value="${Math.round(sel.x)}" /></label>
            <label>y <input type="number" class="qf-pdf-prop" data-k="y" value="${Math.round(sel.y)}" /></label>
            <label>w <input type="number" class="qf-pdf-prop" data-k="w" value="${Math.round(sel.w)}" /></label>
            <label>h <input type="number" class="qf-pdf-prop" data-k="h" value="${Math.round(sel.h)}" /></label>
          </div>
          ${sel.type === 'text' || sel.type === 'variable' ? `
            <label class="qf-pdf-editor-field">ขนาดตัวอักษร (pt)
              <input type="number" class="qf-pdf-prop" data-k="size" value="${sel.size}" min="4" max="72" />
            </label>
            ${sel.type === 'text' ? `
              <label class="qf-pdf-editor-field">ข้อความ
                <input type="text" class="qf-pdf-prop" data-k="text" value="${escapeHtml(sel.text)}" />
              </label>` : `
              <label class="qf-pdf-editor-field">ตัวแปร
                <select class="qf-pdf-prop" data-k="variable">
                  ${['alias','orderId','date','shopName','customerName','trackingNumber'].map(v =>
                    `<option value="${v}" ${sel.variable === v ? 'selected' : ''}>{${v}}</option>`
                  ).join('')}
                </select>
              </label>`}
            <div class="qf-pdf-editor-align-row">
              <label><input type="radio" name="qf-pdf-align" class="qf-pdf-prop" data-k="align" value="left" ${sel.align === 'left' ? 'checked' : ''}> ซ้าย</label>
              <label><input type="radio" name="qf-pdf-align" class="qf-pdf-prop" data-k="align" value="center" ${sel.align === 'center' ? 'checked' : ''}> กลาง</label>
              <label><input type="radio" name="qf-pdf-align" class="qf-pdf-prop" data-k="align" value="right" ${sel.align === 'right' ? 'checked' : ''}> ขวา</label>
            </div>
            <label class="qf-pdf-editor-check"><input type="checkbox" class="qf-pdf-prop" data-k="bold" ${sel.bold ? 'checked' : ''}> ตัวหนา</label>
          ` : ''}
          ${sel.type === 'image' ? `
            <label class="qf-pdf-editor-field">เลือกรูป (PNG/JPG, ≤200KB)
              <input type="file" class="qf-pdf-img-upload" accept="image/png,image/jpeg" />
            </label>` : ''}
          <button class="qf-pdf-editor-delete">🗑 ลบ element</button>
        `;
      }

      // Logo preview
      const logoPrev = overlay.querySelector('.qf-pdf-editor-logo-preview');
      if (tpl.brand.logoDataUrl) {
        logoPrev.innerHTML = `<img src="${tpl.brand.logoDataUrl}" style="max-width:100%;max-height:60px;"> <button class="qf-pdf-editor-logo-remove">ลบโลโก้</button>`;
      } else {
        logoPrev.innerHTML = '<div style="color:#64748b;font-size:11px;">ยังไม่มีโลโก้</div>';
      }
    };
    rerender();

    // ====== Add elements ======
    const addElement = (type, opts = {}) => {
      const id = `el_${Date.now()}_${Math.random().toString(36).slice(2, 5)}`;
      const defaults = {
        text: { x: 10, y: 6, w: 120, h: 14, size: 10, text: 'ข้อความ', align: 'left' },
        image: { x: 10, y: 6, w: 60, h: 30 },
        qrPlaceholder: { x: 10, y: 6, w: 40, h: 40 },
        variable: { x: 10, y: 6, w: 120, h: 14, size: 10, variable: opts.variable || 'alias', align: 'left' },
      };
      const base = defaults[type] || defaults.text;
      const newEl = { id, type, ...base, dataUrl: '', bold: false };
      // Shift into non-LOCKED zone if initial spot overlaps.
      while (intersectsLocked(newEl)) {
        newEl.y += 4;
        if (newEl.y > J_AND_T_LAYOUT.pageSize.h - newEl.h) { newEl.y = 4; newEl.x += 4; }
        if (newEl.x > J_AND_T_LAYOUT.pageSize.w - newEl.w) { newEl.x = 4; break; }
      }
      tpl.elements.push(newEl);
      selectedId = id;
      pushHistory();
      rerender();
    };

    overlay.querySelectorAll('.qf-pdf-editor-add').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const kind = btn.dataset.add;
        const v = btn.dataset.addvar;
        if (v) { addElement('variable', { variable: v }); return; }
        if (kind === 'logo') {
          // "logo" shortcut: inserts an image placeholder and asks for PNG upload.
          addElement('image');
          setTimeout(() => {
            const input = overlay.querySelector('.qf-pdf-img-upload');
            input?.click();
          }, 50);
          return;
        }
        addElement(kind);
      });
    });

    // ====== Canvas drag / select / resize ======
    let dragState = null;
    canvasEl.addEventListener('mousedown', (e) => {
      // PHASE 3 — system override ghost box (drag/resize)
      const sysNode = e.target.closest('.qf-pdf-sys-el');
      if (sysNode) {
        const key = sysNode.dataset.sysKey;
        const ov = tpl.overrides?.[key];
        if (!ov) return;
        selectedSysKey = key;
        selectedId = null;
        const isResize = e.target.classList.contains('qf-pdf-editor-resize');
        dragState = {
          kind: 'sys',
          sysKey: key,
          mode: isResize ? 'resize' : 'move',
          startX: e.clientX, startY: e.clientY,
          origX: ov.x, origY: ov.y, origW: ov.w, origH: ov.h,
        };
        rerender();
        e.preventDefault();
        return;
      }

      const elNode = e.target.closest('.qf-pdf-editor-el');
      if (!elNode) { selectedId = null; selectedSysKey = null; rerender(); return; }
      const id = elNode.dataset.id;
      selectedId = id;
      selectedSysKey = null;
      const el = tpl.elements.find(x => x.id === id);
      if (!el) return;
      const isResize = e.target.classList.contains('qf-pdf-editor-resize');
      const startX = e.clientX;
      const startY = e.clientY;
      dragState = {
        kind: 'element',
        id,
        mode: isResize ? 'resize' : 'move',
        startX, startY,
        origX: el.x, origY: el.y, origW: el.w, origH: el.h,
      };
      rerender();
      e.preventDefault();
    });

    document.addEventListener('mousemove', (e) => {
      if (!dragState) return;
      const dx = (e.clientX - dragState.startX) / PDF_TEMPLATE_SCALE;
      const dy = (e.clientY - dragState.startY) / PDF_TEMPLATE_SCALE;

      // PHASE 3 — system override drag/resize
      if (dragState.kind === 'sys') {
        const ov = tpl.overrides?.[dragState.sysKey];
        if (!ov) return;
        let newX = ov.x, newY = ov.y, newW = ov.w, newH = ov.h;
        if (dragState.mode === 'move') {
          newX = snapToGrid(dragState.origX + dx);
          newY = snapToGrid(dragState.origY + dy);
        } else {
          newW = Math.max(8, snapToGrid(dragState.origW + dx));
          newH = Math.max(8, snapToGrid(dragState.origH + dy));
        }
        newX = Math.max(0, Math.min(J_AND_T_LAYOUT.pageSize.w - newW, newX));
        newY = Math.max(0, Math.min(J_AND_T_LAYOUT.pageSize.h - newH, newY));
        const probe = { x: newX, y: newY, w: newW, h: newH };
        const lockedHit = intersectsLocked(probe);
        if (lockedHit) {
          flashLockedRegion(lockedHit.id);
          return;
        }
        ov.x = newX; ov.y = newY; ov.w = newW; ov.h = newH;
        rerender();
        return;
      }

      const el = tpl.elements.find(x => x.id === dragState.id);
      if (!el) return;
      let newX = el.x, newY = el.y, newW = el.w, newH = el.h;
      if (dragState.mode === 'move') {
        newX = snapToGrid(dragState.origX + dx);
        newY = snapToGrid(dragState.origY + dy);
        newW = el.w; newH = el.h;
      } else {
        newW = Math.max(8, snapToGrid(dragState.origW + dx));
        newH = Math.max(8, snapToGrid(dragState.origH + dy));
      }
      // Clamp to page.
      newX = Math.max(0, Math.min(J_AND_T_LAYOUT.pageSize.w - newW, newX));
      newY = Math.max(0, Math.min(J_AND_T_LAYOUT.pageSize.h - newH, newY));
      // LOCKED check.
      const probe = { x: newX, y: newY, w: newW, h: newH };
      const lockedHit = intersectsLocked(probe);
      if (lockedHit) {
        // Flash the locked region red and reject the move.
        flashLockedRegion(lockedHit.id);
        return;
      }
      el.x = newX; el.y = newY; el.w = newW; el.h = newH;
      rerender();
    });

    document.addEventListener('mouseup', () => {
      if (dragState) {
        pushHistory();
        dragState = null;
      }
    });

    const flashLockedRegion = (lockId) => {
      const nodes = canvasEl.querySelectorAll('.qf-pdf-editor-locked');
      nodes.forEach(n => {
        if (n.textContent && J_AND_T_LAYOUT.locked.find(l => l.id === lockId)?.label === n.textContent) {
          n.classList.add('qf-pdf-editor-locked-flash');
          setTimeout(() => n.classList.remove('qf-pdf-editor-locked-flash'), 300);
        }
      });
    };

    // ====== Keyboard ======
    const onEditorKey = (e) => {
      if (!document.body.contains(overlay)) return;
      // Don't hijack typing into form fields.
      const tgt = e.target;
      if (tgt && (tgt.tagName === 'INPUT' || tgt.tagName === 'SELECT' || tgt.tagName === 'TEXTAREA')) return;
      if (e.key === 'Escape') {
        cleanup();
        return;
      }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') {
        e.preventDefault();
        undo();
        return;
      }
      // PHASE 3 — keyboard ops on selected system override
      if (selectedSysKey) {
        const ov = tpl.overrides?.[selectedSysKey];
        if (!ov) { selectedSysKey = null; return; }
        if (e.key === 'Delete' || e.key === 'Backspace') {
          delete tpl.overrides[selectedSysKey];
          selectedSysKey = null;
          pushHistory();
          rerender();
          e.preventDefault();
          return;
        }
        const step2 = e.shiftKey ? PDF_TEMPLATE_GRID_PT : 1;
        let ddx = 0, ddy = 0;
        if (e.key === 'ArrowLeft') ddx = -step2;
        else if (e.key === 'ArrowRight') ddx = step2;
        else if (e.key === 'ArrowUp') ddy = -step2;
        else if (e.key === 'ArrowDown') ddy = step2;
        else return;
        const newX = Math.max(0, Math.min(J_AND_T_LAYOUT.pageSize.w - ov.w, ov.x + ddx));
        const newY = Math.max(0, Math.min(J_AND_T_LAYOUT.pageSize.h - ov.h, ov.y + ddy));
        const probe = { x: newX, y: newY, w: ov.w, h: ov.h };
        const hit = intersectsLocked(probe);
        if (hit) { flashLockedRegion(hit.id); e.preventDefault(); return; }
        ov.x = newX; ov.y = newY;
        pushHistory();
        rerender();
        e.preventDefault();
        return;
      }

      if (!selectedId) return;
      const el = tpl.elements.find(x => x.id === selectedId);
      if (!el) return;
      if (e.key === 'Delete' || e.key === 'Backspace') {
        tpl.elements = tpl.elements.filter(x => x.id !== selectedId);
        selectedId = null;
        pushHistory();
        rerender();
        e.preventDefault();
        return;
      }
      const step = e.shiftKey ? PDF_TEMPLATE_GRID_PT : 1;
      let dx = 0, dy = 0;
      if (e.key === 'ArrowLeft') dx = -step;
      else if (e.key === 'ArrowRight') dx = step;
      else if (e.key === 'ArrowUp') dy = -step;
      else if (e.key === 'ArrowDown') dy = step;
      else return;
      const newX = Math.max(0, Math.min(J_AND_T_LAYOUT.pageSize.w - el.w, el.x + dx));
      const newY = Math.max(0, Math.min(J_AND_T_LAYOUT.pageSize.h - el.h, el.y + dy));
      const probe = { x: newX, y: newY, w: el.w, h: el.h };
      const hit = intersectsLocked(probe);
      if (hit) { flashLockedRegion(hit.id); e.preventDefault(); return; }
      el.x = newX; el.y = newY;
      pushHistory();
      rerender();
      e.preventDefault();
    };
    document.addEventListener('keydown', onEditorKey);

    // ====== Properties edits ======
    // Use 'change' for text/number (avoid re-render on every keystroke) and
    // 'input' only for checkbox/radio. This keeps typing smooth without
    // losing caret position.
    const handlePropChange = (e) => {
      const target = e.target;
      if (!target.classList.contains('qf-pdf-prop')) return;
      const el = tpl.elements.find(x => x.id === selectedId);
      if (!el) return;
      const k = target.dataset.k;
      let v;
      if (target.type === 'checkbox') v = target.checked;
      else if (target.type === 'radio') { if (!target.checked) return; v = target.value; }
      else if (target.type === 'number') v = Number(target.value) || 0;
      else v = target.value;
      if (['x', 'y', 'w', 'h'].includes(k)) {
        const probe = { x: el.x, y: el.y, w: el.w, h: el.h, [k]: v };
        const hit = intersectsLocked(probe);
        if (hit) {
          flashLockedRegion(hit.id);
          target.value = el[k];
          return;
        }
      }
      el[k] = v;
      pushHistory();
      rerender();
    };
    overlay.addEventListener('change', handlePropChange);
    // Radios + checkboxes don't fire 'change' reliably without blur — also
    // bind 'input' but filter out text/number to avoid focus loss.
    overlay.addEventListener('input', (e) => {
      const t = e.target;
      if (!t.classList.contains('qf-pdf-prop')) return;
      if (t.type === 'text' || t.type === 'number') return;
      handlePropChange(e);
    });

    // Delete button in properties
    overlay.addEventListener('click', (e) => {
      if (e.target.closest('.qf-pdf-editor-delete')) {
        tpl.elements = tpl.elements.filter(x => x.id !== selectedId);
        selectedId = null;
        pushHistory();
        rerender();
      }
    });

    // Image upload
    overlay.addEventListener('change', async (e) => {
      if (e.target.classList.contains('qf-pdf-img-upload')) {
        const file = e.target.files?.[0];
        if (!file) return;
        if (file.size > 200 * 1024) { showToast('รูปเกิน 200KB — บีบอัดก่อนนะ', 2500); return; }
        const dataUrl = await new Promise((resolve, reject) => {
          const r = new FileReader();
          r.onload = () => resolve(r.result);
          r.onerror = reject;
          r.readAsDataURL(file);
        });
        const el = tpl.elements.find(x => x.id === selectedId);
        if (el && el.type === 'image') {
          el.dataUrl = dataUrl;
          pushHistory();
          rerender();
        }
      }
      if (e.target.classList.contains('qf-pdf-editor-logo')) {
        const file = e.target.files?.[0];
        if (!file) return;
        if (file.size > 200 * 1024) { showToast('โลโก้เกิน 200KB — บีบอัดก่อนนะ', 2500); return; }
        const dataUrl = await new Promise((resolve, reject) => {
          const r = new FileReader();
          r.onload = () => resolve(r.result);
          r.onerror = reject;
          r.readAsDataURL(file);
        });
        tpl.brand.logoDataUrl = dataUrl;
        pushHistory();
        rerender();
      }
      if (e.target.classList.contains('qf-pdf-editor-zone')) {
        const z = e.target.dataset.zone;
        tpl.zones[z] = e.target.checked;
        pushHistory();
      }
    });

    overlay.addEventListener('click', (e) => {
      if (e.target.classList.contains('qf-pdf-editor-logo-remove')) {
        tpl.brand.logoDataUrl = null;
        pushHistory();
        rerender();
      }
    });

    // PHASE 3 — system override toggle. Checkbox change seeds defaults into
    // tpl.overrides[key] on enable, removes on disable.
    overlay.addEventListener('change', (e) => {
      const t = e.target;
      if (!t.classList.contains('qf-pdf-editor-sys-toggle')) return;
      const key = t.dataset.sysKey;
      const meta = J_AND_T_LAYOUT.systemElements?.[key];
      if (!meta) return;
      if (!tpl.overrides) tpl.overrides = {};
      if (t.checked) {
        const def = meta.default || meta;
        tpl.overrides[key] = {
          enabled: true,
          x: def.x, y: def.y, w: def.w, h: def.h,
          ...(meta.kind === 'text' ? { size: def.size || 10 } : {}),
          ...(meta.kind === 'block' ? {
            fontSize: def.fontSize || 7,
            ...(key === 'skuTable' ? { columns: def.columns || ['name', 'sku', 'seller', 'qty'] } : {}),
          } : {}),
        };
        selectedSysKey = key;
      } else {
        delete tpl.overrides[key];
        if (selectedSysKey === key) selectedSysKey = null;
      }
      pushHistory();
      rerender();
    });

    // Shop name / tagline / template name
    overlay.querySelector('.qf-pdf-editor-shopname').addEventListener('input', (e) => {
      tpl.brand.shopName = e.target.value;
    });
    overlay.querySelector('.qf-pdf-editor-tagline').addEventListener('input', (e) => {
      tpl.brand.tagline = e.target.value;
    });
    overlay.querySelector('.qf-pdf-editor-name').addEventListener('input', (e) => {
      tpl.name = e.target.value;
    });
    overlay.querySelector('.qf-pdf-editor-grid-toggle').addEventListener('change', (e) => {
      showGrid = e.target.checked;
      rerender();
    });
    overlay.querySelector('.qf-pdf-editor-bg-toggle').addEventListener('change', (e) => {
      showSampleBg = e.target.checked;
      rerender();
    });
    // Re-apply background when asset-bridge finally delivers the URL
    // (editor may have opened before asset-bridge posted the message).
    const onAssetMsg = (msg) => {
      if (msg.source !== window) return;
      if (msg.data?.__qfAsset === 'samplePreview' && msg.data.url && !state.samplePreviewUrl) {
        state.samplePreviewUrl = msg.data.url;
        if (document.body.contains(overlay)) rerender();
      }
    };
    window.addEventListener('message', onAssetMsg);
    overlay.querySelector('.qf-pdf-editor-undo').addEventListener('click', (e) => {
      e.stopPropagation(); undo();
    });

    // PHASE 4a — "เปลี่ยน layout" button opens preset picker in replace mode.
    // On pick, overwrite elements/overrides/zones/brand with the preset's
    // values. The template's name + presetId are updated too.
    const changeLayoutBtn = overlay.querySelector('.qf-pdf-editor-change-layout');
    if (changeLayoutBtn) {
      changeLayoutBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        openPresetPicker({
          mode: 'replace',
          confirmReplace: true,
          onPick: (preset, isBlank) => {
            if (isBlank || !preset) return;
            try {
              const cloned = clonePresetAsDraft(preset);
              tpl.elements = cloned.elements;
              tpl.overrides = cloned.overrides;
              tpl.zones = { ...tpl.zones, ...cloned.zones };
              tpl.brand = { ...tpl.brand, ...cloned.brand };
              tpl.presetId = preset.id;
              selectedId = null;
              selectedSysKey = null;
              pushHistory();
              rerender();
              const nameInput = overlay.querySelector('.qf-pdf-editor-shopname');
              if (nameInput && typeof cloned.brand.shopName === 'string') {
                nameInput.value = cloned.brand.shopName;
              }
              const tagInput = overlay.querySelector('.qf-pdf-editor-tagline');
              if (tagInput && typeof cloned.brand.tagline === 'string') {
                tagInput.value = cloned.brand.tagline;
              }
              showToast(`ใช้ layout "${preset.name}" แล้ว`, 1800);
            } catch (err) {
              console.warn('[qf] preset apply failed:', err);
              showToast('ใช้ layout ล้มเหลว — ใช้ editor เปล่าแทน', 2500);
            }
          },
        });
      });
    }

    // Save / Cancel
    const cleanup = () => {
      document.removeEventListener('keydown', onEditorKey);
      window.removeEventListener('message', onAssetMsg);
      overlay.remove();
    };
    overlay.querySelector('.qf-pdf-editor-close').addEventListener('click', cleanup);
    overlay.querySelector('.qf-pdf-editor-cancel').addEventListener('click', cleanup);
    overlay.querySelector('.qf-pdf-editor-save').addEventListener('click', () => {
      try {
        if (existing) {
          tpl.id = existing.id;
          updatePdfTemplate(tpl);
          showToast('บันทึกเทมเพลตแล้ว', 2000);
        } else {
          addPdfTemplate(tpl);
          showToast('สร้างเทมเพลตใหม่แล้ว', 2000);
        }
        cleanup();
        openPdfTemplateManager();
      } catch (err) {
        showToast(err.message, 3000);
      }
    });
    overlay.querySelector('.qf-pdf-editor-saveas').addEventListener('click', () => {
      try {
        const copy = JSON.parse(JSON.stringify(tpl));
        copy.id = null;
        copy.name = (tpl.name || 'เทมเพลต') + ' (ใหม่)';
        addPdfTemplate(copy);
        showToast('บันทึกเป็นเทมเพลตใหม่แล้ว', 2000);
        cleanup();
        openPdfTemplateManager();
      } catch (err) {
        showToast(err.message, 3000);
      }
    });
  }

  // Expose for debugging / external smoke-test scripts.
  window.__qfPdfTemplates = {
    load: loadPdfTemplates,
    save: savePdfTemplates,
    getActive: getActivePdfTemplate,
    setActive: setActivePdfTemplate,
    add: addPdfTemplate,
    update: updatePdfTemplate,
    delete: deletePdfTemplate,
    duplicate: duplicatePdfTemplate,
    apply: applyPdfTemplate,
    openEditor: openPdfTemplateEditor,
    openManager: openPdfTemplateManager,
    mockData: mockTemplateRecordData,
    presets: LAYOUT_PRESETS,
    verifyPresets,
    openPicker: openPresetPicker,
  };

  // Run preset LOCKED verification once at module init. intersectsLocked is
  // defined above; any overlap is logged via console.error so regressions
  // surface during the test plan.
  try { verifyPresets(); } catch (e) { console.warn('[qf] verifyPresets failed:', e); }

  // ==================== END PDF TEMPLATE BUILDER ====================

  // ==================== INIT ====================
  function init() {
    if (!isOrderPage() && !isLabelsPage()) return;
    // §7.6: Populate seller email from TikTok's session cookie (best-effort).
    try {
      const emailCookie = document.cookie.split(';').map(c => c.trim()).find(c => c.startsWith('passport_user_email='));
      if (emailCookie) state.sellerEmail = decodeURIComponent(emailCookie.split('=').slice(1).join('='));
    } catch {}
    buildWidget();
    // Show recovery banner after widget is built (DOM must exist)
    setTimeout(renderRecoveryBanner, 0);
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
