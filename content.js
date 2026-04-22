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
  // One-shot probe flag: dump first Shopee record shape to console when
  // pre-order detection fails, so the user can report the real field name.
  // Reset per scan in scanAllPages so probing survives across scans.
  let _shopeePreOrderProbeLogged = false;
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
  const WORKERS_STORAGE_KEY = 'qf_workers_v1';
  const OVERLAY_PREF_KEY = 'qf_overlay_enabled_v1';
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
        return { preset: p, showAlias: true, aliasScale: 1.0, showName: true, showVariant: true,
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
    // aliasScale 1.6 (minimal) → size 'l'; 1.0 → 'm'.
    const aliasSize = (p.aliasScale && p.aliasScale >= 1.4) ? 'l' : 'm';
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
    printedUnitIds: new Set(),     // fulfillUnitId พิมพ์แล้วใน session นี้ (labels page only, clear on scan)
    selectMode: false,
    selected: new Map(),           // key → {type, productId, skuId, scenario, sigKey}
    dateFilter: { start: null, end: null, field: 'createTime' }, // field: createTime | shipByTime | autoCancelTime
    advancedOpen: false,
    calendarMonth: null, // {year, month} cursor for the visible month
    overlayEnabled: loadOverlayPref(),
    workers: loadWorkers(),          // [{id, name, icon}] คนแพ็ค
    sellerEmail: null,               // §7.6: populated at widget boot from TikTok session cookie
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
                ${e.workerName ? `<div class="qf-history-packer">
                  <span class="qf-history-packer-icon">${escapeHtml((state.workers.find(w => w.id === e.workerId) || {}).icon || '')}</span>แพ็ค: ${escapeHtml(e.workerName)}</div>` : ''}
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

    overlay.querySelector('.qf-history-modal-close').onclick = (e) => {
      e.stopPropagation();
      cleanup();
    };
    clearAllBtn.onclick = async (e) => {
      e.stopPropagation();
      const ok = await confirmInline('ล้างประวัติทั้งหมด?', 'ล้าง');
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

  // Small inline confirm used by history "ล้างทั้งหมด". Matches qf-modal style.
  function confirmInline(title, dangerLabel) {
    return new Promise(resolve => {
      const overlay = document.createElement('div');
      overlay.className = 'qf-modal-overlay';
      overlay.innerHTML = `
        <div class="qf-modal" role="dialog">
          <div class="qf-modal-title">${escapeHtml(title)}</div>
          <div class="qf-modal-actions">
            <button class="qf-btn-cancel">ยกเลิก</button>
            <button class="qf-btn-confirm qf-btn-danger">${escapeHtml(dangerLabel || 'ยืนยัน')}</button>
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
    if (d.__qfAsset === 'manifest' && d.version) state.localVersion = d.version;
  });
  window.postMessage({ __qfAsset: 'request_font' }, '*');
  // Expose state so the top-level fetch hook can write apiListUrl into it
  window.__qfState = state;

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

  // Labels-page done: computed per active filter, not stored as product-level timestamp.
  // A product/variant/combo is done iff every fulfillUnitId that's currently visible
  // (passes carrier + pre-order filter) has been printed this session OR was already
  // marked printed (labelStatus=50) by the API. Resets on scan.
  //
  // Under "พิมพ์แล้ว" filter, every visible item is inherently printed, so marking
  // them as done/locked would block reprinting — which IS the point of that filter.
  // Treat that filter as a dedicated reprint view: no done state applies.
  function isLabelsDone(idSet) {
    if (state.labelStatusFilter === 'printed') return false;
    const visible = [...idSet].filter(id => passesCarrier(id) && passesPreOrder(id));
    if (!visible.length) return false;
    return visible.every(id => {
      if (state.printedUnitIds.has(id)) return true;
      return state.records.get(id)?.labelStatus === LABEL_STATUS_PRINTED;
    });
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
    // Warn user if an active planning session exists — scan would invalidate its IDs
    if (loadPlanningSession()) {
      const ok = await confirmInline('มีแผนค้างอยู่ — ละทิ้งแผนและ scan ใหม่?', 'ละทิ้ง');
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
        labelStatus: rec.labelStatus || null,
        createTime: rec.createTime || null,
        shipByTime: rec.shipByTime || null,
        autoCancelTime: rec.autoCancelTime || null,
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
    // If none are present (unknown schema), dump the first record shape to
    // the console once so the user can identify the real field and we can
    // pin it down. Guard with a module-level flag to avoid log spam.
    const firstItem = items?.[0] || {};
    let isPreOrder = !!(
      ext?.is_pre_order
      || pkgExt?.is_pre_order
      || fulfilment?.is_pre_order
      || firstItem.is_pre_order
      || firstItem.inner_item_ext_info?.is_pre_order
      || firstItem.item_ext_info?.is_pre_order
    );
    // TODO (needs live verification): Shopee's exact pre-order field is not
    // documented here. If isPreOrder stays false for every record on a tab
    // with known pre-order items, inspect the dump below and add the right
    // path. Dump runs once per scan — cleared in scanAllPages.
    if (!isPreOrder && !_shopeePreOrderProbeLogged && (ext || pkgExt || fulfilment)) {
      _shopeePreOrderProbeLogged = true;
      try {
        console.log('[QF Shopee pre-order probe] ext keys:', Object.keys(ext || {}));
        console.log('[QF Shopee pre-order probe] pkgExt keys:', Object.keys(pkgExt || {}));
        console.log('[QF Shopee pre-order probe] fulfilment keys:', Object.keys(fulfilment || {}));
        console.log('[QF Shopee pre-order probe] item keys:', Object.keys(firstItem));
        console.log('[QF Shopee pre-order probe] inner_item_ext_info:', firstItem.inner_item_ext_info);
        console.log('[QF Shopee pre-order probe] full ext:', ext);
      } catch {}
    }

    const sp = fulfilment || {};
    const carrierName = sp.fulfilment_channel_name || sp.masked_channel_name || 'ไม่ระบุ';
    const carrierId = String(sp.fulfilment_channel_name || ext.masked_channel_id || 'unknown');
    processLabelRecord({
      fulfillUnitId: String(fulfillUnitId),
      batchId: String(batchId || fulfillUnitId),
      orderIds: [String(ext.order_id || '')],
      labelStatus,
      isPreOrder,
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
  // §3.1: Chunk-plan modal thresholds
  const CHUNK_PROMPT_THRESHOLD = 200; // show modal when total > this
  const CHUNK_AUTO_SAFE_SIZE   = 200; // default X value in "แบ่งทุกๆ X ใบ" option

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
          const r = await fetch(state.fontUrl);
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

  function showChunkPlanModal({ total, multiSku = false, defaultPickingList = null }) {
    const lastPickingList = defaultPickingList !== null ? defaultPickingList : loadPickingListPref();
    const defaultN = Math.max(2, Math.ceil(total / CHUNK_AUTO_SAFE_SIZE));
    // §3.7: Default radio: 'single' if total<=200; else 'every' with X=200.
    const defaultMode = total <= CHUNK_PROMPT_THRESHOLD ? 'single' : 'every';

    return new Promise(resolve => {
      const overlay = document.createElement('div');
      overlay.className = 'qf-modal-overlay qf-chunk-plan-overlay';
      overlay.innerHTML = `
        <div class="qf-modal qf-chunk-plan-modal" role="dialog">
          <div class="qf-modal-title">เตรียมพิมพ์ฉลาก</div>
          <div class="qf-modal-body">
            <div class="qf-modal-count">${total} ใบ</div>
            <div class="qf-modal-summary">เลือกวิธีแบ่งไฟล์</div>
            <div class="qf-chunk-plan-options">
              <label class="qf-chunk-plan-opt">
                <input type="radio" name="qf-chunk-mode" value="single"
                  ${defaultMode === 'single' ? 'checked' : ''}>
                <span class="qf-chunk-plan-opt-body">
                  <span class="qf-chunk-plan-opt-title">ไฟล์เดียว</span>
                  <span class="qf-chunk-plan-opt-sub qf-chunk-single-label">${multiSku ? `แยก 1 ไฟล์ต่อสินค้า` : `${total} ใบในไฟล์เดียว`}</span>
                </span>
              </label>
              <label class="qf-chunk-plan-opt">
                <input type="radio" name="qf-chunk-mode" value="even"
                  ${defaultMode === 'even' ? 'checked' : ''}>
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
              <label class="qf-chunk-plan-opt">
                <input type="radio" name="qf-chunk-mode" value="every"
                  ${defaultMode === 'every' ? 'checked' : ''}>
                <span class="qf-chunk-plan-opt-body">
                  <span class="qf-chunk-plan-opt-title">
                    กำหนดจำนวน
                    <input type="number" name="qf-chunk-X" class="qf-chunk-plan-input qf-chunk-x-input"
                      min="1" max="${total}" value="${CHUNK_AUTO_SAFE_SIZE}"/>
                    <span class="qf-chunk-unit">ใบ/ไฟล์</span>
                  </span>
                  <span class="qf-chunk-plan-opt-sub qf-chunk-x-preview"></span>
                </span>
              </label>
            </div>
            <button type="button" class="qf-chunk-plan-more" aria-expanded="false">
              <span class="qf-chunk-plan-more-label">ตัวเลือกเพิ่มเติม</span>
              <span class="qf-chunk-plan-more-count"></span>
              <span class="qf-chunk-plan-chevron">›</span>
            </button>
            <div class="qf-chunk-plan-extras" hidden>
              <label class="qf-chunk-plan-extra-opt">
                <input type="checkbox" name="qf-chunk-picking" class="qf-chunk-picking-chk"
                  ${lastPickingList ? 'checked' : ''}/>
                แนบใบสรุปรายการ
              </label>
              ${multiSku ? `<label class="qf-chunk-plan-extra-opt">
                <input type="checkbox" name="qf-chunk-combined" class="qf-chunk-combined-chk"/>
                รวมทุกสินค้าเป็นไฟล์เดียว
              </label>` : ''}
              <label class="qf-chunk-plan-extra-opt">
                <input type="checkbox" name="qf-chunk-divider" class="qf-chunk-divider-chk" checked/>
                ใส่หน้าคั่นระหว่างสินค้าแต่ละแบบ
              </label>
            </div>
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
      const submitBtn = overlay.querySelector('.qf-chunk-plan-submit');

      const getMode = () => (radioEls().find(r => r.checked) || {}).value || 'single';

      const updatePreviews = () => {
        const n = parseInt(nInput.value) || defaultN;
        const x = parseInt(xInput.value) || CHUNK_AUTO_SAFE_SIZE;
        nPreview.textContent = `~${Math.floor(total/n)}–${Math.ceil(total/n)} ใบ/ไฟล์`;
        xPreview.textContent = `${Math.ceil(total/x)} ไฟล์`;
      };
      updatePreviews();

      // Clicking a number input should auto-select the parent radio.
      [nInput, xInput].forEach(inp => {
        inp.addEventListener('focus', () => {
          const radioVal = inp.name === 'qf-chunk-N' ? 'even' : 'every';
          const radio = overlay.querySelector(`input[value="${radioVal}"]`);
          if (radio) radio.checked = true;
        });
        inp.addEventListener('input', updatePreviews);

        // §3.3: On blur, clamp and flash red border if out-of-range.
        inp.addEventListener('blur', () => {
          const isN = inp.name === 'qf-chunk-N';
          const min = isN ? 2 : 1;
          const max = total;
          let val = parseInt(inp.value);
          if (isNaN(val) || val < min) val = min;
          if (val > max) val = max;
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

      // Update the single-mode sub-label based on combined checkbox state.
      const combinedChk = overlay.querySelector('.qf-chunk-combined-chk');
      const singleLabel = overlay.querySelector('.qf-chunk-single-label');
      if (combinedChk) {
        combinedChk.addEventListener('change', () => {
          if (singleLabel && multiSku) {
            singleLabel.textContent = combinedChk.checked
              ? `รวม ${total} ใบในไฟล์เดียว`
              : `แยก 1 ไฟล์ต่อสินค้า`;
          }
        });
      }

      // Collapsible "ตัวเลือกเพิ่มเติม" — collapsed by default, counter shows how many are on.
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

      submitBtn.onclick = () => {
        const mode = getMode();
        const withPickingList = overlay.querySelector('.qf-chunk-picking-chk')?.checked || false;
        const combined = overlay.querySelector('.qf-chunk-combined-chk')?.checked || false;
        const withDivider = overlay.querySelector('.qf-chunk-divider-chk')?.checked || false;
        savePickingListPref(withPickingList);

        let n = null;
        let x = null;

        if (mode === 'even') {
          n = parseInt(nInput.value);
          const clamped = Math.max(2, Math.min(total, isNaN(n) ? defaultN : n));
          if (clamped !== n) {
            nInput.value = clamped;
            nInput.classList.add('qf-input-error');
            setTimeout(() => nInput.classList.remove('qf-input-error'), 600);
            return; // Don't submit yet — let user see the correction.
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
      completeChunk(i, {url, pageCount}) {
        const row = cards[i];
        const filename = chunks[i].filename || `chunk-${i+1}.pdf`;
        blobUrls[i] = {url, filename, downloaded: false};
        row.querySelector('.qf-chunk-row-fill').style.width = '100%';
        row.querySelector('.qf-chunk-row-status').textContent = `${pageCount} หน้า`;
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

  function buildSkuRender(s) {
    // primary = big top line (alias + qty), secondary = small bottom (variant name only)
    const v = getVariantInfo(s.productId, s.skuId);
    const productAlias = (getAlias(s.productId) || '').trim() || shortName(s.productName);
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
      const productAlias = (getAlias(s.productId) || '').trim() || shortName(s.productName);
      const variantOverride = (v?.alias || '').trim();
      const variantName = variantOverride || variantDisplayName(s);
      const qty = s.quantity || 1;
      if (v?.replace && variantOverride) return `${variantOverride} ${qty}`;
      return variantName ? `${productAlias} ${variantName} ${qty}` : `${productAlias} ${qty}`;
    });
    return { mode: 'multi', text: parts.join(' + ') };
  }

  async function overlayAliasOnPdf(pdfBytes, fulfillUnitIds, onProgress, workerName, workerIcon) {
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
      if (workerName) {
        const { width: pw, height: ph } = page.getSize();
        const wSize = Math.min(ph * 0.04, 18);
        // §6.2: Drop "แพ็ค: " prefix and icon — Sarabun-Bold lacks Unicode symbol
        // glyphs so they render as □. Worker is identified by name only.
        // Position: top-RIGHT (user request) with measured text width.
        const wText = workerName;
        const wWidth = font.widthOfTextAtSize(wText, wSize);
        page.drawText(wText, {
          x: Math.max(6, pw - wWidth - 6),
          y: ph - wSize - 6,
          size: wSize, font,
          color: rgb(0, 0, 0), opacity: 0.85,
        });
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
  // Check if record's labelStatus matches current filter — needed AFTER optimistic
  // updates flip rec.labelStatus from 30 → 50, so counts drop in "ยังไม่พิมพ์".
  // Fallback: if labelStatus is missing (record predates 1.12.1 or normalize
  // missed the field), assume it matches the server-side filter we scanned with.
  function passesLabelStatus(id) {
    if (state.labelStatusFilter === 'all') return true;
    const rec = state.records.get(id);
    if (!rec) return true;
    if (rec.labelStatus == null) return true; // unknown status → trust the scan's server filter
    const target = state.labelStatusFilter === 'printed' ? LABEL_STATUS_PRINTED : LABEL_STATUS_NOT_PRINTED;
    return rec.labelStatus === target;
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
  function selectAllVisible() {
    const tab = state.currentTab;
    if (tab === 'single' || tab === 'multi') {
      const idsKey = tab === 'single' ? 'fulfillUnitIdsSingle' : 'fulfillUnitIdsMulti';
      const type = tab === 'single' ? 'single_item' : 'single_sku';
      const scenario = tab === 'single' ? 'single' : 'multi';
      for (const p of state.products.values()) {
        const count = carrierFilteredSize(p[idsKey]);
        if (count === 0) continue;
        if (isDone(p.productId, null, type)) continue;
        state.selected.set(selectionKey({type:'product', productId:p.productId, scenario}),
                          {type:'product', productId:p.productId, scenario});
      }
    } else if (tab === 'weird') {
      for (const combo of state.weirdCombos.values()) {
        const count = carrierFilteredSize(combo.fulfillUnitIds);
        if (count === 0) continue;
        if (isComboDone(combo.sigKey)) continue;
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

    const confirm = await showPrintConfirm({
      title: multiSku ? `พิมพ์ ${skuBuckets.size} SKU` : `พิมพ์ ${items.length} ไฟล์`,
      summary: multiSku ? `${skuBuckets.size} SKU · ${totalIds} ฉลาก` : `แยก 1 ไฟล์ต่อตัวเลือก`,
      count: totalIds,
      sampleText: sample,
    });
    if (!confirm) return;

    const { workerId, workerName, workerIcon } = confirm;

    // §3.5/§4: Use showChunkPlanModal whenever multi-SKU (so user can choose combined vs split)
    // or for large batches. Default combined=false (split per SKU).
    let plan;
    if (!multiSku && totalIds <= CHUNK_PROMPT_THRESHOLD) {
      plan = { mode: 'single', withPickingList: loadPickingListPref(), combined: false, withDivider: false };
    } else {
      plan = await showChunkPlanModal({ total: totalIds, multiSku, defaultPickingList: loadPickingListPref() });
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
        const slices = planSlice(allSelectedIds, plan);
        const chunkCount = slices.length;
        const exportChunks = [];
        // §UX: show immediate progress so the UI isn't silent during PDF build.
        const prepProgress = showProgress(`กำลังเตรียม PDF รวม (${groups.length} SKU · ${totalIds} ฉลาก)`);
        try {
          for (let i = 0; i < slices.length; i++) {
            const slice = slices[i];
            const idx = i + 1;
            const chunkSuffix = chunkCount > 1 ? `-ชุด${idx}-${chunkCount}` : '';
            const filename = `${baseFilename}${chunkSuffix}.pdf`;
            const label = chunkCount === 1 ? 'ไฟล์เดียว' : `ชุด ${idx}/${chunkCount}`;
            const basePct = (i / chunkCount) * 100;
            prepProgress.update(basePct, `กำลังสร้างชุด ${idx}/${chunkCount}`);

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
              const { bytes } = await buildMultiSkuCombinedPdf(sliceGroups, workerName, workerIcon, plan.withPickingList, (pct) => {
                prepProgress.update(basePct + (pct / chunkCount), `ชุด ${idx}/${chunkCount} · ${pct.toFixed(0)}%`);
              }, plan.withDivider);
              exportChunks.push({ ids: slice, label, filename, prebuiltBytes: bytes });
            } catch (_buildErr) {
              exportChunks.push({ ids: slice, label, filename });
            }
          }
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

  function showPrintConfirm({ title, summary, count, sampleText }) {
    return new Promise(resolve => {
      const overlay = document.createElement('div');
      overlay.className = 'qf-modal-overlay';
      const overlayChecked = state.overlayEnabled ? 'checked' : '';
      const teams = state.teams || [];
      const hasAssignees = state.workers.length > 0 || teams.length > 0;
      const last = loadLastAssignee();
      const workerOpts = state.workers.length ? `<optgroup label="คน">${state.workers.map(w => `<option value="worker:${escapeHtml(w.id)}"${last === 'worker:' + w.id ? ' selected' : ''}>${escapeHtml(w.icon)}  ${escapeHtml(w.name)}</option>`).join('')}</optgroup>` : '';
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
            ${packerRowHtml}
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
        resolve({ ok: true, workerId: a.workerId || null, workerName: a.workerName || null, workerIcon: a.workerIcon || null, teamId: a.teamId || null, teamName: a.teamName || null });
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

    // Split ids into API batches (max PRINT_BATCH_SIZE each)
    const batches = [];
    for (let i = 0; i < ids.length; i += PRINT_BATCH_SIZE) {
      batches.push(ids.slice(i, i + PRINT_BATCH_SIZE));
    }
    const batchCount = batches.length;

    // Per-batch progress [0..1]; averaged into single bar
    const batchProgress = new Float64Array(batchCount);
    const reportProgress = () => {
      const avg = batchProgress.reduce((a, b) => a + b, 0) / batchCount;
      onProgress(avg, batchCount > 1 ? `ประมวลผล ${batchCount} ชุดพร้อมกัน...` : '');
    };

    // Fire all batches in parallel — generate + download + overlay run concurrently
    const batchResults = await Promise.all(batches.map(async (batch, bi) => {
      const setP = (frac) => { batchProgress[bi] = frac; reportProgress(); };
      setP(0.05);

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
        return null;
      }

      setP(0.25);
      const pdfBytes = await fetch(docUrl).then(r => r.arrayBuffer());
      setP(0.40);

      let modifiedBytes;
      if (state.overlayEnabled) {
        try {
          modifiedBytes = await overlayAliasOnPdf(pdfBytes, batch, (cur, totPages) => {
            batchProgress[bi] = 0.40 + 0.50 * (cur / totPages);
            reportProgress();
          }, workerName, workerIcon);
        } catch (e) {
          console.warn('[QF] overlay failed, using original:', e);
          modifiedBytes = pdfBytes;
        }
      } else {
        modifiedBytes = pdfBytes;
      }
      // Apply user-chosen PDF template ON TOP of (after) the alias watermark so
      // custom logos/text can optionally cover or sit alongside it.
      try {
        const _activeTpl = getActivePdfTemplate();
        if (_activeTpl && modifiedBytes) {
          modifiedBytes = await applyPdfTemplate(modifiedBytes, _activeTpl, buildTemplateRecordData(batch));
        }
      } catch (_tplErr) {
        console.warn('[QF] pdf template apply failed, using overlay-only bytes:', _tplErr);
      }

      setP(1.0);
      return modifiedBytes;
    }));

    // Merge pages in original batch order
    const mergedDoc = await PDFDocument.create();
    let pageCount = 0;
    for (const modifiedBytes of batchResults) {
      if (!modifiedBytes) continue;
      const partDoc = await PDFDocument.load(modifiedBytes);
      const indices = partDoc.getPageIndices();
      const copied = await mergedDoc.copyPages(partDoc, indices);
      copied.forEach(p => mergedDoc.addPage(p));
      pageCount += copied.length;
    }

    const bytes = await mergedDoc.save();
    return { bytes, pageCount };
  }

  // Sub-group ids by (productId:skuId) then by carrier so each SKU × carrier
  // combination gets its own divider. Sorted by alias → carrierName.
  function subGroupByCarrier(ids) {
    const bucketMap = new Map();
    for (const id of ids) {
      const rec = state.records.get(id);
      if (!rec?.skuList?.length) continue;
      const s = rec.skuList[0];
      const skuKey = `${s.productId || ''}:${s.skuId || ''}`;
      const carrierId = state.carrierOf.get(id) || 'unknown';
      const carrier = state.carriers.get(carrierId) || { name: 'ไม่ระบุ', iconUrl: '' };
      const key = `${skuKey}|${carrierId}`;
      if (!bucketMap.has(key)) {
        const aliasRaw = (getAlias(s.productId) || '').trim();
        const variantInfo = getVariantInfo(s.productId, s.skuId);
        bucketMap.set(key, {
          skuKey,
          carrierId,
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
      return a.carrierName.localeCompare(b.carrierName, 'th');
    });
  }

  // Prepend per-subgroup divider pages to an existing label PDF. Reorders by
  // (sku, carrier) so dividers always sit directly above their labels.
  // Returns {bytes, pageCount}. idOrder must match the order ids were sent to
  // the API so page index → id can be inferred.
  async function prependDividersToChunk(labelBytes, idOrder, workerName) {
    const { PDFDocument } = window.PDFLib;
    const subs = subGroupByCarrier(idOrder);
    if (subs.length === 0) return { bytes: labelBytes, pageCount: null };

    const labelDoc = await PDFDocument.load(labelBytes);
    const totalPages = labelDoc.getPageCount();
    if (totalPages === 0) return { bytes: labelBytes, pageCount: 0 };

    // idOrder is the order ids were sent to TikTok. Assume 1 id ↔ 1 page.
    // (TikTok returns 1 label per fulfill_unit_id.) Build id → pageIndex map.
    const pageByIdIdx = new Map();
    const limit = Math.min(idOrder.length, totalPages);
    for (let i = 0; i < limit; i++) pageByIdIdx.set(idOrder[i], i);

    const finalDoc = await PDFDocument.create();
    if (window.fontkit) finalDoc.registerFontkit(window.fontkit);
    const fontBytes = await ensureFontBytes();
    const font = await finalDoc.embedFont(fontBytes, { subset: true });

    const W = labelDoc.getPage(0).getWidth();
    const H = labelDoc.getPage(0).getHeight();

    for (const sub of subs) {
      await buildDividerPage(finalDoc, { W, H }, {
        alias: sub.alias,
        officialName: sub.officialName,
        variantName: sub.variantName,
        productImageURL: sub.productImageURL,
        carrierName: sub.carrierName,
        carrierIconURL: sub.carrierIconURL,
        qty: sub.ids.length,
      }, font, workerName);
      const pageIdx = sub.ids
        .map(id => pageByIdIdx.get(id))
        .filter(i => i != null);
      if (pageIdx.length) {
        const copied = await finalDoc.copyPages(labelDoc, pageIdx);
        copied.forEach(p => finalDoc.addPage(p));
      }
    }

    const bytes = await finalDoc.save();
    return { bytes, pageCount: finalDoc.getPageCount() };
  }

  // §4.5 / §4.6: Build a single divider page for the combined multi-SKU PDF.
  // payload = {alias, officialName, variantName, productImageURL, qty,
  //            carrierName?, carrierIconURL?}
  // Layout redesigned (Bug #2 fix) — clean top-down stack, no overlaps.
  // ทุกอย่างขาวดำ (grayscale only).
  async function buildDividerPage(pdfDoc, { W, H }, payload, font, workerName) {
    const { rgb } = window.PDFLib;
    const page = pdfDoc.addPage([W, H]);
    // Phase 2: use per-field config (with preset fallback). Resolve size mult per field.
    const eff = getEffectiveDividerConfig();
    const f = eff.fields;
    const mult = (k) => DIVIDER_SIZE_MAP[f[k]?.size] || 1.0;
    const vis = (k) => !!f[k]?.visible;
    // Map new config → existing cfg.* flags expected by downstream code (minimal churn).
    const cfg = {
      showAlias:   vis('alias'),
      showName:    vis('name'),
      showVariant: vis('variant'),
      showCarrier: vis('carrier'),
      showImage:   vis('image'),
      showQty:     vis('qty'),
      showWorker:  vis('worker'),
      showFooter:  vis('footer'),
      aliasScale:  mult('alias'),
      imageFirst:  !!eff.imageFirst,
    };
    const BLACK = rgb(0, 0, 0);
    const DARK = rgb(0.15, 0.15, 0.15);
    const MID = rgb(0.35, 0.35, 0.35);
    const LIGHT = rgb(0.55, 0.55, 0.55);

    // --- Banners (top + bottom) — ขาวดำ ---
    const BANNER_H = 20;
    page.drawRectangle({ x: 0, y: H - BANNER_H, width: W, height: BANNER_H, color: BLACK });
    page.drawRectangle({ x: 0, y: 0, width: W, height: BANNER_H, color: BLACK });
    const bannerText = '— แบ่งกลุ่ม SKU —';
    const bannerSize = 11;
    const bannerW = font.widthOfTextAtSize(bannerText, bannerSize);
    page.drawText(bannerText, {
      x: (W - bannerW) / 2,
      y: H - BANNER_H + (BANNER_H - bannerSize) / 2 + 1,
      size: bannerSize,
      font,
      color: rgb(1, 1, 1),
    });

    // --- Top area cursor (below top banner) ---
    // Layout: top-down stack with vertical cursor.
    const PAD_X = 12;
    const contentMaxW = W - PAD_X * 2;
    let topCursor = H - BANNER_H - 6; // y of next drawable top edge
    const imgMult = mult('image');

    // Worker/team line at top (text in grey, below banner).
    if (cfg.showWorker && workerName) {
      const wSize = 9 * mult('worker');
      topCursor -= wSize + 2;
      page.drawText(`ผู้แพ็ค: ${workerName}`, {
        x: PAD_X,
        y: topCursor,
        size: wSize,
        font,
        color: MID,
        maxWidth: contentMaxW,
      });
      topCursor -= 4;
    }

    // --- photo-first preset renders image FIRST, at top ---
    // For other presets image is placed between carrier and qty.
    const IMG_MAX = 70 * imgMult; // max side for product image (shrunk from 120 to leave room)
    let topImageBottom = null;
    if (cfg.imageFirst && cfg.showImage && payload.productImageURL) {
      try {
        const imgResp = await _origFetch.call(window, payload.productImageURL, { credentials: 'omit' });
        if (imgResp.ok) {
          const imgBuf = await imgResp.arrayBuffer();
          const ct = imgResp.headers.get('content-type') || '';
          const isPng = ct.includes('png') || payload.productImageURL.toLowerCase().includes('.png');
          const image = isPng ? await pdfDoc.embedPng(imgBuf) : await pdfDoc.embedJpg(imgBuf);
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

    // --- Alias block (always the biggest element when shown) ---
    if (cfg.showAlias && payload.alias) {
      const baseSize = Math.min(H * 0.08, 36) * (cfg.aliasScale || 1);
      const aliasW0 = font.widthOfTextAtSize(payload.alias, baseSize);
      const scale = aliasW0 > contentMaxW ? contentMaxW / aliasW0 : 1;
      const aliasSize = baseSize * scale;
      const aliasW = font.widthOfTextAtSize(payload.alias, aliasSize);
      topCursor -= aliasSize + 4;
      page.drawText(payload.alias, {
        x: (W - aliasW) / 2,
        y: topCursor,
        size: aliasSize,
        font,
        color: BLACK,
      });
      topCursor -= 4;
    }

    // --- Official name (truncate, 10pt) ---
    if (cfg.showName && payload.officialName) {
      const nameSize = 10 * mult('name');
      let nameText = payload.officialName;
      while (nameText.length > 2 && font.widthOfTextAtSize(nameText, nameSize) > contentMaxW) {
        nameText = nameText.slice(0, -1);
      }
      if (nameText !== payload.officialName) nameText = nameText.slice(0, -1) + '…';
      const nameW = font.widthOfTextAtSize(nameText, nameSize);
      topCursor -= nameSize + 2;
      page.drawText(nameText, {
        x: (W - nameW) / 2,
        y: topCursor,
        size: nameSize,
        font,
        color: DARK,
      });
    }

    // --- Variant line (9pt grey) ---
    if (cfg.showVariant && payload.variantName) {
      const varSize = 9 * mult('variant');
      let varText = payload.variantName;
      while (varText.length > 2 && font.widthOfTextAtSize(varText, varSize) > contentMaxW) {
        varText = varText.slice(0, -1);
      }
      if (varText !== payload.variantName) varText = varText.slice(0, -1) + '…';
      const varW = font.widthOfTextAtSize(varText, varSize);
      topCursor -= varSize + 2;
      page.drawText(varText, {
        x: (W - varW) / 2,
        y: topCursor,
        size: varSize,
        font,
        color: MID,
      });
      topCursor -= 4;
    }

    // --- Carrier badge (outlined box, 22pt tall) ---
    if (cfg.showCarrier && payload.carrierName) {
      const carrierSize = 11 * mult('carrier');
      const carrierLabel = `ขนส่ง: ${payload.carrierName}`;
      const labelW = font.widthOfTextAtSize(carrierLabel, carrierSize);
      const iconSize = 16;
      let iconImg = null;
      if (payload.carrierIconURL) {
        try {
          const iconResp = await _origFetch.call(window, payload.carrierIconURL, { credentials: 'omit' });
          if (iconResp.ok) {
            const iconBuf = await iconResp.arrayBuffer();
            const ct = iconResp.headers.get('content-type') || '';
            const isPng = ct.includes('png') || payload.carrierIconURL.toLowerCase().includes('.png');
            iconImg = isPng ? await pdfDoc.embedPng(iconBuf) : await pdfDoc.embedJpg(iconBuf);
          }
        } catch { iconImg = null; }
      }
      const gap = iconImg ? 6 : 0;
      const iconW = iconImg ? iconSize : 0;
      const rowW = iconW + gap + labelW;
      const padX = 10;
      const boxW = rowW + padX * 2;
      const boxH = 22;
      const boxX = (W - boxW) / 2;
      topCursor -= boxH + 6;
      const boxY = topCursor;
      page.drawRectangle({
        x: boxX, y: boxY, width: boxW, height: boxH,
        borderColor: BLACK, borderWidth: 1,
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

    // --- Bottom-up layout: footer (y=18) → qty (y=42) → image above qty ---
    const FOOTER_Y = BANNER_H + 8; // above bottom banner
    const QTY_Y = FOOTER_Y + 22;

    // Footer
    if (cfg.showFooter) {
      const footerText = 'กรุณาตรวจสอบจำนวนสินค้าก่อนแพ็คสินค้า';
      const footerSize = 8 * mult('footer');
      const footerW = font.widthOfTextAtSize(footerText, footerSize);
      page.drawText(footerText, {
        x: (W - footerW) / 2,
        y: FOOTER_Y,
        size: footerSize,
        font,
        color: LIGHT,
      });
    }

    // Qty (always shown unless preset hides)
    if (cfg.showQty) {
      const qtyText = `จำนวนทั้งหมด: ${payload.qty} ใบ`;
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
    if (!cfg.imageFirst && cfg.showImage && payload.productImageURL) {
      const availableTop = topCursor - 6;    // top of available vertical space
      const availableBot = QTY_Y + 18;       // leave room above qty
      const availH = availableTop - availableBot;
      if (availH > 30) {
        try {
          const imgResp = await _origFetch.call(window, payload.productImageURL, { credentials: 'omit' });
          if (imgResp.ok) {
            const imgBuf = await imgResp.arrayBuffer();
            const ct = imgResp.headers.get('content-type') || '';
            const isPng = ct.includes('png') || payload.productImageURL.toLowerCase().includes('.png');
            const image = isPng ? await pdfDoc.embedPng(imgBuf) : await pdfDoc.embedJpg(imgBuf);
            const maxSide = Math.min(IMG_MAX, availH, contentMaxW);
            const ratio = image.width / image.height;
            const iw = ratio >= 1 ? maxSide : maxSide * ratio;
            const ih = ratio >= 1 ? maxSide / ratio : maxSide;
            const iy = availableBot + (availH - ih) / 2;
            page.drawImage(image, { x: (W - iw) / 2, y: iy, width: iw, height: ih });
          }
        } catch { /* skip */ }
      }
    }

    return page;
  }

  // §7: Build picking list pages for prepending to a label PDF.
  // Returns array of PDFPage objects already added to a temporary PDFDocument.
  // Caller uses finalDoc.copyPages(tmpDoc, tmpDoc.getPageIndices()) to embed them.
  async function buildPickingListPages(groups, { W, H }, headerMeta) {
    if (!window.PDFLib || !window.fontkit) return [];
    const { PDFDocument, rgb } = window.PDFLib;

    // Aggregate rows from groups — one row per productId:skuId bucket.
    const aggMap = new Map();
    for (const grp of groups) {
      const key = `${grp.productId || ''}:${grp.skuId || ''}`;
      if (!aggMap.has(key)) {
        const variantInfo = getVariantInfo(grp.productId, grp.skuId);
        aggMap.set(key, {
          alias: grp.alias || '',
          officialName: grp.officialName || '',
          skuAlias: (variantInfo?.alias || '').trim(),
          sellerSku: grp.sellerSku || '',
          productImageURL: grp.productImageURL || null,
          qty: 0,
          orderIds: new Set(),
        });
      }
      const agg = aggMap.get(key);
      // qty = label count in this group (each id = 1 fulfillUnit = 1 order)
      agg.qty += grp.ids.length;
      for (const id of grp.ids) {
        const rec = state.records.get(id);
        const oid = rec?.orderIds?.[0];
        if (oid) agg.orderIds.add(oid);
      }
    }

    let no = 0;
    const rows = [...aggMap.values()]
      .sort((a, b) => a.alias.localeCompare(b.alias) || a.officialName.localeCompare(b.officialName))
      .map(r => ({ ...r, no: ++no, orderIds: [...r.orderIds] }));

    const tmpDoc = await PDFDocument.create();
    if (window.fontkit) tmpDoc.registerFontkit(window.fontkit);
    const fontBytes = await ensureFontBytes();
    const font = await tmpDoc.embedFont(fontBytes, { subset: true });

    const isCompact = W < 420;
    const headerHeight = 120;
    // Bug #3 fix: smaller rows so more fit per page (tabular 38, compact 46).
    const rowHeight = isCompact ? 46 : 38;
    const footerHeight = 30;
    const rowsPerPage = Math.max(1, Math.floor((H - headerHeight - footerHeight) / rowHeight));
    const pageCount = Math.ceil(rows.length / rowsPerPage);

    // Pre-embed images for tabular layout (one fetch per group, best-effort).
    const imageCache = new Map(); // key → embedded image or null
    if (!isCompact) {
      await Promise.all(rows.map(async (row) => {
        if (!row.productImageURL || imageCache.has(row.productImageURL)) return;
        try {
          const resp = await _origFetch.call(window, row.productImageURL, { credentials: 'omit' });
          if (!resp.ok) { imageCache.set(row.productImageURL, null); return; }
          const buf = await resp.arrayBuffer();
          const ct = resp.headers.get('content-type') || '';
          const isPng = ct.includes('png') || row.productImageURL.toLowerCase().includes('.png');
          const img = isPng ? await tmpDoc.embedPng(buf) : await tmpDoc.embedJpg(buf);
          imageCache.set(row.productImageURL, img);
        } catch { imageCache.set(row.productImageURL, null); }
      }));
    }

    for (let pi = 0; pi < pageCount; pi++) {
      const page = tmpDoc.addPage([W, H]);
      renderPickingHeader(page, headerMeta, font, W, H, pi + 1, pageCount);

      const pageRows = rows.slice(pi * rowsPerPage, (pi + 1) * rowsPerPage);
      let curY = H - headerHeight;
      for (const row of pageRows) {
        if (isCompact) {
          renderPickingRowCompact(page, row, curY, W, font);
        } else {
          renderPickingRowTabular(page, row, curY, W, font, imageCache);
        }
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
    if (headerMeta.assigneeName) {
      page.drawText(`รับผิดชอบ: ${headerMeta.assigneeName}`, { x: 24, y: H - 98, size: 10, font, color: black, maxWidth: W - 48 });
    }

    // Separator line
    page.drawLine({ start: { x: 12, y: H - 108 }, end: { x: W - 12, y: H - 108 }, thickness: 0.5, color: grey });

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

  function renderPickingRowTabular(page, row, y, W, font, imageCache) {
    const { rgb } = window.PDFLib;
    const black = rgb(0, 0, 0);
    const grey = rgb(0.5, 0.5, 0.5);
    const rowH = 38; // Bug #3 fix: shrunk from 50 to fit more rows
    const pad = 24;
    // Widths: no=20, img=32, name=flex, sku=46, sellerSku=56, qty=32, orderId=74
    const fixedW = 260;
    const nameW = Math.max(40, W - pad * 2 - fixedW);
    let x = pad;

    // Row number
    page.drawText(String(row.no), { x, y: y - 12, size: 8, font, color: grey });
    x += 20;

    // Image (30×30)
    const img = row.productImageURL ? (imageCache.get(row.productImageURL) || null) : null;
    if (img) {
      const boxSize = 30;
      const ratio = img.width / img.height;
      const iw = ratio >= 1 ? boxSize : boxSize * ratio;
      const ih = ratio >= 1 ? boxSize / ratio : boxSize;
      page.drawImage(img, { x, y: y - rowH + (rowH - ih) / 2, width: iw, height: ih });
    }
    x += 32;

    // Alias (11pt) — single line truncated.
    const aliasSize = 11;
    const aliasRaw = row.alias || row.officialName.slice(0, 12);
    let aliasText = aliasRaw;
    while (aliasText.length > 1 && font.widthOfTextAtSize(aliasText, aliasSize) > nameW) {
      aliasText = aliasText.slice(0, -1);
    }
    page.drawText(aliasText, { x, y: y - 12, size: aliasSize, font, color: black });

    // Official name (8pt) — word-wrap up to 2 lines.
    const nameSize = 8;
    const nameLines = wrapTextToLines(row.officialName || '', 2, font, nameSize, nameW);
    nameLines.forEach((line, i) => {
      page.drawText(line, { x, y: y - 22 - i * (nameSize + 1), size: nameSize, font, color: grey });
    });
    x += nameW;

    // SKU alias
    const skuText = row.skuAlias || '-';
    page.drawText(skuText.slice(0, 10), { x, y: y - 14, size: 8, font, color: grey, maxWidth: 44 });
    x += 46;

    // Seller SKU
    const ssText = row.sellerSku || '-';
    page.drawText(ssText.slice(0, 12), { x, y: y - 14, size: 8, font, color: grey, maxWidth: 54 });
    x += 56;

    // Qty
    page.drawText(String(row.qty), { x, y: y - 14, size: 11, font, color: black });
    x += 32;

    // Order IDs (up to 3, +N more)
    const oids = row.orderIds.slice(0, 3);
    const more = row.orderIds.length - oids.length;
    const oidText = oids.join(', ') + (more > 0 ? ` +${more}` : '');
    page.drawText(oidText, { x, y: y - 14, size: 7, font, color: grey, maxWidth: 72 });

    // Row separator
    page.drawLine({ start: { x: pad, y: y - rowH }, end: { x: W - pad, y: y - rowH }, thickness: 0.25, color: rgb(0.85, 0.85, 0.85) });
  }

  function renderPickingRowCompact(page, row, y, W, font) {
    const { rgb } = window.PDFLib;
    const black = rgb(0, 0, 0);
    const grey = rgb(0.5, 0.5, 0.5);
    const rowH = 46; // Bug #3 fix: shrunk from 60
    const pad = 12;
    const imgSize = 34;

    // Row number + alias (11pt, right of image)
    const textX = pad + imgSize + 6;
    const textW = W - textX - pad;
    const aliasSize = 11;
    const alias = row.alias || row.officialName.slice(0, 16);
    let aliasText = `${row.no}. ${alias}`;
    while (aliasText.length > 1 && font.widthOfTextAtSize(aliasText, aliasSize) > textW) {
      aliasText = aliasText.slice(0, -1);
    }
    page.drawText(aliasText, { x: textX, y: y - 12, size: aliasSize, font, color: black });

    // Official name (8pt) — word-wrap up to 2 lines.
    const ns = 8;
    const nameLines = wrapTextToLines(row.officialName || '', 2, font, ns, textW);
    nameLines.forEach((line, i) => {
      page.drawText(line, { x: textX, y: y - 22 - i * (ns + 1), size: ns, font, color: grey });
    });

    // Bottom line: Qty + SKU + order count
    const bottomY = y - rowH + 6;
    const bottomParts = [`Qty: ${row.qty}`];
    if (row.skuAlias) bottomParts.push(`SKU: ${row.skuAlias}`);
    if (row.orderIds.length > 0) bottomParts.push(`Orders: ${row.orderIds.length}`);
    page.drawText(bottomParts.join('  /  '), { x: textX, y: bottomY, size: 8, font, color: grey, maxWidth: textW });

    // Row separator
    page.drawLine({ start: { x: pad, y: y - rowH }, end: { x: W - pad, y: y - rowH }, thickness: 0.25, color: rgb(0.85, 0.85, 0.85) });
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
    };
  }

  // §4.7: Build a combined multi-SKU PDF: [divider_1][labels_1][divider_2][labels_2]…
  // groups[] shape: [{productId, skuId, alias, officialName, variantName, productImageURL, ids: string[]}]
  //   sorted by alias ascending by caller.
  async function buildMultiSkuCombinedPdf(groups, workerName, workerIcon, withPickingList, onProgress, withDivider = true) {
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
    try {
      firstGroupBytes = await callGenerateApiRaw(groups[0].ids);
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

      // When divider is on, sub-group by carrier so each (SKU × carrier)
      // combination gets its own divider + block of labels — กันแพ็คผิด
      // เมื่อ SKU เดียวกันมีหลายขนส่ง.
      const subs = withDivider ? subGroupByCarrier(grp.ids) : null;
      const parts = (subs && subs.length > 0)
        ? subs.map(s => ({
            ids: s.ids,
            alias: s.alias || grp.alias,
            officialName: s.officialName || grp.officialName,
            variantName: s.variantName || grp.variantName,
            productImageURL: s.productImageURL || grp.productImageURL,
            carrierName: s.carrierName,
            carrierIconURL: s.carrierIconURL,
          }))
        : [{
            ids: grp.ids,
            alias: grp.alias,
            officialName: grp.officialName,
            variantName: grp.variantName,
            productImageURL: grp.productImageURL,
            carrierName: null,
            carrierIconURL: null,
          }];

      for (let pi = 0; pi < parts.length; pi++) {
        const part = parts[pi];
        // Fetch label bytes (reuse first-group probe result only for the
        // first sub-part of the first group when it matches grp.ids).
        let labelsBytes;
        try {
          if (gi === 0 && pi === 0 && firstGroupBytes && part.ids.length === grp.ids.length) {
            labelsBytes = firstGroupBytes;
          } else {
            labelsBytes = await callGenerateApiRaw(part.ids);
          }
        } catch (e) {
          throw new Error(`กลุ่ม "${part.alias}" (${gi + 1}/${groups.length}) ดึง labels ไม่ได้: ${e.message}`);
        }

        // Apply alias overlay if enabled.
        if (state.overlayEnabled && labelsBytes) {
          try {
            labelsBytes = await overlayAliasOnPdf(labelsBytes, part.ids, () => {}, workerName, workerIcon);
          } catch (_oErr) { /* continue unmodified */ }
        }
        // Apply active PDF template (user's custom header/logo/etc).
        if (labelsBytes) {
          try {
            const _activeTpl = getActivePdfTemplate();
            if (_activeTpl) {
              labelsBytes = await applyPdfTemplate(labelsBytes, _activeTpl, buildTemplateRecordData(part.ids));
            }
          } catch (_tErr) { /* continue unmodified */ }
        }

        if (withDivider) {
          await buildDividerPage(finalDoc, { W, H }, {
            alias: part.alias,
            officialName: part.officialName,
            variantName: part.variantName,
            productImageURL: part.productImageURL,
            carrierName: part.carrierName,
            carrierIconURL: part.carrierIconURL,
            qty: part.ids.length,
          }, font, workerName);
        }

        if (labelsBytes) {
          try {
            const partDoc = await PDFDocument.load(labelsBytes);
            const copied = await finalDoc.copyPages(partDoc, partDoc.getPageIndices());
            copied.forEach(p => finalDoc.addPage(p));
          } catch (_copyErr) { /* skip malformed */ }
        }
      }

      if (onProgress) onProgress((gi + 1) / groups.length);
    }

    const bytes = await finalDoc.save();
    return { bytes, pageCount: finalDoc.getPageCount() };
  }

  // Helper: call TikTok's generate API for a list of fulfillUnitIds and return raw PDF bytes.
  // Splits into batches of PRINT_BATCH_SIZE, merges pages, returns combined bytes.
  async function callGenerateApiRaw(ids) {
    const { PDFDocument } = window.PDFLib;
    const batches = [];
    for (let i = 0; i < ids.length; i += PRINT_BATCH_SIZE) {
      batches.push(ids.slice(i, i + PRINT_BATCH_SIZE));
    }
    const batchDocs = await Promise.all(batches.map(async (batch) => {
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
      return pdfResp.arrayBuffer();
    }));

    if (batchDocs.length === 1) return batchDocs[0];
    const merged = await PDFDocument.create();
    for (const bytes of batchDocs) {
      const part = await PDFDocument.load(bytes);
      const copied = await merged.copyPages(part, part.getPageIndices());
      copied.forEach(p => merged.addPage(p));
    }
    return merged.save();
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

    const confirm = await showPrintConfirm({
      title: displayLabel || 'พิมพ์ฉลาก',
      count: ids.length,
      sampleText,
    });
    if (!confirm) return false;
    const { workerId, workerName, workerIcon } = confirm;

    // §3.5: Use showChunkPlanModal when multi-SKU OR above threshold.
    let plan;
    // Detect multi-SKU: count distinct productId:skuId buckets across ALL skus in every record.
    // A single record with >1 skus (weird combo) is already multi-SKU on its own.
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
    if (!multiSku && ids.length <= CHUNK_PROMPT_THRESHOLD) {
      plan = { mode: 'single', withPickingList: loadPickingListPref(), combined: false, withDivider: false };
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

    const runChunk = async (ci) => {
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
          ({ bytes, pageCount } = await buildChunkPdf(chunks[ci].ids, (pct, label) => {
            result.updateChunkProgress(ci, pct, label);
          }, workerName, workerIcon));

          // §Divider: prepend per-subgroup dividers to each split chunk.
          // Runs BEFORE picking list so picking list stays at the very top.
          if (withDivider && window.PDFLib) {
            try {
              const { bytes: dBytes, pageCount: dCount } = await prependDividersToChunk(bytes, chunks[ci].ids, workerName);
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
        result.completeChunk(ci, {url, pageCount});
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

    // Toggle history vs daily option panels
    overlay.querySelectorAll('input[name="qf-csv-type"]').forEach(radio => {
      radio.addEventListener('change', () => {
        const isHistory = overlay.querySelector('input[name="qf-csv-type"]:checked').value === 'history';
        overlay.querySelector('#qf-csv-history-opts').style.display = isHistory ? '' : 'none';
        overlay.querySelector('#qf-csv-daily-opts').style.display = isHistory ? 'none' : '';
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
      } else {
        const range = overlay.querySelector('input[name="qf-csv-daily-range"]:checked').value;
        let cutoff = 0;
        if (range === 'today') cutoff = dayStart(now);
        else if (range === '7d') cutoff = dayStart(now) - 6 * 24 * 60 * 60 * 1000;
        const entries = loadHistory().filter(e => e.timestamp >= cutoff);
        if (entries.length === 0) { showToast('ไม่มีข้อมูลในช่วงที่เลือก', 2000); return; }
        downloadCsv(buildDailyCsv(entries), `quickfilter-daily-${todayStr}.csv`);
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
          <div id="qf-settings-wrap" style="position:relative;">
            <button id="qf-settings-btn" title="ตั้งค่า">⋮</button>
            <div id="qf-settings-menu" class="qf-settings-menu" style="display:none;">
              <button id="qf-menu-csv">ดาวน์โหลดประวัติ CSV</button>
              <button id="qf-menu-plan">🎨 วางแผน</button>
              <button id="qf-menu-divider-preset">🗂 ใบคั่น - รูปแบบ</button>
              <button id="qf-menu-pdf-templates">📄 เทมเพลต PDF</button>
              <button id="qf-menu-pdf-template-new">✏️ สร้างเทมเพลต</button>
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
        openPdfTemplateManager();
      });
    }
    const pdfTplNewBtn = document.getElementById('qf-menu-pdf-template-new');
    if (pdfTplNewBtn) {
      pdfTplNewBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        settingsMenu.style.display = 'none';
        const tpls = loadPdfTemplates();
        if (tpls.length >= MAX_PDF_TEMPLATES) {
          showToast(`ครบ ${MAX_PDF_TEMPLATES} เทมเพลตแล้ว — ลบก่อนจึงสร้างใหม่ได้`, 3000);
          openPdfTemplateManager();
          return;
        }
        openPdfTemplateEditor();
      });
    }
    const dividerPresetBtn = document.getElementById('qf-menu-divider-preset');
    if (dividerPresetBtn) {
      dividerPresetBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        settingsMenu.style.display = 'none';
        openDividerPresetModal();
      });
    }
    renderHistoryBadge();
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
        const cardDone = labels
          ? isLabelsDone(p[idsKey])
          : isDone(p.productId, null, type);
        const productItem = {type: 'product', productId: p.productId, scenario};
        const selected = state.selectMode && isSelected(productItem);
        card.className = 'qf-product-card'
          + (cardDone ? ' qf-done' : '')
          + (state.selectMode ? ' qf-select-mode' : '')
          + (selected ? ' qf-selected' : '');
        card.title = cardDone
          ? `${p.productName}\n\nพิมพ์แล้ว — คลิกเพื่อเลือก: พิมพ์ซ้ำ หรือ ลบเครื่องหมาย ✓`
          : p.productName;
        const variantsRaw = [...p.variants.values()].map(v => ({v, c: variantCount(v)}));
        // Skip the synthetic "no-variant" entry (skuId=null) that Shopee
        // items without model_id produce — it's product-level, not a real
        // variant, and rendering a badge for it would show "null".
        const variants = variantsRaw.filter(x => x.c > 0 && x.v.skuId != null);
        const hasBadges = variants.length >= 1;
        const aliasVal = labels ? (getAlias(p.productId) || '') : '';
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
          for (const {v, c} of variants) {
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
            badge.textContent = `${displayName} (${c})`;
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
          const qtyBuckets = [...p.fulfillUnitIdsByQty.entries()]
            .map(([qty, idSet]) => ({ qty, ids: applyCarrierFilter([...idSet]) }))
            .filter(b => b.ids.length > 0)
            .sort((a, b) => a.qty - b.qty);
          if (qtyBuckets.length >= 1) {
            const chipsRow = document.createElement('div');
            chipsRow.className = 'qf-qty-chips';
            for (const b of qtyBuckets) {
              const allDone = b.ids.every(id =>
                state.printedUnitIds.has(id) || state.records.get(id)?.labelStatus === LABEL_STATUS_PRINTED
              );
              const chip = document.createElement('span');
              chip.className = 'qf-qty-chip' + (allDone ? ' qf-qty-chip--done' : '');
              chip.dataset.qty = b.qty;
              chip.textContent = (allDone ? '\u2713 ' : '') + '\u00d7' + b.qty + ' (' + b.ids.length + ')';
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
    if (already) {
      // Update in place — session may have changed columns since last save today.
      const updated = existing.map(s => {
        if (s.sessionId !== session.sessionId || s.date !== date) return s;
        const cols = {};
        for (const [cid, col] of Object.entries(session.columns || {})) {
          cols[cid] = { assigneeKind: col.kind, assigneeName: col.teamName || col.workerName || null, ids: [...(col.fulfillUnitIds || [])] };
        }
        return { ...s, columns: cols, savedAt: Date.now() };
      });
      savePlanSnapshots(updated);
    } else {
      const cols = {};
      for (const [cid, col] of Object.entries(session.columns || {})) {
        cols[cid] = { assigneeKind: col.kind, assigneeName: col.teamName || col.workerName || null, ids: [...(col.fulfillUnitIds || [])] };
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
        // v1 → v2: add kind + columnId
        if (!migrated.kind) {
          migrated = { ...migrated, kind: 'worker', columnId: wid, workerId: wid };
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
    const allIds = [...state.records.keys()];
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
      const overlay = document.createElement('div');
      overlay.className = 'qf-modal-overlay';
      overlay.innerHTML = `
        <div class="qf-modal qf-plan-chunk-popup" role="dialog">
          <div class="qf-plan-chunk-body">
            <div class="qf-plan-chunk-label">กี่ใบให้ ${escapeHtml(workerName)}?</div>
            <div class="qf-plan-chunk-row">
              <input type="number" class="qf-plan-chunk-input" min="1" max="${totalAvailable}" value="${totalAvailable}"/>
              <button class="qf-plan-chunk-all-btn">ทั้งหมด ${totalAvailable}</button>
            </div>
          </div>
          <div class="qf-modal-actions">
            <button class="qf-btn-cancel">ยกเลิก</button>
            <button class="qf-btn-confirm">ตกลง</button>
          </div>
        </div>
      `;
      document.body.appendChild(overlay);
      const inp = overlay.querySelector('.qf-plan-chunk-input');
      const cleanup = v => { overlay.remove(); resolve(v); };
      overlay.querySelector('.qf-plan-chunk-all-btn').onclick = () => { inp.value = totalAvailable; };
      overlay.querySelector('.qf-btn-cancel').onclick = () => cleanup(null);
      overlay.querySelector('.qf-btn-confirm').onclick = () => {
        const n = Math.max(1, Math.min(totalAvailable, parseInt(inp.value) || totalAvailable));
        cleanup(n);
      };
      overlay.onclick = e => { if (e.target === overlay) cleanup(null); };
      const onKey = e => { if (e.key === 'Escape') { cleanup(null); document.removeEventListener('keydown', onKey); } };
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
      if (!multiSku && printIds_.length <= CHUNK_PROMPT_THRESHOLD) {
        plan = { mode: 'single', withPickingList: loadPickingListPref(), combined: false, withDivider: false };
      } else {
        plan = await showChunkPlanModal({ total: printIds_.length, multiSku, defaultPickingList: loadPickingListPref() });
        if (!plan) {
          // User cancelled — restore column to pending.
          renderFn({ ...newSession, columns: { ...newSession.columns, [workerId]: { ...col, status: 'pending' } } });
          return;
        }
      }

      // §8.1: Column name for filename: teamName (if team kind) or workerName.
      const assigneeName = col.teamName || col.workerName || null;
      const assigneeKind = col.kind === 'team' ? 'team' : (assigneeName ? 'worker' : null);

      // §4.8: If multi-SKU and combined mode, build combined PDF per chunk.
      let ok;
      if (multiSku && plan.combined) {
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
          chunk.withDivider
        );
        const blob = new Blob([bytes], { type: 'application/pdf' });
        const url = URL.createObjectURL(blob);
        return { ...chunk, prebuiltUrl: url };
      } catch (e) {
        return { ...chunk, prebuiltUrl: null, buildError: e };
      }
    }));

    // Fall back to flat runChunkedExport if any combined build failed.
    const allBuilt = exportChunks.every(c => c.prebuiltUrl);
    if (!allBuilt) {
      return printPlanColumnChunked(printIds_, plan, col, assigneeName, assigneeKind, newSession, newCol, workerId, renderFn);
    }

    // Use runChunkedExport with prebuilt chunks (each chunk is already a PDF blob URL).
    const runChunks = exportChunks.map(c => ({ ids: c.ids, label: c.label, filename: c.filename }));
    return runChunkedExport(runChunks, displayTitle, {
      baseFilename,
      totalLabels: printIds_.length,
      workerId: col.workerId,
      workerName: workerNameForPdf,
      workerIcon: workerIconForPdf,
      assigneeKind,
      assigneeName,
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
      withPickingList: plan.withPickingList || false,
      withDivider: plan.withDivider || false,
    });
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

      const badge = document.createElement('span');
      badge.className = 'qf-plan-zone-badge';
      badge.textContent = isTeam ? '👥' : (col.workerIcon || '●');
      headerRow.appendChild(badge);

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
            const ok = await confirmInline(`ลบทีม "${tm.name}"?`, 'ลบ');
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

      // Icon picker (worker only)
      const iconPicker = document.createElement('div');
      iconPicker.className = 'qf-zone-add-icons';
      let selectedIcon = editingWorker ? editingWorker.icon : WORKER_ICONS[0];
      for (const ic of WORKER_ICONS) {
        const btn = document.createElement('button');
        btn.className = 'qf-workers-icon-btn' + (ic === selectedIcon ? ' active' : '');
        btn.dataset.icon = ic;
        btn.textContent = ic;
        btn.type = 'button';
        btn.onclick = () => {
          selectedIcon = ic;
          iconPicker.querySelectorAll('.qf-workers-icon-btn').forEach(b => b.classList.toggle('active', b.dataset.icon === ic));
        };
        iconPicker.appendChild(btn);
      }

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
        lbl.appendChild(document.createTextNode(`${w.icon} ${w.name}`));
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

      // Unassigned zone
      const unassignedZone = document.createElement('div');
      unassignedZone.className = 'qf-plan-zone qf-zone-unassigned';
      const unassignedHeader = document.createElement('div');
      unassignedHeader.className = 'qf-plan-zone-header';
      const unassignedTitle = document.createElement('span');
      unassignedTitle.className = 'qf-plan-zone-title';
      unassignedTitle.textContent = `📦 ยังไม่มอบหมาย (${session.unassignedIds.length} ใบ)`;
      unassignedHeader.appendChild(unassignedTitle);
      unassignedZone.appendChild(unassignedHeader);
      const unassignedScroll = document.createElement('div');
      unassignedScroll.className = 'qf-plan-zone-scroll qf-plan-cards';
      unassignedScroll.id = 'qf-plan-cards-unassigned';
      unassignedScroll.dataset.zone = 'unassigned';
      unassignedScroll.dataset.locked = '0';
      unassignedZone.appendChild(unassignedScroll);
      body.appendChild(unassignedZone);
      renderPlanCardsInto(unassignedScroll, session.unassignedIds, allSessionCards);

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
      const breakdown = nonDone.map(([, c]) => `${c.teamName || c.workerName || c.columnId}: ${c.fulfillUnitIds.length} ใบ`).join(', ');
      const ok = await confirmInline(`พิมพ์ทั้งหมด? ${breakdown}`, 'พิมพ์');
      if (!ok) return;
      let cur = s;
      for (const [cid] of nonDone) {
        const col = cur.columns[cid];
        if (!col) continue;
        printPlanColumn(cur, cid, col.fulfillUnitIds, ns => { render(ns); cur = ns; });
      }
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
              const ok = await confirmInline(`พิมพ์ของ ${colDisplayName} ${ids.length} ใบ?`, 'พิมพ์');
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
            const updatedCol = { ...col, status: 'pending', printedIds: [], failedIds: [], errorMsg: null, retryCount: 0 };
            const next = { ...s, columns: { ...s.columns, [cid]: updatedCol } };
            debouncedSavePlan(next);
            rerender(next);
          } else if (action === 'removecol') {
            if (!col) return;
            const ok = await confirmInline(`ลบ ${colDisplayName}? (ids จะคืนยังไม่มอบหมาย)`, 'ลบ');
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
      root.querySelectorAll('.qf-plan-card[draggable]').forEach(card => {
        card.addEventListener('dragstart', e => {
          _dragCardKey = card.dataset.cardKey;
          _dragSourceZone = card.closest('[data-zone]')?.dataset.zone;
          card.classList.add('qf-dragging');
          e.dataTransfer.effectAllowed = 'move';
        });
        card.addEventListener('dragend', () => {
          card.classList.remove('qf-dragging');
          root.querySelectorAll('.qf-plan-cards').forEach(el => el.classList.remove('qf-drag-over'));
        });
      });

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
