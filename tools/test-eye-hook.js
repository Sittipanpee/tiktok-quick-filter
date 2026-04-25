// ========================================================================
// TEST SCRIPT — Eye-Click Hook + Rate Limit Discovery
//
// เปิดหน้า order detail ที่มี 👁 eye icon เช่น:
//   https://seller-th.tiktok.com/order/detail?order_no=XXXXXX&shop_region=TH
// แล้วเปิด DevTools → Console → paste สคริปต์นี้ → Enter
// ========================================================================
(async () => {
  console.log('🔬 [QF TEST] starting...');

  // ---- 1. Check extension version ----
  const build = window.__qfContentScriptBuild || 'NOT-LOADED (extension needs reload)';
  const hasProbe = typeof window.__qfProbeContactRateLimit === 'function';
  const hasRlState = typeof window.__qfRateLimitState === 'function';
  console.log('build:', build);
  console.log('hasProbe:', hasProbe, '| hasRlState:', hasRlState);
  if (!hasProbe) {
    console.error('❌ ABORT: new content.js not loaded. Go to chrome://extensions/, click Reload on Quick Filter, then hard-refresh this page (Cmd+Shift+R).');
    return;
  }

  // ---- 2. Verify buyer_contact_info template captured ----
  // This requires ONE manual click of an eye icon first (so the hook captures
  // the URL + body template). Check if it's available.
  const templateReady = !!(window.__qfState && typeof window.__qfProbeContactRateLimit === 'function');
  console.log('extension state object:', Object.keys(window.__qfState || {}).length, 'keys');

  // ---- 3. Find eye icons on the page ----
  const eyes = [...document.querySelectorAll('svg.arco-icon-eye_invisible')];
  console.log('eye icons (invisible):', eyes.length);
  if (!eyes.length) {
    console.warn('⚠️ No 👁 eye icons found. Open an order detail page and re-run.');
    return;
  }

  // ---- 4. Hide widget + remove popovers (so clicks go through) ----
  const w = document.getElementById('qf-widget');
  if (w) w.style.display = 'none';
  document.querySelectorAll('[class*="popover"],[class*="tour"]').forEach(el => {
    try { el.remove(); } catch {}
  });

  // ---- 5. Click the FIRST eye to prime the template ----
  const primaryEye = eyes[0];
  primaryEye.scrollIntoView({ block: 'center', behavior: 'instant' });
  await new Promise(r => setTimeout(r, 400));
  const r0 = primaryEye.getBoundingClientRect();
  const cx = r0.x + r0.width / 2, cy = r0.y + r0.height / 2;
  const tgt = document.elementFromPoint(cx, cy);
  if (tgt) {
    ['pointerdown','mousedown','pointerup','mouseup','click'].forEach(type => {
      const ctor = type.startsWith('pointer') ? PointerEvent : MouseEvent;
      tgt.dispatchEvent(new ctor(type, { bubbles: true, cancelable: true, clientX: cx, clientY: cy, button: 0, buttons: 1, pointerType: 'mouse' }));
    });
  }
  console.log('✓ clicked primary eye to prime template');
  await new Promise(r => setTimeout(r, 1500));

  // ---- 6. Read the orderId from the URL ----
  const orderId = new URL(location.href).searchParams.get('order_no');
  console.log('orderId (from URL):', orderId);
  if (!orderId) {
    console.error('❌ No order_no in URL. Script works only on /order/detail?order_no=...');
    return;
  }

  // ---- 7. SAFE probe: 5 calls, 800ms gap (conservative — ~6 calls/5s) ----
  console.log('🧪 starting SAFE rate-limit probe (5 calls at 800ms gap)...');
  const probe1 = await window.__qfProbeContactRateLimit(orderId, { burst: 5, gapMs: 800 });
  console.log('probe1 (safe):', probe1);

  // If safe probe didn't hit a limit, try an aggressive one.
  const safeHit = (probe1 || []).some(x => x.limited);
  if (!safeHit) {
    console.log('✓ no rate limit at 5 calls / 800ms gap. Trying 20 calls at 100ms gap...');
    await new Promise(r => setTimeout(r, 3000));
    const probe2 = await window.__qfProbeContactRateLimit(orderId, { burst: 20, gapMs: 100 });
    const firstHit = (probe2 || []).findIndex(x => x.limited);
    if (firstHit >= 0) {
      console.log(`⚠️ rate limit detected at call #${firstHit} of burst probe`);
    } else {
      console.log('✓ no rate limit at 20 calls / 100ms gap either. Trying 50 @ 0ms...');
      await new Promise(r => setTimeout(r, 5000));
      const probe3 = await window.__qfProbeContactRateLimit(orderId, { burst: 50, gapMs: 0 });
      console.log('probe3 result:', probe3);
    }
  } else {
    console.log('⚠️ already rate-limited on safe probe — back off required');
  }

  // ---- 8. Dump final state ----
  console.log('📊 final rate-limit state:', window.__qfRateLimitState());

  // ---- 9. Check address book DB ----
  const addr = await new Promise(resolve => {
    const req = indexedDB.open('qf_address_book', 1);
    req.onsuccess = () => {
      const db = req.result;
      const tx = db.transaction('addresses', 'readonly');
      const getReq = tx.objectStore('addresses').get(String(orderId));
      getReq.onsuccess = () => resolve(getReq.result || null);
      getReq.onerror = () => resolve('ERROR');
    };
    req.onerror = () => resolve('DB-ERROR');
  });
  console.log('📇 address book record for order', orderId, ':', addr);

  console.log('🏁 [QF TEST] done — copy this entire console output and send to Claude');
})();
