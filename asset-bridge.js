(() => {
  const REMOTE_MANIFEST = 'https://raw.githubusercontent.com/Sittipanpee/tiktok-quick-filter/main/manifest.json';
  const fontUrl = chrome.runtime.getURL('vendor/Sarabun-Bold.ttf');
  const version = chrome.runtime.getManifest().version;

  const sendBasics = () => {
    window.postMessage({ __qfAsset: 'font', url: fontUrl }, '*');
    window.postMessage({ __qfAsset: 'manifest', version }, '*');
  };

  // ISOLATED world fetch isn't subject to the page's CSP — safe for raw.githubusercontent.com
  const checkUpdate = async () => {
    try {
      const r = await fetch(REMOTE_MANIFEST, { cache: 'no-store' });
      if (!r.ok) return;
      const m = await r.json();
      if (m && m.version) {
        window.postMessage({ __qfAsset: 'update', remoteVersion: m.version, localVersion: version }, '*');
      }
    } catch (e) {
      // offline / network error — silent
    }
  };

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', sendBasics);
  }
  sendBasics();
  checkUpdate();

  window.addEventListener('message', (e) => {
    if (e.source === window && e.data?.__qfAsset === 'request_font') sendBasics();
  });
})();
