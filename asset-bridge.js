(() => {
  const fontUrl = chrome.runtime.getURL('vendor/Sarabun-Bold.ttf');
  const send = () => window.postMessage({ __qfAsset: 'font', url: fontUrl }, '*');
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', send);
  }
  send();
  window.addEventListener('message', (e) => {
    if (e.source === window && e.data?.__qfAsset === 'request_font') send();
  });
})();
