(() => {
  const fontUrl = chrome.runtime.getURL('vendor/Sarabun-Bold.ttf');
  const samplePreviewUrl = chrome.runtime.getURL('vendor/jnt_sample_preview.png');
  const tiktokLogoUrl = chrome.runtime.getURL('vendor/tiktok_shop_logo.png');
  const jntLogoUrl = chrome.runtime.getURL('vendor/jnt_express_logo.png');
  const version = chrome.runtime.getManifest().version;

  const sendBasics = () => {
    window.postMessage({ __qfAsset: 'font', url: fontUrl }, '*');
    window.postMessage({ __qfAsset: 'samplePreview', url: samplePreviewUrl }, '*');
    window.postMessage({ __qfAsset: 'carrierLogo', carrier: 'tiktok', url: tiktokLogoUrl }, '*');
    window.postMessage({ __qfAsset: 'carrierLogo', carrier: 'jnt', url: jntLogoUrl }, '*');
    window.postMessage({ __qfAsset: 'manifest', version }, '*');
  };

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', sendBasics);
  }
  sendBasics();

  window.addEventListener('message', (e) => {
    if (e.source === window && (
      e.data?.__qfAsset === 'request_font' ||
      e.data?.__qfAsset === 'request_sample_preview' ||
      e.data?.__qfAsset === 'request_carrier_logos'
    )) sendBasics();
  });
})();
