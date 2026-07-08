// intercept.js — runs in the PAGE (MAIN) world at document_start.
//
// Microsoft's OnePlayer fetches the DASH video manifest from a service worker
// and authorizes the .svc.ms media CDN with an `x-spopactoken` bearer header
// (the URL's P1-P4 signature alone now returns 401 NoAccessToken — the
// "TempAuthRemoval" rollout). An extension's webRequest API can't see that
// header, so we hook window.fetch here, in the page, and forward the manifest
// URL + token to the isolated-world content script via postMessage.
//
// Video-capture portions adapted from brendangooden/ms-teams-sharepoint-downloader
// (MIT License, Copyright (c) 2025 Brendan Gooden).
(function () {
  const originalFetch = window.fetch;

  function extractHeader(args, headerName) {
    const key = headerName.toLowerCase();
    try {
      if (args[0] && typeof args[0] === 'object' && args[0].headers &&
          typeof args[0].headers.get === 'function') {
        return args[0].headers.get(key); // Request object
      }
      const init = args[1];
      if (!init || !init.headers) return null;
      const h = init.headers;
      if (typeof Headers !== 'undefined' && h instanceof Headers) return h.get(key);
      if (Array.isArray(h)) {
        const e = h.find(p => Array.isArray(p) && String(p[0]).toLowerCase() === key);
        return e ? e[1] : null;
      }
      if (typeof h === 'object') {
        for (const k of Object.keys(h)) if (k.toLowerCase() === key) return h[k];
      }
    } catch (_) { /* best-effort */ }
    return null;
  }

  window.fetch = async function (...args) {
    const response = await originalFetch.apply(this, args);
    const url = args[0];

    // Capture the videomanifest request WITHOUT a tempauth param — that's the
    // current P1-P4 + x-spopactoken regime. Trim to the index/dash manifest.
    if (url && typeof url === 'string' && url.includes('videomanifest') &&
        !/tempauth/i.test(url)) {
      let manifestUrl = url;
      const dashIndex = manifestUrl.indexOf('index&format=dash');
      if (dashIndex !== -1) {
        manifestUrl = manifestUrl.substring(0, dashIndex + 'index&format=dash'.length);
      }
      const spopactoken = extractHeader(args, 'x-spopactoken');
      console.log('[Teams DL] Detected videomanifest URL',
        spopactoken ? '(with x-spopactoken)' : '(no token)');
      // Same-frame handoff to the isolated-world content script — scope the
      // message to this frame's origin so it can't leak to other listeners.
      window.postMessage({ type: 'VIDEO_MANIFEST_URL', manifestUrl, spopactoken }, window.location.origin);
    }

    return response;
  };
})();

// Fallback: derive the manifest URL from the page's g_fileInfo global.
(function () {
  function fromFileInfo() {
    if (typeof window.g_fileInfo === 'undefined') return null;
    const transformUrl = window.g_fileInfo['.transformUrl'] ||
                         window.g_fileInfo['.providerCdnTransformUrl'];
    if (!transformUrl) return null;
    try {
      const u = new URL(transformUrl);
      u.pathname = u.pathname.replace(/\/transform\/.*$/, '/transform/videomanifest');
      u.searchParams.set('part', 'index');
      u.searchParams.set('format', 'dash');
      return u.toString();
    } catch (_) { return null; }
  }
  function tryPost() {
    const manifestUrl = fromFileInfo();
    if (!manifestUrl || /tempauth/i.test(manifestUrl)) return false;
    window.postMessage({ type: 'VIDEO_MANIFEST_URL', manifestUrl }, window.location.origin);
    return true;
  }
  if (!tryPost()) {
    const orig = window.OnLoadVideoFileInfo;
    window.OnLoadVideoFileInfo = function () {
      if (orig) orig.apply(this, arguments);
      tryPost();
    };
    window.addEventListener('load', () => setTimeout(tryPost, 1000));
  }
})();
