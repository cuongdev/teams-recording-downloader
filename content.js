// content.js — isolated-world content script.
//
// Receives the manifest URL + x-spopactoken from intercept.js (page world),
// then on click downloads the DASH segments, decrypts the DASH-SEA
// AES-128-CBC "clear-key" encryption in-browser (Web Crypto), and muxes the
// audio+video fMP4 tracks into a flat MP4 via mux-worker.js.
//
// Download/decrypt/mux pipeline adapted from
// brendangooden/ms-teams-sharepoint-downloader (MIT License,
// Copyright (c) 2025 Brendan Gooden).
(function () {
  'use strict';

  let videoManifestUrl = null;
  let videoSpopActoken = null;
  let cueCache = { key: null, cues: null };     // scraped transcript, reused across actions
  let scrapeInFlight = null;                    // shared scrape promise (dedupes concurrent scrapes)
  const concurrency = 4; // global in-flight segment budget; SharePoint throttles in the low teens
  const isTopFrame = window.top === window.self;

  // --- receive captures from the page-world interceptor ---
  window.addEventListener('message', (event) => {
    if (event.source !== window || event.origin !== window.location.origin || !event.data) return;
    if (event.data.type === 'VIDEO_MANIFEST_URL') {
      videoManifestUrl = event.data.manifestUrl;
      if (event.data.spopactoken) videoSpopActoken = event.data.spopactoken;
      showButton();
      // If the recording lives in a child frame, this frame renders its own
      // (working) button — tell the top frame to drop its placeholder so the
      // user doesn't see two buttons.
      if (!isTopFrame) {
        try { chrome.runtime.sendMessage({ type: 'recordingFrameReady' }); } catch (_) {}
      }
    }
  });

  // The top frame's placeholder hides itself once a child frame reports it owns
  // the real recording (and only while the top frame hasn't captured its own).
  try {
    chrome.runtime.onMessage.addListener((request) => {
      if (request && request.type === 'hidePlaceholderButton' && isTopFrame && !videoManifestUrl) {
        hideButton();
      }
    });
  } catch (_) { /* chrome.runtime may be unavailable in some sandboxed frames */ }

  function svcMsFetchInit(extra) {
    const init = Object.assign({}, extra || {});
    if (videoSpopActoken) {
      init.headers = Object.assign({}, init.headers || {}, { 'x-spopactoken': videoSpopActoken });
    }
    return init;
  }

  // ===========================================================================
  // DASH manifest parsing
  // ===========================================================================
  function hexToBytes(hex) {
    const out = new Uint8Array(hex.length / 2);
    for (let i = 0; i < out.length; i++) out[i] = parseInt(hex.substr(i * 2, 2), 16);
    return out;
  }

  function parseDashManifest(xmlText, manifestUrl) {
    const doc = new DOMParser().parseFromString(xmlText, 'application/xml');
    if (doc.querySelector('parsererror')) throw new Error('Failed to parse DASH manifest XML');

    const baseUrlEl = doc.querySelector('BaseURL');
    const manifestDerivedBase = manifestUrl.split('?')[0].replace(/\/[^/]*$/, '/');
    const baseUrl = (baseUrlEl && baseUrlEl.textContent.trim()) || manifestDerivedBase;

    function toAbsolute(url) {
      if (!url) return '';
      if (/^https:\/\//.test(url)) return url;
      if (/^[a-z][a-z0-9+\-.]*:/i.test(url)) throw new Error('Unsafe URL scheme in manifest: ' + url);
      return new URL(url, baseUrl).href;
    }
    function expandTemplate(tpl, repId, bandwidth, number, time) {
      return tpl
        .replace(/\$RepresentationID\$/g, repId)
        .replace(/\$Bandwidth\$/g, bandwidth)
        .replace(/\$Number%0(\d+)d\$/g, (_, w) => String(number).padStart(parseInt(w, 10), '0'))
        .replace(/\$Number\$/g, String(number))
        .replace(/\$Time\$/g, String(time));
    }

    const adaptationSets = Array.from(doc.querySelectorAll('AdaptationSet'));
    const isMuxed = adaptationSets.length === 1;
    const tracks = [];

    for (const as of adaptationSets) {
      let type = as.getAttribute('contentType') || '';
      if (!type) {
        const mime = as.getAttribute('mimeType') || '';
        type = mime.startsWith('video') ? 'video' : mime.startsWith('audio') ? 'audio' : '';
      }
      if (isMuxed) type = 'muxed';

      const reps = Array.from(as.querySelectorAll('Representation'))
        .sort((a, b) => parseInt(b.getAttribute('bandwidth') || '0', 10) - parseInt(a.getAttribute('bandwidth') || '0', 10));
      const rep = reps[0];
      if (!rep) continue;

      const repId = rep.getAttribute('id') || '';
      const bandwidth = rep.getAttribute('bandwidth') || '';
      const mimeType = rep.getAttribute('mimeType') || as.getAttribute('mimeType') || '';
      const segTpl = rep.querySelector('SegmentTemplate') || as.querySelector('SegmentTemplate');
      if (!segTpl) continue;

      const startNumber = parseInt(segTpl.getAttribute('startNumber') || '1', 10);
      const initUrl = toAbsolute(expandTemplate(segTpl.getAttribute('initialization') || '', repId, bandwidth, startNumber, 0));
      const mediaTpl = segTpl.getAttribute('media') || '';
      const segments = [];

      const timeline = segTpl.querySelector('SegmentTimeline');
      if (timeline) {
        let t = 0, segNum = startNumber;
        for (const s of timeline.querySelectorAll('S')) {
          const sT = s.getAttribute('t');
          if (sT !== null) t = parseInt(sT, 10);
          const d = parseInt(s.getAttribute('d') || '0', 10);
          const r = parseInt(s.getAttribute('r') || '0', 10);
          for (let i = 0; i <= r; i++) {
            segments.push(toAbsolute(expandTemplate(mediaTpl, repId, bandwidth, segNum, t)));
            t += d; segNum++;
          }
        }
      } else {
        const duration = parseInt(segTpl.getAttribute('duration') || '0', 10);
        const timescale = parseInt(segTpl.getAttribute('timescale') || '1', 10);
        const period = as.closest('Period');
        const periodDur = period ? (() => {
          const m = (period.getAttribute('duration') || '').match(/PT(?:(\d+)H)?(?:(\d+)M)?(?:([\d.]+)S)?/);
          return m ? parseInt(m[1] || '0') * 3600 + parseInt(m[2] || '0') * 60 + parseFloat(m[3] || '0') : 0;
        })() : 0;
        if (duration > 0 && periodDur > 0) {
          const count = Math.ceil(periodDur / (duration / timescale));
          for (let i = 0; i < count; i++) {
            segments.push(toAbsolute(expandTemplate(mediaTpl, repId, bandwidth, startNumber + i, i * duration)));
          }
        }
      }

      // DASH-SEA AES-128-CBC encryption (HTTP-fetchable key, not hard DRM).
      let encryption = null;
      const seaCp = [...as.querySelectorAll('ContentProtection')].find(cp =>
        cp.getAttribute('schemeIdUri') === 'urn:mpeg:dash:sea:2012');
      if (seaCp) {
        const segEnc = seaCp.querySelector('SegmentEncryption');
        const scheme = segEnc ? segEnc.getAttribute('schemeIdUri') : '';
        const period = seaCp.querySelector('CryptoPeriod');
        const keyUri = period ? period.getAttribute('keyUriTemplate') : null;
        const ivAttr = period ? (period.getAttribute('IV') || '') : '';
        if (/aes128-cbc/i.test(scheme) && keyUri && ivAttr) {
          encryption = { scheme: 'aes-128-cbc', keyUri, iv: hexToBytes(ivAttr.replace(/^0x/i, '')) };
        }
      }

      tracks.push({ type, mimeType, initUrl, segments, encryption });
    }
    return tracks;
  }

  // ===========================================================================
  // Segment download + decrypt
  // ===========================================================================
  function abortableSleep(ms, signal) {
    return new Promise((resolve, reject) => {
      if (signal && signal.aborted) return reject(Object.assign(new Error('Cancelled'), { name: 'AbortError' }));
      const t = setTimeout(() => { if (signal) signal.removeEventListener('abort', onAbort); resolve(); }, ms);
      function onAbort() { clearTimeout(t); reject(Object.assign(new Error('Cancelled'), { name: 'AbortError' })); }
      if (signal) signal.addEventListener('abort', onAbort, { once: true });
    });
  }

  async function fetchWithRetry(url, init, signal, onThrottle, maxAttempts = 6) {
    let attempt = 0;
    for (;;) {
      attempt++;
      if (signal && signal.aborted) throw Object.assign(new Error('Cancelled'), { name: 'AbortError' });
      let resp;
      try {
        resp = await fetch(url, init);
      } catch (e) {
        if (e.name === 'AbortError') throw e;
        if (attempt >= maxAttempts) throw e;
        const delayMs = Math.min(1000 * Math.pow(2, attempt - 1), 30000);
        if (onThrottle) onThrottle({ attempt, delayMs, status: 0 });
        await abortableSleep(delayMs, signal);
        continue;
      }
      if ((resp.status === 429 || resp.status === 503) && attempt < maxAttempts) {
        const headerSecs = parseInt(resp.headers.get('Retry-After'), 10);
        const delayMs = Number.isFinite(headerSecs) && headerSecs > 0
          ? Math.min(headerSecs * 1000, 30000)
          : Math.min(1000 * Math.pow(2, attempt - 1), 30000);
        if (onThrottle) onThrottle({ attempt, delayMs, status: resp.status });
        await abortableSleep(delayMs, signal);
        continue;
      }
      return resp;
    }
  }

  async function downloadDashSegments(tracks, onProgress, signal) {
    const totalSegs = tracks.reduce((s, t) => s + (t.initUrl ? 1 : 0) + t.segments.length, 0);
    let done = 0;
    function reportProgress(text) { onProgress(done, totalSegs, text); }
    function noteThrottle({ attempt, delayMs, status }) {
      reportProgress(`HTTP ${status || 'network'} — backing off ${Math.round(delayMs / 1000)}s (attempt ${attempt})...`);
    }

    const trackStates = await Promise.all(tracks.map(async (track) => {
      const label = tracks.length > 1 ? ` (${track.type} track)` : '';
      let cryptoKey = null;
      if (track.encryption) {
        reportProgress(`Fetching encryption key${label}...`);
        const init = track.encryption.keyUri.includes('svc.ms') && videoSpopActoken
          ? { signal, headers: { 'x-spopactoken': videoSpopActoken } }
          : { signal };
        const keyResp = await fetchWithRetry(track.encryption.keyUri, init, signal, noteThrottle);
        if (!keyResp.ok) throw new Error(`Encryption key fetch failed: HTTP ${keyResp.status}`);
        const keyBuf = await keyResp.arrayBuffer();
        cryptoKey = await crypto.subtle.importKey('raw', keyBuf, { name: 'AES-CBC' }, false, ['decrypt']);
      }
      async function decryptIfNeeded(buf) {
        if (!cryptoKey) return buf;
        return await crypto.subtle.decrypt({ name: 'AES-CBC', iv: track.encryption.iv }, cryptoKey, buf);
      }

      const orderedBufs = new Array((track.initUrl ? 1 : 0) + track.segments.length);
      let segStart = 0;
      if (track.initUrl) {
        reportProgress(`Fetching init segment${label}...`);
        const r = await fetchWithRetry(track.initUrl, { signal }, signal, noteThrottle);
        if (!r.ok) throw new Error(`Init segment failed: HTTP ${r.status}`);
        orderedBufs[0] = await decryptIfNeeded(await r.arrayBuffer());
        done++; segStart = 1;
      }
      return { track, orderedBufs, segStart, decryptIfNeeded };
    }));

    const queue = [];
    for (const st of trackStates) for (let si = 0; si < st.track.segments.length; si++) queue.push({ st, si });
    reportProgress(`Downloading ${queue.length} segments (${concurrency} parallel)...`);

    await new Promise((resolve, reject) => {
      if (queue.length === 0) return resolve();
      let qIdx = 0, inFlight = 0, rejected = false;
      function launch() {
        while (!rejected && inFlight < concurrency && qIdx < queue.length) {
          if (signal && signal.aborted) { rejected = true; return reject(Object.assign(new Error('Cancelled'), { name: 'AbortError' })); }
          const job = queue[qIdx++];
          inFlight++;
          fetchWithRetry(job.st.track.segments[job.si], { signal }, signal, noteThrottle)
            .then(r => { if (!r.ok) throw new Error(`Segment failed: HTTP ${r.status}`); return r.arrayBuffer(); })
            .then(job.st.decryptIfNeeded)
            .then(buf => {
              if (rejected) return;
              job.st.orderedBufs[job.st.segStart + job.si] = buf;
              done++;
              reportProgress(`Downloading segments... (${done}/${totalSegs})`);
              inFlight--;
              if (inFlight === 0 && qIdx >= queue.length) resolve(); else launch();
            })
            .catch(err => { if (!rejected) { rejected = true; reject(err); } });
        }
      }
      launch();
    });

    return trackStates.map(s => s.orderedBufs);
  }

  // ===========================================================================
  // Mux (Web Worker)
  // ===========================================================================
  let _muxWorkerBlobUrl = null;
  async function getMuxWorkerUrl() {
    if (_muxWorkerBlobUrl) return _muxWorkerBlobUrl;
    const resp = await fetch(chrome.runtime.getURL('mux-worker.js'));
    if (!resp.ok) throw new Error(`mux-worker fetch failed: HTTP ${resp.status}`);
    const src = await resp.text();
    _muxWorkerBlobUrl = URL.createObjectURL(new Blob([src], { type: 'application/javascript' }));
    return _muxWorkerBlobUrl;
  }

  function toArrayBuffers(chunks) {
    return chunks.map(b => {
      if (b instanceof ArrayBuffer) return b;
      if (ArrayBuffer.isView(b)) {
        return b.byteOffset === 0 && b.byteLength === b.buffer.byteLength ? b.buffer : b.slice().buffer;
      }
      return b;
    });
  }

  async function muxTracks(videoChunks, audioChunks, onProgress) {
    const workerUrl = await getMuxWorkerUrl();
    return new Promise((resolve, reject) => {
      const worker = new Worker(workerUrl);
      worker.onmessage = (e) => {
        if (e.data.progress) { const p = e.data.progress; onProgress(p.done, p.total, p.text); return; }
        if (e.data.error) { worker.terminate(); reject(new Error(e.data.error)); return; }
        if (e.data.result) { worker.terminate(); resolve(e.data.result); }
      };
      worker.onerror = (e) => { worker.terminate(); reject(new Error(e.message || 'mux-worker crashed')); };
      const video = toArrayBuffers(videoChunks);
      const audio = toArrayBuffers(audioChunks);
      try {
        worker.postMessage({ video, audio }, [...video, ...audio]);
      } catch (_) {
        worker.postMessage({ video, audio }); // structured-clone fallback
      }
    });
  }

  function downloadFile(data, filename, mime = 'video/mp4') {
    const blob = new Blob(Array.isArray(data) ? data : [data], { type: mime });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  // ===========================================================================
  // Orchestration
  // ===========================================================================
  const HARD_DRM_SCHEMES = [
    'edef8ba9-79d6-4ace-a3c8-27dcd51d21ed', // Widevine
    '9a04f079-9840-4286-ab92-e65be0885f95', // PlayReady
    '94ce86fb-07ff-4f43-adb8-93d2fa968ca2'  // FairPlay
  ];

  async function runDownload(filename, onProgress, signal) {
    onProgress(0, 1, 'Fetching manifest...');
    const resp = await fetch(videoManifestUrl, svcMsFetchInit({ signal }));
    if (!resp.ok) throw new Error(`Manifest fetch failed: HTTP ${resp.status}`);
    const xmlText = await resp.text();

    const cpSchemes = [...xmlText.matchAll(/<ContentProtection\b[^>]*schemeIdUri="([^"]+)"/gi)].map(m => m[1].toLowerCase());
    if (cpSchemes.some(s => HARD_DRM_SCHEMES.some(uuid => s.includes(uuid)))) {
      throw Object.assign(new Error('DRM_PROTECTED'), { isDrm: true });
    }

    onProgress(0, 1, 'Parsing manifest...');
    const allTracks = parseDashManifest(xmlText, videoManifestUrl);
    if (!allTracks.length) throw new Error('No tracks found in manifest');

    const videoTrack = allTracks.find(t => t.type === 'video' || t.type === 'muxed');
    const audioTrack = allTracks.find(t => t.type === 'audio');
    const safe = filename.replace(/[^a-z0-9\s_-]/gi, '_');

    if (videoTrack && audioTrack) {
      const trackData = await downloadDashSegments([videoTrack, audioTrack], onProgress, signal);
      const muxed = await muxTracks(trackData[0], trackData[1], onProgress);
      downloadFile(muxed, safe + '.mp4');
    } else {
      const only = videoTrack || audioTrack || allTracks[0];
      const trackData = await downloadDashSegments([only], onProgress, signal);
      downloadFile(trackData[0], safe + (only.type === 'audio' ? '.m4a' : '.mp4'));
    }
    onProgress(1, 1, 'Download complete!');
  }

  // ===========================================================================
  // Transcript scraping (from the rendered Teams/Stream transcript panel) +
  // conversion to VTT / SRT / TXT.
  //
  // The transcript API returns 400/notSupported for these recordings, but the
  // panel renders every line into the DOM. The panel is a Fluent UI virtualized
  // `ms-List` — only on-screen rows exist at any moment — so we auto-scroll the
  // container and accumulate rows keyed by their stable `aria-posinset`. Every
  // row carries speaker + relative timestamp in its `aria-label` (even
  // continuation rows with no visible header), and `aria-setsize` tells us the
  // total row count, which is our completion signal.
  // ===========================================================================
  function delay(ms) { return new Promise(r => setTimeout(r, ms)); }

  function fmtTimestamp(sec, msSep) {
    if (!isFinite(sec) || sec < 0) sec = 0;
    const h = Math.floor(sec / 3600);
    const m = Math.floor((sec % 3600) / 60);
    const s = Math.floor(sec % 60);
    const ms = Math.round((sec - Math.floor(sec)) * 1000);
    const p = (n, l = 2) => String(n).padStart(l, '0');
    return `${p(h)}:${p(m)}:${p(s)}${msSep}${p(ms, 3)}`;
  }

  // "Duc P 0 minutes 8 seconds" / "+15*******37 1 hour 2 minutes 3 seconds"
  // -> { speaker: "Duc P", seconds: 8 }
  function parseTranscriptLabel(label) {
    // Trailing time is one or more "<n> hour|minute|second(s)" tokens. Teams
    // omits units that are zero, so a whole-minute mark reads "48 minutes" with
    // no seconds — the duration must not be required to end in seconds, or the
    // "48 minutes" gets glued onto the speaker name (splitting one person in two).
    const m = label.match(/^(.*?)\s+((?:\d+\s*(?:hours?|minutes?|seconds?)\b\s*)+)$/i);
    if (!m) return { speaker: label.trim(), seconds: null };
    const dur = m[2];
    let seconds = 0;
    const h = dur.match(/(\d+)\s*hours?/i); if (h) seconds += parseInt(h[1], 10) * 3600;
    const mi = dur.match(/(\d+)\s*minutes?/i); if (mi) seconds += parseInt(mi[1], 10) * 60;
    const s = dur.match(/(\d+)\s*seconds?/i); if (s) seconds += parseInt(s[1], 10);
    return { speaker: m[1].trim(), seconds };
  }

  function isTranscriptPanelPresent() {
    return !!document.querySelector('[data-testid="entryRenderer"], [id^="sub-entry-"]');
  }

  // The scrollable ancestor of the transcript list.
  function findTranscriptScroller() {
    const list = document.querySelector('.ms-List[role="list"]') ||
      (document.querySelector('[data-testid="entryRenderer"]') || {}).closest?.('.ms-List') ||
      document.querySelector('.ms-List');
    if (!list) return null;
    const scrollable = list.closest('[data-is-scrollable="true"]');
    if (scrollable) return scrollable;
    let el = list;
    while (el && el !== document.body) {
      const s = getComputedStyle(el);
      if ((s.overflowY === 'auto' || s.overflowY === 'scroll') && el.scrollHeight > el.clientHeight + 20) return el;
      el = el.parentElement;
    }
    return list.parentElement || list;
  }

  function collectTranscriptEntries(map) {
    for (const g of document.querySelectorAll('div[role="group"][id^="entry-"]')) {
      const label = (g.getAttribute('aria-label') || '').trim();
      if (!label) continue; // meeting-event rows ("started transcription") have a blank label
      const sub = g.querySelector('[id^="sub-entry-"]');
      if (!sub) continue;
      const text = (sub.innerText || sub.textContent || '').replace(/\s+/g, ' ').trim();
      if (!text) continue;
      let pos = parseInt(sub.getAttribute('aria-posinset') || '0', 10);
      if (!pos) pos = (parseInt((g.id.match(/\d+/) || [0])[0], 10) || 0) + 1;
      if (map.has(pos)) continue;
      const { speaker, seconds } = parseTranscriptLabel(label);
      map.set(pos, { pos, speaker, seconds, text });
    }
  }

  async function scrapeTranscriptCues(onProgress, signal) {
    const scroller = findTranscriptScroller();
    if (!scroller) throw new Error('NO_TRANSCRIPT');
    const setSizeEl = document.querySelector('[id^="sub-entry-"][aria-setsize]');
    const total = setSizeEl ? parseInt(setSizeEl.getAttribute('aria-setsize'), 10) : 0;
    const totalTxt = total ? `/${total}` : '';
    const map = new Map();

    scroller.scrollTop = 0;
    await delay(150);
    collectTranscriptEntries(map);

    let stale = 0;
    for (let i = 0; i < 3000; i++) {
      if (signal && signal.aborted) throw Object.assign(new Error('Cancelled'), { name: 'AbortError' });
      if (total && map.size >= total) break;
      const before = scroller.scrollTop;
      const atBottom = before + scroller.clientHeight >= scroller.scrollHeight - 2;
      scroller.scrollTop = Math.min(before + Math.max(scroller.clientHeight * 0.75, 200), scroller.scrollHeight);
      await delay(130);
      collectTranscriptEntries(map);
      onProgress(map.size, total || map.size, `Reading transcript… (${map.size}${totalTxt})`);
      if (atBottom || scroller.scrollTop <= before + 1) {
        if (++stale >= 4) break;
      } else {
        stale = 0;
      }
    }
    collectTranscriptEntries(map);

    const cues = [...map.values()].sort((a, b) => a.pos - b.pos);
    if (!cues.length) throw new Error('NO_TRANSCRIPT');
    // Fill any gap in parsed times, then derive each cue's end from the next start.
    for (let i = 0; i < cues.length; i++) {
      if (cues[i].seconds == null) cues[i].seconds = i ? cues[i - 1].seconds : 0;
    }
    for (let i = 0; i < cues.length; i++) {
      const start = cues[i].seconds;
      let end = i + 1 < cues.length ? cues[i + 1].seconds : start + 3;
      if (end <= start) end = start + 2;
      cues[i].start = start;
      cues[i].end = end;
    }
    return cues;
  }

  // Scrape once per recording, then reuse. Keyed by URL so navigating to a
  // different recording re-scrapes. Pass force=true to bypass the cache.
  async function getCues(onProgress, signal, force) {
    const key = location.href;
    if (!force && cueCache.cues && cueCache.key === key) {
      onProgress(cueCache.cues.length, cueCache.cues.length, `Using cached transcript (${cueCache.cues.length} lines).`);
      return cueCache.cues;
    }
    // Only one scroll pass at a time: concurrent callers (e.g. transcript +
    // insights clicked together) share the same in-flight scrape.
    if (!scrapeInFlight) {
      scrapeInFlight = scrapeTranscriptCues(onProgress, signal)
        .then((cues) => { cueCache = { key, cues }; return cues; })
        .finally(() => { scrapeInFlight = null; });
    }
    const cues = await scrapeInFlight;
    if (signal && signal.aborted) throw Object.assign(new Error('Cancelled'), { name: 'AbortError' });
    return cues;
  }

  function cuesToVtt(cues) {
    let out = 'WEBVTT\n\n';
    cues.forEach((c, i) => {
      const v = c.speaker ? `<v ${c.speaker}>` : '';
      out += `${i + 1}\n${fmtTimestamp(c.start, '.')} --> ${fmtTimestamp(c.end, '.')}\n${v}${c.text}\n\n`;
    });
    return out;
  }

  async function runTranscriptDownload(filename, onProgress, signal) {
    onProgress(0, 1, 'Scanning transcript panel…');
    const cues = await getCues(onProgress, signal);
    const safe = filename.replace(/[^a-z0-9\s_-]/gi, '_');
    downloadFile(cuesToVtt(cues), safe + '.vtt', 'text/vtt;charset=utf-8');
    onProgress(1, 1, `Transcript downloaded (${cues.length} lines).`);
  }

  // ===========================================================================
  // Meeting insights — computed locally from the scraped transcript. Nothing
  // leaves the browser; the LLM-only analyses are packaged as a copy-paste
  // prompt embedded in the dashboard (see buildLlmPrompt).
  // ===========================================================================
  const STOPWORDS = new Set((
    'the a an and or but if then else so of to in on at by for with without from up down out over under ' +
    'again once here there all any both each few more most some such no nor not only own same than too very ' +
    'is are was were be been being have has had do does did doing will would shall should can could may might must ' +
    'i me my we our us you your he him his she her it its they them their this that these those which who whom what ' +
    'as at into about against between through during before after above below off am pre re ' +
    'yeah yes yep ok okay right like just kind sort gonna wanna gotta really actually basically literally ' +
    'know think mean going get got go went one two also well thing things stuff something anything everything ' +
    'okay uh um hmm oh hey guys guy folks team let lets say said says thats theres im ive dont doesnt cant wont ' +
    'now new use used using make made makes want need see seen look looks even still back way ways lot lots'
  ).split(/\s+/));

  function secToClock(sec) {
    sec = Math.max(0, Math.round(sec || 0));
    const h = Math.floor(sec / 3600), m = Math.floor((sec % 3600) / 60), s = sec % 60;
    const p = (n) => String(n).padStart(2, '0');
    return h > 0 ? `${h}:${p(m)}:${p(s)}` : `${m}:${p(s)}`;
  }

  function escapeHtml(str) {
    return String(str).replace(/[&<>"']/g, (c) => (
      { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]
    ));
  }

  function computeInsights(cues) {
    const bySpeaker = new Map();
    const ensure = (name) => {
      if (!bySpeaker.has(name)) bySpeaker.set(name, {
        name, talkSec: 0, turns: 0, words: 0, questions: 0, affirmations: 0, responses: 0
      });
      return bySpeaker.get(name);
    };
    const affRe = /\b(yeah|yes|yep|agree|agreed|right|exactly|correct|sure|okay|ok|makes sense|good point|true|totally|definitely|absolutely)\b/i;
    const transitions = new Map(); // "A → B" -> count
    let totalTurns = 0, totalQuestions = 0, totalWords = 0;
    let prev = null;

    for (const c of cues) {
      const sp = ensure(c.speaker || 'Unknown');
      sp.talkSec += Math.max(0, (c.end != null ? c.end : c.start) - c.start);
      const w = c.text.trim().split(/\s+/).filter(Boolean).length;
      sp.words += w; totalWords += w;
      if (c.text.includes('?')) { sp.questions++; totalQuestions++; }
      if (w <= 6 && affRe.test(c.text)) sp.affirmations++;
      if (c.speaker !== prev) {
        sp.turns++; totalTurns++;
        if (prev) {
          const key = prev + ' → ' + c.speaker;
          transitions.set(key, (transitions.get(key) || 0) + 1);
          sp.responses++;
        }
      }
      prev = c.speaker;
    }

    const parts = [...bySpeaker.values()];
    const talkSum = parts.reduce((s, p) => s + p.talkSec, 0) || 1;
    const maxCollabRaw = Math.max(1, ...parts.map(p => p.questions + p.affirmations + p.responses));
    for (const p of parts) {
      p.talkPct = p.talkSec / talkSum;
      const talkShare = p.talkSec / talkSum;
      const turnShare = totalTurns ? p.turns / totalTurns : 0;
      const qShare = totalQuestions ? p.questions / totalQuestions : 0;
      p.engagement = Math.round(100 * (0.5 * talkShare + 0.3 * turnShare + 0.2 * qShare));
      p.collabRaw = p.questions + p.affirmations + p.responses;
      p.collabScore = Math.round(100 * p.collabRaw / maxCollabRaw);
    }
    parts.sort((a, b) => b.talkSec - a.talkSec);

    const topTransitions = [...transitions.entries()]
      .map(([pair, count]) => ({ pair, count }))
      .sort((a, b) => b.count - a.count).slice(0, 12);

    const start0 = cues.length ? cues[0].start : 0;
    const durationSec = cues.length ? (cues[cues.length - 1].end - start0) : 0;

    return {
      parts, topTransitions,
      totalTurns, totalQuestions, totalWords,
      durationSec, start0,
      wpm: durationSec ? Math.round(totalWords / (durationSec / 60)) : 0,
      lineCount: cues.length
    };
  }

  // The prompt the user pastes into their approved AI tool for the qualitative
  // analyses (summary, decisions, action items, sentiment, styles, rubric).
  function buildLlmPrompt(cues, title) {
    const lines = cues.map(c => `[${secToClock(c.start)}] ${c.speaker}: ${c.text}`).join('\n');
    return (
`You are an expert meeting analyst. Analyse the following meeting transcript and reply in Markdown with these sections:

1. **Summary** — 4–6 sentences.
2. **Key Decisions** — bullet list (decision + who drove it).
3. **Action Items** — table: Owner | Task | Due (if stated).
4. **Risks / Blockers** — bullet list.
5. **Topics Discussed** — bullet list, most to least time.
6. **Timeline Highlights** — key moments with [mm:ss].
7. **Meeting Sentiment** — overall tone + any notable shifts.
8. **Communication Style per Participant** — 2–4 adjectives each (e.g. analytical, concise, exploratory, supportive, facilitator).
9. **Contribution Reflection per Participant** — for each person score 1–5 on Leadership, Problem-solving, Ownership, Collaboration, Communication, each with one line of justification quoting the transcript.

IMPORTANT: Section 9 is an AI-generated reflection to help individuals grow. It is NOT a formal performance evaluation and must not be used as one on its own. Note where the transcript is insufficient to judge.

Meeting title: ${title}

Transcript:
${lines}
`);
  }

  function bar(label, valueText, ratio, color) {
    const pct = Math.max(0, Math.min(100, Math.round(ratio * 100)));
    return (
`<div class="row">
  <div class="row-label">${escapeHtml(label)}</div>
  <div class="track"><div class="fill" style="width:${pct}%;background:${color}"></div></div>
  <div class="row-val">${escapeHtml(valueText)}</div>
</div>`);
  }

  const PALETTE = ['#5b8ff9', '#61ddaa', '#f6bd16', '#ff9d4d', '#e8684a', '#9270ca', '#269a99', '#ff99c3', '#6dc8ec', '#d3648b'];

  function buildTimelineSvg(insights, cues) {
    const parts = insights.parts;
    const idx = new Map(parts.map((p, i) => [p.name, i]));
    const W = 1000, ROW = 30, BARH = 22;
    const H = parts.length * ROW;
    const D = insights.durationSec || 1;

    // Nice tick step: ~6–8 gridlines across the meeting.
    const STEPS = [30, 60, 120, 300, 600, 900, 1800, 3600, 7200];
    const step = STEPS.find(s => D / s <= 8) || Math.ceil(D / 8);
    let grid = '', axis = '';
    for (let t = 0; t <= D + 0.5; t += step) {
      const x = (t / D) * W, pct = (t / D) * 100;
      grid += `<line x1="${x.toFixed(1)}" y1="0" x2="${x.toFixed(1)}" y2="${H}"/>`;
      const tx = pct <= 1 ? '0' : (pct >= 99 ? '-100%' : '-50%');
      axis += `<span style="left:${pct.toFixed(2)}%;transform:translateX(${tx})">${secToClock(t)}</span>`;
    }

    let rects = '';
    for (const c of cues) {
      const i = idx.get(c.speaker); if (i == null) continue;
      const rel = c.start - insights.start0;
      const x = (rel / D) * W;
      const w = Math.max(1.5, (Math.max(0, c.end - c.start) / D) * W);
      const y = i * ROW + (ROW - BARH) / 2;
      rects += `<rect x="${x.toFixed(1)}" y="${y}" width="${w.toFixed(1)}" height="${BARH}" rx="2" fill="${PALETTE[i % PALETTE.length]}"><title>${escapeHtml(c.speaker)} · ${secToClock(rel)}</title></rect>`;
    }

    const names = parts.map((p) =>
      `<div class="tl-n" style="height:${ROW}px">${escapeHtml(p.name)}</div>`
    ).join('');

    return (
`<div class="tl">
  <div class="tl-names">${names}</div>
  <div class="tl-chart">
    <svg viewBox="0 0 ${W} ${H}" preserveAspectRatio="none" style="width:100%;height:${H}px" role="img" aria-label="Participation timeline — hover a segment for its time">
      <g class="tl-grid">${grid}</g>${rects}
    </svg>
    <div class="tl-axis">${axis}</div>
  </div>
</div>`);
  }

  const DASHBOARD_CSS =
`:root,:host{--bg:#14151a;--card:#1e2028;--card2:#262935;--line:#333644;--text:#e8eaf0;--muted:#9aa0ad;--accent:#61ddaa}
*{box-sizing:border-box}
body{margin:0}
body,.wrap{background:var(--bg);color:var(--text);font:15px/1.5 "Segoe UI",system-ui,sans-serif}
.wrap{max-width:1040px;margin:0 auto;padding:28px 20px 60px}
h1{font-size:22px;margin:0 0 2px}
.sub{color:var(--muted);font-size:13px;margin-bottom:22px}
h2{font-size:15px;text-transform:uppercase;letter-spacing:.06em;color:var(--muted);margin:30px 0 12px;font-weight:600}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:18px}
.tiles{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:12px}
.tile{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:16px}
.tile-v{font-size:26px;font-weight:700}
.tile-k{color:var(--muted);font-size:12px;text-transform:uppercase;letter-spacing:.05em;margin-top:4px}
.row{display:grid;grid-template-columns:150px 1fr auto;align-items:center;gap:12px;margin:8px 0}
.row-label{color:var(--text);font-size:13px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.row-val{color:var(--muted);font-size:12px;min-width:70px;text-align:right}
.track{background:var(--card2);border-radius:6px;height:16px;overflow:hidden}
.fill{height:100%;border-radius:6px;min-width:2px}
table{width:100%;border-collapse:collapse;font-size:13px}
th,td{text-align:left;padding:8px 10px;border-bottom:1px solid var(--line)}
th{color:var(--muted);font-weight:600;font-size:11px;text-transform:uppercase;letter-spacing:.05em}
td:not(:first-child),th:not(:first-child){text-align:right}
.cols{display:grid;grid-template-columns:1fr 1fr;gap:20px}
@media(max-width:720px){.cols{grid-template-columns:1fr}.row{grid-template-columns:110px 1fr auto}}
ul.trans{list-style:none;margin:0;padding:0}
ul.trans li{display:flex;justify-content:space-between;padding:6px 0;border-bottom:1px solid var(--line);font-size:13px}
ul.trans b{color:var(--accent)}
.tl{display:flex;gap:10px}
.tl-names{flex:none;max-width:170px}
.tl-n{display:flex;align-items:center;font-size:12px;font-weight:600;color:var(--text);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.tl-chart{flex:1;min-width:0;position:relative}
.tl-chart svg{display:block;background:var(--card2);border-radius:8px}
.tl-chart svg rect{cursor:crosshair}
.tl-grid line{stroke:rgba(255,255,255,.06)}
.tl-axis{position:relative;height:14px;margin-top:5px}
.tl-axis span{position:absolute;font-size:10px;color:var(--muted);white-space:nowrap}
.scards{display:grid;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));gap:14px}
.scard{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:16px}
.scard-h{display:flex;justify-content:space-between;align-items:baseline;margin-bottom:10px}
.scard-h .nm{font-weight:700;font-size:14px}
.scard-h .eng{color:var(--accent);font-weight:700;font-size:12px;white-space:nowrap}
.srow{display:grid;grid-template-columns:100px 1fr 30px;align-items:center;gap:8px;margin:6px 0}
.srow .l{font-size:12px;color:var(--muted)}
.srow .v{font-size:12px;text-align:right;color:var(--muted)}
.note{color:var(--muted);font-size:12px;margin-top:8px}
textarea{width:100%;height:260px;background:#0d0e12;color:#cdd3df;border:1px solid var(--line);border-radius:8px;padding:12px;font:12px/1.5 ui-monospace,Menlo,Consolas,monospace;resize:vertical}
button.copy{background:var(--accent);color:#062b1e;border:0;padding:10px 16px;border-radius:8px;font-weight:700;cursor:pointer;margin-bottom:10px}
.disc{background:#2a2410;border:1px solid #5a4d18;color:#e7d9a6;border-radius:10px;padding:12px 14px;font-size:12.5px;margin-top:14px}
.foot{color:var(--muted);font-size:11px;margin-top:34px;text-align:center}`;

  function dashboardBody(insights, cues, title) {
    const p = insights.parts;
    const maxTalk = Math.max(1, ...p.map(x => x.talkSec));
    const maxEng = Math.max(1, ...p.map(x => x.engagement));

    const tiles = [
      ['Duration', secToClock(insights.durationSec)],
      ['Participants', String(p.length)],
      ['Speaking turns', String(insights.totalTurns)],
      ['Questions', String(insights.totalQuestions)],
      ['Total words', insights.totalWords.toLocaleString()],
      ['Words / min', String(insights.wpm)]
    ].map(([k, v]) => `<div class="tile"><div class="tile-v">${escapeHtml(v)}</div><div class="tile-k">${escapeHtml(k)}</div></div>`).join('');

    const talkBars = p.map((x, i) =>
      bar(x.name, `${secToClock(x.talkSec)} · ${Math.round(x.talkPct * 100)}%`, x.talkSec / maxTalk, PALETTE[i % PALETTE.length])
    ).join('');

    const engBars = p.map((x, i) =>
      bar(x.name, String(x.engagement), x.engagement / maxEng, PALETTE[i % PALETTE.length])
    ).join('');

    const contribRows = p.map((x) =>
`<tr><td>${escapeHtml(x.name)}</td><td>${secToClock(x.talkSec)}</td><td>${Math.round(x.talkPct * 100)}%</td><td>${x.turns}</td><td>${x.words.toLocaleString()}</td><td>${x.questions}</td><td>${x.engagement}</td><td>${x.collabScore}</td></tr>`
    ).join('');

    // Per-person scorecards: relative activity signals (0–100 vs the top person).
    const maxOf = (f) => Math.max(1, ...p.map(f));
    const mTalk = maxOf(x => x.talkSec), mTurns = maxOf(x => x.turns),
      mQ = maxOf(x => x.questions), mResp = maxOf(x => x.responses), mCollab = maxOf(x => x.collabRaw);
    const DIMS = [
      ['Presence', x => x.talkSec / mTalk],
      ['Turns', x => x.turns / mTurns],
      ['Inquiry', x => x.questions / mQ],
      ['Responsiveness', x => x.responses / mResp],
      ['Collaboration', x => x.collabRaw / mCollab]
    ];
    const scards = p.map((x, i) => {
      const rows = DIMS.map(([label, f]) => {
        const v = Math.round(Math.max(0, Math.min(1, f(x))) * 100);
        return `<div class="srow"><div class="l">${label}</div><div class="track"><div class="fill" style="width:${v}%;background:${PALETTE[i % PALETTE.length]}"></div></div><div class="v">${v}</div></div>`;
      }).join('');
      return `<div class="scard"><div class="scard-h"><span class="nm">${escapeHtml(x.name)}</span><span class="eng">Engagement ${x.engagement}</span></div>${rows}</div>`;
    }).join('');

    const transRows = insights.topTransitions.length
      ? insights.topTransitions.map(t => `<li><span>${escapeHtml(t.pair)}</span><b>${t.count}</b></li>`).join('')
      : '<li><span>Not enough turns to map interactions</span><b>—</b></li>';

    const timeline = buildTimelineSvg(insights, cues);
    const promptText = buildLlmPrompt(cues, title);

    return (
`<h1>Meeting insights</h1>
<div class="sub">${escapeHtml(title)} · ${insights.lineCount} transcript lines · generated locally in your browser</div>

<h2>Meeting health</h2>
<div class="tiles">${tiles}</div>

<div class="cols">
  <div><h2>Talk-time distribution</h2><div class="card">${talkBars}</div></div>
  <div><h2>Engagement share</h2><div class="card">${engBars}<div class="note">0.5·talk + 0.3·turns + 0.2·questions (relative, heuristic).</div></div></div>
</div>

<h2>Contribution breakdown</h2>
<div class="card" style="overflow-x:auto">
<table>
<thead><tr><th>Participant</th><th>Talk</th><th>%</th><th>Turns</th><th>Words</th><th>Questions</th><th>Engagement</th><th>Collab*</th></tr></thead>
<tbody>${contribRows}</tbody>
</table>
<div class="note">*Collaboration proxy (relative 0–100): questions asked + short agreements + responses to others. A heuristic, not a judgement of quality.</div>
</div>

<h2>Participation timeline</h2>
<div class="card">${timeline}</div>

<h2>Contribution scores per person</h2>
<div class="scards">${scards}</div>
<div class="note">Relative activity signals (0–100 vs the most active participant): Presence = talk time, Turns = speaking turns, Inquiry = questions asked, Responsiveness = replies to others, Collaboration = questions + agreements + replies. Heuristics from the transcript — not a performance rating. For qualitative scoring (leadership, ownership…) use the AI prompt below.</div>

<h2>Who responds to whom</h2>
<div class="card"><ul class="trans">${transRows}</ul></div>

<h2>AI analysis (summary · decisions · action items · sentiment · styles · rubric)</h2>
<div class="card">
  <p class="note" style="margin-top:0">These qualitative analyses need an LLM. Nothing was sent anywhere — copy the prompt below into the AI tool your company permits (Copilot, Claude, ChatGPT Enterprise…) and paste its answer back.</p>
  <button class="copy" id="copyBtn">Copy prompt</button>
  <textarea id="llmPrompt" readonly>${escapeHtml(promptText)}</textarea>
  <div class="disc"><b>Note on the per-person rubric:</b> it is an AI-generated reflection to support growth — not a formal performance review, and should never be used as one on its own. Metrics here (talk time, turns, questions) are objective counts; scores like engagement/collaboration are rough heuristics.</div>
</div>

<div class="foot">Generated by Teams Recording Downloader · all metrics computed on-device from the transcript · timestamps are second-resolution.</div>`);
  }

  // Full standalone document for the "Download HTML" option. The embedded
  // <script> deliberately avoids backticks and ${} so it survives this template.
  function buildDashboardHtml(insights, cues, title) {
    return (
`<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Meeting insights — ${escapeHtml(title)}</title>
<style>${DASHBOARD_CSS}</style></head><body><div class="wrap">${dashboardBody(insights, cues, title)}</div>
<script>
(function(){
  var b=document.getElementById('copyBtn'),t=document.getElementById('llmPrompt');
  b.addEventListener('click',function(){
    t.focus();t.select();
    var done=function(){b.textContent='Copied ✓';setTimeout(function(){b.textContent='Copy prompt';},1800);};
    if(navigator.clipboard&&navigator.clipboard.writeText){navigator.clipboard.writeText(t.value).then(done,function(){document.execCommand('copy');done();});}
    else{document.execCommand('copy');done();}
  });
})();
</script>
</body></html>`);
  }

  // ===========================================================================
  // In-page insights overlay (Shadow DOM — style-isolated from the host page).
  // ===========================================================================
  let _ttPolicy = null, _ttTried = false;
  function safeSetHTML(el, html) {
    try { el.innerHTML = html; return true; } catch (_) { /* Trusted Types */ }
    try {
      if (window.trustedTypes && window.trustedTypes.createPolicy) {
        if (!_ttTried) {
          _ttTried = true;
          try { _ttPolicy = window.trustedTypes.createPolicy('teamsdl-insights', { createHTML: (s) => s }); }
          catch (_) { _ttPolicy = null; }
        }
        if (_ttPolicy) { el.innerHTML = _ttPolicy.createHTML(html); return true; }
      }
    } catch (_) { /* policy disallowed by page CSP */ }
    return false;
  }

  let insightsOverlayEl = null;
  function onOverlayKey(e) { if (e.key === 'Escape') removeInsightsOverlay(); }
  function removeInsightsOverlay() {
    if (!insightsOverlayEl) return;
    insightsOverlayEl.remove();
    insightsOverlayEl = null;
    document.removeEventListener('keydown', onOverlayKey, true);
  }

  function openInsightsOverlay(insights, cues, title) {
    removeInsightsOverlay();
    const host = document.createElement('div');
    host.style.cssText = 'position:fixed;inset:0;z-index:2147483647';
    const root = host.attachShadow({ mode: 'open' });

    const style = document.createElement('style');
    style.textContent = DASHBOARD_CSS +
`\n.ovl-scrim{position:fixed;inset:0;background:rgba(0,0,0,.55)}
.ovl-panel{position:fixed;inset:0;background:var(--bg);color:var(--text);font:14px/1.5 "Segoe UI",system-ui,sans-serif;display:flex;flex-direction:column;overflow:hidden}
.ovl-bar{display:flex;align-items:center;gap:10px;padding:12px 16px;border-bottom:1px solid var(--line);background:var(--card)}
.ovl-bar .t{font-weight:700;font-size:14px;flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.ovl-bar button{background:var(--card2);color:var(--text);border:1px solid var(--line);border-radius:8px;padding:8px 12px;font:600 13px "Segoe UI",system-ui,sans-serif;cursor:pointer}
.ovl-scroll{overflow:auto;flex:1}
.ovl-scroll .wrap{padding:16px 24px 40px}`;
    root.appendChild(style);

    const scrim = document.createElement('div');
    scrim.className = 'ovl-scrim';
    scrim.addEventListener('click', removeInsightsOverlay);
    root.appendChild(scrim);

    const panel = document.createElement('div');
    panel.className = 'ovl-panel';
    const barEl = document.createElement('div');
    barEl.className = 'ovl-bar';
    const tEl = document.createElement('div'); tEl.className = 't'; tEl.textContent = 'Meeting insights — ' + title;
    const dlBtn = document.createElement('button'); dlBtn.textContent = '⤓ Download HTML';
    const closeBtn = document.createElement('button'); closeBtn.textContent = '✕ Close';
    barEl.appendChild(tEl); barEl.appendChild(dlBtn); barEl.appendChild(closeBtn);
    panel.appendChild(barEl);

    const scroll = document.createElement('div'); scroll.className = 'ovl-scroll';
    const wrap = document.createElement('div'); wrap.className = 'wrap';
    if (!safeSetHTML(wrap, dashboardBody(insights, cues, title))) return false; // TT blocked → caller falls back
    scroll.appendChild(wrap);
    panel.appendChild(scroll);
    root.appendChild(panel);

    document.documentElement.appendChild(host);
    insightsOverlayEl = host;
    document.addEventListener('keydown', onOverlayKey, true);

    closeBtn.addEventListener('click', removeInsightsOverlay);
    dlBtn.addEventListener('click', () => {
      const safe = title.replace(/[^a-z0-9\s_-]/gi, '_').trim() || 'teams_meeting';
      downloadFile(buildDashboardHtml(insights, cues, title), safe + '_insights.html', 'text/html;charset=utf-8');
    });
    const copyBtn = root.querySelector('#copyBtn');
    const ta = root.querySelector('#llmPrompt');
    if (copyBtn && ta) {
      copyBtn.addEventListener('click', () => {
        ta.focus(); ta.select();
        const done = () => { copyBtn.textContent = 'Copied ✓'; setTimeout(() => { copyBtn.textContent = 'Copy prompt'; }, 1800); };
        if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(ta.value).then(done, () => { document.execCommand('copy'); done(); });
        else { document.execCommand('copy'); done(); }
      });
    }
    return true;
  }

  // ===========================================================================
  // Ask panel — docked chat/search over the transcript. Local keyword search
  // when no provider is configured; streamed Q&A via the service worker when
  // one is. The API key lives only in the options page + service worker.
  // ===========================================================================
  let askHost = null, askOpen = false, askCuesPromise = null;

  // Reads NON-SECRET config only. The API key is never read here.
  async function getProviderConfig() {
    try {
      const { aiProvider } = await chrome.storage.local.get('aiProvider');
      if (!aiProvider || !aiProvider.connected) return null;
      return { connected: true, host: aiProvider.host, model: aiProvider.model, budget: aiProvider.budget || 48000 };
    } catch (_) { return null; }
  }

  const ASK_CSS = `
.ask-panel{position:fixed;top:0;right:0;bottom:0;width:380px;max-width:92vw;z-index:2147483647;background:#1f1f23;color:#eee;font:14px/1.5 "Segoe UI",system-ui,sans-serif;display:flex;flex-direction:column;box-shadow:-4px 0 18px rgba(0,0,0,.4)}
.ask-bar{display:flex;align-items:center;gap:8px;padding:12px 14px;border-bottom:1px solid #333;background:#2a2a30}
.ask-bar .t{font-weight:700;flex:1}
.ask-bar button{background:#3a3a42;color:#eee;border:0;border-radius:6px;padding:6px 10px;cursor:pointer;font:600 13px "Segoe UI",system-ui,sans-serif}
.ask-lock{font-size:12px;color:#9ad0b0;padding:6px 14px;border-bottom:1px solid #333;display:none}
.ask-log{flex:1;overflow:auto;padding:14px;display:flex;flex-direction:column;gap:10px}
.ask-msg{padding:8px 11px;border-radius:10px;max-width:92%;white-space:pre-wrap;word-wrap:break-word}
.ask-msg.user{align-self:flex-end;background:#3b4a7a}
.ask-msg.assistant{align-self:flex-start;background:#2f2f36}
.ask-msg.error{align-self:flex-start;background:#5a2320;color:#ffd9d5}
.ask-msg.system{align-self:center;background:transparent;color:#9a9aa2;font-size:12px;text-align:center}
.ask-hit{background:#2f2f36;border-radius:8px;padding:8px 10px;cursor:pointer}
.ask-hit:hover{background:#3a3a42}
.ask-hit .m{color:#9ad0b0;font-size:12px}
.ask-hit b{color:#ffe08a}
.ask-cta{background:#2f6f4f;color:#fff;border:0;border-radius:8px;padding:9px 12px;cursor:pointer;font:600 13px "Segoe UI",system-ui,sans-serif;align-self:flex-start}
.ask-form{display:flex;gap:8px;padding:12px 14px;border-top:1px solid #333;background:#2a2a30}
.ask-form input{flex:1;background:#1f1f23;color:#eee;border:1px solid #444;border-radius:8px;padding:9px 11px;font:inherit}
.ask-form button{background:#6264a7;color:#fff;border:0;border-radius:8px;padding:9px 14px;cursor:pointer;font:600 13px "Segoe UI",system-ui,sans-serif}`;

  function ensureAskPanel() {
    if (askHost) return askHost._handles;
    const host = document.createElement('div');
    const root = host.attachShadow({ mode: 'open' });
    const style = document.createElement('style');
    style.textContent = ASK_CSS;
    root.appendChild(style);

    const panel = document.createElement('div');
    panel.className = 'ask-panel';

    const bar = document.createElement('div');
    bar.className = 'ask-bar';
    const title = document.createElement('div'); title.className = 't'; title.textContent = '💬 Ask this meeting';
    const settingsBtn = document.createElement('button'); settingsBtn.textContent = '⚙︎';  settingsBtn.title = 'Provider settings';
    const closeBtn = document.createElement('button'); closeBtn.textContent = '✕';
    bar.appendChild(title); bar.appendChild(settingsBtn); bar.appendChild(closeBtn);

    const lock = document.createElement('div'); lock.className = 'ask-lock';
    const log = document.createElement('div'); log.className = 'ask-log';

    const form = document.createElement('form'); form.className = 'ask-form';
    const input = document.createElement('input'); input.type = 'text'; input.placeholder = 'Search the transcript…';
    const sendBtn = document.createElement('button'); sendBtn.type = 'submit'; sendBtn.textContent = 'Send';
    form.appendChild(input); form.appendChild(sendBtn);

    panel.appendChild(bar); panel.appendChild(lock); panel.appendChild(log); panel.appendChild(form);
    root.appendChild(panel);
    document.documentElement.appendChild(host);

    settingsBtn.addEventListener('click', () => chrome.runtime.sendMessage({ type: 'openOptions' }));
    closeBtn.addEventListener('click', () => toggleAskPanel(false));

    askHost = host;
    host._handles = { root, log, input, form, lock };
    return host._handles;
  }

  function appendMsg(root, cls, text) {
    const log = root.querySelector('.ask-log');
    const el = document.createElement('div');
    el.className = 'ask-msg ' + cls;
    el.textContent = text;
    log.appendChild(el);
    log.scrollTop = log.scrollHeight;
    return el;
  }

  // Best-effort: scroll the live Teams transcript to a cue's position.
  function jumpToCue(cue) {
    const sub = document.querySelector('[id^="sub-entry-"][aria-posinset="' + cue.pos + '"]');
    const target = sub ? sub.closest('div[role="group"]') || sub : null;
    if (target && target.scrollIntoView) target.scrollIntoView({ behavior: 'smooth', block: 'center' });
  }

  function renderSearchResults(root, cues, query) {
    const log = root.querySelector('.ask-log');
    const hits = ChatLib.searchCues(cues, query, 20);
    if (!hits.length) { appendMsg(root, 'system', 'No transcript lines match “' + query + '”.'); return; }
    const terms = ChatLib.tokenize(query);
    for (const h of hits) {
      const div = document.createElement('div');
      div.className = 'ask-hit';
      const meta = document.createElement('div'); meta.className = 'm';
      meta.textContent = secToClock(h.cue.start) + ' · ' + (h.cue.speaker || 'Unknown');
      const body = document.createElement('div');
      // Highlight matched terms without raw innerHTML injection.
      const words = String(h.cue.text).split(/(\s+)/);
      for (const w of words) {
        const bare = w.toLowerCase().replace(/[^a-z0-9]/g, '');
        if (bare && terms.indexOf(bare) !== -1) { const b = document.createElement('b'); b.textContent = w; body.appendChild(b); }
        else body.appendChild(document.createTextNode(w));
      }
      div.appendChild(meta); div.appendChild(body);
      div.addEventListener('click', () => jumpToCue(h.cue));
      log.appendChild(div);
    }
    log.scrollTop = log.scrollHeight;
  }

  async function loadAskCues() {
    if (!askCuesPromise) askCuesPromise = getCues(() => {}, undefined).catch((e) => { askCuesPromise = null; throw e; });
    return askCuesPromise;
  }

  async function handleAsk(handles, query) {
    appendMsg(handles.root, 'user', query);
    let cues;
    try { cues = await loadAskCues(); }
    catch (_) { appendMsg(handles.root, 'error', 'Could not read the transcript. Open the transcript panel (CC) first.'); return; }
    // Task 5 adds the connected branch here. For now: always local search.
    renderSearchResults(handles.root, cues, query);
  }

  async function toggleAskPanel(force) {
    const open = typeof force === 'boolean' ? force : !askOpen;
    askOpen = open;
    if (!open) { if (askHost) askHost.style.display = 'none'; return; }
    const handles = ensureAskPanel();
    askHost.style.display = '';
    const cfg = await getProviderConfig();
    if (cfg) {
      handles.input.placeholder = 'Ask about this meeting…';
      handles.lock.style.display = 'block';
      handles.lock.textContent = '🔒 Connected — questions are sent to ' + cfg.host;
    } else {
      handles.input.placeholder = 'Search the transcript…';
      handles.lock.style.display = 'none';
      if (!handles.log.childElementCount) {
        appendMsg(handles.root, 'system', 'No AI provider connected — running local keyword search. Connect one for full Q&A.');
        const cta = document.createElement('button');
        cta.className = 'ask-cta'; cta.textContent = 'Connect a provider';
        cta.addEventListener('click', () => chrome.runtime.sendMessage({ type: 'openOptions' }));
        handles.log.appendChild(cta);
      }
    }
    if (!handles.form._wired) {
      handles.form._wired = true;
      handles.form.addEventListener('submit', (e) => {
        e.preventDefault();
        const q = handles.input.value.trim();
        if (!q) return;
        handles.input.value = '';
        handleAsk(handles, q);
      });
    }
    handles.input.focus();
  }

  // ===========================================================================
  // UI
  // ===========================================================================
  let btn, transcriptBtn, insightsBtn, askBtn, statusEl, wrapEl;

  // Inline SVG line icons (Feather-style), built via the DOM API rather than
  // innerHTML so Trusted Types CSP on Teams/SharePoint can't block them.
  const SVGNS = 'http://www.w3.org/2000/svg';
  const ICONS = {
    download: [['path', { d: 'M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4' }], ['polyline', { points: '7 10 12 15 17 10' }], ['line', { x1: '12', y1: '15', x2: '12', y2: '3' }]],
    barChart: [['line', { x1: '18', y1: '20', x2: '18', y2: '10' }], ['line', { x1: '12', y1: '20', x2: '12', y2: '4' }], ['line', { x1: '6', y1: '20', x2: '6', y2: '14' }]],
    cancel: [['line', { x1: '18', y1: '6', x2: '6', y2: '18' }], ['line', { x1: '6', y1: '6', x2: '18', y2: '18' }]],
    chat: [['path', { d: 'M21 11.5a8.38 8.38 0 0 1-.9 3.8 8.5 8.5 0 0 1-7.6 4.7 8.38 8.38 0 0 1-3.8-.9L3 21l1.9-5.7a8.38 8.38 0 0 1-.9-3.8 8.5 8.5 0 0 1 4.7-7.6 8.38 8.38 0 0 1 3.8-.9h.5a8.48 8.48 0 0 1 8 8v.5z' }]]
  };
  function makeIcon(children) {
    const svg = document.createElementNS(SVGNS, 'svg');
    const base = { viewBox: '0 0 24 24', width: '15', height: '15', fill: 'none', stroke: 'currentColor', 'stroke-width': '2', 'stroke-linecap': 'round', 'stroke-linejoin': 'round', 'aria-hidden': 'true' };
    for (const k in base) svg.setAttribute(k, base[k]);
    svg.style.cssText = 'flex:none';
    for (const [tag, attrs] of children) {
      const el = document.createElementNS(SVGNS, tag);
      for (const k in attrs) el.setAttribute(k, attrs[k]);
      svg.appendChild(el);
    }
    return svg;
  }
  const CANCEL_BG = '#b3352f';
  function makeToolButton(icon, label, bg, cancelLabel) {
    const b = document.createElement('button');
    b.style.cssText = 'display:none;align-items:center;gap:8px;background:' + bg + ';color:#fff;border:0;padding:10px 16px;border-radius:8px;font:600 14px/1 "Segoe UI",system-ui,sans-serif;cursor:pointer;box-shadow:0 2px 8px rgba(0,0,0,.3)';
    const iconEl = makeIcon(icon);
    b.appendChild(iconEl);
    const span = document.createElement('span');
    span.textContent = label;
    b.appendChild(span);
    b._iconEl = iconEl; b._span = span;
    b._baseIcon = icon; b._baseLabel = label; b._baseBg = bg;
    b._cancelLabel = cancelLabel;
    return b;
  }

  // While an action runs its button turns into a red "Cancel" (same button);
  // clicking again aborts, and it reverts to its original icon/label when done.
  function setButtonRunning(button, running) {
    const icon = running ? ICONS.cancel : button._baseIcon;
    const fresh = makeIcon(icon);
    button.replaceChild(fresh, button._iconEl);
    button._iconEl = fresh;
    button._span.textContent = running ? button._cancelLabel : button._baseLabel;
    button.style.background = running ? CANCEL_BG : button._baseBg;
    button.title = running ? 'Cancel' : '';
  }

  // Wire a toggle action: first click starts runFn(signal); while running the
  // button is a Cancel toggle. `precheck` (optional) may veto the start and set
  // its own status. Buttons run independently, so actions can overlap.
  function makeAction(button, runFn, precheck) {
    button.addEventListener('click', async () => {
      if (button._controller) { button._controller.abort(); return; }
      if (precheck && !precheck()) return;
      const controller = new AbortController();
      button._controller = controller;
      setButtonRunning(button, true);
      try { await runFn(controller.signal); }
      finally { button._controller = null; setButtonRunning(button, false); }
    });
  }

  function injectButton() {
    if (btn) return;
    const wrap = document.createElement('div');
    wrap.style.cssText = 'position:fixed;bottom:20px;right:20px;z-index:2147483647;font-family:Segoe UI,system-ui,sans-serif;display:none;flex-direction:column;align-items:flex-end;gap:6px;';
    statusEl = document.createElement('div');
    statusEl.style.cssText = 'background:#222;color:#fff;padding:4px 10px;border-radius:6px;font-size:12px;max-width:280px;display:none;';
    btn = makeToolButton(ICONS.download, 'Download recording', '#6264a7', 'Cancel recording');
    transcriptBtn = makeToolButton(ICONS.download, 'Download transcript', '#464775', 'Cancel transcript');
    insightsBtn = makeToolButton(ICONS.barChart, 'Meeting insights', '#2f6f4f', 'Cancel insights');
    askBtn = makeToolButton(ICONS.chat, 'Ask this meeting', '#6264a7', 'Ask this meeting');
    askBtn.addEventListener('click', () => {
      if (!isTranscriptPanelPresent()) { setStatus('Open the transcript panel (Transcript / CC) first, then click again.'); return; }
      toggleAskPanel();
    });
    makeAction(btn, runVideoDownload, () => {
      if (videoManifestUrl) return true;
      setStatus('Play the recording for a few seconds first.'); return false;
    });
    makeAction(transcriptBtn, runTranscript, () => {
      if (isTranscriptPanelPresent()) return true;
      setStatus('Open the transcript panel (Transcript / CC) first, then click again.'); return false;
    });
    makeAction(insightsBtn, runInsights, () => {
      if (isTranscriptPanelPresent()) return true;
      setStatus('Open the transcript panel (Transcript / CC) first, then click again.'); return false;
    });
    wrap.appendChild(statusEl); wrap.appendChild(askBtn); wrap.appendChild(insightsBtn); wrap.appendChild(transcriptBtn); wrap.appendChild(btn);
    (document.body || document.documentElement).appendChild(wrap);
    wrapEl = wrap;
  }
  function showButton() {
    injectButton();
    btn.style.display = 'inline-flex';
    wrapEl.style.display = 'flex';
  }
  function showTranscriptButton() {
    injectButton();
    transcriptBtn.style.display = 'inline-flex';
    insightsBtn.style.display = 'inline-flex';
    askBtn.style.display = 'inline-flex';
    wrapEl.style.display = 'flex';
  }
  function hideButton() {
    if (wrapEl) wrapEl.style.display = 'none';
  }
  function setStatus(text) {
    if (!statusEl) return;
    statusEl.textContent = text;
    statusEl.style.display = text ? 'block' : 'none';
  }

  async function runVideoDownload(signal) {
    const filename = (document.title || 'teams_recording').replace(/[^a-z0-9\s_-]/gi, '_').trim() || 'teams_recording';
    const onProgress = (done, total, text) => setStatus(text || `${done}/${total}`);
    try {
      await runDownload(filename, onProgress, signal);
      setStatus('Recording saved to your Downloads.');
    } catch (err) {
      console.error('[Teams DL] download failed', err);
      if (err.name === 'AbortError') setStatus('Recording download cancelled.');
      else if (err.isDrm) setStatus('This recording uses hard DRM (Widevine/PlayReady) — cannot be decrypted client-side.');
      else setStatus('Recording failed: ' + (err.message || err));
    }
  }

  async function runTranscript(signal) {
    const filename = (document.title || 'teams_transcript').replace(/[^a-z0-9\s_-]/gi, '_').trim() || 'teams_transcript';
    const onProgress = (done, total, text) => setStatus(text || `${done}/${total}`);
    try {
      await runTranscriptDownload(filename, onProgress, signal);
      setStatus('Transcript saved (.vtt) to your Downloads.');
    } catch (err) {
      console.error('[Teams DL] transcript failed', err);
      if (err.name === 'AbortError') setStatus('Transcript cancelled.');
      else if (err.message === 'NO_TRANSCRIPT') setStatus('No transcript found. Open the transcript panel (CC) first, then retry.');
      else setStatus('Transcript failed: ' + (err.message || err));
    }
  }

  async function runInsights(signal) {
    const title = (document.title || 'Teams meeting').replace(/\s+/g, ' ').trim() || 'Teams meeting';
    const safe = title.replace(/[^a-z0-9\s_-]/gi, '_').trim() || 'teams_meeting';
    const onProgress = (done, total, text) => setStatus(text || `${done}/${total}`);
    try {
      const cues = await getCues(onProgress, signal);
      setStatus('Building insights dashboard…');
      const insights = computeInsights(cues);
      if (openInsightsOverlay(insights, cues, title)) {
        setStatus(`Insights ready (${insights.parts.length} participants, ${cues.length} lines).`);
      } else {
        // In-page injection blocked by page security (Trusted Types) — fall back
        // to a downloaded file opened in a new tab.
        const html = buildDashboardHtml(insights, cues, title);
        downloadFile(html, safe + '_insights.html', 'text/html;charset=utf-8');
        try { window.open(URL.createObjectURL(new Blob([html], { type: 'text/html' })), '_blank'); } catch (_) {}
        setStatus('In-page view blocked by page security — insights saved to Downloads instead.');
      }
    } catch (err) {
      console.error('[Teams DL] insights failed', err);
      if (err.name === 'AbortError') setStatus('Insights cancelled.');
      else if (err.message === 'NO_TRANSCRIPT') setStatus('No transcript found. Open the transcript panel (CC) first, then retry.');
      else setStatus('Insights failed: ' + (err.message || err));
    }
  }

  // Only the SharePoint Stream player page (…/_layouts/15/stream.aspx, also on
  // sharepoint-df.com) gets the instant placeholder — we don't want a floating
  // button on every SharePoint doc/list. Everywhere else the button still
  // appears the moment a recording is actually detected (capture-based reveal).
  function isStreamPage() {
    try {
      return /https:\/\/[^/]*sharepoint(?:-df)?\.com\/.*\/_layouts\/15\/stream\.aspx/i.test(location.href);
    } catch (_) {
      return false;
    }
  }

  // Surface the overlay button in the top frame immediately on Stream pages,
  // before any manifest is captured. Clicking it before playback prompts the
  // user to play the recording first; once a manifest is seen (here or in a
  // child frame) the button is wired to the real download.
  // Reveal the transcript button as soon as the transcript panel is rendered in
  // this frame (the panel loads lazily, after the user opens it), and re-check on
  // DOM mutations. The video button stays capture-driven; this is independent.
  function watchForTranscriptPanel() {
    if (isTranscriptPanelPresent()) { showTranscriptButton(); return; }
    let done = false;
    const obs = new MutationObserver(() => {
      if (done) return;
      if (isTranscriptPanelPresent()) { done = true; obs.disconnect(); showTranscriptButton(); }
    });
    try { obs.observe(document.documentElement, { childList: true, subtree: true }); } catch (_) {}
  }

  function ready() {
    injectButton();
    if (isTopFrame && isStreamPage()) showButton();
    watchForTranscriptPanel();
  }
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', ready);
  } else {
    ready();
  }
})();
