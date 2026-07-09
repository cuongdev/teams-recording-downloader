'use strict';
importScripts('chat-lib.js'); // provides self.ChatLib

const SYSTEM_PROMPT =
  'You are a helpful assistant answering questions about a single meeting. ' +
  'Base your answers only on the transcript provided. If the transcript does ' +
  'not contain the answer, say so. Cite speaker names and [seconds] timestamps ' +
  'when relevant.';

chrome.runtime.onMessage.addListener((msg) => {
  if (msg && msg.type === 'openOptions') chrome.runtime.openOptionsPage();
});

chrome.runtime.onConnect.addListener((port) => {
  if (port.name !== 'chat') return;
  port.onMessage.addListener((msg) => { handleChat(port, msg); });
});

async function handleChat(port, msg) {
  let cfg, apiKey, consented;
  try {
    const store = await chrome.storage.local.get(['aiProvider', 'aiApiKey', 'aiConsented']);
    cfg = store.aiProvider;
    apiKey = store.aiApiKey;
    consented = store.aiConsented;
  } catch (e) {
    return post(port, { type: 'error', message: 'Could not read settings.' });
  }
  if (!cfg || !cfg.connected || !apiKey) {
    return post(port, { type: 'error', message: 'No provider connected. Open Settings to connect one.' });
  }
  // Defense in depth: never send transcript text without the user's recorded
  // consent, even if some future caller opens the port directly. The panel sets
  // this only after the consent dialog (content.js askConsent).
  if (!consented) {
    return post(port, { type: 'error', message: 'Consent required — ask again from the panel.' });
  }
  const has = await chrome.permissions.contains({ origins: [cfg.host + '/*'] });
  if (!has) {
    return post(port, { type: 'error', message: 'Re-authorize your provider in Settings (host access was removed).' });
  }

  const body = ChatLib.buildChatRequest(cfg.model, SYSTEM_PROMPT, msg.history || [], msg.question, msg.contextText || '');

  let resp;
  try {
    resp = await fetch(cfg.baseUrl + '/chat/completions', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', Authorization: 'Bearer ' + apiKey },
      body: JSON.stringify(body),
    });
  } catch (e) {
    return post(port, { type: 'error', message: "Couldn't reach the endpoint. Verify the URL and that it allows this extension." });
  }

  if (!resp.ok) {
    const map = { 401: 'Check your API key.', 403: 'Check your API key.', 404: 'Check the model name and Base URL.', 400: 'Check the model name and Base URL.', 429: 'Provider rate-limited the request — try again.' };
    let detail = map[resp.status] || ('Provider returned HTTP ' + resp.status + '.');
    return post(port, { type: 'error', message: detail });
  }

  if (!resp.body) {
    return post(port, { type: 'error', message: 'Provider returned an empty response.' });
  }

  // Stream the SSE body, decoding incrementally and forwarding tokens.
  const reader = resp.body.getReader();
  const decoder = new TextDecoder();
  let buffer = '';
  let sawToken = false;
  try {
    for (;;) {
      const { value, done } = await reader.read();
      if (done) break;
      buffer += decoder.decode(value, { stream: true });
      const parsed = ChatLib.parseSSE(buffer);
      buffer = parsed.rest;
      for (const t of parsed.tokens) { sawToken = true; post(port, { type: 'token', text: t }); }
      if (parsed.done) break;
    }
  } catch (e) {
    return post(port, { type: 'error', message: 'Stream interrupted: ' + e.message });
  }
  if (!sawToken) return post(port, { type: 'error', message: 'Provider returned an unexpected (non-streaming) response.' });
  post(port, { type: 'done' });
}

function post(port, msg) {
  try { port.postMessage(msg); } catch (_) { /* panel closed */ }
}
