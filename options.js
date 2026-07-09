'use strict';

const $ = (id) => document.getElementById(id);
const statusEl = $('status');

function setStatus(msg, ok) {
  statusEl.textContent = msg;
  statusEl.className = 'status ' + (ok ? 'ok' : 'err');
}

function originOf(url) {
  try { return new URL(url).origin; } catch (_) { return null; }
}

function normalizeBaseUrl(url) {
  return String(url || '').trim().replace(/\/+$/, '');
}

async function load() {
  const { aiProvider } = await chrome.storage.local.get('aiProvider');
  if (!aiProvider) return;
  $('baseUrl').value = aiProvider.baseUrl || '';
  $('apiKey').value = aiProvider.apiKey || '';
  $('model').value = aiProvider.model || '';
  $('budget').value = aiProvider.budget || 48000;
  if (aiProvider.connected) setStatus('Connected to ' + aiProvider.host, true);
}

function readForm() {
  const baseUrl = normalizeBaseUrl($('baseUrl').value);
  const host = originOf(baseUrl);
  const apiKey = $('apiKey').value.trim();
  const model = $('model').value.trim();
  const budget = Math.max(4000, parseInt($('budget').value, 10) || 48000);
  return { baseUrl, host, apiKey, model, budget };
}

async function save() {
  const cfg = readForm();
  if (!cfg.host) return setStatus('Enter a valid https Base URL.', false);
  if (!cfg.apiKey) return setStatus('Enter an API key.', false);
  if (!cfg.model) return setStatus('Enter a model name.', false);
  let granted = true;
  try {
    granted = await chrome.permissions.request({ origins: [cfg.host + '/*'] });
  } catch (e) {
    return setStatus('Permission request failed: ' + e.message, false);
  }
  if (!granted) return setStatus('Host permission denied — cannot reach this endpoint.', false);
  await chrome.storage.local.set({ aiProvider: { ...cfg, connected: true } });
  setStatus('Saved and authorized for ' + cfg.host, true);
}

// The options page is an extension context, so once the host permission is
// granted it can fetch the endpoint directly (no service worker needed here).
async function test() {
  const cfg = readForm();
  if (!cfg.host || !cfg.apiKey || !cfg.model) return setStatus('Fill in all fields first.', false);
  const has = await chrome.permissions.contains({ origins: [cfg.host + '/*'] });
  if (!has) return setStatus('Click "Save & authorize" first to grant access to ' + cfg.host, false);
  setStatus('Testing…', true);
  try {
    const resp = await fetch(cfg.baseUrl + '/chat/completions', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', Authorization: 'Bearer ' + cfg.apiKey },
      body: JSON.stringify({ model: cfg.model, messages: [{ role: 'user', content: 'ping' }], max_tokens: 1 }),
    });
    if (resp.ok) setStatus('Connection OK (HTTP ' + resp.status + ').', true);
    else setStatus('Endpoint returned HTTP ' + resp.status + ' — check key/model/URL.', false);
  } catch (e) {
    setStatus('Request failed: ' + e.message + ' (verify URL and that it allows this extension).', false);
  }
}

async function disconnect() {
  const { aiProvider } = await chrome.storage.local.get('aiProvider');
  if (aiProvider && aiProvider.host) {
    try { await chrome.permissions.remove({ origins: [aiProvider.host + '/*'] }); } catch (_) {}
  }
  await chrome.storage.local.remove('aiProvider');
  $('apiKey').value = '';
  setStatus('Disconnected. The Ask panel will use local search only.', true);
}

$('save').addEventListener('click', save);
$('test').addEventListener('click', test);
$('disconnect').addEventListener('click', disconnect);
load();
