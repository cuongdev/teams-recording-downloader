// Pure, dependency-free logic shared by the content-script panel (browser) and
// the service worker, and unit-tested under Node. No chrome.* / DOM access here.
(function (root, factory) {
  const lib = factory();
  if (typeof module !== 'undefined' && module.exports) module.exports = lib;
  if (root) root.ChatLib = lib;
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const WORD_RE = /[a-z0-9]+/gi;

  function tokenize(text) {
    return String(text || '').toLowerCase().match(WORD_RE) || [];
  }

  // Rank cues by query relevance. Term coverage (+1/term) + frequency (+0.1)
  // + exact-phrase bonus (+3). Positive scores only, score desc, ties by index.
  function searchCues(cues, query, limit) {
    limit = limit || 20;
    const terms = tokenize(query);
    if (!terms.length) return [];
    const uniq = [...new Set(terms)];
    const phrase = String(query || '').toLowerCase().trim();
    const results = [];
    for (let i = 0; i < cues.length; i++) {
      const text = String(cues[i].text || '');
      const lc = text.toLowerCase();
      const toks = tokenize(text);
      const tokSet = new Set(toks);
      let score = 0;
      for (const t of uniq) if (tokSet.has(t)) score += 1;
      if (score === 0) continue;
      for (const t of toks) if (uniq.indexOf(t) !== -1) score += 0.1;
      if (phrase.length > 2 && lc.indexOf(phrase) !== -1) score += 3;
      results.push({ cue: cues[i], index: i, score });
    }
    results.sort((a, b) => b.score - a.score || a.index - b.index);
    return results.slice(0, limit);
  }

  function lineOf(cue) {
    return '[' + (cue.start | 0) + 's] ' + (cue.speaker || 'Unknown') + ': ' + cue.text;
  }

  function cuesToPlainText(cues) {
    return cues.map(lineOf).join('\n');
  }

  // Smart hybrid: whole transcript if it fits `budget` chars; else the
  // top-matching lines (hit first, then neighbors) packed to budget, output in
  // transcript order with "…" marking skipped stretches.
  function buildContext(cues, query, budget) {
    budget = budget || 48000;
    const full = cuesToPlainText(cues);
    if (full.length <= budget) return full;
    const hits = searchCues(cues, query, 1000);
    const keep = new Set();
    let size = 0;
    for (const h of hits) {
      for (const j of [h.index, h.index - 1, h.index + 1]) {
        if (j < 0 || j >= cues.length || keep.has(j)) continue;
        // +1 for this line's own '\n' join, +2 reserved for a possible '…\n'
        // gap separator the assembly loop below may insert before it. The
        // number of gaps is always < the number of kept lines, so reserving
        // 2 per line is a safe upper bound — guarantees the assembled
        // output never exceeds `budget` even with multiple disjoint clusters.
        const len = lineOf(cues[j]).length + 3;
        if (size + len > budget) continue;
        keep.add(j);
        size += len;
      }
      if (size >= budget) break;
    }
    if (!keep.size) return full.slice(0, budget);
    const ordered = [...keep].sort((a, b) => a - b);
    const lines = [];
    let prev = -2;
    for (const idx of ordered) {
      if (idx !== prev + 1 && lines.length) lines.push('…');
      lines.push(lineOf(cues[idx]));
      prev = idx;
    }
    return lines.join('\n');
  }

  // Parse a growing SSE text buffer. Returns extracted assistant tokens, whether
  // the stream signalled [DONE], and any incomplete trailing line to prepend to
  // the next chunk.
  function parseSSE(buffer) {
    const tokens = [];
    let done = false;
    const parts = String(buffer || '').split('\n');
    const rest = parts.pop();
    for (let line of parts) {
      line = line.trim();
      if (!line || line.indexOf('data:') !== 0) continue;
      const data = line.slice(5).trim();
      if (data === '[DONE]') { done = true; continue; }
      try {
        const json = JSON.parse(data);
        const delta = json.choices && json.choices[0] && json.choices[0].delta;
        if (delta && typeof delta.content === 'string') tokens.push(delta.content);
      } catch (_) { /* keepalive / partial JSON — ignore */ }
    }
    return { tokens, done, rest };
  }

  function buildChatRequest(model, systemPrompt, history, question, transcript) {
    const messages = [{ role: 'system', content: systemPrompt + '\n\nTranscript:\n' + transcript }];
    for (const turn of history || []) messages.push({ role: turn.role, content: turn.content });
    messages.push({ role: 'user', content: question });
    return { model, messages, stream: true };
  }

  // Extract model ids from an OpenAI-compatible GET /models response. Accepts
  // `{data:[{id}]}`, a bare array, or an array of strings. Deduped + sorted.
  function parseModelList(json) {
    const arr = json && Array.isArray(json.data) ? json.data : (Array.isArray(json) ? json : []);
    const ids = [];
    for (const m of arr) {
      const id = typeof m === 'string' ? m : (m && m.id);
      if (typeof id === 'string' && id) ids.push(id);
    }
    return [...new Set(ids)].sort();
  }

  return { tokenize, searchCues, cuesToPlainText, buildContext, parseSSE, buildChatRequest, parseModelList };
});
