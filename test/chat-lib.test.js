const test = require('node:test');
const assert = require('node:assert');
const ChatLib = require('../chat-lib.js');

const cues = [
  { pos: 1, speaker: 'Alice', start: 0, end: 5, text: 'We should ship the login feature next week.' },
  { pos: 2, speaker: 'Bob', start: 5, end: 10, text: 'I agree, the login work is almost done.' },
  { pos: 3, speaker: 'Alice', start: 10, end: 15, text: 'Great, what about the billing page?' },
];

test('tokenize lowercases and splits on non-alphanumerics', () => {
  assert.deepEqual(ChatLib.tokenize('Hello, WORLD-2!'), ['hello', 'world', '2']);
});

test('searchCues returns [] for an empty query', () => {
  assert.deepEqual(ChatLib.searchCues(cues, '   '), []);
});

test('searchCues finds every cue containing a term', () => {
  const r = ChatLib.searchCues(cues, 'login');
  assert.equal(r.length, 2);
  assert.equal(r[0].cue.speaker, 'Alice'); // tie broken by earlier index
});

test('searchCues boosts an exact phrase match to the top', () => {
  const r = ChatLib.searchCues(cues, 'billing page');
  assert.equal(r[0].cue.pos, 3);
});

test('buildContext returns the full transcript when under budget', () => {
  const ctx = ChatLib.buildContext(cues, 'login', 48000);
  assert.ok(ctx.includes('Alice'));
  assert.ok(ctx.includes('billing'));
});

test('buildContext retrieves the relevant passage when over budget', () => {
  const ctx = ChatLib.buildContext(cues, 'billing', 60); // smaller than full transcript
  assert.ok(ctx.includes('billing'));
  assert.ok(ctx.length <= 60);
});

test('parseSSE extracts delta tokens and detects [DONE]', () => {
  const chunk =
    'data: {"choices":[{"delta":{"content":"Hel"}}]}\n' +
    'data: {"choices":[{"delta":{"content":"lo"}}]}\n' +
    'data: [DONE]\n';
  const { tokens, done } = ChatLib.parseSSE(chunk);
  assert.deepEqual(tokens, ['Hel', 'lo']);
  assert.equal(done, true);
});

test('parseSSE keeps an incomplete trailing line as rest', () => {
  const { tokens, rest } = ChatLib.parseSSE(
    'data: {"choices":[{"delta":{"content":"Hi"}}]}\ndata: {"cho'
  );
  assert.deepEqual(tokens, ['Hi']);
  assert.equal(rest, 'data: {"cho');
});

test('buildChatRequest assembles system + history + question', () => {
  const req = ChatLib.buildChatRequest(
    'gpt-4o-mini',
    'You answer questions about a meeting.',
    [{ role: 'user', content: 'hi' }, { role: 'assistant', content: 'hello' }],
    'Who mentioned login?',
    'transcript text'
  );
  assert.equal(req.model, 'gpt-4o-mini');
  assert.equal(req.stream, true);
  assert.equal(req.messages[0].role, 'system');
  assert.ok(req.messages[0].content.includes('transcript text'));
  assert.equal(req.messages[req.messages.length - 1].content, 'Who mentioned login?');
});
