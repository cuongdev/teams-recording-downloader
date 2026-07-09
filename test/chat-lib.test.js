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
  // Both cues cover the same query terms with equal frequency (score 2.2
  // each without the phrase bonus), and the phrase-containing cue sits at
  // the LATER index. Without the +3 phrase bonus the tie breaks to the
  // earlier index (pos 1), so this fails if the bonus is ever removed.
  const phraseCues = [
    { pos: 1, speaker: 'Alice', start: 0, end: 5, text: 'The page about billing needs review.' },
    { pos: 2, speaker: 'Bob', start: 5, end: 10, text: 'Let us finalize the billing page today.' },
  ];
  const r = ChatLib.searchCues(phraseCues, 'billing page');
  assert.equal(r[0].cue.pos, 2);
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

test('buildContext stays within budget across disjoint clusters', () => {
  // Two matching cues far apart in the transcript, everything else is
  // non-matching filler long enough to never be pulled in as a "neighbor"
  // line. This mirrors the original bug: the assembly loop inserts an
  // uncounted '…' separator between the two clusters, which used to let
  // the result exceed `budget` (the original repro used budget=31 and
  // returned a 32-char result). The fixed reservation accounting is a
  // provable-safe upper bound rather than an exact one, so it needs a
  // little more headroom than the true minimum to admit both clusters —
  // 35 is that headroom-inclusive minimum for this fixture; below it the
  // fix conservatively (and still safely) keeps only one cluster.
  const manyCues = [];
  for (let i = 0; i < 16; i++) {
    manyCues.push({
      pos: i,
      speaker: 'S',
      start: i,
      end: i + 1,
      text: (i === 2 || i === 15) ? 'target' : 'this is unrelated padding text here',
    });
  }
  const budget = 35;
  const ctx = ChatLib.buildContext(manyCues, 'target', budget);
  assert.ok(ctx.includes('…'));
  assert.ok(ctx.length <= budget);
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
  assert.equal(req.messages.length, 4);
  assert.equal(req.messages[0].role, 'system');
  assert.ok(req.messages[0].content.includes('transcript text'));
  assert.equal(req.messages[1].role, 'user');
  assert.equal(req.messages[1].content, 'hi');
  assert.equal(req.messages[2].role, 'assistant');
  assert.equal(req.messages[2].content, 'hello');
  assert.equal(req.messages[req.messages.length - 1].content, 'Who mentioned login?');
});
