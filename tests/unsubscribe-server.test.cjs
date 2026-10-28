const { test } = require('node:test');
const assert = require('node:assert/strict');
const { createHash } = require('node:crypto');
const metadataModule = import('../server/unsubscribe-metadata.mjs');
const engineModule = import('../server/unsubscribe-engine.mjs');
const headers = [
  { name: 'From', value: 'Newsletter <news@example.com>' },
  { name: 'List-Unsubscribe', value: '<https://news.example.com/unsubscribe?token=synthetic>' },
  { name: 'List-Unsubscribe-Post', value: 'List-Unsubscribe=One-Click' }
];
function providerMessage(provider, extra = {}) {
  return provider === 'gmail' ? { id: 'message-1', payload: { headers }, ...extra }
    : { id: 'message-1', from: { emailAddress: { address: 'news@example.com' } }, internetMessageHeaders: headers, ...extra };
}
const credentials = { token: 'synthetic-token-for-tests', messageId: 'message-1', expectedSender: 'news@example.com' };

test('both metadata readers restrict origin, fields, redirects and return only selected headers', async () => {
  const { createMetadataReader } = await metadataModule;
  for (const provider of ['gmail', 'outlook']) {
    const calls = [];
    const reader = createMetadataReader(async (url, options) => { calls.push({ url: new URL(url), options }); return Response.json(providerMessage(provider)); });
    const value = await reader({ provider, ...credentials });
    assert.equal(value.sender, 'news@example.com');
    assert.equal(calls.length, 1);
    const { url, options } = calls[0];
    assert.equal(url.origin, provider === 'gmail' ? 'https://gmail.googleapis.com' : 'https://graph.microsoft.com');
    assert.equal(options.method, 'GET');
    assert.equal(options.redirect, 'error');
    assert.equal(options.credentials, 'omit');
    assert.equal(options.headers.Authorization, 'Bearer ' + credentials.token);
    assert.doesNotMatch(url.search, /body|raw|attachment|subject|preview/i);
    assert.deepEqual(Object.keys(value).sort(), ['headers', 'messageId', 'provider', 'sender']);
  }
});

test('reader rejects arbitrary provider, path injection, changed sender and changed message', async () => {
  const { createMetadataReader } = await metadataModule;
  let calls = 0;
  const reader = createMetadataReader(async () => { calls++; return Response.json(providerMessage('gmail')); });
  for (const input of [{ provider: 'other' }, { messageId: '../profile' }, { token: 'bad\ntoken' }]) {
    await assert.rejects(reader({ provider: 'gmail', ...credentials, ...input }), /invalid_request/);
  }
  assert.equal(calls, 0);
  await assert.rejects(reader({ provider: 'gmail', ...credentials, expectedSender: 'other@example.com' }), /sender_changed/);
  const changed = createMetadataReader(async () => Response.json(providerMessage('gmail', { id: 'other' })));
  await assert.rejects(changed({ provider: 'gmail', ...credentials }), /message_changed/);
});

test('provider failures are sanitised, oversized streams rejected and no token appears in errors', async () => {
  const { createMetadataReader, boundedJson } = await metadataModule;
  for (const [status, code] of [[401, 'reconnect_required'], [403, 'metadata_permission'], [404, 'message_missing'], [429, 'provider_limited']]) {
    const reader = createMetadataReader(async () => new Response(credentials.token, { status }));
    await assert.rejects(reader({ provider: 'gmail', ...credentials }), (error) => error.code === code && !String(error).includes(credentials.token));
  }
  const reader = createMetadataReader(async () => { throw new Error(credentials.token); });
  await assert.rejects(reader({ provider: 'gmail', ...credentials }), /^MetadataFailure: metadata_unavailable$/);
  await assert.rejects(boundedJson(new Response('x'.repeat(200)), 100), /metadata_too_large/);
});

async function setup({ verified = true, allowed = true, status = 204, throws = false, duplicate = false, durable = true, changed = false } = {}) {
  const { createUnsubscribeEngine } = await engineModule;
  const calls = [], saved = [];
  let reads = 0;
  const metadata = { provider: 'gmail', messageId: 'message-1', sender: 'news@example.com', headers: headers.map((h) => ({ ...h, name: h.name.toLowerCase() })) };
  const engine = createUnsubscribeEngine({
    readMetadata: async () => { reads++; return changed && reads > 1 ? { ...metadata, sender: 'other@example.com' } : metadata; },
    verifySignature: async () => verified, approveDestination: async () => allowed,
    // Deliberately synthetic test adapter; NOT the production durable ledger.
    ledger: { durable, reserve: async () => duplicate ? 'duplicate' : 'reserved', complete: async (key, value) => saved.push({ key, value }) },
    sendPost: async (url, options) => { calls.push({ url, options }); if (throws) throw new Error('private remote error'); return new Response(null, { status }); },
    digest: async (value) => createHash('sha256').update(value).digest('hex')
  });
  const plan = await engine.prepare({ actor: 'actor-1', provider: 'gmail', token: credentials.token, items: [{ messageId: 'message-1', sender: 'news@example.com' }] });
  return { engine, plan, calls, saved, execute: (extra = {}) => engine.execute({ plan, actor: 'actor-1', token: credentials.token, confirmed: true, ...extra }) };
}

test('preparing a candidate sends nothing; confirmation emits exactly one fixed, context-free POST', async () => {
  const s = await setup();
  assert.equal(s.calls.length, 0);
  assert.doesNotMatch(JSON.stringify(s.plan), /synthetic-token-for-tests/);
  assert.equal((await s.execute({ confirmed: false })).reason, 'plan_invalid');
  const response = await s.execute();
  assert.equal(response.outcomes[0].status, 'requested');
  assert.equal(s.calls.length, 1);
  assert.deepEqual(s.calls[0].options.headers, { 'Content-Type': 'application/x-www-form-urlencoded' });
  assert.equal(s.calls[0].options.body, 'List-Unsubscribe=One-Click');
  assert.equal(s.calls[0].options.redirect, 'manual');
  assert.equal(s.calls[0].options.credentials, 'omit');
  assert.equal((await s.execute()).reason, 'plan_invalid');
});

test('wrong actor, refreshed credential, forged plan, missing signature, egress denial and non-durable state fail closed', async () => {
  for (const args of [{ actor: 'other' }, { token: 'other-synthetic-token' }, { plan: {} }]) {
    const s = await setup();
    assert.equal((await s.execute(args)).reason, 'plan_invalid');
    assert.equal(s.calls.length, 0);
  }
  for (const options of [{ verified: false }, { allowed: false }, { durable: false }, { duplicate: true }, { changed: true }]) {
    const s = await setup(options);
    await s.execute();
    assert.equal(s.calls.length, 0);
  }
});

test('timeouts and redirects are uncertain, never followed/retried; HTTP errors do not prove login', async () => {
  for (const [options, expected] of [[{ throws: true }, 'outcome_unknown'], [{ status: 302 }, 'redirect_received'], [{ status: 403 }, 'request_rejected'], [{ status: 429 }, 'sender_limited']]) {
    const s = await setup(options);
    const response = await s.execute();
    assert.equal(response.outcomes[0].reason, expected);
    await s.execute();
    assert.equal(s.calls.length, 1);
  }
});

test('concurrent execution consumes a plan only once', async () => {
  const s = await setup();
  await Promise.all([s.execute(), s.execute(), s.execute()]);
  assert.equal(s.calls.length, 1);
});

test('unsupported, duplicate and local unsubscribe headers return specific reasons without navigation', async () => {
  const { inspectOneClick } = await engineModule;
  const h = headers.map((entry) => ({ ...entry, name: entry.name.toLowerCase() }));
  assert.equal(inspectOneClick({ headers: [] }).reason, 'link_missing');
  assert.equal(inspectOneClick({ headers: h.slice(0, 2) }).reason, 'automatic_unavailable');
  assert.equal(inspectOneClick({ headers: [...h, h[1]] }).reason, 'headers_ambiguous');
  assert.equal(inspectOneClick({ headers: h.map((entry) => entry.name === 'list-unsubscribe' ? { ...entry, value: '<https://127.0.0.1/>' } : entry) }).reason, 'destination_unsupported');
});
