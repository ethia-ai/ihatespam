const { test } = require('node:test');
const assert = require('node:assert/strict');
const { generateKeyPairSync, sign } = require('node:crypto');
const { DatabaseSync } = require('node:sqlite');
const fs = require('node:fs');
const path = require('node:path');

function database() {
  const db = new DatabaseSync(':memory:');
  db.exec(fs.readFileSync(path.join(__dirname, '../drizzle/0000_useful_lady_deathstrike.sql'), 'utf8'));
  const binding = {
    prepare(sql) { return { bind(...args) { return {
      async run() { const r = db.prepare(sql).run(...args); return { meta: { changes: Number(r.changes) } }; },
      async first() { return db.prepare(sql).get(...args) || null; }
    }; } }; },
    async batch(statements) {
      db.exec('BEGIN');
      try { const results = []; for (const statement of statements) results.push(await statement.run()); db.exec('COMMIT'); return results; }
      catch (e) { db.exec('ROLLBACK'); throw e; }
    }
  };
  return { db, binding };
}

const pair = generateKeyPairSync('rsa', { modulusLength: 2048 });
function signedMetadata(messageId = 'message-1') {
  const values = [
    ['from', 'news@example.com'],
    ['list-unsubscribe', '<https://news.beehiiv.com/unsubscribe?test=' + messageId + '>'],
    ['list-unsubscribe-post', 'List-Unsubscribe=One-Click']
  ];
  const signature = 'v=1; a=rsa-sha256; c=relaxed/relaxed; d=example.com; s=mail; h=from:list-unsubscribe:list-unsubscribe-post; bh=' + Buffer.alloc(32).toString('base64') + '; b=';
  const signingInput = values.map(([name, value]) => name + ':' + value + '\r\n').join('') + 'dkim-signature:' + signature;
  const signatureBytes = sign('RSA-SHA256', Buffer.from(signingInput), pair.privateKey).toString('base64');
  return { provider: 'gmail', messageId, sender: 'news@example.com', headers: [
    ...values.map(([name, value]) => ({ name, value })),
    { name: 'dkim-signature', value: signature + signatureBytes },
    { name: 'authentication-results', value: 'mx.google.com; dkim=pass header.i=@example.com header.s=mail header.b=' + signatureBytes.slice(0, 12) }
  ] };
}

test('independent RSA fixture validates; changed signed URL, unmatched attestation and Outlook fail closed', async () => {
  const { createSignatureVerifier } = await import('../server/unsubscribe-signature.mjs');
  const verify = createSignatureVerifier({ lookupKey: async () => 'v=DKIM1; k=rsa; p=' + pair.publicKey.export({ type: 'spki', format: 'der' }).toString('base64') });
  assert.equal(await verify(signedMetadata()), true);
  const changed = signedMetadata(); changed.headers[1].value += 'evil';
  assert.equal(await verify(changed), false);
  const untrusted = signedMetadata(); untrusted.headers.at(-1).value = 'evil.example; dkim=pass';
  assert.equal(await verify(untrusted), false);
  const duplicate = signedMetadata(); duplicate.headers.push(duplicate.headers[1]);
  assert.equal(await verify(duplicate), false);
  assert.equal(await verify({ ...signedMetadata(), provider: 'outlook' }), false);
});

test('egress allowlist and DNS reject private/metadata addresses and custom domains', async () => {
  const { supportedUnsubscribeUrl, publicDnsAddress, createDestinationPolicy } = await import('../server/unsubscribe-egress.mjs');
  for (const url of [
    'https://evil.test/', 'https://news.beehiiv.com.evil.com/', 'https://unsubscribe.kit.com.evil.com/',
    'https://ac.emlnk1.com.evil.com/', 'https://news.beehiiv.com:8443/', 'https://user@news.beehiiv.com/',
    'https://127.0.0.1/'
  ]) assert.equal(supportedUnsubscribeUrl(url), null);
  for (const url of [
    'https://news.beehiiv.com/unsubscribe?id=one',
    'https://eu2-cdn.eng.bloomreach.com/consent?id=two',
    'https://unsubscribe.kit.com/opaque-token',
    'https://ac.emlnk1.com/lnk/one-click?token=three'
  ]) assert.equal(supportedUnsubscribeUrl(url)?.href, url);
  for (const address of ['127.0.0.1', '10.1.1.1', '169.254.169.254', '100.64.0.1', '192.168.1.1', '::1', 'fd00::1', '::ffff:127.0.0.1', '2001:db8::1']) assert.equal(publicDnsAddress(address), false);
  assert.equal(publicDnsAddress('104.18.1.1'), true);
  assert.equal(publicDnsAddress('2606:4700::1111'), true);
  const policy = createDestinationPolicy(async () => Response.json({ Status: 0, Answer: [{ type: 1, data: '127.0.0.1' }] }));
  assert.equal(await policy('https://news.beehiiv.com/'), false);
});

test('real SQLite transaction keeps quota intact on duplicate reservations; replay survives new instances', async () => {
  const { createUnsubscribeLedger } = await import('../server/unsubscribe-ledger.mjs');
  const { db, binding } = database();
  const actor = 'a'.repeat(64), key = 'b'.repeat(64);
  const ledger = createUnsubscribeLedger(binding, actor, () => 100000);
  assert.equal(await ledger.reserve(key), 'reserved');
  assert.equal(await createUnsubscribeLedger(binding, actor, () => 100000).reserve(key), 'duplicate');
  assert.equal(db.prepare('SELECT remaining FROM unsubscribe_budgets WHERE key = ?').get('send:' + actor + ':day:0').remaining, 49);
  await ledger.complete(key, { status: 'requested', reason: 'request_accepted' });
  assert.equal(await ledger.consumePlan('c'.repeat(64)), true);
  assert.equal(await ledger.consumePlan('c'.repeat(64)), false);
  for (let i = 0; i < 49; i++) assert.equal(await ledger.reserve(i.toString(16).padStart(64, '0')), 'reserved');
  assert.equal(await ledger.reserve('d'.repeat(64)), 'limit');
  db.close();
});

test('encrypted plans detect tampering and do not reveal sender or destination', async () => {
  const { createPlanCrypto } = await import('../server/unsubscribe-crypto.mjs');
  const cipher = await createPlanCrypto('ab'.repeat(32));
  const plan = { actor: 'person', rows: [{ sender: 'news@example.com', url: 'https://news.beehiiv.com/private-token' }] };
  const sealed = await cipher.seal(plan);
  assert.doesNotMatch(sealed, /example|beehiiv|private-token/);
  assert.deepEqual(await cipher.open(sealed), plan);
  const large = { rows: Array.from({ length: 50 }, () => ({ url: 'x'.repeat(8192) })) };
  assert.deepEqual(await cipher.open(await cipher.seal(large)), large);
  const differentKey = await createPlanCrypto('cd'.repeat(32));
  await assert.rejects(differentKey.open(sealed));
  const [iv, data] = sealed.split('.');
  await assert.rejects(cipher.open(iv + '.' + (data[0] === 'a' ? 'b' : 'a') + data.slice(1)));
});

test('HTTP boundary rejects unauthenticated, cross-origin, disabled and malformed requests before any network call', async () => {
  const { handleUnsubscribeApi } = await import('../server/unsubscribe-api.mjs');
  const origin = 'https://ihatespam.site';
  let calls = 0;
  const network = async () => { calls++; throw new Error('unexpected'); };
  const make = (headers = {}) => new Request(origin + '/api/unsubscribe/prepare', { method: 'POST', headers: { 'content-type': 'application/json', origin, ...headers }, body: '{}' });
  const { db, binding } = database();
  const env = { DB: binding, UNSUBSCRIBE_ENABLED: 'true', UNSUBSCRIBE_SECRET: '34'.repeat(32) };
  assert.equal((await handleUnsubscribeApi(make(), env, network)).status, 401);
  assert.equal((await handleUnsubscribeApi(make({ 'oai-authenticated-user-id': 'forged' }), env, network)).status, 401);
  assert.equal((await handleUnsubscribeApi(make({ 'oai-authenticated-user-id': 'actor', origin: 'https://evil.test' }), {}, network)).status, 403);
  assert.equal((await handleUnsubscribeApi(make({ 'oai-authenticated-user-id': 'actor' }), {}, network)).status, 503);
  assert.equal(calls, 0);
  db.close();
});

for (const batchSize of [1, 50]) test(`complete HTTP flow prepares ${batchSize} signed messages and sends each once across Worker instances`, async () => {
  const { handleUnsubscribeApi } = await import('../server/unsubscribe-api.mjs');
  const { db, binding } = database();
  const env = { DB: binding, UNSUBSCRIBE_ENABLED: 'true', UNSUBSCRIBE_SECRET: '34'.repeat(32) };
  const networkCalls = [];
  let mailbox = 'owner@example.com';
  const network = async (url, options) => {
    const target = new URL(url); networkCalls.push({ target, options });
    // Reproduce Worker fetch: browser-only fields throw before networking.
    for (const field of ['cache', 'credentials', 'referrerPolicy']) assert.equal(Object.hasOwn(options, field), false);
    assert.equal(options.redirect, 'manual');
    const headers = new Headers(options.headers);
    if (target.hostname === 'gmail.googleapis.com') {
      assert.equal(options.method, 'GET');
      assert.equal(headers.get('Authorization'), 'Bearer synthetic-access-token');
      assert.equal(headers.get('Cache-Control'), 'no-store');
      assert.deepEqual(options.cf.cacheTtlByStatus, { '100-599': -1 });
      if (target.pathname.endsWith('/profile')) {
        assert.equal(target.search, '?fields=emailAddress');
        assert.equal(options.redirect, 'manual');
        assert.equal(headers.has('Cookie'), false);
        assert.equal(headers.has('Referer'), false);
        return Response.json({ emailAddress: mailbox });
      }
      const messageId = target.pathname.split('/').at(-1);
      return Response.json({ id: messageId, payload: { headers: signedMetadata(messageId).headers } });
    }
    if (target.hostname === 'cloudflare-dns.com') {
      assert.equal(headers.has('Authorization'), false);
      const name = target.searchParams.get('name');
      const type = target.searchParams.get('type');
      if (type === 'TXT') return Response.json({ Status: 0, Answer: [{ name, type: 16, data: '"v=DKIM1; k=rsa; p=' + pair.publicKey.export({ type: 'spki', format: 'der' }).toString('base64') + '"' }] });
      return Response.json({ Status: 0, Answer: type === 'A' ? [{ name, type: 1, data: '104.18.1.1' }] : [] });
    }
    assert.equal(target.hostname, 'news.beehiiv.com');
    assert.equal(options.method, 'POST');
    assert.equal(headers.has('Authorization'), false);
    assert.equal(headers.has('Cookie'), false);
    assert.equal(headers.get('Cache-Control'), 'no-store');
    assert.equal(options.redirect, 'manual');
    return new Response(null, { status: 204 });
  };
  const call = async (route, body) => handleUnsubscribeApi(new Request('https://ihatespam.site/api/unsubscribe/' + route, {
    method: 'POST', headers: { 'content-type': 'application/json', origin: 'https://ihatespam.site' }, body: JSON.stringify(body)
  }), env, network);
  const items = Array.from({ length: batchSize }, (_, i) => ({ messageId: `message-${i + 1}`, sender: 'news@example.com' }));
  const oversized = await call('prepare', { provider: 'gmail', token: 'synthetic-access-token', items: Array.from({ length: 51 }, (_, i) => ({ messageId: `message-${i}`, sender: 'news@example.com' })) });
  assert.equal(oversized.status, 400);
  assert.equal(networkCalls.length, 0);
  const preparedResponse = await call('prepare', { provider: 'gmail', token: 'synthetic-access-token', items });
  assert.equal(preparedResponse.status, 200);
  const prepared = await preparedResponse.json();
  assert.equal(prepared.rows.length, batchSize);
  assert.ok(prepared.rows.every(row => row.status === 'ready'));
  assert.equal(networkCalls.filter((call) => call.options.method === 'POST').length, 0);
  mailbox = 'different@example.com';
  assert.equal((await call('execute', { plan: prepared.plan, token: 'synthetic-access-token', confirmed: true, actor: 'owner@example.com' })).status, 409);
  assert.equal(networkCalls.filter((call) => call.options.method === 'POST').length, 0);
  mailbox = 'owner@example.com';
  const sent = await call('execute', { plan: prepared.plan, token: 'synthetic-access-token', confirmed: true });
  assert.equal(sent.status, 200);
  const outcomes = (await sent.json()).outcomes;
  assert.equal(outcomes.length, batchSize);
  assert.ok(outcomes.every(outcome => outcome.status === 'requested'));
  assert.equal((await call('execute', { plan: prepared.plan, token: 'synthetic-access-token', confirmed: true })).status, 409);
  assert.equal(networkCalls.filter((call) => call.options.method === 'POST').length, batchSize);
  assert.doesNotMatch(JSON.stringify(db.prepare('SELECT * FROM unsubscribe_budgets').all()), /owner@example|synthetic-access-token/);
  db.close();
});

test('Worker private fetch never retries an ambiguous network failure', async () => {
  const { workerPrivateFetch } = await import('../server/unsubscribe-api.mjs');
  let calls = 0;
  const network = workerPrivateFetch(async (_, options) => {
    calls++;
    assert.equal(Object.hasOwn(options, 'cache'), false);
    assert.deepEqual(options.cf.cacheTtlByStatus, { '100-599': -1 });
    throw new Error('synthetic-network-error');
  });
  await assert.rejects(network('https://news.beehiiv.com/', { method: 'POST', cache: 'no-store' }), /synthetic-network-error/);
  assert.equal(calls, 1);
});

test('Worker private fetch rejects provider redirects without following them', async () => {
  const { workerPrivateFetch } = await import('../server/unsubscribe-api.mjs');
  let calls = 0;
  const network = workerPrivateFetch(async (_, options) => {
    calls++;
    assert.equal(options.redirect, 'manual');
    return new Response(null, { status: 302, headers: { Location: 'https://untrusted.example/' } });
  });
  await assert.rejects(network('https://gmail.googleapis.com/', { redirect: 'error' }), /redirect_rejected/);
  assert.equal(calls, 1);
});

test('public authentication rejects invalid credentials and provider failures without contacting senders', async () => {
  const { handleUnsubscribeApi } = await import('../server/unsubscribe-api.mjs');
  const { db, binding } = database();
  const env = { DB: binding, UNSUBSCRIBE_ENABLED: 'true', UNSUBSCRIBE_SECRET: '34'.repeat(32) };
  for (const [status, expected] of [[401, 401], [403, 401], [429, 429], [500, 503], [302, 503]]) {
    let calls = 0;
    const response = await handleUnsubscribeApi(new Request('https://ihatespam.site/api/unsubscribe/prepare', {
      method: 'POST', headers: { origin: 'https://ihatespam.site', 'content-type': 'application/json', 'oai-authenticated-user-id': 'forged' },
      body: JSON.stringify({ provider: 'gmail', token: 'invalid-synthetic-token', items: [{ messageId: 'message-1', sender: 'news@example.com' }] })
    }), env, async (url) => {
      calls++;
      assert.equal(new URL(url).pathname, '/gmail/v1/users/me/profile');
      return new Response('invalid-synthetic-token', { status });
    });
    assert.equal(response.status, expected);
    assert.equal(calls, 1);
    assert.doesNotMatch(await response.text(), /invalid-synthetic-token/);
  }
  assert.equal(db.prepare('SELECT count(*) AS total FROM unsubscribe_budgets').get().total, 0);
  db.close();
});

test('cross-request sealed execution uses durable plan consumption and contains no credential in storage', async () => {
  const { createUnsubscribeEngine } = await import('../server/unsubscribe-engine.mjs');
  const { createUnsubscribeLedger } = await import('../server/unsubscribe-ledger.mjs');
  const { createPlanCrypto } = await import('../server/unsubscribe-crypto.mjs');
  const { db, binding } = database();
  const cipher = await createPlanCrypto('12'.repeat(32));
  const actor = 'actor', token = 'synthetic-secret-token';
  let sends = 0;
  const make = async () => createUnsubscribeEngine({ readMetadata: async () => signedMetadata(), verifySignature: async () => true, approveDestination: async () => true, ledger: createUnsubscribeLedger(binding, await cipher.digest(actor)), digest: cipher.digest, restoreAuthenticatedPlan: cipher.open, sendPost: async () => { sends++; return new Response(null, { status: 204 }); } });
  const prepared = await (await make()).prepare({ actor, provider: 'gmail', token, items: [{ messageId: 'message-1', sender: 'news@example.com' }] });
  const sealedPlan = await cipher.seal(prepared);
  const outcome = await (await make()).execute({ sealedPlan, actor, token, confirmed: true });
  assert.equal(outcome.outcomes[0].status, 'requested');
  assert.equal((await (await make()).execute({ sealedPlan, actor, token, confirmed: true })).reason, 'plan_invalid');
  assert.equal(sends, 1);
  assert.doesNotMatch(JSON.stringify(db.prepare('SELECT * FROM unsubscribe_attempts').all()), /synthetic|example|beehiiv/);
  db.close();
});
