const { test } = require('node:test');
const assert = require('node:assert/strict');
const { pathToFileURL } = require('node:url');
const path = require('node:path');
const source = import(pathToFileURL(path.join(__dirname, '../server/security-worker.mjs')));
const asset = (type, body) => ({ type, data: Buffer.from(body).toString('base64'), etag: '"fixed-test-hash"' });
async function site() {
  return (await source).createStaticSite({
    '/index.html': asset('text/html; charset=utf-8', '<h1>Private</h1>'),
    '/auth-redirect.html': asset('text/html; charset=utf-8', 'bridge'),
    '/app.js': asset('application/javascript; charset=utf-8', '/* app */')
  });
}

test('security headers apply to all page, asset, missing and error responses', async () => {
  const worker = await site();
  for (const route of ['/', '/app.js', '/missing']) {
    const response = await worker.fetch(new Request('https://example.test' + route));
    assert.equal(response.headers.get('x-frame-options'), 'DENY');
    assert.match(response.headers.get('content-security-policy'), /frame-ancestors 'none'/);
    assert.equal(response.headers.get('referrer-policy'), 'no-referrer');
    assert.equal(response.headers.get('x-content-type-options'), 'nosniff');
    assert.match(response.headers.get('permissions-policy'), /camera=\(\)/);
  }
});

test('both callback URLs serve directly, without redirects or caching', async () => {
  const worker = await site();
  for (const route of ['/auth-redirect.html', '/auth-redirect']) {
    const response = await worker.fetch(new Request('https://example.test' + route + '?not-reflected=secret'));
    assert.equal(response.status, 200);
    assert.equal(response.headers.get('location'), null);
    assert.equal(response.headers.get('cache-control'), 'no-store');
    assert.equal(response.headers.get('x-frame-options'), 'SAMEORIGIN');
    assert.match(response.headers.get('content-security-policy'), /frame-ancestors 'self'/);
    assert.equal(await response.text(), 'bridge');
  }
});

test('no POST endpoint, arbitrary file access, host reflection or open redirect', async () => {
  const worker = await site();
  assert.equal((await worker.fetch(new Request('https://evil.test/', { method: 'POST', body: 'private' }))).status, 405);
  for (const route of ['/server/index.js', '/.openai/hosting.json', '/_headers', '/__proto__', '//evil.test/', '/?url=https://evil.test']) {
    const response = await worker.fetch(new Request('https://example.test' + route));
    assert.equal(response.headers.get('location'), null);
    assert.ok([200, 404].includes(response.status));
    assert.ok(!(await response.text()).includes('evil.test'));
  }
});

test('HEAD and conditional asset requests have no body; HTML is never served as cached 304', async () => {
  const worker = await site();
  const head = await worker.fetch(new Request('https://example.test/', { method: 'HEAD' }));
  assert.equal(await head.text(), '');
  const cached = await worker.fetch(new Request('https://example.test/app.js', { headers: { 'If-None-Match': '"fixed-test-hash"' } }));
  assert.equal(cached.status, 304);
  assert.equal(await cached.text(), '');
  const html = await worker.fetch(new Request('https://example.test/', { headers: { 'If-None-Match': '"fixed-test-hash"' } }));
  assert.equal(html.status, 200);
});
