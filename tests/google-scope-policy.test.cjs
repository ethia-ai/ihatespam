const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

test('Google scope inventory stays bounded and excludes implicit old grants', () => {
  const source = fs.readFileSync(path.join(__dirname, '../dist/app.js'), 'utf8');
  const scopes = [...new Set(source.match(/https:\/\/www\.googleapis\.com\/auth\/[a-z.]+/g))].sort();
  assert.deepEqual(scopes, [
    'https://www.googleapis.com/auth/gmail.metadata',
    'https://www.googleapis.com/auth/gmail.modify',
    'https://www.googleapis.com/auth/gmail.settings.basic'
  ]);
  assert.doesNotMatch(source, /include_granted_scopes:\s*true/);
  assert.equal((source.match(/include_granted_scopes:\s*false/g) || []).length, 2);
  assert.match(source, /scope: GMAIL_METADATA_SCOPE/);
  assert.match(source, /new Set\(\[GMAIL_METADATA_SCOPE, \.\.\.scopes\]\)/);
});

test('Privacy disclosure includes concrete mechanisms and Limited Use', () => {
  const policy = fs.readFileSync(path.join(__dirname, '../dist/privacy.html'), 'utf8');
  for (const term of ['HTTPS', 'AES-GCM', 'HMAC', 'Limited Use', 'page memory', 'request origin']) assert.ok(policy.includes(term), term);
});
