const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const context = vm.createContext({ window: {}, URL });
vm.runInContext(fs.readFileSync(path.join(__dirname, '../dist/unsubscribe-safety.js'), 'utf8'), context);
const validate = context.window.UnsubscribeSafety.publicHttpsUrl;
test('unsubscribe navigation accepts public HTTPS syntax, rejects dangerous or local targets', () => {
  assert.equal(validate('https://news.example.com/consent?token=synthetic').hostname, 'news.example.com');
  for (const url of ['http://news.example.com/', 'javascript:alert(1)', 'https://user:pass@example.com/', 'https://127.0.0.1/', 'https://[::1]/', 'https://localhost/', 'https://intranet.local/', 'https://example.com:8443/', 'https://example.com/#fragment', 'https://example.com/\n', 'https://example.com/\\evil', 'https://2130706433/', null]) assert.equal(validate(url), null, String(url));
});
test('no automatic form submission remains in the app', () => {
  const source = fs.readFileSync(path.join(__dirname, '../dist/app.js'), 'utf8');
  assert.doesNotMatch(source, /submitOneClickUnsubscribe|form\.submit\(/);
});
