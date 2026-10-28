// A dependency-free build: preserve dist/ as the handwritten static app, then
// wrap only its public files to enforce headers that static Sites didn't apply.
const fs = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');
const root = path.resolve(__dirname, '..');
const files = [
  'index.html', 'styles.css', 'app.js', 'i18n.js', 'outlook-archive.js', 'unsubscribe-safety.js',
  'privacy.html', 'terms.html', 'og.png', 'oauth-app-icon.png', 'auth-redirect.html', 'auth-redirect.js',
  'google-access-step-1.png', 'google-access-step-2.png',
  'vendor/msal-browser-5.17.3.min.js', 'vendor/msal-redirect-bridge-5.17.3.min.js', 'vendor/MSAL-LICENSE.txt'
];
const types = { '.html': 'text/html; charset=utf-8', '.js': 'application/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8', '.png': 'image/png', '.txt': 'text/plain; charset=utf-8' };
const assets = {};
for (const file of files) {
  const bytes = fs.readFileSync(path.join(root, 'dist', file));
  assets['/' + file] = { type: types[path.extname(file)], data: bytes.toString('base64'), etag: '"' + crypto.createHash('sha256').update(bytes).digest('hex') + '"' };
}
const modules = ['unsubscribe-metadata', 'unsubscribe-engine', 'unsubscribe-signature', 'unsubscribe-egress', 'unsubscribe-crypto', 'unsubscribe-ledger', 'unsubscribe-api', 'security-worker'];
const worker = modules.map((name) => fs.readFileSync(path.join(root, `server/${name}.mjs`), 'utf8').replace(/^import .*;\r?\n/gm, '')).join('\n');
fs.mkdirSync(path.join(root, 'dist/server'), { recursive: true });
fs.writeFileSync(path.join(root, 'dist/server/index.js'), `${worker}\nexport default createStaticSite(${JSON.stringify(assets)});\n`);
console.log(`Built protected Worker for ${files.length} public files and authenticated unsubscribe routes.`);
