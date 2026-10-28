const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.join(__dirname, '../dist');

test('support links use the approved destination and safe external-link attributes', () => {
  const html = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
  const links = html.match(/<a class="support-button[^>]+>/g) || [];
  assert.equal(links.length, 2);
  for (const link of links) {
    assert.match(link, /href="https:\/\/buymeacoffee\.com\/ethia"/);
    assert.match(link, /target="_blank"/);
    assert.match(link, /rel="noopener noreferrer"/);
    assert.match(link, /referrerpolicy="no-referrer"/);
  }
});

test('privacy and support copy is complete in every supported language', () => {
  const nodes = [];
  const document = {
    documentElement: {},
    title: '',
    dispatchEvent() {},
    addEventListener(name, callback) { if (name === 'DOMContentLoaded') callback(); },
    querySelector() { return null; },
    querySelectorAll(selector) {
      if (selector === '[data-i18n]') return nodes;
      return [];
    }
  };
  const context = vm.createContext({
    document,
    navigator: { languages: ['en'], language: 'en' },
    localStorage: { getItem() { return null; }, setItem() {} },
    CustomEvent: class {},
    Intl,
    window: {}
  });
  vm.runInContext(fs.readFileSync(path.join(root, 'i18n.js'), 'utf8'), context);
  for (const language of ['en', 'pt', 'it', 'es', 'fr', 'de']) {
    context.window.WeHateSpamI18n.setLanguage(language, false);
    for (const key of ['privacyCopy', 'buyMeCoffee', 'supportProject', 'supportCopy', 'gmailVerificationNote', 'gmailVerificationHelp', 'gmailVerificationSteps', 'openSourceLink', 'openSourceNote', 'gmailGuideStep1', 'gmailGuideStep2', 'manualLinkOpened', 'manualConfirmSender', 'manualConfirmOpened']) {
      const translated = context.window.WeHateSpamI18n.t(key);
      assert.notEqual(translated, key, `${language}.${key} is missing`);
      assert.ok(translated.length > 3, `${language}.${key} is empty`);
    }
  }
});

test('Gmail warning is linked to consent button and open-source link is transparent', () => {
  const html = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
  assert.match(html, /data-provider="gmail" aria-describedby="gmail-verification-note"/);
  assert.match(html, /id="gmail-verification-note"/);
  assert.match(html, /not verified by Google/);
  assert.match(html, /tutorial below/);
  assert.match(html, /https:\/\/github.com\/ethia-ai\/ihatespam/);
  assert.match(html, /Open source does not guarantee security or Google approval/);
});

test('public pages use the official domain and publish the operator contact', () => {
  const home = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
  const privacy = fs.readFileSync(path.join(root, 'privacy.html'), 'utf8');
  const terms = fs.readFileSync(path.join(root, 'terms.html'), 'utf8');
  assert.match(home, /rel="canonical" href="https:\/\/ihatespam\.site"/);
  assert.doesNotMatch(home, /wehatespam\.ethia-studio\.chatgpt\.site/);
  for (const page of [privacy, terms]) {
    assert.match(page, /Ethia/);
    assert.match(page, /mailto:ethia\.studio@gmail\.com/);
    assert.doesNotMatch(page, /owner-only test|private version/);
  }
});
