const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function harness({ cancelPermission = false, wrongAccount = false, embedded = false, automatic = false, executeFailure = false, mailbox, failVerificationFor, failRule = false, failGmailFilters = false, stored = new Map() } = {}) {
  const nodes = new Map();
  function node() {
    return {
      childNodes: [], dataset: {}, handlers: {}, style: {}, hidden: false, disabled: false,
      classList: { add() {}, remove() {}, toggle() {} }, textContent: '',
      setAttribute() {}, removeAttribute(name) { delete this[name]; }, append(...children) { this.childNodes.push(...children); },
      replaceChildren(...children) { this.childNodes = children; },
      addEventListener(name, handler) { this.handlers[name] = handler; },
      showModal() { this.open = true; }, close() { this.open = false; },
      set innerHTML(_) { throw new Error('Untrusted HTML forbidden'); }
    };
  }
  const get = (selector) => { if (!nodes.has(selector)) nodes.set(selector, node()); return nodes.get(selector); };
  const account = { homeAccountId: 'A', tenantId: 'tenant', localAccountId: 'local', username: 'test@example.com' };
  const calls = [], oauth = [];
  const localStorage = {
    getItem(key) { return stored.has(key) ? stored.get(key) : null; },
    setItem(key, value) { stored.set(key, String(value)); },
    removeItem(key) { stored.delete(key); }
  };
  const messages = new Map((mailbox || [{ id: 'message-one', address: 'sender@example.com' }]).map((message) => [message.id, { ...message, folder: 'inbox-id' }]));
  const authentication = (scopes = ['Mail.ReadWrite']) => ({ account: wrongAccount ? { ...account, homeAccountId: 'B' } : account, scopes, accessToken: 'synthetic-write-token' });
  const client = {
    initialize: async () => {},
    acquireTokenPopup: async (input) => { oauth.push(input); if (cancelPermission) throw new Error('cancelled'); return authentication(input.scopes); },
    acquireTokenSilent: async (input) => authentication(input.scopes)
  };
  const window = {
    localStorage,
    WeHateSpamI18n: { t: (key, values = {}) => key + JSON.stringify(values) },
    msal: { PublicClientApplication: function () { return client; } },
    google: { accounts: { oauth2: { initTokenClient: (options) => ({ requestAccessToken() { oauth.push(options.scope); void options.callback({ access_token: 'synthetic-new-gmail-token' }); } }) } } },
    location: { origin: 'https://example.test', pathname: '/' },
    isSecureContext: true, addEventListener() {}, scrollTo() {}
  };
  window.self = window;
  window.top = embedded ? {} : window;
  const context = vm.createContext({
    window, testAccount: account, URL, URLSearchParams, AbortSignal, AbortController, setTimeout, clearTimeout,
    document: { querySelector: get, querySelectorAll: () => [], createElement: node, createTextNode: (text) => ({ textContent: text }), addEventListener() {} },
    fetch: async (url, options) => {
      url = String(url);
      if (automatic && url.startsWith('/api/unsubscribe/')) {
        calls.push({ url: new URL(url, 'https://example.test'), ...options });
        if (executeFailure && url.endsWith('/execute')) throw new Error('connection lost');
        const items = url.endsWith('/prepare') ? JSON.parse(options.body).items : [];
        if (items.length) context.preparedItems = items;
        return Response.json(url.endsWith('/prepare') ? { plan: 'synthetic-sealed-plan', rows: items.map((item, index) => ({ ...item, status: index ? 'manual' : 'ready', reason: index ? 'automatic_unavailable' : undefined, host: 'news.beehiiv.com' })) } : { status: 'processed', outcomes: context.preparedItems.map((_, index) => index ? { status:'manual', reason:'automatic_unavailable' } : { status:'requested', reason:'request_accepted' }) });
      }
      calls.push({ url: new URL(url), ...options });
      if (new URL(url).hostname === 'gmail.googleapis.com') {
        if (failGmailFilters && new URL(url).pathname.endsWith('/settings/filters')) throw Error('Filters unavailable');
        if (new URL(url).pathname.endsWith('/profile')) return Response.json({ emailAddress: wrongAccount ? 'other@example.com' : 'owner@example.com' });
        return Response.json({});
      }
      const p = new URL(url).pathname;
      if (p.endsWith('/mailFolders/inbox/messageRules') && options.method === 'POST') {
        if (failRule) return Response.json({ error: { code: 'Denied' } }, { status: 403 });
        return Response.json({ id: `rule-${calls.filter((call) => call.url.pathname.endsWith('/messageRules')).length}` }, { status: 201 });
      }
      const id = decodeURIComponent(p.match(/\/messages\/([^/]+)/)?.[1] || '');
      const message = messages.get(id);
      if (options.method === 'POST') { message.folder = { archive: 'archive-id', inbox: 'inbox-id', deleteditems: 'trash-id' }[JSON.parse(options.body).destinationId]; message.moved = true; return new Response('{"body":"discarded, not parsed"}', { status: 201 }); }
      if (message?.moved && id === failVerificationFor) throw Error('Verification unavailable');
      const value = p.endsWith('/mailFolders/inbox') ? { id: 'inbox-id' } : p.endsWith('/mailFolders/archive') ? { id: 'archive-id' } : p.endsWith('/mailFolders/deleteditems') ? { id: 'trash-id' } : { id, parentFolderId: message?.folder, from: { emailAddress: { address: message?.address } } };
      return Response.json(value);
    }
  });
  for (const file of ['unsubscribe-safety.js', 'outlook-archive.js', 'app.js']) vm.runInContext(fs.readFileSync(path.join(__dirname, '../dist', file), 'utf8'), context);
  const run = (code) => vm.runInContext(code, context);
  run(`state.mode = 'outlook'; outlookAccount = testAccount; outlookAccessToken = 'synthetic-read-token';
    setSenders([{ id: 'chosen', name: '<img src=x onerror=alert(1)>', address: 'sender@example.com', messageIds: ['message-one'], count: 1, future: ['review'], current: 'keep', category: 'review' }], 'outlook', testAccount.username);
    state.selected.add('chosen'); state.current.chosen = 'archive';`);
  return { calls, oauth, get, run, stored, click: (selector) => get(selector).handlers.click() };
}

function confirmManualRow(s, container, address) {
  const section = s.get(container).childNodes.find(section => section.childNodes.slice(1).some(row => row.childNodes[0].childNodes[1].textContent === address));
  const row = section.childNodes.slice(1).find(row => row.childNodes[0].childNodes[1].textContent === address);
  const button = row.childNodes.find(child => child.textContent === 'manualConfirmSender{}');
  assert.ok(button, 'Opened manual step offers explicit confirmation');
  button.handlers.click();
}

test('one reviewed click obtains Microsoft permission and then archives exactly once', async () => {
  const s = harness();
  await s.click('#review-actions');
  assert.equal(s.calls.length, 0);
  assert.equal(s.oauth.length, 0);
  await s.click('#confirm-actions');
  assert.equal(s.oauth.length, 1);
  assert.deepEqual(Array.from(s.oauth[0].scopes), ['Mail.ReadWrite']);
  assert.equal(s.calls.filter((call) => call.method === 'POST').length, 1);
  assert.equal(s.run('outlookActionResult.entries[0].status'), 'archived');
  assert.equal(s.get('#undo-outlook-archive').hidden, false);
  for (const call of s.calls) {
    assert.equal(call.headers.Prefer, 'IdType="ImmutableId"');
    assert.equal(call.redirect, 'error');
    assert.equal(call.credentials, 'omit');
    assert.equal(call.cache, 'no-store');
    assert.equal(call.url.origin, 'https://graph.microsoft.com');
    assert.ok(!/body|subject|preview|attachments/i.test(call.url.search));
  }
  await s.click('#undo-outlook-archive');
  assert.equal(s.run('outlookActionResult.entries[0].restored'), true);
});

function selectOutlookBatch(s, secondAction = 'trash') {
  s.run(`const second = { ...senders[0], id: 'second', name: 'Second', address: 'second@example.com', messageIds: ['message-two'] };
    senders.push(second); state.selected.add('second'); state.current.second = '${secondAction}'; state.future.second = new Set(['review']); updateBar();`);
}

test('Outlook review accepts multiple selected senders including No change without moving unselected emails', async () => {
  const s = harness();
  selectOutlookBatch(s, 'keep');
  assert.equal(s.get('#review-actions').disabled, false);
  await s.click('#review-actions');
  assert.equal(s.get('#review-dialog').open, true);
  assert.equal(s.run('autoUnsubscribeContext.actions.length'), 2);
  await s.click('#confirm-actions');
  assert.deepEqual(s.calls.filter((call) => call.method === 'POST').map((call) => call.url.pathname), ['/v1.0/me/messages/message-one/move']);
  assert.equal(s.run('Object.fromEntries(batchResultMetrics(autoUnsubscribeContext)).resultEmailsTreated'), 1);
});

test('one reviewed click creates an exact Outlook rule for future mail from the selected address', async () => {
  const s = harness();
  s.run("state.current.chosen='keep'; state.future.chosen=new Set(['auto_trash']); updateBar()");
  assert.equal(s.run("futureOptions(senders[0]).some(([value]) => value === 'auto_trash')"), true);
  await s.click('#review-actions');
  assert.equal(s.calls.length, 0);
  await s.click('#confirm-actions');
  assert.equal(s.oauth.length, 1);
  assert.deepEqual(Array.from(s.oauth[0].scopes), ['MailboxSettings.ReadWrite']);
  const rules = s.calls.filter((call) => call.url.pathname.endsWith('/mailFolders/inbox/messageRules'));
  assert.equal(rules.length, 1);
  const body = JSON.parse(rules[0].body);
  assert.equal(body.conditions.fromAddresses[0].emailAddress.address, 'sender@example.com');
  assert.deepEqual(body.actions, { delete: true, stopProcessingRules: true });
  assert.equal(s.run("Object.fromEntries(batchResultMetrics(autoUnsubscribeContext)).resultAutomaticDeletedRules"), 1);
  assert.equal(s.run("Object.fromEntries(batchResultMetrics(autoUnsubscribeContext)).resultActionsRemaining"), 0);
  assert.match(s.get('#success-title').textContent, /batchResultTitle/);
  assert.equal(s.get('#result-share').hidden, false);
  assert.equal(s.run("Object.keys(readTreatedSenders()).length"), 1);
});

test('combined Outlook moves and future routing request both permissions in the same popup', async () => {
  const s = harness();
  s.run("state.future.chosen=new Set(['auto_trash']); updateBar()");
  await s.click('#review-actions');
  await s.click('#confirm-actions');
  assert.equal(s.oauth.length, 1);
  assert.deepEqual(Array.from(s.oauth[0].scopes), ['Mail.ReadWrite', 'MailboxSettings.ReadWrite']);
  assert.equal(s.calls.filter((call) => call.url.pathname.endsWith('/move')).length, 1);
  assert.equal(s.calls.filter((call) => call.url.pathname.endsWith('/messageRules')).length, 1);
});

test('grouped Outlook routing creates rules only for individually selected addresses', async () => {
  const s = harness();
  s.run(`const first={...senders[0],id:'first',address:'first@example.com'};
    const second={...senders[0],id:'second',address:'second@example.com',messageIds:['message-two']};
    const grouped={...first,id:'grouped',name:'Example',address:first.address,count:2,members:[first,second],future:['review'],current:'keep'};
    setSenders([grouped],'outlook',testAccount.username); setMemberSelection(senders[0],senders[0].members[1],true);
    state.future.grouped=new Set(['auto_trash']); state.current.grouped='keep'; updateBar();`);
  await s.click('#review-actions');
  await s.click('#confirm-actions');
  const rules = s.calls.filter((call) => call.url.pathname.endsWith('/messageRules'));
  assert.equal(rules.length, 1);
  assert.equal(JSON.parse(rules[0].body).conditions.fromAddresses[0].emailAddress.address, 'second@example.com');
});

test('an unconfirmed Outlook rule remains pending and is never remembered as treated', async () => {
  const s = harness({ failRule: true });
  s.run("state.current.chosen='keep'; state.future.chosen=new Set(['auto_trash']); updateBar()");
  await s.click('#review-actions');
  await s.click('#confirm-actions');
  assert.equal(s.run("Object.fromEntries(batchResultMetrics(autoUnsubscribeContext)).resultActionsRemaining"), 1);
  assert.match(s.get('#success-title').textContent, /batchPendingTitle/);
  assert.equal(s.get('#result-share').hidden, true);
  assert.equal(s.run("Object.keys(readTreatedSenders()).length"), 0);
});

test('a sender without unsubscribe headers can be blocked directly in Gmail or guided to Outlook settings', async () => {
  const gmail = harness();
  gmail.run("state.mode='gmail'; senders[0].canUnsubscribe=false; senders[0].category='review'");
  assert.equal(gmail.run("futureOptions(senders[0]).some(([value]) => value === 'block')"), true);

  const outlook = harness();
  outlook.run("state.current.chosen='keep'; state.future.chosen=new Set(['block']); senders[0].canUnsubscribe=false; senders[0].category='review'; updateBar()");
  assert.equal(outlook.run("futureOptions(senders[0]).some(([value]) => value === 'block')"), true);
  assert.equal(outlook.get('#review-actions').disabled, false);
  await outlook.click('#review-actions');
  assert.equal(outlook.run("autoUnsubscribeContext.rows[0].reason"), 'outlook_block');
  assert.equal(outlook.get('#confirm-actions').hidden, true);
  outlook.run("globalThis.blockLink=manualUnsubscribeLink(senders[0])");
  const link = outlook.run('blockLink');
  assert.equal(link.href, 'https://outlook.live.com/mail/0/options/mail/junkEmail');
  assert.match(link.textContent, /outlookBlockOpen/);
  link.handlers.click({ preventDefault() { throw Error('valid Outlook settings link rejected'); } });
  assert.equal(outlook.run('allManualLinksOpened(autoUnsubscribeContext)'), true);
  assert.equal(outlook.get('#confirm-actions').hidden, false);
  assert.equal(outlook.calls.length, 0);
});

test('Outlook mixed archive, trash and manual unsubscribe batch executes each selected move once after one confirmation', async () => {
  const s = harness({ mailbox: [{ id: 'message-one', address: 'sender@example.com' }, { id: 'message-two', address: 'second@example.com' }] });
  selectOutlookBatch(s);
  s.run("state.future.second=new Set(['unsubscribe']); Object.assign(senders[1], {canUnsubscribe:true, unsubscribeUrl:'https://second.example.com/preferences'}); updateBar()");
  assert.equal(s.get('#review-actions').disabled, false);
  await s.click('#review-actions');
  assert.equal(s.calls.length, 0);
  assert.equal(s.get('#review-unsubscribes').childNodes.length, 1);
  await s.click('#confirm-actions');
  const moves = s.calls.filter((call) => call.method === 'POST');
  assert.deepEqual(moves.map((call) => [call.url.pathname, JSON.parse(call.body).destinationId]), [
    ['/v1.0/me/messages/message-one/move', 'archive'], ['/v1.0/me/messages/message-two/move', 'deleteditems']
  ]);
  const metrics = s.run('Object.fromEntries(batchResultMetrics(autoUnsubscribeContext))');
  assert.equal(metrics.resultEmailsTreated, 2);
  assert.equal(metrics.resultActionsRemaining, 1);
  assert.equal(s.get('#result-share').hidden, true);
  assert.deepEqual(s.get('#success-report').childNodes.map((row) => row.childNodes[0].childNodes[0].textContent), ['✓', '✋']);
  assert.equal(s.calls.some((call) => call.url.hostname !== 'graph.microsoft.com'), false);
});

test('Outlook batch consent cancellation and changed reviewed choices never move messages', async () => {
  const cancelled = harness({ cancelPermission: true }); selectOutlookBatch(cancelled);
  await cancelled.click('#review-actions'); await cancelled.click('#confirm-actions');
  assert.equal(cancelled.calls.length, 0);
  assert.equal(cancelled.run('autoUnsubscribeContext.phase'), 'review');
  const changed = harness(); selectOutlookBatch(changed);
  await changed.click('#review-actions');
  changed.run("state.current.second='archive'");
  await changed.click('#confirm-actions');
  assert.equal(changed.calls.length, 0);
  assert.equal(changed.run('autoUnsubscribeContext.phase'), 'blocked');
});

test('Outlook batch stops after an unverifiable move and does not count remaining senders as treated', async () => {
  const s = harness({ failVerificationFor: 'message-one' }); selectOutlookBatch(s);
  await s.click('#review-actions'); await s.click('#confirm-actions');
  assert.equal(s.calls.filter((call) => call.method === 'POST').length, 1);
  assert.equal(s.calls.some((call) => call.url.pathname.includes('message-two')), false);
  const metrics = s.run('Object.fromEntries(batchResultMetrics(autoUnsubscribeContext))');
  assert.equal(metrics.resultEmailsTreated, 0);
  assert.equal(metrics.resultActionsRemaining, 2);
  assert.equal(s.get('#result-share').hidden, true);
});

test('manual links open directly without confirmation dialogs, network calls or a success claim', () => {
  const s = harness();
  s.run("Object.assign(senders[0], { canUnsubscribe: true, unsubscribeUrl: 'https://news.example.com/preferences?token=synthetic' }); globalThis.link = manualUnsubscribeLink(senders[0]);");
  const link = s.run('link');
  assert.equal(link.href, 'https://news.example.com/preferences?token=synthetic');
  assert.equal(link.target, '_blank');
  assert.equal(link.rel, 'noopener noreferrer');
  link.handlers.click({ preventDefault() { throw Error('valid link rejected'); } });
  assert.equal(s.calls.length, 0);
  s.run('setSenders([])');
  let prevented = false;
  link.handlers.click({ preventDefault() { prevented = true; } });
  assert.equal(prevented, true);
});

test('invalid unsubscribe destinations do not produce an external link', () => {
  const s = harness();
  assert.equal(s.run("manualUnsubscribeLink({ unsubscribeUrl: 'https://127.0.0.1/' })"), null);
});

test('cancelled Microsoft permission never moves or reads emails', async () => {
  const s = harness({ cancelPermission: true });
  await s.click('#review-actions');
  await s.click('#confirm-actions');
  assert.equal(s.calls.length, 0);
  assert.equal(s.run('outlookPlan'), null);
  assert.equal(s.get('#confirm-actions').disabled, true);
});

function useGmailNewsletter(s) {
  s.run("state.mode = 'gmail'; mailboxIdentity = 'owner@example.com'; gmailAccessToken = 'synthetic-read-token'; state.current.chosen = 'keep'; state.future.chosen = new Set(['unsubscribe']); Object.assign(senders[0], { canUnsubscribe: true, unsubscribeUrl: 'https://news.beehiiv.com/preferences' });");
}

for (const size of [50, 51]) test(`Gmail review ${size === 50 ? 'accepts 50' : 'rejects 51'} selected unsubscribe addresses`, async () => {
  const s = harness({ automatic: true });
  useGmailNewsletter(s);
  s.run(`const batch = Array.from({length: ${size}}, (_, i) => ({...senders[0], id: 'sender-' + i,
    address: 'sender-' + i + '@example.com', messageIds: ['message-' + i], future: ['unsubscribe'], current: 'keep'}));
    setSenders(batch, 'gmail', 'owner@example.com');
    senders.forEach(sender => { state.selected.add(sender.id); state.future[sender.id] = new Set(['unsubscribe']); state.current[sender.id] = 'keep'; }); updateBar();`);
  await s.click('#review-actions');
  assert.equal(s.run('autoUnsubscribeContext.senders.length'), size);
  if (size === 50) {
    assert.equal(s.run('autoUnsubscribeContext.phase'), 'review');
    assert.equal(s.get('#confirm-actions').disabled, false);
    const request = s.calls.find(call => call.url.pathname.endsWith('/prepare'));
    assert.equal(JSON.parse(request.body).items.length, 50);
    assert.equal(s.calls.filter(call => call.url.pathname.endsWith('/execute')).length, 0);
  } else {
    assert.equal(s.run('autoUnsubscribeContext.error'), 'autoLimitSelection');
    assert.equal(s.get('#confirm-actions').disabled, true);
    assert.equal(s.calls.length, 0);
    assert.equal(s.oauth.length, 0);
  }
});

test('dashboard selection prepares on Review; one final confirmation executes exactly once', async () => {
  const s = harness({ automatic: true }); useGmailNewsletter(s);
  await s.click('#review-actions');
  assert.equal(s.calls.length, 1);
  assert.match(s.calls[0].url.pathname, /prepare$/);
  assert.equal(s.get('#confirm-actions').disabled, false);
  assert.match(s.get('#confirm-actions').textContent, /batchConfirm/);
  assert.equal(s.run('autoUnsubscribeContext.phase'), 'review');
  assert.equal(s.get('#review-action-status').childNodes[0].textContent, 'i');
  s.run("autoUnsubscribeContext.phase='sending'; renderBatchReview()");
  assert.equal(s.get('#review-loading').hidden, false);
  assert.match(s.get('#review-loading-copy').textContent, /applyingEmailActions/);
  s.run("autoUnsubscribeContext.phase='review'; renderBatchReview()");
  await s.click('#confirm-actions');
  await s.click('#confirm-actions');
  assert.equal(s.calls.length, 2);
  assert.match(s.calls[1].url.pathname, /execute$/);
  assert.equal(JSON.parse(s.calls[1].body).confirmed, true);
  assert.equal(s.run('autoUnsubscribeContext.phase'), 'done');
  assert.equal(s.get('#result-share').hidden, false);
  assert.equal(s.run("Object.keys(readTreatedSenders()).length"), 1);
  s.run('setSenders([])');
  assert.equal(s.run('autoUnsubscribeContext'), null);
});

test('mixed selections retain all senders; manual ones do not block automatic ones', async () => {
  const s = harness({ automatic: true }); useGmailNewsletter(s);
  s.run("const manual = {...senders[0], id:'manual', name:'Manual', address:'manual@example.com', messageIds:['manual-id'], future:['unsubscribe'], current:'keep', unsubscribeUrl:'https://news.example.com/preferences'}; senders.push(manual); state.selected.add('manual'); state.future.manual=new Set(['unsubscribe']); state.current.manual='keep';");
  await s.click('#review-actions');
  assert.equal(s.get('#review-unsubscribes').childNodes.length, 2);
  assert.equal(s.run('autoUnsubscribeContext.senders.length'), 2);
  await s.click('#confirm-actions');
  assert.equal(s.run('autoUnsubscribeContext.results[0].status'), 'requested');
  assert.equal(s.run('autoUnsubscribeContext.results[1].status'), 'manual');
  assert.equal(s.get('#result-unsubscribes').childNodes.length, 2);
  assert.equal(s.get('#result-share').hidden, true);
  assert.match(s.get('#success-title').textContent, /batchPendingTitle/);
});

test('Outlook manual-only review sends nothing; clicking greys the link and completion is confirmed separately', async () => {
  const s = harness();
  s.run("state.current.chosen='keep'; state.future.chosen=new Set(['unsubscribe']); Object.assign(senders[0],{canUnsubscribe:true,unsubscribeUrl:'https://news.example.com/preferences'});");
  await s.click('#review-actions');
  assert.equal(s.calls.length, 0);
  assert.equal(s.get('#confirm-actions').hidden, true);
  assert.equal(s.get('#review-unsubscribes').childNodes.length, 1);
  assert.equal(s.run('autoUnsubscribeContext.phase'), 'review');
  s.run("globalThis.manualLink = manualUnsubscribeLink(senders[0])");
  const link = s.run('manualLink');
  link.handlers.click({ preventDefault() { throw Error('valid link rejected'); } });
  assert.match(link.textContent, /manualLinkOpened/);
  assert.equal(link.tabIndex, -1);
  assert.equal(s.run('autoUnsubscribeContext.manualConfirmed'), false);
  assert.equal(s.run("Object.keys(readTreatedSenders()).length"), 0);
  assert.equal(s.get('#confirm-actions').hidden, false);
  assert.match(s.get('#confirm-actions').textContent, /manualFinish/);
  s.run('renderBatchReview(); globalThis.rerenderedLink = manualUnsubscribeLink(senders[0])');
  let prevented = false;
  s.run('rerenderedLink').handlers.click({ preventDefault() { prevented = true; } });
  assert.equal(prevented, true);
  await s.click('#confirm-actions');
  assert.equal(s.calls.length, 0);
  assert.equal(s.run('autoUnsubscribeContext.manualConfirmed'), true);
  assert.equal(s.run("Object.keys(readTreatedSenders()).length"), 1);
  assert.equal(s.get('#result-share').hidden, false);
  assert.equal(s.run("Object.fromEntries(batchResultMetrics(autoUnsubscribeContext)).resultManualStepsProcessed"), 1);
  assert.equal(s.run("Object.fromEntries(batchResultMetrics(autoUnsubscribeContext)).resultActionsRemaining"), 0);
  assert.equal(s.run("Object.fromEntries(batchResultMetrics(autoUnsubscribeContext)).resultAutomaticUnsubscribes"), 0);
});

test('each confirmed manual sender persists even when other manual steps remain unfinished', async () => {
  const s = harness({ automatic: true }); useGmailNewsletter(s);
  s.run("gmailPreviewResult={profile:{emailAddress:'owner@example.com'}}; state.current.chosen='trash'; const manual={...senders[0],id:'manual',name:'Manual',address:'manual@example.com',messageIds:['manual-id'],unsubscribeUrl:'https://news.example.com/preferences'}; senders.push(manual); state.selected.add('manual'); state.future.manual=new Set(['unsubscribe']); state.current.manual='keep';");
  s.run("const second={...senders[1],id:'second',name:'Second manual',address:'second@example.com',messageIds:['second-id'],unsubscribeUrl:'https://second.example.com/preferences'}; senders.push(second); state.selected.add('second'); state.future.second=new Set(['unsubscribe']); state.current.second='keep';");
  await s.click('#review-actions');
  s.run("globalThis.link = manualUnsubscribeLink(senders[1])");
  s.run('link').handlers.click({ preventDefault() { throw Error('valid link rejected'); } });
  s.run('renderBatchReview()');
  assert.equal(s.run("Object.keys(readTreatedSenders()).length"), 0);
  confirmManualRow(s, '#review-unsubscribes', 'manual@example.com');
  assert.equal(s.run("Object.keys(readTreatedSenders()).length"), 1);
  await s.click('#confirm-actions');
  assert.equal(s.get('#result-share').hidden, true);
  assert.equal(s.get('#result-manual-confirm').disabled, true);
  await s.click('#result-manual-confirm');
  assert.equal(s.run('autoUnsubscribeContext.manualConfirmed'), false);
  assert.equal(s.get('#result-share').hidden, true);
  const requests = s.calls.length;
  const before = s.run("Object.fromEntries(batchResultMetrics(autoUnsubscribeContext))");
  assert.equal(before.resultEmailsTreated, 1);
  assert.equal(before.resultManualStepsProcessed, 1);
  assert.equal(before.resultManualPending, 1);
  assert.equal(before.resultActionsRemaining, 1);
  assert.equal(before.resultAutomaticUnsubscribes, 1);
  const icons = () => s.get('#success-report').childNodes.map((row) => row.childNodes[0].childNodes[0].textContent);
  assert.deepEqual(icons(), ['✓', '✓', '✋']);
  s.run("globalThis.lastLink = manualUnsubscribeLink(senders[2])");
  s.run('lastLink').handlers.click({ preventDefault() { throw Error('valid link rejected'); } });
  await new Promise((resolve) => setTimeout(resolve, 5));
  assert.equal(s.run("Object.fromEntries(batchResultMetrics(autoUnsubscribeContext)).resultActionsRemaining"), 1);
  assert.equal(s.run("Object.keys(readTreatedSenders()).length"), 2);
  assert.equal(s.get('#result-manual-confirm').disabled, false);
  confirmManualRow(s, '#result-unsubscribes', 'second@example.com');
  assert.deepEqual(icons(), ['✓', '✓', '✓']);
  assert.equal(s.run("Object.fromEntries(batchResultMetrics(autoUnsubscribeContext)).resultManualStepsProcessed"), 2);
  assert.equal(s.run("Object.fromEntries(batchResultMetrics(autoUnsubscribeContext)).resultActionsRemaining"), 0);
  assert.equal(s.get('#result-manual-confirm').hidden, true);
  assert.equal(s.get('#result-share').hidden, false);
  assert.match(s.get('#success-title').textContent, /batchResultTitle/);
  assert.equal(s.calls.length, requests);
  assert.equal(s.run('autoUnsubscribeContext.results[1].status'), 'manual');
  const after = s.run("Object.fromEntries(batchResultMetrics(autoUnsubscribeContext))");
  assert.equal(after.resultManualStepsProcessed, 2);
  assert.equal(after.resultActionsRemaining, 0);
  assert.equal(after.resultEmailsTreated, 1);
  assert.equal(s.run("Object.keys(readTreatedSenders()).length"), 3);
});

test('confirmed future actions are remembered per provider and account while existing-email actions remain available', () => {
  const s = harness();
  s.run("state.mode='gmail'; mailboxIdentity='owner@example.com'; rememberTreatedSender(senders[0]); const rescanned={...senders[0],id:'again',future:['unsubscribe'],current:'keep',futureTreated:false}; setSenders([rescanned],'gmail','owner@example.com'); renderSenders()");
  assert.equal(s.run("senders[0].futureTreated"), true);
  assert.equal(s.run("futureValue(senders[0].id)"), 'treated');
  assert.deepEqual(Array.from(s.run("futureOptions(senders[0])")[0]), ['treated', 'senderAlreadyTreated{}']);
  assert.equal(s.get('#sender-list').childNodes[0].childNodes[4].childNodes[1].childNodes[0].disabled, true);
  assert.equal(s.get('#sender-list').childNodes[0].childNodes[5].childNodes[1].childNodes[0].disabled, false);
  s.run("setSenders([{...senders[0],id:'other',future:['unsubscribe'],futureTreated:false}],'gmail','other@example.com')");
  assert.equal(s.run("senders[0].futureTreated"), false);
  s.run("setSenders([{...senders[0],id:'outlook',future:['unsubscribe'],futureTreated:false}],'outlook','owner@example.com')");
  assert.equal(s.run("senders[0].futureTreated"), false);
});

test('manual confirmation survives closing review and reloading before the whole batch is finished', async () => {
  const s = harness();
  s.run(`state.current.chosen='keep'; state.future.chosen=new Set(['unsubscribe']);
    Object.assign(senders[0],{canUnsubscribe:true,unsubscribeUrl:'https://news.example.com/preferences'});
    const second={...senders[0],id:'second',address:'second@example.com',messageIds:['second-id']};
    senders.push(second); state.selected.add('second'); state.current.second='keep'; state.future.second=new Set(['unsubscribe']);`);
  await s.click('#review-actions');
  s.run('globalThis.link=manualUnsubscribeLink(senders[0])');
  s.run('link').handlers.click({ preventDefault() { throw Error('valid link rejected'); } });
  s.run('renderBatchReview()');
  confirmManualRow(s, '#review-unsubscribes', 'sender@example.com');
  await s.click('#cancel-actions');
  assert.equal(s.calls.length, 0);
  const reloaded = harness({ stored: s.stored });
  assert.equal(reloaded.run('senders[0].futureTreated'), true);
  assert.equal(reloaded.run("senderWasTreated({address:'second@example.com'})"), false);
  assert.equal(reloaded.run("senderWasTreated({address:'sender@example.com'}, 'outlook', 'other@example.com')"), false);
});

test('an accepted unsubscribe is remembered even if subsequent Gmail filter loading fails', async () => {
  const s = harness({ automatic: true, failGmailFilters: true });
  useGmailNewsletter(s);
  s.run(`gmailPreviewResult={profile:{emailAddress:'owner@example.com'}};
    state.future.chosen=new Set(['unsubscribe','block']);`);
  await s.click('#review-actions');
  await s.click('#confirm-actions');
  assert.equal(s.run("autoUnsubscribeContext.results[0].status"), 'requested');
  assert.equal(s.run("senderWasTreated(senders[0])"), true);
  assert.equal(s.calls.some(call => call.url.pathname.endsWith('/settings/filters')), true);
  const reloaded = harness({ stored: s.stored });
  reloaded.run(`state.mode='gmail'; setSenders([{...senders[0],id:'new-scan',future:['unsubscribe']}], 'gmail', 'owner@example.com')`);
  assert.equal(reloaded.run("futureValue('new-scan')"), 'treated');
});

test('mixed groups skip previously treated addresses for unsubscribe while cleaning all selected messages', async () => {
  const s = harness({ automatic: true });
  useGmailNewsletter(s);
  s.run(`gmailPreviewResult={profile:{emailAddress:'owner@example.com'}};
    const first={...senders[0],id:'first',name:'Acme Newsletter',address:'news@acme.com',messageIds:['first-message'],future:['unsubscribe']};
    const second={...first,id:'second',address:'updates@acme.com',messageIds:['second-message']};
    rememberTreatedSender(first);
    setSenders(groupRelatedSenders([first,second],'gmail'),'gmail','owner@example.com');
    setSenderSelection(senders[0],true); updateBar();`);
  await s.click('#review-actions');
  const prepared = s.calls.find(call => call.url.pathname.endsWith('/prepare'));
  assert.deepEqual(JSON.parse(prepared.body).items.map(item => item.sender), ['updates@acme.com']);
  assert.equal(s.run('autoUnsubscribeContext.actions[0].members.length'), 2);
  await s.click('#confirm-actions');
  assert.deepEqual(s.calls.filter(call => call.url.pathname.endsWith('/trash')).map(call => call.url.pathname).sort(),
    ['/gmail/v1/users/me/messages/first-message/trash', '/gmail/v1/users/me/messages/second-message/trash']);
  s.run(`setSenders(senders,'gmail','owner@example.com')`);
  assert.equal(s.run('senders[0].futureTreated'), true);
});

for (const provider of ['gmail', 'outlook', 'demo']) test(`${provider} defaults existing emails to trash without selecting senders or offering Protect`, () => {
  const s = harness();
  s.run(`state.mode='${provider}';
    setSenders([{...senders[0],id:'protected',category:'keep',future:['keep'],current:'keep'},
      {...senders[0],id:'newsletter',category:'unsubscribe',canUnsubscribe:true,future:['unsubscribe'],current:'archive'}], '${provider}', 'owner@example.com');
    renderSenders(); updateBar();`);
  assert.deepEqual(Array.from(s.run('Object.values(state.current)')), ['trash', 'trash']);
  assert.equal(s.run("futureValue('protected')"), 'review');
  assert.equal(s.run("senders.some(sender => futureOptions(sender).some(([value]) => value === 'keep'))"), false);
  assert.equal(s.run('state.selected.size'), 0);
  assert.equal(s.get('#review-actions').disabled, true);
  assert.equal(s.calls.length, 0);
  for (const row of s.get('#sender-list').childNodes) {
    const choices = row.childNodes[5].childNodes[1].childNodes[0].childNodes;
    assert.equal(choices.find(option => option.selected).value, 'trash');
  }
});

test('different addresses from the same sender are grouped into one dashboard row', () => {
  const s = harness();
  s.run(`globalThis.grouped = groupRelatedSenders([
    { id:'one', name:'Acme Newsletter', address:'news@acme.com', count:3, messageIds:['a','b','c'], canUnsubscribe:true, unsubscribeUrl:'https://acme.com/u/one', unsubscribeMessageId:'a', oneClick:false, future:['unsubscribe'], current:'keep', category:'unsubscribe' },
    { id:'two', name:'Acme Email', address:'hello@acme.com', count:2, messageIds:['d','e'], canUnsubscribe:true, unsubscribeUrl:'https://acme.com/u/two', unsubscribeMessageId:'d', oneClick:true, future:['unsubscribe'], current:'keep', category:'unsubscribe' }
  ], 'gmail')`);
  assert.equal(s.run('grouped.length'), 1);
  assert.equal(s.run('grouped[0].members.length'), 2);
  assert.equal(s.run('grouped[0].count'), 5);
  assert.deepEqual(Array.from(s.run('grouped[0].messageIds')), ['a', 'b', 'c', 'd', 'e']);
  assert.deepEqual(Array.from(s.run('senderAddresses(grouped[0])')), ['news@acme.com', 'hello@acme.com']);
});

test('related brand domains are grouped and every observed address is blocked', async () => {
  const s = harness();
  s.run(`state.mode='gmail'; mailboxIdentity='owner@example.com'; gmailAccessToken='synthetic-read-token'; gmailPreviewResult={profile:{emailAddress:'owner@example.com'}};
    const grouped=groupRelatedSenders([
      {id:'one',name:'Pinterest',address:'ideas@em.pinterest.com',count:2,messageIds:['a','b'],canUnsubscribe:false,future:['review'],current:'keep',category:'review'},
      {id:'two',name:'Pinterest Updates',address:'recommendations@pinterestmail.com',count:3,messageIds:['c','d','e'],canUnsubscribe:false,future:['review'],current:'keep',category:'review'}
    ],'gmail'); setSenders(grouped,'gmail','owner@example.com'); state.selected.add(grouped[0].id); state.future[grouped[0].id]=new Set(['block']); updateBar();`);
  assert.equal(s.run('senders.length'), 1);
  assert.equal(s.run('senders[0].name'), 'Pinterest');
  assert.deepEqual(Array.from(s.run('senderAddresses(senders[0])')), ['ideas@em.pinterest.com', 'recommendations@pinterestmail.com']);
  await s.click('#review-actions');
  await s.click('#confirm-actions');
  const filters = s.calls.filter((call) => call.method === 'POST' && call.url.pathname.endsWith('/settings/filters'));
  assert.deepEqual(filters.map((call) => JSON.parse(call.body).criteria.from), ['ideas@em.pinterest.com', 'recommendations@pinterestmail.com']);
});

test('individual address selection limits grouped actions to that address and its emails', async () => {
  const s = harness();
  s.run(`state.mode='gmail'; mailboxIdentity='owner@example.com'; gmailAccessToken='synthetic-read-token'; gmailPreviewResult={profile:{emailAddress:'owner@example.com'}};
    const grouped=groupRelatedSenders([
      {id:'one',name:'Eleven Labs Inc.',address:'failed-payments+acct_123@stripe.com',count:2,messageIds:['a','b'],canUnsubscribe:false,future:['review'],current:'keep',category:'review'},
      {id:'two',name:'Eleven Labs Inc.',address:'billing+acct_123@stripe.com',count:3,messageIds:['c','d','e'],canUnsubscribe:false,future:['review'],current:'keep',category:'review'}
    ],'gmail'); setSenders(grouped,'gmail','owner@example.com');
    setMemberSelection(grouped[0], grouped[0].members[1], true); state.future[grouped[0].id]=new Set(['block']); updateBar();`);
  assert.equal(s.run('stats().emails'), 3);
  assert.deepEqual(Array.from(s.run('selectedSenderMembers(senders[0]).map((member) => member.address)')), ['billing+acct_123@stripe.com']);
  await s.click('#review-actions');
  assert.deepEqual(Array.from(s.run('autoUnsubscribeContext.actions[0].messageIds')), ['c', 'd', 'e']);
  await s.click('#confirm-actions');
  const filters = s.calls.filter((call) => call.method === 'POST' && call.url.pathname.endsWith('/settings/filters'));
  assert.deepEqual(filters.map((call) => JSON.parse(call.body).criteria.from), ['billing+acct_123@stripe.com']);
});

test('a partial address selection distinguishes selected emails from the group total', () => {
  const s = harness();
  s.run(`const grouped=groupRelatedSenders([
      {id:'one',name:'Eleven Labs Inc.',address:'failed-payments+acct_123@stripe.com',count:2,messageIds:['a','b'],canUnsubscribe:false,future:['review'],current:'keep',category:'review'},
      {id:'two',name:'Eleven Labs Inc.',address:'billing+acct_123@stripe.com',count:3,messageIds:['c','d','e'],canUnsubscribe:false,future:['review'],current:'keep',category:'review'}
    ],'gmail'); setSenders(grouped,'gmail','owner@example.com');
    setMemberSelection(grouped[0], grouped[0].members[1], true);
    globalThis.partialCount = senderEmailCountDisplay(grouped[0]);`);
  assert.equal(s.run('partialCount.value'), 'selectedEmailFraction{"count":3,"total":5}');
  assert.equal(s.run('partialCount.label'), 'emailsSelected{}');
  assert.equal(s.run('partialCount.partial'), true);
});

test('a grouped sender shows how many addresses were already treated', () => {
  const s = harness();
  s.run(`state.mode='gmail'; mailboxIdentity='owner@example.com';
    const first={id:'one',name:'Pinterest',address:'ideas@em.pinterest.com',count:2,messageIds:['a'],canUnsubscribe:false,future:['review'],current:'keep',category:'review'};
    rememberTreatedSender(first);
    const grouped=groupRelatedSenders([first,{...first,id:'two',address:'recommendations@pinterestmail.com',messageIds:['b']}],'gmail');
    setSenders(grouped,'gmail','owner@example.com'); renderSenders();`);
  assert.equal(s.run('senders[0].futureTreated'), false);
  assert.equal(s.run('senders[0].treatedAddressCount'), 1);
  assert.match(s.get('#sender-list').childNodes[0].childNodes[2].childNodes.at(-1).textContent, /senderAddressesAlreadyTreated/);
});

test('unrelated newsletters on a shared delivery platform are never grouped', () => {
  const s = harness();
  s.run(`globalThis.grouped = groupRelatedSenders([
    { id:'one', name:'Alpha Brief', address:'alpha@mail.beehiiv.com', count:1, messageIds:['a'], canUnsubscribe:true, unsubscribeUrl:'https://alpha.beehiiv.com/u' },
    { id:'two', name:'Beta Daily', address:'beta@mail.beehiiv.com', count:1, messageIds:['b'], canUnsubscribe:true, unsubscribeUrl:'https://beta.beehiiv.com/u' }
  ], 'gmail')`);
  assert.equal(s.run('grouped.length'), 2);
});

test('different services from a multi-service company are not grouped by domain alone', () => {
  const s = harness();
  s.run(`globalThis.grouped = groupRelatedSenders([
    { id:'accounts', name:'Google Accounts', address:'no-reply@accounts.google.com', count:8, messageIds:['a'], canUnsubscribe:false },
    { id:'alerts', name:'Google Alerts', address:'googlealerts-noreply@google.com', count:12, messageIds:['b'], canUnsubscribe:false },
    { id:'play', name:'Google Play', address:'googleplay-noreply@google.com', count:7, messageIds:['c'], canUnsubscribe:false },
    { id:'maps', name:'Google Maps', address:'google-maps-noreply@google.com', count:6, messageIds:['d'], canUnsubscribe:false }
  ], 'gmail')`);
  assert.equal(s.run('grouped.length'), 4);
  assert.deepEqual(Array.from(s.run('grouped.map((sender) => sender.name)')), ['Google Accounts', 'Google Alerts', 'Google Maps', 'Google Play']);
});

test('shared transaction platforms group only an identical merchant brand', () => {
  const s = harness();
  s.run(`globalThis.grouped = groupRelatedSenders([
    { id:'one', name:'Eleven Labs Inc.', address:'failed-payments+acct_123@stripe.com', count:1, messageIds:['a'], canUnsubscribe:false },
    { id:'two', name:'Eleven Labs Inc.', address:'billing+acct_123@stripe.com', count:1, messageIds:['b'], canUnsubscribe:false },
    { id:'three', name:'Another Merchant', address:'billing+acct_456@stripe.com', count:1, messageIds:['c'], canUnsubscribe:false }
  ], 'gmail')`);
  assert.equal(s.run('grouped.length'), 2);
  assert.equal(s.run("grouped.find((sender) => sender.name === 'Eleven Labs Inc.').members.length"), 2);
});

test('a grouped sender exposes unsubscribe only when every address supports it', () => {
  const s = harness();
  s.run(`globalThis.grouped = groupRelatedSenders([
    { id:'one', name:'Acme Newsletter', address:'news@acme.com', count:1, messageIds:['a'], canUnsubscribe:true, unsubscribeUrl:'https://acme.com/u' },
    { id:'two', name:'Acme Email', address:'hello@acme.com', count:1, messageIds:['b'], canUnsubscribe:false, unsubscribeUrl:null }
  ], 'gmail')`);
  assert.equal(s.run('grouped.length'), 1);
  assert.equal(s.run('grouped[0].canUnsubscribe'), false);
  assert.equal(s.run("futureOptions(grouped[0]).some(([value]) => value === 'unsubscribe')"), false);
  assert.equal(s.run("futureOptions(grouped[0]).some(([value]) => value === 'block')"), true);
});

test('uncertain senders cannot select Protect and their explanatory line differs from the title', () => {
  const s = harness();
  s.run("state.mode='gmail'; senders[0].reasonKey='gmailUncertainReason'");
  assert.equal(s.run("futureOptions(senders[0]).some(([value]) => value === 'keep')"), false);
  assert.equal(s.run("senderStatus(senders[0]).childNodes[1].childNodes[1].textContent.startsWith('uncertainSenderDetail')"), true);
});

test('combined Gmail cleanup prepares with the renewed token and writes only after final confirmation', async () => {
  const s = harness({ automatic: true }); useGmailNewsletter(s);
  s.run("gmailPreviewResult={profile:{emailAddress:'owner@example.com'}}; state.current.chosen='trash';");
  await s.click('#review-actions');
  assert.equal(s.oauth.length, 1);
  assert.equal(s.calls.filter((call) => call.method === 'POST' && call.url.hostname === 'gmail.googleapis.com').length, 0);
  const prepared = s.calls.find((call) => call.url.pathname.endsWith('/prepare'));
  assert.equal(JSON.parse(prepared.body).token, 'synthetic-new-gmail-token');
  await s.click('#confirm-actions');
  const mailboxWrites = s.calls.filter((call) => call.method === 'POST' && call.url.hostname === 'gmail.googleapis.com');
  assert.equal(mailboxWrites.length, 1);
  assert.match(mailboxWrites[0].url.pathname, /message-one\/trash$/);
  assert.equal(s.run('autoUnsubscribeContext.results[0].status'), 'requested');
});

test('a different Gmail account during permission refuses both preparation and cleanup', async () => {
  const s = harness({ automatic: true, wrongAccount: true }); useGmailNewsletter(s);
  s.run("gmailPreviewResult={profile:{emailAddress:'owner@example.com'}}; state.current.chosen='trash';");
  await s.click('#review-actions');
  await s.click('#confirm-actions');
  assert.equal(s.calls.length, 1);
  assert.match(s.calls[0].url.pathname, /profile$/);
  assert.equal(s.run('autoUnsubscribeContext.phase'), 'blocked');
});

test('an uncertain unsubscribe response stays uncertain, is not retried and does not invite sharing', async () => {
  const s = harness({ automatic: true, executeFailure: true }); useGmailNewsletter(s);
  await s.click('#review-actions');
  await s.click('#confirm-actions');
  await s.click('#confirm-actions');
  assert.equal(s.calls.length, 2);
  assert.equal(s.run('autoUnsubscribeContext.results[0].status'), 'uncertain');
  assert.equal(s.get('#result-share').hidden, true);
});

test('changing the reviewed action refuses execution; cancellation discards the plan', async () => {
  const s = harness({ automatic: true }); useGmailNewsletter(s);
  await s.click('#review-actions');
  s.run("state.current.chosen='trash'");
  await s.click('#confirm-actions');
  assert.equal(s.calls.length, 1);
  assert.equal(s.run('autoUnsubscribeContext.phase'), 'blocked');
  await s.click('#cancel-actions');
  assert.equal(s.run('autoUnsubscribeContext'), null);
});

test('Microsoft account mismatch refuses the write token', async () => {
  const s = harness({ wrongAccount: true });
  await s.click('#review-actions');
  await s.click('#confirm-actions');
  assert.equal(s.calls.length, 0);
  assert.equal(s.run('outlookWriteAccessToken'), null);
});

test('transport refuses deletion, rules, unapproved moves and arbitrary URLs', async () => {
  const s = harness();
  for (const command of [
    `outlookRequest('me/messages/message-one', {}, {method:'DELETE'})`,
    `outlookRequest('me/mailFolders/inbox/messageRules', {}, {method:'POST', body:{}})`,
    `outlookRequest('me/messages/message-one/move', {}, {method:'POST', body:{destinationId:'archive'}})`,
    `outlookRequest('https://evil.test')`,
    `outlookRequest('../users/another-account/messages')`
  ]) await assert.rejects(s.run(command), /NOT-ALLOWED/);
  assert.equal(s.calls.length, 0);
});

test('embedded context cannot begin the Outlook action flow', async () => {
  const s = harness({ embedded: true });
  await s.click('#review-actions');
  assert.equal(s.run('outlookPlan'), null);
  assert.equal(s.calls.length, 0);
  assert.match(s.get('#provider-dialog-copy').textContent, /secureWindowRequired/);
});

test('unsupported cleanup actions remain disabled', async () => {
  const s = harness();
  s.run(`state.current.chosen = 'spam'`);
  await s.click('#review-actions');
  assert.equal(s.run('outlookPlan'), null);
  assert.equal(s.calls.length, 0);
});

test('all six languages contain every new Outlook archive message', () => {
  const app = fs.readFileSync(path.join(__dirname, '../dist/app.js'), 'utf8');
  const keys = [...new Set(app.match(/"(?:outlook(?:Archive|Trash|Auto|Rule|Undo|Exact|Write|Final|Authorise|Confirm|Permission|Single|Future|NoChanges|Newsletter|Test)[A-Za-z]*|resultAutomaticDeletedRules|unsub[A-Z][A-Za-z]*|auto[A-Z][A-Za-z]*|batch[A-Z][A-Za-z]*|secureWindowRequired)"/g).map((key) => JSON.parse(key)))];
  const context = vm.createContext({ window: {}, localStorage: { getItem() {}, setItem() {} }, navigator: { language: 'en' }, CustomEvent: function () {}, document: { documentElement: {}, querySelector() {}, querySelectorAll: () => [], addEventListener() {}, dispatchEvent() {} } });
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../dist/i18n.js'), 'utf8'), context);
  for (const lang of ['en', 'pt', 'it', 'es', 'fr', 'de']) {
    context.window.WeHateSpamI18n.setLanguage(lang, false);
    for (const key of keys) assert.notEqual(context.window.WeHateSpamI18n.t(key), key, lang + ': ' + key);
  }
});

test('trash has its own destination and executes only after its single reviewed confirmation', async () => {
  const s = harness();
  s.run(`state.current.chosen = 'trash'`);
  await s.click('#review-actions');
  assert.equal(s.run('outlookPlan.action'), 'trash');
  assert.match(s.get('#review-summary').childNodes[1].textContent, /outlookTrashExactPlan/);
  await s.click('#confirm-actions');
  assert.equal(s.run('outlookActionResult.entries[0].status'), 'trashed');
  assert.match(s.get('#success-title').textContent, /outlookTrashComplete/);
  await s.click('#undo-outlook-archive');
  assert.equal(s.run('outlookActionResult.entries[0].restored'), true);
  assert.deepEqual(s.calls.filter((call) => call.method === 'POST').map((call) => JSON.parse(call.body).destinationId), ['deleteditems', 'inbox']);
});
