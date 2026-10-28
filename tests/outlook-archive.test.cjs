const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const source = fs.readFileSync(require('node:path').join(__dirname, '../dist/outlook-archive.js'), 'utf8');

function setup(options = {}) {
  const context = vm.createContext({ window: {} });
  vm.runInContext(source, context);
  const api = context.window.OutlookArchive;
  let session = { accountKey: 'account-A', session: 1 };
  const calls = [];
  const folders = { one: 'inbox-id', two: 'inbox-id', other: 'inbox-id' };
  const addresses = { one: 'newsletter@example.com', two: 'newsletter@example.com', other: 'other@example.com' };
  const request = async (path, input = {}) => {
    calls.push({ path, ...input });
    if (path === 'me/mailFolders/inbox') return { id: 'inbox-id' };
    if (path === 'me/mailFolders/archive') return { id: 'archive-id' };
    if (path === 'me/mailFolders/deleteditems') return { id: 'trash-id' };
    const id = decodeURIComponent(path.split('/')[2]);
    if (input.method === 'POST') {
      assert.equal(input.discardBody, true);
      assert.ok(['archive', 'inbox', 'deleteditems'].includes(input.body.destinationId));
      if (options.rejectMove) throw new Error('rejected');
      folders[id] = { archive: 'archive-id', inbox: 'inbox-id', deleteditems: 'trash-id' }[input.body.destinationId];
      if (options.disconnectOnMove) session = { accountKey: 'account-B', session: 2 };
      if (options.loseResponse) throw new Error('network');
      return {};
    }
    if (options.verificationFails && ['archive-id', 'trash-id'].includes(folders[id])) throw new Error('network');
    return { id, parentFolderId: folders[id], from: { emailAddress: { address: addresses[id] } } };
  };
  const sender = { name: '<script>not HTML</script>', address: 'newsletter@example.com', messageIds: ['one', 'two'] };
  const plan = api.createPlan(sender, session.accountKey, session.session, options.action || 'archive');
  const runtime = { request, currentSession: () => session };
  return { api, plan, runtime, calls, folders, addresses, setSession(value) { session = value; } };
}

test('only selected immutable IDs are archived; undo restores exactly those messages', async () => {
  const s = setup();
  const result = await s.api.run(s.plan, s.runtime);
  assert.ok(result.entries.every((entry) => entry.status === 'archived'));
  assert.equal(s.folders.other, 'inbox-id');
  await s.api.undo(result, s.runtime);
  assert.ok(result.entries.every((entry) => entry.restored));
  assert.equal(s.calls.filter((call) => call.method === 'POST').length, 4);
  await s.api.undo(result, s.runtime);
  assert.equal(s.calls.filter((call) => call.method === 'POST').length, 4);
  for (const call of s.calls) assert.ok(!/body|preview|subject|attachments/i.test(call.params?.$select || ''));
});

test('plan is immutable, deduplicated and not executable twice (including concurrently)', async () => {
  const s = setup();
  assert.ok(Object.isFrozen(s.plan) && Object.isFrozen(s.plan.ids));
  const running = s.api.run(s.plan, s.runtime);
  await assert.rejects(s.api.run(s.plan, s.runtime), /PLAN-EXPIRED/);
  await running;
  assert.equal(s.calls.filter((call) => call.method === 'POST').length, 2);
});

test('account/session change rejects plan before all requests', async () => {
  const s = setup();
  s.setSession({ accountKey: 'account-B', session: 1 });
  await assert.rejects(s.api.run(s.plan, s.runtime), /SESSION-CHANGED/);
  assert.equal(s.calls.length, 0);
});

test('changed sender or folder is not touched', async () => {
  const s = setup();
  s.addresses.one = 'someone-else@example.com';
  s.folders.two = 'deleted-id';
  const result = await s.api.run(s.plan, s.runtime);
  assert.ok(result.entries.every((entry) => entry.status === 'unchanged'));
  assert.equal(s.calls.filter((call) => call.method === 'POST').length, 0);
});

test('lost write response is verified, never retried', async () => {
  const s = setup({ loseResponse: true });
  const result = await s.api.run(s.plan, s.runtime);
  assert.ok(result.entries.every((entry) => entry.status === 'archived'));
  assert.equal(s.calls.filter((call) => call.method === 'POST').length, 2);
});

test('unverifiable writes remain unknown and cannot be undone speculatively', async () => {
  const s = setup({ verificationFails: true });
  const result = await s.api.run(s.plan, s.runtime);
  assert.ok(result.entries.every((entry) => entry.status === 'unknown'));
  await s.api.undo(result, s.runtime);
  assert.equal(s.calls.filter((call) => call.method === 'POST').length, 1);
  assert.equal(s.folders.two, 'inbox-id');
});

test('denied writes are reported unchanged, not successful', async () => {
  const s = setup({ rejectMove: true });
  const result = await s.api.run(s.plan, s.runtime);
  assert.ok(result.entries.every((entry) => entry.status === 'unchanged'));
});

test('disconnect during execution prevents remaining messages from being touched', async () => {
  const s = setup({ disconnectOnMove: true });
  const result = await s.api.run(s.plan, s.runtime);
  assert.equal(result.entries[0].status, 'unknown');
  assert.equal(s.calls.filter((call) => call.method === 'POST').length, 1);
  assert.equal(s.folders.two, 'inbox-id');
  await assert.rejects(s.api.undo(result, s.runtime), /SESSION-CHANGED/);
});

test('undo does not move messages subsequently moved by the user', async () => {
  const s = setup();
  const result = await s.api.run(s.plan, s.runtime);
  s.folders.one = 'some-user-folder';
  await s.api.undo(result, s.runtime);
  assert.equal(s.folders.one, 'some-user-folder');
  assert.equal(result.entries[0].restored, false);
  assert.equal(result.entries[1].restored, true);
});

test('invalid IDs, sender and empty/oversized plans are refused', () => {
  const s = setup();
  for (const messageIds of [[], [null], Array.from({ length: 101 }, (_, i) => String(i))]) {
    assert.throws(() => s.api.createPlan({ address: 'a@example.com', messageIds }, 'account-A', 1));
  }
  assert.throws(() => s.api.createPlan({ address: '<bad>', messageIds: ['one'] }, 'account-A', 1));
});

test('trash moves only the selected emails to Deleted Items and restores them without any DELETE', async () => {
  const s = setup({ action: 'trash' });
  assert.equal(s.plan.destination, 'deleteditems');
  const result = await s.api.run(s.plan, s.runtime);
  assert.ok(result.entries.every((entry) => entry.status === 'trashed'));
  assert.equal(s.folders.other, 'inbox-id');
  await s.api.undo(result, s.runtime);
  assert.ok(result.entries.every((entry) => entry.restored));
  assert.deepEqual(s.calls.filter((call) => call.method === 'POST').map((call) => call.body.destinationId), ['deleteditems', 'deleteditems', 'inbox', 'inbox']);
  assert.ok(s.calls.every((call) => call.method !== 'DELETE' && !call.path.includes('permanentDelete')));
});

test('trash outcome unknown stops the batch and does not retry or speculate about undo', async () => {
  const s = setup({ action: 'trash', verificationFails: true });
  const result = await s.api.run(s.plan, s.runtime);
  assert.equal(result.entries[0].status, 'unknown');
  await s.api.undo(result, s.runtime);
  assert.equal(s.calls.filter((call) => call.method === 'POST').length, 1);
  assert.equal(s.folders.two, 'inbox-id');
});

test('restore from trash skips messages moved elsewhere after this test', async () => {
  const s = setup({ action: 'trash' });
  const result = await s.api.run(s.plan, s.runtime);
  s.folders.one = 'another-folder';
  await s.api.undo(result, s.runtime);
  assert.equal(s.folders.one, 'another-folder');
  assert.equal(result.entries[0].restored, false);
  assert.equal(result.entries[1].restored, true);
});

test('unsupported actions and permanent deletion destinations cannot become a plan', () => {
  const s = setup();
  for (const action of ['delete', 'permanentDelete', 'spam', 'recoverableitemsdeletions', 'inbox']) {
    assert.throws(() => s.api.createPlan({ address: 'a@example.com', messageIds: ['one'] }, 'A', 1, action), /INVALID-PLAN/);
  }
});
