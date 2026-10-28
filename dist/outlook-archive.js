/* Controlled Outlook moves: one sender, Inbox → Archive or Deleted Items, and undo.
 * No permanent deletion, reporting, rules, unsubscribe, retries or persistent state.
 */
(() => {
  const issuedPlans = new WeakSet();
  const consumedPlans = new WeakSet();
  const knownResults = new WeakSet();
  const busyResults = new WeakSet();

  function createPlan(sender, accountKey, session, action = "archive") {
    const ids = [...new Set(sender.messageIds || [])];
    if (!["archive", "trash"].includes(action) || !accountKey || !ids.length || ids.length > 100 ||
        ids.some((id) => typeof id !== "string" || !id || id.length > 2048) ||
        !/^[^\s<>@]+@[^\s<>@]+\.[^\s<>@]+$/.test(sender.address || "")) {
      throw new Error("INVALID-PLAN");
    }
    const plan = Object.freeze({
      accountKey, session, action, destination: action === "trash" ? "deleteditems" : "archive",
      address: sender.address.toLowerCase(), name: sender.name,
      ids: Object.freeze(ids), createdAt: Date.now()
    });
    issuedPlans.add(plan);
    return plan;
  }

  function assertSession(plan, current) {
    if (current.accountKey !== plan.accountKey || current.session !== plan.session) {
      throw new Error("SESSION-CHANGED");
    }
  }

  const messagePath = (id) => `me/messages/${encodeURIComponent(id)}`;
  const senderAddress = (message) => String((message.from?.emailAddress || message.sender?.emailAddress)?.address || "").trim().toLowerCase();

  async function run(plan, { request, currentSession }) {
    if (!issuedPlans.has(plan)) throw new Error("INVALID-PLAN");
    assertSession(plan, currentSession());
    if (consumedPlans.has(plan) || Date.now() - plan.createdAt > 10 * 60 * 1000) throw new Error("PLAN-EXPIRED");
    // Consume before any await: a double click can never run the plan twice.
    consumedPlans.add(plan);
    const checkedRequest = (...args) => { assertSession(plan, currentSession()); return request(...args); };
    const inbox = await checkedRequest("me/mailFolders/inbox", { params: { "$select": "id" } });
    const destination = await checkedRequest(`me/mailFolders/${plan.destination}`, { params: { "$select": "id" } });
    if (!inbox.id || !destination.id || inbox.id === destination.id) throw new Error("FOLDER-UNAVAILABLE");
    const result = { plan, inboxId: inbox.id, destinationFolderId: destination.id, entries: [] };
    knownResults.add(result);

    for (const id of plan.ids) {
      const entry = { id, status: "unchanged", restored: false };
      result.entries.push(entry);
      try {
        const before = await checkedRequest(messagePath(id), { params: { "$select": "id,from,sender,parentFolderId" } });
        if (before.id !== id || before.parentFolderId !== inbox.id || senderAddress(before) !== plan.address) {
          entry.reason = "MESSAGE-CHANGED";
          continue;
        }
        assertSession(plan, currentSession());
        entry.status = "unknown";
        try {
          await checkedRequest(`${messagePath(id)}/move`, {
            method: "POST", body: { destinationId: plan.destination }, discardBody: true,
            params: { "$select": "id,parentFolderId" }
          });
        } catch (_) { /* Never repeat a write: check the actual folder instead. */ }
        const after = await checkedRequest(messagePath(id), { params: { "$select": "id,parentFolderId" } });
        if (after.id === id && after.parentFolderId === destination.id) entry.status = plan.action === "trash" ? "trashed" : "archived";
        else if (after.id === id && after.parentFolderId === inbox.id) entry.status = "unchanged";
      } catch (_) {
        // Preserve unknown for a write with no verifiable outcome; don't claim success.
        entry.reason = "CHECK-FAILED";
      }
      // Stop promptly when disconnected; no remaining writes use the old account.
      try { assertSession(plan, currentSession()); } catch (_) { break; }
      if (entry.status === "unknown" || entry.reason === "CHECK-FAILED") break;
    }
    return result;
  }

  async function undo(result, { request, currentSession }) {
    if (!knownResults.has(result) || busyResults.has(result)) throw new Error("INVALID-UNDO");
    assertSession(result.plan, currentSession());
    busyResults.add(result);
    const checkedRequest = (...args) => { assertSession(result.plan, currentSession()); return request(...args); };
    try {
      for (const entry of result.entries.filter((item) => ["archived", "trashed"].includes(item.status) && !item.restored && !item.undoAttempted)) {
        try {
          const before = await checkedRequest(messagePath(entry.id), { params: { "$select": "id,from,sender,parentFolderId" } });
          if (before.id !== entry.id || before.parentFolderId !== result.destinationFolderId || senderAddress(before) !== result.plan.address) {
            entry.undoReason = "MESSAGE-CHANGED";
            continue;
          }
          assertSession(result.plan, currentSession());
          entry.undoAttempted = true;
          try {
            await checkedRequest(`${messagePath(entry.id)}/move`, {
              method: "POST", body: { destinationId: "inbox" }, discardBody: true,
              params: { "$select": "id,parentFolderId" }
            });
          } catch (_) { /* Verify once, without retrying the move. */ }
          const after = await checkedRequest(messagePath(entry.id), { params: { "$select": "id,parentFolderId" } });
          entry.restored = after.id === entry.id && after.parentFolderId === result.inboxId;
          if (!entry.restored) entry.undoReason = "NOT-CONFIRMED";
        } catch (_) { entry.undoReason = "CHECK-FAILED"; }
        try { assertSession(result.plan, currentSession()); } catch (_) { break; }
      }
      return result;
    } finally { busyResults.delete(result); }
  }

  window.OutlookArchive = Object.freeze({ createPlan, run, undo });
})();
