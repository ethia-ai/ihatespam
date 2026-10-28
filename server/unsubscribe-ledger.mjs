// Only opaque keyed hashes, counters and result codes. No OAuth tokens, URLs,
// headers, addresses or message bodies are persisted. D1 batch is transactional.
export function createUnsubscribeLedger(db, actorHash, now = Date.now) {
  if (!db?.prepare || !db?.batch || !/^[a-f0-9]{64}$/.test(actorHash)) throw new Error('ledger_unavailable');
  const statement = (sql, ...args) => db.prepare(sql).bind(...args);
  function budgets(kind) {
    const time = now();
    const day = Math.floor(time / 86400000), minute = Math.floor(time / 60000);
    const entries = kind === 'request'
      ? [[`request:global:day:${day}`, 200, (day + 1) * 86400000], [`request:${actorHash}:minute:${minute}`, 12, (minute + 1) * 60000]]
      : [[`send:global:day:${day}`, 100, (day + 1) * 86400000], [`send:${actorHash}:day:${day}`, 50, (day + 1) * 86400000]];
    return entries.flatMap(([key, limit, expires]) => [
      statement('INSERT OR IGNORE INTO unsubscribe_budgets (key, remaining, expires) VALUES (?, ?, ?)', key, limit, expires),
      statement('UPDATE unsubscribe_budgets SET remaining = remaining - 1 WHERE key = ?', key)
    ]);
  }
  return Object.freeze({
    durable: true,
    async allowRequest() {
      try {
        await db.batch([
          statement('DELETE FROM unsubscribe_budgets WHERE expires < ?', now()),
          statement('DELETE FROM unsubscribe_attempts WHERE expires < ?', now()),
          ...budgets('request')
        ]);
        return true;
      } catch (_) { return false; }
    },
    async consumePlan(id) {
      if (!/^[a-f0-9]{64}$/.test(id)) return false;
      try {
        await statement('INSERT INTO unsubscribe_attempts (key, status, reason, expires) VALUES (?, ?, ?, ?)', 'plan:' + id, 'consumed', 'plan_used', now() + 86400000).run();
        return true;
      } catch (_) { return false; }
    },
    async reserve(key) {
      if (!/^[a-f0-9]{64}$/.test(key)) return 'unavailable';
      try {
        await db.batch([
          ...budgets('send'),
          statement('INSERT INTO unsubscribe_attempts (key, status, reason, expires) VALUES (?, ?, ?, ?)', key, 'uncertain', 'outcome_unknown', now() + 7 * 86400000)
        ]);
        return 'reserved';
      } catch (_) {
        try { if (await statement('SELECT key FROM unsubscribe_attempts WHERE key = ?', key).first()) return 'duplicate'; } catch (_) { /* Fail closed. */ }
        return 'limit';
      }
    },
    async complete(key, outcome) {
      const result = await statement('UPDATE unsubscribe_attempts SET status = ?, reason = ? WHERE key = ?', outcome.status, outcome.reason, key).run();
      if (result.meta?.changes !== 1) throw new Error('ledger_unavailable');
    }
  });
}
