// Private HTTP adapter supplies verification, restricted egress and durable
// replay/quota controls. This engine has no default external fetch path.
export const MAX_UNSUBSCRIBE_BATCH = 50;
const headerValues = (metadata, name) => metadata.headers.filter((entry) => entry.name === name).map((entry) => entry.value);
const result = (status, reason) => Object.freeze({ status, reason });

export function inspectOneClick(metadata) {
  const links = headerValues(metadata, 'list-unsubscribe');
  const post = headerValues(metadata, 'list-unsubscribe-post');
  if (!links.length) return result('manual', 'link_missing');
  if (links.length !== 1 || post.length > 1) return result('manual', 'headers_ambiguous');
  if (post.length !== 1 || post[0].trim().toLowerCase() !== 'list-unsubscribe=one-click') return result('manual', 'automatic_unavailable');
  const entries = links[0].split(',').map((entry) => entry.trim());
  if (!entries.every((entry) => /^<[^<>]+>$/.test(entry))) return result('manual', 'headers_invalid');
  const https = entries.map((entry) => entry.slice(1, -1)).filter((entry) => /^https:/i.test(entry));
  if (https.length !== 1) return result('manual', 'destination_unsupported');
  let url;
  try {
    if (/[\s\\\u0000-\u001f\u007f]/.test(https[0]) || https[0].length > 8192) throw new Error();
    url = new URL(https[0]);
    if (url.protocol !== 'https:' || url.username || url.password || url.port || url.hash || !/^(?:[a-z0-9-]+\.)+[a-z]{2,63}$/.test(url.hostname) || /(?:^|\.)(?:localhost|local|internal|lan|test|invalid|example|onion)$/.test(url.hostname)) throw new Error();
  } catch (_) { return result('manual', 'destination_unsupported'); }
  // Syntax inspection is NOT proof of a valid DKIM signature or safe egress.
  return Object.freeze({ status: 'candidate', url: url.href, host: url.hostname });
}

export function createUnsubscribeEngine({ readMetadata, verifySignature, approveDestination, ledger, sendPost, digest, restoreAuthenticatedPlan, now = Date.now }) {
  const available = [readMetadata, verifySignature, approveDestination, sendPost, digest].every((fn) => typeof fn === 'function')
    && ledger?.durable === true && typeof ledger.reserve === 'function' && typeof ledger.complete === 'function';
  const issued = new WeakSet();
  const consumed = new WeakSet();

  async function prepare({ actor, provider, token, items }) {
    if (!available) return result('unavailable', 'service_not_ready');
    if (typeof actor !== 'string' || !actor || !['gmail', 'outlook'].includes(provider) || !Array.isArray(items) || !items.length || items.length > MAX_UNSUBSCRIBE_BATCH) return result('unavailable', 'invalid_request');
    if (new Set(items.map((item) => item.messageId)).size !== items.length) return result('unavailable', 'invalid_request');
    const rows = [];
    for (const item of items) {
      try {
        const metadata = await readMetadata({ provider, token, messageId: item.messageId, expectedSender: item.sender });
        const candidate = inspectOneClick(metadata);
        if (candidate.status !== 'candidate') { rows.push(Object.freeze({ messageId: item.messageId, sender: metadata.sender, ...candidate })); continue; }
        // Adapter must validate receiver provenance, the specific DKIM signature
        // and coverage of both unsubscribe headers, not trust a client flag.
        if (await verifySignature(metadata) !== true) { rows.push(Object.freeze({ messageId: item.messageId, sender: metadata.sender, ...result('manual', 'signature_unverified') })); continue; }
        // Must enforce safe DNS/connection egress, not just URL syntax.
        if (await approveDestination(candidate.url) !== true) { rows.push(Object.freeze({ messageId: item.messageId, sender: metadata.sender, ...result('manual', 'destination_unsupported') })); continue; }
        rows.push(Object.freeze({ messageId: item.messageId, sender: metadata.sender, status: 'ready', url: candidate.url, host: candidate.host, fingerprint: await digest(JSON.stringify(metadata)) }));
      } catch (error) {
        const codes = ['reconnect_required', 'metadata_permission', 'message_missing', 'sender_changed', 'message_changed', 'provider_limited', 'headers_invalid', 'metadata_too_large'];
        rows.push(Object.freeze({ messageId: item.messageId, sender: item.sender, ...result('manual', codes.includes(error?.code) ? error.code : 'metadata_unavailable') }));
      }
    }
    const plan = Object.freeze({ id: crypto.randomUUID(), actor, provider, credential: await digest(token), expires: now() + 300000, rows: Object.freeze(rows) });
    issued.add(plan);
    return plan;
  }

  async function execute({ plan, sealedPlan, actor, token, confirmed = false }) {
    if (sealedPlan && typeof restoreAuthenticatedPlan === 'function') {
      try {
        const restored = await restoreAuthenticatedPlan(sealedPlan);
        if (!restored || typeof restored.id !== 'string' || !Array.isArray(restored.rows) || restored.rows.length > MAX_UNSUBSCRIBE_BATCH || !Number.isFinite(restored.expires)) return result('unavailable', 'plan_invalid');
        plan = Object.freeze({ ...restored, rows: Object.freeze(restored.rows.map((row) => Object.freeze(row))) });
        issued.add(plan);
      } catch (_) { return result('unavailable', 'plan_invalid'); }
    }
    if (!available || !issued.has(plan) || consumed.has(plan) || !confirmed || actor !== plan.actor || now() >= plan.expires || await digest(token) !== plan.credential) return result('unavailable', 'plan_invalid');
    // Single-use before any await; concurrent invocations cannot reuse the plan.
    if (consumed.has(plan)) return result('unavailable', 'plan_invalid');
    consumed.add(plan);
    if (typeof ledger.consumePlan === 'function' && !await ledger.consumePlan(await digest('plan:' + plan.id))) return result('unavailable', 'plan_invalid');
    const outcomes = [];
    let stop = false;
    for (const row of plan.rows) {
      if (row.status !== 'ready') { outcomes.push(result(row.status, row.reason)); continue; }
      if (stop) { outcomes.push(result('not_sent', 'batch_stopped')); continue; }
      if (now() >= plan.expires) { outcomes.push(result('not_sent', 'plan_expired')); stop = true; continue; }
      let key;
      try {
        const metadata = await readMetadata({ provider: plan.provider, token, messageId: row.messageId, expectedSender: row.sender });
        if (await digest(JSON.stringify(metadata)) !== row.fingerprint || await verifySignature(metadata) !== true || await approveDestination(row.url) !== true) { outcomes.push(result('not_sent', 'message_changed')); continue; }
        // Key excludes tokens, URLs and message identifiers. Production digest
        // must be server-keyed HMAC; ledger owns atomic quota + replay decisions.
        key = await digest(`${actor}\n${plan.provider}\n${row.url}`);
        const reservation = await ledger.reserve(key);
        if (reservation !== 'reserved') { outcomes.push(result('not_sent', reservation === 'duplicate' ? 'already_attempted' : 'limit_reached')); continue; }
      } catch (_) { outcomes.push(result('not_sent', 'validation_unavailable')); stop = true; continue; }
      let outcome;
      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), 12000);
      try {
        const response = await sendPost(row.url, {
          method: 'POST', body: 'List-Unsubscribe=One-Click', headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
          credentials: 'omit', redirect: 'manual', cache: 'no-store', referrerPolicy: 'no-referrer', signal: controller.signal
        });
        const status = response.status;
        await response.body?.cancel();
        outcome = status >= 200 && status < 300 ? result('requested', 'request_accepted')
          : status >= 300 && status < 400 ? result('uncertain', 'redirect_received')
          : status === 429 ? result('manual', 'sender_limited')
          : status >= 400 && status < 500 ? result('manual', 'request_rejected')
          : result('uncertain', 'outcome_unknown');
      } catch (_) { outcome = result('uncertain', 'outcome_unknown'); }
      finally { clearTimeout(timer); }
      try { await ledger.complete(key, outcome); }
      catch (_) { outcome = result('uncertain', 'outcome_unknown'); }
      outcomes.push(outcome);
      if (outcome.status === 'uncertain' || outcome.reason === 'sender_limited') stop = true;
    }
    return Object.freeze({ status: 'processed', outcomes: Object.freeze(outcomes) });
  }
  return Object.freeze({ prepare, execute });
}
