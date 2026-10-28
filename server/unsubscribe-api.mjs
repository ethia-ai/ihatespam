import { boundedJson, createMetadataReader, readGmailIdentity } from './unsubscribe-metadata.mjs';
import { createUnsubscribeEngine, MAX_UNSUBSCRIBE_BATCH } from './unsubscribe-engine.mjs';
import { createSignatureVerifier, createDkimLookup } from './unsubscribe-signature.mjs';
import { createDestinationPolicy, supportedUnsubscribeUrl } from './unsubscribe-egress.mjs';
import { createPlanCrypto } from './unsubscribe-crypto.mjs';
import { createUnsubscribeLedger } from './unsubscribe-ledger.mjs';

const UNSUBSCRIBE_ORIGIN = 'https://ihatespam.site';
// Workers reject browser-only RequestInit fields such as credentials; older
// versions also reject cache. There is no ambient cookie jar or page referrer.
// Use native no-cache controls and never retry an unsubscribe POST.
export function workerPrivateFetch(fetchNetwork) {
  return async (url, options = {}) => {
    const { cache, credentials, referrerPolicy, ...request } = options;
    if (credentials && credentials !== 'omit') throw new Error('request_rejected');
    const headers = new Headers(request.headers);
    headers.delete('Cookie');
    headers.delete('Referer');
    request.headers = headers;
    // Hosted Workers accept manual/follow, but reject redirect: 'error'.
    // Manual plus an explicit 3xx refusal preserves error-mode semantics.
    if (request.redirect === 'error') request.redirect = 'manual';
    if (cache === 'no-store') {
      headers.set('Cache-Control', 'no-store');
      request.cf = { ...request.cf, cacheTtlByStatus: { '100-599': -1 } };
    }
    const response = await fetchNetwork(url, request);
    if (options.redirect === 'error' && response.status >= 300 && response.status < 400) {
      await response.body?.cancel();
      throw new Error('redirect_rejected');
    }
    return response;
  };
}
export async function handleUnsubscribeApi(request, env, fetchNetwork = fetch) {
  const reply = (value, status = 200) => Response.json(value, { status, headers: { 'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff', 'X-Frame-Options': 'DENY', 'Referrer-Policy': 'no-referrer', 'Content-Security-Policy': "default-src 'none'; frame-ancestors 'none'" } });
  const path = new URL(request.url).pathname;
  if (!['/api/unsubscribe/prepare', '/api/unsubscribe/execute'].includes(path)) return reply({ reason: 'not_found' }, 404);
  if (request.method !== 'POST') return reply({ reason: 'method_not_allowed' }, 405);
  if (request.headers.get('origin') !== UNSUBSCRIBE_ORIGIN || !['same-origin', null].includes(request.headers.get('sec-fetch-site')) || request.headers.get('content-type')?.split(';')[0] !== 'application/json') return reply({ reason: 'request_rejected' }, 403);
  if (env?.UNSUBSCRIBE_ENABLED !== 'true' || !env.DB || !env.UNSUBSCRIBE_SECRET) return reply({ reason: 'service_not_ready' }, 503);
  fetchNetwork = workerPrivateFetch(fetchNetwork);
  try {
    const body = await boundedJson(request, 832000);
    if (typeof body.token !== 'string' || body.token.length < 10 || body.token.length > 16384 || /\s/.test(body.token)) return reply({ reason: 'reconnect_required' }, 401);
    if (path.endsWith('/prepare') && (body.provider !== 'gmail' || !Array.isArray(body.items) || body.items.length < 1 || body.items.length > MAX_UNSUBSCRIBE_BATCH || body.items.some((item) => !item || typeof item.messageId !== 'string' || typeof item.sender !== 'string'))) return reply({ reason: 'invalid_request' }, 400);
    if (path.endsWith('/execute') && (body.confirmed !== true || typeof body.plan !== 'string')) return reply({ reason: 'plan_invalid' }, 400);
    const cipher = await createPlanCrypto(env.UNSUBSCRIBE_SECRET);
    // Stable per-mailbox identity survives token refresh without persisting the
    // address. Both prepare and execute verify the presented token with Google.
    const actor = await cipher.digest('actor:' + await readGmailIdentity(body.token, fetchNetwork));
    const ledger = createUnsubscribeLedger(env.DB, actor);
    if (!await ledger.allowRequest()) return reply({ reason: 'limit_reached' }, 429);
    const destinationPolicy = createDestinationPolicy(fetchNetwork);
    const engine = createUnsubscribeEngine({
      readMetadata: createMetadataReader(fetchNetwork),
      verifySignature: createSignatureVerifier({ lookupKey: createDkimLookup(fetchNetwork) }),
      approveDestination: destinationPolicy, ledger, digest: cipher.digest, restoreAuthenticatedPlan: cipher.open,
      sendPost: async (url, options) => {
        // The fixed, platform-controlled allowlist is enforced again immediately
        // before fetch. No arbitrary URL, client headers or redirects are used.
        if (!supportedUnsubscribeUrl(url) || !await destinationPolicy(url)) throw new Error('destination_unsupported');
        return fetchNetwork(url, options);
      }
    });
    if (path.endsWith('/prepare')) {
      const plan = await engine.prepare({ actor, provider: body.provider, token: body.token, items: body.items });
      body.token = null;
      if (!plan.rows) return reply({ reason: plan.reason }, 400);
      return reply({ plan: await cipher.seal(plan), expires: plan.expires, rows: plan.rows.map(({ messageId, sender, status, reason, host }) => ({ messageId, sender, status, reason, host })) });
    }
    const outcome = await engine.execute({ sealedPlan: body.plan, actor, token: body.token, confirmed: true });
    body.token = null;
    return reply(outcome, outcome.status === 'processed' ? 200 : 409);
  } catch (error) {
    if (error?.code === 'reconnect_required') return reply({ reason: 'reconnect_required' }, 401);
    if (error?.code === 'provider_limited') return reply({ reason: 'provider_limited' }, 429);
    if (error?.code === 'identity_unavailable') return reply({ reason: 'identity_unavailable' }, 503);
    if (error?.message === 'service_not_ready') return reply({ reason: 'service_not_ready' }, 503);
    return reply({ reason: 'service_unavailable' }, 503);
  }
}
