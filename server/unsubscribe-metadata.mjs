// Server-only adapter. Never log arguments or provider responses. No message
// body, attachment, write scope, sender URL or arbitrary API path is accepted.
export class MetadataFailure extends Error {
  constructor(code) { super(code); this.name = 'MetadataFailure'; this.code = code; }
}

const MAX_BYTES = 128 * 1024;
const fail = (code) => { throw new MetadataFailure(code); };

// Authenticate the public API against Google, never a client-supplied identity
// or a hosting header. Request only the mailbox address, not email content.
export async function readGmailIdentity(token, fetchProvider = fetch) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 12000);
  try {
    const response = await fetchProvider('https://gmail.googleapis.com/gmail/v1/users/me/profile?fields=emailAddress', {
      method: 'GET', headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' },
      redirect: 'error', credentials: 'omit', cache: 'no-store', referrerPolicy: 'no-referrer', signal: controller.signal
    });
    if (!response.ok) {
      if (![401, 403, 429].includes(response.status)) console.error('gmail_identity_http_failure', response.status);
      await response.body?.cancel();
      fail(response.status === 401 || response.status === 403 ? 'reconnect_required' : response.status === 429 ? 'provider_limited' : 'identity_unavailable');
    }
    const profile = await boundedJson(response, 4096);
    if (typeof profile.emailAddress !== 'string' || profile.emailAddress.length > 320 || !/^[^\s<>@]+@[^\s<>@]+\.[^\s<>@]+$/.test(profile.emailAddress)) fail('identity_unavailable');
    return 'gmail:' + profile.emailAddress.toLowerCase();
  } catch (error) {
    if (error instanceof MetadataFailure) throw error;
    // Only fixed categories: never log the exception, URL, token or response.
    const category = ['redirect', 'credentials', 'referrerPolicy', 'cache', 'Illegal invocation', 'network', 'fetch', 'signal', 'cancel'].find((name) => String(error?.message).includes(name)) || 'unknown';
    console.error('gmail_identity_transport_failure', category);
    fail('identity_unavailable');
  } finally { clearTimeout(timer); token = null; }
}

export async function boundedJson(response, maximum = MAX_BYTES) {
  const declared = Number(response.headers.get('content-length'));
  if (Number.isFinite(declared) && declared > maximum) { await response.body?.cancel(); fail('metadata_too_large'); }
  if (!response.body) fail('metadata_unavailable');
  const reader = response.body.getReader();
  const chunks = [];
  let total = 0;
  let timedOut = false;
  const readTimer = setTimeout(() => { timedOut = true; void reader.cancel().catch(() => {}); }, 12000);
  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      total += value.byteLength;
      if (total > maximum) { await reader.cancel(); fail('metadata_too_large'); }
      chunks.push(value);
    }
    if (timedOut) fail('metadata_unavailable');
    const bytes = new Uint8Array(total);
    let offset = 0;
    for (const chunk of chunks) { bytes.set(chunk, offset); offset += chunk.byteLength; }
    return JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(bytes));
  } catch (error) {
    if (error instanceof MetadataFailure) throw error;
    fail('metadata_unavailable');
  } finally { clearTimeout(readTimer); reader.releaseLock(); }
}

function normaliseHeaders(raw) {
  if (!Array.isArray(raw) || raw.length > 500) fail('metadata_unavailable');
  return raw.map((entry) => {
    if (typeof entry?.name !== 'string' || !/^[A-Za-z0-9-]{1,100}$/.test(entry.name)) fail('headers_invalid');
    if (typeof entry.value !== 'string' || entry.value.length > 32768 || /\r(?!\n)|\n(?![ \t])/.test(entry.value)) fail('headers_invalid');
    return Object.freeze({ name: entry.name.toLowerCase(), value: entry.value.replace(/\r?\n[ \t]+/g, ' ') });
  });
}

export function createMetadataReader(fetchProvider = fetch) {
  return async function readMetadata({ provider, token, messageId, expectedSender }) {
    if (!['gmail', 'outlook'].includes(provider) || typeof token !== 'string' || token.length < 10 || token.length > 16384 || /\s/.test(token)) fail('invalid_request');
    if (typeof messageId !== 'string' || !/^[A-Za-z0-9_+=-]{1,2048}$/.test(messageId)) fail('invalid_request');
    if (typeof expectedSender !== 'string' || !/^[^\s<>@]+@[^\s<>@]+\.[^\s<>@]+$/.test(expectedSender) || expectedSender.length > 320) fail('invalid_request');
    const url = provider === 'gmail'
      ? new URL(`https://gmail.googleapis.com/gmail/v1/users/me/messages/${encodeURIComponent(messageId)}`)
      : new URL(`https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(messageId)}`);
    if (provider === 'gmail') {
      url.searchParams.set('format', 'metadata');
      url.searchParams.set('fields', 'id,payload/headers');
      // DKIM may sign any header. Retrieve headers only, including the signed
      // fields; never raw MIME, payload parts, body or attachments.
    } else url.searchParams.set('$select', 'id,from,internetMessageHeaders');
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 12000);
    try {
      const response = await fetchProvider(url.href, {
        method: 'GET', headers: { Authorization: `Bearer ${token}`, Accept: 'application/json', ...(provider === 'outlook' ? { Prefer: 'IdType="ImmutableId"' } : {}) },
        redirect: 'error', credentials: 'omit', cache: 'no-store', referrerPolicy: 'no-referrer', signal: controller.signal
      });
      if (!response.ok) {
        await response.body?.cancel();
        fail(response.status === 401 ? 'reconnect_required' : response.status === 403 ? 'metadata_permission' : response.status === 404 ? 'message_missing' : response.status === 429 ? 'provider_limited' : 'metadata_unavailable');
      }
      const data = await boundedJson(response);
      if (data.id !== messageId) fail('message_changed');
      const headers = normaliseHeaders(provider === 'gmail' ? data.payload?.headers : data.internetMessageHeaders);
      const from = headers.filter((entry) => entry.name === 'from');
      if (from.length !== 1) fail('headers_invalid');
      // Deliberately reject complex mailbox syntax rather than guessing.
      const match = from[0].value.match(/^(?:[^<>]*<([^<>\s]+)>|([^<>\s]+))$/);
      const address = (match?.[1] || match?.[2] || '').toLowerCase();
      if (address !== expectedSender.toLowerCase()) fail('sender_changed');
      if (provider === 'outlook' && String(data.from?.emailAddress?.address || '').toLowerCase() !== address) fail('sender_changed');
      return Object.freeze({ provider, messageId, sender: address, headers: Object.freeze(headers) });
    } catch (error) {
      if (error instanceof MetadataFailure) throw error;
      // Provider error objects can contain private URLs. Never propagate them.
      fail('metadata_unavailable');
    } finally { clearTimeout(timer); token = null; }
  };
}
