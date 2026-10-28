// Conservative RSA/SHA-256, relaxed-header verifier. Body validation is NOT
// performed here: a matching trusted Gmail receiver attestation is also required.
// Outlook receiver provenance is not yet established, so it fails closed.
import { boundedJson } from './unsubscribe-metadata.mjs';
const encoder = new TextEncoder();
const tags = (value) => {
  const out = Object.create(null);
  for (const part of value.split(';')) {
    if (!part.trim()) continue;
    const match = part.match(/^\s*([a-z][a-z0-9]*)\s*=\s*([\s\S]*?)\s*$/);
    if (!match || Object.hasOwn(out, match[1])) return null;
    out[match[1]] = match[2];
  }
  return out;
};
const decode = (value) => {
  const compact = value.replace(/[ \t\r\n]/g, '');
  if (!/^[A-Za-z0-9+/]+={0,2}$/.test(compact)) throw new Error('invalid_key');
  return Uint8Array.from(atob(compact), (c) => c.charCodeAt(0));
};
export const relaxedHeader = (name, value) => name.toLowerCase().trim() + ':' + value.replace(/\r?\n[ \t]+/g, ' ').replace(/[ \t]+/g, ' ').trim() + '\r\n';

export function canonicalSignedHeaders(headers, signature) {
  const t = tags(signature.value);
  if (!t?.h || t.h.length > 4096) return null;
  const selected = new Set();
  let data = '';
  for (const field of t.h.toLowerCase().split(':').map((part) => part.trim())) {
    if (!/^[a-z0-9-]+$/.test(field) || field === 'dkim-signature') return null;
    for (let i = headers.length - 1; i >= 0; i--) {
      if (headers[i].name === field && !selected.has(i)) {
        selected.add(i); data += relaxedHeader(field, headers[i].value); break;
      }
    }
  }
  const emptyB = signature.value.replace(/(^|;)([ \t]*b[ \t]*=)[^;]*/, '$1$2');
  // The DKIM-Signature field itself has no trailing CRLF in the signing input.
  return encoder.encode(data + relaxedHeader('dkim-signature', emptyB).slice(0, -2));
}

export function createSignatureVerifier({ lookupKey, subtle = crypto.subtle, now = Date.now }) {
  return async function verify(metadata) {
    try {
      if (metadata.provider !== 'gmail') return false;
      const headers = metadata.headers;
      for (const name of ['from', 'list-unsubscribe', 'list-unsubscribe-post']) {
        if (headers.filter((h) => h.name === name).length !== 1) return false;
      }
      const attestations = headers.filter((h) => h.name === 'authentication-results' && /^mx\.google\.com\s*;/i.test(h.value.trim()));
      if (attestations.length !== 1) return false;
      const signatures = headers.filter((h) => h.name === 'dkim-signature');
      if (!signatures.length || signatures.length > 5) return false;
      const attestation = attestations[0].value;
      for (const signature of signatures) {
        const t = tags(signature.value);
        if (!t || t.v !== '1' || t.a !== 'rsa-sha256' || !/^relaxed\/(?:relaxed|simple)$/.test(t.c || '') || !t.b || !t.bh || t.l !== undefined) continue;
        if (!/^[a-z0-9](?:[a-z0-9.-]{0,250}[a-z0-9])?$/i.test(t.d || '') || !/^[a-z0-9_-]{1,100}$/i.test(t.s || '')) continue;
        const domain = t.d.toLowerCase();
        const fromDomain = metadata.sender.split('@')[1];
        if (fromDomain !== domain && !fromDomain.endsWith('.' + domain)) continue;
        const signed = (t.h || '').toLowerCase().split(':').map((s) => s.trim());
        if (!['from', 'list-unsubscribe', 'list-unsubscribe-post'].every((name) => signed.includes(name))) continue;
        if (t.x && (!/^\d+$/.test(t.x) || Number(t.x) <= now() / 1000)) continue;
        if (t.t && (!/^\d+$/.test(t.t) || Number(t.t) > now() / 1000 + 300)) continue;
        if (decode(t.bh).length !== 32) continue;
        // Match the precise successful signature prefix, not a generic pass
        // somewhere else in a chain. Missing/truncated attestations fail closed.
        const passes = attestation.split(';').filter((part) => /^\s*dkim=pass\b/i.test(part));
        const matched = passes.some((part) => {
          const identity = part.match(/\bheader\.(?:i|d)=([^\s;]+)/i)?.[1]?.replace(/^@/, '').toLowerCase();
          const selector = part.match(/\bheader\.s=([^\s;]+)/i)?.[1];
          const prefix = part.match(/\bheader\.b="?([A-Za-z0-9+/=]{8,})"?/i)?.[1];
          return identity === domain && selector === t.s && prefix && t.b.replace(/\s/g, '').startsWith(prefix);
        });
        if (!matched) continue;
        const keyRecord = tags(await lookupKey(t.s + '._domainkey.' + domain));
        if (!keyRecord || (keyRecord.v && keyRecord.v !== 'DKIM1') || (keyRecord.k && keyRecord.k !== 'rsa') || !keyRecord.p || (keyRecord.t || '').split(':').includes('y')) continue;
        if (keyRecord.h && !keyRecord.h.split(':').includes('sha256')) continue;
        if (keyRecord.s && !keyRecord.s.split(':').some((s) => s === '*' || s === 'email')) continue;
        const key = await subtle.importKey('spki', decode(keyRecord.p), { name: 'RSASSA-PKCS1-v1_5', hash: 'SHA-256' }, false, ['verify']);
        if (key.algorithm.modulusLength < 2048) continue;
        const data = canonicalSignedHeaders(headers, signature);
        if (data && await subtle.verify('RSASSA-PKCS1-v1_5', key, decode(t.b), data)) return true;
      }
    } catch (_) { /* DNS, parser, expired/rotated key: manual, never guessed. */ }
    return false;
  };
}

export function createDkimLookup(fetchDns = fetch) {
  return async (name) => {
    if (!/^[a-z0-9_-]{1,100}\._domainkey\.[a-z0-9.-]{1,253}$/i.test(name)) throw new Error('dns_invalid');
    const url = new URL('https://cloudflare-dns.com/dns-query');
    url.searchParams.set('name', name); url.searchParams.set('type', 'TXT');
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 5000);
    try {
      const response = await fetchDns(url.href, { headers: { Accept: 'application/dns-json' }, redirect: 'error', credentials: 'omit', signal: controller.signal });
      if (!response.ok) { await response.body?.cancel(); throw new Error('dns_unavailable'); }
      const json = await boundedJson(response, 32768);
      if (json.Status !== 0 || !Array.isArray(json.Answer)) throw new Error('dns_unavailable');
      const values = json.Answer.filter((a) => a.type === 16 && a.name?.replace(/\.$/, '').toLowerCase() === name.toLowerCase()).map((a) => {
        if (typeof a.data !== 'string' || !/^(?:"[^"\\]*"\s*)+$/.test(a.data)) throw new Error('dns_invalid');
        return [...a.data.matchAll(/"([^"\\]*)"/g)].map((m) => m[1]).join('');
      });
      if (values.length !== 1) throw new Error('dns_ambiguous');
      return values[0];
    } finally { clearTimeout(timer); }
  };
}
