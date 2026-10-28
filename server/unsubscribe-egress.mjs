import { boundedJson } from './unsubscribe-metadata.mjs';

// Fixed platform-controlled destinations only. No user-supplied custom domains
// or wildcard delegation beyond Beehiiv's single publication label. These DNS
// zones are part of the trust boundary and each addition requires tests.
const PLATFORM_DESTINATIONS = Object.freeze([
  (hostname) => /^[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.beehiiv\.com$/.test(hostname),
  (hostname) => hostname === 'eu2-cdn.eng.bloomreach.com',
  (hostname) => hostname === 'unsubscribe.kit.com',
  (hostname) => hostname === 'ac.emlnk1.com'
]);

export function supportedUnsubscribeUrl(value) {
  try {
    if (typeof value !== 'string' || value.length > 8192 || /[\s\\\u0000-\u001f\u007f]/.test(value)) return null;
    const url = new URL(value);
    if (url.protocol !== 'https:' || url.username || url.password || url.port || url.hash) return null;
    if (!PLATFORM_DESTINATIONS.some((matches) => matches(url.hostname))) return null;
    return url;
  } catch (_) { return null; }
}

export function publicDnsAddress(address) {
  if (/^\d+\.\d+\.\d+\.\d+$/.test(address)) {
    const octets = address.split('.').map(Number);
    if (octets.some((n) => n > 255)) return false;
    const [a, b, c] = octets;
    return !(a === 0 || a === 10 || a === 127 || a >= 224 || (a === 100 && b >= 64 && b <= 127) || (a === 169 && b === 254) || (a === 172 && b >= 16 && b <= 31) || (a === 192 && (b === 168 || b === 0)) || (a === 198 && (b === 18 || b === 19 || b === 51)) || (a === 203 && b === 0 && c === 113));
  }
  if (!/^[0-9a-f:]+$/i.test(address) || (address.match(/::/g) || []).length > 1) return false;
  const pieces = address.split(':').filter(Boolean);
  if (pieces.some((part) => part.length > 4) || (address.includes('::') ? pieces.length >= 8 : pieces.length !== 8)) return false;
  const first = parseInt(pieces[0], 16);
  return first >= 0x2000 && first <= 0x3fff && first !== 0x2002 && first !== 0x3fff && !(first === 0x2001 && [0, 0xdb8].includes(parseInt(pieces[1] || '0', 16)));
}

export function createDestinationPolicy(fetchDns = fetch) {
  return async (value) => {
    const url = supportedUnsubscribeUrl(value);
    if (!url) return false;
    try {
      const addresses = [];
      for (const type of ['A', 'AAAA']) {
        const dns = new URL('https://cloudflare-dns.com/dns-query');
        dns.searchParams.set('name', url.hostname); dns.searchParams.set('type', type);
        const response = await fetchDns(dns.href, { headers: { Accept: 'application/dns-json' }, credentials: 'omit', redirect: 'error', signal: AbortSignal.timeout(5000) });
        if (!response.ok) { await response.body?.cancel(); return false; }
        const json = await boundedJson(response, 32768);
        if (json.Status !== 0) return false;
        for (const answer of json.Answer || []) if ([1, 28].includes(answer.type)) addresses.push(answer.data);
      }
      return addresses.length > 0 && addresses.every((address) => typeof address === 'string' && publicDnsAddress(address));
    } catch (_) { return false; }
  };
}
