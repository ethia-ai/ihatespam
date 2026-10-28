const cryptoEncoder = new TextEncoder();
const encodePlanBytes = (bytes) => {
  let binary = '';
  for (let offset = 0; offset < bytes.length; offset += 8192) binary += String.fromCharCode(...bytes.subarray(offset, offset + 8192));
  return btoa(binary).replaceAll('+', '-').replaceAll('/', '_').replace(/=+$/, '');
};
const decodePlanBytes = (value) => Uint8Array.from(atob(value.replaceAll('-', '+').replaceAll('_', '/')), (c) => c.charCodeAt(0));

export async function createPlanCrypto(secret, subtle = crypto.subtle) {
  if (typeof secret !== 'string' || !/^[a-f0-9]{64}$/.test(secret)) throw new Error('service_not_ready');
  const bytes = Uint8Array.from(secret.match(/../g), (s) => parseInt(s, 16));
  const root = await subtle.importKey('raw', bytes, 'HKDF', false, ['deriveKey']);
  const derive = (purpose, algorithm, usages) => subtle.deriveKey({ name: 'HKDF', hash: 'SHA-256', salt: cryptoEncoder.encode('i-hate-spam/unsubscribe/v1'), info: cryptoEncoder.encode(purpose) }, root, algorithm, false, usages);
  const hmac = await derive('ledger-digests', { name: 'HMAC', hash: 'SHA-256', length: 256 }, ['sign']);
  const aes = await derive('sealed-plans', { name: 'AES-GCM', length: 256 }, ['encrypt', 'decrypt']);
  const context = cryptoEncoder.encode('i-hate-spam/unsubscribe-plan/v1');
  return Object.freeze({
    async digest(value) { return [...new Uint8Array(await subtle.sign('HMAC', hmac, cryptoEncoder.encode(value)))].map((b) => b.toString(16).padStart(2, '0')).join(''); },
    async seal(plan) {
      const nonce = crypto.getRandomValues(new Uint8Array(12));
      const ciphertext = await subtle.encrypt({ name: 'AES-GCM', iv: nonce, additionalData: context }, aes, cryptoEncoder.encode(JSON.stringify(plan)));
      return encodePlanBytes(nonce) + '.' + encodePlanBytes(new Uint8Array(ciphertext));
    },
    async open(value) {
      if (typeof value !== 'string' || value.length > 800000 || !/^[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+$/.test(value)) throw new Error('plan_invalid');
      const [nonce, data] = value.split('.').map(decodePlanBytes);
      if (nonce.length !== 12) throw new Error('plan_invalid');
      return JSON.parse(new TextDecoder().decode(await subtle.decrypt({ name: 'AES-GCM', iv: nonce, additionalData: context }, aes, data)));
    }
  });
}
