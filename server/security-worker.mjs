import { handleUnsubscribeApi } from './unsubscribe-api.mjs';
// Static assets plus authenticated, bounded unsubscribe routes. Never log data.
export function createStaticSite(assets) {
  const appCsp = "default-src 'self'; script-src 'self' https://accounts.google.com; style-src 'self'; img-src 'self' data:; font-src 'self'; connect-src 'self' https://accounts.google.com https://oauth2.googleapis.com https://gmail.googleapis.com https://login.microsoftonline.com https://graph.microsoft.com; frame-src https://accounts.google.com https://login.microsoftonline.com; object-src 'none'; base-uri 'none'; form-action 'self'; frame-ancestors 'none'; upgrade-insecure-requests";
  const callbackCsp = "default-src 'none'; script-src 'self'; base-uri 'none'; form-action 'none'; frame-ancestors 'self'";
  return {
    async fetch(request, env) {
      const path = new URL(request.url).pathname;
      if (path.startsWith('/api/unsubscribe/')) return handleUnsubscribeApi(request, env);
      const callback = path === '/auth-redirect.html' || path === '/auth-redirect';
      const headers = new Headers({
        'Content-Security-Policy': callback ? callbackCsp : appCsp,
        'Referrer-Policy': 'no-referrer',
        'X-Content-Type-Options': 'nosniff',
        'X-Frame-Options': callback ? 'SAMEORIGIN' : 'DENY',
        'Permissions-Policy': 'camera=(), microphone=(), geolocation=(), payment=(), usb=()',
        'Cross-Origin-Opener-Policy': 'same-origin-allow-popups',
        'Cross-Origin-Resource-Policy': 'same-origin',
        'Cache-Control': 'no-store'
      });
      if (!['GET', 'HEAD'].includes(request.method)) {
        headers.set('Allow', 'GET, HEAD');
        return new Response(null, { status: 405, headers });
      }
      const aliases = { '/': '/index.html', '/auth-redirect': '/auth-redirect.html', '/privacy': '/privacy.html', '/terms': '/terms.html' };
      const key = Object.hasOwn(aliases, path) ? aliases[path] : path;
      if (!Object.hasOwn(assets, key)) return new Response(null, { status: 404, headers });
      const asset = assets[key];
      headers.set('Content-Type', asset.type);
      headers.set('ETag', asset.etag);
      const html = asset.type.startsWith('text/html');
      if (!html) headers.set('Cache-Control', 'public, max-age=0, must-revalidate');
      if (!html && request.headers.get('If-None-Match') === asset.etag) {
        return new Response(null, { status: 304, headers });
      }
      // Binary-safe delivery; assets are generated only from the explicit public allowlist.
      const body = request.method === 'HEAD' ? null : Uint8Array.from(atob(asset.data), (char) => char.charCodeAt(0));
      return new Response(body, { status: 200, headers });
    }
  };
}
