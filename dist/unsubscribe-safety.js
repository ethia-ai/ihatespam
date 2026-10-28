// Validation is a navigation guard, not proof of a trustworthy sender or website.
(() => {
  function publicHttpsUrl(value) {
    if (typeof value !== "string" || value.length > 8192 || /[\u0000-\u0020\u007f\\]/.test(value)) return null;
    try {
      const url = new URL(value);
      const host = url.hostname.toLowerCase();
      if (url.protocol !== "https:" || url.username || url.password || url.port || url.hash) return null;
      if (!/^(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z]{2,63}$/.test(host)) return null;
      if (host.split('.').some((label) => label.length > 63) || host.length > 253) return null;
      if (/(?:^|\.)(?:localhost|local|internal|lan|home|test|invalid|example|onion)$/.test(host)) return null;
      return url;
    } catch (_) { return null; }
  }
  window.UnsubscribeSafety = Object.freeze({ publicHttpsUrl });
})();
