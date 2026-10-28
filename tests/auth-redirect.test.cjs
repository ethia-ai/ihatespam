const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const vm = require("node:vm");

const dist = path.join(__dirname, "../dist");
const bridge = fs.readFileSync(path.join(dist, "vendor/msal-redirect-bridge-5.17.3.min.js"), "utf8");
const callback = fs.readFileSync(path.join(dist, "auth-redirect.js"), "utf8");

async function runRedirect(hash) {
  const observations = { broadcasts: [], replacedUrls: [], closed: false };
  const context = {
    document: { title: "", body: { textContent: "" } },
    location: { origin: "https://example.test", pathname: "/auth-redirect.html", search: "", hash },
    history: { replaceState(_state, _title, url) { observations.replacedUrls.push(url); } },
    close() { observations.closed = true; },
    BroadcastChannel: class {
      constructor(name) { this.name = name; }
      postMessage(message) { observations.broadcasts.push({ name: this.name, message }); }
      close() {}
    },
    URLSearchParams,
    TextDecoder,
    atob: (value) => Buffer.from(value, "base64").toString("binary")
  };
  context.window = context;
  context.self = context;
  vm.createContext(context);
  vm.runInContext(bridge, context);
  await vm.runInContext(callback, context);
  return { context, observations };
}

test("Microsoft popup response is cleared, broadcast and closed without storage or network access", async () => {
  const state = Buffer.from(JSON.stringify({ id: "test-popup-channel", meta: { interactionType: "popup" } })).toString("base64");
  const payload = `code=synthetic-test-code&state=${encodeURIComponent(state)}`;
  const { context, observations } = await runRedirect(`#${payload}`);
  assert.equal(context.document.body.textContent, "");
  assert.deepEqual(observations.replacedUrls, ["https://example.test/auth-redirect.html"]);
  assert.equal(observations.broadcasts.length, 1);
  assert.equal(observations.broadcasts[0].name, "test-popup-channel");
  assert.equal(observations.broadcasts[0].message.payload, payload);
  assert.equal(observations.closed, true);
});

test("invalid callback shows a generic message and does not expose the response", async () => {
  const { context, observations } = await runRedirect("#code=synthetic-private-value");
  assert.equal(observations.broadcasts.length, 0);
  assert.equal(context.document.body.textContent.includes("synthetic-private-value"), false);
  assert.match(context.document.body.textContent, /could not be completed/);
});

test("callback loads only the Microsoft bridge and its initializer", () => {
  const html = fs.readFileSync(path.join(dist, "auth-redirect.html"), "utf8");
  const scripts = [...html.matchAll(/<script src="([^"]+)"/g)].map((match) => match[1]);
  assert.deepEqual(scripts, ["./vendor/msal-redirect-bridge-5.17.3.min.js", "./auth-redirect.js"]);
  assert.match(html, /default-src 'none'; script-src 'self'/);
});
