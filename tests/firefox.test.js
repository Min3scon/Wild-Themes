// The Firefox package: same code base as Chrome with a different manifest. Checks the manifest the
// build writes, runs Mozilla's own add-on linter over it, and runs the background script the way
// Firefox does (listed scripts, in order, no importScripts, no service worker).
// This does NOT run Firefox itself: that has to be done in a real Firefox (see extension/README.md).
const test = require("node:test"), assert = require("node:assert/strict");
const fs = require("fs"), path = require("path"), os = require("os"), vm = require("vm"), { spawnSync } = require("child_process");
const { build, GECKO_ID } = require("../scripts/build-extension.js");

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "wt-ff-"));
const chrome = build("chrome", path.join(tmp, "chrome")), firefox = build("firefox", path.join(tmp, "firefox"));
const mf = (d) => JSON.parse(fs.readFileSync(path.join(d, "manifest.json"), "utf8"));

test("Firefox manifest: event page, add-on id, data declaration; Chrome manifest untouched", () => {
  const f = mf(firefox), c = mf(chrome), src = JSON.parse(fs.readFileSync(path.join(__dirname, "..", "extension", "manifest.json"), "utf8"));
  assert.deepEqual(c, src, "Chrome package uses the manifest as it is");
  assert.ok(c.background.service_worker && !c.browser_specific_settings, "no Firefox keys in the Chrome package");
  assert.equal(f.manifest_version, 3);
  assert.deepEqual(f.background, { scripts: ["lib/themes.js", "lib/catalog.js", "background.js"] }, "helpers load before background.js; no service_worker");
  assert.equal(f.browser_specific_settings.gecko.id, GECKO_ID);
  assert.deepEqual(f.browser_specific_settings.gecko.data_collection_permissions, { required: ["none"] });
  for (const k of ["permissions", "host_permissions", "content_scripts", "action", "options_ui", "icons", "version"]) assert.deepEqual(f[k], c[k], k + " identical");
  // every file the manifest points at exists in the package
  const files = [f.action.default_popup, f.options_ui.page, ...f.background.scripts, ...Object.values(f.icons), ...f.content_scripts.flatMap((s) => [...(s.js || []), ...(s.css || [])])];
  for (const p of files) assert.ok(fs.existsSync(path.join(firefox, p)), p + " exists");
});

test("Google Docs, Sheets and Slides are in the content-script matches and host permissions, in both packages", () => {
  for (const d of [chrome, firefox]) {
    const m = mf(d);
    for (const g of ["document", "spreadsheets", "presentation"]) {
      assert.ok(m.content_scripts[0].matches.includes(`https://docs.google.com/${g}/*`), g + " match");
      assert.ok(m.host_permissions.includes(`https://docs.google.com/${g}/*`), g + " host");
    }
    assert.ok(!m.content_scripts[0].matches.some((x) => /docs\.google\.com\/\*|forms|drive/.test(x)), "nothing broader than the three editors");
  }
});

test("Mozilla's add-on linter: no errors, and only the known warning", { timeout: 300000 }, () => {
  const bin = path.join(__dirname, "node_modules", ".bin", "addons-linter");
  if (!fs.existsSync(bin)) return;   // dev dependency not installed
  const r = spawnSync(bin, [firefox, "--output", "json"], { encoding: "utf8", timeout: 280000 });
  const out = (r.stdout || "") + (r.stderr || ""), j = JSON.parse(out.slice(out.indexOf("{")));
  assert.deepEqual(j.errors.map((e) => `${e.code} ${e.file} ${e.message}`), [], "linter errors");
  // jszip (unmodified, from its own release) builds a global-object lookup with Function(); that is the one allowed warning
  const unexpected = j.warnings.filter((w) => !(w.code === "DANGEROUS_EVAL" && /vendor\/jszip/.test(w.file || "")));
  assert.deepEqual(unexpected.map((w) => `${w.code} ${w.file}:${w.line} ${w.message}`), [], "linter warnings");
});

test("background script runs as a Firefox event page: scripts in order, no importScripts, message handling works", async () => {
  const f = mf(firefox), listeners = { message: [], alarm: [], installed: [], startup: [], changed: [] }, calls = [];
  const ctx = {
    self: null, console, fetch: async () => ({ ok: false, status: 404 }), setTimeout, clearTimeout, URL, AbortController, TextEncoder, JSON,
    chrome: {
      runtime: { onMessage: { addListener: (fn) => listeners.message.push(fn) }, onInstalled: { addListener: (fn) => listeners.installed.push(fn) },
        onStartup: { addListener: (fn) => listeners.startup.push(fn) }, getURL: (p) => "moz-extension://x/" + p },
      scripting: { insertCSS: (o) => calls.push(["insertCSS", o]) },
      storage: { local: { get: (k, cb) => cb({}), set: async () => {} }, sync: { get: (d, cb) => cb(d) }, onChanged: { addListener: (fn) => listeners.changed.push(fn) } },
      alarms: { get: (n, cb) => cb(undefined), create: (n, o) => calls.push(["alarm", n, o]), onAlarm: { addListener: (fn) => listeners.alarm.push(fn) } },
      tabs: { create: (o) => calls.push(["tab", o]) },
    },
  };
  ctx.self = ctx; vm.createContext(ctx);
  assert.equal(typeof ctx.importScripts, "undefined", "Firefox event pages have no importScripts");
  for (const s of f.background.scripts) vm.runInContext(fs.readFileSync(path.join(firefox, s), "utf8"), ctx, { filename: s });
  assert.equal(listeners.message.length, 1);
  listeners.message[0]({ css: "synthwave" }, { tab: { id: 7 }, frameId: 3 }, () => {});
  assert.deepEqual(JSON.parse(JSON.stringify(calls.find((c) => c[0] === "insertCSS")[1])), { target: { tabId: 7, frameIds: [3] }, files: ["scenes/synthwave.css"] });
  assert.ok(calls.some((c) => c[0] === "alarm" && c[1] === "catalog" && c[2].periodInMinutes === 360), "catalogue refresh scheduled once");
  listeners.message[0]({ css: "../../evil" }, { tab: { id: 7 }, frameId: 0 }, () => {});
  assert.equal(calls.filter((c) => c[0] === "insertCSS").length, 1, "odd scene names are ignored");
});
