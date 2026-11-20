// The downloads as people use them: unzip into a folder (names with spaces and brackets, like Windows
// makes "Wild Themes (1)") and load it unpacked in Chromium; hand the Firefox zip to Mozilla's linter
// as about:debugging would load it. Checks the zips themselves: manifest at the top, no wrapper folder,
// names safe for Windows, nothing a loader trips over. Real Edge, Opera, Vivaldi, Brave and Firefox aren't run here.
const test = require("node:test"), assert = require("node:assert/strict");
const fs = require("fs"), path = require("path"), os = require("os"), { execFileSync, spawnSync } = require("child_process");
process.env.PW_EXPERIMENTAL_SERVICE_WORKER_NETWORK_EVENTS = "1";
let chromium; try { ({ chromium } = require("playwright")); } catch (_) { /* skipped below */ }

const ROOT = path.join(__dirname, "..");
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "wt-pkg-"));
execFileSync("node", [path.join(ROOT, "scripts", "build-extension.js"), "--zips", tmp], { stdio: "pipe" });
const list = (z) => execFileSync("unzip", ["-Z1", z], { encoding: "utf8", maxBuffer: 1 << 26 }).split("\n").filter(Boolean);
const unzip = (z, to) => { fs.mkdirSync(to, { recursive: true }); execFileSync("unzip", ["-q", z, "-d", to]); return to; };
const ZIPS = { chrome: path.join(tmp, "wild-themes-extension.zip"), firefox: path.join(tmp, "wild-themes-firefox.zip") };

// Chromium's manifest keys we use. Anything else would show as "Unrecognized manifest key" under chrome://extensions.
const CHROME_KEYS = ["manifest_version", "name", "short_name", "version", "description", "icons", "action", "options_ui", "background", "permissions", "content_scripts", "host_permissions"];

for (const [name, zip] of Object.entries(ZIPS)) {
  test(`${name} zip: manifest at the top, no wrapper folder, names every loader and Windows accept`, () => {
    const names = list(zip), files = names.filter((n) => !n.endsWith("/"));
    assert.ok(files.includes("manifest.json"), "manifest.json is at the top of the zip (an extra folder around it gives 'Manifest file is missing')");
    for (const n of names) {
      assert.ok(/^[\x21-\x7e]+$/.test(n.replace(/ /g, "_")), "plain characters only: " + n);
      assert.ok(!/\\|:|\*|\?|"|<|>|\||\/\/|(^|\/)\.\.(\/|$)|(^|\/)\.DS_Store$|^__MACOSX/.test(n), "name that breaks on Windows, Mac or a loader: " + n);
      assert.ok(!/^_/.test(n), "no top-level name starting with an underscore (reserved by Chrome): " + n);
      assert.ok(n.length < 150, "short enough for Windows' path limit once unzipped: " + n);
    }
    const m = JSON.parse(execFileSync("unzip", ["-p", zip, "manifest.json"], { encoding: "utf8" }));
    const referenced = [m.action.default_popup, m.options_ui.page, ...(m.background.scripts || [m.background.service_worker]), ...Object.values(m.icons),
      ...m.content_scripts.flatMap((s) => [...s.js, ...(s.css || [])])];
    for (const f of referenced) assert.ok(files.includes(f), "listed in the manifest and present: " + f);
  });
}

test("Chrome zip uses only keys Chromium browsers know (no warnings under chrome://extensions, edge://extensions, opera://extensions...)", () => {
  const m = JSON.parse(execFileSync("unzip", ["-p", ZIPS.chrome, "manifest.json"], { encoding: "utf8" }));
  assert.deepEqual(Object.keys(m).filter((k) => !CHROME_KEYS.includes(k)), []);
  assert.deepEqual(Object.keys(m.background), ["service_worker"]);
});

test("Chrome zip, unzipped into a folder with spaces and brackets, loads unpacked in Chromium and works", { skip: !chromium && "playwright not installed" }, async () => {
  const folder = unzip(ZIPS.chrome, path.join(tmp, "Wild Themes (1)", "wild-themes-extension"));
  const ctx = await chromium.launchPersistentContext(path.join(tmp, "profile"), { headless: false, args: ["--headless=new", `--disable-extensions-except=${folder}`, `--load-extension=${folder}`] });
  try {
    let [sw] = ctx.serviceWorkers(); if (!sw) sw = await ctx.waitForEvent("serviceworker", { timeout: 30000 });
    const id = sw.url().split("/")[2];
    const version = await sw.evaluate(() => chrome.runtime.getManifest().version);
    assert.match(version, /^\d+\.\d+\.\d+$/);
    const pop = await ctx.newPage(), errs = []; pop.on("pageerror", (e) => errs.push(e.message));
    await pop.goto(`chrome-extension://${id}/popup.html`);
    await pop.waitForSelector(".card", { timeout: 15000 });
    assert.ok((await pop.$$(".card")).length >= 18, "the popup lists the built-in themes");
    const opt = await ctx.newPage(); opt.on("pageerror", (e) => errs.push(e.message));
    await opt.goto(`chrome-extension://${id}/options.html`);
    await opt.waitForSelector("#pick", { timeout: 15000 });
    assert.deepEqual(errs, []);
  } finally { await ctx.close(); }
});

test("Firefox zip is what about:debugging, Developer Edition and signing take: Mozilla's linter passes the zip and the .xpi copy", { timeout: 600000 }, () => {
  const bin = path.join(__dirname, "node_modules", ".bin", "addons-linter");
  assert.deepEqual(fs.readFileSync(path.join(tmp, "wild-themes-firefox.zip")), fs.readFileSync(path.join(tmp, "wild-themes-firefox-unsigned.xpi")), ".xpi is the same file");
  if (!fs.existsSync(bin)) return;   // dev dependency not installed
  for (const f of ["wild-themes-firefox.zip", "wild-themes-firefox-unsigned.xpi"]) {
    const r = spawnSync(bin, [path.join(tmp, f), "--output", "json"], { encoding: "utf8", timeout: 280000 });
    const out = (r.stdout || "") + (r.stderr || ""), j = JSON.parse(out.slice(out.indexOf("{")));
    assert.deepEqual(j.errors.map((e) => `${e.code} ${e.message}`), [], f + ": linter errors");
  }
});
