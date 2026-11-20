// Smoothness: the frame-rate guard's decision, and the guard working in the real extension. On a device that
// can't keep up (simulated by a page that keeps the main thread busy) Auto holds the big and 3D layers still,
// then stops the scene, remembers that for the device, and "Auto" in the popup tests again.
// A healthy page must not be degraded. Real GPUs aren't available here: what's tested is the logic, not the look on your screen.
const test = require("node:test"), assert = require("node:assert/strict");
const fs = require("fs"), path = require("path"), os = require("os"), vm = require("vm");
process.env.PW_EXPERIMENTAL_SERVICE_WORKER_NETWORK_EVENTS = "1";
let chromium; try { ({ chromium } = require("playwright")); } catch (_) { /* skipped below */ }

const EXT = path.join(__dirname, "..", "extension");
const ctx0 = { self: {} }; vm.runInNewContext(fs.readFileSync(path.join(EXT, "lib", "perf.js"), "utf8"), ctx0);
const judge = (gaps, ms = 2000) => ctx0.self.WWPerf.judge(gaps, ms);
const steady = (ms, n) => Array(n).fill(ms);

test("guard decision: steady 60 fps is fine; late, uneven or too few frames are not", () => {
  assert.equal(judge(steady(16.7, 120)), 0, "60 fps");
  assert.equal(judge(steady(8.3, 240)), 0, "120 fps");
  assert.equal(judge(steady(33.3, 60)), 1, "30 fps: the scene should be lighter");
  assert.equal(judge([...steady(16.7, 100), ...steady(70, 20)]), 1, "one frame in six takes 70 ms: jittery");
  assert.equal(judge([...steady(16.7, 108), ...steady(50, 12)]), 0, "one frame in ten takes 50 ms: tolerated");
  assert.equal(judge(steady(16.7, 12)), 1, "only a handful of frames in two seconds");
  assert.equal(judge(null), 1);
});

const PAGE = `<!doctype html><html><body style="margin:0;background:#f9fbfd"><div id="docs-chrome" style="height:84px"><div role="menubar" style="height:34px">File</div><div role="toolbar" style="height:42px">Tools</div></div>
<div id="docs-editor-container" style="position:absolute;top:84px;bottom:0;left:0;right:0;background:#f9fbfd"><div id="docs-editor" style="position:absolute;inset:0"><div class="kix-appview-editor" style="position:absolute;inset:0;overflow:auto;background-color: rgb(249, 251, 253);"><div class="kix-page-paginated" style="position:relative;width:816px;height:1056px;margin:20px auto;background:#fff"></div></div></div></div></body></html>`;

test("guard in the real extension: busy device gets lighter, then still, and it's remembered", { skip: !chromium && "playwright not installed", timeout: 180000 }, async (t) => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "wt-sm-"));
  const ctx = await chromium.launchPersistentContext(dir, { headless: false, args: ["--headless=new", `--disable-extensions-except=${EXT}`, `--load-extension=${EXT}`] });
  try {
    await ctx.route("https://min3scon.github.io/**", (r) => r.fulfill({ status: 404, body: "" }));
    await ctx.route("https://docs.google.com/**", (r) => r.fulfill({ contentType: "text/html", body: PAGE }));
    let [sw] = ctx.serviceWorkers(); if (!sw) sw = await ctx.waitForEvent("serviceworker");
    const extId = sw.url().split("/")[2];
    const perf = () => sw.evaluate(() => chrome.storage.local.get("perf")).then((v) => v.perf && v.perf.level);
    const waitLevel = async (n, ms) => { const t0 = Date.now(); while (Date.now() - t0 < ms) { if ((await perf()) >= n) return true; await new Promise((r) => setTimeout(r, 500)); } return false; };
    const open = async (id) => { const pg = await ctx.newPage(); await pg.goto(`https://docs.google.com/document/d/${id}/edit`); return pg; };
    const busy = (pg) => pg.evaluate(() => setInterval(() => { const t = performance.now(); while (performance.now() - t < 70); }, 90));   // frames take 70-160 ms

    await t.test("a healthy page is never degraded", async () => {
      await sw.evaluate(() => chrome.storage.sync.set({ theme: "winter", on: true, motion: "auto", sparks: false, docs: {} }));
      await sw.evaluate(() => chrome.storage.local.remove("perf"));
      const pg = await open("HEALTHY");
      await pg.waitForSelector('#ww-bg[data-scene="winter"]', { state: "attached", timeout: 15000 });
      await pg.waitForTimeout(14000);                                    // two looks: at 3 s and 8 s
      assert.equal(await perf(), undefined, "nothing recorded");
      assert.equal(await pg.$eval("#ww-bg", (n) => n.classList.contains("still") || n.classList.contains("lite")), false);
      await pg.close();
    });

    await t.test("a busy page: big and 3D layers hold still first (light), then the scene stops", async () => {
      await sw.evaluate(() => chrome.storage.sync.set({ theme: "synthwave", on: true, motion: "auto", sparks: false, docs: {} }));
      await sw.evaluate(() => chrome.storage.local.remove("perf"));
      const pg = await open("BUSY");
      await pg.waitForSelector('#ww-bg[data-scene="synthwave"]', { state: "attached", timeout: 15000 });
      await busy(pg);
      assert.ok(await waitLevel(1, 40000), "level 1 recorded");
      await pg.waitForSelector("#ww-bg.lite", { state: "attached", timeout: 5000 });
      const layers = await pg.evaluate(() => [...document.querySelectorAll("#ww-bg>i")].map((i) => ({ heavy: i.hasAttribute("data-heavy"), states: i.getAnimations().map((a) => a.playState) })));
      assert.ok(layers.some((l) => l.heavy && l.states.every((s) => s === "paused")), "the perspective grid holds still: " + JSON.stringify(layers));
      assert.ok(layers.some((l) => !l.heavy && l.states.includes("running")), "the light layers keep moving: " + JSON.stringify(layers));
      assert.ok(await waitLevel(2, 60000), "level 2 recorded");
      await pg.waitForSelector("#ww-bg.still", { state: "attached", timeout: 5000 });
      assert.equal(await pg.$eval("#ww-bg", (n) => n.getAnimations({ subtree: true }).filter((a) => a.playState === "running").length), 0, "nothing moves");
      await pg.close();
    });

    await t.test("the verdict is remembered: a new page starts still, with no flash of motion", async () => {
      const pg = await open("NEXT");
      await pg.waitForSelector("#ww-bg.still", { state: "attached", timeout: 15000 });
      await pg.close();
    });

    await t.test("'On' ignores the verdict, 'Auto' in the popup forgets it and measures again", async () => {
      await sw.evaluate(() => chrome.storage.sync.set({ motion: "on" }));
      let pg = await open("ON");
      await pg.waitForSelector("#ww-bg", { state: "attached", timeout: 15000 });
      assert.equal(await pg.$eval("#ww-bg", (n) => n.classList.contains("still")), false, "On keeps the scene moving");
      await pg.close();
      const pop = await ctx.newPage(); await pop.goto(`chrome-extension://${extId}/popup.html`);
      await pop.waitForSelector('#motion [data-v="auto"]'); await pop.click('#motion [data-v="auto"]');
      await pop.waitForTimeout(500);
      assert.equal(await perf(), undefined, "verdict cleared");
      assert.equal((await sw.evaluate(() => chrome.storage.sync.get("motion"))).motion, "auto");
      await pop.close();
    });
  } finally { await ctx.close(); }
});
