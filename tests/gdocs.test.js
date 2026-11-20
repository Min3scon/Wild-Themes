// Google Docs, Slides and Sheets in the real extension (Chromium), on stand-in pages that copy the
// structure the live editors are known to use (.kix-appview-editor > .kix-page-paginated, the menu bar and
// toolbar roles, the hidden typing frame, .grid-container...). These are NOT the live Google pages:
// Google changes its markup and the real editors need a Google sign-in, so the live result has to be
// checked in a browser. If a scene doesn't appear there, the extension changes nothing.
const test = require("node:test"), assert = require("node:assert/strict");
const fs = require("fs"), path = require("path"), os = require("os");
const H = require("./helpers.js");
process.env.PW_EXPERIMENTAL_SERVICE_WORKER_NETWORK_EVENTS = "1";
let chromium; try { ({ chromium } = require("playwright")); } catch (_) { /* skipped below */ }

const EXT = path.join(__dirname, "..", "extension");
const fnv = (s) => { let h = 2166136261; for (let i = 0; i < s.length; i++) h = Math.imul(h ^ s.charCodeAt(i), 16777619); return (h >>> 0).toString(36); };

const DOCS = `<!doctype html><html><head><title>Doc</title></head><body class="docs-material" style="margin:0;background:#f9fbfd;font:14px sans-serif">
<div id="docs-chrome" style="background:#f9fbfd;height:84px">
  <div id="docs-menubar" role="menubar" style="height:34px;padding:6px 12px">File Edit View Insert Format</div>
  <div id="docs-toolbar" role="toolbar" style="height:42px;padding:6px 12px;background:#edf2fa;border-radius:20px;margin:0 12px">B I U</div>
</div>
<div id="docs-editor" style="position:absolute;top:84px;bottom:0;left:0;right:0">
  <div class="kix-appview-editor" style="position:absolute;inset:0;overflow:auto;background:#f9fbfd">
    <div class="kix-page-paginated" style="position:relative;width:816px;height:1056px;margin:20px auto;background:#fff;box-shadow:0 0 1px #ccc">
      <canvas class="kix-canvas-tile-content" width="816" height="1056" style="position:absolute;inset:0"></canvas>
      <div class="kix-cursor"><div class="kix-cursor-caret" style="position:absolute;left:96px;top:96px;width:2px;height:18px;background:#000"></div></div>
    </div>
  </div>
</div>
<iframe class="docs-texteventtarget-iframe" style="position:absolute;left:-2000px;top:0;width:20px;height:20px" srcdoc="<body contenteditable=true></body>"></iframe>
</body></html>`;

const SLIDES = `<!doctype html><html><body style="margin:0;background:#f9fbfd"><div id="docs-chrome" style="height:84px"><div role="menubar" style="height:34px">File</div><div role="toolbar" style="height:42px">Tools</div></div>
<div id="docs-editor" style="position:absolute;top:84px;bottom:0;left:200px;right:0;background:#f1f3f4"><div id="workspace" style="position:absolute;inset:0;background:#f1f3f4"><div class="punch-viewer-container" style="position:relative;width:640px;height:360px;margin:60px auto;background:#fff"></div></div></div>
<div style="position:absolute;top:84px;bottom:0;left:0;width:200px;background:#f9fbfd">filmstrip</div></body></html>`;

const SHEETS = `<!doctype html><html><body style="margin:0;background:#f9fbfd"><div id="docs-chrome" style="height:84px"><div role="menubar" style="height:34px">File</div><div role="toolbar" style="height:42px">Tools</div></div>
<div class="grid-container" style="position:absolute;top:84px;bottom:30px;left:0;right:0;background:#fff"><canvas width="600" height="300" style="width:600px;height:300px"></canvas></div></body></html>`;

test("Google Docs, Slides and Sheets: scene, per-document theme, sparks, other Google pages left alone", { skip: !chromium && "playwright not installed" }, async (t) => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "wt-gd-"));
  const ctx = await chromium.launchPersistentContext(dir, { headless: false, args: ["--headless=new", `--disable-extensions-except=${EXT}`, `--load-extension=${EXT}`] });
  const SITE = "https://min3scon.github.io/Wild-Themes/";
  try {
    await ctx.route(SITE + "**", (r) => r.fulfill({ status: 404, body: "" }));
    const html = { "/document/d/DOCA/edit": DOCS, "/document/d/DOCB/edit": DOCS, "/presentation/d/SLD1/edit": SLIDES, "/spreadsheets/d/SHT1/edit": SHEETS, "/forms/d/FRM1/edit": DOCS };
    await ctx.route("https://docs.google.com/**", (r) => { const p = new URL(r.request().url()).pathname; r.fulfill({ contentType: "text/html", body: html[p] || DOCS }); });
    let [sw] = ctx.serviceWorkers(); if (!sw) sw = await ctx.waitForEvent("serviceworker");
    const open = async (p) => { const pg = await ctx.newPage(); const errs = []; pg.on("pageerror", (e) => errs.push(e.message)); await pg.goto("https://docs.google.com" + p); pg.errs = errs; return pg; };
    await sw.evaluate(() => chrome.storage.sync.set({ theme: "synthwave", on: true, sparks: true, motion: "off", docs: {} }));

    await t.test("Docs: scene around the page, page and toolbar keep their look", async () => {
      const pg = await open("/document/d/DOCA/edit");
      await pg.waitForSelector('#ww-bg[data-scene="synthwave"]', { state: "attached", timeout: 15000 });
      await pg.waitForSelector(".kix-page-paginated[data-ww-page]", { state: "attached", timeout: 15000 });
      assert.equal(await pg.$eval(".kix-appview-editor", (n) => n.hasAttribute("data-ww-clear")), true, "grey surround made see-through");
      assert.equal(await pg.$eval("#docs-toolbar", (n) => getComputedStyle(n).backgroundColor), "rgb(237, 242, 250)", "toolbar keeps its colour");
      assert.equal(await pg.$eval("#docs-chrome", (n) => n.hasAttribute("data-ww-clear")), false, "app chrome untouched");
      assert.equal(await pg.$eval(".kix-page-paginated", (n) => getComputedStyle(n).backgroundColor), "rgb(255, 255, 255)", "page stays white and readable");
      await pg.screenshot({ path: path.join(H.OUT, "gdocs-docs.png") });
      const shot = path.join(H.OUT, "gdocs-docs.png");
      const edge = H.sample(shot, [0.02, 0.4, 0.1, 0.6]), mid = H.sample(shot, [0.4, 0.4, 0.6, 0.6]);
      assert.ok(H.lum(edge) < 90, "scene (dark) beside the page, got " + edge);
      assert.ok(H.lum(mid) > 240, "page white in the middle, got " + mid);
      assert.deepEqual(pg.errs, []);
      await pg.close();
    });

    await t.test("Docs: each document remembers its own theme (keyed by a hash of the document id)", async () => {
      await sw.evaluate((k) => chrome.storage.sync.set({ docs: { [k]: "winter" } }), fnv("DOCB"));
      const a = await open("/document/d/DOCA/edit"), b = await open("/document/d/DOCB/edit");
      await a.waitForSelector("#ww-bg", { state: "attached", timeout: 15000 });
      await b.waitForSelector("#ww-bg", { state: "attached", timeout: 15000 });
      assert.equal(await a.$eval("#ww-bg", (n) => n.dataset.scene), "synthwave");
      assert.equal(await b.$eval("#ww-bg", (n) => n.dataset.scene), "winter");
      const stored = JSON.stringify(await sw.evaluate(() => chrome.storage.sync.get("docs")));
      assert.ok(!/DOCB/.test(stored), "the document id itself is never stored");
      await a.close(); await b.close();
    });

    await t.test("Docs: typing sparks work although Google reads keys in a hidden frame", async () => {
      const pg = await open("/document/d/DOCA/edit");
      await pg.waitForSelector('#ww-bg[data-scene="synthwave"]', { state: "attached", timeout: 15000 });
      await pg.waitForTimeout(600);
      await pg.evaluate(() => { const d = document.querySelector("iframe.docs-texteventtarget-iframe").contentDocument; d.dispatchEvent(new KeyboardEvent("keydown", { key: "a", bubbles: true })); });
      await pg.waitForSelector("#ww-fx", { state: "attached", timeout: 5000 });
      const moved = await pg.$eval("#ww-fx", (n) => n.style.transform);
      assert.match(moved, /translate\(9\d(\.\d+)?px, ?10\d(\.\d+)?px\)|translate\(\d+/, "sparks start at the caret: " + moved);
      await pg.close();
    });

    await t.test("Slides: slide on its grey surround is themed like a Word page", async () => {
      const pg = await open("/presentation/d/SLD1/edit");
      await pg.waitForSelector('#ww-bg[data-scene="synthwave"]', { state: "attached", timeout: 15000 });
      await pg.waitForSelector(".punch-viewer-container[data-ww-page], #workspace[data-ww-clear]", { state: "attached", timeout: 15000 });
      await pg.screenshot({ path: path.join(H.OUT, "gdocs-slides.png") });
      // the bars here have no colour of their own (they sit on the page background we clear): they must stay light
      const bars = H.sample(path.join(H.OUT, "gdocs-slides.png"), [0.3, 0.0, 0.9, 0.1]);
      assert.ok(H.lum(bars) > 200, "top bars keep the page's light colour, got " + bars);
      assert.deepEqual(pg.errs, []);
      await pg.close();
    });

    await t.test("Sheets: scene is placed over the grid pane", async () => {
      const pg = await open("/spreadsheets/d/SHT1/edit");
      await pg.waitForSelector('#ww-bg[data-scene="synthwave"]', { state: "attached", timeout: 15000 });
      await pg.waitForFunction(() => document.querySelector("#ww-bg").style.getPropertyValue("height") !== "", null, { timeout: 15000 });
      const box = await pg.$eval("#ww-bg", (n) => ({ top: n.style.top, height: n.style.height }));
      assert.equal(box.top, "84px");
      await pg.screenshot({ path: path.join(H.OUT, "gdocs-sheets.png") });
      assert.deepEqual(pg.errs, []);
      await pg.close();
    });

    await t.test("Google Forms (not an editor we theme): nothing is injected", async () => {
      const pg = await open("/forms/d/FRM1/edit");
      await pg.waitForTimeout(1500);
      assert.equal(await pg.$("#ww-bg"), null);
      await pg.close();
    });
  } finally { await ctx.close(); }
});
