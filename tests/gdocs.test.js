// Google Docs, Slides and Sheets in the real extension (Chromium), on stand-in pages. The Docs page copies the
// structure the live editor is documented to use by other extensions: #docs-chrome (bars with their own colour) above
// #docs-editor-container, a transparent outline panel on the left, #docs-editor > .kix-appview-editor (inline grey
// background) holding .kix-page-paginated (transparent) whose white is painted by .kix-canvas-tile-content tiles,
// a side strip on the right, and the hidden typing iframe. These are NOT the live Google pages: Google changes its
// markup and the real editors need a Google sign-in, so the live result has to be checked in a browser.
// If a scene doesn't appear there, the extension changes nothing, and the popup's "Copy page details" says why.
const test = require("node:test"), assert = require("node:assert/strict");
const fs = require("fs"), path = require("path"), os = require("os");
const H = require("./helpers.js");
process.env.PW_EXPERIMENTAL_SERVICE_WORKER_NETWORK_EVENTS = "1";
let chromium; try { ({ chromium } = require("playwright")); } catch (_) { /* skipped below */ }

const EXT = path.join(__dirname, "..", "extension");
const fnv = (s) => { let h = 2166136261; for (let i = 0; i < s.length; i++) h = Math.imul(h ^ s.charCodeAt(i), 16777619); return (h >>> 0).toString(36); };

// like the live editor: pages are transparent wrappers, white comes from canvas tiles with an inline background
const DOCS = `<!doctype html><html><head><title>Doc</title></head><body class="docs-material" style="margin:0;background:#f9fbfd;font:14px sans-serif">
<div id="docs-chrome" style="background:#f9fbfd;height:84px">
  <div id="docs-menubar" role="menubar" style="box-sizing:border-box;height:34px;padding:6px 12px">File Edit View Insert Format</div>
  <div id="docs-toolbar" role="toolbar" style="box-sizing:border-box;height:42px;padding:6px 12px;background:#edf2fa;border-radius:20px;margin:0 12px">B I U</div>
</div>
<div id="docs-editor-container" style="position:absolute;top:84px;bottom:0;left:0;right:0;background:#f9fbfd">
  <div class="left-sidebar-container" style="position:absolute;left:0;top:0;bottom:0;width:240px;padding:12px">Outline: Heading one</div>
  <div id="kix-horizontal-ruler-container" style="position:absolute;top:0;left:240px;right:56px;height:18px;background:#f9fbfd"><div class="docs-horizontal-ruler">1 2 3 4 5</div></div>
  <div id="kix-vertical-ruler" class="docs-vertical-ruler" style="position:absolute;left:240px;top:18px;bottom:0;width:14px;background:#f9fbfd">1 2 3</div>
  <div id="docs-editor" style="position:absolute;top:18px;bottom:0;left:254px;right:56px;background:#f9fbfd">
    <div class="kix-appview-editor-container" style="position:absolute;inset:0">
      <div class="kix-appview-editor" style="position:absolute;inset:0;overflow:auto;background-color: rgb(249, 251, 253);">
        <div class="kix-page-column" style="position:relative;width:816px;height:1056px;margin:20px auto;overflow:hidden"><div class="kix-page-paginated" style="position:relative;width:816px;height:1056px">
          <canvas class="kix-canvas-tile-content" width="816" height="400" style="position:absolute;left:0;top:0;width:816px;height:400px;background-color: rgb(255, 255, 255)"></canvas>
          <canvas class="kix-canvas-tile-content" width="816" height="400" style="position:absolute;left:0;top:400px;width:816px;height:400px;background-color: rgb(255, 255, 255)"></canvas>
          <canvas class="kix-canvas-tile-content" width="816" height="256" style="position:absolute;left:0;top:800px;width:816px;height:256px;background-color: rgb(255, 255, 255)"></canvas>
          <div class="kix-cursor"><div class="kix-cursor-caret" style="position:absolute;left:96px;top:96px;width:2px;height:18px;background:#000"></div></div>
        </div></div>
      </div>
    </div>
  </div>
  <div id="side-strip" style="position:absolute;right:0;top:0;bottom:0;width:56px;background:#f9fbfd">icons</div>
</div>
<iframe class="docs-texteventtarget-iframe" style="position:absolute;left:-2000px;top:0;width:20px;height:20px" srcdoc="<body contenteditable=true></body>"></iframe>
</body></html>`;

// simple build: the page wrapper itself is white
const DOCS_SIMPLE = DOCS.replace('class="kix-page-paginated" style="position:relative;width:816px;height:1056px"', 'class="kix-page-paginated" style="position:relative;width:816px;height:1056px;background:#fff"');
// Google renamed its classes: only the general scan (white box on a grey surround) can find the page
const DOCS_RENAMED = DOCS_SIMPLE.replace(/kix-page-column/g, "x5").replace(/kix-page-paginated/g, "x1").replace(/kix-appview-editor-container/g, "x2").replace(/kix-appview-editor/g, "x3").replace(/kix-canvas-tile-content/g, "x4");

// the page box is the only thing of its size (no outer column): the glow sits on the box that is itself inverted in dark mode
const DOCS_FLAT = DOCS.replace('class="kix-page-column" style="position:relative;width:816px;height:1056px;margin:20px auto;overflow:hidden"', 'class="kix-page-column" style="position:relative;width:900px;height:1056px;margin:20px auto"');
// text is DOM text on a white page box (no canvas), as in an older rendering mode
const DOCS_DOMTEXT = DOCS_SIMPLE.replace(/<canvas class="kix-canvas-tile-content"[^>]*><\/canvas>/g, "").replace('<div class="kix-cursor">', '<div id="words" style="position:absolute;left:60px;top:60px;font:bold 70px sans-serif;color:#000;line-height:1.2;white-space:nowrap">TEXT TEXT</div><div class="kix-cursor">');
// Google renamed the page box classes but the canvases are still there
const DOCS_NOPAGE = DOCS.replace(/kix-page-paginated/g, "zz1").replace(/kix-page-column/g, "zz2");
// a document with its own dark page colour
const DOCS_ALREADYDARK = DOCS.replace(/background-color: rgb\(255, 255, 255\)/g, "background-color: rgb(30, 30, 30)");

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
    const html = { "/document/d/DOCA/edit": DOCS, "/document/d/DOCB/edit": DOCS, "/document/d/SIMPLE/edit": DOCS_SIMPLE, "/document/d/RENAMED/edit": DOCS_RENAMED, "/document/d/FLAT/edit": DOCS_FLAT, "/document/d/DOMTEXT/edit": DOCS_DOMTEXT, "/document/d/NOPAGE/edit": DOCS_NOPAGE, "/document/d/ALREADYDARK/edit": DOCS_ALREADYDARK,
      "/presentation/d/SLD1/edit": SLIDES, "/spreadsheets/d/SHT1/edit": SHEETS, "/forms/d/FRM1/edit": DOCS };
    await ctx.route("https://docs.google.com/**", (r) => { const p = new URL(r.request().url()).pathname; r.fulfill({ contentType: "text/html", body: html[p] || DOCS }); });
    let [sw] = ctx.serviceWorkers(); if (!sw) sw = await ctx.waitForEvent("serviceworker");
    const open = async (p) => { const pg = await ctx.newPage(); const errs = []; pg.on("pageerror", (e) => errs.push(e.message)); await pg.goto("https://docs.google.com" + p); pg.errs = errs; return pg; };
    await sw.evaluate(() => chrome.storage.sync.set({ theme: "synthwave", on: true, sparks: true, motion: "off", docs: {}, docsDark: "off", docsClean: true }));

    await t.test("Docs (like the live editor): scene over the editor area, page tiles stay white, panels and bars keep their look", async () => {
      const pg = await open("/document/d/DOCA/edit");
      await pg.waitForSelector('#ww-bg[data-scene="synthwave"]', { state: "attached", timeout: 15000 });
      await pg.waitForSelector(".kix-page-column[data-ww-page]", { state: "attached", timeout: 15000 });
      assert.equal(await pg.$eval(".kix-page-paginated", (n) => n.hasAttribute("data-ww-page")), false, "glow sits on the outermost box of the page's size, not an inner wrapper");
      for (const sel of [".kix-appview-editor", "#docs-editor", "#docs-editor-container"]) assert.equal(await pg.$eval(sel, (n) => n.hasAttribute("data-ww-clear")), true, sel + " made see-through");
      assert.equal(await pg.$eval("#docs-chrome", (n) => n.hasAttribute("data-ww-clear")), false, "app chrome untouched");
      assert.equal(await pg.$eval("#docs-toolbar", (n) => getComputedStyle(n).backgroundColor), "rgb(237, 242, 250)", "toolbar keeps its colour");
      assert.equal(await pg.$eval(".kix-canvas-tile-content", (n) => getComputedStyle(n).backgroundColor), "rgb(255, 255, 255)", "page tiles stay white");
      const box = await pg.evaluate(() => { const e = document.querySelector(".kix-appview-editor").getBoundingClientRect(), h = document.querySelector("#kix-horizontal-ruler-container").getBoundingClientRect(), v = document.querySelector("#kix-vertical-ruler").getBoundingClientRect(), b = document.querySelector("#ww-bg").style; return { e, h, v, bg: { top: parseFloat(b.top), left: parseFloat(b.left), right: parseFloat(b.left) + parseFloat(b.width), bottom: parseFloat(b.top) + parseFloat(b.height) } }; });
      assert.deepEqual([box.bg.top, box.bg.left, box.bg.right, box.bg.bottom].map(Math.round), [box.h.top, box.v.left, box.e.right, box.e.bottom].map(Math.round), "scene covers the editor and the space the rulers took");
      await pg.screenshot({ path: path.join(H.OUT, "gdocs-docs.png") });
      const shot = path.join(H.OUT, "gdocs-docs.png"), W = 1280, Hh = 720;
      const px = (x0, y0, x1, y1) => H.sample(shot, [x0 / W, y0 / Hh, x1 / W, y1 / Hh]);
      assert.ok(H.lum(px(270, 300, 330, 360)) < 90, "scene beside the page (inside the editor), got " + px(270, 300, 330, 360));
      assert.ok(H.lum(px(500, 87, 900, 99)) < 90, "scene where the horizontal ruler was, got " + px(500, 87, 900, 99));
      assert.ok(H.lum(px(242, 300, 252, 400)) < 90, "scene where the vertical ruler was, got " + px(242, 300, 252, 400));
      assert.ok(H.lum(px(560, 300, 700, 400)) > 240, "page white in the middle");
      assert.ok(H.lum(px(20, 300, 220, 400)) > 200, "outline panel keeps the light page colour, got " + px(20, 300, 220, 400));
      assert.ok(H.lum(px(1230, 300, 1270, 400)) > 200, "right side strip keeps it too, got " + px(1230, 300, 1270, 400));
      assert.ok(H.lum(px(300, 5, 900, 28)) > 200, "menu bar light");
      assert.deepEqual(pg.errs, []);
      await pg.close();
    });

    await t.test("Docs with a white page wrapper, and with renamed classes (general scan): scene still around the page", async () => {
      for (const p of ["/document/d/SIMPLE/edit", "/document/d/RENAMED/edit"]) {
        const pg = await open(p);
        try {
          await pg.waitForSelector('#ww-bg[data-scene="synthwave"]', { state: "attached", timeout: 15000 });
          await pg.waitForFunction(() => document.querySelector("[data-ww-page]") && document.querySelector("[data-ww-clear]"), null, { timeout: 15000 });
          const shot = path.join(H.OUT, "gdocs-variant.png"); await pg.screenshot({ path: shot });
          assert.ok(H.lum(H.sample(shot, [270 / 1280, 300 / 720, 330 / 1280, 360 / 720])) < 120, p + ": scene beside the page");
          assert.ok(H.lum(H.sample(shot, [560 / 1280, 300 / 720, 700 / 1280, 400 / 720])) > 240, p + ": page white");
        } finally { await pg.close(); }
      }
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
      await pg.close();
    });

    await t.test("Docs: the page is looked at straight away after a change, then at most about six times a second", async () => {
      const pg = await open("/document/d/DOCA/edit");
      await pg.waitForSelector("[data-ww-page]", { state: "attached", timeout: 15000 });
      await pg.waitForTimeout(1500);
      // each look re-sets data-ww-page on the page, which shows as an attribute change we can count
      const counts = await pg.evaluate(async () => {
        const page = document.querySelector("[data-ww-page]"); let n = 0;
        new MutationObserver((r) => { n += r.filter((x) => x.attributeName === "data-ww-page").length; }).observe(page, { attributes: true });
        const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
        await sleep(300);
        document.body.append(document.createElement("i"));                    // one change after a quiet spell
        await sleep(120);
        const quick = n;
        n = 0;
        const t0 = performance.now();
        while (performance.now() - t0 < 1000) { document.body.append(document.createElement("i")); await sleep(4); }   // a page that never stops changing
        await sleep(250);
        return { quick, storm: n };
      });
      assert.ok(counts.quick >= 1, "first change looked at within 120 ms: " + counts.quick);
      assert.ok(counts.storm >= 2 && counts.storm <= 9, "about six looks in a second of constant change, got " + counts.storm);
      await pg.close();
    });

    await t.test("Docs: 'Copy page details' reports what the extension sees, without any document text", async () => {
      const pg = await open("/document/d/DOCA/edit");
      await pg.evaluate(() => { document.querySelector(".kix-appview-editor").insertAdjacentHTML("beforeend", "<p id=secret>my private sentence</p>"); });
      await pg.waitForSelector("[data-ww-page]", { state: "attached", timeout: 15000 });
      const id = sw.url().split("/")[2], pop = await ctx.newPage();
      await pop.goto(`chrome-extension://${id}/popup.html`);
      const d = await pop.evaluate(async () => {
        const [tab] = await chrome.tabs.query({ url: "https://docs.google.com/document/*" });
        const status = await new Promise((ok) => chrome.tabs.sendMessage(tab.id, "ww-status", ok));
        const details = await new Promise((ok) => chrome.tabs.sendMessage(tab.id, "ww-details", ok));
        return { status, details };
      });
      assert.equal(d.status.google, true); assert.equal(d.status.shown, true);
      assert.equal(d.details.scene, true); assert.equal(d.details.editor, true); assert.equal(d.details.pages, 1);
      assert.ok(d.details.at[0].stack.length > 3 && d.details.at[0].stack.some((s) => /kix-appview-editor/.test(s) && /CLEARED/.test(s)), "element stack at the centre");
      assert.ok(!/private sentence/.test(JSON.stringify(d.details)), "no document text");
      assert.ok(!/DOCA/.test(JSON.stringify(d.details)), "document id masked");
      await pop.close(); await pg.close();
    });

    await t.test("Docs: the page gets the same glow outline as in Word, even when an inner wrapper would clip it", async () => {
      await sw.evaluate(() => chrome.storage.sync.set({ theme: "synthwave", docsDark: "off" }));
      const pg = await open("/document/d/DOCA/edit");
      await pg.waitForSelector(".kix-page-column[data-ww-page]", { state: "attached", timeout: 15000 });
      const shadow = await pg.$eval(".kix-page-column", (n) => getComputedStyle(n).boxShadow);
      assert.match(shadow, /rgba?\(/, "box-shadow set: " + shadow);
      await pg.screenshot({ path: path.join(H.OUT, "gdocs-glow.png") });
      const left = await pg.$eval(".kix-page-column", (n) => Math.round(n.getBoundingClientRect().left));
      const shot = path.join(H.OUT, "gdocs-glow.png"), W = 1280, Hh = 720;
      const px = (x0, y0, x1, y1) => H.sample(shot, [x0 / W, y0 / Hh, x1 / W, y1 / Hh]);
      const near = px(left - 8, 160, left - 3, 240), far = px(left - 60, 160, left - 55, 240);
      assert.ok(near[0] - far[0] > 12, `glow: red ${near[0]} just outside the page vs ${far[0]} further away`);
      await pg.close();
    });

    await t.test("Docs dark mode: plain inversion, page colour untouched; Auto follows the theme, On and Off override it, live", async () => {
      const shot = path.join(H.OUT, "gdocs-dark.png"), px = (x0, y0, x1, y1) => H.sample(shot, [x0 / 1280, y0 / 720, x1 / 1280, y1 / 720]);
      const page = async (theme, dark, id = "DOCA") => {
        await sw.evaluate(([th, d]) => chrome.storage.sync.set({ theme: th, docsDark: d }), [theme, dark]);
        const pg = await open(`/document/d/${id}/edit`);
        await pg.waitForSelector("[data-ww-page]", { state: "attached", timeout: 15000 });
        await pg.waitForTimeout(400);
        return pg;
      };
      let pg = await page("synthwave", "auto");                                   // dark theme + Auto
      assert.equal(await pg.$eval("html", (n) => n.hasAttribute("data-ww-dark")), true);
      assert.equal(await pg.$eval(".kix-canvas-tile-content", (n) => getComputedStyle(n).backgroundColor), "rgb(255, 255, 255)", "the page's own colour is not changed");
      assert.match(await pg.$eval(".kix-page-paginated", (n) => getComputedStyle(n).filter), /invert\(1\).*hue-rotate\(180deg\)/);
      assert.equal(await pg.$eval(".kix-canvas-tile-content", (n) => getComputedStyle(n).filter), "none", "inverted once, not twice");
      await pg.screenshot({ path: shot });
      assert.ok(H.lum(px(500, 200, 900, 400)) < 25, "plain dark, got " + px(500, 200, 900, 400));
      await sw.evaluate(() => chrome.storage.sync.set({ docsDark: "off" }));          // live, no reload
      await pg.waitForFunction(() => !document.documentElement.hasAttribute("data-ww-dark"), null, { timeout: 5000 });
      await pg.screenshot({ path: shot });
      assert.ok(H.lum(px(500, 200, 900, 400)) > 240, "back to white");
      await pg.close();
      pg = await page("mint", "auto");                                            // light theme + Auto: the page is left as it is
      assert.equal(await pg.$eval("html", (n) => n.hasAttribute("data-ww-dark")), false);
      await pg.close();
      pg = await page("mint", "on");                                              // On beats a light theme
      assert.equal(await pg.$eval("html", (n) => n.hasAttribute("data-ww-dark")), true);
      await pg.screenshot({ path: shot });
      assert.ok(H.lum(px(500, 200, 900, 400)) < 25, "dark page on a light theme");
      await pg.close();
      await sw.evaluate(() => chrome.storage.sync.set({ theme: "synthwave", docsDark: "auto" }));
    });

    await t.test("Docs dark mode: text stays readable whatever draws it (DOM text on a white box, canvases without the page box), and a document's own dark page isn't flipped", async () => {
      const shot = path.join(H.OUT, "gdocs-dark2.png");
      await sw.evaluate(() => chrome.storage.sync.set({ theme: "synthwave", docsDark: "auto" }));
      // DOM text on a white page box: black text turns light on the dark page
      let pg = await open("/document/d/DOMTEXT/edit");
      await pg.waitForSelector("html[data-ww-dark] [data-ww-page]", { state: "attached", timeout: 15000 });
      await pg.waitForTimeout(400);
      const r = await pg.$eval("#words", (n) => { const b = n.getBoundingClientRect(); return [b.left, b.top, b.right, b.bottom].map(Math.round); });
      await pg.screenshot({ path: shot });
      const words = H.sample(shot, [r[0] / 1280, r[1] / 720, r[2] / 1280, r[3] / 720]), below = H.sample(shot, [r[0] / 1280, (r[3] + 40) / 720, r[2] / 1280, (r[3] + 140) / 720]);
      assert.ok(H.lum(below) < 12, "dark page around the text, got " + below);
      assert.ok(H.lum(words) > 30, "light text on it, got " + words + " (page " + below + ")");
      await pg.close();
      // no page box classes: the canvases are inverted on their own
      pg = await open("/document/d/NOPAGE/edit");
      await pg.waitForSelector("html[data-ww-dark]", { state: "attached", timeout: 15000 });
      assert.match(await pg.$eval(".kix-canvas-tile-content", (n) => getComputedStyle(n).filter), /invert\(1\)/);
      await pg.close();
      // a document whose own page colour is already dark is left alone in Auto, inverted by On
      pg = await open("/document/d/ALREADYDARK/edit");
      await pg.waitForSelector("[data-ww-page]", { state: "attached", timeout: 15000 });
      await pg.waitForTimeout(500);
      assert.equal(await pg.$eval("html", (n) => n.hasAttribute("data-ww-dark")), false, "already dark: Auto leaves it");
      await sw.evaluate(() => chrome.storage.sync.set({ docsDark: "on" }));
      await pg.waitForFunction(() => document.documentElement.hasAttribute("data-ww-dark"), null, { timeout: 5000 });
      await pg.close();
      await sw.evaluate(() => chrome.storage.sync.set({ docsDark: "auto" }));
    });

    await t.test("Docs dark mode: the glow keeps the theme's colour, on an outer box and on the inverted page box itself", async () => {
      await sw.evaluate(() => chrome.storage.sync.set({ theme: "synthwave", docsDark: "auto" }));
      for (const id of ["DOCA", "FLAT"]) {
        const pg = await open(`/document/d/${id}/edit`);
        await pg.waitForSelector("html[data-ww-dark] [data-ww-page]", { state: "attached", timeout: 15000 });
        await pg.waitForTimeout(400);
        const left = await pg.$eval(".kix-page-paginated", (n) => Math.round(n.getBoundingClientRect().left));
        const shot = path.join(H.OUT, "gdocs-dark-glow.png"); await pg.screenshot({ path: shot });
        const px = (x0, y0, x1, y1) => H.sample(shot, [x0 / 1280, y0 / 720, x1 / 1280, y1 / 720]);
        const near = px(left - 8, 160, left - 3, 240), far = px(left - 30, 160, left - 25, 240);
        assert.ok(near[0] - far[0] > 12 && near[0] > near[1], `${id}: glow stays pink (red above green), got near ${near} far ${far}`);
        await pg.close();
      }
    });

    await t.test("Docs clean view: rulers and scrollbar are gone, scrolling still works; the switch brings them back", async () => {
      const pg = await open("/document/d/DOCA/edit");
      await pg.waitForSelector(".kix-page-column[data-ww-page]", { state: "attached", timeout: 15000 });
      await pg.waitForTimeout(400);
      const state = () => pg.evaluate(() => { const e = document.querySelector(".kix-appview-editor"); return { h: getComputedStyle(document.querySelector("#kix-horizontal-ruler-container")).visibility, v: getComputedStyle(document.querySelector("#kix-vertical-ruler")).visibility, bar: e.offsetWidth - e.clientWidth, canScroll: e.scrollHeight > e.clientHeight }; });
      let s0 = await state();
      assert.deepEqual([s0.h, s0.v, s0.bar, s0.canScroll], ["hidden", "hidden", 0, true], JSON.stringify(s0));
      await pg.evaluate(() => { document.querySelector(".kix-appview-editor").scrollTop = 200; });
      assert.equal(await pg.$eval(".kix-appview-editor", (n) => n.scrollTop), 200, "still scrolls");
      await sw.evaluate(() => chrome.storage.sync.set({ docsClean: false }));
      await pg.waitForFunction(() => !document.documentElement.hasAttribute("data-ww-clean"), null, { timeout: 5000 });
      s0 = await state();
      assert.equal(s0.h, "visible"); assert.equal(s0.v, "visible"); assert.ok(s0.bar > 0, "scrollbar back: " + s0.bar);
      await sw.evaluate(() => chrome.storage.sync.set({ docsClean: true }));
      await pg.close();
    });

    await t.test("Popup on a Google tab: Dark page switch shows, reflects the setting and saves it", async () => {
      const id = sw.url().split("/")[2], pop = await ctx.newPage();
      await pop.addInitScript(() => {   // pretend the active tab is a Google Doc
        chrome.tabs.query = (q, cb) => cb([{ id: 1 }]);
        chrome.tabs.sendMessage = (tab, m, cb) => cb(m === "ww-status" ? { docKey: "abc", google: true, shown: true } : {});
      });
      await pop.goto(`chrome-extension://${id}/popup.html`);
      await pop.waitForSelector("#gnote:not([hidden])", { timeout: 5000 });
      assert.equal(await pop.getAttribute('#gdark [data-d="auto"]', "aria-checked"), "true");
      assert.equal(await pop.isChecked("#gclean"), true, "clean view on by default");
      await pop.uncheck("#gclean"); await pop.waitForTimeout(200);
      assert.equal((await sw.evaluate(() => chrome.storage.sync.get("docsClean"))).docsClean, false);
      await pop.check("#gclean"); await pop.waitForTimeout(200);
      await pop.click('#gdark [data-d="off"]');
      await pop.waitForTimeout(300);
      assert.equal((await sw.evaluate(() => chrome.storage.sync.get("docsDark"))).docsDark, "off");
      assert.equal(await pop.getAttribute('#gdark [data-d="off"]', "aria-checked"), "true");
      await sw.evaluate(() => chrome.storage.sync.set({ docsDark: "auto" }));
      await pop.close();
    });

    await t.test("Slides: slide on its grey surround is themed like a Word page, bare top bars stay light", async () => {
      const pg = await open("/presentation/d/SLD1/edit");
      await pg.waitForSelector('#ww-bg[data-scene="synthwave"]', { state: "attached", timeout: 15000 });
      await pg.waitForSelector(".punch-viewer-container[data-ww-page], #workspace[data-ww-clear]", { state: "attached", timeout: 15000 });
      await pg.screenshot({ path: path.join(H.OUT, "gdocs-slides.png") });
      const bars = H.sample(path.join(H.OUT, "gdocs-slides.png"), [0.3, 0.0, 0.9, 0.1]);
      assert.ok(H.lum(bars) > 200, "top bars keep the page's light colour, got " + bars);
      assert.deepEqual(pg.errs, []);
      await pg.close();
    });

    await t.test("Sheets: scene is placed over the grid pane", async () => {
      const pg = await open("/spreadsheets/d/SHT1/edit");
      await pg.waitForSelector('#ww-bg[data-scene="synthwave"]', { state: "attached", timeout: 15000 });
      await pg.waitForFunction(() => document.querySelector("#ww-bg").style.getPropertyValue("height") !== "", null, { timeout: 15000 });
      assert.equal(await pg.$eval("#ww-bg", (n) => n.style.top), "84px");
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
