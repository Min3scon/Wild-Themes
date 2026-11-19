// Loads the real extension in Chromium and checks the live catalogue: themes published on the
// website after this copy was made (a new built-in scene, a community picture theme) show up in the
// popup and render in Word for the web, and unsafe scene CSS is refused. Site and Word are mocked.
const test = require("node:test"), assert = require("node:assert/strict");
const fs = require("fs"), path = require("path"), os = require("os");
const H = require("./helpers.js");
process.env.PW_EXPERIMENTAL_SERVICE_WORKER_NETWORK_EVENTS = "1";   // lets the test mock the service worker's fetches
let chromium; try { ({ chromium } = require("playwright")); } catch (_) { /* skipped below */ }
const step = (m) => process.env.WT_DEBUG && console.log("step:", m);

const SITE = "https://min3scon.github.io/Wild-Themes/";
const EXT = path.join(__dirname, "..", "extension");
const png = fs.readFileSync(path.join(__dirname, "fixtures", "backdrop.jpg"));
const catalog = {
  format: 1, version: "test1", updated: "2026-11-19T00:00:00Z", themes: [
    { id: "synthwave", kind: "builtin", name: "Synthwave Sunset", mode: "dark", accent: "#ff2bd6", accent2: "#21e6ff" },
    { id: "testscene", kind: "builtin", name: "Test Scene", mode: "dark", accent: "#12ab34", accent2: "#3456cd", fx: "dot",
      doc: { dk1: "101010", dk2: "123456", lt2: "EEF5EE", accent1: "12AB34", accent2: "3456CD" }, scene: "extension/scenes/testscene.css", preview: "extension/previews/testscene.webp" },
    { id: "evilscene", kind: "builtin", name: "Evil", mode: "dark", accent: "#000000", scene: "extension/scenes/evilscene.css" },
    { id: "c-harbour", kind: "community", name: "Harbour", author: "tester", mode: "dark", accent: "#4dc3ff", accent2: "#ffb31a", effect: "none",
      colors: null, image: "community/harbour/bg.jpg", thumb: "community/harbour/thumb.webp" },
  ],
};

test("extension picks up themes published on the website", { skip: !chromium && "playwright not installed" }, async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "wt-ext-"));
  const ctx = await chromium.launchPersistentContext(dir, { headless: false, args: ["--headless=new", `--disable-extensions-except=${EXT}`, `--load-extension=${EXT}`] });
  try {
    await ctx.route(SITE + "**", (r) => r.fulfill({ status: 404, body: "" }));   // registered first: later routes win
    await ctx.route(SITE + "catalog.json", (r) => r.fulfill({ json: catalog }));
    await ctx.route(SITE + "extension/scenes/testscene.css", (r) => r.fulfill({ contentType: "text/css", body: '#ww-bg[data-scene="testscene"]{background:linear-gradient(#12ab34,#3456cd)!important}' }));
    await ctx.route(SITE + "extension/scenes/evilscene.css", (r) => r.fulfill({ contentType: "text/css", body: '#ww-bg{background:url(https://evil.example/x.png)}' }));
    await ctx.route(SITE + "community/harbour/bg.jpg", (r) => r.fulfill({ contentType: "image/jpeg", body: png }));
    step("routes");
    let [sw] = ctx.serviceWorkers(); if (!sw) sw = await ctx.waitForEvent("serviceworker");
    const id = sw.url().split("/")[2];

    // the service worker syncs on install; ask again to be sure the mocked catalogue was used
    const r = await sw.evaluate(() => self.sync().catch((e) => ({ error: String(e && e.message || e) })));
    step("synced " + JSON.stringify(r));
    const stored = await sw.evaluate(() => chrome.storage.local.get(["remote", "community", "scene:testscene", "scene:evilscene", "catalog"]));
    assert.ok(stored.remote && stored.remote.testscene, "new built-in theme stored: " + JSON.stringify(r));
    assert.ok(!stored.remote.evilscene, "theme with unsafe CSS refused");
    assert.ok(!stored["scene:evilscene"]);
    assert.match(stored["scene:testscene"], /linear-gradient/);
    assert.equal(stored.community[0].id, "c-harbour");
    assert.equal(stored.catalog.version, "test1");

    step("stored ok");
    // popup lists them
    const pop = await ctx.newPage();
    await pop.goto(`chrome-extension://${id}/popup.html`);
    await pop.waitForSelector('.card[data-id="testscene"]');
    await pop.click('[data-f="community"]');
    await pop.waitForSelector('.card[data-id="c-harbour"]');
    await pop.screenshot({ path: path.join(H.OUT, "ext-popup-community.png") });

    step("popup ok");
    // Word for the web (mocked) renders the new built-in scene from the website's CSS
    await sw.evaluate(() => chrome.storage.sync.set({ theme: "testscene", on: true }));
    const WORD = "https://word-edit.officeapps.live.com/we/wordeditorframe.aspx?WOPISrc=https%3A%2F%2Fexample%2Fwopi%2Ffiles%2F1";
    await ctx.route("https://word-edit.officeapps.live.com/**", (r) => r.fulfill({ contentType: "text/html",
      body: '<!doctype html><html><body style="margin:0;background:#e6e6e6"><div id="canvas" style="position:absolute;inset:0;background:#e6e6e6"><div id="page" contenteditable="true" style="width:600px;height:800px;margin:40px auto;background:#fff">Hello</div></div></body></html>' }));
    const word = await ctx.newPage();
    await word.goto(WORD);
    await word.waitForSelector('#ww-bg[data-scene="testscene"]', { state: "attached", timeout: 15000 });
    await word.waitForFunction(() => /linear-gradient/.test(getComputedStyle(document.querySelector("#ww-bg")).backgroundImage), null, { timeout: 10000 });

    step("word scene ok");
    // community picture theme: fetched once from the website, then drawn
    await sw.evaluate(() => chrome.storage.sync.set({ theme: "c-harbour" }));
    await word.waitForSelector('#ww-bg[data-scene="custom"] canvas', { state: "attached", timeout: 15000 });
    const img = await sw.evaluate(() => chrome.storage.local.get("img:c-harbour"));
    assert.match(img["img:c-harbour"], /^data:image\/jpeg;base64,/);
    await word.screenshot({ path: path.join(H.OUT, "ext-word-community.png") });
  } finally { await ctx.close(); }
});
