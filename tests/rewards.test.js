// Words typed per theme, the secret theme and the reward cutscenes.
// - Only real key presses count, one finished word at a time: pasted text, shortcuts, held-down repeats and events made by
//   a script never do. Paste itself still works in the document; it just doesn't count.
// - Each typed word has a 1 in 37,500 chance to unlock the secret theme. Until then it is a greyed-out "???" card with no
//   picture, can't be picked (not even by setting it directly) and isn't on the website, in the File studio or in the catalogue.
// - Finding it plays a cutscene in the page; passing a word milestone plays a smaller one. The dice are replaced in the
//   browser test (self.wtRandom) so the unlock can be forced; the odds themselves are checked separately.
const test = require("node:test"), assert = require("node:assert/strict");
const fs = require("fs"), path = require("path"), os = require("os"), vm = require("vm");
process.env.PW_EXPERIMENTAL_SERVICE_WORKER_NETWORK_EVENTS = "1";
const H = require("./helpers.js");
let chromium; try { ({ chromium } = require("playwright")); } catch (_) { /* skipped below */ }

const ROOT = path.join(__dirname, ".."), EXT = path.join(ROOT, "extension");
const lib = { self: {} }; vm.runInNewContext(fs.readFileSync(path.join(EXT, "lib", "rewards.js"), "utf8"), lib);
const RW = lib.self.WWRewards;

test("which key presses make words", () => {
  const k = (key, more = {}) => RW.keyKind({ key, isTrusted: true, getModifierState: (m) => !!(more.altGr && m === "AltGraph"), ...more });
  const count = (keys) => { let n = 0, w = false; for (const x of keys) { const r = k(x); if (r === "char") w = true; else if (r === "end" && w) { w = false; n++; } } return n; };
  assert.equal(count([..."hello world, don't stop typing. "]), 5, "apostrophes keep a word whole; punctuation and spaces end it");
  assert.equal(count([..."well-known 123 Ünïcödé 日本 "]), 4, "hyphens, digits and any alphabet");
  assert.equal(count(["a", "Enter", "b", "Tab"]), 2, "Enter and Tab end a word");
  assert.equal(count([..."   ,,, ..."]), 0, "no letters, no word");
  assert.equal(k("v", { ctrlKey: true }), "", "Ctrl+V (paste) is ignored");
  assert.equal(k("v", { metaKey: true }), "", "Cmd+V (paste) is ignored");
  assert.equal(k("a", { repeat: true }), "", "a held-down key is ignored");
  assert.equal(k("a", { isTrusted: false }), "", "events made by a script are ignored");
  assert.equal(k("a", { isComposing: true }), "", "text still being composed is ignored");
  assert.equal(k("Shift"), ""); assert.equal(k("ArrowLeft"), ""); assert.equal(k("Backspace"), "");
  assert.equal(k("ą", { ctrlKey: true, altKey: true, altGr: true }), "char", "AltGr letters count (Polish, German...)");
  assert.equal(RW.ODDS, 37500); assert.equal(RW.SECRET, "legendary");
  assert.deepEqual(Array.from(RW.MILESTONES, (m) => m.at), [1000, 10000, 50000, 67000, 100000]);
  assert.equal(RW.next(50000).id, "w67k"); assert.equal(RW.next(50000).cutscene, "67", "67,000 words has its own cutscene");
  assert.equal(RW.next(9999).id, "w10k"); assert.equal(RW.passed(10000).length, 2); assert.equal(RW.next(1e6), null);
});

// the background script in a stand-in browser: storage with promises, the dice replaceable
function background(dice) {
  const store = { local: {}, sync: {} }, listeners = [];
  const area = (a) => ({
    get: async (k) => { const keys = typeof k === "string" ? [k] : Array.isArray(k) ? k : Object.keys(k || {}); const o = {}; for (const x of keys) if (x in store[a]) o[x] = JSON.parse(JSON.stringify(store[a][x])); return o; },
    set: async (o) => { Object.assign(store[a], JSON.parse(JSON.stringify(o))); },
  });
  const ctx = { console, setTimeout, clearTimeout, URL, JSON, fetch: async () => ({ ok: false }), crypto: require("crypto").webcrypto,
    chrome: { runtime: { onMessage: { addListener: (f) => listeners.push(f) }, onInstalled: { addListener() {} }, onStartup: { addListener() {} }, getURL: (p) => p },
      storage: { local: area("local"), sync: area("sync"), onChanged: { addListener() {} } }, scripting: { insertCSS() {} },
      alarms: { get: (n, cb) => cb({}), create() {}, onAlarm: { addListener() {} } }, tabs: { create() {} } } };
  ctx.self = ctx; if (dice) ctx.wtRandom = dice; vm.createContext(ctx);
  for (const s of ["lib/themes.js", "lib/catalog.js", "lib/rewards.js", "background.js"]) vm.runInContext(fs.readFileSync(path.join(EXT, s), "utf8"), ctx, { filename: s });
  const send = (m) => new Promise((ok) => { const r = listeners[0](m, { tab: { id: 1 }, frameId: 0 }, ok); if (r !== true) setTimeout(() => ok(undefined), 0); });
  return { store, send, ctx };
}

test("background: words add up per theme, one update at a time; bad reports are refused", async () => {
  const b = background(() => 0.5);                                   // dice that never win
  const r = await Promise.all([b.send({ words: 5, theme: "synthwave" }), b.send({ words: 3, theme: "winter" }), b.send({ words: 2, theme: "synthwave" })]);
  assert.deepEqual(b.store.local.words, { synthwave: 7, winter: 3 }, "reports from several tabs at once all count");
  assert.deepEqual(r.map((x) => x.total), [5, 8, 10]);
  assert.ok(!b.store.sync.found, "no unlock with losing dice");
  for (const bad of [{ words: -4, theme: "winter" }, { words: 9999, theme: "winter" }, { words: 2, theme: "../x" }, { words: 2, theme: 5 }])
    assert.equal(await b.send(bad), null, JSON.stringify(bad));
  assert.deepEqual(b.store.local.words, { synthwave: 7, winter: 3 });
});

test("background: one roll per word at 1 in 37,500; found once, kept in sync storage; milestones", async () => {
  const rolls = []; let win = 13;
  const b = background(() => { rolls.push(1); return --win === 0 ? 1 / RW.ODDS - 1e-12 : 1 / RW.ODDS; });   // exactly 1/ODDS loses, just under wins
  let r = await b.send({ words: 10, theme: "synthwave" });
  assert.equal(rolls.length, 10, "ten words, ten rolls"); assert.ok(!r.secret);
  r = await b.send({ words: 6, theme: "winter" });
  assert.equal(r.secret, "legendary", "the 13th word wins"); assert.equal(r.total, 13, "found on the 13th word");
  assert.deepEqual({ ...b.store.sync.found.legendary, at: 0 }, { at: 0, words: 13, theme: "winter" });
  assert.equal(b.store.local.words.winter, 6, "the rest of the report still counts");
  const before = rolls.length;
  r = await b.send({ words: 490, theme: "winter" });
  assert.ok(!r.secret && !r.badge); assert.equal(r.total, 506);
  r = await b.send({ words: 500, theme: "winter" });
  assert.equal(rolls.length, before, "no more rolls once it's found"); assert.ok(!r.secret);
  assert.equal(r.total, 1006); assert.equal(r.badge, "w1k", "passing 1,000 words earns the first milestone");
  r = await b.send({ words: 5, theme: "winter" }); assert.ok(!r.badge, "a milestone is earned once");
  // the real dice: the chance per word is 1 in ODDS (a million rolls of the real generator land near 1,000,000 / 37,500 = 26.7)
  const real = background(null), dice = real.ctx.crypto;
  let wins = 0; const u = new Uint32Array(1);
  for (let i = 0; i < 1e6; i++) { dice.getRandomValues(u); if (u[0] / 4294967296 < 1 / RW.ODDS) wins++; }
  assert.ok(wins > 8 && wins < 50, "wins in a million: " + wins);
});

test("background: the 67,000th word earns the six-seven milestone", async () => {
  const b = background(() => 0.5);
  b.store.local.words = { synthwave: 50000, winter: 16990 };
  let r = await b.send({ words: 9, theme: "winter" }); assert.ok(!r.badge); assert.equal(r.total, 66999);
  r = await b.send({ words: 1, theme: "winter" }); assert.equal(r.badge, "w67k"); assert.equal(r.total, 67000);
});

test("the secret theme isn't on the website, in the File studio or in the live catalogue; it ships inside the extension", () => {
  const cat = JSON.parse(fs.readFileSync(path.join(ROOT, "catalog.json"), "utf8"));
  assert.ok(H.THEMES.legendary && H.THEMES.legendary.secret, "in the extension's theme list, marked secret");
  assert.ok(!cat.themes.some((t) => t.id === "legendary"), "not in catalog.json");
  assert.match(fs.readFileSync(path.join(ROOT, "site.js"), "utf8"), /if \(!t\.secret\) themes\.push/);
  assert.match(fs.readFileSync(path.join(ROOT, "studio.js"), "utf8"), /filter\(\(\[, t\]\) => !t\.secret\)/);
  // the desktop app's tray menu reads entries that start with "name": the secret one starts with "secret"
  assert.match(fs.readFileSync(path.join(EXT, "lib", "themes.js"), "utf8"), /"legendary": \{\s*"secret": true,\s*"name"/);
});

const WORD = '<!doctype html><html><body style="margin:0;background:#e6e6e6"><div id="canvas" style="position:absolute;inset:0;background:#e6e6e6"><div id="page" contenteditable="true" style="width:600px;height:800px;margin:40px auto;background:#fff;font:16px Georgia;padding:20px"></div></div></body></html>';
const DOCS = `<!doctype html><html><body style="margin:0;background:#f9fbfd"><div id="docs-chrome" style="height:84px">File</div>
<div id="docs-editor-container" style="position:absolute;top:84px;bottom:0;left:0;right:0;background:#f9fbfd"><div id="docs-editor" style="position:absolute;inset:0"><div class="kix-appview-editor" style="position:absolute;inset:0;overflow:auto;background-color: rgb(249, 251, 253);"><div class="kix-page-paginated" style="position:relative;width:816px;height:1056px;margin:20px auto;background:#fff"></div></div></div></div>
<iframe class="docs-texteventtarget-iframe" style="position:absolute;left:-2000px;top:0;width:20px;height:20px" srcdoc="<body contenteditable=true></body>"></iframe></body></html>`;

test("in the real extension: typing counts, paste doesn't; the locked card; the unlock cutscene; milestones", { skip: !chromium && "playwright not installed", timeout: 180000 }, async (t) => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "wt-rw-"));
  const ctx = await chromium.launchPersistentContext(dir, { headless: false, args: ["--headless=new", `--disable-extensions-except=${EXT}`, `--load-extension=${EXT}`], permissions: ["clipboard-read", "clipboard-write"] });
  try {
    await ctx.route("https://min3scon.github.io/**", (r) => r.fulfill({ status: 404, body: "" }));
    await ctx.route("https://word-edit.officeapps.live.com/**", (r) => r.fulfill({ contentType: "text/html", body: WORD }));
    await ctx.route("https://docs.google.com/**", (r) => r.fulfill({ contentType: "text/html", body: DOCS }));
    let [sw] = ctx.serviceWorkers(); if (!sw) sw = await ctx.waitForEvent("serviceworker");
    const extId = sw.url().split("/")[2];
    const local = (k) => sw.evaluate((k) => chrome.storage.local.get(k), k).then((v) => v[k]);
    const sync = (k) => sw.evaluate((k) => chrome.storage.sync.get(k), k).then((v) => v[k]);
    const word = async () => { const pg = await ctx.newPage(); await pg.goto("https://word-edit.officeapps.live.com/we/wordeditorframe.aspx?WOPISrc=https%3A%2F%2Fexample%2Fwopi%2Ffiles%2F1"); return pg; };
    const until = async (fn, ms = 10000) => { const t0 = Date.now(); for (;;) { const v = await fn(); if (v || Date.now() - t0 > ms) return v; await new Promise((r) => setTimeout(r, 250)); } };
    await sw.evaluate(() => chrome.storage.sync.set({ theme: "legendary", on: true, motion: "on", sparks: false, docs: {}, found: {} }));
    await sw.evaluate(() => chrome.storage.local.remove(["words", "revealed"]));

    await t.test("a secret theme that hasn't been found can't be picked, even by setting it directly", async () => {
      const pg = await word();
      await pg.waitForSelector("#ww-bg", { state: "attached", timeout: 15000 });
      assert.equal(await pg.$eval("#ww-bg", (n) => n.dataset.scene), "synthwave", "the usual default shows instead");
      await pg.close();
      await sw.evaluate(() => chrome.storage.sync.set({ theme: "synthwave" }));
    });

    await t.test("typed words count for the theme they were typed in; pasted and scripted text doesn't", async () => {
      const pg = await word();
      await pg.waitForSelector('#ww-bg[data-scene="synthwave"]', { state: "attached", timeout: 15000 });
      await pg.click("#page");
      await pg.keyboard.type("hello world, don't stop typing. ");                       // 5 words
      await pg.evaluate(() => navigator.clipboard.writeText("pasted words should never count at all "));
      await pg.keyboard.press("Control+V");                                              // paste works...
      await pg.waitForFunction(() => document.querySelector("#page").textContent.includes("pasted words"), null, { timeout: 5000 });
      await pg.evaluate(() => document.execCommand("insertText", false, "dropped text also never counts "));   // ...and so does inserting text
      await pg.evaluate(() => { for (const key of [..."fake words "]) document.querySelector("#page").dispatchEvent(new KeyboardEvent("keydown", { key, bubbles: true })); });
      assert.ok(await until(async () => (await local("words") || {}).synthwave === 5, 9000), "5 words: " + JSON.stringify(await local("words")));
      await pg.waitForTimeout(4500);
      assert.deepEqual(await local("words"), { synthwave: 5 }, "nothing else was counted");
      assert.match(await pg.$eval("#page", (n) => n.textContent.replace(/\s/g, " ")), /hello world, don't stop typing\. pasted words should never count at all dropped text/, "paste still works in the document");
      await pg.close();
    });

    await t.test("Google Docs: words typed in its hidden typing frame count too", async () => {
      const pg = await ctx.newPage(); await pg.goto("https://docs.google.com/document/d/DOCW/edit");
      await pg.waitForSelector('#ww-bg[data-scene="synthwave"]', { state: "attached", timeout: 15000 });
      await pg.waitForTimeout(800);
      await pg.frames().find((f) => f !== pg.mainFrame()).focus("body");
      await pg.keyboard.type("three more words ");
      assert.ok(await until(async () => (await local("words") || {}).synthwave === 8, 9000), JSON.stringify(await local("words")));
      await pg.close();
    });

    await t.test("popup: words on each card, the total, and the secret theme locked: greyed out, no picture, no name", async () => {
      const pop = await ctx.newPage(); await pop.goto(`chrome-extension://${extId}/popup.html`);
      await pop.waitForSelector('.card[data-id="synthwave"] .wc');
      assert.equal(await pop.$eval('.card[data-id="synthwave"] .wc', (n) => n.textContent), "8");
      assert.equal(await pop.$eval("#wtotal", (n) => n.textContent), "8");
      assert.match(await pop.$eval("#wnote", (n) => n.textContent), /992 more to 🔥 Warmed up/);
      const lock = await pop.$eval(".card.locked", (n) => ({ img: !!n.querySelector("img"), text: n.textContent, grey: getComputedStyle(n).filter, title: n.title }));
      assert.deepEqual({ img: lock.img, text: lock.text, grey: lock.grey }, { img: false, text: "????", grey: "grayscale(1)" }, "no preview, no name");
      assert.match(lock.title, /1 in 37,500/);
      assert.equal(await pop.$('.card[data-id="legendary"]'), null, "not listed as a theme");
      assert.ok(!/Legendary/.test(await pop.content()), "its name isn't anywhere in the popup");
      await pop.click(".card.locked", { force: true });                  // marked aria-disabled for screen readers; a click still explains it
      assert.match(await pop.$eval("#wnote", (n) => n.textContent), /secret theme/i);
      assert.equal(await sync("theme"), "synthwave", "clicking it picks nothing");
      await pop.screenshot({ path: path.join(H.OUT, "popup-locked.png") });
      await pop.close();
    });

    await t.test("the lucky word: the cutscene plays, Use it now switches to the secret theme", async () => {
      await sw.evaluate(() => { self.wtRandom = () => 0; });         // the next roll wins
      const pg = await word();
      await pg.waitForSelector('#ww-bg[data-scene="synthwave"]', { state: "attached", timeout: 15000 });
      await pg.click("#page");
      await pg.keyboard.type("lucky ");
      await pg.waitForSelector('#ww-cut[data-kind="secret"]', { state: "attached", timeout: 9000 });
      assert.equal(await pg.$eval("#ww-cut .wwc-title", (n) => n.textContent), "Legendary");
      assert.match(await pg.$eval("#ww-cut .wwc-sub", (n) => n.textContent), /1 in 37,500 .* after 9 words/);
      const f = await sync("found"); assert.equal(f.legendary.words, 9);
      await pg.keyboard.type("still typing ");                         // the cutscene doesn't take the keyboard
      assert.match(await pg.$eval("#page", (n) => n.textContent), /lucky still typing/);
      await pg.waitForTimeout(4600);                                    // the whole cutscene
      await pg.screenshot({ path: path.join(H.OUT, "cutscene-secret.png") });
      await pg.click("#ww-cut .wwc-go");
      await pg.waitForSelector('#ww-bg[data-scene="legendary"]', { state: "attached", timeout: 8000 });
      await pg.waitForSelector("#ww-cut", { state: "detached", timeout: 3000 });
      assert.equal(await sync("theme"), "legendary");
      await pg.waitForTimeout(400);
      await pg.screenshot({ path: path.join(H.OUT, "legendary-word.png") });
      await pg.close();
    });

    await t.test("popup after the unlock: the card turns over once, with its picture; Watch it again replays the cutscene", async () => {
      let pop = await ctx.newPage(); await pop.goto(`chrome-extension://${extId}/popup.html`);
      await pop.waitForSelector('.card[data-id="legendary"]');
      assert.equal(await pop.$(".card.locked"), null);
      const c = await pop.$eval('.card[data-id="legendary"]', (n) => ({ cls: n.className, img: n.querySelector("img").getAttribute("src") }));
      assert.match(c.cls, /legend/); assert.match(c.cls, /reveal/, "the reveal plays the first time"); assert.equal(c.img, "previews/legendary.webp");
      assert.match(await pop.$eval("#wfound", (n) => (n.hidden ? "" : n.textContent)), /Legendary found after 9 words/);
      await pop.close();
      pop = await ctx.newPage(); await pop.goto(`chrome-extension://${extId}/popup.html`);
      await pop.waitForSelector('.card[data-id="legendary"]');
      assert.doesNotMatch(await pop.$eval('.card[data-id="legendary"]', (n) => n.className), /reveal/, "only once");
      await pop.close();
      const pg = await word();
      await pg.waitForSelector('#ww-bg[data-scene="legendary"]', { state: "attached", timeout: 15000 });
      const tabId = await sw.evaluate(async () => (await chrome.tabs.query({ url: "https://word-edit.officeapps.live.com/*" }))[0].id);
      await sw.evaluate((id) => chrome.tabs.sendMessage(id, { cutscene: "secret" }).catch(() => {}), tabId);   // no answer is sent back
      await pg.waitForSelector('#ww-cut[data-kind="secret"]', { state: "attached", timeout: 5000 });
      await pg.keyboard.press("Escape"); await pg.keyboard.press("Escape");   // skip to the end, then close
      await pg.waitForSelector("#ww-cut", { state: "detached", timeout: 3000 });
      await pg.close();
    });

    await t.test("a word milestone plays its own short cutscene, which goes by itself", async () => {
      await sw.evaluate(() => chrome.storage.local.set({ words: { legendary: 998 } }));
      const pg = await word();
      await pg.waitForSelector('#ww-bg[data-scene="legendary"]', { state: "attached", timeout: 15000 });
      await pg.click("#page");
      await pg.keyboard.type("one two ");
      await pg.waitForSelector('#ww-cut[data-kind="badge"]', { state: "attached", timeout: 9000 });
      assert.match(await pg.$eval("#ww-cut", (n) => n.textContent), /Warmed up.*1,000 words typed/);
      assert.equal(await pg.$eval("#ww-cut", (n) => getComputedStyle(n).pointerEvents), "none", "doesn't block the page");
      await pg.waitForTimeout(1500);
      await pg.screenshot({ path: path.join(H.OUT, "cutscene-milestone.png") });
      await pg.waitForSelector("#ww-cut", { state: "detached", timeout: 8000 });
      await pg.close();
    });

    await t.test("67,000 words: 67 copies of \"67\" all over the screen, the 6 and the 7 bobbing, gone by itself", async () => {
      await sw.evaluate(() => chrome.storage.local.set({ words: { legendary: 66998 } }));
      const pg = await word();
      await pg.waitForSelector('#ww-bg[data-scene="legendary"]', { state: "attached", timeout: 15000 });
      await pg.click("#page");
      await pg.keyboard.type("six seven ");
      await pg.waitForSelector('#ww-cut[data-kind="67"]', { state: "attached", timeout: 9000 });
      const v = await pg.$eval("#ww-cut", (n) => ({ spam: [...n.querySelectorAll(".wwc-67")].map((s) => s.textContent), big: n.querySelector(".wwc-67big").textContent,
        text: n.querySelector(".wwc-ribbon").textContent, pe: getComputedStyle(n).pointerEvents }));
      assert.equal(v.spam.length, 67, "sixty-seven of them"); assert.ok(v.spam.every((s) => s === "67"));
      assert.equal(v.big, "6🤲7"); assert.match(v.text, /SIX SEVEN.*67,000 words typed/);
      assert.equal(v.pe, "none", "doesn't block the page");
      await pg.waitForTimeout(2600);                                    // all 67 are out
      const spread = await pg.$$eval("#ww-cut .wwc-67", (a) => a.map((s) => { const r = s.getBoundingClientRect(); return [r.left + r.width / 2 < innerWidth / 2, r.top + r.height / 2 < innerHeight / 2, +getComputedStyle(s).opacity]; }));
      assert.ok(spread.every((s) => s[2] > .9), "all visible");
      for (const q of [[true, true], [false, true], [true, false], [false, false]]) assert.ok(spread.some((s) => s[0] === q[0] && s[1] === q[1]), "in every corner of the screen");
      const bob = await pg.$eval("#ww-cut .wwc-six", (n) => n.getAnimations().map((a) => a.playState));
      assert.deepEqual(bob, ["running"], "the 6 and the 7 take turns going up and down");
      await pg.keyboard.type("still typing ");
      assert.match(await pg.$eval("#page", (n) => n.textContent), /six seven still typing/, "typing carries on underneath");
      await pg.screenshot({ path: path.join(H.OUT, "cutscene-67.png") });
      await pg.waitForSelector("#ww-cut", { state: "detached", timeout: 9000 });
      await pg.close();
    });
  } finally { await ctx.close(); }
});
