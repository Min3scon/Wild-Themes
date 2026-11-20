// Every theme is complete (scene, picture, backdrops, Teams background, document palette, entry in the catalogue), and the
// newest one, LARP Coder (the word "larpcoder" repeated and animated), works in the real extension and in files.
const test = require("node:test"), assert = require("node:assert/strict");
const fs = require("fs"), path = require("path"), os = require("os");
process.env.PW_EXPERIMENTAL_SERVICE_WORKER_NETWORK_EVENTS = "1";
const H = require("./helpers.js");
let chromium; try { ({ chromium } = require("playwright")); } catch (_) { /* skipped below */ }

const ROOT = path.join(__dirname, ".."), EXT = path.join(ROOT, "extension");
const hex6 = /^[0-9A-Fa-f]{6}$/;

test("every theme has all its parts", () => {
  const ids = Object.keys(H.THEMES);
  assert.ok(ids.length >= 19 && ids.includes("larpcoder"));
  for (const id of ids) {
    const t = H.THEMES[id], d = t.doc;
    assert.ok(t.name && t.blurb && t.emoji, id + ": name, blurb, emoji");
    assert.ok(["dark", "light"].includes(t.mode), id + ": mode");
    assert.match(t.accent, /^#[0-9a-f]{6}$/i, id + ": accent"); assert.match(t.accent2, /^#[0-9a-f]{6}$/i, id + ": accent2");
    for (const k of ["dk1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"]) assert.match(d[k], hex6, `${id}: doc.${k}`);
    assert.ok(d.major && d.minor && /^#[0-9a-f]{6}$/i.test(d.panel), id + ": fonts and panel");
    for (const f of [`scenes/${id}.css`, `previews/${id}.webp`, `backdrops/${id}.gif`, `backdrops-wide/${id}.gif`, `teams-backgrounds/${id}.jpg`]) assert.ok(fs.existsSync(path.join(EXT, f)), `${id}: ${f}`);
    const css = fs.readFileSync(path.join(EXT, "scenes", `${id}.css`), "utf8");
    assert.ok(css.includes(`[data-scene=${id}]`), id + ": scene is keyed by its id");
    assert.ok(!/url\s*\(|@import|expression\s*\(|javascript:/i.test(css), id + ": scene loads nothing (works under any page policy)");
  }
  const cat = JSON.parse(fs.readFileSync(path.join(ROOT, "catalog.json"), "utf8"));
  assert.deepEqual(cat.themes.filter((t) => t.kind === "builtin").map((t) => t.id).sort(), ids.slice().sort(), "committed catalog.json lists every theme (node scripts/build-catalog.js)");
});

test("LARP Coder: the word repeated, files sized for Office, applies to Word, PowerPoint and Excel files", async () => {
  const css = fs.readFileSync(path.join(EXT, "scenes", "larpcoder.css"), "utf8");
  assert.ok((css.match(/larpcoder/g) || []).length > 500, "the word repeated hundreds of times");
  const kb = (f) => fs.statSync(path.join(EXT, f)).size / 1024;
  assert.ok(kb("backdrops/larpcoder.gif") < 400 && kb("backdrops-wide/larpcoder.gif") < 600 && kb("previews/larpcoder.webp") < 20 && kb("teams-backgrounds/larpcoder.jpg") < 600, "picture sizes");
  assert.ok(kb("scenes/larpcoder.css") < 12, "scene stays small");
  for (const name of ["report.docx", "deck.pptx", "book.xlsx"]) {
    const out = await H.WT.apply(H.fixture(name), H.theme("larpcoder"), {}, H.makeAssets, "test");
    const i = await H.WT.inspect(out);
    assert.equal(i.status, "intact", name + ": " + JSON.stringify(i.problems)); assert.equal(i.theme.id, "larpcoder");
    const pal = H.WT.palette(H.theme("larpcoder")), scheme = await H.WT.schemeOf(await H.text(out, { docx: "word/theme/theme1.xml", pptx: "ppt/theme/theme1.xml", xlsx: "xl/theme/theme1.xml" }[H.kindOfName(name)]));
    assert.equal(scheme.accent1, pal.accent1); assert.equal(scheme.major, "Consolas");
  }
});

test("LARP Coder in the real extension: scene shows, three layers move, none is heavy, sparks use code glyphs", { skip: !chromium && "playwright not installed", timeout: 90000 }, async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "wt-th-"));
  const ctx = await chromium.launchPersistentContext(dir, { headless: false, args: ["--headless=new", `--disable-extensions-except=${EXT}`, `--load-extension=${EXT}`] });
  try {
    await ctx.route("https://min3scon.github.io/**", (r) => r.fulfill({ status: 404, body: "" }));
    await ctx.route("https://word-edit.officeapps.live.com/**", (r) => r.fulfill({ contentType: "text/html",
      body: '<!doctype html><html><body style="margin:0;background:#e6e6e6"><div id="canvas" style="position:absolute;inset:0;background:#e6e6e6"><div id="page" contenteditable="true" style="width:600px;height:800px;margin:40px auto;background:#fff">Hello</div></div></body></html>' }));
    let [sw] = ctx.serviceWorkers(); if (!sw) sw = await ctx.waitForEvent("serviceworker");
    await sw.evaluate(() => chrome.storage.sync.set({ theme: "larpcoder", on: true, motion: "on", sparks: true, docs: {} }));
    const pg = await ctx.newPage();
    await pg.goto("https://word-edit.officeapps.live.com/we/wordeditorframe.aspx?WOPISrc=https%3A%2F%2Fexample%2Fwopi%2Ffiles%2F1");
    await pg.waitForSelector('#ww-bg[data-scene="larpcoder"]', { state: "attached", timeout: 15000 });
    await pg.waitForFunction(() => document.querySelectorAll("#ww-bg>i").length === 4 && [...document.querySelectorAll("#ww-bg>i")].filter((i) => i.getAnimations().length).length === 3, null, { timeout: 10000 });
    const layers = await pg.evaluate(() => [...document.querySelectorAll("#ww-bg>i")].map((i) => ({ names: i.getAnimations().map((a) => a.animationName), state: i.getAnimations().map((a) => a.playState), area: i.offsetWidth * i.offsetHeight / (innerWidth * innerHeight) })));
    assert.deepEqual(layers.filter((l) => l.names.length).map((l) => l.names[0]).sort(), ["lc-left", "lc-up", "ww-move"]);
    assert.ok(layers.every((l) => l.state.every((s) => s === "running")), "all animations running");
    assert.ok(layers.every((l) => l.area < 2.2), "no layer is heavy: " + JSON.stringify(layers.map((l) => +l.area.toFixed(2))));
    const text = await pg.evaluate(() => getComputedStyle(document.querySelector("#ww-bg>i:nth-child(2)"), "::before").content.length);
    assert.ok(text > 2000, "text layer holds the repeated word");
    await pg.waitForTimeout(500);
    await pg.screenshot({ path: path.join(H.OUT, "larpcoder-word.png") });
    await pg.click("#page");                                         // the sparks start where the pointer or caret is
    await pg.keyboard.press("a");
    await pg.waitForSelector('#ww-fx[data-fx="glyph"]', { state: "attached", timeout: 5000 });
  } finally { await ctx.close(); }
});
