// Drives the Office add-in in Chromium against a stand-in office.js, for Word, PowerPoint and Excel.
// The add-in's pane is the website itself (studio.html / index.html with ?host=office), so this runs
// the real pages: recognise, apply, change, configure, remove, undo, repair, copy, and the gallery.
const test = require("node:test"), assert = require("node:assert/strict");
const http = require("http"), fs = require("fs"), path = require("path");
const H = require("./helpers.js");
let chromium; try { ({ chromium } = require("playwright")); } catch (_) { /* skipped below */ }

const ROOT = path.join(__dirname, "..");
const TYPES = { ".html": "text/html", ".js": "text/javascript", ".css": "text/css", ".webp": "image/webp", ".gif": "image/gif", ".png": "image/png", ".jpg": "image/jpeg", ".json": "application/json" };
function serve() {
  return new Promise((ok) => {
    const srv = http.createServer((req, res) => {
      const p = path.join(ROOT, decodeURIComponent(new URL(req.url, "http://x").pathname));
      if (!p.startsWith(ROOT) || !fs.existsSync(p) || fs.statSync(p).isDirectory()) { res.writeHead(404); return res.end(); }
      res.writeHead(200, { "Content-Type": TYPES[path.extname(p)] || "application/octet-stream" }); fs.createReadStream(p).pipe(res);
    }).listen(0, () => ok(srv));
  });
}

async function pane(browser, port, host, bytes, page_ = "studio.html") {
  const page = await browser.newPage({ viewport: { width: 340, height: 900 } });
  const errors = []; page.on("pageerror", (e) => errors.push(e.message));
  await page.route("https://appsforoffice.microsoft.com/**", (r) => r.fulfill({ contentType: "text/javascript", body: fs.readFileSync(path.join(__dirname, "addin-mock-office.js"), "utf8") }));
  await page.route("https://fonts.googleapis.com/**", (r) => r.fulfill({ contentType: "text/css", body: "" }));
  await page.addInitScript(([h, d]) => { window.__host = h; window.__doc = new Uint8Array(d); }, [host, Array.from(bytes)]);
  await page.goto(`http://localhost:${port}/${page_}?host=office`);
  if (page_ === "studio.html" && host !== "none") await page.waitForFunction(() => { const t = document.querySelector("#fname").textContent; return t && !/Reading/.test(t); }, null, { timeout: 30000 });
  return { page, errors };
}
const doc = (page) => page.evaluate(() => Array.from(window.__doc)).then((a) => new Uint8Array(a));
const msgDone = (page, sel = "#msg") => page.waitForFunction((s) => /Checked|opened in a new window|only took part|Undone|Couldn|can't/.test(document.querySelector(s).textContent), sel, { timeout: 60000 });

test("Office add-in (the website in Word, PowerPoint and Excel) with stand-in Office", { skip: !chromium && "playwright not installed" }, async (t) => {
  const srv = await serve(), port = srv.address().port, browser = await chromium.launch();
  try {
    await t.test("Word: recognise, change theme, configure, remove, undo, repair", async () => {
      const themed = await H.WT.apply(H.fixture("report.docx"), H.theme("synthwave"), {}, H.makeAssets, "test");
      const { page, errors } = await pane(browser, port, "Word", themed);
      assert.equal(await page.textContent("#fname"), "Synthwave Sunset");
      assert.equal(await page.textContent("#badge"), "Intact");
      assert.ok(await page.isHidden("#drop"), "the open document is the file: no file picker");
      assert.ok(await page.isVisible("#verify"));
      assert.equal(await page.textContent("#download"), "Open a themed copy");
      await page.screenshot({ path: path.join(H.OUT, "addin-word-open.png"), fullPage: true });

      await page.click(".st-card >> nth=9");
      await msgDone(page);
      assert.match(await page.textContent("#msg"), /Checked: the theme is in this file and intact/);
      const after = await H.WT.inspect(await doc(page));
      assert.equal(after.status, "intact"); assert.notEqual(after.theme.id, "synthwave");
      const calls = await page.evaluate(() => window.__calls);
      assert.ok(calls.some((c) => c[0] === "insertFile" && c[1] === "Replace" && c[2].importTheme && c[2].importStyles && c[2].importPageColor));
      assert.ok(calls.some((c) => c[0] === "xml.add"), "live custom XML part kept in step");

      await page.click('#look label.switch:has-text("Heading glow") input');
      await page.click("#applyLook");
      await msgDone(page);
      assert.equal((await H.WT.inspect(await doc(page))).settings.glow, true);

      await page.click("#remove");
      await msgDone(page);
      assert.equal(await page.textContent("#fname"), "No theme in this Word document");
      assert.equal((await H.WT.inspect(await doc(page))).status, "none");
      assert.ok(await page.isVisible("#undo"));
      await page.click("#undo");
      await msgDone(page);
      assert.equal((await H.WT.inspect(await doc(page))).status, "intact");

      // someone changed the theme colours by hand -> Restore theme repairs it from the file itself
      const z = await H.zipOf(await doc(page)), tp = "word/theme/theme1.xml";
      z.file(tp, (await z.file(tp).async("string")).replace(/(<a:accent1>\s*<a:srgbClr val=")[0-9A-Fa-f]{6}/, "$1123456"));
      const tampered = await z.generateAsync({ type: "uint8array" });
      await page.evaluate((d) => { window.__doc = new Uint8Array(d); }, Array.from(tampered));
      await page.click("#verify");
      await page.waitForFunction(() => document.querySelector("#badge").textContent === "Changed");
      assert.ok(await page.isVisible("#repair"));
      await page.click("#repair");
      await msgDone(page);
      assert.equal((await H.WT.inspect(await doc(page))).status, "intact");

      await page.click("#download");
      await msgDone(page);
      assert.equal((await H.WT.inspect(new Uint8Array(await page.evaluate(() => Array.from(window.__copy))))).status, "intact", "themed copy opens as a new document");
      assert.deepEqual(errors, []);
      await page.close();
    });

    await t.test("PowerPoint: apply in place with transition setting", async () => {
      const { page, errors } = await pane(browser, port, "PowerPoint", H.fixture("deck.pptx"));
      assert.equal(await page.textContent("#fname"), "No theme in this PowerPoint deck");
      await page.selectOption("#wt-set-transition", "push");
      await page.click(".st-card >> nth=0");
      await msgDone(page);
      const i = await H.WT.inspect(await doc(page));
      assert.equal(i.status, "intact"); assert.equal(i.settings.transition, "push");
      assert.ok((await page.evaluate(() => window.__calls)).some((c) => c[0] === "insertSlides" && c[1].formatting === "KeepSourceFormatting"));
      await page.screenshot({ path: path.join(H.OUT, "addin-ppt.png") });
      assert.deepEqual(errors, []);
      await page.close();
    });

    await t.test("Excel: themed workbook opens as a copy, original untouched", async () => {
      const orig = H.fixture("book.xlsx");
      const { page, errors } = await pane(browser, port, "Excel", orig);
      await page.click(".st-card >> nth=2");
      await msgDone(page);
      assert.match(await page.textContent("#msg"), /new window/);
      const copy = new Uint8Array(await page.evaluate(() => Array.from(window.__copy)));
      assert.equal((await H.WT.inspect(copy)).status, "intact");
      assert.deepEqual(await doc(page), orig, "open workbook not modified");
      await page.screenshot({ path: path.join(H.OUT, "addin-excel.png") });
      assert.deepEqual(errors, []);
      await page.close();
    });

    await t.test("Gallery in Word: a theme's dialog saves it into the open document", async () => {
      const { page, errors } = await pane(browser, port, "Word", H.fixture("plain.docx"), "index.html");
      await page.waitForFunction(() => document.documentElement.dataset.office === "word");
      await page.click(".card .shot >> nth=0");
      assert.match(await page.textContent("dialog .drop b"), /Save it into this Word document/);
      await page.click("dialog .drop");
      await msgDone(page, "dialog .msg");
      assert.match(await page.textContent("dialog .msg"), /Checked: the theme is in this file and intact/);
      const i = await H.WT.inspect(await doc(page));
      assert.equal(i.status, "intact"); assert.equal(i.tool, "Wild Themes add-in");
      await page.screenshot({ path: path.join(H.OUT, "addin-gallery.png") });
      assert.deepEqual(errors, []);
      await page.close();
    });

    await t.test("?host=office in a normal browser: the ordinary File studio", async () => {
      const { page, errors } = await pane(browser, port, "none", new Uint8Array(0));
      await page.waitForSelector("#drop", { state: "visible" });
      assert.equal(await page.evaluate(() => sessionStorage.getItem("wt-host")), null);
      assert.deepEqual(errors, []);
      await page.close();
    });
  } finally { await browser.close(); srv.close(); }
});
