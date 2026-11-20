// The website's "Add to browser" dialog: picks the steps for the browser in use, every browser button shows its
// own address and steps, the Firefox one-click install appears only once a signed file is published, and every
// download it links to is a file the build really makes.
const test = require("node:test"), assert = require("node:assert/strict");
const http = require("http"), fs = require("fs"), path = require("path"), os = require("os"), { execFileSync } = require("child_process");
const H = require("./helpers.js");
let chromium; try { ({ chromium } = require("playwright")); } catch (_) { /* skipped below */ }

const ROOT = path.join(__dirname, "..");
const TYPES = { ".html": "text/html", ".js": "text/javascript", ".css": "text/css", ".webp": "image/webp", ".json": "application/json", ".png": "image/png", ".jpg": "image/jpeg", ".gif": "image/gif" };
function serve(signed) {
  return new Promise((ok) => {
    const srv = http.createServer((req, res) => {
      const u = decodeURIComponent(new URL(req.url, "http://x").pathname);
      if (u === "/downloads/wild-themes-firefox.xpi") { res.writeHead(signed ? 200 : 404, { "Content-Type": "application/x-xpinstall" }); return res.end(signed ? "x" : ""); }
      const p = path.join(ROOT, u === "/" ? "index.html" : u);
      if (!p.startsWith(ROOT) || !fs.existsSync(p) || fs.statSync(p).isDirectory()) { res.writeHead(404); return res.end(); }
      res.writeHead(200, { "Content-Type": TYPES[path.extname(p)] || "application/octet-stream" }); fs.createReadStream(p).pipe(res);
    }).listen(0, () => ok(srv));
  });
}
const UA = {
  chrome: "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36",
  edge: "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36 Edg/140.0.0.0",
  opera: "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36 OPR/120.0.0.0",
  firefox: "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:140.0) Gecko/20100101 Firefox/140.0",
  safari: "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/18.0 Safari/605.1.15",
  phone: "Mozilla/5.0 (Linux; Android 14; Pixel 8) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Mobile Safari/537.36",
};
const state = (page) => page.evaluate(() => ({
  shown: [...document.querySelectorAll("#dlg [data-browser]")].filter((d) => !d.hidden).map((d) => d.dataset.browser),
  pressed: [...document.querySelectorAll("#dlg [data-b]")].filter((b) => b.getAttribute("aria-pressed") === "true").map((b) => b.dataset.b),
  addr: document.querySelector("#dlg [data-addr]").textContent, copy: document.querySelector("#dlg [data-addr]").parentElement.querySelector(".copy").dataset.copy,
  dev: document.querySelector("#dlg [data-dev]").textContent, name: document.querySelector("#dlg [data-bname]").textContent,
  signedShown: !document.querySelector("#dlg [data-signed]").hidden,
}));
async function open(browser, port, ua) {
  const ctx = await browser.newContext({ userAgent: ua, viewport: { width: 1100, height: 900 } });
  await ctx.route(/fonts\.(googleapis|gstatic)/, (r) => r.abort());
  const page = await ctx.newPage(), errs = []; page.errs = errs; page.on("pageerror", (e) => errs.push(e.message));
  await page.goto(`http://127.0.0.1:${port}/`);
  await page.click('.nav [data-open="install"]'); await page.waitForSelector("#dlg[open]");
  return page;
}

test("site install dialog: steps for the browser in use, one per browser, signed Firefox install when published", { skip: !chromium && "playwright not installed" }, async (t) => {
  const srv = await serve(false), srv2 = await serve(true), browser = await chromium.launch();
  try {
    await t.test("each browser is detected and shows its own address and steps", async () => {
      const want = { chrome: ["chromium", "chrome://extensions"], edge: ["chromium", "edge://extensions"], opera: ["chromium", "opera://extensions"], firefox: ["firefox"], safari: ["safari"], phone: ["safari"] };
      for (const [name, [kind, addr]] of Object.entries(want)) {
        const page = await open(browser, srv.address().port, UA[name]);
        const s = await state(page);
        assert.deepEqual(s.shown, [kind], name + " pane");
        if (addr) { assert.equal(s.addr, addr, name); assert.equal(s.copy, addr, name + ": copy button copies it"); }
        assert.equal(s.signedShown, false, "no one-click Firefox install before a signed file exists");
        assert.deepEqual(page.errs, [], name);
        await page.context().close();
      }
    });

    await t.test("every browser button works", async () => {
      const page = await open(browser, srv.address().port, UA.chrome);
      const addr = { chrome: "chrome://extensions", edge: "edge://extensions", brave: "brave://extensions", opera: "opera://extensions", vivaldi: "vivaldi://extensions", arc: "arc://extensions" };
      for (const [b, a] of Object.entries(addr)) {
        await page.click(`#dlg [data-b="${b}"]`);
        const s = await state(page);
        assert.deepEqual(s.shown, ["chromium"]); assert.deepEqual(s.pressed, [b]); assert.equal(s.addr, a); assert.equal(s.copy, a);
        assert.match(s.dev, /^Developer mode/); assert.match(s.name, /\S/);
      }
      assert.equal((await (async () => { await page.click('#dlg [data-b="edge"]'); return state(page); })()).dev, "Developer mode (left menu)");
      await page.click('#dlg [data-b="firefox"]'); assert.deepEqual((await state(page)).shown, ["firefox"]);
      await page.click('#dlg [data-b="safari"]'); assert.deepEqual((await state(page)).shown, ["safari"]);
      await page.click('#dlg [data-b="vivaldi"]'); assert.deepEqual((await state(page)).shown, ["chromium"]);
      await page.screenshot({ path: path.join(H.OUT, "site-install-vivaldi.png") });
      await page.click('#dlg [data-b="firefox"]'); await page.screenshot({ path: path.join(H.OUT, "site-install-firefox.png") });
      assert.deepEqual(page.errs, []);
      await page.context().close();
    });

    await t.test("Firefox pane explains the 'has not been verified' error and how to get a signed copy", async () => {
      const page = await open(browser, srv.address().port, UA.firefox);
      const text = await page.textContent('#dlg [data-browser="firefox"]');
      assert.match(text, /This add-on could not be installed because it has not been verified\./);
      assert.match(text, /On your own/);
      assert.match(text, /Load Temporary Add-on/);
      await page.context().close();
    });

    await t.test("the signed Firefox install shows up once downloads/wild-themes-firefox.xpi exists", async () => {
      const page = await open(browser, srv2.address().port, UA.firefox);
      await page.waitForFunction(() => !document.querySelector("#dlg [data-signed]").hidden, null, { timeout: 5000 });
      assert.equal(await page.getAttribute("#dlg [data-signed] a", "href"), "downloads/wild-themes-firefox.xpi");
      await page.context().close();
    });

    await t.test("every download the dialog links to is made by the build (the signed .xpi is the one optional file)", async () => {
      const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "wt-dl-"));
      execFileSync("node", [path.join(ROOT, "scripts", "build-extension.js"), "--zips", tmp], { stdio: "pipe" });
      const html = fs.readFileSync(path.join(ROOT, "index.html"), "utf8");
      const dialog = html.slice(html.indexOf('<template id="t-install">'), html.indexOf("</template>", html.indexOf('<template id="t-install">')));
      const links = [...dialog.matchAll(/href="(downloads\/[^"]+)"/g)].map((m) => m[1]);
      assert.ok(links.length >= 4, "links found");
      for (const l of links) {
        const f = path.basename(l);
        if (f === "wild-themes-firefox.xpi") continue;
        assert.ok(fs.existsSync(path.join(tmp, f)) || fs.existsSync(path.join(ROOT, l)), l + " is built or committed");
      }
    });
  } finally { await browser.close(); srv.close(); srv2.close(); }
});
