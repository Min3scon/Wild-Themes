// Shared test helpers: loads the real extension libraries (the same files the site, extension and add-in run).
const fs = require("fs"), path = require("path"), vm = require("vm"), { spawnSync } = require("child_process");
const JSZip = require("jszip");
const LIB = path.join(__dirname, "..", "extension", "lib");
const FIX = path.join(__dirname, "fixtures");
const OUT = path.join(__dirname, "out");
fs.mkdirSync(OUT, { recursive: true });

const ctx = { self: {} };
vm.runInNewContext(fs.readFileSync(path.join(LIB, "themes.js"), "utf8"), ctx);
const THEMES = ctx.self.WW_THEMES;
const theme = (id) => ({ ...THEMES[id], id });
const WT = require(path.join(LIB, "wtcore.js"));
const fixture = (name) => new Uint8Array(fs.readFileSync(path.join(FIX, name)));
const kindOfName = (n) => (/\.pptx$/.test(n) ? "pptx" : /\.xls[xm]$/.test(n) ? "xlsx" : "docx");

// Stand-ins for the canvas-made pictures (browsers draw these; the tests use fixed files).
async function makeAssets(t, kind, section, settings) {
  if (!settings.backdrop) return { backdrop: null, panel: null };
  const jpg = fixture("backdrop.jpg"), gif = new Uint8Array(fs.readFileSync(path.join(__dirname, "..", "extension", "backdrops-wide", "synthwave.gif"))), png = fixture("panel.png");
  const backdrop = kind === "pptx" && settings.animated ? { bytes: gif, ext: "gif" } : { bytes: jpg, ext: "jpeg" };
  return { backdrop, panel: kind === "xlsx" ? null : async () => png };
}

const zipOf = (bytes) => JSZip.loadAsync(bytes);
async function text(bytes, p) { const z = await zipOf(bytes); const f = z.file(p); return f ? f.async("string") : null; }
async function names(bytes) { return Object.keys((await zipOf(bytes)).files).filter((n) => !n.endsWith("/")).sort(); }
function save(name, bytes) { const p = path.join(OUT, name); fs.writeFileSync(p, bytes); return p; }
const hasSoffice = () => spawnSync("which", ["soffice"]).status === 0;
// Re-save a file through LibreOffice (another editor), returning the new bytes.
function libreRoundTrip(name, bytes) {
  const src = save(name, bytes), dir = path.join(OUT, "lo-" + path.basename(name, path.extname(name)));
  fs.mkdirSync(dir, { recursive: true });
  const ext = path.extname(name).slice(1), filter = { docx: "docx:MS Word 2007 XML", pptx: "pptx:Impress MS PowerPoint 2007 XML", xlsx: "xlsx:Calc MS Excel 2007 XML", xlsm: "xlsm:Calc MS Excel 2007 VBA XML" }[ext];
  const r = spawnSync("soffice", ["--headless", "--convert-to", filter, "--outdir", dir, src], { timeout: 120000 });
  if (r.status !== 0) throw new Error("LibreOffice failed: " + r.stderr);
  return new Uint8Array(fs.readFileSync(path.join(dir, path.basename(name))));
}
// Render page/slide 1 to PNG through LibreOffice + pdftoppm; returns the PNG path.
function render(name, bytes, page = 1) {
  const src = save(name, bytes), dir = path.join(OUT, "render-" + path.basename(name, path.extname(name)));
  fs.mkdirSync(dir, { recursive: true });
  const r = spawnSync("soffice", ["--headless", "--convert-to", "pdf", "--outdir", dir, src], { timeout: 120000 });
  if (r.status !== 0) throw new Error("LibreOffice render failed");
  const pdf = path.join(dir, path.basename(name, path.extname(name)) + ".pdf"), png = path.join(dir, `page${page}`);
  spawnSync("pdftoppm", ["-png", "-r", "40", "-f", String(page), "-l", String(page), "-singlefile", pdf, png]);
  return png + ".png";
}
// Average colour of a box (fractions of width/height) in a PNG, via Python/Pillow.
function sample(png, box) {
  const py = `import sys,json\nfrom PIL import Image\nim=Image.open(sys.argv[1]).convert('RGB');w,h=im.size\nx0,y0,x1,y1=[float(v) for v in sys.argv[2:6]]\nc=im.crop((int(x0*w),int(y0*h),max(int(x1*w),int(x0*w)+1),max(int(y1*h),int(y0*h)+1))).resize((1,1),Image.BOX).getpixel((0,0))\nprint(json.dumps(c))`;
  const r = spawnSync("python3", ["-I", "-c", py, png, ...box.map(String)], { encoding: "utf8" });
  return JSON.parse(r.stdout);
}
const lum = ([r, g, b]) => 0.2126 * r + 0.7152 * g + 0.0722 * b;

module.exports = { WT, THEMES, theme, fixture, makeAssets, zipOf, text, names, save, hasSoffice, libreRoundTrip, render, sample, lum, kindOfName, FIX, OUT };
