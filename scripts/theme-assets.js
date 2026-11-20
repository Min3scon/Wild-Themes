// Makes the pictures a theme needs, from its scene (extension/scenes/<id>.css), so a new theme is one scene file plus an entry
// in extension/lib/themes.js:
//   extension/previews/<id>.webp           320x180   gallery card
//   extension/backdrops/<id>.gif           794x1123  1 frame, Word (the page background)
//   extension/backdrops-wide/<id>.gif      640x360   6 frames, PowerPoint
//   extension/teams-backgrounds/<id>.jpg   1920x1080 Teams meeting background
// The GIFs loop seamlessly: for them each layer is run for exactly one cycle of its pattern (one text line, one word...).
// Usage: node scripts/theme-assets.js <id> [base colour]    (needs playwright and python3 with Pillow)
const fs = require("fs"), path = require("path"), os = require("os"), { spawnSync } = require("child_process");
const ROOT = path.join(__dirname, ".."), EXT = path.join(ROOT, "extension");
const [id, base = "#05040d"] = process.argv.slice(2);
if (!id) { console.error("usage: node scripts/theme-assets.js <id> [base colour]"); process.exit(1); }
let chromium; try { ({ chromium } = require("playwright")); } catch (_) { ({ chromium } = require(path.join(ROOT, "tests", "node_modules", "playwright"))); }
const css = fs.readFileSync(path.join(EXT, "content.css"), "utf8") + fs.readFileSync(path.join(EXT, "scenes", `${id}.css`), "utf8");
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "wt-assets-"));

async function frames(browser, w, h, times, loop) {
  const page = await browser.newPage({ viewport: { width: w, height: h } });
  // loop: one cycle of every layer in 1.8 s; the fine scanline layer and the text glow are left out of the GIFs (they don't compress: size matters in .pptx files)
  const extra = loop ? "#ww-bg>i{animation-duration:1.8s!important;--n:1!important;--w:1!important}#ww-bg>i:nth-child(4){background-image:radial-gradient(120% 100% at 50% 50%,#0000 55%,#000000b3 100%)!important}#ww-bg>i{text-shadow:none!important}" : "";
  await page.setContent(`<!doctype html><html data-ww="${id}" style="--ww-base:${base}"><head><style>${css}${extra}</style></head><body style="margin:0"><div id="ww-bg" data-scene="${id}"><i></i><i></i><i></i><i></i></div></body></html>`);
  await page.waitForTimeout(400);
  const files = [];
  for (const [k, t] of times.entries()) {
    await page.evaluate((ms) => document.getAnimations().forEach((a) => { a.pause(); a.currentTime = ms; }), t);
    await page.waitForTimeout(60);
    const f = path.join(tmp, `${w}x${h}-${k}.png`); await page.screenshot({ path: f }); files.push(f);
  }
  await page.close(); return files;
}
const py = (code, args) => { const r = spawnSync("python3", ["-I", "-c", code, ...args], { encoding: "utf8" }); if (r.status) throw new Error(r.stderr); };
const GIF = `import sys
from PIL import Image
out,delay=sys.argv[1],int(sys.argv[2]); fr=[Image.open(f).convert("RGB") for f in sys.argv[3:]]
w,h=fr[0].size; m=Image.new("RGB",(w,h*len(fr)))
for i,f in enumerate(fr): m.paste(f,(0,h*i))
pal=m.quantize(colors=int(__import__("os").environ.get("GIF_COLORS","64")),method=Image.Quantize.MEDIANCUT,dither=Image.Dither.NONE)
q=[f.quantize(palette=pal,dither=Image.Dither.NONE) for f in fr]
q[0].save(out,save_all=True,append_images=q[1:],duration=delay,loop=0,optimize=True,disposal=1)`;
const STILL = `import sys
from PIL import Image
src,out,w,h,kind,q=sys.argv[1],sys.argv[2],int(sys.argv[3]),int(sys.argv[4]),sys.argv[5],int(sys.argv[6])
im=Image.open(src).convert("RGB")
if im.size!=(w,h): im=im.resize((w,h),Image.LANCZOS)
im.save(out,kind,quality=q,**({"method":6} if kind=="WEBP" else {"optimize":True}))`;

(async () => {
  const browser = await chromium.launch();
  const at = 4000;
  const [still] = await frames(browser, 1280, 720, [at], false);
  py(STILL, [still, path.join(EXT, "previews", `${id}.webp`), 320, 180, "WEBP", 78]);
  const [big] = await frames(browser, 1920, 1080, [at], false);
  py(STILL, [big, path.join(EXT, "teams-backgrounds", `${id}.jpg`), 1920, 1080, "JPEG", 88]);
  const wide = await frames(browser, 640, 360, Array.from({ length: 6 }, (_, k) => k * 300), true);
  py(GIF, [path.join(EXT, "backdrops-wide", `${id}.gif`), "300", ...wide]);
  // Word only ever uses the first frame of this one (it plays a GIF as a video on top of the text), so it is a single frame
  const tall = await frames(browser, 794, 1123, [0], true);
  py(GIF, [path.join(EXT, "backdrops", `${id}.gif`), "1000", ...tall]);
  await browser.close();
  for (const f of [`previews/${id}.webp`, `backdrops/${id}.gif`, `backdrops-wide/${id}.gif`, `teams-backgrounds/${id}.jpg`]) console.log(f, (fs.statSync(path.join(EXT, f)).size / 1024).toFixed(0) + " KB");
  fs.rmSync(tmp, { recursive: true, force: true });
})();
