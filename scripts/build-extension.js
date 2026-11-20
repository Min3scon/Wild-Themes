// Builds the browser extension packages from extension/:
//   downloads/wild-themes-extension.zip   Chrome, Edge, Brave, Arc (the manifest as it is)
//   downloads/wild-themes-firefox.zip     Firefox (event-page background, add-on id, data-collection declaration)
//   downloads/wild-themes-firefox-unsigned.xpi   the same file under the name Firefox's "install from file" expects
// Both are plain copies of extension/ with a different manifest: one code base, nothing forked.
// Usage: node scripts/build-extension.js [--zips <folder>] [--dir <folder>]
//   --zips: where the zips go (default downloads/); --dir: write unzipped folders there instead of zips (for loading by hand)
const fs = require("fs"), path = require("path"), { execFileSync } = require("child_process");
const ROOT = path.join(__dirname, ".."), SRC = path.join(ROOT, "extension");

// Firefox: no service workers (a non-persistent event page instead, with the helper scripts listed
// first because there is no importScripts), and the add-on needs a stable id plus, since late 2025,
// a statement of what data it collects. Wild Themes sends nothing anywhere, so: none.
const GECKO_ID = "wild-themes@min3scon.github.io", GECKO_MIN = "140.0", GECKO_ANDROID_MIN = "142.0";
function firefoxManifest(m) {
  const f = JSON.parse(JSON.stringify(m));
  f.background = { scripts: ["lib/themes.js", "lib/catalog.js", "background.js"] };
  f.browser_specific_settings = {
    gecko: { id: GECKO_ID, strict_min_version: GECKO_MIN, data_collection_permissions: { required: ["none"] } },
    gecko_android: { strict_min_version: GECKO_ANDROID_MIN },
  };
  return f;
}
const chromeManifest = (m) => JSON.parse(JSON.stringify(m));

function build(target, outDir) {
  fs.rmSync(outDir, { recursive: true, force: true });
  fs.cpSync(SRC, outDir, { recursive: true });
  const m = JSON.parse(fs.readFileSync(path.join(SRC, "manifest.json"), "utf8"));
  fs.writeFileSync(path.join(outDir, "manifest.json"), JSON.stringify(target === "firefox" ? firefoxManifest(m) : chromeManifest(m), null, 2) + "\n");
  return outDir;
}
function zip(dir, file) {
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.rmSync(file, { force: true });
  execFileSync("zip", ["-rqX", path.resolve(file), "."], { cwd: dir });
}

module.exports = { build, firefoxManifest, chromeManifest, GECKO_ID, GECKO_MIN };

if (require.main === module) {
  const arg = (n) => { const i = process.argv.indexOf(n); return i > 0 ? path.resolve(process.argv[i + 1]) : null; };
  const dir = arg("--dir"), zips = arg("--zips") || path.join(ROOT, "downloads");
  const tmp = dir || fs.mkdtempSync(path.join(require("os").tmpdir(), "wt-build-"));
  for (const [target, file] of [["chrome", "wild-themes-extension.zip"], ["firefox", "wild-themes-firefox.zip"]]) {
    const out = build(target, path.join(tmp, target));
    if (dir) { console.log(`${target}: ${out}`); continue; }
    zip(out, path.join(zips, file));
    console.log(`${file}: ${(fs.statSync(path.join(zips, file)).size / 1048576).toFixed(1)} MB`);
    // An .xpi is a zip with another name. Firefox Developer Edition, Nightly and ESR (signature check off) install from a
    // file called .xpi. Not the name of the signed copy (wild-themes-firefox.xpi), which is added by hand after signing.
    if (target === "firefox") fs.copyFileSync(path.join(zips, file), path.join(zips, "wild-themes-firefox-unsigned.xpi"));
  }
  if (!dir) fs.rmSync(tmp, { recursive: true, force: true });
}
