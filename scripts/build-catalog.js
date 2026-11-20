// Builds catalog.json: the live list of every theme (built-in + approved community themes) that the
// Office add-in and the browser extension read from the website, so new themes reach users
// without a new download. Runs on every Pages deploy (static.yml). Usage: node scripts/build-catalog.js
const fs = require("fs"), path = require("path"), vm = require("vm"), crypto = require("crypto");
const ROOT = path.join(__dirname, "..");
const ctx = { self: {} };
vm.runInNewContext(fs.readFileSync(path.join(ROOT, "extension/lib/themes.js"), "utf8"), ctx);
const exists = (p) => fs.existsSync(path.join(ROOT, p));
const opt = (p) => (exists(p) ? p : undefined);

const themes = [];
for (const [id, t] of Object.entries(ctx.self.WW_THEMES)) {
  if (t.secret) continue;                                           // found by typing; it ships inside the extension
  themes.push({ id, kind: "builtin", ...t,
    preview: opt(`extension/previews/${id}.webp`), scene: opt(`extension/scenes/${id}.css`),
    backdrop: opt(`extension/backdrops/${id}.gif`), backdropWide: opt(`extension/backdrops-wide/${id}.gif`),
    teams: opt(`extension/teams-backgrounds/${id}.jpg`) });
}
let community = { themes: [] };
try { community = JSON.parse(fs.readFileSync(path.join(ROOT, "community/index.json"), "utf8")); } catch (_) { /* none yet */ }
for (const c of community.themes || []) {
  themes.push({ id: c.id, kind: "community", name: c.name, author: c.author, mode: c.mode, motion: c.motion, effect: c.effect,
    colors: c.colors, accent: c.accent, accent2: c.accent2, image: c.image || null, thumb: c.thumb, added: c.added });
}
const body = JSON.stringify(themes);
const out = { format: 1, site: "https://min3scon.github.io/Wild-Themes/", version: crypto.createHash("sha1").update(body).digest("hex").slice(0, 12),
  updated: new Date().toISOString(), themes };
fs.writeFileSync(path.join(ROOT, "catalog.json"), JSON.stringify(out, null, 1) + "\n");
console.log(`catalog.json: ${themes.length} themes, version ${out.version}`);
