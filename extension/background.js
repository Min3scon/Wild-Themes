// Wild Themes background script (Chrome/Edge/Brave: service worker; Firefox: event page).
// 1. Loads only the active scene's CSS into the Word frame that asked for it.
// 2. Keeps the theme list in step with the website (catalog.json): new built-in scenes and
//    approved community themes arrive by themselves, checked first, every few hours.
// Chrome runs this as a service worker and loads the helpers here; Firefox loads them from the
// manifest's "scripts" list, before this file, and has no importScripts.
if (typeof importScripts === "function") importScripts("lib/themes.js", "lib/catalog.js", "lib/rewards.js");
const BUNDLED = self.WW_THEMES, C = self.WTCatalog, RW = self.WWRewards;

chrome.runtime.onMessage.addListener((m, s, reply) => {
  if (m.css && /^[a-z0-9-]+$/.test(m.css)) {
    const target = { tabId: s.tab.id, frameIds: [s.frameId] };
    if (BUNDLED[m.css]) chrome.scripting.insertCSS({ target, files: [`scenes/${m.css}.css`] });
    else chrome.storage.local.get("scene:" + m.css, (v) => { const css = C.safeCss(v["scene:" + m.css]); if (css) chrome.scripting.insertCSS({ target, css }); });
  }
  if (m.sync) { sync().then((r) => reply(r), (e) => reply({ error: String(e && e.message || e) })); return true; }
  if (m.picture && /^[a-z0-9-]+$/.test(m.picture)) { cachePicture(m.picture).then((ok) => reply({ ok }), () => reply({ ok: false })); return true; }
  if (m.words) {
    const n = Math.floor(m.words);
    if (!(n > 0 && n <= 500) || typeof m.theme != "string" || !/^[\w-]{1,64}$/.test(m.theme)) return void reply(null);
    addWords(m.theme, n).then(reply, () => reply(null));
    return true;
  }
});

// Words typed, per theme (this browser only), and the secret theme's dice: each word is one roll with a 1 in RW.ODDS chance.
// Once found it's saved in sync storage, so it's unlocked wherever you're signed in. One update at a time: pages in several
// tabs can report together. self.wtRandom replaces the dice in the tests.
let wordQueue = Promise.resolve();
const roll = () => (self.wtRandom ? self.wtRandom() : crypto.getRandomValues(new Uint32Array(1))[0] / 4294967296);
function addWords(theme, n) {
  const job = wordQueue.then(async () => {
    const { words = {} } = await chrome.storage.local.get("words");
    const before = RW.total(words), total = before + n;
    words[theme] = (+words[theme] || 0) + n;
    await chrome.storage.local.set({ words });
    const out = { total, theme: words[theme] };
    const { found = {} } = await chrome.storage.sync.get("found");
    if (!found[RW.SECRET]) for (let i = 0; i < n; i++) if (roll() < 1 / RW.ODDS) {
      found[RW.SECRET] = { at: Date.now(), words: before + i + 1, theme };
      await chrome.storage.sync.set({ found });
      out.secret = RW.SECRET; out.total = before + i + 1;
      break;
    }
    const m = RW.MILESTONES.filter((x) => before < x.at && total >= x.at).pop();
    if (m) out.badge = m.id;
    return out;
  });
  wordQueue = job.catch(() => {});
  return job;
}

// Fetch the catalogue; store new built-in themes (with their scene CSS) and community themes.
async function sync() {
  const cat = await C.load();
  const remote = {}, community = [];
  for (const t of C.newBuiltins(cat, BUNDLED)) {
    if (!t.scene) continue;
    const r = await fetch(C.abs(t.scene), { cache: "no-cache" });
    const css = r.ok ? C.safeCss(await r.text()) : null;
    if (!css) continue;                                             // unsafe or missing: skip that theme
    await chrome.storage.local.set({ ["scene:" + t.id]: css });
    const { kind, scene, backdrop, backdropWide, teams, ...def } = t;
    remote[t.id] = { ...def, preview: C.abs(t.preview), backdropUrl: C.abs(backdrop), backdropWideUrl: C.abs(backdropWide), teamsUrl: C.abs(teams), remote: 1 };
  }
  for (const t of cat.themes) if (t.kind === "community") community.push(C.asCustom(t));
  await chrome.storage.local.set({ remote, community, catalog: { version: cat.version, updated: cat.updated, at: Date.now() } });
  return { version: cat.version, builtins: Object.keys(remote).length, community: community.length };
}
// Community picture themes: the picture is fetched once, when someone first picks the theme.
async function cachePicture(id) {
  const { community = [] } = await chrome.storage.local.get("community");
  const t = community.find((c) => c.id === id);
  if (!t || !t.imageUrl || !t.imageUrl.startsWith(C.SITE)) return false;
  if ((await chrome.storage.local.get("img:" + id))["img:" + id]) return true;
  const r = await fetch(t.imageUrl);
  if (!r.ok) return false;
  const b = await r.blob();
  if (b.size > 3e6 || !/^image\/(jpeg|png|webp)$/.test(b.type)) return false;
  const bytes = new Uint8Array(await b.arrayBuffer());
  let s = ""; for (let i = 0; i < bytes.length; i += 32768) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 32768));
  await chrome.storage.local.set({ ["img:" + id]: `data:${b.type};base64,${btoa(s)}` });
  return true;
}

// create it once: re-creating it each time the worker wakes would keep pushing the first run back
chrome.alarms.get("catalog", (a) => { if (!a) chrome.alarms.create("catalog", { periodInMinutes: 360 }); });
chrome.alarms.onAlarm.addListener((a) => { if (a.name === "catalog") sync().catch(() => {}); });
chrome.runtime.onStartup.addListener(() => sync().catch(() => {}));
chrome.runtime.onInstalled.addListener(({ reason }) => {
  sync().catch(() => {});
  if (reason == "install") chrome.tabs.create({ url: chrome.runtime.getURL("options.html#file") });
});

// Wild Themes for Word (desktop app) listens on this PC only; tell it the chosen theme so desktop
// Word matches the browser. Nothing happens when the app isn't running.
const tellDesktop = () => chrome.storage.sync.get({ theme: "", on: true }, (v) => {
  fetch(`http://127.0.0.1:47615/theme?id=${encodeURIComponent(v.on ? v.theme : "")}`).catch(() => {});
});
chrome.storage.onChanged.addListener((c, area) => { if (area === "sync" && (c.theme || c.on)) tellDesktop(); });
chrome.runtime.onStartup.addListener(tellDesktop);
chrome.runtime.onInstalled.addListener(tellDesktop);
