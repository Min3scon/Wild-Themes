// Wild Themes add-in for Word, PowerPoint and Excel (desktop, web and Mac).
// Same document core as the website and the browser extension (../extension/lib/wtcore.js):
// it reads the open file, shows which theme it carries and whether it's intact, and applies,
// changes, configures, repairs or removes the theme by rewriting the file itself.
// Nothing changes until you click; after every change the file is read back and checked.
const T = self.WW_THEMES, WT = self.WTCore, A = self.WWAssets, $ = (s) => document.querySelector(s);
const BASE = new URL("../extension/", location.href).href;
let host = "word", kind = "docx", info = null, busy = false, filter = "all", draft = null, lastRemoved = null;

// ------------------------------------------------------------------ small storage (this device only)
const store = {
  get(k, d) { try { const v = localStorage.getItem("wt:" + k); return v == null ? d : JSON.parse(v); } catch (_) { return d; } },
  set(k, v) { try { localStorage.setItem("wt:" + k, JSON.stringify(v)); } catch (_) { /* private mode */ } },
};
const mine = () => store.get("mine", {});
const makeAssets = A.maker(BASE, async (t) => (mine()[t.id] || {}).image || null);
const say = (t) => { $("#msg").textContent = t || ""; };
const themeById = (id) => T[id] ? { ...T[id], id } : mine()[id] ? { ...mine()[id].theme, id } : info && info.theme && info.theme.id === id && info.theme.doc ? info.theme : null;

// ------------------------------------------------------------------ the open file
function getFile() {
  return new Promise((ok, fail) => Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 4194304 }, (r) => {
    if (r.status !== Office.AsyncResultStatus.Succeeded) return fail(r.error);
    const file = r.value, parts = []; let i = 0;
    const next = () => file.getSliceAsync(i, (s) => {
      if (s.status !== Office.AsyncResultStatus.Succeeded) { file.closeAsync(); return fail(s.error); }
      parts.push(new Uint8Array(s.value.data));
      if (++i < file.sliceCount) next();
      else { file.closeAsync(); const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0)); let o = 0; for (const p of parts) { out.set(p, o); o += p.length; } ok(out); }
    });
    next();
  }));
}
function toBase64(bytes) {
  let s = ""; for (let i = 0; i < bytes.length; i += 32768) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 32768));
  return btoa(s);
}
const supports = (set, v) => Office.context.requirements.isSetSupported(set, v);
async function metaXmlOf(bytes) {
  const found = await WT.findMetaPart(await JSZip.loadAsync(bytes));
  if (!found) return null;
  return (await (await JSZip.loadAsync(bytes)).file(found.path).async("string")).replace(/^<\?xml[^>]*\?>\s*/, "");
}

// Word: the theme's pictures sit in locked content controls; only this add-in unlocks them.
async function unlockWord() {
  try {
    await Word.run(async (ctx) => {
      const lists = [ctx.document.body.contentControls.getByTag("WildThemeMarker")];
      const secs = ctx.document.sections; secs.load("items"); await ctx.sync();
      for (const sec of secs.items) for (const k of ["Primary", "FirstPage", "EvenPages"]) lists.push(sec.getHeader(k).contentControls.getByTag("WildTheme"));
      for (const l of lists) l.load("items"); await ctx.sync();
      for (const l of lists) for (const cc of l.items) { cc.cannotDelete = false; cc.cannotEdit = false; }
      await ctx.sync();
    });
  } catch (_) { /* nothing locked */ }
}

// Keep the live document's own custom XML part + properties in step with the rebuilt file
// (Word's insert and PowerPoint's slide insert don't carry package-level parts across).
async function syncMeta(bytes) {
  const xml = await metaXmlOf(bytes);
  const props = xml ? { WildTheme: info && info.theme ? info.theme.id : "", WildThemeVersion: String(WT.VERSION) } : null;
  const run = host === "word" ? Word.run : host === "ppt" ? PowerPoint.run : null;
  const can = host === "word" ? supports("WordApi", "1.4") : supports("PowerPointApi", "1.7");
  if (!run || !can) return;
  await run(async (ctx) => {
    const doc = host === "word" ? ctx.document : ctx.presentation;
    const parts = doc.customXmlParts.getByNamespace(WT.NS); parts.load("items"); await ctx.sync();
    for (const p of parts.items) p.delete();
    if (xml) doc.customXmlParts.add(xml);
    await ctx.sync();
    try {   // custom properties: secondary marker that survives editors which drop custom XML
      const cp = doc.properties.customProperties; cp.load("items/key"); await ctx.sync();
      for (const p of cp.items) if (/^WildTheme/.test(p.key)) p.delete();
      if (props) { const name = await nameFromXml(xml); cp.add("WildTheme", name.id); cp.add("WildThemeName", name.name); cp.add("WildThemeVersion", props.WildThemeVersion); }
      await ctx.sync();
    } catch (_) { /* older Office: the custom XML part is enough */ }
  });
}
async function nameFromXml(xml) { const m = WT.parseMeta(xml); return { id: m.id, name: m.name }; }

// Put a rebuilt file in place of the open one. Excel can't swap a whole workbook from an add-in.
async function swapIn(bytes, { removing } = {}) {
  const b64 = toBase64(bytes);
  if (b64.length > 4 * 1024 * 1024 * 1.37) throw new Error("This file is over about 4 MB, too big for Office to swap in place. Use “Open a themed copy” instead.");
  if (host === "word") {
    const opts = { importTheme: true, importStyles: true, importPageColor: true, importParagraphSpacing: true };
    if (supports("WordApi", "1.6")) opts.importCustomProperties = true;
    if (supports("WordApi", "1.7")) opts.importDifferentOddEvenPages = true;
    await unlockWord();
    await Word.run(async (ctx) => { ctx.document.insertFileFromBase64(b64, "Replace", opts); await ctx.sync(); });
    if (removing && supports("WordApiDesktop", "1.4")) {   // the page background isn't cleared by an import that has none
      try { await Word.run(async (ctx) => { ctx.document.background.fill.clear(); await ctx.sync(); }); } catch (_) { /* checked below */ }
    }
  } else if (host === "ppt") {
    await PowerPoint.run(async (ctx) => {
      const slides = ctx.presentation.slides; slides.load("items/id"); await ctx.sync();
      const old = slides.items.map((s) => s.id);
      ctx.presentation.insertSlidesFromBase64(b64, { formatting: "KeepSourceFormatting", targetSlideId: old[old.length - 1] });
      await ctx.sync();
      for (const id of old) slides.getItem(id).delete();
      await ctx.sync();
    });
  } else throw new Error("Excel opens the themed workbook as a copy.");
  await syncMeta(bytes);
}
// Open the rebuilt file as a new document (exactly the file Wild Themes wrote, nothing re-imported).
async function openCopy(bytes) {
  const b64 = toBase64(bytes);
  if (host === "word") await Word.run(async (ctx) => { ctx.application.createDocument(b64).open(); await ctx.sync(); });
  else if (host === "ppt") await PowerPoint.createPresentation(b64);
  else await Excel.createWorkbook(b64);
}
const canSwap = () => host === "word" ? supports("WordApi", "1.5") : host === "ppt" ? supports("PowerPointApi", "1.2") : false;
const canCopy = () => host === "word" ? supports("WordApi", "1.3") : host === "ppt" ? supports("PowerPointApi", "1.1") : supports("ExcelApi", "1.8");

// ------------------------------------------------------------------ checking the file
async function check(quiet) {
  try {
    if (!quiet) say("Reading this file…");
    info = await WT.inspect(await getFile());
    if (!quiet) say("");
  } catch (e) {
    info = null; say("Couldn't read this file: " + ((e && e.message) || e));
  }
  draft = null;
  renderFile(); renderGrid(); renderLook();
  return info;
}
const swatch = (t) => WTUI.swatch(t);
function renderFile() {
  const d = WTUI.describe(info), st = info ? info.status : "none";
  $("#sw").style.background = info && info.theme ? swatch(info.theme) : "";
  $("#fname").textContent = d.title; $("#fsub").textContent = d.sub;
  const badge = $("#badge"); badge.hidden = !d.badge; badge.className = "badge " + d.cls; badge.textContent = d.badge;
  const pr = $("#problems"); pr.textContent = "";
  for (const p of d.problems || []) { const li = document.createElement("li"); li.textContent = p; pr.append(li); }
  pr.hidden = !pr.children.length;
  $("#repair").hidden = !d.canRepair;
  $("#repair").textContent = st === "legacy" ? "Upgrade theme" : "Restore theme";
  $("#remove").hidden = !d.canRemove;
  $("#undo").hidden = !(lastRemoved && st === "none");
  $("#altrow").hidden = !canCopy() || host === "excel";
}

// ------------------------------------------------------------------ gallery
function renderGrid() {
  for (const b of document.querySelectorAll(".seg button")) b.setAttribute("aria-checked", b.dataset.f === filter);
  const grid = $("#grid"); grid.textContent = "";
  const cur = info && info.theme ? info.theme.id : null, items = [];
  if (filter === "mine") {
    for (const [id, m] of Object.entries(mine())) items.push([id, { ...m.theme, id }, null]);
    if (info && info.theme && info.theme.custom && !mine()[cur]) items.unshift([cur, info.theme, "In this file"]);
  } else for (const [id, t] of Object.entries(T)) if (filter === "all" || t.mode === filter) items.push([id, t, null]);
  if (!items.length) { const p = document.createElement("p"); p.className = "empty"; p.textContent = filter === "mine" ? "No custom themes yet. Make one on the website or in the browser extension, export it as a .wildtheme file and import it here." : "No themes."; grid.append(p); }
  for (const [id, t, tag] of items) {
    const b = document.createElement("button"), s = document.createElement("span");
    b.className = "card"; b.setAttribute("role", "radio"); b.setAttribute("aria-checked", id === cur); b.title = t.blurb || t.name;
    let pic;
    if (T[id]) { pic = new Image(); pic.src = `${BASE}previews/${id}.webp`; pic.alt = ""; }
    else { pic = document.createElement("i"); pic.style.background = swatch(t); }
    s.textContent = t.name;
    b.append(pic, s);
    if (tag) { const em = document.createElement("em"); em.textContent = tag; b.append(em); }
    b.onclick = () => applyTheme(id === cur ? info.theme : themeById(id) || t, b);
    grid.append(b);
  }
}
for (const b of document.querySelectorAll(".seg button")) b.onclick = () => { filter = b.dataset.f; renderGrid(); };
$("#import").onchange = async (e) => {
  const f = e.target.files[0]; e.target.value = "";
  if (!f) return;
  try {
    const j = JSON.parse(await f.text());
    if (!j.wildtheme || !j.theme || !j.theme.name) throw new Error("not a theme");
    const id = "custom-" + j.theme.name.toLowerCase().replace(/[^a-z0-9]+/g, "-").slice(0, 40);
    const all = mine(); all[id] = { theme: { ...j.theme, id, custom: 1 }, image: typeof j.image === "string" && j.image.length < 3e6 ? j.image : null };
    store.set("mine", all); filter = "mine"; renderGrid(); say(`Imported “${j.theme.name}”.`);
  } catch (_) { say("That file isn't a Wild Themes theme (.wildtheme)."); }
};

// ------------------------------------------------------------------ look & effects (form shared with the website and extension)
const current = () => draft || (info && info.settings) || WT.settingsFor(kind, store.get("settings:" + kind, {}));
function renderLook() {
  WTUI.settingsForm($("#look"), kind, current(), edit);
  $("#applyLook").disabled = !draft || !info || !info.theme || busy;
  $("#applyLook").textContent = info && info.theme ? "Apply changes" : "Used for the next theme";
}
function edit(k, v) { draft = { ...current(), [k]: v }; store.set("settings:" + kind, draft); renderLook(); }
$("#resetLook").onclick = () => { draft = WT.settingsFor(kind, {}); store.set("settings:" + kind, draft); renderLook(); };
$("#applyLook").onclick = () => info && info.theme && applyTheme(themeById(info.theme.id) || info.theme, null, draft);

for (const tab of document.querySelectorAll('[role="tab"]')) tab.onclick = () => {
  for (const t of document.querySelectorAll('[role="tab"]')) { const on = t === tab; t.setAttribute("aria-selected", on); $("#" + t.getAttribute("aria-controls")).hidden = !on; }
};

// ------------------------------------------------------------------ actions
async function guarded(fn, card) {
  if (busy) return;
  busy = true; if (card) card.setAttribute("aria-busy", "true");
  for (const b of document.querySelectorAll(".btn")) b.disabled = true;
  try { await fn(); }
  catch (e) { say((e && e.message) || String(e)); }
  finally {
    busy = false; if (card) card.removeAttribute("aria-busy");
    for (const b of document.querySelectorAll(".btn")) b.disabled = false;
    renderLook();
  }
}
// Write `bytes` back (in place, or as a copy), then read the file again and report what's really in it.
async function writeBack(bytes, verb, expectId, opts = {}) {
  if (host === "excel" || opts.copy || !canSwap()) {
    if (!canCopy()) throw new Error("This version of Office can't open a rebuilt file. Use the Wild Themes website or browser extension instead.");
    say("Opening the result in a new window…");
    await openCopy(bytes);
    say(`${verb} The result opened in a new window: save it there (File › Save As, you can replace this file). This window is unchanged.`);
    return;
  }
  say("Saving it into this file…");
  await swapIn(bytes, opts);
  const now = await check(true);
  if (expectId === null) {
    say(now && now.status === "none" ? `${verb} Checked: no theme left in this file.` : "Most of the theme was removed, but Office kept some of it in this window. Use “Open a themed copy” for a clean file, or remove it from the website.");
  } else if (now && now.theme && now.theme.id === expectId && now.status === "intact") {
    say(`${verb} Checked: the theme is in this file and intact.`);
  } else {
    const why = now && now.problems && now.problems.length ? " " + now.problems.map((p) => p.text).join(" ") : "";
    say(`Office only took part of the theme into this window.${why} “Open a themed copy” gives you the complete file.`);
  }
}
let lastBytes = null;
function applyTheme(t, card, settings) {
  if (!t) return say("That theme isn't available here.");
  return guarded(async () => {
    say("Reading the file…");
    const bytes = await getFile();
    say(`Adding ${t.name}…`);
    const s = settings || current();
    const out = await WT.apply(bytes, t, s, makeAssets, "Wild Themes add-in");
    lastBytes = out;
    if (!info || !info.theme || info.theme.id !== t.id) info = { ...(info || {}), theme: t };
    await writeBack(out, `${t.name} is saved in the file.`, t.id);
  }, card);
}
$("#verify").onclick = () => guarded(() => check());
$("#repair").onclick = () => guarded(async () => {
  say("Rebuilding the theme from the description inside this file…");
  const out = await WT.repair(await getFile(), makeAssets, "Wild Themes add-in", T);
  lastBytes = out;
  await writeBack(out, "Theme restored.", info.theme.id);
});
$("#remove").onclick = () => guarded(async () => {
  lastRemoved = info && info.theme ? { theme: themeById(info.theme.id) || info.theme, settings: info.settings } : null;
  say("Removing the theme and putting back the original look…");
  const out = await WT.remove(await getFile());
  lastBytes = out;
  await writeBack(out, "Theme removed.", null, { removing: true });
});
$("#undo").onclick = () => lastRemoved && applyTheme(lastRemoved.theme, null, lastRemoved.settings);
$("#copy").onclick = () => guarded(async () => {
  let bytes = lastBytes;
  if (!bytes) {
    const t = info && info.theme && (themeById(info.theme.id) || info.theme);
    if (!t) throw new Error("Pick a theme first.");
    bytes = await WT.apply(await getFile(), t, current(), makeAssets, "Wild Themes add-in");
  }
  await writeBack(bytes, "Done.", null, { copy: true });
});

// ------------------------------------------------------------------ open by itself when a themed file opens
const SR = () => supports("SharedRuntime", "1.1") && Office.addin && Office.addin.setStartupBehavior;
$("#auto").onchange = async (e) => {
  try { await Office.addin.setStartupBehavior(e.target.checked ? Office.StartupBehavior.load : Office.StartupBehavior.none); store.set("auto", e.target.checked); }
  catch (_) { e.target.checked = false; say("This version of Office can't open add-ins by themselves."); }
};
// Cheap check (no full file read): does the document hold our custom XML part?
async function hasThemePart() {
  const run = host === "word" ? Word.run : host === "ppt" ? PowerPoint.run : Excel.run;
  const ok = host === "word" ? supports("WordApi", "1.4") : host === "ppt" ? supports("PowerPointApi", "1.7") : supports("ExcelApi", "1.5");
  if (!ok) return null;
  return run(async (ctx) => {
    const doc = host === "word" ? ctx.document : host === "ppt" ? ctx.presentation : ctx.workbook;
    const parts = doc.customXmlParts.getByNamespace(WT.NS); parts.load("items"); await ctx.sync();
    return parts.items.length > 0;
  });
}

Office.onReady(async (o) => {
  host = o.host === Office.HostType.PowerPoint ? "ppt" : o.host === Office.HostType.Excel ? "excel" : "word";
  kind = { word: "docx", ppt: "pptx", excel: "xlsx" }[host];
  $("#hostnote").textContent = {
    word: "In Word the scene fills the page behind a light reading panel, and headings, lists, tables, links, headers and footers follow the theme. Changes are made in this window; “Open a themed copy” gives an exact copy instead.",
    ppt: "In PowerPoint the scene sits on the slide masters behind a see-through panel; titles, bullets, tables and charts follow the theme, and slides can get a transition. Animated scenes play in slide shows.",
    excel: "In Excel the theme colours and fonts drive cell styles, tables, charts and tab colours, with a pale picture behind every sheet (desktop Excel shows it). Office doesn't let add-ins swap a whole workbook in place, so the themed workbook opens in a new window for you to save.",
  }[host];
  const sr = SR();
  $("#auto").closest("label").hidden = !sr;
  if (sr) $("#auto").checked = store.get("auto", false);
  // Started in the background (shared runtime) for a file that carries a theme: show the pane.
  try {
    if (sr && store.get("auto", false) && await hasThemePart()) await Office.addin.showAsTaskpane();
  } catch (_) { /* pane stays as it is */ }
  renderGrid(); renderLook();
  await check();
});
