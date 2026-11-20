// File studio: the website's tool for themes already inside files. Same core as the
// browser extension (extension/lib/wtcore.js), so a file changed here is recognised there, and the
// other way round. Runs entirely in the page. The Office add-in is this same page: inside Word,
// PowerPoint or Excel (addin/office-host.js) the file is the open document instead of a download.
const $ = (s, r = document) => r.querySelector(s);
const EXT = new URL("extension/", location.href).href, T = self.WW_THEMES, WT = self.WTCore;
let file = null, bytes = null, info = null, history = [], draft = null, busy = false, community = [];
let office = null, exact = null;   // office: the open document's host; exact: the last file Wild Themes wrote for it
const want = new URLSearchParams(location.search).get("theme");

const say = (t, cls = "") => { const m = $("#msg"); m.textContent = t || ""; m.className = "msg " + cls; };
const store = { get: (k, d) => { try { const v = localStorage.getItem("wt:" + k); return v == null ? d : JSON.parse(v); } catch (_) { return d; } },
  set: (k, v) => { try { localStorage.setItem("wt:" + k, JSON.stringify(v)); } catch (_) { /* private mode */ } } };
const mine = () => store.get("mine", {});   // imported .wildtheme files, on this device only
const toDataURL = (blob) => new Promise((ok) => { const r = new FileReader(); r.onload = () => ok(r.result); r.readAsDataURL(blob); });
const imageFor = async (t) => {
  if ((mine()[t.id] || {}).image) return mine()[t.id].image;
  const c = community.find((x) => x.id === t.id); return c && c.image ? toDataURL(await (await fetch(c.image)).blob()) : null;
};
const makeAssets = WWAssets.maker(EXT, imageFor);
const tool = () => office ? "Wild Themes add-in" : "Wild Themes website";
const themeById = (id) => T[id] ? { ...T[id], id } : mine()[id] ? { ...mine()[id].theme, id } : (() => { const c = community.find((x) => x.id === id); return c ? WWAssets.forCustom({ ...c, image: c.image ? 1 : 0 }) : info && info.theme && info.theme.id === id ? info.theme : null; })();
fetch("community/index.json").then((r) => (r.ok ? r.json() : { themes: [] })).then((j) => { community = j.themes || []; renderGrid(); }).catch(() => {});

// ---------- file in / file out
function wire() {
  const drop = $("#drop"), input = $("input", drop);
  input.onchange = () => { open(input.files[0]); input.value = ""; };
  drop.ondragover = (e) => { e.preventDefault(); drop.classList.add("over"); };
  drop.ondragleave = () => drop.classList.remove("over");
  drop.ondrop = (e) => { e.preventDefault(); drop.classList.remove("over"); open(e.dataTransfer.files[0]); };
}
async function open(f) {
  if (!f) return;
  file = f; history = []; draft = null;
  await set(new Uint8Array(await f.arrayBuffer()), "");
  if (info && want && T[want] && info.status === "none") say(`Pick “${T[want].name}” below (or any theme) to save it into this file.`);
}
// Every change makes new bytes; the file is read back so what's shown is what's really in it.
async function set(next, message, cls) {
  try {
    info = await WT.inspect(next);
    bytes = next;
    $("#file").hidden = $("#tools").hidden = false;
    renderFile(); renderGrid(); renderLook();
    say(message, cls);
  } catch (e) {
    say("That file can't be opened: " + ((e && e.message) || e), "bad");
  }
}
const outName = () => {
  const m = /^(.*?)(\.[^.]+)$/.exec(file.name) || [0, file.name, ""];
  const tag = info && info.theme ? ` (${info.theme.name})` : " (no theme)";
  return m[1].replace(/ \([^)]*\)$/, "") + tag + m[2];
};
$("#download").onclick = async () => {
  if (office) return guard(async () => {
    const out = exact || (info.theme ? await WT.apply(bytes, themeById(info.theme.id) || info.theme, current(), makeAssets, tool()) : bytes);
    if (!office.canCopy()) throw new Error("This version of Office can't open a rebuilt file. Use the Wild Themes website or browser extension instead.");
    say("Opening the file in a new window…");
    await office.openCopy(out);
    say("The file opened in a new window: save it there (File › Save As, you can replace this file). This window is unchanged.", "ok");
  });
  const url = URL.createObjectURL(new Blob([bytes], { type: "application/octet-stream" }));
  Object.assign(document.createElement("a"), { href: url, download: outName() }).click();
  setTimeout(() => URL.revokeObjectURL(url), 30000);
  say(`Saved “${outName()}”.`, "ok");
};

// ---------- what's in the file
function renderFile() {
  const d = WTUI.describe(info);
  $("#sw").style.background = info.theme ? WTUI.swatch(info.theme) : "";
  $("#fname").textContent = d.title; $("#fsub").textContent = d.sub; $("#fpath").textContent = file.name;
  const b = $("#badge"); b.hidden = !d.badge; b.textContent = d.badge; b.className = "st-badge " + d.cls;
  const pr = $("#problems"); pr.replaceChildren(...(d.problems || []).map((t) => Object.assign(document.createElement("li"), { textContent: t })));
  pr.hidden = !pr.children.length;
  $("#repair").hidden = !d.canRepair; $("#repair").textContent = info.status === "legacy" ? "Upgrade theme" : "Restore theme";
  $("#remove").hidden = !d.canRemove;
  $("#undo").hidden = !history.length;
  $("#download").disabled = office ? !(exact || info.theme) : !history.length;
}
function renderGrid() {
  const g = $("#grid"); if (!g) return;
  const cur = info && info.theme ? info.theme.id : null, my = mine();
  const items = Object.entries(T).filter(([, t]) => !t.secret).map(([id, t]) => [id, t, `extension/previews/${id}.webp`])
    .concat(community.map((c) => [c.id, c, c.thumb]), Object.entries(my).map(([id, m]) => [id, { ...m.theme, id }, null, "Mine"]));
  if (cur && !items.some((x) => x[0] === cur)) items.unshift([cur, info.theme, null, "In this file"]);
  g.replaceChildren(...items.map(([id, t, thumb, tag]) => {
    const b = document.createElement("button"), s = document.createElement("span");
    let pic;
    if (thumb) { pic = new Image(); pic.src = thumb; pic.alt = ""; pic.loading = "lazy"; }
    else { pic = document.createElement("i"); pic.style.background = WTUI.swatch(t); }
    b.className = "st-card" + (id === want ? " want" : ""); b.setAttribute("role", "radio"); b.setAttribute("aria-checked", id === cur);
    s.textContent = t.name; b.title = t.blurb || t.name;
    b.append(pic, s);
    if (tag) b.append(Object.assign(document.createElement("em"), { textContent: tag }));
    b.onclick = () => applyTheme(id === cur ? themeById(id) || info.theme : themeById(id) || t, b);
    return b;
  }));
}
$("#import").onchange = async (e) => {
  const f = e.target.files[0]; e.target.value = "";
  if (!f) return;
  try {
    const j = JSON.parse(await f.text());
    if (!j.wildtheme || !j.theme || !j.theme.name) throw new Error("not a theme");
    const id = "custom-" + j.theme.name.toLowerCase().replace(/[^a-z0-9]+/g, "-").slice(0, 40);
    const all = mine(); all[id] = { theme: { ...j.theme, id, custom: 1 }, image: typeof j.image === "string" && j.image.length < 3e6 ? j.image : null };
    store.set("mine", all); renderGrid(); say(`Imported “${j.theme.name}”. It's in the list below the built-in themes.`, "ok");
  } catch (_) { say("That file isn't a Wild Themes theme (.wildtheme).", "bad"); }
};
const current = () => draft || (info && info.settings) || WT.settingsFor(info.kind, {});
function renderLook() {
  WTUI.settingsForm($("#look"), info.kind, current(), (k, v) => { draft = { ...current(), [k]: v }; renderLook(); });
  $("#applyLook").disabled = !draft || !info.theme || busy;
}
$("#resetLook").onclick = () => { if (info) { draft = WT.settingsFor(info.kind, {}); renderLook(); } };

// ---------- changes
async function guard(fn, card) {
  if (busy || !bytes) return;
  busy = true; if (card) card.setAttribute("aria-busy", "true");
  try { await fn(); } catch (e) { say((e && e.message) || String(e), "bad"); }
  finally { busy = false; if (card) card.removeAttribute("aria-busy"); renderLook(); }
}
// expect: the theme id that should be in the file afterwards (null: none); verb: what was done.
const run = (label, fn, card, expect, verb) => guard(async () => {
  say(label);
  const before = bytes, out = await fn();
  draft = null;
  if (office) {
    say(office.canSwap() ? "Saving it into this file…" : "Opening the result in a new window…");
    const res = await office.commit(out, { removing: expect === null });
    if (res.copied) { exact = out; return say(...office.report(res, null, expect, verb)); }
    history.push(before); exact = out;
    await set(res.bytes);
    return say(...office.report(res, info, expect, verb));
  }
  history.push(before);
  await set(out, "", "");
  const ok = info.status === "intact" || (info.status === "none" && expect === null);
  say(ok ? (info.status === "none" ? "Theme removed and the original look restored. Download the file to keep it." : `Checked: ${info.theme.name} is in the file and intact. Download the file to keep it.`)
    : "Done, but the check found problems: " + (info.problems || []).map((p) => p.text).join(" "), ok ? "ok" : "bad");
}, card);
const applyTheme = (t, card) => t && run(`Adding ${t.name}…`, () => WT.apply(bytes, t, current(), makeAssets, tool()), card, t.id, `${t.name} is saved in the file.`);
$("#applyLook").onclick = () => info.theme && run("Applying your changes…", () => WT.apply(bytes, themeById(info.theme.id) || info.theme, draft, makeAssets, tool()), null, info.theme.id, "Changes applied.");
$("#repair").onclick = () => run("Rebuilding the theme from the description inside the file…", () => WT.repair(bytes, makeAssets, tool(), T), null, info.theme.id, "Theme restored.");
$("#remove").onclick = () => run("Removing the theme…", () => WT.remove(bytes), null, null, "Theme removed.");
$("#undo").onclick = () => guard(async () => {
  if (!history.length) return;
  const prev = history.pop(); draft = null;
  if (!office) return set(prev, "Undone.", "");
  say("Putting the file back…");
  const res = await office.commit(prev, { removing: !(await WT.inspect(prev)).theme });
  exact = prev; await set(res.bytes, "Undone.", "ok");
});

// ---------- inside Word, PowerPoint or Excel: the open document is the file
async function readDocument(message, cls) {
  file = { name: office.name || "This " + office.where() };
  await set(await office.read(), message, cls);
  if (info && want && T[want] && info.status === "none") say(`Pick “${T[want].name}” below (or any theme) to save it into this file.`);
}
function startOffice() {
  $("#drop").hidden = true;
  $(".sub").textContent = office.note();
  $("#download").textContent = "Open a themed copy"; $("#download").hidden = !office.canCopy();
  $("#verify").hidden = false;
  $("#verify").onclick = () => guard(() => readDocument("Checked again.", ""));
  const auto = $("#auto");
  $("#autoRow").hidden = !office.autoSupported(); auto.checked = office.auto();
  auto.onchange = () => office.setAuto(auto.checked).catch(() => { auto.checked = false; say("This version of Office can't open add-ins by themselves.", "bad"); });
  $("#file").hidden = false; $("#fname").textContent = "Reading this file…";
  readDocument("").catch((e) => say("Couldn't read this file: " + ((e && e.message) || e), "bad"));
}

// Inside the Teams app the tab can't always save downloads: offer the same page in the browser.
if (new URLSearchParams(location.search).get("host") === "teams") {
  const p = document.createElement("p"); p.className = "sub";
  p.innerHTML = 'In the Teams desktop app, downloads from tabs can be blocked. If “Download the file” does nothing, <a target="_blank" rel="noopener" href="studio.html">open the File studio in your browser</a>.';
  $("#drop").after(p);
}
wire();
if (self.WTOffice) {
  $("#drop").hidden = true;            // no flash of the file picker while Office starts
  WTOffice.ready.then((o) => { office = o; if (o) startOffice(); else $("#drop").hidden = false; });
}
