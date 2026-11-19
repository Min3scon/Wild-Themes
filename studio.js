// File studio: the website's tool for themes already inside files. Same core as the
// browser extension and the Office add-in (extension/lib/wtcore.js), so a file changed here
// is recognised by them, and the other way round. Runs entirely in the page.
const $ = (s, r = document) => r.querySelector(s);
const EXT = new URL("extension/", location.href).href, T = self.WW_THEMES, WT = self.WTCore;
let file = null, bytes = null, info = null, history = [], draft = null, busy = false, community = [];
const want = new URLSearchParams(location.search).get("theme");

const say = (t, cls = "") => { const m = $("#msg"); m.textContent = t || ""; m.className = "msg " + cls; };
const toDataURL = (blob) => new Promise((ok) => { const r = new FileReader(); r.onload = () => ok(r.result); r.readAsDataURL(blob); });
const imageFor = async (t) => { const c = community.find((x) => x.id === t.id); return c && c.image ? toDataURL(await (await fetch(c.image)).blob()) : null; };
const makeAssets = WWAssets.maker(EXT, imageFor);
const themeById = (id) => T[id] ? { ...T[id], id } : (() => { const c = community.find((x) => x.id === id); return c ? WWAssets.forCustom({ ...c, image: c.image ? 1 : 0 }) : info && info.theme && info.theme.id === id ? info.theme : null; })();
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
$("#download").onclick = () => {
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
  $("#download").disabled = !history.length;
}
function renderGrid() {
  const g = $("#grid"); if (!g) return;
  const cur = info && info.theme ? info.theme.id : null;
  const items = Object.entries(T).map(([id, t]) => [id, t, `extension/previews/${id}.webp`]).concat(community.map((c) => [c.id, c, c.thumb]));
  g.replaceChildren(...items.map(([id, t, thumb]) => {
    const b = document.createElement("button"), img = new Image(), s = document.createElement("span");
    b.className = "st-card" + (id === want ? " want" : ""); b.setAttribute("role", "radio"); b.setAttribute("aria-checked", id === cur);
    img.src = thumb; img.alt = ""; img.loading = "lazy"; s.textContent = t.name; b.title = t.blurb || t.name;
    b.append(img, s); b.onclick = () => applyTheme(themeById(id), b);
    return b;
  }));
}
const current = () => draft || (info && info.settings) || WT.settingsFor(info.kind, {});
function renderLook() {
  WTUI.settingsForm($("#look"), info.kind, current(), (k, v) => { draft = { ...current(), [k]: v }; renderLook(); });
  $("#applyLook").disabled = !draft || !info.theme || busy;
}

// ---------- changes
async function run(label, fn, card) {
  if (busy || !bytes) return;
  busy = true; if (card) card.setAttribute("aria-busy", "true"); say(label);
  try {
    const before = bytes, out = await fn();
    history.push(before); draft = null;
    await set(out, "", "");
    const ok = info.status === "intact" || (info.status === "none" && /Remov/.test(label));
    say(ok ? (info.status === "none" ? "Theme removed and the original look restored. Download the file to keep it." : `Checked: ${info.theme.name} is in the file and intact. Download the file to keep it.`)
      : "Done, but the check found problems: " + (info.problems || []).map((p) => p.text).join(" "), ok ? "ok" : "bad");
  } catch (e) { say((e && e.message) || String(e), "bad"); }
  finally { busy = false; if (card) card.removeAttribute("aria-busy"); renderLook(); }
}
const applyTheme = (t, card) => t && run(`Adding ${t.name}…`, () => WT.apply(bytes, t, current(), makeAssets, "Wild Themes website"), card);
$("#applyLook").onclick = () => info.theme && run("Applying your changes…", () => WT.apply(bytes, themeById(info.theme.id) || info.theme, draft, makeAssets, "Wild Themes website"));
$("#repair").onclick = () => run("Rebuilding the theme from the description inside the file…", () => WT.repair(bytes, makeAssets, "Wild Themes website", T));
$("#remove").onclick = () => run("Removing the theme…", () => WT.remove(bytes));
$("#undo").onclick = async () => { if (!history.length) return; const prev = history.pop(); draft = null; await set(prev, "Undone.", ""); };

// Inside the Teams app the tab can't always save downloads: offer the same page in the browser.
if (new URLSearchParams(location.search).get("host") === "teams") {
  const p = document.createElement("p"); p.className = "sub";
  p.innerHTML = 'In the Teams desktop app, downloads from tabs can be blocked. If “Download the file” does nothing, <a target="_blank" rel="noopener" href="studio.html">open the File studio in your browser</a>.';
  $("#drop").after(p);
}
wire();
