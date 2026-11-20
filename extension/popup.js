const T = self.WW_THEMES, $ = (s) => document.querySelector(s);
const DEF = { on: true, theme: "synthwave", motion: "auto", sparks: true, docs: {}, docsDark: "auto", docsClean: true };
let S = DEF, customs = [], community = [], remote = {}, status = null, scope = "all", filter = "all";

const current = () => (scope === "doc" && status && status.docKey && S.docs[status.docKey]) || S.theme;
const lookup = (id) => T[id] || remote[id] || customs.find((c) => c.id === id) || community.find((c) => c.id === id);
const save = (patch) => { Object.assign(S, patch); chrome.storage.sync.set(patch); render(); };

async function choose(id) {
  const c = community.find((x) => x.id === id);   // community picture: fetched once from the website
  if (c && c.image) await new Promise((ok) => chrome.runtime.sendMessage({ picture: id }, ok));
  if (scope === "doc" && status && status.docKey) {
    const docs = { ...S.docs, [status.docKey]: id };
    const keys = Object.keys(docs);
    if (keys.length > 300) delete docs[keys[0]];                 // keep sync storage small
    save({ docs, on: true });
  } else save({ theme: id, on: true });
}

function card(id, name, thumb) {
  const b = document.createElement("button");
  b.className = "card"; b.setAttribute("role", "radio"); b.dataset.id = id; b.title = name;
  const img = document.createElement("img");
  img.alt = ""; img.loading = "lazy"; img.src = thumb;
  img.onerror = () => { const t = lookup(id); img.removeAttribute("src"); if (t) img.style.background = `linear-gradient(135deg,${t.accent},${t.accent2 || t.accent})`; };
  const s = document.createElement("span"); s.textContent = name;
  b.append(img, s);
  b.onclick = () => choose(id);
  return b;
}

function render() {
  const t = lookup(current()) || T[DEF.theme];
  document.documentElement.style.setProperty("--a", t.accent);
  document.documentElement.style.setProperty("--b", t.accent2 || t.accent);
  document.body.classList.toggle("off", !S.on);
  $("#on").checked = S.on; $("#sparks").checked = S.sparks;
  for (const b of document.querySelectorAll("#motion button")) b.setAttribute("aria-checked", b.dataset.v === S.motion);
  for (const b of document.querySelectorAll(".filter button")) b.setAttribute("aria-checked", b.dataset.f === filter);
  $("#scope-doc").setAttribute("aria-checked", scope === "doc");
  $("#scope-all").setAttribute("aria-checked", scope === "all");
  const grid = $("#grid"); grid.textContent = "";
  const items = filter === "custom" || filter === "community" ? [] : Object.entries(T).filter(([, v]) => filter === "all" || v.mode === filter)
    .map(([id, v]) => card(id, v.name, `previews/${id}.webp`))
    .concat(Object.entries(remote).filter(([, v]) => filter === "all" || v.mode === filter).map(([id, v]) => card(id, v.name, v.preview)));
  if (filter === "all" || filter === "community") for (const c of community) items.push(card(c.id, c.name, c.thumb));
  if (filter === "all" || filter === "custom") for (const c of customs) items.push(card(c.id, c.name, c.thumb));
  if (!items.length) {
    const p = document.createElement("p"); p.className = "empty";
    p.textContent = filter === "community" ? "No community themes yet. New ones appear here by themselves once they're approved on the website."
      : "No custom themes yet. Click Custom themes to make one from your own picture or colours.";
    items.push(p);
  }
  grid.append(...items);
  for (const c of grid.querySelectorAll(".card")) c.setAttribute("aria-checked", c.dataset.id === current());
}

const open = (hash) => chrome.tabs.create({ url: chrome.runtime.getURL("options.html#" + hash) });
$("#on").onchange = (e) => save({ on: e.target.checked });
$("#sparks").onchange = (e) => save({ sparks: e.target.checked });
for (const b of document.querySelectorAll("#motion button")) b.onclick = () => { if (b.dataset.v === "auto") chrome.storage.local.remove("perf"); save({ motion: b.dataset.v }); };   // Auto measures again
for (const b of document.querySelectorAll(".filter button")) b.onclick = () => { filter = b.dataset.f; render(); };
$("#scope-doc").onclick = () => { scope = "doc"; render(); };
$("#scope-all").onclick = () => { scope = "all"; render(); };
$("#save-file").onclick = () => open("file");
$("#customs").onclick = () => open("custom");

chrome.storage.sync.get(DEF, (v) => {
  S = { ...DEF, ...v };
  chrome.storage.local.get(["customs", "community", "remote", "catalog"], (c) => {
    customs = c.customs || []; community = c.community || []; remote = c.remote || {};
    render();
    // new themes from the website: refresh the list if it's more than an hour old
    if (!c.catalog || Date.now() - c.catalog.at > 3600e3) chrome.runtime.sendMessage({ sync: 1 }, () => chrome.storage.local.get(["community", "remote"], (n) => {
      if (chrome.runtime.lastError) return;
      community = n.community || []; remote = n.remote || {}; render();
    }));
    // ask the open Word document which file it is and whether the file carries a theme
    chrome.tabs.query({ active: true, currentWindow: true }, ([tab]) => {
      if (!tab) return;
      chrome.tabs.sendMessage(tab.id, "ww-status", (r) => {
        if (chrome.runtime.lastError || !r || !r.docKey) return;
        status = r; scope = "doc";
        if (r.google) {   // Google pages: if the scene isn't showing, copy what the extension sees so it can be fixed
          $("#gnote").hidden = false;
          $("#gclean").checked = S.docsClean !== false; $("#gclean").onchange = (e) => save({ docsClean: e.target.checked });
          for (const b of document.querySelectorAll("#gdark button")) { b.setAttribute("aria-checked", b.dataset.d === S.docsDark); b.onclick = () => { save({ docsDark: b.dataset.d }); for (const x of document.querySelectorAll("#gdark button")) x.setAttribute("aria-checked", x === b); }; }
          $("#gstate").textContent = r.shown ? "Scene is on in this Google page." : "Scene isn't showing in this Google page.";
          $("#gcopy").onclick = () => chrome.tabs.sendMessage(tab.id, "ww-details", (d) => navigator.clipboard.writeText(JSON.stringify(d, null, 1)).then(() => ($("#gstate").textContent = "Copied. Paste it in the chat."), () => ($("#gstate").textContent = "Couldn't copy.")));
        }
        $("#doc").hidden = false;
        if (r.file) {
          const f = lookup(r.file.id);
          $("#filenote").hidden = false;
          $("#filenote").textContent = `This file has the ${f ? f.name : r.file.name || "custom"} theme saved in it.`;
        }
        render();
      });
    });
  });
});
