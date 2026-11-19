// Wild Themes site: gallery, live preview, and theming files right in the browser.
// The file never leaves the computer: all the work happens on this page.
const T = self.WW_THEMES, $ = (s) => document.querySelector(s), BASE = new URL("extension/", location.href).href;
let pick = "synthwave", filter = "all";

function select(id) {
  pick = id;
  const t = T[id];
  $("#ww-bg").dataset.scene = id;
  $("#scene-css").href = `extension/scenes/${id}.css`;
  document.documentElement.style.setProperty("--a", t.accent);
  document.documentElement.style.setProperty("--b", t.accent2);
  $("#now").textContent = t.name; $("#pickname").textContent = t.name;
  for (const c of document.querySelectorAll(".card")) c.setAttribute("aria-checked", c.dataset.id === id);
}

function grid() {
  const g = $("#grid"); g.textContent = "";
  for (const [id, t] of Object.entries(T)) {
    if (filter !== "all" && t.mode !== filter) continue;
    const b = document.createElement("button"), img = new Image(), d = document.createElement("div");
    b.className = "card"; b.dataset.id = id; b.setAttribute("role", "radio"); b.setAttribute("aria-checked", id === pick);
    img.src = `extension/previews/${id}.webp`; img.alt = ""; img.loading = "lazy";
    const n = document.createElement("b"), s = document.createElement("span"), m = document.createElement("small");
    n.textContent = t.name; s.textContent = t.blurb; m.textContent = t.mode;
    d.append(m, n, s); b.append(img, d);
    b.onclick = () => select(id);
    g.append(b);
  }
}
for (const b of document.querySelectorAll(".chips button")) b.onclick = () => {
  filter = b.dataset.f;
  for (const x of document.querySelectorAll(".chips button")) x.setAttribute("aria-checked", x === b);
  grid();
};

// ---------- theme a file: drop it and the themed copy downloads straight away
const msg = (t) => { $("#msg").textContent = t; };
function save(bytes, name) {
  const url = URL.createObjectURL(new Blob([bytes]));
  Object.assign(document.createElement("a"), { href: url, download: name }).click();
  setTimeout(() => URL.revokeObjectURL(url), 30000);
}
async function themeFile(file, remove) {
  if (!file) return;
  if (!/\.(docx|pptx|xlsx)$/i.test(file.name)) return msg("Use a .docx, .pptx or .xlsx file. Older .doc, .ppt and .xls files need saving in the new format first.");
  const kind = WWAssets.kindOf(file.name), S = WWAssets.stamperFor(kind), base = file.name.replace(/\.\w+$/, "");
  const t = { ...T[pick], id: pick };
  msg(remove ? "Removing the theme…" : `Adding ${t.name}…`);
  try {
    const bytes = new Uint8Array(await file.arrayBuffer());
    if (remove) { save(await S.unstamp(bytes), `${base} (no theme).${kind}`); return msg(`Done. "${base} (no theme).${kind}" is in your downloads.`); }
    const info = await S.analyze(bytes);
    save(await S.stamp(bytes, t, await WWAssets.assetsFor(t, kind, info.sections[0], BASE)), `${base} (${t.name}).${kind}`);
    msg(`Done. "${base} (${t.name}).${kind}" is in your downloads. Open it anywhere: the theme is inside.`);
  } catch (e) { msg("That file couldn't be themed: " + e.message); }
}
$("#in").onchange = (e) => themeFile(e.target.files[0]);
const drop = $("#drop");
drop.ondragover = (e) => { e.preventDefault(); drop.classList.add("over"); };
drop.ondragleave = () => drop.classList.remove("over");
drop.ondrop = (e) => { e.preventDefault(); drop.classList.remove("over"); themeFile(e.dataTransfer.files[0]); };
$("#remove").onclick = () => $("#rm").click();
$("#rm").onchange = (e) => themeFile(e.target.files[0], true);
$("#blank").onclick = async () => {
  const t = { ...T[pick], id: pick };
  try { save(await WWStamper.newDocument(t, await WWAssets.assetsFor(t, "docx", WWStamper.NEW_DOC_SECTION, BASE)), `${t.name}.docx`); msg(`"${t.name}.docx" is in your downloads.`); }
  catch (e) { msg("Couldn't make the document: " + e.message); }
};
$("#teams").onclick = async () => {
  const t = { ...T[pick], id: pick };
  try { save(await WWAssets.teams(t, BASE), `${t.name} (Teams background).jpg`); msg("Saved. In a Teams meeting: Video effects › Add new."); }
  catch (e) { msg("Couldn't make the background: " + e.message); }
};

grid(); select(pick);
