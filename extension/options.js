const T = self.WW_THEMES, $ = (s) => document.querySelector(s);
let customs = [], docx = null, docName = "", docKind = "docx", kind = "colours", picData = null;

// ---------- tabs
function tab(name) {
  for (const b of document.querySelectorAll("[data-tab]")) b.setAttribute("aria-selected", b.dataset.tab === name);
  $("#tab-file").hidden = name !== "file"; $("#tab-custom").hidden = name !== "custom";
  history.replaceState(null, "", "#" + name);
}
for (const b of document.querySelectorAll("[data-tab]")) b.onclick = () => tab(b.dataset.tab);
const fromHash = () => tab(location.hash === "#custom" ? "custom" : "file");
addEventListener("hashchange", fromHash);
fromHash();

// ---------- theme list (built-in + yours)
function fillPick() {
  const sel = $("#pick"), keep = sel.value; sel.textContent = "";
  const group = (label, list) => {
    const g = document.createElement("optgroup"); g.label = label;
    for (const [id, name] of list) g.append(new Option(name, id));
    if (list.length) sel.append(g);
  };
  group("Dark", Object.entries(T).filter(([, t]) => t.mode === "dark").map(([id, t]) => [id, t.name]));
  group("Light", Object.entries(T).filter(([, t]) => t.mode === "light").map(([id, t]) => [id, t.name]));
  group("Yours", customs.map((c) => [c.id, c.name]));
  if (keep) sel.value = keep;
}
const themeFor = (id) => T[id] ? { ...T[id], id } : WWAssets.forCustom({ ...customs.find((c) => c.id === id) });

// ---------- save a theme into a Word file
const msg = (t) => { $("#msg").textContent = t; };
async function loadDocx(file) {
  if (!file) return;
  if (!/\.(docx|pptx|xlsx)$/i.test(file.name)) return msg("Choose a .docx, .pptx or .xlsx file. Older .doc, .ppt and .xls files need saving in the new format first.");
  docKind = WWAssets.kindOf(file.name);
  docx = new Uint8Array(await file.arrayBuffer()); docName = file.name.replace(/\.(docx|pptx|xlsx)$/i, "");
  $("#dropname").textContent = file.name; $("#stamp").disabled = false; $("#unstamp").disabled = false;
  try {
    const info = await WWAssets.stamperFor(docKind).analyze(docx);
    msg(info.existingTheme ? `This file already has the ${(T[info.existingTheme] || {}).name || "a"} theme. Picking another one replaces it.` : "");
  } catch (e) { return msg("This file couldn't be read. Is it really a Word, PowerPoint or Excel file?"), ($("#stamp").disabled = true); }
  stampNow();                                    // no extra click: the themed copy is made straight away
}
$("#docx").onchange = (e) => loadDocx(e.target.files[0]);
const drop = $("#drop");
drop.ondragover = (e) => { e.preventDefault(); drop.classList.add("over"); };
drop.ondragleave = () => drop.classList.remove("over");
drop.ondrop = (e) => { e.preventDefault(); drop.classList.remove("over"); loadDocx(e.dataTransfer.files[0]); };

const imageOf = async (t) => t.image ? (await chrome.storage.local.get("img:" + t.image))["img:" + t.image] : null;
const assetsFor = async (t, sec, k = docKind) => WWAssets.assetsFor(t, k, sec, chrome.runtime.getURL(""), await imageOf(t));
function download(bytes, name) {
  const url = URL.createObjectURL(new Blob([bytes]));
  Object.assign(document.createElement("a"), { href: url, download: name }).click();
  setTimeout(() => URL.revokeObjectURL(url), 30000);
}
async function stampNow() {
  const t = themeFor($("#pick").value);
  msg("Working…");
  try {
    const S = WWAssets.stamperFor(docKind), info = await S.analyze(docx);
    const out = await S.stamp(docx, t, await assetsFor(t, info.sections[0]));
    download(out, `${docName} (${t.name}).${docKind}`);
    msg(`Done. "${docName} (${t.name}).${docKind}" is in your downloads. Upload it to OneDrive and open it.`);
  } catch (e) { msg("Something went wrong theming this file: " + e.message); }
}
$("#stamp").onclick = stampNow;
$("#unstamp").onclick = async () => {
  try { download(await WWAssets.stamperFor(docKind).unstamp(docx), `${docName} (no theme).${docKind}`); msg(`"${docName} (no theme).${docKind}" is in your downloads.`); }
  catch (e) { msg("Couldn't remove the theme: " + e.message); }
};
$("#teamsbg").onclick = async () => {
  const t = themeFor($("#pick").value);
  try { download(await WWAssets.teams(t, chrome.runtime.getURL(""), await imageOf(t)), `${t.name} (Teams background).jpg`); msg(`"${t.name} (Teams background).jpg" is in your downloads. In a Teams meeting: Video effects › Add new.`); }
  catch (e) { msg("Couldn't make the background: " + e.message); }
};
$("#blank").onclick = async () => {
  const t = themeFor($("#pick").value);
  try { download(await WWStamper.newDocument(t, await assetsFor(t, WWStamper.NEW_DOC_SECTION, "docx")), `${t.name}.docx`); msg(`"${t.name}.docx" is in your downloads.`); }
  catch (e) { msg("Couldn't make the document: " + e.message); }
};

// ---------- custom themes
const cmsg = (t) => { $("#cmsg").textContent = t; };
for (const b of document.querySelectorAll("[data-k]")) b.onclick = () => {
  kind = b.dataset.k;
  for (const x of document.querySelectorAll("[data-k]")) x.setAttribute("aria-checked", x === b);
  $("#k-colours").hidden = kind !== "colours"; $("#k-picture").hidden = kind !== "picture"; preview();
};
document.querySelector('[data-k="colours"]').setAttribute("aria-checked", "true");

// Pictures are shrunk to 1920px and saved as JPEG: small to store, quick to draw.
$("#pic").onchange = async (e) => {
  const f = e.target.files[0]; if (!f) return;
  try {
    const bmp = await createImageBitmap(f), s = Math.min(1, 1920 / bmp.width);
    const c = Object.assign(document.createElement("canvas"), { width: bmp.width * s, height: bmp.height * s });
    const g = c.getContext("2d"); g.drawImage(bmp, 0, 0, c.width, c.height);
    picData = c.toDataURL("image/jpeg", .85);
    const px = g.getImageData(0, 0, c.width, c.height).data; let r = 0, gg = 0, b = 0, n = 0;
    for (let i = 0; i < px.length; i += 4 * 97) { r += px[i]; gg += px[i + 1]; b += px[i + 2]; n++; }
    const avg = [r / n, gg / n, b / n];
    $("#acc").value = "#" + avg.map((v) => Math.min(255, v * 1.4 + 30) | 0).map((v) => v.toString(16).padStart(2, "0")).join("");
    $("#picname").textContent = f.name; $("#k-picture").dataset.light = (avg[0] + avg[1] + avg[2]) / 3 > 150 ? "1" : "";
    preview();
  } catch (_) { cmsg("That picture couldn't be opened. Try a PNG or JPG."); }
};
function draft() {
  const name = $("#cname").value.trim() || "My theme";
  if (kind === "picture") return { name, image: 1, dim: +$("#dim").value, accent: $("#acc").value, accent2: $("#acc").value,
    mode: $("#k-picture").dataset.light ? "light" : "dark", fx: $("#cfx").value };
  const colors = [$("#c1").value, $("#c2").value, $("#c3").value];
  const light = parseInt(colors[0].slice(1, 3), 16) + parseInt(colors[0].slice(3, 5), 16) + parseInt(colors[0].slice(5, 7), 16) > 450;
  return { name, colors, particles: $("#part").value, accent: colors[1], accent2: colors[2], mode: light ? "light" : "dark", fx: $("#cfx").value };
}
function preview() {
  const d = draft(), p = $("#prev");
  p.style.backgroundImage = kind === "picture" ? (picData ? `linear-gradient(rgba(0,0,0,${d.dim}),rgba(0,0,0,${d.dim})),url(${picData})` : "none")
    : `radial-gradient(120% 90% at 0 0,${d.colors[1]},transparent 60%),radial-gradient(120% 90% at 100% 100%,${d.colors[2]},transparent 60%)`;
  p.style.backgroundColor = kind === "picture" ? "#222" : d.colors[0];
}
for (const el of document.querySelectorAll("#c1,#c2,#c3,#dim,#cname")) el.oninput = preview;

async function thumb(d) {
  const c = Object.assign(document.createElement("canvas"), { width: 192, height: 108 }), g = c.getContext("2d");
  if (d.image && picData) {
    const bmp = await createImageBitmap(await (await fetch(picData)).blob()), s = Math.max(192 / bmp.width, 108 / bmp.height);
    g.drawImage(bmp, (192 - bmp.width * s) / 2, (108 - bmp.height * s) / 2, bmp.width * s, bmp.height * s);
    g.fillStyle = `rgba(0,0,0,${d.dim})`; g.fillRect(0, 0, 192, 108);
  } else {
    g.fillStyle = d.colors[0]; g.fillRect(0, 0, 192, 108);
    for (const [x, y, k] of [[0, 0, 1], [192, 108, 2]]) { const r = g.createRadialGradient(x, y, 0, x, y, 200); r.addColorStop(0, d.colors[k]); r.addColorStop(1, d.colors[k] + "00"); g.fillStyle = r; g.fillRect(0, 0, 192, 108); }
  }
  return c.toDataURL("image/jpeg", .8);
}
async function addCustom(d, imageData) {
  const id = "c_" + Date.now().toString(36);
  const t = { ...d, id, thumb: d.thumb || await thumb(d) };
  if (d.image) { t.image = id; await chrome.storage.local.set({ ["img:" + id]: imageData }); }
  customs = [...customs, t];
  await chrome.storage.local.set({ customs });
  renderMine(); fillPick();
  return t;
}
$("#csave").onclick = async () => {
  if (kind === "picture" && !picData) return cmsg("Choose a picture first.");
  const t = await addCustom(draft(), picData);
  cmsg(`Saved "${t.name}". Pick it from the popup in Word, or save it into a file.`);
};

function renderMine() {
  const ul = $("#mine"); ul.textContent = "";
  if (!customs.length) { const li = document.createElement("li"); li.textContent = "None yet."; ul.append(li); return; }
  for (const c of customs) {
    const li = document.createElement("li"), img = new Image(), b = document.createElement("b");
    img.src = c.thumb; img.alt = ""; b.textContent = c.name;
    const exp = document.createElement("button"); exp.textContent = "Export";
    exp.onclick = async () => {
      const data = c.image ? (await chrome.storage.local.get("img:" + c.image))["img:" + c.image] : null;
      const url = URL.createObjectURL(new Blob([JSON.stringify({ wildtheme: 1, theme: { ...c, id: undefined }, image: data })], { type: "application/json" }));
      Object.assign(document.createElement("a"), { href: url, download: c.name + ".wildtheme" }).click();
    };
    const del = document.createElement("button"); del.textContent = "Delete";
    del.onclick = async () => {
      customs = customs.filter((x) => x !== c);
      await chrome.storage.local.set({ customs });
      if (c.image) await chrome.storage.local.remove("img:" + c.image);
      renderMine(); fillPick();
    };
    li.append(img, b, exp, del); ul.append(li);
  }
}
$("#imp").onchange = async (e) => {
  try {
    const j = JSON.parse(await e.target.files[0].text());
    if (!j.wildtheme || !j.theme || !j.theme.name) throw 0;
    picData = j.image || null;
    const t = await addCustom(j.theme, j.image);
    cmsg(`Imported "${t.name}".`);
  } catch (_) { cmsg("That isn't a Wild Themes theme file."); }
};

chrome.storage.local.get("customs", (c) => { customs = c.customs || []; fillPick(); renderMine(); preview(); });
