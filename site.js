// Wild Themes site. The gallery loads first; file tools load only when used;
// one shared iframe plays hover previews, so idle pages do no work.
const $ = (s, r = document) => r.querySelector(s), $$ = (s, r = document) => [...r.querySelectorAll(s)];
const REPO = "https://github.com/Min3scon/Wild-Themes", EXT = new URL("extension/", location.href).href;
const store = { get: (k, d) => { try { return localStorage.getItem(k) ?? d; } catch (_) { return d; } }, set: (k, v) => { try { localStorage.setItem(k, v); } catch (_) {} } };
const reduced = matchMedia("(prefers-reduced-motion: reduce)").matches;
let themes = [], filter = "all", query = "", motion = store.get("motion", reduced ? "0" : "1") === "1";

for (const [id, t] of Object.entries(self.WW_THEMES)) themes.push({ ...t, id, kind: "builtin", author: "Wild Themes", thumb: `previews/${id}.webp` });
fetch("community/index.json").then((r) => (r.ok ? r.json() : { themes: [] })).then((j) => {
  for (const c of j.themes || []) themes.push({ ...c, kind: "community", blurb: c.image ? "A picture theme from the community." : "A colour theme from the community." });
  render();
}).catch(() => {});

// ---------- gallery
const FILTERS = [["all", "All"], ["dark", "Dark"], ["light", "Light"], ["community", "Community"]];
const inFilter = (t, f) => f === "all" || (f === "community" ? t.kind === "community" : t.mode === f);
const inQuery = (t) => !query || [t.name, t.blurb, t.author, t.mode, t.kind].join(" ").toLowerCase().includes(query);
const el = (tag, cls, text) => { const n = document.createElement(tag); if (cls) n.className = cls; if (text != null) n.textContent = text; return n; };
const pills = (t) => [
  [motion && t.motion !== false ? "Animated" : "Still", true],
  [t.mode === "light" ? "Light" : "Dark", false, t.accent],
  [t.kind === "community" ? "Community" : "Official", false],
].map(([label, on, dot]) => { const p = el("span", "pill" + (on ? " on" : ""), label); if (dot) { const i = el("i"); i.style.background = dot; p.prepend(i); } return p; });

function render() {
  const g = $("#filters");
  g.replaceChildren(...FILTERS.map(([f, label]) => {
    const b = el("button", "btn btn-sm", label + " ");
    b.append(el("small", null, themes.filter((t) => inFilter(t, f)).length));
    b.setAttribute("aria-pressed", f === filter);
    b.onclick = () => { filter = f; render(); };
    return b;
  }));
  $$("#motion .btn").forEach((b) => b.setAttribute("aria-pressed", (b.dataset.m === "1") === motion));
  const list = themes.filter((t) => inFilter(t, filter) && inQuery(t));
  $("#grid").replaceChildren(...list.map(card));
  $("#empty").hidden = list.length > 0;
  $("#empty p").textContent = filter === "community" && !query ? "No community themes yet. Make one and share it." : `No themes match “${query}”.`;
}
function card(t) {
  const c = el("article", "card"), shot = el("button", "shot"), img = new Image(), body = el("div", "body");
  shot.setAttribute("aria-label", "Open " + t.name);
  img.src = t.thumb; img.alt = ""; img.loading = "lazy"; img.decoding = "async";
  shot.append(img);
  const row = el("div", "row"), use = el("button", "btn btn-outline btn-sm", "Use theme"), doc = el("button", "btn btn-ghost btn-sm", "Word starter");
  use.onclick = shot.onclick = () => openTheme(t);
  doc.onclick = () => quick(t, "starter", doc);
  row.append(use, doc);
  const pl = el("div", "pills"); pl.append(...pills(t));
  body.append(el("h3", null, t.name), el("p", "by", "by " + (t.author || "the community")), pl, el("p", "d", t.blurb), row);
  c.append(shot, body);
  shot.onpointerenter = () => peek(t, shot);
  shot.onpointerleave = unpeek;
  return c;
}
$("#q").oninput = (e) => { query = e.target.value.trim().toLowerCase(); render(); };
$$("#motion .btn").forEach((b) => b.onclick = () => { motion = b.dataset.m === "1"; store.set("motion", b.dataset.m); render(); hero(); });
addEventListener("keydown", (e) => { if (e.key === "/" && !/INPUT|TEXTAREA|SELECT/.test(document.activeElement.tagName)) { e.preventDefault(); $("#q").focus(); } });

// ---------- live scenes: the extension's own CSS inside an iframe sized like its target
function sceneDoc(t, win) {
  const custom = t.kind !== "builtin", c = t.colors || ["#111", t.accent, t.accent2];
  const back = custom && t.image ? `<div style="position:fixed;inset:0;background:url('${t.image}') center/cover"></div>`
    : `<div id="ww-bg" data-scene="${custom ? "custom" : t.id}"${custom && t.effect && t.effect !== "none" ? ` data-part="${t.effect}"` : ""} class="${motion ? "" : "still"}"
      style="--ww-c1:${c[0]};--ww-c2:${c[1]};--ww-c3:${c[2]}"><i></i><i></i><i></i><i></i></div>`;
  const page = win ? `<div class="rb"><s></s><s></s><s class="on"></s><s></s><s></s></div><div class="pg"><b class="h"></b><b class="t"></b><b></b><b></b><b class="s"></b><b></b><b class="t"></b><b></b><b class="s"></b><b></b></div>` : "";
  return `<!doctype html><html><head><link rel="stylesheet" href="extension/content.css">${custom ? "" : `<link rel="stylesheet" href="extension/scenes/${t.id}.css">`}<style>
html,body{margin:0;height:100%;overflow:hidden;background:transparent}
.rb{position:fixed;inset:0 0 auto;height:10%;background:#f3f2f1;display:flex;gap:3%;align-items:center;padding:0 4%;box-shadow:0 1px 0 #e1dfdd}
.rb s{width:7%;height:18%;border-radius:4px;background:#c8c6c4}.rb .on{background:${t.accent}}
.pg{position:fixed;left:50%;top:16%;bottom:-6%;width:46%;transform:translateX(-50%);background:#fff;padding:5.5% 6.5%;display:grid;gap:4.6%;align-content:start;
box-shadow:0 0 0 1px ${t.accent}99,0 0 28px ${t.accent}66}
.pg b{height:6px;border-radius:3px;background:#dedce4}.pg .t{width:82%}.pg .h{height:14px;width:58%;background:#2c2840;margin-bottom:2%}.pg .s{height:8px;width:36%;background:${t.accent};margin-top:3%}
</style></head><body>${back}${page}</body></html>`;
}
const peekEl = $("#peek");
let peekTimer = 0;
function peek(t, target) {
  if (!motion || reduced || (t.kind !== "builtin" && t.image)) return;
  clearTimeout(peekTimer);
  peekTimer = setTimeout(() => {                         // short delay: no work when the pointer only passes over
    const r = target.getBoundingClientRect();
    Object.assign(peekEl.style, { left: r.left + scrollX + "px", top: r.top + scrollY + "px", width: r.width + "px", height: r.height + "px" });
    peekEl.classList.remove("on");
    peekEl.onload = () => peekEl.classList.add("on");
    peekEl.srcdoc = sceneDoc(t, false);
  }, 140);
}
function unpeek() { clearTimeout(peekTimer); peekEl.classList.remove("on"); }

// hero: a framed live preview that steps through the themes while it's on screen
const heroIds = ["synthwave", "aurora", "sakura", "winter", "lava", "goldenhour", "matrix", "ocean"];
let heroIx = 0, heroTimer = 0, heroVisible = true;
function hero() {
  const t = themes.find((x) => x.id === heroIds[heroIx % heroIds.length]);
  if (!t) return;
  $("#hero-stage").srcdoc = sceneDoc(t, true);
  $("#hero-name").textContent = t.name;
  clearTimeout(heroTimer);
  if (motion && !reduced) heroTimer = setTimeout(() => { if (heroVisible && !document.hidden) heroIx++; hero(); }, 6000);
}
new IntersectionObserver(([e]) => { heroVisible = e.isIntersecting; }).observe($("#hero-stage"));

// ---------- dialogs
const dlg = $("#dlg");
function openDialog(tpl) {
  unpeek();
  const body = $("#dlg-body");
  body.replaceChildren($(tpl).content.cloneNode(true));
  dlg.showModal();
  body.scrollTop = 0;
  return body;
}
$("#close").onclick = () => dlg.close();
dlg.addEventListener("click", (e) => { if (e.target === dlg) dlg.close(); });     // click outside closes

// file tools: ~200 KB the gallery never needs, loaded on first use
let tools = null;
const script = (src) => new Promise((ok, no) => document.head.append(Object.assign(document.createElement("script"), { src, onload: ok, onerror: no })));
const loadTools = () => tools ||= ["vendor/jszip.min.js", "lib/stamper.js", "lib/stamper-pptx.js", "lib/stamper-xlsx.js", "lib/assets.js", "lib/wtcore.js"]
  .reduce((p, f) => p.then(() => script(EXT + f)), Promise.resolve());
// The Office add-in is this page inside Word, PowerPoint or Excel (addin/office-host.js).
let office = null;
if (self.WTOffice) WTOffice.ready.then((o) => { office = o; });
async function save(data, name, type) {
  if (office && office.host === "word" && /\.docx$/.test(name)) {   // a new Word document opens in Word rather than downloading
    return office.openCopy(new Uint8Array(await new Blob([data]).arrayBuffer()), "docx");
  }
  const url = URL.createObjectURL(data instanceof Blob ? data : new Blob([data], { type: type || "application/octet-stream" }));
  Object.assign(document.createElement("a"), { href: url, download: name }).click();
  setTimeout(() => URL.revokeObjectURL(url), 30000);
}
const toDataURL = (blob) => new Promise((ok) => { const r = new FileReader(); r.onload = () => ok(r.result); r.readAsDataURL(blob); });
async function forFiles(t) {
  if (t.kind === "builtin") return { th: { ...self.WW_THEMES[t.id], id: t.id }, img: null };
  const th = WWAssets.forCustom({ id: t.id, name: t.name, mode: t.mode, accent: t.accent, accent2: t.accent2, colors: t.colors, particles: t.effect, image: t.image ? 1 : 0 });
  return { th, img: t.image ? (t.image.startsWith("data:") ? t.image : await toDataURL(await (await fetch(t.image)).blob())) : null };
}
async function stillify(a) {                             // Still: first frame as PNG instead of an animated GIF
  if (motion || !a.backdrop || a.backdrop.ext !== "gif") return a;
  const bmp = await createImageBitmap(new Blob([a.backdrop.bytes], { type: "image/gif" }));
  const c = Object.assign(document.createElement("canvas"), { width: bmp.width, height: bmp.height });
  c.getContext("2d").drawImage(bmp, 0, 0);
  return { ...a, backdrop: { bytes: new Uint8Array(await (await new Promise((ok) => c.toBlob(ok, "image/png"))).arrayBuffer()), ext: "png" } };
}
// Same core as the extension and the add-in: the file gets the full embedded definition, backup and checks.
async function themeFile(t, file, say) {
  if (!file) return;
  if (!/\.(docx|docm)$/i.test(file.name)) return say("Use a Word file (.docx).", "bad");
  say("Adding the theme…");
  try {
    await loadTools();
    const m = /^(.*?)(\.[^.]+)$/.exec(file.name), { th, img } = await forFiles(t);
    const out = await WTCore.apply(new Uint8Array(await file.arrayBuffer()), th, { animated: motion }, WWAssets.maker(EXT, async () => img), "Wild Themes website");
    const check = await WTCore.inspect(out);
    if (check.status !== "intact") throw new Error(check.problems.map((p) => p.text).join(" "));
    save(out, `${m[1]} (${t.name})${m[2]}`);
    say(`Saved “${m[1]} (${t.name})${m[2]}”. The theme is inside it and checked.`, "ok");
  } catch (e) { say("That file couldn't be themed: " + e.message, "bad"); }
}
// In the add-in the theme goes into the open document (checked after), not into a download.
let docBusy = false;
async function themeDoc(t, say) {
  if (docBusy) return;
  docBusy = true; say("Adding the theme…");
  try {
    await loadTools();
    const { th, img } = await forFiles(t), before = await office.read(), was = await WTCore.inspect(before);
    const out = await WTCore.apply(before, th, { ...(was.settings || {}), animated: motion }, WWAssets.maker(EXT, async () => img), "Wild Themes add-in");
    say(office.canSwap() ? "Saving it into this file…" : "Opening the result in a new window…");
    const res = await office.commit(out);
    say(...office.report(res, res.copied ? null : await WTCore.inspect(res.bytes), th.id, `${t.name} is saved in the file.`));
  } catch (e) { say("That file couldn't be themed: " + ((e && e.message) || e), "bad"); }
  finally { docBusy = false; }
}
function docDrop(drop, go) {
  $("input", drop).remove();
  $("b", drop).textContent = `Save it into this ${office.where()}`;
  $("span", drop).textContent = office.canSwap() ? "Written into the open file, then checked. Nothing is uploaded." : "Opens as a themed copy for you to save. Nothing is uploaded.";
  drop.tabIndex = 0; drop.setAttribute("role", "button");
  drop.onclick = go;
  drop.onkeydown = (e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); go(); } };
}
function wireDrop(drop, onFile) {
  const input = $("input", drop);
  input.onchange = () => { onFile(input.files[0]); input.value = ""; };
  drop.ondragover = (e) => { e.preventDefault(); drop.classList.add("over"); };
  drop.ondragleave = () => drop.classList.remove("over");
  drop.ondrop = (e) => { e.preventDefault(); drop.classList.remove("over"); onFile(e.dataTransfer.files[0]); };
}
const sayer = (m) => (text, cls = "") => { m.textContent = text; m.className = "msg " + cls; };
const actions = {
  async starter(t, say) {
    await loadTools(); const { th, img } = await forFiles(t);
    const blank = await WWStamper.newDocument(th, await stillify(await WWAssets.assetsFor(th, "docx", WWStamper.NEW_DOC_SECTION, EXT, img)));
    // full core on top: definition, fingerprints and backup inside, like every other themed file
    const doc = await WTCore.apply(new Uint8Array(await new Blob([blank]).arrayBuffer()), th, { animated: motion }, WWAssets.maker(EXT, async () => img), "Wild Themes website");
    const opened = await save(doc, `${t.name}.docx`);
    say && say(opened ? `“${t.name}” opened as a new document.` : `Saved “${t.name}.docx”.`, "ok");
  },
  async teams(t, say) {
    await loadTools(); const { th, img } = await forFiles(t);
    save(await WWAssets.teams(th, EXT, img), `${t.name} (Teams).jpg`, "image/jpeg"); say("Saved. In a Teams meeting: Video effects › Add new.", "ok");
  },
  async ext(t, say) {
    // Wild Themes for Word (the desktop .exe) listens on this PC only; it takes built-in scenes.
    const desk = t.kind === "builtin" ? fetch(`http://127.0.0.1:47615/theme?id=${encodeURIComponent(t.id)}`).then(() => true, () => false) : Promise.resolve(false);
    // With the extension installed, its site bridge switches the theme right away.
    if (!document.documentElement.dataset.wildThemesExtension) {
      if (await desk) return say(`Done. Desktop Word now shows ${t.name} around the page.`, "ok");
      say("Add the extension first, then press this again.", "bad");
      return setTimeout(openInstall, 900);
    }
    let theme = { builtin: true, id: t.id, name: t.name };
    if (t.kind !== "builtin") {
      await loadTools(); const { th, img } = await forFiles(t);
      theme = { key: t.id, name: th.name, mode: th.mode, accent: th.accent, accent2: th.accent2, colors: th.colors, effect: th.particles, image: img, thumb: await thumbOf(th, img) };
    }
    const done = new Promise((ok) => {
      const on = (e) => { if (e.source === window && e.data && e.data.wt === "used") { removeEventListener("message", on); ok(e.data.ok); } };
      addEventListener("message", on);
      setTimeout(() => { removeEventListener("message", on); ok(null); }, 3000);
    });
    postMessage({ wt: "use", theme }, location.origin);
    const r = await done;
    say(r ? `Done. Wild Themes now uses ${t.name}. Open or switch to a document to see it.` : r === false ? "The extension couldn't take that theme." : "The extension didn't answer. Reload this page and try again.", r ? "ok" : "bad");
  },
};
// small card picture for the extension popup
async function thumbOf(th, img) {
  const c = Object.assign(document.createElement("canvas"), { width: 192, height: 108 }), g = c.getContext("2d");
  if (img) { const b = await createImageBitmap(await (await fetch(img)).blob()), s = Math.max(192 / b.width, 108 / b.height); g.drawImage(b, (192 - b.width * s) / 2, (108 - b.height * s) / 2, b.width * s, b.height * s); }
  else {
    const col = th.colors || ["#111", th.accent, th.accent2]; g.fillStyle = col[0]; g.fillRect(0, 0, 192, 108);
    for (const [x, y, k] of [[0, 0, 1], [192, 108, 2]]) { const r = g.createRadialGradient(x, y, 0, x, y, 200); r.addColorStop(0, col[k]); r.addColorStop(1, col[k] + "00"); g.fillStyle = r; g.fillRect(0, 0, 192, 108); }
  }
  return c.toDataURL("image/jpeg", .8);
}
async function quick(t, a, btn) {
  const was = btn.textContent; btn.textContent = "Making…"; btn.disabled = true;
  try { await actions[a](t); btn.textContent = "Saved"; } catch (_) { btn.textContent = "Try again"; }
  setTimeout(() => { btn.textContent = was; btn.disabled = false; }, 1600);
}

function openTheme(t) {
  const body = openDialog("#t-theme"), say = sayer($(".msg", body));
  $("h3", body).textContent = t.name;
  $(".by", body).textContent = "by " + (t.author || "the community");
  $(".pills", body).append(...pills(t));
  $(".desc", body).textContent = t.blurb;
  $(".stage", body).srcdoc = sceneDoc(t, true);
  if (office) docDrop($(".drop", body), () => themeDoc(t, say));
  else wireDrop($(".drop", body), (f) => themeFile(t, f, say));
  const st = $("[data-studio]", body); if (st && t.kind !== "local") st.href = "studio.html?theme=" + encodeURIComponent(t.id);
  $$("[data-a]", body).forEach((b) => b.onclick = () => {
    actions[b.dataset.a](t, say).catch((e) => say(e.message, "bad"));
  });
}

function openInstall() {
  const body = openDialog("#t-install");
  $$(".copy", body).forEach((b) => b.onclick = async () => { try { await navigator.clipboard.writeText(b.dataset.copy); b.textContent = "Copied"; } catch (_) { b.textContent = "Select it"; } });
  // show the steps for the browser being used; the other one is a click away
  const pick = (name) => { $$("[data-browser]", body).forEach((d) => (d.hidden = d.dataset.browser !== name)); $$("[data-b]", body).forEach((b) => b.setAttribute("aria-pressed", b.dataset.b === name)); };
  $$("[data-b]", body).forEach((b) => (b.onclick = () => pick(b.dataset.b)));
  pick(/firefox/i.test(navigator.userAgent) && !/seamonkey/i.test(navigator.userAgent) ? "firefox" : "chrome");
}

// ---------- your own theme: built and used locally; shared only if you press Share
function openMaker(share) {
  const body = openDialog("#t-maker"), say = sayer($("#m-msg", body));
  const t = { id: "local", kind: "local", name: "My theme", author: "you", mode: "dark", accent: "#ff2bd6", accent2: "#21e6ff", colors: null, effect: "none", image: null, blurb: "" };
  let kind = "picture", blob = null;
  const stage = $("#m-stage", body), seg = $$("[data-k]", body);
  const colours = () => [$("#m-c1", body).value, $("#m-c2", body).value, $("#m-c3", body).value];
  const sync = () => {
    if (kind === "colours") {
      const c = colours(); Object.assign(t, { colors: c, accent: c[1], accent2: c[2], image: null, effect: $("#m-fx", body).value,
        mode: parseInt(c[0].slice(1, 3), 16) + parseInt(c[0].slice(3, 5), 16) + parseInt(c[0].slice(5, 7), 16) > 450 ? "light" : "dark" });
    }
    stage.srcdoc = sceneDoc(t, true);
  };
  const setKind = (k) => { kind = k; seg.forEach((b) => b.setAttribute("aria-pressed", b.dataset.k === k)); $$("[data-for]", body).forEach((d) => d.hidden = d.dataset.for !== k); if (k === "picture" && !t.image) t.colors = null; sync(); };
  seg.forEach((b) => b.onclick = () => setKind(b.dataset.k));
  for (const id of ["#m-c1", "#m-c2", "#m-c3", "#m-fx"]) $(id, body).oninput = sync;
  $("#m-pic", body).onchange = async (e) => {
    const f = e.target.files[0]; if (!f) return;
    try {
      const bmp = await createImageBitmap(f);
      // same budget as the built-in themes: 1920×1080 at most, under 450 KB
      const s = Math.min(1, 1920 / bmp.width, 1080 / bmp.height), w = Math.round(bmp.width * s), h = Math.round(bmp.height * s);
      const c = Object.assign(document.createElement("canvas"), { width: w, height: h }), g = c.getContext("2d");
      g.drawImage(bmp, 0, 0, w, h);
      let q = .86; do { blob = await new Promise((ok) => c.toBlob(ok, "image/jpeg", q)); q -= .07; } while (blob.size > 450e3 && q > .35);
      const px = g.getImageData(0, 0, w, h).data; let r = 0, gg = 0, b = 0, n = 0;
      for (let i = 0; i < px.length; i += 4 * 199) { r += px[i]; gg += px[i + 1]; b += px[i + 2]; n++; }
      Object.assign(t, { image: await toDataURL(blob), colors: null, mode: (r + gg + b) / n / 3 > 140 ? "light" : "dark",
        accent: "#" + [r, gg, b].map((v) => Math.min(255, v / n * 1.35 + 25) | 0).map((v) => v.toString(16).padStart(2, "0")).join("") });
      t.accent2 = t.accent;
      say(`Ready: ${w}×${h}, ${Math.round(blob.size / 1024)} KB. It stays on this computer.`, "ok");
      sync();
    } catch (_) { say("That picture couldn't be opened. Try a PNG or JPG.", "bad"); }
  };
  const ready = () => kind === "colours" || t.image || (say("Choose a picture first, or switch to Colours.", "bad"), false);
  if (office) docDrop($("#m-drop", body), () => ready() && themeDoc({ ...t, name: $("#m-name", body).value.trim() || t.name }, say));
  else wireDrop($("#m-drop", body), (f) => ready() && themeFile({ ...t, name: $("#m-name", body).value.trim() || t.name }, f, say));
  $("#m-ext", body).onclick = () => ready() && actions.ext({ ...t, id: "local-" + Date.now().toString(36), name: $("#m-name", body).value.trim() || t.name }, say);
  $("#m-teams", body).onclick = () => ready() && actions.teams(t, say);
  $("#m-go", body).onclick = () => {
    const name = $("#m-name", body).value.trim().slice(0, 30);
    if (!name) return say("Give it a name first.", "bad"), $("#m-name", body).focus();
    if (!ready()) return;
    const slug = name.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "") || "theme";
    const p = new URLSearchParams({ template: "theme.yml", title: "Theme: " + name, name, mode: t.mode === "light" ? "Light" : "Dark", accent: t.accent });
    if (kind === "picture") { save(blob, slug + ".jpg"); p.set("kind", "Picture (still)"); }
    else { const fx = $("#m-fx", body); p.set("kind", "Colours (animated)"); p.set("colours", t.colors.join(" ")); p.set("effect", fx.options[fx.selectedIndex].text.replace("No effect", "None")); }
    open(`${REPO}/issues/new?${p}`, "_blank", "noopener");
    say(kind === "picture" ? "The picture is in your downloads. Drag it into the Picture box on GitHub, then press Create." : "Press Create on GitHub to send it.", "ok");
  };
  setKind("picture");
  if (share) setTimeout(() => $("#m-share", body).scrollIntoView({ block: "nearest" }), 50);
}

const opener = (b) => b.dataset.open === "install" ? openInstall() : b.dataset.open === "desktop" ? openDesktop() : openMaker(!!b.dataset.share);
$$("[data-open]").forEach((b) => b.onclick = () => opener(b));
function openDesktop() { $$("[data-open]", openDialog("#t-desktop")).forEach((b) => b.onclick = () => opener(b)); }
render(); hero();
