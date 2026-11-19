// Wild Themes site. Light on purpose: the gallery loads first; the file tools load
// only when someone drops a file, and one shared iframe plays hover previews.
const $ = (s, r = document) => r.querySelector(s), $$ = (s, r = document) => [...r.querySelectorAll(s)];
const REPO = "https://github.com/Min3scon/Wild-Themes", EXT = new URL("extension/", location.href).href;
const store = { get: (k, d) => { try { return localStorage.getItem(k) ?? d; } catch (_) { return d; } }, set: (k, v) => { try { localStorage.setItem(k, v); } catch (_) {} } };
let themes = [], filter = "all", query = "", motion = store.get("motion", "1") === "1";

// ---------- data: built-in themes + approved community themes
for (const [id, t] of Object.entries(self.WW_THEMES)) themes.push({ ...t, id, kind: "builtin", thumb: `extension/previews/${id}.webp` });
fetch("community/index.json").then((r) => r.ok ? r.json() : { themes: [] }).then((j) => {
  for (const c of j.themes || []) themes.push({ ...c, kind: "community", blurb: c.author ? `by ${c.author}` : "Community theme" });
  render();
}).catch(() => {});

// ---------- gallery
function matches(t) {
  if (filter === "community" ? t.kind !== "community" : filter !== "all" && t.mode !== filter) return false;
  if (!query) return true;
  return [t.name, t.blurb, t.author, t.mode, t.kind].join(" ").toLowerCase().includes(query);
}
function render() {
  const list = themes.filter(matches), grid = $("#grid");
  grid.replaceChildren(...list.map((t) => {
    const b = document.createElement("button"), th = document.createElement("div"), img = new Image(), l = document.createElement("div");
    b.className = "card"; b.dataset.id = t.id;
    th.className = "thumb"; img.src = t.thumb; img.alt = ""; img.loading = "lazy"; img.decoding = "async"; th.append(img);
    const n = document.createElement("b"), m = document.createElement("span");
    n.textContent = t.name; m.textContent = t.kind === "community" ? (t.author ? "by " + t.author : "community") : t.mode;
    l.className = "label"; l.append(n, m);
    b.append(th, l);
    b.onclick = () => openTheme(t);
    b.onpointerenter = () => peek(t, th);
    b.onpointerleave = () => unpeek();
    return b;
  }));
  const e = $("#empty");
  e.hidden = list.length > 0;
  e.textContent = filter === "community" && !query ? "No community themes yet. Be the first: click Submit." : `Nothing matches “${query}”.`;
}
$$(".filters button").forEach((b) => b.onclick = () => {
  filter = b.dataset.f; $$(".filters button").forEach((x) => x.setAttribute("aria-checked", x === b)); render();
});
$$(".motion-all button").forEach((b) => {
  b.setAttribute("aria-checked", (b.dataset.m === "1") === motion);
  b.onclick = () => { motion = b.dataset.m === "1"; store.set("motion", b.dataset.m); $$(".motion-all button").forEach((x) => x.setAttribute("aria-checked", x === b)); };
});
$("#q").oninput = (e) => { query = e.target.value.trim().toLowerCase(); render(); };
addEventListener("keydown", (e) => {
  if (e.key === "/" && document.activeElement.tagName !== "INPUT") { e.preventDefault(); $("#q").focus(); }
  if (e.key === "Escape") closeSheet();
});
addEventListener("scroll", () => $(".bar").classList.toggle("scrolled", scrollY > 4), { passive: true });

// ---------- live scenes, drawn by the extension's own CSS inside an iframe sized like the target
function sceneDoc(t, withWindow) {
  const custom = t.kind === "community";
  const c = t.colors || ["#111", t.accent, t.accent2];
  const bg = custom && t.image ? `<div style="position:fixed;inset:0;background:url(${t.image}) center/cover"></div>` :
    `<div id="ww-bg" data-scene="${custom ? "custom" : t.id}"${custom && t.effect && t.effect !== "none" ? ` data-part="${t.effect}"` : ""}
      style="--ww-c1:${c[0]};--ww-c2:${c[1]};--ww-c3:${c[2]}" class="${motion ? "" : "still"}"><i></i><i></i><i></i><i></i></div>`;
  const win = withWindow ? `<div class="rb"><s></s><s></s><s class="on"></s><s></s></div><div class="pg"><b class="h"></b><b></b><b></b><b class="s"></b><b></b><b></b><b class="s"></b><b></b></div>` : "";
  return `<!doctype html><html><head><link rel="stylesheet" href="extension/content.css">${custom ? "" : `<link rel="stylesheet" href="extension/scenes/${t.id}.css">`}
<style>html,body{margin:0;height:100%;background:transparent;overflow:hidden}
.rb{position:fixed;inset:0 0 auto;height:11%;background:#f3f2f1;display:flex;gap:4%;align-items:center;padding:0 4%}
.rb s{width:9%;height:16%;border-radius:4px;background:#c8c6c4}.rb .on{background:${t.accent}}
.pg{position:fixed;left:50%;top:17%;bottom:-4%;width:48%;transform:translateX(-50%);background:#fff;padding:5% 6%;display:grid;gap:5%;align-content:start;
box-shadow:0 0 0 1px ${t.accent}88,0 0 24px ${t.accent}55}
.pg b{height:5px;border-radius:3px;background:#dcdae2}.pg .h{height:12px;width:60%;background:#3b3355}.pg .s{height:7px;width:38%;background:${t.accent};opacity:.8}</style>
</head><body>${bg}${win}</body></html>`;
}
const peekEl = $("#peek");
let peekTimer = 0;
function peek(t, el) {
  if (!motion || matchMedia("(prefers-reduced-motion: reduce)").matches || (t.kind === "community" && t.image)) return;
  clearTimeout(peekTimer);
  peekTimer = setTimeout(() => {                         // a short delay: no work when the pointer just passes over
    const r = el.getBoundingClientRect();
    Object.assign(peekEl.style, { left: r.left + scrollX + "px", top: r.top + scrollY + "px", width: r.width + "px", height: r.height + "px" });
    peekEl.classList.remove("on");
    peekEl.onload = () => peekEl.classList.add("on");
    peekEl.srcdoc = sceneDoc(t, false);
  }, 120);
}
function unpeek() { clearTimeout(peekTimer); peekEl.classList.remove("on"); }

// ---------- side sheet
let lastFocus = null;
function openSheet(tpl) {
  lastFocus = document.activeElement;
  const body = $("#sheet-body");
  body.replaceChildren($(tpl).content.cloneNode(true));
  $("#sheet").hidden = $("#scrim").hidden = false;
  document.body.style.overflow = "hidden";
  $("#close").focus();
  return body;
}
function closeSheet() {
  if ($("#sheet").hidden) return;
  $("#sheet").hidden = $("#scrim").hidden = true;
  document.body.style.overflow = "";
  if (lastFocus) lastFocus.focus();
}
$("#close").onclick = closeSheet; $("#scrim").onclick = closeSheet;

// file tools load on first use only (~200 KB the gallery never needs)
let tools = null;
const script = (src) => new Promise((ok, no) => document.head.append(Object.assign(document.createElement("script"), { src, onload: ok, onerror: no })));
const loadTools = () => tools ||= ["vendor/jszip.min.js", "lib/stamper.js", "lib/stamper-pptx.js", "lib/stamper-xlsx.js", "lib/assets.js"]
  .reduce((p, f) => p.then(() => script(EXT + f)), Promise.resolve());

function save(bytes, name, type) {
  const url = URL.createObjectURL(new Blob([bytes], { type: type || "application/octet-stream" }));
  Object.assign(document.createElement("a"), { href: url, download: name }).click();
  setTimeout(() => URL.revokeObjectURL(url), 30000);
}
const toDataURL = (blob) => new Promise((ok) => { const r = new FileReader(); r.onload = () => ok(r.result); r.readAsDataURL(blob); });

// What the stamper needs: built-ins as they are, community themes as custom themes.
async function forFiles(t) {
  if (t.kind === "builtin") return { th: { ...self.WW_THEMES[t.id], id: t.id }, img: null };
  const th = WWAssets.forCustom({ id: t.id, name: t.name, mode: t.mode, accent: t.accent, accent2: t.accent2, colors: t.colors, particles: t.effect, image: t.image ? 1 : 0 });
  return { th, img: t.image ? await toDataURL(await (await fetch(t.image)).blob()) : null };
}
// Still mode: the backdrop's first frame as a PNG instead of an animated GIF.
async function stillify(a) {
  if (motion || !a.backdrop || a.backdrop.ext !== "gif") return a;
  const bmp = await createImageBitmap(new Blob([a.backdrop.bytes], { type: "image/gif" }));
  const c = Object.assign(document.createElement("canvas"), { width: bmp.width, height: bmp.height });
  c.getContext("2d").drawImage(bmp, 0, 0);
  const png = await new Promise((ok) => c.toBlob(ok, "image/png"));
  return { ...a, backdrop: { bytes: new Uint8Array(await png.arrayBuffer()), ext: "png" } };
}

function openTheme(t) {
  unpeek();
  const body = openSheet("#t-theme"), msg = $(".msg", body);
  const say = (text, cls = "") => { msg.textContent = text; msg.className = "msg " + cls; };
  $("h2", body).textContent = t.name;
  $(".meta", body).textContent = [t.blurb, t.mode === "light" ? "Light" : "Dark", motion ? "Animated" : "Still"].filter(Boolean).join(" · ");
  $(".stage", body).srcdoc = sceneDoc(t, true);

  async function theme(file, remove) {
    if (!file) return;
    if (!/\.(docx|pptx|xlsx)$/i.test(file.name)) return say("Use a .docx, .pptx or .xlsx file.", "bad");
    say(remove ? "Removing the theme…" : "Adding the theme…");
    try {
      await loadTools();
      const kind = WWAssets.kindOf(file.name), S = WWAssets.stamperFor(kind), base = file.name.replace(/\.\w+$/, "");
      const bytes = new Uint8Array(await file.arrayBuffer());
      if (remove) { save(await S.unstamp(bytes), `${base} (no theme).${kind}`); return say(`Saved “${base} (no theme).${kind}”.`, "ok"); }
      const { th, img } = await forFiles(t), info = await S.analyze(bytes);
      const assets = await stillify(await WWAssets.assetsFor(th, kind, info.sections[0], EXT, img));
      save(await S.stamp(bytes, th, assets), `${base} (${t.name}).${kind}`);
      say(`Saved “${base} (${t.name}).${kind}”. The theme is inside it now.`, "ok");
    } catch (e) { say("That file couldn't be themed: " + e.message, "bad"); }
  }
  const drop = $(".drop", body), input = $("input", drop);
  input.onchange = () => theme(input.files[0]);
  drop.ondragover = (e) => { e.preventDefault(); drop.classList.add("over"); };
  drop.ondragleave = () => drop.classList.remove("over");
  drop.ondrop = (e) => { e.preventDefault(); drop.classList.remove("over"); theme(e.dataTransfer.files[0]); };

  const acts = {
    async starter() {
      await loadTools(); const { th, img } = await forFiles(t);
      const a = await stillify(await WWAssets.assetsFor(th, "docx", WWStamper.NEW_DOC_SECTION, EXT, img));
      save(await WWStamper.newDocument(th, a), `${t.name}.docx`); say(`Saved “${t.name}.docx”.`, "ok");
    },
    async teams() {
      await loadTools(); const { th, img } = await forFiles(t);
      save(await WWAssets.teams(th, EXT, img), `${t.name} (Teams).jpg`, "image/jpeg"); say("Saved. In a Teams meeting: Video effects › Add new.", "ok");
    },
    async ext() {
      if (t.kind === "builtin") return say("It's built into the extension: click the Wild Themes icon in Word and pick it.", "ok");
      const { th, img } = await (loadTools().then(() => forFiles(t)));
      save(JSON.stringify({ wildtheme: 1, theme: { name: th.name, mode: th.mode, accent: th.accent, accent2: th.accent2, colors: th.colors, particles: th.particles, image: th.image, fx: "dot" }, image: img }),
        `${t.name}.wildtheme`, "application/json");
      say("Saved. In the extension: Custom themes › Import a theme file.", "ok");
    },
    remove() { const i = Object.assign(document.createElement("input"), { type: "file", accept: ".docx,.pptx,.xlsx" }); i.onchange = () => theme(i.files[0], true); i.click(); },
  };
  $$(".acts button", body).forEach((b) => b.onclick = () => acts[b.dataset.a]().catch((e) => say(e.message, "bad")));
}

// ---------- install
$("#open-install").onclick = () => {
  const body = openSheet("#t-install");
  $$(".copy", body).forEach((b) => b.onclick = async () => {
    try { await navigator.clipboard.writeText(b.dataset.copy); b.textContent = "Copied"; } catch (_) { b.textContent = "Select it"; }
  });
};

// ---------- submit: build the picture here (size-checked), then hand over to a prefilled GitHub form
$("#open-submit").onclick = () => {
  const body = openSheet("#t-submit"), msg = $("#s-msg", body), go = $("#s-go", body), prev = $("#s-prev", body);
  let kind = "picture", pic = null, mode = "dark", accent = "#8a5cff";
  const say = (text, cls = "") => { msg.textContent = text; msg.className = "msg " + cls; };
  const ready = () => { go.disabled = !$("#s-name", body).value.trim() || (kind === "picture" && !pic); };
  const colours = () => [$("#s-c1", body).value, $("#s-c2", body).value, $("#s-c3", body).value];
  const showColours = () => {
    const [a, b, c] = colours();
    prev.style.background = `radial-gradient(120% 90% at 0 0,${b},transparent 60%),radial-gradient(120% 90% at 100% 100%,${c},transparent 60%),${a}`;
  };
  $$(".seg button", body).forEach((b) => b.onclick = () => {
    kind = b.dataset.k; $$(".seg button", body).forEach((x) => x.setAttribute("aria-checked", x === b));
    $$("[data-for]", body).forEach((d) => d.hidden = d.dataset.for !== kind);
    if (kind === "colours") showColours(); else prev.style.background = pic ? `center/cover url(${pic.url})` : "";
    ready();
  });
  $("#s-name", body).oninput = ready;
  for (const id of ["#s-c1", "#s-c2", "#s-c3"]) $(id, body).oninput = showColours;
  $("#s-pic", body).onchange = async (e) => {
    const f = e.target.files[0]; if (!f) return;
    if (f.size > 25e6) return say("That picture is over 25 MB. Pick a smaller one.", "bad");
    try {
      const bmp = await createImageBitmap(f);
      if (bmp.width < 1280 || bmp.height < 720) return say(`It's ${bmp.width}×${bmp.height}. Pictures need to be at least 1280×720.`, "bad");
      // cover-crop to 16:9 at 1920×1080, then lower the quality until it's under 2 MB
      const c = Object.assign(document.createElement("canvas"), { width: 1920, height: 1080 }), g = c.getContext("2d");
      const s = Math.max(1920 / bmp.width, 1080 / bmp.height);
      g.drawImage(bmp, (1920 - bmp.width * s) / 2, (1080 - bmp.height * s) / 2, bmp.width * s, bmp.height * s);
      let q = .88, blob;
      do { blob = await new Promise((ok) => c.toBlob(ok, "image/jpeg", q)); q -= .08; } while (blob.size > 2e6 && q > .3);
      if (blob.size > 2e6) return say("Couldn't get that picture under 2 MB. Try a simpler one.", "bad");
      const px = g.getImageData(0, 0, 1920, 1080).data; let r = 0, gg = 0, bb = 0, n = 0;
      for (let i = 0; i < px.length; i += 4 * 211) { r += px[i]; gg += px[i + 1]; bb += px[i + 2]; n++; }
      mode = (r + gg + bb) / n / 3 > 140 ? "light" : "dark";
      accent = "#" + [r, gg, bb].map((v) => Math.min(255, v / n * 1.35 + 25) | 0).map((v) => v.toString(16).padStart(2, "0")).join("");
      pic = { blob, url: URL.createObjectURL(blob) };
      prev.style.background = `center/cover url(${pic.url})`;
      say(`Ready: 1920×1080, ${(blob.size / 1e6).toFixed(1)} MB.`, "ok");
    } catch (_) { say("That picture couldn't be opened. Try a PNG or JPG.", "bad"); }
    ready();
  };
  go.onclick = () => {
    const name = $("#s-name", body).value.trim().slice(0, 30), slug = name.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "") || "theme";
    const p = new URLSearchParams({ template: "theme.yml", title: "Theme: " + name, name, mode: mode === "light" ? "Light" : "Dark" });
    if (kind === "picture") { save(pic.blob, slug + ".jpg", "image/jpeg"); p.set("kind", "Picture (still)"); p.set("accent", accent); }
    else {
      const c = colours(), fx = $("#s-fx", body);
      p.set("kind", "Colours (animated)"); p.set("colours", c.join(" ")); p.set("accent", c[1]);
      p.set("mode", parseInt(c[0].slice(1, 3), 16) + parseInt(c[0].slice(3, 5), 16) + parseInt(c[0].slice(5, 7), 16) > 450 ? "Light" : "Dark");
      p.set("effect", fx.options[fx.selectedIndex].text.replace("No effect", "None"));
    }
    open(`${REPO}/issues/new?${p}`, "_blank", "noopener");
    say(kind === "picture" ? "The picture is in your downloads. Drag it into the Picture box on GitHub, then press Create." : "Press Create on GitHub to send it.", "ok");
  };
  showColours(); prev.style.background = "";
};

render();
