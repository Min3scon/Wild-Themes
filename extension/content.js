// Wild Themes for Word Online — runs only inside the Word editor.
// No animation loop and nothing per frame: the scene is CSS gradients painted once,
// and the browser slides/fades whole layers. Only the active theme's CSS is loaded.
// JavaScript runs on events only: load, resize, scroll end, a key press, a setting change.
(() => {
  const D = document, H_ = D.documentElement, L = location, P = L.pathname, HN = L.hostname;
  // Which app? Word and PowerPoint have a page/slide on a canvas; Excel and Teams have panes.
  const app = /\/we\//i.test(P) || HN == "word.cloud.microsoft" ? "doc"
    : /\/p\//i.test(P) || HN == "powerpoint.cloud.microsoft" ? "doc"
    : /\/x\//i.test(P) || HN == "excel.cloud.microsoft" || /^teams\./.test(HN) ? "pane" : "";
  if (window.__ww || !app || innerWidth < 300) return;
  window.__ww = 1;
  const T = self.WW_THEMES, DEF = { on: true, theme: "synthwave", motion: "auto", sparks: true, docs: {} };
  const ATTR = (n, a, v) => v == null ? n.removeAttribute(a) : n.setAttribute(a, v);
  let S = DEF, custom = {}, id, t, bg, fx, file, pages = [], tm, last = 0, ptr, k = 0;

  // Which file is this? A short hash of the WOPISrc file address (the address itself is never stored).
  let key = (new URLSearchParams(L.search.toLowerCase()).get("wopisrc") || "").split("?")[0], h = 2166136261;
  for (let i = 0; i < key.length; i++) h = Math.imul(h ^ key.charCodeAt(i), 16777619);
  key = key && (h >>> 0).toString(36);

  // A theme saved inside the file shows up as picture alt text in Word's page.
  const readFile = () => {
    const n = D.querySelector('[alt^="Wild Theme:"],[aria-label^="Wild Theme:"]');
    if (!n) return null;
    const f = {};
    (n.alt || n.getAttribute("aria-label")).slice(11).split(";").forEach((p, i) => {
      const [a, b] = p.split("=").map((x) => x.trim());
      i ? (f[a] = b) : (f.id = a);
    });
    return f;
  };

  // this document's choice > the theme saved in the file > your default
  const pick = () => {
    const w = (key && S.docs[key]) || (file && file.id) || S.theme;
    return T[w] ? [w, T[w]] : custom[w] ? [w, custom[w]]
      : file && w == file.id ? [w, { custom: 1, mode: file.mode, accent: file.accent, accent2: file.accent2, fx: "dot",
        colors: [file.mode == "light" ? "#f4f1ff" : "#0d0b1a", file.accent, file.accent2] }]
      : T[S.theme] ? [S.theme, T[S.theme]] : [];
  };
  const still = () => S.motion == "off" || (S.motion == "auto" && (matchMedia("(prefers-reduced-motion:reduce)").matches ||
    navigator.hardwareConcurrency <= 2 || navigator.deviceMemory <= 2));

  // Don't show anything until we know which theme applies: in Word/PowerPoint a theme saved in
  // the file only shows up once the page is drawn, and showing the default first made it flash.
  let ready = app == "pane", seen = 0;
  function probe(last) {
    if (ready) return;
    file = readFile();
    const editable = !!D.querySelector('[contenteditable="true"]');
    if (file || (editable && seen) || last) { ready = true; apply(); }
    seen = editable;
  }

  function apply() {
    const [nid, nt] = S.on && (ready || (key && S.docs[key])) ? pick() : [];
    if (nid != id) {
      clear();
      if (!nid) return;
      id = nid; t = nt;
      bg = D.createElement("div");
      bg.id = "ww-bg";
      bg.innerHTML = "<i></i><i></i><i></i><i></i>";
      if (t.custom) {
        const c = t.colors || [];
        bg.dataset.scene = "custom";
        bg.style.cssText = `--ww-c1:${c[0]};--ww-c2:${c[1]};--ww-c3:${c[2]};--ww-dim:${t.dim || 0};--ww-dimc:${t.mode == "light" ? "#fff" : "#000"}`;
        if (t.particles && t.particles != "none") bg.dataset.part = t.particles;
        if (t.image) picture();
      } else {
        bg.dataset.scene = id;
        chrome.runtime.sendMessage({ css: id });                 // load just this scene's CSS
      }
      H_.append(bg);
      H_.style.cssText += `;--ww-a:${t.accent};--ww-edge:${t.accent}88;--ww-glow:${t.accent}55;--ww-base:${t.custom ? t.colors && t.colors[0] : t.mode == "light" ? "#f3eefa" : "#08060f"}`;
      H_.dataset.ww = id;
    }
    if (!bg) return;                                                // nothing shown yet (still waiting)
    bg.classList.toggle("still", still());
    scan();
  }

  function clear() {
    if (bg) bg.remove(), (bg = 0);
    if (fx) fx.remove(), (fx = 0);
    id = t = null;
    delete H_.dataset.ww;
    for (const v of ["--ww-a", "--ww-edge", "--ww-glow", "--ww-base"]) H_.style.removeProperty(v);
    D.querySelectorAll("[data-ww-clear],[data-ww-page]").forEach((n) => (ATTR(n, "data-ww-clear"), ATTR(n, "data-ww-page")));
    pages = [];
  }

  // Custom picture themes: decoded once, drawn once onto a canvas. Nothing after that.
  async function picture() {
    const own = bg, b64 = (await chrome.storage.local.get("img:" + t.image))["img:" + t.image];
    if (!b64 || own != bg) return;
    const bin = atob(b64.slice(b64.indexOf(",") + 1)), u8 = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);              // no fetch: works under any page policy
    const bmp = await createImageBitmap(new Blob([u8])), cv = D.createElement("canvas"), r = Math.min(1, 1600 / innerWidth);
    cv.width = innerWidth * r; cv.height = innerHeight * r;
    const s = Math.max(cv.width / bmp.width, cv.height / bmp.height);
    cv.getContext("2d").drawImage(bmp, (cv.width - bmp.width * s) / 2, (cv.height - bmp.height * s) / 2, bmp.width * s, bmp.height * s);
    bmp.close();
    own.prepend(cv);
  }

  // Find Word's page and the grey canvas around it. Strict rules; if unsure, change nothing.
  const solid = (n) => {
    if (n.hasAttribute("data-ww-clear")) return 1;
    const s = getComputedStyle(n);
    return s.backgroundImage != "none" || !/, 0\)|^transparent/.test(s.backgroundColor);
  };
  const ok = (n) => n != bg && n != H_ && n != D.body;
  // Excel/Teams: the scene shows only inside the main pane. Everything outside keeps its
  // look because the page behind is painted in the app's own background colour.
  function scanPane() {
    const root = D.querySelector("[role=main]"), rr = root && root.getBoundingClientRect();
    if (!rr || rr.width * rr.height < 1e5) return;
    let shell = "";
    for (let a = root.parentElement; a && a != D.body && a != H_; a = a.parentElement) {
      if (a.hasAttribute("data-ww-clear") || solid(a)) { shell = shell || getComputedStyle(a).backgroundColor; ATTR(a, "data-ww-clear", ""); }
    }
    if (shell && !/, 0\)/.test(shell)) H_.style.setProperty("--ww-base", shell, "important");
    // a see-through veil in the pane's own colour keeps text readable that has no card behind it
    const own = getComputedStyle(root).backgroundColor, m = own.match(/\d+(\.\d+)?/g);
    if (m && !root.hasAttribute("data-ww-clear") && solid(root)) bg.style.setProperty("--ww-veil", `rgba(${m[0]},${m[1]},${m[2]},${HN.includes("teams") ? .45 : .8})`);
    const big = rr.width * rr.height * .45;
    for (let i = 0; i < 16; i++) {
      for (const n of D.elementsFromPoint(rr.left + rr.width * ((i % 4) + .5) / 4, rr.top + rr.height * ((i >> 2) + .5) / 4)) {
        if (!ok(n) || !(n == root || root.contains(n))) continue;
        const r = n.getBoundingClientRect();
        if (r.width * r.height >= big && (n.hasAttribute("data-ww-clear") || solid(n))) ATTR(n, "data-ww-clear", "");
      }
    }
    bg.style.setProperty("inset", "auto", "important");               // first, or it would reset top/left
    for (const [k, v] of [["top", rr.top], ["left", rr.left], ["width", rr.width], ["height", rr.height]]) bg.style.setProperty(k, v + "px", "important");
    bg.classList.add("veil");
  }

  function scan() {
    if (!bg) return;
    if (!bg.isConnected) H_.append(bg);
    if (app == "pane") return scanPane();
    const W = innerWidth, H = innerHeight, bars = [...D.querySelectorAll("[role=tablist],[role=toolbar],[role=menubar]")];
    let rb = 0, plate = "";
    for (const n of bars) { const r = n.getBoundingClientRect(); if (r.width > 40 && r.top < H * .4 && r.bottom > rb) rb = r.bottom; }
    const found = [];
    for (let y = rb + 30; y < H - 10; y += 110) {
      const st = D.elementsFromPoint(W / 2, y).filter(ok);
      if (st[0] && st[0].closest("[role=dialog],[role=menu],[role=listbox]")) continue;
      for (let i = 0; i < st.length; i++) {
        const n = st[i], r = n.getBoundingClientRect();
        if (r.width > W * .92) break;
        if (r.width < 260 || r.height < 110 || r.top < rb - 4 || n.hasAttribute("data-ww-clear") || !solid(n)) continue;
        const u = st.slice(i + 1).find(solid), ur = u ? u.getBoundingClientRect() : { left: 0, right: W };
        if (r.left - ur.left > 7 && ur.right - r.right > 7 && !found.includes(n)) found.push(n);
        break;
      }
    }
    const pg = found.filter((p) => !found.some((q) => q != p && p.contains(q)));
    if (pg.length) {
      pages.forEach((p) => pg.includes(p) || p.getBoundingClientRect().bottom < 0 || ATTR(p, "data-ww-page"));
      pg.forEach((p) => ATTR(p, "data-ww-page", ""));
      pages = pg;
      const pr = pg.map((p) => p.getBoundingClientRect());
      // canvas = solid boxes beside the page that contain it; if one also holds the ribbon,
      // a plate in its colour goes behind the ribbon so the ribbon stays solid
      for (const x of [W * .03, W * .97]) for (let y = rb + 20; y < H - 10; y += (H - rb) / 4) {
        if (pr.some((r) => x >= r.left && x <= r.right && y >= r.top && y <= r.bottom)) continue;
        for (const n of D.elementsFromPoint(x, y)) {
          if (!ok(n) || pg.includes(n) || !solid(n) || !pg.some((p) => n.contains(p))) continue;
          if (!n.hasAttribute("data-ww-clear") && bars.some((c) => n.contains(c))) plate = getComputedStyle(n).backgroundColor;
          ATTR(n, "data-ww-clear", "");
        }
      }
      if (plate) bg.lastChild.style.cssText = `inset:0 0 auto;height:${rb}px;background:${plate}`;
    }
    const f = readFile();
    if (JSON.stringify(f) != JSON.stringify(file)) file = f, apply();
  }
  const later = (ms) => { clearTimeout(tm); tm = setTimeout(scan, ms); };

  // Typing sparks: 12 reusable dots, coloured once per theme, moved by the compositor.
  const G = "01<>{}#*+✦✧★";
  addEventListener("keydown", (e) => {
    const now = e.timeStamp;
    if (!t || !S.sparks || now - last < 90 || e.ctrlKey || e.metaKey || e.altKey || (e.key.length > 1 && e.key != "Enter")) return;
    last = now;
    requestAnimationFrame(() => {
      const sel = getSelection(), r = sel.rangeCount && sel.getRangeAt(0).getClientRects()[0];
      const p = r && r.top > 0 ? [r.right, r.top + r.height / 2] : ptr;
      if (!p || !t) return;
      if (!fx) {
        fx = D.createElement("div");
        fx.id = "ww-fx";
        fx.dataset.fx = t.fx || "dot";
        fx.innerHTML = "<b></b>".repeat(12);
        [...fx.children].forEach((b, i) => {
          b.style.setProperty("--c", i % 2 ? t.accent2 : t.accent);
          if (/glyph|star/.test(t.fx)) b.textContent = t.fx == "star" ? G[9 + i % 3] : G[i % 9];
        });
        H_.append(fx);
      }
      fx.style.transform = `translate(${p[0]}px,${p[1]}px)`;
      const down = /snow|petal|confetti|drop/.test(t.fx) ? 26 : -26;
      for (let i = e.key == "Enter" ? 8 : 4; i--;) {
        const a = Math.random() * 6.28, d = 18 + Math.random() * 30;
        fx.children[k++ % 12].animate([{ opacity: 1, transform: "none" },
          { opacity: 0, transform: `translate(${Math.cos(a) * d}px,${Math.sin(a) * d * .6 + down}px) rotate(${a * 57 | 0}deg)` }],
        { duration: 500 + d * 10, easing: "ease-out" });
      }
    });
  }, true);
  addEventListener("pointerdown", (e) => { ptr = [e.clientX, e.clientY]; }, { capture: true, passive: true });
  addEventListener("resize", () => later(300), { passive: true });
  addEventListener("scroll", () => later(400), { capture: true, passive: true });
  chrome.runtime.onMessage.addListener((m, _, reply) => { m == "ww-status" && reply({ docKey: key, active: id, file }); });

  const load = () => chrome.storage.sync.get(DEF, (v) => chrome.storage.local.get("customs", (c) => {
    S = { ...DEF, ...v };
    const before = id && !T[id] && JSON.stringify(custom[id]);
    custom = {};
    (c.customs || []).forEach((x) => (custom[x.id] = { ...x, custom: 1 }));
    if (before && before != JSON.stringify(custom[id])) clear();   // rebuild only if that custom theme changed
    file = readFile() || file;
    apply();
  }));
  chrome.storage.onChanged.addListener(load);
  load();
  [250, 600, 1200, 2000, 3500].forEach((ms) => setTimeout(probe, ms));
  setTimeout(() => probe(true), 5000);                              // no file theme by now: use the default
  [800, 2500, 6000].forEach((ms) => setTimeout(scan, ms));         // Word builds its page in stages
})();
