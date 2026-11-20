// Wild Themes for Word Online — runs only inside the Word editor.
// No animation loop and nothing per frame: the scene is CSS gradients painted once,
// and the browser slides/fades whole layers. Only the active theme's CSS is loaded.
// JavaScript runs on events only: load, resize, scroll end, a key press, a setting change.
(() => {
  const D = document, H_ = D.documentElement, L = location, P = L.pathname, HN = L.hostname;
  // Which app? Word and PowerPoint have a page/slide on a canvas; Excel and Teams have panes.
  // Google Docs and Slides: a page/slide on a grey surround, like Word. Sheets: a grid pane, like Excel.
  const GH = HN == "docs.google.com", GK = GH ? (/^\/(document|spreadsheets|presentation)\//.exec(P) || [])[1] : "";
  const app = GK ? (GK == "spreadsheets" ? "pane" : "doc")
    : /\/we\//i.test(P) || HN == "word.cloud.microsoft" ? "doc"
    : /\/p\//i.test(P) || HN == "powerpoint.cloud.microsoft" ? "doc"
    : /\/x\//i.test(P) || HN == "excel.cloud.microsoft" || /^teams\./.test(HN) || HN == "assignments.edu.cloud.microsoft" ? "pane" : "";   // Assignments: its own site, in a frame inside Teams
  if (window.__ww || !app || innerWidth < 300) return;
  window.__ww = 1;
  const T = self.WW_THEMES, RW = self.WWRewards, DEF = { on: true, theme: "synthwave", motion: "auto", sparks: true, docs: {}, docsDark: "auto", docsClean: true, found: {} };
  const ATTR = (n, a, v) => v == null ? n.removeAttribute(a) : n.setAttribute(a, v);
  let S = DEF, custom = {}, R = {}, id, t, bg, fx, file, pages = [], tm, last = 0, ptr, k = 0;

  // Which file is this? A short hash of the WOPISrc file address, or of Google's document id
  // (the address / id itself is never stored).
  let key = GH ? (/\/d\/([\w-]+)/.exec(P) || [])[1] || "" : (new URLSearchParams(L.search.toLowerCase()).get("wopisrc") || "").split("?")[0], h = 2166136261;
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

  // this document's choice > the theme saved in the file > your default. A secret theme counts only once it's been found.
  const has = (w) => !!T[w] && !(T[w].secret && !(S.found || {})[w]);
  const pick = () => {
    const w = (key && S.docs[key]) || (file && file.id) || S.theme;
    return has(w) ? [w, T[w]] : R[w] ? [w, R[w]] : custom[w] ? [w, custom[w]]
      : file && w == file.id ? [w, { custom: 1, mode: file.mode, accent: file.accent, accent2: file.accent2, fx: "dot",
        colors: [file.mode == "light" ? "#f4f1ff" : "#0d0b1a", file.accent, file.accent2] }]
      : has(S.theme) ? [S.theme, T[S.theme]] : T[S.theme] ? [DEF.theme, T[DEF.theme]] : [];   // a secret not found yet: the usual default
  };
  // Smoothness guard (Auto only): the scene's real frame rate is measured; if this device can't keep up, the big
  // or 3D layers stop first (level 1: "lite"), then everything (level 2). The verdict is remembered per device.
  let level = 0, bad = 0, gt = 0;
  const lite = () => S.motion == "auto" && level >= 1;
  const still = () => S.motion == "off" || (S.motion == "auto" && (level >= 2 || matchMedia("(prefers-reduced-motion:reduce)").matches ||
    navigator.hardwareConcurrency <= 2 || navigator.deviceMemory <= 2));

  const sample = (ms) => new Promise((ok) => {                       // frame gaps over ms; null if the tab was hidden meanwhile
    const t0 = [];
    const tick = (now) => {
      if (D.hidden) return ok(null);
      t0.push(now);
      if (now - t0[0] >= ms) return ok(t0.slice(1).map((x, i) => x - t0[i]));
      requestAnimationFrame(tick);
    };
    requestAnimationFrame(tick);
  });
  // Layers that are huge, or that rotate, scale or use perspective, cost the most to draw: those are the ones level 1 holds still.
  const markHeavy = () => bg && [...bg.children].forEach((i) => {
    let heavy = i.offsetWidth * i.offsetHeight > 2.2 * innerWidth * innerHeight;
    try { heavy = heavy || i.getAnimations().some((a) => /rotate|perspective|scale/.test(JSON.stringify(a.effect.getKeyframes().map((k) => k.transform)))); } catch (_) { /* no animation info */ }
    i.toggleAttribute("data-heavy", heavy);
  });
  const guard = (ms) => {
    clearTimeout(gt);
    gt = setTimeout(async () => {
      if (!bg || S.motion != "auto" || still()) return;
      if (D.hidden) return void D.addEventListener("visibilitychange", () => guard(3000), { once: true });
      const gaps = await sample(2000);
      if (!gaps) return guard(3000);
      if (!WWPerf.judge(gaps, 2000)) { bad = 0; return; }
      if (++bad < 2) return guard(5000);                                 // a page that's still loading drops frames: look once more
      bad = 0; level++;
      if (level == 1) markHeavy();
      chrome.storage.local.set({ perf: { level, at: Date.now() } });
      apply();
      if (level < 2) guard(4000);
    }, ms);
  };

  // The colour that comes out as `hex` after the dark page's invert(1) hue-rotate(180deg) (that filter is its own inverse, up to
  // the 0-1 clamp). Used for the glow when it sits on the page box that is itself inverted, so the glow keeps the theme's colour.
  const inv = (hex) => {
    const n = parseInt(String(hex).replace("#", "").slice(0, 6), 16) || 0, d = [n >> 16 & 255, n >> 8 & 255, n & 255].map((v) => v / 255), M = [[-.574, 1.43, .144], [.426, .43, .144], [.426, 1.43, -.856]];
    return M.map((r) => (Math.max(0, Math.min(1, 1 - (r[0] * d[0] + r[1] * d[1] + r[2] * d[2]))) * 255 + .5 | 0).toString(16).padStart(2, "0")).join("");
  };

  // Don't show anything until we know which theme applies: in Word/PowerPoint a theme saved in
  // the file only shows up once the page is drawn, and showing the default first made it flash.
  let ready = app == "pane" || GH, seen = 0;   // Google files carry no theme we can read: nothing to wait for
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
      flush();                                                      // words so far belong to the theme they were typed in
      clear();
      if (!nid) return;
      id = nid; t = nt;
      bg = D.createElement("div");
      bg.id = "ww-bg";
      for (let i = 0; i < 4; i++) bg.append(D.createElement("i"));
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
      const vs = D.getElementById("ww-vars") || D.head.appendChild(Object.assign(D.createElement("style"), { id: "ww-vars" }));
      vs.textContent = `html[data-ww]{--ww-a:${t.accent};--ww-edge:${t.accent}88;--ww-glow:${t.accent}55;--ww-edge-i:#${inv(t.accent)}88;--ww-glow-i:#${inv(t.accent)}55}`;
      guard(3000);
    }
    if (!bg) return;                                                // nothing shown yet (still waiting)
    if (lite()) markHeavy();
    bg.classList.toggle("lite", lite());
    bg.classList.toggle("still", still());
    scan();
  }

  function clear() {
    if (bg) bg.remove(), (bg = 0);
    if (fx) fx.remove(), (fx = 0);
    id = t = null;
    delete H_.dataset.ww;
    H_.removeAttribute("data-ww-dark"); H_.removeAttribute("data-ww-clean");
    const vs = D.getElementById("ww-vars"); if (vs) vs.remove();
    for (const v of ["--ww-a", "--ww-edge", "--ww-glow", "--ww-base"]) H_.style.removeProperty(v);
    D.querySelectorAll("[data-ww-clear],[data-ww-page]").forEach((n) => (ATTR(n, "data-ww-clear"), ATTR(n, "data-ww-page")));
    pages = [];
  }

  // Custom picture themes: decoded once, drawn once onto a canvas. Nothing after that.
  async function picture() {
    const own = bg, b64 = (await chrome.storage.local.get("img:" + t.image))["img:" + t.image];
    if (!b64 && t.community) {                                    // community picture not fetched yet: fetch once, then redraw
      chrome.runtime.sendMessage({ picture: t.image }, (r) => { if (r && r.ok && own == bg) { clear(); apply(); } });
      return;
    }
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
    // Teams v2 has no role=main; its chat/channel area is app-layout-area--main
    const root = D.querySelector('[data-tid="app-layout-area--main"]') || D.querySelector("[role=main]") || (GH && D.querySelector(".grid-container, #docs-editor, #docs-editor-container")) || (HN.startsWith("assignments.") ? D.body : null), rr = root && root.getBoundingClientRect();
    if (!rr || rr.width * rr.height < 1e5) return;
    let shell = "";
    for (let a = root.parentElement; a && a != D.body && a != H_; a = a.parentElement) {
      if (a.hasAttribute("data-ww-clear") || solid(a)) { shell = shell || getComputedStyle(a).backgroundColor; ATTR(a, "data-ww-clear", ""); }
    }
    if (!shell || /, 0\)/.test(shell)) {   // Teams v2: its colour is on body/html, which our own rules clear; read it with them off
      const was = H_.getAttribute("data-ww"); H_.removeAttribute("data-ww");
      shell = [D.body, H_].map((n) => getComputedStyle(n).backgroundColor).find((c) => !/, 0\)|^transparent/.test(c)) || "";
      H_.setAttribute("data-ww", was);
    }
    if (shell) H_.style.setProperty("--ww-base", shell, "important");
    // a see-through veil in the pane's own colour keeps text readable that has no card behind it
    const own = getComputedStyle(root).backgroundColor, m = own.match(/\d+(\.\d+)?/g);
    if (m && !root.hasAttribute("data-ww-clear") && solid(root)) bg.style.setProperty("--ww-veil", `rgba(${m[0]},${m[1]},${m[2]},${HN.includes("teams") || HN.startsWith("assignments.") ? .45 : .8})`);
    const big = rr.width * rr.height * .45;
    for (let i = 0; i < 16; i++) {
      for (const n of D.elementsFromPoint(rr.left + rr.width * ((i % 4) + .5) / 4, rr.top + rr.height * ((i >> 2) + .5) / 4)) {
        if (!ok(n) || !(n == root || root.contains(n))) continue;
        const r = n.getBoundingClientRect();
        // never a message, post or list item (a long post is big too): only the area's own layers
        if (n.closest("li,[role=listitem],[role=article],[data-tid=channel-pane-message],[data-tid=chat-pane-message],[data-tid=chat-pane-item]")) continue;
        if (r.width * r.height >= big && (n.hasAttribute("data-ww-clear") || solid(n))) ATTR(n, "data-ww-clear", "");
      }
    }
    bg.style.setProperty("inset", "auto", "important");               // first, or it would reset top/left
    for (const [k, v] of [["top", rr.top], ["left", rr.left], ["width", rr.width], ["height", rr.height]]) bg.style.setProperty(k, v + "px", "important");
    bg.classList.add("veil");
  }

  // The page's own background colour (on body or html), read with our rules switched off.
  let shell0;
  const shellColour = () => {
    if (shell0) return shell0;
    const was = H_.getAttribute("data-ww"); H_.removeAttribute("data-ww");
    shell0 = [D.body, H_].map((n) => getComputedStyle(n).backgroundColor).find((c) => !/, 0\)|^transparent/.test(c)) || "";
    if (was != null) H_.setAttribute("data-ww", was);
    return shell0;
  };

  // Google Docs (print layout). Its pages are white tiles inside .kix-appview-editor (the scrolling area), which sits
  // between the app's own bars and side panels. The scene covers that area only, the page keeps its own white, and
  // everything outside the area looks as before because the page's own colour is painted behind it by html.
  function scanDocs() {
    const ed = D.querySelector(".kix-appview-editor"), er = ed && ed.getBoundingClientRect();
    const pg = [...D.querySelectorAll(".kix-page-paginated")];
    if (!er || er.width * er.height < 1e5 || !pg.length) return false;     // pageless, or Google changed the layout: the general scan tries
    const base = shellColour();
    for (let a = ed; a && a != D.body && a != H_; a = a.parentElement) if (a.hasAttribute("data-ww-clear") || solid(a)) ATTR(a, "data-ww-clear", "");
    if (base) H_.style.setProperty("--ww-base", base, "important");
    // Clean view (see docsFlags): the scene takes the place of the rulers and the scrollbar
    const clean = H_.hasAttribute("data-ww-clean");
    let top = er.top, left = er.left;
    if (clean) {
      const hr = D.querySelector("#kix-horizontal-ruler-container, .docs-horizontal-ruler"), vr = D.querySelector("#kix-vertical-ruler, .docs-vertical-ruler");
      const hb = hr && hr.getBoundingClientRect(), vb = vr && vr.getBoundingClientRect();
      if (hb && hb.height && hb.bottom <= er.top + 40) top = Math.min(top, hb.top);        // the ruler above the page
      if (vb && vb.width && vb.right <= er.left + 40) left = Math.min(left, vb.left);      // the ruler beside it
    }
    bg.style.setProperty("inset", "auto", "important");
    for (const [k, v] of [["top", top], ["left", left], ["width", er.right - left], ["height", er.bottom - top]]) bg.style.setProperty(k, v + "px", "important");
    bg.classList.remove("veil");
    // the glow goes on the outermost box that is exactly the page's size: an inner wrapper may be clipped by an ancestor
    const glow = pg.map((p) => {
      const pr = p.getBoundingClientRect(); let g = p;
      for (let a = p.parentElement; a && a != ed; a = a.parentElement) {
        const r = a.getBoundingClientRect();
        if (Math.abs(r.width - pr.width) > 3 || Math.abs(r.left - pr.left) > 3) break;
        g = a;
      }
      return g;
    });
    pages.forEach((p) => glow.includes(p) || ATTR(p, "data-ww-page"));
    glow.forEach((p) => ATTR(p, "data-ww-page", ""));
    pages = glow;
    return true;
  }

  // Google Docs: dark page (Auto follows the theme) and clean view (no rulers, no scrollbar) are switched by attributes on
  // html, whether or not the scan below finds Docs' page boxes. The CSS in content.css does the work.
  function docsFlags() {
    // Auto: a dark theme gives a dark page, a light theme leaves the page as it is. A page that is already dark (the document's own
    // page colour) is left alone in Auto, so it never turns light; On inverts whatever the page is.
    const tile = D.querySelector(".kix-canvas-tile-content") || D.querySelector(".kix-page-paginated"), pc = tile ? getComputedStyle(tile).backgroundColor : "", m = pc.match(/[\d.]+/g);
    const alreadyDark = !!m && !/, 0\)|^transparent/.test(pc) && (.2126 * m[0] + .7152 * m[1] + .0722 * m[2]) / 255 < .35;
    const dark = S.docsDark == "on" || (S.docsDark != "off" && !!t && t.mode == "dark" && !alreadyDark), clean = S.docsClean !== false;
    if (H_.hasAttribute("data-ww-dark") != dark) ATTR(H_, "data-ww-dark", dark ? "" : null);
    if (H_.hasAttribute("data-ww-clean") != clean) { ATTR(H_, "data-ww-clean", clean ? "" : null); setTimeout(() => dispatchEvent(new Event("resize")), 50); }   // Docs centres the page again without the scrollbar
  }

  function scan() {
    hookTyping();
    if (!bg) return;
    if (!bg.isConnected) H_.append(bg);
    if (GK == "document") docsFlags();
    if (app == "pane") return scanPane();
    if (GK == "document" && scanDocs()) return;
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
        if (r.left - ur.left > 7 && ur.right - r.right > 7) { if (!found.includes(n)) found.push(n); break; }
        // no wider box behind it: this is a layer inside the page (Google Docs draws its white in tiles), try the box around it
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
      // beside the page as well as at the window's edges: side panels can sit between the edge and the page's backdrop
      for (const x of [W * .03, W * .97, ...pr.flatMap((r) => [r.left - 12, r.right + 12])].filter((x) => x > 1 && x < W - 1)) for (let y = rb + 20; y < H - 10; y += (H - rb) / 4) {
        if (pr.some((r) => x >= r.left && x <= r.right && y >= r.top && y <= r.bottom)) continue;
        for (const n of D.elementsFromPoint(x, y)) {
          if (!ok(n) || pg.includes(n) || !solid(n) || !pg.some((p) => n.contains(p))) continue;
          if (!n.hasAttribute("data-ww-clear") && bars.some((c) => n.contains(c))) plate = getComputedStyle(n).backgroundColor;
          ATTR(n, "data-ww-clear", "");
        }
      }
      // Google's top bars may have no colour of their own and sit on the page's background, which we clear:
      // when nothing solid is behind them, keep the page's original colour there
      if (!plate && GH && rb && !D.elementsFromPoint(W / 2, Math.max(2, rb / 2)).filter(ok).some((n) => n.hasAttribute("data-ww-clear") || solid(n))) plate = shellColour();
      if (plate) bg.lastChild.style.cssText = `inset:0 0 auto;height:${rb}px;background:${plate}`;
    }
    const f = readFile();
    if (JSON.stringify(f) != JSON.stringify(file)) file = f, apply();
  }
  const later = (ms) => { clearTimeout(tm); tm = setTimeout(scan, ms); };

  // Words typed with a theme showing, per theme. Counted from real key presses only (letters or digits, then a space,
  // Enter or punctuation), so pasted or dropped text never counts; paste itself still works as usual. Sent every few
  // seconds; the background adds them up and rolls the dice for the secret theme, one roll per word.
  let inWord = false, typed = 0, typedFor = null, ft = 0;
  const countKey = (e) => {
    const k = RW.keyKind(e);
    if (k == "char") inWord = !!id;
    else if (k == "end" && inWord) {
      inWord = false;
      if (typedFor && typedFor != id) flush();
      typed++; typedFor = id;
      if (!ft) ft = setTimeout(flush, 4000);
    }
  };
  function flush() {
    clearTimeout(ft); ft = 0;
    if (!typed || !typedFor) return;
    const n = typed, th = typedFor; typed = 0;
    try { chrome.runtime.sendMessage({ words: n, theme: th }, (r) => !chrome.runtime.lastError && r && reward(r)); } catch (_) { /* the extension was updated: this page's script is orphaned */ }
  }
  addEventListener("pagehide", flush);
  D.addEventListener("visibilitychange", () => D.hidden && flush());
  // A reward plays its cutscene here, where it was earned
  const calm = () => S.motion == "off" || matchMedia("(prefers-reduced-motion:reduce)").matches;
  const reward = (r) => {
    const s = T[RW.SECRET];
    if (r.secret && s) self.WWCut.secret({ ...s, words: r.total, odds: RW.ODDS, calm: calm(), use: () => chrome.storage.sync.set({ theme: RW.SECRET, on: true }) });
    else if (r.badge) {
      const m = RW.MILESTONES.find((x) => x.id == r.badge);
      if (m) self.WWCut[m.cutscene == "67" ? "sixSeven" : "badge"]({ ...m, total: r.total, accent: t ? t.accent : s.accent, accent2: t ? t.accent2 : s.accent2, calm: calm() });
    }
  };

  // Typing sparks: 12 reusable dots, coloured once per theme, moved by the compositor.
  const G = "01<>{}#*+✦✧★";
  const onKey = (e) => {
    countKey(e);
    const now = e.timeStamp;
    if (!t || !S.sparks || now - last < 90 || e.ctrlKey || e.metaKey || e.altKey || (e.key.length > 1 && e.key != "Enter")) return;
    last = now;
    requestAnimationFrame(() => {
      // Word for the web draws its own text cursor (a thin element), not a real selection; use it
      // first, then the selection, and the mouse only as a last resort.
      let p = null;
      for (const n of D.querySelectorAll('[class*="ursor" i],[class*="aret" i],[id*="ursor" i]')) {
        const b = n.getBoundingClientRect();
        if (b.width <= 4 && b.height >= 8 && b.height <= 90 && b.top > 0 && b.left > 0) { p = [b.left + 2, b.top + b.height / 2]; break; }
      }
      if (!p) { const sel = getSelection(), r = sel.rangeCount && sel.getRangeAt(0).getClientRects()[0]; p = r && r.top > 0 ? [r.right, r.top + r.height / 2] : ptr; }
      if (!p || !t) return;
      if (!fx) {
        fx = D.createElement("div");
        fx.id = "ww-fx";
        fx.dataset.fx = t.fx || "dot";
        for (let i = 0; i < 12; i++) fx.append(D.createElement("b"));
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
  };
  addEventListener("keydown", onKey, true);
  // Google Docs reads typing in a hidden same-origin frame, so keys never reach this page: listen there too.
  let typingDoc = null;
  const hookTyping = () => {
    if (!GH) return;
    try {
      const d = (D.querySelector("iframe.docs-texteventtarget-iframe") || {}).contentDocument;
      if (d && d != typingDoc) { typingDoc = d; d.addEventListener("keydown", onKey, true); d.addEventListener("keydown", (e) => e.key == "Escape" && self.WWCut.esc(), true); }
    } catch (_) { /* not reachable: no sparks */ }
  };
  addEventListener("pointerdown", (e) => { ptr = [e.clientX, e.clientY]; }, { capture: true, passive: true });
  addEventListener("resize", () => later(300), { passive: true });
  addEventListener("scroll", () => later(400), { capture: true, passive: true });
  // Teams and Word swap whole panes without reloading (new chat, channel, tab): look again right away
  // throttled to one check per frame (not pushed back while Teams keeps redrawing), so a new pane is themed at once
  // A page that never stops changing (Docs, Teams) would otherwise be looked at every frame: first change at once,
  // then at most about six looks a second.
  let pend = 0, lastScan = 0;
  const look = () => { pend = 0; lastScan = performance.now(); scan(); };
  new MutationObserver(() => {
    if (pend) return;
    pend = 1;
    const wait = 160 - (performance.now() - lastScan);
    wait > 0 ? setTimeout(() => requestAnimationFrame(look), wait) : requestAnimationFrame(look);
  }).observe(D.documentElement, { childList: true, subtree: true });
  // What the page looks like to the extension: element names, sizes and colours only, never document text.
  const sig = (n) => {
    const r = n.getBoundingClientRect(), c = typeof n.className == "string" ? n.className.trim().split(/\s+/).slice(0, 3).join(".") : "";
    return `${n.tagName.toLowerCase()}${n.id ? "#" + n.id : ""}${c ? "." + c : ""} ${Math.round(r.left)},${Math.round(r.top)} ${Math.round(r.width)}x${Math.round(r.height)} bg=${getComputedStyle(n).backgroundColor}${n.hasAttribute("data-ww-clear") ? " CLEARED" : ""}${n.hasAttribute("data-ww-page") ? " PAGE" : ""}`;
  };
  const details = () => ({
    version: chrome.runtime.getManifest().version, page: HN + P.replace(/\/d\/[\w-]+/, "/d/ID"), app, theme: id, level, motion: S.motion, still: still(),
    scene: !!(bg && bg.isConnected), sceneBox: bg ? bg.style.cssText.slice(0, 200) : "", viewport: [innerWidth, innerHeight],
    editor: !!D.querySelector(".kix-appview-editor"), pages: D.querySelectorAll(".kix-page-paginated").length,
    rulers: [...D.querySelectorAll("#kix-horizontal-ruler-container,#kix-vertical-ruler,.docs-horizontal-ruler,.docs-vertical-ruler")].map(sig), dark: H_.hasAttribute("data-ww-dark"), clean: H_.hasAttribute("data-ww-clean"),
    canvases: [...D.querySelectorAll("canvas")].slice(0, 6).map((n) => sig(n) + " filter=" + getComputedStyle(n).filter),
    scrollbar: (() => { const e = D.querySelector(".kix-appview-editor"); return e ? e.offsetWidth - e.clientWidth : -1; })(),
    at: [[.5, .5], [.08, .5], [.92, .5], [.5, .2], [.5, .85]].map(([x, y]) => ({ x, y, stack: D.elementsFromPoint(innerWidth * x, innerHeight * y).slice(0, 12).map(sig) })),
  });
  chrome.runtime.onMessage.addListener((m, _, reply) => {
    if (m == "ww-status") reply({ docKey: key, active: id, file, google: GH, shown: !!(bg && bg.isConnected) });
    if (m == "ww-details") reply(details());
    // "Watch it again" in the popup: only the frame that shows the scene, and only once the theme has been found
    if (m && m.cutscene == "secret" && bg && bg.isConnected && innerWidth >= 400 && (S.found || {})[RW.SECRET])
      reward({ secret: RW.SECRET, total: S.found[RW.SECRET].words || 0 });
  });

  const load = () => chrome.storage.sync.get(DEF, (v) => chrome.storage.local.get(["customs", "community", "remote", "perf"], (c) => {
    S = { ...DEF, ...v };
    level = c.perf && Date.now() - c.perf.at < 14 * 864e5 ? c.perf.level : 0;     // re-tested after two weeks
    const before = id && !T[id] && JSON.stringify(custom[id] || R[id]);
    custom = {}; R = c.remote || {};                                   // R: built-in themes added on the website since
    (c.community || []).concat(c.customs || []).forEach((x) => (custom[x.id] = { ...x, custom: 1 }));
    if (before && before != JSON.stringify(custom[id] || R[id])) clear();   // rebuild only if that custom theme changed
    file = readFile() || file;
    apply();
  }));
  chrome.storage.onChanged.addListener((c) => Object.keys(c).some((k) => k != "words") && load());   // word counts change often and alter nothing here
  load();
  // Wild Themes for Word and Teams (the desktop .exe), when it's running, says which theme is on:
  // use the same one, so the browser, Teams and Word all match.
  fetch("http://127.0.0.1:47615/current").then((r) => r.json()).then((j) => chrome.storage.sync.get(DEF, (v) => {
    if (j.theme && T[j.theme] && j.theme !== v.theme) chrome.storage.sync.set({ theme: j.theme, on: true });
  })).catch(() => { /* app not running */ });
  [250, 600, 1200, 2000, 3500].forEach((ms) => setTimeout(probe, ms));
  setTimeout(() => probe(true), 5000);                              // no file theme by now: use the default
  [800, 2500, 6000].forEach((ms) => setTimeout(scan, ms));         // Word builds its page in stages
})();
