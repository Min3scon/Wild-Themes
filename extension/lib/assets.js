// Builds what a Word file needs to carry a theme: the backdrop picture, the readable
// panel behind the text, and a text-safe colour palette. Runs only when you click,
// in the options page or the Office add-in. Never inside Word while you type.
(() => {
  const cv = (w, h) => Object.assign(document.createElement("canvas"), { width: Math.round(w), height: Math.round(h) });
  const bytesOf = (c, type, q) => new Promise((ok) => c.toBlob(async (b) => ok(new Uint8Array(await b.arrayBuffer())), type, q));
  const rgb = (h) => { const n = parseInt(h.replace("#", "").slice(0, 6), 16); return [n >> 16 & 255, n >> 8 & 255, n & 255]; };
  const hex = (c) => c.map((v) => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, "0")).join("");
  const lum = (c) => { const f = (v) => (v /= 255) <= .03928 ? v / 12.92 : ((v + .055) / 1.055) ** 2.4; return .2126 * f(c[0]) + .7152 * f(c[1]) + .0722 * f(c[2]); };
  // darken a colour until it reads as text on white (contrast >= 4.5)
  const textSafe = (h) => { let c = rgb(h); while ((1.05) / (lum(c) + .05) < 4.5) c = c.map((v) => v * .9); return hex(c); };

  // Turn a custom theme into the same shape as a built-in one (palette for the document).
  function forCustom(t) {
    const a = t.accent || "#8a5cff", b = t.accent2 || a;
    return { ...t, custom: 1, doc: { dk1: "1E1B26", dk2: textSafe(a === b ? "#3a2f5a" : b), lt2: "F4F1FA", accent1: textSafe(a), accent2: textSafe(b),
      accent3: a.slice(1), accent4: b.slice(1), accent5: textSafe(a), accent6: b.slice(1), major: "Segoe UI", minor: "Aptos", panel: "#ffffff" } };
  }

  // Backdrop: built-in themes ship animated GIFs (portrait for Word, 16:9 for PowerPoint);
  // custom ones are drawn here. Excel gets a pale still version so cell text stays readable.
  async function backdrop(t, sec, base, imageB64, kind) {
    if (kind == "xlsx") return washed(t, base, imageB64);
    if (!t.custom) try {
      // themes added on the website since this copy was made carry their own picture addresses
      const r = await fetch((kind == "pptx" ? t.backdropWideUrl : t.backdropUrl) || base + (kind == "pptx" ? "backdrops-wide/" : "backdrops/") + t.id + ".gif");
      if (r.ok) return { bytes: new Uint8Array(await r.arrayBuffer()), ext: "gif" };
    } catch (_) { /* no shipped backdrop: draw one below */ }
    const w = 794, h = Math.round(794 * sec.pageH / sec.pageW), c = cv(w, h), g = c.getContext("2d");
    if (t.image && imageB64) {
      const img = await createImageBitmap(await (await fetch(imageB64)).blob());
      const s = Math.max(w / img.width, h / img.height);
      g.drawImage(img, (w - img.width * s) / 2, (h - img.height * s) / 2, img.width * s, img.height * s);
      g.fillStyle = t.mode === "light" ? `rgba(255,255,255,${t.dim || 0})` : `rgba(0,0,0,${t.dim || 0})`;
      g.fillRect(0, 0, w, h);
      return { bytes: await bytesOf(c, "image/jpeg", .86), ext: "jpeg" };
    }
    const col = t.colors || ["#120a2a", t.accent || "#ff2bd6", t.accent2 || "#21e6ff"];
    g.fillStyle = col[0]; g.fillRect(0, 0, w, h);
    for (const [x, y, k] of [[0, 0, 1], [w, h, 2]]) {
      const r = g.createRadialGradient(x, y, 0, x, y, w * 1.1);
      r.addColorStop(0, col[k]); r.addColorStop(1, col[k] + "00"); g.fillStyle = r; g.fillRect(0, 0, w, h);
    }
    return { bytes: await bytesOf(c, "image/png"), ext: "png" };
  }

  // Excel sheet background: the scene's first frame washed 80% towards white, 1280x720.
  async function washed(t, base, imageB64) {
    const c = cv(1280, 720), g = c.getContext("2d");
    const src = await backdrop(t, { pageW: 16, pageH: 9 }, base, imageB64, "pptx");
    const img = await createImageBitmap(new Blob([src.bytes]));
    const s = Math.max(1280 / img.width, 720 / img.height);
    g.drawImage(img, (1280 - img.width * s) / 2, (720 - img.height * s) / 2, img.width * s, img.height * s);
    g.fillStyle = "rgba(255,255,255,.8)"; g.fillRect(0, 0, 1280, 720);
    return { bytes: await bytesOf(c, "image/jpeg", .85), ext: "jpeg" };
  }

  // Teams meeting background: 1920x1080 still of the scene.
  async function teams(t, base, imageB64) {
    if (!t.custom) try {
      const r = await fetch(t.teamsUrl || base + "teams-backgrounds/" + t.id + ".jpg");
      if (r.ok) return new Uint8Array(await r.arrayBuffer());
    } catch (_) { /* draw one */ }
    const c = cv(1920, 1080), g = c.getContext("2d");
    const img = await createImageBitmap(new Blob([(await backdrop(t, { pageW: 16, pageH: 9 }, base, imageB64, "pptx")).bytes]));
    const s = Math.max(1920 / img.width, 1080 / img.height);
    g.drawImage(img, (1920 - img.width * s) / 2, (1080 - img.height * s) / 2, img.width * s, img.height * s);
    return bytesOf(c, "image/jpeg", .88);
  }

  // Panel: a light card with a soft accent glow, sized to the stamper's panel rectangle.
  // Slides get a slightly see-through one so the scene still shows a little.
  async function panel(t, rect, alpha = 1) {
    const s = 96 / 1440, pad = 14, c = cv(rect.w * s, rect.h * s), g = c.getContext("2d");
    g.shadowColor = (t.accent || "#8a5cff") + "aa"; g.shadowBlur = 12;
    g.globalAlpha = alpha;
    g.fillStyle = (t.doc && t.doc.panel) || "#ffffff";
    g.beginPath(); g.roundRect(pad, pad, c.width - pad * 2, c.height - pad * 2, 18); g.fill();
    if (alpha > 0) { g.globalAlpha = 1; g.shadowBlur = 0; g.lineWidth = 2; g.strokeStyle = t.accent || "#8a5cff"; g.stroke(); }   // 0 = no panel at all
    return { bytes: await bytesOf(c, "image/png") };
  }

  // One place that knows which stamper and which pictures each file type needs.
  const kindOf = (name) => (/\.pptx$/i.test(name) ? "pptx" : /\.xls[xm]$/i.test(name) ? "xlsx" : "docx");
  const stamperFor = (kind) => ({ docx: self.WWStamper, pptx: self.WWStamperPptx, xlsx: self.WWStamperXlsx })[kind];
  // Word plays an animated GIF as a video, drawn ON TOP of the text with a play button,
  // whatever its wrap setting. So Word files always get the scene's first frame as a JPEG.
  async function stillJpeg(b) {
    const img = await createImageBitmap(new Blob([b.bytes], { type: "image/gif" }));
    const c = cv(img.width, img.height); c.getContext("2d").drawImage(img, 0, 0);
    return { bytes: await bytesOf(c, "image/jpeg", .88), ext: "jpeg" };
  }
  // settings (WTCore.settingsFor): backdrop off = no pictures; panel = panel opacity (60-100);
  // animated (PowerPoint) = keep the moving GIF, otherwise its first frame as a still.
  async function assetsFor(t, kind, sec, base, imageB64, settings) {
    const s = settings || {};
    if (s.backdrop === false) return { backdrop: null, panel: null };
    let bd = await backdrop(t, sec || { pageW: 11906, pageH: 16838 }, base, imageB64, kind);
    if (bd.ext == "gif" && (kind == "docx" || (kind == "pptx" && s.animated === false))) bd = await stillJpeg(bd);
    const alpha = (s.panel ?? (kind == "pptx" ? 90 : 100)) / 100;
    return { backdrop: bd, panel: kind == "xlsx" ? null : async (_s, rect) => (await panel(t, rect, alpha)).bytes };
  }
  // The picture-maker WTCore.apply() wants. imageFor(theme) returns a custom theme's picture (data URL) or null.
  const maker = (base, imageFor) => async (t, kind, sec, settings) =>
    assetsFor(t.custom || !self.WW_THEMES[t.id] ? (t.doc ? t : forCustom(t)) : t, kind, sec, base, imageFor ? await imageFor(t) : null, settings);

  self.WWAssets = { forCustom, backdrop, panel, washed, teams, textSafe, kindOf, stamperFor, assetsFor, maker };
})();
