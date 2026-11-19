// Wild Themes — writes a theme INTO a PowerPoint (.pptx) file, deeply and reversibly.
// The theme lives in the presentation itself: theme colours and fonts, the masters' text styles
// (titles, body text, bullets), the default table style, theme colours that had been baked into
// shapes and charts, slide transitions, and a backdrop plus a see-through reading panel on every
// slide master (locked, behind everything), plus file metadata. Anyone who opens the file sees it,
// with nothing installed. Everything changed is recorded in a backup, so unstamp() can put the
// original back exactly.
//
// Browser: load vendor/jszip.min.js, this file and wtcore.js -> self.WWStamperPptx.
// Node:    const S = require('./stamper-pptx.js')  (needs jszip + @xmldom/xmldom).
(function (root, factory) {
  if (typeof module === "object" && module.exports) {
    let xmldom = null;
    try { xmldom = require("@xmldom/xmldom"); } catch (_) { /* rely on globals */ }
    module.exports = factory(require("jszip"), xmldom);
  } else {
    root.WWStamperPptx = factory(root.JSZip, null);
  }
})(typeof self !== "undefined" ? self : this, function (JSZip, xmldom) {
  "use strict";

  const DOMParserImpl = (typeof DOMParser !== "undefined" && DOMParser) || (xmldom && xmldom.DOMParser);
  const SerializerImpl = (typeof XMLSerializer !== "undefined" && XMLSerializer) || (xmldom && xmldom.XMLSerializer);
  // The shared core (palette, settings), fetched at call time: wtcore.js itself requires this file.
  const core = () => {
    const c = (typeof self !== "undefined" && self.WTCore) || (typeof require === "function" && require("./wtcore.js"));
    if (!c) throw new Error("Wild Themes: load wtcore.js before using the PowerPoint stamper.");
    return c;
  };

  const NS = {
    p: "http://schemas.openxmlformats.org/presentationml/2006/main",
    a: "http://schemas.openxmlformats.org/drawingml/2006/main",
    r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    pr: "http://schemas.openxmlformats.org/package/2006/relationships",
    ct: "http://schemas.openxmlformats.org/package/2006/content-types",
    cp: "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
    vt: "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
    dc: "http://purl.org/dc/elements/1.1/",
    mc: "http://schemas.openxmlformats.org/markup-compatibility/2006",
  };
  const REL = {
    officeDocument: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
    slideMaster: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster",
    slideLayout: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
    slide: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
    theme: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
    image: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
    chart: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
    tableStyles: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles",
    custom: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties",
  };
  const CT_CUSTOM = "application/vnd.openxmlformats-officedocument.custom-properties+xml";
  const CT_TABLESTYLES = "application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml";
  const IMG_CT = { png: "image/png", gif: "image/gif", jpeg: "image/jpeg", jpg: "image/jpeg" };
  const TWIP_EMU = 635;
  const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  const FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";
  const PROPS = ["WildTheme", "WildThemeName", "WildThemeVersion"];
  const FORMAT_VERSION = "4";
  // Built-in table style that follows the theme: Medium Style 2 - Accent 1 (accent header row, tinted banding).
  const TABLE_STYLE = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}";
  const TRANSITIONS = {
    fade: { el: "fade", attrs: {} }, push: { el: "push", attrs: { dir: "u" } }, wipe: { el: "wipe", attrs: { dir: "r" } },
    split: { el: "split", attrs: { orient: "vert", dir: "out" } }, dissolve: { el: "dissolve", attrs: {} },
  };
  // Theme slots whose fixed (srgbClr) copies are turned back into theme colours ("Match old colours").
  const DIRECT_SLOTS = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "dk2", "hlink"];
  // Office's own default themes (2007, 2013-2022, 2023+): shapes pasted from such decks keep these as fixed
  // colours, so they count as theme colours too (the file's own theme wins). Left out on purpose: values that are
  // also on Office's "Standard Colors" row or are plain grey / blue (FFC000, A5A5A5, 0000FF), and black / white.
  const OFFICE_SLOTS = {
    accent1: ["4F81BD", "4472C4", "156082"], accent2: ["C0504D", "ED7D31", "E97132"], accent3: ["9BBB59", "196B24"],
    accent4: ["8064A2", "0F9ED5"], accent5: ["4BACC6", "5B9BD5", "A02B93"], accent6: ["F79646", "70AD47", "4EA72E"],
    dk2: ["1F497D", "44546A", "0E2841"], hlink: ["0563C1", "467886"],
  };
  const NEVER_MAP = new Set(["000000", "FFFFFF"]);

  // ------------------------------------------------------------------ helpers
  const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const hex = (h, fb) => {
    h = String(h || fb || "000000").replace("#", "").toUpperCase();
    if (h.length === 3) h = h.split("").map((c) => c + c).join("");
    return /^[0-9A-F]{6}$/.test(h) ? h : String(fb || "000000").replace("#", "").toUpperCase();
  };
  function parseXml(str) {
    const doc = new DOMParserImpl().parseFromString(str, "application/xml");
    const err = doc.getElementsByTagName("parsererror");
    if (err && err.length) throw new Error("Could not read the PowerPoint file (XML error).");
    return doc;
  }
  function serialize(doc) {
    return XML_DECL + new SerializerImpl().serializeToString(doc).replace(/^<\?xml[^>]*\?>\s*/, "");
  }
  const serializeEl = (el) => new SerializerImpl().serializeToString(el);
  function fragment(doc, xml) {
    const f = parseXml(`<root xmlns:p="${NS.p}" xmlns:a="${NS.a}" xmlns:r="${NS.r}">${xml}</root>`).documentElement;
    const out = [];
    for (let n = f.firstChild; n; n = n.nextSibling) if (n.nodeType === 1) out.push(doc.importNode(n, true));
    return out;
  }
  const kids = (el, ns, name) => {
    const out = [];
    for (let n = el && el.firstChild; n; n = n.nextSibling) if (n.nodeType === 1 && n.namespaceURI === ns && n.localName === name) out.push(n);
    return out;
  };
  const kid = (el, ns, name) => kids(el, ns, name)[0] || null;
  const elKids = (el) => { const out = []; for (let n = el && el.firstChild; n; n = n.nextSibling) if (n.nodeType === 1) out.push(n); return out; };
  const all = (el, ns, name) => Array.from(el.getElementsByTagNameNS(ns, name));
  // Insert `el` before the first child named in `before` (keeps the schema's element order), else at the end.
  const insertBeforeAny = (parent, el, before) => parent.insertBefore(el, elKids(parent).find((c) => before.includes(c.localName)) || null);
  const make = (doc, ns, prefix, name) => doc.createElementNS(ns, prefix + ":" + name);
  const dirOf = (p) => (p.includes("/") ? p.slice(0, p.lastIndexOf("/")) : "");
  const relsPathOf = (p) => (dirOf(p) ? dirOf(p) + "/" : "") + "_rels/" + p.split("/").pop() + ".rels";
  function resolveTarget(fromPart, target) {
    if (target.startsWith("/")) return target.slice(1);
    const parts = dirOf(fromPart) ? dirOf(fromPart).split("/") : [];
    for (const seg of target.split("/")) { if (seg === "..") parts.pop(); else if (seg && seg !== ".") parts.push(seg); }
    return parts.join("/");
  }
  function relativeTarget(fromPart, toPart) {
    const from = dirOf(fromPart) ? dirOf(fromPart).split("/") : [];
    const to = toPart.split("/");
    let i = 0;
    while (i < from.length && i < to.length - 1 && from[i] === to[i]) i++;
    return "../".repeat(from.length - i) + to.slice(i).join("/");
  }
  // Short content hash for media names: the same picture always gets the same name (deterministic re-stamps).
  function fnv(bytes) {
    const u = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    let h = 2166136261;
    for (let i = 0; i < u.length; i++) h = Math.imul(h ^ u[i], 16777619);
    return (h >>> 0).toString(16).padStart(8, "0");
  }
  // Namespace declarations repeated from an ancestor (left behind when an element is serialised on its own and
  // put back) are dropped, so a restored element serialises exactly as the original did.
  function dropRedundantNs(el) {
    const inScope = (node, attr) => { for (let n = node; n && n.nodeType === 1; n = n.parentNode) if (n.hasAttribute(attr)) return n.getAttribute(attr); return null; };
    for (const a of Array.from(el.attributes)) {
      if (a.name !== "xmlns" && !a.name.startsWith("xmlns:")) continue;
      if (inScope(el.parentNode, a.name) === a.value) el.removeAttribute(a.name);
    }
    for (const c of elKids(el)) dropRedundantNs(c);
  }
  // WCAG contrast ratio of two hex colours.
  function contrast(h1, h2) {
    const L = (h) => { const c = [0, 2, 4].map((i) => parseInt(h.slice(i, i + 2), 16) / 255).map((v) => (v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4)); return 0.2126 * c[0] + 0.7152 * c[1] + 0.0722 * c[2]; };
    const a = L(hex(h1)), b = L(hex(h2));
    return (Math.max(a, b) + 0.05) / (Math.min(a, b) + 0.05);
  }

  // ------------------------------------------------------------------ package wrapper
  class Pkg {
    constructor(zip) { this.zip = zip; this.cache = new Map(); this.dirty = new Set(); }
    static async open(bytes) {
      let zip;
      try { zip = await JSZip.loadAsync(bytes); } catch (_) { throw new Error("This isn't a PowerPoint file (.pptx)."); }
      if (!zip.file("[Content_Types].xml")) throw new Error("This isn't a PowerPoint file (.pptx).");
      return new Pkg(zip);
    }
    async raw(p) { const f = this.zip.file(p); return f ? f.async("string") : null; }
    async xml(p) {
      if (this.cache.has(p)) return this.cache.get(p);
      const f = this.zip.file(p);
      if (!f) return null;
      const doc = parseXml(await f.async("string"));
      this.cache.set(p, doc);
      return doc;
    }
    setXml(p, doc) { this.cache.set(p, doc); this.dirty.add(p); }
    touch(p) { this.dirty.add(p); }
    putText(p, str) { this.cache.delete(p); this.dirty.delete(p); this.zip.file(p, str, { createFolders: false }); }
    putBinary(p, bytes) { this.zip.file(p, bytes, { binary: true, createFolders: false, compression: "STORE" }); }
    remove(p) { this.zip.remove(p); this.cache.delete(p); this.dirty.delete(p); }
    has(p) { return !!this.zip.file(p); }
    async rels(part) {
      const p = relsPathOf(part);
      let doc = await this.xml(p);
      // a missing .rels file is only written once something is added to it
      if (!doc) { doc = parseXml(`<Relationships xmlns="${NS.pr}"/>`); this.cache.set(p, doc); }
      const root = doc.documentElement, self_ = this;
      const list = () => Array.from(root.getElementsByTagName("Relationship")).map((el) => ({
        el, id: el.getAttribute("Id"), type: el.getAttribute("Type"), target: el.getAttribute("Target") || "",
        external: el.getAttribute("TargetMode") === "External",
      }));
      return {
        list,
        add(type, target) {
          const ids = new Set(list().map((r) => r.id));
          let n = 1; while (ids.has("rIdWT" + n)) n++;
          const el = doc.createElementNS(NS.pr, "Relationship");
          el.setAttribute("Id", "rIdWT" + n); el.setAttribute("Type", type); el.setAttribute("Target", target);
          root.appendChild(el); self_.touch(p);
          return "rIdWT" + n;
        },
        remove(id) { for (const r of list()) if (r.id === id) { root.removeChild(r.el); self_.touch(p); } },
        byType(type) { return list().filter((r) => r.type === type && !r.external); },
        target(id) { const r = list().find((x) => x.id === id && !x.external); return r ? resolveTarget(part, r.target) : null; },
      };
    }
    async contentTypes() {
      const doc = await this.xml("[Content_Types].xml");
      const root = doc.documentElement, self_ = this;
      const defaults = () => Array.from(root.getElementsByTagName("Default"));
      return {
        defaults: () => defaults().map((d) => (d.getAttribute("Extension") || "").toLowerCase()),
        ensureDefault(ext, type) {
          if (defaults().some((d) => (d.getAttribute("Extension") || "").toLowerCase() === ext)) return;
          const el = doc.createElementNS(NS.ct, "Default");
          el.setAttribute("Extension", ext); el.setAttribute("ContentType", type);
          root.insertBefore(el, root.getElementsByTagName("Override")[0] || null); self_.touch("[Content_Types].xml");
        },
        removeDefault(ext) {
          for (const d of defaults()) if ((d.getAttribute("Extension") || "").toLowerCase() === ext) { root.removeChild(d); self_.touch("[Content_Types].xml"); }
        },
        ensureOverride(part, type) {
          const ex = Array.from(root.getElementsByTagName("Override")).find((o) => o.getAttribute("PartName") === "/" + part);
          if (ex) return;
          const el = doc.createElementNS(NS.ct, "Override");
          el.setAttribute("PartName", "/" + part); el.setAttribute("ContentType", type);
          root.appendChild(el); self_.touch("[Content_Types].xml");
        },
        removeOverride(part) {
          for (const o of Array.from(root.getElementsByTagName("Override"))) if (o.getAttribute("PartName") === "/" + part) { root.removeChild(o); self_.touch("[Content_Types].xml"); }
        },
      };
    }
    async save() {
      for (const p of this.dirty) if (this.cache.has(p)) this.zip.file(p, serialize(this.cache.get(p)), { createFolders: false });
      for (const name of Object.keys(this.zip.files)) { const f = this.zip.files[name]; if (f && f.dir) delete this.zip.files[name]; }
      return this.zip.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 6 } });
    }
  }

  // ------------------------------------------------------------------ the presentation's structure
  async function structure(pkg) {
    const rootRels = await pkg.rels("");
    const r = rootRels.byType(REL.officeDocument)[0];
    const pres = r && resolveTarget("", r.target);
    const presDoc = pres && await pkg.xml(pres);
    if (!presDoc || presDoc.documentElement.namespaceURI !== NS.p || presDoc.documentElement.localName !== "presentation") {
      throw new Error("This isn't a PowerPoint file (.pptx).");
    }
    const presRels = await pkg.rels(pres);
    const P = presDoc.documentElement;
    // masters and slides in the presentation's own order (its id lists), falling back to relationship order
    const ordered = (listName, itemName, relType) => {
      const out = [], ids = new Map();
      const lst = kid(P, NS.p, listName);
      for (const it of kids(lst, NS.p, itemName)) {
        const t = presRels.target(it.getAttributeNS(NS.r, "id") || it.getAttribute("r:id") || "");
        if (t && !out.includes(t)) { out.push(t); ids.set(t, it.getAttribute("id")); }
      }
      for (const x of presRels.byType(relType)) { const t = resolveTarget(pres, x.target); if (!out.includes(t)) out.push(t); }
      return { out, ids };
    };
    const M = ordered("sldMasterIdLst", "sldMasterId", REL.slideMaster), SL = ordered("sldIdLst", "sldId", REL.slide);
    const masters = M.out, slides = SL.out;
    const layouts = [], themes = [], themeOf = new Map(), masterOf = new Map(), layoutOf = new Map();
    for (const m of masters) {
      const mr = await pkg.rels(m);
      const t = mr.byType(REL.theme)[0];
      if (t) { const tp = resolveTarget(m, t.target); themeOf.set(m, tp); if (!themes.includes(tp)) themes.push(tp); }
      for (const l of mr.byType(REL.slideLayout)) { const lp = resolveTarget(m, l.target); if (!layouts.includes(lp)) layouts.push(lp); if (!masterOf.has(lp)) masterOf.set(lp, m); }
    }
    for (const l of layouts) { const m = (await pkg.rels(l)).byType(REL.slideMaster)[0]; if (m) masterOf.set(l, resolveTarget(l, m.target)); }
    for (const sl of slides) { const l = (await pkg.rels(sl)).byType(REL.slideLayout)[0]; if (l) layoutOf.set(sl, resolveTarget(sl, l.target)); }
    // masters actually used by a slide (slide -> layout -> master); with no slides, all of them
    let used = slides.length ? masters.filter((m) => slides.some((sl) => masterOf.get(layoutOf.get(sl)) === m)) : masters.slice();
    if (!used.length) used = masters.slice();                     // broken links: don't hide every theme
    const usedThemes = [];
    for (const m of used) { const tp = themeOf.get(m); if (tp && !usedThemes.includes(tp)) usedThemes.push(tp); }
    // charts (and their colour/style parts), with the slide or layout that shows them
    const charts = [], ownerOf = new Map();
    for (const owner of [...slides, ...layouts]) {
      for (const c of (await pkg.rels(owner)).byType(REL.chart)) {
        const cp = resolveTarget(owner, c.target);
        if (!ownerOf.has(cp)) ownerOf.set(cp, owner);
        if (!pkg.has(cp)) continue;
        for (const x of (await pkg.rels(cp)).list()) { const xp = resolveTarget(cp, x.target); if (!x.external && !ownerOf.has(xp)) ownerOf.set(xp, owner); }
      }
    }
    for (const name of Object.keys(pkg.zip.files)) if (/^ppt\/charts\/[^/]+\.xml$/.test(name)) charts.push(name);
    charts.sort();

    const sz = kid(P, NS.p, "sldSz");
    const cx = parseInt(sz && sz.getAttribute("cx"), 10) || 12192000, cy = parseInt(sz && sz.getAttribute("cy"), 10) || 6858000;
    const section = { pageW: Math.round(cx / TWIP_EMU), pageH: Math.round(cy / TWIP_EMU) };
    // where the text placeholders sit (master + layouts), so the reading panel can cover them
    let box = null;
    for (const part of [...masters, ...layouts]) {
      const doc = await pkg.xml(part);
      if (!doc) continue;
      for (const sp of all(doc, NS.p, "sp")) {
        const ph = all(sp, NS.p, "ph")[0];
        if (!ph || ["dt", "ftr", "sldNum"].includes(ph.getAttribute("type"))) continue;
        const spPr = kid(sp, NS.p, "spPr"), xf = spPr && kid(spPr, NS.a, "xfrm");
        const off = xf && kid(xf, NS.a, "off"), ext = xf && kid(xf, NS.a, "ext");
        if (!off || !ext) continue;
        const r = { x: +off.getAttribute("x"), y: +off.getAttribute("y"), r: +off.getAttribute("x") + +ext.getAttribute("cx"), b: +off.getAttribute("y") + +ext.getAttribute("cy") };
        if (![r.x, r.y, r.r, r.b].every(Number.isFinite)) continue;
        box = box ? { x: Math.min(box.x, r.x), y: Math.min(box.y, r.y), r: Math.max(box.r, r.r), b: Math.max(box.b, r.b) } : r;
      }
    }
    if (box) section.content = { x: Math.round(box.x / TWIP_EMU), y: Math.round(box.y / TWIP_EMU), w: Math.round((box.r - box.x) / TWIP_EMU), h: Math.round((box.b - box.y) / TWIP_EMU) };

    // Slides are recorded in the backup by their slide id, which survives PowerPoint renumbering the parts.
    const slideById = new Map();
    for (const [part, id] of SL.ids) slideById.set(id, part);
    const keyOf = (part) => (SL.ids.has(part) ? "sld:" + SL.ids.get(part) : part);
    const partOf = (key) => (key.startsWith("sld:") ? slideById.get(key.slice(4)) || null : key);
    const themeFor = (part) => {
      const owner = ownerOf.get(part) || part;
      const m = masterOf.get(layoutOf.get(owner)) || masterOf.get(owner) || (masters.includes(owner) ? owner : null);
      return (m && themeOf.get(m)) || themes[0] || null;
    };
    return { pres, masters, layouts, slides, themes, usedThemes, themeOf, charts, section, keyOf, partOf, themeFor };
  }

  // Reading panel: most of the slide (about 4% in from the sides, 7% from top and bottom), grown
  // to cover the layouts' text placeholders, always leaving at least a 1.5% border of backdrop.
  function panelRect(section) {
    const W = section.pageW, H = section.pageH, pad = 173;            // 0.12 in around the text
    let x = W * 0.04, y = H * 0.07, r = W - x, b = H - y;
    const c = section.content;
    if (c) { x = Math.min(x, c.x - pad); y = Math.min(y, c.y - pad); r = Math.max(r, c.x + c.w + pad); b = Math.max(b, c.y + c.h + pad); }
    x = Math.max(x, W * 0.015); y = Math.max(y, H * 0.015); r = Math.min(r, W * 0.985); b = Math.min(b, H * 0.985);
    return { x: Math.round(x), y: Math.round(y), w: Math.round(r - x), h: Math.round(b - y) };
  }

  // ------------------------------------------------------------------ theme colours and fonts
  const SCHEME_SLOTS = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
  // Colour + font scheme from the shared palette (wtcore.js), so Word, PowerPoint and Excel agree.
  function applyTheme(doc, pal, name) {
    const els = all(doc, NS.a, "themeElements")[0];
    if (!els) return false;
    const [clr] = fragment(doc, `<a:clrScheme name="${esc(name)}">${SCHEME_SLOTS.map((k) => `<a:${k}><a:srgbClr val="${pal[k]}"/></a:${k}>`).join("")}</a:clrScheme>`);
    const oldClr = kid(els, NS.a, "clrScheme");
    if (oldClr) els.replaceChild(clr, oldClr); else els.insertBefore(clr, els.firstChild);
    // fonts: the Latin typefaces change; the per-script fallbacks (Japanese, Arabic, ...) are kept
    let fs = kid(els, NS.a, "fontScheme");
    if (!fs || !kid(fs, NS.a, "majorFont") || !kid(fs, NS.a, "minorFont")) {
      const blank = `<a:latin typeface=""/><a:ea typeface=""/><a:cs typeface=""/>`;
      const [fresh] = fragment(doc, `<a:fontScheme name=""><a:majorFont>${blank}</a:majorFont><a:minorFont>${blank}</a:minorFont></a:fontScheme>`);
      if (fs) els.replaceChild(fresh, fs); else els.insertBefore(fresh, kid(els, NS.a, "fmtScheme"));
      fs = fresh;
    }
    fs.setAttribute("name", name);
    for (const [k, face] of [["majorFont", pal.major], ["minorFont", pal.minor]]) {
      const f = kid(fs, NS.a, k);
      let latin = kid(f, NS.a, "latin");
      if (!latin) { latin = make(doc, NS.a, "a", "latin"); f.insertBefore(latin, f.firstChild); }
      for (const at of ["panose", "pitchFamily", "charset"]) latin.removeAttribute(at);
      latin.setAttribute("typeface", face);
    }
    return true;
  }
  // value -> slot for the colours that count as theme colours in a theme part (see DIRECT_SLOTS / OFFICE_SLOTS)
  function themeColours(doc) {
    const out = new Map();
    const cs = doc && all(doc, NS.a, "clrScheme")[0];
    if (cs) for (const slot of DIRECT_SLOTS) {
      const v = elKids(kid(cs, NS.a, slot))[0];
      if (!v) continue;
      const val = String((v.localName === "sysClr" ? v.getAttribute("lastClr") : v.getAttribute("val")) || "").toUpperCase();
      if (/^[0-9A-F]{6}$/.test(val) && !NEVER_MAP.has(val) && !out.has(val)) out.set(val, slot);
    }
    for (const [slot, vals] of Object.entries(OFFICE_SLOTS)) for (const v of vals) if (!out.has(v)) out.set(v, slot);
    return out;
  }

  // ------------------------------------------------------------------ text styles (master txStyles)
  const FILLS = ["noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"];
  const AFTER_FILL = ["effectLst", "effectDag", "highlight", "uLnTx", "uLn", "uFillTx", "uFill", "latin", "ea", "cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst"];
  const AFTER_LATIN = ["ea", "cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst"];
  const AFTER_BUCLR = ["buSzTx", "buSzPct", "buSzPts", "buFontTx", "buFont", "buNone", "buAutoNum", "buChar", "buBlip", "tabLst", "defRPr", "extLst"];
  const LVLS = [1, 2, 3, 4, 5, 6, 7, 8, 9].map((i) => `lvl${i}pPr`);
  const solid = (doc, slot) => fragment(doc, `<a:solidFill><a:schemeClr val="${slot}"/></a:solidFill>`)[0];
  function styleLevels(doc, list, { font, colour, bullet }) {
    if (!kid(list, NS.a, "lvl1pPr")) insertBeforeAny(list, make(doc, NS.a, "a", "lvl1pPr"), [...LVLS.slice(1), "extLst"]);
    for (const name of LVLS) {
      const lvl = kid(list, NS.a, name);
      if (!lvl) continue;
      let rpr = kid(lvl, NS.a, "defRPr");
      if (!rpr) { rpr = make(doc, NS.a, "a", "defRPr"); insertBeforeAny(lvl, rpr, ["extLst"]); }
      for (const f of elKids(rpr).filter((c) => c.namespaceURI === NS.a && FILLS.includes(c.localName))) rpr.removeChild(f);
      insertBeforeAny(rpr, solid(doc, colour), AFTER_FILL);
      let latin = kid(rpr, NS.a, "latin");
      if (!latin) { latin = make(doc, NS.a, "a", "latin"); insertBeforeAny(rpr, latin, AFTER_LATIN); }
      for (const at of ["panose", "pitchFamily", "charset"]) latin.removeAttribute(at);
      latin.setAttribute("typeface", font);
      if (bullet && !kid(lvl, NS.a, "buNone")) {
        for (const b of [...kids(lvl, NS.a, "buClr"), ...kids(lvl, NS.a, "buClrTx")]) lvl.removeChild(b);
        const bc = make(doc, NS.a, "a", "buClr");
        bc.appendChild(fragment(doc, `<a:schemeClr val="${bullet}"/>`)[0]);
        insertBeforeAny(lvl, bc, AFTER_BUCLR);
      }
    }
  }
  // Titles: major font in accent 1 (or the second text colour when accent 1 is too pale for the panel).
  // Body and other text: minor font in the main text colour; bullets in accent 1.
  function styleMaster(doc, pal) {
    const M = doc.documentElement;
    let tx = kid(M, NS.p, "txStyles");
    if (!tx) { tx = make(doc, NS.p, "p", "txStyles"); insertBeforeAny(M, tx, ["extLst"]); }
    const cm = kid(M, NS.p, "clrMap");
    // tx1/tx2 only while the master maps them to dark colours (the reading panel is light)
    const dark = (slot, fb) => { const v = cm && cm.getAttribute(slot); return !v || v === "dk1" || v === "dk2" ? slot : fb; };
    const spec = {
      titleStyle: { font: "+mj-lt", colour: contrast(pal.accent1, pal.panel) >= 3 ? "accent1" : dark("tx2", "dk2") },
      bodyStyle: { font: "+mn-lt", colour: dark("tx1", "dk1"), bullet: "accent1" },
      otherStyle: { font: "+mn-lt", colour: dark("tx1", "dk1") },
    };
    const order = ["titleStyle", "bodyStyle", "otherStyle", "extLst"];
    for (const [name, o] of Object.entries(spec)) {
      let list = kid(tx, NS.p, name);
      if (!list) { list = make(doc, NS.p, "p", name); insertBeforeAny(tx, list, order.slice(order.indexOf(name) + 1)); }
      styleLevels(doc, list, o);
    }
  }

  // ------------------------------------------------------------------ fixed colours -> theme colours
  // Every DrawingML colour element in document order; positions are what the backup records.
  function colourEls(doc) {
    const out = [];
    const walk = (el) => {
      for (let n = el.firstChild; n; n = n.nextSibling) {
        if (n.nodeType !== 1) continue;
        if (n.namespaceURI === NS.a && (n.localName === "srgbClr" || n.localName === "schemeClr")) out.push(n);
        walk(n);
      }
    };
    walk(doc.documentElement);
    return out;
  }
  // Swap a colour element for another kind, keeping its modifiers (lumMod, alpha, ...).
  function swapColour(doc, el, localName, val) {
    const fresh = doc.createElementNS(NS.a, (el.prefix ? el.prefix + ":" : "") + localName);
    fresh.setAttribute("val", val);
    while (el.firstChild) fresh.appendChild(el.firstChild);
    el.parentNode.replaceChild(fresh, el);
  }

  // ------------------------------------------------------------------ slide transitions
  const isEl = (n, ns, name) => n && n.nodeType === 1 && n.namespaceURI === ns && n.localName === name;
  const transitionsIn = (el) => (isEl(el, NS.p, "transition") ? [el] : isEl(el, NS.mc, "AlternateContent") ? all(el, NS.p, "transition") : []);
  const hasTransition = (doc) => elKids(doc.documentElement).some((c) => transitionsIn(c).length > 0);
  function addTransition(doc, kind) {
    const t = TRANSITIONS[kind];
    const attrs = Object.entries(t.attrs).map(([k, v]) => ` ${k}="${v}"`).join("");
    const [el] = fragment(doc, `<p:transition spd="med"><p:${t.el}${attrs}/></p:transition>`);
    // schema order: cSld, clrMapOvr, transition, timing, extLst
    const S = doc.documentElement;
    const ref = elKids(S).find((c) => isEl(c, NS.p, "timing") || isEl(c, NS.p, "extLst") || (isEl(c, NS.mc, "AlternateContent") && all(c, NS.p, "timing").length));
    S.insertBefore(el, ref || null);
  }
  // One of ours, as written (or as PowerPoint re-saves it, with a p14 duration): speed "med" + one of our effects.
  function isOurTransition(tr) {
    const attrs = Array.from(tr.attributes).filter((a) => !/^xmlns(:|$)/.test(a.name) && !(a.localName === "dur" && a.prefix === "p14"));
    if (attrs.length !== 1 || attrs[0].name !== "spd" || attrs[0].value !== "med") return false;
    const fx = elKids(tr);
    if (fx.length !== 1 || fx[0].namespaceURI !== NS.p) return false;
    return Object.values(TRANSITIONS).some((t) => {
      if (fx[0].localName !== t.el) return false;
      const got = Array.from(fx[0].attributes).filter((a) => !/^xmlns(:|$)/.test(a.name));
      return got.length === Object.keys(t.attrs).length && got.every((a) => t.attrs[a.name] === a.value);
    });
  }

  // ------------------------------------------------------------------ table style
  async function tableStylesPart(pkg, s) {
    const rels = await pkg.rels(s.pres);
    const r = rels.byType(REL.tableStyles)[0];
    return { rels, rel: r || null, part: r ? resolveTarget(s.pres, r.target) : (dirOf(s.pres) ? dirOf(s.pres) + "/" : "") + "tableStyles.xml" };
  }

  // ------------------------------------------------------------------ pictures
  const markerDescr = (theme) =>
    `Wild Theme: ${theme.id}; name=${theme.name}; accent=#${hex(theme.accent)}; accent2=#${hex(theme.accent2)}; mode=${theme.mode || "dark"}`;
  const LOCKS = 'noGrp="1" noSelect="1" noRot="1" noChangeAspect="1" noMove="1" noResize="1" noEditPoints="1" noAdjustHandles="1" noChangeArrowheads="1" noChangeShapeType="1" noCrop="1"';
  function picXml({ id, name, descr, rId, x, y, cx, cy }) {
    return `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${esc(name)}" descr="${esc(descr)}"/>` +
      `<p:cNvPicPr><a:picLocks ${LOCKS}/></p:cNvPicPr><p:nvPr userDrawn="1"/></p:nvPicPr>` +
      `<p:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
      `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
  }
  const spTreeOf = (doc) => { const cSld = kid(doc.documentElement, NS.p, "cSld"); return cSld && kid(cSld, NS.p, "spTree"); };
  const isOurs = (pic) => {
    const nv = kid(pic, NS.p, "nvPicPr"), c = nv && kid(nv, NS.p, "cNvPr");
    return !!c && (c.getAttribute("name") || "").startsWith("WildTheme");
  };
  // Take our pictures out of a part; returns the image relationship ids they used.
  function removeOurs(doc) {
    const tree = spTreeOf(doc), freed = [];
    if (!tree) return freed;
    for (const pic of kids(tree, NS.p, "pic")) {
      if (!isOurs(pic)) continue;
      for (const b of all(pic, NS.a, "blip")) { const id = b.getAttributeNS(NS.r, "embed") || b.getAttribute("r:embed"); if (id) freed.push(id); }
      tree.removeChild(pic);
    }
    return freed;
  }
  function maxShapeId(doc) {
    let max = 0;
    const walk = (el) => {
      for (let n = el.firstChild; n; n = n.nextSibling) {
        if (n.nodeType !== 1) continue;
        if (n.localName === "cNvPr") max = Math.max(max, parseInt(n.getAttribute("id"), 10) || 0);
        walk(n);
      }
    };
    walk(doc.documentElement);
    return max;
  }
  // Our pictures go right after the group properties: the very back of the slide, behind everything.
  function insertAtBack(doc, xmls) {
    const tree = spTreeOf(doc);
    if (!tree) return false;
    const after = kid(tree, NS.p, "grpSpPr") || kid(tree, NS.p, "nvGrpSpPr");
    const ref = after ? after.nextSibling : tree.firstChild;
    for (const el of fragment(doc, xmls.join(""))) tree.insertBefore(el, ref);
    return true;
  }

  // ------------------------------------------------------------------ custom properties (file metadata)
  async function setCustomProps(pkg, props) {
    const path = "docProps/custom.xml";
    let doc = await pkg.xml(path);
    if (!doc) {
      doc = parseXml(`<Properties xmlns="${NS.cp}" xmlns:vt="${NS.vt}"/>`);
      pkg.setXml(path, doc);
      const rels = await pkg.rels("");
      if (!rels.byType(REL.custom).length) rels.add(REL.custom, "docProps/custom.xml");
      (await pkg.contentTypes()).ensureOverride(path, CT_CUSTOM);
    }
    const root = doc.documentElement;
    const els = () => Array.from(root.getElementsByTagNameNS(NS.cp, "property"));
    let pid = 1;
    for (const el of els()) pid = Math.max(pid, parseInt(el.getAttribute("pid"), 10) || 1);
    for (const [name, value] of Object.entries(props)) {
      let el = els().find((e) => e.getAttribute("name") === name);
      if (el) while (el.firstChild) el.removeChild(el.firstChild);
      else {
        el = doc.createElementNS(NS.cp, "property");
        el.setAttribute("fmtid", FMTID); el.setAttribute("pid", String(++pid)); el.setAttribute("name", name);
        root.appendChild(el);
      }
      const v = doc.createElementNS(NS.vt, "vt:lpwstr");
      v.appendChild(doc.createTextNode(String(value)));
      el.appendChild(v);
    }
    pkg.touch(path);
  }
  async function removeCustomProps(pkg) {
    const path = "docProps/custom.xml";
    const doc = await pkg.xml(path);
    if (!doc) return;
    const root = doc.documentElement;
    let changed = false;
    for (const el of Array.from(root.getElementsByTagNameNS(NS.cp, "property"))) if (PROPS.includes(el.getAttribute("name"))) { root.removeChild(el); changed = true; }
    if (!changed) return;
    if (root.getElementsByTagNameNS(NS.cp, "property").length) { pkg.touch(path); return; }
    pkg.remove(path);
    const rels = await pkg.rels("");
    for (const r of rels.byType(REL.custom)) rels.remove(r.id);
    (await pkg.contentTypes()).removeOverride(path);
  }
  async function readCustomProp(pkg, name) {
    const doc = await pkg.xml("docProps/custom.xml");
    if (!doc) return null;
    const el = Array.from(doc.getElementsByTagNameNS(NS.cp, "property")).find((e) => e.getAttribute("name") === name);
    return el ? el.textContent.trim() || null : null;
  }
  // Which theme does the file say it carries? File property first, then the backdrop's alt text.
  async function readMarker(pkg, s) {
    let id = await readCustomProp(pkg, "WildTheme"), name = await readCustomProp(pkg, "WildThemeName");
    if (!id || !name) {
      for (const part of [...s.masters, ...s.layouts, ...s.slides]) {
        const doc = await pkg.xml(part);
        const c = doc && all(doc, NS.p, "cNvPr").find((e) => /^Wild Theme:/.test(e.getAttribute("descr") || ""));
        if (!c) continue;
        const bits = c.getAttribute("descr").slice(11).split(";");
        id = id || bits[0].trim() || null;
        const n = bits.map((b) => /^\s*name=(.*)$/.exec(b)).find(Boolean);
        name = name || (n ? n[1].trim() : null);
        break;
      }
    }
    return { id: id || null, name: name || null };
  }

  // ------------------------------------------------------------------ cleaning up after ourselves
  const isOurMedia = (target) => /(^|\/)media\/wildtheme-/.test(target);
  async function cleanParts(pkg, parts) {
    const freedMedia = new Set();
    for (const part of parts) {
      const doc = await pkg.xml(part);
      if (!doc) continue;
      const freed = removeOurs(doc);
      if (freed.length) pkg.touch(part);
      const rels = await pkg.rels(part);
      for (const r of rels.list()) {
        if (r.type !== REL.image || r.external) continue;
        if (isOurMedia(r.target) || freed.includes(r.id)) { freedMedia.add(resolveTarget(part, r.target)); rels.remove(r.id); }
      }
    }
    if (!freedMedia.size && !Object.keys(pkg.zip.files).some((n) => /^ppt\/media\/wildtheme-/.test(n))) return 0;
    // delete our pictures that nothing points at any more
    const used = new Set();
    for (const name of Object.keys(pkg.zip.files)) {
      if (!/_rels\/[^/]+\.rels$/.test(name)) continue;
      const owner = name.replace(/_rels\/([^/]+)\.rels$/, "$1");
      const rd = pkg.cache.get(name) || parseXml(await pkg.zip.file(name).async("string"));
      for (const el of Array.from(rd.getElementsByTagName("Relationship"))) {
        if (el.getAttribute("TargetMode") !== "External") used.add(resolveTarget(owner, el.getAttribute("Target") || ""));
      }
    }
    for (const m of freedMedia) if (!used.has(m)) pkg.remove(m);
    for (const name of Object.keys(pkg.zip.files)) if (/^ppt\/media\/wildtheme-/.test(name) && !used.has(name)) pkg.remove(name);
    return freedMedia.size;
  }
  const showsMaster = (doc) => doc.documentElement.getAttribute("showMasterSp") !== "0" && doc.documentElement.getAttribute("showMasterSp") !== "false";

  // ------------------------------------------------------------------ backup of the original
  // Describes the ORIGINAL file whatever the settings, so one backup can undo any later re-stamp.
  async function takeBackup(pkg, s) {
    const b = { kind: "pptx", v: 1, themes: {}, txStyles: {}, tableStyles: null, noTransition: [], direct: {}, ctDefaults: [] };
    for (const tp of s.themes) { const x = await pkg.raw(tp); if (x != null) b.themes[tp] = x; }
    for (const m of s.masters) {
      const doc = await pkg.xml(m);
      if (!doc) continue;
      const tx = kid(doc.documentElement, NS.p, "txStyles");
      b.txStyles[m] = tx ? serializeEl(tx) : null;
    }
    const ts = await tableStylesPart(pkg, s);
    const tsDoc = await pkg.xml(ts.part);
    b.tableStyles = { part: ts.part, existed: !!tsDoc, rel: !!ts.rel, def: tsDoc && tsDoc.documentElement.hasAttribute("def") ? tsDoc.documentElement.getAttribute("def") : null };
    for (const sl of s.slides) { const doc = await pkg.xml(sl); if (doc && !hasTransition(doc)) b.noTransition.push(s.keyOf(sl)); }
    // fixed colours equal to a theme colour, by position among the part's colour elements
    const palettes = new Map();
    for (const part of [...s.layouts, ...s.slides, ...s.charts]) {
      const doc = await pkg.xml(part);
      if (!doc) continue;
      const tp = s.themeFor(part);
      if (!palettes.has(tp)) palettes.set(tp, themeColours(tp && await pkg.xml(tp)));
      const map = palettes.get(tp), recs = [];
      colourEls(doc).forEach((el, i) => {
        if (el.localName !== "srgbClr") return;
        const v = el.getAttribute("val") || "", slot = map.get(v.toUpperCase());
        if (slot) recs.push([i, v, slot]);
      });
      if (recs.length) b.direct[s.keyOf(part)] = recs;
    }
    b.ctDefaults = (await pkg.contentTypes()).defaults();
    return b;
  }

  async function restore(pkg, s, b) {
    const themesB = b.themes || {}, txB = Object.assign({}, b.txStyles || {});
    // Copies of a themed master made after stamping (LibreOffice saves one master per layout; PowerPoint can
    // duplicate masters) get the same original back as the master they were copied from.
    const schemeKey = (x) => { try { return JSON.stringify(core().schemeOf(x)); } catch (_) { return null; } };
    const fromTheme = new Map();
    for (const [tp, orig] of Object.entries(themesB)) {
      const cur = typeof orig === "string" && await pkg.raw(tp), k = cur && schemeKey(cur);
      if (k && k !== schemeKey(orig) && !fromTheme.has(k)) fromTheme.set(k, orig);
    }
    for (const tp of s.themes) {
      if (tp in themesB) continue;
      const cur = await pkg.raw(tp), orig = cur && fromTheme.get(schemeKey(cur));
      if (orig) pkg.putText(tp, orig);
    }
    const txNow = async (m) => { const d = await pkg.xml(m), t = d && kid(d.documentElement, NS.p, "txStyles"); return t ? serializeEl(t) : null; };
    const fromTx = new Map();
    for (const [m, orig] of Object.entries(txB)) { const cur = await txNow(m); if (cur != null && cur !== orig && !fromTx.has(cur)) fromTx.set(cur, orig); }
    for (const m of s.masters) {
      if (m in txB) continue;
      const cur = await txNow(m);
      if (cur != null && fromTx.has(cur)) txB[m] = fromTx.get(cur);
    }
    // theme parts: the original XML, byte for byte
    for (const [tp, xml] of Object.entries(themesB)) if (typeof xml === "string" && pkg.has(tp)) pkg.putText(tp, xml);
    // the masters' text styles
    for (const [m, xml] of Object.entries(txB)) {
      const doc = await pkg.xml(m);
      if (!doc) continue;
      const M = doc.documentElement, cur = kid(M, NS.p, "txStyles");
      if (xml == null) { if (cur) { M.removeChild(cur); pkg.touch(m); } continue; }
      const [el] = fragment(doc, xml);
      if (!el) continue;
      if (cur) M.replaceChild(el, cur); else insertBeforeAny(M, el, ["extLst"]);
      dropRedundantNs(el);
      pkg.touch(m);
    }
    // default table style
    const tb = b.tableStyles;
    if (tb && tb.part) {
      const doc = await pkg.xml(tb.part);
      if (doc && tb.existed) {
        const R = doc.documentElement, before = R.getAttribute("def");
        if (tb.def == null) R.removeAttribute("def"); else R.setAttribute("def", tb.def);
        if (before !== R.getAttribute("def")) pkg.touch(tb.part);
      } else if (doc && !tb.existed && !all(doc, NS.a, "tblStyle").length) {
        pkg.remove(tb.part);
        (await pkg.contentTypes()).removeOverride(tb.part);
        const rels = await pkg.rels(s.pres);
        for (const r of rels.byType(REL.tableStyles)) if (resolveTarget(s.pres, r.target) === tb.part) rels.remove(r.id);
      }
    }
    // transitions we added (only while they're still exactly ours)
    for (const key of b.noTransition || []) {
      const part = s.partOf(key), doc = part && await pkg.xml(part);
      if (!doc) continue;
      for (const c of elKids(doc.documentElement)) {
        const trs = transitionsIn(c);
        if (trs.length && trs.every(isOurTransition)) { doc.documentElement.removeChild(c); pkg.touch(part); }
      }
    }
    // theme colours that were fixed colours (only where the element is still the theme colour we made)
    for (const [key, recs] of Object.entries(b.direct || {})) {
      const part = s.partOf(key), doc = part && await pkg.xml(part);
      if (!doc) continue;
      const els = colourEls(doc);
      let n = 0;
      for (const [i, from, slot] of recs) {
        const el = els[i];
        if (el && el.localName === "schemeClr" && el.getAttribute("val") === slot) { swapColour(doc, el, "srgbClr", from); n++; }
      }
      if (n) pkg.touch(part);
    }
    // content types we added for our pictures
    if (Array.isArray(b.ctDefaults)) {
      const ct = await pkg.contentTypes(), names = Object.keys(pkg.zip.files);
      for (const ext of ct.defaults()) {
        if (b.ctDefaults.includes(ext) || names.some((n) => n.toLowerCase().endsWith("." + ext))) continue;
        ct.removeDefault(ext);
      }
    }
  }

  // ------------------------------------------------------------------ public API
  async function analyze(bytes) {
    const pkg = await Pkg.open(bytes);
    const s = await structure(pkg);
    const marker = await readMarker(pkg, s);
    const coreProps = await pkg.xml("docProps/core.xml");
    const t = coreProps && coreProps.getElementsByTagNameNS(NS.dc, "title")[0];
    return { kind: "pptx", sections: [s.section], slides: s.slides.length, title: t ? t.textContent.trim() : "", existingTheme: marker.id };
  }

  // settings: WTCore.settingsFor("pptx", ...) -> { backdrop, panel, styles, tables, direct, animated, transition }
  async function stamp(bytes, theme, assets, settings) {
    if (!theme || !theme.id) throw new Error("No theme chosen.");
    const C = core(), opts = C.settingsFor("pptx", settings || {}), pal = C.palette(theme);
    assets = assets || {};
    const pkg = await Pkg.open(bytes);
    const s = await structure(pkg);
    if (!s.masters.length) throw new Error("This presentation has no slide master.");
    const failures = [];
    const stats = { masters: s.masters.length, slides: s.slides.length, themes: 0, placed: 0, pictures: 0, styles: 0, tables: false, direct: 0, transitions: 0, transitionsKept: 0, animated: false };

    // 1. start clean (re-stamping swaps the theme instead of stacking it), then record the original
    await cleanParts(pkg, [...s.masters, ...s.layouts, ...s.slides]);
    const backup = await takeBackup(pkg, s);

    // 2. theme colours and fonts, on every master's theme
    const name = theme.name || theme.id;
    for (const m of s.masters) if (!s.themeOf.has(m)) failures.push(`Slide master ${m} has no theme part, so its colours and fonts can't change.`);
    for (const tp of s.themes) {
      const doc = await pkg.xml(tp);
      if (doc && applyTheme(doc, pal, name)) { pkg.touch(tp); stats.themes++; } else failures.push(`The theme part ${tp} could not be updated.`);
    }

    // 3. text styles: titles, body text and bullets on every master
    if (opts.styles) {
      for (const m of s.masters) {
        const doc = await pkg.xml(m);
        if (!doc) { failures.push(`The text styles of ${m} could not be updated.`); continue; }
        styleMaster(doc, pal); pkg.touch(m); stats.styles++;
      }
    }

    // 4. default table style (new tables, and tables without their own style)
    if (opts.tables) {
      const ts = await tableStylesPart(pkg, s);
      let doc = await pkg.xml(ts.part);
      if (!doc) {
        doc = parseXml(`<a:tblStyleLst xmlns:a="${NS.a}" def="${TABLE_STYLE}"/>`);
        pkg.setXml(ts.part, doc);
        if (!ts.rel) ts.rels.add(REL.tableStyles, relativeTarget(s.pres, ts.part));
        (await pkg.contentTypes()).ensureOverride(ts.part, CT_TABLESTYLES);
      } else if (doc.documentElement.getAttribute("def") !== TABLE_STYLE) {
        doc.documentElement.setAttribute("def", TABLE_STYLE); pkg.touch(ts.part);
      }
      stats.tables = true;
    }

    // 5. fixed copies of the old theme colours become theme colours (modifiers such as lumMod are kept)
    if (opts.direct) {
      for (const [key, recs] of Object.entries(backup.direct)) {
        const part = s.partOf(key), doc = await pkg.xml(part), els = colourEls(doc);
        for (const [i, , slot] of recs) swapColour(doc, els[i], "schemeClr", slot);
        pkg.touch(part); stats.direct += recs.length;
      }
    }

    // 6. slide transition on slides that have none (existing transitions are never replaced)
    if (opts.transition && opts.transition !== "none") {
      if (!TRANSITIONS[opts.transition]) failures.push(`Unknown slide transition "${opts.transition}".`);
      else for (const sl of s.slides) {
        const doc = await pkg.xml(sl);
        if (!doc) continue;
        if (!backup.noTransition.includes(s.keyOf(sl))) { stats.transitionsKept++; continue; }
        addTransition(doc, opts.transition); pkg.touch(sl); stats.transitions++;
      }
    }

    // 7. pictures: one backdrop (animated GIF or still) + one reading panel, shared by every master
    if (opts.backdrop) {
      if (!assets.backdrop || !assets.backdrop.bytes) failures.push("No scene picture was supplied.");
      else {
        const ext = (assets.backdrop.ext || "png").toLowerCase().replace(/^jpg$/, "jpeg");
        const ct = await pkg.contentTypes();
        ct.ensureDefault(ext, IMG_CT[ext] || "image/png");
        const backdropPath = `ppt/media/wildtheme-backdrop-${fnv(assets.backdrop.bytes)}.${ext}`;
        pkg.putBinary(backdropPath, assets.backdrop.bytes);
        stats.animated = ext === "gif";
        const rect = panelRect(s.section);
        let panelBytes = null;
        if (typeof assets.panel === "function") panelBytes = await assets.panel(s.section, rect);
        else if (assets.panel && assets.panel.bytes) panelBytes = assets.panel.bytes;
        const panelPath = panelBytes ? `ppt/media/wildtheme-panel-${fnv(panelBytes)}.png` : null;
        if (panelPath) { ct.ensureDefault("png", "image/png"); pkg.putBinary(panelPath, panelBytes); }

        const cx = s.section.pageW * TWIP_EMU, cy = s.section.pageH * TWIP_EMU;
        const descr = markerDescr(theme);
        const addTo = async (part) => {
          const doc = await pkg.xml(part);
          if (!doc || !spTreeOf(doc)) { failures.push(`The scene picture could not be placed on ${part}.`); return; }
          const rels = await pkg.rels(part);
          let id = maxShapeId(doc);
          const xmls = [picXml({ id: ++id, name: "WildTheme Backdrop", descr, rId: rels.add(REL.image, relativeTarget(part, backdropPath)), x: 0, y: 0, cx, cy })];
          if (panelPath) {
            xmls.push(picXml({ id: ++id, name: "WildTheme Panel", descr: "Reading panel (Wild Themes)", rId: rels.add(REL.image, relativeTarget(part, panelPath)),
              x: rect.x * TWIP_EMU, y: rect.y * TWIP_EMU, cx: rect.w * TWIP_EMU, cy: rect.h * TWIP_EMU }));
          }
          insertAtBack(doc, xmls);
          pkg.touch(part);
          stats.placed++; stats.pictures += xmls.length;
        };
        for (const m of s.masters) await addTo(m);
        // layouts and slides that hide the master's pictures get their own copy, so every slide is themed
        for (const p of [...s.layouts, ...s.slides]) { const doc = await pkg.xml(p); if (doc && !showsMaster(doc)) await addTo(p); }
      }
    }

    // 8. file metadata
    await setCustomProps(pkg, { WildTheme: theme.id, WildThemeName: name, WildThemeVersion: FORMAT_VERSION });

    const out = await pkg.save();
    out.backup = backup;
    out.stats = stats;
    out.failures = failures;
    return out;
  }

  // Remove everything ours. With the backup the original theme, text styles, table style, colours and
  // transitions come back exactly; without one (older files) only the pictures and metadata go.
  async function unstamp(bytes, backup) {
    const pkg = await Pkg.open(bytes);
    const s = await structure(pkg);
    await cleanParts(pkg, [...s.masters, ...s.layouts, ...s.slides]);
    await removeCustomProps(pkg);
    if (backup && typeof backup === "object" && backup.kind === "pptx") await restore(pkg, s, backup);
    return pkg.save();
  }

  // What the file carries now (for the core's fingerprints and checks). Themes: those of the masters that
  // slides use (unused masters don't count), one per distinct colour + font scheme, so an editor that splits a
  // master into identical copies (LibreOffice makes one per layout) doesn't look like a change.
  async function probe(bytes) {
    const pkg = await Pkg.open(bytes);
    const s = await structure(pkg);
    const themes = [], seen = new Set(), C = core();
    for (const tp of s.usedThemes) {
      const x = await pkg.raw(tp);
      if (x == null) continue;
      let key = x;
      try { key = JSON.stringify(C.schemeOf(x)); } catch (_) { /* unreadable: counts as its own */ }
      if (!seen.has(key)) { seen.add(key); themes.push(x); }
    }
    let pictures = 0;
    for (const part of [...s.masters, ...s.layouts, ...s.slides]) {
      const doc = await pkg.xml(part), tree = doc && spTreeOf(doc);
      if (tree) pictures += kids(tree, NS.p, "pic").filter(isOurs).length;
    }
    const media = [];
    for (const name of Object.keys(pkg.zip.files).filter((n) => /^ppt\/media\/wildtheme-backdrop/.test(n)).sort()) media.push(await pkg.zip.file(name).async("uint8array"));
    const marker = await readMarker(pkg, s);
    return { themes, pictures, media, marker: marker.id, markerName: marker.name };
  }

  return { analyze, stamp, unstamp, probe, panelRect, markerDescr };
});
