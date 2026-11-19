// Wild Themes — writes a theme INTO a PowerPoint (.pptx) file.
// The theme lives in the presentation itself: theme colours and fonts, and a backdrop plus a
// see-through reading panel on every slide master (locked, behind everything), plus file
// metadata. Anyone who opens the file sees it, with nothing installed.
//
// Browser: load vendor/jszip.min.js first, then this file -> self.WWStamperPptx.
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

  const NS = {
    p: "http://schemas.openxmlformats.org/presentationml/2006/main",
    a: "http://schemas.openxmlformats.org/drawingml/2006/main",
    r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    pr: "http://schemas.openxmlformats.org/package/2006/relationships",
    ct: "http://schemas.openxmlformats.org/package/2006/content-types",
    cp: "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
    vt: "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
    dc: "http://purl.org/dc/elements/1.1/",
  };
  const REL = {
    officeDocument: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
    slideMaster: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster",
    slideLayout: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
    slide: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
    theme: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
    image: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
    custom: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties",
    core: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
  };
  const CT_CUSTOM = "application/vnd.openxmlformats-officedocument.custom-properties+xml";
  const IMG_CT = { png: "image/png", gif: "image/gif", jpeg: "image/jpeg", jpg: "image/jpeg" };
  const TWIP_EMU = 635;
  const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  const FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";
  const PROPS = ["WildTheme", "WildThemeName", "WildThemeVersion"];

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
  const all = (el, ns, name) => Array.from(el.getElementsByTagNameNS(ns, name));
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

  // ------------------------------------------------------------------ package wrapper
  class Pkg {
    constructor(zip) { this.zip = zip; this.cache = new Map(); this.dirty = new Set(); }
    static async open(bytes) {
      let zip;
      try { zip = await JSZip.loadAsync(bytes); } catch (_) { throw new Error("This isn't a PowerPoint file (.pptx)."); }
      if (!zip.file("[Content_Types].xml")) throw new Error("This isn't a PowerPoint file (.pptx).");
      return new Pkg(zip);
    }
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
    putBinary(p, bytes) { this.zip.file(p, bytes, { binary: true, createFolders: false, compression: "STORE" }); }
    remove(p) { this.zip.remove(p); this.cache.delete(p); this.dirty.delete(p); }
    async rels(part) {
      const p = relsPathOf(part);
      let doc = await this.xml(p);
      if (!doc) { doc = parseXml(`<Relationships xmlns="${NS.pr}"/>`); this.setXml(p, doc); }
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
      };
    }
    async contentTypes() {
      const doc = await this.xml("[Content_Types].xml");
      const root = doc.documentElement, self_ = this;
      return {
        ensureDefault(ext, type) {
          if (Array.from(root.getElementsByTagName("Default")).some((d) => (d.getAttribute("Extension") || "").toLowerCase() === ext)) return;
          const el = doc.createElementNS(NS.ct, "Default");
          el.setAttribute("Extension", ext); el.setAttribute("ContentType", type);
          root.insertBefore(el, root.getElementsByTagName("Override")[0] || null); self_.touch("[Content_Types].xml");
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
    const masters = presRels.byType(REL.slideMaster).map((x) => resolveTarget(pres, x.target));
    const slides = presRels.byType(REL.slide).map((x) => resolveTarget(pres, x.target));
    const layouts = [], themes = new Set();
    for (const m of masters) {
      const mr = await pkg.rels(m);
      for (const t of mr.byType(REL.theme)) themes.add(resolveTarget(m, t.target));
      for (const l of mr.byType(REL.slideLayout)) { const lp = resolveTarget(m, l.target); if (!layouts.includes(lp)) layouts.push(lp); }
    }
    const sz = kid(presDoc.documentElement, NS.p, "sldSz");
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
    return { pres, masters, layouts, slides, themes: [...themes], section };
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
  function schemesXml(theme) {
    const d = theme.doc || {};
    const C = {
      dk1: hex(d.dk1, "1F1F1F"), dk2: hex(d.dk2, "1F3864"), lt2: hex(d.lt2, "EEEEEE"),
      accent1: hex(d.accent1 || theme.accent, "4472C4"), accent2: hex(d.accent2 || theme.accent2, "ED7D31"),
      accent3: hex(d.accent3 || theme.accent, "A5A5A5"), accent4: hex(d.accent4 || theme.accent2, "FFC000"),
      accent5: hex(d.accent5, "5B9BD5"), accent6: hex(d.accent6, "70AD47"),
    };
    const name = esc(theme.name || "Wild Theme");
    const slots = [["dk1", C.dk1], ["lt1", "FFFFFF"], ["dk2", C.dk2], ["lt2", C.lt2], ["accent1", C.accent1], ["accent2", C.accent2],
      ["accent3", C.accent3], ["accent4", C.accent4], ["accent5", C.accent5], ["accent6", C.accent6], ["hlink", C.accent2], ["folHlink", C.dk2]]
      .map(([k, v]) => `<a:${k}><a:srgbClr val="${v}"/></a:${k}>`).join("");
    const font = (f) => `<a:latin typeface="${esc(f)}"/><a:ea typeface=""/><a:cs typeface=""/>`;
    return {
      clr: `<a:clrScheme name="${name}">${slots}</a:clrScheme>`,
      font: `<a:fontScheme name="${name}"><a:majorFont>${font(d.major || "Aptos Display")}</a:majorFont><a:minorFont>${font(d.minor || "Aptos")}</a:minorFont></a:fontScheme>`,
    };
  }
  function applyTheme(doc, theme) {
    const els = all(doc, NS.a, "themeElements")[0];
    if (!els) return;
    const s = schemesXml(theme);
    for (const [name, xml] of [["clrScheme", s.clr], ["fontScheme", s.font]]) {
      const [fresh] = fragment(doc, xml);
      const old = kid(els, NS.a, name);
      if (old) els.replaceChild(fresh, old);
      else els.insertBefore(fresh, name === "clrScheme" ? els.firstChild : kid(els, NS.a, "fmtScheme"));
    }
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
    let ref = after ? after.nextSibling : tree.firstChild;
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
    for (const el of Array.from(root.getElementsByTagNameNS(NS.cp, "property"))) if (PROPS.includes(el.getAttribute("name"))) root.removeChild(el);
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
  }
  const showsMaster = (doc) => doc.documentElement.getAttribute("showMasterSp") !== "0" && doc.documentElement.getAttribute("showMasterSp") !== "false";

  // ------------------------------------------------------------------ public API
  async function analyze(bytes) {
    const pkg = await Pkg.open(bytes);
    const s = await structure(pkg);
    let existing = await readCustomProp(pkg, "WildTheme");
    if (!existing) {
      for (const m of s.masters) {
        const doc = await pkg.xml(m);
        const c = doc && all(doc, NS.p, "cNvPr").find((e) => /^Wild Theme:/.test(e.getAttribute("descr") || ""));
        if (c) { existing = c.getAttribute("descr").slice(11).split(";")[0].trim() || null; break; }
      }
    }
    const core = await pkg.xml("docProps/core.xml");
    const t = core && core.getElementsByTagNameNS(NS.dc, "title")[0];
    return { kind: "pptx", sections: [s.section], slides: s.slides.length, title: t ? t.textContent.trim() : "", existingTheme: existing };
  }

  async function stamp(bytes, theme, assets) {
    if (!theme || !theme.id) throw new Error("No theme chosen.");
    if (!assets || !assets.backdrop || !assets.backdrop.bytes) throw new Error("Missing backdrop picture.");
    const pkg = await Pkg.open(bytes);
    const s = await structure(pkg);
    if (!s.masters.length) throw new Error("This presentation has no slide master.");

    // 1. start clean (re-stamping swaps the theme instead of stacking it)
    await cleanParts(pkg, [...s.masters, ...s.layouts, ...s.slides]);

    // 2. theme colours and fonts
    for (const tp of s.themes) {
      const doc = await pkg.xml(tp);
      if (doc) { applyTheme(doc, theme); pkg.touch(tp); }
    }

    // 3. pictures: one backdrop + one reading panel, shared by every master
    const tag = Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
    const ext = (assets.backdrop.ext || "png").toLowerCase().replace("jpg", "jpeg");
    const ct = await pkg.contentTypes();
    ct.ensureDefault(ext, IMG_CT[ext] || "image/png");
    ct.ensureDefault("png", "image/png");
    const backdropPath = `ppt/media/wildtheme-backdrop-${tag}.${ext}`;
    pkg.putBinary(backdropPath, assets.backdrop.bytes);
    const rect = panelRect(s.section);
    let panelBytes = null;
    if (typeof assets.panel === "function") panelBytes = await assets.panel(s.section, rect);
    else if (assets.panel && assets.panel.bytes) panelBytes = assets.panel.bytes;
    const panelPath = panelBytes ? `ppt/media/wildtheme-panel-${tag}.png` : null;
    if (panelPath) pkg.putBinary(panelPath, panelBytes);

    const cx = s.section.pageW * TWIP_EMU, cy = s.section.pageH * TWIP_EMU;
    const descr = markerDescr(theme);
    async function addTo(part) {
      const doc = await pkg.xml(part);
      if (!doc) return 0;
      const rels = await pkg.rels(part);
      let id = maxShapeId(doc);
      const xmls = [picXml({ id: ++id, name: "WildTheme Backdrop", descr, rId: rels.add(REL.image, relativeTarget(part, backdropPath)), x: 0, y: 0, cx, cy })];
      if (panelPath) {
        xmls.push(picXml({ id: ++id, name: "WildTheme Panel", descr: "Reading panel (Wild Themes)", rId: rels.add(REL.image, relativeTarget(part, panelPath)),
          x: rect.x * TWIP_EMU, y: rect.y * TWIP_EMU, cx: rect.w * TWIP_EMU, cy: rect.h * TWIP_EMU }));
      }
      if (!insertAtBack(doc, xmls)) return 0;
      pkg.touch(part);
      return 1;
    }
    let placed = 0;
    for (const m of s.masters) placed += await addTo(m);
    // layouts and slides that hide the master's pictures get their own copy, so every slide is themed
    for (const p of [...s.layouts, ...s.slides]) { const doc = await pkg.xml(p); if (doc && !showsMaster(doc)) placed += await addTo(p); }

    // 4. file metadata
    await setCustomProps(pkg, { WildTheme: theme.id, WildThemeName: theme.name || theme.id, WildThemeVersion: "2" });

    const out = await pkg.save();
    out.stats = { masters: s.masters.length, placed, slides: s.slides.length };
    return out;
  }

  async function unstamp(bytes) {
    const pkg = await Pkg.open(bytes);
    const s = await structure(pkg);
    await cleanParts(pkg, [...s.masters, ...s.layouts, ...s.slides]);
    await removeCustomProps(pkg);
    return pkg.save();
  }

  return { analyze, stamp, unstamp, panelRect, markerDescr };
});
