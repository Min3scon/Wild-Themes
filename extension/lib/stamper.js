// Wild Themes — writes a theme INTO a .docx file.
// The theme lives in the document itself (theme colours and fonts, heading styles,
// a full-page backdrop behind the text in every page header, and metadata), so
// anyone who opens the file in Word sees it with nothing installed.
//
// Browser: load vendor/jszip.min.js first, then this file -> self.WWStamper.
// Node:    const WWStamper = require('./stamper.js')  (needs jszip + @xmldom/xmldom).
(function (root, factory) {
  if (typeof module === "object" && module.exports) {
    let xmldom = null;
    try { xmldom = require("@xmldom/xmldom"); } catch (_) { /* rely on globals */ }
    module.exports = factory(require("jszip"), xmldom);
  } else {
    root.WWStamper = factory(root.JSZip, null);
  }
})(typeof self !== "undefined" ? self : this, function (JSZip, xmldom) {
  "use strict";

  const DOMParserImpl = (typeof DOMParser !== "undefined" && DOMParser) || (xmldom && xmldom.DOMParser);
  const SerializerImpl = (typeof XMLSerializer !== "undefined" && XMLSerializer) || (xmldom && xmldom.XMLSerializer);

  // ------------------------------------------------------------------ namespaces
  const NS = {
    w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    a: "http://schemas.openxmlformats.org/drawingml/2006/main",
    pic: "http://schemas.openxmlformats.org/drawingml/2006/picture",
    mc: "http://schemas.openxmlformats.org/markup-compatibility/2006",
    w15: "http://schemas.microsoft.com/office/word/2012/wordml",
    v: "urn:schemas-microsoft-com:vml",
    o: "urn:schemas-microsoft-com:office:office",
    xmlns: "http://www.w3.org/2000/xmlns/",
    pr: "http://schemas.openxmlformats.org/package/2006/relationships",
    ct: "http://schemas.openxmlformats.org/package/2006/content-types",
    cp: "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
    vt: "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
  };
  const REL = {
    officeDocument: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
    styles: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
    theme: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
    header: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
    image: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
    settings: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings",
    custom: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties",
    core: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
  };
  const CT = {
    header: "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
    theme: "application/vnd.openxmlformats-officedocument.theme+xml",
    styles: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
    custom: "application/vnd.openxmlformats-officedocument.custom-properties+xml",
  };
  const IMG_CT = { png: "image/png", gif: "image/gif", jpeg: "image/jpeg", jpg: "image/jpeg" };
  const OUR_RSID = "57494C44";                 // "WILD": marks runs/paragraphs we add
  const HIDE_SDT = true;                       // hide the content-control box (w15:appearance)
  const TWIP_EMU = 635;
  const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  // 1x1 transparent PNG used as the body marker
  const MARKER_PNG_B64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==";
  const FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";

  // ------------------------------------------------------------------ small helpers
  const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const hex = (h, fb) => {
    h = String(h || fb || "000000").replace("#", "").toUpperCase();
    if (h.length === 3) h = h.split("").map((c) => c + c).join("");
    return /^[0-9A-F]{6}$/.test(h) ? h : String(fb || "000000").replace("#", "").toUpperCase();
  };
  const b64ToBytes = (b64) => {
    if (typeof Buffer !== "undefined") return new Uint8Array(Buffer.from(b64, "base64"));
    const s = atob(b64); const u = new Uint8Array(s.length);
    for (let i = 0; i < s.length; i++) u[i] = s.charCodeAt(i);
    return u;
  };
  function parseXml(str) {
    const doc = new DOMParserImpl().parseFromString(str, "application/xml");
    const err = doc.getElementsByTagName("parsererror");
    if (err && err.length) throw new Error("Could not read the Word file (XML error).");
    return doc;
  }
  function serialize(doc) {
    let s = new SerializerImpl().serializeToString(doc);
    s = s.replace(/^<\?xml[^>]*\?>\s*/, "");
    return XML_DECL + s;
  }
  // parse a fragment written with prefixes w:, wp:, a:, pic:, r:
  function fragment(doc, xml) {
    const wrap = `<root xmlns:w="${NS.w}" xmlns:r="${NS.r}" xmlns:wp="${NS.wp}" xmlns:a="${NS.a}" xmlns:pic="${NS.pic}" xmlns:w15="${NS.w15}" xmlns:v="${NS.v}" xmlns:o="${NS.o}">${xml}</root>`;
    const f = parseXml(wrap).documentElement;
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
  const wAttr = (el, name) => (el ? el.getAttributeNS(NS.w, name) || el.getAttribute("w:" + name) || "" : "");
  const onOff = (el) => {
    if (!el) return false;
    const v = wAttr(el, "val");
    return !(v === "0" || v === "false" || v === "off");
  };
  function ancestor(el, ns, name) {
    for (let n = el.parentNode; n && n.nodeType === 1; n = n.parentNode) if (n.namespaceURI === ns && n.localName === name) return n;
    return null;
  }
  const dirOf = (p) => (p.includes("/") ? p.slice(0, p.lastIndexOf("/")) : "");
  const relsPathOf = (p) => (dirOf(p) ? dirOf(p) + "/" : "") + "_rels/" + p.split("/").pop() + ".rels";
  function resolveTarget(fromPart, target) {
    if (target.startsWith("/")) return target.slice(1);
    const parts = (dirOf(fromPart) ? dirOf(fromPart).split("/") : []);
    for (const seg of target.split("/")) {
      if (seg === "..") parts.pop(); else if (seg && seg !== ".") parts.push(seg);
    }
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
      const zip = await JSZip.loadAsync(bytes);
      if (!zip.file("[Content_Types].xml")) throw new Error("This isn't a Word document (.docx).");
      return new Pkg(zip);
    }
    has(p) { return !!this.zip.file(p) || this.cache.has(p); }
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
      const root = doc.documentElement;
      const list = () => Array.from(root.getElementsByTagName("Relationship")).map((el) => ({
        el, id: el.getAttribute("Id"), type: el.getAttribute("Type"), target: el.getAttribute("Target"),
        external: el.getAttribute("TargetMode") === "External",
      }));
      const self_ = this;
      return {
        path: p, doc, list,
        add(type, target) {
          const ids = new Set(list().map((r) => r.id));
          let n = 1; while (ids.has("rIdWT" + n)) n++;
          const el = doc.createElementNS(NS.pr, "Relationship");
          el.setAttribute("Id", "rIdWT" + n); el.setAttribute("Type", type); el.setAttribute("Target", target);
          root.appendChild(el); self_.touch(p);
          return "rIdWT" + n;
        },
        remove(id) { for (const r of list()) if (r.id === id) { root.removeChild(r.el); self_.touch(p); } },
        byType(type) { return list().filter((r) => r.type === type); },
        byId(id) { return list().find((r) => r.id === id) || null; },
      };
    }
    async contentTypes() {
      const doc = await this.xml("[Content_Types].xml");
      const root = doc.documentElement;
      const self_ = this;
      return {
        ensureDefault(ext, type) {
          const ex = Array.from(root.getElementsByTagName("Default")).find((d) => (d.getAttribute("Extension") || "").toLowerCase() === ext);
          if (ex) return;
          const el = doc.createElementNS(NS.ct, "Default");
          el.setAttribute("Extension", ext); el.setAttribute("ContentType", type);
          const firstOverride = root.getElementsByTagName("Override")[0];
          root.insertBefore(el, firstOverride || null); self_.touch("[Content_Types].xml");
        },
        ensureOverride(part, type) {
          const name = "/" + part;
          const ex = Array.from(root.getElementsByTagName("Override")).find((o) => o.getAttribute("PartName") === name);
          if (ex) { if (ex.getAttribute("ContentType") !== type) { ex.setAttribute("ContentType", type); self_.touch("[Content_Types].xml"); } return; }
          const el = doc.createElementNS(NS.ct, "Override");
          el.setAttribute("PartName", name); el.setAttribute("ContentType", type);
          root.appendChild(el); self_.touch("[Content_Types].xml");
        },
        removeOverride(part) {
          for (const o of Array.from(root.getElementsByTagName("Override"))) if (o.getAttribute("PartName") === "/" + part) { root.removeChild(o); self_.touch("[Content_Types].xml"); }
        },
      };
    }
    async mainPart() {
      const rels = await this.rels("");
      const r = rels.byType(REL.officeDocument)[0];
      if (!r) throw new Error("This Word file has no main document.");
      return resolveTarget("", r.target);
    }
    async save() {
      for (const p of this.dirty) if (this.cache.has(p)) this.zip.file(p, serialize(this.cache.get(p)), { createFolders: false });
      // drop folder entries (zip.remove() on a folder would delete its contents too)
      for (const name of Object.keys(this.zip.files)) { const f = this.zip.files[name]; if (f && f.dir) delete this.zip.files[name]; }
      return this.zip.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 6 } });
    }
  }

  // ------------------------------------------------------------------ geometry
  function sectionGeometry(sectPr) {
    const pgSz = sectPr && kid(sectPr, NS.w, "pgSz");
    const pgMar = sectPr && kid(sectPr, NS.w, "pgMar");
    const num = (el, name, fb) => { const v = parseInt(wAttr(el, name), 10); return Number.isFinite(v) ? v : fb; };
    let pageW = num(pgSz, "w", 12240), pageH = num(pgSz, "h", 15840);
    const orient = wAttr(pgSz, "orient");
    if (orient === "landscape" && pageW < pageH) [pageW, pageH] = [pageH, pageW];
    return {
      pageW, pageH,
      margins: {
        top: Math.abs(num(pgMar, "top", 1440)), right: num(pgMar, "right", 1440),
        bottom: Math.abs(num(pgMar, "bottom", 1440)), left: num(pgMar, "left", 1440),
        header: num(pgMar, "header", 720), footer: num(pgMar, "footer", 720),
      },
    };
  }
  // Text area grown by ~0.3in each side, kept at least ~0.3in inside the page edges.
  function panelRect(section) {
    const { pageW, pageH, margins: m } = section;
    const side = (margin) => {
      const edge = Math.min(432, margin * 0.6);
      return Math.max(0, Math.min(432, margin - edge));
    };
    const x = Math.max(0, m.left - side(m.left));
    let y = Math.max(0, m.top - side(m.top));
    const r = Math.min(pageW, pageW - m.right + side(m.right));
    let b = Math.min(pageH, pageH - m.bottom + side(m.bottom));
    // keep header / footer text on the panel too, so it never sits on the backdrop
    if (section.headerText) y = Math.max(Math.min(216, m.header), Math.min(y, m.header - 288));
    if (section.footerText) b = Math.min(pageH - Math.min(216, m.footer), Math.max(b, pageH - m.footer + 288));
    return { x: Math.round(x), y: Math.round(y), w: Math.round(Math.max(1, r - x)), h: Math.round(Math.max(1, b - y)) };
  }
  const sameGeom = (a, b) => a.pageW === b.pageW && a.pageH === b.pageH &&
    JSON.stringify(panelRect(a)) === JSON.stringify(panelRect(b));

  // Does any header/footer this section shows contain text or a field (page numbers)?
  const partHasText = (d) => !!d && (all(d, NS.w, "t").some((t) => t.textContent.trim()) || all(d, NS.w, "instrText").length > 0 || all(d, NS.w, "fldSimple").length > 0);
  async function sectionInfos(pkg, main, doc, mainRels, evenOdd) {
    const out = [];
    const inherit = { header: {}, footer: {} };
    for (const sectPr of sections(doc)) {
      const g = sectionGeometry(sectPr);
      const types = ["default"];
      if (onOff(kid(sectPr, NS.w, "titlePg"))) types.push("first");
      if (evenOdd) types.push("even");
      for (const kind of ["header", "footer"]) {
        for (const ref of kids(sectPr, NS.w, kind + "Reference")) {
          const rel = mainRels.byId(ref.getAttributeNS(NS.r, "id") || ref.getAttribute("r:id"));
          if (rel) inherit[kind][wAttr(ref, "type") || "default"] = resolveTarget(main, rel.target);
        }
        let has = false;
        for (const t of types) if (inherit[kind][t]) has = has || partHasText(await pkg.xml(inherit[kind][t]));
        g[kind + "Text"] = has;
      }
      out.push(g);
    }
    if (!out.length) out.push(sectionGeometry(null));
    return out;
  }
  async function evenOddOf(pkg, main, mainRels) {
    const settingsRel = mainRels.byType(REL.settings)[0];
    if (!settingsRel) return false;
    const sd = await pkg.xml(resolveTarget(main, settingsRel.target));
    return !!sd && onOff(all(sd, NS.w, "evenAndOddHeaders")[0] || null);
  }

  // All real section properties in body order (skip the old copies inside w:sectPrChange).
  function sections(doc) {
    return all(doc, NS.w, "sectPr").filter((s) => !ancestor(s, NS.w, "sectPrChange"));
  }

  // ------------------------------------------------------------------ theme part
  function themeXml(theme) {
    const d = theme.doc || {};
    const C = {
      dk1: hex(d.dk1, "1F1F1F"), dk2: hex(d.dk2, "1F3864"), lt2: hex(d.lt2, "EEEEEE"),
      accent1: hex(d.accent1 || theme.accent, "4472C4"), accent2: hex(d.accent2 || theme.accent2, "ED7D31"),
      accent3: hex(d.accent3 || theme.accent, "A5A5A5"), accent4: hex(d.accent4 || theme.accent2, "FFC000"),
      accent5: hex(d.accent5, "5B9BD5"), accent6: hex(d.accent6, "70AD47"),
    };
    const name = esc(theme.name || "Wild Theme");
    const major = esc(d.major || "Aptos Display"), minor = esc(d.minor || "Aptos");
    const g = (stops) => `<a:gradFill rotWithShape="1"><a:gsLst>${stops.map(([p, m]) => `<a:gs pos="${p}"><a:schemeClr val="phClr">${m}</a:schemeClr></a:gs>`).join("")}</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>`;
    const ln = (w) => `<a:ln w="${w}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>`;
    const slots = [["dk1", C.dk1], ["lt1", "FFFFFF"], ["dk2", C.dk2], ["lt2", C.lt2], ["accent1", C.accent1], ["accent2", C.accent2],
      ["accent3", C.accent3], ["accent4", C.accent4], ["accent5", C.accent5], ["accent6", C.accent6], ["hlink", C.accent2], ["folHlink", C.dk2]]
      .map(([k, v]) => `<a:${k}><a:srgbClr val="${v}"/></a:${k}>`).join("");
    return XML_DECL + `<a:theme xmlns:a="${NS.a}" name="Wild Theme ${name}"><a:themeElements>` +
      `<a:clrScheme name="${name}">${slots}</a:clrScheme>` +
      `<a:fontScheme name="${name}"><a:majorFont><a:latin typeface="${major}"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>` +
      `<a:minorFont><a:latin typeface="${minor}"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>` +
      `<a:fmtScheme name="${name}"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>` +
      g([[0, '<a:tint val="67000"/>'], [100000, '<a:tint val="81000"/>']]) + g([[0, '<a:shade val="94000"/>'], [100000, '<a:shade val="78000"/>']]) +
      `</a:fillStyleLst><a:lnStyleLst>${ln(6350)}${ln(12700)}${ln(19050)}</a:lnStyleLst>` +
      `<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>` +
      `<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/></a:schemeClr></a:solidFill>` +
      g([[0, '<a:tint val="93000"/>'], [100000, '<a:shade val="63000"/>']]) +
      `</a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`;
  }

  // ------------------------------------------------------------------ styles
  const MAJOR = '<w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:cstheme="majorBidi"/>';
  const MINOR = '<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:cstheme="minorBidi"/>';

  function styleDefs(theme) {
    const d = theme.doc || {};
    const C = {
      dk2: hex(d.dk2, "1F3864"), lt2: hex(d.lt2, "EEEEEE"),
      accent1: hex(d.accent1 || theme.accent, "4472C4"), accent2: hex(d.accent2 || theme.accent2, "ED7D31"),
      accent3: hex(d.accent3 || theme.accent, "A5A5A5"),
    };
    const col = (slot, v) => `<w:color w:val="${v}" w:themeColor="${slot}"/>`;
    const bdr = (side, slot, v, sz, sp) => `<w:${side} w:val="single" w:sz="${sz}" w:space="${sp}" w:color="${v}" w:themeColor="${slot}"/>`;
    const shd = `<w:shd w:val="clear" w:color="auto" w:fill="${C.lt2}" w:themeFill="background2"/>`;
    // name (lowercase w:name) -> default styleId, priority, pPr, rPr
    return [
      ["title", "Title", 10, '<w:spacing w:after="40" w:line="240" w:lineRule="auto"/><w:contextualSpacing/>',
        MAJOR + col("text2", C.dk2) + '<w:kern w:val="28"/><w:sz w:val="56"/><w:szCs w:val="56"/>'],
      ["subtitle", "Subtitle", 11, '<w:spacing w:after="320"/>',
        MINOR + col("accent1", C.accent1) + '<w:sz w:val="26"/><w:szCs w:val="26"/>'],
      ["heading 1", "Heading1", 9, '<w:keepNext/><w:keepLines/><w:pBdr>' + bdr("bottom", "accent3", C.accent3, 12, 4) +
        '</w:pBdr><w:spacing w:before="360" w:after="140"/><w:outlineLvl w:val="0"/>',
        MAJOR + col("accent1", C.accent1) + '<w:sz w:val="32"/><w:szCs w:val="32"/>'],
      ["heading 2", "Heading2", 9, '<w:keepNext/><w:keepLines/><w:spacing w:before="260" w:after="80"/><w:outlineLvl w:val="1"/>',
        MAJOR + col("accent2", C.accent2) + '<w:sz w:val="26"/><w:szCs w:val="26"/>'],
      ["heading 3", "Heading3", 9, '<w:keepNext/><w:keepLines/><w:spacing w:before="200" w:after="60"/><w:outlineLvl w:val="2"/>',
        MINOR + '<w:b/><w:bCs/>' + col("text2", C.dk2) + '<w:sz w:val="24"/><w:szCs w:val="24"/>'],
      ["quote", "Quote", 29, '<w:pBdr>' + bdr("left", "accent3", C.accent3, 36, 12) + '</w:pBdr>' + shd +
        '<w:spacing w:before="200" w:after="200"/><w:ind w:left="300" w:right="300"/>',
        MAJOR + col("text2", C.dk2) + '<w:sz w:val="26"/><w:szCs w:val="26"/>'],
      ["intense quote", "IntenseQuote", 30, '<w:pBdr>' + bdr("left", "accent1", C.accent1, 48, 12) + '</w:pBdr>' + shd +
        '<w:spacing w:before="240" w:after="240"/><w:ind w:left="300" w:right="300"/>',
        MINOR + '<w:b/><w:bCs/>' + col("accent1", C.accent1) + '<w:sz w:val="24"/><w:szCs w:val="24"/>'],
    ];
  }

  function applyStyles(doc, theme) {
    const root = doc.documentElement;
    const styleEls = kids(root, NS.w, "style");
    const nameOf = (s) => wAttr(kid(s, NS.w, "name"), "val").toLowerCase();
    const idOf = (s) => wAttr(s, "styleId");
    const normal = styleEls.find((s) => wAttr(s, "type") === "paragraph" && ["1", "true", "on"].includes(wAttr(s, "default")))
      || styleEls.find((s) => wAttr(s, "type") === "paragraph" && nameOf(s) === "normal");
    const normalId = normal ? idOf(normal) : "Normal";

    // 1. document default run font -> theme body font (direct formatting is untouched)
    let docDefaults = kid(root, NS.w, "docDefaults");
    if (!docDefaults) {
      docDefaults = fragment(doc, "<w:docDefaults/>")[0];
      root.insertBefore(docDefaults, root.firstChild);
    }
    let rPrDefault = kid(docDefaults, NS.w, "rPrDefault");
    if (!rPrDefault) { rPrDefault = fragment(doc, "<w:rPrDefault/>")[0]; docDefaults.insertBefore(rPrDefault, docDefaults.firstChild); }
    let rPr = kid(rPrDefault, NS.w, "rPr");
    if (!rPr) { rPr = fragment(doc, "<w:rPr/>")[0]; rPrDefault.appendChild(rPr); }
    setRunFonts(doc, rPr, MINOR);
    // Normal style: swap any fixed font for the theme body font
    if (normal) {
      const nr = kid(normal, NS.w, "rPr");
      if (nr && kid(nr, NS.w, "rFonts")) setRunFonts(doc, nr, MINOR);
    }

    // 2. headings, title, quotes
    for (const [lname, defId, prio, pPr, rPrXml] of styleDefs(theme)) {
      const existing = styleEls.find((s) => wAttr(s, "type") === "paragraph" && nameOf(s) === lname);
      const id = existing ? idOf(existing) : defId;
      const link = existing ? wAttr(kid(existing, NS.w, "link"), "val") : "";
      const displayName = { title: "Title", subtitle: "Subtitle", quote: "Quote", "intense quote": "Intense Quote" }[lname] || lname;
      const xml = `<w:style w:type="paragraph" w:styleId="${esc(id)}"><w:name w:val="${esc(displayName)}"/>` +
        `<w:basedOn w:val="${esc(normalId)}"/><w:next w:val="${esc(normalId)}"/>` + (link ? `<w:link w:val="${esc(link)}"/>` : "") +
        `<w:uiPriority w:val="${prio}"/>` + (lname.startsWith("heading") ? "<w:unhideWhenUsed/>" : "") +
        `<w:qFormat/><w:pPr>${pPr}</w:pPr><w:rPr>${rPrXml}</w:rPr></w:style>`;
      const el = fragment(doc, xml)[0];
      if (existing) root.replaceChild(el, existing); else root.appendChild(el);
      // keep the matching character style ("Heading 1 Char") in step
      if (link) {
        const ch = kids(root, NS.w, "style").find((s) => wAttr(s, "type") === "character" && idOf(s) === link);
        if (ch) {
          const old = kid(ch, NS.w, "rPr");
          const nr = fragment(doc, `<w:rPr>${rPrXml}</w:rPr>`)[0];
          if (old) ch.replaceChild(nr, old); else ch.appendChild(nr);
        }
      }
    }
  }
  // Replace a run-properties element's w:rFonts (keeping schema order: rStyle first, then rFonts).
  function setRunFonts(doc, rPr, fontsXml) {
    const el = fragment(doc, fontsXml)[0];
    const old = kid(rPr, NS.w, "rFonts");
    if (old) { rPr.replaceChild(el, old); return; }
    const rStyle = kid(rPr, NS.w, "rStyle");
    rPr.insertBefore(el, rStyle ? rStyle.nextSibling : rPr.firstChild);
  }

  // ------------------------------------------------------------------ drawings
  function anchorXml({ id, name, descr, title, rId, x, y, cx, cy, z }) {
    return `<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${z}" behindDoc="1" locked="1" layoutInCell="0" allowOverlap="1">` +
      `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>${x}</wp:posOffset></wp:positionH>` +
      `<wp:positionV relativeFrom="page"><wp:posOffset>${y}</wp:posOffset></wp:positionV>` +
      `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>` +
      `<wp:docPr id="${id}" name="${esc(name)}" descr="${esc(descr)}"${title ? ` title="${esc(title)}"` : ""}/>` +
      `<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>` +
      `<pic:nvPicPr><pic:cNvPr id="${id}" name="${esc(name)}"/><pic:cNvPicPr/></pic:nvPicPr>` +
      `<pic:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
      `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
      `</pic:pic></a:graphicData></a:graphic></wp:anchor>`;
  }
  const markerDescr = (theme) =>
    `Wild Theme: ${theme.id}; name=${theme.name}; accent=#${hex(theme.accent)}; accent2=#${hex(theme.accent2)}; mode=${theme.mode || "dark"}`;

  // Is this drawing one of ours (current or v1 builder)?
  function isOurDocPr(docPr) {
    const name = docPr.getAttribute("name") || "";
    const descr = docPr.getAttribute("descr") || "";
    return name.startsWith("WildTheme") || (name === "Theme backdrop" && /^Decorative animated/.test(descr));
  }
  const isOurTag = (v) => /^WildTheme/.test(v || "");
  // A paragraph we created ourselves that has nothing left in it is dropped (unless it is the only block).
  function dropEmptyOwnParagraph(p) {
    if (!p || p.namespaceURI !== NS.w || p.localName !== "p" || wAttr(p, "rsidR") !== OUR_RSID) return;
    if (Array.from(p.childNodes).some((c) => c.nodeType === 1 && !(c.namespaceURI === NS.w && c.localName === "pPr"))) return;
    const parent = p.parentNode;
    const blocks = Array.from(parent.childNodes).filter((c) => c.nodeType === 1 && (c.localName === "p" || c.localName === "tbl" || c.localName === "sdt"));
    if (blocks.length > 1) parent.removeChild(p);
  }
  const blipIds = (el) => all(el, NS.a, "blip").map((b) => b.getAttributeNS(NS.r, "embed") || b.getAttribute("r:embed")).filter(Boolean);
  // Remove everything Wild Theme put in a part: locked content controls first, then any loose
  // drawings from older versions. Returns the image relationship ids that were freed.
  function removeOurs(doc) {
    const freed = [];
    for (const sdt of all(doc, NS.w, "sdt")) {
      const pr = kid(sdt, NS.w, "sdtPr");
      if (!pr || !isOurTag(wAttr(kid(pr, NS.w, "tag"), "val"))) continue;
      if (!sdt.parentNode) continue;
      freed.push(...blipIds(sdt));
      const p = sdt.parentNode;
      p.removeChild(sdt);
      dropEmptyOwnParagraph(p);
    }
    for (const docPr of all(doc, NS.wp, "docPr")) {
      if (!isOurDocPr(docPr)) continue;
      const drawing = ancestor(docPr, NS.w, "drawing");
      if (!drawing || !drawing.parentNode) continue;
      freed.push(...blipIds(drawing));
      const alt = ancestor(drawing, NS.mc, "AlternateContent");
      const target = alt || drawing;
      const run = ancestor(target, NS.w, "r");
      target.parentNode.removeChild(target);
      if (run && !Array.from(run.childNodes).some((c) => c.nodeType === 1 && !(c.namespaceURI === NS.w && c.localName === "rPr"))) {
        const p = run.parentNode;
        p.removeChild(run);
        dropEmptyOwnParagraph(p);
      }
    }
    return freed;
  }
  const removeOurDrawings = removeOurs;
  // Declare w15 (and list it as ignorable) on a part's root so we can hide a content control's box.
  function ensureW15(doc) {
    const root = doc.documentElement;
    if (!root.getAttribute("xmlns:w15")) root.setAttributeNS(NS.xmlns, "xmlns:w15", NS.w15);
    if (!root.getAttribute("xmlns:mc")) root.setAttributeNS(NS.xmlns, "xmlns:mc", NS.mc);
    const ign = root.getAttributeNS(NS.mc, "Ignorable") || root.getAttribute("mc:Ignorable") || "";
    if (!ign.split(/\s+/).includes("w15")) root.setAttributeNS(NS.mc, "mc:Ignorable", (ign + " w15").trim());
  }
  // A locked inline content control around our run: it can't be deleted or edited in Word.
  function lockedSdt(tag, id, innerXml, hidden) {
    return `<w:sdt><w:sdtPr><w:alias w:val="Wild Theme"/><w:tag w:val="${tag}"/><w:id w:val="${id}"/><w:lock w:val="sdtContentLocked"/>` +
      (hidden ? `<w15:appearance w15:val="hidden"/>` : "") + `</w:sdtPr><w:sdtContent>${innerXml}</w:sdtContent></w:sdt>`;
  }
  function sdtIdPicker(docs) {
    const used = new Set();
    for (const d of docs) for (const pr of all(d, NS.w, "sdtPr")) { const v = wAttr(kid(pr, NS.w, "id"), "val"); if (v) used.add(v); }
    return () => {
      let v;
      do { v = String(Math.floor(Math.random() * 2000000000) - 1000000000); } while (used.has(v));
      used.add(v);
      return v;
    };
  }
  // Old visible tag lines from v1 documents ("✦ Wild Theme: …") are removed.
  function removeTagLines(doc) {
    const body = all(doc, NS.w, "body")[0];
    if (!body) return;
    for (const p of kids(body, NS.w, "p")) {
      const text = all(p, NS.w, "t").map((t) => t.textContent).join("");
      if (/^\s*(✦\s*)?Wild Theme:\s*.{0,60}$/.test(text) && !all(p, NS.w, "drawing").length) body.removeChild(p);
    }
  }
  function maxDocPrId(docs) {
    let max = 0;
    for (const d of docs) for (const tag of ["docPr"]) for (const el of all(d, NS.wp, tag)) max = Math.max(max, parseInt(el.getAttribute("id"), 10) || 0);
    return max;
  }
  // Put a run at the start of a part's first paragraph (or a new tiny paragraph if the part starts with a table).
  function insertLeadRun(doc, container, runXml) {
    const blocks = Array.from(container.childNodes).filter((c) => c.nodeType === 1);
    const first = blocks.find((c) => c.namespaceURI === NS.w && (c.localName === "p" || c.localName === "tbl" || c.localName === "sdt"));
    let p = first && first.localName === "p" ? first : null;
    if (!p) {
      p = fragment(doc, `<w:p w:rsidR="${OUR_RSID}"><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>`)[0];
      container.insertBefore(p, first || null);
    }
    const run = fragment(doc, runXml)[0];
    const pPr = kid(p, NS.w, "pPr");
    p.insertBefore(run, pPr ? pPr.nextSibling : p.firstChild);
    return p;
  }


  // ------------------------------------------------------------------ settings
  // CT_Settings is an ordered sequence; this is its order up to defaultTabStop.
  const SETTINGS_ORDER = ["writeProtection", "view", "zoom", "removePersonalInformation", "removeDateAndTime",
    "doNotDisplayPageBoundaries", "displayBackgroundShape", "printPostScriptOverText", "printFractionalCharacterWidth",
    "printFormsData", "embedTrueTypeFonts", "embedSystemFonts", "saveSubsetFonts", "saveFormsData", "mirrorMargins",
    "alignBordersAndEdges", "bordersDoNotSurroundHeader", "bordersDoNotSurroundFooter", "gutterAtTop", "hideSpellingErrors",
    "hideGrammaticalErrors", "activeWritingStyle", "proofState", "formsDesign", "attachedTemplate", "linkStyles",
    "stylePaneFormatFilter", "stylePaneSortMethod", "documentType", "mailMerge", "revisionView", "trackRevisions",
    "doNotTrackMoves", "doNotTrackFormatting", "documentProtection", "autoFormatOverride", "styleLockTheme", "styleLockQFSet"];
  // Add an empty on/off setting (<w:name/>) in its schema position, or switch an existing one on.
  function ensureSetting(doc, name) {
    const root = doc.documentElement;
    const ex = kid(root, NS.w, name);
    if (ex) { ex.removeAttributeNS(NS.w, "val"); ex.removeAttribute("w:val"); return; }
    const earlier = new Set(SETTINGS_ORDER.slice(0, SETTINGS_ORDER.indexOf(name)));
    let before = null;
    for (let n = root.firstChild; n; n = n.nextSibling) {
      if (n.nodeType !== 1) continue;
      if (n.namespaceURI === NS.w && earlier.has(n.localName)) continue;
      before = n; break;
    }
    root.insertBefore(fragment(doc, `<w:${name}/>`)[0], before);
  }
  function removeSetting(doc, name) {
    let n = 0;
    for (const el of kids(doc.documentElement, NS.w, name)) { doc.documentElement.removeChild(el); n++; }
    return n;
  }
  // Theme + style-set locks, and show the page background picture.
  function lockTheme(doc) { for (const name of ["styleLockTheme", "styleLockQFSet"]) ensureSetting(doc, name); }

  // ------------------------------------------------------------------ page background
  // Declare a namespace prefix on a part's root.
  function ensureNs(doc, prefix, uri) {
    const root = doc.documentElement;
    if (!root.getAttribute("xmlns:" + prefix)) root.setAttributeNS(NS.xmlns, "xmlns:" + prefix, uri);
  }
  // Relationship ids used anywhere inside an element (r:id, r:embed, ...).
  function relIdsIn(el) {
    const ids = [];
    for (const n of [el, ...Array.from(el.getElementsByTagNameNS("*", "*"))]) {
      for (const a of Array.from(n.attributes || [])) if (a.namespaceURI === NS.r && a.value) ids.push(a.value);
    }
    return ids;
  }
  const backgroundOf = (doc) => kid(doc.documentElement, NS.w, "background");
  // Ours if its picture fill is titled "Wild Theme" or points at one of our media files.
  function isOurBackground(bg, rels) {
    if (!bg) return false;
    for (const f of all(bg, NS.v, "fill")) if ((f.getAttributeNS(NS.o, "title") || f.getAttribute("o:title")) === "Wild Theme") return true;
    return relIdsIn(bg).some((id) => { const r = rels.byId(id); return !!r && isOurMedia(r.target); });
  }
  function pageBackgroundXml(rId) {
    return `<w:background w:color="FFFFFF"><v:background id="_x0000_s1025" o:bwmode="white" o:targetscreensize="1024,768">` +
      `<v:fill r:id="${rId}" o:title="Wild Theme" recolor="t" type="frame"/></v:background></w:background>`;
  }

  // ------------------------------------------------------------------ custom properties
  async function setCustomProps(pkg, props) {
    const path = "docProps/custom.xml";
    let doc = await pkg.xml(path);
    if (!doc) {
      doc = parseXml(`<Properties xmlns="${NS.cp}" xmlns:vt="${NS.vt}"/>`);
      pkg.setXml(path, doc);
      const rels = await pkg.rels("");
      if (!rels.byType(REL.custom).length) rels.add(REL.custom, "docProps/custom.xml");
      (await pkg.contentTypes()).ensureOverride(path, CT.custom);
    }
    const root = doc.documentElement;
    const propEls = () => Array.from(root.getElementsByTagNameNS(NS.cp, "property"));
    let pid = 1;
    for (const el of propEls()) pid = Math.max(pid, parseInt(el.getAttribute("pid"), 10) || 1);
    for (const [name, value] of Object.entries(props)) {
      let el = propEls().find((e) => e.getAttribute("name") === name);
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
  async function readCustomProp(pkg, name) {
    const doc = await pkg.xml("docProps/custom.xml");
    if (!doc) return null;
    const el = Array.from(doc.getElementsByTagNameNS(NS.cp, "property")).find((e) => e.getAttribute("name") === name);
    return el ? el.textContent.trim() || null : null;
  }

  // Delete pictures we freed (and any of our own media) that no relationship points at any more.
  async function pruneMedia(pkg, freedMedia) {
    const stillUsed = new Set();
    for (const name of Object.keys(pkg.zip.files)) {
      if (!/_rels\/[^/]+\.rels$/.test(name)) continue;
      const owner = name.replace(/_rels\/([^/]+)\.rels$/, "$1");
      const rd = pkg.cache.get(name) || parseXml(await pkg.zip.file(name).async("string"));
      for (const el of Array.from(rd.getElementsByTagName("Relationship"))) {
        if (el.getAttribute("TargetMode") !== "External") stillUsed.add(resolveTarget(owner, el.getAttribute("Target") || ""));
      }
    }
    for (const m of freedMedia) if (!stillUsed.has(m)) pkg.remove(m);
    for (const name of Object.keys(pkg.zip.files)) if (/^word\/media\/wildtheme-/.test(name) && !stillUsed.has(name)) pkg.remove(name);
  }
  const isOurMedia = (target) => /(^|\/)media\/wildtheme-/.test(target);
  // Remove our pictures (and their relationships) from the main part and every header part,
  // plus our page background. { anyBackground: true } also removes someone else's page background
  // (stamping replaces it). Returns whether a page background was removed.
  async function cleanParts(pkg, main, mainRels, { anyBackground = false } = {}) {
    const freedMedia = new Set();
    let removedBackground = false;
    for (const part of [main, ...mainRels.byType(REL.header).map((r) => resolveTarget(main, r.target))]) {
      const pd = await pkg.xml(part);
      if (!pd) continue;
      const freed = removeOurs(pd);
      if (part === main) {
        removeTagLines(pd);
        const bg = backgroundOf(pd);
        if (bg && (anyBackground || isOurBackground(bg, mainRels))) {
          freed.push(...relIdsIn(bg));
          pd.documentElement.removeChild(bg);
          removedBackground = true;
        }
      }
      pkg.touch(part);
      const rels = await pkg.rels(part);
      for (const r of rels.list()) {
        if (r.type !== REL.image) continue;
        if (isOurMedia(r.target) || freed.includes(r.id)) { freedMedia.add(resolveTarget(part, r.target)); rels.remove(r.id); }
      }
    }
    await pruneMedia(pkg, freedMedia);
    return removedBackground;
  }

  // ------------------------------------------------------------------ public: analyze
  async function analyze(bytes) {
    const pkg = await Pkg.open(bytes);
    const main = await pkg.mainPart();
    const doc = await pkg.xml(main);
    const mainRels = await pkg.rels(main);
    const secs = await sectionInfos(pkg, main, doc, mainRels, await evenOddOf(pkg, main, mainRels));
    let title = null;
    const core = await pkg.xml("docProps/core.xml");
    if (core) {
      const t = core.getElementsByTagNameNS("http://purl.org/dc/elements/1.1/", "title")[0];
      title = t ? t.textContent.trim() || null : null;
    }
    let existingTheme = await readCustomProp(pkg, "WildTheme");
    if (!existingTheme) {
      for (const docPr of all(doc, NS.wp, "docPr")) {
        const m = /^Wild Theme:\s*([^;]+)/.exec(docPr.getAttribute("descr") || "");
        if (m) { existingTheme = m[1].trim(); break; }
      }
    }
    return { sections: secs, title, existingTheme: existingTheme || null };
  }

  // ------------------------------------------------------------------ public: stamp
  async function stamp(bytes, theme, assets) {
    if (!theme || !theme.id) throw new Error("No theme chosen.");
    if (!assets || !assets.backdrop || !assets.backdrop.bytes) throw new Error("Missing backdrop image.");
    const pkg = await Pkg.open(bytes);
    const ct = await pkg.contentTypes();
    const main = await pkg.mainPart();
    const doc = await pkg.xml(main);
    const mainRels = await pkg.rels(main);
    const descr = markerDescr(theme);

    // ---- clean up any earlier Wild Theme: drawings, page background, media, tag lines
    await cleanParts(pkg, main, mainRels, { anyBackground: true });

    // ---- media
    const ext = (assets.backdrop.ext || "png").toLowerCase().replace("jpg", "jpeg");
    if (!IMG_CT[ext]) throw new Error("Backdrop must be GIF, PNG or JPEG.");
    const stampId = Date.now().toString(36);
    const media = {
      backdrop: `word/media/wildtheme-backdrop-${stampId}.${ext}`,
      panel: `word/media/wildtheme-panel-${stampId}.png`,
      marker: `word/media/wildtheme-marker-${stampId}.png`,
    };
    pkg.putBinary(media.backdrop, assets.backdrop.bytes);
    // assets.panel: { bytes } used everywhere, or a function (section, rectTwips) -> PNG bytes
    // called once per distinct page shape so landscape sections get an undistorted panel.
    const panelFn = typeof assets.panel === "function" ? assets.panel : null;
    const panelMedia = new Map();     // rect key -> media path
    async function panelPath(geom) {
      if (!panelFn) return assets.panel && assets.panel.bytes ? media.panel : null;
      const rect = panelRect(geom);
      const key = `${rect.w}x${rect.h}`;
      if (!panelMedia.has(key)) {
        const bytes = await panelFn(geom, rect);
        if (!bytes) { panelMedia.set(key, null); return null; }
        const path = `word/media/wildtheme-panel-${stampId}-${panelMedia.size + 1}.png`;
        pkg.putBinary(path, bytes);
        panelMedia.set(key, path);
      }
      return panelMedia.get(key);
    }
    if (!panelFn && assets.panel && assets.panel.bytes) pkg.putBinary(media.panel, assets.panel.bytes);
    pkg.putBinary(media.marker, b64ToBytes(MARKER_PNG_B64));
    ct.ensureDefault(ext, IMG_CT[ext]);
    ct.ensureDefault("png", IMG_CT.png);

    // ---- theme colours + fonts
    let themeRel = mainRels.byType(REL.theme)[0];
    let themePath = themeRel ? resolveTarget(main, themeRel.target) : "word/theme/theme1.xml";
    if (!themeRel) mainRels.add(REL.theme, relativeTarget(main, themePath));
    pkg.zip.file(themePath, themeXml(theme), { createFolders: false });
    pkg.cache.delete(themePath);
    ct.ensureOverride(themePath, CT.theme);

    // ---- styles
    let stylesRel = mainRels.byType(REL.styles)[0];
    let stylesPath;
    if (stylesRel) stylesPath = resolveTarget(main, stylesRel.target);
    else {
      stylesPath = "word/styles.xml";
      pkg.setXml(stylesPath, parseXml(`<w:styles xmlns:w="${NS.w}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style></w:styles>`));
      mainRels.add(REL.styles, relativeTarget(main, stylesPath));
      ct.ensureOverride(stylesPath, CT.styles);
    }
    const styles = await pkg.xml(stylesPath);
    applyStyles(styles, theme);
    pkg.touch(stylesPath);

    // ---- settings: lock the theme and style set so they're hard to swap out by accident
    let settingsPath;
    const sRel = mainRels.byType(REL.settings)[0];
    if (sRel) settingsPath = resolveTarget(main, sRel.target);
    else {
      settingsPath = "word/settings.xml";
      pkg.setXml(settingsPath, parseXml(`<w:settings xmlns:w="${NS.w}"/>`));
      mainRels.add(REL.settings, relativeTarget(main, settingsPath));
      ct.ensureOverride(settingsPath, "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml");
    }
    const settingsDoc = await pkg.xml(settingsPath);
    if (settingsDoc) { lockTheme(settingsDoc); ensureSetting(settingsDoc, "displayBackgroundShape"); pkg.touch(settingsPath); }

    // ---- which header types are in play; page geometry per section
    const evenOdd = await evenOddOf(pkg, main, mainRels);
    const infos = await sectionInfos(pkg, main, doc, mainRels, evenOdd);

    // ---- the scene on every page: the document's page background (full colour, never dimmed;
    //      Word shows the first frame of a GIF here)
    const bgRid = mainRels.add(REL.image, relativeTarget(main, media.backdrop));
    ensureNs(doc, "v", NS.v); ensureNs(doc, "o", NS.o);
    doc.documentElement.insertBefore(fragment(doc, pageBackgroundXml(bgRid))[0], doc.documentElement.firstChild);

    // ---- headers: the light panel, section by section (Word dims header content while you
    //      edit the body, which is fine for a pale panel but would wash out the scene)
    const headerDocs = [];
    for (const r of mainRels.byType(REL.header)) { const d = await pkg.xml(resolveTarget(main, r.target)); if (d) headerDocs.push(d); }
    let nextId = maxDocPrId([doc, ...headerDocs]) + 1;
    const nextSdtId = sdtIdPicker([doc, ...headerDocs]);
    const stampedGeom = new Map();          // header part path -> geometry it was stamped for
    const inherited = {};                   // header type -> part path currently in force
    let headerCount = 0;
    const freshHeaderPath = () => {
      let n = 1;
      while (pkg.has(`word/header${n}.xml`)) n++;
      return `word/header${n}.xml`;
    };
    // One picture per run: Word only draws the first <w:drawing> in a run.
    const pictureRun = (anchor) => `<w:r w:rsidR="${OUR_RSID}"><w:rPr><w:noProof/></w:rPr><w:drawing>${anchor}</w:drawing></w:r>`;
    const panelAnchor = (geom, rId, z) => {
      const pr = panelRect(geom);
      return anchorXml({ id: nextId++, name: "WildThemePanel", descr: "Decorative page panel (Wild Theme)", title: "",
        rId, x: pr.x * TWIP_EMU, y: pr.y * TWIP_EMU, cx: pr.w * TWIP_EMU, cy: pr.h * TWIP_EMU, z });
    };
    async function stampHeader(path, geom) {
      stampedGeom.set(path, geom);
      const pPath = await panelPath(geom);
      if (!pPath) return;
      const hd = await pkg.xml(path);
      const rels = await pkg.rels(path);
      const rId = rels.add(REL.image, relativeTarget(path, pPath));
      if (HIDE_SDT) ensureW15(hd);
      insertLeadRun(hd, hd.documentElement, lockedSdt("WildTheme", nextSdtId(), pictureRun(panelAnchor(geom, rId, 2)), HIDE_SDT));
      pkg.touch(path);
      headerCount++;
    }
    async function newHeader(sourcePath) {
      const path = freshHeaderPath();
      if (sourcePath) {
        const src = await pkg.xml(sourcePath);
        const copy = parseXml(serialize(src));
        removeOurDrawings(copy);
        pkg.setXml(path, copy);
        const srcRels = pkg.zip.file(relsPathOf(sourcePath));
        if (srcRels) {
          const rd = parseXml(await srcRels.async("string"));
          for (const el of Array.from(rd.getElementsByTagName("Relationship"))) {
            if (isOurMedia(el.getAttribute("Target") || "")) el.parentNode.removeChild(el);
          }
          pkg.setXml(relsPathOf(path), rd);
        }
      } else {
        pkg.setXml(path, parseXml(`<w:hdr xmlns:w="${NS.w}" xmlns:r="${NS.r}"><w:p w:rsidR="${OUR_RSID}"><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p></w:hdr>`));
      }
      ct.ensureOverride(path, CT.header);
      return path;
    }
    function addHeaderRef(sectPr, type, rId) {
      const ref = fragment(doc, `<w:headerReference w:type="${type}" r:id="${rId}"/>`)[0];
      // headerReference elements come first in sectPr
      let before = sectPr.firstChild;
      while (before && before.nodeType === 1 && before.namespaceURI === NS.w && (before.localName === "headerReference")) before = before.nextSibling;
      while (before && before.nodeType !== 1) before = before.nextSibling;
      sectPr.insertBefore(ref, before || null);
    }

    const secEls = sections(doc);
    if (!secEls.length) {
      const body = all(doc, NS.w, "body")[0];
      const s = fragment(doc, `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`)[0];
      body.appendChild(s);
      secEls.push(s);
    }
    const havePanel = !!(panelFn || (assets.panel && assets.panel.bytes));
    for (let si = 0; si < (havePanel ? secEls.length : 0); si++) {
      const sectPr = secEls[si];
      const geom = infos[si] || sectionGeometry(sectPr);
      const types = ["default"];
      if (onOff(kid(sectPr, NS.w, "titlePg"))) types.push("first");
      if (evenOdd) types.push("even");
      for (const type of types) {
        const ref = kids(sectPr, NS.w, "headerReference").find((h) => wAttr(h, "type") === type);
        if (ref) {
          const rid = ref.getAttributeNS(NS.r, "id") || ref.getAttribute("r:id");
          const rel = mainRels.byId(rid);
          if (!rel) continue;
          let path = resolveTarget(main, rel.target);
          if (stampedGeom.has(path) && !sameGeom(stampedGeom.get(path), geom)) {
            // the same header is shared with a differently-sized section: give this one its own copy
            path = await newHeader(path);
            ref.setAttributeNS(NS.r, "r:id", mainRels.add(REL.header, relativeTarget(main, path)));
          }
          if (!stampedGeom.has(path)) await stampHeader(path, geom);
          inherited[type] = path;
        } else if (inherited[type]) {
          // Word reuses the previous section's header; only split it off if the page size differs
          if (sameGeom(stampedGeom.get(inherited[type]), geom)) continue;
          const path = await newHeader(inherited[type]);
          addHeaderRef(sectPr, type, mainRels.add(REL.header, relativeTarget(main, path)));
          await stampHeader(path, geom);
          inherited[type] = path;
        } else {
          // nothing defined yet: Word would show an empty header here, so start from an empty one
          const path = await newHeader(null);
          addHeaderRef(sectPr, type, mainRels.add(REL.header, relativeTarget(main, path)));
          await stampHeader(path, geom);
          inherited[type] = path;
        }
      }
    }
    pkg.touch(main);

    // ---- page 1, in the body's first paragraph (locked): the marker the extension finds in
    //      Word for the web, the animated backdrop (body pictures aren't dimmed) and its panel.
    //      Body behind-text pictures sit above the header layer, so page 1 needs its own panel.
    const body = all(doc, NS.w, "body")[0];
    const first = infos[0] || sectionGeometry(secEls[0]);
    const markerRid = mainRels.add(REL.image, relativeTarget(main, media.marker));
    const page1Panel = await panelPath(first);
    let page1 = pictureRun(anchorXml({ id: nextId++, name: "WildThemeMarker", descr, title: "Wild Theme", rId: markerRid, x: 0, y: 0, cx: 9525, cy: 9525, z: 3 })) +
      pictureRun(anchorXml({ id: nextId++, name: "WildThemeBackdrop", descr, title: "Wild Theme", rId: bgRid, x: 0, y: 0,
        cx: first.pageW * TWIP_EMU, cy: first.pageH * TWIP_EMU, z: 1 }));
    if (page1Panel) page1 += pictureRun(panelAnchor(first, mainRels.add(REL.image, relativeTarget(main, page1Panel)), 2));
    if (HIDE_SDT) ensureW15(doc);
    insertLeadRun(doc, body, lockedSdt("WildThemeMarker", nextSdtId(), page1, HIDE_SDT));

    // ---- metadata
    await setCustomProps(pkg, { WildTheme: theme.id, WildThemeName: theme.name || theme.id, WildThemeVersion: "3" });

    const out = await pkg.save();
    out.stats = { sections: secEls.length, headersStamped: headerCount };
    return out;
  }

  // ------------------------------------------------------------------ public: unstamp
  // Removes everything Wild Theme added (pictures, locked controls, marker, metadata, theme locks).
  // Theme colours/fonts and heading styles stay as they are: the originals aren't known.
  async function unstamp(bytes) {
    const pkg = await Pkg.open(bytes);
    const ct = await pkg.contentTypes();
    const main = await pkg.mainPart();
    const doc = await pkg.xml(main);
    const mainRels = await pkg.rels(main);
    const removedBackground = await cleanParts(pkg, main, mainRels);

    // header parts that only ever held our picture: drop them and their references
    for (const r of mainRels.byType(REL.header)) {
      const path = resolveTarget(main, r.target);
      const hd = await pkg.xml(path);
      if (!hd) continue;
      const blocks = Array.from(hd.documentElement.childNodes).filter((c) => c.nodeType === 1);
      const onlyOurs = blocks.length === 1 && blocks[0].localName === "p" && wAttr(blocks[0], "rsidR") === OUR_RSID &&
        !Array.from(blocks[0].childNodes).some((c) => c.nodeType === 1 && !(c.namespaceURI === NS.w && c.localName === "pPr"));
      if (!onlyOurs) continue;
      const stillShared = mainRels.byType(REL.header).some((o) => o.id !== r.id && resolveTarget(main, o.target) === path);
      for (const ref of all(doc, NS.w, "headerReference")) {
        if ((ref.getAttributeNS(NS.r, "id") || ref.getAttribute("r:id")) === r.id) ref.parentNode.removeChild(ref);
      }
      mainRels.remove(r.id);
      if (!stillShared) {
        pkg.remove(path);
        pkg.remove(relsPathOf(path));
        ct.removeOverride(path);
      }
    }
    pkg.touch(main);

    // metadata
    const custom = await pkg.xml("docProps/custom.xml");
    if (custom) {
      const root = custom.documentElement;
      for (const el of Array.from(root.getElementsByTagNameNS(NS.cp, "property"))) {
        if (/^WildTheme/.test(el.getAttribute("name") || "")) root.removeChild(el);
      }
      if (!root.getElementsByTagNameNS(NS.cp, "property").length) {
        pkg.remove("docProps/custom.xml");
        ct.removeOverride("docProps/custom.xml");
        const rootRels = await pkg.rels("");
        for (const rr of rootRels.byType(REL.custom)) rootRels.remove(rr.id);
      } else pkg.touch("docProps/custom.xml");
    }

    // v1 documents put "Wild Theme: <name>" in the keywords
    const core = await pkg.xml("docProps/core.xml");
    if (core) {
      for (const k of Array.from(core.getElementsByTagNameNS("http://schemas.openxmlformats.org/package/2006/metadata/core-properties", "keywords"))) {
        if (/^\s*Wild Theme:/.test(k.textContent)) { k.parentNode.removeChild(k); pkg.touch("docProps/core.xml"); }
      }
    }

    // theme locks
    const sRel = mainRels.byType(REL.settings)[0];
    if (sRel) {
      const sp = resolveTarget(main, sRel.target);
      const sd = await pkg.xml(sp);
      if (sd) {
        removeSetting(sd, "styleLockTheme");
        removeSetting(sd, "styleLockQFSet");
        if (removedBackground && !backgroundOf(doc)) removeSetting(sd, "displayBackgroundShape");
        pkg.touch(sp);
      }
    }

    // our name on the colour scheme
    const tRel = mainRels.byType(REL.theme)[0];
    if (tRel) {
      const tp = resolveTarget(main, tRel.target);
      const td = await pkg.xml(tp);
      if (td) {
        const n = td.documentElement.getAttribute("name") || "";
        if (/^Wild Theme /.test(n)) { td.documentElement.setAttribute("name", n.replace(/^Wild Theme /, "")); pkg.touch(tp); }
      }
    }
    return pkg.save();
  }

  // ------------------------------------------------------------------ public: new document
  async function newDocument(theme, assets) {
    const zip = new JSZip();
    const name = theme.name || "Wild Theme";
    const p = (text, style, extra = "") =>
      `<w:p><w:pPr>${style ? `<w:pStyle w:val="${style}"/>` : ""}${extra}</w:pPr><w:r><w:t xml:space="preserve">${esc(text)}</w:t></w:r></w:p>`;
    const bullet = (t) => p(t, "ListParagraph", '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>');
    const body = [
      p("Your Title Here", "Title"),
      p(`${theme.blurb || name}. Replace this with your own subtitle.`, "Subtitle"),
      p("Start writing", "Heading1"),
      p(`This document has the ${name} theme built into the file. The backdrop, colours and fonts travel with it, so anyone who opens it sees the same look, with no extensions or add-ins needed.`),
      bullet("Type over this text, or delete it all and start fresh."),
      bullet("Use the Styles gallery on the Home tab for headings, quotes and lists in the theme’s colours."),
      bullet("The backdrop sits behind every page, so your text always stays on a clean, readable panel."),
      p("A pull quote", "Heading2"),
      p("Make it bold. Make it weird. Keep it readable.", "Quote"),
      p("Notes", "Heading2"),
      p("The theme is locked into the file, and every new page gets the backdrop too. To change or remove it, use the Wild Themes add-in or browser extension."),
    ].join("");
    const W = `xmlns:w="${NS.w}" xmlns:r="${NS.r}"`;
    zip.file("[Content_Types].xml", XML_DECL + `<Types xmlns="${NS.ct}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `<Override PartName="/word/styles.xml" ContentType="${CT.styles}"/>` +
      `<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>` +
      `<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>` +
      `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>` +
      `<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`, { createFolders: false });
    zip.file("_rels/.rels", XML_DECL + `<Relationships xmlns="${NS.pr}"><Relationship Id="rId1" Type="${REL.officeDocument}" Target="word/document.xml"/>` +
      `<Relationship Id="rId2" Type="${REL.core}" Target="docProps/core.xml"/>` +
      `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`, { createFolders: false });
    zip.file("word/_rels/document.xml.rels", XML_DECL + `<Relationships xmlns="${NS.pr}"><Relationship Id="rId1" Type="${REL.styles}" Target="styles.xml"/>` +
      `<Relationship Id="rId2" Type="${REL.settings}" Target="settings.xml"/>` +
      `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>`, { createFolders: false });
    zip.file("word/document.xml", XML_DECL + `<w:document ${W}><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>` +
      `<w:pgMar w:top="1800" w:right="1400" w:bottom="1500" w:left="1400" w:header="432" w:footer="432" w:gutter="0"/><w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr></w:body></w:document>`, { createFolders: false });
    zip.file("word/styles.xml", XML_DECL + `<w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr>${MINOR}<w:color w:val="${hex((theme.doc || {}).dk1, "1F1F1F")}" w:themeColor="text1"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-GB" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>` +
      `<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>` +
      `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>` +
      `<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="34"/><w:qFormat/><w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>` +
      `<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>` +
      `<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>` +
      `<w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style></w:styles>`, { createFolders: false });
    const lv = Array.from({ length: 9 }, (_, i) => `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="${"●○■"[i % 3]}"/><w:lvlJc w:val="left"/>` +
      `<w:pPr><w:ind w:left="${720 + 360 * i}" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:hint="default"/><w:color w:val="${hex((theme.doc || {}).accent1 || theme.accent)}" w:themeColor="accent1"/></w:rPr></w:lvl>`).join("");
    zip.file("word/numbering.xml", XML_DECL + `<w:numbering ${W}><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>${lv}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`, { createFolders: false });
    zip.file("word/settings.xml", XML_DECL + `<w:settings ${W}><w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/>` +
      `<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>` +
      `<w:themeFontLang w:val="en-GB"/><w:clrSchemeMapping w:bg1="light1" w:t1="dark1" w:bg2="light2" w:t2="dark2" w:accent1="accent1" w:accent2="accent2" w:accent3="accent3" w:accent4="accent4" w:accent5="accent5" w:accent6="accent6" w:hyperlink="hyperlink" w:followedHyperlink="followedHyperlink"/></w:settings>`, { createFolders: false });
    zip.file("docProps/core.xml", XML_DECL + `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${esc(name)}</dc:title></cp:coreProperties>`, { createFolders: false });
    zip.file("docProps/app.xml", XML_DECL + `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Office Word</Application></Properties>`, { createFolders: false });
    const base = await zip.generateAsync({ type: "uint8array" });
    return stamp(base, theme, assets);
  }

  // The page geometry newDocument() uses (A4), so callers can render assets before calling it.
  const NEW_DOC_SECTION = { pageW: 11906, pageH: 16838, margins: { top: 1800, right: 1400, bottom: 1500, left: 1400, header: 432, footer: 432 } };

  return { analyze, stamp, unstamp, newDocument, panelRect, NEW_DOC_SECTION, markerDescr };
});
