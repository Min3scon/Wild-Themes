// Wild Themes — writes a theme INTO a .docx file, deeply and reversibly.
// The theme lives in the document itself: theme colours and fonts, the document's styles
// (body text, title, headings, quotes, links, captions, header/footer), table and list colours,
// old hard-coded theme colours turned into theme references, a still scene behind every page
// with a light reading panel, and metadata. Anyone who opens the file in Word sees it with
// nothing installed. stamp() also returns a backup of everything it replaced, and
// unstamp(bytes, backup) puts the original look back exactly.
//
// Browser: load vendor/jszip.min.js first (and wtcore.js for the shared palette), then this file -> self.WWStamper.
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
  // The shared core (palette, settings). Looked up at call time: wtcore.js requires this file while loading.
  const core = () => (typeof self !== "undefined" && self.WTCore) || (typeof require === "function" ? require("./wtcore.js") : null);
  function theCore() { try { return core(); } catch (_) { return null; } }

  // ------------------------------------------------------------------ namespaces
  const NS = {
    w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    a: "http://schemas.openxmlformats.org/drawingml/2006/main",
    pic: "http://schemas.openxmlformats.org/drawingml/2006/picture",
    mc: "http://schemas.openxmlformats.org/markup-compatibility/2006",
    w14: "http://schemas.microsoft.com/office/word/2010/wordml",
    w15: "http://schemas.microsoft.com/office/word/2012/wordml",
    v: "urn:schemas-microsoft-com:vml",
    o: "urn:schemas-microsoft-com:office:office",
    xmlns: "http://www.w3.org/2000/xmlns/",
    pr: "http://schemas.openxmlformats.org/package/2006/relationships",
    ct: "http://schemas.openxmlformats.org/package/2006/content-types",
    cp: "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
    vt: "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
    coreProps: "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
  };
  const RELS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
  const REL = {
    officeDocument: RELS + "officeDocument", styles: RELS + "styles", theme: RELS + "theme", header: RELS + "header",
    footer: RELS + "footer", footnotes: RELS + "footnotes", endnotes: RELS + "endnotes", numbering: RELS + "numbering",
    image: RELS + "image", settings: RELS + "settings", custom: RELS + "custom-properties",
    core: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
  };
  const CT = {
    header: "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
    theme: "application/vnd.openxmlformats-officedocument.theme+xml",
    styles: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
    settings: "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
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
  const SLOTS = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
  const ACCENTS = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"];

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
  const toU8 = (b) => (b instanceof Uint8Array ? b : new Uint8Array(b));
  // short content hash: media names follow the picture, so re-applying gives the same part names
  function fnv(bytes) {
    const u = typeof bytes === "string" ? Array.from(bytes, (c) => c.charCodeAt(0) & 255) : toU8(bytes);
    let h = 0x811c9dc5;
    for (let i = 0; i < u.length; i++) { h ^= u[i]; h = Math.imul(h, 0x01000193); }
    return (h >>> 0).toString(16).padStart(8, "0");
  }
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
  // One element as self-contained XML (for the backup).
  const elXml = (el) => new SerializerImpl().serializeToString(el);
  // Bring a backed-up element back into a part, dropping namespace declarations the part's root already makes.
  function restoredEl(doc, xml) {
    const el = doc.importNode(parseXml(xml).documentElement, true);
    const root = doc.documentElement;
    for (const n of [el, ...Array.from(el.getElementsByTagNameNS("*", "*"))]) {
      for (const a of Array.from(n.attributes || [])) {
        if (!/^xmlns(:|$)/.test(a.name)) continue;
        if (root.hasAttribute(a.name) && root.getAttribute(a.name) === a.value) n.removeAttribute(a.name);
      }
    }
    return el;
  }
  // parse a fragment written with prefixes w:, wp:, a:, pic:, r:, w14:, w15:, v:, o:
  function fragment(doc, xml) {
    const wrap = `<root xmlns:w="${NS.w}" xmlns:r="${NS.r}" xmlns:wp="${NS.wp}" xmlns:a="${NS.a}" xmlns:pic="${NS.pic}" xmlns:w14="${NS.w14}" xmlns:w15="${NS.w15}" xmlns:v="${NS.v}" xmlns:o="${NS.o}">${xml}</root>`;
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
  const wSet = (el, name, v) => el.setAttributeNS(NS.w, "w:" + name, v);
  const wDel = (el, name) => { el.removeAttributeNS(NS.w, name); if (el.hasAttribute("w:" + name)) el.removeAttribute("w:" + name); };
  const onOff = (el) => {
    if (!el) return false;
    const v = wAttr(el, "val");
    return !(v === "0" || v === "false" || v === "off");
  };
  const ridOf = (el) => el.getAttributeNS(NS.r, "id") || el.getAttribute("r:id");
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
  // Insert (or replace) a child, keeping the schema's element order. Children not in `order`
  // (other namespaces such as w14, or later elements like rPrChange) count as coming last.
  function setOrdered(parent, el, order) {
    if (el.namespaceURI === NS.w || el.namespaceURI === NS.w14) for (const old of kids(parent, el.namespaceURI, el.localName)) parent.removeChild(old);
    const rank = (n) => (n.namespaceURI === NS.w && order.includes(n.localName) ? order.indexOf(n.localName) : order.length);
    const mine = rank(el);
    let before = null;
    for (let n = parent.firstChild; n; n = n.nextSibling) if (n.nodeType === 1 && rank(n) > mine) { before = n; break; }
    parent.insertBefore(el, before);
    return el;
  }
  const RPR_ORDER = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow", "emboss",
    "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u",
    "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath"];
  const STYLE_ORDER = ["name", "aliases", "basedOn", "next", "link", "autoRedefine", "hidden", "uiPriority", "semiHidden", "unhideWhenUsed",
    "qFormat", "locked", "personal", "personalCompose", "personalReply", "rsid", "pPr", "rPr", "tblPr", "trPr", "tcPr", "tblStylePr"];
  // CT_Settings is an ordered sequence.
  const SETTINGS_ORDER = ["writeProtection", "view", "zoom", "removePersonalInformation", "removeDateAndTime",
    "doNotDisplayPageBoundaries", "displayBackgroundShape", "printPostScriptOverText", "printFractionalCharacterWidth",
    "printFormsData", "embedTrueTypeFonts", "embedSystemFonts", "saveSubsetFonts", "saveFormsData", "mirrorMargins",
    "alignBordersAndEdges", "bordersDoNotSurroundHeader", "bordersDoNotSurroundFooter", "gutterAtTop", "hideSpellingErrors",
    "hideGrammaticalErrors", "activeWritingStyle", "proofState", "formsDesign", "attachedTemplate", "linkStyles",
    "stylePaneFormatFilter", "stylePaneSortMethod", "documentType", "mailMerge", "revisionView", "trackRevisions",
    "doNotTrackMoves", "doNotTrackFormatting", "documentProtection", "autoFormatOverride", "styleLockTheme", "styleLockQFSet",
    "defaultTabStop", "autoHyphenation", "consecutiveHyphenLimit", "hyphenationZone", "doNotHyphenateCaps", "showEnvelope",
    "summaryLength", "clickAndTypeStyle", "defaultTableStyle", "evenAndOddHeaders", "bookFoldRevPrinting", "bookFoldPrinting",
    "bookFoldPrintingSheets", "drawingGridHorizontalSpacing", "drawingGridVerticalSpacing", "displayHorizontalDrawingGridEvery",
    "displayVerticalDrawingGridEvery", "doNotUseMarginsForDrawingGridOrigin", "drawingGridHorizontalOrigin",
    "drawingGridVerticalOrigin", "doNotShadeFormData", "noPunctuationKerning", "characterSpacingControl", "printTwoOnOne",
    "strictFirstAndLastChars", "noLineBreaksAfter", "noLineBreaksBefore", "savePreviewPicture", "doNotValidateAgainstSchema",
    "saveInvalidXml", "ignoreMixedContent", "alwaysShowPlaceholderText", "doNotDemarcateInvalidXml", "saveXmlDataOnly",
    "useXSLTWhenSaving", "saveThroughXslt", "showXMLTags", "alwaysMergeEmptyNamespace", "updateFields", "hdrShapeDefaults",
    "footnotePr", "endnotePr", "compat", "docVars", "rsids", "mathPr", "attachedSchema", "themeFontLang", "clrSchemeMapping",
    "doNotIncludeSubdocsInStats", "doNotAutoCompressPictures", "forceUpgrade", "captions", "readModeInkLockDown", "smartTagType",
    "schemaLibrary", "shapeDefaults", "doNotEmbedSmartTags", "decimalSymbol", "listSeparator"];

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
    // A part's relationships. A missing .rels part is created in memory and only written once something is added.
    async rels(part) {
      const p = relsPathOf(part);
      let doc = await this.xml(p);
      if (!doc) { doc = parseXml(`<Relationships xmlns="${NS.pr}"/>`); this.cache.set(p, doc); }
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
      const defaults = () => Array.from(root.getElementsByTagName("Default"));
      return {
        // returns true when it had to add the entry
        ensureDefault(ext, type) {
          if (defaults().some((d) => (d.getAttribute("Extension") || "").toLowerCase() === ext)) return false;
          const el = doc.createElementNS(NS.ct, "Default");
          el.setAttribute("Extension", ext); el.setAttribute("ContentType", type);
          const firstOverride = root.getElementsByTagName("Override")[0];
          root.insertBefore(el, firstOverride || null); self_.touch("[Content_Types].xml");
          return true;
        },
        removeDefault(ext) {
          for (const d of defaults()) if ((d.getAttribute("Extension") || "").toLowerCase() === ext) { root.removeChild(d); self_.touch("[Content_Types].xml"); }
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
          const rel = mainRels.byId(ridOf(ref));
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

  // ------------------------------------------------------------------ palette + settings (shared with wtcore.js)
  function paletteOf(theme) {
    const c = theCore();
    if (c && c.palette) return c.palette(theme);
    // the same rules as WTCore.palette, for pages that load only this file
    const d = theme.doc || {};
    const accent1 = hex(d.accent1 || theme.accent, "4472C4"), accent2 = hex(d.accent2 || theme.accent2, "ED7D31");
    return {
      dk1: hex(d.dk1, "1F1F1F"), lt1: "FFFFFF", dk2: hex(d.dk2, "1F3864"), lt2: hex(d.lt2, "EEEEEE"), accent1, accent2,
      accent3: hex(d.accent3 || theme.accent, "A5A5A5"), accent4: hex(d.accent4 || theme.accent2, "FFC000"),
      accent5: hex(d.accent5, "5B9BD5"), accent6: hex(d.accent6, "70AD47"), hlink: accent2, folHlink: hex(d.dk2, "1F3864"),
      major: d.major || "Aptos Display", minor: d.minor || "Aptos", panel: hex(d.panel, "FFFFFF"),
    };
  }
  const DEFAULT_SETTINGS = { backdrop: true, panel: 100, styles: true, tables: true, lists: true, direct: true, lock: true, glow: false };
  function settingsOf(given) {
    const c = theCore();
    if (c && c.settingsFor) return c.settingsFor("docx", given || {});
    const out = { ...DEFAULT_SETTINGS };
    for (const k of Object.keys(out)) {
      if (!given || given[k] === undefined) continue;
      out[k] = typeof out[k] === "boolean" ? (given[k] === true || given[k] === "true" || given[k] === 1) : given[k];
    }
    return out;
  }

  // ------------------------------------------------------------------ theme part
  function themeXml(theme, P) {
    const name = esc(theme.name || "Wild Theme");
    const major = esc(P.major), minor = esc(P.minor);
    const g = (stops) => `<a:gradFill rotWithShape="1"><a:gsLst>${stops.map(([p, m]) => `<a:gs pos="${p}"><a:schemeClr val="phClr">${m}</a:schemeClr></a:gs>`).join("")}</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>`;
    const ln = (w) => `<a:ln w="${w}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>`;
    const slots = SLOTS.map((k) => `<a:${k}><a:srgbClr val="${P[k]}"/></a:${k}>`).join("");
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
  // Colour slots and latin fonts of a theme part (the file's original theme, before we replace it).
  function themeInfo(xml) {
    let d;
    try { d = parseXml(xml); } catch (_) { return null; }
    const cs = all(d, NS.a, "clrScheme")[0], fs = all(d, NS.a, "fontScheme")[0];
    const slots = {};
    if (cs) for (const slot of SLOTS) {
      const el = kid(cs, NS.a, slot);
      const c = el && Array.from(el.childNodes).find((n) => n.nodeType === 1);
      if (c) slots[slot] = String((c.localName === "sysClr" ? c.getAttribute("lastClr") : c.getAttribute("val")) || "").toUpperCase();
    }
    const font = (k) => { const f = fs && kid(fs, NS.a, k); const l = f && kid(f, NS.a, "latin"); return (l && l.getAttribute("typeface")) || null; };
    return { slots, major: font("majorFont"), minor: font("minorFont") };
  }

  // ------------------------------------------------------------------ "match old colours": direct formatting -> theme references
  // Office's stock palettes (2013-2022, 2023+, 2007-2010): text pasted from other documents often carries these
  // as hard-coded colours. Only their accents and Text 2 are matched; the file's own theme is matched in full.
  const STOCK = [
    { dk2: "44546A", accent1: "4472C4", accent2: "ED7D31", accent3: "A5A5A5", accent4: "FFC000", accent5: "5B9BD5", accent6: "70AD47" },
    { dk2: "0E2841", accent1: "156082", accent2: "E97132", accent3: "196B24", accent4: "0F9ED5", accent5: "A02B93", accent6: "4EA72E" },
    { dk2: "1F497D", accent1: "4F81BD", accent2: "C0504D", accent3: "9BBB59", accent4: "8064A2", accent5: "4BACC6", accent6: "F79646" },
  ];
  const W_THEME = { dk2: "text2", lt2: "background2", hlink: "hyperlink", folHlink: "followedHyperlink" };
  const A_SCHEME = { dk2: "tx2", lt2: "bg2", hlink: "hlink", folHlink: "folHlink" };
  function directMaps(orig) {
    const colors = new Map(), accents = new Map();
    const add = (m, v, slot) => { v = String(v || "").toUpperCase(); if (/^[0-9A-F]{6}$/.test(v) && !m.has(v)) m.set(v, slot); };
    if (orig) {
      for (const s of ACCENTS) { add(colors, orig.slots[s], s); add(accents, orig.slots[s], s); }
      for (const s of ["dk2", "lt2", "hlink", "folHlink"]) add(colors, orig.slots[s], s);
    }
    for (const set of STOCK) for (const s of [...ACCENTS, "dk2"]) { add(colors, set[s], s); if (s !== "dk2") add(accents, set[s], s); }
    for (const v of ["000000", "FFFFFF"]) { colors.delete(v); accents.delete(v); }   // never plain black / white
    return { colors, accents, minor: orig && orig.minor, major: orig && orig.major };
  }
  // Formatting inside tracked-change records (w:rPrChange ...) describes the old state: leave it.
  const inRevision = (el) => { for (let n = el.parentNode; n && n.nodeType === 1; n = n.parentNode) if (n.namespaceURI === NS.w && /Change$/.test(n.localName)) return true; return false; };
  const dmlColours = (doc) => Array.from(doc.getElementsByTagNameNS(NS.a, "*")).filter((e) => e.localName === "srgbClr" || e.localName === "schemeClr");
  function renameColour(doc, el, local, val) {
    const n = doc.createElementNS(NS.a, (el.prefix ? el.prefix + ":" : "") + local);
    n.setAttribute("val", val);
    while (el.firstChild) n.appendChild(el.firstChild);
    el.parentNode.replaceChild(n, el);
    return n;
  }
  // The conversions a file needs, found from the file alone: [kind, index, original value, theme reference, slot].
  // The same list goes in the backup whatever the theme or settings, so unstamp can undo them exactly.
  function findDirect(doc, maps) {
    const rec = [];
    all(doc, NS.w, "color").forEach((el, i) => {
      if (wAttr(el, "themeColor") || inRevision(el)) return;
      const v = wAttr(el, "val").toUpperCase(), slot = maps.colors.get(v);
      if (slot) rec.push(["c", i, v, W_THEME[slot] || slot, slot]);
    });
    all(doc, NS.w, "shd").forEach((el, i) => {
      if (wAttr(el, "themeFill") || inRevision(el)) return;
      const v = wAttr(el, "fill").toUpperCase(), slot = maps.colors.get(v);
      if (slot) rec.push(["s", i, v, W_THEME[slot] || slot, slot]);
    });
    if (maps.minor || maps.major) all(doc, NS.w, "rFonts").forEach((el, i) => {
      if (inRevision(el)) return;
      const set = [];
      for (const [a, ta] of [["ascii", "asciiTheme"], ["hAnsi", "hAnsiTheme"]]) {
        if (wAttr(el, ta)) continue;
        const f = wAttr(el, a).toLowerCase();
        const t = !f ? null : maps.minor && f === maps.minor.toLowerCase() ? "minorHAnsi" : maps.major && f === maps.major.toLowerCase() ? "majorHAnsi" : null;
        if (t) set.push(ta + "=" + t);
      }
      if (set.length) rec.push(["f", i, "", set.join(","), ""]);
    });
    dmlColours(doc).forEach((el, i) => {
      if (el.localName !== "srgbClr") return;
      const v = String(el.getAttribute("val") || "").toUpperCase(), slot = maps.accents.get(v);
      if (slot) rec.push(["d", i, v, A_SCHEME[slot] || slot, slot]);
    });
    return rec;
  }
  const directLists = (doc) => {
    const lists = {};
    return (k) => lists[k] || (lists[k] = k === "c" ? all(doc, NS.w, "color") : k === "s" ? all(doc, NS.w, "shd") : k === "f" ? all(doc, NS.w, "rFonts") : dmlColours(doc));
  };
  function applyDirect(doc, recs, P, st) {
    const get = directLists(doc);
    for (const [k, i, , ours, slot] of recs) {
      const el = get(k)[i];
      if (k === "c") { wSet(el, "val", P[slot]); wSet(el, "themeColor", ours); st.colors++; }
      else if (k === "s") { wSet(el, "fill", P[slot]); wSet(el, "themeFill", ours); st.fills++; }
      else if (k === "f") { for (const pv of ours.split(",")) { const [a, v] = pv.split("="); wSet(el, a, v); } st.fonts++; }
      else { get("d")[i] = renameColour(doc, el, "schemeClr", ours); st.drawing++; }
    }
  }
  // Undo applyDirect(). Only when every recorded change is still in place: a part another editor has
  // rewritten (or that was edited in between) is left alone, and since its theme references then point at
  // the restored original theme, Word still shows those runs in the original theme's colours.
  function revertDirect(doc, recs) {
    const get = directLists(doc);
    const still = ([k, i, , ours]) => {
      const el = get(k)[i];
      if (!el) return false;
      if (k === "c") return wAttr(el, "themeColor") === ours;
      if (k === "s") return wAttr(el, "themeFill") === ours;
      if (k === "f") return ours.split(",").every((pv) => { const [a, v] = pv.split("="); return wAttr(el, a) === v; });
      return el.localName === "schemeClr" && el.getAttribute("val") === ours;
    };
    if (!recs.length || !recs.every(still)) return 0;
    for (const [k, i, orig, ours] of recs) {
      const el = get(k)[i];
      if (k === "c") { wSet(el, "val", orig); wDel(el, "themeColor"); }
      else if (k === "s") { wSet(el, "fill", orig); wDel(el, "themeFill"); }
      else if (k === "f") for (const pv of ours.split(",")) wDel(el, pv.split("=")[0]);
      else get("d")[i] = renameColour(doc, el, "srgbClr", orig);
    }
    return recs.length;
  }

  // ------------------------------------------------------------------ styles
  const MAJOR = '<w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:cstheme="majorBidi"/>';
  const MINOR = '<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:cstheme="minorBidi"/>';
  const GLOW = '<w14:glow w14:rad="63500"><w14:schemeClr w14:val="accent3"><w14:alpha w14:val="60000"/></w14:schemeClr></w14:glow>';
  const col = (slot, v) => `<w:color w:val="${v}" w:themeColor="${slot}"/>`;
  const szXml = (n) => `<w:sz w:val="${n}"/><w:szCs w:val="${n}"/>`;

  // Paragraph styles replaced as a whole (lowercase w:name -> default styleId, display name, priority, pPr, rPr).
  function paraSpecs(P, glow) {
    const bdr = (side, slot, v, sz, sp) => `<w:${side} w:val="single" w:sz="${sz}" w:space="${sp}" w:color="${v}" w:themeColor="${slot}"/>`;
    const shd = `<w:shd w:val="clear" w:color="auto" w:fill="${P.lt2}" w:themeFill="background2"/>`;
    const head = (lvl, before, after, extra) => `<w:keepNext/><w:keepLines/>${extra || ""}<w:spacing w:before="${before}" w:after="${after}"/><w:outlineLvl w:val="${lvl}"/>`;
    const G = glow ? GLOW : "";
    return [
      { lname: "title", id: "Title", name: "Title", prio: 10, pPr: '<w:spacing w:after="40" w:line="240" w:lineRule="auto"/><w:contextualSpacing/>',
        rPr: MAJOR + col("text2", P.dk2) + '<w:kern w:val="28"/>' + szXml(56) + G },
      { lname: "subtitle", id: "Subtitle", name: "Subtitle", prio: 11, pPr: '<w:spacing w:after="320"/>', rPr: MINOR + col("accent1", P.accent1) + szXml(26) },
      { lname: "heading 1", id: "Heading1", name: "heading 1", prio: 9, unhide: true,
        pPr: head(0, 360, 140, "<w:pBdr>" + bdr("bottom", "accent3", P.accent3, 12, 4) + "</w:pBdr>"), rPr: MAJOR + col("accent1", P.accent1) + szXml(32) + G },
      { lname: "heading 2", id: "Heading2", name: "heading 2", prio: 9, unhide: true, pPr: head(1, 260, 80), rPr: MAJOR + col("accent2", P.accent2) + szXml(26) + G },
      { lname: "heading 3", id: "Heading3", name: "heading 3", prio: 9, unhide: true, pPr: head(2, 200, 60), rPr: MINOR + "<w:b/><w:bCs/>" + col("text2", P.dk2) + szXml(24) },
      { lname: "heading 4", id: "Heading4", name: "heading 4", prio: 9, semi: true, unhide: true, pPr: head(3, 200, 40),
        rPr: MINOR + "<w:b/><w:bCs/><w:i/><w:iCs/>" + col("accent1", P.accent1) + szXml(22) },
      { lname: "heading 5", id: "Heading5", name: "heading 5", prio: 9, semi: true, unhide: true, pPr: head(4, 160, 40), rPr: MINOR + "<w:b/><w:bCs/>" + col("accent2", P.accent2) + szXml(22) },
      { lname: "heading 6", id: "Heading6", name: "heading 6", prio: 9, semi: true, unhide: true, pPr: head(5, 160, 40), rPr: MINOR + "<w:i/><w:iCs/>" + col("text2", P.dk2) + szXml(22) },
      { lname: "quote", id: "Quote", name: "Quote", prio: 29, pPr: "<w:pBdr>" + bdr("left", "accent3", P.accent3, 36, 12) + "</w:pBdr>" + shd +
        '<w:spacing w:before="200" w:after="200"/><w:ind w:left="300" w:right="300"/>', rPr: MAJOR + col("text2", P.dk2) + szXml(26) },
      { lname: "intense quote", id: "IntenseQuote", name: "Intense Quote", prio: 30, pPr: "<w:pBdr>" + bdr("left", "accent1", P.accent1, 48, 12) + "</w:pBdr>" + shd +
        '<w:spacing w:before="240" w:after="240"/><w:ind w:left="300" w:right="300"/>', rPr: MINOR + "<w:b/><w:bCs/>" + col("accent1", P.accent1) + szXml(24) },
      { lname: "caption", id: "Caption", name: "caption", prio: 35, semi: true, unhide: true, pPr: '<w:spacing w:after="200" w:line="240" w:lineRule="auto"/>',
        rPr: "<w:i/><w:iCs/>" + col("text2", P.dk2) + szXml(18) },
    ];
  }
  function tableGridXml(id, name, basedOn, prio, P) {
    const b = (side) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="${P.accent1}" w:themeColor="accent1"/>`;
    return `<w:style w:type="table" w:styleId="${esc(id)}"><w:name w:val="${esc(name)}"/>` + (basedOn ? `<w:basedOn w:val="${esc(basedOn)}"/>` : "") +
      `<w:uiPriority w:val="${esc(prio)}"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>` +
      `<w:tblPr><w:tblStyleRowBandSize w:val="1"/><w:tblInd w:w="0" w:type="dxa"/><w:tblBorders>${["top", "left", "bottom", "right", "insideH", "insideV"].map(b).join("")}</w:tblBorders>` +
      `<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>` +
      `<w:tblStylePr w:type="band1Horz"><w:tblPr/><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="${P.lt2}" w:themeFill="background2"/></w:tcPr></w:tblStylePr>` +
      `<w:tblStylePr w:type="firstRow"><w:rPr><w:b/><w:bCs/>${col("background1", "FFFFFF")}</w:rPr><w:tblPr/>` +
      `<w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="${P.accent1}" w:themeFill="accent1"/></w:tcPr></w:tblStylePr>` +
      `<w:tblStylePr w:type="lastRow"><w:rPr><w:b/><w:bCs/></w:rPr><w:tblPr/><w:tcPr><w:tcBorders>` +
      `<w:top w:val="double" w:sz="4" w:space="0" w:color="${P.accent1}" w:themeColor="accent1"/></w:tcBorders></w:tcPr></w:tblStylePr></w:style>`;
  }

  // Every style a theme may change, decided from the file alone, so the backup is the same whatever
  // the theme or settings (the core keeps the first backup across re-themes). slots: "type:name" -> styleId.
  const STYLE_SLOTS = [["paragraph", "normal", "Normal"], ...["title", "subtitle", "heading 1", "heading 2", "heading 3", "heading 4", "heading 5", "heading 6",
    "quote", "intense quote", "caption"].map((n) => ["paragraph", n, null]), ["paragraph", "header", "Header"], ["paragraph", "footer", "Footer"],
  ["paragraph", "toc heading", "TOCHeading"], ["character", "hyperlink", "Hyperlink"], ["character", "followedhyperlink", "FollowedHyperlink"],
  ["character", "strong", "Strong"], ["character", "emphasis", "Emphasis"], ["table", "table grid", "TableGrid"]];
  const typeOf = (s) => wAttr(s, "type") || "paragraph";
  const idOf = (s) => wAttr(s, "styleId");
  const styleEls = (doc) => kids(doc.documentElement, NS.w, "style");
  const styleById = (doc, id) => styleEls(doc).find((s) => idOf(s) === id) || null;
  function planStyles(doc) {
    const list = styleEls(doc);
    const nameOf = (s) => wAttr(kid(s, NS.w, "name"), "val").toLowerCase();
    const isDefault = (s) => ["1", "true", "on"].includes(wAttr(s, "default"));
    const find = (type, lname) => list.find((s) => typeOf(s) === type && nameOf(s) === lname) || null;
    const taken = new Set(list.map(idOf));
    const defIds = Object.fromEntries(paraSpecs({}, false).map((sp) => [sp.lname, sp.id]));
    const normal = list.find((s) => typeOf(s) === "paragraph" && isDefault(s)) || find("paragraph", "normal");
    const plan = { slots: {}, links: [], exists: {} };
    for (const [type, lname, defId] of STYLE_SLOTS) {
      const el = lname === "normal" ? normal : find(type, lname);
      if (!el && lname === "toc heading") continue;          // only themed when the document has one
      let id = el ? idOf(el) : defId || defIds[lname];
      if (!el) { const base = id; for (let i = 1; taken.has(id); i++) id = base + "WT" + i; taken.add(id); }
      plan.slots[type + ":" + lname] = id;
      plan.exists[id] = !!el;
      const link = el && wAttr(kid(el, NS.w, "link"), "val");
      const ch = link && list.find((x) => idOf(x) === link);
      if (ch && typeOf(ch) === "character" && !plan.links.includes(link)) plan.links.push(link);
    }
    const dpf = list.find((s) => typeOf(s) === "character" && isDefault(s));
    const tnorm = list.find((s) => typeOf(s) === "table" && isDefault(s));
    plan.normalId = plan.slots["paragraph:normal"];
    plan.dpfId = dpf ? idOf(dpf) : "";
    plan.tableNormalId = tnorm ? idOf(tnorm) : "";
    return plan;
  }
  // The styles backup: each planned style element as it was (null = it didn't exist), and the document defaults.
  function backupStyles(doc, plan) {
    const items = {};
    for (const id of [...Object.values(plan.slots), ...plan.links]) { const el = styleById(doc, id); items[id] = el ? elXml(el) : null; }
    const dd = kid(doc.documentElement, NS.w, "docDefaults");
    return { items, docDefaults: dd ? elXml(dd) : null };
  }

  // Theme the styles part as the settings ask. Returns the table style id new tables should use (or null).
  function applyStyles(doc, P, S, plan, stats) {
    const root = doc.documentElement;
    const byId = (id) => styleById(doc, id);
    const rPrOf = (el) => kid(el, NS.w, "rPr") || setOrdered(el, fragment(doc, "<w:rPr/>")[0], STYLE_ORDER);
    const setRPr = (rPr, xml) => { for (const el of fragment(doc, xml)) setOrdered(rPr, el, RPR_ORDER); };
    const linkOf = (el) => wAttr(kid(el, NS.w, "link"), "val");
    const normalId = plan.normalId;

    // replace (or add) a whole style; returns its linked character style id
    function put(key, build) {
      const id = plan.slots[key];
      const ex = byId(id);
      const el = fragment(doc, build(id, ex))[0];
      if (ex) root.replaceChild(el, ex); else root.appendChild(el);
      stats.styles++;
      return ex ? linkOf(ex) : "";
    }
    // change some run properties of a style in place (or add it from defXml when it's missing)
    function tweak(key, defXml, rPrXml) {
      const id = plan.slots[key];
      const ex = id && byId(id);
      if (!ex) {
        if (!id || !defXml) return "";
        root.appendChild(fragment(doc, defXml(id))[0]);
        stats.styles++;
        return "";
      }
      setRPr(rPrOf(ex), rPrXml);
      stats.styles++;
      return linkOf(ex);
    }
    // keep a linked character style ("Heading 1 Char") in step with its paragraph style
    function step(linkId, rPrXml, replace) {
      const ch = linkId && plan.links.includes(linkId) && byId(linkId);
      if (!ch || typeOf(ch) !== "character") return;
      if (replace) {
        const old = kid(ch, NS.w, "rPr");
        if (old) ch.removeChild(old);
        setOrdered(ch, fragment(doc, `<w:rPr>${rPrXml}</w:rPr>`)[0], STYLE_ORDER);
      } else setRPr(rPrOf(ch), rPrXml);
    }

    if (S.styles) {
      // document defaults: the theme's body font
      let dd = kid(root, NS.w, "docDefaults");
      if (!dd) dd = root.insertBefore(fragment(doc, "<w:docDefaults/>")[0], root.firstChild);
      let rd = kid(dd, NS.w, "rPrDefault");
      if (!rd) rd = dd.insertBefore(fragment(doc, "<w:rPrDefault/>")[0], dd.firstChild);
      setRPr(kid(rd, NS.w, "rPr") || rd.appendChild(fragment(doc, "<w:rPr/>")[0]), MINOR + col("text1", P.dk1));

      // Normal: body font; its text colour (text1) comes from the defaults above. A colour on Normal itself
      // would outrank the table style's header-row colour (white on the accent fill) in Word and LibreOffice.
      tweak("paragraph:normal", (id) => `<w:style w:type="paragraph" w:default="1" w:styleId="${esc(id)}"><w:name w:val="Normal"/><w:qFormat/><w:rPr>${MINOR}</w:rPr></w:style>`, MINOR);
      const nrm = byId(plan.slots["paragraph:normal"]), nrPr = nrm && kid(nrm, NS.w, "rPr");
      if (nrPr) for (const c of kids(nrPr, NS.w, "color")) nrPr.removeChild(c);

      // title, subtitle, headings, quotes, caption
      for (const sp of paraSpecs(P, S.glow)) {
        const link = put("paragraph:" + sp.lname, (id, ex) => {
          const name = (ex && wAttr(kid(ex, NS.w, "name"), "val")) || sp.name;
          const lk = ex ? linkOf(ex) : "";
          return `<w:style w:type="paragraph" w:styleId="${esc(id)}"><w:name w:val="${esc(name)}"/>` +
            `<w:basedOn w:val="${esc(normalId)}"/><w:next w:val="${esc(normalId)}"/>` + (lk ? `<w:link w:val="${esc(lk)}"/>` : "") +
            `<w:uiPriority w:val="${sp.prio}"/>` + (sp.semi ? "<w:semiHidden/>" : "") + (sp.unhide ? "<w:unhideWhenUsed/>" : "") +
            `<w:qFormat/><w:pPr>${sp.pPr}</w:pPr><w:rPr>${sp.rPr}</w:rPr></w:style>`;
        });
        step(link, sp.rPr, true);
      }

      // header / footer text
      const t2 = col("text2", P.dk2);
      for (const ln of ["header", "footer"]) {
        const link = tweak("paragraph:" + ln, (id) => `<w:style w:type="paragraph" w:styleId="${esc(id)}"><w:name w:val="${ln}"/><w:basedOn w:val="${esc(normalId)}"/>` +
          `<w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:pPr><w:tabs><w:tab w:val="center" w:pos="4680"/><w:tab w:val="right" w:pos="9360"/></w:tabs>` +
          `<w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr>${t2}</w:rPr></w:style>`, t2);
        step(link, t2, false);
      }
      // TOC Heading (only when the document has one)
      tweak("paragraph:toc heading", null, MAJOR + col("accent1", P.accent1));

      // links, strong, emphasis
      const charDef = (name, prio, extra, rPr) => (id) => `<w:style w:type="character" w:styleId="${esc(id)}"><w:name w:val="${name}"/>` +
        (plan.dpfId ? `<w:basedOn w:val="${esc(plan.dpfId)}"/>` : "") + `<w:uiPriority w:val="${prio}"/>${extra}<w:rPr>${rPr}</w:rPr></w:style>`;
      const chars = [
        ["hyperlink", "Hyperlink", 99, "<w:unhideWhenUsed/>", col("hyperlink", P.hlink) + '<w:u w:val="single"/>'],
        ["followedhyperlink", "FollowedHyperlink", 99, "<w:semiHidden/><w:unhideWhenUsed/>", col("followedHyperlink", P.folHlink) + '<w:u w:val="single"/>'],
        ["strong", "Strong", 22, "<w:qFormat/>", "<w:b/><w:bCs/>" + col("accent1", P.accent1)],
        ["emphasis", "Emphasis", 20, "<w:qFormat/>", "<w:i/><w:iCs/>" + col("accent2", P.accent2)],
      ];
      for (const [ln, name, prio, extra, rPr] of chars) tweak("character:" + ln, charDef(name, prio, extra, rPr), rPr);
    } else if (S.glow) {
      // glow on the document's own title and top headings, the rest left as it is
      for (const ln of ["title", "heading 1", "heading 2"]) {
        const id = plan.slots["paragraph:" + ln];
        const ex = plan.exists[id] && byId(id);
        if (!ex) continue;
        setRPr(rPrOf(ex), GLOW);
        step(linkOf(ex), GLOW, false);
        stats.styles++;
      }
    }

    if (!S.tables) return null;
    put("table:table grid", (id, ex) => tableGridXml(id, (ex && wAttr(kid(ex, NS.w, "name"), "val")) || "Table Grid",
      plan.tableNormalId, (ex && wAttr(kid(ex, NS.w, "uiPriority"), "val")) || "39", P));
    stats.tables++;
    return plan.slots["table:table grid"];
  }
  function restoreStyles(doc, SB) {
    const root = doc.documentElement;
    for (const [id, xml] of Object.entries(SB.items || {})) {
      const cur = styleById(doc, id);
      if (xml == null) { if (cur) root.removeChild(cur); continue; }
      const el = restoredEl(doc, xml);
      if (cur) root.replaceChild(el, cur); else root.appendChild(el);
    }
    if ("docDefaults" in SB) {
      const cur = kid(root, NS.w, "docDefaults");
      if (SB.docDefaults == null) { if (cur) root.removeChild(cur); }
      else { const el = restoredEl(doc, SB.docDefaults); if (cur) root.replaceChild(el, cur); else root.insertBefore(el, root.firstChild); }
    }
  }

  // ------------------------------------------------------------------ lists
  const LIST_COLOR = "accent1";
  const listLevels = (doc) => {
    const out = [];
    for (const an of kids(doc.documentElement, NS.w, "abstractNum")) {
      for (const lvl of kids(an, NS.w, "lvl")) if (wAttr(kid(lvl, NS.w, "numFmt"), "val") !== "none") out.push({ a: wAttr(an, "abstractNumId"), l: wAttr(lvl, "ilvl"), lvl });
    }
    return out;
  };
  // Every drawn list level's run properties as they were (null = none), whatever the settings.
  const backupLists = (doc) => listLevels(doc).map(({ a, l, lvl }) => { const r = kid(lvl, NS.w, "rPr"); return { a, l, r: r ? elXml(r) : null }; });
  function colourLists(doc, P, stats) {
    for (const { lvl } of listLevels(doc)) {
      const rPr = kid(lvl, NS.w, "rPr") || lvl.appendChild(fragment(doc, "<w:rPr/>")[0]);   // rPr is the last child of w:lvl
      setOrdered(rPr, fragment(doc, col(LIST_COLOR, P[LIST_COLOR]))[0], RPR_ORDER);
      stats.listLevels++;
    }
  }
  // A level is put back when it holds the original, or the original with only our theme colour added.
  // A numbering part another editor renumbered never gets the wrong list's glyph fonts: there only our colour goes.
  function restoreLists(doc, levels) {
    const abs = kids(doc.documentElement, NS.w, "abstractNum");
    const withoutColor = (el) => { const c = el.cloneNode(true); for (const x of kids(c, NS.w, "color")) c.removeChild(x); return elXml(c); };
    let n = 0;
    for (const { a, l, r } of levels || []) {
      const an = abs.find((x) => wAttr(x, "abstractNumId") === a);
      const lvl = an && kids(an, NS.w, "lvl").find((x) => wAttr(x, "ilvl") === l);
      const cur = lvl && kid(lvl, NS.w, "rPr");
      if (!cur) continue;
      const c = kid(cur, NS.w, "color");
      const ours = !!c && wAttr(c, "themeColor") === LIST_COLOR;
      const orig = r != null ? restoredEl(doc, r) : fragment(doc, "<w:rPr/>")[0];
      if (elXml(orig) === elXml(cur)) continue;                        // untouched
      if (ours && withoutColor(orig) === withoutColor(cur)) {
        lvl.removeChild(cur);
        if (r != null) lvl.appendChild(orig);
        n++;
      } else if (ours) {
        // rewritten since: just take our colour back out (or put the level's own colour back)
        const own = kid(orig, NS.w, "color");
        if (own) cur.replaceChild(own, c); else cur.removeChild(c);
        if (r == null && !Array.from(cur.childNodes).some((x) => x.nodeType === 1)) lvl.removeChild(cur);
        n++;
      }
    }
    return n;
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

  const PANEL_DESCR = "Decorative page panel (Wild Theme)";
  const markerIn = (descr) => /(?:^|\n)Wild Theme:\s*([^;\n]+)(?:;\s*name=([^;\n]*))?/.exec(descr || "");
  // Is this drawing one of ours (current or v1 builder)? Another editor may rename our pictures but keeps their alt text.
  function isOurDocPr(docPr) {
    const name = docPr.getAttribute("name") || "";
    const descr = docPr.getAttribute("descr") || "";
    return name.startsWith("WildTheme") || (name === "Theme backdrop" && /^Decorative animated/.test(descr)) ||
      descr === PANEL_DESCR || descr.endsWith("\n" + PANEL_DESCR) || !!markerIn(descr);
  }
  // Which scene picture is it: "backdrop", "panel", or null (the 1px marker, or not ours).
  function sceneKind(docPr) {
    const name = docPr.getAttribute("name") || "";
    if (name === "WildThemeBackdrop") return "backdrop";
    if (name === "WildThemePanel") return "panel";
    if (name.startsWith("WildTheme") || !isOurDocPr(docPr)) return null;
    const descr = docPr.getAttribute("descr") || "";
    if (descr === PANEL_DESCR || descr.endsWith("\n" + PANEL_DESCR)) return "panel";
    const ext = kid(docPr.parentNode, NS.wp, "extent");
    return ext && parseInt(ext.getAttribute("cx"), 10) > 9525 * 4 ? "backdrop" : null;
  }
  const isOurTag = (v) => /^WildTheme/.test(v || "");
  const isOurSdt = (sdt) => { const pr = kid(sdt, NS.w, "sdtPr"); return !!pr && isOurTag(wAttr(kid(pr, NS.w, "tag"), "val")); };
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
  // drawings from older versions. Returns the image relationship ids that were freed (.removed = how many things went).
  function removeOurs(doc) {
    const freed = [];
    let removed = 0;
    for (const sdt of all(doc, NS.w, "sdt")) {
      if (!isOurSdt(sdt) || !sdt.parentNode) continue;
      freed.push(...blipIds(sdt));
      const p = sdt.parentNode;
      p.removeChild(sdt);
      removed++;
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
      removed++;
      if (run && !Array.from(run.childNodes).some((c) => c.nodeType === 1 && !(c.namespaceURI === NS.w && c.localName === "rPr"))) {
        const p = run.parentNode;
        p.removeChild(run);
        dropEmptyOwnParagraph(p);
      }
    }
    freed.removed = removed;
    return freed;
  }
  const removeOurDrawings = removeOurs;
  // Declare a namespace prefix on a part's root.
  function ensureNs(doc, prefix, uri) {
    const root = doc.documentElement;
    if (!root.getAttribute("xmlns:" + prefix)) root.setAttributeNS(NS.xmlns, "xmlns:" + prefix, uri);
  }
  // Declare a namespace and list it as ignorable (mc:Ignorable), so older Word skips what it doesn't know.
  function ensureIgnorable(doc, prefix, uri) {
    const root = doc.documentElement;
    ensureNs(doc, prefix, uri);
    ensureNs(doc, "mc", NS.mc);
    const ign = root.getAttributeNS(NS.mc, "Ignorable") || root.getAttribute("mc:Ignorable") || "";
    if (!ign.split(/\s+/).includes(prefix)) root.setAttributeNS(NS.mc, "mc:Ignorable", (ign + " " + prefix).trim());
  }
  const ensureW15 = (doc) => ensureIgnorable(doc, "w15", NS.w15);
  // Root attributes we may add; their original values go in the backup.
  const ROOT_ATTRS = { "xmlns:w14": NS.xmlns, "xmlns:w15": NS.xmlns, "xmlns:mc": NS.xmlns, "xmlns:v": NS.xmlns, "xmlns:o": NS.xmlns, "mc:Ignorable": NS.mc };
  function noteRoot(B, part, doc) {
    const r = (B.roots[part] = B.roots[part] || {});
    const root = doc.documentElement;
    for (const name of Object.keys(ROOT_ATTRS)) if (!(name in r)) r[name] = root.hasAttribute(name) ? root.getAttribute(name) : null;
  }
  // Is a namespace still used anywhere in the part (elements or attributes)?
  function nsInUse(doc, uri) {
    for (const n of Array.from(doc.getElementsByTagNameNS("*", "*"))) {
      if (n === doc.documentElement) continue;      // the root's own attributes are what we're restoring
      if (n.namespaceURI === uri) return true;
      for (const a of Array.from(n.attributes || [])) if (a.namespaceURI === uri) return true;
    }
    return false;
  }
  function restoreRoot(doc, saved) {
    const root = doc.documentElement;
    const used = new Set();
    for (const [name, v] of Object.entries(saved)) {
      if (!name.startsWith("xmlns:")) continue;
      const cur = root.getAttribute(name);
      if (v == null && cur && nsInUse(doc, cur)) used.add(name.slice(6));
    }
    const keepIgnorable = [...used].some((p) => p !== "mc");
    let changed = false;
    for (const [name, v] of Object.entries(saved)) {
      if (name.startsWith("xmlns:") && used.has(name.slice(6))) continue;
      if (name === "mc:Ignorable" && keepIgnorable) continue;   // something still needs its ignorable prefix
      const cur = root.hasAttribute(name) ? root.getAttribute(name) : null;
      if (cur === v) continue;
      if (v == null) root.removeAttribute(name);
      else root.setAttributeNS(ROOT_ATTRS[name], name, v);
      changed = true;
    }
    return changed;
  }
  // A locked inline content control around our run: it can't be deleted or edited in Word.
  function lockedSdt(tag, id, innerXml, hidden) {
    return `<w:sdt><w:sdtPr><w:alias w:val="Wild Theme"/><w:tag w:val="${tag}"/><w:id w:val="${id}"/><w:lock w:val="sdtContentLocked"/>` +
      (hidden ? `<w15:appearance w15:val="hidden"/>` : "") + `</w:sdtPr><w:sdtContent>${innerXml}</w:sdtContent></w:sdt>`;
  }
  // Content-control ids: a fixed sequence per theme (deterministic), skipping ids already in use.
  function sdtIdPicker(docs, seed) {
    const used = new Set();
    for (const d of docs) for (const pr of all(d, NS.w, "sdtPr")) { const v = wAttr(kid(pr, NS.w, "id"), "val"); if (v) used.add(v); }
    let k = 0;
    return () => {
      let v;
      do { v = String(((parseInt(seed, 16) + (k++) * 2654435761) >>> 0) % 2000000000 - 1000000000); } while (used.has(v));
      used.add(v);
      return v;
    };
  }
  // Old visible tag lines from v1 documents ("✦ Wild Theme: …") are removed.
  function removeTagLines(doc) {
    const body = all(doc, NS.w, "body")[0];
    if (!body) return 0;
    let n = 0;
    for (const p of kids(body, NS.w, "p")) {
      const text = all(p, NS.w, "t").map((t) => t.textContent).join("");
      if (/^\s*(✦\s*)?Wild Theme:\s*.{0,60}$/.test(text) && !all(p, NS.w, "drawing").length) { body.removeChild(p); n++; }
    }
    return n;
  }
  function maxDocPrId(docs) {
    let max = 0;
    for (const d of docs) for (const el of all(d, NS.wp, "docPr")) max = Math.max(max, parseInt(el.getAttribute("id"), 10) || 0);
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

  // ------------------------------------------------------------------ settings part
  // Put an element in its schema position (replacing any element of the same name).
  function putSetting(doc, el) {
    const root = doc.documentElement;
    for (const old of kids(root, NS.w, el.localName)) root.removeChild(old);
    const rank = (n) => (n.namespaceURI === NS.w && SETTINGS_ORDER.includes(n.localName) ? SETTINGS_ORDER.indexOf(n.localName) : SETTINGS_ORDER.length);
    const mine = rank(el);
    let before = null;
    for (let n = root.firstChild; n; n = n.nextSibling) if (n.nodeType === 1 && rank(n) > mine) { before = n; break; }
    root.insertBefore(el, before);
  }
  function removeSetting(doc, name) {
    let n = 0;
    for (const el of kids(doc.documentElement, NS.w, name)) { doc.documentElement.removeChild(el); n++; }
    return n;
  }

  // ------------------------------------------------------------------ page background
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
  // w:color is what apps that show a page colour but not its picture use: the theme's light panel tint.
  function pageBackgroundXml(rId, color) {
    return `<w:background w:color="${color}"><v:background id="_x0000_s1025" o:bwmode="white" o:targetscreensize="1024,768">` +
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
    for (const name of new Set([...Object.keys(pkg.zip.files), ...pkg.cache.keys()])) {
      if (!/_rels\/[^/]+\.rels$/.test(name)) continue;
      const owner = name.replace(/_rels\/([^/]+)\.rels$/, "$1");
      const f = pkg.zip.file(name);
      const rd = pkg.cache.get(name) || (f && parseXml(await f.async("string")));
      if (!rd) continue;
      for (const el of Array.from(rd.getElementsByTagName("Relationship"))) {
        if (el.getAttribute("TargetMode") !== "External") stillUsed.add(resolveTarget(owner, el.getAttribute("Target") || ""));
      }
    }
    for (const m of freedMedia) if (!stillUsed.has(m)) pkg.remove(m);
    for (const name of Object.keys(pkg.zip.files)) if (/^word\/media\/wildtheme-/.test(name) && !stillUsed.has(name)) pkg.remove(name);
  }
  const isOurMedia = (target) => /(^|\/)media\/wildtheme-/.test(target);
  // Remove our pictures (and their relationships) from the main part and every header part,
  // plus our page background (someone else's is left alone). Returns whether a page background was removed.
  async function cleanParts(pkg, main, mainRels, { tagLines = false } = {}) {
    const freedMedia = new Set();
    let removedBackground = false;
    for (const part of [main, ...mainRels.byType(REL.header).map((r) => resolveTarget(main, r.target))]) {
      const pd = await pkg.xml(part);
      if (!pd) continue;
      const freed = removeOurs(pd);
      let changed = freed.removed > 0;
      if (part === main) {
        if (tagLines && removeTagLines(pd)) changed = true;
        const bg = backgroundOf(pd);
        if (bg && isOurBackground(bg, mainRels)) {
          freed.push(...relIdsIn(bg));
          pd.documentElement.removeChild(bg);
          removedBackground = changed = true;
        }
      }
      if (changed) pkg.touch(part);
      const rels = await pkg.rels(part);
      for (const r of rels.list()) {
        if (r.type !== REL.image) continue;
        if (isOurMedia(r.target) || freed.includes(r.id)) { freedMedia.add(resolveTarget(part, r.target)); rels.remove(r.id); }
      }
    }
    await pruneMedia(pkg, freedMedia);
    return removedBackground;
  }
  // Any sign of a Wild Theme in the file (any version)?
  async function looksOurs(pkg, main, doc, mainRels) {
    if (await readCustomProp(pkg, "WildTheme")) return true;
    const docs = [doc];
    for (const r of mainRels.byType(REL.header)) { const d = await pkg.xml(resolveTarget(main, r.target)); if (d) docs.push(d); }
    for (const d of docs) if (all(d, NS.wp, "docPr").some(isOurDocPr) || all(d, NS.w, "sdt").some(isOurSdt)) return true;
    if (isOurBackground(backgroundOf(doc), mainRels)) return true;
    const tRel = mainRels.byType(REL.theme)[0];
    const td = tRel && (await pkg.xml(resolveTarget(main, tRel.target)));
    if (td && /^Wild Theme /.test(td.documentElement.getAttribute("name") || "")) return true;
    const cp = await pkg.xml("docProps/core.xml");
    return !!cp && Array.from(cp.getElementsByTagNameNS(NS.coreProps, "keywords")).some((k) => /^\s*Wild Theme:/.test(k.textContent));
  }
  // The parts whose direct formatting "match old colours" converts.
  function contentParts(pkg, main, mainRels) {
    const out = [main];
    for (const t of [REL.header, REL.footer, REL.footnotes, REL.endnotes]) {
      for (const r of mainRels.byType(t)) { const p = resolveTarget(main, r.target); if (!out.includes(p) && pkg.zip.file(p)) out.push(p); }
    }
    for (const name of Object.keys(pkg.zip.files).sort()) if (/^word\/charts\/[^/]+\.xml$/.test(name)) out.push(name);
    return out;
  }
  // A part from the main part's relationships, created (and recorded) when it's missing.
  async function partFor(pkg, ct, main, mainRels, type, defPath, contentType, emptyXml) {
    const rel = mainRels.byType(type)[0];
    if (rel) return { path: resolveTarget(main, rel.target), created: false };
    pkg.setXml(defPath, parseXml(emptyXml));
    mainRels.add(type, relativeTarget(main, defPath));
    ct.ensureOverride(defPath, contentType);
    return { path: defPath, created: true };
  }

  const EMPTY_STYLES = `<w:styles xmlns:w="${NS.w}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style></w:styles>`;
  const OUR_SETTINGS = ["styleLockTheme", "styleLockQFSet", "displayBackgroundShape", "defaultTableStyle"];
  // The backup: everything a theme may change, as it is in the (clean) file. It depends on the file alone,
  // not on the theme or the settings, because the core keeps the first backup across re-themes.
  async function snapshot(pkg, main, doc, mainRels) {
    const relPath = (type) => { const r = mainRels.byType(type)[0]; return r ? resolveTarget(main, r.target) : null; };
    const B = { v: 4, kind: "docx" };
    // theme part
    const tp = relPath(REL.theme), tf = tp && pkg.zip.file(tp);
    B.theme = { path: tp || "word/theme/theme1.xml", xml: tf ? await tf.async("string") : null, rel: !!tp };
    // styles: the ones a theme may change (null = it didn't exist), and the document defaults
    const sp = relPath(REL.styles), sd = sp && (await pkg.xml(sp));
    const plan = sd ? planStyles(sd) : null;
    B.styles = sd ? { path: sp, created: false, ...backupStyles(sd, plan) } : { path: sp || "word/styles.xml", created: !sp, items: {} };
    // settings we may add
    const setp = relPath(REL.settings), setd = setp && (await pkg.xml(setp));
    B.settings = { path: setp || "word/settings.xml", created: !setp, items: {} };
    for (const name of OUR_SETTINGS) { const ex = setd && kid(setd.documentElement, NS.w, name); B.settings.items[name] = ex ? elXml(ex) : null; }
    // list levels
    const np = relPath(REL.numbering), nd = np && (await pkg.xml(np));
    B.numbering = nd ? { path: np, levels: backupLists(nd) } : null;
    // someone else's page background (the scene replaces it)
    const bg = backgroundOf(doc);
    B.background = bg && !isOurBackground(bg, mainRels) ? elXml(bg) : null;
    // header layout: the header relationships, and which section shows which header
    B.headers = { rels: {}, refs: [] };
    for (const r of mainRels.byType(REL.header)) B.headers.rels[r.id] = resolveTarget(main, r.target);
    const secs = sections(doc);
    secs.forEach((sect, si) => { for (const ref of kids(sect, NS.w, "headerReference")) B.headers.refs.push([si, wAttr(ref, "type") || "default", ridOf(ref)]); });
    B.sectPr = !secs.length;
    // direct formatting in the original theme's colours and fonts (or Office's stock palettes)
    B.direct = {};
    const maps = directMaps(B.theme.xml ? themeInfo(B.theme.xml) : null);
    for (const part of contentParts(pkg, main, mainRels)) {
      const d = await pkg.xml(part);
      const recs = d ? findDirect(d, maps) : [];
      if (recs.length) B.direct[part] = recs;
    }
    // root attributes of the parts we may declare namespaces on
    B.roots = {};
    for (const part of [main, ...Object.values(B.headers.rels), ...(sd ? [sp] : [])]) { const d = await pkg.xml(part); if (d) noteRoot(B, part, d); }
    // content types and relationship parts as they are
    const ctDoc = await pkg.xml("[Content_Types].xml");
    B.ctDefaults = Array.from(ctDoc.documentElement.getElementsByTagName("Default")).map((d) => (d.getAttribute("Extension") || "").toLowerCase()).sort();
    B.rels = Object.keys(pkg.zip.files).filter((n) => /\.rels$/.test(n)).sort();
    return { B, plan };
  }

  // ------------------------------------------------------------------ public: analyze
  async function analyze(bytes) {
    const pkg = await Pkg.open(bytes);
    const main = await pkg.mainPart();
    const doc = await pkg.xml(main);
    const mainRels = await pkg.rels(main);
    const secs = await sectionInfos(pkg, main, doc, mainRels, await evenOddOf(pkg, main, mainRels));
    let title = null;
    const coreDoc = await pkg.xml("docProps/core.xml");
    if (coreDoc) {
      const t = coreDoc.getElementsByTagNameNS("http://purl.org/dc/elements/1.1/", "title")[0];
      title = t ? t.textContent.trim() || null : null;
    }
    let existingTheme = await readCustomProp(pkg, "WildTheme");
    if (!existingTheme) {
      for (const docPr of all(doc, NS.wp, "docPr")) {
        const m = markerIn(docPr.getAttribute("descr"));
        if (m) { existingTheme = m[1].trim(); break; }
      }
    }
    return { sections: secs, title, existingTheme: existingTheme || null };
  }

  // ------------------------------------------------------------------ public: probe
  // What the core fingerprints: the theme part, how many of our scene pictures are present, the scene's
  // media bytes, and the theme named by the file's custom property or picture alt text.
  async function probe(bytes) {
    const pkg = await Pkg.open(bytes);
    const main = await pkg.mainPart();
    const doc = await pkg.xml(main);
    const mainRels = await pkg.rels(main);
    const out = { themes: [], pictures: 0, media: [], marker: null, markerName: null };
    const tRel = mainRels.byType(REL.theme)[0];
    const tf = tRel && pkg.zip.file(resolveTarget(main, tRel.target));
    if (tf) out.themes.push(await tf.async("string"));
    const media = new Set();
    let descr = null;
    for (const part of [main, ...mainRels.byType(REL.header).map((r) => resolveTarget(main, r.target))]) {
      const d = part === main ? doc : await pkg.xml(part);
      if (!d) continue;
      const rels = await pkg.rels(part);
      for (const docPr of all(d, NS.wp, "docPr")) {
        if (!descr && markerIn(docPr.getAttribute("descr"))) descr = docPr.getAttribute("descr");
        const kind = sceneKind(docPr);
        if (!kind) continue;
        out.pictures++;
        if (kind !== "backdrop") continue;
        const drawing = ancestor(docPr, NS.w, "drawing");
        for (const id of drawing ? blipIds(drawing) : []) { const r = rels.byId(id); if (r) media.add(resolveTarget(part, r.target)); }
      }
    }
    const bg = backgroundOf(doc);
    if (bg && isOurBackground(bg, mainRels)) {
      out.pictures++;
      for (const id of relIdsIn(bg)) { const r = mainRels.byId(id); if (r && r.type === REL.image) media.add(resolveTarget(main, r.target)); }
    }
    for (const m of Array.from(media).sort()) { const f = pkg.zip.file(m); if (f) out.media.push(await f.async("uint8array")); }
    out.marker = await readCustomProp(pkg, "WildTheme");
    out.markerName = await readCustomProp(pkg, "WildThemeName");
    if (descr) {
      const m = markerIn(descr);
      if (m && !out.marker) out.marker = m[1].trim();
      if (m && !out.markerName && m[2]) out.markerName = m[2].trim();
    }
    if (out.marker && !out.markerName) out.markerName = out.marker;
    return out;
  }

  // ------------------------------------------------------------------ public: stamp
  // settings: WTCore.settingsFor("docx", ...) -> { backdrop, panel, styles, tables, lists, direct, lock, glow }.
  // Returns the new file with .backup (everything needed to put the original back), .stats and .failures.
  async function stamp(bytes, theme, assets, settings) {
    if (!theme || !theme.id) throw new Error("No theme chosen.");
    const S = settingsOf(settings);
    const P = paletteOf(theme);
    assets = assets || {};
    const failures = [];
    const stats = { sections: 0, headersStamped: 0, styles: 0, tables: 0, listLevels: 0, direct: { colors: 0, fills: 0, fonts: 0, drawing: 0 } };
    const pkg = await Pkg.open(bytes);
    const ct = await pkg.contentTypes();
    const main = await pkg.mainPart();
    const doc = await pkg.xml(main);
    const mainRels = await pkg.rels(main);
    const descr = markerDescr(theme);
    const backdrop = assets.backdrop && assets.backdrop.bytes ? assets.backdrop : null;
    const scene = !!(S.backdrop && backdrop);
    if (S.backdrop && !backdrop) failures.push("No scene picture was supplied, so none was placed.");

    // ---- anything an earlier Wild Theme left behind (stamp is normally given a clean file)
    await cleanParts(pkg, main, mainRels, { tagLines: await looksOurs(pkg, main, doc, mainRels) });

    // ---- the backup: worked out from the file alone, before anything changes
    const { B, plan } = await snapshot(pkg, main, doc, mainRels);

    // ---- direct formatting in the old theme's colours/fonts -> theme references
    if (S.direct) {
      for (const [part, recs] of Object.entries(B.direct)) { applyDirect(await pkg.xml(part), recs, P, stats.direct); pkg.touch(part); }
    }

    // ---- theme colours + fonts
    if (!B.theme.rel) mainRels.add(REL.theme, relativeTarget(main, B.theme.path));
    pkg.zip.file(B.theme.path, themeXml(theme, P), { createFolders: false });
    pkg.cache.delete(B.theme.path);
    ct.ensureOverride(B.theme.path, CT.theme);

    // ---- styles: text styles, glow, table style
    let tableStyle = null;
    if (S.styles || S.tables || S.glow) {
      const sp = await partFor(pkg, ct, main, mainRels, REL.styles, B.styles.path, CT.styles, EMPTY_STYLES);
      const styles = await pkg.xml(sp.path);
      if (styles) {
        if (S.glow) ensureIgnorable(styles, "w14", NS.w14);
        tableStyle = applyStyles(styles, P, S, plan || planStyles(styles), stats);
        pkg.touch(sp.path);
      } else failures.push("The document's styles part is missing, so its styles could not be themed.");
    }

    // ---- bullet and number colours
    if (S.lists && B.numbering) {
      const nd = await pkg.xml(B.numbering.path);
      colourLists(nd, P, stats);
      if (stats.listLevels) pkg.touch(B.numbering.path);
    }

    // ---- settings: theme / style-set locks, show the page background, default table style
    const want = [];
    if (S.lock) want.push(["styleLockTheme", null], ["styleLockQFSet", null]);
    if (scene) want.push(["displayBackgroundShape", null]);
    if (tableStyle) want.push(["defaultTableStyle", tableStyle]);
    if (want.length) {
      const sp = await partFor(pkg, ct, main, mainRels, REL.settings, B.settings.path, CT.settings, `<w:settings xmlns:w="${NS.w}"/>`);
      const sd = await pkg.xml(sp.path);
      for (const [name, val] of want) putSetting(sd, fragment(sd, val == null ? `<w:${name}/>` : `<w:${name} w:val="${esc(val)}"/>`)[0]);
      pkg.touch(sp.path);
    }

    // ---- which header types are in play; page geometry per section
    const evenOdd = await evenOddOf(pkg, main, mainRels);
    const infos = await sectionInfos(pkg, main, doc, mainRels, evenOdd);
    const secEls = sections(doc);
    if (!secEls.length) {
      const body = all(doc, NS.w, "body")[0];
      const s = fragment(doc, `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`)[0];
      body.appendChild(s);
      secEls.push(s);
    }
    stats.sections = secEls.length;

    // ---- media (named by content, so the same theme gives the same part names)
    const ext = backdrop ? (backdrop.ext || "png").toLowerCase().replace("jpg", "jpeg") : "png";
    if (scene && !IMG_CT[ext]) throw new Error("Backdrop must be GIF, PNG or JPEG.");
    const media = { backdrop: scene ? `word/media/wildtheme-backdrop-${fnv(backdrop.bytes)}.${ext}` : null, marker: "word/media/wildtheme-marker.png" };
    if (scene) { pkg.putBinary(media.backdrop, backdrop.bytes); ct.ensureDefault(ext, IMG_CT[ext]); }
    pkg.putBinary(media.marker, b64ToBytes(MARKER_PNG_B64));
    ct.ensureDefault("png", IMG_CT.png);
    // assets.panel: { bytes } used everywhere, or a function (section, rectTwips) -> PNG bytes
    // called once per distinct page shape so landscape sections get an undistorted panel.
    const panelFn = typeof assets.panel === "function" ? assets.panel : null;
    const panelMedia = new Map();     // rect key -> media path
    async function panelPath(geom) {
      const key = panelFn ? (({ w, h }) => `${w}x${h}`)(panelRect(geom)) : "all";
      if (!panelMedia.has(key)) {
        const bytes = panelFn ? await panelFn(geom, panelRect(geom)) : assets.panel && assets.panel.bytes;
        let path = null;
        if (bytes) { path = `word/media/wildtheme-panel-${fnv(bytes)}.png`; pkg.putBinary(path, bytes); }
        panelMedia.set(key, path);
      }
      return panelMedia.get(key);
    }
    const havePanel = !!(panelFn || (assets.panel && assets.panel.bytes));
    if (scene && !havePanel) failures.push("No reading panel picture was supplied, so the text has no panel behind it.");

    const headerDocs = [];
    for (const r of mainRels.byType(REL.header)) { const d = await pkg.xml(resolveTarget(main, r.target)); if (d) headerDocs.push(d); }
    let nextId = maxDocPrId([doc, ...headerDocs]) + 1;
    const nextSdtId = sdtIdPicker([doc, ...headerDocs], fnv(theme.id));
    const pictureRun = (anchor) => `<w:r w:rsidR="${OUR_RSID}"><w:rPr><w:noProof/></w:rPr><w:drawing>${anchor}</w:drawing></w:r>`;
    const panelAnchor = (geom, rId, z) => {
      const pr = panelRect(geom);
      return anchorXml({ id: nextId++, name: "WildThemePanel", descr: "Decorative page panel (Wild Theme)", title: "",
        rId, x: pr.x * TWIP_EMU, y: pr.y * TWIP_EMU, cx: pr.w * TWIP_EMU, cy: pr.h * TWIP_EMU, z });
    };

    let bgRid = null;
    if (scene) {
      // ---- the scene on every page: the document's page background (full colour, never dimmed).
      //      Someone else's page background is backed up (its picture stays in the file for the restore).
      const old = backgroundOf(doc);
      if (old) doc.documentElement.removeChild(old);         // someone else's (it's in the backup)
      bgRid = mainRels.add(REL.image, relativeTarget(main, media.backdrop));
      ensureNs(doc, "v", NS.v); ensureNs(doc, "o", NS.o);
      doc.documentElement.insertBefore(fragment(doc, pageBackgroundXml(bgRid, P.panel))[0], doc.documentElement.firstChild);

      // ---- headers: the light panel, section by section (Word dims header content while you
      //      edit the body, which is fine for a pale panel but would wash out the scene)
      const stampedGeom = new Map();          // header part path -> geometry it was stamped for
      const inherited = {};                   // header type -> part path currently in force
      const freshHeaderPath = () => { let n = 1; while (pkg.has(`word/header${n}.xml`)) n++; return `word/header${n}.xml`; };
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
        stats.headersStamped++;
      }
      async function newHeader(sourcePath) {
        const path = freshHeaderPath();
        if (sourcePath) {
          const copy = parseXml(serialize(await pkg.xml(sourcePath)));
          removeOurDrawings(copy);
          pkg.setXml(path, copy);
          const srcRels = pkg.zip.file(relsPathOf(sourcePath));
          if (srcRels) {
            const rd = parseXml(await srcRels.async("string"));
            for (const el of Array.from(rd.getElementsByTagName("Relationship"))) if (isOurMedia(el.getAttribute("Target") || "")) el.parentNode.removeChild(el);
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
      for (let si = 0; si < (havePanel ? secEls.length : 0); si++) {
        const sectPr = secEls[si];
        const geom = infos[si] || sectionGeometry(sectPr);
        const types = ["default"];
        if (onOff(kid(sectPr, NS.w, "titlePg"))) types.push("first");
        if (evenOdd) types.push("even");
        for (const type of types) {
          const ref = kids(sectPr, NS.w, "headerReference").find((h) => (wAttr(h, "type") || "default") === type);
          if (ref) {
            const rel = mainRels.byId(ridOf(ref));
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
    }

    // ---- page 1, in the body's first paragraph (locked): the marker the extension finds in Word for
    //      the web (always, it's an invisible 1px picture), the scene and its panel (body behind-text
    //      pictures sit above the header layer, so page 1 needs its own panel).
    const body = all(doc, NS.w, "body")[0];
    const first = infos[0] || sectionGeometry(secEls[0]);
    const markerRid = mainRels.add(REL.image, relativeTarget(main, media.marker));
    let page1 = pictureRun(anchorXml({ id: nextId++, name: "WildThemeMarker", descr, title: "Wild Theme", rId: markerRid, x: 0, y: 0, cx: 9525, cy: 9525, z: 3 }));
    if (scene) {
      page1 += pictureRun(anchorXml({ id: nextId++, name: "WildThemeBackdrop", descr, title: "Wild Theme", rId: bgRid, x: 0, y: 0,
        cx: first.pageW * TWIP_EMU, cy: first.pageH * TWIP_EMU, z: 1 }));
      const page1Panel = havePanel ? await panelPath(first) : null;
      if (page1Panel) page1 += pictureRun(panelAnchor(first, mainRels.add(REL.image, relativeTarget(main, page1Panel)), 2));
      else if (havePanel) failures.push("The reading panel picture for page 1 could not be made.");
    }
    if (HIDE_SDT) ensureW15(doc);
    insertLeadRun(doc, body, lockedSdt("WildThemeMarker", nextSdtId(), page1, HIDE_SDT));
    pkg.touch(main);

    // ---- metadata
    await setCustomProps(pkg, { WildTheme: theme.id, WildThemeName: theme.name || theme.id, WildThemeVersion: "4" });

    const out = await pkg.save();
    out.backup = B;
    out.stats = stats;
    out.failures = failures;
    return out;
  }

  // ------------------------------------------------------------------ public: unstamp
  // Removes everything Wild Theme added (pictures, locked controls, marker, metadata). With the backup
  // stamp() returned, the original theme part, styles, settings, list colours, page background,
  // headers and direct formatting come back exactly. Without one (files from older versions) the
  // theme colours/fonts and styles stay as they are, since the originals aren't known.
  async function unstamp(bytes, backup) {
    const B = backup && typeof backup === "object" && backup.v >= 4 ? backup : null;
    const pkg = await Pkg.open(bytes);
    const ct = await pkg.contentTypes();
    const main = await pkg.mainPart();
    const doc = await pkg.xml(main);
    const mainRels = await pkg.rels(main);
    const ours = await looksOurs(pkg, main, doc, mainRels);
    if (!ours && !B) return toU8(bytes);           // no theme here: leave the file exactly as it is
    const removedBackground = await cleanParts(pkg, main, mainRels, { tagLines: ours });

    if (B) await restore(pkg, ct, main, doc, mainRels, B);

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
      for (const ref of all(doc, NS.w, "headerReference")) if (ridOf(ref) === r.id) ref.parentNode.removeChild(ref);
      mainRels.remove(r.id);
      pkg.touch(main);
      if (!stillShared) { pkg.remove(path); pkg.remove(relsPathOf(path)); ct.removeOverride(path); }
    }

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
    const coreDoc = await pkg.xml("docProps/core.xml");
    if (coreDoc) {
      for (const k of Array.from(coreDoc.getElementsByTagNameNS(NS.coreProps, "keywords"))) {
        if (/^\s*Wild Theme:/.test(k.textContent)) { k.parentNode.removeChild(k); pkg.touch("docProps/core.xml"); }
      }
    }

    if (!B) {
      // older versions: theme locks, the page-background switch, our name on the colour scheme
      const sRel = mainRels.byType(REL.settings)[0];
      const sd = sRel && (await pkg.xml(resolveTarget(main, sRel.target)));
      if (sd) {
        let n = removeSetting(sd, "styleLockTheme") + removeSetting(sd, "styleLockQFSet");
        if (removedBackground && !backgroundOf(doc)) n += removeSetting(sd, "displayBackgroundShape");
        if (n) pkg.touch(resolveTarget(main, sRel.target));
      }
      const tRel = mainRels.byType(REL.theme)[0];
      const tp = tRel && resolveTarget(main, tRel.target);
      const td = tp && (await pkg.xml(tp));
      if (td) {
        const n = td.documentElement.getAttribute("name") || "";
        if (/^Wild Theme /.test(n)) { td.documentElement.setAttribute("name", n.replace(/^Wild Theme /, "")); pkg.touch(tp); }
      }
    }
    return pkg.save();
  }

  // Put back what stamp() replaced, from its backup.
  async function restore(pkg, ct, main, doc, mainRels, B) {
    const relPath = (type) => { const r = mainRels.byType(type)[0]; return r ? resolveTarget(main, r.target) : null; };
    const dropPart = (type, path) => {
      pkg.remove(path); ct.removeOverride(path);
      for (const r of mainRels.byType(type)) if (resolveTarget(main, r.target) === path) mainRels.remove(r.id);
    };

    // headers we added, or split off for a differently-sized section: references and relationships with
    // our ids go back to the original layout (only while the original header relationships are intact:
    // an editor that renumbers them has made its own header parts, which are left as they are)
    const HL = B.headers && B.headers.rels ? B.headers : null;
    const intact = HL && Object.entries(HL.rels).every(([id, t]) => { const r = mainRels.byId(id); return !!r && r.type === REL.header && resolveTarget(main, r.target) === t; });
    if (intact) {
      const ours = (id) => /^rIdWT/.test(id || "") && !(id in HL.rels);
      sections(doc).forEach((sect, si) => {
        for (const ref of kids(sect, NS.w, "headerReference")) {
          if (!ours(ridOf(ref))) continue;
          const was = HL.refs.find(([s, t]) => s === si && t === (wAttr(ref, "type") || "default"));
          if (was) ref.setAttributeNS(NS.r, "r:id", was[2]); else sect.removeChild(ref);
        }
      });
      const used = new Set(all(doc, NS.w, "headerReference").map(ridOf));
      for (const r of mainRels.byType(REL.header)) {
        if (!ours(r.id) || used.has(r.id)) continue;
        mainRels.remove(r.id);
        const p = resolveTarget(main, r.target);
        if (!Object.values(HL.rels).includes(p) && !mainRels.byType(REL.header).some((o) => resolveTarget(main, o.target) === p)) {
          pkg.remove(p); pkg.remove(relsPathOf(p)); ct.removeOverride(p);
        }
      }
    }
    if (B.sectPr) {
      const body = all(doc, NS.w, "body")[0];
      const last = body && kids(body, NS.w, "sectPr").pop();
      if (last) body.removeChild(last);
    }
    // someone else's page background we replaced
    if (B.background && !backgroundOf(doc)) {
      const el = restoredEl(doc, B.background);
      if (relIdsIn(el).some((id) => !mainRels.byId(id))) for (const vb of all(el, NS.v, "background")) vb.parentNode.removeChild(vb);
      doc.documentElement.insertBefore(el, doc.documentElement.firstChild);
    }
    pkg.touch(main);

    // theme part
    if (B.theme) {
      const tp = relPath(REL.theme) || B.theme.path;
      if (B.theme.xml != null) { pkg.zip.file(tp, B.theme.xml, { createFolders: false }); pkg.cache.delete(tp); ct.ensureOverride(tp, CT.theme); }
      else if (!B.theme.rel) dropPart(REL.theme, tp);
    }
    // styles
    if (B.styles) {
      const sp = relPath(REL.styles);
      if (sp && B.styles.created) dropPart(REL.styles, sp);
      else if (sp) { const sd = await pkg.xml(sp); if (sd) { restoreStyles(sd, B.styles); pkg.touch(sp); } }
    }
    // settings
    if (B.settings) {
      const sp = relPath(REL.settings);
      if (sp && B.settings.created) dropPart(REL.settings, sp);
      else if (sp) {
        const sd = await pkg.xml(sp);
        if (sd) {
          for (const [name, xml] of Object.entries(B.settings.items || {})) {
            removeSetting(sd, name);
            if (xml != null) putSetting(sd, restoredEl(sd, xml));
          }
          pkg.touch(sp);
        }
      }
    }
    // bullet / number colours
    if (B.numbering) {
      const np = relPath(REL.numbering) || B.numbering.path;
      const nd = await pkg.xml(np);
      if (nd && restoreLists(nd, B.numbering.levels)) pkg.touch(np);
    }
    // direct formatting we pointed at the theme
    for (const [part, recs] of Object.entries(B.direct || {})) {
      const d = await pkg.xml(part);
      if (d && revertDirect(d, recs)) pkg.touch(part);
    }
    // namespace declarations we added to part roots
    for (const [part, saved] of Object.entries(B.roots || {})) {
      const d = pkg.has(part) && (await pkg.xml(part));
      if (d && restoreRoot(d, saved)) pkg.touch(part);
    }
    // content types we added for our pictures
    for (const ext of ["png", "jpeg", "jpg", "gif"]) {
      if ((B.ctDefaults || []).includes(ext)) continue;
      if (!Object.keys(pkg.zip.files).some((n) => n.toLowerCase().endsWith("." + ext))) ct.removeDefault(ext);
    }
    // relationship parts we created that are empty again
    for (const rp of new Set([...Object.keys(pkg.zip.files), ...pkg.cache.keys()])) {
      if (!/\.rels$/.test(rp) || (B.rels || []).includes(rp)) continue;
      const rd = await pkg.xml(rp);
      if (rd && !rd.getElementsByTagName("Relationship").length) pkg.remove(rp);
    }
  }

  // ------------------------------------------------------------------ public: new document
  async function newDocument(theme, assets, settings) {
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
      `<Override PartName="/word/settings.xml" ContentType="${CT.settings}"/>` +
      `<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>` +
      `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>` +
      `<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`, { createFolders: false });
    zip.file("_rels/.rels", XML_DECL + `<Relationships xmlns="${NS.pr}"><Relationship Id="rId1" Type="${REL.officeDocument}" Target="word/document.xml"/>` +
      `<Relationship Id="rId2" Type="${REL.core}" Target="docProps/core.xml"/>` +
      `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`, { createFolders: false });
    zip.file("word/_rels/document.xml.rels", XML_DECL + `<Relationships xmlns="${NS.pr}"><Relationship Id="rId1" Type="${REL.styles}" Target="styles.xml"/>` +
      `<Relationship Id="rId2" Type="${REL.settings}" Target="settings.xml"/>` +
      `<Relationship Id="rId3" Type="${REL.numbering}" Target="numbering.xml"/></Relationships>`, { createFolders: false });
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
    return stamp(base, theme, assets, settings);
  }

  // The page geometry newDocument() uses (A4), so callers can render assets before calling it.
  const NEW_DOC_SECTION = { pageW: 11906, pageH: 16838, margins: { top: 1800, right: 1400, bottom: 1500, left: 1400, header: 432, footer: 432 } };

  return { analyze, stamp, unstamp, probe, newDocument, panelRect, NEW_DOC_SECTION, markerDescr };
});
