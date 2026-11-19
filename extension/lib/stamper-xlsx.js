// Wild Themes — writes a theme INTO an Excel workbook (.xlsx / .xlsm).
// The theme lives in the file itself: theme colours and fonts (which Excel's built-in
// cell styles, tables and charts follow), cached font names for readers that ignore the
// theme, built-in cell styles on theme colours, sheet tab colours, a pale background
// picture on every worksheet, and metadata (custom properties).
// Everything is reversible: stamp() returns a backup of exactly what it changed, and
// unstamp(bytes, backup) puts it back. Cell data, formulas, tables, charts' data,
// conditional formatting, validation, names and VBA are never touched; worksheets are
// edited as text, in place, so their cell XML stays byte-identical.
//
// Browser: load vendor/jszip.min.js and wtcore.js, then this file -> self.WWStamperXlsx.
// Node:    const WWStamperXlsx = require('./stamper-xlsx.js')  (needs jszip + @xmldom/xmldom).
(function (root, factory) {
  if (typeof module === "object" && module.exports) {
    let xmldom = null;
    try { xmldom = require("@xmldom/xmldom"); } catch (_) { /* rely on globals */ }
    module.exports = factory(require("jszip"), xmldom);
  } else {
    root.WWStamperXlsx = factory(root.JSZip, null);
  }
})(typeof self !== "undefined" ? self : this, function (JSZip, xmldom) {
  "use strict";

  const DOMParserImpl = (typeof DOMParser !== "undefined" && DOMParser) || (xmldom && xmldom.DOMParser);
  const SerializerImpl = (typeof XMLSerializer !== "undefined" && XMLSerializer) || (xmldom && xmldom.XMLSerializer);
  // The shared core, fetched when used (it requires this file, so not at load time).
  const core = () => (typeof self !== "undefined" && self.WTCore) || require("./wtcore.js");

  const NS = {
    s: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    a: "http://schemas.openxmlformats.org/drawingml/2006/main",
    pr: "http://schemas.openxmlformats.org/package/2006/relationships",
    ct: "http://schemas.openxmlformats.org/package/2006/content-types",
    cp: "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
    vt: "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
    dc: "http://purl.org/dc/elements/1.1/",
  };
  const RELNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
  const REL = {
    officeDocument: RELNS + "officeDocument", worksheet: RELNS + "worksheet", chartsheet: RELNS + "chartsheet",
    theme: RELNS + "theme", styles: RELNS + "styles", image: RELNS + "image", custom: RELNS + "custom-properties",
  };
  const CT = {
    theme: "application/vnd.openxmlformats-officedocument.theme+xml",
    styles: "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
    custom: "application/vnd.openxmlformats-officedocument.custom-properties+xml",
  };
  const IMG_CT = { png: "image/png", jpeg: "image/jpeg" };
  const FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";
  const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  const FIXED_DATE = new Date(Date.UTC(1980, 0, 1, 12)); // new zip entries: same bytes every time
  const DEFAULT_TABLE_STYLE = "TableStyleMedium2";        // follows accent1
  const FEATURES = ["backdrop", "styles", "tables", "direct", "tabs"];
  // CT_Worksheet children that must come AFTER <picture> (everything else comes before it)
  const AFTER_PICTURE = new Set(["oleObjects", "controls", "webPublishItems", "tableParts", "extLst"]);
  // Distinctive accents of Office's standard themes (2007-10, 2013-22, 2023+). Hard-coded copies of these
  // count as "old theme colours" too. Greys and Excel's "standard colours" (FFC000...) are left out on purpose.
  const OFFICE_ACCENTS = [
    ["4F81BD", "C0504D", "9BBB59", "8064A2", "4BACC6", "F79646"],
    ["4472C4", "ED7D31", null, null, "5B9BD5", "70AD47"],
    ["156082", "E97132", "196B24", "0F9ED5", "A02B93", "4EA72E"],
  ];

  // ------------------------------------------------------------------ helpers
  const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  function parseXml(str) {
    const doc = new DOMParserImpl().parseFromString(str, "application/xml");
    const err = doc.getElementsByTagName("parsererror");
    if (!doc.documentElement || (err && err.length)) throw new Error("Could not read the Excel file (XML error).");
    return doc;
  }
  const prologOf = (s) => /^\uFEFF?(?:<\?xml[^>]*\?>\s*)?/.exec(s)[0];
  const serializeDoc = (doc, prolog) => (prolog == null ? XML_DECL : prolog) + new SerializerImpl().serializeToString(doc).replace(/^<\?xml[^>]*\?>\s*/, "");
  const elementKids = (el) => { const out = []; for (let n = el && el.firstChild; n; n = n.nextSibling) if (n.nodeType === 1) out.push(n); return out; };
  const kid = (el, local) => elementKids(el).find((c) => c.localName === local) || null;
  const kidsNamed = (el, local) => elementKids(el).filter((c) => c.localName === local);
  const mk = (doc, local) => { const r = doc.documentElement; return doc.createElementNS(r.namespaceURI, r.prefix ? r.prefix + ":" + local : local); };
  const num = (el, attr) => Number(el.getAttribute(attr) || 0);
  const rId = (el) => el.getAttributeNS(NS.r, "id") || el.getAttribute("r:id") || "";
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
  const isOurMedia = (target) => /(^|\/)media\/wildtheme-/.test(target);
  const mediaId = (bytes) => fnv(bytes) + "-" + (bytes.length || bytes.byteLength || 0);
  function fnv(bytes) {
    const b = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    let h = 0x811c9dc5;
    for (let i = 0; i < b.length; i++) { h ^= b[i]; h = Math.imul(h, 0x01000193) >>> 0; }
    return h.toString(16).padStart(8, "0");
  }
  function paletteOf(theme) {
    let C = null;
    try { C = core(); } catch (_) { /* not loaded */ }
    if (!C || !C.palette) throw new Error("Load wtcore.js before writing a theme into a workbook.");
    return C.palette(theme);
  }
  function optionsOf(s) {
    const o = {};
    for (const k of FEATURES) { const v = s ? s[k] : undefined; o[k] = v === undefined ? true : v === true || v === "true" || v === 1; }
    return o;
  }

  // ------------------------------------------------------------------ text-level XML editing
  // Worksheets can be huge and hold the user's data, so they are never re-serialised: we find
  // element boundaries with this scanner and splice small fragments in or out.
  function tagEnd(x, lt) {
    let q = 0;
    for (let i = lt + 1; i < x.length; i++) {
      const c = x.charCodeAt(i);
      if (q) { if (c === q) q = 0; } else if (c === 34 || c === 39) q = c; else if (c === 62) return i;
    }
    throw new Error("Could not read the Excel file (XML error).");
  }
  const need = (i) => { if (i < 0) throw new Error("Could not read the Excel file (XML error)."); return i; };
  // Elements one level deep in x[from, to): { q, prefix, local, start, openEnd, end, empty }.
  function scan(x, from, to) {
    const out = [];
    let depth = 0, cur = null, i = from;
    while (i < to) {
      const lt = x.indexOf("<", i);
      if (lt < 0 || lt >= to) break;
      const c = x.charCodeAt(lt + 1);
      if (c === 63) { i = need(x.indexOf("?>", lt)) + 2; continue; }
      if (c === 33) {
        if (x.startsWith("<!--", lt)) i = need(x.indexOf("-->", lt)) + 3;
        else if (x.startsWith("<![CDATA[", lt)) i = need(x.indexOf("]]>", lt)) + 3;
        else i = need(x.indexOf(">", lt)) + 1;
        continue;
      }
      const gt = tagEnd(x, lt);
      if (c === 47) {
        depth--;
        if (depth === 0 && cur) { cur.end = gt + 1; out.push(cur); cur = null; }
        i = gt + 1; continue;
      }
      const empty = x.charCodeAt(gt - 1) === 47;
      if (depth === 0) {
        let j = lt + 1;
        while (j < gt && !/[\s/>]/.test(x[j])) j++;
        const q = x.slice(lt + 1, j), k = q.indexOf(":");
        cur = { q, prefix: k < 0 ? "" : q.slice(0, k), local: k < 0 ? q : q.slice(k + 1), start: lt, openEnd: gt + 1, empty };
        // sheetData never nests and holds all the cells: jump straight to its end
        const close = !empty && cur.local === "sheetData" ? x.indexOf(`</${q}>`, gt + 1) : -1;
        if (empty || close >= 0) { cur.end = empty ? gt + 1 : close + q.length + 3; out.push(cur); i = cur.end; cur = null; continue; }
        depth = 1;
      } else if (!empty) depth++;
      i = gt + 1;
    }
    return out;
  }
  const closeOf = (x, el) => x.lastIndexOf("<", el.end - 1);
  const kidsOf = (x, el) => (el.empty ? [] : scan(x, el.openEnd, closeOf(x, el)));
  const splice = (x, a, b, s) => x.slice(0, a) + s + x.slice(b);
  function attrsOf(tag) {
    const m = new Map(), re = /\s([^\s=/>]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
    let r;
    while ((r = re.exec(tag))) m.set(r[1], r[2] !== undefined ? r[2] : r[3]);
    return m;
  }
  const idAttr = (tag) => { const m = /\s[\w.-]+:id\s*=\s*["']([^"']*)["']/.exec(tag); return m ? m[1] : ""; };
  function sheetTop(x) {
    const root = scan(x, 0, x.length)[0];
    if (!root || root.empty) throw new Error("Could not read a worksheet (XML error).");
    return { root, kids: kidsOf(x, root), close: closeOf(x, root), p: root.prefix ? root.prefix + ":" : "" };
  }
  // name of a top-level child, looking inside mc:AlternateContent wrappers
  function effLocal(x, k) {
    if (k.local !== "AlternateContent") return k.local;
    for (const branch of kidsOf(x, k)) { const inner = kidsOf(x, branch)[0]; if (inner) return inner.local; }
    return "";
  }
  function findPicture(x) {
    const k = sheetTop(x).kids.find((c) => c.local === "picture");
    return k ? { start: k.start, end: k.end, id: idAttr(x.slice(k.start, k.openEnd)), str: x.slice(k.start, k.end) } : null;
  }
  function pictureTag(x, id) {
    const { root, p } = sheetTop(x);
    const decl = [...attrsOf(x.slice(root.start, root.openEnd))].find(([n, v]) => n.startsWith("xmlns:") && v === NS.r);
    return decl ? `<${p}picture ${decl[0].slice(6)}:id="${id}"/>` : `<${p}picture xmlns:r="${NS.r}" r:id="${id}"/>`;
  }
  function insertPicture(x, markup) {
    const { kids, close } = sheetTop(x);
    const before = kids.find((k) => AFTER_PICTURE.has(effLocal(x, k)));
    const at = before ? before.start : close;
    return splice(x, at, at, markup);
  }
  // <sheetPr><tabColor theme="n"/>...: tabColor is sheetPr's first child, sheetPr the sheet's first.
  function setTab(x, idx) {
    const { root, kids, p } = sheetTop(x);
    const sp = kids.find((k) => k.local === "sheetPr");
    if (!sp) return { xml: splice(x, root.openEnd, root.openEnd, `<${p}sheetPr><${p}tabColor theme="${idx}"/></${p}sheetPr>`), sp: null, tab: null };
    const sp0 = x.slice(sp.start, sp.end), tab = `<${sp.prefix ? sp.prefix + ":" : ""}tabColor theme="${idx}"/>`;
    if (sp.empty) return { xml: splice(x, sp.start, sp.end, x.slice(sp.start, sp.end - 2).replace(/\s+$/, "") + ">" + tab + `</${sp.q}>`), sp: sp0, tab: null };
    const tc = kidsOf(x, sp).find((k) => k.local === "tabColor");
    if (tc) return { xml: splice(x, tc.start, tc.end, tab), sp: sp0, tab: x.slice(tc.start, tc.end) };
    return { xml: splice(x, sp.openEnd, sp.openEnd, tab), sp: sp0, tab: null };
  }
  function revertTab(x, b, rgb) {
    const sp = sheetTop(x).kids.find((k) => k.local === "sheetPr");
    if (!sp || sp.empty) return x;
    const tc = kidsOf(x, sp).find((k) => k.local === "tabColor");
    if (!tc) return x;
    const a = attrsOf(x.slice(tc.start, tc.end));
    const same = a.get("theme") === String(b.t) || (!a.has("theme") && !!rgb && (a.get("rgb") || "").slice(-6).toUpperCase() === rgb);
    if (!same || a.has("tint")) return x; // recoloured since: keep the user's choice
    let s = x.slice(sp.start, tc.start) + (b.tab || "") + x.slice(tc.end, sp.end);
    const tidy = (m) => m.replace(/^(<[^>]*?)\s*>\s*<\/[^>]+>$/, "$1/>");
    if (b.sp != null && tidy(s) === tidy(b.sp)) s = b.sp;                 // exactly as it was
    else if (b.sp == null && /^<[^\s>]+>\s*<\/[^>]+>$/.test(s)) s = "";  // we added an empty sheetPr
    else s = tidy(s);
    return splice(x, sp.start, sp.end, s);
  }

  // ------------------------------------------------------------------ package wrapper
  class Pkg {
    constructor(zip, input) {
      this.zip = zip; this.input = input;
      this.cache = new Map(); this.strs = new Map(); this.prologs = new Map(); this.dirty = new Set(); this.changed = false;
    }
    static async open(bytes) {
      let zip;
      try { zip = await JSZip.loadAsync(bytes); } catch (_) { throw new Error("This isn't an Excel workbook (.xlsx)."); }
      if (!zip.file("[Content_Types].xml")) throw new Error("This isn't an Excel workbook (.xlsx).");
      return new Pkg(zip, bytes);
    }
    async raw(p) {
      if (this.strs.has(p)) return this.strs.get(p);
      const f = this.zip.file(p);
      if (!f) return null;
      const s = await f.async("string");
      this.strs.set(p, s);
      return s;
    }
    // current text of a part (re-serialised only if we changed its DOM)
    async str(p) { return this.cache.has(p) && this.dirty.has(p) ? this.serialize(p) : this.raw(p); }
    async xml(p) {
      if (this.cache.has(p)) return this.cache.get(p);
      const s = await this.raw(p);
      if (s == null) return null;
      const doc = parseXml(s);
      this.cache.set(p, doc); this.prologs.set(p, prologOf(s));
      return doc;
    }
    serialize(p) { return serializeDoc(this.cache.get(p), this.prologs.has(p) ? this.prologs.get(p) : XML_DECL); }
    setXml(p, doc, prolog) { this.cache.set(p, doc); if (prolog !== undefined) this.prologs.set(p, prolog); this.strs.delete(p); this.dirty.add(p); }
    setStr(p, s) { this.strs.set(p, s); this.cache.delete(p); this.dirty.add(p); }
    touch(p) { this.dirty.add(p); }
    putBinary(p, bytes) { this.zip.file(p, bytes, { binary: true, createFolders: false, compression: "STORE", date: FIXED_DATE }); this.changed = true; }
    remove(p) { this.zip.remove(p); this.cache.delete(p); this.strs.delete(p); this.dirty.delete(p); this.changed = true; }
    async rels(part) {
      const p = relsPathOf(part);
      let doc = await this.xml(p);
      if (!doc) { doc = parseXml(`<Relationships xmlns="${NS.pr}"/>`); this.cache.set(p, doc); }
      const root = doc.documentElement, self_ = this;
      const list = () => Array.from(root.getElementsByTagName("Relationship")).map((el) => ({
        el, id: el.getAttribute("Id"), type: el.getAttribute("Type"), target: el.getAttribute("Target") || "",
        external: el.getAttribute("TargetMode") === "External",
      }));
      return {
        path: p, list,
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
      const root = doc.documentElement, self_ = this, P = "[Content_Types].xml";
      const defaults = () => Array.from(root.getElementsByTagName("Default"));
      const overrides = () => Array.from(root.getElementsByTagName("Override"));
      return {
        ensureDefault(ext, type) {
          if (defaults().some((d) => (d.getAttribute("Extension") || "").toLowerCase() === ext)) return false;
          const el = doc.createElementNS(NS.ct, "Default");
          el.setAttribute("Extension", ext); el.setAttribute("ContentType", type);
          root.insertBefore(el, overrides()[0] || null); self_.touch(P);
          return true;
        },
        removeDefault(ext) { for (const d of defaults()) if ((d.getAttribute("Extension") || "").toLowerCase() === ext) { root.removeChild(d); self_.touch(P); } },
        ensureOverride(part, type) {
          const ex = overrides().find((o) => o.getAttribute("PartName") === "/" + part);
          if (ex) { if (ex.getAttribute("ContentType") !== type) { ex.setAttribute("ContentType", type); self_.touch(P); } return; }
          const el = doc.createElementNS(NS.ct, "Override");
          el.setAttribute("PartName", "/" + part); el.setAttribute("ContentType", type);
          root.appendChild(el); self_.touch(P);
        },
        removeOverride(part) { for (const o of overrides()) if (o.getAttribute("PartName") === "/" + part) { root.removeChild(o); self_.touch(P); } },
        overrideOf(part) { const o = overrides().find((x) => x.getAttribute("PartName") === "/" + part); return o ? o.getAttribute("ContentType") || "" : ""; },
      };
    }
    async save() {
      if (!this.dirty.size && !this.changed) return new Uint8Array(this.input instanceof ArrayBuffer ? this.input.slice(0) : this.input);
      for (const p of this.dirty) {
        let content;
        if (this.cache.has(p)) {
          const doc = this.cache.get(p);
          // a relationships part we emptied (or only ever planned) is dropped, not left behind empty
          if (/\.rels$/.test(p) && !doc.documentElement.getElementsByTagName("Relationship").length) { if (this.zip.file(p)) this.zip.remove(p); continue; }
          content = this.serialize(p);
        } else if (this.strs.has(p)) content = this.strs.get(p);
        else continue;
        const old = this.zip.file(p);
        this.zip.file(p, content, { createFolders: false, date: old ? old.date : FIXED_DATE });
      }
      for (const name of Object.keys(this.zip.files)) { const f = this.zip.files[name]; if (f && f.dir) delete this.zip.files[name]; }
      return this.zip.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 6 } });
    }
  }

  // The workbook part and its sheets (dialog and macro sheets are left alone).
  async function workbookOf(pkg) {
    const rootRels = await pkg.rels("");
    const r = rootRels.byType(REL.officeDocument)[0];
    if (!r) throw new Error("This isn't an Excel workbook (.xlsx).");
    const wbPath = resolveTarget("", r.target);
    const ct = await pkg.contentTypes();
    if (!/spreadsheetml|ms-excel/.test(ct.overrideOf(wbPath))) throw new Error("This isn't an Excel workbook (.xlsx).");
    const wb = await pkg.xml(wbPath);
    if (!wb || wb.documentElement.namespaceURI !== NS.s) throw new Error("This workbook is saved in Strict Open XML format, which isn't supported. Save it as a normal Excel Workbook (.xlsx) first.");
    const wbRels = await pkg.rels(wbPath);
    const sheets = [];
    for (const el of Array.from(wb.getElementsByTagNameNS(NS.s, "sheet"))) {
      const rel = wbRels.byId(rId(el));
      sheets.push({
        name: el.getAttribute("name") || "", path: rel && !rel.external ? resolveTarget(wbPath, rel.target) : null,
        isWorksheet: !!rel && rel.type === REL.worksheet, isChartsheet: !!rel && rel.type === REL.chartsheet,
      });
    }
    return { wbPath, wbRels, sheets, ct };
  }
  const partOfRel = (wb, type) => { const r = wb.wbRels.byType(type)[0]; return r && !r.external ? resolveTarget(wb.wbPath, r.target) : null; };

  // ------------------------------------------------------------------ theme part
  const SLOTS = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
  function newThemeXml(theme, P) {
    const name = esc(theme.name || "Wild Theme");
    const g = (stops) => `<a:gradFill rotWithShape="1"><a:gsLst>${stops.map(([p, m]) => `<a:gs pos="${p}"><a:schemeClr val="phClr">${m}</a:schemeClr></a:gs>`).join("")}</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>`;
    const ln = (w) => `<a:ln w="${w}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>`;
    return XML_DECL + `<a:theme xmlns:a="${NS.a}" name="Wild Theme ${name}"><a:themeElements>` +
      `<a:clrScheme name="Wild Theme ${name}">${SLOTS.map((k) => `<a:${k}><a:srgbClr val="${P[k]}"/></a:${k}>`).join("")}</a:clrScheme>` +
      `<a:fontScheme name="Wild Theme ${name}"><a:majorFont><a:latin typeface="${esc(P.major)}"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>` +
      `<a:minorFont><a:latin typeface="${esc(P.minor)}"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>` +
      `<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>` +
      g([[0, '<a:tint val="67000"/>'], [100000, '<a:tint val="81000"/>']]) + g([[0, '<a:shade val="94000"/>'], [100000, '<a:shade val="78000"/>']]) +
      `</a:fillStyleLst><a:lnStyleLst>${ln(6350)}${ln(12700)}${ln(19050)}</a:lnStyleLst>` +
      `<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>` +
      `<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/></a:schemeClr></a:solidFill>` +
      g([[0, '<a:tint val="93000"/>'], [100000, '<a:shade val="63000"/>']]) +
      `</a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`;
  }
  // Existing theme: swap only the colour scheme and the major/minor Latin fonts; keep the rest.
  function patchTheme(doc, theme, P) {
    const name = "Wild Theme " + (theme.name || "Wild Theme");
    const te = doc.getElementsByTagNameNS(NS.a, "themeElements")[0];
    if (!te) return false;
    const fresh = parseXml(newThemeXml(theme, P));
    const oldClr = doc.getElementsByTagNameNS(NS.a, "clrScheme")[0];
    const clr = doc.importNode(fresh.getElementsByTagNameNS(NS.a, "clrScheme")[0], true);
    if (oldClr) te.replaceChild(clr, oldClr); else te.insertBefore(clr, te.firstChild);
    let fs = doc.getElementsByTagNameNS(NS.a, "fontScheme")[0];
    if (!fs) {
      fs = doc.importNode(fresh.getElementsByTagNameNS(NS.a, "fontScheme")[0], true);
      te.insertBefore(fs, clr.nextSibling);
    } else {
      fs.setAttribute("name", name);
      for (const [tag, face] of [["majorFont", P.major], ["minorFont", P.minor]]) {
        const f = fs.getElementsByTagNameNS(NS.a, tag)[0];
        if (!f) continue;
        let latin = f.getElementsByTagNameNS(NS.a, "latin")[0];
        if (!latin) { latin = doc.createElementNS(NS.a, (f.prefix ? f.prefix + ":" : "") + "latin"); f.insertBefore(latin, f.firstChild); }
        latin.setAttribute("typeface", face);
      }
    }
    doc.documentElement.setAttribute("name", name);
    return true;
  }
  // Returns the backup: { path, xml: original text or null (no part), rel: we added the relationship, ct: original override }
  async function writeTheme(pkg, wb, theme, P) {
    let path = partOfRel(wb, REL.theme), made = false;
    if (!path) { path = "xl/theme/theme1.xml"; let n = 1; while (pkg.zip.file(path)) path = `xl/theme/theme${++n}.xml`; made = true; }
    const before = await pkg.raw(path);
    const b = { path, xml: before, rel: made, ct: wb.ct.overrideOf(path) || null };
    let doc = null;
    if (before != null) try { doc = parseXml(before); } catch (_) { doc = null; }
    if (doc && patchTheme(doc, theme, P)) pkg.setXml(path, doc, prologOf(before));
    else pkg.setXml(path, parseXml(newThemeXml(theme, P)), XML_DECL);
    if (made) wb.wbRels.add(REL.theme, relativeTarget(wb.wbPath, path));
    wb.ct.ensureOverride(path, CT.theme);
    return b;
  }
  async function restoreTheme(pkg, wb, t) {
    if (t.xml != null) { if ((await pkg.str(t.path)) !== t.xml) pkg.setStr(t.path, t.xml); } else if (pkg.zip.file(t.path)) pkg.remove(t.path);
    if (t.rel) for (const r of wb.wbRels.byType(REL.theme)) if (resolveTarget(wb.wbPath, r.target) === t.path) wb.wbRels.remove(r.id);
    if (!t.ct) wb.ct.removeOverride(t.path); else if (wb.ct.overrideOf(t.path) !== t.ct) wb.ct.ensureOverride(t.path, t.ct);
  }
  // Slot colours of a theme part (sysClr -> its lastClr).
  function schemeColours(themeXml) {
    const out = {};
    if (!themeXml) return out;
    let doc;
    try { doc = parseXml(themeXml); } catch (_) { return out; }
    const cs = doc.getElementsByTagNameNS(NS.a, "clrScheme")[0];
    for (const el of elementKids(cs)) {
      const v = elementKids(el)[0];
      const h = v ? (v.localName === "sysClr" ? v.getAttribute("lastClr") : v.getAttribute("val")) || "" : "";
      if (/^[0-9A-Fa-f]{6}$/.test(h)) out[el.localName] = h.toUpperCase();
    }
    return out;
  }
  // Hard-coded colours that are really the OLD theme's colours -> where they go in the new one.
  // cell: hex -> SpreadsheetML theme index (0=lt1 1=dk1 2=lt2 3=dk2 4-9=accent1-6 10=hlink 11=folHlink); chart: hex -> accentN.
  function colourMaps(themeXml) {
    const s = schemeColours(themeXml), cell = new Map(), chart = new Map();
    const skip = new Set(["000000", "FFFFFF", s.dk1, s.lt1]); // text and paper stay as they are
    const put = (h, idx, name) => { if (!h || skip.has(h)) return; if (!cell.has(h)) cell.set(h, idx); if (name && !chart.has(h)) chart.set(h, name); };
    for (let i = 1; i <= 6; i++) put(s["accent" + i], 3 + i, "accent" + i);
    put(s.dk2, 3); put(s.lt2, 2); put(s.hlink, 10); put(s.folHlink, 11);
    for (const row of OFFICE_ACCENTS) row.forEach((h, i) => put(h, 4 + i, "accent" + (i + 1)));
    return { cell, chart };
  }

  // ------------------------------------------------------------------ styles.xml
  const ITEM = { fonts: "font", fills: "fill", borders: "border", cellStyleXfs: "xf", cellXfs: "xf", cellStyles: "cellStyle" };
  const STYLE_ORDER = ["numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "cellStyles", "dxfs", "tableStyles", "colors", "extLst"];
  const COLOUR_TAGS = new Set(["color", "fgColor", "bgColor"]);
  // builtinId -> [font colour theme index, border sides to put on accent1 (side -> tint)], as in Excel's own defaults
  const BUILTIN = {
    15: [3], 16: [3, { bottom: null }], 17: [3, { bottom: "0.499984740745262" }], 18: [3, { bottom: "0.39997558519241921" }],
    19: [3], 25: [1, { top: null, bottom: null }], 8: [10], 9: [11],
  };
  const BUILTIN_BY_NAME = { title: 15, "heading 1": 16, "heading 2": 17, "heading 3": 18, "heading 4": 19, total: 25, hyperlink: 8, "followed hyperlink": 9 };
  // What a later save by Excel keeps of our changes; reverting only happens while these still match.
  // Nothing theme-dependent (font names) goes in: the core keeps the FIRST stamp's backup across re-themes.
  const SIG_ATTRS = ["rgb", "theme", "tint", "indexed", "auto", "fontId", "borderId"];
  function sig(el) {
    const out = [];
    const walk = (e, path) => {
      for (const a of SIG_ATTRS) if (e.hasAttribute(a)) out.push(`${path}@${a}=${a === "rgb" ? e.getAttribute(a).slice(-6).toUpperCase() : e.getAttribute(a)}`);
      if (e.localName === "scheme" && e.hasAttribute("val")) out.push(`${path}@val=${e.getAttribute("val")}`);
      for (const c of elementKids(e)) walk(c, path + "/" + c.localName);
    };
    walk(el, el.localName);
    return out.sort().join(";");
  }
  const snippet = (el) => new SerializerImpl().serializeToString(el);
  function stripDecls(el, rootEl) {
    for (const a of Array.from(el.attributes)) if ((a.name === "xmlns" || a.name.startsWith("xmlns:")) && rootEl.getAttribute(a.name) === a.value) el.removeAttribute(a.name);
    for (const c of elementKids(el)) stripDecls(c, rootEl);
  }
  // Parse a snippet back in the context of doc's root namespaces.
  function fromSnippet(doc, str) {
    const root = doc.documentElement;
    const decls = Array.from(root.attributes).filter((a) => a.name === "xmlns" || a.name.startsWith("xmlns:")).map((a) => ` ${a.name}="${esc(a.value)}"`).join("");
    const w = parseXml(`<wtwrap${decls}>${str}</wtwrap>`);
    const el = doc.importNode(elementKids(w.documentElement)[0], true);
    stripDecls(el, root);
    return el;
  }
  class Styles {
    constructor(doc) { this.doc = doc; this.root = doc.documentElement; this.recs = new Map(); this.added = {}; this.memo = {}; }
    list(name) { return kid(this.root, name); }
    items(name) { if (!this.memo[name]) { const l = this.list(name); this.memo[name] = l ? kidsNamed(l, ITEM[name]) : []; } return this.memo[name]; }
    // remember an element's original markup before its first change
    touch(name, i) { const k = name + "#" + i; if (!this.recs.has(k)) this.recs.set(k, { l: name, i, o: snippet(this.items(name)[i]) }); }
    append(name, el) {
      const l = this.list(name), items = this.items(name), n = items.length;
      if (!this.added[name]) this.added[name] = { at: n, n: 0, count: l.hasAttribute("count") ? l.getAttribute("count") : null };
      const last = items[n - 1];
      l.insertBefore(el, last ? last.nextSibling : l.firstChild);
      this.added[name].n++;
      if (l.hasAttribute("count")) l.setAttribute("count", String(n + 1));
      delete this.memo[name];
      return n;
    }
    record() { return [...this.recs.values()].map((r) => ({ l: r.l, i: r.i, o: r.o, s: sig(this.items(r.l)[r.i]) })); }
  }
  function setThemeColour(doc, parent, local, idx, tint) {
    let c = kid(parent, local);
    if (!c) { c = mk(doc, local); parent.appendChild(c); }
    for (const a of ["rgb", "indexed", "auto", "theme", "tint"]) c.removeAttribute(a);
    c.setAttribute("theme", String(idx));
    if (tint) c.setAttribute("tint", tint);
  }
  const hasThemeColour = (c) => !!c && c.hasAttribute("theme");
  // Point cell style k's font/border at a theme-coloured one: in place when only this style uses it,
  // otherwise as a themed copy (so Normal and other cells keep their look).
  function themeRef(st, list, attr, k, needs, fix) {
    const xfs = st.items("cellStyleXfs"), xf = xfs[k], id = num(xf, attr), el = st.items(list)[id];
    if (!el || !needs(el)) return false;
    const shared = id === 0 || xfs.some((x, i) => i !== k && num(x, attr) === id) || st.items("cellXfs").some((x) => num(x, "xfId") !== k && num(x, attr) === id);
    if (!shared) { st.touch(list, id); fix(el); return true; }
    const copy = el.cloneNode(true);
    fix(copy);
    const nid = st.append(list, copy);
    st.touch("cellStyleXfs", k); xf.setAttribute(attr, String(nid));
    st.items("cellXfs").forEach((x, j) => { if (num(x, "xfId") === k && num(x, attr) === id) { st.touch("cellXfs", j); x.setAttribute(attr, String(nid)); } });
    return true;
  }
  // Cached font names follow the theme fonts (LibreOffice and older readers use the name); Normal is the body font.
  function themeFonts(st, P, stats) {
    st.items("fonts").forEach((f, i) => {
      const sch = kid(f, "scheme"), cur = sch ? sch.getAttribute("val") : null, want = i === 0 ? "minor" : cur;
      if (want !== "minor" && want !== "major") return;
      const face = want === "major" ? P.major : P.minor, nm = kid(f, "name");
      if (nm && nm.getAttribute("val") === face && cur === want) return;
      st.touch("fonts", i);
      if (nm) nm.setAttribute("val", face);
      else { const n = mk(st.doc, "name"); n.setAttribute("val", face); f.insertBefore(n, kid(f, "family") || kid(f, "charset") || sch || null); }
      if (cur !== want) { if (sch) sch.setAttribute("val", want); else { const s = mk(st.doc, "scheme"); s.setAttribute("val", want); f.appendChild(s); } }
      stats.fonts++;
    });
  }
  function themeBuiltins(st, stats) {
    for (const cs of st.items("cellStyles")) {
      const id = cs.hasAttribute("builtinId") ? Number(cs.getAttribute("builtinId")) : BUILTIN_BY_NAME[(cs.getAttribute("name") || "").trim().toLowerCase()];
      const spec = BUILTIN[id];
      const k = num(cs, "xfId");
      if (!spec || !st.items("cellStyleXfs")[k]) continue;
      let did = themeRef(st, "fonts", "fontId", k, (f) => !hasThemeColour(kid(f, "color")), (f) => setThemeColour(st.doc, f, "color", spec[0], null));
      if (spec[1]) {
        const sides = (b) => Object.entries(spec[1]).map(([n, tint]) => [kid(b, n), tint])
          .filter(([s]) => s && s.getAttribute("style") && s.getAttribute("style") !== "none" && !hasThemeColour(kid(s, "color")));
        did = themeRef(st, "borders", "borderId", k, (b) => sides(b).length > 0, (b) => { for (const [s, tint] of sides(b)) setThemeColour(st.doc, s, "color", 4, tint); }) || did;
      }
      if (did) stats.styles++;
    }
  }
  // Font and fill colours that are the old theme's colours written out as RGB -> theme references (tint kept).
  function directStyleColours(st, map, stats) {
    const idxOf = (c) => {
      if (!c || !c.hasAttribute("rgb") || c.hasAttribute("theme")) return null;
      const v = map.get(c.getAttribute("rgb").slice(-6).toUpperCase());
      return v === undefined ? null : v;
    };
    const conv = (c, idx) => { c.removeAttribute("rgb"); c.setAttribute("theme", String(idx)); };
    st.items("fonts").forEach((f, i) => {
      const c = kid(f, "color"), idx = idxOf(c);
      if (idx == null) return;
      st.touch("fonts", i); conv(c, idx); stats.direct.fonts++;
    });
    st.items("fills").forEach((f, i) => {
      for (const c of Array.from(f.getElementsByTagName("*")).filter((e) => COLOUR_TAGS.has(e.localName))) {
        const idx = idxOf(c);
        if (idx == null) continue;
        st.touch("fills", i); conv(c, idx); stats.direct.fills++;
      }
    });
  }
  function setTableDefault(st) {
    let ts = st.list("tableStyles");
    const b = { had: !!ts, v: ts && ts.hasAttribute("defaultTableStyle") ? ts.getAttribute("defaultTableStyle") : null };
    if (!ts) {
      ts = mk(st.doc, "tableStyles"); ts.setAttribute("count", "0");
      const after = STYLE_ORDER.slice(STYLE_ORDER.indexOf("tableStyles") + 1);
      st.root.insertBefore(ts, elementKids(st.root).find((c) => after.includes(c.localName)) || null);
    }
    ts.setAttribute("defaultTableStyle", DEFAULT_TABLE_STYLE);
    return b;
  }
  function minimalStyles(P) {
    return XML_DECL + `<styleSheet xmlns="${NS.s}"><fonts count="1"><font><sz val="11"/><color theme="1"/><name val="${esc(P.minor)}"/><family val="2"/><scheme val="minor"/></font></fonts>` +
      `<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>` +
      `<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>` +
      `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
      `<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>` +
      `<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/>` +
      `<tableStyles count="0" defaultTableStyle="${DEFAULT_TABLE_STYLE}" defaultPivotStyle="PivotStyleLight16"/></styleSheet>`;
  }
  async function themeStyles(pkg, wb, P, opts, cellMap, stats) {
    const path = partOfRel(wb, REL.styles);
    if (!path || !pkg.zip.file(path)) {
      if (!opts.styles && !opts.tables) return null; // nothing to recolour without a styles part
      // no styles part (some generators): write Excel's default one, on the theme fonts
      let p = path || "xl/styles.xml", n = 1;
      if (!path) while (pkg.zip.file(p)) p = `xl/styles${++n}.xml`;
      pkg.setXml(p, parseXml(minimalStyles(P)), XML_DECL);
      if (!path) wb.wbRels.add(REL.styles, relativeTarget(wb.wbPath, p));
      const ct = wb.ct.overrideOf(p) || null;
      wb.ct.ensureOverride(p, CT.styles);
      stats.fonts++; stats.tables = true;
      return { path: p, created: true, rel: !path, ct };
    }
    const doc = await pkg.xml(path), st = new Styles(doc), b = { path };
    if (opts.styles) { themeFonts(st, P, stats); themeBuiltins(st, stats); }
    if (opts.direct) directStyleColours(st, cellMap, stats);
    if (opts.tables) { b.tsd = setTableDefault(st); stats.tables = true; }
    b.els = st.record(); b.add = st.added;
    if (b.els.length || Object.keys(b.add).length || b.tsd) pkg.touch(path);
    return b;
  }
  async function revertStyles(pkg, wb, b) {
    if (b.created) {
      if (pkg.zip.file(b.path)) pkg.remove(b.path);
      if (b.rel) for (const r of wb.wbRels.byType(REL.styles)) if (resolveTarget(wb.wbPath, r.target) === b.path) wb.wbRels.remove(r.id);
      if (!b.ct) wb.ct.removeOverride(b.path);
      return;
    }
    const doc = await pkg.xml(b.path);
    if (!doc) return;
    const st = new Styles(doc);
    let changed = false;
    for (const r of (b.els || []).slice().reverse()) {
      const el = st.items(r.l)[r.i];
      if (!el || sig(el) !== r.s) continue; // changed since in Excel: leave it
      el.parentNode.replaceChild(fromSnippet(doc, r.o), el);
      st.memo = {}; changed = true;
    }
    for (const [name, a] of Object.entries(b.add || {})) {
      const items = st.items(name);
      if (items.length !== a.at + a.n) continue;
      const attr = name === "fonts" ? "fontId" : "borderId";
      if ([...st.items("cellStyleXfs"), ...st.items("cellXfs")].some((x) => num(x, attr) >= a.at)) continue;
      for (const el of items.slice(a.at)) el.parentNode.removeChild(el);
      const l = st.list(name);
      if (a.count == null) l.removeAttribute("count"); else l.setAttribute("count", a.count);
      st.memo = {}; changed = true;
    }
    const ts = b.tsd && st.list("tableStyles");
    if (ts && ts.getAttribute("defaultTableStyle") === DEFAULT_TABLE_STYLE) {
      if (!b.tsd.had) { if (!elementKids(ts).length) ts.parentNode.removeChild(ts); else ts.removeAttribute("defaultTableStyle"); }
      else if (b.tsd.v == null) ts.removeAttribute("defaultTableStyle");
      else ts.setAttribute("defaultTableStyle", b.tsd.v);
      changed = true;
    }
    if (changed) pkg.touch(b.path);
  }

  // ------------------------------------------------------------------ charts
  const chartParts = (pkg) => Object.keys(pkg.zip.files).filter((n) => /^xl\/charts\/chart[^/]*\.xml$/i.test(n));
  function swapColour(doc, el, local, val) {
    const n = doc.createElementNS(NS.a, (el.prefix ? el.prefix + ":" : "") + local);
    n.setAttribute("val", val);
    while (el.firstChild) n.appendChild(el.firstChild);
    el.parentNode.replaceChild(n, el);
    return n;
  }
  // a:srgbClr equal to an old accent -> a:schemeClr accentN. Backup: positions among the chart's schemeClr elements.
  async function themeCharts(pkg, map, backup, stats) {
    for (const name of chartParts(pkg)) {
      const doc = await pkg.xml(name);
      const made = new Map();
      for (const el of Array.from(doc.getElementsByTagNameNS(NS.a, "srgbClr"))) {
        const v = el.getAttribute("val") || "", slot = map.get(v.toUpperCase());
        if (slot) made.set(swapColour(doc, el, "schemeClr", slot), v);
      }
      if (!made.size) continue;
      const recs = [];
      Array.from(doc.getElementsByTagNameNS(NS.a, "schemeClr")).forEach((e, n) => { if (made.has(e)) recs.push({ n, v: made.get(e), s: e.getAttribute("val") }); });
      backup.charts[name] = recs;
      stats.direct.charts += made.size;
      pkg.touch(name);
    }
  }
  async function revertCharts(pkg, charts) {
    for (const [name, recs] of Object.entries(charts)) {
      const doc = pkg.zip.file(name) && (await pkg.xml(name));
      if (!doc) continue;
      const all = Array.from(doc.getElementsByTagNameNS(NS.a, "schemeClr"));
      let changed = false;
      for (const r of recs) { const el = all[r.n]; if (el && el.getAttribute("val") === r.s) { swapColour(doc, el, "srgbClr", r.v); changed = true; } }
      if (changed) pkg.touch(name);
    }
  }

  // ------------------------------------------------------------------ media
  // Delete our media nobody points at any more.
  async function pruneMedia(pkg, candidates) {
    const used = new Set();
    // relationship parts on disk plus ones created in this run (still only in the cache)
    for (const name of new Set([...Object.keys(pkg.zip.files), ...pkg.cache.keys()])) {
      if (!/(^|\/)_rels\/[^/]+\.rels$/.test(name)) continue;
      const owner = name.replace(/(^|\/)_rels\/([^/]+)\.rels$/, (m, pre, f) => pre + f).replace(/^\//, "");
      const rd = await pkg.xml(name);
      if (!rd) continue;
      for (const el of Array.from(rd.getElementsByTagName("Relationship"))) {
        if (el.getAttribute("TargetMode") !== "External") used.add(resolveTarget(owner, el.getAttribute("Target") || ""));
      }
    }
    for (const name of Object.keys(pkg.zip.files)) if ((candidates.has(name) || isOurMedia(name)) && !used.has(name) && /\/media\//.test(name)) pkg.remove(name);
  }

  // ------------------------------------------------------------------ custom properties
  async function setCustomProps(pkg, props) {
    const path = "docProps/custom.xml";
    let doc = await pkg.xml(path);
    if (!doc) {
      doc = parseXml(`<Properties xmlns="${NS.cp}" xmlns:vt="${NS.vt}"/>`);
      pkg.setXml(path, doc, XML_DECL);
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
  async function removeCustomProps(pkg) {
    const doc = await pkg.xml("docProps/custom.xml");
    if (!doc) return;
    const root = doc.documentElement, ours = Array.from(root.getElementsByTagNameNS(NS.cp, "property")).filter((el) => /^WildTheme/.test(el.getAttribute("name") || ""));
    if (!ours.length) return;
    for (const el of ours) root.removeChild(el);
    if (!root.getElementsByTagNameNS(NS.cp, "property").length) {
      pkg.remove("docProps/custom.xml");
      (await pkg.contentTypes()).removeOverride("docProps/custom.xml");
      const rr = await pkg.rels("");
      for (const r of rr.byType(REL.custom)) rr.remove(r.id);
    } else pkg.touch("docProps/custom.xml");
  }

  // Our background media: named wildtheme-* by us or, once Excel or LibreOffice has renamed it
  // (xl/media/image1.jpeg), the same bytes as recorded in the WildThemeMedia property.
  async function mediaJudge(pkg) {
    const ids = new Set(((await readCustomProp(pkg, "WildThemeMedia")) || "").split(/[\s,;]+/).filter(Boolean));
    const memo = new Map();
    return async (part) => {
      if (isOurMedia(part)) return true;
      if (!ids.size) return false;
      if (!memo.has(part)) { const f = pkg.zip.file(part); memo.set(part, f ? ids.has(mediaId(await f.async("uint8array"))) : false); }
      return memo.get(part);
    };
  }

  // ------------------------------------------------------------------ public API
  async function analyze(bytes) {
    const pkg = await Pkg.open(bytes);
    const wb = await workbookOf(pkg);
    const core_ = await pkg.xml("docProps/core.xml");
    const t = core_ && core_.getElementsByTagNameNS(NS.dc, "title")[0];
    return {
      kind: "xlsx",
      macro: /macroEnabled/i.test(wb.ct.overrideOf(wb.wbPath)),
      sheets: wb.sheets.map((s) => s.name),
      sections: [{ pageW: 19200, pageH: 10800 }],
      title: t ? t.textContent.trim() || null : null,
      existingTheme: await readCustomProp(pkg, "WildTheme"),
    };
  }

  // settings (WTCore.settingsFor("xlsx")): { backdrop, styles, tables, direct, tabs }; missing = on.
  // Returns the new bytes with .backup (JSON: everything needed to restore the original), .stats and .failures.
  async function stamp(bytes, theme, assets, settings) {
    if (!theme || !theme.id) throw new Error("No theme chosen.");
    const opts = optionsOf(settings), P = paletteOf(theme), failures = [];
    const stats = { sheets: 0, pictures: 0, tabs: 0, fonts: 0, styles: 0, tables: false, direct: { fonts: 0, fills: 0, charts: 0 } };
    const fail = (what, e) => failures.push(what + (e && e.message ? ` (${e.message})` : "") + ".");
    let bd = opts.backdrop && assets && assets.backdrop && assets.backdrop.bytes ? assets.backdrop : null, ext = null;
    if (bd) {
      ext = String(bd.ext || "").toLowerCase().replace(/^jpg$/, "jpeg");
      if (!IMG_CT[ext]) { failures.push("Excel sheet backgrounds must be PNG or JPEG."); bd = null; }
    } else if (opts.backdrop && settings && settings.backdrop === true) failures.push("No scene picture was supplied for the sheets.");

    const pkg = await Pkg.open(bytes);
    const wb = await workbookOf(pkg);
    const backup = { v: 1, theme: null, styles: null, sheets: {}, charts: {}, ct: [] };
    // the workbook's own colours before the change: hard-coded copies of them follow the new theme
    const oldTheme = partOfRel(wb, REL.theme);
    const maps = colourMaps(oldTheme ? await pkg.raw(oldTheme) : null);

    try { backup.theme = await writeTheme(pkg, wb, theme, P); } catch (e) { fail("The theme colours and fonts could not be written", e); }
    if (opts.styles || opts.tables || opts.direct) try { backup.styles = await themeStyles(pkg, wb, P, opts, maps.cell, stats); } catch (e) { fail("The cell styles could not be updated", e); }
    if (opts.direct) try { await themeCharts(pkg, maps.chart, backup, stats); } catch (e) { fail("The chart colours could not be matched", e); }

    let media = null;
    if (bd) {
      media = `xl/media/wildtheme-backdrop-${fnv(bd.bytes)}.${ext}`;
      pkg.putBinary(media, bd.bytes);
      if (wb.ct.ensureDefault(ext, IMG_CT[ext])) backup.ct.push(ext);
    }
    const freed = new Set();
    let k = 0;
    for (const sh of wb.sheets) {
      if (!sh.path || !(sh.isWorksheet || sh.isChartsheet)) continue;
      if (sh.isWorksheet) stats.sheets++;
      const x0 = await pkg.str(sh.path);
      if (x0 == null) { if (opts.tabs || (media && sh.isWorksheet)) failures.push(`The sheet "${sh.name}" is missing from the file.`); continue; }
      let x = x0;
      const b = {};
      if (opts.tabs) {
        try { const t = 4 + (k % 6), r = setTab(x, t); x = r.xml; b.t = t; b.sp = r.sp; b.tab = r.tab; stats.tabs++; } catch (e) { fail(`The tab colour of "${sh.name}" could not be set`, e); }
        k++;
      }
      if (media && sh.isWorksheet) {
        try {
          const rels = await pkg.rels(sh.path), old = findPicture(x);
          if (old) {
            const rel = rels.byId(old.id);
            if (rel && isOurMedia(rel.target)) { if (!rel.external) freed.add(resolveTarget(sh.path, rel.target)); rels.remove(rel.id); }
            else b.pic = old.str; // the user's own background: its relationship and picture stay in the file for unstamp
            x = splice(x, old.start, old.end, "");
          }
          x = insertPicture(x, pictureTag(x, rels.add(REL.image, relativeTarget(sh.path, media))));
          stats.pictures++;
        } catch (e) { fail(`The scene picture could not be placed on "${sh.name}"`, e); }
      }
      if (x !== x0) { pkg.setStr(sh.path, x); backup.sheets[sh.path] = b; }
    }
    if (media) await pruneMedia(pkg, freed);
    const props = { WildTheme: theme.id, WildThemeName: theme.name || theme.id, WildThemeVersion: "4" };
    if (media) props.WildThemeMedia = mediaId(bd.bytes);
    try { await setCustomProps(pkg, props); } catch (e) { fail("The theme name could not be stored", e); }

    const out = await pkg.save();
    out.backup = backup; out.stats = stats; out.failures = failures;
    return out;
  }

  // Remove everything stamp() added. With its backup, the original theme, styles and tabs come back exactly;
  // without one (older files) the pictures and properties go and the theme stays, unlabelled.
  async function unstamp(bytes, backup) {
    const pkg = await Pkg.open(bytes);
    const wb = await workbookOf(pkg);
    const B = backup && typeof backup === "object" && backup.v ? backup : null;
    const freed = new Set(), ours = await mediaJudge(pkg);
    // tab colours other apps turned into RGB (LibreOffice) are recognised by the current accents
    const tp = partOfRel(wb, REL.theme), acc = schemeColours(tp ? await pkg.raw(tp) : null);
    for (const sh of wb.sheets) {
      if (!sh.path || !(sh.isWorksheet || sh.isChartsheet)) continue;
      const x0 = await pkg.str(sh.path);
      if (x0 == null) continue;
      let x = x0;
      const b = B && B.sheets && B.sheets[sh.path];
      if (sh.isWorksheet) {
        const rels = await pkg.rels(sh.path), pic = findPicture(x);
        if (pic) {
          const rel = rels.byId(pic.id);
          if (rel && !rel.external && (await ours(resolveTarget(sh.path, rel.target)))) { x = splice(x, pic.start, pic.end, ""); freed.add(resolveTarget(sh.path, rel.target)); rels.remove(rel.id); }
        }
        // stray relationships to our pictures (e.g. from an interrupted earlier stamp)
        for (const r of rels.byType(REL.image)) {
          if (r.external || x.indexOf(`:id="${r.id}"`) >= 0 || !(await ours(resolveTarget(sh.path, r.target)))) continue;
          freed.add(resolveTarget(sh.path, r.target)); rels.remove(r.id);
        }
        if (b && b.pic && !findPicture(x) && rels.byId(idAttr(b.pic))) x = insertPicture(x, b.pic);
      }
      if (b && b.t) x = revertTab(x, b, acc["accent" + (b.t - 3)]);
      if (x !== x0) pkg.setStr(sh.path, x);
    }
    await pruneMedia(pkg, freed);
    if (B && B.styles) await revertStyles(pkg, wb, B.styles);
    if (B && B.charts) await revertCharts(pkg, B.charts);
    if (B && B.theme) await restoreTheme(pkg, wb, B.theme);
    else {
      // "Wild Theme " prefix on the scheme names marks our theme; drop the label, keep the colours/fonts
      const path = partOfRel(wb, REL.theme), doc = path && pkg.zip.file(path) ? await pkg.xml(path) : null;
      if (doc) {
        for (const el of [doc.documentElement, ...Array.from(doc.getElementsByTagNameNS(NS.a, "clrScheme")), ...Array.from(doc.getElementsByTagNameNS(NS.a, "fontScheme"))]) {
          const n = el.getAttribute("name") || "";
          if (n.startsWith("Wild Theme ")) { el.setAttribute("name", n.slice(11) || "Custom"); pkg.touch(path); }
        }
      }
    }
    await removeCustomProps(pkg);
    for (const ext of (B && B.ct) || []) if (!Object.keys(pkg.zip.files).some((n) => n.toLowerCase().endsWith("." + ext))) wb.ct.removeDefault(ext);
    return pkg.save();
  }

  // What carries the theme in this file now (read by the core's integrity check).
  async function probe(bytes) {
    const pkg = await Pkg.open(bytes);
    const wb = await workbookOf(pkg);
    const themes = [];
    const tp = partOfRel(wb, REL.theme), tx = tp ? await pkg.raw(tp) : null;
    if (tx) themes.push(tx);
    let pictures = 0;
    const seen = new Set(), media = [], ours = await mediaJudge(pkg);
    for (const sh of wb.sheets) {
      if (!sh.path || !sh.isWorksheet) continue;
      const x = await pkg.raw(sh.path);
      const pic = x && findPicture(x);
      if (!pic) continue;
      const rel = (await pkg.rels(sh.path)).byId(pic.id);
      const m = rel && !rel.external ? resolveTarget(sh.path, rel.target) : null;
      if (!m || !(await ours(m))) continue;
      pictures++;
      const f = pkg.zip.file(m);
      if (f && !seen.has(m)) { seen.add(m); media.push(await f.async("uint8array")); }
    }
    return { themes, pictures, media, marker: await readCustomProp(pkg, "WildTheme"), markerName: await readCustomProp(pkg, "WildThemeName") };
  }

  return { analyze, stamp, unstamp, probe };
});
