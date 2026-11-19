// Wild Themes — writes a theme INTO an Excel workbook (.xlsx).
// The theme lives in the file itself: theme colours and fonts (which Excel's built-in
// cell styles, tables and charts follow), a pale background picture on every worksheet,
// and metadata (custom properties), so anyone who opens the workbook sees it.
//
// Browser: load vendor/jszip.min.js first, then this file -> self.WWStamperXlsx.
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

  const NS = {
    s: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    a: "http://schemas.openxmlformats.org/drawingml/2006/main",
    mc: "http://schemas.openxmlformats.org/markup-compatibility/2006",
    pr: "http://schemas.openxmlformats.org/package/2006/relationships",
    ct: "http://schemas.openxmlformats.org/package/2006/content-types",
    cp: "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
    vt: "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
    dc: "http://purl.org/dc/elements/1.1/",
  };
  const REL = {
    officeDocument: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
    worksheet: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
    theme: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
    image: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
    custom: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties",
  };
  const CT = {
    theme: "application/vnd.openxmlformats-officedocument.theme+xml",
    custom: "application/vnd.openxmlformats-officedocument.custom-properties+xml",
  };
  const IMG_CT = { png: "image/png", jpeg: "image/jpeg", jpg: "image/jpeg" };
  const FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";
  const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  // CT_Worksheet children that must come AFTER <picture> (everything else comes before it)
  const AFTER_PICTURE = new Set(["oleObjects", "controls", "webPublishItems", "tableParts", "extLst"]);

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
    if (err && err.length) throw new Error("Could not read the Excel file (XML error).");
    return doc;
  }
  function serialize(doc) {
    return XML_DECL + new SerializerImpl().serializeToString(doc).replace(/^<\?xml[^>]*\?>\s*/, "");
  }
  const elementKids = (el) => { const out = []; for (let n = el && el.firstChild; n; n = n.nextSibling) if (n.nodeType === 1) out.push(n); return out; };
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

  // ------------------------------------------------------------------ package wrapper
  class Pkg {
    constructor(zip) { this.zip = zip; this.cache = new Map(); this.dirty = new Set(); }
    static async open(bytes) {
      let zip;
      try { zip = await JSZip.loadAsync(bytes); } catch (_) { throw new Error("This isn't an Excel workbook (.xlsx)."); }
      if (!zip.file("[Content_Types].xml")) throw new Error("This isn't an Excel workbook (.xlsx).");
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
      const root = doc.documentElement, self_ = this;
      return {
        doc,
        ensureDefault(ext, type) {
          if (Array.from(root.getElementsByTagName("Default")).some((d) => (d.getAttribute("Extension") || "").toLowerCase() === ext)) return;
          const el = doc.createElementNS(NS.ct, "Default");
          el.setAttribute("Extension", ext); el.setAttribute("ContentType", type);
          root.insertBefore(el, root.getElementsByTagName("Override")[0] || null); self_.touch("[Content_Types].xml");
        },
        ensureOverride(part, type) {
          const ex = Array.from(root.getElementsByTagName("Override")).find((o) => o.getAttribute("PartName") === "/" + part);
          if (ex) { if (ex.getAttribute("ContentType") !== type) { ex.setAttribute("ContentType", type); self_.touch("[Content_Types].xml"); } return; }
          const el = doc.createElementNS(NS.ct, "Override");
          el.setAttribute("PartName", "/" + part); el.setAttribute("ContentType", type);
          root.appendChild(el); self_.touch("[Content_Types].xml");
        },
        removeOverride(part) {
          for (const o of Array.from(root.getElementsByTagName("Override"))) if (o.getAttribute("PartName") === "/" + part) { root.removeChild(o); self_.touch("[Content_Types].xml"); }
        },
        overrideOf(part) {
          const o = Array.from(root.getElementsByTagName("Override")).find((x) => x.getAttribute("PartName") === "/" + part);
          return o ? o.getAttribute("ContentType") || "" : "";
        },
      };
    }
    async save() {
      for (const p of this.dirty) if (this.cache.has(p)) this.zip.file(p, serialize(this.cache.get(p)), { createFolders: false });
      for (const name of Object.keys(this.zip.files)) { const f = this.zip.files[name]; if (f && f.dir) delete this.zip.files[name]; }
      return this.zip.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 6 } });
    }
  }

  // The workbook part and its worksheets (chartsheets / dialog sheets are left alone).
  async function workbookOf(pkg) {
    const rootRels = await pkg.rels("");
    const r = rootRels.byType(REL.officeDocument)[0];
    if (!r) throw new Error("This isn't an Excel workbook (.xlsx).");
    const wbPath = resolveTarget("", r.target);
    const ct = await pkg.contentTypes();
    if (!/spreadsheetml/.test(ct.overrideOf(wbPath))) throw new Error("This isn't an Excel workbook (.xlsx).");
    const wb = await pkg.xml(wbPath);
    if (!wb || wb.documentElement.namespaceURI !== NS.s) throw new Error("This workbook is saved in Strict Open XML format, which isn't supported. Save it as a normal Excel Workbook (.xlsx) first.");
    const wbRels = await pkg.rels(wbPath);
    const sheets = [];
    for (const el of Array.from(wb.getElementsByTagNameNS(NS.s, "sheet"))) {
      const rel = wbRels.byId(rId(el));
      sheets.push({ name: el.getAttribute("name") || "", path: rel && !rel.external ? resolveTarget(wbPath, rel.target) : null, isWorksheet: !!rel && rel.type === REL.worksheet });
    }
    return { wbPath, wbRels, sheets, ct };
  }

  // ------------------------------------------------------------------ theme part
  function colours(theme) {
    const d = theme.doc || {};
    return [["dk1", hex(d.dk1, "1F1F1F")], ["lt1", "FFFFFF"], ["dk2", hex(d.dk2, "1F3864")], ["lt2", hex(d.lt2, "EEEEEE")],
      ["accent1", hex(d.accent1 || theme.accent, "4472C4")], ["accent2", hex(d.accent2 || theme.accent2, "ED7D31")],
      ["accent3", hex(d.accent3 || theme.accent, "A5A5A5")], ["accent4", hex(d.accent4 || theme.accent2, "FFC000")],
      ["accent5", hex(d.accent5, "5B9BD5")], ["accent6", hex(d.accent6, "70AD47")],
      ["hlink", hex(d.accent2 || theme.accent2, "0563C1")], ["folHlink", hex(d.dk2, "954F72")]];
  }
  function newThemeXml(theme) {
    const d = theme.doc || {}, name = esc(theme.name || "Wild Theme");
    const g = (stops) => `<a:gradFill rotWithShape="1"><a:gsLst>${stops.map(([p, m]) => `<a:gs pos="${p}"><a:schemeClr val="phClr">${m}</a:schemeClr></a:gs>`).join("")}</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>`;
    const ln = (w) => `<a:ln w="${w}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>`;
    return XML_DECL + `<a:theme xmlns:a="${NS.a}" name="Wild Theme ${name}"><a:themeElements>` +
      `<a:clrScheme name="Wild Theme ${name}">${colours(theme).map(([k, v]) => `<a:${k}><a:srgbClr val="${v}"/></a:${k}>`).join("")}</a:clrScheme>` +
      `<a:fontScheme name="Wild Theme ${name}"><a:majorFont><a:latin typeface="${esc(d.major || "Aptos Display")}"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>` +
      `<a:minorFont><a:latin typeface="${esc(d.minor || "Aptos")}"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>` +
      `<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>` +
      g([[0, '<a:tint val="67000"/>'], [100000, '<a:tint val="81000"/>']]) + g([[0, '<a:shade val="94000"/>'], [100000, '<a:shade val="78000"/>']]) +
      `</a:fillStyleLst><a:lnStyleLst>${ln(6350)}${ln(12700)}${ln(19050)}</a:lnStyleLst>` +
      `<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>` +
      `<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/></a:schemeClr></a:solidFill>` +
      g([[0, '<a:tint val="93000"/>'], [100000, '<a:shade val="63000"/>']]) +
      `</a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`;
  }
  // Existing theme: swap only the colour scheme and the major/minor Latin fonts; keep the rest.
  function patchTheme(doc, theme) {
    const d = theme.doc || {}, name = "Wild Theme " + (theme.name || "Wild Theme");
    const te = doc.getElementsByTagNameNS(NS.a, "themeElements")[0];
    if (!te) return false;
    const fresh = parseXml(newThemeXml(theme));
    const freshClr = fresh.getElementsByTagNameNS(NS.a, "clrScheme")[0];
    const oldClr = doc.getElementsByTagNameNS(NS.a, "clrScheme")[0];
    const clr = doc.importNode(freshClr, true);
    if (oldClr) te.replaceChild(clr, oldClr); else te.insertBefore(clr, te.firstChild);
    let fs = doc.getElementsByTagNameNS(NS.a, "fontScheme")[0];
    if (!fs) {
      fs = doc.importNode(fresh.getElementsByTagNameNS(NS.a, "fontScheme")[0], true);
      te.insertBefore(fs, clr.nextSibling);
    } else {
      fs.setAttribute("name", name);
      for (const [tag, face] of [["majorFont", d.major || "Aptos Display"], ["minorFont", d.minor || "Aptos"]]) {
        const f = fs.getElementsByTagNameNS(NS.a, tag)[0];
        const latin = f && f.getElementsByTagNameNS(NS.a, "latin")[0];
        if (latin) latin.setAttribute("typeface", face);
      }
    }
    doc.documentElement.setAttribute("name", name);
    return true;
  }
  async function writeTheme(pkg, wb, theme) {
    const rel = wb.wbRels.byType(REL.theme)[0];
    if (rel) {
      const path = resolveTarget(wb.wbPath, rel.target);
      const doc = await pkg.xml(path);
      if (doc && patchTheme(doc, theme)) { pkg.touch(path); return; }
      pkg.setXml(path, parseXml(newThemeXml(theme)));
      wb.ct.ensureOverride(path, CT.theme);
      return;
    }
    let path = "xl/theme/theme1.xml", n = 1;
    while (pkg.zip.file(path)) path = `xl/theme/theme${++n}.xml`;
    pkg.setXml(path, parseXml(newThemeXml(theme)));
    wb.wbRels.add(REL.theme, relativeTarget(wb.wbPath, path));
    wb.ct.ensureOverride(path, CT.theme);
  }

  // ------------------------------------------------------------------ sheet backgrounds
  // localName of a worksheet child, looking inside mc:AlternateContent wrappers
  function childName(el) {
    if (el.namespaceURI === NS.mc && el.localName === "AlternateContent") {
      const inner = elementKids(el).map((c) => elementKids(c)[0]).find(Boolean);
      return inner ? inner.localName : "";
    }
    return el.localName;
  }
  // Point every worksheet's background at `media` (null = remove our backgrounds).
  // Returns media parts that are no longer referenced by a sheet we touched.
  async function setBackgrounds(pkg, wb, media) {
    const freed = new Set();
    for (const sh of wb.sheets) {
      if (!sh.isWorksheet || !sh.path) continue;
      const doc = await pkg.xml(sh.path);
      if (!doc) continue;
      const root = doc.documentElement, rels = await pkg.rels(sh.path);
      const existing = elementKids(root).find((c) => c.namespaceURI === NS.s && c.localName === "picture");
      if (existing) {
        const rel = rels.byId(rId(existing));
        if (!media && !(rel && isOurMedia(rel.target))) continue;          // unstamp: keep the user's own background
        if (rel) { if (!rel.external) freed.add(resolveTarget(sh.path, rel.target)); rels.remove(rel.id); }
        root.removeChild(existing);
      }
      // drop stray relationships to our pictures (e.g. from an interrupted earlier stamp)
      for (const r of rels.byType(REL.image)) {
        if (!isOurMedia(r.target)) continue;
        const used = Array.from(root.getElementsByTagName("*")).some((e) => rId(e) === r.id);
        if (!used) { freed.add(resolveTarget(sh.path, r.target)); rels.remove(r.id); }
      }
      if (media) {
        const id = rels.add(REL.image, relativeTarget(sh.path, media));
        const pic = doc.createElementNS(NS.s, "picture");
        pic.setAttributeNS("http://www.w3.org/2000/xmlns/", "xmlns:r", NS.r);
        pic.setAttributeNS(NS.r, "r:id", id);
        const before = elementKids(root).find((c) => AFTER_PICTURE.has(childName(c))) || null;
        root.insertBefore(pic, before);
      }
      pkg.touch(sh.path);
    }
    return freed;
  }
  // Delete media nobody points at any more (and any stray media of ours).
  async function pruneMedia(pkg, candidates) {
    const used = new Set();
    // relationship parts on disk plus ones created in this run (still only in the cache)
    for (const name of new Set([...Object.keys(pkg.zip.files), ...pkg.cache.keys()])) {
      if (!/(^|\/)_rels\/[^/]+\.rels$/.test(name)) continue;
      const owner = name.replace(/(^|\/)_rels\/([^/]+)\.rels$/, (m, pre, f) => pre + f).replace(/^\//, "");
      const rd = pkg.cache.get(name) || (pkg.zip.file(name) && parseXml(await pkg.zip.file(name).async("string")));
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
  async function removeCustomProps(pkg) {
    const doc = await pkg.xml("docProps/custom.xml");
    if (!doc) return;
    const root = doc.documentElement;
    for (const el of Array.from(root.getElementsByTagNameNS(NS.cp, "property"))) if (/^WildTheme/.test(el.getAttribute("name") || "")) root.removeChild(el);
    if (!root.getElementsByTagNameNS(NS.cp, "property").length) {
      pkg.remove("docProps/custom.xml");
      (await pkg.contentTypes()).removeOverride("docProps/custom.xml");
      const rr = await pkg.rels("");
      for (const r of rr.byType(REL.custom)) rr.remove(r.id);
    } else pkg.touch("docProps/custom.xml");
  }

  // ------------------------------------------------------------------ public API
  async function analyze(bytes) {
    const pkg = await Pkg.open(bytes);
    const wb = await workbookOf(pkg);
    const core = await pkg.xml("docProps/core.xml");
    const t = core && core.getElementsByTagNameNS(NS.dc, "title")[0];
    return {
      kind: "xlsx",
      sheets: wb.sheets.map((s) => s.name),
      sections: [{ pageW: 19200, pageH: 10800 }],
      title: t ? t.textContent.trim() || null : null,
      existingTheme: await readCustomProp(pkg, "WildTheme"),
    };
  }

  async function stamp(bytes, theme, assets) {
    if (!theme || !theme.id) throw new Error("No theme chosen.");
    const bd = assets && assets.backdrop;
    if (!bd || !bd.bytes) throw new Error("No background picture for the workbook.");
    const ext = bd.ext === "jpg" ? "jpeg" : bd.ext;
    if (!IMG_CT[ext]) throw new Error("Excel sheet backgrounds must be PNG or JPEG.");
    const pkg = await Pkg.open(bytes);
    const wb = await workbookOf(pkg);
    await writeTheme(pkg, wb, theme);
    const tag = Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
    const media = `xl/media/wildtheme-backdrop-${tag}.${ext}`;
    pkg.putBinary(media, bd.bytes);
    wb.ct.ensureDefault(ext, IMG_CT[ext]);
    const freed = await setBackgrounds(pkg, wb, media);
    await pruneMedia(pkg, freed);
    await setCustomProps(pkg, { WildTheme: theme.id, WildThemeName: theme.name || theme.id, WildThemeVersion: "2" });
    const out = await pkg.save();
    out.stats = { sheets: wb.sheets.filter((s) => s.isWorksheet).length };
    return out;
  }

  async function unstamp(bytes) {
    const pkg = await Pkg.open(bytes);
    const wb = await workbookOf(pkg);
    const freed = await setBackgrounds(pkg, wb, null);
    await pruneMedia(pkg, freed);
    await removeCustomProps(pkg);
    // "Wild Theme " prefix on the scheme names marks our theme; drop the label, keep the colours/fonts
    const rel = wb.wbRels.byType(REL.theme)[0];
    if (rel) {
      const path = resolveTarget(wb.wbPath, rel.target), doc = await pkg.xml(path);
      if (doc) {
        for (const el of [doc.documentElement, ...Array.from(doc.getElementsByTagNameNS(NS.a, "clrScheme")), ...Array.from(doc.getElementsByTagNameNS(NS.a, "fontScheme"))]) {
          const n = el.getAttribute("name") || "";
          if (n.startsWith("Wild Theme ")) { el.setAttribute("name", n.slice(11) || "Custom"); pkg.touch(path); }
        }
      }
    }
    return pkg.save();
  }

  return { analyze, stamp, unstamp };
});
