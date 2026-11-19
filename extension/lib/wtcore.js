// Wild Themes — document core. ONE implementation shared by the website, the browser
// extension and the Office add-in, so all three read, write, check and remove themes the same way.
//
// A themed file carries everything itself:
//   • native theme parts (colours + fonts), styles, backdrop pictures — written by the per-format stampers;
//   • a custom XML part (namespace below) holding the theme definition, its settings, integrity
//     fingerprints and a backup of what was there before, so any copy of the tools can recognise,
//     verify, change, repair or remove the theme without the creator's account or theme library.
// Ordinary Office files are not tamper-proof: the fingerprints detect accidental changes, nothing more.
//
// Browser: load vendor/jszip.min.js, themes.js, the three stampers, then this file -> self.WTCore.
// Node:    const WTCore = require('./wtcore.js')  (needs jszip + @xmldom/xmldom).
(function (root, factory) {
  if (typeof module === "object" && module.exports) {
    let xmldom = null;
    try { xmldom = require("@xmldom/xmldom"); } catch (_) { /* rely on globals */ }
    const S = { docx: require("./stamper.js"), pptx: require("./stamper-pptx.js"), xlsx: require("./stamper-xlsx.js") };
    module.exports = factory(require("jszip"), xmldom, S, require("crypto"));
  } else {
    root.WTCore = factory(root.JSZip, null, null, null);
  }
})(typeof self !== "undefined" ? self : this, function (JSZip, xmldom, nodeStampers, nodeCrypto) {
  "use strict";

  const VERSION = 4;
  const NS = "urn:wild-themes:document-theme";
  const DS = "http://schemas.openxmlformats.org/officeDocument/2006/customXml";
  const REL_CUSTOMXML = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
  const REL_CUSTOMXML_PROPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps";
  const REL_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
  const CT_PROPS = "application/vnd.openxmlformats-officedocument.customXmlProperties+xml";
  const PR = "http://schemas.openxmlformats.org/package/2006/relationships";
  const CTNS = "http://schemas.openxmlformats.org/package/2006/content-types";
  const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  const MAIN_CT = {
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml": "docx",
    "application/vnd.ms-word.document.macroEnabled.main+xml": "docx",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml": "docx",
    "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml": "pptx",
    "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml": "pptx",
    "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml": "pptx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml": "xlsx",
    "application/vnd.ms-excel.sheet.macroEnabled.main+xml": "xlsx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml": "xlsx",
  };

  const DOMParserImpl = (typeof DOMParser !== "undefined" && DOMParser) || (xmldom && xmldom.DOMParser);
  const SerializerImpl = (typeof XMLSerializer !== "undefined" && XMLSerializer) || (xmldom && xmldom.XMLSerializer);
  const parseXml = (s) => new DOMParserImpl().parseFromString(s, "application/xml");
  const serialize = (d) => { const s = new SerializerImpl().serializeToString(d); return s.startsWith("<?xml") ? s : XML_DECL + s; };
  const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const stampers = () => nodeStampers || { docx: self.WWStamper, pptx: self.WWStamperPptx, xlsx: self.WWStamperXlsx };

  // ------------------------------------------------------------------ settings and features
  // Every switch the tools offer. Each format supports a subset (FEATURES); the rest are ignored.
  const SETTINGS = {
    backdrop: { type: "bool", def: true, label: "Scene picture", tip: "The theme's scene behind the page, slide or sheet." },
    panel: { type: "range", def: 100, min: 60, max: 100, label: "Reading panel", tip: "How solid the light panel behind the text is." },
    styles: { type: "bool", def: true, label: "Text styles", tip: "Title, headings, quotes, links and cell styles in the theme's colours and fonts." },
    tables: { type: "bool", def: true, label: "Tables", tip: "Table borders, header rows and banding in the theme's colours." },
    lists: { type: "bool", def: true, label: "Bullets and numbering", tip: "Bullet and number colours follow the theme." },
    direct: { type: "bool", def: true, label: "Match old colours", tip: "Text and shapes coloured with the file's old theme colours move to the new ones." },
    lock: { type: "bool", def: true, label: "Lock style set", tip: "Stops Word's Design gallery from swapping the theme by accident." },
    glow: { type: "bool", def: false, label: "Heading glow", tip: "A soft glow in the accent colour around headings (Word 2010 and later)." },
    animated: { type: "bool", def: true, label: "Animated scene", tip: "The scene moves while the slides are shown." },
    transition: { type: "choice", def: "fade", options: ["none", "fade", "push", "wipe", "split", "dissolve"], label: "Slide transition", tip: "Played between slides in a slide show." },
    tabs: { type: "bool", def: true, label: "Sheet tab colours", tip: "Worksheet tabs in the theme's accent colours." },
  };
  const FEATURES = {
    docx: ["backdrop", "panel", "styles", "tables", "lists", "direct", "lock", "glow"],
    pptx: ["backdrop", "panel", "styles", "tables", "direct", "animated", "transition"],
    xlsx: ["backdrop", "styles", "tables", "direct", "tabs"],
  };
  const KIND_DEFAULTS = { pptx: { panel: 90 } };
  function settingsFor(kind, given) {
    const out = {};
    for (const k of FEATURES[kind] || []) {
      const spec = SETTINGS[k];
      let v = given && given[k] !== undefined ? given[k] : (KIND_DEFAULTS[kind] || {})[k] !== undefined ? KIND_DEFAULTS[kind][k] : spec.def;
      if (spec.type === "bool") v = v === true || v === "true" || v === 1;
      else if (spec.type === "range") v = Math.max(spec.min, Math.min(spec.max, Math.round(Number(v) || spec.def)));
      else if (spec.type === "choice" && !spec.options.includes(v)) v = spec.def;
      out[k] = v;
    }
    return out;
  }

  // ------------------------------------------------------------------ theme definition
  const hex6 = (v, d) => { const s = String(v || "").replace("#", "").toUpperCase(); return /^[0-9A-F]{6}$/.test(s.slice(0, 6)) ? s.slice(0, 6) : d; };
  // The native colour + font scheme written into every format. Same rules everywhere.
  function palette(theme) {
    const d = (theme && theme.doc) || {};
    const accent1 = hex6(d.accent1 || theme.accent, "4472C4"), accent2 = hex6(d.accent2 || theme.accent2, "ED7D31");
    return {
      dk1: hex6(d.dk1, "1F1F1F"), lt1: "FFFFFF", dk2: hex6(d.dk2, "1F3864"), lt2: hex6(d.lt2, "EEEEEE"),
      accent1, accent2, accent3: hex6(d.accent3 || theme.accent, "A5A5A5"), accent4: hex6(d.accent4 || theme.accent2, "FFC000"),
      accent5: hex6(d.accent5, "5B9BD5"), accent6: hex6(d.accent6, "70AD47"), hlink: accent2, folHlink: hex6(d.dk2, "1F3864"),
      major: d.major || "Aptos Display", minor: d.minor || "Aptos", panel: hex6(d.panel, "FFFFFF"),
    };
  }
  const SLOTS = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
  // What goes into the file: the definition minus anything bulky (custom pictures live in the file as media).
  function portableDef(theme) {
    const keep = ["id", "name", "emoji", "mode", "blurb", "accent", "accent2", "fx", "doc", "custom", "colors", "dim", "part", "anim"];
    const out = {};
    for (const k of keep) if (theme[k] !== undefined) out[k] = theme[k];
    if (theme.image) out.hadImage = true;
    return out;
  }

  // ------------------------------------------------------------------ hashing
  async function sha256(data) {
    const bytes = typeof data === "string" ? new TextEncoder().encode(data) : data;
    if (nodeCrypto) return nodeCrypto.createHash("sha256").update(bytes).digest("hex");
    const h = await crypto.subtle.digest("SHA-256", bytes);
    return Array.from(new Uint8Array(h), (b) => b.toString(16).padStart(2, "0")).join("");
  }
  // Canonical colour + font scheme of a theme part, so a re-save by Office (which reformats the XML) doesn't count as a change.
  function schemeOf(themeXmlString) {
    const doc = parseXml(themeXmlString);
    const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const cs = doc.getElementsByTagNameNS(A, "clrScheme")[0], fs = doc.getElementsByTagNameNS(A, "fontScheme")[0];
    const out = {};
    if (cs) for (const slot of SLOTS) {
      const el = Array.from(cs.childNodes).find((c) => c.nodeType === 1 && c.localName === slot);
      const v = el && Array.from(el.childNodes).find((c) => c.nodeType === 1);
      out[slot] = v ? (v.getAttribute("val") === "windowText" || v.getAttribute("val") === "window" ? (v.getAttribute("lastClr") || v.getAttribute("val")) : v.getAttribute("val") || "").toUpperCase() : "";
    }
    if (fs) for (const k of ["majorFont", "minorFont"]) {
      const f = fs.getElementsByTagNameNS(A, k)[0], l = f && f.getElementsByTagNameNS(A, "latin")[0];
      out[k === "majorFont" ? "major" : "minor"] = l ? l.getAttribute("typeface") : "";
    }
    return out;
  }
  const schemeString = (s) => [...SLOTS, "major", "minor"].map((k) => `${k}=${s[k] || ""}`).join(";");

  // ------------------------------------------------------------------ package helpers
  const relsPathOf = (p) => { const i = p.lastIndexOf("/"); return `${p.slice(0, i + 1)}_rels/${p.slice(i + 1)}.rels`; };
  function resolve(from, target) {
    if (target.startsWith("/")) return target.slice(1);
    const parts = from.split("/"); parts.pop();
    for (const seg of target.split("/")) { if (seg === "..") parts.pop(); else if (seg && seg !== ".") parts.push(seg); }
    return parts.join("/");
  }
  function relative(from, to) {
    const a = from.split("/"); a.pop(); const b = to.split("/");
    while (a.length && b.length > 1 && a[0] === b[0]) { a.shift(); b.shift(); }
    return "../".repeat(a.length) + b.join("/");
  }
  async function readXml(zip, path) { const f = zip.file(path); return f ? parseXml(await f.async("string")) : null; }
  async function mainPartOf(zip) {
    const rels = await readXml(zip, "_rels/.rels");
    if (!rels) throw new Error("This isn't an Office file (no package relationships).");
    const r = Array.from(rels.getElementsByTagNameNS(PR, "Relationship")).find((e) => e.getAttribute("Type") === REL_DOC);
    if (!r) throw new Error("This Office file has no main document part.");
    return resolve("", r.getAttribute("Target"));
  }
  async function kindOf(zip) {
    const ct = await readXml(zip, "[Content_Types].xml");
    const main = await mainPartOf(zip);
    const o = ct && Array.from(ct.getElementsByTagNameNS(CTNS, "Override")).find((e) => e.getAttribute("PartName") === "/" + main);
    const kind = o && MAIN_CT[o.getAttribute("ContentType")];
    if (!kind) throw new Error("Only Word (.docx), PowerPoint (.pptx) and Excel (.xlsx) files can carry a theme.");
    return { kind, main };
  }
  function relEls(doc) { return Array.from(doc.getElementsByTagNameNS(PR, "Relationship")); }
  function newRelId(doc) { let n = 1; const ids = new Set(relEls(doc).map((e) => e.getAttribute("Id"))); while (ids.has("rIdWT" + n)) n++; return "rIdWT" + n; }
  function emptyRels() { return parseXml(`<Relationships xmlns="${PR}"/>`); }
  function guid() {
    const b = new Uint8Array(16);
    (nodeCrypto ? nodeCrypto.webcrypto : crypto).getRandomValues(b);
    const h = Array.from(b, (x) => x.toString(16).padStart(2, "0")).join("").toUpperCase();
    return `{${h.slice(0, 8)}-${h.slice(8, 12)}-4${h.slice(13, 16)}-${h.slice(16, 20)}-${h.slice(20)}}`;
  }

  // ------------------------------------------------------------------ the embedded definition (custom XML part)
  function metaXml(m) {
    const j = (o) => esc(JSON.stringify(o));
    return XML_DECL + `<wt:wildTheme xmlns:wt="${NS}" wt:version="${VERSION}" wt:id="${esc(m.id)}" wt:name="${esc(m.name)}" wt:kind="${m.kind}" wt:applied="${esc(m.applied)}" wt:tool="${esc(m.tool || "")}">` +
      `<wt:definition>${j(m.def)}</wt:definition><wt:settings>${j(m.settings)}</wt:settings>` +
      `<wt:integrity>${j(m.integrity)}</wt:integrity><wt:backup>${j(m.backup || null)}</wt:backup></wt:wildTheme>`;
  }
  function parseMeta(xml) {
    const doc = parseXml(xml), r = doc.documentElement;
    if (!r || r.namespaceURI !== NS || r.localName !== "wildTheme") return null;
    const a = (n) => r.getAttributeNS(NS, n) || r.getAttribute("wt:" + n) || "";
    const j = (n) => { const el = r.getElementsByTagNameNS(NS, n)[0]; try { return el ? JSON.parse(el.textContent) : null; } catch (_) { return { corrupt: true }; } };
    return { version: Number(a("version")) || 0, id: a("id"), name: a("name"), kind: a("kind"), applied: a("applied"), tool: a("tool"),
      def: j("definition"), settings: j("settings"), integrity: j("integrity"), backup: j("backup") };
  }
  // Find our part among the file's custom XML parts.
  async function findMetaPart(zip) {
    for (const name of Object.keys(zip.files)) {
      if (!/^customXml\/item\d+\.xml$/.test(name)) continue;
      const s = await zip.file(name).async("string");
      if (s.indexOf(NS) < 0) continue;
      const m = parseMeta(s);
      if (m) return { path: name, meta: m };
    }
    return null;
  }
  async function writeMetaPart(zip, main, m) {
    const found = await findMetaPart(zip);
    let path = found && found.path;
    if (!path) {
      let n = 1; while (zip.file(`customXml/item${n}.xml`)) n++;
      path = `customXml/item${n}.xml`;
      const props = `customXml/itemProps${n}.xml`;
      zip.file(props, XML_DECL + `<ds:datastoreItem ds:itemID="${guid()}" xmlns:ds="${DS}"><ds:schemaRefs><ds:schemaRef ds:uri="${NS}"/></ds:schemaRefs></ds:datastoreItem>`);
      zip.file(relsPathOf(path), XML_DECL + `<Relationships xmlns="${PR}"><Relationship Id="rId1" Type="${REL_CUSTOMXML_PROPS}" Target="itemProps${n}.xml"/></Relationships>`);
      const ct = await readXml(zip, "[Content_Types].xml"), root = ct.documentElement;
      const has = (tag, attr, v) => Array.from(ct.getElementsByTagNameNS(CTNS, tag)).some((e) => (e.getAttribute(attr) || "").toLowerCase() === v.toLowerCase());
      if (!has("Default", "Extension", "xml")) { const d = ct.createElementNS(CTNS, "Default"); d.setAttribute("Extension", "xml"); d.setAttribute("ContentType", "application/xml"); root.insertBefore(d, root.firstChild); }
      if (!has("Override", "PartName", "/" + props)) { const o = ct.createElementNS(CTNS, "Override"); o.setAttribute("PartName", "/" + props); o.setAttribute("ContentType", CT_PROPS); root.appendChild(o); }
      zip.file("[Content_Types].xml", serialize(ct));
      const rp = relsPathOf(main), rels = (await readXml(zip, rp)) || emptyRels();
      const el = rels.createElementNS(PR, "Relationship");
      el.setAttribute("Id", newRelId(rels)); el.setAttribute("Type", REL_CUSTOMXML); el.setAttribute("Target", relative(main, path));
      rels.documentElement.appendChild(el);
      zip.file(rp, serialize(rels));
    }
    zip.file(path, metaXml(m));
    return path;
  }
  async function removeMetaPart(zip, main) {
    const found = await findMetaPart(zip);
    if (!found) return false;
    const path = found.path, relsP = relsPathOf(path);
    const ir = await readXml(zip, relsP);
    const props = ir && relEls(ir).map((e) => resolve(path, e.getAttribute("Target")));
    zip.remove(path); zip.remove(relsP);
    const ct = await readXml(zip, "[Content_Types].xml");
    for (const p of props || []) {
      zip.remove(p);
      for (const o of Array.from(ct.getElementsByTagNameNS(CTNS, "Override"))) if (o.getAttribute("PartName") === "/" + p) o.parentNode.removeChild(o);
    }
    zip.file("[Content_Types].xml", serialize(ct));
    // drop the relationship from whichever part pointed at it (normally the main part)
    for (const name of Object.keys(zip.files)) {
      if (!/_rels\/[^/]*\.rels$/.test(name) || name === relsP) continue;
      const owner = name.replace(/_rels\/([^/]*)\.rels$/, "$1");
      const rd = await readXml(zip, name);
      let changed = false;
      for (const e of relEls(rd)) if (e.getAttribute("Type") === REL_CUSTOMXML && resolve(owner, e.getAttribute("Target")) === path) { e.parentNode.removeChild(e); changed = true; }
      if (changed) zip.file(name, serialize(rd));
    }
    return true;
  }

  // ------------------------------------------------------------------ integrity
  // probe = what the format's stamper sees in the file now: { themes: [xml], pictures: n, media: [bytes], marker }
  async function seal(probe, def, settings) {
    const schemes = probe.themes.map((x) => schemeString(schemeOf(x)));
    return {
      alg: "SHA-256",
      scheme: await sha256(schemes.join("|")),
      themes: probe.themes.length,
      pictures: probe.pictures,
      media: await Promise.all(probe.media.map((b) => sha256(b))),
      definition: await sha256(JSON.stringify(def) + JSON.stringify(settings)),
      expect: schemeString(palette(def)),
    };
  }
  async function verify(meta, probe) {
    const problems = [], i = meta.integrity || {};
    if (!i.scheme) return { status: "unverified", problems: ["The file has no fingerprints to check."] };
    if ((await sha256(JSON.stringify(meta.def) + JSON.stringify(meta.settings))) !== i.definition) problems.push({ part: "definition", level: "modified", text: "The embedded theme definition was edited outside Wild Themes." });
    const schemes = probe.themes.map((x) => schemeString(schemeOf(x)));
    if (!probe.themes.length) problems.push({ part: "theme", level: "partial", text: "The theme colours and fonts part is missing." });
    else if ((await sha256(schemes.join("|"))) !== i.scheme) problems.push({ part: "theme", level: "modified", text: "The theme colours or fonts were changed (for example with Design › Colours)." });
    if (probe.pictures < (i.pictures || 0)) problems.push({ part: "pictures", level: "partial", text: `${i.pictures - probe.pictures} of the theme's ${i.pictures} scene pictures are missing.` });
    if (i.media && i.media.length && probe.media.length) {
      const now = new Set(await Promise.all(probe.media.map((b) => sha256(b))));
      if (!i.media.every((h) => now.has(h))) problems.push({ part: "media", level: "info", text: "Office re-encoded the scene pictures (normal after some saves)." });
    }
    const worst = problems.some((p) => p.level === "partial") ? "partial" : problems.some((p) => p.level === "modified") ? "modified" : "intact";
    return { status: worst, problems };
  }

  // ------------------------------------------------------------------ public API
  async function open(bytes) {
    const zip = await JSZip.loadAsync(bytes);
    const { kind, main } = await kindOf(zip);
    return { zip, kind, main, S: stampers()[kind] };
  }
  const save = (zip) => zip.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 6 }, mimeType: undefined });

  // What theme does this file carry, and is it intact? Works on any copy, from anyone.
  async function inspect(bytes) {
    const { zip, kind, S } = await open(bytes);
    const found = await findMetaPart(zip);
    const probe = await S.probe(bytes);
    if (!found) {
      if (probe.marker) return { kind, status: "legacy", theme: { id: probe.marker, name: probe.markerName || probe.marker }, settings: null, problems: [{ part: "definition", level: "info", text: "Made by an older version: re-apply to embed the full definition and backup." }], probe };
      return { kind, status: "none", theme: null, settings: null, problems: [], probe };
    }
    const m = found.meta;
    if (m.def && m.def.corrupt) return { kind, status: "damaged", theme: { id: m.id, name: m.name }, settings: null, problems: [{ part: "definition", level: "partial", text: "The embedded definition can't be read." }], meta: m, probe };
    const v = await verify(m, probe);
    return { kind, status: v.status, problems: v.problems, theme: m.def || { id: m.id, name: m.name }, settings: m.settings, applied: m.applied, tool: m.tool, version: m.version, hasBackup: !!m.backup, meta: m, probe };
  }

  // Take the file back to how it was before any Wild Theme (restores the backed-up theme and styles).
  async function stripTo(bytes) {
    const { zip, kind, main, S } = await open(bytes);
    const found = await findMetaPart(zip);
    const backup = found ? found.meta.backup : null;
    let clean = await S.unstamp(bytes, backup || undefined);
    if (found) { const z2 = await JSZip.loadAsync(clean); await removeMetaPart(z2, main); clean = await save(z2); }
    return { clean, backup, kind, main, meta: found && found.meta };
  }

  // Apply, change or reconfigure: always rebuilt from the untouched original, so repeating it never stacks.
  // makeAssets(theme, kind, section, settings) -> { backdrop, panel } (pictures are made by the caller: canvas in browsers).
  async function apply(bytes, theme, settings, makeAssets, tool) {
    if (!theme || !theme.id) throw new Error("No theme chosen.");
    const first = await open(bytes);
    const kind = first.kind, S = first.S;
    const opts = settingsFor(kind, settings);
    const base = await stripTo(bytes);
    const info = await S.analyze(base.clean);
    const assets = await makeAssets(theme, kind, (info.sections && info.sections[0]) || null, opts);
    const out = await S.stamp(base.clean, theme, assets, opts);
    // the backup always describes the ORIGINAL file, even when re-themed many times
    const backup = base.backup || out.backup || null;
    const probe = await S.probe(out);
    const failed = (out.failures || []).concat(opts.backdrop && assets.backdrop && probe.pictures === 0 ? ["The scene picture could not be placed."] : []);
    if (!probe.themes.length) failed.push("The theme colours and fonts could not be written.");
    if (failed.length) { const e = new Error("The theme was only partly applied: " + failed.join(" ")); e.partial = failed; throw e; }
    const def = portableDef(theme);
    const zip = await JSZip.loadAsync(out);
    await writeMetaPart(zip, base.main, {
      id: theme.id, name: theme.name || theme.id, kind, applied: new Date().toISOString(), tool: tool || "",
      def, settings: opts, integrity: await seal(probe, def, opts), backup,
    });
    const result = await save(zip);
    result.stats = out.stats; result.kind = kind;
    return result;
  }

  // Remove: restores a native look; content is untouched.
  async function remove(bytes) { return (await stripTo(bytes)).clean; }

  // Restore / repair: re-applies the definition embedded in the file itself (no theme library needed).
  async function repair(bytes, makeAssets, tool, catalogue) {
    const r = await inspect(bytes);
    if (r.status === "none") throw new Error("This file has no theme to restore.");
    let theme = r.meta && r.meta.def && r.meta.def.id ? r.meta.def : null;
    if (!theme && catalogue && catalogue[r.theme.id]) theme = { ...catalogue[r.theme.id], id: r.theme.id };
    if (!theme) throw new Error("This file's theme definition is unreadable and isn't one of the built-in themes.");
    return apply(bytes, theme, r.settings || {}, makeAssets, tool);
  }

  return { VERSION, NS, SETTINGS, FEATURES, settingsFor, palette, portableDef, schemeOf, sha256,
    inspect, apply, remove, repair, findMetaPart, parseMeta, metaXml, kindOf: async (bytes) => (await open(bytes)).kind };
});
