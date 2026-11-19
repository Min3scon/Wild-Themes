// Word (.docx) deep theming: what a theme changes inside the file, that every setting is honoured,
// and that removing the theme gives the original file back exactly. Runs the real pipeline (wtcore.js -> stamper.js).
const test = require("node:test"), assert = require("node:assert/strict");
const path = require("path"), fs = require("fs");
const JSZip = require("jszip");
const { DOMParser, XMLSerializer } = require("@xmldom/xmldom");
const H = require("./helpers.js");
const { WT, theme, fixture, makeAssets } = H;
const WWStamper = require(path.join(__dirname, "..", "extension", "lib", "stamper.js"));

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const parse = (s) => new DOMParser().parseFromString(s, "application/xml");
const canon = (s) => new XMLSerializer().serializeToString(parse(s).documentElement);
const ser = (n) => new XMLSerializer().serializeToString(n);
const kids = (el, name, ns = W) => Array.from(el.childNodes).filter((n) => n.nodeType === 1 && n.namespaceURI === ns && n.localName === name);
const kid = (el, name, ns = W) => (el ? kids(el, name, ns)[0] || null : null);
const all = (el, name, ns = W) => Array.from(el.getElementsByTagNameNS(ns, name));
const wv = (el, name = "val") => (el ? el.getAttributeNS(W, name) : "");
const apply = (bytes, id, s) => WT.apply(bytes, theme(id), s || {}, makeAssets, "test");
const xml = async (bytes, p) => parse(await H.text(bytes, p));

const styleList = async (bytes) => kids((await xml(bytes, "word/styles.xml")).documentElement, "style");
const styleMap = async (bytes) => new Map((await styleList(bytes)).map((s) => [wv(s, "styleId"), ser(s)]));
const style = async (bytes, id) => (await styleList(bytes)).find((s) => wv(s, "styleId") === id) || null;
const colorOf = (rPrParent) => kid(kid(rPrParent, "rPr"), "color");
// The document's text without anything Wild Theme put in (its locked content controls).
async function bodyText(bytes) {
  const d = await xml(bytes, "word/document.xml");
  for (const sdt of all(d, "sdt")) {
    const tag = wv(kid(kid(sdt, "sdtPr"), "tag"));
    if (/^WildTheme/.test(tag) && sdt.parentNode) sdt.parentNode.removeChild(sdt);
  }
  return all(d, "p").map((p) => all(p, "t").map((t) => t.textContent).join("")).join("\n");
}
const runWithText = (d, text) => all(d, "r").find((r) => all(r, "t").map((t) => t.textContent).join("") === text);
// Every XML part of a package, canonicalised (an XML re-save that only changes formatting compares equal).
async function xmlParts(bytes) {
  const z = await H.zipOf(bytes), out = {};
  for (const n of Object.keys(z.files).sort()) if (!z.files[n].dir && /\.(xml|rels)$/.test(n)) out[n] = canon(await z.file(n).async("string"));
  return out;
}
async function sameAsOriginal(orig, now, label) {
  assert.deepEqual(await H.names(now), await H.names(orig), `${label}: same parts as the original`);
  const a = await xmlParts(orig), b = await xmlParts(now);
  for (const n of Object.keys(a)) assert.equal(b[n], a[n], `${label}: ${n} is back to the original`);
}

// ------------------------------------------------------------------ what a theme changes
test("report.docx: styles, links, lists, tables and old colours follow the theme", async () => {
  const orig = fixture("report.docx");
  const out = await apply(orig, "synthwave");
  H.save("docx-report-synthwave.docx", out);
  const P = WT.palette(theme("synthwave"));
  assert.ok(out.stats && out.stats.direct.colors >= 1 && out.stats.direct.fonts >= 1 && out.stats.listLevels > 0, JSON.stringify(out.stats));

  // theme part = shared palette
  const sch = WT.schemeOf(await H.text(out, "word/theme/theme1.xml"));
  for (const k of ["dk1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink", "major", "minor"]) assert.equal(sch[k], P[k], k);

  // body font everywhere, text colour from the theme
  const sd = await xml(out, "word/styles.xml");
  const ddRPr = kid(kid(kid(sd.documentElement, "docDefaults"), "rPrDefault"), "rPr");
  assert.equal(wv(kid(ddRPr, "rFonts"), "asciiTheme"), "minorHAnsi");
  // body text colour: text1 on the document defaults, which Normal inherits (a colour on Normal itself
  // would outrank the table header row's white text)
  assert.equal(wv(kid(ddRPr, "color"), "themeColor"), "text1"); assert.equal(wv(kid(ddRPr, "color")), P.dk1);
  const normal = await style(out, "Normal");
  assert.equal(wv(kid(kid(normal, "rPr"), "rFonts"), "asciiTheme"), "minorHAnsi");
  assert.equal(colorOf(normal), null, "Normal inherits the text colour");

  // every themed style uses theme colours (w:themeColor with the matching hex)
  const expect = { Title: ["text2", P.dk2], Subtitle: ["accent1", P.accent1], Heading1: ["accent1", P.accent1], Heading2: ["accent2", P.accent2],
    Heading3: ["text2", P.dk2], Heading4: ["accent1", P.accent1], Heading5: ["accent2", P.accent2], Heading6: ["text2", P.dk2],
    Quote: ["text2", P.dk2], IntenseQuote: ["accent1", P.accent1], Caption: ["text2", P.dk2], Header: ["text2", P.dk2], Footer: ["text2", P.dk2],
    Hyperlink: ["hyperlink", P.hlink], FollowedHyperlink: ["followedHyperlink", P.folHlink], Strong: ["accent1", P.accent1], Emphasis: ["accent2", P.accent2],
    TOCHeading: ["accent1", P.accent1] };
  for (const [id, [slot, hex]] of Object.entries(expect)) {
    const s = await style(out, id);
    assert.ok(s, `style ${id} present`);
    assert.equal(wv(colorOf(s), "themeColor"), slot, `${id} colour slot`);
    assert.equal(wv(colorOf(s)), hex, `${id} colour value`);
  }
  assert.equal(wv(kid(kid(await style(out, "Hyperlink"), "rPr"), "u")), "single");
  // linked character styles stay in step with their paragraph styles
  for (const id of ["Title", "Heading1", "Heading2", "Heading3", "Quote", "IntenseQuote", "Subtitle"]) {
    const p = await style(out, id), c = await style(out, wv(kid(p, "link")));
    assert.ok(c, `${id} has its linked character style`);
    assert.equal(ser(kid(c, "rPr")), ser(kid(p, "rPr")), `${id} Char matches ${id}`);
  }
  assert.equal(wv(colorOf(await style(out, "HeaderChar")), "themeColor"), "text2");

  // the hyperlink and its relationship are untouched
  const rels = await H.text(out, "word/_rels/document.xml.rels");
  assert.match(rels, /Id="rId11"[^>]*Target="https:\/\/example\.com\/"[^>]*TargetMode="External"|Id="rId11"[^>]*TargetMode="External"[^>]*Target="https:\/\/example\.com\/"/);
  const d = await xml(out, "word/document.xml");
  const link = all(d, "hyperlink")[0];
  assert.equal(link.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id"), "rId11");
  assert.equal(all(link, "t")[0].textContent, "the website");

  // hard-coded colours/fonts from a theme -> theme references; other colours stay
  const blue = colorOf(runWithText(d, "hard-coded Office blue"));
  assert.equal(wv(blue, "themeColor"), "accent1"); assert.equal(wv(blue), P.accent1);
  const red = colorOf(runWithText(d, "red text"));
  assert.equal(wv(red), "C00000"); assert.equal(red.getAttributeNS(W, "themeColor"), "", "C00000 is not a theme colour");
  const cal = kid(kid(runWithText(d, "fixed Calibri run"), "rPr"), "rFonts");
  assert.equal(wv(cal, "asciiTheme"), "majorHAnsi", "Calibri is this file's heading (major) font"); assert.equal(wv(cal, "ascii"), "Calibri");

  // bullets and numbers in the accent colour, glyph fonts kept
  const nd = await xml(out, "word/numbering.xml");
  const lvls = all(nd, "lvl");
  assert.ok(lvls.length >= 9);
  for (const l of lvls) {
    assert.equal(wv(colorOf(l), "themeColor"), "accent1", "list level colour");
    if (wv(kid(l, "numFmt")) === "bullet") assert.equal(wv(kid(kid(l, "rPr"), "rFonts"), "ascii"), "Symbol", "bullet glyph font kept");
  }

  // Table Grid: theme borders, accent header row with white bold text, total row rule; new tables use it
  const tg = await style(out, "TableGrid");
  const borders = Array.from(kid(kid(tg, "tblPr"), "tblBorders").childNodes).filter((n) => n.nodeType === 1);
  assert.equal(borders.length, 6);
  for (const b of borders) assert.equal(wv(b, "themeColor"), "accent1", `table border ${b.localName}`);
  const cond = (type) => kids(tg, "tblStylePr").find((x) => wv(x, "type") === type);
  const fr = cond("firstRow");
  assert.ok(fr, "firstRow conditional format");
  assert.equal(wv(kid(kid(fr, "tcPr"), "shd"), "themeFill"), "accent1"); assert.equal(wv(kid(kid(fr, "tcPr"), "shd"), "fill"), P.accent1);
  assert.ok(kid(kid(fr, "rPr"), "b"), "header row bold");
  assert.equal(wv(colorOf(fr), "themeColor"), "background1"); assert.equal(wv(colorOf(fr)), "FFFFFF");
  assert.equal(wv(kid(kid(kid(cond("lastRow"), "tcPr"), "tcBorders"), "top"), "themeColor"), "accent1");
  const st = await xml(out, "word/settings.xml");
  assert.equal(wv(kid(st.documentElement, "defaultTableStyle")), "TableGrid");
  for (const k of ["styleLockTheme", "styleLockQFSet", "displayBackgroundShape"]) assert.ok(kid(st.documentElement, k), `${k} set`);

  // the scene: page background (tinted with the panel colour for apps that only show the colour) + pictures
  const bg = kid(d.documentElement, "background");
  assert.equal(wv(bg, "color"), P.panel);
  const probe = await WWStamper.probe(out);
  assert.equal(probe.marker, "synthwave"); assert.equal(probe.markerName, "Synthwave Sunset");
  assert.equal(probe.themes.length, 1);
  assert.equal(probe.media.length, 1, "one scene picture file");
  assert.ok(probe.pictures >= 4, `page background + page-1 scene + panels (got ${probe.pictures})`);
  assert.match(await H.text(out, "docProps/custom.xml"), /name="WildThemeVersion"><vt:lpwstr>4</);
  assert.match(ser(d), /descr="Wild Theme: synthwave; name=Synthwave Sunset; accent=#FF2BD6; accent2=#21E6FF; mode=dark"/, "alt-text marker kept");
});

test("re-applying and changing theme never stacks; text and links are untouched", async () => {
  const orig = fixture("report.docx");
  const a = await apply(orig, "synthwave");
  const b = await apply(a, "synthwave");
  const c = await apply(b, "mint", { glow: true });
  const count = async (x) => (await styleList(x)).length;
  assert.equal(await count(b), await count(a), "same number of styles after re-applying");
  assert.equal(await count(c), await count(a), "same number of styles after changing theme");
  const ids = (await styleList(c)).map((s) => wv(s, "styleId"));
  assert.equal(new Set(ids).size, ids.length, "no duplicated style ids");
  assert.deepEqual(await H.names(b), await H.names(a), "same part names");
  // stamping is deterministic: same theme, same file -> the same XML (except the core's timestamp)
  for (const p of ["word/document.xml", "word/styles.xml", "word/numbering.xml", "word/settings.xml", "word/header1.xml", "word/header2.xml"]) {
    assert.equal(await H.text(b, p), await H.text(a, p), `${p} identical after re-applying`);
  }
  const text = await bodyText(orig);
  assert.ok(text.includes("hard-coded Office blue") && text.includes("A quote that matters."));
  for (const [x, label] of [[a, "applied"], [b, "re-applied"], [c, "re-themed"]]) {
    assert.equal(await bodyText(x), text, `all body text kept (${label})`);
    assert.match(await H.text(x, "word/_rels/document.xml.rels"), /Id="rId11"[^>]*example\.com/, `hyperlink relationship kept (${label})`);
  }
  // the second theme's colours, not a mix
  const P = WT.palette(theme("mint"));
  assert.equal(wv(colorOf(runWithText(await xml(c, "word/document.xml"), "hard-coded Office blue"))), P.accent1);
  assert.equal((await WT.inspect(c)).status, "intact");
});

// ------------------------------------------------------------------ removing gives the original back
for (const name of ["report.docx", "plain.docx"]) {
  test(`${name}: remove restores the original exactly (theme, styles, lists, settings, document)`, async () => {
    const orig = fixture(name);
    // the core keeps the FIRST backup across re-themes, so the first apply here switches everything off
    const off = { backdrop: false, styles: false, tables: false, lists: false, direct: false, lock: false };
    const themed = await apply(await apply(await apply(orig, "synthwave", off), "aurora"), "mint", { glow: true });
    const r = await WT.remove(themed);
    H.save(`docx-removed-${name}`, r);
    assert.equal(await H.text(r, "word/theme/theme1.xml"), await H.text(orig, "word/theme/theme1.xml"), "theme1.xml byte for byte");
    // the parts the task names, element by element
    const so = await styleList(orig), sr = await styleList(r);
    assert.deepEqual(sr.map(ser), so.map(ser), "styles.xml style elements");
    const dd = async (x) => ser(kid((await xml(x, "word/styles.xml")).documentElement, "docDefaults"));
    assert.equal(await dd(r), await dd(orig), "docDefaults");
    assert.equal(canon(await H.text(r, "word/numbering.xml")), canon(await H.text(orig, "word/numbering.xml")), "numbering.xml");
    // and everything else in the package
    await sameAsOriginal(orig, r, name);
  });
}

test("drawing colours, shading and charts: converted to theme colours and converted back", async () => {
  // plain.docx (Office 2007 theme) plus a shaded run in an old theme colour, a Text 2 run, and a chart part
  const z = await JSZip.loadAsync(fixture("plain.docx"));
  let d = await z.file("word/document.xml").async("string");
  d = d.replace("<w:body>", '<w:body><w:p><w:r><w:rPr><w:color w:val="1F497D"/><w:shd w:val="clear" w:color="auto" w:fill="4F81BD"/></w:rPr><w:t>shaded</w:t></w:r>' +
    '<w:r><w:rPr><w:color w:val="000000"/><w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/></w:rPr><w:t>plain</w:t></w:r></w:p>');
  z.file("word/document.xml", d);
  z.file("word/charts/chart1.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="${A}"><c:chart><c:plotArea>` +
    `<c:spPr><a:solidFill><a:srgbClr val="4472C4"><a:lumMod val="75000"/></a:srgbClr></a:solidFill></c:spPr>` +
    `<c:spPr><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></c:spPr><c:spPr><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></c:spPr></c:plotArea></c:chart></c:chartSpace>`);
  const orig = await z.generateAsync({ type: "uint8array" });
  const out = await apply(orig, "aurora");
  const P = WT.palette(theme("aurora"));
  const od = await xml(out, "word/document.xml");
  const sh = runWithText(od, "shaded"), pl = runWithText(od, "plain");
  assert.equal(wv(kid(kid(sh, "rPr"), "shd"), "themeFill"), "accent1"); assert.equal(wv(kid(kid(sh, "rPr"), "shd"), "fill"), P.accent1);
  assert.equal(wv(colorOf(sh), "themeColor"), "text2"); assert.equal(wv(colorOf(sh)), P.dk2);
  assert.equal(wv(colorOf(pl)), "000000"); assert.equal(colorOf(pl).getAttributeNS(W, "themeColor"), "", "black stays black");
  assert.equal(wv(kid(kid(pl, "rPr"), "shd"), "fill"), "FFFF00");
  const ch = await xml(out, "word/charts/chart1.xml");
  const clrs = Array.from(ch.getElementsByTagNameNS(A, "*")).filter((e) => /^(srgbClr|schemeClr)$/.test(e.localName));
  assert.deepEqual(clrs.map((e) => `${e.localName}:${e.getAttribute("val")}`), ["schemeClr:accent1", "srgbClr:C00000", "schemeClr:accent2"]);
  assert.equal(clrs[0].getElementsByTagNameNS(A, "lumMod").length, 1, "colour adjustments kept");
  assert.ok(out.stats.direct.drawing === 1 && out.stats.direct.fills === 1 && out.stats.direct.colors === 1, JSON.stringify(out.stats.direct));
  await sameAsOriginal(orig, await WT.remove(out), "with chart");
});

// ------------------------------------------------------------------ settings
test("settings: each part can be switched off, glow on", async () => {
  const orig = fixture("report.docx");
  const os = await styleMap(orig);
  const changed = (a, b) => [...b.keys()].filter((k) => a.get(k) !== b.get(k)).concat([...a.keys()].filter((k) => !b.has(k))).sort();

  // styles off (and tables off): styles.xml untouched, the theme part still themed
  let out = await apply(orig, "synthwave", { styles: false, tables: false });
  assert.equal(canon(await H.text(out, "word/styles.xml")), canon(await H.text(orig, "word/styles.xml")), "styles.xml untouched");
  assert.equal(WT.schemeOf(await H.text(out, "word/theme/theme1.xml")).accent1, WT.palette(theme("synthwave")).accent1);
  assert.equal(kid((await xml(out, "word/settings.xml")).documentElement, "defaultTableStyle"), null);

  // styles off, tables on: only Table Grid changes
  out = await apply(orig, "synthwave", { styles: false });
  assert.deepEqual(changed(os, await styleMap(out)), ["TableGrid"]);

  // tables off: Table Grid as it was, no default table style
  out = await apply(orig, "synthwave", { tables: false });
  assert.equal((await styleMap(out)).get("TableGrid"), os.get("TableGrid"));
  assert.equal(kid((await xml(out, "word/settings.xml")).documentElement, "defaultTableStyle"), null);

  // lists off: numbering untouched
  out = await apply(orig, "synthwave", { lists: false });
  assert.equal(canon(await H.text(out, "word/numbering.xml")), canon(await H.text(orig, "word/numbering.xml")));

  // lock off: no theme / style-set locks
  out = await apply(orig, "synthwave", { lock: false });
  let st = (await xml(out, "word/settings.xml")).documentElement;
  assert.equal(kid(st, "styleLockTheme"), null); assert.equal(kid(st, "styleLockQFSet"), null);
  assert.ok(kid(st, "displayBackgroundShape"), "the page background is still shown");

  // direct off: hard-coded colours stay as they are
  out = await apply(orig, "synthwave", { direct: false });
  const blue = colorOf(runWithText(await xml(out, "word/document.xml"), "hard-coded Office blue"));
  assert.equal(wv(blue), "4472C4"); assert.equal(blue.getAttributeNS(W, "themeColor"), "");

  // glow on: Title / Heading 1 / Heading 2 (and their Char styles) glow in the accent; w14 declared + ignorable
  out = await apply(orig, "synthwave", { glow: true });
  const sroot = (await xml(out, "word/styles.xml")).documentElement;
  assert.equal(sroot.getAttribute("xmlns:w14"), W14);
  assert.ok(sroot.getAttribute("mc:Ignorable").split(/\s+/).includes("w14"));
  for (const id of ["Title", "Heading1", "Heading2", "Heading1Char", "TitleChar"]) {
    const g = kid(kid(await style(out, id), "rPr"), "glow", W14);
    assert.ok(g, `${id} glows`);
    assert.equal(kid(g, "schemeClr", W14).getAttributeNS(W14, "val"), "accent3");
  }
  assert.equal(kid(kid(await style(out, "Heading3"), "rPr"), "glow", W14), null);
  const noGlow = await apply(orig, "synthwave", {});
  assert.equal(kid(kid(await style(noGlow, "Title"), "rPr"), "glow", W14), null, "glow is off by default");

  // glow without the text styles: added to the document's own headings only
  out = await apply(orig, "synthwave", { styles: false, tables: false, glow: true });
  assert.deepEqual(changed(os, await styleMap(out)), ["Heading1", "Heading1Char", "Heading2", "Heading2Char", "Title", "TitleChar"]);
  await sameAsOriginal(orig, await WT.remove(out), "glow only");

  // backdrop off: no pictures, no page background, but the theme, styles and marker are there
  out = await apply(orig, "synthwave", { backdrop: false });
  const d = await xml(out, "word/document.xml");
  assert.equal(kid(d.documentElement, "background"), null);
  assert.ok(!(await H.names(out)).some((n) => /wildtheme-(backdrop|panel)/.test(n)), "no scene media");
  assert.equal(kid((await xml(out, "word/settings.xml")).documentElement, "displayBackgroundShape"), null);
  const pr = await WWStamper.probe(out);
  assert.equal(pr.pictures, 0); assert.equal(pr.marker, "synthwave");
  assert.ok((await style(out, "Heading1")) && wv(colorOf(await style(out, "Heading1")), "themeColor") === "accent1");
  await sameAsOriginal(orig, await WT.remove(out), "backdrop off");
});

// ------------------------------------------------------------------ the stamper on its own
test("stamper: deterministic, clean files untouched by unstamp, failures reported", async () => {
  const assets = await makeAssets(theme("galaxy"), "docx", null, { backdrop: true });
  const clean = fixture("report.docx");
  const x = await WWStamper.stamp(clean, theme("galaxy"), assets);       // no settings: defaults
  const y = await WWStamper.stamp(clean, theme("galaxy"), assets);
  assert.deepEqual(await H.names(x), await H.names(y));
  assert.deepEqual(await xmlParts(x), await xmlParts(y), "same file both times");
  assert.deepEqual(x.failures, []);
  assert.deepEqual(JSON.parse(JSON.stringify(x.backup)), x.backup, "backup is plain JSON");
  // the backup describes the file, not the theme or the settings (the core reuses the first one)
  const other = await WWStamper.stamp(clean, theme("lava"), { backdrop: null, panel: null }, { backdrop: false, styles: false, tables: false, lists: false, direct: false, lock: false, glow: true });
  assert.deepEqual(other.backup, x.backup, "same backup for another theme and other settings");
  // stamping a stamped file again (without the core's unstamp first) still doesn't stack
  const z2 = await WWStamper.stamp(x, theme("galaxy"), assets);
  assert.equal((await styleList(z2)).length, (await styleList(x)).length);
  assert.equal((await WWStamper.probe(z2)).pictures, (await WWStamper.probe(x)).pictures);
  // unstamp with the backup = the original; on a file without a theme it changes nothing at all
  await sameAsOriginal(clean, await WWStamper.unstamp(x, x.backup), "unstamp(backup)");
  const same = await WWStamper.unstamp(clean);
  assert.deepEqual(await xmlParts(same), await xmlParts(clean));
  assert.equal((await WWStamper.probe(clean)).marker, null);
  // asked for a scene but given no picture: reported, not silently skipped
  const noPic = await WWStamper.stamp(clean, theme("galaxy"), { backdrop: null, panel: null }, { backdrop: true });
  assert.equal(noPic.failures.length, 1);
  await assert.rejects(WT.apply(clean, theme("galaxy"), {}, async () => ({ backdrop: null, panel: null }), "test"), /partly applied/);
});

test("legacy v3 file: recognised by probe, unstamp without a backup removes the theme pieces", async () => {
  const dir = path.join(__dirname, "..", "documents");
  const file = fs.readdirSync(dir).find((f) => f.endsWith(".docx"));
  const v3 = new Uint8Array(fs.readFileSync(path.join(dir, file)));
  const p = await WWStamper.probe(v3);
  assert.ok(p.marker && p.pictures >= 2, JSON.stringify({ marker: p.marker, pictures: p.pictures }));
  const r = await WWStamper.unstamp(v3);
  const pr = await WWStamper.probe(r);
  assert.equal(pr.marker, null); assert.equal(pr.pictures, 0);
  assert.ok(!(await H.names(r)).some((n) => /wildtheme|custom\.xml/.test(n)));
  assert.doesNotMatch(await H.text(r, "word/theme/theme1.xml"), /name="Wild Theme /);
  assert.ok((await bodyText(r)).includes("Your Title Here"), "text kept");
});

// ------------------------------------------------------------------ another editor: LibreOffice
test("LibreOffice round trip: the theme is still recognised, and can be removed", { skip: !H.hasSoffice() && "LibreOffice not installed" }, async (t) => {
  const orig = fixture("report.docx");
  const a = await apply(orig, "synthwave");
  const lo = H.libreRoundTrip("docx-lo-synthwave.docx", a);
  const i = await WT.inspect(lo);
  assert.equal(i.theme.id, "synthwave", "theme found in the embedded definition");
  assert.ok(["intact", "partial"].includes(i.status), `${i.status} ${JSON.stringify(i.problems)}`);
  assert.ok(i.hasBackup);
  assert.ok(!i.problems.some((p) => p.part === "theme"), "LibreOffice keeps the theme colours and fonts");
  assert.equal(i.probe.marker, "synthwave", "custom properties kept");
  // What LibreOffice 24.2 really does, noted rather than required (it can change between versions):
  const ld = await xml(lo, "word/document.xml"), lbg = kid(ld.documentElement, "background");
  const ls = (await xml(lo, "word/settings.xml")).documentElement;
  t.diagnostic(`LibreOffice: status ${i.status}; page background picture ${lbg && all(lbg, "fill", "urn:schemas-microsoft-com:vml").length ? "kept" : "dropped (colour kept)"}; ` +
    `styleLockTheme ${kid(ls, "styleLockTheme") ? "kept" : "dropped"}; defaultTableStyle ${kid(ls, "defaultTableStyle") ? "kept" : "dropped"}; ` +
    `header parts ${(await H.names(lo)).filter((n) => /word\/header\d+\.xml$/.test(n)).length} (was ${(await H.names(a)).filter((n) => /word\/header\d+\.xml$/.test(n)).length})`);
  // newer LibreOffice writes "auto" here; Word keeps the colour (checked in Word 365)
  if (lbg && wv(lbg, "color") !== "auto") assert.equal(wv(lbg, "color"), WT.palette(theme("synthwave")).panel, "page colour fallback survives");

  // remove after the round trip: nothing of ours left, the original theme and styles back
  const r = await WT.remove(lo);
  assert.equal((await WT.inspect(r)).status, "none");
  const z = await H.zipOf(r);
  for (const n of Object.keys(z.files)) {
    if (z.files[n].dir || !/\.(xml|rels)$/.test(n)) continue;
    assert.doesNotMatch(await z.file(n).async("string"), /WildTheme|Wild Theme|wildtheme/, `nothing of ours left in ${n}`);
  }
  assert.deepEqual(WT.schemeOf(await H.text(r, "word/theme/theme1.xml")), WT.schemeOf(await H.text(orig, "word/theme/theme1.xml")));
  assert.equal((await styleMap(r)).get("Heading1"), (await styleMap(orig)).get("Heading1"), "Heading 1 back to the original");
  // (LibreOffice itself merges the empty paragraph that carries a section break, so compare non-empty paragraphs)
  const nonEmpty = async (x) => (await bodyText(x)).split("\n").filter(Boolean).join("\n");
  assert.equal(await nonEmpty(r), await nonEmpty(orig), "text kept");
  // LibreOffice bakes table styles into the cells, so the header-row fill stays (it follows the restored theme in Word)
  const firstCell = all(all(await xml(r, "word/document.xml"), "tbl")[0], "tc")[0];
  t.diagnostic(`after remove: first table cell fill ${wv(kid(kid(firstCell, "tcPr"), "shd"), "fill") || "none"} (${wv(kid(kid(firstCell, "tcPr"), "shd"), "themeFill") || "no theme fill"})`);
});

for (const id of ["synthwave", "mint"]) {
  test(`LibreOffice render (${id}): dark scene at the page corner, light panel behind the text`, { skip: !H.hasSoffice() && "LibreOffice not installed" }, async (t) => {
    const out = await apply(fixture("report.docx"), id);
    const png = H.render(`docx-render-${id}.docx`, out);
    const corner = H.sample(png, [0.005, 0.005, 0.04, 0.04]), panel = H.sample(png, [0.45, 0.78, 0.8, 0.9]);
    assert.ok(H.lum(corner) < 80, `corner is the dark scene: ${corner}`);
    assert.ok(H.lum(panel) > 200, `panel is light: ${panel}`);
    const plain = H.render(`docx-render-plain-${id}.docx`, await apply(fixture("plain.docx"), id));
    assert.ok(H.lum(H.sample(plain, [0.4, 0.4, 0.6, 0.6])) > 200, "page centre is the light panel");
    assert.ok(H.lum(H.sample(plain, [0.005, 0.005, 0.04, 0.04])) < 80, "page corner is the scene");
    // pages after the first: LibreOffice shows the page colour, not the page picture (Word shows the scene)
    const p2 = H.render(`docx-render-p2-${id}.docx`, out, 2);
    t.diagnostic(`page 2 corner in LibreOffice: ${H.sample(p2, [0.005, 0.005, 0.04, 0.04])} (theme panel colour #${WT.palette(theme(id)).panel})`);
  });
}
