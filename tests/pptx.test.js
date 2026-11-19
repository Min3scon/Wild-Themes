// PowerPoint stamper: deep theming (theme, text styles, table style, fixed colours, transitions, scene) and exact removal.
const test = require("node:test"), assert = require("node:assert/strict");
const path = require("path");
const { DOMParser, XMLSerializer } = require("@xmldom/xmldom");
const H = require("./helpers.js");
const { WT, theme, fixture, makeAssets } = H;
const S = require(path.join(__dirname, "..", "extension", "lib", "stamper-pptx.js"));

const MEDIUM2_ACCENT1 = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}";
const NO_STYLE = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
const OURS = { fade: '<p:transition spd="med"><p:fade/></p:transition>', push: '<p:transition spd="med"><p:push dir="u"/></p:transition>',
  wipe: '<p:transition spd="med"><p:wipe dir="r"/></p:transition>', split: '<p:transition spd="med"><p:split orient="vert" dir="out"/></p:transition>',
  dissolve: '<p:transition spd="med"><p:dissolve/></p:transition>' };
const SLIDE3_WIPE = '<p:transition spd="slow"><p:wipe dir="d"/></p:transition>';

// ------------------------------------------------------------------ helpers
const apply = (bytes, settings = {}, id = "synthwave") => WT.apply(bytes, theme(id), settings, makeAssets, "test");
async function edit(bytes, fn) { const z = await H.zipOf(bytes); await fn(z); return new Uint8Array(await z.generateAsync({ type: "uint8array" })); }
const slide = (bytes, n) => H.text(bytes, `ppt/slides/slide${n}.xml`);
const MASTER = "ppt/slideMasters/slideMaster1.xml";
const texts = (xml) => [...xml.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((m) => m[1]);
const transitions = (xml) => xml.match(/<p:transition\b[\s\S]*?<\/p:transition>/g) || [];
const timing = (xml) => (xml.match(/<p:timing>[\s\S]*<\/p:timing>/) || [null])[0];
const parse = (xml) => new DOMParser().parseFromString(xml, "application/xml");
// canonical serialisation of a whole part (the XML declaration's quoting doesn't count)
const canonDoc = (xml) => new XMLSerializer().serializeToString(parse(xml.replace(/^<\?xml[^>]*\?>\s*/, "")));
// canonical serialisation of the first element with this tag (namespace declarations as in the whole document)
const canon = (xml, tag) => { const el = parse(xml).getElementsByTagName(tag)[0]; return el ? new XMLSerializer().serializeToString(el) : null; };
const tblDef = async (bytes) => { const x = await H.text(bytes, "ppt/tableStyles.xml"); return x ? parse(x).documentElement.getAttribute("def") : null; };
const used = async (bytes) => (await WT.inspect(bytes)).probe;
const rels = (list) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  list.map(([id, type, target]) => `<Relationship Id="${id}" Type="${R}${type}" Target="${target}"/>`).join("") + "</Relationships>";

// deck.pptx plus: a table default that isn't theme-following, a theme colour with a modifier and a non-theme colour on
// slide 2, a theme colour in the chart, and a user transition on slide 5 that looks exactly like one of ours.
const variant = () => edit(fixture("deck.pptx"), async (z) => {
  const ts = await z.file("ppt/tableStyles.xml").async("string");
  z.file("ppt/tableStyles.xml", ts.replace(MEDIUM2_ACCENT1, NO_STYLE));
  let s2 = await z.file("ppt/slides/slide2.xml").async("string");
  s2 = s2.replace('<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></p:spPr>',
    '<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="4F81BD"><a:lumMod val="75000"/></a:srgbClr></a:solidFill></a:ln></p:spPr>');
  s2 = s2.replace("<a:r><a:t>Callout</a:t></a:r>", '<a:r><a:rPr lang="en-US"><a:solidFill><a:srgbClr val="123456"/></a:solidFill></a:rPr><a:t>Callout</a:t></a:r>');
  z.file("ppt/slides/slide2.xml", s2);
  const c = await z.file("ppt/charts/chart1.xml").async("string");
  const i = c.indexOf("</c:tx>", c.indexOf("<c:idx val=\"1\"/>"));
  z.file("ppt/charts/chart1.xml", c.slice(0, i + 7) + '<c:spPr><a:solidFill><a:srgbClr val="C0504D"/></a:solidFill></c:spPr>' + c.slice(i + 7));
  const s5 = await z.file("ppt/slides/slide5.xml").async("string");
  z.file("ppt/slides/slide5.xml", s5.replace("</p:clrMapOvr>", "</p:clrMapOvr>" + OURS.fade));
});
// A second slide master (+ layout + theme) copied from the ORIGINAL deck; optionally slide 5 uses it.
const withSecondMaster = (bytes, useForSlide5) => edit(bytes, async (z) => {
  const o = await H.zipOf(fixture("deck.pptx")), get = (p) => o.file(p).async("string");
  z.file("ppt/theme/theme2.xml", await get("ppt/theme/theme1.xml"));
  z.file("ppt/slideMasters/slideMaster2.xml", (await get(MASTER)).replace(/<p:sldLayoutIdLst>[\s\S]*?<\/p:sldLayoutIdLst>/, '<p:sldLayoutIdLst><p:sldLayoutId id="2147483661" r:id="rId1"/></p:sldLayoutIdLst>'));
  z.file("ppt/slideMasters/_rels/slideMaster2.xml.rels", rels([["rId1", "slideLayout", "../slideLayouts/slideLayout12.xml"], ["rId2", "theme", "../theme/theme2.xml"]]));
  z.file("ppt/slideLayouts/slideLayout12.xml", await get("ppt/slideLayouts/slideLayout7.xml"));
  z.file("ppt/slideLayouts/_rels/slideLayout12.xml.rels", rels([["rId1", "slideMaster", "../slideMasters/slideMaster2.xml"]]));
  const pres = await z.file("ppt/presentation.xml").async("string");
  z.file("ppt/presentation.xml", pres.replace("</p:sldMasterIdLst>", '<p:sldMasterId id="2147483660" r:id="rIdM2"/></p:sldMasterIdLst>'));
  const pr = await z.file("ppt/_rels/presentation.xml.rels").async("string");
  z.file("ppt/_rels/presentation.xml.rels", pr.replace("</Relationships>", `<Relationship Id="rIdM2" Type="${R}slideMaster" Target="slideMasters/slideMaster2.xml"/></Relationships>`));
  const ct = await z.file("[Content_Types].xml").async("string");
  z.file("[Content_Types].xml", ct.replace("</Types>", [["ppt/slideMasters/slideMaster2.xml", "presentationml.slideMaster+xml"],
    ["ppt/slideLayouts/slideLayout12.xml", "presentationml.slideLayout+xml"], ["ppt/theme/theme2.xml", "theme+xml"]]
    .map(([p, t]) => `<Override PartName="/${p}" ContentType="application/vnd.openxmlformats-officedocument.${t}"/>`).join("") + "</Types>"));
  if (useForSlide5) {
    const r5 = await z.file("ppt/slides/_rels/slide5.xml.rels").async("string");
    z.file("ppt/slides/_rels/slide5.xml.rels", r5.replace("../slideLayouts/slideLayout7.xml", "../slideLayouts/slideLayout12.xml"));
  }
});

// ------------------------------------------------------------------ tests
test("pptx: content, animation and existing transition kept; other slides get the chosen transition", async () => {
  const orig = fixture("deck.pptx");
  const a = await apply(orig);
  H.save("pptx-themed.pptx", a);
  for (let n = 1; n <= 5; n++) assert.deepEqual(texts(await slide(a, n)), texts(await slide(orig, n)), `slide ${n} text`);
  const o2 = await slide(orig, 2), a2 = await slide(a, 2);
  assert.ok(timing(o2), "fixture has an animation");
  assert.equal(timing(a2), timing(o2), "slide 2 animation byte-identical");
  assert.deepEqual(transitions(await slide(a, 3)), [SLIDE3_WIPE], "slide 3 keeps its own wipe");
  for (const n of [1, 2, 4, 5]) assert.deepEqual(transitions(await slide(a, n)), [OURS.fade], `slide ${n} gets the fade`);
  // schema position: after p:clrMapOvr, before p:timing
  assert.ok(a2.indexOf("</p:clrMapOvr>") < a2.indexOf("<p:transition") && a2.indexOf("<p:transition") < a2.indexOf("<p:timing>"));
  // charts, tables and pictures are still there and untouched
  const zo = await H.zipOf(orig), za = await H.zipOf(a);
  for (const p of ["ppt/charts/chart1.xml", "ppt/embeddings/Microsoft_Excel_Sheet1.xlsx", "ppt/media/image1.png"]) {
    assert.deepEqual(await za.file(p).async("uint8array"), await zo.file(p).async("uint8array"), p);
  }
  assert.equal(canon(await slide(a, 4), "a:tbl"), canon(await slide(orig, 4), "a:tbl"), "table unchanged");
  assert.match(await slide(a, 5), /<p:pic>/, "slide 5 picture kept");
});

test("pptx: every transition choice, and 'none' adds nothing", async () => {
  for (const kind of ["push", "wipe", "split", "dissolve"]) {
    const a = await apply(fixture("deck.pptx"), { transition: kind });
    for (const n of [1, 2, 4, 5]) assert.deepEqual(transitions(await slide(a, n)), [OURS[kind]], `${kind} on slide ${n}`);
    assert.deepEqual(transitions(await slide(a, 3)), [SLIDE3_WIPE]);
  }
  const none = await apply(fixture("deck.pptx"), { transition: "none" });
  for (const n of [1, 2, 4, 5]) assert.deepEqual(transitions(await slide(none, n)), [], `no transition on slide ${n}`);
  assert.deepEqual(transitions(await slide(none, 3)), [SLIDE3_WIPE]);
  // a transition wrapped in mc:AlternateContent (as PowerPoint saves newer effects) counts as existing
  const alt = await edit(fixture("deck.pptx"), async (z) => {
    const s1 = await z.file("ppt/slides/slide1.xml").async("string");
    z.file("ppt/slides/slide1.xml", s1.replace("<p:sld ", '<p:sld xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" ')
      .replace("</p:clrMapOvr>", '</p:clrMapOvr><mc:AlternateContent><mc:Choice Requires="p14"><p:transition spd="slow" p14:dur="2000"><p14:vortex dir="r"/></p:transition></mc:Choice><mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>'));
  });
  const b = await apply(alt);
  const s1 = await slide(b, 1);
  assert.equal(transitions(s1).length, 2, "only the existing (choice + fallback) transitions");
  assert.match(s1, /p14:vortex/);
  const back = await WT.remove(b);
  assert.equal(await slide(back, 1), await slide(alt, 1), "an existing AlternateContent transition survives removal");
});

test("pptx: theme colours, text styles, table style and fixed colours follow the theme", async () => {
  const v = await variant();
  const a = await apply(v);
  H.save("pptx-variant-themed.pptx", a);
  const pal = WT.palette(theme("synthwave"));
  const sc = WT.schemeOf(await H.text(a, "ppt/theme/theme1.xml"));
  for (const k of ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink", "major", "minor"]) assert.equal(sc[k], pal[k], k);
  // text styles
  const m = await H.text(a, MASTER);
  const title = canon(m, "p:titleStyle"), body = canon(m, "p:bodyStyle"), other = canon(m, "p:otherStyle");
  assert.match(title, /<a:solidFill><a:schemeClr val="accent1"\/><\/a:solidFill><a:latin typeface="\+mj-lt"\/>/);
  assert.match(body, /<a:buClr><a:schemeClr val="accent1"\/><\/a:buClr>/);
  assert.equal((body.match(/<a:buClr>/g) || []).length, 9, "every body level's bullet");
  assert.equal((body.match(/<a:latin typeface="\+mn-lt"\/>/g) || []).length, 9);
  assert.match(other, /<a:schemeClr val="tx1"\/><\/a:solidFill><a:latin typeface="\+mn-lt"\/>/);
  // table style
  assert.equal(await tblDef(a), MEDIUM2_ACCENT1);
  // fixed colours: theme ones become theme colours (modifiers kept), others stay
  const s2 = await slide(a, 2);
  assert.match(s2, /<a:solidFill><a:schemeClr val="accent1"\/><\/a:solidFill><a:ln>/, "4472C4 callout -> accent1");
  assert.match(s2, /<a:ln><a:solidFill><a:schemeClr val="accent1"><a:lumMod val="75000"\/><\/a:schemeClr>/, "modifier kept");
  assert.match(s2, /<a:srgbClr val="123456"\/>/, "non-theme colour stays");
  assert.doesNotMatch(s2, /4472C4|4F81BD/);
  assert.match(await H.text(a, "ppt/charts/chart1.xml"), /<c:spPr><a:solidFill><a:schemeClr val="accent2"\/>/, "chart colour -> accent2");
  assert.equal(a.stats.direct, 3);
  // slide 5 had a transition of its own (identical to ours): kept, not doubled
  assert.deepEqual(transitions(await slide(a, 5)), [OURS.fade]);
});

test("pptx: remove restores theme, text styles, table style, colours, and only our transitions", async () => {
  const v = await variant();
  let a = await apply(v, { transition: "push" });
  a = await apply(a, { transition: "fade" }, "winter");                // re-theme: the backup still describes the original
  // the user changes our transition on slide 4 afterwards: it's theirs now
  a = await edit(a, async (z) => {
    const s4 = await z.file("ppt/slides/slide4.xml").async("string");
    z.file("ppt/slides/slide4.xml", s4.replace(OURS.fade, '<p:transition spd="fast"><p:cover dir="l"/></p:transition>'));
  });
  const r = await WT.remove(a);
  H.save("pptx-variant-removed.pptx", r);
  assert.equal((await WT.inspect(r)).status, "none");
  assert.equal(await H.text(r, "ppt/theme/theme1.xml"), await H.text(v, "ppt/theme/theme1.xml"), "theme1.xml byte for byte");
  const mo = await H.text(v, MASTER), mr = await H.text(r, MASTER);
  assert.equal(canon(mr, "p:txStyles"), canon(mo, "p:txStyles"), "txStyles exactly");
  assert.equal(canon(mr, "p:sldMaster"), canon(mo, "p:sldMaster"), "whole master exactly");
  assert.equal(await tblDef(r), NO_STYLE, "table style default restored");
  assert.equal(canon(await slide(r, 2), "p:sld"), canon(await slide(v, 2), "p:sld"), "slide 2 colours exactly");
  assert.equal(canon(await H.text(r, "ppt/charts/chart1.xml"), "c:chartSpace"), canon(await H.text(v, "ppt/charts/chart1.xml"), "c:chartSpace"));
  for (const n of [1, 2]) assert.deepEqual(transitions(await slide(r, n)), [], `our transition removed from slide ${n}`);
  assert.deepEqual(transitions(await slide(r, 3)), [SLIDE3_WIPE], "slide 3's own wipe kept");
  assert.deepEqual(transitions(await slide(r, 4)), ['<p:transition spd="fast"><p:cover dir="l"/></p:transition>'], "a transition the user changed is kept");
  assert.deepEqual(transitions(await slide(r, 5)), [OURS.fade], "slide 5's own transition kept although it looks like ours");
  assert.deepEqual(await H.names(r), await H.names(v), "same parts as the original");
  assert.doesNotMatch(await H.text(r, "[Content_Types].xml"), /Extension="gif"/, "content type added for the GIF is gone");
});

test("pptx: styles, tables and direct off leave those parts untouched", async () => {
  const v = await variant();
  const a = await apply(v, { styles: false, tables: false, direct: false });
  assert.equal(canon(await H.text(a, MASTER), "p:txStyles"), canon(await H.text(v, MASTER), "p:txStyles"));
  assert.equal(await tblDef(a), NO_STYLE);
  const s2 = await slide(a, 2);
  assert.match(s2, /<a:srgbClr val="4472C4"\/>/);
  assert.match(s2, /<a:srgbClr val="4F81BD"><a:lumMod val="75000"\/><\/a:srgbClr>/);
  assert.match(await H.text(a, "ppt/charts/chart1.xml"), /<a:srgbClr val="C0504D"\/>/);
  assert.equal(WT.schemeOf(await H.text(a, "ppt/theme/theme1.xml")).accent1, WT.palette(theme("synthwave")).accent1, "theme itself still applied");
  assert.equal((await WT.inspect(a)).status, "intact");
  // switching them on later and removing still gives the original back
  const b = await apply(a, {});
  assert.equal(await tblDef(b), MEDIUM2_ACCENT1);
  const r = await WT.remove(b);
  assert.equal(canon(await H.text(r, MASTER), "p:sldMaster"), canon(await H.text(v, MASTER), "p:sldMaster"));
  assert.equal(await tblDef(r), NO_STYLE);
  assert.equal(canon(await slide(r, 2), "p:sld"), canon(await slide(v, 2), "p:sld"));
});

test("pptx: a deck without tableStyles.xml gets one, and loses it again on removal", async () => {
  const bare = await edit(fixture("deck.pptx"), async (z) => {
    z.remove("ppt/tableStyles.xml");
    const pr = await z.file("ppt/_rels/presentation.xml.rels").async("string");
    z.file("ppt/_rels/presentation.xml.rels", pr.replace(/<Relationship [^>]*tableStyles[^>]*\/>/, ""));
    const ct = await z.file("[Content_Types].xml").async("string");
    z.file("[Content_Types].xml", ct.replace(/<Override PartName="\/ppt\/tableStyles.xml"[^>]*\/>/, ""));
  });
  const a = await apply(bare);
  assert.equal(await tblDef(a), MEDIUM2_ACCENT1);
  assert.match(await H.text(a, "ppt/_rels/presentation.xml.rels"), /relationships\/tableStyles" Target="tableStyles.xml"/);
  assert.match(await H.text(a, "[Content_Types].xml"), /PartName="\/ppt\/tableStyles.xml"/);
  const r = await WT.remove(a);
  assert.deepEqual(await H.names(r), await H.names(bare));
  assert.doesNotMatch(await H.text(r, "ppt/_rels/presentation.xml.rels"), /tableStyles/);
  assert.doesNotMatch(await H.text(r, "[Content_Types].xml"), /tableStyles/);
});

test("pptx: scene off, animated GIF or still picture; probe and markers", async () => {
  const off = await apply(fixture("deck.pptx"), { backdrop: false });
  const po = await used(off);
  assert.equal(po.pictures, 0);
  assert.equal(po.media.length, 0);
  assert.ok(!(await H.names(off)).some((n) => /wildtheme/.test(n)), "no scene media");
  assert.equal(po.marker, "synthwave", "the file property still names the theme");
  assert.match(canon(await H.text(off, MASTER), "p:titleStyle"), /schemeClr val="accent1"/, "styles still applied");

  const gif = await apply(fixture("deck.pptx"), { animated: true });
  const pg = await S.probe(gif);
  assert.equal(pg.pictures, 2, "backdrop + reading panel on the one master");
  const gifBytes = (await makeAssets(theme("synthwave"), "pptx", null, WT.settingsFor("pptx", {}))).backdrop.bytes;
  assert.deepEqual(pg.media, [gifBytes], "the animated GIF, unchanged");
  assert.equal(pg.marker, "synthwave"); assert.equal(pg.markerName, "Synthwave Sunset");
  assert.equal(pg.themes.length, 1);
  assert.match(await H.text(gif, "[Content_Types].xml"), /Extension="gif" ContentType="image\/gif"/);
  assert.match(await H.text(gif, "docProps/custom.xml"), /name="WildThemeVersion"><vt:lpwstr>4</);
  const m = await H.text(gif, MASTER);
  assert.ok(m.includes(`descr="${S.markerDescr(theme("synthwave")).replace(/&/g, "&amp;")}"`), "alt-text marker for PowerPoint for the web");
  assert.ok(m.indexOf("WildTheme Backdrop") < m.indexOf("Title Placeholder"), "scene behind everything");
  assert.match(m, /<a:picLocks [^>]*noSelect="1"/);

  const still = await apply(fixture("deck.pptx"), { animated: false });
  assert.ok((await H.names(still)).some((n) => /^ppt\/media\/wildtheme-backdrop-[0-9a-f]+\.jpeg$/.test(n)));
  assert.ok(!(await H.names(still)).some((n) => /\.gif$/.test(n)));

  // without the file property the alt text still identifies the theme
  const noProps = await edit(gif, async (z) => z.file("docProps/custom.xml", (await z.file("docProps/custom.xml").async("string")).replace(/<property [^>]*name="WildTheme(Name)?"[^>]*>.*?<\/property>/g, "")));
  const pn = await S.probe(noProps);
  assert.equal(pn.marker, "synthwave"); assert.equal(pn.markerName, "Synthwave Sunset");
});

test("pptx: deterministic, idempotent, and unstamp leaves an unthemed file alone", async () => {
  const orig = fixture("deck.pptx");
  const opts = WT.settingsFor("pptx", {});
  const assets = await makeAssets(theme("synthwave"), "pptx", (await S.analyze(orig)).sections[0], opts);
  const x = await S.stamp(orig, theme("synthwave"), assets, opts), y = await S.stamp(orig, theme("synthwave"), assets, opts);
  assert.deepEqual(await H.names(x), await H.names(y));
  for (const n of await H.names(x)) assert.equal(await H.text(x, n), await H.text(y, n), n);
  assert.deepEqual(x.failures, []);
  assert.deepEqual(JSON.parse(JSON.stringify(x.backup)), x.backup, "backup is JSON-serialisable");
  // stamping its own output again doesn't stack anything
  const z = await S.stamp(x, theme("synthwave"), assets, opts);
  assert.deepEqual(await H.names(z), await H.names(x));
  assert.equal((await S.probe(z)).pictures, 2);
  for (let n = 1; n <= 5; n++) assert.equal(transitions(await slide(z, n)).length, 1);
  // unstamp with its backup -> the original, part for part
  const back = await S.unstamp(x, x.backup);
  assert.deepEqual(await H.names(back), await H.names(orig));
  for (const n of (await H.names(orig)).filter((p) => /\.(xml|rels)$/.test(p))) {
    assert.equal(canonDoc(await H.text(back, n)), canonDoc(await H.text(orig, n)), n);
  }
  // unstamp without a backup on an unthemed file: nothing changes
  const same = await S.unstamp(orig);
  assert.deepEqual(await H.names(same), await H.names(orig));
  for (const n of (await H.names(orig)).filter((p) => /\.(xml|rels)$/.test(p))) {
    assert.equal(canonDoc(await H.text(same, n)), canonDoc(await H.text(orig, n)), n);
  }
});

test("pptx: an unused master with the original theme doesn't count; a second used master is themed and restored", async () => {
  // the Office add-in leaves the old master behind, unused
  const a = await apply(fixture("deck.pptx"));
  const extra = await withSecondMaster(a, false);
  const i = await WT.inspect(extra);
  assert.equal(i.status, "intact", JSON.stringify(i.problems));
  assert.equal(i.probe.themes.length, 1);
  // a deck whose slide 5 uses a second master: both masters themed, both restored exactly
  const two = await withSecondMaster(fixture("deck.pptx"), true);
  const b = await apply(two);
  H.save("pptx-two-masters.pptx", b);
  const ib = await WT.inspect(b);
  assert.equal(ib.status, "intact", JSON.stringify(ib.problems));
  assert.equal(ib.probe.pictures, 4, "scene on both masters");
  for (const t of ["theme1", "theme2"]) assert.equal(WT.schemeOf(await H.text(b, `ppt/theme/${t}.xml`)).accent1, WT.palette(theme("synthwave")).accent1, t);
  assert.match(canon(await H.text(b, "ppt/slideMasters/slideMaster2.xml"), "p:bodyStyle"), /<a:buClr>/);
  const r = await WT.remove(b);
  for (const p of ["ppt/theme/theme1.xml", "ppt/theme/theme2.xml"]) assert.equal(await H.text(r, p), await H.text(two, p), p);
  for (const p of [MASTER, "ppt/slideMasters/slideMaster2.xml"]) assert.equal(canon(await H.text(r, p), "p:sldMaster"), canon(await H.text(two, p), "p:sldMaster"), p);
});

test("pptx: LibreOffice round trip keeps the theme, and it can still be removed", { skip: !H.hasSoffice() && "LibreOffice not installed" }, async () => {
  const orig = fixture("deck.pptx");
  const a = await apply(orig, { animated: false });
  const lo = H.libreRoundTrip("pptx-lo.pptx", a);
  const i = await WT.inspect(lo);
  // LibreOffice keeps the custom XML part and the file properties (an older one may drop the custom XML -> "legacy");
  // it splits the master into one per layout (identical copies) and renames the media.
  assert.ok(["intact", "legacy"].includes(i.status), i.status + " " + JSON.stringify(i.problems));
  assert.equal(i.theme.id, "synthwave");
  assert.equal(i.probe.marker, "synthwave", "file property survives");
  assert.ok(i.probe.pictures >= 2, "scene pictures survive");
  for (const x of i.probe.themes) assert.equal(WT.schemeOf(x).accent1, WT.palette(theme("synthwave")).accent1);
  assert.equal(timing(await slide(lo, 2)) !== null, true, "animation survives");
  assert.match(transitions(await slide(lo, 1)).join(""), /<p:fade\/>/, "our fade survives");
  if (i.status === "intact") {
    const r = await WT.remove(lo);
    const ir = await WT.inspect(r);
    assert.equal(ir.status, "none");
    assert.equal(ir.probe.pictures, 0);
    for (const x of ir.probe.themes) assert.equal(WT.schemeOf(x).accent1, "4F81BD", "original theme on every master copy");
    for (const n of [1, 2, 4, 5]) assert.deepEqual(transitions(await slide(r, n)), [], `slide ${n}`);
    assert.deepEqual(transitions(await slide(r, 3)), [SLIDE3_WIPE]);
    assert.match(await slide(r, 2), /<a:srgbClr val="4472C4"\/>/);
  }
});

test("pptx: slide 1 renders with the dark scene at the edges and the light panel in the middle", { skip: !H.hasSoffice() && "LibreOffice not installed" }, async () => {
  const a = await apply(fixture("deck.pptx"));
  const png = H.render("pptx-render.pptx", a, 1);
  for (const box of [[0, 0, 0.025, 1], [0.975, 0, 1, 1], [0, 0, 1, 0.02], [0, 0.98, 1, 1]]) assert.ok(H.lum(H.sample(png, box)) < 90, `edge ${box} is the dark scene: ${H.sample(png, box)}`);
  for (const box of [[0.1, 0.15, 0.3, 0.35], [0.7, 0.65, 0.9, 0.85]]) assert.ok(H.lum(H.sample(png, box)) > 180, `centre ${box} is the light panel: ${H.sample(png, box)}`);
  const plain = H.render("pptx-render-plain.pptx", fixture("deck.pptx"), 1);
  assert.ok(H.lum(H.sample(plain, [0, 0, 0.025, 1])) > 200, "the unthemed slide is white there");
});
