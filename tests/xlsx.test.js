// Excel (.xlsx / .xlsm): deep, reversible theming. What the stamper changes, what it must never
// touch (data, formulas, tables, charts' data, validation, conditional formats, names, VBA), and
// how exactly it comes back out — through the shared core and directly.
const test = require("node:test"), assert = require("node:assert/strict");
const fs = require("fs"), path = require("path"), { spawnSync } = require("child_process");
const { DOMParser, XMLSerializer } = require("@xmldom/xmldom");
const H = require("./helpers.js");
const { WT, theme, fixture, makeAssets } = H;
const S = require(path.join(__dirname, "..", "extension", "lib", "stamper-xlsx.js"));

const SNS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const ANS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const CNS = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const parse = (x) => new DOMParser().parseFromString(x, "application/xml");
const ser = (n) => new XMLSerializer().serializeToString(n);
const canon = (x) => (x == null ? null : ser(parse(x)).replace(/^<\?xml[^>]*\?>\s*/, ""));
const els = (x, local, ns = SNS) => Array.from(parse(x).getElementsByTagNameNS(ns, local));
const kids = (el, local) => Array.from(el.childNodes).filter((n) => n.nodeType === 1 && (!local || n.localName === local));
const FIXTURES = ["book.xlsx", "macro.xlsm"];
const SHEETS = ["xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml"];

async function parts(bytes) {
  const z = await H.zipOf(bytes), out = new Map();
  for (const n of Object.keys(z.files)) if (!z.files[n].dir) out.set(n, Buffer.from(await z.file(n).async("uint8array")));
  return out;
}
// every cell: ref, type, style, formula markup, cached value, inline text
const cells = (x) => els(x, "c").map((c) => {
  const f = kids(c, "f")[0], v = kids(c, "v")[0], is = kids(c, "is")[0];
  return [c.getAttribute("r"), c.getAttribute("t"), c.getAttribute("s"), f ? ser(f) : null, v ? v.textContent : null, is ? is.textContent : null];
});
const sheetData = (x) => x.match(/<sheetData[\s\S]*<\/sheetData>/)[0];
const block = (x, local) => els(x, local).map(ser);
const styles = async (bytes) => parse(await H.text(bytes, "xl/styles.xml")).documentElement;
const list = (root, name) => kids(kids(root, name)[0] || root, { fonts: "font", fills: "fill", borders: "border", cellStyleXfs: "xf", cellXfs: "xf", cellStyles: "cellStyle" }[name]);
const fontOfCell = (root, s) => list(root, "fonts")[Number(list(root, "cellXfs")[s].getAttribute("fontId"))];
const styleXf = (root, name) => list(root, "cellStyleXfs")[Number(list(root, "cellStyles").find((c) => c.getAttribute("name") === name).getAttribute("xfId"))];
const attr = (el, a) => (el && el.hasAttribute(a) ? el.getAttribute(a) : null);
const tabOf = (x) => (x.match(/<tabColor[^>]*\/>/) || [null])[0];
const opts = (o) => WT.settingsFor("xlsx", o);
async function edit(bytes, fn) { const z = await H.zipOf(bytes); await fn(z); return new Uint8Array(await z.generateAsync({ type: "uint8array" })); }
const editText = (z, p, fn) => z.file(p).async("string").then((s) => z.file(p, fn(s)));

// Same part names, every part byte-identical — except the listed ones, which must be canonically equal XML.
async function assertSameFile(actual, expected, label, canonicalOnly = []) {
  const [a, e] = [await parts(actual), await parts(expected)];
  assert.deepEqual([...a.keys()].sort(), [...e.keys()].sort(), `${label}: same parts`);
  for (const [n, b] of e) {
    if (b.equals(a.get(n))) continue;
    if (/\.(xml|rels)$/.test(n)) assert.equal(canon(a.get(n).toString()), canon(b.toString()), `${label}: ${n} restored`);
    assert.ok(canonicalOnly.includes(n), `${label}: ${n} is byte-identical`);
  }
}

for (const name of FIXTURES) {
  test(`${name}: data, formulas, tables, charts, names, validation and VBA are untouched`, async () => {
    const orig = fixture(name);
    const a = await WT.apply(orig, theme("synthwave"), {}, makeAssets, "test");
    H.save(`xlsx-${name}`, a);
    const [po, pa] = [await parts(orig), await parts(a)];
    for (const p of SHEETS) {
      const [xo, xa] = [po.get(p).toString(), pa.get(p).toString()];
      assert.deepEqual(cells(xa), cells(xo), `${p}: every cell formula, value and style index`);
      assert.equal(sheetData(xa), sheetData(xo), `${p}: sheetData byte-identical`);
      for (const b of ["conditionalFormatting", "dataValidations", "pageMargins", "drawing", "tableParts", "sheetViews"]) assert.deepEqual(block(xa, b), block(xo, b), `${p}: ${b}`);
    }
    assert.ok(cells(po.get(SHEETS[0]).toString()).some((c) => c[3] && c[3].includes("SUM(B2:C2)")), "fixture has formulas");
    assert.deepEqual(els(pa.get(SHEETS[1]).toString(), "f").map((f) => f.textContent), ["Sales!D2*2"], "cross-sheet formula");
    // parts the theme has no business in: byte-identical
    for (const p of ["xl/workbook.xml", "xl/tables/table1.xml", "xl/charts/chart1.xml", "xl/drawings/drawing1.xml", "xl/drawings/_rels/drawing1.xml.rels", "docProps/app.xml", "docProps/core.xml"]) {
      assert.ok(pa.get(p).equals(po.get(p)), `${p} byte-identical`);
    }
    assert.deepEqual(block(pa.get("xl/workbook.xml").toString(), "definedNames"), ['<definedNames xmlns="' + SNS + '"><definedName name="Totals">Sales!$D$2:$D$5</definedName></definedNames>']);
    assert.deepEqual(els(pa.get("xl/charts/chart1.xml").toString(), "f", CNS).map((f) => f.textContent), els(po.get("xl/charts/chart1.xml").toString(), "f", CNS).map((f) => f.textContent), "chart data references");
    assert.ok(!pa.has("xl/calcChain.xml") || pa.get("xl/calcChain.xml").equals(po.get("xl/calcChain.xml")));
    // content types of every original part are unchanged (macro-enabled workbook stays macro-enabled)
    const ct = (b) => new Map(els(b.get("[Content_Types].xml").toString(), "Override", "http://schemas.openxmlformats.org/package/2006/content-types").map((o) => [o.getAttribute("PartName"), o.getAttribute("ContentType")]));
    const [co, ca] = [ct(po), ct(pa)];
    for (const [k, v] of co) assert.equal(ca.get(k), v, `content type of ${k}`);
    if (name.endsWith(".xlsm")) {
      assert.ok(pa.get("xl/vbaProject.bin").equals(po.get("xl/vbaProject.bin")), "vbaProject.bin byte-identical");
      assert.equal(ca.get("/xl/workbook.xml"), "application/vnd.ms-excel.sheet.macroEnabled.main+xml");
      assert.match(pa.get("[Content_Types].xml").toString(), /Extension="bin" ContentType="application\/vnd.ms-office.vbaProject"/);
      assert.match(pa.get("xl/_rels/workbook.xml.rels").toString(), /Id="rIdVBA"[^>]*Target="vbaProject.bin"/);
    }
  });

  test(`${name}: the theme reaches fonts, built-in styles, tables, tabs and old colours`, async () => {
    const orig = fixture(name), pal = WT.palette(theme("synthwave"));
    const a = await WT.apply(orig, theme("synthwave"), {}, makeAssets, "test");
    assert.deepEqual(a.stats, { sheets: 2, pictures: 2, tabs: 2, fonts: 2, styles: 0, tables: true, direct: { fonts: 1, fills: 0, charts: 0 } });
    const st = await styles(a), fonts = list(st, "fonts");
    // cached font names follow the theme fonts; Normal is the minor (body) font
    assert.equal(kids(fonts[0], "name")[0].getAttribute("val"), pal.minor);
    assert.equal(kids(fonts[0], "scheme")[0].getAttribute("val"), "minor");
    const title = list(st, "fonts")[Number(styleXf(st, "Title").getAttribute("fontId"))];
    assert.equal(kids(title, "name")[0].getAttribute("val"), pal.major);
    assert.equal(kids(title, "color")[0].getAttribute("theme"), "3", "Title on dk2 (theme index 3)");
    // F1: hard-coded Office blue -> accent1 reference; F2: dark red is not a theme colour and stays
    const f1 = kids(fontOfCell(st, 2), "color")[0], f2 = kids(fontOfCell(st, 3), "color")[0];
    assert.equal(f1.getAttribute("theme"), "4"); assert.ok(!f1.hasAttribute("rgb"));
    assert.equal(f2.getAttribute("rgb"), "FFC00000"); assert.ok(!f2.hasAttribute("theme"));
    // new tables follow accent1; the existing table keeps its own style
    assert.equal(kids(st, "tableStyles")[0].getAttribute("defaultTableStyle"), "TableStyleMedium2");
    assert.match(await H.text(a, "xl/tables/table1.xml"), /<tableStyleInfo name="TableStyleMedium9"/);
    // tabs cycle through the accents; every worksheet carries the backdrop
    const [s1, s2] = [await H.text(a, SHEETS[0]), await H.text(a, SHEETS[1])];
    assert.equal(tabOf(s1), '<tabColor theme="4"/>'); assert.equal(tabOf(s2), '<tabColor theme="5"/>');
    assert.match(s1, /^<worksheet[^>]*><sheetPr><tabColor theme="4"\/><outlinePr /, "tabColor is sheetPr's first child");
    assert.match(s1, /<drawing [^>]*\/><picture [^>]*r:id="rIdWT1"\/><tableParts /, "picture between drawing and tableParts");
    const pr = await S.probe(a);
    assert.equal(pr.pictures, 2); assert.equal(pr.media.length, 1); assert.equal(pr.marker, "synthwave"); assert.equal(pr.markerName, "Synthwave Sunset");
    assert.equal(WT.schemeOf(pr.themes[0]).accent1, pal.accent1);
    const custom = await H.text(a, "docProps/custom.xml");
    assert.match(custom, /name="WildThemeVersion"><vt:lpwstr>4</);
    assert.match(custom, /name="WildThemeMedia"><vt:lpwstr>[0-9a-f]{8}-\d+</);
  });

  test(`${name}: remove restores theme, styles and sheets exactly, even after re-theming`, async () => {
    const orig = fixture(name);
    const a = await WT.apply(orig, theme("synthwave"), {}, makeAssets, "test");
    const b = await WT.apply(a, theme("winter"), { tabs: true }, makeAssets, "test");
    const c = await WT.apply(b, theme("aurora"), { backdrop: false, direct: false }, makeAssets, "test");
    for (const [label, file] of [["after one theme", a], ["after re-themes", c]]) {
      const r = await WT.remove(file);
      assert.equal((await WT.inspect(r)).status, "none");
      // the core re-serialises the two parts its own custom XML part hangs from; the stamper's parts come back byte for byte
      await assertSameFile(r, orig, label, ["[Content_Types].xml", "xl/_rels/workbook.xml.rels"]);
    }
  });
}

test("xlsx: each setting switches its part on and off, and unstamp undoes exactly what was done", async () => {
  const orig = fixture("book.xlsx"), t = theme("galaxy"), pal = WT.palette(t);
  const assets = { backdrop: { bytes: fixture("backdrop.jpg"), ext: "jpeg" }, panel: null };
  const runs = { all: {}, backdrop: { backdrop: false }, styles: { styles: false }, tables: { tables: false }, direct: { direct: false }, tabs: { tabs: false },
    none: { backdrop: false, styles: false, tables: false, direct: false, tabs: false } };
  for (const [label, o] of Object.entries(runs)) {
    const s = opts(o), out = await S.stamp(orig, t, s.backdrop ? assets : { backdrop: null }, s);
    assert.deepEqual(out.failures, [], label);
    assert.doesNotThrow(() => JSON.parse(JSON.stringify(out.backup)), "backup is JSON");
    const st = await styles(out), s1 = await H.text(out, SHEETS[0]), names = await H.names(out);
    const font0 = kids(list(st, "fonts")[0], "name")[0].getAttribute("val");
    assert.equal(font0, s.styles ? pal.minor : "Calibri", `${label}: Normal font`);
    assert.equal(kids(st, "tableStyles")[0].getAttribute("defaultTableStyle"), s.tables ? "TableStyleMedium2" : "TableStyleMedium9", `${label}: default table style`);
    assert.equal(attr(kids(fontOfCell(st, 2), "color")[0], "theme"), s.direct ? "4" : null, `${label}: old blue`);
    assert.equal(tabOf(s1), s.tabs ? '<tabColor theme="4"/>' : null, `${label}: tab colour`);
    assert.equal(/<picture /.test(s1), s.backdrop, `${label}: picture`);
    assert.equal(names.some((n) => /wildtheme-backdrop/.test(n)), s.backdrop, `${label}: media`);
    assert.equal(/Extension="jpeg"/.test(await H.text(out, "[Content_Types].xml")), s.backdrop, `${label}: jpeg content type`);
    assert.equal((await S.probe(out)).pictures, s.backdrop ? 2 : 0);
    assert.equal(WT.schemeOf(await H.text(out, "xl/theme/theme1.xml")).accent1, pal.accent1, `${label}: theme always written`);
    await assertSameFile(await S.unstamp(out, JSON.parse(JSON.stringify(out.backup))), orig, `${label} unstamped`);
  }
  // requested parts that can't be done are reported, never skipped silently
  const missing = await S.stamp(orig, t, { backdrop: null }, opts({}));
  assert.deepEqual(missing.failures, ["No scene picture was supplied for the sheets."]);
  const gif = await S.stamp(orig, t, { backdrop: { bytes: new Uint8Array([0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 1, 0, 1, 0]), ext: "gif" } }, opts({}));
  assert.deepEqual(gif.failures, ["Excel sheet backgrounds must be PNG or JPEG."]);
  await assert.rejects(WT.apply(orig, t, {}, async () => ({ backdrop: null, panel: null }), "test"), /partly applied: No scene picture/);
});

test("xlsx: deterministic and idempotent", async () => {
  const orig = fixture("macro.xlsm"), t = theme("synthwave"), s = opts({});
  const a1 = await S.stamp(orig, t, await makeAssets(t, "xlsx", null, s), s);
  const a2 = await S.stamp(orig, t, await makeAssets(t, "xlsx", null, s), s);
  assert.ok(Buffer.from(a1).equals(Buffer.from(a2)), "same input, same bytes");
  assert.deepEqual(a1.backup, a2.backup);
  const back = await S.unstamp(a1, a1.backup);
  const a3 = await S.stamp(back, t, await makeAssets(t, "xlsx", null, s), s);
  assert.ok(Buffer.from(a3).equals(Buffer.from(a1)), "stamp(unstamp(stamp(x))) == stamp(x)");
  // unstamping a file that never had a theme changes nothing (with or without a backup)
  assert.ok(Buffer.from(await S.unstamp(orig)).equals(Buffer.from(orig)));
  await assertSameFile(await S.unstamp(orig, a1.backup), orig, "unthemed + backup", ["xl/styles.xml"]);
});

test("xlsx: old chart colours follow the new accents and come back", async () => {
  const solid = (v) => `<a:solidFill xmlns:a="${ANS}"><a:srgbClr val="${v}"><a:alpha val="80000"/></a:srgbClr></a:solidFill>`;
  const orig = await edit(fixture("book.xlsx"), (z) => editText(z, "xl/charts/chart1.xml", (x) => {
    let i = 0;
    return x.replace(/<spPr>/g, () => `<spPr>${solid(["4F81BD", "C0504D"][i++])}`).replace("</plotArea>", `<spPr>${solid("C00000")}</spPr></plotArea>`);
  }));
  const a = await WT.apply(orig, theme("synthwave"), {}, makeAssets, "test");
  assert.equal(a.stats.direct.charts, 2);
  const cx = await H.text(a, "xl/charts/chart1.xml");
  const vals = els(cx, "schemeClr", ANS).map((e) => e.getAttribute("val"));
  assert.deepEqual(vals, ["accent1", "accent2"], "series fills on theme accents");
  assert.deepEqual(els(cx, "srgbClr", ANS).map((e) => e.getAttribute("val")), ["C00000"], "non-theme colour stays");
  assert.equal(els(cx, "alpha", ANS).length, 3, "colour modifiers kept");
  assert.deepEqual(els(cx, "f", CNS).map((f) => f.textContent), els(await H.text(orig, "xl/charts/chart1.xml"), "f", CNS).map((f) => f.textContent));
  await assertSameFile(await WT.remove(a), orig, "chart", ["[Content_Types].xml", "xl/_rels/workbook.xml.rels", "xl/charts/chart1.xml"]);
});

test("xlsx: a workbook without a theme gets one, and loses it again", async () => {
  const orig = await edit(fixture("book.xlsx"), async (z) => {
    z.remove("xl/theme/theme1.xml");
    await editText(z, "xl/_rels/workbook.xml.rels", (x) => x.replace(/<Relationship [^>]*relationships\/theme"[^>]*\/>/, ""));
    await editText(z, "[Content_Types].xml", (x) => x.replace(/<Override PartName="\/xl\/theme\/theme1.xml"[^>]*\/>/, ""));
  });
  assert.ok(Buffer.from(await S.unstamp(orig)).equals(Buffer.from(orig)), "unstamp of a theme-less file changes nothing");
  assert.equal((await S.probe(orig)).themes.length, 0);
  const a = await WT.apply(orig, theme("winter"), {}, makeAssets, "test");
  assert.equal((await WT.inspect(a)).status, "intact");
  assert.equal(WT.schemeOf(await H.text(a, "xl/theme/theme1.xml")).accent1, WT.palette(theme("winter")).accent1);
  assert.match(await H.text(a, "xl/_rels/workbook.xml.rels"), /Target="theme\/theme1.xml"/);
  assert.match(await H.text(a, "[Content_Types].xml"), /PartName="\/xl\/theme\/theme1.xml"/);
  await assertSameFile(await WT.remove(a), orig, "theme-less", ["[Content_Types].xml", "xl/_rels/workbook.xml.rels"]);
});

test("xlsx: a workbook without styles.xml gets Excel's default one, and loses it again", async () => {
  const orig = await edit(fixture("book.xlsx"), async (z) => {
    z.remove("xl/styles.xml");
    await editText(z, "xl/_rels/workbook.xml.rels", (x) => x.replace(/<Relationship [^>]*relationships\/styles"[^>]*\/>/, ""));
    await editText(z, "[Content_Types].xml", (x) => x.replace(/<Override PartName="\/xl\/styles.xml"[^>]*\/>/, ""));
    await editText(z, SHEETS[0], (x) => x.replace(/ s="\d+"/g, ""));
  });
  const a = await WT.apply(orig, theme("aurora"), {}, makeAssets, "test");
  const st = await styles(a);
  assert.equal(kids(list(st, "fonts")[0], "name")[0].getAttribute("val"), WT.palette(theme("aurora")).minor);
  assert.equal(kids(st, "tableStyles")[0].getAttribute("defaultTableStyle"), "TableStyleMedium2");
  await assertSameFile(await WT.remove(a), orig, "styles-less", ["[Content_Types].xml", "xl/_rels/workbook.xml.rels"]);
});

test("xlsx: the user's own sheet background and tab colour come back on remove", async () => {
  const orig = await edit(fixture("book.xlsx"), async (z) => {
    z.file("xl/media/user-bg.png", fixture("panel.png"));
    z.file("xl/worksheets/_rels/sheet2.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdU1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/user-bg.png"/></Relationships>`);
    await editText(z, SHEETS[1], (x) => x.replace("<sheetPr>", '<sheetPr codeName="Notes"><tabColor rgb="FF00B050"/>').replace("</worksheet>", '<picture xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdU1"/></worksheet>'));
    await editText(z, SHEETS[0], (x) => x.replace(/<sheetPr>[\s\S]*?<\/sheetPr>/, '<sheetPr filterMode="0"/>'));
    await editText(z, "[Content_Types].xml", (x) => x.replace("<Override", '<Default Extension="png" ContentType="image/png"/><Override'));
  });
  const a = await WT.apply(orig, theme("synthwave"), {}, makeAssets, "test");
  const [s1, s2] = [await H.text(a, SHEETS[0]), await H.text(a, SHEETS[1])];
  assert.match(s1, /<sheetPr filterMode="0"><tabColor theme="4"\/><\/sheetPr>/);
  assert.match(s2, /<sheetPr codeName="Notes"><tabColor theme="5"\/>/);
  assert.equal((s2.match(/<picture /g) || []).length, 1, "one background per sheet");
  assert.match(s2, /<picture [^>]*r:id="rIdWT1"\/>/);
  assert.ok((await H.names(a)).includes("xl/media/user-bg.png"), "the user's picture stays in the file for later");
  await assertSameFile(await WT.remove(a), orig, "user background", ["[Content_Types].xml", "xl/_rels/workbook.xml.rels"]);
});

test("xlsx: built-in styles move onto theme colours without recolouring Normal or other cells", async () => {
  const orig = await edit(fixture("book.xlsx"), (z) => editText(z, "xl/styles.xml", (x) => x
    .replace('<color theme="3"/><sz val="18"/>', '<color rgb="FF1F497D"/><sz val="18"/>') // Title: hard-coded colour, own font
    .replace('<borders count="1">', '<borders count="2">').replace("</borders>", '<border><left/><right/><top style="thin"><color rgb="FF4F81BD"/></top><bottom style="double"><color rgb="FF4F81BD"/></bottom><diagonal/></border></borders>')
    .replace('<cellStyleXfs count="2">', '<cellStyleXfs count="4">').replace("</cellStyleXfs>", '<xf numFmtId="0" fontId="2" fillId="0" borderId="0"/><xf numFmtId="0" fontId="3" fillId="0" borderId="1"/></cellStyleXfs>')
    .replace('<cellXfs count="4">', '<cellXfs count="6">').replace("</cellXfs>", '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="2"/><xf numFmtId="0" fontId="3" fillId="0" borderId="1" xfId="3"/></cellXfs>')
    .replace('<cellStyles count="2">', '<cellStyles count="4">').replace("</cellStyles>", '<cellStyle name="Heading 1" xfId="2" builtinId="16"/><cellStyle name="Total" xfId="3" builtinId="25"/></cellStyles>')));
  const a = await WT.apply(orig, theme("synthwave"), {}, makeAssets, "test");
  assert.equal(a.stats.styles, 3);
  const st = await styles(a), fonts = list(st, "fonts");
  // Title owns its font: recoloured in place
  assert.equal(kids(fonts[Number(styleXf(st, "Title").getAttribute("fontId"))], "color")[0].getAttribute("theme"), "3");
  // Heading 1 shared the blue font 2 with cell F1: it gets a themed copy (dk2); F1 keeps font 2 (moved to accent1 as an old colour)
  assert.equal(fonts.length, 6); assert.equal(kids(st, "fonts")[0].getAttribute("count"), "6");
  assert.equal(styleXf(st, "Heading 1").getAttribute("fontId"), "4");
  assert.equal(kids(fonts[4], "color")[0].getAttribute("theme"), "3");
  assert.equal(list(st, "cellXfs")[4].getAttribute("fontId"), "4", "cells in Heading 1 follow it");
  assert.equal(list(st, "cellXfs")[2].getAttribute("fontId"), "2", "F1 keeps its font");
  assert.equal(kids(fonts[2], "color")[0].getAttribute("theme"), "4");
  // Normal is never recoloured
  assert.equal(kids(fonts[0], "color")[0].getAttribute("theme"), "1", "Normal still dk1");
  assert.equal(list(st, "cellXfs")[0].getAttribute("fontId"), "0", "Normal cells untouched");
  // Total shared the red font 3 with cell F2: themed copy (dk1); F2 stays red; its border sides go on accent1
  assert.equal(styleXf(st, "Total").getAttribute("fontId"), "5");
  assert.equal(kids(fonts[5], "color")[0].getAttribute("theme"), "1");
  assert.equal(list(st, "cellXfs")[5].getAttribute("fontId"), "5");
  assert.equal(kids(fontOfCell(st, 3), "color")[0].getAttribute("rgb"), "FFC00000");
  const border = list(st, "borders")[1];
  for (const side of ["top", "bottom"]) assert.equal(kids(kids(border, side)[0], "color")[0].getAttribute("theme"), "4");
  await assertSameFile(await WT.remove(a), orig, "built-in styles", ["[Content_Types].xml", "xl/_rels/workbook.xml.rels"]);
});

// A real third-party reader: openpyxl must open the themed workbook and see the same data.
const PY = `
import sys, json, warnings
warnings.simplefilter("ignore")
from openpyxl import load_workbook
p = sys.argv[1]
wb = load_workbook(p, keep_vba=p.endswith(".xlsm"))
ws = wb["Sales"]
col = lambda c: [c.type, c.theme if c.type == "theme" else c.value]
print(json.dumps({
  "sheets": wb.sheetnames, "formulas": [ws[f"D{i}"].value for i in range(2, 6)], "notes": wb["Notes"]["A2"].value,
  "tables": {k: ws.tables[k].tableStyleInfo.name for k in ws.tables.keys()}, "names": wb.defined_names["Totals"].attr_text,
  "dv": [str(d.sqref) for d in ws.data_validations.dataValidation], "cf": [str(r.sqref) for r in ws.conditional_formatting],
  "freeze": ws.freeze_panes, "charts": len(ws._charts),
  "tabs": [s.sheet_properties.tabColor.theme if s.sheet_properties.tabColor is not None else None for s in wb.worksheets],
  "f1": col(ws["F1"].font.color), "f2": col(ws["F2"].font.color), "body": ws["A2"].font.name, "title": ws["A1"].font.name,
  "theme": b"Wild Theme" in (wb.loaded_theme or b""), "vba": wb.vba_archive is not None,
}))
`;
const hasOpenpyxl = () => spawnSync("python3", ["-I", "-c", "import openpyxl"]).status === 0;
for (const name of FIXTURES) {
  test(`${name}: openpyxl opens the themed workbook with its data intact`, { skip: !hasOpenpyxl() && "openpyxl not installed" }, async () => {
    const t = theme("synthwave"), pal = WT.palette(t);
    const p = H.save(`xlsx-openpyxl-${name}`, await WT.apply(fixture(name), t, {}, makeAssets, "test"));
    const r = spawnSync("python3", ["-I", "-c", PY, p], { encoding: "utf8" });
    assert.equal(r.status, 0, r.stderr);
    assert.deepEqual(JSON.parse(r.stdout), {
      sheets: ["Sales", "Notes"], formulas: ["=SUM(B2:C2)", "=SUM(B3:C3)", "=SUM(B4:C4)", "=SUM(B5:C5)"], notes: "=Sales!D2*2",
      tables: { SalesTable: "TableStyleMedium9" }, names: "Sales!$D$2:$D$5", dv: ["E2:E5"], cf: ["B2:C5"], freeze: "A2", charts: 1,
      tabs: [4, 5], f1: ["theme", 4], f2: ["rgb", "FFC00000"], body: pal.minor, title: pal.major, theme: true, vba: name.endsWith(".xlsm"),
    });
  });
}

// Another editor re-saves the file: what of the theme really survives LibreOffice, and does remove still work?
for (const name of FIXTURES) {
  test(`${name}: LibreOffice round trip keeps the theme recognisable and removable`, { skip: !H.hasSoffice() && "LibreOffice not installed" }, async () => {
    const t = theme("synthwave"), pal = WT.palette(t);
    const a = await WT.apply(fixture(name), t, {}, makeAssets, "test");
    const lo = H.libreRoundTrip(`xlsx-lo-${name}`, a);
    const names = await H.names(lo), i = await WT.inspect(lo);
    // LibreOffice keeps: the theme part (re-written, same scheme), our custom properties, the custom XML part,
    // the sheet backgrounds (media renamed to image1.jpeg — recognised by fingerprint) and the tab colours (as RGB).
    const keptMeta = names.some((n) => /^customXml\/item\d+\.xml$/.test(n));
    assert.equal(i.status, keptMeta ? "intact" : "legacy", JSON.stringify(i.problems));
    assert.equal(i.theme.id, "synthwave");
    assert.equal(i.probe.marker, "synthwave", "custom-property marker survives");
    assert.equal(i.probe.pictures, 2, "backgrounds survive and are still recognised");
    const sc = WT.schemeOf(i.probe.themes[0]);
    for (const k of ["dk2", "accent1", "accent6", "major", "minor"]) assert.equal(sc[k], pal[k], `theme ${k} survives`);
    const s1 = await H.text(lo, SHEETS[0]);
    assert.ok([`<tabColor theme="4"/>`, `<tabColor rgb="FF${pal.accent1}"/>`].includes(tabOf(s1)), tabOf(s1));
    assert.deepEqual(els(s1, "f").map((f) => f.textContent), ["SUM(B2:C2)", "SUM(B3:C3)", "SUM(B4:C4)", "SUM(B5:C5)"]);
    // what doesn't: LibreOffice drops the (dummy) VBA project and the default table style
    if (name.endsWith(".xlsm")) assert.ok(!names.includes("xl/vbaProject.bin"), "LibreOffice drops the dummy VBA project");
    // remove still takes everything out
    const r = await WT.remove(lo);
    assert.equal((await WT.inspect(r)).status, "none");
    assert.deepEqual(WT.schemeOf(await H.text(r, "xl/theme/theme1.xml")), WT.schemeOf(await H.text(fixture(name), "xl/theme/theme1.xml")), "original theme back");
    const r1 = await H.text(r, SHEETS[0]);
    assert.equal(tabOf(r1), null, "tab colour gone"); assert.ok(!/<picture /.test(r1), "background gone");
    assert.ok(!(await H.names(r)).some((n) => /\/media\//.test(n)), "no picture left behind");
    assert.equal(await H.text(r, "docProps/custom.xml"), null);
  });
}

test("xlsx: LibreOffice renders the themed sheet readably", { skip: !H.hasSoffice() && "LibreOffice not installed" }, async () => {
  // the real tool gives Excel a scene washed 80% towards white (assets.js); do the same with the test scene
  const washed = path.join(H.OUT, "xlsx-washed.jpg");
  const py = "import sys\nfrom PIL import Image\nim=Image.open(sys.argv[1]).convert('RGB').resize((1280,720))\nImage.blend(im,Image.new('RGB',im.size,(255,255,255)),0.8).save(sys.argv[2],quality=85)";
  assert.equal(spawnSync("python3", ["-I", "-c", py, path.join(H.FIX, "backdrop.jpg"), washed]).status, 0);
  const assets = async (t, k, s, o) => ({ backdrop: o.backdrop ? { bytes: new Uint8Array(fs.readFileSync(washed)), ext: "jpeg" } : null, panel: null });
  const t = theme("synthwave"), pal = WT.palette(t);
  const png = H.render("xlsx-render.xlsx", await WT.apply(fixture("book.xlsx"), t, {}, assets, "test"));
  const bg = H.sample(png, [0.15, 0.27, 0.45, 0.33]);
  assert.ok(H.lum(bg) > 180 && H.lum(bg) < 250, `pale backdrop behind the cells: ${bg}`);
  // the chart's first bar is drawn in the theme's accent1 (old accent1 -> new accent1 through the theme)
  const bar = H.sample(png, [0.875, 0.27, 0.89, 0.32]), want = [0, 2, 4].map((i) => parseInt(pal.accent1.slice(i, i + 2), 16));
  assert.ok(bar.every((v, i) => Math.abs(v - want[i]) < 40), `chart bar ${bar} ~ accent1 ${want}`);
});
