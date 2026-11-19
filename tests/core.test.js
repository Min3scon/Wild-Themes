// Contract every format must meet, through the one shared pipeline (wtcore.js) the site, extension and add-in use.
const test = require("node:test"), assert = require("node:assert/strict");
const H = require("./helpers.js");
const { WT, theme, fixture, makeAssets } = H;

const FILES = ["report.docx", "plain.docx", "deck.pptx", "book.xlsx", "macro.xlsm"];
const themePartOf = { docx: "word/theme/theme1.xml", pptx: "ppt/theme/theme1.xml", xlsx: "xl/theme/theme1.xml" };
const scheme = async (bytes, kind) => { const x = await H.text(bytes, themePartOf[kind]); return x ? WT.schemeOf(x) : null; };
const ours = (ns) => ns.filter((n) => /customXml\/item\d+\.xml$/.test(n));

for (const name of FILES) {
  const kind = H.kindOfName(name);
  test(`${name}: apply, inspect, re-apply, change, remove, repair`, async () => {
    const orig = fixture(name);
    const before = await WT.inspect(orig);
    assert.equal(before.status, "none");
    const origScheme = await scheme(orig, kind);

    // apply
    const a = await WT.apply(orig, theme("synthwave"), {}, makeAssets, "test");
    H.save(`core-${name}`, a);
    const ia = await WT.inspect(a);
    assert.equal(ia.status, "intact", JSON.stringify(ia.problems));
    assert.equal(ia.theme.id, "synthwave");
    assert.equal(ia.kind, kind);
    assert.ok(ia.hasBackup, "backup of the original is embedded");
    const pal = WT.palette(theme("synthwave"));
    const sa = await scheme(a, kind);
    for (const k of ["dk1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"]) assert.equal(sa[k], pal[k], `theme slot ${k}`);
    assert.equal(sa.major, pal.major); assert.equal(sa.minor, pal.minor);

    // re-apply the same theme: identical structure, nothing stacked
    const b = await WT.apply(a, theme("synthwave"), {}, makeAssets, "test");
    assert.deepEqual(await H.names(b), await H.names(a), "same parts after re-applying");
    assert.equal(ours(await H.names(b)).length, ours(await H.names(a)).length);
    assert.equal((await WT.inspect(b)).status, "intact");

    // change to another theme: backup still describes the ORIGINAL file
    const c = await WT.apply(b, theme("winter"), {}, makeAssets, "test");
    const ic = await WT.inspect(c);
    assert.equal(ic.theme.id, "winter"); assert.equal(ic.status, "intact");
    assert.deepEqual(ic.meta.backup, ia.meta.backup, "backup unchanged across re-themes");

    // remove: native look back, no Wild Theme parts left
    const r = await WT.remove(c);
    H.save(`core-removed-${name}`, r);
    assert.equal((await WT.inspect(r)).status, "none");
    assert.deepEqual(await scheme(r, kind), origScheme, "original theme colours and fonts restored");
    assert.ok(!(await H.names(r)).some((n) => /wildtheme/i.test(n)), "no Wild Theme media left");

    // tamper with the theme colours -> modified; repair from the file's own definition -> intact
    const z = await H.zipOf(a);
    const tx = await z.file(themePartOf[kind]).async("string");
    z.file(themePartOf[kind], tx.replace(/(<a:accent1>\s*<a:srgbClr val=")[0-9A-Fa-f]{6}/, "$1123456"));
    const tampered = await z.generateAsync({ type: "uint8array" });
    const it = await WT.inspect(tampered);
    assert.equal(it.status, "modified", JSON.stringify(it.problems));
    const fixed = await WT.repair(tampered, makeAssets, "test");
    assert.equal((await WT.inspect(fixed)).status, "intact");
  });

  test(`${name}: settings are stored, read back and honoured`, async () => {
    const s = { backdrop: false, styles: true, tables: false, transition: "push", panel: 70, tabs: false, glow: true };
    const out = await WT.apply(fixture(name), theme("aurora"), s, makeAssets, "test");
    const i = await WT.inspect(out);
    assert.equal(i.status, "intact", JSON.stringify(i.problems));
    for (const k of WT.FEATURES[kind]) assert.ok(k in i.settings, `setting ${k} stored`);
    assert.equal(i.settings.backdrop, false);
    assert.equal(i.probe.pictures, 0, "no scene pictures when the backdrop is off");
  });
}

test("legacy (v3) documents are recognised and upgrade cleanly", async () => {
  const fs = require("fs"), path = require("path");
  const dir = path.join(__dirname, "..", "documents");
  const file = fs.readdirSync(dir).find((f) => f.endsWith(".docx"));
  const bytes = new Uint8Array(fs.readFileSync(path.join(dir, file)));
  const i = await WT.inspect(bytes);
  assert.ok(["legacy", "intact"].includes(i.status), i.status);
  assert.ok(i.theme && i.theme.id);
  const up = await WT.apply(bytes, theme(i.theme.id), {}, makeAssets, "test");
  assert.equal((await WT.inspect(up)).status, "intact");
});

test("non-Office input is refused with a clear message", async () => {
  await assert.rejects(WT.inspect(new Uint8Array([1, 2, 3])));
});
