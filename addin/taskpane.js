// Wild Themes add-in for Word, PowerPoint and Excel. Reads the open file, writes the theme
// into it, and puts it back. Nothing runs until you click a theme.
const T = self.WW_THEMES, $ = (s) => document.querySelector(s);
let current = null, filter = "all", busy = false, host = "word";
const msg = (t) => { $("#msg").textContent = t; };
const kind = () => ({ word: "docx", ppt: "pptx", excel: "xlsx" })[host];

function render() {
  for (const b of document.querySelectorAll(".seg button")) b.setAttribute("aria-checked", b.dataset.f === filter);
  const grid = $("#grid"); grid.textContent = "";
  for (const [id, t] of Object.entries(T)) {
    if (filter !== "all" && t.mode !== filter) continue;
    const b = document.createElement("button"), img = new Image(), s = document.createElement("span");
    b.className = "card"; b.setAttribute("role", "radio"); b.setAttribute("aria-checked", id === current); b.title = t.blurb;
    img.src = `../extension/previews/${id}.webp`; img.alt = ""; s.textContent = t.name;
    b.append(img, s); b.onclick = () => apply(id, b);
    grid.append(b);
  }
  $("#now").textContent = current ? `This file: ${T[current] ? T[current].name : "a custom theme"}` : "This file has no theme yet.";
  $("#remove").hidden = !current || host === "excel";
}
for (const b of document.querySelectorAll(".seg button")) b.onclick = () => { filter = b.dataset.f; render(); };

// The whole file, in slices (Office hands files over in chunks).
function getFile() {
  return new Promise((ok, fail) => Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 4194304 }, (r) => {
    if (r.status !== Office.AsyncResultStatus.Succeeded) return fail(r.error);
    const file = r.value, parts = []; let i = 0;
    const next = () => file.getSliceAsync(i, (s) => {
      if (s.status !== Office.AsyncResultStatus.Succeeded) { file.closeAsync(); return fail(s.error); }
      parts.push(new Uint8Array(s.value.data));
      if (++i < file.sliceCount) next();
      else { file.closeAsync(); const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0)); let o = 0; for (const p of parts) { out.set(p, o); o += p.length; } ok(out); }
    });
    next();
  }));
}
function toBase64(bytes) {
  let s = ""; for (let i = 0; i < bytes.length; i += 32768) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 32768));
  return btoa(s);
}
const tooBig = (b64) => b64.length > 4 * 1024 * 1024 * 1.37;

// Word: theme pictures sit in locked content controls; only this add-in unlocks them.
async function unlockWord() {
  try {
    await Word.run(async (ctx) => {
      const lists = [ctx.document.body.contentControls.getByTag("WildThemeMarker")];
      const secs = ctx.document.sections; secs.load("items"); await ctx.sync();
      for (const sec of secs.items) for (const k of ["Primary", "FirstPage", "EvenPages"]) lists.push(sec.getHeader(k).contentControls.getByTag("WildTheme"));
      for (const l of lists) l.load("items"); await ctx.sync();
      for (const l of lists) for (const cc of l.items) { cc.cannotDelete = false; cc.cannotEdit = false; }
      await ctx.sync();
    });
  } catch (_) { /* nothing to unlock */ }
}

// Put a rebuilt file back in place of the open one.
async function swapIn(bytes) {
  const b64 = toBase64(bytes);
  if (tooBig(b64)) throw new Error("This file is too big for Office to swap in place (over about 4 MB). Use the browser extension's “Save theme into a file” instead.");
  if (host === "word") {
    const opts = { importTheme: true, importStyles: true, importPageColor: true, importParagraphSpacing: true };
    if (Office.context.requirements.isSetSupported("WordApi", "1.6")) opts.importCustomProperties = true;
    if (Office.context.requirements.isSetSupported("WordApi", "1.7")) opts.importDifferentOddEvenPages = true;
    await unlockWord();
    await Word.run(async (ctx) => { ctx.document.insertFileFromBase64(b64, "Replace", opts); await ctx.sync(); });
  } else if (host === "ppt") {
    // PowerPoint: add the themed slides (with their own master), then remove the old ones.
    await PowerPoint.run(async (ctx) => {
      const slides = ctx.presentation.slides; slides.load("items/id"); await ctx.sync();
      const old = slides.items.map((s) => s.id);
      ctx.presentation.insertSlidesFromBase64(b64, { formatting: "KeepSourceFormatting", targetSlideId: old[old.length - 1] });
      await ctx.sync();
      for (const id of old) slides.getItem(id).delete();
      await ctx.sync();
    });
  }
}

async function apply(id, card) {
  if (busy) return;
  if (host === "excel") return msg("Excel can't swap a whole workbook from an add-in. Use the browser extension's “Save theme into a file” page for .xlsx files, then upload the themed copy.");
  busy = true; card.setAttribute("aria-busy", "true");
  const t = { ...T[id], id };
  try {
    msg("Reading the file…");
    const bytes = await getFile(), S = WWAssets.stamperFor(kind());
    const info = await S.analyze(bytes);
    msg("Adding the theme…");
    const out = await S.stamp(bytes, t, await WWAssets.assetsFor(t, kind(), info.sections[0], new URL("../extension/", location.href).href));
    msg("Saving it into the file…");
    await swapIn(out);
    current = id; msg(`Done. ${t.name} is saved in this file.`); render();
  } catch (e) {
    msg((e && e.message) || "Couldn't apply the theme.");
  } finally { busy = false; card.removeAttribute("aria-busy"); }
}

async function removeTheme() {
  if (busy) return;
  busy = true;
  try {
    msg("Removing the theme…");
    await swapIn(await WWAssets.stamperFor(kind()).unstamp(await getFile()));
    current = null; msg("The theme has been removed from this file."); render();
  } catch (e) { msg((e && e.message) || "Couldn't remove the theme."); }
  finally { busy = false; }
}
$("#remove").onclick = removeTheme;

Office.onReady(async (info) => {
  host = info.host === Office.HostType.PowerPoint ? "ppt" : info.host === Office.HostType.Excel ? "excel" : "word";
  const need = { word: ["WordApi", "1.5"], ppt: ["PowerPointApi", "1.2"] }[host];
  render();
  if (host === "excel") {
    $("#lede").textContent = "In Excel, themes are saved into workbooks with the Wild Themes browser extension (“Save theme into a file”). Office doesn't let add-ins swap a whole workbook.";
    return;
  }
  if (!Office.context.requirements.isSetSupported(need[0], need[1])) {
    msg("This version of Office can't save themes in place. Update it, or use the browser extension's “Save theme into a file”.");
    for (const c of document.querySelectorAll(".card")) c.disabled = true;
    return;
  }
  try {   // which theme is in this file? (the file reports it in its own properties)
    current = (await WWAssets.stamperFor(kind()).analyze(await getFile())).existingTheme;
  } catch (_) { /* unknown: leave blank */ }
  render();
});
