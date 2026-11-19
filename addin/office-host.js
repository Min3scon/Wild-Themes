// Wild Themes inside Word, PowerPoint and Excel. The add-in's task pane is the website itself
// (the gallery and the File studio, same HTML, CSS and scripts); this file only swaps "open a file
// / download it" for "read the open document / write it back", through the same core as everywhere
// else (extension/lib/wtcore.js). Loaded by the pages' <head> only when they run as the add-in.
(() => {
  const W = { active: false, host: null, kind: null, name: "" };
  const supports = (set, v) => Office.context.requirements.isSetSupported(set, v);
  const WT = () => self.WTCore;
  const run = () => W.host === "word" ? Word.run : W.host === "ppt" ? PowerPoint.run : Excel.run;
  const docOf = (ctx) => W.host === "word" ? ctx.document : W.host === "ppt" ? ctx.presentation : ctx.workbook;
  const forget = () => { try { sessionStorage.removeItem("wt-host"); } catch (_) { /* storage blocked */ } };

  // ---------------------------------------------------------------- the open file
  W.read = () => new Promise((ok, fail) => Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 4194304 }, (r) => {
    if (r.status !== Office.AsyncResultStatus.Succeeded) return fail(r.error);
    const file = r.value, parts = []; let i = 0;
    const next = () => file.getSliceAsync(i, (s) => {
      if (s.status !== Office.AsyncResultStatus.Succeeded) { file.closeAsync(); return fail(s.error); }
      parts.push(new Uint8Array(s.value.data));
      if (++i < file.sliceCount) return next();
      file.closeAsync();
      const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0)); let o = 0;
      for (const p of parts) { out.set(p, o); o += p.length; }
      ok(out);
    });
    next();
  }));
  function toBase64(bytes) {
    let s = ""; for (let i = 0; i < bytes.length; i += 32768) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 32768));
    return btoa(s);
  }
  W.canSwap = () => W.host === "word" ? supports("WordApi", "1.5") : W.host === "ppt" ? supports("PowerPointApi", "1.2") : false;
  W.canCopy = () => W.host === "word" ? supports("WordApi", "1.3") : W.host === "ppt" ? supports("PowerPointApi", "1.1") : supports("ExcelApi", "1.8");

  // Word: the theme's pictures sit in locked content controls; only Wild Themes unlocks them.
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
    } catch (_) { /* nothing locked */ }
  }
  // Keep the live document's custom XML part + properties in step with the rebuilt file
  // (Word's insert and PowerPoint's slide insert don't carry package-level parts across).
  async function syncMeta(bytes) {
    if (W.host === "excel" || !(W.host === "word" ? supports("WordApi", "1.4") : supports("PowerPointApi", "1.7"))) return;
    const zip = await JSZip.loadAsync(bytes), found = await WT().findMetaPart(zip);
    const xml = found ? (await zip.file(found.path).async("string")).replace(/^<\?xml[^>]*\?>\s*/, "") : null;
    await run()(async (ctx) => {
      const doc = docOf(ctx), parts = doc.customXmlParts.getByNamespace(WT().NS); parts.load("items"); await ctx.sync();
      for (const p of parts.items) p.delete();
      if (xml) doc.customXmlParts.add(xml);
      await ctx.sync();
      try {   // custom properties: second marker that survives editors which drop custom XML
        const cp = doc.properties.customProperties; cp.load("items/key"); await ctx.sync();
        for (const p of cp.items) if (/^WildTheme/.test(p.key)) p.delete();
        if (xml) { const m = WT().parseMeta(xml); cp.add("WildTheme", m.id); cp.add("WildThemeName", m.name); cp.add("WildThemeVersion", String(WT().VERSION)); }
        await ctx.sync();
      } catch (_) { /* older Office: the custom XML part is enough */ }
    });
  }
  // Put rebuilt bytes in place of the open document (Word, PowerPoint).
  async function swapIn(bytes, removing) {
    const b64 = toBase64(bytes);
    if (b64.length > 4 * 1024 * 1024 * 1.37) throw new Error("This file is over about 4 MB, too big for Office to swap in place. Use “Open a themed copy” instead.");
    if (W.host === "word") {
      const opts = { importTheme: true, importStyles: true, importPageColor: true, importParagraphSpacing: true };
      if (supports("WordApi", "1.6")) opts.importCustomProperties = true;
      if (supports("WordApi", "1.7")) opts.importDifferentOddEvenPages = true;
      await unlockWord();
      await Word.run(async (ctx) => { ctx.document.insertFileFromBase64(b64, "Replace", opts); await ctx.sync(); });
      if (removing && supports("WordApiDesktop", "1.4")) {   // an import without a page background doesn't clear the old one
        try { await Word.run(async (ctx) => { ctx.document.background.fill.clear(); await ctx.sync(); }); } catch (_) { /* checked by the caller */ }
      }
    } else {
      await PowerPoint.run(async (ctx) => {
        const slides = ctx.presentation.slides; slides.load("items/id"); await ctx.sync();
        const old = slides.items.map((s) => s.id);
        ctx.presentation.insertSlidesFromBase64(b64, { formatting: "KeepSourceFormatting", targetSlideId: old[old.length - 1] });
        await ctx.sync();
        for (const id of old) slides.getItem(id).delete();
        await ctx.sync();
      });
    }
    await syncMeta(bytes);
  }
  // Open bytes as a new document: exactly the file Wild Themes wrote, nothing re-imported.
  W.openCopy = async (bytes, kind = W.kind) => {
    const b64 = toBase64(bytes);
    if (kind === "docx" && W.host === "word") await Word.run(async (ctx) => { ctx.application.createDocument(b64).open(); await ctx.sync(); });
    else if (kind === "pptx" && W.host === "ppt") await PowerPoint.createPresentation(b64);
    else if (kind === "xlsx" && W.host === "excel") await Excel.createWorkbook(b64);
    else return false;
    return true;
  };
  // Write a rebuilt file into the open document (or, where Office can't, open it as a copy) and
  // read the document back, so callers report what is really in it.
  W.commit = async (bytes, { removing = false, copy = false } = {}) => {
    if (copy || !W.canSwap()) {
      if (!W.canCopy()) throw new Error("This version of Office can't open a rebuilt file. Use the Wild Themes website or browser extension instead.");
      await W.openCopy(bytes);
      return { copied: true, bytes };
    }
    await swapIn(bytes, removing);
    return { copied: false, bytes: await W.read() };
  };
  // One line on what happened, in the same words wherever the pane is.
  W.report = (res, info, expectId, verb) => {
    if (res.copied) return [`${verb} The result opened in a new window: save it there (File › Save As, you can replace this file). This window is unchanged.`, "ok"];
    if (expectId === null) return info.status === "none" ? [`${verb} Checked: no theme left in this file.`, "ok"]
      : ["Most of the theme was removed, but Office kept some of it in this window. Use “Open a themed copy” for a clean file.", "bad"];
    if (info.theme && info.theme.id === expectId && info.status === "intact") return [`${verb} Checked: the theme is in this file and intact.`, "ok"];
    const why = (info.problems || []).map((p) => p.text).join(" ");
    return [`Office only took part of the theme into this window.${why ? " " + why : ""} “Open a themed copy” gives you the complete file.`, "bad"];
  };
  W.where = () => ({ word: "Word document", ppt: "PowerPoint deck", excel: "Excel workbook" })[W.host];
  W.note = () => ({
    word: "In Word the scene fills the page behind a light reading panel, and headings, lists, tables, links, headers and footers follow the theme. Changes are made in this window; “Open a themed copy” gives an exact copy instead.",
    ppt: "In PowerPoint the scene sits on the slide masters behind a see-through panel; titles, bullets, tables and charts follow the theme, and slides can get a transition. Animated scenes play in slide shows.",
    excel: "In Excel the theme colours and fonts drive cell styles, tables, charts and tab colours, with a pale picture behind every sheet. Office doesn't let add-ins swap a whole workbook in place, so the themed workbook opens in a new window for you to save.",
  })[W.host];

  // ---------------------------------------------------------------- open by itself when a themed file opens
  const store = { get: (k, d) => { try { const v = localStorage.getItem("wt:" + k); return v == null ? d : JSON.parse(v); } catch (_) { return d; } },
    set: (k, v) => { try { localStorage.setItem("wt:" + k, JSON.stringify(v)); } catch (_) { /* private mode */ } } };
  W.autoSupported = () => supports("SharedRuntime", "1.1") && !!(Office.addin && Office.addin.setStartupBehavior);
  W.auto = () => store.get("auto", false);
  W.setAuto = async (on) => { await Office.addin.setStartupBehavior(on ? Office.StartupBehavior.load : Office.StartupBehavior.none); store.set("auto", on); };
  async function hasThemePart() {
    const ok = W.host === "word" ? supports("WordApi", "1.4") : W.host === "ppt" ? supports("PowerPointApi", "1.7") : supports("ExcelApi", "1.5");
    if (!ok) return false;
    return run()(async (ctx) => {
      const parts = docOf(ctx).customXmlParts.getByNamespace("urn:wild-themes:document-theme"); parts.load("items"); await ctx.sync();
      return parts.items.length > 0;
    });
  }

  W.ready = new Promise((ok) => Office.onReady((o) => {
    const host = o && o.host === Office.HostType.Word ? "word" : o && o.host === Office.HostType.PowerPoint ? "ppt" : o && o.host === Office.HostType.Excel ? "excel" : null;
    if (!host) { forget(); return ok(null); }        // ?host=office opened in a normal browser
    const url = Office.context.document && Office.context.document.url;
    Object.assign(W, { active: true, host, kind: { word: "docx", ppt: "pptx", excel: "xlsx" }[host], name: url ? decodeURIComponent(url.split(/[\\/]/).pop()) : "" });
    document.documentElement.dataset.office = host;
    ok(W);
    // Started in the background (shared runtime) for a file that carries a theme: show the pane.
    if (W.autoSupported() && W.auto()) hasThemePart().then((has) => has && Office.addin.showAsTaskpane()).catch(() => {});
  }));
  self.WTOffice = W;
})();
