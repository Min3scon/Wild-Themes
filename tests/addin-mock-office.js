// Stand-in for office.js used by addin.test.js: just enough of Office/Word/PowerPoint/Excel for the
// task pane, backed by window.__doc (the "open file"). Imports are simulated as perfect, so these
// tests check the pane's own logic; real-Office behaviour still needs testing in Office itself.
(() => {
  const host = window.__host || "Word";
  const calls = (window.__calls = []);
  const b64ToBytes = (b) => Uint8Array.from(atob(b), (c) => c.charCodeAt(0));
  const xmlParts = (window.__xmlParts = window.__xmlParts || []);
  const coll = (items) => ({ items, load() { return this; } });
  const ctxBase = { sync: async () => {} };
  const customXml = {
    getByNamespace: (ns) => coll(xmlParts.filter((p) => p.xml.includes(ns)).map((p) => ({ delete() { xmlParts.splice(xmlParts.indexOf(p), 1); calls.push(["xml.delete"]); } }))),
    add(xml) { xmlParts.push({ xml }); calls.push(["xml.add", xml.length]); },
  };
  const props = { load() { return this; }, items: [], add(k, v) { calls.push(["prop.add", k, v]); } };
  window.Office = {
    HostType: { Word: "Word", PowerPoint: "PowerPoint", Excel: "Excel" },
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    FileType: { Compressed: "compressed" }, EventType: {}, StartupBehavior: { load: "load", none: "none" },
    onReady(cb) { setTimeout(() => cb({ host }), 0); return Promise.resolve({ host }); },
    addin: { async setStartupBehavior(b) { calls.push(["startup", b]); }, async showAsTaskpane() { calls.push(["show"]); } },
    context: {
      requirements: { isSetSupported: (set, v) => !(window.__unsupported || []).includes(set) },
      document: {
        getFileAsync(type, opts, cb) {
          const data = window.__doc, size = opts.sliceSize, n = Math.max(1, Math.ceil(data.length / size));
          cb({ status: "succeeded", value: { sliceCount: n, closeAsync() {}, getSliceAsync(i, cb2) { cb2({ status: "succeeded", value: { data: Array.from(data.subarray(i * size, (i + 1) * size)) } }); } } });
        },
      },
    },
  };
  window.Word = {
    async run(fn) {
      const ctx = { ...ctxBase,
        document: {
          body: { contentControls: { getByTag: () => coll([]) } },
          sections: coll([]),
          insertFileFromBase64(b64, where, opts) { calls.push(["insertFile", where, opts]); window.__doc = b64ToBytes(b64); },
          customXmlParts: customXml, properties: { customProperties: props },
          background: { fill: { clear() { calls.push(["bg.clear"]); } } },
        },
        application: { createDocument(b64) { return { open() { calls.push(["openCopy", b64.length]); window.__copy = b64ToBytes(b64); } }; } },
      };
      return fn(ctx);
    },
  };
  window.PowerPoint = {
    async run(fn) {
      const slides = { items: [{ id: "256" }, { id: "257" }], load() { return this; }, getItem: (id) => ({ delete() { calls.push(["slide.delete", id]); } }) };
      return fn({ ...ctxBase, presentation: { slides, customXmlParts: customXml, properties: { customProperties: props },
        insertSlidesFromBase64(b64, o) { calls.push(["insertSlides", o]); window.__doc = b64ToBytes(b64); } } });
    },
    async createPresentation(b64) { calls.push(["openCopy", b64.length]); window.__copy = b64ToBytes(b64); },
  };
  window.Excel = {
    async run(fn) { return fn({ ...ctxBase, workbook: { customXmlParts: customXml } }); },
    async createWorkbook(b64) { calls.push(["openCopy", b64.length]); window.__copy = b64ToBytes(b64); },
  };
})();
