// Lets the Wild Themes website switch the extension's theme. Runs only on that site.
// The page posts {wt:"use", theme} to itself; this script checks it and saves it.
(() => {
  const ok = (h) => typeof h == "string" && /^#[0-9a-f]{6}$/i.test(h);
  document.documentElement.dataset.wildThemesExtension = chrome.runtime.getManifest().version;
  addEventListener("message", async (e) => {
    const m = e.data;
    if (e.source !== window || e.origin !== location.origin || !m || m.wt !== "use" || !m.theme) return;
    const t = m.theme, reply = (ok, text) => postMessage({ wt: "used", ok, text }, location.origin);
    try {
      if (t.builtin) {                                         // a built-in theme: just select it
        if (!/^[a-z0-9]+$/.test(t.id)) throw 0;
        await chrome.storage.sync.set({ theme: t.id, on: true });
        return reply(true, t.name);
      }
      // your own or a community theme: store it like one made in the extension, then select it
      if (!t.name || String(t.name).length > 40 || !ok(t.accent)) throw 0;
      if (t.image && (typeof t.image != "string" || !t.image.startsWith("data:image/") || t.image.length > 3e6)) throw 0;
      const { customs = [] } = await chrome.storage.local.get("customs");
      const id = "c_" + (t.key || Date.now().toString(36)).replace(/[^a-z0-9-]/gi, "").slice(0, 40);
      const theme = { id, name: String(t.name).slice(0, 40), mode: t.mode == "light" ? "light" : "dark", accent: t.accent, accent2: ok(t.accent2) ? t.accent2 : t.accent,
        fx: "dot", thumb: t.thumb || "", dim: 0 };
      if (t.image) { theme.image = id; await chrome.storage.local.set({ ["img:" + id]: t.image }); }
      else { if (!Array.isArray(t.colors) || !t.colors.every(ok)) throw 0; theme.colors = t.colors.slice(0, 3); theme.particles = t.effect || "none"; }
      await chrome.storage.local.set({ customs: [...customs.filter((c) => c.id !== id), theme] });
      await chrome.storage.sync.set({ theme: id, on: true });
      reply(true, theme.name);
    } catch (_) { reply(false); }
  });
})();
