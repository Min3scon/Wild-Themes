// Wild Themes — the live theme catalogue. The website publishes catalog.json (built-in themes plus
// approved community themes); the Office add-in and the browser extension read it so new themes
// show up without a new download. Everything fetched is checked before use, and the tools fall back
// to the themes they ship with when the site can't be reached.
(() => {
  const SITE = "https://min3scon.github.io/Wild-Themes/";
  const HEX = /^#[0-9a-fA-F]{6}$/, ID = /^[a-z0-9][a-z0-9-]{0,47}$/, REL = /^(extension|community)\/[A-Za-z0-9_\/.-]+$/;
  const okPath = (p) => p == null || (typeof p === "string" && REL.test(p) && !p.includes(".."));
  const str = (s, n) => typeof s === "string" && s.length > 0 && s.length <= n;

  // Keep only well-formed entries; anything odd is dropped, never trusted.
  function clean(cat) {
    if (!cat || cat.format !== 1 || !Array.isArray(cat.themes)) throw new Error("Not a Wild Themes catalogue.");
    const themes = cat.themes.filter((t) => t && ID.test(t.id) && str(t.name, 60) && (t.mode === "dark" || t.mode === "light") &&
      HEX.test(t.accent || "") && (!t.accent2 || HEX.test(t.accent2)) && (!t.colors || (Array.isArray(t.colors) && t.colors.length <= 4 && t.colors.every((c) => HEX.test(c)))) &&
      ["builtin", "community"].includes(t.kind) && ["preview", "scene", "backdrop", "backdropWide", "teams", "image", "thumb"].every((k) => okPath(t[k])) &&
      (!t.doc || Object.entries(t.doc).every(([k, v]) => typeof v === "string" && v.length <= 40 && /^[a-z0-9]+$/i.test(k))));
    return { version: String(cat.version || ""), updated: String(cat.updated || ""), themes };
  }
  async function load(base = SITE, ms = 8000) {
    const c = new AbortController(), timer = setTimeout(() => c.abort(), ms);
    try {
      const r = await fetch(new URL("catalog.json", base).href, { cache: "no-cache", signal: c.signal });
      if (!r.ok) throw new Error("HTTP " + r.status);
      return clean(await r.json());
    } finally { clearTimeout(timer); }
  }
  // Scene CSS from the site is only ever gradients and animations: refuse anything that could load or run something.
  function safeCss(css) {
    if (typeof css !== "string" || css.length > 65536) return null;
    return /url\s*\(|@import|expression\s*\(|javascript:|behavior\s*:|-moz-binding|\[value/i.test(css) ? null : css;
  }
  const abs = (p, base = SITE) => (p ? new URL(p, base).href : null);
  // Built-in themes this copy doesn't ship with yet (added on the site since).
  const newBuiltins = (cat, local) => cat.themes.filter((t) => t.kind === "builtin" && !local[t.id]);
  // A community theme in the shape the tools use for "your own" themes.
  function asCustom(t, base = SITE) {
    const colors = t.colors && t.colors.length === 3 ? t.colors : [t.mode === "light" ? "#f4f1ff" : "#0d0b1a", t.accent, t.accent2 || t.accent];
    return { id: t.id, name: t.name, mode: t.mode, accent: t.accent, accent2: t.accent2 || t.accent, colors, particles: t.effect || "none",
      image: t.image ? t.id : null, imageUrl: abs(t.image, base), thumb: abs(t.thumb, base), author: t.author, community: 1 };
  }
  self.WTCatalog = { SITE, load, clean, safeCss, abs, newBuiltins, asCustom };
})();
