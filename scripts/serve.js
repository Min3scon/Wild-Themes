// Serve the site from this checkout for trying the Office add-in against local changes:
//   node scripts/serve.js [port]         (default 8080; put an HTTPS tunnel in front for Office)
// Only site files: no dotfiles, node_modules or test output.
const http = require("http"), fs = require("fs"), path = require("path");
const ROOT = path.join(__dirname, ".."), PORT = Number(process.argv[2]) || 8080;
const TYPES = { ".html": "text/html; charset=utf-8", ".js": "text/javascript", ".css": "text/css", ".json": "application/json", ".xml": "application/xml",
  ".webp": "image/webp", ".gif": "image/gif", ".png": "image/png", ".jpg": "image/jpeg", ".svg": "image/svg+xml", ".zip": "application/zip",
  ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document" };
http.createServer((req, res) => {
  let rel = decodeURIComponent(new URL(req.url, "http://x").pathname);
  if (rel.endsWith("/")) rel += "index.html";
  const p = path.join(ROOT, rel), parts = path.relative(ROOT, p).split(path.sep);
  if (parts[0] === ".." || parts.some((s) => s.startsWith(".") || s === "node_modules") || parts[0] === "tests" || !fs.existsSync(p) || fs.statSync(p).isDirectory()) {
    res.writeHead(404); return res.end("Not found");
  }
  res.writeHead(200, { "Content-Type": TYPES[path.extname(p)] || "application/octet-stream", "Cache-Control": "no-store" });
  fs.createReadStream(p).pipe(res);
}).listen(PORT, "127.0.0.1", () => console.log(`Wild Themes on http://localhost:${PORT}/`));
