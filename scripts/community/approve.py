#!/usr/bin/env python3
"""Validate an approved theme submission (GitHub issue form) and add it to the gallery.

Run by .github/workflows/community.yml when a maintainer adds the `approved` label.
Inputs (env): ISSUE_NUMBER, ISSUE_AUTHOR, ISSUE_BODY. Optional: REPO_ROOT (default: repo
this script lives in), RESULT_FILE (markdown message for the issue comment).
Outputs: community/<slug>/{bg.jpg,thumb.webp}, an updated community/index.json, and
`status` (ok|error) + `name` on $GITHUB_OUTPUT.
Exit code 0 for handled outcomes (including a rejected submission); non-zero only on bugs.
"""
import colorsys, datetime, io, json, os, random, re, sys, unicodedata, urllib.parse, urllib.request
from pathlib import Path

from PIL import Image, ImageDraw

MAX_BYTES = 2 * 1024 * 1024          # picture upload limit (checked while downloading)
MIN_W, MIN_H = 1280, 720             # smallest picture accepted
MAX_SIDE = 8000                      # largest picture side accepted
BG_MAX_BYTES = 450 * 1024            # published background size cap
THUMB_MAX_BYTES = 60 * 1024          # published thumbnail size cap (960x540 card picture)
Image.MAX_IMAGE_PIXELS = MAX_SIDE * MAX_SIDE

LABELS = {"Theme name": "name", "Your name": "author", "Dark or light": "mode", "Type": "kind", "Colours": "colours",
          "Effect": "effect", "Accent colour": "accent", "Picture": "picture", "Rights": "rights"}
EFFECTS = {"none": "none", "twinkling stars": "stars", "snow": "snow", "bubbles": "bubbles", "confetti": "confetti"}
ALLOWED_HOSTS = ("github.com", "githubusercontent.com")
HEX = re.compile(r"^#(?:[0-9a-fA-F]{3}|[0-9a-fA-F]{6})$")


class Rejected(Exception):
    """A problem the submitter can fix."""


# ---------------------------------------------------------------- parsing
def parse_form(body):
    """Issue-form markdown ("### Label\\n\\nvalue") -> {field id: value}."""
    out = {}
    for block in re.split(r"^### ", body.replace("\r\n", "\n"), flags=re.M)[1:]:
        label, _, value = block.partition("\n")
        key = LABELS.get(label.strip())
        if not key:
            continue
        value = value.strip()
        out[key] = "" if value == "_No response_" else value
    return out


def norm_hex(h):
    h = h.strip()
    if not HEX.match(h):
        return None
    h = h[1:].lower()
    if len(h) == 3:
        h = "".join(c * 2 for c in h)
    return "#" + h


def slugify(name):
    s = unicodedata.normalize("NFKD", name).encode("ascii", "ignore").decode()
    s = re.sub(r"[^a-z0-9]+", "-", s.lower()).strip("-")
    return s[:40].strip("-")


# ---------------------------------------------------------------- download
def host_ok(url):
    p = urllib.parse.urlsplit(url)
    host = (p.hostname or "").lower()
    if os.environ.get("ALLOW_TEST_HOST") == "1" and p.scheme == "http" and host in ("127.0.0.1", "localhost"):
        return True                                            # tests only, never set in the workflow
    if p.scheme != "https":
        return False
    if host == "github.com":
        return p.path.startswith("/user-attachments/")
    return any(host == h or host.endswith("." + h) for h in ALLOWED_HOSTS if h != "github.com")


class SafeRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        if not host_ok(newurl):
            raise Rejected("The picture link points outside GitHub. Drag the picture straight into the issue form instead of linking to another site.")
        return super().redirect_request(req, fp, code, msg, headers, newurl)


def find_picture_url(text):
    urls = re.findall(r'!\[[^\]]*\]\((\S+?)\)', text) + re.findall(r'<img[^>]+src="([^"]+)"', text)
    urls += [u for u in re.findall(r"https?://[^\s)\"'<>]+", text) if u not in urls]
    urls = list(dict.fromkeys(urls))
    if not urls:
        raise Rejected("No picture found. Drag your picture into the **Picture** box so GitHub uploads it.")
    if len(urls) > 1:
        raise Rejected("Please add just one picture.")
    if not host_ok(urls[0]):
        raise Rejected("The picture has to be uploaded to GitHub: drag it into the **Picture** box instead of pasting a link to another site.")
    return urls[0]


def download(url):
    opener = urllib.request.build_opener(SafeRedirect)
    req = urllib.request.Request(url, headers={"User-Agent": "wild-themes-approval"})
    with opener.open(req, timeout=30) as r:
        size = r.headers.get("Content-Length")
        if size and size.isdigit() and int(size) > MAX_BYTES:
            raise Rejected(f"The picture is {int(size) / 1048576:.1f} MB. The limit is 2 MB: save it as a JPG or make it smaller.")
        data = bytearray()
        while chunk := r.read(65536):
            data += chunk
            if len(data) > MAX_BYTES:
                raise Rejected("The picture is over 2 MB. Save it as a JPG or make it smaller, then upload it again.")
    return bytes(data)


# ---------------------------------------------------------------- images
def save_capped(img, path, fmt, cap, quality, floor):
    while True:
        buf = io.BytesIO()
        if fmt == "JPEG":
            img.save(buf, "JPEG", quality=quality, optimize=True, progressive=True)
        else:
            img.save(buf, "WEBP", quality=quality, method=6)
        if buf.tell() <= cap or quality <= floor:
            break
        quality -= 6
    path.write_bytes(buf.getvalue())
    return buf.tell()


def cover(img, w, h):
    sw, sh = img.size
    s = max(w / sw, h / sh)
    img = img.resize((max(w, round(sw * s)), max(h, round(sh * s))), Image.LANCZOS)
    l, t = (img.width - w) // 2, (img.height - h) // 2
    return img.crop((l, t, l + w, t + h))


def picture_theme(raw, mode, outdir):
    try:
        im = Image.open(io.BytesIO(raw))
        fmt = im.format
    except Exception:
        raise Rejected("That file isn't a picture we can read. Use a PNG, JPG or WebP.")
    if fmt not in ("PNG", "JPEG", "WEBP", "GIF"):
        raise Rejected(f"{fmt or 'This'} files aren't supported. Use a PNG, JPG or WebP.")
    w, h = im.size
    if w < MIN_W or h < MIN_H:
        raise Rejected(f"The picture is {w}×{h}. It needs to be at least {MIN_W}×{MIN_H} so it stays sharp.")
    if max(w, h) > MAX_SIDE:
        raise Rejected(f"The picture is {w}×{h}. The longest side can be at most {MAX_SIDE} pixels.")
    try:
        im.seek(0)
        im.load()
    except Exception:
        raise Rejected("That picture looks damaged. Try saving it again.")
    if im.mode in ("RGBA", "LA", "P"):
        im = im.convert("RGBA")
        base = Image.new("RGB", im.size, (0, 0, 0) if mode == "dark" else (255, 255, 255))
        base.paste(im, mask=im.split()[3])
        im = base
    else:
        im = im.convert("RGB")
    out_w = min(1920, w if w * 9 <= h * 16 else round(h * 16 / 9))
    out_w = max(MIN_W, out_w)
    bg = cover(im, out_w, round(out_w * 9 / 16))           # fresh image: no EXIF/metadata carried over
    bg_bytes = save_capped(bg, outdir / "bg.jpg", "JPEG", BG_MAX_BYTES, 82, 40)
    th_bytes = save_capped(cover(bg, 960, 540), outdir / "thumb.webp", "WEBP", THUMB_MAX_BYTES, 78, 30)
    r, g, b = bg.resize((1, 1), Image.BOX).getpixel((0, 0))
    hh, ll, ss = colorsys.rgb_to_hls(r / 255, g / 255, b / 255)
    accent = hls_hex(hh, 0.62, max(ss, 0.6))
    accent2 = hls_hex((hh + 0.12) % 1, 0.7, max(ss, 0.5))
    return accent, accent2, bg_bytes, th_bytes


def hls_hex(h, l, s):
    return "#" + "".join(f"{round(v * 255):02x}" for v in colorsys.hls_to_rgb(h, l, s))


def rgb(h):
    return tuple(int(h[i:i + 2], 16) for i in (1, 3, 5))


def glow(img, color, cx, cy, rx, ry):
    """Paste a soft radial glow fading to transparent (like CSS radial-gradient)."""
    mask = Image.radial_gradient("L").point(lambda v: max(0, 255 - v * 255 // 179))  # 255 centre -> 0 at radius (edge = 179)
    mask = mask.resize((max(1, int(rx * 2)), max(1, int(ry * 2))), Image.BILINEAR)
    img.paste(Image.new("RGB", mask.size, rgb(color)), (int(cx - rx), int(cy - ry)), mask)


def colours_thumb(colors, effect, seed, outdir):
    w, h = 960, 540
    img = Image.new("RGB", (w, h), rgb(colors[0]))
    glow(img, colors[1], 0, 0, w * 0.72, h * 0.54)          # matches the site: 120%×90% ellipse, fade by 60%
    glow(img, colors[2], w, h, w * 0.72, h * 0.54)
    d, rnd, k = ImageDraw.Draw(img), random.Random(seed), w / 320      # dots were tuned at 320 px wide
    for _ in range({"stars": 60, "snow": 70, "bubbles": 18, "confetti": 44}.get(effect, 0)):
        x, y = rnd.uniform(0, w), rnd.uniform(0, h)
        if effect == "stars":
            r = rnd.choice((0.7, 1, 1.4)) * k; d.ellipse((x - r, y - r, x + r, y + r), fill=(255, 255, 255))
        elif effect == "snow":
            r = rnd.uniform(1, 2.4) * k; d.ellipse((x - r, y - r, x + r, y + r), fill=(255, 255, 255))
        elif effect == "bubbles":
            r = rnd.uniform(3, 8) * k; d.ellipse((x - r, y - r, x + r, y + r), outline=(255, 255, 255), width=2)
        else:
            c = rnd.choice(("#ff3d7f", "#ffbe0b", "#3a86ff", "#06d6a0")); d.rectangle((x, y, x + 3 * k, y + 5 * k), fill=rgb(c))
    return save_capped(img, outdir / "thumb.webp", "WEBP", THUMB_MAX_BYTES, 82, 30)


# ---------------------------------------------------------------- main flow
def build(form, issue, login, root):
    problems = []
    name = re.sub(r"\s+", " ", form.get("name", "")).strip()
    if not name:
        problems.append("Add a **Theme name**.")
    elif len(name) > 30:
        problems.append(f"The theme name is {len(name)} characters; the limit is 30.")
    elif not all(c.isalnum() or c in " -'&.!" for c in name):
        problems.append("The theme name can only use letters, numbers, spaces and - ' & . !")
    slug = slugify(name) if name else ""
    if name and not slug:
        problems.append("The theme name needs at least one plain letter or number.")
    index_path = root / "community" / "index.json"
    index = json.loads(index_path.read_text()) if index_path.exists() else {"themes": []}
    if slug and (any(t["id"] == "c-" + slug for t in index["themes"]) or (root / "community" / slug).exists()):
        problems.append(f"There's already a theme called “{name}”. Pick a different name.")

    mode = form.get("mode", "").strip().lower()
    if mode not in ("dark", "light"):
        problems.append("Choose **Dark** or **Light**.")
    kind_raw = form.get("kind", "").lower()
    kind = "picture" if kind_raw.startswith("picture") else "colours" if kind_raw.startswith("colour") else ""
    if not kind:
        problems.append("Choose a **Type**: Picture or Colours.")
    effect = EFFECTS.get(form.get("effect", "").strip().lower() or "none")
    if effect is None:
        problems.append("Choose an **Effect** from the list.")
    accent = None
    if form.get("accent"):
        accent = norm_hex(form["accent"])
        if not accent:
            problems.append(f"The accent colour “{form['accent']}” isn't a hex colour like #ff2bd6.")
    colors = None
    if kind == "colours":
        parts = re.split(r"[\s,;]+", form.get("colours", "").strip())
        parts = [p for p in parts if p]
        colors = [norm_hex(p) for p in parts]
        if len(parts) != 3 or not all(colors):
            problems.append("For Colours, give exactly 3 hex colours (background, glow 1, glow 2), like `#120a2a #ff2bd6 #21e6ff`.")
    if not re.search(r"- \[[xX]\]", form.get("rights", "")):
        problems.append("Tick the box to confirm you made the picture or have the right to share it.")
    url = None
    if kind == "picture":
        try:
            url = find_picture_url(form.get("picture", ""))
        except Rejected as e:
            problems.append(str(e))
    if problems:
        raise Rejected("\n".join("- " + p for p in problems))

    outdir = root / "community" / slug
    outdir.mkdir(parents=True)
    try:
        if kind == "picture":
            auto1, auto2, bg_b, th_b = picture_theme(download(url), mode, outdir)
            image, a1, a2 = f"community/{slug}/bg.jpg", accent or auto1, auto2
        else:
            th_b, bg_b, image = colours_thumb(colors, effect, slug, outdir), 0, None
            a1, a2 = accent or colors[1], colors[2]
    except BaseException:
        for f in outdir.glob("*"):
            f.unlink()
        outdir.rmdir()
        raise
    entry = {"id": "c-" + slug, "name": name, "author": (form.get("author") or "").strip()[:40] or login,
             "mode": mode, "motion": kind == "colours", "effect": effect, "colors": colors,
             "accent": a1, "accent2": a2, "image": image, "thumb": f"community/{slug}/thumb.webp",
             "issue": issue, "added": datetime.date.today().isoformat()}
    index["themes"].insert(0, entry)
    index["themes"].sort(key=lambda t: (t["added"], t["issue"]), reverse=True)
    index_path.parent.mkdir(parents=True, exist_ok=True)
    index_path.write_text(json.dumps(index, indent=2, ensure_ascii=False) + "\n")
    return entry, bg_b, th_b


def emit(status, message, name=""):
    result = os.environ.get("RESULT_FILE", "community-result.md")
    Path(result).write_text(message + "\n")
    out = os.environ.get("GITHUB_OUTPUT")
    if out:
        with open(out, "a") as f:
            f.write(f"status={status}\nname={name}\n")
    print(f"[{status}] {message}")


def main():
    root = Path(os.environ.get("REPO_ROOT") or Path(__file__).resolve().parents[2])
    issue = int(os.environ["ISSUE_NUMBER"])
    login = os.environ.get("ISSUE_AUTHOR", "someone")
    try:
        entry, bg_b, th_b = build(parse_form(os.environ.get("ISSUE_BODY", "")), issue, login, root)
    except Rejected as e:
        emit("error", "Thanks for sending a theme! A few things need fixing before it can go live:\n\n" + str(e) +
             "\n\nEdit this issue to fix them, and a maintainer will approve it again.")
        return 0
    emit("ok", f"Your theme **{entry['name']}** is now in the Wild Themes gallery. Thanks for sharing it!", entry["name"])
    print(f"bg {bg_b} B, thumb {th_b} B")
    return 0


if __name__ == "__main__":
    sys.exit(main())
