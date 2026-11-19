# Wild Themes

Animated themes for Word, PowerPoint, Excel and Teams, and a way to save a theme **inside** a file so everyone who opens it sees it, with nothing installed.

**Website:** https://min3scon.github.io/Wild-Themes/ (theme a file right in your browser)

| Folder | What it is |
|---|---|
| `/` (`index.html`, `site.*`) | The website: theme gallery, live preview, and "theme a file" (drop a .docx / .pptx / .xlsx, the themed copy downloads straight away) |
| `extension/` | Browser extension for Word, PowerPoint, Excel and Teams on the web (Chrome, Edge, Brave, Arc). Load this folder unpacked, or use `downloads/wild-themes-extension.zip` |
| `addin/` | Office add-in for Word and PowerPoint, desktop and web: one click saves a theme into the open file. Upload `addin/manifest.xml` in Office (Home › Add-ins › My Add-ins › Upload My Add-in) |
| `documents/` | 18 Word documents with a theme already inside |
| `downloads/` | Zips for the extension and the documents |

## What goes into a themed file

- Theme colours and fonts, and matching title, heading and quote styles
- A backdrop behind the text on every page or slide, with a reading panel so text stays readable
- Word: the backdrop sits in locked content controls; PowerPoint: on the slide master, locked
- Hidden file properties naming the theme (`WildTheme`, `WildThemeName`)

The backdrop animates in desktop Word and PowerPoint. Office on the web shows pictures still. Excel gets a pale sheet background (desktop) plus theme colours and fonts.

## Performance

The extension has no animation loop and no per-frame code. Scenes are CSS gradients painted once; the browser only slides or fades whole layers. Word loads about 9 KB of code plus about 2 KB for the theme in use. Measured cost while typing: about 0.06 ms of script per keystroke with sparks on, none with sparks off.

Unofficial. Not made by or affiliated with Microsoft.
