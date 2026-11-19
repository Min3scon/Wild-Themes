# Wild Themes add-in for Word, PowerPoint and Excel

Adds **Home › Themes** to Word, PowerPoint and Excel on Windows, Mac and the web. The pane **is the website**: the same File studio and theme gallery (same HTML, CSS and scripts), opened as `studio.html?host=office`. Inside Office the pages work on the open document instead of a downloaded file:

- the File studio reads the open file, shows its theme and whether it's intact, and applies, changes, tunes (look and effects), restores, removes or undoes it, then reads the file back and shows exactly what's in it;
- in the gallery (Themes in the site menu), a theme's **Save it into this document** and Your own theme's drop box write into the open file; Word starter opens as a new Word document;
- "Open a themed copy" opens the exact file Wild Themes wrote in a new window; Excel always works this way, as Office doesn't let add-ins swap a whole workbook;
- every theme on the website is listed, including approved community themes and imported `.wildtheme` files, without updating the add-in.

`office-host.js` is the only add-in code: it reads and writes the open document through Office.js and is loaded by the pages only when they run as the add-in. Everything else is the site and the shared core (`../extension/lib/wtcore.js`), served from the website, so the add-in updates itself. `taskpane.html` forwards manifests from before 3.1.

Install: see the main README ("Office add-in"). Requirements: Word API 1.5 (in-place) or 1.3 (copy), PowerPoint API 1.2 (in-place) or 1.1 (copy), Excel API 1.8 (copy). Custom XML sync needs Word API 1.4 / PowerPoint API 1.7. Opening by itself needs the shared runtime (SharedRuntime 1.1).

Word swaps documents in place up to about 4 MB; use "Open a themed copy" for bigger files. Comments and tracked changes may not survive an in-place swap; the copy keeps everything.

## Trying local changes in desktop Office (Windows)

Office only loads add-ins over HTTPS. Serve the checkout and put any HTTPS tunnel in front of it, for example:

```bash
node scripts/serve.js 8091
cloudflared tunnel --url http://127.0.0.1:8091
```

Copy `manifest.xml`, replace `https://min3scon.github.io/Wild-Themes` and the AppDomain with the tunnel's address, and give it a different `<Id>`. Register it under `HKCU\Software\Microsoft\Office\16.0\WEF\Developer` (value name: the Id, value: the manifest's path) and restart Word, PowerPoint or Excel. Remove the value when you're done.
