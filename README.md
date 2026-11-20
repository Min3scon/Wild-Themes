# Wild Themes

Animated themes for Word, PowerPoint, Excel and Teams, saved **inside** your files so everyone who opens them sees the theme, with nothing installed.

**Website:** https://min3scon.github.io/Wild-Themes/ · **File studio:** https://min3scon.github.io/Wild-Themes/studio.html

| Folder | What it is |
|---|---|
| `/` (`index.html`, `site.*`) | Theme gallery, live previews, "theme a file", make and share your own theme |
| `studio.html` | File studio: open any Word, PowerPoint or Excel file to see its theme, check it's intact, change it, tune it, restore it or remove it. Runs in the browser; nothing is uploaded |
| `extension/` | Browser extension (Chrome, Edge, Brave, Arc, Firefox): live scenes in Word, PowerPoint, Excel, Teams and Google Docs, Sheets and Slides on the web, plus the file tools. `extension/lib/` is the shared core every tool uses. `scripts/build-extension.js` builds the Chrome-family and Firefox packages (the site's downloads are rebuilt on every deploy). How to load them in each browser: `extension/README.md` |
| `addin/` | Office add-in for Word, PowerPoint and Excel (Windows, Mac, web) |
| `teams/` | Teams app: File studio and the theme gallery as tabs |
| `catalog.json` | Live theme list (built-in + approved community themes), rebuilt on every deploy. The add-in and the extension read it, so new themes arrive without a new download |
| `scripts/theme-assets.js` | Makes a new theme's pictures (gallery card, Word and PowerPoint backdrops, Teams background) from its scene file: `node scripts/theme-assets.js <id>`. A theme is one scene in `extension/scenes/`, an entry in `extension/lib/themes.js`, and these pictures |
| `documents/` | One Word document per theme, ready to use |
| `tests/` | Regression tests (run in CI on every push) |

## What goes into a themed file

The theme becomes part of the file's own structure:

| | Word (.docx) | PowerPoint (.pptx) | Excel (.xlsx) |
|---|---|---|---|
| Theme part | Colours + fonts (`word/theme/theme1.xml`) | Every used master's theme | `xl/theme/theme1.xml` |
| Styles | Normal, Title, Subtitle, Heading 1–6, quotes, links, header/footer, caption | Master title/body text styles, bullet colours | Normal font, Title, Heading 1–4, Total, Hyperlink |
| Tables | Table Grid + "Wild Theme Table" with header row; default for new tables | Default table style follows the accent | Default table style follows the accent |
| Lists | Bullet and number colours | Bullet colours | — |
| Existing colours | Text, shading, shapes and charts using the old theme's colours move to the new ones | Shapes and charts likewise | Cells and charts likewise |
| Scene | Page background + reading panel on every page | Scene + see-through panel on the masters (animated in slide shows) | Pale picture behind every sheet |
| Effects | Optional heading glow | Optional slide transitions; animated or still scene | Sheet tab colours |

Plus a custom XML part (`urn:wild-themes:document-theme`) holding the theme's full definition, its settings, SHA-256 fingerprints and a backup of the file's original theme and styles. Any copy of the tools can recognise the theme, tell whether it was changed or partly deleted, rebuild it from the file alone, or remove it and put the original look back. Custom properties (`WildTheme`, `WildThemeName`, `WildThemeVersion`) are a second marker that survives editors which drop custom XML.

Office files aren't tamper-proof: the fingerprints detect accidental changes, they don't prevent edits.

## Words typed and the secret theme (extension)

The extension counts the words you type in each theme (shown on each card in the popup, with the total). Only real key presses count, one finished word at a time: pasted, dropped or dictated text doesn't, though paste still works as usual in your documents. Every typed word has a 1 in 37,500 chance (`ODDS` in `extension/lib/rewards.js`) to unlock a secret theme. Until then it's a greyed-out "???" card with no picture or name. Finding it plays a cutscene in the page; passing 1,000, 10,000, 50,000 and 100,000 words plays a shorter one. Counts stay in the browser (they're never sent anywhere); the unlock is saved to your browser sync. The secret theme isn't on the website, in the File studio or in `catalog.json`. That makes it a surprise, not a lock: the extension's code is open.

## Office add-in

Install once; the code and the theme list update themselves from the website.

- **Web (Word, PowerPoint, Excel):** Home › Add-ins › More Add-ins › My Add-ins › **Upload My Add-in** › `addin/manifest.xml`.
- **Windows:** put `manifest.xml` in a shared folder, add it under File › Options › Trust Center › Trust Center Settings › Trusted Add-in Catalogs (tick "Show in Menu"), restart, then Home › Add-ins › Shared Folder.
- **Mac:** copy `manifest.xml` into `~/Library/Containers/com.microsoft.Word/Data/Documents/wef` (and the same for `com.microsoft.Powerpoint`, `com.microsoft.Excel`), restart.

Then Home › **Themes**. The pane is the website itself (the File studio and the gallery), working on the open file: it reads it, shows its theme and whether it's intact, and lets you apply, change, tune (look and effects), restore or remove it. Word and PowerPoint are changed in place; "Open a themed copy" gives an exact rebuilt file in a new window instead. Excel doesn't let add-ins replace a whole workbook, so the themed workbook opens as a copy to save. Turn on "Open Wild Themes by itself…" to have the pane open when a themed file is opened (Office versions with the shared runtime).

## Teams

Teams previews Office files with Office for the web, so a themed file shows its theme in Teams the way Office for the web draws it; the downloaded file is unchanged and keeps everything. To manage themes inside Teams, upload `downloads/wild-themes-teams.zip` (Apps › Manage your apps › Upload an app, if your organisation allows it). For meetings, every theme has a 1920×1080 background (Video effects › Add new).

## Tests

```bash
cd tests && npm ci && npm test
```

Covers apply / re-apply / change / remove / restore for every format, exact restore of the original parts, preservation of content (formulas, validation, names, tables, charts, macros, animations, transitions, links, lists), settings, tamper detection, LibreOffice round trips and renders, the add-in task pane against a stand-in Office, and the extension's live catalogue.

Unofficial. Not made by or affiliated with Microsoft.
