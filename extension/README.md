# Wild Themes for Microsoft 365

Word, PowerPoint, Excel, Teams and Google Docs. Chrome, Edge, Brave, Arc and Firefox.

Three parts that work together:

| Part | Where it works | What it does |
|---|---|---|
| **Themed files** | Word, PowerPoint and Excel: desktop apps, web, phones | The theme is saved inside the .docx / .pptx / .xlsx, so anyone who opens it sees it with nothing installed |
| **Browser extension** (this folder) | Word, PowerPoint, Excel, Teams and Google Docs on the web, in Chrome, Edge, Brave, Arc and Firefox | Live animated scene, sparks while you type, the tool that saves themes into files, and Teams meeting backgrounds |
| **Office add-in** (`wild-themes-addin`) | Word and PowerPoint (desktop and web); Excel shows how to theme a workbook | One click inside the app saves a theme into the open file |

## What each app gets

| App | On the web (extension) | In the file (everyone, no install) |
|---|---|---|
| Word | Scene around the page | Backdrop + reading panel on every page, locked; theme colours and fonts |
| PowerPoint | Scene around the slide | Backdrop + see-through reading panel on every slide (on the slide master, locked); theme colours and fonts |
| Excel | Accent colours and sparks only: the grid fills the window, so there's no free space for a scene | Pale sheet background on every sheet (desktop Excel); theme colours and fonts |
| Google Docs | Scene around the page (the page stays white and readable). Typing sparks included | Google Docs keeps no theme of its own: the extension remembers your choice for each document on your own computers |
| Google Slides | Scene around the slide | As Google Docs |
| Google Sheets | Accent colours and sparks; the scene shows only where the grid leaves room | As Google Docs |
| Teams | Scene behind the chat, with a soft veil so text stays readable | Teams has no files to theme: use the 1920×1080 meeting backgrounds instead (Video effects › Add new) |

## Install the extension

Browsers install add-ons in one click only from their own stores, so until Wild Themes is in them you load the files yourself. Each browser does it a little differently; the website's **Add to browser** button shows these steps for the browser you're using.

**Chrome, Edge, Brave, Opera, Vivaldi, Arc** (all load the same package, `wild-themes-extension.zip`)

1. Unzip it (Windows: right-click › Extract All…). Keep the folder: the browser reads the add-on from it. `manifest.json` must be directly inside the folder you pick.
2. Open the browser's extensions page and turn on **Developer mode**:

   | Browser | Address | Developer mode switch |
   |---|---|---|
   | Chrome | `chrome://extensions` | top right |
   | Edge | `edge://extensions` | left menu |
   | Brave | `brave://extensions` | top right |
   | Opera | `opera://extensions` | top right |
   | Vivaldi | `vivaldi://extensions` | top right |
   | Arc | `arc://extensions` | top right |

3. Click **Load unpacked** and choose the folder. Pin it from the puzzle-piece menu if you like. Expect a "developer mode extensions" notice when the browser starts.

**Firefox** (version 140 or newer; `wild-themes-firefox.zip`)

Firefox keeps an add-on after a restart only when Mozilla has signed it. Opening or dragging an unsigned file into Firefox fails with "This add-on could not be installed because it has not been verified": that is this rule, not a broken download. Ways around it:

- *Signed copy:* when `downloads/wild-themes-firefox.xpi` is published the website offers **Install in Firefox**. The repository's **Sign Firefox add-on** workflow makes that file with a free Mozilla account (`.github/workflows/firefox-sign.yml` says how).
- *Until Firefox restarts:* open `about:debugging#/runtime/this-firefox`, click **Load Temporary Add-on…** and choose `wild-themes-firefox.zip` (don't unzip it), or the `manifest.json` inside an unzipped copy. Firefox treats access to websites as a permission you grant: allow it if asked; if the scene doesn't show, open `about:addons`, click Wild Themes › Permissions and switch on access to Word, Google Docs and Teams.
- *Keep it in regular Firefox, without GitHub:* sign it on Mozilla's website. Free account at addons.mozilla.org/developers › Submit a New Add-on › **On your own** (not listed) › upload `wild-themes-firefox.zip` (new version number each time) › download the signed .xpi when the automatic check finishes › open it in Firefox.
- *Developer Edition, Nightly or ESR:* set `xpinstall.signatures.required` to `false` in `about:config` (regular Firefox ignores it), then **about:addons › ⚙ › Install Add-on From File…** and choose `wild-themes-firefox-unsigned.xpi` (the same file as the zip).

**Safari and phones:** not supported. Safari only runs add-ons built into a Mac app with Xcode and signed through Apple's developer program (Apple's `safari-web-extension-converter` can start one from this folder; not tried). Phone browsers can't load add-ons from a file; Firefox for Android only takes add-ons from Mozilla's store. Themed files, the File studio and the Office add-in work everywhere.

**Building the packages:** `node scripts/build-extension.js` writes both zips and the `.xpi` copy to `downloads/` (the site rebuilds them on every deploy); `--dir <folder>` writes unzipped folders instead, to load by hand. Loading `extension/` itself works in Chrome-family browsers; for Firefox use the built Firefox package, since its manifest differs (an event page instead of a service worker, an add-on id, a declaration that nothing is collected).

## Use it

- **Pick a theme:** open a file in Word, PowerPoint or Excel on the web, Google Docs, Sheets or Slides, or Teams, click the icon and pick one. With **This document** selected, the extension remembers the theme for that file on your own computers, automatically.
- **Save a theme into a file (for everyone):** click **Save theme into a file**, choose a .docx, .pptx or .xlsx (on the web: File › Save As › Download a Copy), pick a theme, then **Create themed copy** and upload it. The theme is written into the file: backdrop, colours, fonts and hidden file properties. When you open a themed file with the extension installed, it picks the theme up from the file on its own, with no tag line.
- **Custom themes:** click **Custom themes** to make one from your own colours or a picture. You can export a theme as a `.wildtheme` file to share and import other people's.

## Built to run on anything

- No animation loop and no per-frame code. Scenes are colour gradients painted once; the browser only slides or fades whole layers, which needs almost no processor or graphics power.
- **Animation › Auto** adapts to your device. It starts with low-powered devices (2 processor cores or 2 GB memory or less) and systems that ask for reduced motion standing still. On every other device it measures the scene's real frame rate for two seconds, a few seconds after the page loads, and looks a second time before deciding. If frames are late or uneven it first holds the biggest, rotating and 3D layers still and keeps the rest moving, then, if that isn't enough, stops the scene. The verdict is remembered for two weeks on that device; clicking **Auto** again forgets it. **On** always animates, **Still** never does.
- The extension looks at the page's layout when it changes, not every frame: the first change at once, then at most about six times a second, so a page that never stops changing (Google Docs, Teams) isn't slowed down.
- The apps' own colours and text are never changed. The extension only makes the empty background behind the page, slide or chat see-through, and only once it is sure where that is. If it can't tell, it changes nothing.

## Good to know

- It's unofficial and reads each web app's page layout, which Microsoft and Google can change. If the scene stops showing, the safe behaviour is that the app just looks normal.
- Browser extensions can't run inside the Word, PowerPoint, Excel or Teams desktop apps, and Microsoft doesn't let anything else restyle them. For an app-like window with the theme, install the web version as an app: in Edge, ⋮ › Apps › Install this site as an app; in Chrome, ⋮ › Cast, save and share › Install page as app. Extensions run inside those windows. On the desktop apps, themes come from the files themselves (and the add-in).
- Typing sparks are a small fixed set of dots reused over and over; turn them off in the popup if you like.
- **Google Docs not showing the scene?** Open the Wild Themes popup on the Docs tab: it says whether the scene is on, and **Copy page details** copies what the extension sees (element names, sizes and colours, never your text or the document id) so the layout can be matched to what Google serves now.
- **Google Docs, Sheets and Slides:** the extension can't read a theme out of a Google file (Google doesn't keep the theme parts of a Word file), so it uses the theme you pick, remembered per document. A themed .docx opened in Google Docs keeps what Google's importer keeps (fonts, heading styles and colours, roughly); the scene picture and the embedded theme definition aren't kept, so keep your original .docx. The Google pages are matched by their layout, which hasn't been checked against the live editors from here; if the scene doesn't show, the page just looks normal.
- Privacy: it runs only on Word, PowerPoint, Excel, Teams and Google Docs, Sheets and Slides on the web, sends nothing anywhere, and stores your choices in your browser's extension storage. Documents are identified by a short code made from the file's address, never the address itself.
