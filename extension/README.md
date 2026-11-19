# Wild Themes for Microsoft 365

Word, PowerPoint, Excel and Teams.

Three parts that work together:

| Part | Where it works | What it does |
|---|---|---|
| **Themed files** | Word, PowerPoint and Excel: desktop apps, web, phones | The theme is saved inside the .docx / .pptx / .xlsx, so anyone who opens it sees it with nothing installed |
| **Browser extension** (this folder) | Word, PowerPoint, Excel and Teams on the web, in Chrome, Edge, Brave, Arc | Live animated scene, sparks while you type, the tool that saves themes into files, and Teams meeting backgrounds |
| **Office add-in** (`wild-themes-addin`) | Word and PowerPoint (desktop and web); Excel shows how to theme a workbook | One click inside the app saves a theme into the open file |

## What each app gets

| App | On the web (extension) | In the file (everyone, no install) |
|---|---|---|
| Word | Scene around the page | Backdrop + reading panel on every page, locked; theme colours and fonts |
| PowerPoint | Scene around the slide | Backdrop + see-through reading panel on every slide (on the slide master, locked); theme colours and fonts |
| Excel | Accent colours and sparks only: the grid fills the window, so there's no free space for a scene | Pale sheet background on every sheet (desktop Excel); theme colours and fonts |
| Teams | Scene behind the chat, with a soft veil so text stays readable | Teams has no files to theme: use the 1920×1080 meeting backgrounds instead (Video effects › Add new) |

## Install the extension

1. Unzip `wild-themes-for-word.zip`. You get one folder, `wild-themes-for-word`, with `manifest.json` directly inside it.
2. Open `chrome://extensions` (Chrome) or `edge://extensions` (Edge) and turn on **Developer mode**.
3. Click **Load unpacked** and choose that folder (the one with `manifest.json` in it, not `icons` or `lib`).
4. Pin it from the puzzle-piece menu.

## Use it

- **Pick a theme:** open a file in Word, PowerPoint or Excel on the web (or Teams), click the icon and pick one. With **This document** selected, the extension remembers the theme for that file on your own computers, automatically.
- **Save a theme into a file (for everyone):** click **Save theme into a file**, choose a .docx, .pptx or .xlsx (on the web: File › Save As › Download a Copy), pick a theme, then **Create themed copy** and upload it. The theme is written into the file: backdrop, colours, fonts and hidden file properties. When you open a themed file with the extension installed, it picks the theme up from the file on its own, with no tag line.
- **Custom themes:** click **Custom themes** to make one from your own colours or a picture. You can export a theme as a `.wildtheme` file to share and import other people's.

## Built to run on anything

- No animation loop and no per-frame code. Scenes are colour gradients painted once; the browser only slides or fades whole layers, which needs almost no processor or graphics power.
- **Animation › Auto** turns animation off on low-powered devices (2 processor cores or 2 GB memory or less) and when your system asks for reduced motion. **Still** keeps the scene but stops all movement.
- The apps' own colours and text are never changed. The extension only makes the empty background behind the page, slide or chat see-through, and only once it is sure where that is. If it can't tell, it changes nothing.

## Good to know

- It's unofficial and reads each web app's page layout, which Microsoft can change. If the scene stops showing, the safe behaviour is that the app just looks normal.
- Browser extensions can't run inside the Word, PowerPoint, Excel or Teams desktop apps, and Microsoft doesn't let anything else restyle them. For an app-like window with the theme, install the web version as an app: in Edge, ⋮ › Apps › Install this site as an app; in Chrome, ⋮ › Cast, save and share › Install page as app. Extensions run inside those windows. On the desktop apps, themes come from the files themselves (and the add-in).
- Typing sparks are a small fixed set of dots reused over and over; turn them off in the popup if you like.
- Privacy: it runs only on Word, PowerPoint, Excel and Teams on the web, sends nothing anywhere, and stores your choices in your browser's extension storage. Documents are identified by a short code made from the file's address, never the address itself.
