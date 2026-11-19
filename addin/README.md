# Wild Themes add-in for Word, PowerPoint and Excel (desktop and web)

Adds a **Themes** button to the Home tab in Word, PowerPoint and Excel. In Word and PowerPoint, pick a theme and it's saved into the open file itself: backdrop, colours, fonts, locked theme pictures and hidden file properties. Anyone who opens the document sees the theme, with nothing installed. Removing it is done with the same add-in (**Remove theme from this document**).

Works in Microsoft 365 on Windows and Mac, and on the web (Word needs Word API 1.5; PowerPoint needs PowerPoint API 1.2). Excel doesn't let add-ins swap a whole workbook, so in Excel the button points you to the browser extension's “Save theme into a file”, which handles .xlsx.

## 1. Put the add-in online (one time, free)

Word only loads add-ins from an HTTPS web address. GitHub Pages is free:

1. Create a GitHub repository called `wild-themes-addin` and upload everything in this folder.
2. In the repository: **Settings › Pages › Deploy from a branch › main / (root)**. After a minute it's live at `https://YOUR-NAME.github.io/wild-themes-addin/`.
3. Open `manifest.xml`, replace every `https://YOUR-SITE` with `https://YOUR-NAME.github.io/wild-themes-addin`, and upload the changed file.

## 2. Add it to Word

- **On the web (Word, PowerPoint, Excel):** Home › Add-ins › More Add-ins › My Add-ins › **Upload My Add-in** › choose `manifest.xml`. (It's remembered in that browser.)
- **Mac:** copy `manifest.xml` into `~/Library/Containers/com.microsoft.Word/Data/Documents/wef` (and the same under `com.microsoft.Powerpoint` / `com.microsoft.Excel`; create `wef` if needed), restart the app, then Home › Add-ins.
- **Windows:** put `manifest.xml` in a shared folder, add that folder under File › Options › Trust Center › Trust Center Settings › Trusted Add-in Catalogs (tick "Show in Menu"), restart the app, then Insert › My Add-ins › Shared Folder.

## Good to know

- In PowerPoint, applying a theme re-adds your slides with the themed master and removes the old copies; speaker notes come along. In Word, it rebuilds the document with the same text, pictures and formatting. Comments and tracked changes may not carry over, so save a copy first if you need them; Word's version history can undo it.
- Word swaps documents in place only up to about 4 MB. For bigger files, use the browser extension's "Save theme into a Word file".
- The backdrop animates in the Word desktop app. Word for the web shows it as a still picture.
- After changing anything in `../wild-themes-for-word`, run `./build.sh` to copy the shared files across.
