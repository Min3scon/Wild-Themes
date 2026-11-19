# Wild Themes add-in for Word, PowerPoint and Excel

Adds **Home › Themes** to Word, PowerPoint and Excel on Windows, Mac and the web. The pane:

- recognises the theme saved in the open file (anyone's file: the theme's description travels inside it) and checks it's intact;
- applies, changes, tunes (look and effects), restores or removes it, then reads the file back and shows exactly what's in it;
- lists every theme on the website, including approved community themes, without updating the add-in.

It runs the same code as the website's File studio and the browser extension (`../extension/lib/wtcore.js`), served from the website, so it updates itself.

Install: see the main README ("Office add-in"). Requirements: Word API 1.5 (in-place) or 1.3 (copy), PowerPoint API 1.2 (in-place) or 1.1 (copy), Excel API 1.8 (copy). Custom XML sync needs Word API 1.4 / PowerPoint API 1.7. Opening by itself needs the shared runtime (SharedRuntime 1.1).

Word swaps documents in place up to about 4 MB; use "Open a themed copy" (or the File studio) for bigger files. Comments and tracked changes may not survive an in-place swap; the copy keeps everything.
