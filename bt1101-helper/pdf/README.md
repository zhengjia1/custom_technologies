# BT1101 Prose Lab — PDF edition

`npm run build` turns `../index.html` into `../BT1101 Prose Lab.pdf`, a fillable PDF that runs the website's own JavaScript inside **Adobe Acrobat Reader** (desktop, JavaScript enabled). Browser PDF viewers, macOS Preview and phone apps can open the file but cannot run it.

## Using the PDF

- Each operation has its own page, laid out like the website. Click an operation in the sidebar, or type into the sidebar search and press Tab.
- Type or paste into a box, then press Tab or click elsewhere; the page updates when you leave the box.
- Dropdowns open a pop-up menu with the current choice ticked; plot types are grouped under their categories.
- Acrobat cannot write to the clipboard or select a whole field from a script. A Copy button puts the exact copied text into the yellow box in its panel; click the box, press Ctrl+A, then Ctrl+C (Cmd+A, Cmd+C on a Mac). Each box and the status line say so.
- The website's warnings (chart axes, regression units, placeholder names, K-means scaling, known population SD) appear as Acrobat dialogs.
- The status line at the top right shows the website's messages ("Sample loaded", "Prose copied") and any errors.

Differences a PDF cannot avoid: pages do not reflow, so fields that appear only for some choices keep their space; the Paste buttons, dark mode and sidebar resizing are left out; completed cards stay open and show **Done!** instead of collapsing.

## How it works

1. `src/extract.js` takes the page markup and scripts from `index.html` unchanged.
2. `src/minidom.js` and `src/env.js` provide the browser objects the website needs (DOM, events, selectors, timers, clipboard) so its code runs in Acrobat, which has JavaScript but no web page.
3. `src/explore.js` tries every option, toggle and sample of each operation to learn which elements can appear, and every combination of the controls that change the layout (plot type, interval type, …). A PDF cannot reflow, so each distinct layout becomes its own page: the main pages follow the website's operation list, the other layouts come after them, and the PDF switches page when your choices change the layout. The search is cached in `.build/layouts-cache.json` until `index.html` changes.
4. `src/capture-page.js` opens `index.html` in headless Edge or Chrome and measures every box, word and control.
5. `src/pdfwriter.js` draws those pages with the website's Source Sans Pro and colours, and places form fields over the controls.
6. `src/bridge.js` connects the fields to the website: edits and clicks become DOM events, and the page's state is copied back into the fields.
7. `build.js` compiles everything to ES5 with core-js polyfills (Acrobat's engine predates modern JavaScript) and embeds it as a document-level script.

## Building and testing

```bash
npm install
npm run build            # version from package.json "pdfVersion" (ddmmyy_hhmm)
npm run build:dark       # dark theme, written to ../Proselab_darkmode.pdf
node build.js 021126_2300 --theme=dark
npm test
```

Building needs Microsoft Edge or Google Chrome installed. The first build after `index.html` changes searches every layout and takes several minutes; later builds reuse `.build/layouts-cache.json`. If the PDF is open in Acrobat (which locks it), the build writes `… (new).pdf` beside it instead. Tests compare minidom with jsdom step by step, and run the PDF's own script in an ES5 sandbox that imitates Acrobat's API. `node test/live-preview.mjs <dir> [operation…]` renders pages as Acrobat would show them after loading samples.

Rebuild the PDF after every change to `index.html`; it embeds a copy of the website.
