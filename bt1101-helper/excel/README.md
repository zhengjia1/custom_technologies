# BT1101 Prose Lab — Excel edition

`../BT1101 Prose Lab.xlsm` is a pure-VBA port of `../index.html` for desktop Excel on
Windows and Mac. Open it with macros enabled; the app window opens automatically (the
start sheet's button reopens it).

How it works: the website's body markup, stylesheet rules and fonts are stored on a very
hidden sheet. At startup the markup is parsed into a VBA DOM, the operation code (ported
nearly line by line from the website's script) runs against it, and a renderer draws the
DOM into an MSForms window using the website's CSS. Details: `PORTING.md` (rules and core
libraries), `FRAMEWORK.md` (DOM), `OPS.md` (module ownership and porting conventions).

## Build

Needs Windows, desktop Excel with "Trust access to the VBA project object model", and the
development dependencies installed in `../pdf` (`npm install` there).

```
node build.js              # version from package.json (excelVersion)
node build.js 021126_2300  # or ddmmyy_hhmm
```

## Tests (from this folder)

| Command | Checks |
|---|---|
| `node test/compare.js test/steps/*/*.txt --only <modules> --no-form` | Every operation against the real website in Node (66 step files; the module list is in `OPS.md`) |
| `node test/e2e/run.js` | The real window's controls, clipboard, modals, keyboard, resizing (~4,300 checks) |
| `node test/layout-compare.js --run` | Renderer layout against the reference engine |
| `node test/regex-cases.js`, `node test/js-cases.js --run`, `node test/dom-test.js`, `node test/variants/equiv.js` | Core libraries and the answer-option engine |
| `node test/perf.js` | Per-action timings |

`node tools/vba.js check` compiles everything and reports the exact module, line and
message of any compile error.

## Known differences

- The website's own behaviour is reproduced exactly, including its known bugs (see the
  differential tests). Fix those in `index.html` first, then port.
- MSForms limits: no rounded corners, no textarea line-height, system-coloured scrollbars.
- Mac: untested on real hardware. Fonts fall back to Arial (no private font install), the
  update check opens the download page instead of checking, and there is no window resize.
  First Mac checks, in order: Debug > Compile; the window appears with readable text;
  page scrolling (trackpad, PageUp/PageDown); dark theme; paste into an input and copy an
  answer into RStudio; Escape/arrows/Tab; buttons reachable with Tab.
