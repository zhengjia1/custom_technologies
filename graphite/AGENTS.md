# AGENTS.md — Graphite

Handoff notes for anyone (human or agent) picking up this project.

## What this is

Graphite turns a spreadsheet into an interactive chart, entirely in the browser. It replaced an
old Flask backend: there is no server for rendering, and the user's data never leaves their device
unless they choose **Save to Cloud**.

Four pages, each a single self-contained `index.html` (no build step, no framework):

| Page | Folder | Chart | Live at |
|---|---|---|---|
| Graphite | `index.html` | Line chart, searchable series | `zhengjia.dev/graphite` |
| Graphite Combo | `combo/index.html` | Bars (left axis) + one line (right axis) | under `zhengjia.dev/graphite/…` |
| Graphite Histogram | `histogram/index.html` | Bars per bin, value labels | under `zhengjia.dev/graphite/…` |
| Graphite Results | `results/index.html` | Students per score group (bars) + calculated percentile (line, 0–100%) | under `zhengjia.dev/graphite/…` |

Look and feel follows zhengjia.dev: random background photo from `cdn.zhengjia.dev/static/N.jpg`,
Adobe Clean font, dark inputs with a gradient border that glows only on focus.

**Dark mode is the default.** A round ☀/☾ button (top-right of the start screen, end of the chart
toolbar; `[data-theme-toggle]`) switches to light. The choice is saved in `localStorage`
(`graphite-theme`, shared by all four pages) and applied by a small `<head>` script before first
paint (`<html data-theme="dark|light">`); light mode is a block of `html[data-theme="light"]` CSS
overrides. Charts get theme colours from `chartTheme()` (Flourish dotted-key settings: layout
background/font, axis lines/gridlines/ticks/titles, popups, Results' quartile lines), merged into
the settings in `buildGraphHtml`, plus the chart page's body background. Switching theme while a
chart is open rebuilds it; Save > SVG/HTML use the current theme (the SVG's background is the chart
page's body colour).

## Files

```
index.html                 Graphite (line chart)
combo/index.html           Graphite Combo
combo/static/              Local mirror of the combo/histogram chart files (hosted on the CDN, see below)
combo/samples/*.xlsx       Combo sample datasets
combo/test.csv             Example data for the combo chart (also the built-in ?id=sample)
histogram/index.html       Graphite Histogram
histogram/samples/*.xlsx   Histogram sample datasets
results/index.html         Graphite Results
results/data.xlsx          Example exam-results data (also the built-in ?id=sample)
samples/*.xlsx             Graphite (line chart) sample datasets
tools/update_samples.py    Downloads official data and rebuilds all sample files + the pages' sample lists
update_samples.bat         Double-click wrapper for the above (installs openpyxl if missing)
test.xlsx                  Ad-hoc test spreadsheet
```

**The four pages are near-copies with no shared code.** Combo was generated from the Graphite page,
Histogram and Results from Combo, and all have since diverged. A fix to shared behaviour (parsing,
validation, cloud save, sandboxing…) must be applied to **all four files** by hand.

## How a page works

1. **Input**: file upload (drag & drop / browse) or pasted cells, parsed with SheetJS
   (`xlsx@0.18.5` from cdnjs). Plus title fields: Graphite = graph title + search-bar title;
   Combo = graph, x, y (bars), optional line-axis title (defaults to the line column's name);
   Histogram = graph, x, y; Results = graph, x, y (the line axis is always "Percentile").
2. **Validation** (`readFile` → `readXlsx` / `readCsv`): only `.xlsx` / `.csv`; contents must match
   the extension (xlsx = zip containing `xl/workbook.xml`, csv = text, UTF-8 or Windows-1252);
   10 MB, 10,000 rows, 100 series; table checked on upload. Combo also needs ≥ 2 number columns;
   Results needs exactly 2 columns (score or score range, number of students), counts ≥ 0, total > 0.
3. **Parsing** (`sheetRows` → `toDataset`): works on SheetJS cell objects. Labels keep Excel's
   displayed text; values use the raw number (`$1,234.50` → 1234.5, `25%` cell → 25, `(500)` → −500,
   text numbers incl. `1.234,56`). Skips title rows above the table, empty/formatted-only columns,
   blank rows, `Total`/`Sum` rows and note rows with no numbers. Excel's default date format is
   shown as e.g. "1 May 2024" (SheetJS would print US-style 5/1/24). Pasted text and CSV are read
   with `raw: true` so labels are kept as typed.
4. **Rendering** (`buildGraphHtml`): builds a complete standalone HTML page for the chart and shows
   it in `<iframe id="frame" sandbox="allow-scripts allow-popups">` via `srcdoc`.
   The toolbar's **Save ▾** menu offers **Cloud** (share link), **SVG** (image) and **HTML** (that
   standalone page, interactive). Shared graphs hide Cloud.
5. **Share links**: `?id=<uuid>` loads a cloud-saved graph; `?id=sample` shows a built-in sample.
   Shared graphs open with **Save** (SVG / HTML) / **Create your own copy** instead of Save / New.
6. **SVG export** (`svgExporter`, injected into the chart page via `toString()`): the parent posts
   `{graphiteSvg}` to the sandboxed frame (re-sending every 500 ms until `{graphiteSvgAck}`, since
   the chart page may still be loading). The chart page waits until the drawing stops changing
   (shape fingerprint stable for 2 checks, max 10 s — charts animate in), lets the library convert
   its HTML title/subtitle/legend to SVG (`template.screenshot({format:"svg"})` on the newer
   library; on Graphite's older one, which lacks it, `htmlParts()` does it), then copies every
   on-screen `<svg>` to its position with computed styles inlined (keeping each element's own
   inline style — clip-path rects are sized there) and replies `{graphiteSvgResult}`. Skips
   `.fl-ignore-svg-download` and popups; copies are sized to at least their container because
   the line chart's `#graphic` is a 300×150 box that overflows.

### Chart files (on the CDN)

| Page | Chart library |
|---|---|
| Graphite | `cdn.zhengjia.dev/graph/static/supp_1.js`, `supp_2.css`, `supp_3a.css`, `embedded.js` (old Flourish line template) |
| Combo, Histogram, Results | `cdn.zhengjia.dev/graphite/combo/template.js`, `bundle.css`, `choices.min.css`, `icons/` (Flourish "Line, bar and pie charts" template 118910 v4) + Graphite's `graph/static/embedded.js` |

Histogram and Results deliberately **reuse** the combo chart files (they point at `/graphite/combo/`);
the Flourish exports were identical apart from a hidden credit string.

### Chart settings worth knowing

- Combo: the template draws the **first** value column as the line (`secondary_series: 1`). Users put
  the line in the **last** column, so `buildGraphHtml` moves it to the front.
- Combo/Histogram set `layout.body_font` to Source Sans Pro so they match Graphite's charts
  (the template's default is Canva Sans).
- Histogram: `legend_mode` is `off` for one series (as in the export), `auto` for several.
- Results: percentile = **% of students at or below** each group (running total ÷ total, 1 dp).
  Groups are sorted by the number their label starts with (`85`, `41-45`); if any label doesn't
  start with a number (e.g. letter grades) the table's order is used, lowest first. Only the
  counts are saved; percentiles are recalculated whenever the chart is built.
- Results also draws dotted vertical lines at **1Q, Median and 3Q** (`axes_highlights.x_lines`,
  styling taken from the user's Flourish export): each is the **first group whose percentile
  reaches 25% / 50% / 75%**. Lines are written as `"1Q: 27.5 :: 27.5"` (label :: position); the
  position is the group's label, which works for single marks (number axis) and ranges like
  `26-30` (category axis) alike. The CDN template (118910) supports these settings, so nothing
  new needs hosting.
- Results merges groups **on narrow screens** (`fitGroups`, injected into the chart page via
  `toString()` so it runs where the width is known): at most ~1 bar per 16 px, neighbouring groups
  merged (counts summed, percentiles recalculated, labels like `27.5–29.5`). Quartile values come
  from the full data; their lines move to the merged bar containing them, and quartiles on the same
  bar share one line. When merged, `layout.subtitle` says "Some groups have been merged to fit this
  narrow screen. Best viewed on desktop." Re-fits on resize (`template.update()`). Desktop is unchanged.
- **Set nested chart settings as dotted keys, not objects** (`"number_format_secondary.suffix": "%"`).
  The template builds some formatters once at load around the original object; replacing the
  object (`"number_format_secondary": { suffix: "%" }`) is silently ignored by the axis.

## Gotchas (each of these cost real debugging time)

- **Short charts on phones.** Flourish's `embedded.js` decides `fixed_height` by reading the parent
  page's URL. In a sandboxed `srcdoc` frame that throws (console shows "Script error"), and the
  chart falls back to 400 px on phones / 650 px on desktop. Every page sets
  `Flourish.fixed_height = true; Flourish.checkHeight = function () {};` **after** the scripts
  load and **before** `template.draw()`. Setting it earlier is useless — `embedded.js` overwrites it.
- **Never write a literal `</body>`, `<body`, `<head>` or `</head>` inside the page's `<script>`.**
  VS Code Live Server (`127.0.0.1:5500`, used for local testing) injects its reload script before the
  *first* `</body>` in the file; if that's inside a JS string the whole page script breaks (blank
  page, no greeting). Build tags as `"</" + "body>"`. Comments count too.
- **Flourish renders titles as HTML.** Shared graphs come from other people, hence the sandbox
  (opaque origin) and `fromSaved()` stripping `<` `>` and keeping only the expected JSON shape.
- **Testing the sandbox locally:** Chrome blocks a sandboxed frame from loading scripts from
  `localhost`, so a local copy of `template.js` won't load in the real frame (CDN files are fine).
  To test, inline the script into a test `srcdoc`, or test against the CDN copy. When inlining via
  `String.replace`, use a function replacer — `$` sequences in minified code corrupt it otherwise.
- **Screenshots of the chart** often catch the 1.5 s entrance animation (empty or half-drawn bars).
  Wait or re-take before concluding something is broken.
- **The Claude browser pane runs no animation frames while it isn't being drawn** (0 rAF/s), so
  charts freeze at their start state (opacity 0) and SVG exports come out blank *in testing only*.
  Keep the pane drawing (take screenshots while exporting) before trusting an export test.
- **Flourish label types:** with `"auto"`, labels like "May 1" were read as years (2001…). Graphite
  only lets plain 4-digit years be `auto`; everything else is `string`.
- **Real time axis (Graphite + Combo only).** A label is a date if it's a real Excel date cell (any
  display format) or text as `DD MMM YYYY` (`3 Nov 2026`, full month names OK) or `DD/MM/YYYY`
  (`03/11/2026`, day first). Only if **every** label is a date: rows are sorted by date, labels are
  stored as `03 Nov 2026`, `labelType = "date"`, and the chart gets ISO dates (`DATE_META`:
  `datetime$%Y-%m-%d` in, `datetime$%d %b %Y` out). Graphite's old library also needs
  `x_axis_date_format`, or ticks show the input format. Invalid dates (31/02), month-first text and
  dates without a year stay text. The newer library only accepts `datetime$` ids from its built-in
  list (`%Y-%m-%d` and `%d %b %Y` are on it) and silently skips points whose label won't parse.
  It also does **not** parse label strings itself: on a time axis it sorts with `label.getTime()`,
  so the Combo chart page turns labels into `Date` objects before drawing (otherwise: blank chart,
  console "t.label.getTime is not a function"). It formats with `utcFormat`, so they must be **UTC
  midnight** (`Date.UTC`) — local midnight shows every tick a day early east of Greenwich.

## Save to Cloud

- Flow: Cloudflare Turnstile → `POST https://graphite-upload.zhengjia.workers.dev/graphite/sign`
  `{token, size}` → Worker returns a presigned S3 PUT URL (2-minute expiry, exact `content-length`
  and `content-type` signed) → browser PUTs the JSON straight to the Backblaze B2 bucket behind the CDN, at
  `graphite/shared/<uuid>.json` → share link `?id=<uuid>`.
- Loading: `GET …/graphite/load?id=<uuid>` (Worker proxies `cdn.zhengjia.dev`, adds CORS).
- Saved JSON: `{ v: 1, kind, title, …, columns, data, labelType }`. `kind` is absent for Graphite
  (line), `"combo"`, `"histogram"` or `"results"` for the others; each page refuses the other kinds
  with a message.
- Worker config: vars `B2_BUCKET`, `B2_ENDPOINT`, `B2_REGION`, `PUBLIC_BASE`, `ALLOWED_ORIGINS`;
  secrets `B2_KEY_ID`, `B2_APP_KEY`, `TURNSTILE_SECRET`. The public Turnstile site key is in each page
  (`TURNSTILE_SITEKEY`) and must list the domain the pages are served from.
  B2 CORS rule `graphite-upload` allows `s3_put` from `https://zhengjia.dev`.
- **The Worker source (`graphite-worker.js`) is not in this folder.** It was written in an earlier
  working folder that has since moved; the deployed copy lives in the Cloudflare dashboard.

## Sample datasets

All samples come from official Singapore open data, downloaded by `tools/update_samples.py`
(run `update_samples.bat`). It never uploads anything.

- `GROUPS` in the script = Graphite → `samples/`, Combo → `combo/samples/`,
  Histogram → `histogram/samples/`, each with its page and CDN folder. Every run processes all
  groups (18 files today) and ends with "Checked all N sample files across M pages".
- Before downloading anything it checks coverage and stops if something would be missed: any
  `index.html` in the project with a `// SAMPLES:BEGIN` list that isn't in `GROUPS`, or any `.xlsx`
  in a sample folder that no dataset builds. **Adding a new chart type = add its datasets and a
  `GROUPS` entry**, otherwise the script refuses to run.
- Each `.xlsx`: A1 title, A2 summary (shown in the picker), A3 source citation, blank row,
  header on row 5, data from row 6. A data fingerprint is stored in the workbook properties
  (`keywords`); unchanged data is not rewritten.
- Each page's sample list sits between `// SAMPLES:BEGIN` and `// SAMPLES:END` and is rewritten by
  the script: file, `v` (the fingerprint, used as `?v=` so Cloudflare's 7-day cache can't serve a
  stale file), title, summary, and for Combo/Histogram the axis titles.
- Sources: SingStat Table Builder API and data.gov.sg datastore API.

Data-source quirks the script already handles:
- SingStat rejects Python's default client — requests send a browser `User-Agent`.
- SingStat's `limit`/`offset` count **data points**, not rows — page until empty.
- Some SingStat tables repeat row names across groups (age table: total/male/female) — use
  `singstat_rows()` with `seriesNo`, not `singstat()`.
- openpyxl read-only workbooks must be closed, or Windows can't replace the file.
- HDB resale applications (`d_02aa…`) sometimes repeat a quarter with revised figures and publish
  new quarters flat type by flat type: keep the latest record, skip incomplete quarters.
- Census tables are rounded per cell, so bins can differ from the stated total by 1.
- The HDB resale-price histogram always uses the last 3 *complete* months.

## Deploying

Everything is uploaded manually (do **not** upload to Backblaze/the CDN from scripts or agents):

- Pages → wherever zhengjia.dev serves them. Results needs no extra CDN files and has no samples.
- Sample files → `cdn.zhengjia.dev/graphite/samples/`, `/graphite/combo/samples/`,
  `/graphite/histogram/samples/`. Upload the files **before** the page whose `v=` tags point at them.
- Chart files → `cdn.zhengjia.dev/graphite/combo/` (from `combo/static/`; `embedded.js` there is unused).

Last checked (19 Oct 2026): chart files `template.js`, `bundle.css`, `choices.min.css` and the
combo + histogram sample files are on the CDN. Still missing (404): `graphite/combo/icons/cross.svg`
and `cross-inverse.svg` — only the "×" icons in the chart's search box. Whether the latest page
HTML has been uploaded isn't verifiable from here.

## Testing

No automated tests. Serve the folder and use a browser:

```
python -m http.server 8765
```

(Not Live Server for sandbox checks — see the localhost note above.) Useful checks: upload
`combo/test.csv`; open `?id=sample`; run every sample in the picker; phone width 412 px; confirm the
chart reaches the bottom of the frame.

## Workflow

- Uploading/deploying is done by hand; don't push to Backblaze or the CDN automatically.
- Local testing uses VS Code Live Server; mobile checks use Chrome device emulation on the live site.
- No secrets belong in this repo: Worker keys live only in Cloudflare's secret settings.
