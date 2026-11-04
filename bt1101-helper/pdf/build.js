// Builds "BT1101 Prose Lab.pdf" from ../index.html.
//   node build.js              uses pdfVersion from package.json
//   node build.js 021126_2300  sets the version (ddmmyy_hhmm)
//
// 1. Run the website on minidom to learn which elements each operation can show.
// 2. Open the website in headless Edge and measure every box, word and control.
// 3. Draw those pages into the PDF and lay form fields over the controls.
// 4. Compile the website script, minidom and the bridge to ES5 for Acrobat.

const fs = require("fs");
const path = require("path");
const { pathToFileURL } = require("url");
const esbuild = require("esbuild");
const babel = require("@babel/core");
const acorn = require("acorn");
const { extractWebsite, writeAppModule } = require("./src/extract.js");
const { buildPdf, labelFields } = require("./src/pdfwriter.js");
const { loadWebsiteFonts, drawableCodePoints } = require("./src/fonts.js");

const ROOT = __dirname;
const BUILD = path.join(ROOT, ".build");
const INDEX = path.join(ROOT, "..", "index.html");
// node build.js [ddmmyy_hhmm] [--theme=dark]
const THEME = (process.argv.find(arg => arg.startsWith("--theme=")) || "--theme=light").slice(8) === "dark" ? "dark" : "light";
const OUTPUT = path.join(ROOT, "..", THEME === "dark" ? "BT1101 Prose Lab (Dark Mode).pdf" : "BT1101 Prose Lab.pdf");
const BROWSERS = ["C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe", "C:/Program Files/Google/Chrome/Application/chrome.exe", "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome", "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge", "/usr/bin/google-chrome", "/usr/bin/chromium"];
const VIEWPORT_WIDTH = 1280;
const NO_SCRIPT_MESSAGE = "JavaScript is not running. Open this file in Adobe Acrobat Reader on a computer, with JavaScript enabled.";

function parseVersion(value) {
  const match = String(value || "").trim().match(/^v?(\d{2})(\d{2})(\d{2})_(\d{2})(\d{2})h?$/i);
  if (!match) throw new Error(`Version must look like ddmmyy_hhmm, got "${value}"`);
  const [, day, month, year, hour, minute] = match.map(Number);
  const date = new Date(Date.UTC(2000 + year, month - 1, day, hour, minute));
  if (date.getUTCDate() !== day || date.getUTCMonth() !== month - 1 || date.getUTCHours() !== hour || date.getUTCMinutes() !== minute) throw new Error(`Version "${value}" is not a real date and time`);
  return `${match[1]}${match[2]}${match[3]}_${match[4]}${match[5]}`;
}

function bundleNodeLibraries(site) {
  writeAppModule(site, path.join(BUILD, "app.js"));
  esbuild.buildSync({ entryPoints: [path.join(ROOT, "src", "refs.js")], bundle: true, format: "cjs", platform: "node", outfile: path.join(BUILD, "refs-node.js"), logLevel: "error" });
  fs.writeFileSync(path.join(BUILD, "node-entry.js"), `
    import * as dom from "../src/minidom.js";
    import { createBrowser, installNumberLocale } from "../src/env.js";
    import { runApp } from "./app.js";
    export { dom, createBrowser, installNumberLocale, runApp };
  `);
  esbuild.buildSync({ entryPoints: [path.join(BUILD, "node-entry.js")], bundle: true, format: "cjs", platform: "node", outfile: path.join(BUILD, "node-bundle.js"), logLevel: "error" });
  [path.join(BUILD, "refs-node.js"), path.join(BUILD, "node-bundle.js")].forEach(file => { delete require.cache[file]; });
  return require(path.join(BUILD, "node-bundle.js"));
}

function makeBoot(lib, site) {
  return () => {
    const tasks = { micro: [], timers: [], seq: 0, now: 0 };
    const errors = [];
    const browser = lib.createBrowser(site.body, tasks, { reportError: error => errors.push(error) });
    const app = lib.runApp(browser.globals);
    browser.drain();
    if (errors.length) throw errors[0];
    return { document: browser.document, drain: browser.drain, app, Event: lib.dom.Event };
  };
}

function operationList(boot) {
  const { app } = boot();
  const list = [{ id: "universal", title: "Universal detector" }];
  app.groups.forEach(group => group.items.forEach(([id, title]) => list.push({ id, title })));
  return list;
}

// Pages to capture: each layout of each operation, plus, where an operation has more than
// one layout, a catch-all page with every element visible for any state not foreseen.
// The document reads as the website's operation list: each operation's default layout, in
// order. Other layouts follow at the end, reached only when the operation's state needs them,
// so scrolling never passes through an operation's alternative layouts.
function pagePlan(operations, explorations, layouts) {
  const main = [], extra = [];
  operations.forEach(operation => {
    const layout = layouts[operation.id];
    layout.variants.forEach((variant, index) => (index === 0 ? main : extra).push({ operation, settings: variant.settings, sig: variant.sig, unhide: false, first: index === 0 }));
    if (layout.variants.length > 1) extra.push({ operation, settings: [], sig: null, unhide: true, first: false });
  });
  return main.concat(extra);
}

// Icons the embedded font cannot draw are copied from the browser at 3x and stored as RGB.
async function decodeIcon(png) {
  const { loadImage, createCanvas } = require("@napi-rs/canvas");
  const image = await loadImage(png);
  const canvas = createCanvas(image.width, image.height);
  const context = canvas.getContext("2d");
  context.drawImage(image, 0, 0);
  const rgba = context.getImageData(0, 0, image.width, image.height).data;
  const rgb = Buffer.alloc(image.width * image.height * 3);
  for (let i = 0, j = 0; i < rgba.length; i += 4, j += 3) { rgb[j] = rgba[i]; rgb[j + 1] = rgba[i + 1]; rgb[j + 2] = rgba[i + 2]; }
  return { width: image.width, height: image.height, rgb, key: require("crypto").createHash("sha1").update(rgb).digest("hex").slice(0, 12) };
}

async function captureAll(plan, explorations, layouts, version, drawable) {
  const puppeteer = require("puppeteer-core");
  const executablePath = BROWSERS.find(file => fs.existsSync(file));
  if (!executablePath) throw new Error("Install Microsoft Edge or Google Chrome to capture the page layout.");
  const script = esbuild.buildSync({ entryPoints: [path.join(ROOT, "src", "capture-page.js")], bundle: true, format: "iife", globalName: "PDFCapture", write: false, target: "es2020", logLevel: "error" }).outputFiles[0].text;
  const browser = await puppeteer.launch({ executablePath, headless: true, args: ["--no-first-run", "--disable-extensions", "--force-color-profile=srgb"] });
  const captures = [];
  try {
    for (const entry of plan) {
      const operation = entry.operation;
      const page = await browser.newPage();
      await page.setViewport({ width: VIEWPORT_WIDTH, height: 900 });
      await page.emulateMediaFeatures([{ name: "prefers-color-scheme", value: THEME }]);
      // Keep the build offline: the website's update check must not reach the network.
      await page.setRequestInterception(true);
      page.on("request", request => (/^(file|data|blob):/.test(request.url()) ? request.continue() : request.abort()));
      await page.goto(pathToFileURL(INDEX).href, { waitUntil: "load" });
      await page.evaluate(theme => { localStorage.setItem("bt1101-theme", theme); document.documentElement.dataset.theme = theme; }, THEME);
      await page.addScriptTag({ content: script });
      await page.evaluate(() => document.fonts.ready);
      const config = Object.assign({ version: `v${version}h`, unhide: entry.unhide, drawable }, explorations[operation.id]);
      await page.evaluate(id => PDFCapture.open(id), operation.id);
      // Reach the layout, load samples for realistic content, then make sure the layout held.
      await page.evaluate(settings => PDFCapture.applySettings(settings), entry.settings);
      await page.evaluate(id => PDFCapture.loadSamples(id), operation.id);
      await page.evaluate(settings => PDFCapture.applySettings(settings), entry.settings);
      if (entry.sig) {
        const reached = await page.evaluate(tracked => PDFCapture.signature(tracked), layouts[operation.id].tracked);
        if (reached.join("|") !== entry.sig.join("|")) {
          const onlyEdge = reached.filter(ref => entry.sig.indexOf(ref) < 0), onlyNode = entry.sig.filter(ref => reached.indexOf(ref) < 0);
          console.log(`\n  note: ${operation.id} layout ${captures.length} differs in Edge: shown only in Edge ${onlyEdge.join(", ") || "-"}; only in Node ${onlyNode.join(", ") || "-"}`);
        }
      }
      await page.evaluate(value => PDFCapture.prepare(value), config);
      // A viewport as tall as the page, at 3x, so icons can be copied anywhere on it, crisply.
      const fullHeight = await page.evaluate(() => document.documentElement.scrollHeight);
      await page.setViewport({ width: VIEWPORT_WIDTH, height: fullHeight, deviceScaleFactor: 3 });
      await page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
      const capture = await page.evaluate(value => PDFCapture.measure(value), config);
      capture.palette = await page.evaluate(() => PDFCapture.palette());
      for (const icon of capture.icons) {
        icon.image = await decodeIcon(await page.screenshot({ type: "png", clip: { x: icon.x, y: icon.y, width: icon.w, height: icon.h } }));
      }
      // Hover over each button and sidebar entry and record the colours the browser applies.
      const hovers = [];
      for (let index = 0; index < capture.hovers; index++) {
        const handle = await page.$(`[data-pdf-hover="${index}"]`);
        if (!handle) { hovers.push(null); continue; }
        try { await handle.hover(); hovers.push(await page.evaluate(i => PDFCapture.hoverStyle(i), index)); } catch (error) { hovers.push(null); }
      }
      await page.mouse.move(0, 0);
      capture.fields.forEach(field => { if (field.hover !== undefined && field.hover !== null) field.hoverStyle = hovers[field.hover]; });
      capture.links.forEach(link => { link.hoverStyle = hovers[link.hover]; });
      captures.push(capture);
      await page.close();
      process.stdout.write(".");
    }
  } finally {
    await browser.close();
  }
  process.stdout.write("\n");
  return captures;
}

// Polyfills that replace a whole built-in constructor (RegExp, Number, Error, typed arrays).
// The website uses none of the features they add, and Acrobat may refuse the replacement.
const REPLACES_CONSTRUCTORS = ["es.regexp.constructor", "es.regexp.dot-all", "es.regexp.sticky", "es.regexp.flags", "es.number.constructor", "es.error.cause", /^es\.typed-array\./, /^es\.array-buffer\./, /^esnext\./];
// Polyfills that wrap a method every ES5 engine already has, to add behaviour the website
// never uses (Symbol.replace/split matchers, species, isConcatSpreadable, well-formed JSON).
// Each call would run through slow wrapper code; the engine's own method is used instead.
// Polyfills that add a missing method (includes, padStart, Map, Promise …) are kept.
const WRAPS_NATIVE_METHODS = [
  "es.array.concat", "es.array.filter", "es.array.map", "es.array.slice", "es.array.splice", "es.array.join", "es.array.push", "es.array.sort",
  "es.function.name", "es.json.parse", "es.json.stringify", "es.number.to-fixed", "es.object.freeze", "es.object.get-own-property-descriptor",
  "es.object.get-own-property-names", "es.object.keys", "es.object.to-string", "es.object.get-prototype-of", "es.reflect.construct",
  "es.regexp.exec", "es.regexp.test", "es.regexp.to-string", "es.string.match", "es.string.replace", "es.string.search", "es.string.split", "es.string.trim",
  "es.symbol.to-primitive", "es.date.to-primitive", /^es\.iterator\./, /^es\.set\..*\.v2$/, /^es\.(weak-)?map\.get-or-insert/
];

function transpileForAcrobat(entryPath) {
  // Bundle modern sources, lower them to ES5 with core-js polyfills, bundle the
  // polyfills, lower esbuild's own helpers, then prove the result parses as ES5.
  const modern = esbuild.buildSync({ entryPoints: [entryPath], bundle: true, format: "iife", globalName: "BT_LIB", write: false, target: "es2020", charset: "utf8", logLevel: "error" }).outputFiles[0].text;
  const lowered = babel.transformSync(modern, {
    babelrc: false, configFile: false, sourceType: "script", compact: false, comments: false,
    presets: [["@babel/preset-env", { targets: { ie: "11" }, useBuiltIns: "usage", corejs: "3.51", exclude: ["web.dom-collections.iterator", "web.dom-collections.for-each", "transform-typeof-symbol", ...REPLACES_CONSTRUCTORS, ...WRAPS_NATIVE_METHODS] }]]
  }).code;
  fs.writeFileSync(path.join(BUILD, "acrobat-lowered.js"), lowered + "\nmodule.exports = BT_LIB;\n");
  const withPolyfills = esbuild.buildSync({ entryPoints: [path.join(BUILD, "acrobat-lowered.js")], bundle: true, format: "iife", globalName: "BT_LIB", write: false, platform: "neutral", mainFields: ["main"], charset: "utf8", logLevel: "error" }).outputFiles[0].text;
  // The second pass sees core-js too; rewriting its typeof checks breaks the Symbol polyfill.
  const lowered2 = babel.transformSync(withPolyfills, { babelrc: false, configFile: false, sourceType: "script", compact: false, comments: false, presets: [["@babel/preset-env", { targets: { ie: "11" }, useBuiltIns: false, exclude: ["transform-typeof-symbol"] }]] }).code;
  const final = tolerateReadOnlyBuiltIns(lowered2);
  acorn.parse(final, { ecmaVersion: 5, sourceType: "script" });
  return final;
}

// Acrobat makes some built-ins read-only (Function.prototype.toString, at least), and
// core-js's strict-mode code throws when it tries to replace them. Where a replacement is
// refused, keep Acrobat's own built-in and carry on; adding missing methods is unaffected.
function tolerateReadOnlyBuiltIns(code) {
  const replace = (pattern, replacement, description) => {
    if (!pattern.test(code)) throw new Error("core-js layout changed; cannot patch " + description);
    code = code.replace(pattern, replacement);
  };
  code = code.replace(/^(\s*)(\w+Prototype\.constructor = \w+;)$/gm, "$1try { $2 } catch (readOnly) {}");
  replace(/(Function\.prototype\.toString = makeBuiltIn\(function toString\(\) \{[\s\S]*?\}, "toString"\);)/,
    "try { $1 } catch (readOnly) {}", "Function.prototype.toString");
  replace(/("node_modules\/core-js\/internals\/define-built-in\.js": function [\w$]+\(exports, module\) \{[\s\S]*?)module\.exports = function \(O, key, value, options\) \{([\s\S]*?\n\s*return O;\n\s*\};)/,
    "$1var defineBuiltInStrict = function (O, key, value, options) {$2\n      module.exports = function (O, key, value, options) { try { return defineBuiltInStrict(O, key, value, options); } catch (readOnly) { return O; } };",
    "define-built-in");
  // Load each polyfill on its own: one Acrobat refuses is skipped (and logged), keeping
  // Acrobat's native version, instead of stopping the whole program from loading.
  const before = code;
  code = code.replace(/^(\s*)(require_(?:es|esnext|web)_[\w$]+\(\);)$/gm, "$1try { $2 } catch (polyfillError) { BT_polyfillErrors.push(String(polyfillError)); }");
  if (code === before) throw new Error("core-js layout changed; cannot guard polyfill loading");
  return code;
}

// Acrobat has no queueMicrotask or setTimeout. core-js's Promise looks for both on the
// global object, so they must exist before the polyfills load; both feed BT_tasks.
const PRELUDE = [
  "var BT_tasks = { micro: [], timers: [], seq: 0, now: 0 }, BT_polyfillErrors = [];",
  "(function (g) {",
  "  g.queueMicrotask = function (callback) { BT_tasks.micro.push(callback); };",
  "  if (typeof g.setTimeout !== 'function') {",
  "    g.setTimeout = function (callback, delay) { var id = ++BT_tasks.seq; if (typeof callback === 'function') BT_tasks.timers.push({ id: id, due: BT_tasks.now + Math.max(0, Number(delay) || 0), seq: id, fn: callback, args: Array.prototype.slice.call(arguments, 2), interval: 0 }); return id; };",
  "    g.clearTimeout = function (id) { for (var i = 0; i < BT_tasks.timers.length; i++) if (BT_tasks.timers[i].id === id) { BT_tasks.timers.splice(i, 1); return; } };",
  "  }",
  "})(Function('return this')());"
].join("\n");
// Every stage reports into the status line, so a failure in Acrobat says where it stopped:
// "Starting…" stuck means the program failed to load; the build-time text still showing
// means the document script never ran (the page actions then say JavaScript is blocked).
const statusScript = pageCount => [
  "var BT_doc = this, BT_loadStarted = Date.now();",
  "function BT_status(text) { for (var i = 0; i < " + pageCount + "; i++) { try { var f = BT_doc.getField('p' + i + '.status'); if (f) f.value = text; } catch (ignored) {} } }",
  "BT_status('Starting BT1101 Prose Lab…');",
  "var BT_loadError = null;"
].join("\n");
const pageOpenScript = page => `if (typeof BT_page === 'function') BT_page(${page}); else { try { this.getField('p${page}.status').value = 'The document script did not run. Check Preferences > JavaScript > Enable Acrobat JavaScript, and that the file is not open in Protected View.'; } catch (e) {} }`;
// How long a page must stay in view before its operation opens (scrolling past is free).
const PAGE_SETTLE_MS = 250;
// Acrobat tints every field blue ("Highlight Existing Fields"). Turn that off while this
// document is open so it looks like the website; the close script restores the setting.
const RESTORE_HIGHLIGHT = "try { if (typeof BT_highlight === 'boolean') app.runtimeHighlight = BT_highlight; } catch (ignored) {}";
const GLUE = [
  "var BT1101 = null, BT_message = '', BT_highlight = null;",
  "function BT_commit(e) { if (BT1101) BT1101.commit(e.target.name, e.value, e); }",
  "function BT_toggle(e) { if (BT1101) BT1101.toggle(e.target.name, e.target.value !== 'Off'); }",
  "function BT_press(e) { if (BT1101) BT1101.press(e.target.name); }",
  "var BT_pendingPage = null, BT_pageTimer = null;",
  "function BT_pageNow() { BT_pageTimer = null; var n = BT_pendingPage; BT_pendingPage = null; if (BT1101 && n !== null) BT1101.pageOpen(n); }",
  "function BT_page(n) {",
  "  if (!BT1101) { if (BT_message) BT_status(BT_message); return; }",
  "  BT_pendingPage = n;",
  "  try { if (BT_pageTimer) app.clearTimeOut(BT_pageTimer); BT_pageTimer = app.setTimeOut('BT_pageNow()', " + PAGE_SETTLE_MS + "); }",
  "  catch (error) { BT_pageNow(); }",
  "}",
  "function BT_describe(error) { return String(error) + (error && error.lineNumber ? ' (line ' + error.lineNumber + ')' : ''); }",
  "if (BT_loadError || typeof BT_LIB === 'undefined') BT_message = 'Could not load: ' + BT_describe(BT_loadError);",
  "else { try { BT1101 = BT_LIB.boot({ doc: this, app: app, display: display, console: console }); } catch (error) { BT_message = 'Could not start: ' + BT_describe(error); } }",
  "if (BT_polyfillErrors.length) { try { console.println('BT1101: ' + BT_polyfillErrors.length + ' polyfill(s) skipped: ' + BT_polyfillErrors.join(' | ')); } catch (ignored) {} }",
  "if (BT_message) { BT_status(BT_message); try { console.println(BT_message); app.alert('BT1101 Prose Lab ' + BT_message); } catch (ignored) {} }",
  "else { try { BT_highlight = app.runtimeHighlight; app.runtimeHighlight = false; } catch (ignored) {} }"
].join("\n");

const colourList = colour => (colour && colour.a !== 0 ? [colour.r, colour.g, colour.b].map(value => Math.round(value * 1000) / 1000) : null);

async function main() {
  const packageJson = JSON.parse(fs.readFileSync(path.join(ROOT, "package.json"), "utf8"));
  const version = parseVersion((process.argv.slice(2).filter(arg => !arg.startsWith("--"))[0]) || packageJson.pdfVersion);
  fs.mkdirSync(BUILD, { recursive: true });
  const site = extractWebsite(INDEX);
  const lib = bundleNodeLibraries(site);
  const { explore, layoutVariants } = require("./src/explore.js");
  const { resolveRef } = require(path.join(BUILD, "refs-node.js"));
  const boot = makeBoot(lib, site);
  const operations = operationList(boot);

  console.log("Exploring operations on minidom…");
  // Exploring every layout takes a while; the result depends only on the website and the
  // exploring code, so it is cached against a hash of both.
  const cacheFile = path.join(BUILD, "layouts-cache.json");
  const cacheKey = require("crypto").createHash("sha1").update(site.html).update(fs.readFileSync(path.join(ROOT, "src", "explore.js"))).update(fs.readFileSync(path.join(ROOT, "src", "refs.js"))).update(fs.readFileSync(path.join(ROOT, "src", "minidom.js"))).digest("hex");
  let cached = null;
  try { cached = JSON.parse(fs.readFileSync(cacheFile, "utf8")); } catch (error) { cached = null; }
  let explorations = {}, layouts = {};
  if (cached && cached.key === cacheKey) {
    ({ explorations, layouts } = cached);
    console.log("Using cached layouts (website unchanged)");
  } else {
    operations.forEach(operation => {
      const started = Date.now();
      explorations[operation.id] = explore(boot, operation.id);
      layouts[operation.id] = layoutVariants(boot, operation.id, explorations[operation.id]);
      console.log(`  ${operation.id.padEnd(20)} ${layouts[operation.id].variants.length} layout(s)  ${((Date.now() - started) / 1000).toFixed(1)} s`);
    });
    fs.writeFileSync(cacheFile, JSON.stringify({ key: cacheKey, explorations, layouts }));
  }
  const plan = pagePlan(operations, explorations, layouts);
  console.log(`Measuring the website in headless Edge (${plan.length} pages)`);
  const siteFonts = loadWebsiteFonts(site.html);
  const captures = await captureAll(plan, explorations, layouts, version, drawableCodePoints(siteFonts.regular));

  const registry = { pages: [], fields: {}, layout: {}, palette: captures[0].palette };
  const pageOfOperation = {};
  plan.forEach((entry, index) => { if (entry.first) pageOfOperation[entry.operation.id] = index; });
  operations.forEach(operation => { if (layouts[operation.id].variants.length > 1) registry.layout[operation.id] = layouts[operation.id].tracked; });
  const problems = [];
  const pages = captures.map((capture, pageIndex) => {
    const operation = plan[pageIndex].operation;
    const session = boot();
    const entry = operation.id === "universal" ? session.document.querySelector(".universal-entry") : session.document.querySelector(`#tutorialList .operation[data-id="${operation.id}"]`);
    entry.click(); session.drain();
    let count = 0;
    const clipNames = {};
    capture.fields.forEach(field => { field.name = `p${pageIndex}.${field.k}${++count}`; if (field.k === "clip") clipNames[field.clip] = field.name; });
    const muted = capture.palette.muted || [0.373, 0.388, 0.408];
    const status = { k: "status", name: `p${pageIndex}.status`, rect: capture.status, text: { size: 12, weight: 600, colour: { r: muted[0], g: muted[1], b: muted[2] }, align: "right" }, value: NO_SCRIPT_MESSAGE, multiline: true };
    const pickers = [];
    capture.fields.forEach(field => {
      if (field.k !== "select" && field.k !== "plot") return;
      field.pick = field.name.replace(/\.(select|plot)(\d+)$/, ".pick$2");
      pickers.push({ k: "pick", name: field.pick, r: field.r, rect: field.rect, d: field.name, plot: field.k === "plot", text: field.text, label: field.label });
    });
    const fields = capture.fields.concat(pickers, [status]);
    const names = [];
    fields.forEach(field => {
      const spec = { k: field.k, r: field.r || null };
      if (field.k === "button") { spec.c = clipNames[field.clip] || null; spec.t = colourList(field.text && field.text.colour); spec.cap = field.caption; }
      if (field.fill && !field.backdrop && ["button", "out", "text", "search", "select", "cardbody"].indexOf(field.k) >= 0) spec.f = colourList(field.fill);
      if (field.last) spec.last = true;
      if (field.pick) { spec.p = field.pick; spec.t = colourList(field.text && field.text.colour); }
      if (field.k === "pick") { spec.d = field.d; if (field.plot) spec.plot = true; }
      registry.fields[field.name] = spec;
      names.push(field.name);
      // Every reference must find its element in the page Acrobat will run. Rows the
      // website appends while typing (summary columns, constraints) appear later.
      const appendedRow = /^q:#(columnFields|lpConstraintRows)\|/.test(spec.r || "");
      if (spec.r && !/^(slot|row):/.test(spec.r) && !appendedRow && field.k !== "label" && !resolveRef(session.document, spec.r)) problems.push(`${operation.id}: ${field.k} ${spec.r} does not resolve`);
    });
    // Words that come and go are turned into label fields by the writer; register them too.
    const owners = [];
    capture.words.forEach(word => { if (word.owner && owners.indexOf(word.owner) < 0) owners.push(word.owner); });
    owners.forEach(owner => { if (!resolveRef(session.document, owner)) problems.push(`${operation.id}: label owner ${owner} does not resolve`); });
    registry.pages.push({ op: operation.id, fields: names, clip: Object.values(clipNames)[0] || null, sig: plan[pageIndex].sig });
    return Object.assign(capture, { fields, open: pageOpenScript(pageIndex) });
  });

  const fontSet = loadWebsiteFonts(site.html);
  const fonts = { SSP: fontSet.regular, SSPS: fontSet.semibold, SSPB: fontSet.bold };
  pages.forEach((page, pageIndex) => {
    labelFields(page, fonts).forEach((label, index) => {
      label.name = `p${pageIndex}.label${index + 1}`;
      page.fields.push(label);
      registry.fields[label.name] = { k: "label", r: label.r };
      registry.pages[pageIndex].fields.push(label.name);
    });
  });
  const outline = operations.map(operation => ({ title: operation.title, page: pageOfOperation[operation.id] }));
  const assemble = scripts => buildPdf({
    pages, fonts, scripts, outline, pageOfOperation, closeScript: RESTORE_HIGHLIGHT, palette: captures[0].palette,
    title: "BT1101 Prose Lab", subject: "RStudio output to prose and R code, for Adobe Acrobat Reader",
    version: `v${version}h`, headerComment: `BT1101 Prose Lab version: v${version}h`
  });
  fs.writeFileSync(path.join(BUILD, "site-data.js"), `export const BODY = ${JSON.stringify(site.body)};\nexport const REGISTRY = ${JSON.stringify(registry)};\nexport const VERSION = ${JSON.stringify("v" + version + "h")};\n`);
  const runtime = transpileForAcrobat(path.join(ROOT, "src", "acrobat-entry.js"));
  // A load failure inside the program must not stop the status reporting that follows it.
  // Acrobat reads the script bytes with the system's text encoding, not UTF-8; on a
  // double-byte code page a non-ASCII character can swallow a closing quote. Every such
  // character is written as \uXXXX, which means the same in strings and regular expressions.
  const docScript = `${statusScript(pages.length)}
${PRELUDE}
try {
${runtime}
} catch (BT_error) { BT_loadError = BT_error; }
${GLUE}
`.replace(/[^\x00-\x7e]/g, char => "\\u" + char.charCodeAt(0).toString(16).padStart(4, "0"));
  if (/[^\x09\x0a\x0d\x20-\x7e]/.test(docScript)) throw new Error("The document script must be plain ASCII for Acrobat.");
  acorn.parse(docScript, { ecmaVersion: 5, sourceType: "script" });
  fs.writeFileSync(path.join(BUILD, "document-script.js"), docScript);
  const pdf = assemble([{ name: "BT1101 Prose Lab", code: docScript }]);
  // Acrobat locks a PDF it has open; write a second copy rather than lose the build.
  let written = OUTPUT;
  try { fs.writeFileSync(OUTPUT, pdf); }
  catch (error) {
    if (error.code !== "EBUSY" && error.code !== "EPERM") throw error;
    written = OUTPUT.replace(/\.pdf$/, " (new).pdf");
    fs.writeFileSync(written, pdf);
    console.log(`\n${path.basename(OUTPUT)} is open in another program, so this build went to ${path.basename(written)}.`);
  }
  pages.forEach((page, index) => console.log(`${plan[index].operation.id.padEnd(20)} ${plan[index].sig ? "layout  " : "catch-all"} ${String(registry.pages[index].fields.length).padStart(3)} fields  ${page.width}x${page.height}px`));
  if (problems.length) { console.log("\nReference problems:\n" + problems.join("\n")); process.exitCode = 1; }
  console.log(`\nWrote ${written} (${(pdf.length / 1024).toFixed(0)} KB, script ${(docScript.length / 1024).toFixed(0)} KB, ${Object.keys(registry.fields).length} fields)`);
}

if (require.main === module) main().catch(error => { console.error(error); process.exit(1); });
module.exports = { parseVersion, transpileForAcrobat };
