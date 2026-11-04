// Builds "BT1101 Prose Lab.xlsm" (desktop Excel, Windows and Mac) from ../index.html and
// the VBA port in vba/.
//   node build.js              uses excelVersion from package.json
//   node build.js 021126_2300  sets the version (ddmmyy_hhmm)
//
// The website's body markup (version-stamped), the stylesheet rules (tools/css-rules.js)
// and the Source Sans Pro fonts go on the very hidden "BTX" sheet; every non-test VBA
// source is imported; the start sheet carries the version and an open button.

const fs = require("fs");
const path = require("path");
const { spawnSync } = require("child_process");
const { extractWebsite } = require("../pdf/src/extract.js");
const { parseVersion } = require("../pdf/build.js");
const { prepareSource, excelTask } = require("./tools/vba.js");

const ROOT = __dirname;
const BUILD = path.join(ROOT, ".build", "release");
const VBA = path.join(ROOT, "vba");
const OUTPUT = path.join(ROOT, "..", "BT1101 Prose Lab.xlsm");
const CHUNK = 30000;   // characters per cell (Excel's limit is 32,767)
const VERSION_CONSTANT = /(Const\s+APP_VERSION\s+As\s+String\s*=\s*")\d{6}_\d{4}(")/i;

function writeChunks(name, text) {
  const files = [];
  for (let i = 0, n = 0; i < text.length || n === 0; i += CHUNK, n++) {
    const file = path.join(BUILD, "cells", `${name}-${String(n).padStart(3, "0")}.txt`);
    fs.writeFileSync(file, text.slice(i, i + CHUNK), "utf8");
    files.push(file);
  }
  return { name, files };
}

function siteFonts(html) {
  const fonts = {};
  for (const match of html.matchAll(/@font-face\s*{([^}]*)}/g)) {
    const data = /base64,([A-Za-z0-9+/=]+)/.exec(match[1]);
    const weight = /font-weight:\s*(\d+)/.exec(match[1]);
    if (data) fonts[weight ? weight[1] : "400"] = data[1];
  }
  return fonts;
}

function main() {
  const packageJson = JSON.parse(fs.readFileSync(path.join(ROOT, "package.json"), "utf8"));
  const version = parseVersion(process.argv[2] || packageJson.excelVersion);
  const display = `v${version}h`;
  fs.rmSync(BUILD, { recursive: true, force: true });
  ["cells", "vba"].forEach(dir => fs.mkdirSync(path.join(BUILD, dir), { recursive: true }));

  const site = extractWebsite(path.join(ROOT, "..", "index.html"));
  const body = site.body.replace(/(<span class="app-version">)[^<]*(<\/span>)/, `$1${display}$2`);
  if (!body.includes(display)) throw new Error("index.html no longer has the .app-version span");

  // Stylesheet rules for the renderer.
  const cssTool = path.join(ROOT, "tools", "css-rules.js");
  if (!fs.existsSync(cssTool)) throw new Error("tools/css-rules.js is missing");
  const css = spawnSync(process.execPath, [cssTool], { encoding: "utf8", cwd: ROOT });
  if (css.status !== 0) throw new Error("css-rules.js failed: " + css.stderr);
  const rules = fs.readFileSync(path.join(ROOT, ".build", "css-rules.txt"), "utf8");

  const fonts = siteFonts(site.html);
  const store = [
    writeChunks("body", body),
    writeChunks("css", rules),
    writeChunks("version", version),
    ...Object.keys(fonts).sort().map(weight => writeChunks("font" + weight, fonts[weight]))
  ];

  // Every non-test source; the edition's version is stamped into APP_VERSION.
  const sources = fs.readdirSync(VBA).filter(name => /\.(bas|cls)$/i.test(name) && !/^test/i.test(name));
  let stamped = 0;
  const modules = sources.map(name => {
    const file = prepareSource(name, path.join(BUILD, "vba"));
    const text = fs.readFileSync(file, "latin1");
    if (VERSION_CONSTANT.test(text)) { fs.writeFileSync(file, text.replace(VERSION_CONSTANT, `$1${version}$2`), "latin1"); stamped++; }
    return file;
  });
  if (stamped !== 1) throw new Error(`Expected exactly one "Const APP_VERSION As String" in vba/, found ${stamped}`);

  const job = {
    action: "build", output: OUTPUT, modules, store,
    startSheet: { version: display },
    form: { name: "frmApp", caption: "BT1101 Prose Lab", code: prepareSource("frmApp.code.vba", path.join(BUILD, "vba")) },
    workbookCode: prepareSource("ThisWorkbook.code.vba", path.join(BUILD, "vba")),
    properties: { Title: "BT1101 Prose Lab", Subject: "RStudio output to prose and R code", Comments: `BT1101 Prose Lab version: ${display}` }
  };
  console.log(`Building ${display}: ${modules.length} modules, ${store.map(item => `${item.name} ${item.files.length}`).join(", ")} chunks`);
  const built = excelTask(job, BUILD);
  if (built.status !== 0) throw new Error(`Excel build failed: ${built.stdout} ${built.stderr} ${built.dialogs}`);
  const compiled = excelTask({ action: "compile", workbook: OUTPUT }, BUILD);
  if (compiled.status !== 0) throw new Error(`Compile failed: ${compiled.stdout}\n${compiled.dialogs}`);
  console.log(`Wrote ${OUTPUT} (${(fs.statSync(OUTPUT).size / 1024).toFixed(0)} KB)`);
}

if (require.main === module) {
  try { main(); } catch (error) { console.error(error.message); process.exitCode = 1; }
}
