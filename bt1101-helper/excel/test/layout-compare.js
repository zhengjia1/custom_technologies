// Compares the VBA layout engine (clsRenderLayout via testRender.RunLayoutSuite and
// modRender.LayoutSnapshot) with the reference JavaScript engine
// (excel/.build/obsolete/src/layout.js) on the same DOM states: every operation bench made
// visible, light and dark, 1280/960/800/375 px.
//
//   node test/layout-compare.js --run [--detail bench/theme/width]
//       prepares the inputs (css-rules.txt, body.html, base64 fonts in .build/fonts),
//       runs the VBA suite in Excel (writes .build/vba-layout.txt) and compares
//   node test/layout-compare.js <vba-snapshots.txt> [--detail bench/theme/width]
//       compares an existing VBA snapshot file
//
// Known reference difference: the reference resolves vw as 0 inside font-size, so the
// main heading (clamp(24px, 2.4vw, 32px)) is ~7px shorter there than on the website and
// in the VBA engine; everything below it is offset by that amount.
//
// Both sides lay out excel/.build/body.html (node test/body.js) with only the bench's
// "hidden" class toggled and data-theme set. Items are matched by kind and id (element
// paths), positions compared in window coordinates. Text wrapping differs a little
// (the VBA side measures with real MSForms labels, the reference with font tables),
// so the report gives the share of items within 2 px / 8 px and the largest offsets.

const fs = require("fs");
const path = require("path");

const EXCEL = path.join(__dirname, "..");
const ROOT = path.join(EXCEL, "..");
const MODULES = path.join(ROOT, "pdf", "node_modules");
const OBSOLETE = path.join(EXCEL, ".build", "obsolete", "src");
const BENCHES = "standardBench summarizeBench shapeBench wideTableBench intervalBench oneSampleBench twoSampleBench anovaBench corrMultipleBench simpleRegressionBench advancedRegressionBench stepwiseBench kfoldBench pcaBench kmeansBench lpBench ggplotBench".split(" ");
const WIDTHS = [1280, 960, 800, 375];
const THEMES = ["light", "dark"];

async function loadReference() {
  const esbuild = require(path.join(MODULES, "esbuild"));
  const out = path.join(EXCEL, ".build", "layout-ref-bundle.js");
  const minidom = path.join(ROOT, "pdf", "src", "minidom.js");
  await esbuild.build({
    stdin: { contents: `export * as dom from ${JSON.stringify(minidom)}; export { createLayoutEngine } from ${JSON.stringify(path.join(OBSOLETE, "layout.js"))};`, resolveDir: EXCEL },
    bundle: true, format: "cjs", platform: "node", outfile: out, logLevel: "error", nodePaths: [MODULES],
    plugins: [{ name: "minidom", setup(build) { build.onResolve({ filter: /minidom\.js$/ }, () => ({ path: minidom })); } }]
  });
  delete require.cache[out];
  return require(out);
}

function pathOf(element) {
  const parts = [];
  for (let node = element; node && node.parentNode && node.parentNode.nodeType === 1; node = node.parentNode) {
    parts.unshift(node.parentNode.children.indexOf(node));
  }
  return "p" + parts.join(".");
}

// Unique ids: the reference gives pseudo-element boxes the id "anon"; number repeats.
function uniqueIds(items) {
  const seen = {};
  items.forEach(item => { const n = seen[item.key] = (seen[item.key] || 0) + 1; if (n > 1) item.key += "#" + n; });
}

async function referenceSnapshots() {
  const lib = await loadReference();
  const { loadSiteFontMetrics } = require(path.join(OBSOLETE, "font-metrics.js"));
  const html = fs.readFileSync(path.join(ROOT, "index.html"), "utf8");
  const cssText = /<style>([\s\S]*?)<\/style>/i.exec(html)[1];
  const body = fs.readFileSync(path.join(EXCEL, ".build", "body.html"), "utf8");
  const engine = lib.createLayoutEngine({ cssText, fonts: loadSiteFontMetrics(html) });
  const document = lib.dom.createDocument(body);
  const ids = new Map();
  const idOf = element => { let id = ids.get(element); if (!id) { id = pathOf(element); ids.set(element, id); } return id; };
  const result = {};
  BENCHES.forEach(bench => {
    document.querySelectorAll("main > section").forEach(section => section.classList.toggle("hidden", section.id !== bench));
    THEMES.forEach(theme => {
      document.documentElement.setAttribute("data-theme", theme);
      WIDTHS.forEach(width => {
        const list = engine.layout(document, { width, height: 860, idOf });
        result[`${bench} ${theme} ${width}`] = toBoxes(list.regions, list.items.map(item => ({ region: item.region, id: item.id, kind: item.kind, x: item.x, y: item.y, w: item.w, h: item.h, text: item.text || "" })));
      });
    });
  });
  return result;
}

function toBoxes(regions, items) {
  const byId = {};
  regions.forEach(region => { byId[region.id] = region; });
  const origin = id => { const r = byId[id]; if (!r) return { x: 0, y: 0 }; const p = r.parent ? origin(r.parent) : { x: 0, y: 0 }; return { x: p.x + r.x, y: p.y + r.y }; };
  const out = items.map(item => {
    const o = origin(item.region);
    return { key: item.kind + " " + normalise(item.id), kind: item.kind, region: item.region, x: o.x + item.x, y: o.y + item.y, w: item.w, h: item.h, text: item.text };
  });
  uniqueIds(out);
  return { regions: regions.map(r => ({ id: r.id, x: r.x, y: r.y, w: r.w, h: r.h, contentH: r.contentH })), items: out };
}

function vbaSnapshots(file) {
  const result = {};
  let regions = null, items = null;
  fs.readFileSync(file, "utf8").split(/\r?\n/).forEach(line => {
    if (line.startsWith("=== ")) {
      regions = []; items = [];
      result[line.slice(4).trim()] = { regions, items };
      return;
    }
    const f = line.split("\t");
    if (f[0] === "R") regions.push({ id: f[1], parent: f[2] || null, x: +f[3], y: +f[4], w: +f[5], h: +f[6], contentH: +f[7] });
    else if (f[0] === "I") items.push({ region: f[1], id: f[2], kind: f[3], x: +f[4], y: +f[5], w: +f[6], h: +f[7], text: f[8] || "" });
  });
  Object.keys(result).forEach(key => { const snap = result[key]; result[key] = toBoxes(snap.regions, snap.items.map(item => Object.assign(item, { id: item.id.replace(/^anon(?=:)/, "anon") }))); });
  return result;
}

function compare(ref, vba) {
  const refMap = new Map(ref.items.filter(item => item.kind !== "text").map(item => [item.key, item]));
  const vbaMap = new Map(vba.items.filter(item => item.kind !== "text").map(item => [item.key, item]));
  let matched = 0, within2 = 0, within8 = 0;
  const offsets = [];
  refMap.forEach((r, key) => {
    const v = vbaMap.get(key);
    if (!v) return;
    matched++;
    const d = Math.max(Math.abs(r.x - v.x), Math.abs(r.y - v.y), Math.abs(r.w - v.w), Math.abs(r.h - v.h));
    if (d <= 2) within2++;
    if (d <= 8) within8++;
    offsets.push({ key, d, ref: r, vba: v });
  });
  offsets.sort((a, b) => b.d - a.d);
  const textRef = ref.items.filter(item => item.kind === "text"), textVba = vba.items.filter(item => item.kind === "text");
  const regionDiffs = [];
  ref.regions.forEach(r => {
    const v = vba.regions.find(x => x.id === r.id);
    if (!v) regionDiffs.push(`missing region ${r.id}`);
    else if (Math.max(Math.abs(r.x - v.x), Math.abs(r.y - v.y), Math.abs(r.w - v.w), Math.abs(r.h - v.h)) > 2) regionDiffs.push(`region ${r.id}: ref ${r.x},${r.y} ${r.w}x${r.h} vba ${v.x},${v.y} ${v.w}x${v.h}`);
  });
  return {
    refCount: refMap.size, vbaCount: vbaMap.size, matched, within2, within8, offsets,
    missing: [...refMap.keys()].filter(key => !vbaMap.has(key)), extra: [...vbaMap.keys()].filter(key => !refMap.has(key)),
    textRef: textRef.length, textVba: textVba.length, regionDiffs
  };
}

// The reference ids pseudo-element boxes "anon"; the VBA engine names them after their
// element (":before"/":after"). Compare them by kind and order instead.
function normalise(id) { return id.replace(/^p[\d.]+:(before|after)(?=:|$)/, "anon").replace(/#\d+$/, ""); }

// Writes the VBA suite's inputs: the rule table, the body markup and the fonts.
function prepare() {
  const { spawnSync } = require("child_process");
  const run = (args) => { const r = spawnSync(process.execPath, args, { cwd: EXCEL, encoding: "utf8" }); if (r.status !== 0) throw new Error(args.join(" ") + ": " + r.stderr); };
  run([path.join("tools", "css-rules.js")]);
  run([path.join("test", "body.js")]);
  const html = fs.readFileSync(path.join(ROOT, "index.html"), "utf8");
  fs.mkdirSync(path.join(EXCEL, ".build", "fonts"), { recursive: true });
  for (const match of html.matchAll(/@font-face\s*{([^}]*)}/g)) {
    const data = /base64,([A-Za-z0-9+/=]+)/.exec(match[1]);
    const weight = /font-weight:\s*(\d+)/.exec(match[1]);
    if (data) fs.writeFileSync(path.join(EXCEL, ".build", "fonts", "font-" + (weight ? weight[1] : "400") + ".b64"), data[1]);
  }
}

// Runs testRender.RunLayoutSuite in Excel with every non-test module (the renderer
// depends on the shell for window.innerWidth) and returns the snapshot file path.
function runVbaSuite() {
  const { spawnSync } = require("child_process");
  const build = path.join(EXCEL, ".build");
  const out = path.join(build, "vba-layout.txt");
  const only = fs.readdirSync(path.join(EXCEL, "vba")).filter(name => /\.(bas|cls)$/i.test(name) && !/^test/i.test(name)).concat(["testUtil.bas", "testRender.bas", "frmApp.code.vba"]);
  const r = spawnSync(process.execPath, [path.join("tools", "vba.js"), "run", "testRender.RunLayoutSuite", path.join(build, "css-rules.txt"), path.join(build, "body.html"), path.join(build, "fonts"), out, "--only", only.join(",")], { cwd: EXCEL, encoding: "utf8", maxBuffer: 1 << 26 });
  process.stdout.write(r.stdout || "");
  if (r.status !== 0) throw new Error("VBA suite failed: " + r.stderr);
  return out;
}

async function main() {
  let vbaFile = process.argv[2];
  if (vbaFile === "--run") { prepare(); vbaFile = runVbaSuite(); }
  const detailIndex = process.argv.indexOf("--detail");
  const detail = detailIndex > 0 ? process.argv[detailIndex + 1].replace(/\//g, " ") : null;
  if (!vbaFile) { console.error("Usage: node test/layout-compare.js <vba-snapshots.txt> [--detail bench/theme/width]"); process.exit(2); }
  const ref = await referenceSnapshots();
  const vba = vbaSnapshots(vbaFile);
  let totalMatched = 0, total2 = 0, total8 = 0, totalRef = 0;
  Object.keys(ref).forEach(key => {
    if (!vba[key]) { console.log(`${key}: no VBA snapshot`); return; }
    const c = compare(ref[key], vba[key]);
    totalMatched += c.matched; total2 += c.within2; total8 += c.within8; totalRef += c.refCount;
    const worst = c.offsets[0];
    console.log(`${key.padEnd(36)} boxes ${String(c.matched).padStart(3)}/${String(c.refCount).padStart(3)} matched, ${pct(c.within2, c.matched)} <=2px, ${pct(c.within8, c.matched)} <=8px, worst ${worst ? worst.d.toFixed(1) + "px " + worst.key : "-"}; text runs ref ${c.textRef} vba ${c.textVba}${c.regionDiffs.length ? "; " + c.regionDiffs.join("; ") : ""}`);
    if (detail === key) {
      console.log("  missing in VBA: " + c.missing.join(", "));
      console.log("  extra in VBA: " + c.extra.join(", "));
      c.offsets.slice(0, 25).forEach(o => console.log(`  ${o.d.toFixed(1)}px ${o.key}: ref ${fmt(o.ref)} vba ${fmt(o.vba)}`));
    }
  });
  console.log(`TOTAL boxes matched ${totalMatched}/${totalRef}, ${pct(total2, totalMatched)} within 2px, ${pct(total8, totalMatched)} within 8px`);
}

const pct = (a, b) => (b ? (100 * a / b).toFixed(0) : "0") + "%";
const fmt = b => `${b.x.toFixed(1)},${b.y.toFixed(1)} ${b.w.toFixed(1)}x${b.h.toFixed(1)}`;

if (require.main === module) main();
module.exports = { referenceSnapshots, vbaSnapshots, compare };
