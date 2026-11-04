// End-to-end tests of the Excel app through its real window (vba/testE2E.bas).
//
//   node test/e2e/run.js [scenario ...] [--parallel N] [--out dir]
//
// With no scenarios the whole suite runs in batches (one Excel instance per batch, so a
// batch stays well inside the toolchain's time limit). Scenario names:
//   universal search letter theme resizer resize plotsearch stats axis kmeans knownsd
//   lprows lpsens anova units revisit   op:<operation id>   ops (every operation)   all
// A batch is a group joined with "+", e.g. "op:ggplot+op:summarize".
//
// Writes <out>/<batch>.tsv (STATUS, scenario, check, detail) and <out>/summary.txt, and
// prints the failures and warnings. Exit code 1 when any check failed. The VBA sources and
// the toolchain are copied to <out>/snapshot first, so every batch builds the same sources
// even while other people edit excel/vba.
const fs = require("fs");
const os = require("os");
const path = require("path");
const { spawn, spawnSync } = require("child_process");

const EXCEL = path.join(__dirname, "..", "..");
const ROOT = path.join(EXCEL, "..");
const OPS = ["universal", "ggplot", "summarize", "shape", "wide-table", "interval", "one-sample", "two-sample", "anova-many",
  "lm-predict", "multi-regression", "model-significance", "model-comparison", "dummy", "interaction", "glm", "confusion",
  "pca", "stepwise", "kmeans", "kfold", "lpsolve", "shapiro", "pearson"];
const DEFAULT_BATCHES = [
  "universal search letter theme resizer plotsearch",
  "stats axis kmeans knownsd",
  "lprows lpsens anova units revisit resize",
  ...chunk(OPS.map(id => "op:" + id), 3).map(list => list.join(" "))
];

function chunk(list, size) {
  const out = [];
  for (let i = 0; i < list.length; i += size) out.push(list.slice(i, i + size));
  return out;
}

function parseArgs(argv) {
  const options = { parallel: 1, out: path.join(os.tmpdir(), "bt1101-e2e"), scenarios: [] };
  for (let i = 0; i < argv.length; i++) {
    if (argv[i] === "--parallel") options.parallel = Math.max(1, Number(argv[++i]) || 1);
    else if (argv[i] === "--out") options.out = path.resolve(argv[++i]);
    else options.scenarios.push(argv[i]);
  }
  return options;
}

// css-rules.txt, body.html and fonts/font-*.b64 from the website, in a private folder.
function prepareBuild(dir) {
  fs.mkdirSync(path.join(dir, "fonts"), { recursive: true });
  const css = spawnSync(process.execPath, [path.join(EXCEL, "tools", "css-rules.js"), path.join(ROOT, "index.html"), path.join(dir, "css-rules.txt")], { encoding: "utf8" });
  if (css.status !== 0) throw new Error("css-rules.js failed: " + css.stderr);
  const { extractWebsite } = require(path.join(ROOT, "pdf", "src", "extract.js"));
  const site = extractWebsite(path.join(ROOT, "index.html"));
  fs.writeFileSync(path.join(dir, "body.html"), site.body);
  const html = fs.readFileSync(path.join(ROOT, "index.html"), "utf8");
  for (const match of html.matchAll(/@font-face\s*{([^}]*)}/g)) {
    const data = /base64,([^")]+)/.exec(match[1]);
    const weight = /font-weight:\s*(\d+)/.exec(match[1]);
    if (data) fs.writeFileSync(path.join(dir, "fonts", "font-" + (weight ? weight[1] : "400") + ".b64"), data[1]);
  }
}

// Copies excel/tools and excel/vba to <out>/snapshot (vba.js finds its sources next to itself).
function snapshotSources(outDir) {
  const snap = path.join(outDir, "snapshot");
  fs.rmSync(snap, { recursive: true, force: true });
  for (const sub of ["tools", "vba"]) {
    fs.mkdirSync(path.join(snap, sub), { recursive: true });
    for (const name of fs.readdirSync(path.join(EXCEL, sub))) {
      const from = path.join(EXCEL, sub, name);
      if (fs.statSync(from).isFile()) fs.copyFileSync(from, path.join(snap, sub, name));
    }
  }
  return snap;
}

function runBatch(name, scenarios, buildDir, outDir, snap) {
  return new Promise(resolve => {
    const outFile = path.join(outDir, name + ".tsv");
    const keep = path.join(outDir, name + ".xlsm");
    if (fs.existsSync(outFile)) fs.unlinkSync(outFile);
    const started = Date.now();
    const child = spawn(process.execPath, [path.join(snap, "tools", "vba.js"), "run", "testE2E.RunSuite", buildDir, outFile, scenarios, "--keep", keep], { cwd: snap });
    let log = "";
    child.stdout.on("data", d => { log += d; });
    child.stderr.on("data", d => { log += d; });
    child.on("close", code => {
      const rows = fs.existsSync(outFile) ? fs.readFileSync(outFile, "utf8").replace(/^﻿/, "").split("\n").filter(Boolean).map(line => line.split("\t")) : [];
      try { fs.unlinkSync(keep); } catch (ignored) { /* kept by a crash */ }
      resolve({ name, scenarios, code, log: log.trim(), rows, seconds: Math.round((Date.now() - started) / 1000) });
    });
  });
}

async function main() {
  const options = parseArgs(process.argv.slice(2));
  fs.mkdirSync(options.out, { recursive: true });
  const buildDir = path.join(options.out, "build");
  prepareBuild(buildDir);
  const snap = snapshotSources(options.out);
  let batches = options.scenarios.length ? options.scenarios.map(s => s.split("+").join(" ")) : DEFAULT_BATCHES;
  if (batches.length === 1 && batches[0] === "all") batches = DEFAULT_BATCHES;
  const named = batches.map((scenarios, i) => ({ name: "batch" + String(i + 1).padStart(2, "0"), scenarios }));
  const results = [];
  const queue = named.slice();
  const workers = Array.from({ length: Math.min(options.parallel, queue.length) }, async () => {
    while (queue.length) {
      const batch = queue.shift();
      console.log(`running ${batch.name}: ${batch.scenarios}`);
      const result = await runBatch(batch.name, batch.scenarios, buildDir, options.out, snap);
      console.log(`  ${batch.name} done in ${result.seconds} s: ${result.log.split("\n")[0]}`);
      results.push(result);
    }
  });
  await Promise.all(workers);
  results.sort((a, b) => a.name.localeCompare(b.name));
  let pass = 0, fail = 0, warn = 0;
  const lines = [];
  const perScenario = new Map();
  for (const result of results) {
    if (!result.rows.length) { fail++; lines.push(`FAIL\t${result.name}\tbatch produced no results\t${result.log.replace(/\s+/g, " ").slice(0, 600)}`); }
    else if (result.code !== 0) { fail++; lines.push(`FAIL\t${result.name}\tbatch exit code ${result.code}\t${result.log.replace(/\s+/g, " ").slice(0, 600)}`); }
    for (const [status, scenario, check, detail] of result.rows) {
      const s = perScenario.get(scenario) || { pass: 0, fail: 0, warn: 0 };
      if (status === "PASS") { pass++; s.pass++; } else if (status === "FAIL") { fail++; s.fail++; } else { warn++; s.warn++; }
      perScenario.set(scenario, s);
      if (status !== "PASS") lines.push([status, scenario, check, detail || ""].join("\t"));
    }
  }
  const table = ["scenario\tpass\tfail\twarn", ...[...perScenario].map(([name, s]) => `${name}\t${s.pass}\t${s.fail}\t${s.warn}`)].join("\n");
  const summary = `${pass} passed, ${fail} failed, ${warn} warnings\n\n${table}\n\n${lines.join("\n")}\n`;
  fs.writeFileSync(path.join(options.out, "summary.txt"), summary);
  console.log("\n" + summary);
  console.log("Results: " + options.out);
  if (fail) process.exitCode = 1;
}

main().catch(error => { console.error(error); process.exitCode = 1; });
