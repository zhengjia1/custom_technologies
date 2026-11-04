// Engine equivalence: runs the website's proseAnswerVariants/codeAnswerVariants (engine.js,
// Math.random = mulberry32) and the VBA port (testVariants.RunEngineCases) on the same
// cases and compares all six outputs of every case.
//
//   node test/variants/equiv.js [--seeds 1,2] [--limit N] [--filter text] [--only-node]
//
// Cases: every corpus text (excel/.build/variants-corpus.json plus synthetic.js) for each
// seed with no history, plus a regeneration case per text and seed whose history is the
// previous variant set (what regenerateSelectedAnswerVariant passes), and a second
// regeneration on top of it.

const fs = require("fs");
const os = require("os");
const path = require("path");
const { spawnSync } = require("child_process");
const { loadEngine } = require("./engine.js");
const synthetic = require("./synthetic.js");

const ROOT = path.join(__dirname, "..", "..");
const args = process.argv.slice(2);
const option = (name, fallback) => { const i = args.indexOf(name); return i >= 0 ? args[i + 1] : fallback; };
const seeds = option("--seeds", "1,2").split(",").map(Number);
const limit = Number(option("--limit", "0"));
const filter = option("--filter", "");

const escape = text => String(text).replace(/\\/g, "\\\\").replace(/\n/g, "\\n").replace(/\t/g, "\\t").replace(/\r/g, "\\r");
const unescape = text => text.replace(/\\(\\|n|t|r)/g, (_, c) => ({ "\\": "\\", n: "\n", t: "\t", r: "\r" })[c]);

const corpus = JSON.parse(fs.readFileSync(path.join(ROOT, ".build", "variants-corpus.json"), "utf8")).concat(synthetic.items());
let texts = corpus.filter(item => !filter || item.text.includes(filter) || item.source.includes(filter));
if (limit) texts = texts.slice(0, limit);

const engine = loadEngine();
const cases = [];
texts.forEach((item, t) => seeds.forEach(seed => {
  const s = seed * 100003 + t;
  engine.seed(s);
  const first = item.kind === "code" ? engine.code(item.text) : engine.prose(item.text);
  cases.push({ seed: s, kind: item.kind, text: item.text, history: null, expected: first, source: item.source });
  engine.seed(s + 7);
  const second = item.kind === "code" ? engine.code(item.text, first) : engine.prose(item.text, first);
  cases.push({ seed: s + 7, kind: item.kind, text: item.text, history: first, expected: second, source: item.source });
  const history = second.slice(); history[2] = first[2];
  engine.seed(s + 13);
  const third = item.kind === "code" ? engine.code(item.text, history) : engine.prose(item.text, history);
  cases.push({ seed: s + 13, kind: item.kind, text: item.text, history, expected: third, source: item.source });
}));

const counts = cases.reduce((acc, c) => { acc[c.kind] = (acc[c.kind] || 0) + 1; return acc; }, {});
console.log(`${texts.length} texts, ${cases.length} cases`, counts);
const dir = fs.mkdtempSync(path.join(os.tmpdir(), "bt1101-variants-"));
const inFile = path.join(dir, "cases.txt");
fs.writeFileSync(inFile, cases.map(c => [c.seed, c.kind, escape(c.text), c.history ? c.history.length : -1, ...(c.history || []).map(escape)].join("\t")).join("\n"), "utf8");
fs.writeFileSync(path.join(ROOT, ".build", "variants-expected.json"), JSON.stringify(cases, null, 1));
if (args.includes("--only-node")) process.exit(0);

const only = "modJs.bas,clsList.cls,clsMap.cls,modRegex.bas,clsRegex.cls,clsMatch.cls,modDom.bas,clsNode.cls,clsEvent.cls,clsDomHash.cls,clsDomSelector.cls,modShared.bas,modVariants.bas,modVariants2.bas,modVariants3.bas,modClipboard.bas,testUtil.bas,testVariants.bas";
const started = Date.now();
const run = spawnSync("node", [path.join(ROOT, "tools", "vba.js"), "run", "testVariants.RunEngineCases", inFile, "--only", only, "--no-form"], { encoding: "utf8", maxBuffer: 1 << 28 });
console.log((run.stdout || "").trim().slice(-400), (run.stderr || "").trim().slice(-400), `(${Math.round((Date.now() - started) / 1000)} s)`);
const outFile = inFile + ".out";
if (!fs.existsSync(outFile)) { console.error("No VBA output"); process.exit(1); }
const lines = fs.readFileSync(outFile, "utf8").split("\n");
let matched = 0, mismatched = 0, shown = 0;
const failingSources = new Map();
cases.forEach((c, i) => {
  const got = (lines[i] || "").split("\t").map(unescape);
  const same = got.length === 6 && got.every((text, k) => text === c.expected[k]);
  if (same) { matched++; return; }
  mismatched++;
  failingSources.set(c.source, (failingSources.get(c.source) || 0) + 1);
  if (shown++ < 4) {
    const k = c.expected.findIndex((text, k) => text !== got[k]);
    console.log(`\n--- case ${i} (${c.kind}, seed ${c.seed}, ${c.source}) option ${k + 1}`);
    if (got.length !== 6) console.log("VBA:", lines[i].slice(0, 600));
    else {
      const a = c.expected[k], b = got[k];
      let p = 0; while (p < a.length && a[p] === b[p]) p++;
      console.log("canonical:", JSON.stringify(c.text.slice(0, 300)));
      console.log("JS :", JSON.stringify(a.slice(Math.max(0, p - 80), p + 160)));
      console.log("VBA:", JSON.stringify(b.slice(Math.max(0, p - 80), p + 160)));
    }
  }
});
console.log(`\nmatched ${matched} / ${cases.length} cases (${mismatched} mismatched)`);
if (failingSources.size) console.log([...failingSources].slice(0, 20));
fs.writeFileSync(path.join(ROOT, ".build", "variants-equiv-result.txt"), `matched ${matched} / ${cases.length}\n`);
