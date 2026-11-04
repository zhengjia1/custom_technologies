// Unit test of the Tutorial 5 distribution functions (vba/opInference2.bas) against the
// website's own JavaScript (excel/.build/app-src.js), over many arguments.
//
//   node test/opInference-numeric.js [--only a.bas,...]
//
// Compares raw results bit for bit (reported), formatted outputs exactly (toFixed(2),
// toFixed(6), formatFinalNumber), and raw values for near-equality (relative 1e-12).

const fs = require("fs");
const os = require("os");
const path = require("path");
const { spawnSync } = require("child_process");

const source = fs.readFileSync(path.join(__dirname, "..", ".build", "app-src.js"), "utf8");
const start = source.indexOf("function inverseNormal(p)");
const end = source.indexOf("function intervalValues()");
if (start < 0 || end < 0) throw new Error("distribution functions not found in app-src.js");
const lib = new Function(source.slice(start, end) + "\nfunction formatFinalNumber(value) { if (!Number.isFinite(value)) return String(value); const absolute = Math.abs(value); if (absolute === 0) return '0'; if (absolute > 1) return value.toFixed(2); return value.toPrecision(3).replace('e+', 'e'); }\nreturn { inverseNormal, logGamma, betaFraction, regularizedBeta, studentTCdf, inverseStudentT, formatFinalNumber, exp: Math.exp };")();

const hex = x => { const b = Buffer.alloc(8); b.writeDoubleBE(x); return b.toString("hex").toUpperCase(); };
const fromHex = h => Buffer.from(h, "hex").readDoubleBE(0);

let seed = 12345;
const rand = () => { seed = (seed * 1103515245 + 12345) % 2147483648; return seed / 2147483648; };
const cases = [];
const add = (fn, ...args) => cases.push({ fn, args });

// inverseNormal: tails, centre, the interval tool's confidence levels, edges
[0, 1, -0.1, 1.1, 0.5, 0.02425, 0.97575, 1e-10, 1 - 1e-10, 0.975, 0.95, 0.995, 0.9, 0.9995].forEach(p => add("inverseNormal", p));
for (let c = 1; c < 100; c += 0.5) add("inverseNormal", 1 - (1 - c / 100) / 2);
for (let i = 0; i < 200; i++) add("inverseNormal", rand());
for (let i = 0; i < 50; i++) add("inverseNormal", rand() * 0.03);
// exp (the fdlibm port used for Math.exp)
[0, 1, -1, 0.5, -0.5, 0.3, 1e-9, -1e-9, 709, 709.7, -700, -745, -746, 50, -50, 2.5, 0.34657].forEach(x => add("exp", x));
for (let i = 0; i < 300; i++) add("exp", (rand() - 0.5) * 80);
for (let i = 0; i < 100; i++) add("exp", (rand() - 0.5) * 2);
// logGamma
[0.5, 1, 1.5, 2, 3, 10, 100, 1e6, 0.75, 5.5, 50.5].forEach(z => add("logGamma", z));
for (let i = 0; i < 200; i++) add("logGamma", 0.5 + rand() * 300);
for (let i = 0; i < 20; i++) add("logGamma", 0.05 + rand() * 0.44);
// betaFraction / regularizedBeta
for (let i = 0; i < 150; i++) { const a = 0.5 + rand() * 60, b = 0.5 + rand() * 3, x = rand(); add("betaFraction", a, b, x); add("regularizedBeta", x, a, b); }
[[0, 2, 0.5], [1, 2, 0.5], [-0.5, 2, 0.5], [1.5, 2, 0.5], [0.5, 0.5, 0.5], [0.999, 30, 0.5]].forEach(args => add("regularizedBeta", ...args));
// studentTCdf
for (let i = 0; i < 200; i++) add("studentTCdf", (rand() - 0.5) * 12, 1 + Math.floor(rand() * 200));
[[0, 1], [2.5, 5], [-2.5, 5], [100, 3], [-100, 3], [1.96, 1e6]].forEach(args => add("studentTCdf", ...args));
// inverseStudentT: the confidence levels the interval tool uses, many df
const levels = [80, 90, 95, 98, 99, 99.9, 50, 97.5, 92.5, 60];
for (const c of levels) for (const n of [2, 3, 5, 10, 11, 25, 30, 64, 100, 1000]) add("inverseStudentT", 1 - (1 - c / 100) / 2, n - 1);
for (let i = 0; i < 60; i++) add("inverseStudentT", rand(), 1 + Math.floor(rand() * 100));
[[0, 5], [1, 5], [0.5, 5], [0.3, 0], [0.975, -1], [NaN, 5]].forEach(args => add("inverseStudentT", ...args));

const lines = cases.map(c => [c.fn === "regularizedBeta" || c.fn === "betaFraction" ? c.fn : c.fn, ...c.args.map(hex)].join("\t"));
const dir = fs.mkdtempSync(path.join(os.tmpdir(), "bt1101-inference-numeric-"));
const argFile = path.join(dir, "cases.txt");
fs.writeFileSync(argFile, lines.join("\n"));
const variants = fs.readdirSync(path.join(__dirname, "..", "vba")).filter(name => /^modVariants\d*\.bas$/.test(name)).join(",");
const only = process.argv.includes("--only") ? process.argv[process.argv.indexOf("--only") + 1] :
  "modJs.bas,clsList.cls,clsMap.cls,modRegex.bas,clsRegex.cls,clsMatch.cls,modDom.bas,clsNode.cls,clsEvent.cls,clsDomHash.cls,clsDomSelector.cls,modShared.bas,modClipboard.bas,testUtil.bas,testDriver.bas,opInference.bas,opInference2.bas,testOpInference.bas," + variants;
const run = spawnSync(process.execPath, [path.join(__dirname, "..", "tools", "vba.js"), "run", "testOpInference.NumericCases", "--argfile", argFile, "--only", only, "--no-form"], { encoding: "utf8", maxBuffer: 64 * 1024 * 1024 });
fs.rmSync(dir, { recursive: true, force: true });
const output = (run.stdout || "").replace(/\r/g, "").split("\n").filter(line => /^[0-9A-F]{16}\t|^ERROR/.test(line));
if (output.length !== cases.length) { console.log((run.stdout || "") + (run.stderr || "")); throw new Error(`expected ${cases.length} results, got ${output.length}`); }

const counts = {};
let formatFailures = 0, nearFailures = 0, bitMismatches = 0;
cases.forEach((c, i) => {
  const expected = lib[c.fn](...c.args);
  const [h, f2, f6, ffn] = output[i].split("\t");
  const actual = fromHex(h || "7FF8000000000001");
  const row = counts[c.fn] || (counts[c.fn] = { total: 0, bits: 0, formatted: 0, near: 0 });
  row.total++;
  const sameBits = hex(expected) === h || (Number.isNaN(expected) && Number.isNaN(actual));
  if (sameBits) row.bits++; else bitMismatches++;
  const formattedOk = output[i].startsWith("ERROR") ? false : f2 === expected.toFixed(2) && f6 === expected.toFixed(6) && ffn === lib.formatFinalNumber(expected);
  if (formattedOk) row.formatted++; else { formatFailures++; console.log(`FORMAT ${c.fn}(${c.args.join(", ")}): js ${expected} -> ${expected.toFixed(2)}|${expected.toFixed(6)}|${lib.formatFinalNumber(expected)}  vba ${output[i]}`); }
  const near = sameBits || (Number.isFinite(expected) && Number.isFinite(actual) && Math.abs(actual - expected) <= 1e-12 * Math.max(1, Math.abs(expected)));
  if (near) row.near++; else { nearFailures++; console.log(`VALUE ${c.fn}(${c.args.join(", ")}): js ${expected} vba ${actual}`); }
  if (!sameBits && near && process.argv.includes("--verbose")) console.log(`BITS ${c.fn}(${c.args.join(", ")}): js ${hex(expected)} vba ${h}`);
});
for (const [fn, row] of Object.entries(counts)) console.log(`${fn.padEnd(16)} ${row.total} cases: ${row.bits} bit-identical, ${row.formatted} formatted-identical, ${row.near} near-equal`);
console.log(`${cases.length} cases: ${cases.length - bitMismatches} bit-identical, ${formatFailures} formatting differences, ${nearFailures} value differences`);
process.exitCode = formatFailures || nearFailures ? 1 : 0;
