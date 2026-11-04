// Differential tests for vba/modJs.bas: generates cases with Node's (V8's) results and,
// with --run, executes them in Excel through testJs.RunCases and prints the report.
//
//   node test/js-cases.js [--out file] [--seed n] [--run]
//
// File format (UTF-8): one case per line, fields separated by TAB:
//   TAG <TAB> arg1 <TAB> ... <TAB> expected
// Doubles are written as "#" + 16 hex digits of their IEEE bits ("#NaN" for NaN);
// "u" marks an omitted (undefined) optional argument. Inside fields, control characters,
// backslash and lone surrogates are written as \uXXXX; everything else is raw UTF-8.

const fs = require("fs");
const os = require("os");
const path = require("path");
const { spawnSync } = require("child_process");

const argv = process.argv.slice(2);
const opt = (name, dflt) => { const i = argv.indexOf(name); return i >= 0 ? argv[i + 1] : dflt; };
const OUT = path.resolve(opt("--out", path.join(os.tmpdir(), "bt1101-js-cases.txt")));
let seed = Number(opt("--seed", 20261103)) >>> 0;

function mulberry32(a) {
  return function () {
    a |= 0; a = a + 0x6D2B79F5 | 0;
    let t = Math.imul(a ^ a >>> 15, 1 | a);
    t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t;
    return ((t ^ t >>> 14) >>> 0) / 4294967296;
  };
}
const rnd = mulberry32(seed);
const randInt = (lo, hi) => lo + Math.floor(rnd() * (hi - lo + 1));
const pick = arr => arr[Math.floor(rnd() * arr.length)];

const dv = new DataView(new ArrayBuffer(8));
function hex(x) {
  if (Number.isNaN(x)) return "#NaN";
  dv.setFloat64(0, x);
  return "#" + dv.getUint32(0).toString(16).toUpperCase().padStart(8, "0") + dv.getUint32(4).toString(16).toUpperCase().padStart(8, "0");
}
function fromBits(hi, lo) { dv.setUint32(0, hi >>> 0); dv.setUint32(4, lo >>> 0); return dv.getFloat64(0); }
function esc(s) {
  s = String(s);
  let out = "";
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    const hiS = c >= 0xd800 && c <= 0xdbff, loS = c >= 0xdc00 && c <= 0xdfff;
    let lone = false;
    if (hiS) { const n = s.charCodeAt(i + 1); if (n >= 0xdc00 && n <= 0xdfff) { out += s[i] + s[i + 1]; i++; continue; } lone = true; }
    if (loS) lone = true;
    if (c < 0x20 || c === 0x5c || c === 0x7f || lone) out += "\\u" + c.toString(16).toUpperCase().padStart(4, "0");
    else out += s[i];
  }
  return out;
}

const lines = [];
const add = (...fields) => lines.push(fields.map(esc).join("\t"));
const tryJs = fn => { try { return fn(); } catch (e) { return "!" + e.name; } };

// ---------------------------------------------------------------- doubles
const specials = [0, -0, NaN, Infinity, -Infinity, 5e-324, -5e-324, 1e-323, 2.2250738585072014e-308,
  2.225073858507201e-308, 1.7976931348623157e308, -1.7976931348623157e308, 1e21, 999999999999999900000,
  1e21 - 65536 * 2, 1e-7, 1e-6, 0.000001, 1.5e-7, 123e-20, 2 ** 53, 2 ** 53 + 2, 2 ** 53 - 1, 2 ** 52 + 0.5,
  2 ** 63, 2 ** 64, 2 ** 70, 1e15, 1e16, 1e17, 1e20, 0.1, 0.2, 0.3, 0.1 + 0.2, 1 / 3, 2 / 3, 1.005, 2.675, 0.5,
  1.45, 8.345, 1.255, 1.335, 0.045, 0.615, 10.235, 1.0049999999999999, 4.35, 0.0005, 0.00015, 0.000015, 9.995,
  99.995, 0.995, 0.9995, 1e-10, 123456789.123456789, 4503599627370495.5, 4503599627370497, 0.49999999999999994,
  -0.5, -1.5, 2.5, -2.5, 1.5, 3.5, -3.5, 0.05, 0.25, 0.125, 0.0625, 1e100, 1e-100, 1e300, 1e-300, 123.456,
  -123.456, 1e22, 1e23, 9.999999999999999e22, 5e-7, 1.25e-7, 0.00001, 4.9406564584124654e-324,
  2 ** -1022, 2 ** -1023, 2 ** -1074 * 3, 0.1 * 3, 100, 1000, 12345, 2147483647, 2147483648, 4294967295,
  4294967296, -2147483648, 1.7976931348623155e308, 5.551115123125783e-17, 0.30000000000000004];
const doubles = [...specials];
const halfwayDigits = new Map();
for (let i = 0; i < 700; i++) { // random bit patterns (all magnitudes)
  const hi = Math.floor(rnd() * 4294967296), lo = Math.floor(rnd() * 4294967296);
  const x = fromBits(hi, lo);
  if (!Number.isNaN(x)) doubles.push(x);
}
for (let i = 0; i < 1000; i++) { // short decimals, like data values
  const digits = randInt(1, 9), scale = randInt(0, 8);
  let n = Math.floor(rnd() * 10 ** digits);
  const x = (rnd() < 0.3 ? -1 : 1) * Number(n + "e-" + scale);
  doubles.push(x);
}
for (let i = 0; i < 800; i++) { // halfway decimals: n.dddd5 for toFixed(d)
  const d = randInt(0, 6), n = Math.floor(rnd() * 10 ** randInt(1, 7));
  const x = (rnd() < 0.3 ? -1 : 1) * Number((n * 10 + 5) + "e-" + (d + 1));
  doubles.push(x);
  halfwayDigits.set(doubles.length - 1, d);
}
for (let i = 0; i < 300; i++) doubles.push((rnd() - 0.5) * 10 ** randInt(-8, 12));
for (let i = 0; i < 150; i++) doubles.push(2 ** 53 + randInt(-2000, 2000) * (rnd() < 0.5 ? 1 : 2));
for (let i = 0; i < 100; i++) doubles.push(10 ** randInt(-25, 25) * (rnd() < 0.5 ? 1 : -1));
for (let i = 0; i < 100; i++) doubles.push(randInt(-1e6, 1e6) + pick([0, 0.5, -0.5, 0.25, 0.49999999999999994]));

doubles.forEach((x, idx) => {
  const h = hex(x);
  add("S", h, String(x));
  const fd = [randInt(0, 20), randInt(0, 4)];
  if (halfwayDigits.has(idx)) fd.push(halfwayDigits.get(idx));
  for (const d of fd) add("F", h, String(d), tryJs(() => x.toFixed(d)));
  const p = randInt(1, 21);
  add("P", h, String(p), tryJs(() => x.toPrecision(p)));
  if (rnd() < 0.25) add("P", h, "u", x.toPrecision());
  if (rnd() < 0.25) add("E", h, "u", x.toExponential());
  else { const f = randInt(0, 20); add("E", h, String(f), tryJs(() => x.toExponential(f))); }
  const mn = randInt(0, 5), mx = mn + randInt(0, 5), g = rnd() < 0.7;
  add("L", h, String(mn), String(mx), g ? "1" : "0", x.toLocaleString("en-US", { minimumFractionDigits: mn, maximumFractionDigits: mx, useGrouping: g }));
  if (rnd() < 0.2) add("L", h, "u", "u", "u", x.toLocaleString("en-US"));
  if (rnd() < 0.15 && halfwayDigits.has(idx)) { const d = halfwayDigits.get(idx); add("L", h, String(d), String(d), "1", x.toLocaleString("en-US", { minimumFractionDigits: d, maximumFractionDigits: d })); }
  add("R", h, hex(Math.round(x)));
});
for (const x of [5e-324, 1.7976931348623157e308, 0.1, 1 / 3, 1e-300, 123.456, -0, 0, 1e20, 2 ** -1074 * 12345]) {
  for (const d of [100, 60]) add("F", hex(x), String(d), tryJs(() => x.toFixed(d)));
  add("P", hex(x), "100", tryJs(() => x.toPrecision(100)));
  add("E", hex(x), "100", tryJs(() => x.toExponential(100)));
}
for (const [x, d] of [[1, 101], [1, -1], [NaN, 101], [Infinity, 5]]) add("F", hex(x), String(d), tryJs(() => x.toFixed(d)));
for (const [x, d] of [[1, 0], [1, 101], [NaN, 0], [Infinity, 200]]) add("P", hex(x), String(d), tryJs(() => x.toPrecision(d)));
for (const [x, d] of [[1, 101], [NaN, -5], [-Infinity, 200]]) add("E", hex(x), String(d), tryJs(() => x.toExponential(d)));

// ---------------------------------------------------------------- string -> number
function decompose(x) {
  dv.setFloat64(0, x);
  const hi = dv.getUint32(0), lo = dv.getUint32(4), ex = (hi >>> 20) & 0x7ff;
  let m = BigInt(hi & 0xfffff) * 2n ** 32n + BigInt(lo), e;
  if (ex === 0) e = -1074; else { m += 2n ** 52n; e = ex - 1075; }
  return { m, e };
}
function midpointStrings(x) { // exact decimal midpoint between x and the next double, and neighbours
  const { m, e } = decompose(x);
  const N = 2n * m + 1n, E = e - 1;
  if (E >= 0) { const v = N * 2n ** BigInt(E); return [v.toString(), v.toString() + ".0000000000001", (v - 1n).toString() + ".9999999999"]; }
  const k = -E, digits = (N * 5n ** BigInt(k)).toString();
  const below = digits.slice(0, -1) + String(Number(digits.slice(-1)) - 1);
  return [digits + "e-" + k, digits + "0001e-" + (k + 4), below + "e-" + k];
}
const numStrings = [];
const ns = s => numStrings.push(s);
doubles.slice(0, 2500).forEach(x => {
  if (!Number.isFinite(x)) return;
  ns(String(x));
  if (rnd() < 0.3) ns(x.toExponential());
  if (rnd() < 0.3) ns(x.toPrecision(17));
  if (rnd() < 0.2) ns(x.toPrecision(randInt(1, 25)));
  if (rnd() < 0.2 && Math.abs(x) < 1e21) ns(x.toFixed(randInt(0, 25)));
});
for (let i = 0; i < 500; i++) {
  let x = doubles[randInt(0, doubles.length - 1)];
  if (!Number.isFinite(x) || x === 0) continue;
  x = Math.abs(x);
  midpointStrings(x).forEach(ns);
}
midpointStrings(1.7976931348623157e308).forEach(ns);
midpointStrings(5e-324).forEach(ns);
midpointStrings(2 ** -1022).forEach(ns);
midpointStrings(2 ** 53).forEach(ns);
for (let i = 0; i < 300; i++) { // random long digit strings
  let s = String(randInt(1, 9));
  const len = randInt(1, 40);
  for (let j = 0; j < len; j++) s += randInt(0, 9);
  if (rnd() < 0.5) s = s.slice(0, randInt(1, s.length)) + "." + s.slice(randInt(1, s.length));
  if (rnd() < 0.6) s += "e" + pick(["", "+", "-"]) + randInt(0, 340);
  ns(s);
}
["2.2250738585072011e-308", "2.2250738585072012e-308", "4.9406564584124654e-324", "2.4703282292062327e-324",
  "2.4703282292062328e-324", "1.7976931348623157e308", "1.7976931348623158e308", "1.7976931348623159e308",
  "9007199254740993", "9007199254740995", "9007199254740993.0000000001", "0.1", "0.3", "1e23", "8.98846567431158e307",
  "1e-324", "3e-324", "2.5e-324", "1e309", "1e-400", "0e999999999", "123456789012345678901234567890",
  "0." + "0".repeat(320) + "247032822920623272088284396434110686182529901307162382212792841250337753635104375932649918180817996189898282347722858865463328355177969898199387398005390939063150356595155702263922908583924491051844359318028499365361525003193704576782492193656236698636584807570015857692699037063119282795585513329278343384093519780155312465972635795746227664652728272200563740064854999770965994704540208281662262378573934507363390079677619305775067401763246736009689513405355374585166611342237666786041621596804619144672918403005300575308490487653917113865916462395249126236538818796362393732804238910186723484976682350898633885879256283027559956575244555072551893136908362547791869486679949683240497058210285131854513962138377228261454376934125320985913276672363281255",
  "1" + "0".repeat(400), "0." + "0".repeat(400) + "1", "1" + "0".repeat(308) + ".5", "4503599627370496.5", "4503599627370497.5",
  "1.00000000000000011102230246251565404236316680908203125", "1.00000000000000011102230246251565404236316680908203124",
  "1.00000000000000011102230246251565404236316680908203126"].forEach(ns);
// syntax variants
const wsChars = [" ", "\t", "\n", "\r", "\v", "\f", " ", "﻿", " ", " ", " ", " ", " ", " ", " ", "　"];
const notWs = ["᠎", "​", "\u0085"];
["", " ", "  \t ", "1e", "1e+", "1e-", ".", "+.5", "-.5", "-.5e-3", ".e5", "e5", "1.e5", "1e5.5", "1.", "-1.", "--1", "+-1",
  "1 2", "1_000", "1,5", "12px", "0x", "0x1g", "0X1F", "0x1f", "-0x1f", "+0x1f", "0xFFFFFFFFFFFFFFFF", "0x1fffffffffffff",
  "0x20000000000001", "0x20000000000003", "0x2000000000000100000001", "0x" + "f".repeat(260), "0x" + "8".repeat(256) + "7",
  "0b101", "0b2", "0B11", "0o17", "0O8", "0o", "0b", "Infinity", "-Infinity", "+Infinity", "infinity", "INFINITY", "Infinityx",
  "NaN", "-0", "+0", "-0.0e5", "00012", "-00012.500", "0.0000", "1e1000", "-1e1000", "1e-1000", "١٢", "１",
  "1e+00000000000000000000000002", "0.000000000000000000000000000000000000000000000000001e51", "12abc", "abc", "1.2.3",
  "  +Infinityxyz", "-.e1", "0.5e", "5e-", "1d2", "&H10", "1e2e3"].forEach(ns);
for (let i = 0; i < 200; i++) {
  const core = pick(["42", "-3.5", "1e3", "0x1A", "Infinity", "", ".5", "abc", "1 2", "0b11", "-0"]);
  const pre = Array.from({ length: randInt(0, 2) }, () => pick(rnd() < 0.85 ? wsChars : notWs)).join("");
  const post = Array.from({ length: randInt(0, 2) }, () => pick(rnd() < 0.85 ? wsChars : notWs)).join("");
  ns(pre + core + post);
}
for (let i = 0; i < 100; i++) { // long radix literals
  const r = pick([[16, "0x", "0123456789abcdefABCDEF"], [2, "0b", "01"], [8, "0o", "01234567"]]);
  let s = r[2][randInt(1, r[2].length - 1)];
  const len = randInt(1, r[0] === 2 ? 80 : 30);
  for (let j = 0; j < len; j++) s += pick(r[2].split(""));
  ns(r[1] + s);
}
// Clinger fast-path inputs (<= 15 digits, |q| <= 22) where rounding the exact result to
// 64 bits first (x87 extended precision, as VBA does) and then to 53 bits gives the wrong
// double.
function round64(m) { // m: positive BigInt -> nearest BigInt with <= 64 significant bits (ties to even)
  const len = m.toString(2).length;
  if (len <= 64) return m;
  const shift = BigInt(len - 64), low = m & ((1n << shift) - 1n), half = 1n << (shift - 1n);
  let top = m >> shift;
  if (low > half || (low === half && (top & 1n))) top += 1n;
  return top << shift;
}
let found = 0;
for (let i = 0; i < 400000 && found < 150; i++) {
  const v = BigInt(randInt(1, 999999999)) * BigInt(randInt(1, 999999)) % 1000000000000000n || 1n, q = randInt(1, 22);
  let s, correct, dr;
  if (rnd() < 0.5) {
    const P = v * 10n ** BigInt(q);
    correct = Number(P); dr = Number(round64(P)); s = v + "e" + q;
  } else {
    const K = 200n, num = (v << K), den = 10n ** BigInt(q), quo = num / den, rem = num % den;
    const m = round64(quo * 2n + (rem ? 1n : 0n)); // sticky bit keeps ties honest
    dr = Number(m) / 2 ** 201; s = v + "e-" + q; correct = Number(s);
  }
  if (dr !== correct) { ns(s); found++; }
}
const junk = ["", "px", "abc", "e", "e+", "e5x", ".5", " 1", "%", "x", ".", "..", " "];
numStrings.forEach(s => {
  add("N", s, hex(Number(s)));
  if (rnd() < 0.5) { const t = s + pick(junk); add("PF", t, hex(parseFloat(t))); }
  else add("PF", s, hex(parseFloat(s)));
});
const intStrings = ["0x1f", "0X1F", "  -42abc", "+17", "1e3", "123456789012345678901234567890", "-0", "", "  ", "z", "Z", "zz",
  "0x", "0b101", "  0012", "9007199254740993", "9007199254740995", "1" + "0".repeat(320), "123.9", "-123.9", "  7",
  "ffffffffffffffffff", "777777777777777777777", "1010101010101010101010101010101010101010101010101010101111", "0.000001",
  "deadbeef", "-deadbeef", "Infinity", "NaN", "12_3", "１"];
for (let i = 0; i < 60; i++) { let s = String(randInt(1, 9)); for (let j = 0; j < randInt(0, 30); j++) s += randInt(0, 9); intStrings.push(s); }
for (const s of intStrings) for (const r of ["u", "10", "16", "2", "8", "36", "0", "1", "37", "32", "4", "3", "-1", "16.9"]) {
  const v = r === "u" ? parseInt(s) : parseInt(s, Number(r));
  if (Number.isFinite(v) && Math.abs(v) >= 2 ** 53 && !["u", "10", "16", "2", "8", "32", "4"].includes(r)) continue; // non-power-of-2 radix may be approximate
  add("PI", s, r, hex(v));
}

// ---------------------------------------------------------------- math helpers
const sv = [0, -0, 1, -1, 0.5, -0.5, 2, 3, -3, 1.5, -2.5, 7, 0.1, NaN, Infinity, -Infinity, 1e308, -1e308, 5e-324, 1e-300, 4.5, -7.25];
for (const a of sv) for (const b of sv) {
  add("ADD", hex(a), hex(b), hex(a + b));
  add("SUB", hex(a), hex(b), hex(a - b));
  add("MUL", hex(a), hex(b), hex(a * b));
  add("DIV", hex(a), hex(b), hex(a / b));
  add("MOD", hex(a), hex(b), hex(a % b));
  add("MAX", hex(a), hex(b), hex(Math.max(a, b)));
  add("MIN", hex(a), hex(b), hex(Math.min(a, b)));
  add("LT", hex(a), hex(b), String(a < b));
  add("EQ", hex(a), hex(b), String(a === b));
  add("GE", hex(a), hex(b), String(a >= b));
  const p = a ** b;
  if (!Number.isFinite(p) || p === 0 || [0, -0, 1, -1, 2, 0.5].includes(b) || [0, -0, 1, -1].includes(a)) add("POW", hex(a), hex(b), hex(p));
  else add("POWR", hex(a), hex(b), hex(p));
}
for (let i = 0; i < 300; i++) {
  const a = (rnd() - 0.5) * 10 ** randInt(-5, 20), b = (rnd() - 0.5) * 10 ** randInt(-5, 10);
  add("MOD", hex(a), hex(b), hex(a % b));
}
add("MOD", hex(1e308), hex(3), hex(1e308 % 3));
add("MOD", hex(-1e300), hex(7e-300), hex(-1e300 % 7e-300));
add("MOD", hex(5.5), hex(2), hex(5.5 % 2));
for (const x of [...sv, 1e21, 2 ** 31, 2 ** 32 + 5, -(2 ** 31) - 1, 4294967295.7, -4294967296.5, 2 ** 84 + 2 ** 33]) {
  add("SQRT", hex(x), hex(Math.sqrt(x)));
  add("FLOOR", hex(x), hex(Math.floor(x)));
  add("CEIL", hex(x), hex(Math.ceil(x)));
  add("TRUNC", hex(x), hex(Math.trunc(x)));
  add("SIGN", hex(x), hex(Math.sign(x)));
  add("ABS", hex(x), hex(Math.abs(x)));
  add("I32", hex(x), String(x | 0));
  add("U32", hex(x), hex(x >>> 0));
  if (!Number.isFinite(Math.log(x)) || x === 1) add("LOG", hex(x), hex(Math.log(x)));
  if (!Number.isFinite(Math.exp(x)) || x === 0) add("EXP", hex(x), hex(Math.exp(x)));
}
for (let k = -330; k <= 320; k++) add("POW", hex(10), hex(k), hex(10 ** k));
for (let k = -30; k <= 30; k++) { add("POW", hex(2), hex(k * 37), hex(2 ** (k * 37))); }
for (let i = 0; i < 400; i++) { // transcendental functions: V8 uses fdlibm ports
  const x = rnd() * 10 ** randInt(-6, 6);
  add("LOGR", hex(x), hex(Math.log(x)));
  const y = (rnd() - 0.5) * 10 ** randInt(-4, 2.85);
  add("EXPR", hex(y), hex(Math.exp(y)));
  const z = (rnd() - 0.5) * 10 ** randInt(-3, 3);
  add("SINR", hex(z), hex(Math.sin(z)));
  add("COSR", hex(z), hex(Math.cos(z)));
  add("SQRT", hex(x), hex(Math.sqrt(x)));
  const a = rnd() * 10 ** randInt(-3, 3), b = (rnd() - 0.5) * 10 ** randInt(-2, 2);
  add("POWR", hex(a), hex(b), hex(a ** b));
}
for (const x of [Math.PI / 4, Math.PI / 2, -Math.PI / 2, Math.PI, 3 * Math.PI / 4, 2 * Math.PI, 100 * Math.PI, 1e5 * Math.PI / 2, 1, 0.5, 1e-9, -1e-9, 0, -0, 0.3, 0.78125, 0.7853981633974483, 10, 700, 709.78, -745.1, 710, -746, 1e-300, 5e-324]) {
  add("SINR", hex(x), hex(Math.sin(x))); add("COSR", hex(x), hex(Math.cos(x))); add("EXPR", hex(x), hex(Math.exp(x)));
  if (x > 0) add("LOGR", hex(x), hex(Math.log(x)));
}
for (let k = -22; k <= 22; k++) add("LOG10", hex(Number("1e" + k)), hex(Math.log10(Number("1e" + k))));
for (let k = -1074; k <= 1023; k += 37) add("LOG2", hex(2 ** k), hex(Math.log2(2 ** k)));
for (const x of [NaN, Infinity, 0, -1, 1e-300, 1234.5, 5e-324, 0.3, 2, 1e300]) { add("LOG10", hex(x), hex(Math.log10(x))); }
for (let i = 0; i < 300; i++) { const x = rnd() * 10 ** randInt(-8, 8); add("LOG10", hex(x), hex(Math.log10(x))); }
add("NAN", hex(NaN), "true");
for (const x of [0, 1, Infinity, -Infinity, -0, 5e-324]) add("NAN", hex(x), "false");
for (const x of [0, 1, NaN, Infinity, -Infinity, -0, 1.7976931348623157e308]) add("FIN", hex(x), String(Number.isFinite(x)));
for (const x of [0, 1, NaN, Infinity, -0, 1.5, 2 ** 60, -3]) add("INT", hex(x), String(Number.isInteger(x)));
for (const s of [0, 1, 42, 123456789, 4294967295, 2147483648, 7]) {
  const r = mulberry32(s), vals = [];
  for (let i = 0; i < 8; i++) vals.push(hex(r()));
  add("RAND", String(s), vals.join(","));
}

// ---------------------------------------------------------------- strings
const alphabet = ["a", "b", "c", "A", "B", "Z", "x", "y", "z", " ", ",", ";", ".", "-", "_", "0", "9", "ab", "ba", "\t", "\n",
  " ", "﻿", " ", "　", "᠎", "​", "é", "É", "μ", "₀", "😀", "\\", "\"", "'"];
const randStr = (min, max) => Array.from({ length: randInt(min, max) }, () => pick(alphabet)).join("");
const idx = n => randInt(-n - 3, n + 3);
const optIdx = n => (rnd() < 0.2 ? "u" : String(idx(n)));
const call = (fn, ...a) => fn(...a);
for (let i = 0; i < 400; i++) {
  const s = randStr(0, 14), n = s.length;
  add("TRIM", s, s.trim());
  add("TRIMS", s, s.trimStart());
  add("TRIME", s, s.trimEnd());
  let a = String(idx(n)), b = optIdx(n);
  add("SLICE", s, a, b, b === "u" ? s.slice(Number(a)) : s.slice(Number(a), Number(b)));
  a = String(idx(n)); b = optIdx(n);
  add("SUBSTRING", s, a, b, b === "u" ? s.substring(Number(a)) : s.substring(Number(a), Number(b)));
  const search = rnd() < 0.6 && n > 0 ? s.substr(randInt(0, n - 1), randInt(0, 3)) : randStr(0, 2);
  let f = optIdx(n);
  add("INDEXOF", s, search, f, String(f === "u" ? s.indexOf(search) : s.indexOf(search, Number(f))));
  f = optIdx(n);
  add("LASTINDEXOF", s, search, f, String(f === "u" ? s.lastIndexOf(search) : s.lastIndexOf(search, Number(f))));
  f = optIdx(n);
  add("INCLUDES", s, search, f, String(f === "u" ? s.includes(search) : s.includes(search, Number(f))));
  f = optIdx(n);
  add("STARTS", s, search, f, String(f === "u" ? s.startsWith(search) : s.startsWith(search, Number(f))));
  f = optIdx(n);
  add("ENDS", s, search, f, String(f === "u" ? s.endsWith(search) : s.endsWith(search, Number(f))));
  const len = randInt(-1, n + 6), pad = rnd() < 0.3 ? "u" : pick(["", "0", "ab", " ", "xyz"]);
  add("PADS", s, String(len), pad, pad === "u" ? s.padStart(len) : s.padStart(len, pad));
  add("PADE", s, String(len), pad, pad === "u" ? s.padEnd(len) : s.padEnd(len, pad));
  const cnt = randInt(-1, 4);
  add("REPEAT", s, String(cnt), tryJs(() => s.repeat(cnt)));
  const ci = randInt(-2, n + 1);
  add("CHARAT", s, String(ci), s.charAt(ci));
  add("CHARCODEAT", s, String(ci), hex(s.charCodeAt(ci)));
  add("AT", s, String(ci), s.at(ci) === undefined ? "" : s.at(ci));
  const sep = pick(["u", "", ",", " ", "ab", "a", ";", "\t", "zz"]), lim = pick(["u", "u", "0", "1", "2", "5", "-1"]);
  const parts = sep === "u" ? (lim === "u" ? s.split() : s.split(undefined, Number(lim))) : (lim === "u" ? s.split(sep) : s.split(sep, Number(lim)));
  add("SPLIT", s, sep, lim, parts.length + ":" + parts.join("\x1f"));
  const ra = rnd() < 0.7 && n > 0 ? s.substr(randInt(0, n - 1), randInt(1, 2)) : pick(["", "q"]);
  const rb = pick(["", "-", "$&", "XY"]);
  if (!rb.includes("$")) add("REPLACEALL", s, ra, rb, s.replaceAll(ra, rb));
}
const caseStrings = ["Hello World", "ABCxyz09_-", "", "éÉàÀüÜñÑ", "αβγΑΒ", "mIxEd CaSe 123", "ça va", "Zz@[`{"];
for (const s of caseStrings) { add("LOWER", s, s.toLowerCase()); add("UPPER", s, s.toUpperCase()); }

const text = lines.join("\n") + "\n";
fs.writeFileSync(OUT, text, "utf8");
const counts = {};
lines.forEach(l => { const t = l.slice(0, l.indexOf("\t")); counts[t] = (counts[t] || 0) + 1; });
console.log(`Wrote ${lines.length} cases to ${OUT}`);
console.log(Object.entries(counts).map(([k, v]) => `${k}=${v}`).join(" "));

if (argv.includes("--run")) {
  const tools = path.join(__dirname, "..", "tools", "vba.js");
  const r = spawnSync(process.execPath, [tools, "run", "testJs.RunCases", OUT, "--only", "modJs.bas,clsList.cls,clsMap.cls,testJs.bas", "--no-form"], { cwd: path.join(__dirname, ".."), encoding: "utf8", maxBuffer: 64 * 1024 * 1024 });
  process.stdout.write(r.stdout || "");
  process.stderr.write(r.stderr || "");
  process.exitCode = r.status;
}
