// Differential test cases for the VBA regular-expression engine (vba/modRegex.bas).
//
//   node test/regex-cases.js [out-file]        (run from excel/; default .build/regex-cases.txt)
//   node tools/vba.js run testRegex.RunRegexCases <absolute out-file> --only modRegex.bas,clsRegex.cls,clsMatch.cls,testRegex.bas --no-form
//
// Every regex literal of the website's main script (plus its `new RegExp` patterns and a set
// of hand-written edge-case patterns) is paired with subject strings: the website's own string
// literals (samples, R outputs, prose), their individual lines, and targeted strings. Node's
// results for exec (with group indices), exec from the middle, matchAll, split and replace are
// recorded in a UTF-8 file that testRegex.bas replays in Excel.
//
// File format (one record per line, fields separated by TAB, text fields escaped: "\\" for a
// backslash and "\uXXXX" for any character below U+0020):
//   S <id> <subject>
//   P <id> <flags> <pattern>                                   flags include "-" when empty
//   C <pattern id> <subject id> <exec> <exec from middle> <matchAll> <split>
//   R <pattern id> <subject id> <template> <result>
// exec: "-" for null, else "start,end;start,end;..." for groups 0..n (-1,-1 = undefined).
// matchAll: exec encodings joined by "|". split: items joined by U+001F, U+001E = undefined.

const fs = require("fs");
const path = require("path");

const EXCEL = path.join(__dirname, "..");
const ROOT = path.join(EXCEL, "..");
const { extractWebsite } = require(path.join(ROOT, "pdf", "src", "extract.js"));
const acorn = require(path.join(ROOT, "pdf", "node_modules", "acorn"));

const OUT = path.resolve(process.argv[2] || path.join(EXCEL, ".build", "regex-cases.txt"));
const MAX_MATCHING = 24, MAX_NONMATCHING = 8, MAX_SUBJECT = 60000;

// ---------------------------------------------------------------- website patterns
const site = extractWebsite(path.join(ROOT, "index.html"));
const code = site.scripts[site.scripts.length - 1];
const tokens = [...acorn.tokenizer(code, { ecmaVersion: "latest" })];

const patterns = new Map(); // key -> { source, flags, templates:Set, custom }
function addPattern(source, flags, extra = {}) {
  const key = flags + "/" + source;
  if (!patterns.has(key)) patterns.set(key, { source, flags, templates: new Set(), subjects: null, ...extra });
  return patterns.get(key);
}

let literalCount = 0;
for (let i = 0; i < tokens.length; i++) {
  const t = tokens[i];
  if (t.type.label !== "regexp") continue;
  literalCount++;
  const entry = addPattern(t.value.pattern, t.value.flags);
  const next = tokens[i + 1], arg = tokens[i + 2];
  if (next && next.type.label === "," && arg && (arg.type.label === "string" || arg.type.label === "template") && typeof arg.value === "string") {
    entry.templates.add(arg.value);
  }
}

// The website's `new RegExp(...)` calls, rebuilt with their actual interpolated values.
const NUMBER = "[-+]?(?:\\d+\\.?\\d*|\\.\\d+)(?:[eE][-+]?\\d+)?";
const NUMBER_PLAIN = "[-+]?(?:\\d+\\.?\\d*|\\.\\d+)";
const NUMBER_GROUP = "([+-]?(?:\\d+(?:\\.\\d*)?|\\.\\d+)(?:e[+-]?\\d+)?)";
for (const modelNumber of [1, 2]) {
  addPattern(`^\\s*Model\\s*${modelNumber}\\s*:`, "i");
  addPattern(`^\\s*Model\\s*${modelNumber}\\s*:\\s*`, "i").templates.add("");
}
addPattern(`^\\s*(.+?)\\s+(${NUMBER})\\s+(${NUMBER})\\s+(${NUMBER})\\s+(<=|>=|<|>)?\\s*(${NUMBER})(?:\\s+\\S+)?\\s*$`, "i");
addPattern(`Multiple\\s+R-squared\\s*:\\s*(${NUMBER})`, "i");
addPattern(`F-statistic\\s*:\\s*(${NUMBER})\\s+on\\s+(${NUMBER})\\s+and\\s+(${NUMBER})\\s+DF\\s*,?\\s*p-value\\s*:\\s*(<=|>=|<|>|=)?\\s*(${NUMBER})`, "i");
addPattern(`^\\s*(.+?)\\s+(${NUMBER})\\s+(${NUMBER})\\s*$`, "i");
addPattern(`^\\s*(\\S+)\\s+(${NUMBER_PLAIN})\\s+(${NUMBER_PLAIN})\\s*$`, "");
for (const name of ["Sensitivity", "Specificity", "Accuracy"]) addPattern(`${name}\\s{0,3}:\\s{0,3}(${NUMBER_PLAIN})`, "i");
addPattern("p(?:-|\\s{0,3})value\\s{0,3}(<=|>=|<|>|=)\\s{0,3}" + NUMBER_GROUP, "i");
addPattern("sample\\s{0,3}estimates\\s{0,3}:\\s{0,3}(?:cor\\s{0,3})?" + NUMBER_GROUP, "i");
addPattern("\\bt\\s{0,3}=\\s{0,3}" + NUMBER_GROUP, "i");
addPattern("\\bdf\\s{0,3}=\\s{0,3}" + NUMBER_GROUP, "i");
addPattern("95\\s{0,3}percent\\s{0,3}confidence\\s{0,3}interval\\s{0,3}:\\s{0,3}" + NUMBER_GROUP + "\\s{1,3}" + NUMBER_GROUP, "i");
const sitePatternCount = patterns.size;

// ---------------------------------------------------------------- subjects
const corpus = new Set();
for (const t of tokens) {
  if ((t.type.label === "string" || t.type.label === "template") && typeof t.value === "string" && t.value.length <= MAX_SUBJECT) corpus.add(t.value);
}
for (const value of [...corpus]) {
  if (!value.includes("\n")) continue;
  for (const line of value.split("\n")) corpus.add(line);
  for (const line of value.split(/\r?\n/)) corpus.add(line);
}
// Extra strings (offered to every pattern; kept when they match) for patterns that the
// website's own strings never match.
[
  "A minimum sample size of 385 is required to estimate the mean.", "This corresponds to a 25% increase in odds.",
  "x1 is thus it is significant in explaining y", "An increase of 1 unit in x1, while holding all other predictors constant, was statistically significant",
  "while holding other predictors constant, which was not statistically significant", "The optimal solution is X1 = 2. The optimal objective value is 10.",
  "The model explains 65.1% of the variation.", "print", "cat", "library", "printx", "a\n# __BT1101_SECTION_12__\nb",
  "<!-- BT1101 Prose Lab version: v021126_2300h -->", "<!--BT1101 Prose Lab version:021126_2300-->", "<!-- BT1101 Prose Lab version: 21126_2300 -->",
  "<script type=\"abc123-text/javascript\" src=\"x\"></script>", "<script type='abc-text/javascript\"></script>",
  "<script src=\"/cdn-cgi/scripts/7d0fa10a/cloudflare-static/rocket-loader.min.js\" defer></script><script>x</script>",
  "<!DOCTYPE html>", "<title> BT1101 Prose Lab </title>", ".5", "1.2e+ 05 and 3E- 2", ".y.group1group2estimateconf.lowconf.highp.adj",
  "a\\_b\\*c\\d", "\\text{Max } 3 \\times x \\cdot y \\quad z\\, w\\; v\\! \\qquad", "Minimize the objective", "maximise total objective value",
  "x = NaN", "NULL", "null", "plot(x,\n  main = \"Title\"\n)", "as.numeric(dataset$col)", "as.numeric(x)\n",
  "alternative hypothesis: true mean is not equal to 0", "alternative hypothesis: true p is greater than 0.5\nnull probability 0.5",
  "alternative hypothesis: true proportion is less than -1.5e-3", "sample estimates:\ncor \n0.8123", "sample estimates: 0.5",
].forEach(text => corpus.add(text));
const samplesLong = [...corpus].filter(value => value.includes("\n") && value.length > 60);

const TARGETED = [
  "", " ", "a", "\n", "\r\n", "a\nb\r\nc\rd e f", "  leading and trailing  ", "\t\v\f        　﻿|",
  "H₀: μ = 0 vs H₁: μ ≠ 0", "x ≥ 3, y ≤ 4, σ² = 1.5, μ₁ − μ₂", "µ μ Μ ς σ Σ ſ s S K k K",
  "p-value < 2.2e-16", "p-value = 0.05", "p-value=0.0312", "t = -2.345, df = 18, p-value = 0.0312", "W = 0.97, p-value = 0.4",
  "F-statistic: 12.5 on 2 and 47 DF,  p-value: 3.1e-05", "Multiple R-squared:  0.6512,\tAdjusted R-squared:  0.636",
  "X_1 + 2X_{2} >= 10", "3x1 + 4.5 x2 \\leq 20 \\geq 1 ≤ ≥", "Maximize Z = 1000X1 + 1200X2 subject to", "The excluded variables are a and b.",
  "ÀÉÎõü straße ÿŸ İı", "word_word-word.word", "dataset$column1", "ab".repeat(40) + "c",
  "aaa", "abcabc", "Copy table", "Copy ASCII table now", "## [1] 0.95\n## [2] 1.96", "1.5e+03 -2.0E-4 .5 7.", "## Model 1: y ~ x1\n## Model 2: y ~ x1 + x2",
  "Sensitivity : 0.8333\nSpecificity : 0.75\nAccuracy : 0.8", "PC1 = 0.5 x1 + 0.3 x2\nPC2 = -0.1 x1", "<script src=\"a.js\"></script>",
];

function hash(text) {
  let h = 2166136261;
  for (let i = 0; i < text.length; i++) { h ^= text.charCodeAt(i); h = Math.imul(h, 16777619) >>> 0; }
  return h;
}

// ---------------------------------------------------------------- hand-written edge cases
const LONG_TEXT = (samplesLong.join("\n").repeat(60)).slice(0, 50000);
const LONG_X = "x".repeat(50000);
const LONG_SPACES = "a" + " \t\n".repeat(16000) + "b";
// Failing searches with patterns like x*y are quadratic in every backtracking engine
// (JavaScript included), so the long failing subjects stay short.
const CUSTOM = [
  ["a*", "g", ["", "baaab", "aaa"]],
  ["a*?", "g", ["", "baaab"]],
  ["(a*)*", "", ["", "b", "aaab"]],
  ["(a*)+", "", ["", "b", "aaab"]],
  ["(a|)+b", "", ["aab", "b"]],
  ["(?:(a)|b)+", "", ["ab", "ba", "aba"]],
  ["(?:(a)|(b))*", "", ["abab", "aab", "bba"]],
  ["(a)|(b)", "g", ["ab", "ba"]],
  ["(z)?y", "g", ["yzy"]],
  ["x*", "", [""]],
  ["", "g", ["", "abc", "😀x"]],
  ["", "gu", ["😀x", "a😀"]],
  ["(?:)", "g", ["ab"]],
  ["\\b", "g", ["", "ab cd", " ab ", "a"]],
  ["\\B", "g", ["", "ab cd", " ab "]],
  ["^", "gm", ["a\nb\r\nc\rd e", "\n\n"]],
  ["$", "gm", ["a\nb\r\nc\rd e", "\n\n"]],
  ["^\\w+$", "gm", ["one\ntwo three\nfour\r\nfive"]],
  ["^.*$", "gm", ["one\ntwo three\nfour\r\nfive\n"]],
  ["^.*$", "g", ["one\ntwo", "single"]],
  [".", "g", ["a\nb\r  c"]],
  [".", "gs", ["a\nb\r c"]],
  ["a.c", "s", ["a\nc"]],
  ["a+?b", "", ["aaab"]],
  ["a+b", "", ["aaab"]],
  ["a{2,3}", "g", ["aaaaaaa"]],
  ["a{2,3}?", "g", ["aaaaaaa"]],
  ["a{2,}", "g", ["a aa aaa aaaa"]],
  ["a{0}b", "g", ["ab b"]],
  ["(ab){2}", "g", ["ababab"]],
  ["(ab){1,2}?c", "g", ["ababc abc"]],
  ["(a|ab)(c|bcd)(d*)", "", ["abcd"]],
  ["(.*?)x(.*)", "", ["aaxbbxcc"]],
  ["(.*)x(.*?)", "", ["aaxbbxcc"]],
  ["<(.+?)>", "g", ["<a><bb></bb>"]],
  ["\\d+(?=%)", "g", ["10% 20 30%"]],
  ["\\d+(?!%)", "g", ["10% 20 30%"]],
  ["(?=(\\w+))\\1:", "", ["abc:"]],
  ["(?!(a))\\w", "g", ["ab"]],
  ["(?=(a))?", "", ["a"]],
  ["(a)\\1", "gi", ["aA aa Ab"]],
  ["(\\w)\\1", "g", ["hello  bookkeeper"]],
  ["\\1(a)", "", ["aa"]],
  ["[^]", "g", ["a\nb"]],
  ["[]", "g", ["a"]],
  ["[\\b]", "g", ["a\bb"]],
  ["[a-z-]+", "g", ["ab-cd_ef"]],
  ["[\\w-]+", "g", ["ab-cd ef"]],
  ["[\\d-z]+", "g", ["1-z9a"]],
  ["[-a]", "g", ["-a b"]],
  ["[\\]\\\\]", "g", ["a]b\\c"]],
  ["[^\\s\\S]", "g", ["abc"]],
  ["[^\\d\\s]+", "g", ["ab 12 c d"]],
  ["\\S+", "g", ["a b c﻿d e"]],
  ["\\s+", "g", ["a b c﻿d᠎e"]],
  ["\\W+", "g", ["aéb c_d"]],
  ["\\x41\\u0042\\0\\t\\v\\f", "", ["AB\u0000\t\u000b\u000c"]],
  ["\\cJ\\cm", "", ["\n\r"]],
  ["\\c1", "", ["\\c1"]],
  ["[\\c1]", "g", ["\u0011c1\\"]],
  ["\\101\\7\\08", "", ["A\u0007\u00008"]],
  ["\\8\\9", "", ["89"]],
  ["a{", "", ["a{"]],
  ["a{1", "", ["a{1"]],
  ["a{,2}", "", ["a{,2}"]],
  ["x{1}{", "", ["x{"]],
  ["}]", "", ["}]"]],
  ["\\/\\-\\!\\<\\@", "", ["/-!<@"]],
  ["\\u00e9", "i", ["É"]],
  ["é", "i", ["É é"]],
  ["[à-ÿ]+", "gi", ["ÀÉÎõü ×÷ Ÿ"]],
  ["[a-z]+", "gi", ["ABC def ſ K İı"]],
  ["σ", "gi", ["Σ σ ς"]],
  ["μ", "gi", ["µ μ Μ"]],
  ["\\u03a3+", "gi", ["σςΣ"]],
  ["[^a]", "gi", ["aAb"]],
  ["[^σ]", "gi", ["Σςx"]],
  ["k", "gi", ["k K K"]],
  ["\\w", "gi", ["ſK"]],
  ["straße", "i", ["STRASSE straße STRAẞE"]],
  ["да", "gi", ["ДА да Ёё"]],
  ["ё", "gi", ["Ё ё"]],
  ["(^|[\\s,(])([\\p{L}])", "gu", ["mean_value (x) été,βeta ж 中文 1a"]],
  ["\\p{L}+", "gu", ["abc éè αβ жз 中文 あ ٣ 12 Ⅰ"]],
  ["\\P{L}+", "gu", ["ab12 cdⅠ"]],
  ["[^\\p{L}\\d]+", "gu", ["ab12, cd!"]],
  ["\\u{48}\\u{2080}", "u", ["H₀"]],
  ["H₀|H₁", "g", ["H₀: x; H₁: y"]],
  ["[≤≥⩽⩾]", "g", ["a ≤ b ≥ c ⩽ ⩾"]],
  ["μ₁\\s*[−-]\\s*μ₂", "g", ["μ₁ − μ₂ and μ₁-μ₂"]],
  ["[−–—]", "g", ["a−b–c—d-e"]],
  ["\\bThe\\b", "gi", ["the theme The. THE"]],
  ["^\\s*$", "", ["", "   ", " x "]],
  ["(\\d+)-(\\d+)", "g", ["10-20 30-40"]],
  ["(?:a|b)*?c", "", ["ababc"]],
  ["(a+)+b", "", ["aaaaaaaaaaaaaaaaaaab"]],
  ["(x+x+)+y", "", ["xxxxxxxxxxxy"]],
  ["\\s*,\\s*", "", ["a , b,c ,d"]],
  ["(-)|(,)", "", ["a-b,c"]],
  ["x", "", ["axbxc"]],
  ["(x)", "", ["axbxc"]],
  ["x*", "", ["axbxc"]],
  ["(?:)", "", ["abc"]],
  ["[\\s\\S]*?END", "", [LONG_TEXT + "END", LONG_TEXT.slice(0, 3000)]],
  ["\\s+", "g", [LONG_SPACES, LONG_TEXT]],
  ["^.*$", "gm", [LONG_TEXT]],
  ["(a|b)*c", "", ["ab".repeat(20000) + "c", "ab".repeat(1500)]],
  ["(?:ab)+$", "", ["ab".repeat(20000)]],
  ["x*y", "", [LONG_X.slice(0, 3000), LONG_X.slice(0, 3000) + "y"]],
  ["[\\s\\S]*", "", [LONG_X]],
  ["(?:x|y)*?z", "", [LONG_X.slice(0, 5000) + "z"]],
  ["\\bp-value\\b", "gi", [LONG_TEXT]],
  ["(\\w+)\\s*=\\s*([-+0-9.e]+)", "g", [LONG_TEXT]],
];
for (const [source, flags, subjects] of CUSTOM) {
  const entry = addPattern(source, flags, { custom: true });
  entry.subjects = (entry.subjects || []).concat(subjects);
  entry.templates.add("[$&]").add("($1|$2)").add("$$-$`-$'").add("$0$01$10$9$<n>$");
}

// ---------------------------------------------------------------- results
const subjectIds = new Map(), subjectList = [];
function subjectId(text) {
  if (!subjectIds.has(text)) { subjectIds.set(text, subjectList.length); subjectList.push(text); }
  return subjectIds.get(text);
}

function encodeExec(m) {
  if (!m) return "-";
  return m.indices.map(pair => pair ? pair[0] + "," + pair[1] : "-1,-1").join(";");
}

function escapeField(text) {
  // Code-unit escaping: control characters and lone surrogates (which UTF-8 cannot carry).
  let out = "";
  for (let i = 0; i < text.length; i++) {
    const c = text.charCodeAt(i);
    if (c === 92) out += "\\\\";
    else if (c < 0x20 || (c >= 0xd800 && c <= 0xdfff && !isPair(text, i))) out += "\\u" + c.toString(16).padStart(4, "0");
    else if (c >= 0xd800 && c <= 0xdbff) { out += text.slice(i, i + 2); i++; }
    else out += text[i];
  }
  return out;
}

function isPair(text, i) {
  const c = text.charCodeAt(i);
  if (c >= 0xd800 && c <= 0xdbff) { const d = text.charCodeAt(i + 1); return d >= 0xdc00 && d <= 0xdfff; }
  if (c >= 0xdc00 && c <= 0xdfff) { const b = text.charCodeAt(i - 1); return b >= 0xd800 && b <= 0xdbff; }
  return false;
}

function execFrom(source, flags, text, start) {
  const re = new RegExp(source, flags.replace(/[gyd]/g, "") + "gd");
  re.lastIndex = start;
  return re.exec(text);
}

function matchAll(source, flags, text) {
  const re = new RegExp(source, flags.replace(/[gyd]/g, "") + "gd");
  return [...text.matchAll(re)].map(encodeExec).join("|");
}

function encodeSplit(items) {
  return items.map(item => item === undefined ? "\u001e" : escapeField(item)).join("\u001f");
}

const lines = [];
let caseCount = 0, replaceCount = 0, patternIndex = 0;
const genericTemplates = ["<$&>", "{$1}{$2}"];
for (const entry of patterns.values()) {
  const id = patternIndex++;
  const re = new RegExp(entry.source, entry.flags);
  const groups = new RegExp(entry.source + "|", entry.flags.replace(/[gy]/g, "")).exec("").length - 1;
  let subjects;
  if (entry.subjects) {
    subjects = entry.subjects;
  } else {
    const plain = new RegExp(entry.source, entry.flags.replace(/[gy]/g, ""));
    const all = [...corpus, ...TARGETED];
    const matching = all.filter(text => plain.test(text)).sort((a, b) => hash(entry.source + a) - hash(entry.source + b)).slice(0, MAX_MATCHING);
    const nonMatching = samplesLong.filter(text => !plain.test(text)).sort((a, b) => hash(entry.source + a) - hash(entry.source + b)).slice(0, MAX_NONMATCHING);
    subjects = [...new Set([...matching, ...nonMatching, ...TARGETED])];
  }
  lines.push(["P", id, entry.flags || "-", escapeField(entry.source)].join("\t"));
  const templates = [...entry.templates, ...genericTemplates.slice(0, groups ? 2 : 1)];
  for (const text of subjects) {
    const sid = subjectId(text);
    const first = encodeExec(execFrom(entry.source, entry.flags, text, 0));
    const middle = encodeExec(execFrom(entry.source, entry.flags, text, Math.floor(text.length / 2)));
    lines.push(["C", id, sid, first, middle, matchAll(entry.source, entry.flags, text), encodeSplit(text.split(re))].join("\t"));
    caseCount++;
    if (text.length > 5000 && !entry.custom) continue;
    for (const template of templates) {
      if (text.length > 2000 && /\$[`']/.test(template)) continue;
      lines.push(["R", id, sid, escapeField(template), escapeField(text.replace(re, template))].join("\t"));
      replaceCount++;
    }
  }
}

const header = subjectList.map((text, i) => ["S", i, escapeField(text)].join("\t"));
fs.writeFileSync(OUT, header.concat(lines).join("\n") + "\n", "utf8");
console.log(`regex literals: ${literalCount} (${sitePatternCount} distinct site patterns incl. new RegExp), custom patterns: ${CUSTOM.length}`);
console.log(`patterns: ${patterns.size}, subjects: ${subjectList.length}, cases: ${caseCount}, replace cases: ${replaceCount}`);
console.log(`wrote ${OUT} (${(fs.statSync(OUT).size / 1048576).toFixed(1)} MB)`);
