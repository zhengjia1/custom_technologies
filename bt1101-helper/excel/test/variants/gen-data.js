// Generates the VBA data tables of the answer-options engine from the website's own engine
// source, so the port's tables are exact copies: the phrase pools (modVariants3, between
// its "BEGIN/END GENERATED DATA" markers) and the R-object name pools and argument-shuffle
// policies (modVariants2).
//
//   node test/variants/gen-data.js            print both blocks
//   node test/variants/gen-data.js --check    verify the VBA modules contain them verbatim
//   node test/variants/gen-data.js --write    rewrite the blocks in the VBA modules

const fs = require("fs");
const path = require("path");
const vm = require("vm");
const { engineSource } = require("./engine.js");

const context = vm.createContext({ window: {}, visibleOperationBench: () => null, HTMLButtonElement: function () {} });
vm.runInContext(engineSource() + `
;globalThis.__data = { prosePhrasePools, rObjectNamePools, rShuffleSettings, rPureArgumentCalls, rEffectfulArgumentCalls };`, context);
const data = context.__data;

// A VBA string expression for any text (non-ASCII through ChrW$).
function vbString(text) {
  const parts = [];
  let literal = "";
  for (const ch of text) {
    const code = ch.codePointAt(0);
    if (code >= 32 && code < 127) literal += ch === '"' ? '""' : ch;
    else {
      if (literal) { parts.push('"' + literal + '"'); literal = ""; }
      if (code > 0xffff) {
        const high = 0xd800 + ((code - 0x10000) >> 10), low = 0xdc00 + ((code - 0x10000) & 0x3ff);
        parts.push(`ChrW$(&H${high.toString(16).toUpperCase()})`, `ChrW$(&H${low.toString(16).toUpperCase()})`);
      } else if (code === 10) parts.push("vbLf");
      else parts.push(`ChrW$(&H${code.toString(16).toUpperCase()})`);
    }
  }
  if (literal || !parts.length) parts.push('"' + literal + '"');
  return parts.join(" & ");
}

// The literal text every match must contain (a sound pre-filter; "" = always run the regex).
function requiredLiteral(source) {
  let i = 0, out = "";
  if (source.startsWith("\\b")) i = 2;
  else if (source.startsWith("^")) i = 1;
  while (i < source.length) {
    const ch = source[i];
    if (ch === "\\") {
      const next = source[i + 1];
      if (/[().\[\]{}*+?|^$\\\/]/.test(next)) { out += next; i += 2; continue; }
      break;
    }
    if (/[()[\]{}*+?|^$.]/.test(ch)) {
      if (/[*+?{]/.test(ch)) out = out.slice(0, -1);
      break;
    }
    out += ch; i += 1;
  }
  const next = source[i];
  if (next !== undefined && /[*+?{]/.test(next) && out.length && source[i - 1] !== ")") out = out.slice(0, -1);
  return out;
}

const lines = [];
lines.push("Private Sub InitProsePools()");
lines.push("    mPoolCount = 0");
data.prosePhrasePools.forEach(([pattern, pool]) => {
  if (pattern.flags.replace(/[gm]/g, "")) throw new Error("Unexpected flags " + pattern.flags);
  if (pool.length !== 5) throw new Error("Pool size");
  const flags = pattern.flags;
  const literal = requiredLiteral(pattern.source);
  lines.push(`    AddProsePool ${vbString(pattern.source)}, "${flags}", ${vbString(literal)}, _`);
  lines.push(`        ${pool.slice(0, 3).map(vbString).join(", ")}, _`);
  lines.push(`        ${pool.slice(3).map(vbString).join(", ")}`);
});
lines.push("End Sub");
lines.push("");
lines.push("Private Sub InitNamePools()");
Object.entries(data.rObjectNamePools).forEach(([type, pool]) => {
  if (pool.some(name => /\s/.test(name))) throw new Error("space in name");
  lines.push(`    AddNamePool "${type}", "${pool.join(" ")}"`);
});
lines.push("End Sub");
lines.push("");
lines.push("Private Sub InitShufflePolicies()");
for (const [call, policy] of data.rShuffleSettings) {
  lines.push(`    AddShufflePolicy "${call}", "${[...policy.packages].join(" ")}", "${[...policy.names].join(" ")}"`);
}
lines.push(`    mPureCalls = " ${[...data.rPureArgumentCalls].join(" ")} "`);
lines.push(`    mEffectfulCalls = " ${[...data.rEffectfulArgumentCalls].join(" ")} "`);
lines.push("End Sub");

const text = lines.join("\n");
const split = text.indexOf("Private Sub InitNamePools()");
const targets = [["modVariants3.bas", text.slice(0, split).trimEnd()], ["modVariants2.bas", text.slice(split).trimEnd()]];
const mode = process.argv[2];
let stale = 0;
if (!mode) console.log(text);
else for (const [name, block] of targets) {
  const file = path.join(__dirname, "..", "..", "vba", name);
  const source = fs.readFileSync(file, "utf8").replace(/\r\n/g, "\n");
  const begin = source.indexOf("\n", source.indexOf("' BEGIN GENERATED DATA")) + 1, end = source.indexOf("' END GENERATED DATA");
  if (mode === "--write") fs.writeFileSync(file, source.slice(0, begin) + block + "\n" + source.slice(end));
  else if (source.slice(begin, end).trimEnd() !== block) { stale++; console.log(name + ": generated data differs from the website's engine (run --write)"); }
}
if (mode === "--check") console.log(stale ? "STALE" : "VBA data tables match the website's engine");
process.exitCode = stale ? 1 : 0;
