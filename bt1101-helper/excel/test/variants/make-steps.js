// Writes test/steps/modVariants/<name>.txt from test/variants/steps-src.txt ("@@ name"
// starts a file; " | " separates the tab-separated fields of a step).
const fs = require("fs");
const path = require("path");
const src = fs.readFileSync(path.join(__dirname, "steps-src.txt"), "utf8").split(/\r?\n/);
const out = path.join(__dirname, "..", "steps", "modVariants");
let name = null, lines = [];
const flush = () => { if (name) fs.writeFileSync(path.join(out, name + ".txt"), lines.join("\n") + "\n"); lines = []; };
for (const line of src) {
  if (line.startsWith("##")) continue;
  const m = line.match(/^@@ (\S+)/);
  if (m) { flush(); name = m[1]; continue; }
  if (name && line.trim()) lines.push(line.startsWith("#") ? line : line.split(" | ").join("\t"));
}
flush();
console.log(fs.readdirSync(out).join(" "));
