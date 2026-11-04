// Differential test: runs a step script against the website (test/ref-run.js) and against
// the VBA port (vba/testDriver.bas through tools/vba.js), then reports differing records.
//
//   node test/compare.js <steps.txt> [more steps files...] [--boot Module.Proc] [--only a.bas,...]
//
// Several step files run in one Excel session (one VBA call per file, one workbook build).

const fs = require("fs");
const os = require("os");
const path = require("path");
const { spawnSync } = require("child_process");
const { run } = require("./ref-run.js");

async function main() {
  const argv = process.argv.slice(2);
  const files = [], passthrough = [];
  let boot = "InitApp";
  for (let i = 0; i < argv.length; i++) {
    if (argv[i] === "--boot") boot = argv[++i];
    else if (argv[i].startsWith("--")) { passthrough.push(argv[i]); if (!["--no-form"].includes(argv[i])) passthrough.push(argv[++i]); }
    else files.push(path.resolve(argv[i]));
  }
  if (!files.length) throw new Error("Usage: node test/compare.js <steps.txt>... [--boot Module.Proc] [--only ...]");
  const body = path.join(__dirname, "..", ".build", "body.html");
  if (!fs.existsSync(body)) spawnSync(process.execPath, [path.join(__dirname, "body.js")], { stdio: "inherit" });
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "bt1101-compare-"));
  // One VBA call runs every file: a list file names each prepared steps file.
  const prepared = files.map((file, index) => {
    const target = path.join(dir, `steps-${index}.txt`);
    fs.writeFileSync(target, `load\t${body.replace(/\\/g, "\\\\")}\nboot\t${boot}\n` + fs.readFileSync(file, "utf8"));
    return target;
  });
  const listFile = path.join(dir, "list.txt");
  fs.writeFileSync(listFile, prepared.join("\n"));
  const vba = spawnSync(process.execPath, [path.join(__dirname, "..", "tools", "vba.js"), "run", "testDriver.RunStepsList", listFile, ...passthrough], { encoding: "utf8", maxBuffer: 64 * 1024 * 1024 });
  const vbaOut = (vba.stdout || "") + (vba.stderr || "");
  const sections = vbaOut.split(/\n?=====FILE=====\n/);
  let failures = 0;
  for (let i = 0; i < files.length; i++) {
    const expected = (await run(fs.readFileSync(files[i], "utf8"), 1)).split("\n");
    const actual = (sections[i + 1] || "").replace(/\s+$/, "").split("\n");
    const name = path.relative(process.cwd(), files[i]);
    const differences = [];
    for (let k = 0; k < Math.max(expected.length, actual.length); k++) {
      if ((expected[k] || "") !== (actual[k] || "")) differences.push(`  #${k + 1}\n    website: ${expected[k] || "<missing>"}\n    vba:     ${actual[k] || "<missing>"}`);
    }
    if (differences.length) { failures++; console.log(`FAIL ${name} (${differences.length} differing records)\n${differences.slice(0, 20).join("\n")}`); }
    else console.log(`PASS ${name} (${expected.length} records)`);
  }
  if (sections.length < files.length + 1) console.log("VBA output:\n" + vbaOut.slice(0, 4000));
  fs.rmSync(dir, { recursive: true, force: true });
  process.exitCode = failures ? 1 : 0;
}

main().catch(error => { console.error(error.message); process.exitCode = 2; });
