// Runs a step script against the real website logic (minidom in Node) and prints the
// recorded results. vba/testDriver.bas runs the same script against the VBA port, so the
// two outputs can be compared line for line (test/compare.js).
//
//   node test/ref-run.js <steps.txt> [--seed 1]
//
// Step lines are tab-separated; arguments use \\ \n \t \r escapes.
//   open <operationId>                click its sidebar entry (universal = the detector)
//   value <selector> <text>           set .value without events
//   input <selector> [text]           optionally set .value, then dispatch "input"
//   change <selector> [text]          optionally set .value, then dispatch "change"
//   select <selector> <index>         set selectedIndex, dispatch input + change
//   click <selector>                  element.click()
//   key <selector> <key> [s|c|a...]   dispatch keydown
//   canon                             set every answer option to 1 (canonical answers)
//   read <selector>                   record textContent
//   val <selector>                    record .value
//   class <selector>                  record the class attribute
//   attr <selector> <name>            record an attribute (or <null>)
//   checked <selector>                record .checked
//   disabled <selector>               record .disabled
//   count <selector>                  record the number of matches
//   clip                              record the last clipboard write
//   toast                             record the toast text and whether it is shown
// Every action drains timers and microtasks afterwards, like a browser between events.

const fs = require("fs");
const path = require("path");
const os = require("os");
const esbuild = require(require.resolve("esbuild", { paths: [path.join(__dirname, "..", "..", "pdf")] }));
const { extractWebsite, writeAppModule } = require("../../pdf/src/extract.js");

// Bundles the website for Node once per process into a private folder, so parallel runs
// never overwrite each other's bundle (pdf/test/harness.js shares pdf/.build).
let bundled = null;
function bootNodeApp() {
  if (!bundled) {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), "bt1101-ref-"));
    const site = extractWebsite(path.join(__dirname, "..", "..", "index.html"));
    writeAppModule(site, path.join(dir, "app.js"));
    const src = path.join(__dirname, "..", "..", "pdf", "src").split(path.sep).join("/");
    fs.writeFileSync(path.join(dir, "entry.js"), `
      export { createBrowser } from "${src}/env.js";
      export { runApp } from "./app.js";
    `);
    const out = path.join(dir, "bundle.js");
    esbuild.buildSync({ entryPoints: [path.join(dir, "entry.js")], bundle: true, format: "cjs", platform: "node", outfile: out, logLevel: "error" });
    bundled = { site, lib: require(out), dir };
    process.on("exit", () => fs.rmSync(dir, { recursive: true, force: true }));
  }
  const { site, lib } = bundled;
  const tasks = { micro: [], timers: [], seq: 0, now: 0 };
  const errors = [];
  const browser = lib.createBrowser(site.body, tasks, { reportError: error => errors.push(error) });
  try { lib.runApp(browser.globals); } catch (error) { errors.push(error); }
  return { browser, errors, tasks };
}

const unescape = text => text.replace(/\\(\\|n|t|r)/g, (_, c) => ({ "\\": "\\", n: "\n", t: "\t", r: "\r" })[c]);
const escape = text => String(text).replace(/\\/g, "\\\\").replace(/\n/g, "\\n").replace(/\t/g, "\\t").replace(/\r/g, "\\r");

function mulberry32(seed) {
  let state = seed >>> 0;
  return () => { state = (state + 0x6D2B79F5) >>> 0; let t = state; t = Math.imul(t ^ (t >>> 15), t | 1); t ^= t + Math.imul(t ^ (t >>> 7), t | 61); return ((t ^ (t >>> 14)) >>> 0) / 4294967296; };
}

// Like the VBA framework's Settle 700: native promise jobs, then timers due within the
// horizon in due order; later timers (toast hiding, "Copied" captions) stay pending.
const HORIZON = 700;
const flushNative = () => new Promise(resolve => setImmediate(resolve));
async function settle(tasks, report) {
  const limit = tasks.now + HORIZON;
  for (let steps = 0; steps < 20000; steps++) {
    await flushNative();
    if (tasks.micro.length) { const fn = tasks.micro.shift(); try { fn(); } catch (error) { report(error); } continue; }
    let next = -1;
    tasks.timers.forEach((timer, i) => { if (timer.due <= limit && (next < 0 || timer.due < tasks.timers[next].due || (timer.due === tasks.timers[next].due && timer.seq < tasks.timers[next].seq))) next = i; });
    if (next < 0) return;
    const timer = tasks.timers.splice(next, 1)[0];
    tasks.now = Math.max(tasks.now, timer.due);
    try { timer.fn.apply(null, timer.args); } catch (error) { report(error); }
  }
  throw new Error("tasks did not settle");
}

async function run(stepsText, seed) {
  Math.random = mulberry32(seed);
  const { browser, errors, tasks } = bootNodeApp();
  await settle(tasks, error => errors.push(error));
  const document = browser.document, Event = browser.globals.Event, KeyboardEvent = browser.globals.KeyboardEvent;
  const out = [];
  const one = selector => {
    const element = document.querySelector(selector);
    if (!element) throw new Error("No element matches " + selector);
    return element;
  };
  const record = (step, value) => out.push(step + "\t" + escape(value));
  const lines = stepsText.split(/\r?\n/);
  for (let index = 0; index < lines.length; index++) {
    const line = lines[index];
    if (!line.trim() || line.startsWith("#")) continue;
    const [command, ...rawArgs] = line.split("\t");
    const args = rawArgs.map(unescape);
    const where = `line ${index + 1} (${command})`;
    try {
      switch (command) {
        case "open": {
          const entry = args[0] === "universal" ? document.querySelector(".universal-entry") : document.querySelector(`#tutorialList .operation[data-id="${args[0]}"]`);
          if (!entry) throw new Error("No operation " + args[0]);
          entry.click();
          break;
        }
        case "value": one(args[0]).value = args[1]; break;
        case "input": { const el = one(args[0]); if (args.length > 1) el.value = args[1]; el.dispatchEvent(new Event("input", { bubbles: true })); break; }
        case "change": { const el = one(args[0]); if (args.length > 1) el.value = args[1]; el.dispatchEvent(new Event("change", { bubbles: true })); break; }
        case "select": { const el = one(args[0]); el.selectedIndex = Number(args[1]); el.dispatchEvent(new Event("input", { bubbles: true })); el.dispatchEvent(new Event("change", { bubbles: true })); break; }
        case "click": one(args[0]).click(); break;
        case "key": {
          const mods = args[2] || "";
          one(args[0]).dispatchEvent(new KeyboardEvent("keydown", { key: args[1], shiftKey: mods.includes("s"), ctrlKey: mods.includes("c"), altKey: mods.includes("a"), bubbles: true, cancelable: true }));
          break;
        }
        case "canon":
          document.querySelectorAll(".answer-variant-select").forEach(select => { if (select.value !== "1") { select.value = "1"; select.dispatchEvent(new Event("change", { bubbles: true })); } });
          break;
        case "read": record(line, one(args[0]).textContent); break;
        case "val": record(line, one(args[0]).value); break;
        case "class": record(line, one(args[0]).getAttribute("class") || ""); break;
        case "attr": { const value = one(args[0]).getAttribute(args[1]); record(line, value === null ? "<null>" : value); break; }
        case "checked": record(line, one(args[0]).checked ? "1" : "0"); break;
        case "disabled": record(line, one(args[0]).disabled ? "1" : "0"); break;
        case "count": record(line, document.querySelectorAll(args[0]).length); break;
        case "clip": record(line, browser.clipboard.writes.length ? browser.clipboard.writes[browser.clipboard.writes.length - 1] : "<none>"); break;
        case "toast": { const toast = document.getElementById("toast"); record(line, (toast.classList.contains("show") ? "shown: " : "hidden: ") + toast.textContent); break; }
        default: throw new Error("Unknown step");
      }
      if (!["read", "val", "class", "attr", "checked", "disabled", "count", "clip", "toast"].includes(command)) await settle(tasks, error => errors.push(error));
    } catch (error) {
      record(line, "ERROR " + where + ": " + error.message);
    }
  }
  errors.forEach(error => out.push("JSERROR\t" + escape(error && error.message ? error.message : error)));
  return out.join("\n");
}

if (require.main === module) {
  const file = process.argv[2];
  const seedIndex = process.argv.indexOf("--seed");
  const seed = seedIndex > 0 ? Number(process.argv[seedIndex + 1]) : 1;
  run(fs.readFileSync(path.resolve(file), "utf8"), seed).then(text => process.stdout.write(text + "\n"));
}
module.exports = { run, escape, unescape };
