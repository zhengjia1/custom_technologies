// VBA toolchain for the Excel edition. Builds a scratch workbook from excel/vba sources,
// compiles it, and runs macros headlessly in Excel, reporting the exact compile or
// run-time error (dialogs are read and dismissed by dialog-watch.ps1).
//
//   node tools/vba.js check                         compile every module
//   node tools/vba.js run <Macro> [args...]         build, compile, run; print the result
//   options: --only a.bas,b.cls   use only these sources (plus anything named by --with)
//            --no-form           leave out frmApp
//            --keep <file.xlsm>  keep the built workbook at this path
//            --argfile <path>    read the macro's first argument from a UTF-8 file
//
// VBA sources must be plain ASCII; they are imported with CRLF line endings.

const fs = require("fs");
const os = require("os");
const path = require("path");
const { spawnSync, spawn } = require("child_process");

const ROOT = path.join(__dirname, "..");
const VBA = path.join(ROOT, "vba");
const FORM_CODE = "frmApp.code.vba";
const WORKBOOK_CODE = "ThisWorkbook.code.vba";

function parseArgs(argv) {
  const options = { only: null, form: true, keep: null, argfile: null, rest: [] };
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if (arg === "--only") options.only = argv[++i].split(",").map(name => name.trim()).filter(Boolean);
    else if (arg === "--no-form") options.form = false;
    else if (arg === "--keep") options.keep = path.resolve(argv[++i]);
    else if (arg === "--argfile") options.argfile = path.resolve(argv[++i]);
    else options.rest.push(arg);
  }
  return options;
}

function prepareSource(name, dir) {
  const source = fs.readFileSync(path.join(VBA, name), "utf8");
  const bad = source.split(/\r?\n/).findIndex(line => /[^\x00-\x7f]/.test(line));
  if (bad >= 0) throw new Error(`vba/${name} line ${bad + 1} is not plain ASCII (use ChrW$ for other characters)`);
  const file = path.join(dir, name);
  fs.writeFileSync(file, source.replace(/\r?\n/g, "\r\n"), "latin1");
  return file;
}

// At most MAX_EXCEL Excel instances run at once across every caller on this machine; the
// others wait for a slot (a directory under the temp folder; stale slots expire).
const MAX_EXCEL = Number(process.env.BT1101_MAX_EXCEL) || 3;
const SLOTS = path.join(os.tmpdir(), "bt1101-excel-slots");
const sleep = ms => Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms);
function acquireSlot() {
  fs.mkdirSync(SLOTS, { recursive: true });
  for (;;) {
    for (let i = 0; i < MAX_EXCEL; i++) {
      const slot = path.join(SLOTS, "slot-" + i);
      try { fs.mkdirSync(slot); return slot; } catch (error) {
        try { if (Date.now() - fs.statSync(slot).mtimeMs > 20 * 60 * 1000) fs.rmSync(slot, { recursive: true, force: true }); } catch (ignored) { /* raced */ }
      }
    }
    sleep(1000);
  }
}

// Runs excel-task.wsf with a dialog watcher attached to the Excel instance it starts.
function excelTask(job, dir) {
  const slot = acquireSlot();
  try { return runExcelTask(job, dir); } finally { fs.rmSync(slot, { recursive: true, force: true }); }
}

function runExcelTask(job, dir) {
  const hwndFile = path.join(dir, `hwnd-${job.action}.txt`), logFile = path.join(dir, `dialogs-${job.action}.txt`), stopFile = path.join(dir, `stop-${job.action}.txt`);
  [hwndFile, logFile, stopFile].forEach(file => { if (fs.existsSync(file)) fs.unlinkSync(file); });
  job.hwndFile = hwndFile;
  const jobFile = path.join(dir, `job-${job.action}.json`);
  fs.writeFileSync(jobFile, JSON.stringify(job));
  const watcher = spawn("powershell", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", path.join(__dirname, "dialog-watch.ps1"), "-HwndFile", hwndFile, "-LogFile", logFile, "-StopFile", stopFile], { stdio: "ignore" });
  const result = spawnSync("cscript", ["//nologo", path.join(__dirname, "excel-task.wsf"), jobFile], { encoding: "utf8", timeout: 15 * 60 * 1000 });
  fs.writeFileSync(stopFile, "stop");
  const started = Date.now();
  while (watcher.exitCode === null && Date.now() - started < 3000) spawnSync("powershell", ["-NoProfile", "-Command", "Start-Sleep -Milliseconds 200"]);
  if (watcher.exitCode === null) watcher.kill();
  const dialogs = fs.existsSync(logFile) ? fs.readFileSync(logFile, "utf8").replace(/^﻿/, "").trim() : "";
  return { status: result.status, stdout: (result.stdout || "").trim(), stderr: (result.stderr || "").trim(), dialogs };
}

function buildWorkbook(options, dir) {
  const names = fs.readdirSync(VBA).filter(name => /\.(bas|cls)$/i.test(name));
  const chosen = options.only ? names.filter(name => options.only.includes(name)) : names;
  if (options.only) options.only.filter(name => /\.(bas|cls)$/i.test(name) && !names.includes(name)).forEach(name => { throw new Error(`vba/${name} does not exist`); });
  const job = { action: "build", output: options.keep || path.join(dir, "test.xlsm"), modules: chosen.map(name => prepareSource(name, dir)) };
  const wantForm = options.form && fs.existsSync(path.join(VBA, FORM_CODE)) && (!options.only || options.only.includes(FORM_CODE));
  if (wantForm) job.form = { name: "frmApp", caption: "BT1101 Prose Lab", code: prepareSource(FORM_CODE, dir) };
  if (fs.existsSync(path.join(VBA, WORKBOOK_CODE)) && (!options.only || options.only.includes(WORKBOOK_CODE))) job.workbookCode = prepareSource(WORKBOOK_CODE, dir);
  const built = excelTask(job, dir);
  if (built.status !== 0) throw new Error(`Build failed: ${built.stdout} ${built.stderr} ${built.dialogs}`);
  return job.output;
}

function main() {
  const [command, ...argv] = process.argv.slice(2);
  const options = parseArgs(argv);
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "bt1101-vba-"));
  try {
    const workbook = buildWorkbook(options, dir);
    const compiled = excelTask({ action: "compile", workbook }, dir);
    if (compiled.status !== 0) {
      console.log(compiled.stdout + (compiled.dialogs ? "\n" + compiled.dialogs : ""));
      process.exitCode = 1;
      return;
    }
    if (command === "check") { console.log("Compiled OK"); return; }
    if (command === "run") {
      const [macro, ...args] = options.rest;
      if (options.argfile) args.unshift(fs.readFileSync(options.argfile, "utf8"));
      const resultFile = path.join(dir, "result.txt");
      const ran = excelTask({ action: "run", workbook, macro, args, resultFile }, dir);
      if (fs.existsSync(resultFile)) process.stdout.write(fs.readFileSync(resultFile, "utf8").replace(/^﻿/, "") + "\n");
      if (ran.stdout) console.log(ran.stdout);
      if (ran.dialogs) { console.log("DIALOGS:\n" + ran.dialogs); process.exitCode = 1; }
      if (ran.status !== 0) process.exitCode = 1;
      return;
    }
    throw new Error("Usage: node tools/vba.js check|run <Macro> [args] [--only a.bas,b.cls] [--no-form] [--keep file.xlsm] [--argfile path]");
  } finally {
    // Excel may still be releasing the workbook for a moment after quitting.
    if (!process.env.BT1101_KEEP_TMP) fs.rmSync(dir, { recursive: true, force: true, maxRetries: 20, retryDelay: 250 });
  }
}

if (require.main === module) {
  try { main(); } catch (error) { console.error(error.message); process.exitCode = 2; }
}
module.exports = { prepareSource, excelTask };
