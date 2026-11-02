// Runs identical scripted sessions against the real page (jsdom) and against minidom,
// with the same random sequence and virtual clock, and reports any difference in
// generated text, form state or clipboard writes.

const path = require("path");
const { JSDOM } = require("jsdom");
const { bundleForNode, ROOT } = require("./harness.js");

function seededRandom(seed, counter) {
  let state = seed >>> 0;
  return () => { if (counter) counter.calls++; state = (state + 0x6D2B79F5) >>> 0; let t = state; t = Math.imul(t ^ (t >>> 15), t | 1); t ^= t + Math.imul(t ^ (t >>> 7), t | 61); return ((t ^ (t >>> 14)) >>> 0) / 4294967296; };
}

function makeTasks() { return { micro: [], timers: [], seq: 0, now: 0 }; }

function virtualTimers(tasks) {
  const setTimeout = (fn, delay, ...args) => { const id = ++tasks.seq; if (typeof fn === "function") tasks.timers.push({ id, due: tasks.now + Math.max(0, Number(delay) || 0), seq: id, fn, args, interval: 0 }); return id; };
  const clearTimeout = id => { const i = tasks.timers.findIndex(t => t.id === id); if (i >= 0) tasks.timers.splice(i, 1); };
  return { setTimeout, clearTimeout, requestAnimationFrame: fn => setTimeout(() => fn(tasks.now), 16), queueMicrotask: fn => tasks.micro.push(fn) };
}

function drainTasks(tasks, report) {
  for (let steps = 0; steps < 20000; steps++) {
    if (tasks.micro.length) { const fn = tasks.micro.shift(); try { fn(); } catch (e) { report(e); } continue; }
    if (!tasks.timers.length) return;
    let next = 0;
    tasks.timers.forEach((t, i) => { const b = tasks.timers[next]; if (t.due < b.due || (t.due === b.due && t.seq < b.seq)) next = i; });
    const timer = tasks.timers.splice(next, 1)[0];
    tasks.now = Math.max(tasks.now, timer.due);
    try { timer.fn.apply(null, timer.args); } catch (e) { report(e); }
  }
  throw new Error("tasks did not settle");
}

// Browser order: all microtasks (native and queued) settle before the next timer fires.
const flushNative = () => new Promise(resolve => setImmediate(resolve));
async function settleJsdom(session) {
  for (let steps = 0; steps < 20000; steps++) {
    await flushNative();
    const tasks = session.tasks;
    if (tasks.micro.length) { const fn = tasks.micro.shift(); try { fn(); } catch (e) { session.report(e); } continue; }
    if (!tasks.timers.length) return;
    let next = 0;
    tasks.timers.forEach((t, i) => { const b = tasks.timers[next]; if (t.due < b.due || (t.due === b.due && t.seq < b.seq)) next = i; });
    const timer = tasks.timers.splice(next, 1)[0];
    tasks.now = Math.max(tasks.now, timer.due);
    try { timer.fn.apply(null, timer.args); } catch (e) { session.report(e); }
  }
  throw new Error("tasks did not settle");
}

function bootJsdom(site, seed) {
  const tasks = makeTasks();
  const errors = [];
  const clipboard = [];
  const timers = virtualTimers(tasks);
  const html = site.html;
  const counter = { calls: 0 };
  const dom = new JSDOM(html, {
    runScripts: "dangerously",
    url: "https://bt1101-pdf.invalid/",
    beforeParse(window) {
      window.Math.random = seededRandom(seed, counter);
      Object.defineProperty(window, "crypto", { value: undefined, configurable: true });
      window.setTimeout = timers.setTimeout; window.clearTimeout = timers.clearTimeout;
      window.requestAnimationFrame = timers.requestAnimationFrame; window.cancelAnimationFrame = timers.clearTimeout;
      Object.defineProperty(window.navigator, "clipboard", { value: { writeText: text => { clipboard.push(String(text)); return Promise.resolve(); }, readText: () => Promise.reject(new Error("no")) }, configurable: true });
      window.CSS = { escape: require("css.escape") };
      window.matchMedia = () => ({ matches: false, addEventListener() {}, removeEventListener() {}, addListener() {}, removeListener() {} });
      window.fetch = () => Promise.reject(new Error("offline"));
      window.scrollTo = () => {};
      window.HTMLElement.prototype.scrollIntoView = function () {};
      window.Number.prototype.toLocaleString = Number.prototype.toLocaleString;
      window.addEventListener("error", event => errors.push(event.error || event.message));
    }
  });
  const report = error => errors.push(error);
  return { kind: "jsdom", window: dom.window, document: dom.window.document, tasks, errors, clipboard, report, counter, Event: dom.window.Event };
}

function bootMinidom(lib, site, seed) {
  const tasks = makeTasks();
  const errors = [];
  const report = error => errors.push(error);
  const originalRandom = Math.random;
  const counter = { calls: 0 };
  Math.random = seededRandom(seed, counter);
  lib.installNumberLocale();
  const browser = lib.createBrowser(site.body, tasks, { reportError: report });
  const session = { kind: "minidom", window: browser.window, document: browser.document, tasks, errors, clipboard: browser.clipboard.writes, report, counter, Event: lib.dom.Event, restoreRandom: () => { Math.random = originalRandom; } };
  try { session.app = lib.runApp(browser.globals); } catch (error) { errors.push(error); }
  return session;
}

// ---------------------------------------------------------------- observation

const OUTPUT_SELECTOR = "pre, .output-body, .pca-variance-prose, .sensitivity-prose-grid, .hypothesis-card, .mapping-card, .shape-guide, .lp-parser-status, .lp-constraint-error, .lp-table-wrap, #toast, #activeTitle, .completion-badge, .panel-title";

function snapshot(session) {
  const document = session.document;
  const bench = document.querySelector("main > .bench:not(.hidden), main > #lpBench:not(.hidden)");
  const scope = bench || document.body;
  const lines = [];
  lines.push("active=" + document.querySelector("#activeTitle").textContent.trim());
  if (session.trackRandom !== false) lines.push("random=" + session.counter.calls);
  [...scope.querySelectorAll(OUTPUT_SELECTOR)].forEach((node, index) => {
    lines.push(`out[${index}]#${node.id}.${node.className}${node.closest(".hidden") ? " (hidden)" : ""}: ${node.textContent.replace(/\s+/g, " ").trim()}`);
  });
  [...scope.querySelectorAll("input, select, textarea, button")].forEach((node, index) => {
    const state = node.type === "checkbox" ? node.checked : node.value;
    lines.push(`ctl[${index}]${node.tagName}#${node.id}.${node.className} disabled=${node.disabled} hidden=${Boolean(node.closest(".hidden"))} value=${JSON.stringify(state)} text=${node.tagName === "BUTTON" ? node.textContent.trim() : ""}`);
  });
  lines.push("toast=" + document.querySelector("#toast").textContent);
  [...document.querySelectorAll(".variable-rename-alert:not(.hidden)")].forEach(alert => lines.push("alert=" + alert.id));
  lines.push("clipboard=" + JSON.stringify(session.clipboard.slice(-3)));
  return lines;
}

// ---------------------------------------------------------------- actions (identical API in both DOMs)

function fire(session, element, type) {
  const Event = session.kind === "jsdom" ? session.window.Event : session.Event;
  element.dispatchEvent(new Event(type, { bubbles: true }));
}
function setValue(session, element, value) {
  if (element.tagName === "TEXTAREA") fire(session, element, "paste");
  element.value = value;
  fire(session, element, "input");
  fire(session, element, "change");
}
function openOperation(session, id) {
  const entry = id === "universal" ? session.document.querySelector(".universal-entry") : session.document.querySelector(`#tutorialList .operation[data-id="${id}"]`);
  entry.click();
}

const ACTIONS = [];
function visibleBench(session) { return session.document.querySelector("main > .bench:not(.hidden), main > #lpBench:not(.hidden)"); }
function visibleButtons(session, pattern) {
  const bench = visibleBench(session);
  return bench ? [...bench.querySelectorAll("button")].filter(button => !button.disabled && !button.closest(".hidden") && pattern.test(button.textContent.trim())) : [];
}

async function step(session, action) {
  action(session);
  await settleJsdom(session);
}

// A scripted session: visit each operation, load its samples, toggle controls and copy.
function buildScript(operationIds) {
  const script = [];
  const add = (label, action) => script.push({ label, action });
  operationIds.forEach(id => {
    add(`open ${id}`, s => openOperation(s, id));
    add(`${id}: samples`, s => visibleButtons(s, /^Use sample/i).forEach(button => button.click()));
    add(`${id}: universal samples`, s => {
      const select = s.document.querySelector("#universalSampleSelect");
      if (id === "universal" && select) { select.value = select.options[1].value; fire(s, select, "change"); }
    });
    add(`${id}: copy all`, s => visibleButtons(s, /^Copy\b/i).forEach(button => button.click()));
    add(`${id}: dismiss alerts`, s => [...s.document.querySelectorAll(".variable-rename-alert:not(.hidden) .variable-rename-dismiss")].slice(-1).forEach(button => button.click()));
    add(`${id}: answer options`, s => {
      const bench = visibleBench(s);
      if (!bench) return;
      [...bench.querySelectorAll(".answer-variant-select")].forEach((select, index) => { select.value = String(((index + 2) % 6) + 1); fire(s, select, "change"); });
    });
    add(`${id}: cycle selects`, s => {
      const bench = visibleBench(s);
      if (!bench) return;
      [...bench.querySelectorAll("select:not(.answer-variant-select)")].filter(select => !select.disabled && !select.closest(".hidden")).slice(0, 4).forEach(select => {
        const options = [...select.options].filter(option => !option.disabled);
        if (options.length > 1) { select.value = options[options.length - 1].value; fire(s, select, "input"); fire(s, select, "change"); }
      });
    });
    add(`${id}: dismiss alerts 2`, s => [...s.document.querySelectorAll(".variable-rename-alert:not(.hidden) .variable-rename-dismiss")].slice(-1).forEach(button => button.click()));
    add(`${id}: toggle checkboxes`, s => {
      const bench = visibleBench(s);
      if (!bench) return;
      [...bench.querySelectorAll('input[type="checkbox"]')].filter(box => !box.disabled && !box.closest(".hidden")).slice(0, 3).forEach(box => box.click());
    });
    add(`${id}: dismiss alerts 3`, s => [...s.document.querySelectorAll(".variable-rename-alert:not(.hidden) .variable-rename-dismiss")].slice(-1).forEach(button => button.click()));
    add(`${id}: edit text`, s => {
      const bench = visibleBench(s);
      if (!bench) return;
      [...bench.querySelectorAll('input[type="text"], input:not([type])')].filter(input => !input.disabled && !input.closest(".hidden")).slice(0, 2).forEach(input => setValue(s, input, input.value ? input.value + "_b" : "column9"));
    });
    add(`${id}: copy again`, s => visibleButtons(s, /^Copy\b/i).slice(0, 2).forEach(button => button.click()));
    add(`${id}: dismiss alerts 4`, s => [...s.document.querySelectorAll(".variable-rename-alert:not(.hidden) .variable-rename-dismiss")].slice(-1).forEach(button => button.click()));
  });
  return script;
}

async function runDifferential(options) {
  options = options || {};
  const { site, lib } = bundleForNode();
  const seed = options.seed || 12345;
  const jsdom = bootJsdom(site, seed);
  await settleJsdom(jsdom);
  const mini = bootMinidom(lib, site, seed);
  drainTasks(mini.tasks, mini.report);
  const operationIds = ["universal"].concat([...jsdom.document.querySelectorAll("#tutorialList .operation")].map(button => button.dataset.id));
  const script = buildScript(options.operations || operationIds);
  const differences = [];
  for (const entry of script) {
    await step(jsdom, entry.action);
    await step(mini, entry.action);
    const a = snapshot(jsdom), b = snapshot(mini);
    const length = Math.max(a.length, b.length);
    for (let i = 0; i < length; i++) {
      if (a[i] !== b[i]) { differences.push({ step: entry.label, jsdom: a[i], minidom: b[i] }); break; }
    }
    if (differences.length >= (options.maxDifferences || 15)) break;
  }
  mini.restoreRandom();
  return { differences, jsdomErrors: jsdom.errors, minidomErrors: mini.errors, steps: script.length };
}

module.exports = { runDifferential, seededRandom };

if (require.main === module) {
  runDifferential({ operations: process.argv[2] ? process.argv.slice(2) : undefined }).then(result => {
    console.log(`steps=${result.steps} differences=${result.differences.length} jsdomErrors=${result.jsdomErrors.length} minidomErrors=${result.minidomErrors.length}`);
    result.differences.forEach(d => console.log(`\n[${d.step}]\n  jsdom:   ${String(d.jsdom).slice(0, 400)}\n  minidom: ${String(d.minidom).slice(0, 400)}`));
    result.jsdomErrors.slice(0, 5).forEach(e => console.log("jsdom error:", e && e.stack ? e.stack.split("\n").slice(0, 3).join(" | ") : e));
    result.minidomErrors.slice(0, 5).forEach(e => console.log("minidom error:", e && e.stack ? e.stack.split("\n").slice(0, 3).join(" | ") : e));
  });
}
