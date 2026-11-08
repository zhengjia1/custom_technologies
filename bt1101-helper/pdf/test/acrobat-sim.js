// Runs the exact document-level script from the PDF in a sandbox shaped like Acrobat:
// an ES5-era global (no Promise, Map, Symbol, timers…) plus mock Doc, Field, app,
// display and console objects. Field events are delivered the way Acrobat does.

const fs = require("fs");
const path = require("path");
const vm = require("vm");
const { seededRandom } = require("./differential.js");

const MODERN = {
  global: ["Promise", "Map", "Set", "WeakMap", "WeakSet", "WeakRef", "FinalizationRegistry", "Symbol", "Reflect", "Proxy", "globalThis",
    "queueMicrotask", "setTimeout", "clearTimeout", "setInterval", "clearInterval", "setImmediate", "structuredClone", "URL", "URLSearchParams",
    "TextEncoder", "TextDecoder", "AggregateError", "BigInt", "Intl", "SharedArrayBuffer", "Atomics", "Iterator", "AsyncIterator"],
  Object: ["assign", "entries", "values", "fromEntries", "getOwnPropertySymbols", "is", "setPrototypeOf", "getOwnPropertyDescriptors", "hasOwn", "groupBy"],
  Array: ["from", "of"],
  "Array.prototype": ["find", "findIndex", "findLast", "findLastIndex", "includes", "flat", "flatMap", "fill", "copyWithin", "at", "entries", "keys", "values", "toSorted", "toReversed", "with"],
  String: ["fromCodePoint", "raw"],
  "String.prototype": ["includes", "startsWith", "endsWith", "padStart", "padEnd", "repeat", "at", "codePointAt", "normalize", "trimStart", "trimEnd", "replaceAll", "matchAll"],
  Number: ["isFinite", "isInteger", "isNaN", "isSafeInteger", "parseFloat", "parseInt", "EPSILON", "MAX_SAFE_INTEGER", "MIN_SAFE_INTEGER"],
  Math: ["trunc", "sign", "log10", "log2", "log1p", "expm1", "cbrt", "hypot", "imul", "clz32", "fround", "sinh", "cosh", "tanh", "asinh", "acosh", "atanh"]
};

function stripToES5(context) {
  vm.runInContext(`(function (spec) {
    var g = Function('return this')();
    function target(path) { var parts = path.split('.'), node = g; for (var i = 0; i < parts.length; i++) node = node[parts[i]]; return node; }
    Object.keys(spec).forEach(function (key) {
      var holder = key === 'global' ? g : target(key);
      spec[key].forEach(function (name) { try { delete holder[name]; } catch (e) {} if (name in holder) { try { Object.defineProperty(holder, name, { value: undefined, configurable: true, writable: true }); } catch (e) {} } });
    });
  })(${JSON.stringify(MODERN)});`, context);
}

// Acrobat locks at least Function.prototype.toString. Lock every existing built-in method
// read-only and non-configurable, the strictest case, so core-js must cope.
function lockBuiltIns(context) {
  vm.runInContext(`(function () {
    var names = ['Object', 'Function', 'Array', 'String', 'Number', 'Boolean', 'Math', 'RegExp', 'Date', 'JSON', 'Error'];
    names.forEach(function (name) {
      var holder = this[name];
      [holder, holder && holder.prototype].forEach(function (target) {
        if (!target) return;
        Object.getOwnPropertyNames(target).forEach(function (key) {
          var descriptor = Object.getOwnPropertyDescriptor(target, key);
          if (!descriptor || typeof descriptor.value !== 'function') return;
          try { Object.defineProperty(target, key, { writable: false, configurable: false }); } catch (e) {}
        });
      });
    }, Function('return this')());
  })();`, context);
}

function createAcrobat(options) {
  options = options || {};
  const fields = {};
  const log = [];
  const alerts = [];
  let focused = null, pageNum = 0, pendingPage = null;
  const display = { visible: 0, hidden: 1, noPrint: 2, noView: 3 };
  function field(name) {
    if (!fields[name]) fields[name] = {
      name, value: "", display: 0, readonly: false, items: null, caption: null, fillColor: null, textColor: null,
      setItems(items) { this.items = items.map(item => Array.isArray(item) ? item.slice() : [item, item]); this.value = ""; },
      buttonSetCaption(caption) { this.caption = String(caption); },
      buttonGetCaption() { return this.caption; },
      setFocus() { focused = name; }
    };
    return fields[name];
  }
  const doc = {
    delay: false,
    getField: name => (options.known && !options.known.has(name) ? null : field(name)),
    get pageNum() { return pageNum; },
    set pageNum(value) { pageNum = value; pendingPage = value; }
  };
  const app = {
    viewerType: "Reader", viewerVersion: 24,
    alert(options) { alerts.push(options); return options.nType === 2 ? (this.answer || 4) : 1; },
    menus: [], nextMenuChoice: null,
    popUpMenuEx() { const items = Array.prototype.slice.call(arguments); this.menus.push(items); const choice = this.nextMenuChoice; this.nextMenuChoice = null; return choice; },
    timers: [],
    setTimeOut(code) { const timer = { code }; this.timers.push(timer); return timer; },
    clearTimeOut(timer) { const index = this.timers.indexOf(timer); if (index >= 0) this.timers.splice(index, 1); }
  };
  const console = { println: message => log.push(String(message)), show() {}, clear() {} };
  return { doc, app, display, console, fields, log, alerts, field, get focused() { return focused; }, clearFocus() { focused = null; }, takePendingPage() { const page = pendingPage; pendingPage = null; return page; } };
}

function loadPdfScript(seed) {
  const script = fs.readFileSync(path.join(__dirname, "..", ".build", "document-script.js"), "utf8");
  const registry = JSON.parse(/export const REGISTRY = (.*);\n/.exec(fs.readFileSync(path.join(__dirname, "..", ".build", "site-data.js"), "utf8"))[1]);
  const acrobat = createAcrobat({ known: new Set(Object.keys(registry.fields)) });
  const sandbox = { app: acrobat.app, display: acrobat.display, console: acrobat.console, color: {}, util: {} };
  const context = vm.createContext(sandbox);
  stripToES5(context);
  if (process.env.BT_LOCK_BUILTINS !== "0") lockBuiltIns(context);
  vm.runInContext("Math.random = __random;", Object.assign(context, { __random: seededRandom(seed || 777) }));
  // Acrobat runs document scripts with `this` bound to the Doc.
  Object.defineProperty(context, "getField", { value: acrobat.doc.getField });
  Object.defineProperty(context, "pageNum", { get: () => acrobat.doc.pageNum, set: value => { acrobat.doc.pageNum = value; } });
  const started = Date.now();
  vm.runInContext(script, context, { filename: "document-script.js" });
  const bootMs = Date.now() - started;
  const api = vm.runInContext("({ commit: BT_commit, toggle: BT_toggle, press: BT_press, page: BT_page, bridge: BT1101 })", context);

  // Delivers events like Acrobat: Validate for text/combo, Mouse Up for buttons and checkboxes,
  // then Page Open for any page the script navigated to.
  // Runs Acrobat timers (deferred page opens) and follows any navigation they cause.
  const runTimers = () => { for (let guard = 0; guard < 20 && acrobat.app.timers.length; guard++) { const timer = acrobat.app.timers.shift(); vm.runInContext(timer.code, context); } };
  const followNavigation = () => { let page; runTimers(); while ((page = acrobat.takePendingPage()) !== null) { api.page(page); runTimers(); } };
  const session = {
    acrobat, registry, api, bootMs, context,
    page(index) { acrobat.doc.pageNum = index; acrobat.takePendingPage(); api.page(index); followNavigation(); },
    // Opens pages in quick succession, as scrolling does, without letting time pass between them.
    scrollThrough(indexes) { indexes.forEach(index => { acrobat.doc.pageNum = index; api.page(index); }); acrobat.takePendingPage(); followNavigation(); },
    // Acrobat hands a multi-line field's text to scripts with "\r" between lines.
    type(name, value) { value = String(value).replace(/\r?\n/g, "\r"); const event = { target: acrobat.fields[name] || acrobat.field(name), value, rc: true }; api.commit(event); acrobat.field(name).value = event.value; followNavigation(); },
    // Opens a dropdown's pop-up menu and picks the item with that return value ("@index").
    choose(name, exportValue) { const spec = registry.fields[name]; const picker = acrobat.field(spec && spec.p ? spec.p : name); if (picker.readonly) return false; acrobat.app.nextMenuChoice = exportValue; api.press({ target: picker }); followNavigation(); return true; },
    menuFor(name) { const spec = registry.fields[name]; acrobat.app.nextMenuChoice = null; api.press({ target: acrobat.field(spec.p) }); return acrobat.app.menus[acrobat.app.menus.length - 1] || []; },
    // Acrobat does not run actions on read-only buttons.
    click(name) { if (acrobat.field(name).readonly) return; api.press({ target: acrobat.field(name) }); followNavigation(); },
    check(name, on) { const target = acrobat.field(name); target.value = on ? "Yes" : "Off"; api.toggle({ target }); followNavigation(); },
    fieldsOn(pageIndex) { return registry.pages[pageIndex].fields.map(name => Object.assign({ spec: registry.fields[name] }, acrobat.fields[name] || { name })); },
    pageOf(id) { return registry.pages.findIndex(page => page.op === id); },
    current() { return acrobat.doc.pageNum; },
    errors() { return api.bridge ? api.bridge.errors.slice() : ["bridge did not start"]; }
  };
  return session;
}

module.exports = { loadPdfScript, createAcrobat, stripToES5 };
