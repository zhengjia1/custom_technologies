// The browser globals index.html expects, backed by minidom and a virtual clock.
// Timers never wait: drain() runs queued microtasks and timers in due-time order.

import * as dom from "./minidom.js";

function englishNumber(value, options) {
  options = options || {};
  if (!isFinite(value)) return String(value);
  const minimum = options.minimumFractionDigits !== undefined ? options.minimumFractionDigits : 0;
  const maximum = options.maximumFractionDigits !== undefined ? Math.max(options.maximumFractionDigits, minimum) : Math.max(3, minimum);
  const negative = value < 0 || (value === 0 && 1 / value < 0);
  let text = Math.abs(value).toFixed(maximum);
  if (maximum > minimum) {
    const parts = text.split(".");
    if (parts[1] !== undefined) {
      let fraction = parts[1];
      while (fraction.length > minimum && fraction[fraction.length - 1] === "0") fraction = fraction.slice(0, -1);
      text = fraction ? parts[0] + "." + fraction : parts[0];
    }
  }
  const parts = text.split(".");
  if (options.useGrouping !== false) parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  text = parts.join(".");
  return negative && Number(text.replace(/,/g, "")) !== 0 ? "-" + text : text;
}

export function installNumberLocale() {
  // Acrobat formats with the operating-system locale; the website's prose assumes en-US.
  // eslint-disable-next-line no-extend-native
  Number.prototype.toLocaleString = function (locales, options) { return englishNumber(Number(this), options); };
}

export function createBrowser(bodyHtml, tasks, hooks) {
  hooks = hooks || {};
  const report = hooks.reportError || (error => { throw error; });
  dom.setErrorReporter(report);
  dom.setMicrotaskScheduler(callback => tasks.micro.push(callback));

  const document = dom.createDocument(bodyHtml);
  document.documentElement.setAttribute("lang", "en");
  const window = new dom.EventTarget();

  function setTimeoutImpl(callback, delay) {
    const args = Array.prototype.slice.call(arguments, 2);
    const id = ++tasks.seq;
    if (typeof callback === "function") tasks.timers.push({ id, due: tasks.now + Math.max(0, Number(delay) || 0), seq: id, fn: callback, args, interval: 0 });
    return id;
  }
  function setIntervalImpl(callback, delay) {
    const id = setTimeoutImpl(callback, delay);
    const timer = tasks.timers[tasks.timers.length - 1];
    if (timer && timer.id === id) { timer.interval = Math.max(1, Number(delay) || 1); timer.runs = 0; }
    return id;
  }
  function clearTimeoutImpl(id) {
    for (let i = 0; i < tasks.timers.length; i++) if (tasks.timers[i].id === id) { tasks.timers.splice(i, 1); return; }
  }
  const requestAnimationFrame = callback => setTimeoutImpl(() => callback(tasks.now), 16);
  const queueMicrotask = callback => { tasks.micro.push(callback); };

  const clipboard = {
    writes: [],
    writeText(text) { clipboard.writes.push(String(text)); return Promise.resolve(); },
    readText() { return Promise.reject(new Error("Clipboard reading is unavailable in Acrobat")); },
    write() { return Promise.reject(new Error("Clipboard is unavailable")); }
  };
  const navigator = { clipboard, userAgent: "Mozilla/5.0 (Acrobat) BT1101PDF", platform: "Acrobat", language: "en-US", languages: ["en-US"], onLine: false, maxTouchPoints: 0 };
  const location = { hostname: "bt1101-pdf.invalid", host: "bt1101-pdf.invalid", href: "file:///BT1101.pdf", protocol: "file:", pathname: "/BT1101.pdf", search: "", hash: "", origin: "null", reload() {}, assign() {}, replace() {} };

  function makeStorage() {
    const values = {};
    return {
      getItem: key => Object.prototype.hasOwnProperty.call(values, key) ? values[key] : null,
      setItem: (key, value) => { values[key] = String(value); },
      removeItem: key => { delete values[key]; },
      clear: () => { Object.keys(values).forEach(key => { delete values[key]; }); },
      key: index => Object.keys(values)[index] || null,
      get length() { return Object.keys(values).length; }
    };
  }
  const localStorage = makeStorage();
  const sessionStorage = makeStorage();
  const matchMedia = media => ({ matches: false, media, onchange: null, addEventListener() {}, removeEventListener() {}, addListener() {}, removeListener() {} });
  const getComputedStyle = element => {
    const hidden = element && element.closest && element.closest(".hidden, [hidden]");
    return { display: hidden ? "none" : "block", visibility: "visible", getPropertyValue: name => (name === "display" ? (hidden ? "none" : "block") : "") };
  };
  const CSS = { escape: dom.cssEscape, supports: () => false };
  function Option(text, value, defaultSelected, selected) {
    const option = document.createElement("option");
    if (text !== undefined) option.text = String(text);
    if (value !== undefined) option.value = String(value);
    if (defaultSelected) option.defaultSelected = true;
    if (selected) option.selected = true;
    return option;
  }
  class Blob { constructor(parts, options) { this.parts = parts || []; this.type = (options && options.type) || ""; this.size = this.parts.join("").length; } text() { return Promise.resolve(this.parts.join("")); } }
  const URL = { createObjectURL: () => "blob:bt1101-pdf", revokeObjectURL() {} };
  const fetch = () => Promise.reject(new Error("Network access is unavailable in the PDF"));
  const noop = () => {};
  const consoleLog = [];
  const consoleObject = {};
  ["log", "info", "warn", "error", "debug", "trace"].forEach(level => {
    consoleObject[level] = function () { consoleLog.push(level + ": " + Array.prototype.slice.call(arguments).map(String).join(" ")); if (consoleLog.length > 200) consoleLog.shift(); };
  });
  class ResizeObserver { observe() {} unobserve() {} disconnect() {} }
  const performance = { now: () => tasks.now };
  const history = { pushState: noop, replaceState: noop, back: noop, state: null };

  const globals = {
    window, self: window, document, navigator, location, localStorage, sessionStorage, matchMedia,
    requestAnimationFrame, cancelAnimationFrame: clearTimeoutImpl,
    setTimeout: setTimeoutImpl, clearTimeout: clearTimeoutImpl, setInterval: setIntervalImpl, clearInterval: clearTimeoutImpl,
    queueMicrotask, CSS, MutationObserver: dom.MutationObserver, ResizeObserver, IntersectionObserver: ResizeObserver,
    Event: dom.Event, CustomEvent: dom.CustomEvent, KeyboardEvent: dom.KeyboardEvent, MouseEvent: dom.MouseEvent, InputEvent: dom.InputEvent, FocusEvent: dom.FocusEvent, UIEvent: dom.UIEvent,
    Node: dom.Node, Element: dom.Element, HTMLElement: dom.HTMLElement, HTMLButtonElement: dom.HTMLButtonElement,
    HTMLInputElement: dom.HTMLInputElement, HTMLSelectElement: dom.HTMLSelectElement, HTMLTextAreaElement: dom.HTMLTextAreaElement,
    HTMLOptionElement: dom.HTMLOptionElement, Text: dom.Text, DocumentFragment: dom.DocumentFragment, Option,
    Blob, URL, fetch, getComputedStyle, alert: noop, confirm: () => false, prompt: () => null,
    console: consoleObject, history, performance, crypto: undefined, scrollTo: noop, scrollBy: noop,
    innerWidth: 1280, innerHeight: 900, outerWidth: 1280, outerHeight: 900, devicePixelRatio: 1, screen: { width: 1280, height: 900 }
  };
  Object.keys(globals).forEach(key => { if (key !== "window") window[key] = globals[key]; });
  window._parentForEvents = () => null;
  document.defaultView = window;

  function drain(limit) {
    limit = limit || 20000;
    let steps = 0;
    for (;;) {
      while (tasks.micro.length) {
        const callback = tasks.micro.shift();
        try { callback(); } catch (error) { report(error); }
        if (++steps > limit) throw new Error("Task queue did not settle");
      }
      if (!tasks.timers.length) return;
      let next = 0;
      for (let i = 1; i < tasks.timers.length; i++) {
        const timer = tasks.timers[i], best = tasks.timers[next];
        if (timer.due < best.due || (timer.due === best.due && timer.seq < best.seq)) next = i;
      }
      const timer = tasks.timers[next];
      tasks.timers.splice(next, 1);
      tasks.now = Math.max(tasks.now, timer.due);
      if (timer.interval && timer.runs < 20) { timer.runs++; timer.due = tasks.now + timer.interval; timer.seq = ++tasks.seq; tasks.timers.push(timer); }
      try { timer.fn.apply(null, timer.args); } catch (error) { report(error); }
      if (++steps > limit) throw new Error("Task queue did not settle");
    }
  }

  return { globals, window, document, clipboard, drain, consoleLog };
}
