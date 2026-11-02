// A small, synchronous DOM used to run the website's script inside Adobe Acrobat,
// which has a JavaScript engine but no browser document. It implements only what
// index.html uses: parsing, selectors, events, form controls and mutation records.

const VOID_TAGS = { area: 1, base: 1, br: 1, col: 1, embed: 1, hr: 1, img: 1, input: 1, link: 1, meta: 1, source: 1, track: 1, wbr: 1 };
const RAW_TEXT_TAGS = { script: 1, style: 1 };
const RCDATA_TAGS = { textarea: 1, title: 1 };
const NAMED_ENTITIES = {
  amp: "&", lt: "<", gt: ">", quot: '"', apos: "'", nbsp: " ", middot: "·", times: "×", hellip: "…",
  ndash: "–", mdash: "—", rarr: "→", larr: "←", le: "≤", ge: "≥", ne: "≠", copy: "©", deg: "°", plusmn: "±",
  minus: "−", sup2: "²", micro: "µ", sigma: "σ", mu: "μ", rho: "ρ", alpha: "α", beta: "β", ldquo: "“", rdquo: "”",
  lsquo: "‘", rsquo: "’", bull: "•", para: "¶"
};
// Elements that never contain flow content implied by an unclosed sibling.
const AUTO_CLOSE = {
  p: { p: 1, div: 1, section: 1, article: 1, h1: 1, h2: 1, h3: 1, h4: 1, ul: 1, ol: 1, pre: 1, table: 1, header: 1, footer: 1 },
  li: { li: 1 }, option: { option: 1, optgroup: 1 }, optgroup: { optgroup: 1 },
  tr: { tr: 1 }, td: { td: 1, th: 1, tr: 1 }, th: { td: 1, th: 1, tr: 1 }
};

export function decodeEntities(text) {
  if (text.indexOf("&") < 0) return text;
  return text.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|[A-Za-z][A-Za-z0-9]*);/g, (whole, body) => {
    if (body[0] === "#") {
      const code = body[1] === "x" || body[1] === "X" ? parseInt(body.slice(2), 16) : parseInt(body.slice(1), 10);
      return isFinite(code) ? String.fromCodePoint(code) : whole;
    }
    return Object.prototype.hasOwnProperty.call(NAMED_ENTITIES, body) ? NAMED_ENTITIES[body] : whole;
  });
}

function escapeText(text) { return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/ /g, "&nbsp;"); }
function escapeAttribute(text) { return text.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/ /g, "&nbsp;"); }
// data-* names are converted on every attribute read; the website reuses a handful of them.
const camelCache = Object.create(null), dashCache = Object.create(null);
function camelData(name) { return camelCache[name] || (camelCache[name] = name.slice(5).replace(/-([a-z])/g, (_, letter) => letter.toUpperCase())); }
function dashData(key) { return dashCache[key] || (dashCache[key] = "data-" + key.replace(/[A-Z]/g, letter => "-" + letter.toLowerCase())); }

// ---------------------------------------------------------------- events

export class Event {
  constructor(type, init) {
    init = init || {};
    this.type = String(type);
    this.bubbles = Boolean(init.bubbles);
    this.cancelable = Boolean(init.cancelable);
    this.composed = Boolean(init.composed);
    this.defaultPrevented = false;
    this.target = null;
    this.currentTarget = null;
    this.eventPhase = 0;
    this.isTrusted = false;
    this.timeStamp = Date.now();
    this._stop = false;
    this._stopNow = false;
  }
  preventDefault() { if (this.cancelable) this.defaultPrevented = true; }
  stopPropagation() { this._stop = true; }
  stopImmediatePropagation() { this._stop = true; this._stopNow = true; }
  composedPath() { return this._path ? this._path.slice() : []; }
}
export class CustomEvent extends Event {
  constructor(type, init) { super(type, init); this.detail = init && init.detail !== undefined ? init.detail : null; }
}
export class UIEvent extends Event {
  constructor(type, init) { super(type, init); this.detail = init && init.detail !== undefined ? init.detail : 0; }
}
export class MouseEvent extends UIEvent {
  constructor(type, init) {
    super(type, init); init = init || {};
    this.clientX = init.clientX || 0; this.clientY = init.clientY || 0; this.button = init.button || 0;
    this.shiftKey = Boolean(init.shiftKey); this.ctrlKey = Boolean(init.ctrlKey); this.metaKey = Boolean(init.metaKey); this.altKey = Boolean(init.altKey);
  }
}
export class KeyboardEvent extends UIEvent {
  constructor(type, init) {
    super(type, init); init = init || {};
    this.key = init.key || ""; this.code = init.code || "";
    this.shiftKey = Boolean(init.shiftKey); this.ctrlKey = Boolean(init.ctrlKey); this.metaKey = Boolean(init.metaKey); this.altKey = Boolean(init.altKey);
    this.isComposing = false; this.repeat = false;
  }
}
export class InputEvent extends UIEvent {
  constructor(type, init) { super(type, init); init = init || {}; this.data = init.data || null; this.inputType = init.inputType || ""; }
}
export class FocusEvent extends UIEvent {}

export class EventTarget {
  addEventListener(type, listener, options) {
    if (!listener) return;
    const capture = typeof options === "boolean" ? options : Boolean(options && options.capture);
    const once = Boolean(options && typeof options === "object" && options.once);
    if (!this._listeners) this._listeners = {};
    const list = this._listeners[type] || (this._listeners[type] = []);
    for (let i = 0; i < list.length; i++) if (list[i].listener === listener && list[i].capture === capture) return;
    list.push({ listener, capture, once });
  }
  removeEventListener(type, listener, options) {
    const capture = typeof options === "boolean" ? options : Boolean(options && options.capture);
    const list = this._listeners && this._listeners[type];
    if (!list) return;
    for (let i = 0; i < list.length; i++) if (list[i].listener === listener && list[i].capture === capture) { list[i].removed = true; list.splice(i, 1); return; }
  }
  _invoke(event, phase) {
    const list = this._listeners && this._listeners[event.type];
    if (!list || !list.length) return;
    const snapshot = list.slice();
    event.currentTarget = this;
    for (let i = 0; i < snapshot.length; i++) {
      const entry = snapshot[i];
      if (entry.removed) continue;
      if (phase === 1 && !entry.capture) continue;
      if (phase === 3 && entry.capture) continue;
      if (entry.once) this.removeEventListener(event.type, entry.listener, entry.capture);
      const listener = entry.listener;
      try {
        if (typeof listener === "function") listener.call(this, event);
        else if (listener && typeof listener.handleEvent === "function") listener.handleEvent(event);
      } catch (error) {
        reportError(error);
      }
      if (event._stopNow) return;
    }
  }
  dispatchEvent(event) {
    const path = [];
    let node = this;
    while (node) { path.push(node); node = node._parentForEvents ? node._parentForEvents() : null; }
    event.target = this;
    event._path = path;
    event._stop = false; event._stopNow = false;
    for (let i = path.length - 1; i > 0 && !event._stop; i--) { event.eventPhase = 1; path[i]._invoke(event, 1); }
    if (!event._stop) {
      event.eventPhase = 2;
      // At the target, capturing listeners run before bubbling ones.
      this._invoke(event, 1);
      if (!event._stopNow) this._invoke(event, 3);
    }
    if (event.bubbles) for (let i = 1; i < path.length && !event._stop; i++) { event.eventPhase = 3; path[i]._invoke(event, 3); }
    event.eventPhase = 0;
    event.currentTarget = null;
    return !event.defaultPrevented;
  }
}

let errorReporter = null;
export function setErrorReporter(reporter) { errorReporter = reporter; }
function reportError(error) { if (errorReporter) errorReporter(error); else throw error; }

// ---------------------------------------------------------------- mutation observers

const observers = [];
let scheduleMicrotask = callback => callback();
export function setMicrotaskScheduler(scheduler) { scheduleMicrotask = scheduler; }

export class MutationObserver {
  constructor(callback) { this._callback = callback; this._targets = []; this._records = []; this._scheduled = false; }
  observe(target, options) { this._targets.push({ target, subtree: Boolean(options && options.subtree), childList: Boolean(options && options.childList) }); if (observers.indexOf(this) < 0) observers.push(this); }
  disconnect() { this._targets = []; this._records = []; const index = observers.indexOf(this); if (index >= 0) observers.splice(index, 1); }
  takeRecords() { const records = this._records; this._records = []; return records; }
  _queue(record) {
    this._records.push(record);
    if (this._scheduled) return;
    this._scheduled = true;
    scheduleMicrotask(() => {
      this._scheduled = false;
      const records = this.takeRecords();
      if (records.length) {
        try { this._callback(records, this); } catch (error) { reportError(error); }
      }
    });
  }
}

function queueChildListRecord(target, added, removed) {
  if (!observers.length) return;
  for (let i = 0; i < observers.length; i++) {
    const observer = observers[i];
    for (let j = 0; j < observer._targets.length; j++) {
      const entry = observer._targets[j];
      if (!entry.childList) continue;
      if (entry.target === target || (entry.subtree && entry.target.contains(target))) {
        observer._queue({ type: "childList", target, addedNodes: added, removedNodes: removed, previousSibling: null, nextSibling: null, attributeName: null });
        break;
      }
    }
  }
}

// ---------------------------------------------------------------- nodes

export class Node extends EventTarget {
  constructor(document) { super(); this.ownerDocument = document || null; this.parentNode = null; this.childNodes = []; }
  _parentForEvents() { return this.parentNode || (this.nodeType === 9 ? this.defaultView : null); }
  get firstChild() { return this.childNodes[0] || null; }
  get lastChild() { return this.childNodes[this.childNodes.length - 1] || null; }
  get nextSibling() { const p = this.parentNode; if (!p) return null; return p.childNodes[p.childNodes.indexOf(this) + 1] || null; }
  get previousSibling() { const p = this.parentNode; if (!p) return null; return p.childNodes[p.childNodes.indexOf(this) - 1] || null; }
  get parentElement() { return this.parentNode && this.parentNode.nodeType === 1 ? this.parentNode : null; }
  get isConnected() { let node = this; while (node.parentNode) node = node.parentNode; return node.nodeType === 9; }
  getRootNode() { let node = this; while (node.parentNode) node = node.parentNode; return node; }
  hasChildNodes() { return this.childNodes.length > 0; }
  contains(other) { while (other) { if (other === this) return true; other = other.parentNode; } return false; }
  get nodeValue() { return this.nodeType === 3 || this.nodeType === 8 ? this.data : null; }
  set nodeValue(value) { if (this.nodeType === 3 || this.nodeType === 8) this.data = String(value); }

  _insertNodes(nodes, reference) {
    const inserted = [];
    for (let i = 0; i < nodes.length; i++) {
      let node = nodes[i];
      if (typeof node === "string") node = (this.ownerDocument || this).createTextNode(node);
      if (node.nodeType === 11) { const children = node.childNodes.slice(); node.childNodes = []; children.forEach(child => { child.parentNode = null; inserted.push(child); }); continue; }
      if (node.parentNode) node.parentNode.removeChild(node);
      inserted.push(node);
    }
    let index = reference ? this.childNodes.indexOf(reference) : this.childNodes.length;
    if (index < 0) throw new Error("NotFoundError: reference node is not a child");
    for (let i = 0; i < inserted.length; i++) {
      inserted[i].parentNode = this;
      this.childNodes.splice(index++, 0, inserted[i]);
    }
    if (inserted.length) {
      const document = this.ownerDocument || this;
      if (document._ids && this.isConnected) for (let i = 0; i < inserted.length; i++) document._registerTree(inserted[i]);
      queueChildListRecord(this, inserted, []);
    }
    return inserted;
  }
  appendChild(node) { this._insertNodes([node], null); return node; }
  insertBefore(node, reference) { this._insertNodes([node], reference || null); return node; }
  removeChild(node) {
    const index = this.childNodes.indexOf(node);
    if (index < 0) throw new Error("NotFoundError: node is not a child");
    this.childNodes.splice(index, 1);
    node.parentNode = null;
    const document = this.ownerDocument || this;
    if (document._activeElement && node.contains(document._activeElement)) document._activeElement = null;
    queueChildListRecord(this, [], [node]);
    return node;
  }
  replaceChild(node, old) { this.insertBefore(node, old); this.removeChild(old); return old; }
  append() { this._insertNodes(Array.prototype.slice.call(arguments), null); }
  prepend() { this._insertNodes(Array.prototype.slice.call(arguments), this.firstChild); }
  replaceChildren() { this.childNodes.slice().forEach(child => this.removeChild(child)); this._insertNodes(Array.prototype.slice.call(arguments), null); }
  before() { if (this.parentNode) this.parentNode._insertNodes(Array.prototype.slice.call(arguments), this); }
  after() { if (this.parentNode) this.parentNode._insertNodes(Array.prototype.slice.call(arguments), this.nextSibling); }
  replaceWith() { const parent = this.parentNode; if (!parent) return; parent._insertNodes(Array.prototype.slice.call(arguments), this); parent.removeChild(this); }
  remove() { if (this.parentNode) this.parentNode.removeChild(this); }

  get textContent() {
    if (this.nodeType === 3 || this.nodeType === 8) return this.data;
    if (this.nodeType === 9) return null;
    let text = "";
    const walk = node => { for (let i = 0; i < node.childNodes.length; i++) { const child = node.childNodes[i]; if (child.nodeType === 3) text += child.data; else if (child.nodeType === 1 || child.nodeType === 11) walk(child); } };
    walk(this);
    return text;
  }
  set textContent(value) {
    if (this.nodeType === 3 || this.nodeType === 8) { this.data = String(value); return; }
    this.childNodes.slice().forEach(child => this.removeChild(child));
    const text = value === null || value === undefined ? "" : String(value);
    if (text) this.appendChild((this.ownerDocument || this).createTextNode(text));
  }
}
Node.ELEMENT_NODE = 1; Node.TEXT_NODE = 3; Node.COMMENT_NODE = 8; Node.DOCUMENT_NODE = 9; Node.DOCUMENT_FRAGMENT_NODE = 11;

export class CharacterData extends Node {
  constructor(document, data) { super(document); this.data = String(data); }
  get length() { return this.data.length; }
  cloneNode() { return this.nodeType === 3 ? this.ownerDocument.createTextNode(this.data) : this.ownerDocument.createComment(this.data); }
}
export class Text extends CharacterData { get nodeType() { return 3; } get nodeName() { return "#text"; } get wholeText() { return this.data; } }
export class Comment extends CharacterData { get nodeType() { return 8; } get nodeName() { return "#comment"; } }

export class DocumentFragment extends Node {
  get nodeType() { return 11; }
  get nodeName() { return "#document-fragment"; }
  get children() { return this.childNodes.filter(node => node.nodeType === 1); }
  get firstElementChild() { return this.children[0] || null; }
  querySelector(selector) { return querySelectorFrom(this, selector, true)[0] || null; }
  querySelectorAll(selector) { return querySelectorFrom(this, selector, false); }
  getElementById(id) { return querySelectorFrom(this, "#" + cssEscape(id), true)[0] || null; }
  cloneNode(deep) { const copy = this.ownerDocument.createDocumentFragment(); if (deep) this.childNodes.forEach(child => copy.appendChild(child.cloneNode(true))); return copy; }
}

// ---------------------------------------------------------------- token lists and style

class DOMTokenList {
  constructor(element, attribute) { this._element = element; this._attribute = attribute; }
  _tokens() {
    if (this._attribute === "class") return this._element._classTokens().slice();
    const value = this._element.getAttribute(this._attribute); return value ? value.split(/\s+/).filter(Boolean) : [];
  }
  _write(tokens) { this._element.setAttribute(this._attribute, tokens.join(" ")); }
  get length() { return this._tokens().length; }
  get value() { return this._element.getAttribute(this._attribute) || ""; }
  set value(text) { this._element.setAttribute(this._attribute, text); }
  item(index) { return this._tokens()[index] || null; }
  contains(token) { return (this._attribute === "class" ? this._element._classTokens() : this._tokens()).indexOf(token) >= 0; }
  add() { const tokens = this._tokens(); for (let i = 0; i < arguments.length; i++) if (tokens.indexOf(arguments[i]) < 0) tokens.push(arguments[i]); this._write(tokens); }
  remove() { const remove = Array.prototype.slice.call(arguments); const tokens = this._tokens(); const kept = tokens.filter(token => remove.indexOf(token) < 0); if (kept.length !== tokens.length || this._element.hasAttribute(this._attribute)) this._write(kept); }
  toggle(token, force) {
    const has = this.contains(token);
    const want = force === undefined ? !has : Boolean(force);
    if (want && !has) this.add(token); else if (!want && has) this.remove(token);
    return want;
  }
  replace(oldToken, newToken) { const tokens = this._tokens(); const index = tokens.indexOf(oldToken); if (index < 0) return false; tokens[index] = newToken; this._write(tokens); return true; }
  forEach(callback, thisArg) { this._tokens().forEach(callback, thisArg); }
  toString() { return this.value; }
}

class CSSStyleDeclaration {
  setProperty(name, value) { this[camelCss(name)] = String(value); }
  getPropertyValue(name) { const value = this[camelCss(name)]; return typeof value === "string" ? value : ""; }
  removeProperty(name) { const key = camelCss(name), old = this.getPropertyValue(name); delete this[key]; return old; }
  get cssText() { return this._cssText || ""; }
  set cssText(text) { this._cssText = String(text); String(text).split(";").forEach(part => { const at = part.indexOf(":"); if (at > 0) this.setProperty(part.slice(0, at).trim(), part.slice(at + 1).trim()); }); }
}
function camelCss(name) { return name.indexOf("--") === 0 ? name : name.replace(/-([a-z])/g, (_, letter) => letter.toUpperCase()); }

// ---------------------------------------------------------------- elements

export class Element extends Node {
  constructor(document, tagName) {
    super(document);
    this.localName = tagName;
    this._attributes = [];
    // Lookup map and parsed class list: selectors read these constantly.
    this._values = {};
    this._classes = null;
    this.dataset = {};
    this.style = new CSSStyleDeclaration();
  }
  get nodeType() { return 1; }
  get tagName() { return this.localName.toUpperCase(); }
  get nodeName() { return this.tagName; }

  // data-* attributes live in dataset so plain property writes stay visible to selectors.
  getAttribute(name) {
    name = String(name).toLowerCase();
    if (name.lastIndexOf("data-", 0) === 0) { const key = camelData(name); return Object.prototype.hasOwnProperty.call(this.dataset, key) && this.dataset[key] !== undefined ? String(this.dataset[key]) : null; }
    const value = this._values[name];
    return value === undefined ? null : value;
  }
  setAttribute(name, value) {
    name = String(name).toLowerCase(); value = String(value);
    if (name.lastIndexOf("data-", 0) === 0) { this.dataset[camelData(name)] = value; return; }
    if (this._values[name] === undefined) this._attributes.push({ name, value });
    else for (let i = 0; i < this._attributes.length; i++) if (this._attributes[i].name === name) this._attributes[i].value = value;
    this._values[name] = value;
    this._changed(name);
  }
  removeAttribute(name) {
    name = String(name).toLowerCase();
    if (name.lastIndexOf("data-", 0) === 0) { delete this.dataset[camelData(name)]; return; }
    if (this._values[name] === undefined) return;
    delete this._values[name];
    for (let i = 0; i < this._attributes.length; i++) if (this._attributes[i].name === name) { this._attributes.splice(i, 1); break; }
    this._changed(name);
  }
  _changed(name) {
    if (name === "class") this._classes = null;
    if (name === "id" && this._values.id !== undefined && this.ownerDocument && this.isConnected) this.ownerDocument._registerId(this);
    this._attributeChanged(name);
  }
  // Class tokens, parsed once per change of the class attribute.
  _classTokens() {
    if (!this._classes) { const value = this._values["class"]; this._classes = value ? value.split(/\s+/).filter(Boolean) : []; }
    return this._classes;
  }
  hasAttribute(name) { return this.getAttribute(name) !== null; }
  toggleAttribute(name, force) { const has = this.hasAttribute(name), want = force === undefined ? !has : Boolean(force); if (want && !has) this.setAttribute(name, ""); else if (!want && has) this.removeAttribute(name); return want; }
  getAttributeNames() { return this._allAttributes().map(attribute => attribute.name); }
  _allAttributes() {
    const list = this._attributes.slice();
    Object.keys(this.dataset).forEach(key => { if (this.dataset[key] !== undefined) list.push({ name: dashData(key), value: String(this.dataset[key]) }); });
    return list;
  }
  get attributes() { return this._allAttributes(); }
  _attributeChanged() {}

  get id() { return this.getAttribute("id") || ""; }
  set id(value) { this.setAttribute("id", value); }
  get className() { return this.getAttribute("class") || ""; }
  set className(value) { this.setAttribute("class", value); }
  get classList() { return this._classList || (this._classList = new DOMTokenList(this, "class")); }
  get title() { return this.getAttribute("title") || ""; }
  set title(value) { this.setAttribute("title", value); }
  get hidden() { return this.hasAttribute("hidden"); }
  set hidden(value) { this.toggleAttribute("hidden", Boolean(value)); }
  get role() { return this.getAttribute("role"); }
  get tabIndex() { const value = this.getAttribute("tabindex"); return value === null ? -1 : Number(value); }
  set tabIndex(value) { this.setAttribute("tabindex", value); }
  get lang() { return this.getAttribute("lang") || ""; }
  get dir() { return this.getAttribute("dir") || ""; }
  get href() { return this.getAttribute("href") || ""; }
  set href(value) { this.setAttribute("href", value); }
  get download() { return this.getAttribute("download") || ""; }
  set download(value) { this.setAttribute("download", value); }
  get htmlFor() { return this.getAttribute("for") || ""; }
  set htmlFor(value) { this.setAttribute("for", value); }
  get open() { return this.hasAttribute("open"); }
  set open(value) { this.toggleAttribute("open", Boolean(value)); }

  get children() { return this.childNodes.filter(node => node.nodeType === 1); }
  get childElementCount() { return this.children.length; }
  get firstElementChild() { for (let i = 0; i < this.childNodes.length; i++) if (this.childNodes[i].nodeType === 1) return this.childNodes[i]; return null; }
  get lastElementChild() { for (let i = this.childNodes.length - 1; i >= 0; i--) if (this.childNodes[i].nodeType === 1) return this.childNodes[i]; return null; }
  get nextElementSibling() { let node = this.nextSibling; while (node && node.nodeType !== 1) node = node.nextSibling; return node; }
  get previousElementSibling() { let node = this.previousSibling; while (node && node.nodeType !== 1) node = node.previousSibling; return node; }

  get innerHTML() { return this.childNodes.map(serializeNode).join(""); }
  set innerHTML(html) {
    this.childNodes.slice().forEach(child => this.removeChild(child));
    const fragment = parseFragment(this.ownerDocument, String(html), this.localName);
    this.appendChild(fragment);
  }
  get outerHTML() { return serializeNode(this); }
  set outerHTML(html) { const fragment = parseFragment(this.ownerDocument, String(html), this.parentNode ? this.parentNode.localName : "div"); this.replaceWith(fragment); }
  get innerText() { return this.textContent; }
  set innerText(value) { this.textContent = value; }
  insertAdjacentHTML(position, html) {
    const context = position === "beforebegin" || position === "afterend" ? (this.parentNode && this.parentNode.localName) || "div" : this.localName;
    this.insertAdjacentElement(position, parseFragment(this.ownerDocument, String(html), context));
  }
  insertAdjacentElement(position, node) {
    position = String(position).toLowerCase();
    if (position === "beforebegin") { if (this.parentNode) this.parentNode.insertBefore(node, this); }
    else if (position === "afterbegin") this.insertBefore(node, this.firstChild);
    else if (position === "beforeend") this.appendChild(node);
    else if (position === "afterend") { if (this.parentNode) this.parentNode.insertBefore(node, this.nextSibling); }
    return node;
  }
  insertAdjacentText(position, text) { this.insertAdjacentElement(position, this.ownerDocument.createTextNode(text)); }

  querySelector(selector) { return querySelectorFrom(this, selector, true)[0] || null; }
  querySelectorAll(selector) { return querySelectorFrom(this, selector, false); }
  getElementsByTagName(tag) { return querySelectorFrom(this, tag, false); }
  getElementsByClassName(names) { return querySelectorFrom(this, names.split(/\s+/).filter(Boolean).map(name => "." + cssEscape(name)).join(""), false); }
  matches(selector) { return matchesSelectorList(this, parseSelector(selector), this); }
  closest(selector) { const list = parseSelector(selector); let node = this; while (node && node.nodeType === 1) { if (matchesSelectorList(node, list, node)) return node; node = node.parentNode; } return null; }
  webkitMatchesSelector(selector) { return this.matches(selector); }

  cloneNode(deep) {
    const copy = this.ownerDocument.createElement(this.localName);
    this._attributes.forEach(attribute => copy.setAttribute(attribute.name, attribute.value));
    Object.keys(this.dataset).forEach(key => { copy.dataset[key] = this.dataset[key]; });
    this._cloneState(copy);
    if (deep) this.childNodes.forEach(child => copy.appendChild(child.cloneNode(true)));
    return copy;
  }
  _cloneState() {}

  // Layout is not modelled; geometry reads return empty boxes.
  getBoundingClientRect() { return { x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0 }; }
  getClientRects() { return []; }
  get offsetWidth() { return 0; } get offsetHeight() { return 0; } get offsetTop() { return 0; } get offsetLeft() { return 0; } get offsetParent() { return null; }
  get clientWidth() { return 0; } get clientHeight() { return 0; }
  get scrollHeight() { return 0; } get scrollWidth() { return 0; }
  scrollIntoView() {} scrollTo() {} scrollBy() {}
  focus() { if (this.ownerDocument) this.ownerDocument._activeElement = this; }
  blur() { if (this.ownerDocument && this.ownerDocument._activeElement === this) this.ownerDocument._activeElement = null; }
  animate() { return { finished: Promise.resolve(), cancel() {}, finish() {}, onfinish: null }; }
  get isContentEditable() { return false; }

  click() {
    if (this.disabled) return;
    const event = new MouseEvent("click", { bubbles: true, cancelable: true, composed: true });
    const notCancelled = this.dispatchEvent(event);
    if (notCancelled) this._activationBehaviour();
  }
  _activationBehaviour() {
    if (this.localName === "summary" && this.parentNode && this.parentNode.localName === "details") this.parentNode.open = !this.parentNode.open;
    else if (this.localName === "label") {
      const control = this.htmlFor ? this.ownerDocument.getElementById(this.htmlFor) : this.querySelector("input, select, textarea, button");
      if (control && control !== this && typeof control.click === "function") control.click();
    }
  }
}

export class HTMLElement extends Element {}
export class HTMLDivElement extends HTMLElement {}
export class HTMLSpanElement extends HTMLElement {}
export class HTMLLabelElement extends HTMLElement {
  get control() { return this.htmlFor ? this.ownerDocument.getElementById(this.htmlFor) : this.querySelector("input, select, textarea, button"); }
}
export class HTMLAnchorElement extends HTMLElement {}
export class HTMLDetailsElement extends HTMLElement {}

function booleanReflect(Class, property, attribute) {
  Object.defineProperty(Class.prototype, property, {
    configurable: true,
    get() { return this.hasAttribute(attribute); },
    set(value) { this.toggleAttribute(attribute, Boolean(value)); }
  });
}
function stringReflect(Class, property, attribute, fallback) {
  Object.defineProperty(Class.prototype, property, {
    configurable: true,
    get() { const value = this.getAttribute(attribute); return value === null ? (fallback || "") : value; },
    set(value) { this.setAttribute(attribute, value); }
  });
}

export class HTMLButtonElement extends HTMLElement {
  get type() { const type = (this.getAttribute("type") || "submit").toLowerCase(); return type === "button" || type === "reset" ? type : "submit"; }
  set type(value) { this.setAttribute("type", value); }
  get value() { return this.getAttribute("value") || ""; }
  set value(value) { this.setAttribute("value", value); }
  get form() { return this.closest("form"); }
}
booleanReflect(HTMLButtonElement, "disabled", "disabled");
stringReflect(HTMLButtonElement, "name", "name");

export class HTMLInputElement extends HTMLElement {
  constructor(document, tagName) { super(document, tagName); this._value = null; this._checked = null; this.selectionStart = 0; this.selectionEnd = 0; }
  get type() { const type = (this.getAttribute("type") || "text").toLowerCase(); return type; }
  set type(value) { this.setAttribute("type", value); }
  get value() {
    if (this.type === "checkbox" || this.type === "radio") { const value = this.getAttribute("value"); return value === null ? "on" : value; }
    return this._value !== null ? this._value : (this.getAttribute("value") || "");
  }
  set value(value) {
    if (this.type === "checkbox" || this.type === "radio") { this.setAttribute("value", value); return; }
    this._value = value === null || value === undefined ? "" : String(value);
  }
  get defaultValue() { return this.getAttribute("value") || ""; }
  set defaultValue(value) { this.setAttribute("value", value); }
  get valueAsNumber() { const value = this.value; return value === "" ? NaN : Number(value); }
  set valueAsNumber(number) { this.value = String(number); }
  get checked() { return this._checked !== null ? this._checked : this.hasAttribute("checked"); }
  set checked(value) {
    this._checked = Boolean(value);
    if (this._checked && this.type === "radio" && this.name) {
      const root = this.closest("form") || this.ownerDocument;
      root.querySelectorAll('input[type="radio"]').forEach(radio => { if (radio !== this && radio.name === this.name) radio._checked = false; });
    }
  }
  get defaultChecked() { return this.hasAttribute("checked"); }
  set defaultChecked(value) { this.toggleAttribute("checked", Boolean(value)); }
  get form() { return this.closest("form"); }
  get labels() { const labels = []; const parent = this.closest("label"); if (parent) labels.push(parent); return labels; }
  select() {} setSelectionRange(start, end) { this.selectionStart = start; this.selectionEnd = end; } setRangeText() {}
  checkValidity() { return true; } reportValidity() { return true; } setCustomValidity() {}
  get validity() { return { valid: true, badInput: false, valueMissing: false }; }
  stepUp(n) { const step = Number(this.getAttribute("step")) || 1; this.value = String((Number(this.value) || 0) + step * (n || 1)); }
  stepDown(n) { this.stepUp(-(n || 1)); }
  _cloneState(copy) { copy._value = this._value; copy._checked = this._checked; }
  _activationBehaviour() {}
  click() {
    if (this.disabled) return;
    const toggles = this.type === "checkbox" || this.type === "radio";
    const before = this.checked;
    if (this.type === "checkbox") this.checked = !before;
    else if (this.type === "radio") this.checked = true;
    const event = new MouseEvent("click", { bubbles: true, cancelable: true, composed: true });
    const notCancelled = this.dispatchEvent(event);
    if (!toggles) return;
    if (!notCancelled) { this.checked = before; return; }
    if (this.checked !== before) {
      this.dispatchEvent(new InputEvent("input", { bubbles: true, composed: true }));
      this.dispatchEvent(new Event("change", { bubbles: true }));
    }
  }
}
booleanReflect(HTMLInputElement, "disabled", "disabled");
booleanReflect(HTMLInputElement, "readOnly", "readonly");
booleanReflect(HTMLInputElement, "required", "required");
stringReflect(HTMLInputElement, "name", "name");
stringReflect(HTMLInputElement, "placeholder", "placeholder");
stringReflect(HTMLInputElement, "min", "min");
stringReflect(HTMLInputElement, "max", "max");
stringReflect(HTMLInputElement, "step", "step");
stringReflect(HTMLInputElement, "autocomplete", "autocomplete");
stringReflect(HTMLInputElement, "inputMode", "inputmode");

export class HTMLTextAreaElement extends HTMLElement {
  constructor(document, tagName) { super(document, tagName); this._value = null; this.selectionStart = 0; this.selectionEnd = 0; }
  get type() { return "textarea"; }
  get value() { return this._value !== null ? this._value : this.textContent; }
  set value(value) { this._value = value === null || value === undefined ? "" : String(value); }
  get defaultValue() { return this.textContent; }
  set defaultValue(value) { this.textContent = value; }
  get form() { return this.closest("form"); }
  select() {} setSelectionRange(start, end) { this.selectionStart = start; this.selectionEnd = end; } setRangeText() {}
  checkValidity() { return true; }
  _cloneState(copy) { copy._value = this._value; }
}
booleanReflect(HTMLTextAreaElement, "disabled", "disabled");
booleanReflect(HTMLTextAreaElement, "readOnly", "readonly");
stringReflect(HTMLTextAreaElement, "name", "name");
stringReflect(HTMLTextAreaElement, "placeholder", "placeholder");
Object.defineProperty(HTMLTextAreaElement.prototype, "rows", { get() { return Number(this.getAttribute("rows")) || 2; }, set(value) { this.setAttribute("rows", value); } });

export class HTMLOptionElement extends HTMLElement {
  constructor(document, tagName) { super(document, tagName); this._selected = null; }
  get value() { const value = this.getAttribute("value"); return value === null ? this.text : value; }
  set value(value) { this.setAttribute("value", value); }
  get text() { return this.textContent.replace(/\s+/g, " ").trim(); }
  set text(value) { this.textContent = value; }
  get label() { const label = this.getAttribute("label"); return label === null ? this.text : label; }
  set label(value) { this.setAttribute("label", value); }
  get defaultSelected() { return this.hasAttribute("selected"); }
  set defaultSelected(value) { this.toggleAttribute("selected", Boolean(value)); }
  get selected() {
    const select = this._select();
    if (select) return select._selectedOption() === this;
    return this._selected !== null ? this._selected : this.defaultSelected;
  }
  set selected(value) {
    const select = this._select();
    if (select && value) select._options().forEach(option => { option._selected = option === this; });
    else this._selected = Boolean(value);
    if (select) select._dirty = true;
  }
  get index() { const select = this._select(); return select ? select._options().indexOf(this) : 0; }
  get form() { const select = this._select(); return select ? select.form : null; }
  _select() { let node = this.parentNode; while (node && node.nodeType === 1) { if (node.localName === "select") return node; if (node.localName !== "optgroup") return null; node = node.parentNode; } return null; }
  _cloneState(copy) { copy._selected = this._selected; }
}
booleanReflect(HTMLOptionElement, "disabled", "disabled");

export class HTMLSelectElement extends HTMLElement {
  constructor(document, tagName) { super(document, tagName); this._dirty = false; }
  get type() { return this.multiple ? "select-multiple" : "select-one"; }
  _options() { return querySelectorFrom(this, "option", false).filter(option => option._select() === this); }
  // Single-select selectedness: the last explicitly selected option wins, else the first enabled option.
  _selectedOption() {
    const options = this._options();
    let chosen = null;
    for (let i = 0; i < options.length; i++) {
      const option = options[i];
      const selected = option._selected !== null ? option._selected : option.defaultSelected;
      if (selected) chosen = option;
    }
    if (chosen) return chosen;
    if (this._dirty) return null;
    for (let i = 0; i < options.length; i++) if (!options[i].disabled) return options[i];
    return null;
  }
  get options() {
    const list = this._options();
    list.add = (option, before) => this.add(option, before);
    list.remove = index => this.remove(index);
    list.item = index => list[index] || null;
    Object.defineProperty(list, "selectedIndex", { get: () => this.selectedIndex, set: value => { this.selectedIndex = value; } });
    return list;
  }
  get length() { return this._options().length; }
  get selectedOptions() { const option = this._selectedOption(); return option ? [option] : []; }
  get selectedIndex() { const option = this._selectedOption(); return option ? this._options().indexOf(option) : -1; }
  set selectedIndex(index) { const options = this._options(); options.forEach((option, i) => { option._selected = i === Number(index); }); this._dirty = true; }
  get value() { const option = this._selectedOption(); return option ? option.value : ""; }
  set value(value) {
    value = String(value);
    const options = this._options();
    let found = false;
    options.forEach(option => { const match = !found && option.value === value; option._selected = match; if (match) found = true; });
    this._dirty = true;
  }
  item(index) { return this._options()[index] || null; }
  namedItem(name) { return this._options().filter(option => option.id === name || option.getAttribute("name") === name)[0] || null; }
  add(option, before) {
    if (typeof before === "number") before = this._options()[before] || null;
    if (before && before.parentNode) before.parentNode.insertBefore(option, before);
    else this.appendChild(option);
  }
  remove(index) {
    if (index === undefined) { Element.prototype.remove.call(this); return; }
    const option = this._options()[index];
    if (option) option.remove();
  }
  get form() { return this.closest("form"); }
  checkValidity() { return true; }
  _cloneState(copy) { copy._dirty = this._dirty; }
}
booleanReflect(HTMLSelectElement, "disabled", "disabled");
booleanReflect(HTMLSelectElement, "multiple", "multiple");
booleanReflect(HTMLSelectElement, "required", "required");
stringReflect(HTMLSelectElement, "name", "name");

export class HTMLOptGroupElement extends HTMLElement {}
booleanReflect(HTMLOptGroupElement, "disabled", "disabled");
stringReflect(HTMLOptGroupElement, "label", "label");

const ELEMENT_CLASSES = {
  button: HTMLButtonElement, input: HTMLInputElement, textarea: HTMLTextAreaElement, select: HTMLSelectElement,
  option: HTMLOptionElement, optgroup: HTMLOptGroupElement, div: HTMLDivElement, span: HTMLSpanElement,
  label: HTMLLabelElement, a: HTMLAnchorElement, details: HTMLDetailsElement
};

// ---------------------------------------------------------------- document

export class Document extends Node {
  constructor() {
    super(null);
    // id -> element, filled as elements join the document; checked on every read.
    this._ids = {};
    this._idMisses = {};
    this._idVersion = 0;
    this._activeElement = null;
    this.readyState = "complete";
    this.visibilityState = "visible";
    this.hidden = false;
    this.defaultView = null;
    this.cookie = "";
    this.documentElement = this.createElement("html");
    this.appendChild(this.documentElement);
    this.head = this.createElement("head");
    this.body = this.createElement("body");
    this.documentElement.appendChild(this.head);
    this.documentElement.appendChild(this.body);
  }
  get nodeType() { return 9; }
  get nodeName() { return "#document"; }
  get activeElement() { return this._activeElement && this._activeElement.isConnected ? this._activeElement : this.body; }
  get title() { const title = this.querySelector("title"); return title ? title.textContent : ""; }
  get children() { return [this.documentElement]; }
  createElement(tagName) {
    const tag = String(tagName).toLowerCase();
    const Class = ELEMENT_CLASSES[tag] || HTMLElement;
    return new Class(this, tag);
  }
  createElementNS(_, tagName) { return this.createElement(tagName); }
  createTextNode(text) { return new Text(this, text); }
  createComment(text) { return new Comment(this, text); }
  createDocumentFragment() { const fragment = new DocumentFragment(this); return fragment; }
  createEvent() { return new Event(""); }
  getElementById(id) { return this._lookupId(String(id)); }
  querySelector(selector) { return querySelectorFrom(this, selector, true)[0] || null; }
  querySelectorAll(selector) { return querySelectorFrom(this, selector, false); }
  getElementsByTagName(tag) { return querySelectorFrom(this, tag, false); }
  getElementsByClassName(names) { return this.documentElement.getElementsByClassName(names); }
  _registerId(element) {
    const current = this._ids[element._values.id];
    if (!current || current === element || !current.isConnected || current._values.id !== element._values.id) this._ids[element._values.id] = element;
    this._idVersion++;
  }
  _registerTree(node) {
    if (node.nodeType !== 1) return;
    if (node._values.id !== undefined) this._registerId(node);
    for (let i = 0; i < node.childNodes.length; i++) this._registerTree(node.childNodes[i]);
  }
  _lookupId(id) {
    const found = this._ids[id];
    if (found && found._values.id === id && found.isConnected) return found;
    if (this._idMisses[id] === this._idVersion) return null;
    // Stale or missing: rebuild from the tree once, then remember a miss until something changes.
    this._ids = {};
    const walk = node => {
      for (let i = 0; i < node.childNodes.length; i++) {
        const child = node.childNodes[i];
        if (child.nodeType !== 1) continue;
        if (child._values.id !== undefined && !this._ids[child._values.id]) this._ids[child._values.id] = child;
        walk(child);
      }
    };
    walk(this);
    const result = this._ids[id] || null;
    if (!result) this._idMisses[id] = this._idVersion;
    return result;
  }
  hasFocus() { return true; }
  execCommand() { return false; }
  elementFromPoint() { return null; }
  getSelection() { return { toString() { return ""; }, removeAllRanges() {}, addRange() {}, rangeCount: 0 }; }
}

// ---------------------------------------------------------------- HTML parsing

export function parseFragment(document, html, contextTag) {
  const fragment = document.createDocumentFragment();
  buildTree(document, fragment, html, contextTag);
  return fragment;
}

// Parses body markup into an existing container. The website's markup is well formed,
// so this keeps to the HTML rules that matter for it: void, raw-text and implied end tags.
export function buildTree(document, root, html) {
  const stack = [root];
  const current = () => stack[stack.length - 1];
  let index = 0;
  const length = html.length;
  const appendText = text => { if (text) current().appendChild(document.createTextNode(decodeEntities(text))); };
  while (index < length) {
    const open = html.indexOf("<", index);
    if (open < 0) { appendText(html.slice(index)); break; }
    if (open > index) appendText(html.slice(index, open));
    if (html.startsWith("<!--", open)) {
      const end = html.indexOf("-->", open + 4);
      const stop = end < 0 ? length : end;
      current().appendChild(document.createComment(html.slice(open + 4, stop)));
      index = end < 0 ? length : end + 3;
      continue;
    }
    if (html[open + 1] === "!" || html[open + 1] === "?") { const end = html.indexOf(">", open); index = end < 0 ? length : end + 1; continue; }
    if (html[open + 1] === "/") {
      const end = html.indexOf(">", open);
      const name = html.slice(open + 2, end < 0 ? length : end).trim().toLowerCase();
      index = end < 0 ? length : end + 1;
      for (let i = stack.length - 1; i > 0; i--) {
        if (stack[i].localName === name) { stack.length = i; break; }
      }
      continue;
    }
    const tagMatch = /^<([A-Za-z][A-Za-z0-9-]*)/.exec(html.slice(open, open + 64));
    if (!tagMatch) { appendText("<"); index = open + 1; continue; }
    const name = tagMatch[1].toLowerCase();
    let cursor = open + tagMatch[0].length;
    const attributes = [];
    let selfClosing = false;
    while (cursor < length) {
      while (cursor < length && /\s/.test(html[cursor])) cursor++;
      if (html[cursor] === ">") { cursor++; break; }
      if (html[cursor] === "/" && html[cursor + 1] === ">") { selfClosing = true; cursor += 2; break; }
      const attrMatch = /^[^\s"'>\/=]+/.exec(html.slice(cursor, cursor + 256));
      if (!attrMatch) { cursor++; continue; }
      const attrName = attrMatch[0].toLowerCase();
      cursor += attrMatch[0].length;
      while (cursor < length && /\s/.test(html[cursor])) cursor++;
      let value = "";
      if (html[cursor] === "=") {
        cursor++;
        while (cursor < length && /\s/.test(html[cursor])) cursor++;
        const quote = html[cursor];
        if (quote === '"' || quote === "'") {
          const close = html.indexOf(quote, cursor + 1);
          value = html.slice(cursor + 1, close < 0 ? length : close);
          cursor = close < 0 ? length : close + 1;
        } else {
          const unquoted = /^[^\s>]+/.exec(html.slice(cursor, cursor + 2048));
          value = unquoted ? unquoted[0] : "";
          cursor += value.length;
        }
      }
      attributes.push([attrName, decodeEntities(value)]);
    }
    index = cursor;
    // Implied end tags for the few elements that need them.
    const top = current();
    if (top.nodeType === 1 && AUTO_CLOSE[top.localName] && AUTO_CLOSE[top.localName][name]) stack.pop();
    if (name === "option" || name === "optgroup") { const parent = current(); if (parent.nodeType === 1 && parent.localName === "option") stack.pop(); }
    const element = document.createElement(name);
    attributes.forEach(([attrName, value]) => { if (!element.hasAttribute(attrName)) element.setAttribute(attrName, value); });
    current().appendChild(element);
    if (VOID_TAGS[name] || selfClosing) continue;
    if (RAW_TEXT_TAGS[name] || RCDATA_TAGS[name]) {
      const closeTag = "</" + name;
      const close = html.toLowerCase().indexOf(closeTag, index);
      const text = html.slice(index, close < 0 ? length : close);
      if (text) element.appendChild(document.createTextNode(RCDATA_TAGS[name] ? decodeEntities(text.replace(/^\n/, "")) : text));
      const end = close < 0 ? length : html.indexOf(">", close);
      index = end < 0 ? length : end + 1;
      continue;
    }
    stack.push(element);
  }
}

function serializeNode(node) {
  if (node.nodeType === 3) {
    const parent = node.parentNode;
    return parent && parent.nodeType === 1 && RAW_TEXT_TAGS[parent.localName] ? node.data : escapeText(node.data);
  }
  if (node.nodeType === 8) return "<!--" + node.data + "-->";
  if (node.nodeType === 11) return node.childNodes.map(serializeNode).join("");
  if (node.nodeType !== 1) return "";
  let html = "<" + node.localName;
  node._allAttributes().forEach(attribute => { html += " " + attribute.name + '="' + escapeAttribute(attribute.value) + '"'; });
  html += ">";
  if (VOID_TAGS[node.localName]) return html;
  return html + node.childNodes.map(serializeNode).join("") + "</" + node.localName + ">";
}

// ---------------------------------------------------------------- selectors

export function cssEscape(value) {
  value = String(value);
  let result = "";
  for (let i = 0; i < value.length; i++) {
    const char = value[i], code = value.charCodeAt(i);
    if (code === 0) result += "�";
    else if ((code >= 1 && code <= 31) || code === 127 || (i === 0 && code >= 48 && code <= 57) || (i === 1 && code >= 48 && code <= 57 && value[0] === "-")) result += "\\" + code.toString(16) + " ";
    else if (i === 0 && char === "-" && value.length === 1) result += "\\" + char;
    else if (code >= 128 || char === "-" || char === "_" || /[0-9A-Za-z]/.test(char)) result += char;
    else result += "\\" + char;
  }
  return result;
}

const selectorCache = {};
export function parseSelector(text) {
  text = String(text);
  if (Object.prototype.hasOwnProperty.call(selectorCache, text)) return selectorCache[text];
  const parser = { text, index: 0 };
  const list = parseSelectorList(parser);
  skipSpace(parser);
  if (parser.index < text.length) throw new SyntaxError("Unsupported selector: " + text);
  selectorCache[text] = list;
  return list;
}
function skipSpace(parser) { while (parser.index < parser.text.length && /\s/.test(parser.text[parser.index])) parser.index++; }
function readIdentifier(parser) {
  let result = "";
  const text = parser.text;
  while (parser.index < text.length) {
    const char = text[parser.index];
    if (char === "\\") {
      const hex = /^[0-9a-fA-F]{1,6} ?/.exec(text.slice(parser.index + 1, parser.index + 8));
      if (hex) { result += String.fromCodePoint(parseInt(hex[0], 16)); parser.index += 1 + hex[0].length; }
      else { result += text[parser.index + 1]; parser.index += 2; }
    } else if (/[A-Za-z0-9_\-]/.test(char) || char.charCodeAt(0) >= 128) { result += char; parser.index++; }
    else break;
  }
  return result;
}
function parseSelectorList(parser) {
  const list = [];
  for (;;) {
    skipSpace(parser);
    list.push(parseComplex(parser));
    skipSpace(parser);
    if (parser.text[parser.index] === ",") { parser.index++; continue; }
    return list;
  }
}
function parseComplex(parser) {
  // Stored right-to-left: [{compound, combinator-to-the-left}, ...]
  const parts = [];
  let combinator = null;
  for (;;) {
    skipSpace(parser);
    const compound = parseCompound(parser);
    parts.push({ compound, combinator });
    const before = parser.index;
    skipSpace(parser);
    const char = parser.text[parser.index];
    if (char === ">" || char === "+" || char === "~") { combinator = char; parser.index++; continue; }
    if (parser.index < parser.text.length && char !== "," && char !== ")" && parser.index > before) { combinator = " "; continue; }
    if (parser.index < parser.text.length && char !== "," && char !== ")") { combinator = " "; continue; }
    break;
  }
  return parts.reverse();
}
function parseCompound(parser) {
  const compound = { tag: null, ids: [], classes: [], attributes: [], pseudos: [] };
  const text = parser.text;
  let any = false;
  if (text[parser.index] === "*") { parser.index++; any = true; }
  else if (/[A-Za-z]/.test(text[parser.index] || "")) { compound.tag = readIdentifier(parser).toLowerCase(); any = true; }
  while (parser.index < text.length) {
    const char = text[parser.index];
    if (char === "#") { parser.index++; compound.ids.push(readIdentifier(parser)); any = true; }
    else if (char === ".") { parser.index++; compound.classes.push(readIdentifier(parser)); any = true; }
    else if (char === "[") {
      parser.index++; skipSpace(parser);
      const name = readIdentifier(parser).toLowerCase(); skipSpace(parser);
      let operator = null, value = null;
      const opMatch = /^([~|^$*]?=)/.exec(text.slice(parser.index, parser.index + 2));
      if (opMatch) {
        operator = opMatch[1]; parser.index += operator.length; skipSpace(parser);
        const quote = text[parser.index];
        if (quote === '"' || quote === "'") {
          let value_ = ""; parser.index++;
          while (parser.index < text.length && text[parser.index] !== quote) {
            if (text[parser.index] === "\\") { value_ += text[parser.index + 1]; parser.index += 2; } else value_ += text[parser.index++];
          }
          parser.index++; value = value_;
        } else value = readIdentifier(parser);
        skipSpace(parser);
      }
      if (text[parser.index] !== "]") throw new SyntaxError("Bad attribute selector: " + text);
      parser.index++;
      compound.attributes.push({ name, operator, value });
      any = true;
    } else if (char === ":") {
      parser.index++;
      if (text[parser.index] === ":") parser.index++;
      const name = readIdentifier(parser).toLowerCase();
      let argument = null;
      if (text[parser.index] === "(") {
        parser.index++;
        if (name === "not" || name === "is" || name === "where" || name === "has") argument = parseSelectorList(parser);
        else { const close = text.indexOf(")", parser.index); argument = text.slice(parser.index, close).trim(); parser.index = close; }
        skipSpace(parser);
        if (text[parser.index] !== ")") throw new SyntaxError("Bad pseudo-class: " + text);
        parser.index++;
      }
      compound.pseudos.push({ name, argument });
      any = true;
    } else break;
  }
  if (!any) throw new SyntaxError("Unsupported selector: " + text);
  return compound;
}
function matchesCompound(element, compound, scope) {
  if (compound.tag && element.localName !== compound.tag) return false;
  for (let i = 0; i < compound.ids.length; i++) if (element._values.id !== compound.ids[i]) return false;
  if (compound.classes.length) {
    const classes = element._classTokens();
    for (let i = 0; i < compound.classes.length; i++) if (classes.indexOf(compound.classes[i]) < 0) return false;
  }
  for (let i = 0; i < compound.attributes.length; i++) {
    const rule = compound.attributes[i];
    const actual = element.getAttribute(rule.name);
    if (actual === null) return false;
    if (!rule.operator) continue;
    const expected = rule.value;
    if (rule.operator === "=" && actual !== expected) return false;
    if (rule.operator === "~=" && actual.split(/\s+/).indexOf(expected) < 0) return false;
    if (rule.operator === "^=" && (expected === "" || actual.indexOf(expected) !== 0)) return false;
    if (rule.operator === "$=" && (expected === "" || actual.slice(-expected.length) !== expected)) return false;
    if (rule.operator === "*=" && (expected === "" || actual.indexOf(expected) < 0)) return false;
    if (rule.operator === "|=" && actual !== expected && actual.indexOf(expected + "-") !== 0) return false;
  }
  for (let i = 0; i < compound.pseudos.length; i++) if (!matchesPseudo(element, compound.pseudos[i], scope)) return false;
  return true;
}
function matchesPseudo(element, pseudo, scope) {
  switch (pseudo.name) {
    case "not": return !matchesSelectorList(element, pseudo.argument, scope);
    case "is": case "where": return matchesSelectorList(element, pseudo.argument, scope);
    case "has": return pseudo.argument.some(complex => querySelectorFrom(element, complex, true, true).length > 0);
    case "scope": return element === scope;
    case "checked": return element.localName === "option" ? element.selected : (element.localName === "input" && (element.type === "checkbox" || element.type === "radio") && element.checked);
    case "disabled": return Boolean(element.disabled);
    case "enabled": return "disabled" in element && !element.disabled;
    case "empty": return element.childNodes.every(child => child.nodeType === 8);
    case "first-child": return element.parentNode && element.parentNode.children[0] === element;
    case "last-child": { const siblings = element.parentNode ? element.parentNode.children : []; return siblings[siblings.length - 1] === element; }
    case "only-child": return element.parentNode && element.parentNode.children.length === 1;
    case "nth-child": return matchesNth(element.parentNode ? element.parentNode.children.indexOf(element) + 1 : 1, pseudo.argument);
    case "first-of-type": return element.parentNode && element.parentNode.children.filter(child => child.localName === element.localName)[0] === element;
    case "focus": case "focus-visible": return element.ownerDocument && element.ownerDocument._activeElement === element;
    case "focus-within": return element.ownerDocument && element.contains(element.ownerDocument._activeElement);
    case "placeholder-shown": return "value" in element && element.value === "" && element.hasAttribute("placeholder");
    case "hover": case "active": case "visited": return false;
    case "link": case "any-link": return element.localName === "a" && element.hasAttribute("href");
    case "root": return element.ownerDocument && element.ownerDocument.documentElement === element;
    default: throw new SyntaxError("Unsupported pseudo-class :" + pseudo.name);
  }
}
function matchesNth(position, argument) {
  const text = String(argument).replace(/\s+/g, "").toLowerCase();
  if (text === "odd") return position % 2 === 1;
  if (text === "even") return position % 2 === 0;
  const match = /^([+-]?\d*)n([+-]\d+)?$/.exec(text);
  if (!match) return position === Number(text);
  const a = match[1] === "" || match[1] === "+" ? 1 : match[1] === "-" ? -1 : Number(match[1]);
  const b = Number(match[2] || 0);
  if (a === 0) return position === b;
  return (position - b) / a >= 0 && (position - b) % a === 0;
}
function matchesComplex(element, parts, scope) {
  // parts are right-to-left; parts[i].combinator joins parts[i] to parts[i+1].
  const step = (node, index) => {
    if (!matchesCompound(node, parts[index].compound, scope)) return false;
    if (index === parts.length - 1) return true;
    const combinator = parts[index].combinator;
    if (combinator === ">") { const parent = node.parentNode; return Boolean(parent && parent.nodeType === 1 && step(parent, index + 1)); }
    if (combinator === " ") { let parent = node.parentNode; while (parent && parent.nodeType === 1) { if (step(parent, index + 1)) return true; parent = parent.parentNode; } return false; }
    if (combinator === "+") { const previous = node.previousElementSibling; return Boolean(previous && step(previous, index + 1)); }
    if (combinator === "~") { let previous = node.previousElementSibling; while (previous) { if (step(previous, index + 1)) return true; previous = previous.previousElementSibling; } return false; }
    return false;
  };
  return step(element, 0);
}
function matchesSelectorList(element, list, scope) {
  for (let i = 0; i < list.length; i++) if (matchesComplex(element, list[i], scope)) return true;
  return false;
}
// "#id" alone is the website's commonest lookup ($("#id")); answer it from the index.
const SIMPLE_ID = /^#([A-Za-z_][\w-]*)$/;
function querySelectorFrom(root, selector, firstOnly, preParsed) {
  if (!preParsed) {
    const simple = SIMPLE_ID.exec(selector);
    const document = root.nodeType === 9 ? root : root.ownerDocument;
    if (simple && document && document._ids && (root.nodeType === 9 || root.isConnected)) {
      const found = document._lookupId(simple[1]);
      return found && found !== root && root.contains(found) ? [found] : [];
    }
  }
  const list = preParsed ? [selector] : parseSelector(selector);
  const scope = root.nodeType === 1 ? root : null;
  const results = [];
  const walk = node => {
    const children = node.childNodes;
    for (let i = 0; i < children.length; i++) {
      const child = children[i];
      if (child.nodeType !== 1) continue;
      if (matchesSelectorList(child, list, scope)) { results.push(child); if (firstOnly) return true; }
      if (walk(child)) return true;
    }
    return false;
  };
  walk(root);
  return results;
}

export function createDocument(bodyHtml) {
  const document = new Document();
  if (bodyHtml) {
    const bodyMatch = /^\s*<body([^>]*)>/i.exec(bodyHtml);
    if (bodyMatch) {
      const attributes = /([A-Za-z-]+)\s*=\s*"([^"]*)"/g;
      let match;
      while ((match = attributes.exec(bodyMatch[1]))) document.body.setAttribute(match[1], decodeEntities(match[2]));
      bodyHtml = bodyHtml.slice(bodyMatch[0].length).replace(/<\/body>\s*$/i, "");
    }
    buildTree(document, document.body, bodyHtml);
  }
  return document;
}
