// Differential test of the VBA DOM (excel/vba/modDom.bas, clsNode.cls, ...) against
// pdf/src/minidom.js. Builds a case file from the website body and a scripted set of
// DOM operations, computes minidom's answers, then runs testDom.RunCases in Excel.
//
//   node test/dom-test.js            (from excel/)  generate + run in Excel
//   node test/dom-test.js --gen-only <out.txt>      only write the case file

const fs = require("fs");
const os = require("os");
const path = require("path");
const { spawnSync } = require("child_process");

const EXCEL = path.join(__dirname, "..");
const REPO = path.join(EXCEL, "..");
const esbuild = require(path.join(REPO, "pdf", "node_modules", "esbuild"));
const { extractWebsite } = require(path.join(REPO, "pdf", "src", "extract.js"));

function loadMinidom() {
  const out = path.join(os.tmpdir(), "bt1101-minidom-" + process.pid + ".cjs");
  esbuild.buildSync({ entryPoints: [path.join(REPO, "pdf", "src", "minidom.js")], bundle: true, format: "cjs", platform: "node", outfile: out, logLevel: "error" });
  const lib = require(out);
  fs.unlinkSync(out);
  return lib;
}

const dom = loadMinidom();
const site = extractWebsite(path.join(REPO, "index.html"));
const RS = "\x1e", FS = "\x1f", NULL = "\u0000null";

function pathOf(node) {
  if (!node) return "null";
  if (node === win) return "win";
  if (node.nodeType === 9) return "doc";
  const parts = [];
  while (node.parentNode) { parts.unshift(node.parentNode.childNodes.indexOf(node)); node = node.parentNode; }
  return (node.nodeType === 9 ? "" : "~") + parts.join(".");
}
function paths(list) { return list.map(pathOf).join(","); }
function nodeAt(p) {
  let node = document;
  if (p === "") return node;
  for (const index of p.split(".")) node = node.childNodes[Number(index)];
  return node;
}

let document = dom.createDocument(site.body);
let win = new dom.EventTarget();
document.defaultView = win;
win._parentForEvents = () => null;

const records = [];
const add = fields => records.push(fields.map(String).join(FS));

// ---------------------------------------------------------------- static cases

const body = site.body;
add(["B", body, document.body.innerHTML, document.querySelectorAll("*").length]);

// Selector strings from the website's main script.
const code = site.scripts[site.scripts.length - 1];
const literal = /"((?:[^"\\\n]|\\.)*)"|'((?:[^'\\\n]|\\.)*)'/g;
const candidates = new Set();
let m;
while ((m = literal.exec(code))) {
  let text = m[1] !== undefined ? m[1] : m[2];
  try { text = JSON.parse('"' + text.replace(/\\'/g, "'").replace(/"/g, '\\"') + '"'); } catch (error) { continue; }
  if (!text || text.length > 300) continue;
  if (!/^[#.\[:*a-z]/i.test(text)) continue;
  if (!/[#.\[:>*]/.test(text) && !/^(table|select|p|label|button|input|option|textarea|strong|h3|h2|pre|code|summary|details|span|div|section|article|li|ul|main|aside|header|small|fieldset|legend)$/.test(text)) continue;
  candidates.add(text);
}
const extra = [
  "#tutorialList .operation[data-id]", ".panel", "*", "div", "body > div", "div, span", ":is(h2, h3)",
  ".panel > .panel-head .panel-title strong", "main > .bench:not(.hidden)", "section:has(.panel-head)",
  "label + input", "h2 ~ p", "li:nth-child(2n+1)", ":not(div):not(span)", "[class~=panel]", "[data-id^=\"t\"]",
  "[data-id$=\"s\"]", "[data-id*=\"reg\"]", "[lang|=en]", "input:checked", "option:checked", "button:disabled",
  "input:enabled", ":empty", "p:first-child", "li:last-child", "span:only-child", "h3:first-of-type",
  "details[open] > summary", "textarea:placeholder-shown", "input:placeholder-shown", "#themeToggle .theme-toggle-icon", "a:link", ":root",
  ":where(.panel) .panel-title", "input:not([type=\"number\"]):not([type=\"checkbox\"]), textarea.column-input",
  "[data-answer-variant-ready]", "select option:nth-child(odd)", "li:nth-child(even)", ":nth-child(3)",
  ":nth-child(-n+2)", "td:nth-child(n)", "option:nth-child(3n-1)", "div:nth-child( 2n + 1 )", "[ type = \"button\" ]", "[type='text' ]",
  "button[aria-pressed=false]", "[aria-hidden]", "aside label input", "header > button > span", "section > *:first-child",
  "*:not(:empty) > strong", "details > summary + *", "main *:last-child", "select:has(option[selected])",
  ":has(> span)", "DIV.panel", "INPUT[TYPE=checkbox]", "[class|=panel]", "[class^=\"panel \"]", "[class$=\"-head\"]",
  ".bench .panel .panel-head", "#lpBench .panel, #pcaBench > .panel", "fieldset legend ~ label", "p ~ p ~ p", "pre + pre",
  "#doesNotExist", "#doesNotExist .x", ".a\\:b", "#\\31 23", "[data-tutorial]", "[data-search-exact]", ":not([hidden])",
  ".hidden:not(.panel) small", "option:not(:checked)", "input[type=checkbox]:not(:checked)", ":disabled", ":enabled"
];
extra.forEach(text => candidates.add(text));

const selectors = [];
for (const text of candidates) {
  try {
    const list = document.querySelectorAll(text);
    selectors.push(text);
    add(["Q", text, paths(list)]);
  } catch (error) { /* not a selector minidom supports */ }
}

// The website's body after its script has run (dynamic lists filled in): reload it and
// repeat the document queries. Benchmarks in VBA run on this document.
const { bootNodeApp } = require(path.join(REPO, "pdf", "test", "harness.js"));
const booted = bootNodeApp();
const bootedHtml = booted.browser.document.body.outerHTML;
document = dom.createDocument(bootedHtml);
document.defaultView = win;
add(["B", bootedHtml, document.body.innerHTML, document.querySelectorAll("*").length]);
for (const text of selectors) add(["Q", text, paths(document.querySelectorAll(text))]);

// Element-scoped queries, closest and matches from a sample of elements.
const all = document.querySelectorAll("*");
const sample = all.filter((_, i) => i % 37 === 5).concat(document.querySelectorAll(".panel").slice(0, 6)).concat(document.querySelectorAll("label, option, summary").slice(0, 10));
const scoped = [":scope > .panel-head", ":scope > .panel-head .panel-title strong", ":scope > *", "*", "div > span", ":scope span", "input, select, textarea, button", ":scope > .completion-badge", "label", "option", ":not(:scope)"];
const probe = selectors.filter((_, i) => i % 4 === 0).concat([".panel", ".bench", "main", "body", "aside", "section", "div", ":scope", "[data-id]", "label, .panel", "details", "body > div *"]);
for (const el of sample) {
  for (const sel of scoped) add(["E", pathOf(el), sel, paths(el.querySelectorAll(sel))]);
  for (const sel of probe) {
    let closest, matches;
    try { closest = pathOf(el.closest(sel)); matches = el.matches(sel) ? "1" : "0"; } catch (error) { continue; }
    add(["C", pathOf(el), sel, closest]);
    add(["M", pathOf(el), sel, matches]);
  }
}

// ---------------------------------------------------------------- scripted scenarios

const vars = {};
const behaviour = {};
let log = [];
const handlers = {};
for (let i = 1; i <= 8; i++) {
  const name = "H" + i;
  handlers[name] = function (event) {
    let line = name + ":" + event.type + "@" + pathOf(event.currentTarget) + ":" + event.eventPhase + ":" + pathOf(event.target);
    if (event.target && event.target.localName === "input") line += ":c" + (event.target.checked ? 1 : 0);
    log.push(line);
    const act = behaviour[name] || "none";
    if (act === "stop") event.stopPropagation();
    else if (act === "stopnow") event.stopImmediatePropagation();
    else if (act === "prevent") event.preventDefault();
  };
}

function ref(text) {
  if (text === "doc") return document;
  if (text === "win") return win;
  if (text === "body") return document.body;
  if (text[0] === "$") return vars[text.slice(1)] || null;
  if (text[0] === "@") return nodeAt(text.slice(1));
  return document.querySelector(text);
}
const bool = text => text === "1";
function flag(value) { return value ? "1" : "0"; }

function runOp(op, a, b, c, d) {
  const el = () => ref(a);
  switch (op) {
    case "let": vars[a] = ref(b); return pathOf(vars[a]);
    case "create": vars[a] = document.createElement(b); return "";
    case "textnode": vars[a] = document.createTextNode(b); return "";
    case "frag": vars[a] = document.createDocumentFragment(); return "";
    case "parse": { const f = document.createDocumentFragment(); dom.buildTree(document, f, b); vars[a] = f; return String(f.childNodes.length); }
    case "sethtml": el().innerHTML = b; return el().innerHTML;
    case "html": return el().innerHTML;
    case "outer": return el().outerHTML;
    case "setouter": el().outerHTML = b; return "";
    case "adj": el().insertAdjacentHTML(b, c); return "";
    case "adjel": el().insertAdjacentElement(b, ref(c)); return "";
    case "settext": el().textContent = b; return "";
    case "gettext": return el().textContent;
    case "setval": el().value = b; return "";
    case "val": return el().value;
    case "defval": return el().defaultValue;
    case "setchecked": el().checked = bool(b); return "";
    case "checked": return flag(el().checked);
    case "setselidx": el().selectedIndex = Number(b); return "";
    case "selidx": return String(el().selectedIndex);
    case "opts": return el().options.map(o => o.value + (o.selected ? "*" : "")).join("|");
    case "optsel": return flag(el().options[Number(b)].selected);
    case "setoptsel": el().options[Number(b)].selected = bool(c); return "";
    case "opttext": return el().text;
    case "click": el().click(); return "";
    case "listen": el().addEventListener(b, handlers[c], { capture: d && d.indexOf("c") >= 0, once: d && d.indexOf("o") >= 0 }); return "";
    case "unlisten": el().removeEventListener(b, handlers[c], { capture: d === "c" }); return "";
    case "behave": behaviour[a] = b; return "";
    case "dispatch": return flag(el().dispatchEvent(new dom.Event(b, { bubbles: bool(c), cancelable: bool(d) })));
    case "log": { const text = log.join(" "); log = []; return text; }
    case "attr": { const v = el().getAttribute(b); return v === null ? NULL : v; }
    case "hasattr": return flag(el().hasAttribute(b));
    case "setattr": el().setAttribute(b, c); return "";
    case "rmattr": el().removeAttribute(b); return "";
    case "toggleattr": return flag(c === "" ? el().toggleAttribute(b) : el().toggleAttribute(b, bool(c)));
    case "attrs": return el().getAttributeNames().join(",");
    case "dget": { const v = el().dataset[b]; return v === undefined ? NULL : String(v); }
    case "dset": el().dataset[b] = c; return "";
    case "ddel": delete el().dataset[b]; return "";
    case "addcls": el().classList.add(...b.split(" ")); return "";
    case "rmcls": el().classList.remove(...b.split(" ")); return "";
    case "togglecls": return flag(c === "" ? el().classList.toggle(b) : el().classList.toggle(b, bool(c)));
    case "replcls": return flag(el().classList.replace(b, c));
    case "hascls": return flag(el().classList.contains(b));
    case "cls": return el().className;
    case "qsa": return paths(el().querySelectorAll(b));
    case "qs": return pathOf(el().querySelector(b));
    case "closest": return pathOf(el().closest(b));
    case "matches": return flag(el().matches(b));
    case "append": el().append(ref(b)); return "";
    case "appendtext": el().append(b); return "";
    case "appendmix": el().append(b, ref(c), d); return "";
    case "prepend": el().prepend(ref(b)); return "";
    case "before": el().before(ref(b)); return "";
    case "after": el().after(ref(b)); return "";
    case "replacewith": el().replaceWith(ref(b)); return "";
    case "replacechildren": el().replaceChildren(ref(b)); return "";
    case "remove": el().remove(); return "";
    case "appendchild": el().appendChild(ref(b)); return "";
    case "insertbefore": el().insertBefore(ref(b), b === c ? null : ref(c)); return "";
    case "removechild": el().removeChild(ref(b)); return "";
    case "clone": vars[a] = ref(b).cloneNode(bool(c)); return vars[a].nodeType === 1 ? vars[a].outerHTML : String(vars[a].childNodes.length);
    case "focus": el().focus(); return "";
    case "blur": el().blur(); return "";
    case "active": return pathOf(document.activeElement);
    case "kids": return el().childNodes.map(n => n.nodeType + (n.nodeType === 3 ? "'" + n.data + "'" : n.nodeType === 1 ? n.localName : "")).join("|");
    case "type": return el().type === undefined ? "" : String(el().type);
    case "disabled": return flag(el().disabled);
    case "setdisabled": el().disabled = bool(b); return "";
    case "hidden": return flag(el().hidden);
    case "sethidden": el().hidden = bool(b); return "";
    case "byid": return pathOf(document.getElementById(a));
    case "contains": return flag(el().contains(ref(b)));
    case "open": return flag(el().open);
    case "setopen": el().open = bool(b); return "";
    case "connected": return flag(el().isConnected);
    case "path": return pathOf(el());
    case "parent": return pathOf(el().parentNode);
    case "next": return pathOf(el().nextSibling);
    case "prev": return pathOf(el().previousSibling);
    case "nextel": return pathOf(el().nextElementSibling);
    case "prevel": return pathOf(el().previousElementSibling);
    case "firstel": return pathOf(el().firstElementChild);
    case "lastel": return pathOf(el().lastElementChild);
    case "children": return paths(el().children);
    case "tag": return el().tagName;
    case "setid": el().id = b; return "";
    case "id": return el().id;
    case "settitle": el().title = b; return "";
    case "title": return el().title;
    case "placeholder": return el().placeholder;
    case "htmlfor": return el().htmlFor;
    case "data": return el().data;
    case "replacechild": el().replaceChild(ref(b), ref(c)); return "";
    case "index": return String(el().index);
    case "control": return pathOf(el().control);
    case "byidfrag": return pathOf(el().getElementById(b));
    default: throw new Error("Unknown op " + op);
  }
}

const S = [];
const s = (...step) => S.push(step);

// -- markup, entities, implied end tags, serialization
s("let", "b", "body");
s("adj", "$b", "beforeend", '<div id="w"><div id="t1" class="a  b" data-foo-bar="1" title="x&amp;y"><p>one<p>two<ul><li>a<li>b</ul><span>&lt;&nbsp;&copy;&#x41;&#66;&bogus;&amp &#X41; &#128512;</span><!-- c --></div></div>');
s("html", "#w");
s("let", "t1", "#t1");
s("kids", "$t1");
s("cls", "$t1"); s("addcls", "$t1", "c a"); s("cls", "$t1"); s("rmcls", "$t1", "a"); s("cls", "$t1");
s("togglecls", "$t1", "b", ""); s("togglecls", "$t1", "z", "1"); s("togglecls", "$t1", "z", "0"); s("togglecls", "$t1", "q", "0");
s("replcls", "$t1", "c", "d"); s("replcls", "$t1", "nope", "d"); s("cls", "$t1"); s("hascls", "$t1", "d"); s("hascls", "$t1", "c");
s("dget", "$t1", "fooBar"); s("dset", "$t1", "answerVariantId", "x1"); s("attr", "$t1", "data-answer-variant-id");
s("attr", "$t1", "DATA-FOO-BAR"); s("attrs", "$t1"); s("setattr", "$t1", "data-some-thing", "v"); s("dget", "$t1", "someThing");
s("ddel", "$t1", "fooBar"); s("attrs", "$t1"); s("outer", "$t1"); s("hasattr", "$t1", "data-foo-bar");
s("toggleattr", "$t1", "hidden", ""); s("attr", "$t1", "hidden"); s("hidden", "$t1"); s("toggleattr", "$t1", "hidden", ""); s("sethidden", "$t1", "1"); s("outer", "$t1"); s("sethidden", "$t1", "0");
s("rmattr", "$t1", "title"); s("setattr", "$t1", "Title", "a\"b&c d"); s("title", "$t1"); s("outer", "$t1");
s("sethtml", "$t1", "<b>x</b>text<i>y</i>"); s("gettext", "$t1"); s("settext", "$t1", "a<b>&c "); s("html", "$t1"); s("kids", "$t1"); s("settext", "$t1", ""); s("kids", "$t1");
s("sethtml", "$t1", '<textarea id="ta">\nline&amp;1\n</textarea><select id="s1"><option>A<option value="b" selected>B<optgroup label="g"><option disabled>C</optgroup></select><script>if (a<b && c>d) {}</script><style>a>b{}</style><br><img src="x"/><input id="i1" value="v&quot;1"><table><tr><td>1<td>2<tr><th>3</table><p>para<div>after p</div><TEXTAREA>\n</TEXTAREA><title>T&amp;</title>');
s("html", "$t1"); s("kids", "$t1");
s("val", "#ta"); s("defval", "#ta"); s("setval", "#ta", "new"); s("val", "#ta"); s("defval", "#ta"); s("html", "#ta");
s("val", "#s1"); s("selidx", "#s1"); s("opts", "#s1"); s("setselidx", "#s1", "0"); s("val", "#s1"); s("setval", "#s1", "C"); s("selidx", "#s1"); s("setval", "#s1", "zzz"); s("selidx", "#s1"); s("val", "#s1"); s("opts", "#s1");
s("setoptsel", "#s1", "1", "1"); s("opts", "#s1"); s("optsel", "#s1", "0"); s("qsa", "#s1", "option:checked"); s("qsa", "#s1", "option:disabled");
s("let", "o0", "#s1 option"); s("opttext", "$o0"); s("type", "#s1");
s("val", "#i1"); s("defval", "#i1"); s("setval", "#i1", "typed"); s("val", "#i1"); s("attr", "#i1", "value"); s("outer", "#i1"); s("type", "#i1");
s("adj", "$t1", "afterbegin", "<em>first</em>"); s("adj", "$t1", "beforebegin", "<hr>"); s("adj", "$t1", "afterend", '<span id="after">z</span>'); s("adj", "$t1", "beforeend", "tail &amp; more");
s("html", "#w");
s("setouter", "#after", "<b>one</b><b>two</b>"); s("html", "#w");
s("parse", "pf", "<p>a<p>b<option>x<option>y</option></p><li>q<li>r"); s("clone", "pfc", "$pf", "1"); s("appendchild", "#w", "$pf"); s("html", "#w"); s("kids", "$pf");
s("sethtml", "#w", "a<b <c d=e f>x</c>< y>&amp&lt;<!doctype z><?pi?><br/><div / class=q>k</div><input disabled checked value=v name=nm><x-y data-q=1>w</x-y><!-- unterminated");
s("html", "#w"); s("kids", "#w");
s("sethtml", "#w", "<p dup=1 DUP=2 a='x\"y' b=\"p'q\">t</p><ol><li>1<li>2</ol><div><span>unclosed</div>tail");
s("html", "#w");

// -- forms and click behaviour
s("sethtml", "#w", '<form id="f"><input type="checkbox" id="cb"><label id="lb" for="cb">L</label><label id="lw"><input type="radio" name="r" id="r1" value="1">R1</label><input type="radio" name="r" id="r2" checked><input type="radio" name="q" id="q1"></form><input type="radio" name="r" id="r3"><details id="d"><summary id="sm">S</summary>body</details><button id="btn" disabled>B</button><button id="btn2" type="button">C</button><input id="txt" placeholder="p"><textarea id="ta2" placeholder="x"></textarea><label id="le">empty</label><label id="lx" for="nope">x</label>');
s("listen", "#w", "click", "H1", ""); s("listen", "#w", "input", "H2", ""); s("listen", "#w", "change", "H3", "");
s("click", "#cb"); s("log"); s("checked", "#cb"); s("val", "#cb");
s("behave", "H1", "prevent"); s("click", "#cb"); s("log"); s("checked", "#cb"); s("behave", "H1", "none");
s("click", "#lb"); s("log"); s("checked", "#cb");
s("click", "#lw"); s("log"); s("checked", "#r1"); s("checked", "#r2"); s("checked", "#q1");
s("click", "#r2"); s("log"); s("checked", "#r1"); s("checked", "#r2");
s("click", "#r2"); s("log");
s("click", "#r3"); s("log"); s("checked", "#r2"); s("checked", "#r3");
s("click", "#le"); s("click", "#lx"); s("log");
s("click", "#sm"); s("open", "#d"); s("outer", "#d"); s("log"); s("click", "#sm"); s("open", "#d"); s("setopen", "#d", "1"); s("qsa", "#w", "details[open] > summary");
s("click", "#btn"); s("log"); s("click", "#btn2"); s("log");
s("setchecked", "#q1", "1"); s("checked", "#q1"); s("qsa", "doc", "input:checked"); s("qsa", "#w", "input:not(:checked)");
s("setval", "#cb", "yes"); s("attr", "#cb", "value"); s("val", "#cb");
s("type", "#cb"); s("type", "#btn2"); s("type", "#btn"); s("type", "#ta2"); s("type", "#d");
s("disabled", "#btn"); s("setdisabled", "#btn", "0"); s("disabled", "#btn"); s("qsa", "#w", "button:enabled"); s("qsa", "#w", ":disabled"); s("setdisabled", "#cb", "1"); s("click", "#cb"); s("log"); s("checked", "#cb");
s("qsa", "#w", ":placeholder-shown"); s("setval", "#txt", "typed"); s("qsa", "#w", ":placeholder-shown"); s("placeholder", "#txt"); s("htmlfor", "#lb");
s("setchecked", "#r1", "1"); s("checked", "#r2"); s("checked", "#r3");
s("unlisten", "#w", "click", "H1", ""); s("click", "#btn2"); s("log");

// -- event propagation
s("sethtml", "#w", '<div id="o"><div id="m"><span id="in">x</span></div></div>');
s("listen", "doc", "click", "H1", "c"); s("listen", "win", "click", "H2", "c"); s("listen", "#o", "click", "H3", "c"); s("listen", "#o", "click", "H4", "");
s("listen", "#m", "click", "H5", ""); s("listen", "#in", "click", "H6", "c"); s("listen", "#in", "click", "H7", ""); s("listen", "win", "click", "H8", "");
s("listen", "#in", "click", "H5", "c");
s("dispatch", "#in", "click", "1", "1"); s("log");
s("dispatch", "#in", "click", "0", "0"); s("log");
s("dispatch", "#m", "click", "1", "0"); s("log");
s("behave", "H3", "stop"); s("dispatch", "#in", "click", "1", "1"); s("log"); s("behave", "H3", "none");
s("behave", "H6", "stopnow"); s("dispatch", "#in", "click", "1", "1"); s("log"); s("behave", "H6", "none");
s("behave", "H5", "stop"); s("dispatch", "#in", "click", "1", "1"); s("log"); s("behave", "H5", "none");
s("behave", "H7", "prevent"); s("dispatch", "#in", "click", "1", "1"); s("log"); s("dispatch", "#in", "click", "1", "0"); s("log"); s("behave", "H7", "none");
s("listen", "#in", "click", "H7", ""); s("unlisten", "#in", "click", "H7", "c"); s("dispatch", "#in", "click", "0", "0"); s("log");
s("unlisten", "#in", "click", "H7", ""); s("dispatch", "#in", "click", "0", "0"); s("log");
s("listen", "#in", "focus", "H4", "o"); s("dispatch", "#in", "focus", "0", "0"); s("dispatch", "#in", "focus", "0", "0"); s("log");
s("dispatch", "doc", "click", "1", "1"); s("log"); s("dispatch", "win", "click", "1", "1"); s("log");
s("create", "det", "div"); s("listen", "$det", "click", "H6", ""); s("dispatch", "$det", "click", "1", "1"); s("log");
s("unlisten", "doc", "click", "H1", "c"); s("unlisten", "win", "click", "H2", "c"); s("unlisten", "win", "click", "H8", "");

// -- tree manipulation, cloning, ids, focus
s("sethtml", "#w", '<ul id="list"><li id="a">A</li><li id="b">B</li><li id="c">C</li></ul><p id="dup">1</p><p id="dup">2</p><input id="inp" value="x" type="checkbox" checked>');
s("byid", "dup"); s("byid", "a"); s("byid", "missing"); s("let", "pa", "#a"); s("let", "pb", "#b"); s("let", "pc", "#c");
s("next", "$pa"); s("prev", "$pa"); s("nextel", "$pb"); s("prevel", "$pb"); s("firstel", "#list"); s("lastel", "#list"); s("children", "#list"); s("parent", "$pa"); s("tag", "$pa");
s("appendchild", "#list", "$pa"); s("html", "#list"); s("insertbefore", "#list", "$pc", "$pb"); s("html", "#list");
s("insertbefore", "#list", "$pb", "$pb"); s("html", "#list");
s("create", "n1", "li"); s("settext", "$n1", "N1"); s("before", "$pa", "$n1"); s("html", "#list");
s("create", "n2", "li"); s("after", "$pa", "$n2"); s("prepend", "#list", "$pa"); s("html", "#list");
s("frag", "fr"); s("create", "x1", "b"); s("create", "x2", "i"); s("appendchild", "$fr", "$x1"); s("appendchild", "$fr", "$x2"); s("kids", "$fr");
s("replacewith", "$n2", "$fr"); s("html", "#list"); s("kids", "$fr"); s("connected", "$n2"); s("connected", "$x1");
s("appendtext", "$pc", " more"); s("kids", "$pc"); s("appendmix", "$pc", "s1", "$n2", "s2"); s("html", "$pc");
s("replacechildren", "$pc", "$n1"); s("html", "#list");
s("removechild", "#list", "$x1"); s("html", "#list"); s("remove", "$pb"); s("html", "#list"); s("connected", "$pb"); s("contains", "#w", "$pa"); s("contains", "$pa", "#w"); s("contains", "$pa", "$pa");
s("clone", "cl", "#inp", "0"); s("setval", "#inp", "changed"); s("clone", "cl2", "#inp", "1"); s("val", "$cl2"); s("checked", "$cl2");
s("setchecked", "#inp", "0"); s("clone", "cl3", "#inp", "0"); s("checked", "$cl3"); s("checked", "$cl");
s("clone", "cl4", "#list", "1"); s("clone", "cl5", "#list", "0");
s("let", "d1", "#dup"); s("remove", "$d1"); s("byid", "dup"); s("setid", "#dup", "dupe2"); s("byid", "dup"); s("byid", "dupe2"); s("id", "#dupe2");
s("create", "nn", "p"); s("setid", "$nn", "fresh"); s("byid", "fresh"); s("appendchild", "#w", "$nn"); s("byid", "fresh");
s("prepend", "#w", "$d1"); s("byid", "dup"); s("setattr", "$d1", "id", "fresh"); s("byid", "fresh"); s("qsa", "doc", "#fresh"); s("qs", "#w", "#fresh");
s("rmattr", "$d1", "id"); s("byid", "fresh"); s("id", "$d1");
s("focus", "#inp"); s("active", ""); s("qsa", "#w", ":focus"); s("qsa", "doc", ":focus-within"); s("blur", "#inp"); s("active", ""); s("focus", "#inp"); s("remove", "#inp"); s("active", "");
s("focus", "$pa"); s("active", ""); s("blur", "$pc"); s("active", "");
s("qsa", "#w", ":scope > *"); s("qs", "#w", ":scope li"); s("closest", "$pa", "#w"); s("closest", "$pa", ":scope"); s("matches", "$pa", "li"); s("matches", "$pa", "#list > li:first-child");
s("closest", "$cl4", "ul"); s("qsa", "$cl4", "li"); s("qs", "$cl4", "#a"); s("qsa", "$fr", "*");
s("setattr", "$pa", "data-x", "1"); s("qsa", "#w", "[data-x]"); s("dset", "$pa", "x", "2"); s("qsa", "#w", "[data-x=\"2\"]");
s("sethtml", "#w", '<div class="x y"></div><div class=" x "></div><div class="xy"></div><div class=""></div><div></div><p lang="en-US"></p><p lang="en"></p><p lang="english"></p><p title=" a  b "></p><p title=""></p>');
s("qsa", "#w", ".x"); s("qsa", "#w", "[class~=x]"); s("qsa", "#w", "[class~=\"\"]"); s("qsa", "#w", "[class=\"\"]"); s("qsa", "#w", "[class]");
s("qsa", "#w", "[lang|=en]"); s("qsa", "#w", "[title~=a]"); s("qsa", "#w", "[title~=\"\"]"); s("qsa", "#w", "[title^=\"\"]"); s("qsa", "#w", "div:empty"); s("qsa", "#w", "p:nth-child(2n)");
s("qsa", "#w", "div:nth-child(-n+3)"); s("qsa", "#w", ":nth-child(0n+2)"); s("qsa", "#w", ":nth-child(+5)"); s("qsa", "#w", "p:not(:first-of-type)"); s("qsa", "#w", "div + p, p + p");
s("let", "dx", "#w div"); s("addcls", "$dx", "x"); s("cls", "$dx"); s("let", "de", "#w div:nth-child(4)"); s("rmcls", "$de", "zz"); s("cls", "$de"); s("attr", "$de", "class");
s("let", "dn", "#w div:nth-child(5)"); s("rmcls", "$dn", "zz"); s("attr", "$dn", "class"); s("addcls", "$dn", "k"); s("outer", "$dn");
s("html", "#w");

// -- more tree and form coverage
s("sethtml", "#w", '<label id="lab"><span id="ls">t</span><input type="checkbox" id="lc"></label><label id="lf" for="lc2">f</label><input id="lc2" type="checkbox"><select id="s2"><optgroup><option id="oa">a</option><option id="ob" disabled>b</option></optgroup><option id="oc">c</option></select><div id="z"><i id="zi">1</i></div>');
s("listen", "#w", "click", "H1", ""); s("listen", "#w", "change", "H3", "");
s("click", "#lc"); s("log"); s("checked", "#lc"); s("click", "#ls"); s("log"); s("checked", "#lc"); s("click", "#lab"); s("log"); s("checked", "#lc");
s("control", "#lab"); s("control", "#lf"); s("click", "#lf"); s("log"); s("checked", "#lc2");
s("index", "#ob"); s("index", "#oc"); s("selidx", "#s2"); s("setoptsel", "#s2", "0", "0"); s("selidx", "#s2"); s("opts", "#s2"); s("val", "#s2");
s("setattr", "#oa", "disabled", ""); s("sethtml", "#s2", '<option disabled>x</option><option>y</option>'); s("selidx", "#s2"); s("val", "#s2"); s("opttext", "#s2 option");
s("create", "q1", "em"); s("adjel", "#zi", "beforebegin", "$q1"); s("create", "q2", "em"); s("adjel", "#zi", "afterend", "$q2"); s("create", "q3", "u"); s("adjel", "#z", "afterbegin", "$q3"); s("html", "#z");
s("create", "q4", "s"); s("replacechild", "#z", "$q4", "#zi"); s("html", "#z"); s("connected", "#zi");
s("parse", "pf2", '<div id="fx"><p id="fy">1</p></div>'); s("byidfrag", "$pf2", "fy"); s("qs", "$pf2", "#fy"); s("qsa", "$pf2", "div > p, #fx");
s("unlisten", "#w", "click", "H1", ""); s("unlisten", "#w", "change", "H3", "");

const scenario = [];
for (const step of S) {
  const [op, a = "", b = "", c = "", d = ""] = step;
  let result;
  try { result = runOp(op, a, b, c, d); } catch (error) { result = "ERR"; }
  scenario.push([op, a, b, c, d, result]);
  add(["S", op, a, b, c, d, result]);
}

const content = records.join(RS);
const counts = {};
records.forEach(r => { const k = r.slice(0, 1); counts[k] = (counts[k] || 0) + 1; });

const args = process.argv.slice(2);
if (args[0] === "--gen-only") {
  fs.writeFileSync(args[1], content, "utf8");
  console.log("Wrote", args[1], counts);
  return;
}
const caseFile = path.join(os.tmpdir(), "bt1101-dom-cases-" + process.pid + ".txt");
fs.writeFileSync(caseFile, content, "utf8");
console.log("Cases:", JSON.stringify(counts), "selectors:", selectors.length);
const only = "modJs.bas,clsList.cls,clsMap.cls,modDom.bas,clsNode.cls,clsEvent.cls,clsDomHash.cls,clsDomSelector.cls,testDom.bas";
const result = spawnSync("node", [path.join(EXCEL, "tools", "vba.js"), "run", "testDom.RunCases", caseFile, "--only", only, "--no-form"], { encoding: "utf8", cwd: EXCEL, timeout: 20 * 60 * 1000 });
process.stdout.write(result.stdout || "");
process.stderr.write(result.stderr || "");
fs.unlinkSync(caseFile);
process.exitCode = result.status;
