// Injected into index.html in headless Edge at build time. It records where the website
// draws every box, word and control so the PDF can reproduce the page, then reports which
// elements become PDF fields. Coordinates are CSS pixels from the page's top-left corner.

import { refFor, isRendered, resolveRef, applySetting, acceptAlerts } from "./refs.js";

const OUTPUT = "pre, .output-body, .pca-variance-prose, .sensitivity-prose-grid, .hypothesis-card, .mapping-card, .shape-guide, .lp-parser-status, .lp-constraint-error, .lp-table-wrap";
const CARD = ".sensitivity-prose-card";
// Browser-only controls. Grid children are made invisible rather than removed, so the
// website's grid keeps its columns.
const INVISIBLE = "#themeToggle, #sidebarResizer";
const REMOVED = "#mobileLibraryBtn, #updateCheck, #clearSearch, #toast, #referencePanel, #pasteBtn, #pasteOneSampleBtn, #pasteTwoSampleBtn, .lp-collapse-toggle, .plot-workflow-menu, .variable-rename-alert, #rankedSearchResults, #noResults";

const wait = ms => new Promise(resolve => setTimeout(resolve, ms));
const bench = () => document.querySelector("main > .bench:not(.hidden), main > #lpBench:not(.hidden)");

export async function open(operation) {
  const entry = operation === "universal" ? document.querySelector(".universal-entry") : document.querySelector(`#tutorialList .operation[data-id="${operation}"]`);
  entry.click();
  await wait(150);
}

// Fill the page with representative content so outputs get realistic sizes.
export async function loadSamples(operation) {
  const root = bench();
  if (operation === "universal") {
    const select = document.querySelector("#universalSampleSelect");
    select.value = "shapiro";
    select.dispatchEvent(new Event("change", { bubbles: true }));
    await wait(700);
    return;
  }
  if (operation === "shapiro" || operation === "pearson") { document.querySelector("#singleSampleBtn").click(); await wait(700); return; }
  for (let pass = 0; pass < 2; pass++) {
    Array.from(root.querySelectorAll("button")).filter(button => /^Use sample/i.test(button.textContent.trim()) && !button.disabled && isRendered(button)).forEach(button => button.click());
    await wait(500);
  }
  // Rows the website appends as you type: make room for several in the PDF.
  if (operation === "summarize") await fillRows("#columnFields .column-row", ".column-input", 8, index => "column" + (index + 1));
  if (operation === "lpsolve") await fillRows("#lpConstraintRows .lp-constraint-row", ".lp-name", 10, index => "Constraint " + (index + 1));
}

async function fillRows(rowSelector, inputSelector, count, value) {
  for (let guard = 0; guard < count * 2 && document.querySelectorAll(rowSelector).length < count; guard++) {
    const rows = document.querySelectorAll(rowSelector);
    const input = rows[rows.length - 1].querySelector(inputSelector);
    input.value = value(rows.length - 1);
    input.dispatchEvent(new Event("input", { bubbles: true }));
    await wait(30);
  }
}

// Puts the page into one of its layouts (see explore.js).
export async function applySettings(settings) {
  for (const setting of settings) { applySetting(document, setting, Event); acceptAlerts(document); await wait(60); }
}

// Which of the given elements are showing: the layout actually reached.
export function signature(tracked) {
  return tracked.filter(ref => { const element = resolveRef(document, ref); return element && isRendered(element, true); });
}

// Records references before any capture-only change. For the catch-all page it also makes
// every element that can appear visible, so each has a place somewhere in the PDF.
export function prepare(config) {
  const refs = new Map();
  document.querySelectorAll("body *").forEach(element => refs.set(element, refFor(element, document)));
  const conditional = new Set(config.conditional), ever = new Set(config.ever);
  const root = bench();
  root.querySelectorAll("*").forEach(element => {
    const ref = refs.get(element);
    if (config.unhide && ever.has(ref) && (element.classList.contains("hidden") || element.hasAttribute("hidden")) && !element.matches("#plotWorkflow")) {
      element.classList.remove("hidden");
      element.removeAttribute("hidden");
    }
  });
  root.querySelectorAll(".lp-collapsible.is-collapsed").forEach(panel => panel.classList.remove("is-collapsed"));
  root.querySelectorAll(".input-error").forEach(element => element.classList.remove("input-error"));
  document.querySelectorAll(REMOVED).forEach(element => { element.style.display = "none"; });
  document.querySelectorAll(INVISIBLE).forEach(element => { element.style.visibility = "hidden"; });
  // Show every page at full height: no sticky header, no scrolling sidebar.
  const style = document.createElement("style");
  style.textContent = `.topbar, .library { position: static !important; height: auto !important; max-height: none !important; overflow: visible !important; }
    .workspace { min-height: 0 !important; } body { background: var(--paper) !important; }
    .pdf-clip { display: flex; flex-wrap: wrap; align-items: center; gap: 6px 10px; margin: 10px 16px 0; }
    .pdf-clip-hint { flex: none; font-size: 12px; font-weight: 600; color: var(--c-5f4b1e); white-space: nowrap; }
    .pdf-clip-box { flex: 1 1 240px; min-width: 0; height: 26px; border: 1px dashed var(--c-c8a84d); background: var(--c-fff5dc); }
    .pdf-status { position: absolute; }
    textarea, .generated-code, .output-body, .pca-variance-prose { scrollbar-width: none; }`;
  document.head.appendChild(style);
  // A copy box under each panel heading that holds a Copy button.
  root.querySelectorAll(".panel, .bench > article, .bench > section").forEach(panel => {
    if (panel.querySelector(":scope > .pdf-clip")) return;
    const copyButtons = Array.from(panel.querySelectorAll("button")).filter(button => /^Copy\b/i.test(button.textContent.trim()) || button.id === "copyBtn");
    if (!copyButtons.length || panel.querySelector(".panel .pdf-clip")) return;
    if (Array.from(panel.querySelectorAll(".panel")).some(inner => inner.contains(copyButtons[0]))) return;
    const clip = document.createElement("div");
    clip.className = "pdf-clip";
    clip.dataset.pdfClip = String(root.querySelectorAll(".pdf-clip").length);
    // Acrobat cannot put text on the clipboard or select a whole field from a script.
    clip.innerHTML = '<span class="pdf-clip-hint">Copied text appears here. Click it, press Ctrl+A, then Ctrl+C.</span><div class="pdf-clip-box"></div>';
    const head = panel.querySelector(":scope > .panel-head");
    if (head) head.after(clip); else panel.prepend(clip);
  });
  const version = document.querySelector(".app-version");
  if (version) version.textContent = config.version;
  materialisePseudoElements();
  window.__pdfRefs = refs;
  return { refs: refs.size };
}

// CSS-generated content (the search magnifier, the rule after "OPERATIONS", the document
// icon's fold and lines) is not in the DOM. Each visible ::before/::after is replaced by a
// real span with the same computed style, then measured and drawn like any element.
const PSEUDO_PROPERTIES = ["position", "left", "top", "right", "bottom", "width", "height", "minWidth", "minHeight", "display", "flex", "flexGrow", "flexShrink", "flexBasis", "alignSelf",
  "margin", "marginTop", "marginRight", "marginBottom", "marginLeft", "padding", "fontSize", "fontWeight", "fontFamily", "lineHeight", "color", "backgroundColor", "borderRadius",
  "border", "boxShadow", "zIndex", "transform", "verticalAlign", "letterSpacing", "textTransform", "opacity", "pointerEvents", "whiteSpace"];
function materialisePseudoElements() {
  const skip = "input, select, textarea, .sidebar-resizer, #themeToggle, .lp-collapse-toggle, .variable-rename-alert, .reference";
  const hosts = [];
  document.querySelectorAll("body *").forEach(element => {
    if (element.closest(skip)) return;
    ["::before", "::after"].forEach(which => {
      const style = getComputedStyle(element, which);
      const content = style.content;
      if (!content || content === "none" || content === "normal" || style.display === "none") return;
      const text = /^["']/.test(content) ? content.slice(1, -1).replace(/\\([0-9a-fA-F]{1,6}) ?/g, (_, hex) => String.fromCodePoint(parseInt(hex, 16))).replace(/\\(.)/g, "$1") : "";
      if (!text && !colour(style.backgroundColor) && style.boxShadow === "none") return;
      const span = document.createElement("span");
      span.className = "pdf-pseudo";
      span.textContent = text;
      PSEUDO_PROPERTIES.forEach(property => { span.style[property] = style[property]; });
      hosts.push({ element, which, span });
    });
  });
  hosts.forEach(({ element, which, span }) => {
    element.classList.add(which === "::before" ? "pdf-no-before" : "pdf-no-after");
    if (which === "::before") element.prepend(span); else element.append(span);
  });
  const style = document.createElement("style");
  style.textContent = ".pdf-no-before::before, .pdf-no-after::after { content: none !important; display: none !important; }";
  document.head.appendChild(style);
}

function colour(value) {
  const match = /rgba?\(([^)]+)\)/.exec(value || "");
  if (!match) return null;
  const parts = match[1].split(/[,\s/]+/).filter(Boolean).map(Number);
  if (parts.length > 3 && parts[3] === 0) return null;
  return { r: parts[0] / 255, g: parts[1] / 255, b: parts[2] / 255, a: parts.length > 3 ? parts[3] : 1 };
}

function box(element) {
  const rect = element.getBoundingClientRect();
  return { x: rect.left + window.scrollX, y: rect.top + window.scrollY, w: rect.width, h: rect.height };
}

function textStyle(style) {
  return {
    size: parseFloat(style.fontSize), weight: Number(style.fontWeight) || 400, colour: colour(style.color),
    spacing: style.letterSpacing === "normal" ? 0 : parseFloat(style.letterSpacing) || 0,
    transform: style.textTransform, align: style.textAlign
  };
}

// Words of a text node with their positions, one entry per word. Words with a character
// the embedded font cannot draw are kept apart as icons, to be copied from the browser.
let drawable = null, icons = null;
function words(node, style, out, owner) {
  const text = node.data;
  const pattern = /\S+/g;
  let match;
  const range = document.createRange();
  while ((match = pattern.exec(text))) {
    range.setStart(node, match.index);
    range.setEnd(node, match.index + match[0].length);
    const rects = range.getClientRects();
    if (!rects.length) continue;
    const rect = rects[0];
    if (rect.width === 0) continue;
    let word = match[0];
    if (style.transform === "uppercase") word = word.toUpperCase();
    if (drawable && Array.from(word).some(char => !drawable.has(char.codePointAt(0)))) {
      icons.push({ x: rect.left + window.scrollX - 1, y: rect.top + window.scrollY - 1, w: rect.width + 2, h: rect.height + 2, owner });
      continue;
    }
    out.push({ x: rect.left + window.scrollX, y: rect.top + window.scrollY, h: rect.height, text: word, style, owner });
  }
}

export function measure(config) {
  const refs = window.__pdfRefs;
  const conditional = new Set(config.conditional), dynamic = new Set(config.dynamic);
  const result = { width: document.documentElement.scrollWidth, height: document.documentElement.scrollHeight, background: colour(getComputedStyle(document.body).backgroundColor), boxes: [], words: [], fields: [], links: [], icons: [], hovers: 0 };
  drawable = new Set(config.drawable || []);
  icons = result.icons;
  // Elements whose hover style is read later by hovering over them in the browser.
  const hoverTarget = element => { element.dataset.pdfHover = String(result.hovers); return result.hovers++; };
  const root = bench();
  const clipFor = element => {
    if (element.classList.contains("pdf-clip")) return element.dataset.pdfClip;
    if (element.classList.contains("pdf-clip-box")) return element.parentNode.dataset.pdfClip;
    for (let node = element; node && node !== document.body; node = node.parentElement) {
      const clip = node.querySelector(".pdf-clip");
      if (clip) return clip.dataset.pdfClip;
    }
    return null;
  };
  const field = (kind, element, extra) => {
    const style = getComputedStyle(element);
    const ref = refs.get(element);
    const backdrop = ["text", "search", "select", "plot", "out"].indexOf(kind) >= 0 && !(ref && conditional.has(ref)) && !element.closest("#columnFields, #lpConstraintRows");
    if (backdrop) drawBox(element, style, null);
    result.fields.push(Object.assign({
      k: kind, r: refs.get(element) || null, rect: box(element),
      text: textStyle(style), fill: colour(style.backgroundColor), border: colour(style.borderTopColor), borderWidth: parseFloat(style.borderTopWidth) || 0,
      caption: element.localName === "button" ? element.textContent.replace(/\s+/g, " ").trim() : undefined,
      label: element.getAttribute("aria-label") || element.getAttribute("placeholder") || undefined,
      multiline: element.localName === "textarea" || kind === "out" || kind === "cardbody",
      clip: kind === "button" || kind === "clip" ? clipFor(element) : null,
      backdrop,
      inset: ["Top", "Right", "Bottom", "Left"].map(side => (parseFloat(style["padding" + side]) || 0) + (parseFloat(style["border" + side + "Width"]) || 0)),
      value: kind === "note" ? element.textContent.replace(/\s+/g, " ").trim() : undefined,
      placeholder: (kind === "text" || kind === "search") && element.getAttribute("placeholder") ? { text: element.getAttribute("placeholder"), colour: colour(getComputedStyle(element, "::placeholder").color) } : undefined
    }, extra || {}));
  };
  const drawBox = (element, style, owner) => {
    const rect = box(element);
    if (rect.w < 0.5 || rect.h < 0.5) return;
    const fill = colour(style.backgroundColor);
    const sides = ["Top", "Right", "Bottom", "Left"].map(side => {
      const width = parseFloat(style["border" + side + "Width"]) || 0;
      return width && style["border" + side + "Style"] !== "none" ? { width, colour: colour(style["border" + side + "Color"]) } : null;
    });
    const radius = parseFloat(style.borderTopLeftRadius) || 0;
    if (fill || sides.some(side => side && side.colour)) result.boxes.push({ rect, fill, sides, radius: Math.min(radius, rect.h / 2, rect.w / 2), owner });
    // Hard-edged shadows ("0 4px 0 white") repeat the box; soft shadows are left out.
    if (style.boxShadow && style.boxShadow !== "none") {
      style.boxShadow.split(/,(?![^(]*\))/).forEach(shadow => {
        const shade = colour(shadow);
        const numbers = (shadow.replace(/rgba?\([^)]*\)/, "").match(/-?[\d.]+px/g) || []).map(parseFloat);
        if (!shade || shade.a < 0.5 || numbers.length < 2 || (numbers[2] || 0) > 0 || /inset/.test(shadow)) return;
        result.boxes.push({ rect: { x: rect.x + numbers[0], y: rect.y + numbers[1], w: rect.w, h: rect.h }, fill: shade, sides: [], radius: Math.min(radius, rect.h / 2, rect.w / 2), owner });
      });
    }
  };

  const walk = (element, owner) => {
    const style = getComputedStyle(element);
    if (style.display === "none" || style.visibility === "hidden") return;
    if (element.matches(".sr-only")) return;
    const ref = refs.get(element);
    const inBench = root.contains(element);
    // Text inside elements that come and go becomes a field that can be hidden with them.
    const nextOwner = inBench && ref && conditional.has(ref) ? ref : owner;

    if (element.classList.contains("pdf-clip-box")) { field("clip", element); return; }
    if (element.id === "searchInput") { field("search", element); return; }
    if (element.matches(".operation, .universal-entry")) {
      result.links.push({ rect: box(element), op: element.classList.contains("universal-entry") ? "universal" : element.dataset.id, hover: hoverTarget(element), radius: parseFloat(getComputedStyle(element).borderTopLeftRadius) || 0 });
    }
    if (inBench) {
      if (element.matches('input[type="checkbox"]')) { field("check", element, { switch: element.matches('[role="switch"], .kmeans-scale-toggle input') }); return; }
      if (element.id === "plotWorkflowSearch") { field("plot", element); return; }
      if (element.matches("input, textarea")) { field("text", element); return; }
      if (element.matches("select")) { field("select", element, { variant: element.classList.contains("answer-variant-select") }); return; }
      if (element.localName === "button") { field("button", element, { hover: element.disabled ? null : hoverTarget(element), radius: parseFloat(style.borderTopLeftRadius) || 0 }); return; }
      if (element.matches(CARD)) {
        const grid = element.parentNode, index = Array.prototype.indexOf.call(grid.querySelectorAll(CARD), element);
        const gridRef = refs.get(grid), cards = grid.querySelectorAll(CARD).length;
        drawBox(element, style, nextOwner);
        const head = element.querySelector(".sensitivity-prose-head");
        const heading = head && head.querySelector("h3"), button = head && head.querySelector("button");
        if (heading) field("note", heading, { r: `row:${gridRef}~${CARD}~${index}~h3` });
        if (button) field("button", button, { r: `row:${gridRef}~${CARD}~${index}~button` });
        const top = head ? box(head).y + box(head).h : box(element).y;
        const outer = box(element);
        const firstText = Array.from(element.children).filter(child => child !== head)[0];
        const textStyleSource = firstText || element;
        result.fields.push({ k: "cardbody", r: `slot:${gridRef}~${CARD}~${index}`, last: index === cards - 1, rect: { x: outer.x + 12, y: top + 4, w: outer.w - 24, h: Math.max(28, outer.y + outer.h - top - 12) }, text: textStyle(getComputedStyle(textStyleSource)), fill: null, multiline: true });
        return;
      }
      if (element.matches(OUTPUT) && !element.querySelector(CARD)) {
        const inner = element.querySelector(".prose-card p, p") || element;
        field("out", element, { text: textStyle(getComputedStyle(element.matches("pre") ? element : inner)), code: element.matches("pre") });
        return;
      }
      // Live text: only small text elements, never containers with their own headers.
      const textOnly = Array.from(element.children).every(child => /^(strong|em|b|i|code|span|small|sub|sup|br)$/.test(child.localName) && !child.children.length);
      if (ref && dynamic.has(ref) && textOnly && !element.querySelector("input, select, textarea, button")) { field("note", element, { owner: nextOwner }); return; }
    }
    if (!(inBench && nextOwner)) drawBox(element, style, nextOwner);
    const textOwner = inBench ? nextOwner : null;
    element.childNodes.forEach(child => {
      if (child.nodeType === 3 && child.data.trim()) words(child, textStyle(style), result.words, textOwner);
      else if (child.nodeType === 1) walk(child, nextOwner);
    });
  };
  walk(document.body, null);
  // Panels that complete on copy get a "Done!" badge beside their title.
  root.querySelectorAll(".lp-collapsible").forEach(panel => {
    const title = panel.querySelector(":scope > .panel-head .panel-title strong");
    if (!title) return;
    const rect = box(title);
    result.fields.push({ k: "badge", r: refs.get(panel), rect: { x: rect.x + rect.w + 10, y: rect.y - 3, w: 56, h: rect.h + 6 }, text: { size: 12, weight: 800, colour: { r: 1, g: 1, b: 1 } } });
  });
  const toggle = document.querySelector("#themeToggle");
  const bar = document.querySelector(".topbar, header");
  const barRect = bar ? box(bar) : { x: 0, y: 0, w: result.width, h: 60 };
  result.status = { x: result.width - 560, y: barRect.y + barRect.h / 2 - 15, w: 532, h: 30 };
  result.toggleHidden = Boolean(toggle);
  return result;
}

// Colours the PDF applies itself (completed cards, errors, the switch, the copy box …),
// read from the website's tokens so each theme gets its own values.
const PALETTE_TOKENS = {
  done: "--c-e9f6ed", errorFill: "--c-fce8e6", active: "--c-d3e3fd", disabled: "--c-f1f3f4", ink: "--c-1f1f1f", muted: "--c-5f6368",
  error: "--c-b3261e", badge: "--c-2f7d52", badgeText: "--c-ffffff", accent: "--accent", track: "--c-dadce0", trackBorder: "--c-80868b",
  clipFill: "--c-fff5dc", clipBorder: "--c-c8a84d", panel: "--panel"
};
export function palette() {
  const probe = document.createElement("span");
  document.body.appendChild(probe);
  const result = {};
  Object.keys(PALETTE_TOKENS).forEach(name => {
    probe.style.color = "";
    probe.style.color = `var(${PALETTE_TOKENS[name]})`;
    const value = colour(getComputedStyle(probe).color);
    if (value) result[name] = [value.r, value.g, value.b].map(part => Math.round(part * 1000) / 1000);
  });
  probe.remove();
  return result;
}

export function hoverStyle(index) {
  const element = document.querySelector(`[data-pdf-hover="${index}"]`);
  if (!element) return null;
  const style = getComputedStyle(element);
  return { fill: colour(style.backgroundColor), border: colour(style.borderTopColor), text: colour(style.color) };
}
