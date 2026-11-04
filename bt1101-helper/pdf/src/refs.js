// Shared by the build-time layout and the in-Acrobat bridge, so a field's reference
// resolves to the same element in both places.
//
// Reference forms:
//   "#id"                              element by id
//   "q:<scope>|<selector>|<n>"         nth match of selector inside scope ("" = document)
//   "slot:<container>~<selector>~<n>"  nth dynamic child created by the website
//   "row:<container>~<rowSel>~<n>~<childSel>"  child of the nth dynamic row
// slot and row use "~" so the container may itself be a "q:" reference.

export function resolveRef(document, ref) {
  if (!ref) return null;
  if (ref[0] === "#") return document.getElementById(ref.slice(1));
  const colon = ref.indexOf(":");
  const kind = ref.slice(0, colon);
  const parts = ref.slice(colon + 1).split(kind === "q" ? "|" : "~");
  if (kind === "q" || kind === "slot") {
    const scope = parts[0] ? resolveRef(document, parts[0]) : document;
    if (!scope) return null;
    return scope.querySelectorAll(parts[1])[Number(parts[2])] || null;
  }
  if (kind === "row") {
    const container = resolveRef(document, parts[0]);
    if (!container) return null;
    const row = container.querySelectorAll(parts[1])[Number(parts[2])];
    return row ? row.querySelector(parts[3]) : null;
  }
  return null;
}

// Classes the website toggles to show state. References ignore them so an element keeps
// the same reference whether it is active, hidden, completed or showing an error.
const STATE_CLASS = /^(active|hidden|show|open|dragging|inapplicable|waiting-note|empty|input-error|is-[\w-]+|has-[\w-]+)$/;

function escapeIdentifier(value) {
  return String(value).replace(/([^A-Za-z0-9_-])/g, "\\$1").replace(/^(\d)/, "\\3$1 ");
}

export function selectorFor(element) {
  const classes = (element.getAttribute("class") || "").split(/\s+/).filter(name => name && !STATE_CLASS.test(name));
  return element.localName + classes.map(name => "." + escapeIdentifier(name)).join("");
}

export function refFor(element, document) {
  if (element.id && document.querySelectorAll("#" + escapeIdentifier(element.id)).length === 1) return "#" + element.id;
  let scope = element.parentNode;
  while (scope && scope.nodeType === 1 && !(scope.id && document.querySelectorAll("#" + escapeIdentifier(scope.id)).length === 1)) scope = scope.parentNode;
  const root = scope && scope.nodeType === 1 ? scope : document;
  const selector = selectorFor(element);
  const matches = Array.prototype.slice.call(root.querySelectorAll(selector));
  return "q:" + (root === document ? "" : "#" + scope.id) + "|" + selector + "|" + matches.indexOf(element);
}

// Puts one control into a given state the way a user would; disabled controls and
// options are left alone, as a user could not change them. Returns whether it applied.
export function applySetting(document, setting, Event) {
  const fire = (target, type) => target.dispatchEvent(new Event(type, { bubbles: true }));
  if (setting.type === "plot") {
    const search = document.getElementById("plotWorkflowSearch");
    if (!search || search.value.trim() === setting.value) return false;
    search.value = setting.value; fire(search, "input"); fire(search, "change");
    return true;
  }
  const target = resolveRef(document, setting.ref);
  if (!target || target.disabled) return false;
  if (setting.type === "select") {
    const option = Array.prototype.filter.call(target.querySelectorAll("option"), item => item.value === setting.value)[0];
    if (!option || option.disabled || target.value === setting.value) return false;
    target.value = setting.value; fire(target, "input"); fire(target, "change");
    return true;
  }
  if (setting.type === "check") { if (target.checked === setting.value) return false; target.click(); return true; }
  if (setting.type === "metric") { if (target.classList.contains("active")) return false; target.click(); return true; }
  return false;
}

// Confirms whichever warning dialog the website opened (for example, disabling K-means
// scaling), so the state asked for is the state reached.
export function acceptAlerts(document) {
  for (let guard = 0; guard < 5; guard++) {
    const alert = document.querySelector(".variable-rename-alert:not(.hidden)");
    if (!alert) return;
    const buttons = alert.querySelectorAll("button");
    buttons[buttons.length - 1].click();
  }
}

// The CSS rules in index.html that hide content, expressed as element checks. The PDF
// keeps completed cards open (it cannot reflow), so collapsing is ignored on request.
export function isRendered(element, keepCollapsed) {
  if (!element || !element.isConnected) return false;
  for (let node = element; node && node.nodeType === 1; node = node.parentNode) {
    if (node.hasAttribute("hidden") || node.classList.contains("hidden")) return false;
    if (node.classList.contains("metric-tooltip")) return false;
    const parent = node.parentNode;
    if (parent && parent.nodeType === 1) {
      if (!keepCollapsed && parent.classList.contains("lp-collapsible") && parent.classList.contains("is-collapsed") && !node.classList.contains("panel-head")) return false;
      if (parent.localName === "details" && !parent.hasAttribute("open") && node.localName !== "summary") return false;
    }
  }
  if (element.matches(".lp-constraint-error, .test-interpretation") && !element.textContent.trim()) return false;
  return true;
}

const BLOCK_TAGS = { div: 1, p: 1, section: 1, article: 1, header: 1, footer: 1, h1: 1, h2: 1, h3: 1, h4: 1, h5: 1, pre: 1, ul: 1, ol: 1, li: 1, table: 1, tr: 1, thead: 1, tbody: 1, label: 1, details: 1, summary: 1 };
const GRID_CARDS = ".shape-guide, .hypothesis-card, .mapping-card, .route-card, .empty-state-inner, .prose-card, .lp-parser-status";
const SKIPPED_TEXT = "button, select, input, textarea, .answer-variant-control, .sr-only, .metric-tooltip, .copyable-stage-status, script, style";

// Approximates innerText for the output containers: block boundaries become line breaks,
// <pre> keeps its spacing, controls are left to their own PDF fields.
export function visibleText(element, extraSkip) {
  if (!element) return "";
  const skipped = extraSkip ? SKIPPED_TEXT + ", " + extraSkip : SKIPPED_TEXT;
  let text = "";
  const pre = [];
  const newline = count => {
    const trailing = /\n*$/.exec(text)[0].length;
    if (text && trailing < count) text += count - trailing === 2 ? "\n\n" : "\n";
  };
  const walk = node => {
    if (node.nodeType === 3) {
      if (pre.length) text += node.data;
      else {
        const collapsed = node.data.replace(/[ \t\r\n\f]+/g, " ");
        if (collapsed === " " && (!text || /[\s]$/.test(text))) return;
        text += /\n$/.test(text) || !text ? collapsed.replace(/^ /, "") : collapsed;
      }
      return;
    }
    if (node.nodeType !== 1) return;
    if (node !== element && (node.matches(skipped) || !isRendered(node, true))) return;
    const tag = node.localName;
    if (tag === "br") { text += "\n"; return; }
    // Children of the website's grid cards sit on their own lines.
    const block = BLOCK_TAGS[tag] || Boolean(node.parentNode && node.parentNode.nodeType === 1 && node.parentNode.matches(GRID_CARDS));
    const paragraph = tag === "p" || tag === "pre" || tag === "section" || tag === "article";
    if (block) newline(paragraph ? 2 : 1);
    if (tag === "pre") pre.push(node);
    if (tag === "td" || tag === "th") { if (node.previousElementSibling) text += " | "; }
    // Fields cannot lower text, so a subscript is written as the website's ASCII copy writes it (X_i).
    if (tag === "sub") text += "_";
    for (let i = 0; i < node.childNodes.length; i++) walk(node.childNodes[i]);
    if (tag === "pre") pre.pop();
    if (block) newline(paragraph ? 2 : 1);
  };
  walk(element);
  return text.replace(/[ \t]+\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim();
}

export function buttonCaption(button) {
  let caption = button.textContent.replace(/\s+/g, " ").trim();
  if (!/[A-Za-z0-9]/.test(caption)) {
    const label = button.getAttribute("aria-label") || button.getAttribute("title") || "";
    if (button.classList.contains("lp-collapse-toggle")) caption = /^Expand/i.test(label) ? "Expand" : "Collapse";
    else if (button.id === "resetBtn") caption = "Reset";
    else if (button.id === "copyBtn") caption = "Copy prose";
    else if (button.classList.contains("lp-remove-constraint") || button.classList.contains("remove-column")) caption = "Remove";
    else caption = label || caption;
  }
  // Section copy buttons all read "Copy"; name the section they belong to.
  const card = button.closest(".sensitivity-prose-card, .pca-code-block");
  const heading = card && card.querySelector("h3");
  if (heading && /^(Copy|Copied)$/i.test(caption)) caption += ": " + heading.textContent.replace(/\s+/g, " ").trim();
  return caption;
}
