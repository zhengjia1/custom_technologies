// Build time: turns the website's <style> into a compact, line-based rule table that the
// VBA style engine (modStyle) loads at startup without parsing CSS itself.
//
//   node tools/css-rules.js [index.html] [out.txt]      (default ../index.html -> .build/css-rules.txt)
//
// What happens here so VBA does not have to:
//   - @font-face blocks and comments are dropped; @media/@container conditions are kept
//     as numeric width ranges (prefers-* features evaluate as in a default browser).
//   - var() tokens defined on :root are resolved for the light theme and for
//     :root[data-theme="dark"]; a declaration whose value differs between themes carries
//     both values. var() names that are not defined on :root (only --sidebar-width, set
//     inline at run time) stay in the value for VBA to resolve.
//   - Shorthands are expanded to the longhands the layout engine reads; everything else
//     (shadows, radii, transitions, cursors, font families...) is dropped. Colours are
//     normalised to #rrggbb or #rrggbbaa.
//   - Selectors are compiled (specificity, subject key) and rules sorted into cascade
//     order (browser defaults first, then specificity, then source order), so VBA only
//     matches (clsNode.Matches with the selector text) and applies them in index order.
//     Rules that can never match in Excel (:hover, :focus-within, ::placeholder, vendor
//     pseudo-elements, functional pseudo-classes other than :not) are left out.
//   - The Unicode coverage of the regular Source Sans Pro face is listed so the renderer
//     can draw characters the font lacks with a symbol font.
//
// Output format (UTF-8, one record per line, fields separated by TAB):
//   V  <format version>
//   P  <property name>                                (property table, index = order)
//   C  <id> <parent id or 0> <m|c> <alternatives>     media / container condition;
//      alternatives are ";"-separated "min:max" pixel ranges (-1 = unbounded, "x" = never)
//   B  <id>                                           declaration block, followed by
//   D  <property index> <important 0|1> <value> [<dark value>]
//   R  <block id> <condition id or 0> <pseudo 0|1|2> <specificity> <key> <simple 0|1> <selector>
//      key = "#id", ".class", tag or "*" (the subject's bucket); simple = 1 when the
//      selector is exactly the key (a bucket hit is a match); selector = text for
//      clsNode.Matches (pseudo-elements removed: pseudo 1 = ::before, 2 = ::after).
//   G  <ranges>                                       font coverage: "start-end" hex code
//      point ranges separated by ","
//   E  end
// Rules appear in cascade order. Values are escaped: "\t" -> "\\t", "\n" -> "\\n", "\\" -> "\\\\".

const fs = require("fs");
const path = require("path");

const ROOT = path.join(__dirname, "..");
const FORMAT_VERSION = 2;

// Browser defaults the site relies on (same as the reference engine's UA sheet).
const UA_CSS = `
html, body, div, p, section, article, header, footer, main, aside, nav, form, fieldset, legend, pre, details, h1, h2, h3, h4, h5, h6, ul, ol, dl, dt, dd, figure, blockquote, hr, address { display: block; }
li { display: list-item; } summary { display: list-item; }
head, script, style, template, option, optgroup, datalist, title, meta, link, [hidden], input[type="hidden"] { display: none; }
span, a, strong, b, em, i, small, code, kbd, samp, sub, sup, abbr, label, mark, u, s, q, cite, var, time, data, output, br, wbr { display: inline; }
button, input, select, textarea, meter, progress { display: inline-block; }
table { display: table; } thead { display: table-header-group; } tbody { display: table-row-group; } tfoot { display: table-footer-group; }
tr { display: table-row; } td, th { display: table-cell; } caption { display: table-caption; } col, colgroup { display: none; }
body { margin: 8px; }
p, ul, ol, dl, blockquote, figure { margin: 1em 0; }
h1 { font-size: 2em; margin: .67em 0; font-weight: 700; } h2 { font-size: 1.5em; margin: .83em 0; font-weight: 700; }
h3 { font-size: 1.17em; margin: 1em 0; font-weight: 700; } h4 { margin: 1.33em 0; font-weight: 700; }
h5 { font-size: .83em; margin: 1.67em 0; font-weight: 700; } h6 { font-size: .67em; margin: 2.33em 0; font-weight: 700; }
strong, b, th, legend { font-weight: 700; } em, i, cite, var { font-style: italic; } small { font-size: .833em; }
pre { white-space: pre; margin: 1em 0; }
ul, ol { padding-left: 40px; } ul { list-style-type: disc; } ol { list-style-type: decimal; }
fieldset { margin: 0 2px; padding: .35em .75em .625em; border: 2px solid #c0c0c0; } legend { padding: 0 2px; }
th { text-align: center; } td, th { padding: 1px; } table { border-spacing: 2px; }
button { padding: 1px 6px; border: 2px solid #767676; background-color: #efefef; text-align: center; }
input { padding: 1px 2px; border: 2px solid #767676; background-color: #ffffff; }
input[type="checkbox"], input[type="radio"] { width: 13px; height: 13px; margin: 3px 3px 3px 4px; padding: 0; border: 0; }
select { padding: 1px 0; border: 1px solid #767676; background-color: #ffffff; font-size: 13.333px; }
textarea { padding: 2px; border: 1px solid #767676; background-color: #ffffff; white-space: pre-wrap; }
input, select, textarea, button { font-size: 13.333px; font-weight: 400; letter-spacing: normal; text-transform: none; text-align: start; }
button { text-align: center; }
button:focus-visible, select:focus-visible, summary:focus-visible, input[type="checkbox"]:focus-visible, input[type="radio"]:focus-visible { outline: 2px solid #1a73e8; outline-offset: 1px; }
`;

// Longhands the VBA layout engine reads (index = property id in VBA, see modStyle's P_*
// constants, which must list the same names in the same order).
const PROPERTIES = [
  "display", "position", "top", "right", "bottom", "left", "z-index",
  "width", "height", "min-width", "min-height", "max-width", "max-height", "box-sizing",
  "margin-top", "margin-right", "margin-bottom", "margin-left",
  "padding-top", "padding-right", "padding-bottom", "padding-left",
  "border-top-width", "border-right-width", "border-bottom-width", "border-left-width",
  "border-top-style", "border-right-style", "border-bottom-style", "border-left-style",
  "border-top-color", "border-right-color", "border-bottom-color", "border-left-color",
  "color", "background-color", "opacity", "visibility",
  "font-size", "font-weight", "font-style", "line-height", "text-align", "text-transform",
  "white-space", "overflow-wrap", "word-break", "letter-spacing", "list-style-type",
  "overflow-x", "overflow-y",
  "flex-direction", "flex-wrap", "flex-grow", "flex-shrink", "flex-basis",
  "align-items", "align-self", "align-content", "justify-content", "justify-items", "justify-self",
  "row-gap", "column-gap", "grid-template-columns", "grid-column-start", "grid-column-end",
  "container-type", "content", "transform", "clip", "border-collapse", "border-spacing", "text-overflow",
  "grid-template-rows",
  // paint only: zero-blur box shadows (copies and spread rings) and focus outlines
  "box-shadow", "outline-width", "outline-style", "outline-color", "outline-offset"
];
const PROPERTY_INDEX = {};
PROPERTIES.forEach((name, index) => { PROPERTY_INDEX[name] = index; });
const COLOUR_PROPERTIES = /^(color|background-color|border-(top|right|bottom|left)-color|outline-color)$/;

// ---------------------------------------------------------------- parsing helpers

function stripComments(text) { return text.replace(/\/\*[\s\S]*?\*\//g, ""); }

function splitTop(text, separator) {
  const parts = [];
  let depth = 0, quote = null, start = 0;
  for (let i = 0; i < text.length; i++) {
    const char = text[i];
    if (quote) { if (char === "\\") i++; else if (char === quote) quote = null; continue; }
    if (char === '"' || char === "'") quote = char;
    else if (char === "(" || char === "[") depth++;
    else if (char === ")" || char === "]") depth--;
    else if (depth === 0 && (separator === " " ? /\s/.test(char) : char === separator)) { parts.push(text.slice(start, i)); start = i + 1; }
  }
  parts.push(text.slice(start));
  return separator === " " ? parts.filter(part => part !== "") : parts;
}

function findBlockEnd(text, open) {
  let depth = 0, quote = null;
  for (let i = open; i < text.length; i++) {
    const char = text[i];
    if (quote) { if (char === "\\") i++; else if (char === quote) quote = null; continue; }
    if (char === '"' || char === "'") quote = char;
    else if (char === "{") depth++;
    else if (char === "}") { depth--; if (depth === 0) return i; }
  }
  return text.length;
}

function parseDeclarations(text) {
  const declarations = [];
  splitTop(text, ";").forEach(part => {
    const colon = part.indexOf(":");
    if (colon < 0) return;
    const property = part.slice(0, colon).trim();
    let value = part.slice(colon + 1).trim();
    const important = /!\s*important\s*$/i.test(value);
    if (important) value = value.replace(/!\s*important\s*$/i, "").trim();
    if (property && value !== "") declarations.push({ property: property.indexOf("--") === 0 ? property : property.toLowerCase(), value, important });
  });
  return declarations;
}

// ---------------------------------------------------------------- selectors

const UNSUPPORTED_PSEUDO = /^(hover|active|visited|focus-within|placeholder|placeholder-shown|selection|marker|backdrop|first-line|first-letter|target|autofill)$/i;
// :focus and :focus-visible match the element with the window's focus (modRender restyles
// it when focus moves).
const SUPPORTED_PSEUDO = /^(first-child|last-child|only-child|checked|disabled|enabled|empty|root|focus|focus-visible)$/i;
const IDENT = /^-?[A-Za-z_ -￿][\w -￿-]*|^--[\w-]*/;

// Returns { tokens, spec, key, pseudo, base } or null when the selector can never match in
// Excel. base is the selector text without its pseudo-element.
function compileSelector(text) {
  let source = text.trim().replace(/\s+/g, " ");
  let pseudo = 0;
  const pe = /::?(before|after)\s*$/i.exec(source);
  if (pe) {
    pseudo = pe[1].toLowerCase() === "before" ? 1 : 2;
    source = source.slice(0, pe.index).trim();
    if (!source || /[>+~]$/.test(source)) source += (source ? " " : "") + "*";
  }
  const state = { text: source, i: 0 };
  const tokens = [];
  let spec = 0;
  let lastCompound = null;
  try {
    skipSpace(state);
    for (;;) {
      const compound = parseCompound(state);
      if (!compound) return null;
      compound.tokens.forEach(token => tokens.push(token));
      spec += compound.spec;
      lastCompound = compound;
      const hadSpace = skipSpace(state);
      if (state.i >= state.text.length) break;
      const char = state.text[state.i];
      if (char === ">" || char === "+" || char === "~") { tokens.push(char); state.i++; skipSpace(state); }
      else if (hadSpace) tokens.push("_");
      else throw new Error("Unexpected " + char + " in " + text);
    }
  } catch (error) {
    if (error.unsupported) return null;
    throw error;
  }
  const subject = lastCompound;
  const key = subject.id ? "#" + subject.id : subject.classes.length ? "." + subject.classes[0] : subject.tag ? subject.tag : "*";
  const only = tokens.length === 1 ? tokens[0] : null;
  const simple = only !== null && (only === "*" || only === key || only === "T" + key) ? 1 : 0;
  return { tokens, spec, key, pseudo, base: source, simple };
}

function skipSpace(state) {
  const start = state.i;
  while (state.i < state.text.length && /\s/.test(state.text[state.i])) state.i++;
  return state.i > start;
}

function readIdent(state) {
  const match = IDENT.exec(state.text.slice(state.i));
  if (!match) throw new Error("Expected identifier at " + state.text.slice(state.i));
  state.i += match[0].length;
  return match[0];
}

function unsupported(message) { const error = new Error(message); error.unsupported = true; return error; }

function parseCompound(state, inner) {
  const compound = { tokens: [], spec: 0, tag: null, id: null, classes: [] };
  let any = false;
  for (;;) {
    if (state.i >= state.text.length) break;
    const char = state.text[state.i];
    if (char === "*") { state.i++; any = true; continue; }
    if (/[A-Za-z]/.test(char) && !any && !compound.tokens.length) {
      compound.tag = readIdent(state).toLowerCase();
      compound.tokens.push("T" + compound.tag);
      compound.spec += 1;
      any = true;
      continue;
    }
    if (char === "#") { state.i++; compound.id = readIdent(state); compound.tokens.push("#" + compound.id); compound.spec += 10000; any = true; continue; }
    if (char === ".") { state.i++; const name = readIdent(state); compound.classes.push(name); compound.tokens.push("." + name); compound.spec += 100; any = true; continue; }
    if (char === "[") {
      const close = state.text.indexOf("]", state.i);
      const body = state.text.slice(state.i + 1, close).trim();
      state.i = close + 1;
      const match = /^([\w-]+)\s*(?:([~^$*|]?=)\s*(?:"([^"]*)"|'([^']*)'|([^\s\]]+)))?\s*(i)?$/.exec(body);
      if (!match) throw new Error("Bad attribute selector [" + body + "]");
      const op = match[2] ? (match[2] === "=" ? "=" : match[2][0]) : "E";
      const value = match[3] !== undefined ? match[3] : match[4] !== undefined ? match[4] : match[5] !== undefined ? match[5] : "";
      compound.tokens.push("A" + op + match[1].toLowerCase() + (op === "E" ? "" : "=" + value));
      compound.spec += 100;
      any = true;
      continue;
    }
    if (char === ":") {
      if (state.text[state.i + 1] === ":") throw unsupported("pseudo-element");
      state.i++;
      const name = readIdent(state).toLowerCase();
      if (name === "not") {
        if (state.text[state.i] !== "(") throw new Error(":not without argument");
        state.i++;
        skipSpace(state);
        const argument = parseCompound(state, true);
        skipSpace(state);
        if (state.text[state.i] !== ")") throw new Error(":not with a complex argument is not supported");
        state.i++;
        compound.tokens.push("N(");
        argument.tokens.forEach(token => compound.tokens.push(token));
        compound.tokens.push("N)");
        compound.spec += argument.spec;
        any = true;
        continue;
      }
      if (state.text[state.i] === "(") throw unsupported("functional pseudo-class " + name);
      if (UNSUPPORTED_PSEUDO.test(name) || /^-/.test(name)) throw unsupported(name);
      if (!SUPPORTED_PSEUDO.test(name)) throw new Error("Unknown pseudo-class :" + name);
      compound.tokens.push("P" + name);
      compound.spec += 100;
      any = true;
      continue;
    }
    break;
  }
  if (!any && !inner) return null;
  if (!compound.tokens.length) compound.tokens.push("*");
  return compound;
}

// ---------------------------------------------------------------- conditions

function parseCondition(prelude, type) {
  const alternatives = splitTop(prelude, ",").map(alternative => {
    let min = -1, max = -1, impossible = false;
    const words = alternative.replace(/\([^)]*\)/g, " ").toLowerCase().split(/\s+/).filter(word => word && word !== "and" && word !== "only" && word !== "screen" && word !== "all");
    if (words.length) impossible = true;
    alternative.replace(/\(([^)]*)\)/g, (_, feature) => {
      const parts = feature.split(":");
      const name = parts[0].trim().toLowerCase(), value = (parts[1] || "").trim();
      const px = parseFloat(value) * (/em$/.test(value) ? 16 : 1);
      if (name === "max-width") max = max < 0 ? px : Math.min(max, px);
      else if (name === "min-width") min = Math.max(min, px);
      else if (name === "prefers-color-scheme") { if (value !== "light") impossible = true; }
      else if (name === "hover" || name === "pointer") { if (value === "none" || value === "coarse") impossible = true; }
      else impossible = true; // prefers-reduced-motion and anything else: not in a default browser
      return "";
    });
    return impossible ? "x" : min + ":" + max;
  });
  return { type, alternatives: alternatives.join(";") };
}

// ---------------------------------------------------------------- stylesheet

function parseStylesheet(text, ua, conditions) {
  const rules = [];
  const walk = (source, conditionId) => {
    let index = 0;
    while (index < source.length) {
      const open = source.indexOf("{", index);
      if (open < 0) break;
      const prelude = source.slice(index, open).trim();
      const close = findBlockEnd(source, open);
      const body = source.slice(open + 1, close);
      index = close + 1;
      if (prelude[0] === "@") {
        const name = /^@([\w-]+)/.exec(prelude)[1].toLowerCase();
        const query = prelude.slice(name.length + 1).trim();
        if (name === "media" || name === "container") {
          const condition = parseCondition(name === "container" ? query.replace(/^[A-Za-z_][\w-]*\s+(?=\()/, "") : query, name === "media" ? "m" : "c");
          condition.parent = conditionId;
          conditions.push(condition);
          walk(body, conditions.length);
        } else if (name === "supports") walk(body, conditionId);
        continue;
      }
      const selectorText = prelude.replace(/^@[^;]*;/, "").trim();
      const declarations = parseDeclarations(body);
      if (!declarations.length) continue;
      const block = { declarations };
      splitTop(selectorText, ",").forEach(selector => {
        const compiled = compileSelector(selector);
        if (compiled) rules.push(Object.assign(compiled, { block, order: rules.length, condition: conditionId, ua, text: selector.trim() }));
      });
    }
  };
  walk(stripComments(text), 0);
  return rules;
}

// ---------------------------------------------------------------- shorthands

const SIDES = ["top", "right", "bottom", "left"];
function fourSides(values) { const [a, b = a, c = a, d = b] = values; return [a, b, c, d]; }
const BORDER_STYLES = /^(none|hidden|solid|dashed|dotted|double|groove|ridge|inset|outset)$/i;
function isLengthToken(token) { return /^-?[\d.]/.test(token) || /^(thin|medium|thick)$/i.test(token) || /^(calc|min|max|clamp)\(/i.test(token); }
function borderParts(value) {
  if (value.trim() === "0" || /^none$/i.test(value.trim())) return { width: "0", style: "none", color: "currentcolor" };
  const result = {};
  splitTop(value, " ").forEach(token => {
    if (BORDER_STYLES.test(token)) result.style = token;
    else if (isLengthToken(token)) result.width = token;
    else result.color = token;
  });
  return { width: result.width || "medium", style: result.style || "none", color: result.color || "currentcolor" };
}

function expand(property, value, out) {
  const tokens = () => splitTop(value, " ");
  switch (property) {
    case "margin": case "padding": fourSides(tokens()).forEach((v, i) => { out[property + "-" + SIDES[i]] = v; }); return;
    case "margin-block": case "padding-block": { const t = tokens(), base = property.split("-")[0]; out[base + "-top"] = t[0]; out[base + "-bottom"] = t[1] || t[0]; return; }
    case "margin-inline": case "padding-inline": { const t = tokens(), base = property.split("-")[0]; out[base + "-left"] = t[0]; out[base + "-right"] = t[1] || t[0]; return; }
    case "inset": fourSides(tokens()).forEach((v, i) => { out[SIDES[i]] = v; }); return;
    case "border": { const p = borderParts(value); SIDES.forEach(side => { out["border-" + side + "-width"] = p.width; out["border-" + side + "-style"] = p.style; out["border-" + side + "-color"] = p.color; }); return; }
    case "border-width": case "border-style": case "border-color": { const kind = property.split("-")[1]; fourSides(tokens()).forEach((v, i) => { out["border-" + SIDES[i] + "-" + kind] = v; }); return; }
    case "border-top": case "border-right": case "border-bottom": case "border-left": { const p = borderParts(value); out[property + "-width"] = p.width; out[property + "-style"] = p.style; out[property + "-color"] = p.color; return; }
    case "background": {
      const layers = splitTop(value, ",");
      const last = splitTop(layers[layers.length - 1].trim(), " ");
      const colour = last.filter(token => /^(#|rgba?\(|hsla?\(|var\(|transparent$|white$|black$|currentcolor$)/i.test(token)).pop();
      out["background-color"] = colour || "transparent";
      return;
    }
    case "font": {
      if (/^(inherit|initial|unset)$/i.test(value.trim())) { ["font-size", "font-weight", "font-style", "line-height"].forEach(name => { out[name] = value.trim(); }); return; }
      const t = tokens();
      let i = 0;
      out["font-style"] = "normal"; out["font-weight"] = "400"; out["line-height"] = "normal";
      for (; i < t.length; i++) {
        if (/^(italic|oblique)$/i.test(t[i])) out["font-style"] = "italic";
        else if (/^(bold|bolder|lighter|\d{3})$/i.test(t[i])) out["font-weight"] = t[i];
        else if (/^(normal|small-caps)$/i.test(t[i])) continue;
        else break;
      }
      if (i < t.length) {
        const size = t[i].split("/");
        out["font-size"] = size[0];
        if (size[1]) out["line-height"] = size[1];
        else if (t[i + 1] === "/" && t[i + 2]) out["line-height"] = t[i + 2];
        else if (t[i + 1] && t[i + 1][0] === "/") out["line-height"] = t[i + 1].slice(1);
      }
      return;
    }
    case "flex": {
      const t = tokens();
      if (t.length === 1 && t[0] === "none") { out["flex-grow"] = "0"; out["flex-shrink"] = "0"; out["flex-basis"] = "auto"; return; }
      if (t.length === 1 && t[0] === "auto") { out["flex-grow"] = "1"; out["flex-shrink"] = "1"; out["flex-basis"] = "auto"; return; }
      if (t.length === 1 && /^[\d.]+$/.test(t[0])) { out["flex-grow"] = t[0]; out["flex-shrink"] = "1"; out["flex-basis"] = "0px"; return; }
      out["flex-grow"] = t[0]; out["flex-shrink"] = t[1] !== undefined && /^[\d.]+$/.test(t[1]) ? t[1] : "1";
      out["flex-basis"] = t.length === 3 ? t[2] : t[1] !== undefined && !/^[\d.]+$/.test(t[1]) ? t[1] : "0px";
      return;
    }
    case "flex-flow": tokens().forEach(token => { if (/wrap/.test(token)) out["flex-wrap"] = token; else out["flex-direction"] = token; }); return;
    case "gap": case "grid-gap": { const t = tokens(); out["row-gap"] = t[0]; out["column-gap"] = t[1] || t[0]; return; }
    case "place-items": { const t = tokens(); out["align-items"] = t[0]; out["justify-items"] = t[1] || t[0]; return; }
    case "place-self": { const t = tokens(); out["align-self"] = t[0]; out["justify-self"] = t[1] || t[0]; return; }
    case "place-content": { const t = tokens(); out["align-content"] = t[0]; out["justify-content"] = t[1] || t[0]; return; }
    case "overflow": { const t = tokens(); out["overflow-x"] = t[0]; out["overflow-y"] = t[1] || t[0]; return; }
    case "grid-column": { const parts = value.split("/").map(part => part.trim()); out["grid-column-start"] = parts[0]; out["grid-column-end"] = parts[1] || "auto"; return; }
    case "list-style": out["list-style-type"] = /none/.test(value) ? "none" : tokens()[0]; return;
    case "word-wrap": out["overflow-wrap"] = value; return;
    case "container": { const t = value.split("/"); out["container-type"] = (t[1] || t[0]).trim(); return; }
    case "outline": { const p = borderParts(value); out["outline-width"] = p.width; out["outline-style"] = p.style; out["outline-color"] = p.color; return; }
    default: out[property] = value;
  }
}

// ---------------------------------------------------------------- colours

const NAMED_COLOURS = { white: "#ffffff", black: "#000000", red: "#ff0000", blue: "#0000ff", green: "#008000", gray: "#808080", grey: "#808080", silver: "#c0c0c0" };
const hex2 = n => { const v = Math.max(0, Math.min(255, Math.round(n))); return (v < 16 ? "0" : "") + v.toString(16); };
// "#rrggbb", "#rrggbbaa", "transparent", "currentcolor"; unknown text is kept.
function normaliseColour(value) {
  const text = value.trim().toLowerCase();
  if (text === "transparent" || text === "currentcolor" || text === "inherit" || text === "initial" || text === "unset") return text;
  if (NAMED_COLOURS[text]) return NAMED_COLOURS[text];
  let match = /^#([0-9a-f]{3,8})$/.exec(text);
  if (match) {
    let hex = match[1];
    if (hex.length === 3 || hex.length === 4) hex = hex.split("").map(char => char + char).join("");
    if (hex.length === 8 && hex.slice(6) === "ff") hex = hex.slice(0, 6);
    return "#" + hex;
  }
  match = /^rgba?\(([^)]*)\)$/.exec(text);
  if (match) {
    const parts = match[1].split(/[\s,/]+/).filter(Boolean);
    const channel = part => /%$/.test(part) ? parseFloat(part) * 2.55 : parseFloat(part);
    const alpha = parts[3] === undefined ? 1 : /%$/.test(parts[3]) ? parseFloat(parts[3]) / 100 : parseFloat(parts[3]);
    const rgb = "#" + hex2(channel(parts[0])) + hex2(channel(parts[1])) + hex2(channel(parts[2]));
    return alpha >= 1 ? rgb : rgb + hex2(alpha * 255);
  }
  return text;
}

// Box shadows keep their layers; each layer's colour is normalised ("inset" layers and
// unknown text pass through for the renderer to skip).
function normaliseShadow(value) {
  if (/^none$/i.test(value.trim())) return "none";
  return splitTop(value, ",").map(layer => splitTop(layer.trim(), " ").map(token =>
    /^(#|rgba?\(|[a-z]+$)/i.test(token) && !/^inset$/i.test(token) ? normaliseColour(token) : token).join(" ")).join(", ");
}

// ---------------------------------------------------------------- var()

function substituteVars(value, lookup, depth) {
  if (value.indexOf("var(") < 0 || depth > 12) return value;
  let result = "", index = 0;
  for (;;) {
    const start = value.indexOf("var(", index);
    if (start < 0) { result += value.slice(index); break; }
    result += value.slice(index, start);
    let count = 0, end = start + 3;
    for (; end < value.length; end++) {
      if (value[end] === "(") count++;
      else if (value[end] === ")") { count--; if (count === 0) break; }
    }
    const inner = value.slice(start + 4, end);
    const parts = splitTop(inner, ",");
    const name = parts[0].trim();
    const found = lookup(name);
    if (found === undefined) result += value.slice(start, end + 1); // resolved at run time
    else result += substituteVars(found, lookup, depth + 1);
    index = end + 1;
  }
  return result;
}

// ---------------------------------------------------------------- font coverage

function fontCoverage(html) {
  for (const match of html.matchAll(/@font-face\s*{([^}]*)}/g)) {
    const weight = /font-weight:\s*(\d+)/.exec(match[1]);
    const data = /base64,([A-Za-z0-9+/=]+)/.exec(match[1]);
    if (!data || (weight && weight[1] !== "400")) continue;
    const buffer = Buffer.from(data[1], "base64");
    const tables = {};
    for (let i = 0; i < buffer.readUInt16BE(4); i++) tables[buffer.toString("latin1", 12 + i * 16, 16 + i * 16)] = buffer.readUInt32BE(12 + i * 16 + 8);
    const cmap = tables.cmap;
    const covered = new Set();
    for (let i = 0; i < buffer.readUInt16BE(cmap + 2); i++) {
      const platform = buffer.readUInt16BE(cmap + 4 + i * 8), encoding = buffer.readUInt16BE(cmap + 6 + i * 8);
      const offset = cmap + buffer.readUInt32BE(cmap + 8 + i * 8);
      if (platform !== 3 || encoding !== 1 || buffer.readUInt16BE(offset) !== 4) continue;
      const segments = buffer.readUInt16BE(offset + 6) / 2;
      const ends = offset + 14, starts = ends + segments * 2 + 2, deltas = starts + segments * 2, ranges = deltas + segments * 2;
      for (let s = 0; s < segments; s++) {
        const end = buffer.readUInt16BE(ends + s * 2), start = buffer.readUInt16BE(starts + s * 2);
        const delta = buffer.readInt16BE(deltas + s * 2), rangeOffset = buffer.readUInt16BE(ranges + s * 2);
        for (let code = start; code <= end && code !== 0xffff; code++) {
          let glyph;
          if (!rangeOffset) glyph = (code + delta) & 0xffff;
          else { const raw = buffer.readUInt16BE(ranges + s * 2 + rangeOffset + (code - start) * 2); glyph = raw ? (raw + delta) & 0xffff : 0; }
          if (glyph) covered.add(code);
        }
      }
    }
    const codes = [...covered].sort((a, b) => a - b);
    const out = [];
    for (let i = 0; i < codes.length; i++) {
      let j = i;
      while (j + 1 < codes.length && codes[j + 1] === codes[j] + 1) j++;
      out.push(codes[i].toString(16) + "-" + codes[j].toString(16));
      i = j;
    }
    return out.join(",");
  }
  return "";
}

// ---------------------------------------------------------------- main

function escapeField(value) { return String(value).replace(/\\/g, "\\\\").replace(/\t/g, "\\t").replace(/\r?\n/g, "\\n"); }

function buildRuleTable(indexPath) {
  const html = fs.readFileSync(indexPath, "utf8");
  const styleMatch = /<style>([\s\S]*?)<\/style>/i.exec(html);
  if (!styleMatch) throw new Error("No <style> in " + indexPath);
  const cssText = styleMatch[1].replace(/@font-face\s*{[^}]*}/g, "");
  const conditions = [];
  const rules = parseStylesheet(UA_CSS, true, conditions).concat(parseStylesheet(cssText, false, conditions));

  // Custom properties on :root, light then dark.
  const light = {}, dark = {};
  const isRoot = rule => rule.tokens.length === 1 && rule.tokens[0] === "Proot";
  const isDarkRoot = rule => rule.tokens.length === 2 && rule.tokens[0] === "Proot" && rule.tokens[1] === "A=data-theme=dark";
  rules.filter(rule => !rule.ua && !rule.condition && !rule.pseudo).forEach(rule => {
    const target = isRoot(rule) ? [light, dark] : isDarkRoot(rule) ? [dark] : null;
    rule.block.declarations.forEach(d => {
      if (d.property.indexOf("--") !== 0) return;
      if (!target) throw new Error("Custom property " + d.property + " outside :root (" + rule.text + ") is not supported");
      target.forEach(map => { map[d.property] = d.value; });
    });
  });
  const resolve = (map, value) => substituteVars(value, name => map[name], 0);

  // Expand each block once per theme.
  const blocks = new Map();
  const unknown = new Set();
  rules.forEach(rule => {
    if (blocks.has(rule.block)) return;
    const lightOut = [], darkOut = [];
    rule.block.declarations.forEach(d => {
      if (d.property.indexOf("--") === 0) return;
      [[light, lightOut], [dark, darkOut]].forEach(([map, out]) => {
        const expanded = {};
        expand(d.property, resolve(map, d.value).trim(), expanded);
        Object.keys(expanded).forEach(name => {
          if (PROPERTY_INDEX[name] === undefined) { unknown.add(name); return; }
          let value = String(expanded[name]).trim();
          if (COLOUR_PROPERTIES.test(name)) value = normaliseColour(value);
          else if (name === "box-shadow") value = normaliseShadow(value);
          out.push({ property: name, value, important: d.important });
        });
      });
    });
    const declarations = lightOut.map((d, i) => ({ property: d.property, important: d.important, light: d.value, dark: darkOut[i].value }));
    blocks.set(rule.block, { id: 0, declarations });
  });

  // Cascade order: browser defaults first, then specificity, then source order.
  const kept = rules.filter(rule => blocks.get(rule.block).declarations.length);
  kept.forEach((rule, index) => { rule.seq = index; });
  kept.sort((a, b) => (a.ua === b.ua ? 0 : a.ua ? -1 : 1) || a.spec - b.spec || a.seq - b.seq);

  const lines = ["V\t" + FORMAT_VERSION];
  PROPERTIES.forEach(name => lines.push("P\t" + name));
  conditions.forEach((condition, index) => lines.push(["C", index + 1, condition.parent || 0, condition.type, condition.alternatives].join("\t")));
  let nextBlock = 0;
  kept.forEach(rule => {
    const block = blocks.get(rule.block);
    if (!block.id) {
      block.id = ++nextBlock;
      lines.push("B\t" + block.id);
      block.declarations.forEach(d => {
        const fields = ["D", PROPERTY_INDEX[d.property], d.important ? 1 : 0, escapeField(d.light)];
        if (d.dark !== d.light) fields.push(escapeField(d.dark));
        lines.push(fields.join("\t"));
      });
    }
    lines.push(["R", block.id, rule.condition || 0, rule.pseudo, rule.spec, rule.key, rule.simple, escapeField(rule.base)].join("\t"));
  });
  lines.push("G\t" + fontCoverage(html));
  lines.push("E");
  return { text: lines.join("\n") + "\n", ruleCount: kept.length, blockCount: nextBlock, unknown: Array.from(unknown).sort(), cssText };
}

if (require.main === module) {
  const input = path.resolve(process.argv[2] || path.join(ROOT, "..", "index.html"));
  const output = path.resolve(process.argv[3] || path.join(ROOT, ".build", "css-rules.txt"));
  const result = buildRuleTable(input);
  fs.mkdirSync(path.dirname(output), { recursive: true });
  fs.writeFileSync(output, result.text);
  console.log(`Wrote ${path.relative(process.cwd(), output)}: ${result.ruleCount} rules, ${result.blockCount} blocks, ${result.text.length} chars`);
  if (result.unknown.length) console.log("Dropped properties: " + result.unknown.join(", "));
}

module.exports = { buildRuleTable, PROPERTIES, UA_CSS, compileSelector, normaliseColour };
