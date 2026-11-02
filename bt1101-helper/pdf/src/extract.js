// Pulls the page markup and application script out of index.html so the PDF always
// runs the website's own logic rather than a copy of it.

const fs = require("fs");
const path = require("path");

const GLOBAL_NAMES = [
  "window", "self", "document", "navigator", "location", "localStorage", "sessionStorage", "matchMedia",
  "requestAnimationFrame", "cancelAnimationFrame", "setTimeout", "clearTimeout", "setInterval", "clearInterval",
  "queueMicrotask", "CSS", "MutationObserver", "ResizeObserver", "IntersectionObserver",
  "Event", "CustomEvent", "KeyboardEvent", "MouseEvent", "InputEvent", "FocusEvent", "UIEvent",
  "Node", "Element", "HTMLElement", "HTMLButtonElement", "HTMLInputElement", "HTMLSelectElement", "HTMLTextAreaElement",
  "HTMLOptionElement", "Text", "DocumentFragment", "Option", "Blob", "URL", "fetch", "getComputedStyle",
  "alert", "confirm", "prompt", "console", "history", "performance", "crypto", "scrollTo", "scrollBy",
  "innerWidth", "innerHeight"
];

// Functions the PDF bridge calls directly; everything else is driven through DOM events.
const EXPORTS = ["setActive", "state", "groups", "samples", "universalSamples", "generate", "appVersion"];

function extractWebsite(indexPath) {
  const html = fs.readFileSync(indexPath, "utf8");
  const firstLine = html.split(/\r?\n/, 1)[0];
  const versionMatch = firstLine.match(/BT1101 Prose Lab version:\s*(v?\d{6}_\d{4}h?)/i);
  const bodyStart = html.indexOf("<body");
  const scripts = [];
  const scriptPattern = /<script\b([^>]*)>([\s\S]*?)<\/script>/gi;
  let match;
  while ((match = scriptPattern.exec(html))) {
    if (/\bsrc\s*=/.test(match[1])) throw new Error("External scripts are not supported: " + match[0].slice(0, 80));
    scripts.push({ index: match.index, code: match[2] });
  }
  const bodyScriptStart = scripts.filter(script => script.index > bodyStart).map(script => script.index)[0];
  const bodyEnd = html.lastIndexOf("</body>");
  let body = html.slice(bodyStart, bodyScriptStart !== undefined ? bodyScriptStart : bodyEnd);
  body = body.replace(/<script\b[\s\S]*?<\/script>/gi, "");
  const title = (html.match(/<title>([\s\S]*?)<\/title>/i) || [])[1];
  return { html, body, scripts: scripts.map(script => script.code), websiteVersion: versionMatch ? versionMatch[1] : "", title };
}

function writeAppModule(site, outPath) {
  const code = site.scripts.join("\n;\n");
  const exports = EXPORTS.map(name => `${name}: typeof ${name} === "undefined" ? undefined : ${name}`).join(", ");
  const module = `// Generated from index.html by pdf/build.js. Do not edit.
export function runApp(g) {
  const { ${GLOBAL_NAMES.join(", ")} } = g;
  return (function () {
${code}
;
    return { ${exports}, getAppState: function () { return { calculatePlotAxisIntervals: typeof calculatePlotAxisIntervals === "undefined" ? false : calculatePlotAxisIntervals }; } };
  })();
}
`;
  fs.mkdirSync(path.dirname(outPath), { recursive: true });
  fs.writeFileSync(outPath, module);
}

module.exports = { extractWebsite, writeAppModule, GLOBAL_NAMES };
