// PDF 1.7 writer for the captured pages: draws the website's boxes and words in
// Source Sans Pro, then lays AcroForm fields over the controls, with document-level
// JavaScript. Captured coordinates are CSS pixels from the top left; 1px = 0.75pt.

const zlib = require("zlib");
const { toWinAnsi, measure, DIFFERENCES } = require("./fonts.js");

const PX = 0.75;
const FLAGS = { readOnly: 1, multiline: 1 << 12, doNotSpellCheck: 1 << 22, doNotScroll: 1 << 23, pushButton: 1 << 16, combo: 1 << 17, commitOnSelChange: 1 << 26 };
const ANNOT_PRINT = 4, ANNOT_HIDDEN = 2;
// Field text uses the website's font (see embedFont), at the website's size.
const FIELD_SCALE = 1;

const num = value => (Math.round(value * 100) / 100).toString();
// Theme colours captured from the website (see capture-page.js palette); light defaults.
let PALETTE = {};
const LIGHT = { accent: [0.043, 0.341, 0.816], track: [0.855, 0.863, 0.878], trackBorder: [0.502, 0.525, 0.545], muted: [0.373, 0.388, 0.408], clipFill: [1, 0.96, 0.86], clipBorder: [0.78, 0.66, 0.3], panel: [1, 1, 1] };
const paint = name => PALETTE[name] || LIGHT[name];
function literal(bytes) { return "(" + bytes.replace(/\\/g, "\\\\").replace(/\(/g, "\\(").replace(/\)/g, "\\)").replace(/\r/g, "\\r").replace(/\n/g, "\\n") + ")"; }
function textString(text) {
  text = String(text);
  if (/^[\x20-\x7e\n\r\t]*$/.test(text)) return literal(text);
  let hex = "FEFF";
  for (let i = 0; i < text.length; i++) hex += text.charCodeAt(i).toString(16).padStart(4, "0").toUpperCase();
  return "<" + hex + ">";
}
// Colours with alpha are flattened onto the page background, which is what they sit on.
function flatten(colour, under) {
  if (!colour) return null;
  const a = colour.a === undefined ? 1 : colour.a, base = under || { r: 1, g: 1, b: 1 };
  return [colour.r * a + base.r * (1 - a), colour.g * a + base.g * (1 - a), colour.b * a + base.b * (1 - a)];
}
const rgb = values => values.map(num).join(" ");

class PdfDocument {
  constructor() { this.objects = []; }
  reserve() { this.objects.push(null); return this.objects.length; }
  set(id, body) { this.objects[id - 1] = body; return id; }
  add(body) { const id = this.reserve(); return this.set(id, body); }
  stream(dictionary, data, compress) {
    let buffer = Buffer.isBuffer(data) ? data : Buffer.from(data, "latin1");
    let extra = "";
    if (compress) { buffer = zlib.deflateSync(buffer, { level: 9 }); extra = " /Filter /FlateDecode"; }
    return this.add({ stream: buffer, dictionary: `<< ${dictionary}${extra} /Length ${buffer.length} >>` });
  }
  serialize(rootId, infoId, headerComment) {
    const chunks = [];
    let offset = 0;
    const push = data => { const buffer = Buffer.isBuffer(data) ? data : Buffer.from(data, "latin1"); chunks.push(buffer); offset += buffer.length; };
    push(`%PDF-1.7\n${headerComment ? "% " + headerComment + "\n" : ""}%\xe2\xe3\xcf\xd3\n`);
    const offsets = [];
    this.objects.forEach((body, index) => {
      offsets.push(offset);
      push(`${index + 1} 0 obj\n`);
      if (body && body.stream) { push(body.dictionary + "\nstream\n"); push(body.stream); push("\nendstream"); }
      else push(body);
      push("\nendobj\n");
    });
    const xref = offset;
    push(`xref\n0 ${this.objects.length + 1}\n0000000000 65535 f \n`);
    offsets.forEach(value => push(String(value).padStart(10, "0") + " 00000 n \n"));
    push(`trailer\n<< /Size ${this.objects.length + 1} /Root ${rootId} 0 R /Info ${infoId} 0 R >>\nstartxref\n${xref}\n%%EOF\n`);
    return Buffer.concat(chunks);
  }
}

function jsAction(document, code) {
  if (code.length < 400) return `<< /S /JavaScript /JS ${literal(code)} >>`;
  return `<< /S /JavaScript /JS ${document.stream("", code, true)} 0 R >>`;
}

// Each font file is embedded once and given two encodings. Text drawn by this writer uses
// WinAnsi plus /Differences (so H0 subscripts, mu and <= draw from the font). Form fields use
// plain WinAnsi: that is the setup Acrobat can type and redraw with, and it then supplies
// the characters WinAnsi lacks itself (with /Differences, Acrobat shows dots).
function embedFont(document, font, name, plain) {
  const file = document.stream(`/Length1 ${font.buffer.length}`, font.buffer, true);
  const flags = 32; // nonsymbolic
  const descriptor = document.add(`<< /Type /FontDescriptor /FontName /${name} /Flags ${flags} /FontBBox [${font.bbox.join(" ")}] /ItalicAngle 0 /Ascent ${font.ascent} /Descent ${font.descent} /CapHeight ${font.capHeight} /StemV ${font.weight >= 600 ? 120 : 80} /FontFile2 ${file} 0 R >>`);
  const drawn = document.add(`<< /Type /Font /Subtype /TrueType /BaseFont /${name} /FirstChar 32 /LastChar 255 /Widths [${font.widths.join(" ")}] /Encoding << /Type /Encoding /BaseEncoding /WinAnsiEncoding /Differences [${DIFFERENCES.map(entry => entry.code + " /" + entry.name).join(" ")}] >> /FontDescriptor ${descriptor} 0 R >>`);
  if (!plain) return drawn;
  const fieldWidths = font.widths.map((width, index) => (EXTRA_CODES.indexOf(index + 32) >= 0 ? font.plainWidths[index] : width));
  const field = document.add(`<< /Type /Font /Subtype /TrueType /BaseFont /${name} /FirstChar 32 /LastChar 255 /Widths [${fieldWidths.join(" ")}] /Encoding /WinAnsiEncoding /FontDescriptor ${descriptor} 0 R >>`);
  return { drawn, field };
}
const EXTRA_CODES = DIFFERENCES.map(entry => entry.code);

function fontKey(weight) { return weight >= 650 ? "SSPB" : weight >= 550 ? "SSPS" : "SSP"; }

function roundedRect(x, y, w, h, r) {
  if (!r) return `${num(x)} ${num(y)} ${num(w)} ${num(h)} re`;
  const k = 0.5523 * r;
  return [
    `${num(x + r)} ${num(y)} m`, `${num(x + w - r)} ${num(y)} l`, `${num(x + w - r + k)} ${num(y)} ${num(x + w)} ${num(y + r - k)} ${num(x + w)} ${num(y + r)} c`,
    `${num(x + w)} ${num(y + h - r)} l`, `${num(x + w)} ${num(y + h - r + k)} ${num(x + w - r + k)} ${num(y + h)} ${num(x + w - r)} ${num(y + h)} c`,
    `${num(x + r)} ${num(y + h)} l`, `${num(x + r - k)} ${num(y + h)} ${num(x)} ${num(y + h - r + k)} ${num(x)} ${num(y + h - r)} c`,
    `${num(x)} ${num(y + r)} l`, `${num(x)} ${num(y + r - k)} ${num(x + r - k)} ${num(y)} ${num(x + r)} ${num(y)} c`, "h"
  ].join(" ");
}

function buttonAppearance(document, fonts, field, width, height, background) {
  const text = field.text || {};
  const size = (text.size || 11) * PX;
  const key = fontKey(text.weight || 700);
  const caption = field.caption || "";
  const captionWidth = measure(fonts[key], caption, size);
  // A button without a background of its own (icon buttons) stays transparent.
  const fill = flatten(field.fill, background);
  const border = flatten(field.border, background);
  const radius = Math.min((field.radius === undefined ? 3 : field.radius) * PX, height / 2);
  let content = fill ? `q ${rgb(fill)} rg ${roundedRect(0.4, 0.4, width - 0.8, height - 0.8, radius)} f Q` : "";
  if (border && field.borderWidth) content += ` q ${rgb(border)} RG ${num(field.borderWidth * PX)} w ${roundedRect(0.4, 0.4, width - 0.8, height - 0.8, radius)} S Q`;
  const colour = flatten(text.colour, background) || [0, 0, 0];
  if (caption === "↓" || caption === "↑") {
    // Arrows are outside WinAnsi, so the steppers' arrows are drawn as shapes.
    const cx = width / 2, cy = height / 2, s = Math.min(width, height) * 0.22, down = caption === "↓" ? -1 : 1;
    content += ` q ${rgb(colour)} RG ${rgb(colour)} rg 1.3 w 1 J 1 j ${num(cx)} ${num(cy - down * s)} m ${num(cx)} ${num(cy + down * s)} l S ${num(cx - s * 0.7)} ${num(cy + down * s * 0.25)} m ${num(cx)} ${num(cy + down * s)} l ${num(cx + s * 0.7)} ${num(cy + down * s * 0.25)} l S Q`;
  } else if (caption === "↺" || caption === "⧉") {
    // Icon buttons: a circular arrow (reset) and two overlapping sheets (copy).
    const cx = width / 2, cy = height / 2, r = Math.min(width, height) * 0.22;
    if (caption === "↺") {
      const k = 0.5523 * r;
      content += ` q ${rgb(colour)} RG ${rgb(colour)} rg 1.2 w 1 J ${num(cx)} ${num(cy + r)} m ${num(cx + k)} ${num(cy + r)} ${num(cx + r)} ${num(cy + k)} ${num(cx + r)} ${num(cy)} c ${num(cx + r)} ${num(cy - k)} ${num(cx + k)} ${num(cy - r)} ${num(cx)} ${num(cy - r)} c ${num(cx - k)} ${num(cy - r)} ${num(cx - r)} ${num(cy - k)} ${num(cx - r)} ${num(cy)} c S ${num(cx - r - r * 0.45)} ${num(cy + r * 0.35)} m ${num(cx - r)} ${num(cy - r * 0.15)} l ${num(cx - r + r * 0.5)} ${num(cy + r * 0.3)} l f Q`;
    } else {
      content += ` q ${rgb(colour)} RG 1.1 w ${num(cx - r * 1.1)} ${num(cy - r * 1.1)} ${num(r * 1.5)} ${num(r * 1.5)} re S ${num(cx - r * 0.4)} ${num(cy - r * 0.4)} ${num(r * 1.5)} ${num(r * 1.5)} re S Q`;
    }
  } else {
    content += ` BT /${key} ${num(size)} Tf ${rgb(colour)} rg ${num(Math.max(2, (width - captionWidth) / 2))} ${num((height - size * 0.7) / 2)} Td ${literal(toWinAnsi(caption))} Tj ET`;
  }
  return document.stream(`/Type /XObject /Subtype /Form /BBox [0 0 ${num(width)} ${num(height)}] /Resources << /Font << /${key} ${document.fontIds[key]} 0 R >> >>`, content, false);
}

// The website's switch: a grey pill with a white knob, blue with the knob right when on.
function checkAppearances(document, size, accent) {
  const box = `q ${rgb(paint("panel"))} rg ${rgb(paint("trackBorder"))} RG 0.75 w ${roundedRect(0.5, 0.5, size - 1, size - 1, 2)} B Q`;
  const on = document.stream(`/Type /XObject /Subtype /Form /BBox [0 0 ${num(size)} ${num(size)}]`, `q ${rgb(accent)} rg ${roundedRect(0.5, 0.5, size - 1, size - 1, 2)} f Q q 1 1 1 RG ${num(size * 0.12)} w 1 J 1 j ${num(size * 0.25)} ${num(size * 0.52)} m ${num(size * 0.43)} ${num(size * 0.32)} l ${num(size * 0.76)} ${num(size * 0.7)} l S Q`, false);
  const off = document.stream(`/Type /XObject /Subtype /Form /BBox [0 0 ${num(size)} ${num(size)}]`, box, false);
  return { on, off };
}

function switchAppearances(document, width, height) {
  const knob = height - 4 * PX * 2;
  const pill = (fill, stroke) => `q ${rgb(fill)} rg ${rgb(stroke)} RG 0.75 w ${roundedRect(0.4, 0.4, width - 0.8, height - 0.8, (height - 0.8) / 2)} B Q`;
  const circle = x => {
    const r = knob / 2, cx = x + r, cy = height / 2, k = 0.5523 * r;
    return `q 1 1 1 rg ${num(cx + r)} ${num(cy)} m ${num(cx + r)} ${num(cy + k)} ${num(cx + k)} ${num(cy + r)} ${num(cx)} ${num(cy + r)} c ${num(cx - k)} ${num(cy + r)} ${num(cx - r)} ${num(cy + k)} ${num(cx - r)} ${num(cy)} c ${num(cx - r)} ${num(cy - k)} ${num(cx - k)} ${num(cy - r)} ${num(cx)} ${num(cy - r)} c ${num(cx + k)} ${num(cy - r)} ${num(cx + r)} ${num(cy - k)} ${num(cx + r)} ${num(cy)} c f Q`;
  };
  const bbox = `/Type /XObject /Subtype /Form /BBox [0 0 ${num(width)} ${num(height)}]`;
  return {
    on: document.stream(bbox, pill(paint("accent"), paint("accent")) + " " + circle(width - knob - 3 * PX), false),
    off: document.stream(bbox, pill(paint("track"), paint("trackBorder")) + " " + circle(3 * PX), false)
  };
}

// What a text or choice field shows until Acrobat redraws it: its box and any starting
// text, in Source Sans Pro. Acrobat redraws a field only when a script changes it, and
// then uses the field's /DA font, which is Helvetica (Acrobat cannot type with the
// embedded Source Sans Pro and shows dots instead).
function fieldAppearance(document, fonts, options) {
  const { width, height } = options;
  let content = "";
  const box = roundedRect(0.4, 0.4, width - 0.8, height - 0.8, options.radius || 0);
  if (options.fill) content += `q ${rgb(options.fill)} rg ${box} f Q `;
  if (options.border) content += `q ${rgb(options.border)} RG ${num(options.borderWidth || 0.75)} w${options.dashed ? " [2 2] 0 d" : ""} ${box} S Q `;
  const placeholder = !options.value && options.placeholder && options.placeholder.text;
  const value = options.value ? String(options.value) : placeholder ? options.placeholder.text : "";
  if (value) {
    const font = fonts[options.font], size = options.size;
    if (placeholder) options = Object.assign({}, options, { colour: options.placeholder.colour ? [options.placeholder.colour.r, options.placeholder.colour.g, options.placeholder.colour.b] : [0.5, 0.53, 0.55] });
    const lines = [];
    value.split("\n").forEach(paragraph => {
      if (!options.multiline) { lines.push(paragraph); return; }
      let line = "";
      paragraph.split(/\s+/).forEach(word => {
        const candidate = line ? line + " " + word : word;
        if (line && measure(font, candidate, size) > width - 4) { lines.push(line); line = word; } else line = candidate;
      });
      lines.push(line);
    });
    content += `/Tx BMC q 1 1 ${num(width - 2)} ${num(height - 2)} re W n BT /${options.font} ${num(size)} Tf ${rgb(options.colour)} rg `;
    const leading = size * 1.15;
    lines.forEach((line, index) => {
      const lineWidth = measure(font, line, size);
      const x = options.align === 2 ? width - 2 - lineWidth : options.align === 1 ? (width - lineWidth) / 2 : 2;
      const y = options.multiline || lines.length > 1 ? height - 2 - size * 0.95 - index * leading : (height - size * 0.7) / 2;
      content += `1 0 0 1 ${num(x)} ${num(y)} Tm ${literal(toWinAnsi(line))} Tj `;
    });
    content += "ET Q EMC";
  }
  const resources = Object.keys(document.fontIds).map(key => `/${key} ${document.fontIds[key]} 0 R`).join(" ");
  return document.stream(`/Type /XObject /Subtype /Form /BBox [0 0 ${num(width)} ${num(height)}] /Resources << /Font << ${resources} >> >>`, content, true);
}

function drawPage(page, fonts) {
  const H = page.height * PX;
  const background = page.background || { r: 1, g: 1, b: 1 };
  let content = `q ${rgb(flatten(background))} rg 0 0 ${num(page.width * PX)} ${num(H)} re f Q\n`;
  page.boxes.forEach(item => {
    const r = item.rect, x = r.x * PX, y = H - (r.y + r.h) * PX, w = r.w * PX, h = r.h * PX, radius = (item.radius || 0) * PX;
    const fill = flatten(item.fill, background);
    if (fill) content += `q ${rgb(fill)} rg ${roundedRect(x, y, w, h, radius)} f Q\n`;
    const sides = item.sides || [];
    const uniform = sides.every(side => side && side.colour && sides[0] && side.width === sides[0].width && JSON.stringify(side.colour) === JSON.stringify(sides[0].colour));
    if (uniform && sides[0]) {
      const width = sides[0].width * PX;
      content += `q ${rgb(flatten(sides[0].colour, background))} RG ${num(width)} w ${roundedRect(x + width / 2, y + width / 2, w - width, h - width, Math.max(0, radius - width / 2))} S Q\n`;
    } else {
      sides.forEach((side, index) => {
        if (!side || !side.colour) return;
        const width = side.width * PX, colour = rgb(flatten(side.colour, background));
        const strip = [[x, y + h - width, w, width], [x + w - width, y, width, h], [x, y, w, width], [x, y, width, h]][index];
        content += `q ${colour} rg ${strip.map(num).join(" ")} re f Q\n`;
      });
    }
  });
  (page.icons || []).forEach(icon => {
    if (!icon.image) return;
    content += `q ${num(icon.w * PX)} 0 0 ${num(icon.h * PX)} ${num(icon.x * PX)} ${num(H - (icon.y + icon.h) * PX)} cm /I${icon.image.key} Do Q\n`;
  });
  // All words go into one text object. Words on the same line in the same style become one
  // TJ run whose spacing numbers put each word exactly where the browser measured it, and
  // font, colour and spacing are only set when they change. Fewer, longer text runs keep
  // Acrobat's scrolling smooth.
  const runs = [];
  page.words.forEach(word => {
    if (word.owner) return;
    const style = word.style, key = fontKey(style.weight), font = fonts[key];
    const size = style.size * PX;
    const baseline = H - (word.y + word.h * (font.ascent / (font.ascent - font.descent))) * PX;
    const colour = rgb(flatten(style.colour, background) || [0, 0, 0]);
    const spacing = (style.spacing || 0) * PX;
    const last = runs[runs.length - 1];
    const x = word.x * PX;
    if (last && last.key === key && last.size === size && last.colour === colour && last.spacing === spacing && Math.abs(last.baseline - baseline) < 0.3 && x >= last.end - 0.5) {
      // TJ numbers are thousandths of the font size, subtracted from the position.
      last.parts.push(Math.round(-(x - last.end) / size * 1000));
      last.parts.push(word.text);
    } else {
      runs.push({ key, size, colour, spacing, baseline, x, parts: [word.text], end: 0 });
    }
    const run = runs[runs.length - 1];
    run.end = x + measure(font, word.text, size) + spacing * word.text.length;
  });
  let state = {};
  content += "BT\n";
  runs.forEach(run => {
    if (state.key !== run.key || state.size !== run.size) content += `/${run.key} ${num(run.size)} Tf `;
    if (state.colour !== run.colour) content += `${run.colour} rg `;
    if (state.spacing !== run.spacing) content += `${num(run.spacing)} Tc `;
    state = run;
    const parts = run.parts.map(part => (typeof part === "number" ? String(part) : literal(toWinAnsi(part)))).join(" ");
    content += `1 0 0 1 ${num(run.x)} ${num(run.baseline)} Tm [${parts}] TJ\n`;
  });
  content += "ET\n";
  return content;
}

// The words and icons inside a rectangle, positioned relative to it, in a given colour.
function regionText(page, fonts, rect, colour) {
  const inside = item => item.x >= rect.x - 1 && item.x < rect.x + rect.w && item.y >= rect.y - 1 && item.y < rect.y + rect.h;
  let content = "";
  (page.icons || []).filter(inside).forEach(icon => {
    if (icon.image) content += `q ${num(icon.w * PX)} 0 0 ${num(icon.h * PX)} ${num((icon.x - rect.x) * PX)} ${num((rect.y + rect.h - icon.y - icon.h) * PX)} cm /I${icon.image.key} Do Q `;
  });
  const words = page.words.filter(word => !word.owner && inside(word));
  if (!words.length) return content;
  content += "BT ";
  words.forEach(word => {
    const key = fontKey(word.style.weight), font = fonts[key], size = word.style.size * PX;
    const baseline = (rect.y + rect.h - (word.y + word.h * (font.ascent / (font.ascent - font.descent)))) * PX;
    const ink = colour || word.style.colour;
    content += `/${key} ${num(size)} Tf ${num((word.style.spacing || 0) * PX)} Tc ${ink ? rgb([ink.r, ink.g, ink.b]) : "0 0 0"} rg 1 0 0 1 ${num((word.x - rect.x) * PX)} ${num(baseline)} Tm ${literal(toWinAnsi(word.text))} Tj `;
  });
  return content + "ET";
}

// Words inside elements that come and go become read-only fields, one per line, so
// they can be hidden together with their element.
function labelFields(page, fonts) {
  const width = word => measure(fonts[fontKey(word.style.weight)], word.text, word.style.size) + (word.style.spacing || 0) * word.text.length;
  const lines = [];
  page.words.forEach(word => {
    if (!word.owner) return;
    const line = lines.find(entry => entry.owner === word.owner && Math.abs(entry.y - word.y) < 2 && word.x >= entry.x && word.x - (entry.x + entry.w) < 14);
    if (line) { line.text += " " + word.text; line.w = Math.max(line.w, word.x + width(word) - line.x); }
    else lines.push({ owner: word.owner, x: word.x, y: word.y, h: word.h, w: width(word), text: word.text, style: word.style });
  });
  return lines.map(line => ({ k: "label", r: line.owner, value: line.text, rect: { x: line.x - 1, y: line.y, w: line.w + 6, h: line.h }, text: line.style }));
}

function buildPdf(spec) {
  const fonts = spec.fonts;
  PALETTE = spec.palette || {};
  const document = new PdfDocument();
  const catalogId = document.reserve(), pagesId = document.reserve(), acroFormId = document.reserve();
  document.fontIds = {
  };
  [["SSP", "SourceSansPro-Regular"], ["SSPS", "SourceSansPro-Semibold"], ["SSPB", "SourceSansPro-Bold"]].forEach(([key, name]) => {
    const ids = embedFont(document, fonts[key], name, true);
    document.fontIds[key] = ids.drawn;
    document.fontIds[key + "F"] = ids.field;
  });
  Object.assign(document.fontIds, {
    Helv: document.add("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>"),
    HeBo: document.add("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>"),
    ZaDb: document.add("<< /Type /Font /Subtype /Type1 /BaseFont /ZapfDingbats >>")
  });
  const fontResources = `<< /Font << ${Object.keys(document.fontIds).map(key => `/${key} ${document.fontIds[key]} 0 R`).join(" ")} >> >>`;
  const pageIds = spec.pages.map(() => document.reserve());
  const fieldIds = [];
  // Each distinct icon is stored once and reused (the sidebar's icon is on every page).
  const imageIds = {};
  spec.pages.forEach(page => (page.icons || []).forEach(icon => {
    const image = icon.image;
    if (!image || imageIds[image.key]) return;
    imageIds[image.key] = document.stream(`/Type /XObject /Subtype /Image /Width ${image.width} /Height ${image.height} /ColorSpace /DeviceRGB /BitsPerComponent 8`, image.rgb, true);
  }));
  const alphaStates = {};
  const alphaState = alpha => { const key = "A" + Math.round(alpha * 100); if (!alphaStates[key]) alphaStates[key] = document.add(`<< /Type /ExtGState /ca ${num(alpha)} /CA ${num(alpha)} >>`); return key; };
  const resourcesFor = keys => `<< /Font << ${Object.keys(document.fontIds).map(key => `/${key} ${document.fontIds[key]} 0 R`).join(" ")} >> /XObject << ${keys.map(key => `/I${key} ${imageIds[key]} 0 R`).join(" ")} >> /ExtGState << ${Object.keys(alphaStates).map(key => `/${key} ${alphaStates[key]} 0 R`).join(" ")} >> >>`;

  spec.pages.forEach((page, pageIndex) => {
    const H = page.height * PX;
    const background = page.background || { r: 1, g: 1, b: 1 };
    const annots = [];
    const pageIcons = Array.from(new Set((page.icons || []).filter(icon => icon.image).map(icon => icon.image.key)));
    page.links.forEach((link, index) => {
      const target = spec.pageOfOperation[link.op];
      if (target === undefined) return;
      const r = link.rect, width = r.w * PX, height = r.h * PX;
      const bbox = `/Type /XObject /Subtype /Form /BBox [0 0 ${num(width)} ${num(height)}]`;
      const normal = document.stream(bbox, "", false);
      const hover = link.hoverStyle || {};
      let rollover = "";
      if (hover.fill) {
        const shape = roundedRect(0, 0, width, height, Math.min((link.radius || 0) * PX, height / 2));
        rollover += `q /${alphaState(hover.fill.a === undefined ? 1 : hover.fill.a)} gs ${rgb([hover.fill.r, hover.fill.g, hover.fill.b])} rg ${shape} f Q `;
        if (hover.fill.a === undefined || hover.fill.a > 0.5) rollover += regionText(page, fonts, r, hover.text);
      }
      const over = document.stream(`${bbox} /Resources ${resourcesFor(pageIcons)}`, rollover, true);
      const id = document.add(`<< /Type /Annot /Subtype /Widget /F 4 /P ${pageIds[pageIndex]} 0 R /Rect [${num(r.x * PX)} ${num(H - (r.y + r.h) * PX)} ${num((r.x + r.w) * PX)} ${num(H - r.y * PX)}] /T (p${pageIndex}.nav${index + 1}) /FT /Btn /Ff ${FLAGS.pushButton} /H /N /MK << >> /BS << /W 0 >> /AP << /N ${normal} 0 R /R ${over} 0 R >> /A << /S /GoTo /D [${pageIds[target]} 0 R /XYZ 0 ${num(spec.pages[target].height * PX)} null] >> >>`);
      fieldIds.push(id);
      annots.push(id);
    });
    page.fields.forEach(field => {
      // Text fields pad their content by about 2pt a side; widen read-only text to match the browser.
      const pad = field.k === "note" || field.k === "label" ? 8 : 0;
      let r = { x: field.rect.x - pad / 2, y: field.rect.y, w: field.rect.w + pad, h: field.rect.h };
      // Over a drawn box, the field covers only the box's content area, so text starts where
      // the website's padding puts it. Dropdown values stop short of the chevron.
      if (field.inset && (field.backdrop || field.k === "cardbody")) {
        const [top, right, bottom, left] = field.inset.map(value => Math.max(0, value - 3));
        const chevron = field.k === "select" || field.k === "plot" ? 18 : 0;
        const multi = field.multiline || field.k === "out";
        r = { x: r.x + left, y: r.y + (multi ? top : 0), w: Math.max(12, r.w - left - right - chevron), h: multi ? Math.max(14, r.h - top - bottom) : r.h };
      }
      const rect = `[${num(r.x * PX)} ${num(H - (r.y + r.h) * PX)} ${num((r.x + r.w) * PX)} ${num(H - r.y * PX)}]`;
      const text = field.text || {};
      const size = Math.max(6, (text.size || 13) * PX);
      const fieldSize = num(Math.max(6, size * FIELD_SCALE));
      const textColour = flatten(text.colour, background) || [0.12, 0.12, 0.12];
      const flagsF = field.k === "badge" ? ANNOT_PRINT | ANNOT_HIDDEN : ANNOT_PRINT;
      const common = `/Type /Annot /Subtype /Widget /F ${flagsF} /P ${pageIds[pageIndex]} 0 R /Rect ${rect} /T ${literal(field.name)}${field.label ? " /TU " + textString(field.label) : ""}`;
      const font = fontKey(text.weight || 400);
      const daFont = fontKey(text.weight || 400) + "F";
      // A field drawn over its own captured box needs no widget background or border.
      const fill = field.backdrop ? null : flatten(field.fill, background), border = field.backdrop || !field.borderWidth ? null : flatten(field.border, background);
      const mk = `/MK << ${fill ? `/BG [${rgb(fill)}] ` : ""}${border ? `/BC [${rgb(border)}] ` : ""}>> /BS << /W ${border ? num(Math.max(0.75, field.borderWidth * PX)) : 0} /S /S >>`;
      let body;
      if (field.k === "button") {
        const appearance = buttonAppearance(document, fonts, field, r.w * PX, r.h * PX, background);
        const hover = field.hoverStyle;
        const rollover = hover ? buttonAppearance(document, fonts, Object.assign({}, field, { fill: hover.fill || field.fill, border: hover.border || field.border, borderWidth: field.borderWidth || (hover.border ? 1 : 0), text: Object.assign({}, field.text, { colour: hover.text || (field.text || {}).colour }) }), r.w * PX, r.h * PX, background) : null;
        body = `<< ${common} /FT /Btn /Ff ${FLAGS.pushButton} /H /P /DA (/${fontKey(text.weight || 700)}F ${fieldSize} Tf ${rgb(textColour)} rg) /MK << /CA ${textString(field.caption || "")} ${fill ? `/BG [${rgb(fill)}] ` : ""}${border ? `/BC [${rgb(border)}]` : ""} >> /BS << /W ${border ? 0.75 : 0} /S /S >> /AP << /N ${appearance} 0 R${rollover ? ` /R ${rollover} 0 R` : ""} >> /A ${jsAction(document, "BT_press(event);")} >>`;
      } else if (field.k === "check") {
        if (!field.switch) {
          const side = Math.min(r.w, r.h) * PX, x0 = r.x * PX, y0 = H - (r.y + r.h) * PX + (r.h * PX - side) / 2;
          const box = checkAppearances(document, side, paint("accent"));
          body = `<< /Type /Annot /Subtype /Widget /F 4 /P ${pageIds[pageIndex]} 0 R /Rect [${num(x0)} ${num(y0)} ${num(x0 + side)} ${num(y0 + side)}] /T ${literal(field.name)} /FT /Btn /V /Off /AS /Off /DA (/ZaDb 0 Tf 0 g) /MK << /CA (4) /BG [1 1 1] /BC [0.5 0.53 0.55] >> /BS << /W 1 /S /S >> /AP << /N << /Yes ${box.on} 0 R /Off ${box.off} 0 R >> >> /A ${jsAction(document, "BT_toggle(event);")} >>`;
          const id = document.add(body); fieldIds.push(id); annots.push(id); return;
        }
        const appearance = switchAppearances(document, r.w * PX, r.h * PX);
        body = `<< ${common} /FT /Btn /V /Off /AS /Off /DA (/ZaDb 0 Tf ${rgb(paint("accent"))} rg) /MK << /CA (4) /BG [${rgb(paint("panel"))}] /BC [${rgb(paint("trackBorder"))}] >> /BS << /W 0 >> /AP << /N << /Yes ${appearance.on} 0 R /Off ${appearance.off} 0 R >> >> /A ${jsAction(document, "BT_toggle(event);")} >>`;
      } else if (field.k === "pick") {
        // Invisible button over the value box: a chevron, and a click opens the pop-up menu.
        const width = r.w * PX, height = r.h * PX, cx = width - 11, cy = height / 2;
        const chevron = document.stream(`/Type /XObject /Subtype /Form /BBox [0 0 ${num(width)} ${num(height)}]`, `q ${rgb(paint("muted"))} RG 1.1 w 1 J 1 j ${num(cx - 3.2)} ${num(cy + 1.6)} m ${num(cx)} ${num(cy - 1.6)} l ${num(cx + 3.2)} ${num(cy + 1.6)} l S Q`, false);
        body = `<< ${common} /FT /Btn /Ff ${FLAGS.pushButton} /H /N /DA (/Helv 9 Tf 0 g) /MK << >> /BS << /W 0 >> /AP << /N ${chevron} 0 R >> /A ${jsAction(document, "BT_press(event);")} >>`;
      } else if (field.k === "select" || field.k === "plot") {
        // The value box shows the chosen option, left-aligned like the website's select.
        const appearance = fieldAppearance(document, fonts, { width: r.w * PX, height: r.h * PX, fill, border, borderWidth: field.borderWidth * PX });
        body = `<< ${common} /FT /Tx /Ff ${FLAGS.readOnly | FLAGS.doNotScroll | FLAGS.doNotSpellCheck} /Q 0 /DA (/${daFont} ${fieldSize} Tf ${rgb(textColour)} rg) ${mk} /AP << /N ${appearance} 0 R >> >>`;
      } else {
        const readOnly = ["status", "note", "label", "badge"].indexOf(field.k) >= 0;
        const multiline = field.multiline || field.k === "clip" || (field.k === "note" && r.h > text.size * 1.9);
        const flags = (readOnly ? FLAGS.readOnly : 0) | (multiline ? FLAGS.multiline : 0) | FLAGS.doNotSpellCheck | (readOnly && !multiline ? FLAGS.doNotScroll : 0);
        const align = field.k === "badge" || text.align === "center" ? 1 : text.align === "right" ? 2 : 0;
        let style = mk, look = { fill, border, borderWidth: field.borderWidth * PX };
        if (field.k === "badge") { style = `/MK << >> /BS << /W 0 >>`; look = {}; }
        if (field.k === "clip") { style = `/MK << /BG [${rgb(paint("clipFill"))}] /BC [${rgb(paint("clipBorder"))}] >> /BS << /W 0.75 /S /D /D [2 2] >>`; look = { fill: paint("clipFill"), border: paint("clipBorder"), borderWidth: 0.75, dashed: true }; }
        if (field.k === "status") { style = `/MK << >> /BS << /W 0 >>`; look = {}; }
        const appearance = fieldAppearance(document, fonts, Object.assign({ width: r.w * PX, height: r.h * PX, value: field.value, placeholder: field.placeholder, font, size, colour: textColour, align, multiline }, look));
        const action = field.k === "text" || field.k === "search" ? ` /AA << /V ${jsAction(document, "BT_commit(event);")} >>` : "";
        const value = field.value ? ` /V ${textString(field.value)}` : "";
        body = `<< ${common} /FT /Tx /Ff ${flags} /Q ${align} /DA (/${field.k === "badge" ? "SSPBF" : daFont} ${field.k === "label" ? num(size) : fieldSize} Tf ${rgb(field.k === "badge" ? [1, 1, 1] : textColour)} rg) ${style} /AP << /N ${appearance} 0 R >>${action}${value} >>`;
      }
      const id = document.add(body);
      fieldIds.push(id);
      annots.push(id);
    });
    const contentId = document.stream("", drawPage(page, fonts), true);
    const openAction = page.open ? ` /AA << /O ${jsAction(document, page.open)} >>` : "";
    document.set(pageIds[pageIndex], `<< /Type /Page /Parent ${pagesId} 0 R /MediaBox [0 0 ${num(page.width * PX)} ${num(H)}] /Resources ${resourcesFor(pageIcons)} /Contents ${contentId} 0 R /Annots [${annots.map(id => id + " 0 R").join(" ")}]${openAction} >>`);
  });

  document.set(pagesId, `<< /Type /Pages /Kids [${pageIds.map(id => id + " 0 R").join(" ")}] /Count ${pageIds.length} >>`);
  document.set(acroFormId, `<< /Fields [${fieldIds.map(id => id + " 0 R").join(" ")}] /DR ${fontResources} /DA (/Helv 0 Tf 0 g) >>`);
  const scripts = spec.scripts.map(script => `${textString(script.name)} ${document.add(jsAction(document, script.code))} 0 R`).join(" ");
  const outlineRoot = document.reserve();
  const outlineItems = spec.outline.map(() => document.reserve());
  spec.outline.forEach((item, index) => {
    const links = [index > 0 ? `/Prev ${outlineItems[index - 1]} 0 R` : "", index < outlineItems.length - 1 ? `/Next ${outlineItems[index + 1]} 0 R` : ""].join(" ");
    document.set(outlineItems[index], `<< /Title ${textString(item.title)} /Parent ${outlineRoot} 0 R ${links} /Dest [${pageIds[item.page]} 0 R /XYZ 0 ${num(spec.pages[item.page].height * PX)} null] >>`);
  });
  document.set(outlineRoot, `<< /Type /Outlines /First ${outlineItems[0]} 0 R /Last ${outlineItems[outlineItems.length - 1]} 0 R /Count ${outlineItems.length} >>`);
  const infoId = document.add(`<< /Title ${textString(spec.title)} /Subject ${textString(spec.subject)} /Creator (BT1101 Prose Lab pdf/build.js) /Producer (BT1101 Prose Lab) /BT1101Version ${textString(spec.version)} >>`);
  document.set(catalogId, `<< /Type /Catalog /Pages ${pagesId} 0 R /AcroForm ${acroFormId} 0 R /Names << /JavaScript << /Names [${scripts}] >> >> /Outlines ${outlineRoot} 0 R /PageMode /UseNone /PageLayout /OneColumn /ViewerPreferences << /DisplayDocTitle true >> /OpenAction [${pageIds[0]} 0 R /FitH ${num(spec.pages[0].height * PX)}] /Lang (en-SG)${spec.closeScript ? ` /AA << /WC ${jsAction(document, spec.closeScript)} >>` : ""} >>`);
  return document.serialize(catalogId, infoId, spec.headerComment);
}

module.exports = { buildPdf, labelFields, PX };
