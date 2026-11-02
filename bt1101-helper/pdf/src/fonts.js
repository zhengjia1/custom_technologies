// Reads the Source Sans Pro TrueType fonts embedded in index.html and prepares them for
// embedding in the PDF as WinAnsi-encoded simple fonts.

const WIN_ANSI_HIGH = { 0x80: 0x20ac, 0x82: 0x201a, 0x83: 0x0192, 0x84: 0x201e, 0x85: 0x2026, 0x86: 0x2020, 0x87: 0x2021, 0x88: 0x02c6, 0x89: 0x2030, 0x8a: 0x0160, 0x8b: 0x2039, 0x8c: 0x0152, 0x8e: 0x017d, 0x91: 0x2018, 0x92: 0x2019, 0x93: 0x201c, 0x94: 0x201d, 0x95: 0x2022, 0x96: 0x2013, 0x97: 0x2014, 0x98: 0x02dc, 0x99: 0x2122, 0x9a: 0x0161, 0x9b: 0x203a, 0x9c: 0x0153, 0x9e: 0x017e, 0x9f: 0x0178 };
// Symbols the website writes that WinAnsi lacks (H₀, μ, ≤ …) take over code slots that are
// undefined or practically unused in WinAnsi. The PDF font declares them in /Differences.
const EXTRA_SLOTS = [0x7f, 0x81, 0x8d, 0x8f, 0x90, 0x9d, 0x83, 0x88, 0x98, 0x8a, 0x8c, 0x8e, 0x9a, 0x9c, 0x9e, 0x9f, 0xa4, 0xa6, 0xa8, 0xaa, 0xaf, 0xb8, 0xba, 0xbc, 0xbd, 0xbe];
const EXTRA_CHARACTERS = [0x2080, 0x2081, 0x2082, 0x2083, 0x2084, 0x2085, 0x2086, 0x2087, 0x2088, 0x2089, 0x3bc, 0x3c1, 0x3c3, 0x3b2, 0x3b1, 0x3b5, 0x2265, 0x2264, 0x2260, 0x2212, 0x221a, 0x2248, 0x302, 0x304, 0x2193, 0x2191];
const CODE_TO_UNICODE = {};
for (let code = 32; code < 256; code++) {
  const unicode = code >= 0x80 && code < 0xa0 ? WIN_ANSI_HIGH[code] : code === 0x7f ? 0 : code;
  if (unicode) CODE_TO_UNICODE[code] = unicode;
}
EXTRA_SLOTS.forEach((code, index) => { CODE_TO_UNICODE[code] = EXTRA_CHARACTERS[index]; });
const UNICODE_TO_WIN_ANSI = {};
Object.keys(CODE_TO_UNICODE).forEach(code => { UNICODE_TO_WIN_ANSI[CODE_TO_UNICODE[code]] = Number(code); });
const DIFFERENCES = EXTRA_SLOTS.map((code, index) => ({ code, name: "uni" + EXTRA_CHARACTERS[index].toString(16).toUpperCase().padStart(4, "0") })).sort((a, b) => a.code - b.code);

const FALLBACK = { "→": "->", "←": "<-", "⌁": "", "✓": "", "☾": "", "☰": "", "↺": "", "⧉": "", "ₖ": "k", "⩾": "≥", "⩽": "≤", "ℹ": "i" };

function parseTrueType(buffer) {
  const tables = {};
  const count = buffer.readUInt16BE(4);
  for (let i = 0; i < count; i++) {
    const offset = 12 + i * 16;
    tables[buffer.toString("latin1", offset, offset + 4)] = { offset: buffer.readUInt32BE(offset + 8), length: buffer.readUInt32BE(offset + 12) };
  }
  const head = tables.head.offset, hhea = tables.hhea.offset, hmtx = tables.hmtx.offset;
  const unitsPerEm = buffer.readUInt16BE(head + 18);
  const bbox = [buffer.readInt16BE(head + 36), buffer.readInt16BE(head + 38), buffer.readInt16BE(head + 40), buffer.readInt16BE(head + 42)];
  const ascender = buffer.readInt16BE(hhea + 4), descender = buffer.readInt16BE(hhea + 6);
  const metrics = buffer.readUInt16BE(hhea + 34);
  const advance = glyph => buffer.readUInt16BE(hmtx + 4 * Math.min(glyph, metrics - 1));
  let capHeight = ascender * 0.7, weight = 400;
  if (tables["OS/2"]) {
    const os2 = tables["OS/2"].offset, version = buffer.readUInt16BE(os2);
    weight = buffer.readUInt16BE(os2 + 4);
    if (version >= 2) capHeight = buffer.readInt16BE(os2 + 88);
  }
  // cmap format 4 (Windows Unicode BMP)
  const cmap = tables.cmap.offset;
  const subtables = buffer.readUInt16BE(cmap + 2);
  let format4 = null;
  for (let i = 0; i < subtables; i++) {
    const platform = buffer.readUInt16BE(cmap + 4 + i * 8), encoding = buffer.readUInt16BE(cmap + 6 + i * 8);
    const offset = cmap + buffer.readUInt32BE(cmap + 8 + i * 8);
    if (platform === 3 && encoding === 1 && buffer.readUInt16BE(offset) === 4) format4 = offset;
  }
  const glyphFor = code => {
    const segments = buffer.readUInt16BE(format4 + 6) / 2;
    const ends = format4 + 14, starts = ends + segments * 2 + 2, deltas = starts + segments * 2, ranges = deltas + segments * 2;
    for (let s = 0; s < segments; s++) {
      if (buffer.readUInt16BE(ends + s * 2) < code) continue;
      const start = buffer.readUInt16BE(starts + s * 2);
      if (start > code) return 0;
      const delta = buffer.readInt16BE(deltas + s * 2), rangeOffset = buffer.readUInt16BE(ranges + s * 2);
      if (!rangeOffset) return (code + delta) & 0xffff;
      const glyph = buffer.readUInt16BE(ranges + s * 2 + rangeOffset + (code - start) * 2);
      return glyph ? (glyph + delta) & 0xffff : 0;
    }
    return 0;
  };
  const scale = 1000 / unitsPerEm;
  const widths = [];
  for (let code = 32; code < 256; code++) {
    const unicode = CODE_TO_UNICODE[code];
    const glyph = unicode ? glyphFor(unicode) : 0;
    widths.push(Math.round(advance(glyph) * scale));
  }
  const plainWidths = [];
  for (let code = 32; code < 256; code++) {
    const unicode = code >= 0x80 && code < 0xa0 ? WIN_ANSI_HIGH[code] : code === 0x7f ? 0 : code;
    plainWidths.push(Math.round(advance(unicode ? glyphFor(unicode) : 0) * scale));
  }
  return {
    buffer, widths, plainWidths, weight,
    ascent: Math.round(ascender * scale), descent: Math.round(descender * scale), capHeight: Math.round(capHeight * scale),
    bbox: bbox.map(value => Math.round(value * scale)), hasGlyph: unicode => glyphFor(unicode) > 0,
    glyphFor, advance: glyph => Math.round(advance(glyph) * scale), glyphCount: tables.maxp ? buffer.readUInt16BE(tables.maxp.offset + 4) : metrics
  };
}

function loadWebsiteFonts(html) {
  const fonts = [];
  for (const match of html.matchAll(/@font-face\s*{([^}]*)}/g)) {
    const data = /base64,([A-Za-z0-9+/=]+)/.exec(match[1]);
    const weight = /font-weight:\s*(\d+)/.exec(match[1]);
    if (data) fonts.push(Object.assign(parseTrueType(Buffer.from(data[1], "base64")), { cssWeight: Number(weight ? weight[1] : 400) }));
  }
  fonts.sort((a, b) => a.cssWeight - b.cssWeight);
  return { regular: fonts[0], semibold: fonts[1] || fonts[0], bold: fonts[2] || fonts[fonts.length - 1] };
}

function toWinAnsi(text) {
  let bytes = "";
  for (const char of String(text)) {
    const code = UNICODE_TO_WIN_ANSI[char.codePointAt(0)];
    if (code !== undefined) bytes += String.fromCharCode(code);
    else if (FALLBACK[char] !== undefined) bytes += toWinAnsi(FALLBACK[char]);
    else bytes += "?";
  }
  return bytes;
}

function measure(font, text, size) {
  const bytes = toWinAnsi(text);
  let width = 0;
  for (let i = 0; i < bytes.length; i++) { const code = bytes.charCodeAt(i); width += code >= 32 ? font.widths[code - 32] : 0; }
  return width * size / 1000;
}

// Characters the drawn text can show: the encoding's characters that the font has a glyph for.
function drawableCodePoints(font) {
  return Object.keys(UNICODE_TO_WIN_ANSI).map(Number).filter(unicode => unicode === 32 || font.hasGlyph(unicode));
}

module.exports = { loadWebsiteFonts, toWinAnsi, measure, DIFFERENCES, drawableCodePoints };
