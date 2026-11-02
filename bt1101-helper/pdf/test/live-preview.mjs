// Renders pages the way Acrobat would show them after use: runs the PDF's script in the
// Acrobat sandbox, presses "Use sample", then paints the resulting field values.
//   node test/live-preview.mjs <outDir> [operation ids…] [--scale=0.8]
import fs from "fs";
import path from "path";
import { createRequire } from "module";
import { fileURLToPath } from "url";
import { createCanvas } from "@napi-rs/canvas";
import * as pdfjs from "pdfjs-dist/legacy/build/pdf.mjs";

const require = createRequire(import.meta.url);
const { loadPdfScript } = require("./acrobat-sim.js");
const [outDir, ...rest] = process.argv.slice(2);
const scale = Number((rest.find(arg => arg.startsWith("--scale=")) || "--scale=0.8").slice(8));
const wanted = rest.filter(arg => !arg.startsWith("--"));
const pdfPath = path.join(path.dirname(fileURLToPath(import.meta.url)), "..", "..", "BT1101 Prose Lab.pdf");

const session = loadPdfScript(4242);
const doc = await pdfjs.getDocument({ data: new Uint8Array(fs.readFileSync(pdfPath)), disableFontFace: true }).promise;
fs.mkdirSync(outDir, { recursive: true });
// Each operation, plus any "<operation>:<option>" request, e.g. "ggplot:9" picks plot type 9.
const operations = session.registry.pages.map(page => page.op).filter((op, i, all) => all.indexOf(op) === i);
const requests = wanted.length ? wanted : operations;
for (const request of requests) {
  const [op, option] = request.split(":");
  session.page(session.pageOf(op));
  let index = session.current();
  const fields = () => session.fieldsOn(session.current());
  fields().filter(field => field.spec.k === "button" && field.display === 0 && /^Use sample/i.test(field.caption || field.spec.cap || "")).forEach(field => session.click(field.name));
  if (op === "universal") { const select = fields().find(field => field.spec.r === "#universalSampleSelect"); if (select) session.choose(select.name, "@1"); session.page(session.pageOf(op)); }
  if (op === "shapiro" || op === "pearson") fields().filter(field => field.spec.k === "button" && /^Use sample/i.test(field.spec.cap || "")).forEach(field => session.click(field.name));
  if (option !== undefined) {
    const picker = fields().find(field => (field.spec.k === "plot" || field.spec.k === "select") && field.spec.p && !/^Option \d/.test(field.value || ""));
    if (picker) session.choose(picker.name, "@" + option);
    fields().filter(field => field.spec.k === "button" && field.display === 0 && /^Use sample/i.test(field.caption || field.spec.cap || "")).forEach(field => session.click(field.name));
  }
  index = session.current();
  const state = {};
  fields().forEach(field => { state[field.name] = field; });
  const page = await doc.getPage(index + 1);
  const storage = doc.annotationStorage;
  for (const annotation of await page.getAnnotations()) {
    const field = state[annotation.fieldName];
    if (!field) continue;
    if (field.spec.k === "label") { if (field.display === 1) storage.setValue(annotation.id, { value: "", hidden: true }); continue; }
    if (field.display === 1) { storage.setValue(annotation.id, { value: "", hidden: true }); continue; }
    if (annotation.checkBox) storage.setValue(annotation.id, { value: field.value === "Yes" });
    else if (annotation.fieldType === "Tx") storage.setValue(annotation.id, { value: String(field.value || "") });
  }
  const viewport = page.getViewport({ scale });
  const canvas = createCanvas(Math.ceil(viewport.width), Math.ceil(viewport.height));
  await page.render({ canvasContext: canvas.getContext("2d"), viewport, annotationMode: pdfjs.AnnotationMode.ENABLE_STORAGE }).promise;
  // Dropdowns get their items at run time; paint the selected label over the empty box.
  const context = canvas.getContext("2d");
  for (const annotation of await page.getAnnotations()) {
    const field = state[annotation.fieldName];
    if (!field || annotation.fieldType !== "Ch" || field.display === 1 || !field.items) continue;
    const item = field.items.find(entry => entry[1] === field.value);
    if (!item) continue;
    const [x1, y1, x2, y2] = annotation.rect;
    const [left, top] = viewport.convertToViewportPoint(x1, y2);
    const [right, bottom] = viewport.convertToViewportPoint(x2, y1);
    context.font = `${Math.round(10 * scale * 1.4)}px sans-serif`;
    context.fillStyle = "#1f1f1f";
    context.fillText(item[0].trim() + "  ▾", left + 6 * scale, (top + bottom) / 2 + 4 * scale, right - left - 8);
  }
  const file = path.join(outDir, `live-${request.replace(":", "-")}-page${index + 1}.png`);
  fs.writeFileSync(file, canvas.toBuffer("image/png"));
  console.log(file);
}
console.log("bridge errors:", session.errors().length ? session.errors() : "none");
