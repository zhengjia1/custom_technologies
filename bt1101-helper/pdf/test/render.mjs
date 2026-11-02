// Renders PDF pages to PNG (with form-field appearances) for visual checks.
//   node test/render.mjs <pdf> <outDir> [page numbers…] [--scale=1]
import fs from "fs";
import path from "path";
import { createCanvas } from "@napi-rs/canvas";
import * as pdfjs from "pdfjs-dist/legacy/build/pdf.mjs";

const [input, outDir, ...rest] = process.argv.slice(2);
const scale = Number((rest.find(arg => arg.startsWith("--scale=")) || "--scale=1").slice(8));
const pages = rest.filter(arg => /^\d+$/.test(arg)).map(Number);
const doc = await pdfjs.getDocument({ data: new Uint8Array(fs.readFileSync(input)), disableFontFace: true, useSystemFonts: false }).promise;
fs.mkdirSync(outDir, { recursive: true });
for (const number of pages.length ? pages : Array.from({ length: doc.numPages }, (_, i) => i + 1)) {
  const page = await doc.getPage(number);
  const viewport = page.getViewport({ scale });
  const canvas = createCanvas(Math.ceil(viewport.width), Math.ceil(viewport.height));
  await page.render({ canvasContext: canvas.getContext("2d"), viewport, annotationMode: pdfjs.AnnotationMode.ENABLE_STORAGE }).promise;
  const file = path.join(outDir, `page-${String(number).padStart(2, "0")}.png`);
  fs.writeFileSync(file, canvas.toBuffer("image/png"));
  console.log(file, Math.ceil(viewport.width), "x", Math.ceil(viewport.height));
}
