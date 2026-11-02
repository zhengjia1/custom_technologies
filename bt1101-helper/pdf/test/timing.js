// Times typical actions on the PDF script in the Acrobat sandbox. Node is far faster than
// Acrobat, so compare the numbers between builds rather than reading them as Acrobat times.
//   node test/timing.js            node --cpu-prof test/timing.js

const { loadPdfScript } = require("./acrobat-sim.js");

const session = loadPdfScript(11);
const results = [];
const time = (label, action, repeat) => {
  const start = process.hrtime.bigint();
  for (let i = 0; i < (repeat || 1); i++) action(i);
  results.push([label, Number(process.hrtime.bigint() - start) / 1e6 / (repeat || 1)]);
};
const pageOf = id => session.pageOf(id);
const find = (page, test) => session.fieldsOn(page).find(test);

console.log(`boot ${session.bootMs} ms`);
time("open page (cold, each operation)", i => session.page(i % session.registry.pages.length), session.registry.pages.length);
time("open page (warm, each operation)", i => session.page(i % session.registry.pages.length), session.registry.pages.length);
const oneSample = pageOf("one-sample");
session.page(oneSample);
const dataset = find(oneSample, field => field.spec.r === "#oneSampleDataset");
time("type into a text box", i => session.type(dataset.name, "data" + i), 10);
const parameter = find(oneSample, field => field.spec.r === "#oneSampleParameter");
time("change a dropdown", i => session.choose(parameter.name, "@" + (i % 2)), 10);
const ggplot = pageOf("ggplot");
session.page(ggplot);
const sample = find(ggplot, field => field.spec.k === "button" && /^Use sample/.test(field.spec.cap || ""));
time("press Use sample (plotting)", () => session.click(sample.name), 5);
const copy = find(ggplot, field => field.spec.k === "button" && /^Copy code/.test(field.spec.cap || ""));
time("press Copy code (plotting)", () => session.click(copy.name), 5);
const universal = pageOf("universal");
session.page(universal);
const input = find(universal, field => field.spec.r === "#inputText");
time("paste output into Universal detector", i => session.type(input.name, "##  Shapiro-Wilk normality test\n## data:  dataset$column1\n## W = 0.4" + i + ",\n## p-value < 2.2e-16"), 5);
results.forEach(([label, ms]) => console.log(`${label.padEnd(40)} ${ms.toFixed(1).padStart(8)} ms`));
console.log("errors:", session.errors().length);
