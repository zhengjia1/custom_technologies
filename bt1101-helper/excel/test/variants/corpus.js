// Collects the canonical prose and R code the website produces (Option 1 of every answer)
// by booting the real site on minidom and exercising every operation with its samples,
// selects and toggles, and every plot type in both axis modes. Writes
// excel/.build/variants-corpus.json: [{ kind: "prose"|"code", text, source }].
//
//   node test/variants/corpus.js

const fs = require("fs");
const path = require("path");
const { bootNodeApp } = require("../../../pdf/test/harness.js");
const { mulberry32 } = require("./engine.js");

const OUT = path.join(__dirname, "..", "..", ".build", "variants-corpus.json");
const HORIZON = 700;

function settle(tasks, errors) {
  const limit = tasks.now + HORIZON;
  for (let steps = 0; steps < 20000; steps++) {
    if (tasks.micro.length) { const fn = tasks.micro.shift(); try { fn(); } catch (error) { errors.push(error); } continue; }
    let next = -1;
    tasks.timers.forEach((timer, i) => { if (timer.due <= limit && (next < 0 || timer.due < tasks.timers[next].due || (timer.due === tasks.timers[next].due && timer.seq < tasks.timers[next].seq))) next = i; });
    if (next < 0) return;
    const timer = tasks.timers.splice(next, 1)[0];
    tasks.now = Math.max(tasks.now, timer.due);
    try { timer.fn.apply(null, timer.args); } catch (error) { errors.push(error); }
  }
}

const megaGroups = [
  { scope: "#pcaCodeSections", children: "[data-copy-pca-section]", kind: "code" },
  { scope: "#advancedRegressionBench", children: "#copyAdvancedRegressionCodeBtn, #copyAdvancedRegressionPredictionCodeBtn, #copyAdvancedRegressionConfidenceCodeBtn", kind: "code" }
];

async function main() {
  Math.random = mulberry32(1);
  const { browser, errors, tasks } = bootNodeApp();
  settle(tasks, errors);
  const document = browser.document, Event = browser.globals.Event;
  const drain = () => settle(tasks, errors);
  const seen = new Map();
  const add = (kind, text, source) => {
    if (!text || !text.trim()) return;
    const key = kind + "\u0000" + text;
    if (!seen.has(key)) seen.set(key, { kind, text, source });
  };
  const fire = (element, type) => element.dispatchEvent(new Event(type, { bubbles: true }));
  const canon = () => {
    document.querySelectorAll(".answer-variant-select").forEach(select => { if (select.value !== "1") { select.value = "1"; fire(select, "change"); } });
    drain();
  };
  const outputNodes = (button) => {
    const codeSection = button.closest(".pca-code-block");
    if (codeSection) return [...codeSection.querySelectorAll(".generated-code")];
    const proseSection = button.closest(".sensitivity-prose-card, .prose-card");
    if (proseSection) return [...proseSection.querySelectorAll("p")];
    const panel = button.closest(".panel");
    if (!panel) return [];
    const code = /code/.test(`${button.id} ${button.textContent}`.toLowerCase()) || button.matches("[data-copy-pca-section]");
    return [...panel.querySelectorAll(code ? ".generated-code, .interval-code" : ".prose-card p, .pca-variance-prose:not(.is-error), .sensitivity-prose-card p")];
  };
  const harvest = (source) => {
    canon();
    document.querySelectorAll(".generated-code, .interval-code").forEach(node => add("code", node.textContent, source));
    document.querySelectorAll(".prose-card p, .pca-variance-prose:not(.is-error), .sensitivity-prose-card p, #generatedProse").forEach(node => add("prose", node.textContent, source));
    megaGroups.forEach(group => {
      const scope = document.querySelector(group.scope);
      if (!scope) return;
      const values = [...scope.querySelectorAll(group.children)].filter(button => !button.disabled && !button.closest(".hidden"))
        .flatMap(button => outputNodes(button)).map(node => node.textContent);
      if (values.length > 1) {
        const markers = values.slice(0, -1).map((_, index) => `# __BT1101_SECTION_${index + 1}__`);
        add("code", values.map((value, index) => index < markers.length ? `${value}\n${markers[index]}` : value).join("\n"), source + " (combined)");
      }
    });
  };
  const visibleBench = () => document.querySelector("main > .bench:not(.hidden), main > #lpBench:not(.hidden)");
  const exerciseControls = (id) => {
    const bench = visibleBench();
    if (!bench) return;
    const selects = [...bench.querySelectorAll("select")].filter(select => !select.matches(".answer-variant-select, .sample-select") && select.id !== "plotWorkflow");
    selects.forEach(select => {
      const original = select.value;
      [...select.options].forEach(option => {
        if (option.disabled) return;
        select.value = option.value; fire(select, "input"); fire(select, "change"); drain();
        harvest(`${id} ${select.id}=${option.value}`);
      });
      select.value = original; fire(select, "input"); fire(select, "change"); drain();
    });
    const checks = [...bench.querySelectorAll('input[type="checkbox"]')];
    checks.forEach(check => {
      check.click(); drain();
      harvest(`${id} toggle ${check.id || check.className}`);
      const alert = document.querySelector(".variable-rename-alert:not(.hidden) .variable-rename-dismiss, .variable-rename-alert:not(.hidden) button");
      if (alert) { alert.click(); drain(); }
      check.click(); drain();
      const alert2 = document.querySelector(".variable-rename-alert:not(.hidden) button");
      if (alert2) { alert2.click(); drain(); }
    });
  };

  const entries = [document.querySelector(".universal-entry"), ...document.querySelectorAll("#tutorialList .operation")];
  for (const entry of entries) {
    const id = entry.dataset.id || "universal";
    entry.click(); drain();
    harvest(id);
    const bench = visibleBench();
    if (id === "universal") {
      const select = document.querySelector("#universalSampleSelect");
      [...select.options].filter(option => option.value).forEach(option => { select.value = option.value; fire(select, "change"); drain(); harvest(`universal sample ${option.value}`); });
      continue;
    }
    if (bench && bench.id === "standardBench") {
      const single = document.querySelector("#singleSampleBtn");
      if (!single.classList.contains("hidden")) { single.click(); drain(); harvest(id + " sample"); }
      continue;
    }
    if (!bench) continue;
    const loaders = [...bench.querySelectorAll("button")].filter(button => /^load/i.test(button.id) || /^Use (parser )?sample/i.test(button.textContent.trim()));
    loaders.forEach(button => { if (!button.disabled) { button.click(); drain(); harvest(`${id} ${button.id}`); } });
    if (id === "ggplot") continue;
    exerciseControls(id);
  }

  // Summarize for stats: one column, two columns, then each statistic toggle.
  document.querySelector('#tutorialList .operation[data-id="summarize"]').click(); drain();
  const typeColumn = (value) => {
    const inputs = [...document.querySelectorAll("#columnFields .column-input")];
    const target = inputs.find(input => !input.value) || inputs[inputs.length - 1];
    target.value = value; fire(target, "input"); drain();
  };
  typeColumn("column1"); harvest("summarize one");
  typeColumn("column2"); harvest("summarize two");
  exerciseControls("summarize");
  const summaryDataset = document.querySelector("#summaryDataset");
  summaryDataset.value = "survey_data"; fire(summaryDataset, "input"); drain(); harvest("summarize dataset");
  // ANOVA standard route (Step 3a) after a non-significant variance test.
  document.querySelector('#tutorialList .operation[data-id="anova-many"]').click(); drain();
  for (const id of ["loadAnovaSampleBtn", "loadAnovaStandardOutputBtn", "loadAnovaWelchOutputsBtn"]) {
    const button = document.querySelector("#" + id);
    if (button && !button.disabled) { button.click(); drain(); harvest("anova " + id); }
  }
  const groupCount = document.querySelector("#anovaGroupCount");
  if (groupCount) { groupCount.value = "2"; fire(groupCount, "input"); drain(); harvest("anova two groups");
    const v = document.querySelector("#loadAnovaVarianceOutputBtn"); if (v && !v.disabled) { v.click(); drain(); harvest("anova two groups variance"); }
    const st = document.querySelector("#loadAnovaStandardOutputBtn"); if (st && !st.disabled) { st.click(); drain(); harvest("anova two groups standard"); } }

  // Every plot type, before and after its sample, in both axis modes and with transformations.
  document.querySelector('#tutorialList .operation[data-id="ggplot"]').click(); drain();
  const workflow = document.querySelector("#plotWorkflow"), search = document.querySelector("#plotWorkflowSearch");
  const axisToggle = () => document.querySelector("#ggplotBench .plot-axis-interval-toggle");
  for (const option of [...workflow.options]) {
    for (const axis of [false, true]) {
      if (axisToggle() && axisToggle().checked !== axis) { axisToggle().click(); drain(); }
      search.value = option.textContent.trim(); fire(search, "input"); fire(search, "change"); drain();
      harvest(`plot ${option.value} axis=${axis}`);
      const sample = document.querySelector("#loadGgplotSampleBtn");
      if (!sample.disabled) { sample.click(); drain(); harvest(`plot ${option.value} sample axis=${axis}`); }
      const style = document.querySelector("#plotStyle");
      if (!document.querySelector("#plotStyleField").classList.contains("hidden")) {
        for (const value of ["grouped", "stacked", "mosaic"]) { style.value = value; fire(style, "change"); drain(); harvest(`plot ${option.value} style=${value} axis=${axis}`); }
        style.value = "grouped"; fire(style, "change"); drain();
      }
      for (const [selector, values] of [["#plotXTransform", ["log", "sqrt", "inverse"]], ["#plotYTransform", ["log10", "log2"]]]) {
        const select = document.querySelector(selector);
        if (select.disabled || select.closest(".hidden")) continue;
        for (const value of values) {
          if (![...select.options].some(o => o.value === value && !o.disabled)) continue;
          select.value = value; fire(select, "input"); fire(select, "change"); drain(); harvest(`plot ${option.value} ${selector}=${value} axis=${axis}`);
        }
        select.value = "none"; fire(select, "input"); fire(select, "change"); drain();
      }
      // Custom title and other dataset name.
      const title = document.querySelector("#plotTitle"), dataset = document.querySelector("#plotDataset");
      title.value = "My custom chart"; fire(title, "input"); dataset.value = "survey_data"; fire(dataset, "input"); drain();
      harvest(`plot ${option.value} custom axis=${axis}`);
      title.value = ""; fire(title, "input"); dataset.value = "dataset"; fire(dataset, "input"); drain();
    }
  }
  // Simple regression and K-means also follow the axis toggle.
  for (const id of ["lm-predict", "kmeans", "pca"]) {
    document.querySelector(`#tutorialList .operation[data-id="${id}"]`).click(); drain();
    const toggle = visibleBench()?.querySelector(".plot-axis-interval-toggle");
    if (toggle && !toggle.checked) { toggle.click(); drain(); }
    harvest(id + " axis=on");
    exerciseControls(id + " axis=on");
    if (toggle && toggle.checked) { toggle.click(); drain(); }
  }

  const corpus = [...seen.values()];
  fs.writeFileSync(OUT, JSON.stringify(corpus, null, 1));
  const counts = corpus.reduce((acc, item) => { acc[item.kind] = (acc[item.kind] || 0) + 1; return acc; }, {});
  console.log(`Wrote ${OUT}: ${corpus.length} texts`, counts, errors.length ? `(${errors.length} website errors: ${errors.slice(0, 3).map(e => e && e.message).join(" | ")})` : "");
}

main().catch(error => { console.error(error); process.exitCode = 1; });
