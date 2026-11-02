// Build-time exploration on minidom. For one operation it tries every select option,
// checkbox, toggle, sample and copy button, and reports which elements:
//   ever        appear in some state (they get a place in the PDF),
//   conditional appear in some states only (their text is drawn as hideable fields),
//   dynamic     change their text (they are drawn as live fields).

const { refFor, isRendered } = require("../.build/refs-node.js");

const OUTPUT = "pre, .output-body, .pca-variance-prose, .sensitivity-prose-grid, .hypothesis-card, .mapping-card, .shape-guide, .lp-parser-status, .lp-constraint-error, .lp-table-wrap";

function activeBench(document) {
  return document.querySelector("main > .bench:not(.hidden), main > #lpBench:not(.hidden)");
}

function explore(boot, operation) {
  const session = boot();
  const { document, drain } = session;
  const Event = session.Event;
  const open = id => { const entry = id === "universal" ? document.querySelector(".universal-entry") : document.querySelector(`#tutorialList .operation[data-id="${id}"]`); entry.click(); drain(); };
  const fire = (target, type) => target.dispatchEvent(new Event(type, { bubbles: true }));
  const dismiss = accept => {
    for (let guard = 0; guard < 5; guard++) {
      const alert = document.querySelector(".variable-rename-alert:not(.hidden)");
      if (!alert) return;
      const buttons = alert.querySelectorAll("button");
      (accept ? buttons[buttons.length - 1] : buttons[0]).click();
      drain();
    }
  };
  open(operation);
  const bench = activeBench(document);
  const seenCount = {}, renderedCount = {}, texts = {};
  let states = 0;
  const record = () => {
    if (activeBench(document) !== bench) open(operation);
    states++;
    bench.querySelectorAll("*").forEach(element => {
      const ref = refFor(element, document);
      seenCount[ref] = (seenCount[ref] || 0) + 1;
      if (!isRendered(element, true)) return;
      renderedCount[ref] = (renderedCount[ref] || 0) + 1;
      if (!element.children.length || element.matches(OUTPUT) || element.id) {
        (texts[ref] || (texts[ref] = new Set())).add(element.textContent.replace(/\s+/g, " ").trim());
      }
    });
  };
  record();
  const controls = () => bench.querySelectorAll("select, input[type=checkbox], button.metric-option").filter(control => !control.matches(".answer-variant-select, .lp-operator") && isRendered(control, true));
  controls().forEach(control => {
    if (control.localName === "select") {
      const original = control.value;
      control._options().forEach(option => {
        if (option.disabled) return;
        control.value = option.value; fire(control, "input"); fire(control, "change"); drain(); dismiss(false); record();
      });
      control.value = original; fire(control, "input"); fire(control, "change"); drain(); dismiss(false);
    } else if (control.type === "checkbox") {
      control.click(); drain(); dismiss(true); record();
      control.click(); drain(); dismiss(true); record();
    } else { control.click(); drain(); record(); }
  });
  const search = bench.querySelector("#plotWorkflowSearch");
  if (search) {
    const original = search.value;
    document.querySelector("#plotWorkflow")._options().forEach(option => {
      search.value = option.textContent.trim(); fire(search, "input"); fire(search, "change"); drain(); record();
      bench.querySelectorAll("button").filter(button => /^Use sample/i.test(button.textContent.trim()) && isRendered(button, true)).forEach(button => { button.click(); drain(); });
      record();
    });
    search.value = original; fire(search, "input"); fire(search, "change"); drain();
  }
  for (let pass = 0; pass < 2; pass++) {
    bench.querySelectorAll("button").filter(button => /^Use sample/i.test(button.textContent.trim()) && isRendered(button, true) && !button.disabled).forEach(button => { button.click(); drain(); dismiss(true); record(); });
    if (operation === "universal") {
      const select = document.querySelector("#universalSampleSelect");
      select._options().slice(1).forEach(option => { open("universal"); select.value = option.value; fire(select, "change"); drain(); if (activeBench(document) === bench) record(); });
      open(operation);
      select.value = "shapiro"; fire(select, "change"); drain(); record();
    }
    if (operation === "shapiro" || operation === "pearson") { document.querySelector("#singleSampleBtn").click(); drain(); record(); }
  }
  bench.querySelectorAll("button").filter(button => /^Copy\b/i.test(button.textContent.trim()) && isRendered(button, true) && !button.disabled).forEach(button => { button.click(); drain(); dismiss(false); record(); });

  const ever = Object.keys(renderedCount);
  const conditional = ever.filter(ref => renderedCount[ref] < seenCount[ref] || seenCount[ref] < states);
  const dynamic = Object.keys(texts).filter(ref => texts[ref].size > 1);
  return { ever, conditional, dynamic };
}

// ---------------------------------------------------------------- layouts
//
// A PDF page cannot reflow, so each layout the website shows becomes its own page. A
// layout is the set of shown elements among those that come and go (outputs excluded:
// their content changes, not the page structure). This tries every combination of the
// controls that change the layout and keeps one way of reaching each distinct layout.

const NOT_LAYOUT = OUTPUT + ", .answer-variant-control, .completion-badge, .copyable-stage-status, #columnFields, #lpConstraintRows, .variable-rename-alert, .plot-workflow-menu";
const MAX_COMBINATIONS = 600;

function layoutVariants(boot, operation, exploration) {
  const { applySetting, acceptAlerts, resolveRef } = require("../.build/refs-node.js");
  const session = boot();
  const { document, drain } = session;
  const entry = operation === "universal" ? document.querySelector(".universal-entry") : document.querySelector(`#tutorialList .operation[data-id="${operation}"]`);
  entry.click(); drain();
  const bench = activeBench(document);
  bench.querySelectorAll("button").filter(button => /^Use sample/i.test(button.textContent.trim()) && isRendered(button, true) && !button.disabled).forEach(button => { button.click(); drain(); acceptAlerts(document); drain(); });

  const tracked = exploration.conditional.filter(ref => {
    const element = resolveRef(document, ref);
    // Buttons come and go with content (Copy prose once there is prose), not with the layout.
    return element && bench.contains(element) && !element.closest(NOT_LAYOUT) && !element.matches("option, optgroup, button") && !element.closest(".panel-actions");
  });
  const signature = () => tracked.filter(ref => { const element = resolveRef(document, ref); return element && isRendered(element, true); });
  const apply = settings => settings.filter(setting => { const done = applySetting(document, setting, session.Event); drain(); acceptAlerts(document); drain(); return done; });

  // Every control a user can set, with its possible values.
  const controls = [];
  bench.querySelectorAll("select").forEach(select => {
    if (select.matches(".answer-variant-select, .lp-operator") || select.id === "plotWorkflow" || select.closest("#columnFields, #lpConstraintRows")) return;
    controls.push({ values: select._options().map(option => ({ type: "select", ref: refFor(select, document), value: option.value })) });
  });
  bench.querySelectorAll('input[type="checkbox"]').forEach(box => {
    if (box.closest("#summaryStatisticFields")) return;
    const ref = refFor(box, document);
    controls.push({ values: [{ type: "check", ref, value: box.checked }, { type: "check", ref, value: !box.checked }] });
  });
  const metricGroups = new Set(bench.querySelectorAll(".metric-option").map(button => button.parentNode));
  metricGroups.forEach(group => controls.push({ values: group.querySelectorAll(".metric-option").map(button => ({ type: "metric", ref: refFor(button, document) })) }));
  if (bench.querySelector("#plotWorkflowSearch")) {
    controls.unshift({ values: document.querySelector("#plotWorkflow")._options().map(option => ({ type: "plot", value: option.textContent.trim() })) });
  }

  // Keep only controls that change the layout on their own.
  const baseline = signature().join("|");
  const affecting = controls.filter(control => control.values.some(setting => {
    const reset = boot();
    const doc = reset.document;
    const open = operation === "universal" ? doc.querySelector(".universal-entry") : doc.querySelector(`#tutorialList .operation[data-id="${operation}"]`);
    open.click(); reset.drain();
    applySetting(doc, setting, reset.Event); reset.drain(); acceptAlerts(doc); reset.drain();
    return tracked.filter(ref => { const element = resolveRef(doc, ref); return element && isRendered(element, true); }).join("|") !== baseline;
  }));

  // Walk the combinations of those controls, in page order, and record each layout.
  const variants = new Map();
  variants.set(baseline, { settings: [], sig: signature() });
  let combinations = 1;
  affecting.forEach(control => { combinations *= control.values.length; });
  const limit = Math.min(combinations, MAX_COMBINATIONS);
  const startSession = () => {
    const fresh = boot();
    const open = operation === "universal" ? fresh.document.querySelector(".universal-entry") : fresh.document.querySelector(`#tutorialList .operation[data-id="${operation}"]`);
    open.click(); fresh.drain();
    return fresh;
  };
  const sigOf = doc => tracked.filter(ref => { const element = resolveRef(doc, ref); return element && isRendered(element, true); });
  // One session walks every combination (each sets all the controls, in page order); a
  // fresh start per combination is far slower. The settings kept for a layout are those
  // that differ from the operation's defaults, so they reproduce it from a fresh start.
  const walker = startSession();
  const defaultState = startSession();
  const isDefault = setting => {
    if (setting.type === "plot") return defaultState.document.getElementById("plotWorkflowSearch").value.trim() === setting.value;
    const target = resolveRef(defaultState.document, setting.ref);
    if (!target) return false;
    if (setting.type === "select") return target.value === setting.value;
    if (setting.type === "check") return target.checked === setting.value;
    if (setting.type === "metric") return target.classList.contains("active");
    return false;
  };
  for (let index = 0; index < limit; index++) {
    let rest = index;
    const chosen = affecting.map(control => { const setting = control.values[rest % control.values.length]; rest = Math.floor(rest / control.values.length); return setting; });
    chosen.forEach(setting => { applySetting(walker.document, setting, walker.Event); walker.drain(); acceptAlerts(walker.document); walker.drain(); });
    const key = sigOf(walker.document).join("|");
    const settings = chosen.filter(setting => !isDefault(setting));
    const known = variants.get(key);
    if (!known || known.settings.length > settings.length) variants.set(key, { settings, sig: key ? key.split("|") : [] });
  }
  // Check each layout from a fresh start; keep what the settings really produce.
  const verified = new Map();
  variants.forEach(variant => {
    const fresh = startSession();
    variant.settings.forEach(setting => { applySetting(fresh.document, setting, fresh.Event); fresh.drain(); acceptAlerts(fresh.document); fresh.drain(); });
    const sig = sigOf(fresh.document), key = sig.join("|");
    const known = verified.get(key);
    if (!known || known.settings.length > variant.settings.length) verified.set(key, { settings: variant.settings, sig });
  });
  // The default layout stays first.
  const ordered = [verified.get(baseline)].filter(Boolean).concat(Array.from(verified.entries()).filter(([key]) => key !== baseline).map(([, value]) => value));
  return { tracked, variants: ordered, combinations, affecting: affecting.length };
}

module.exports = { explore, layoutVariants };
