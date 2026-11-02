// Drives every operation page of the PDF script in the ES5 Acrobat sandbox:
// load samples, change answer options, press copy buttons, and check the results.

const { loadPdfScript } = require("./acrobat-sim.js");

function sweep(seed) {
  const session = loadPdfScript(seed);
  const problems = [];
  const report = [];
  session.page(0);
  const status = session.acrobat.fields["p0.status"];
  if (!status || !/^Ready/.test(status.value)) problems.push("status line: " + (status && status.value));
  const visible = page => session.fieldsOn(page).filter(field => field.display === 0);
  const buttons = (page, pattern) => visible(page).filter(field => field.spec.k === "button" && !field.readonly && pattern.test(field.caption || field.spec.cap || ""));

  // A page in its exact layout shows all of its input fields; hidden ones mean a mismatch.
  // (Rows the website adds while typing are left out: they appear one at a time.)
  const hiddenInputs = page => session.fieldsOn(page).filter(field => /^(text|select|plot|check)$/.test(field.spec.k) && field.display !== 0 && !/#(columnFields|lpConstraintRows)\|/.test(field.spec.r || "")).length;
  const operations = session.registry.pages.map(page => page.op).filter((op, index, all) => all.indexOf(op) === index);

  for (const op of operations) {
    session.page(session.pageOf(op));
    let page = session.current();
    if (session.registry.pages[page].op !== op) problems.push(`${op}: opening it showed page ${page} of ${session.registry.pages[page].op}`);
    buttons(page, /^Use sample/i).forEach(field => session.click(field.name));
    page = session.current();
    if (op === "universal") {
      const select = visible(page).find(field => field.spec.r === "#universalSampleSelect");
      if (select) { session.choose(select.name, "@1"); session.page(session.pageOf(op)); page = session.current(); }
    }
    if (op === "shapiro" || op === "pearson") buttons(page, /^Use sample/i).forEach(field => session.click(field.name));
    // Every layout-changing choice must land on a page captured in that layout.
    if (session.registry.layout && session.registry.layout[op]) {
      const pickers = visible(page).filter(field => (field.spec.k === "select" || field.spec.k === "plot") && field.spec.p && !/^Option \d/.test(field.value || ""));
      const layouts = new Set();
      pickers.forEach(picker => {
        const count = session.menuFor(picker.name).filter(item => item.cReturn).length;
        for (let index = 0; index < count; index++) {
          if (!session.choose(picker.name, "@" + index)) continue;
          const here = session.current(), entry = session.registry.pages[here];
          layouts.add(here);
          if (entry.op !== op) problems.push(`${op}: choosing option ${index} of ${picker.name} left the operation`);
          else if (!entry.sig) report.push(`${"".padEnd(20)} option ${index} of ${picker.spec.r || picker.name} used the catch-all page`);
          else if (hiddenInputs(here)) problems.push(`${op}: option ${index} of ${picker.spec.r || picker.name} shows page ${here} with ${hiddenInputs(here)} hidden input fields`);
        }
        session.choose(picker.name, "@0");
      });
      report.push(`${op.padEnd(20)} reached ${layouts.size} layout pages from its dropdowns`);
      session.page(session.pageOf(op));
    }
    page = session.current();
    // Change every visible answer option, then copy.
    visible(page).filter(field => field.spec.k === "select" && /^Option \d/.test(field.value || "")).forEach((field, index) => {
      const target = (index + 2) % 6;
      if (session.choose(field.name, "@" + target) && !session.acrobat.fields[field.name].value.startsWith("Option " + (target + 1))) problems.push(`${op}: answer option ${field.name} shows "${session.acrobat.fields[field.name].value}", expected Option ${target + 1}`);
    });
    // The pop-up menu lists the select's options with the current one ticked.
    const anyPicker = visible(page).find(field => field.spec.k === "select" && field.spec.p);
    if (anyPicker) {
      const menu = session.menuFor(anyPicker.name);
      if (!menu.length || menu.filter(item => item.bMarked).length !== 1) problems.push(`${op}: pop-up menu for ${anyPicker.name} has ${menu.length} items and ${menu.filter(item => item.bMarked).length} ticked`);
    }
    page = session.current();
    let copies = 0, emptyCopies = 0;
    const uncopied = [];
    buttons(page, /^(Copy\b|⧉)/i).forEach(field => {
      const clip = field.spec.c;
      if (clip) session.acrobat.field(clip).value = "";
      session.acrobat.clearFocus();
      session.click(field.name);
      if (clip && session.acrobat.focused === clip) { copies++; if (!session.acrobat.fields[clip].value.trim()) emptyCopies++; }
      else if (!session.acrobat.fields[field.name] || !session.acrobat.fields[field.name].readonly) uncopied.push(field.spec.cap || field.caption);
    });
    const outputs = visible(page).filter(field => field.spec.k === "out");
    const filled = outputs.filter(field => String(field.value).trim()).length;
    report.push(`${op.padEnd(20)} outputs ${filled}/${outputs.length}  copies ${copies}  alerts ${session.acrobat.alerts.length}`);
    if (emptyCopies) problems.push(`${op}: ${emptyCopies} copy boxes were empty`);
    if (uncopied.length) report.push(`${"".padEnd(20)} no copy from: ${uncopied.join(", ")}`);
  }
  // The detector must route regression output to its interpretation page.
  const universal = session.pageOf("universal");
  session.page(universal);
  const textarea = session.fieldsOn(universal).find(field => field.spec.k === "text" && field.spec.r === "#inputText");
  session.type(textarea.name, "Call:\nlm(formula = y ~ x1 + x2, data = d)\n\nCoefficients:\n            Estimate Std. Error t value Pr(>|t|)\n(Intercept)   1.0000     0.5000   2.000   0.0500\nx1            0.5000     0.1000   5.000 2.0e-05\nx2           -0.2000     0.1000  -2.000   0.0510\n\nResidual standard error: 1 on 47 degrees of freedom\nMultiple R-squared:  0.5,\tAdjusted R-squared:  0.48\nF-statistic: 23.5 on 2 and 47 DF,  p-value: 7.9e-08");
  const regressionPage = session.pageOf("multi-regression");
  if (session.acrobat.doc.pageNum !== regressionPage) problems.push(`detector routed regression output to page ${session.acrobat.doc.pageNum}, expected ${regressionPage}`);
  const prose = session.fieldsOn(regressionPage).filter(field => field.spec.k === "out" && /holding/i.test(field.value || ""));
  if (!prose.length) problems.push("regression prose missing after detector routing");
  session.errors().forEach(error => problems.push("bridge error: " + error));
  // Timing lines and skipped polyfills (expected where the engine locks a built-in) are
  // information, not failures.
  session.acrobat.log.filter(line => !/^BT1101 timing:/.test(line)).forEach(line => (/^BT1101: \d+ polyfill/.test(line) ? report : problems).push("console: " + line.split("\n")[0]));
  return { problems, report, bootMs: session.bootMs, alerts: session.acrobat.alerts };
}

module.exports = { sweep };

if (require.main === module) {
  const result = sweep(Number(process.argv[2]) || 31);
  console.log(`boot ${result.bootMs} ms`);
  result.report.forEach(line => console.log(line));
  console.log(`alerts shown: ${result.alerts.map(alert => alert.cTitle).join(" | ")}`);
  console.log(result.problems.length ? "PROBLEMS:\n" + result.problems.join("\n") : "No problems.");
}
