// npm test: checks the PDF built by `npm run build`.
//   1. minidom runs the website exactly like a real DOM (jsdom), step for step.
//   2. The PDF's own document script, in an ES5 Acrobat-like sandbox, drives every page.

const { runDifferential } = require("./differential.js");
const { sweep } = require("./sweep.js");

(async () => {
  let failed = false;
  console.log("1. Website on minidom vs jsdom");
  const differential = await runDifferential();
  console.log(`   ${differential.steps} steps, ${differential.differences.length} differences, ${differential.minidomErrors.length} minidom errors`);
  differential.differences.slice(0, 5).forEach(d => console.log(`   [${d.step}]\n     jsdom:   ${String(d.jsdom).slice(0, 200)}\n     minidom: ${String(d.minidom).slice(0, 200)}`));
  if (differential.differences.length || differential.minidomErrors.length) failed = true;

  console.log("2. PDF script in the Acrobat sandbox");
  const result = sweep(31);
  console.log(`   boot ${result.bootMs} ms, ${result.alerts.length} warning dialogs`);
  result.problems.forEach(problem => console.log("   " + problem));
  if (result.problems.length) failed = true;

  console.log(failed ? "\nFAILED" : "\nAll checks passed.");
  process.exit(failed ? 1 : 0);
})();
