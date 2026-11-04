// Parse-checks every generated R-code option of the equivalence cases with R (parse(),
// which compiles the code without running it), as AGENTS.md requires. Texts whose
// canonical answer is not R code (placeholder messages) are reported separately.
//
//   node test/variants/equiv.js --only-node   (writes .build/variants-expected.json)
//   node test/variants/rparse.js [--rscript path]

const fs = require("fs");
const os = require("os");
const path = require("path");
const { spawnSync } = require("child_process");

const ROOT = path.join(__dirname, "..", "..");
const args = process.argv.slice(2);
const rscriptIndex = args.indexOf("--rscript");
const RSCRIPT = rscriptIndex >= 0 ? args[rscriptIndex + 1] : "C:\\Program Files\\R\\R-4.5.2\\bin\\Rscript.exe";
const SEP = "#@@BT1101_RECORD@@";

const cases = JSON.parse(fs.readFileSync(path.join(ROOT, ".build", "variants-expected.json"), "utf8")).filter(c => c.kind === "code");
const texts = [], index = new Map();
const id = text => { if (!index.has(text)) { index.set(text, texts.length); texts.push(text); } return index.get(text); };
const groups = cases.map(c => ({ source: c.source, canonical: id(c.text), options: c.expected.map(id) }));

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "bt1101-rparse-"));
const input = path.join(dir, "codes.txt"), script = path.join(dir, "check.R"), output = path.join(dir, "result.txt");
fs.writeFileSync(input, texts.map(text => text + "\n" + SEP).join("\n"), "utf8");
fs.writeFileSync(script, `
lines <- readLines(${JSON.stringify(input.replace(/\\/g, "/"))}, encoding = "UTF-8", warn = FALSE)
ends <- which(lines == "${SEP}")
starts <- c(1, head(ends, -1) + 1)
res <- character(length(ends))
for (i in seq_along(ends)) {
  code <- if (ends[i] > starts[i]) lines[starts[i]:(ends[i] - 1)] else character(0)
  res[i] <- tryCatch({ parse(text = code, keep.source = FALSE); "OK" }, error = function(e) gsub("[\\r\\n]+", " ", conditionMessage(e)))
}
writeLines(res, ${JSON.stringify(output.replace(/\\/g, "/"))}, useBytes = TRUE)
`, "utf8");
const run = spawnSync(RSCRIPT, ["--vanilla", script], { encoding: "utf8" });
if (run.status !== 0) { console.error(run.stderr); process.exit(1); }
const results = fs.readFileSync(output, "utf8").split(/\r?\n/);

let options = 0, failures = 0, nonCode = 0;
const shown = [];
groups.forEach(group => {
  if (results[group.canonical] !== "OK") { nonCode++; return; }
  group.options.forEach((option, k) => {
    options++;
    if (results[option] !== "OK") {
      failures++;
      if (shown.length < 5) shown.push(`${group.source} option ${k + 1}: ${results[option]}\n${texts[option].slice(0, 400)}`);
    }
  });
});
console.log(`${texts.length} distinct texts; ${options} options of ${groups.length - nonCode} R-code cases parsed, ${failures} failures; ${nonCode} cases whose canonical text is not R code (placeholders)`);
shown.forEach(text => console.log("\n" + text));
process.exitCode = failures ? 1 : 0;
