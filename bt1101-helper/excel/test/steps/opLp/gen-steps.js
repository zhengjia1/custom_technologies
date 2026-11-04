// Generates reset-all.txt, parser-extra.txt, sensitivity-extra.txt and detector.txt in this
// folder (the other opLp step files are hand-written). Run: node test/steps/opLp/gen-steps.js
const fs = require("fs");
const path = require("path");
const dir = process.argv[2] || __dirname;
const esc = text => String(text).replace(/\\/g, "\\\\").replace(/\n/g, "\\n").replace(/\t/g, "\\t");
const step = (...parts) => parts.map((p, i) => (i ? esc(p) : p)).join("\t");

const row = n => `#lpConstraintRows .lp-constraint-row:nth-child(${n})`;
const sec = key => `#sensitivityInterpretation [data-sensitivity-copy="${key}"]`;

const snapshot = () => [
  step("class", "#lpBench .panel:nth-child(1)"),
  step("class", "#lpBench .code-panel"),
  step("class", "#lpBench .lp-table-panel"),
  step("class", "#sensitivityInterpretationPanel"),
  step("count", "#lpBench .completion-badge"),
  step("count", "#lpBench .is-copy-complete"),
  step("count", "#lpBench .lp-collapsible.is-collapsed"),
  step("attr", "#sensitivityInterpretationPanel .lp-collapse-toggle", "aria-expanded"),
  step("attr", "#lpBench .code-panel .lp-collapse-toggle", "aria-expanded"),
  step("count", "#lpConstraintRows .lp-constraint-row"),
  step("read", "#lpParseStatus"),
];
const sectionButtons = () => ["solution", "shadow", "reduced", "slack"].map(key => step("count", `${sec(key)}`));
const buttonTexts = () => [step("count", "#sensitivityInterpretation [data-sensitivity-copy]"), step("count", "#sensitivityInterpretation [data-sensitivity-copy][disabled]"), step("count", "#sensitivityInterpretation .sensitivity-prose-card.is-copy-complete")];
const model = () => [
  step("val", "#lpDirection"), step("val", "#lpVariableType"), step("val", "#lpVariables"), step("val", "#lpObjective"),
  step("val", "#lpBinary"), step("checked", "#lpSensitivity"), step("disabled", "#lpSensitivity"),
  step("class", "#lpBinaryField"), step("class", "#lpSensitivityField"),
  step("read", "#lpConstraintRows"), step("read", "#lpConstraintError"), step("read", "#lpTablePreview"),
  step("canon"), step("read", "#lpCode"), step("disabled", "#copyLpCodeBtn"), step("disabled", "#copyLpAsciiBtn"),
];
const completeAll = () => [
  step("canon"),
  step("click", "#copyLpCodeBtn"), step("click", "#dismissPlotAxisRangeAlert"),
  step("click", "#copyLpAsciiBtn"),
  step("click", sec("solution")), step("click", "#dismissVariableRenameAlert"),
  step("click", sec("shadow")),
  ...snapshot(), ...buttonTexts(),
];

// ---------------------------------------------------------------- reset-all.txt
{
  const out = ["# Every optimisation-model change resets the whole Tutorial 8 workflow.", step("open", "lpsolve"),
    step("click", "#loadLpSampleBtn"), step("click", "#loadSensitivitySampleBtn")];
  const changes = [
    [step("change", "#lpDirection", "min")],
    [step("change", "#lpDirection", "max")],
    [step("input", "#lpObjective", "0.15, 0.40, 0.65")],
    [step("input", `${row(2)} .lp-name`, "Storage space")],
    [step("input", `${row(1)} .lp-rhs`, "210")],
    [step("input", `${row(3)} .lp-coefficients`, "1, 0, 0.5")],
    [step("select", `${row(2)} .lp-operator`, "2")],
    [step("click", "#lpSensitivity")],
    [step("click", "#lpSensitivity")],
    [step("change", "#lpVariableType", "integer")],
    [step("change", "#lpVariableType", "continuous"), step("click", "#lpSensitivity")],
    [step("input", "#lpBinary", "1")],
    [step("input", "#lpVariables", "X1, X2, X3")],
    [step("input", `${row(4)} .lp-name`, "Extra")],
    [step("click", `${row(4)} .lp-remove-constraint`)],
    [step("input", "#sensDuals", "0.01, 0.02, 0.03, -0.1, 0, 0")],
    [step("input", "#sensRawOutput", "X1 X2 X3\n7 2.2 0\n[1] 100\n[1] 0\n[1] 0.01 0.02 0.03 -0.1 0 0\n[1] 0.1 0.2 0.3\n[1] 0.3 0.4 0.5\nNULL\nNULL")],
    [step("input", "#lpSmartInput", "Maximise 3X1 +")],
    [step("input", "#lpSmartInput", "Maximise 0.15X1 + 0.4X2 + 0.6X3\nSubject to\nBudget: 15X1 + 30X2 + 40X3 <= 200\nStorage: 7X1 + 5X2 + 8X3 <= 60\nDiet: X1 <= 7")],
    [step("click", "#loadSensitivitySampleBtn")],
    [step("click", "#resetLpBtn")],
    [step("click", "#loadLpSampleBtn"), step("click", "#loadSensitivitySampleBtn")],
    [step("click", "#loadLpParserSampleBtn")],
  ];
  for (const change of changes) {
    if (!change.includes(step("click", "#loadSensitivitySampleBtn")) || change.length === 1) out.push(...completeAll());
    out.push(...change);
    out.push(...snapshot(), ...buttonTexts(), step("val", "#sensRawOutput"), step("read", "#sensParseStatus"));
  }
  // Partial section copies keep the card open; Copy prose completes every section.
  out.push(step("click", "#loadLpSampleBtn"), step("click", "#loadSensitivitySampleBtn"), step("canon"));
  for (const key of ["slack", "reduced", "shadow"]) out.push(step("click", sec(key)), step("clip"), step("toast"), ...snapshot(), ...buttonTexts());
  out.push(step("click", sec("solution")), step("clip"), step("toast"), step("class", "#variableRenameAlert"), ...snapshot(), ...buttonTexts());
  out.push(step("click", "#dismissVariableRenameAlert"), step("class", "#variableRenameAlert"));
  out.push(step("input", "#sensCoefTo", "0.3, 0.4, 0.55"), ...snapshot(), ...buttonTexts());
  out.push(step("click", "#copySensitivityProseBtn"), step("clip"), step("toast"), step("class", "#variableRenameAlert"), ...snapshot(), ...buttonTexts());
  out.push(step("click", "#dismissVariableRenameAlert"));
  out.push(step("click", "#sensitivityInterpretationPanel .lp-collapse-toggle"), ...snapshot());
  out.push(step("click", "#sensitivityInterpretationPanel .lp-collapse-toggle"), ...snapshot());
  fs.writeFileSync(path.join(dir, "reset-all.txt"), out.join("\n") + "\n");
}

// ---------------------------------------------------------------- parser-extra.txt
{
  const models = [
    // LaTeX short comparison forms, \times, spacing commands, braces.
    "\\text{Maximise } P = 3 \\times X_{1} + 2\\cdot X_{2}\nSubject to\n\\text{Steel}: 2X_{1} + X_{2} \\le 100\n\\text{Paint}: X_{1} \\,+\\, 3X_{2} \\leq{} 90\n\\text{Demand}: X_{1} \\ge 10\n\\text{Floor}: X_{2} \\geq{} 5\nX_{1}, X_{2} \\ge 0",
    // Lower-case subscripts, unicode minus, en dash and em dash.
    "Minimize cost = 4x_1 + 3x_2 − x_3\nsubject to\nc1: x_1 + x_2 + x_3 ≥ 10\nc2: 2x_1 – x_2 ≤ 8\nc3: x_2 — x_3 = 1",
    // Markdown bold names and code spans; non-negativity named row with two variables.
    "Maximise Z = 5A + 4B\nSubject to\n**Labour hours**: 6A + 4B <= 24\n`Material`: A + 2B <= 6\n_Mix_: -A + B <= 1\nNon negativity: 1A + 0B >= 0",
    // A unit lower bound on one of several variables; non-unit bound stays a constraint.
    "Maximise 2X1 + 3X2 + X3\nSubject to\nCap: X1 + X2 + X3 <= 10\nX3 >= 0\nMinimum X2: 2X2 >= 0\nAt least: X1 >= 1\nNonnegativity: X1 + X2 >= 0",
    // Objective written with the expression on the left; multiline with trailing plus.
    "Maximise\n12X1 + 9X2 +\n  7X3 = Profit\nSubject to\nA: X1 + X2 + X3 <= 30\nB: 2X1 + X2 <= 40",
    // Short direction words are not recognised.
    "Max: 3X1 + 2X2\nSubject to\nA: X1 + X2 <= 4",
    "Min 3X1 + 2X2\nst\nA: X1 + X2 >= 4",
    // Constraint without numeric RHS is skipped.
    "Minimise 2X1 + 5X2\nSubject to\nA: X1 + X2 >= budget\nB: X1 + 3X2 >= 9",
    // Binary declarations: some, all, both orders; integer only.
    "Maximise 4X1 + 5X2 + 3X3\nSubject to\nW: 2X1 + 3X2 + X3 <= 5\nX2 and X3 are binary",
    "Maximise 4X1 + 5X2 + 3X3\nSubject to\nW: 2X1 + 3X2 + X3 <= 5\nbinary: all",
    "Maximise 4X1 + 5X2\nSubject to\nW: 2X1 + 3X2 <= 5\nBinary variables",
    "Maximise 4X1 + 5X2\nSubject to\nW: 2X1 + 3X2 <= 5\nInteger: X1, X2",
    "Minimise 4X1 + 5X2\nSubject to\nW: 2X1 + 3X2 >= 5\nNon-negativity and integer constraints",
    // Markdown table without "Subject to", no objective keyword row (falls back to text lines).
    "| Minimise | 3X1 + 4X2 |\n|---|---|\n| Demand | X1 + X2 >= 10 |\n| Ratio | X1 - 2X2 <= 0 |",
    // Exported two-column table without a "Subject to" row (Minimise objective row).
    "| Model component | Mathematical expression |\n|:---|---:|\n| Minimise objective | 3X1 + 4X2 = Z |\n| Demand | 1X1 + 1X2 >= 10 |\n| Ratio | 1X1 + -2X2 <= 0 |\n| Non-negativity: X1 | X1 >= 0 |\n| Non-negativity: X2 | X2 >= 0 |",
    // ASCII export of the website's own table, with Unicode comparisons.
    "+--------------------+--------------------+\n| Model component    | Mathematical expression |\n+--------------------+--------------------+\n| Maximise objective | 0.15X1 + 0.4X2 = Z |\n+--------------------+--------------------+\n| Budget             | 15X1 + 30X2 ≤ 200   |\n+--------------------+--------------------+\n| Diet               | 1X1 + 0X2 ≥ 2      |\n+--------------------+--------------------+\n| Non-negativity: X1 | X1 ≥ 0             |\n+--------------------+--------------------+",
    // Table with Subject to and row labels containing words; parentheses and decimals.
    "| Maximise total return | R = 0.08Stocks + 0.05Bonds |\n|---|---|\n| Subject to | |\n| Total funds | Stocks + Bonds <= 100000 |\n| Risk limit (stocks) | Stocks <= 0.6(Stocks + Bonds) |\n| Minimum bonds | Bonds >= 20000 |",
    // Dotted variable names and a coefficient-only objective with Z on the left.
    "Minimise Z = 2.5 prod.a + 1.25 prod.b\nSubject to\nNeed: prod.a + prod.b >= 8\nprod.a, prod.b >= 0",
    // Windows line endings and $ math wrappers.
    "Maximise $Z = 7X_1 + 5X_2$\r\nSubject to\r\nLabour: $3X_1 + 4X_2 \\leq 2400$\r\nMachine: $2X_1 + X_2 \\leq 1000$\r\n",
    // Overflowing coefficients (Infinity and Infinity - Infinity = NaN).
    "Maximise 2X1 + 3X2\nSubject to\nA: X1 + X2 <= 4\nB: " + "9".repeat(400) + "X1 - " + "9".repeat(400) + "X1 + X2 >= 0\nC: " + "9".repeat(400) + "X2 >= 0",
    // Only direction and constraints but no objective coefficients.
    "Maximise profit\nSubject to\nA: X1 + X2 <= 4\nB: X1 <= 3",
  ];
  const out = ["# Smart parser: further formulation formats.", step("open", "lpsolve")];
  for (const text of models) {
    out.push(step("click", "#resetLpBtn"), step("input", "#lpSmartInput", text), ...model(), step("read", "#lpParseStatus"));
    out.push(step("click", "#copyLpAsciiBtn"), step("clip"));
  }
  fs.writeFileSync(path.join(dir, "parser-extra.txt"), out.join("\n") + "\n");
}

// ---------------------------------------------------------------- sensitivity-extra.txt
{
  const out = ["# Sensitivity interpretation: wording branches, parser errors and sample sizes.", step("open", "lpsolve")];
  const prose = () => [step("read", "#sensitivityInterpretation"), step("disabled", "#copySensitivityProseBtn"), step("read", "#sensParseStatus"),
    step("val", "#sensSolution"), step("val", "#sensObjectiveValue"), step("val", "#sensDuals"), step("val", "#sensCoefFrom"),
    step("val", "#sensCoefTo"), step("val", "#sensRhsFrom"), step("val", "#sensRhsTo")];
  // Empty model: sample refuses, pasted output refuses.
  out.push(step("click", "#resetLpBtn"), step("click", "#loadSensitivitySampleBtn"), step("toast"), ...prose());
  out.push(step("input", "#sensRawOutput", "X1 X2\n1 2\n[1] 3\n[1] 0\n[1] 0 0 0\n[1] 0 0\n[1] 1 1"), ...prose());
  // One variable, one constraint (sample has only one solution value).
  out.push(step("input", "#lpVariables", "X1"), step("input", "#lpObjective", "5"), step("input", `${row(1)} .lp-name`, "Cap"),
    step("input", `${row(1)} .lp-coefficients`, "2"), step("input", `${row(1)} .lp-rhs`, "14"),
    step("click", "#loadSensitivitySampleBtn"), step("toast"), step("val", "#sensRawOutput"), ...prose());
  // Two variables, minimisation, >= and = constraints, a non-negativity row, NULL RHS ranges.
  out.push(step("change", "#lpDirection", "min"), step("input", "#lpVariables", "A, B"), step("input", "#lpObjective", "3, 2.5"),
    step("input", `${row(1)} .lp-name`, "Protein"), step("input", `${row(1)} .lp-coefficients`, "2, 1"), step("input", `${row(1)} .lp-rhs`, "16"),
    step("select", `${row(1)} .lp-operator`, "1"),
    step("input", `${row(2)} .lp-name`, "Blend"), step("input", `${row(2)} .lp-coefficients`, "1, 1"), step("input", `${row(2)} .lp-rhs`, "9.2"),
    step("select", `${row(2)} .lp-operator`, "2"),
    step("input", `${row(3)} .lp-name`, "Non-negativity A"), step("input", `${row(3)} .lp-coefficients`, "1, 0"), step("input", `${row(3)} .lp-rhs`, "0"),
    step("select", `${row(3)} .lp-operator`, "1"),
    step("click", "#loadSensitivitySampleBtn"), step("val", "#sensRawOutput"), ...prose());
  out.push(step("input", "#sensRawOutput", "> lp_sol$solution\n  A   B\n  6.8 2.4\n> lp_sol$objval\n[1] 26.4\n> lp_sol$status\n[1] 0\n> lp_sol$duals\n[1] 1.25 -0.0000001 0 0.5 -1e-10\n> lp_sol$sens.coef.from\n[1] 2.5 -1e+30\n> lp_sol$sens.coef.to\n[1] 1e+30 3\n> lp_sol$sens.rhs.from\nNULL\n> lp_sol$sens.rhs.to\nNULL"), ...prose());
  out.push(step("input", "#sensRawOutput", "## \"A\" \"B\"\n## 6.8 2.4\n## [1] 26.4\n## [1] 0\n## [1] 1.25 0 0 0.5 0\n## [1] 2.5 -1.5e+01\n## [1] 1e+30 3\n## [1] 9 8 -1e+30\n## [1] 1e+30 10 0"), ...prose());
  // Long vectors wrapped across lines with [n] indices.
  out.push(step("input", "#sensRawOutput", "A B\n6.8 2.4\n[1] 26.4\n[1] 0\n[1] 1.25 0\n[3] 0 0.5\n[5] 0\n[1] 2.5\n[2] 1\n[1] 4 3\n[1] 9 8\n[4] 1\n[1] 20 10 0"), ...prose());
  // Errors: missing blocks, wrong lengths, non-finite values.
  out.push(step("input", "#sensRawOutput", "A B\n6.8\n[1] 26.4\n[1] 0\n[1] 1 2 3 4 5\n[1] 1 2\n[1] 3 4"), ...prose());
  out.push(step("input", "#sensRawOutput", "A B\n6.8 2.4\n[1] 26.4\n[1] 0\n[1] 1 2 3 4\n[1] 1 2\n[1] 3 4"), ...prose());
  out.push(step("input", "#sensRawOutput", "A B\n6.8 2.4\n[1] 26.4\n[1] 0\n[1] 1 2 3 4 5\n[1] 1 2\n[1] 3 4\n[1] 1 2\n[1] 3 4 5"), ...prose());
  out.push(step("input", "#sensRawOutput", "A B\n6.8 2.4\n[1] 26.4\n[1] 0\n[1] 1 2 3 4 1e999\n[1] 1 2\n[1] 3 4"), ...prose());
  // Manual edits: no binding constraints, negative shadow price, rhs given for one side only.
  out.push(step("input", "#sensSolution", "10, 0"), ...prose());
  out.push(step("input", "#sensDuals", "-0.004, 0.006, 2, -3.5, 0"), ...prose());
  out.push(step("input", "#sensRhsFrom", "1, 2, 3"), ...prose());
  out.push(step("input", "#sensRhsTo", "4, 5, 6"), ...prose());
  out.push(step("input", "#sensRhsFrom", "1, 2"), ...prose());
  out.push(step("input", "#sensObjectiveValue", "abc"), ...prose());
  out.push(step("input", "#sensObjectiveValue", "1234.5678"), ...prose());
  out.push(step("change", "#lpDirection", "max"), ...prose());
  // Three binding resource constraints listed with a serial comma.
  out.push(step("click", "#loadLpSampleBtn"), step("input", "#sensSolution", "4, 4, 0.5"), step("input", "#sensObjectiveValue", "2.5"),
    step("input", "#sensDuals", "0, 0, 0, 0, 0, 0"), step("input", "#sensCoefFrom", "0.1, 0.2, 0.3"), step("input", "#sensCoefTo", "0.4, 0.5, 0.6"),
    step("input", "#sensRhsFrom", ""), step("input", "#sensRhsTo", ""), ...prose());
  out.push(step("input", `${row(1)} .lp-rhs`, "200"), step("input", `${row(2)} .lp-rhs`, "52"), step("input", `${row(3)} .lp-rhs`, "4"), ...prose());
  out.push(step("input", "#sensSolution", "4.004, 3.996, 0.5"), ...prose());
  out.push(step("input", "#sensSolution", "-0.001, 0, 0"), ...prose());
  out.push(step("count", "#sensitivityInterpretation [data-sensitivity-copy]"));
  fs.writeFileSync(path.join(dir, "sensitivity-extra.txt"), out.join("\n") + "\n");
}

// ---------------------------------------------------------------- detector.txt
{
  const out = ["# Universal detector routes lpSolve models into the operation.", step("open", "universal")];
  const after = () => [step("class", "#lpBench"), step("class", "#lpSensitivityAnalysis"), step("val", "#lpSmartInput"), ...model(),
    step("read", "#lpParseStatus"), step("read", "#sensitivityInterpretation"), ...snapshot()];
  out.push(step("click", "#loadLpSampleBtn"));
  out.push(step("open", "universal"), step("input", "#inputText", "Minimise 4X1 + 6X2\nSubject to\nProtein: 2X1 + 3X2 >= 12\nFat: X1 + 2X2 <= 10"), ...after());
  out.push(step("click", "#loadSensitivitySampleBtn"), ...completeAll());
  out.push(step("open", "universal"), step("input", "#inputText", "| Maximise objective | Z = 3X1 + 5X2 |\n|---|---|\n| Plant 1 | X1 <= 4 |\n| Plant 2 | 2X2 <= 12 |\n| Plant 3 | 3X1 + 2X2 <= 18 |"), ...after());
  fs.writeFileSync(path.join(dir, "detector.txt"), out.join("\n") + "\n");
}
