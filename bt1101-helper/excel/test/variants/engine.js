// Loads the website's answer-variant engine (proseAnswerVariants, codeAnswerVariants and
// their R helpers) from excel/.build/app-src.js into a Node vm context, with Math.random
// replaced by the same mulberry32 generator modJs.JsRandom implements.
//
//   const engine = loadEngine(); engine.seed(7); engine.prose(text); engine.code(text, history)

const fs = require("fs");
const path = require("path");
const vm = require("vm");

function mulberry32(seed) {
  let state = seed >>> 0;
  return () => { state = (state + 0x6D2B79F5) >>> 0; let t = state; t = Math.imul(t ^ (t >>> 15), t | 1); t ^= t + Math.imul(t ^ (t >>> 7), t | 61); return ((t ^ (t >>> 14)) >>> 0) / 4294967296; };
}

function engineSource() {
  const file = path.join(__dirname, "..", "..", ".build", "app-src.js");
  const lines = fs.readFileSync(file, "utf8").split("\n");
  const start = lines.findIndex(line => /^\s*const answerVariantLabels = \{/.test(line));
  const end = lines.findIndex(line => /^\s*function cachedCodeAnswerVariants\(/.test(line));
  if (start < 0 || end < 0) throw new Error("Answer-variant engine not found in app-src.js");
  return lines.slice(start, end).join("\n");
}

function loadEngine() {
  const context = vm.createContext({ window: {}, visibleOperationBench: () => null, HTMLButtonElement: function HTMLButtonElement() {} });
  vm.runInContext(engineSource() + `
;globalThis.__engine = { proseAnswerVariants, codeAnswerVariants, rTokens, rStatements, protectedRNames, assignedRObjects,
  renameRObjects, storeRReports, varyRStructure, randomPlotStylePlans, changeRPlotStyle, shuffleRNamedArguments, freshRName };`, context);
  const engine = context.__engine;
  return {
    raw: engine,
    seed(seed) { context.__random = mulberry32(seed); vm.runInContext("Math.random = () => __random()", context); },
    prose(text, history) { return Array.from(engine.proseAnswerVariants(text, history)); },
    code(text, history) { return Array.from(engine.codeAnswerVariants(text, history || [])); },
    // regenerateSelectedAnswerVariant's per-cache step (website lines "Remember each slot's last wording").
    regenerate(kind, stored, option, previousOption) {
      const generate = kind === "code" ? engine.codeAnswerVariants : engine.proseAnswerVariants;
      const history = stored.history || [...stored.variants];
      const previous = history[option - 1], visible = stored.variants[previousOption - 1];
      const variants = Array.from(generate(stored.canonical, history));
      if (variants[option - 1] === previous || variants[option - 1] === visible) {
        const different = variants.findIndex((text, index) => index > 0 && text !== previous && text !== visible);
        if (different > 0) [variants[option - 1], variants[different]] = [variants[different], variants[option - 1]];
      }
      history[option - 1] = variants[option - 1];
      return { canonical: stored.canonical, variants, history: Array.from(history) };
    }
  };
}

module.exports = { loadEngine, mulberry32, engineSource };
