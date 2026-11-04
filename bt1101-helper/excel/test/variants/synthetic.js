// Synthetic texts for the engine-equivalence test: prose that triggers every phrase pool
// (alone and combined, with paragraphs and the protected Pearson ending) and R code that
// stresses the tokenizer, renaming, report storing, plot styles and argument shuffling.

const { engineSource } = require("./engine.js");
const vm = require("vm");

// One matching sentence per phrase pool, built from the pool's own pattern.
function poolSentences() {
  const context = vm.createContext({ window: {}, visibleOperationBench: () => null, HTMLButtonElement: function () {} });
  vm.runInContext(engineSource() + ";globalThis.__pools = prosePhrasePools;", context);
  return context.__pools.map(([pattern]) => {
    let s = pattern.source
      .replace(/^\^/, "").replace(/\\b/g, "")
      .replace(/\((variables\|variable)\)/, "variables")
      .replace(/\(\\d\+\)/, "30").replace(/\(PC\\d\+\)/, "PC1")
      .replace(/\(all \)\?/, "all ").replace(/\\s\*/g, " ")
      .replace(/\\\(/g, "(").replace(/\\\)/g, ")");
    if (/^PC1 = /.test(s)) s = s + "0.5 * column1 + 0.25 * column2";
    else if (/^H/.test(s)) s = s + "the mean is 5.";
    else s = s.charAt(0) === s.charAt(0).toLowerCase() ? "Overall, " + s + " here." : s + " the measured value.";
    return s;
  });
}

function proseItems() {
  const sentences = poolSentences();
  const items = sentences.map((text, i) => ({ kind: "prose", text, source: "synthetic pool " + i }));
  // Combined paragraphs (several pools, paragraph breaks, sentence starts for layout 1).
  for (let i = 0; i < sentences.length; i += 4) {
    const group = sentences.slice(i, i + 4);
    items.push({ kind: "prose", text: group.join(" ") + "\n\n" + group.slice(0, 2).reverse().join("\n\n\n"), source: "synthetic combined " + i });
  }
  const ending = "There is no statistically significant correlation between the two variables.";
  items.push({ kind: "prose", text: "A Pearson correlation test was conducted between column1 and column2. Since the p-value is 0.312, which is above 0.05, we fail to reject the null hypothesis. " + ending, source: "synthetic pearson ending" });
  items.push({ kind: "prose", text: ending, source: "synthetic ending only" });
  items.push({ kind: "prose", text: "No phrase from any pool appears here at all.", source: "synthetic no pool" });
  items.push({ kind: "prose", text: "   ", source: "synthetic blank" });
  items.push({ kind: "prose", text: "H₀: μ = 5\nH₁: μ ≠ 5\n\nThe test statistic is t = 2.1. Therefore, we reject the null hypothesis.  Wait. Next sentence! Another? Yes.", source: "synthetic unicode" });
  return items;
}

const code = [
  'library(dplyr)\nresult <- dataset %>%\n  group_by(column1) %>%\n  summarise(mean_value = mean(column2, na.rm = TRUE), sd_value = sd(column2, na.rm = TRUE))\nprint(result)',
  'model <- lm(column2 ~ column1, data = dataset)\nsummary(model)\nconfint(model, level = 0.95, parm = "column1")',
  'for (i in seq_len(nrow(dataset))) {\n  total <- total + dataset$column1[i]\n}\ntotal',
  'counts <- table(dataset$column1, useNA = "ifany", dnn = c("Group"))\nbarplot(counts, main = "Frequency", col = "steelblue", horiz = TRUE, las = 1, xlab = "Count")',
  'shapiro.test(dataset$column1)\nt.test(dataset$column1, mu = 5, alternative = "greater", conf.level = 0.95)',
  'x <- c(1, 2.5, .5, 1e-3, 2L, 0x10)\ny <- x %in% c(1, 2)\nz <- `my var` + 1 # comment with "quote"\nw <<- 3\n4 -> v\nv ->> u',
  's <- "string with \\"escape\\" and # hash"\nt <- \'single \\\' quote\'\nlabel <- glue::glue("Mean of {round(m, 2)} and {{literal}} {n}")\nm <- 1\nn <- 2\nmsg <- glue("Total {m + n} {{x}}")',
  'tab <- cbind(mean_value, sd_value)\nmean_value <- 1\nsd_value <- 2\nres <- cbind(mean_value, sd_value)\nres',
  'h <- hist(dataset$column1, breaks = "Sturges", plot = FALSE)\nplot(h, main = "Histogram", xlab = "Value", ylab = "Count", col = "grey", labels = TRUE, ylim = c(0, max(h$counts) * 1.2))',
  'pie(slices, labels = paste(names(slices), "\\n", round(100 * slices / sum(slices), 1), "%"), col = rainbow(length(slices)), main = "Share")',
  'pal <- RColorBrewer::brewer.pal(3, "Set2")\npal2 <- brewer.pal(n = 4, name = "Blues")\nggplot(dataset, aes(x = column1, fill = column2)) + geom_bar(position = "dodge", width = 0.7) + scale_fill_brewer(palette = "Set1", name = "Group")',
  'plot(dataset$column1, dataset$column2, pch = 19, col = "red", main = "Scatter", xlab = "X", ylab = "Y")\nabline(lm(column2 ~ column1, data = dataset), col = "blue", lwd = 2)',
  'km <- kmeans(X_scaled, centers = 3, nstart = 10, iter.max = 50)\nfviz_cluster(km, data = X_scaled)\nfviz_eig(pca_model)\nbiplot(pca_model)\npairs.panels(dataset[, 1:3], method = "pearson", hist.col = "orange")',
  'lp_sol <- lp(direction = "max", objective.in = c(3, 5), const.mat = matrix(c(1, 0, 0, 2), nrow = 2, byrow = TRUE), const.dir = c("<=", "<="), const.rhs = c(4, 12), compute.sens = TRUE)\nlp_sol$solution\nlp_sol$duals',
  'out <- mean(x, na.rm = TRUE, trim = 0.1)\nq <- quantile(x, probs = c(0.25, 0.75), na.rm = TRUE, names = FALSE)\nr <- round(out, digits = 2)\nstats::sd(x, na.rm = TRUE)\nbase::seq(from = 1, to = 10, by = 2)\nfoo::seq(from = 1, to = 10, by = 2)',
  'v <- seq(from = 0, to = max(x), length.out = 5) # spaced\nk <- kable(summary_table, digits = 2, caption = "Table", row.names = FALSE, align = "c")\nkable(summary_table,\n  # keep order\n  digits = 2, caption = "Table")',
  'f <- function(a, b = 2) {\n  a + b\n}\nf(1, b = 3)\nmatrix(data = f(1), nrow = 1, ncol = 1)\nmatrix(data = runif(4), nrow = 2, ncol = 2)\nmatrix(data = x[1], nrow = 1, ncol = 1)\nmatrix(data = (1), nrow = 1, ncol = 1)',
  'data.frame(a = 1, b = 2, stringsAsFactors = FALSE, check.names = TRUE)\ncut(x, breaks = 3, labels = FALSE, right = TRUE)\nprcomp(dataset, center = TRUE, scale. = TRUE)\nstep(model, direction = "both", trace = 0)',
  'anova(model1, model2)\nvar.test(dataset$column1 ~ dataset$column2)\nbartlett.test(column1 ~ column2, data = dataset)\ncor.test(dataset$column1, dataset$column2, method = "pearson", alternative = "two.sided")',
  'dataset <- dataset %>% mutate(ratio = column1 / column2)\nnew_data <- dataset %>% filter(ratio > 1)\nnew_data2 = new_data\nnrow(new_data2)',
  'x[["a"]] <- 1\nobj@slot\nf()(1)\ng(,1)\nh(a = 1, a = 2, b = 3)\nplot(x = 1, y = 2, col = )\n{ a <- 1; b <- 2 }\nif (a > 1) b else c',
  '',
  '   \n  ',
  'summary(lm(y ~ x, data = d))',
  'x <- 1',
  'print(t.test(x, mu = 0))'
];

function codeItems() {
  return code.map((text, i) => ({ kind: "code", text, source: "synthetic code " + i }));
}

function items() { return proseItems().concat(codeItems()); }

module.exports = { items };
if (require.main === module) items().forEach(item => console.log(item.kind, JSON.stringify(item.text).slice(0, 150)));
