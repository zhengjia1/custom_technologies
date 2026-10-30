# BT1101 Prose Lab — production instructions

These instructions replace every earlier `AGENTS.md` instruction for this workspace. Treat `index.html` as production code and preserve all behaviour below unless the user explicitly changes it.

## Product objective

Build a CyberChef-style website where students paste RStudio output or configure an analysis and receive accurate, context-specific prose or R code.

- The complete application must remain in one standalone `index.html` file.
- Keep all HTML, CSS, JavaScript, fonts, samples, and templates embedded in that file.
- All parsing and computation must run locally in the browser.
- Do not add network calls, CDNs, analytics, remote fonts, external APIs, build steps, or runtime dependencies.
- Preserve placeholder names such as `dataset`, `column1`, and `column2`.
- Accept ordinary R console prefixes and tolerate up to three added or missing spaces around expected tokens.
- Generated R code may assume the supplied data is correct and valid for the selected analysis. Omit defensive data-validation `if`/`stop` blocks. Retain conditions that perform the analysis or preserve valid-data behaviour, including classification, plot orientation, constant-axis handling, and undefined trend lines. Keep the website's setup validation and required warnings.

## Application structure

- Keep the CyberChef-style recipe library and workbench layout.
- Universal Detector must always be the first operation.
- Separate tutorials into collapsible recipe groups and substantial workflows into collapsible cards.
- Keep recipe search by test name, topic, and R function.
- Keep the resizable sidebar, keyboard controls, responsive mobile library, and locally persisted sidebar width.
- Preserve keyboard focus, accessible labels, live status/error regions, and responsive layouts.
- Use embedded Source Sans Pro throughout; do not introduce a separate code font.
- Generated prose should grow to fit its content.
- Auto-grow compact textareas to a sensible maximum, then allow scrolling.

Avoid decorative or duplicated interface content. Do not restore:

- the `100% LOCAL · NO UPLOADS` badge;
- the global rounding-conventions callout;
- route-match banners, character counts, auto-generation labels, or duplicated copy controls;
- yellow reminder panels unless essential guidance is required.

Use a plain **Use sample** button when an operation has one sample. Use a dropdown only when multiple samples are genuinely useful, such as Universal Detector.

## Shared completion and warning behaviour

- Ordinary completed cards may turn green, show **Done!**, and collapse after their required content is copied.
- Staged workflows must wait for every required copy action before completing.
- Editing information that affects an answer must clear the affected completion state.
- Use large dismissible modal warnings—not transient toasts—for decisions that can materially change an answer:
  - disabling K-means scaling;
  - selecting a known population standard deviation;
  - copying prose containing placeholder decision-variable names;
  - copying regression prose whose units may need editing.

## Answer permutations

- Every copyable prose block and R-code block must provide exactly six answer options.
- Option 1 must always be the current canonical answer, unchanged, so users can fall back to the most authoritative wording or code.
- Options 2–6 may rephrase prose or change semantically irrelevant R-code structure/formatting, but must not change statistical meaning, calculations, decisions, required warnings, or runtime behaviour.
- Make the six R-code options meaningfully different—not merely different in braces, spacing, or assignment operators. Plausible differences include freshman-appropriate variable names, operation order, intermediate steps, and equivalent base-R structures, provided the final output is identical and the generated code remains valid R.
- Generate R-code permutations locally in the user's browser from randomized pools; do not permanently map an option number to one fixed set of temporary variable names, plot colours, or point shapes.
- Keep Option 1 canonical. For Options 2–6, choose context-appropriate temporary names at generation time—for example, model-like names for fitted models, table-like names for tables, and data-like names for temporary data frames—and update every downstream reference to the chosen name.
- Never rename the user's supplied dataset names, column names, formula variables, function names, or named function arguments. In particular, a user data frame updated in place, such as `dataset <- dataset %>% ...`, must retain its original name because it must already exist before the generated code runs.
- Explicitly selecting any of Options 2–6 must generate a fresh local variation of the selected code/prose. Returning from Option 2 to Option 3 and then back to Option 2 must produce a different answer from the previously displayed Option 2 whenever equivalent alternatives are available. Remember each option's last displayed variation across intervening selections.
- Cache the displayed answer between explicit regeneration actions. Copying, repeated copying, incidental rendering, and returning to an edited operation must preserve its exact visible answer. Never randomize during copying. Option 1 always restores the current canonical answer.
- Do not generate permutations or an answer-option dropdown for tables, including the lpSolve ASCII model table. Table copying must always preserve the canonical table exactly.
- Choose a random option when an answer control is created and choose another random option when the user pastes input. Option 1 remains available in the dropdown even when another option is selected initially.
- Re-randomize the selected option and generate a fresh local code-variant set whenever the user revisits an untouched operation or selects a different plotting/table recipe. If the user has already entered or changed data in that operation, preserve its existing answer selection and exact generated code mapping when they leave and return.
- In a large workflow or “mega” output made of related child code/prose blocks, including lpSolve and staged ANOVA cards, use one shared answer-option control for the whole related output rather than separate controls for every child block.
- Apply answer options only to generated answers. Static instructional copy, such as **R code ready** directions, must remain fixed and must not receive a prose option control.
- Changing the dropdown must immediately update the visible prose/code. The copied text must match the visibly selected option and must not be transformed a second time during copying.
- Regenerating an answer after an input edit must derive all six options from the new canonical output; permutations must never accumulate or become nested through repeated edits.
- Changing an answer option affects the answer and must clear the relevant copied/completion state, including staged-workflow tracking.
- Preserve exact mandated endings across every prose option. In particular, a non-significant Pearson conclusion must still end exactly with: **There is no statistically significant correlation between the two variables.**

R argument ordering:

- Apply fresh argument-order shuffling throughout the shared R-code variation generator for Options 2–6, including eligible nested calls. On reselection, change the eligible argument order from that option's previous displayed order whenever another safe order exists.
- Shuffle only verified order-independent named arguments. Keep positional arguments in their original slots and preserve every argument's name, value, and associated expression. For example, `barplot(summary.table, beside = TRUE, horiz = outputValue)` may become `barplot(summary.table, horiz = outputValue, beside = TRUE)`.
- Use token-aware parsing and explicit function/argument policies; do not split R calls with comma-based regular expressions. Preserve strings, comments, formulas, namespaces, nested calls, and indexing expressions.
- Keep ordered vector/list entries, data-frame columns, sequential `mutate()`/`summarise()` definitions, and order-sensitive `par()` settings fixed. Do not reorder arguments whose evaluation can perform assignments, draw random numbers, draw plots, or invoke unknown functions. Unknown or ambiguous calls retain their argument order.

Further variation techniques:

- Prioritize independent statement ordering, alternative intermediate calculations, and multiple prose sentence structures when extending variation. These are implementation directions, not a claim that every recipe already supports every technique.
- Reorder setup statements only when they are independent. Every value must be assigned before use; preserve dependencies across related child blocks, random-number generation order, printed-output order, and plot-drawing order.
- Vary inline calculations versus named intermediate steps, or use alternative pipelines and equivalent R implementations, only after verifying the same values, missing-value handling, row/column ordering, labels, and downstream behaviour. Do not repeat a calculation with side effects or introduce extra console output.
- Vary line wrapping, indentation, and argument grouping as supplementary variation, without making generated code difficult for beginners to read.
- Build prose alternatives from different sentence structures, not only word substitutions. Combine or separate observations into sentences and paragraphs, and reorder independent observations where the explanation remains coherent.
- Numerical wording may vary only when it preserves the required precision and meaning. Do not replace a precise percentage with an approximate fraction, remove required statistics, alter strict comparisons, or change decisions, warnings, or mandated endings.
- Preserve the current canonical answer in Option 1 when adding permutation techniques. Explicitly requested correctness fixes must apply to the canonical generator and all options.

## Parsing and rounding

- Do not round intermediate calculations.
- For final answers greater than 1, normally show two decimal places.
- For final answers below 1, normally show three significant figures.
- Use natural rounding points where appropriate, such as two decimal places for currency.
- Descriptive-statistics tables may use two or three decimal places depending on scale.
- Statistical decisions must use the supplied significance level and provide hypotheses, decisions, and context-specific conclusions.
- Use one consistent p-value decision rule across detectors and all prose options: reject H₀ when `p < alpha`; otherwise fail to reject H₀. A non-significant result means insufficient evidence for the alternative hypothesis, not evidence proving the null hypothesis or insufficient evidence for the null hypothesis.
- For Shapiro–Wilk, the conclusion always evaluates the alternative that the variable is **not normally distributed**: sufficient evidence when significant, insufficient evidence when non-significant. Never turn the latter into “insufficient evidence ... is normally distributed.” Apply the same alternative-hypothesis wording discipline to one-/two-sample, variance, ANOVA, and regression conclusions.
- Respect printed p-value bounds. `p < alpha` establishes significance, but `p <= alpha` alone does not distinguish significance from equality. Bounds crossing the cutoff must request a more precise value rather than infer a decision. Preserve one-sided alternatives and the supplied alpha. Do not round a displayed p-value across the significance cutoff.
- Describe the observed numerical comparison exactly: use “greater than” when `p > alpha` and “less than” when `p < alpha`; do not replace a strict observed comparison with “greater than or equal to” or “less than or equal to.”

## Required operations

### Core tests

- Shapiro–Wilk
- Pearson correlation using `cor.test()`
- Multiple correlations using `psych::corr.test()`

Pearson significance and correlation description are separate decisions:

- If `p < 0.05`, report `r`, direction, and strength.
- If `p >= 0.05`, do not report `r`, direction, or strength in the conclusion. End exactly with: **There is no statistically significant correlation between the two variables.**

### Universal Detector

- Keep its trained-sample dropdown.
- Identify recognized R code/output locally and route it to the correct parser.
- Do not navigate away for self-contained parsers: Shapiro–Wilk, confidence/prediction intervals, Pearson correlation, one-sample mean/proportion, and two-sample mean/variance.
- Navigate directly to the relevant interpretation block for:
  - multi-variable regression output;
  - binary GLM;
  - confusion matrices;
  - PCA importance output;
  - K-means;
  - lpSolve.
- Regression output routes to **Multi-variable linear regression → Interpret regression output**.
- Confusion matrices route to their interpretation card.
- PCA importance output routes to **Variance explained by retained components**.
- Stepwise model selection is not a Universal Detector sample or destination.

### Tutorial 2 — Plotting and tables

Default requirements for every table/chart, in both axis modes and all answer options:

- Tables must have appropriate table-specific titles and meaningful column names. Frequency outputs identify the category or interval and the Frequency column.
- Charts must have appropriate titles, axis labels on meaningful axes, and legends wherever needed to identify groups or additional marks (including group-mean points). Bar values on barplots are optional.
- Pie charts must have appropriate titles, with every slice labelled by category name and percentage. Keep those labels tied to the correct slices across all colour/code variants.


Write plot setup as a sentence, adapting fields to the plot type:

**I would like to plot a [plot type] of [variable] against [variable], titled [title]. The dataset is called [dataset] (default is "dataset").**

- Derive a title from the plot type and variables when the title field is empty. A custom title overrides it.
- Format variable names in automatically generated chart titles, axis labels and legend titles by replacing underscores with spaces and capitalizing each word (e.g. `daily_screen_time_hours` → **Daily Screen Time Hours**). Apply this in both axis modes and all six options, including one-variable regression plots. Preserve explicit custom chart titles, actual dataset/column names, formulas, selectors and data category values. Perform display formatting in the browser, without adding R formatting code.
- Generated table captions must identify tables, never inherit chart titles or arrangement names such as “bar plot”, “contingency plot”, or “histogram”. Derive table-specific captions from the variables independently of the chart title in both axis modes and all six code options. Use “Frequency table”, “Table of mean …”, or “Descriptive statistics table” as appropriate. Keep explicit user-entered captions in the dedicated wide-table caption field.
- Use the categorized searchable plot combobox with headings:
  - Summaries
  - Categorical charts
  - Distributions
  - Relationships and trends
  - Other
- Keep an explicit searchable **Grouped barplot** entry under Summaries, plus side-by-side bar plot, stacked bar plot, mosaic plot, and contingency plot.
- Route side-by-side bar plot, stacked bar plot, and mosaic plot to the contingency workflow and preselect the correct arrangement.

Grouped barplot:

- Accept one numeric variable and one categorical grouping variable.
- Calculate each category's mean and standard deviation with missing values removed, then print the result as a two-decimal `knitr::kable()` table.
- Convert the table's `Mean` and `SD` columns to a matrix without transposing it. Use the category values as the matrix rows and legend entries, so each x-axis statistic displays the categories side by side.
- Label the two x-axis groups **Mean** and **Standard deviation**, add value labels above the bars, provide y-axis headroom, use bounded whole-number ticks, and restore temporary `par()` settings.

Plot-search keyboard behaviour:

- With one result, Tab commits it and advances naturally.
- With two or more results, Tab remains in search and highlights the first result.
- Up/Down moves the highlight and Enter commits it.
- Escape closes the menu.
- Apply the same multi-result behaviour to the unfiltered list.

Axis transformations:

- Keep a separate collapsible block.
- Independently support none, `log`, `log10`, `log2`, `sqrt`, and reciprocal (`1/x`) on applicable numeric axes.
- Disable categorical or unavailable axes.
- Preserve domain guidance for logarithms, square roots, and reciprocals.

Plot requirements:

- Keep scatterplot and line-plot code short: use one ordinary `plot()` call in basic mode; when interval calculation is enabled, add only coordinate assignments and `range(pretty(...))` limits. Let R's class-aware methods handle numeric, Date and POSIXt inputs without generated `if`/`else` dispatch. Do not add an unsolicited regression line or adaptive-margin code to these recipes. Numeric chart ranges use `pretty()` to handle constant values without a redundant constant-range `if`. Use fixed layout and pie-label size instead of conditional presentation branches. Preserve conditions genuinely required by other analyses; never add branches simply to vary an answer.

- Across every graph generator and all six answer options, including **Calculate axis intervals** on, never emit `xaxt`, `yaxt`, `xaxs`, `yaxs`, `xpd`, `args.legend`, or `inset`. Use ordinary axes with appropriate labels and `xlim`/`ylim` where needed. Do not add manual `axis()`/`Axis()` calls over automatic axes. Preserve Date/POSIXt coordinates for automatic calendar labelling. This overrides earlier detailed tick-rendering and clipping-parameter requirements; calculating limits must not restore those unsupported parameters.

- Provide a synchronized **Calculate axis intervals** toggle for Plotting & tables, one-variable regression, and K-means, off by default. Turning it off selects basic plotting code based on the user's `1 - Plotting.pdf`: use ordinary plot defaults and the reference's titles, labels, colours, arrangement, legends, bin/outlier settings, and simple limits where appropriate. Omit `par()`, `mar()`, `xaxt`, `yaxt`, `xaxs`, `yaxs`, explicit tick vectors, custom `Axis()`/`axis()` calls, adaptive margins/orientation, `xpd`, `args.legend`, `inset`, and advanced theme/axis-scale configuration. The explicit exclusion of `xaxt` overrides its presence in the PDF's line-plot example. Preserve analysis/table calculations, date/time classes when passing coordinates to automatic plotting, and necessary numeric conversion for models. Native specialized plots remain available. Detailed layout and tick-grid requirements below apply to toggle-on mode; basic mode follows the simpler reference examples. Toggling regenerates affected answers and clears their copied/completion state.

- In basic mode, frequency bars, pie charts, grouped/stacked contingency bars, histograms, scatterplots, and line plots follow `Data Visualisation 2 handout.pdf`. Use `count()`, `spread()`, `as.matrix()`, `paste()`, `brewer.pal()`, ordinary plotting calls and the parameters taught there. Histogram counts use `labels = TRUE`; do not append `text()` labels or an unsolicited fitted line to these basic recipes. Keep explicitly requested transformations and advanced-axis mode, and retain recipes from other tutorials. Colour permutations must reuse named colours or `brewer.pal()` without injecting `hcl.colors()` or replacing `scale_fill_brewer()` with `discrete_scale()`. Do not add namespaces solely for variation.

- When **Calculate axis intervals** is off, every attempt to copy generated chart code (including combined code containing charts) must show a large dismissible modal reminding the user to check that both axis ranges cover the entire dataset, plotted geometry and labels, including stacked totals. Keep copying the exact displayed code through the normal clipboard/completion flow; this is an informational warning, not an additional copy confirmation. Show it on each copy attempt, omit it when the toggle is on or the copied answer contains no chart, and include the small footnote **Press Escape to dismiss this warning.** Support Escape and an **I understand** button, keep keyboard focus inside the modal, and restore focus on dismissal.

- Every graph-producing recipe must use a visibly different colour scheme in each of its six R-code options. Choose the five non-canonical schemes locally and without replacement from suitable colour/palette pools so the schemes are distinct within histograms and independently within every other graph type.
- Audit both axes of every graph-producing recipe, including charts inside regression, PCA, and K-means workflows. Preserve appropriate automatic limits for specialized plotting methods; do not impose limits that exclude their plotted geometry.
- For every suitable graph, generate readable, evenly spaced ticks on each numeric axis. Derive a round tick interval from the actual plotted values or geometry after transformations, and put both limits on that same tick grid. End the upper axis at a round tick above the maximum. Add label headroom in whole tick intervals, not fixed percentages, and never append an irregular endpoint to an otherwise regular tick sequence. Include required zero baselines; use whole-number ticks for counts and cluster numbers, and appropriately scaled decimal ticks for continuous values. Keep explicitly controlled plot limits aligned with their endpoint ticks. Pie charts, mosaic plots, and other charts without meaningful numeric axes do not need numeric ticks.
- Include zero in bar-chart numeric ranges. Size stacked charts from the largest stacked total, not the largest individual segment; side-by-side charts use the individual bar values. Accommodate negative bars and leave space for value labels.
- Derive histogram x-axis limits from `h$breaks`, not just the observations, so the outer bins fit completely. Use count-based heights consistently with frequency labels and provide y-axis headroom.
- Expand constant-valued numeric ranges to a finite, non-zero span. Assume the selected axis has valid data rather than emitting a defensive no-finite-values check. A scatterplot with a constant predictor may still render, but must not attempt to draw a regression line with non-finite coefficients.
- Scatterplots, line plots, and regression scatterplots must support `Date` and `POSIXt` values on either axis. Use class-aware calendar ticks and labels, preserve the time zone, and use numeric coordinates consistently for plot limits and fitted lines. Never apply numeric `floor()`/division to date-time objects. Necessary class dispatch is not defensive data validation.
- Respect R plotting-method conventions: base `boxplot()`/`bxp()` uses `ylim` for the numeric data range even when horizontal, and `yaxs` controls expansion of that range. Keep categorical limits, numeric limits, and manually drawn ticks aligned with the actual orientation.
- Keep bars and histogram rectangles inside the plotting region, including when the user's R session has changed clipping settings. Preserve space for labels instead of using clipping to hide incorrectly bounded data.
- Frequency bars are always horizontal. Do not restore an orientation branch.
- Reserve left margin using the longest category label, label the categorical axis appropriately, preserve descriptors, and leave right-side headroom for count labels.
- Histogram code must calculate bins with `plot = FALSE`, plot with explicit y-axis headroom, and then add labels.
- Do not allow bar or histogram labels to be clipped.
- Preserve categorical-chart alignment safeguards, readable labels/margins, inset or resized legends, inset pie labels, and restoration of temporary base-R `par()` settings.
- Keep summarize-then-plot, frequency/bar, pie, contingency, grouped means, histogram, single/grouped boxplots, scatterplots, line plots, and Pareto analysis.

Pareto analysis:

- Accept a target contribution from 1–100, defaulting to 80.
- Require a non-negative value column.
- Find the smallest leading set of highest-value observations that reaches the target.
- Generate only the calculation and one console sentence:
  **X observations (Y% of the dataset) account for Z% of total [variable].**
- Do not generate a Pareto table or chart.

### Tutorial 4 — Statistical measures and probability distributions

**Summarize for stats**

- Use dynamically appended column-name fields.
- One populated column generates the single-column workflow; two or more generate the multi-column table.
- Show transition toasts in both directions.
- Only the dedicated visible column inputs may feed this code.
- Do not generate `detach()`, file-import templates, covariance placeholders, exports, or unrelated controls.
- By default, create `tab.summary` with columns in this order: `vars`, `n`, `mean`, `sd`, `median`, `skew`, `kurtosis`.
- Provide statistic toggles shared by single- and multi-column summaries: non-missing count, mean, SD, median, skewness, excess kurtosis, minimum, maximum, total rows, missing count, variance, Q1, Q3, IQR, and sum. Keep the six current statistics selected by default; always retain `vars`. Generate only selected statistics in a stable column order. Total rows includes missing values; `n` excludes them. Disable code copying if no statistic is selected. Changing selections must refresh every answer option and reset affected copied/completion state while preserving selections across operation visits and single/multi-column transitions.
- Explicitly print `tab.summary`.
- Render `knitr::kable()` with `row.names = FALSE`, `digits = 2`, and caption **Descriptive Statistics Table for [column]**.
- Apply the same field order and `kable()` formatting in multi-column mode.
- Construct multi-column descriptive summaries with one row per selected variable using `data.frame()` and explicit source-column references. Do not return multi-row vectors from `summarise()`, which requires one row per group; output names such as `n` and `mean` must never mask input columns.

Also keep:

- separate skewness interpretation;
- a kurtosis/excess-kurtosis toggle with mesokurtic, platykurtic, and leptokurtic explanations;
- the narrow-to-wide table generator and matching R code.

### Tutorial 5 — Statistical inference

Keep confidence intervals, prediction intervals, and sample-size calculations in one operation.

- Grey out inapplicable target, parameter-type, and standard-deviation combinations.
- Generate usable code once dataset, column/parameter, and confidence level are present; numeric summaries are optional.
- Parse returned R intervals into prose.
- Support:
  - mean confidence intervals with unknown population standard deviation;
  - mean confidence intervals with known population standard deviation;
  - proportion confidence intervals, including filtered conditions such as `Amount > 22`;
  - prediction intervals for a mean;
  - sample-size calculations for means and proportions.
- Confirm with a dismissible modal that “known standard deviation” means the population—not sample—standard deviation.
- Keep one-sample mean/proportion, two-sample mean/variance, and correlation-test parsers and prose.

### ANOVA workflow

Use four staged collapsible steps. Step 0 is the single source of truth for all code and sample output.

1. **Step 0 — Setup:** dataset, response, grouping variable, significance level, and compact normality/large-sample CLT toggle. Do not add a yellow explanation below it.
2. **Step 1 — Normality:** grouped Shapiro–Wilk code. Disable, grey, and collapse it when normality/CLT is assumed.
3. **Step 2 — Homogeneity of variance:** ask for group count. Use `var.test(dataset$response ~ dataset$group)` for two groups and `bartlett.test(response ~ group, data = dataset)` for more than two. Parse F-test/Bartlett output with hypotheses, decision, and prose. Route `p > alpha` to Step 3a and `p < alpha` to Step 3b while greying the irrelevant route.
4. **Step 3a — Standard ANOVA + Tukey HSD:** generate only `model <- aov(...)`, `summary(model)`, and `TukeyHSD(model)`. Parse degrees of freedom, F, p, and Tukey rows. State omnibus significance and significant adjusted pairwise comparisons, or state that Tukey found none.
5. **Step 3b — Welch ANOVA + Games–Howell:** generate only `dataset %>% welch_anova_test(...)` and `dataset %>% games_howell_test(...)`. Parse both outputs and describe significant pairwise differences.

Additional ANOVA requirements:

- Do not include explanatory `#` comments in generated ANOVA code.
- Substitute current Step 0 names into every sample output.
- Loading sample output must never overwrite Step 0.
- Keep textareas compact and remove excessive blank lines.
- Steps 2, 3a, and 3b require staged completion: copying code completes only the code panel; the card completes and collapses only after both code and prose are copied. Editing resets the affected state.

### Tutorial 6 — Regression and classification

- One-variable regression supports no transform, square root, inverse, natural log, and base-10 log.
- Multi-variable regression uses collapsible setup, generated-code, and interpretation cards.
- Parse each predictor’s estimate, t-value, and p-value.
- Describe coefficient effects while holding other predictors constant, followed by a separate population-parameter significance decision.
- Provide separate multiple R-squared and overall F-test blocks.
- Do not allow pasted R output decorations such as **Show in New Window** to appear in the RStudio-output field or generated interpretation.
- Warn on regression-prose copy that “1 unit” may need replacement with a natural unit such as minutes or dollars.
- Parse pairwise/nested-model ANOVA F and p values and name predictors added by the full model.
- Keep dummy/reference groups, interactions, prediction, model significance, and binomial GLM workflows.
- Confusion-matrix interpretation must explain sensitivity, specificity, and accuracy with information hovers.
- Recompute metrics from the four cells and reject inconsistent reported values explicitly.

### Tutorial 7 — Data mining

- PCA uses small collapsible workflow blocks with individual copy controls and **Copy all**.
- Keep downstream regression/GLM component-significance modelling.
- Support individual/cumulative variance toggles, retained-component prose, loading-matrix parsing into readable PC equations, and downstream significance parsing.
- Title stepwise selection **Forward/Backward Stepwise Selection** and provide a direction toggle.
- Forward mode starts from the null model with scope.
- Backward mode starts from the full model while excluding named columns.
- Report adjusted R-squared.
- K-means keeps its scaling toggle.
- Disabling scaling requires a large warning explaining that variables should be directly comparable.
- When scaling is disabled, omit `X_scaled <- scale(X)` and use `X` in every downstream call.
- Keep K-fold cross-validation.

### Tutorial 8 — lpSolve and sensitivity analysis

Use one sidebar operation named **lpSolve / Sensitivity analysis**.

Model builder and parser:

- Accept Markdown, ASCII, and plain-text formulations.
- Accept multiline objectives, named constraints, parentheses, Unicode/LaTeX comparisons, and variable forms such as `X_1`, `X_{1}`, and `X1`.
- Accept exported two-column model tables with or without a “Subject to” row.
- For equations such as `1000X1 + 1200X2 = Z`, select the coefficient-bearing side as the objective.
- Do not interpret row labels such as “Contract for cars” as decision variables.
- Treat unit lower bounds such as `X1 >= 0` or `1X1 + 0X2 >= 0` as non-negativity declarations, not ordinary constraint rows.
- Do not duplicate explicit non-negativity declarations in `constraint_mat` or the model table. `lpSolve` already applies non-negativity to continuous variables.
- Detect binary/integer declarations.
- A complete parse may replace the manual builder; an incomplete parse must not overwrite it.
- Constraint rows append automatically while typing. Ignore the trailing blank row.
- Show only a remove control and provide **Reset all**. Do not add custom RHS steppers or an add button.
- Every populated constraint must contain exactly one numeric coefficient per decision variable. Require `0` for absent variables.
- Highlight invalid rows and refuse table/code generation or copying until valid.
- Put each constraint on its own readable R-code line.
- Avoid unnecessary `.00` suffixes in mathematical expressions.
- Hide sensitivity controls for all-integer and selected-binary models because `lpSolve` cannot combine them with `compute.sens`.
- Provide ASCII model-table copying.

Sensitivity interpretation:

- Interpret `lp_sol$duals` as genuine constraint shadow prices followed by one reduced cost per decision variable.
- Do not discuss shadow prices for non-negativity declarations.
- Split output into **Solution**, **Shadow Price**, **Reduced Cost**, and **Slack**, each with its own copy button.
- State every decision-variable quantity and the objective value in Solution prose.
- When a shadow price is zero, say the resource change “will not change the optimal objective value.” Do not append “by $0.00”.
- Determine binding status from slack, not from whether the shadow price is zero.
- Round displayed slack to two decimals before classifying it. Numerical residue such as `-0.0001` displays as `0` and is binding.
- Every resource paragraph must explicitly say binding or non-binding.
- A binding constraint is fully utilised; a non-binding constraint is not fully utilised.
- Copying Solution/prose must warn users to replace `X1`, `X2`, and similar placeholders with meaningful names.

Sensitivity completion behaviour:

- Copying one prose section must not complete or collapse the interpretation card.
- Mark only that section green and change its button to **Copied**.
- Collapse and show **Done!** only after all available prose sections have been copied.
- **Copy prose** counts as copying every section and may complete immediately.
- Editing sensitivity information resets section-copy tracking.
- Any optimization-model change must reset the entire Tutorial 8 workflow:
  - remove every **Done!** badge;
  - clear copied-state highlighting and paragraph tracking;
  - restore **Copy** buttons;
  - expand every Tutorial 8 card so regenerated information is visible.
- Apply that reset to manual settings, constraint edits/removals, reset/sample actions, and successful smart-parser imports.

## Release versioning

- Every change to `index.html` requires a new version number from the user before the edit is considered complete.
- Prompt the user for the new version number whenever an `index.html` change is requested and no new version was supplied.
- Use the exact format `ddmmyy_hhmm`, with an optional display-only leading `v` and optional display-only trailing `h`.
- Store the canonical version in JavaScript as `ddmmyy_hhmm` without the leading `v` or trailing `h`.
- Display it beneath **BT1101 Prose Lab** as `vddmmyy_hhmmh`.
- Put `<!-- BT1101 Prose Lab version: vddmmyy_hhmmh -->` on the first line of `index.html`.
- Update the first-line marker, displayed version, and JavaScript version together.
- Do not use a separate version API endpoint.
- When the page hostname is `zhengjia.dev` or one of its subdomains, automatically treat the page as current without making an update-check request.
- On every other hostname, fetch `https://zhengjia.dev/bt1101`, parse the version only from the fetched file's first-line marker, and reject files whose first line does not contain a valid marker.
- Parse versions as day, month, two-digit year, hour, and minute. Compare their resulting date-times chronologically; never determine recency from raw string ordering.
- Treat a remote version as an available update only when its parsed date-time is later than the local version. Equal or earlier remote versions mean the local file is current.
- Reject impossible dates or times as invalid versions rather than comparing them.

## Production verification

After every `index.html` edit:

1. Extract the embedded JavaScript and run `node --check`.
2. Confirm there are no duplicate static HTML IDs.
3. Confirm active JavaScript ID targets exist, including intentionally generated dynamic controls.
4. Exercise valid and invalid states for every affected parser or generator.
5. Exercise affected copy/completion/reset behaviour.
6. Remove temporary data, debug logging, highlighted development text, and obsolete UI branches.
7. Parse-check every generated R-code option affected by the change with R, not only JavaScript syntax, and fix any option that R cannot compile.
8. For permutation changes, exercise Option 2 → Option 3 → Option 2, repeated copying, canonical restoration, and shared workflow references. Verify argument bindings and ordered structures remain intact; compare executable variants' results with the canonical answer. Isolated argument-order changes must preserve rendered chart output.
9. For chart changes, test both orientations where supported, positive and negative values, constant values, extreme observations, long category labels, applicable transformations, and more than one plot size. Check both-axis geometry, histogram bin edges, label placement, bounded ticks, and restoration of temporary graphics settings across all six options.

Do not split production assets into extra runtime files while testing or implementing changes.
