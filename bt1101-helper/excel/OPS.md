# Porting the website's operations

Read `PORTING.md` (rules, toolchain, core libraries) and `FRAMEWORK.md` (DOM) first.

## Ownership

The website's main script (`excel/.build/app-src.js`, regenerate with
`node test/extract-app.js`) is split by line range. Each range is ported into one VBA
module; every top-level JS function keeps its name in PascalCase (`renderSummaryCode` →
`RenderSummaryCode`), top-level `const`/`let` state becomes module-level variables.

| Module | Website lines (app-src.js) | Covers |
|---|---|---|
| `modShared` | (done) | formatting, p-values, R names, axis helpers, completion, toasts |
| `modVariants` (+ `modVariants2` R-code engine, `modVariants3` prose pools) | 71–962 | answer options engine (prose/code variants, R-code variation); `VisibleOperationBench` is public |
| `modApp` | 1–70, 963–1590, 3889–4258, 4743–end | shell: groups/samples, update check, alerts, setActive, detector, Shapiro/Pearson/universal parsers, generate, sidebar, search, theme, number inputs |
| `opStats` | 1591–1695, 3782–3888, 4704–4742 | summarize, skewness & kurtosis, wide table |
| `opInference` | 1696–1863, 3453–3781, 4563–4703 | one-sample, two-sample, intervals (incl. distribution functions) |
| `opAnova` | 1864–2057, 4516–4562 | ANOVA workflow, multiple correlations code |
| `opRegression` | 2058–2540, 4494–4515 | regression code/interpretation, model comparison, GLM, coefficient intervals, confusion matrix |
| `opMining` | 2541–2764, 4468–4493 | stepwise, k-fold, PCA, k-means |
| `opLp` | 2765–3197, 4374–4467 | lpSolve model builder/parser, sensitivity analysis |
| `opPlot` | 3198–3452, 4259–4373 | plotting & tables code, plot-type combobox |

`modShared` already holds the cross-module helpers: `FormatFinalNumber`, `FormatPValue`,
`PValueDecision` (Variant True/False/Null), `PValueComparison`, `PValueDecisionError`,
`RColumnName`, `RString`, `ChartDisplayLabel`, `PredictorList`, the axis helpers
(`BoundedTickCode(stem, range, [headroom], [wholeNumbers])`, `CoordinateTickCode`,
`NumericAxisArguments`, `CoordinateAxisArguments`, …), `SetPanelCompletion`,
`SetLpPanelCollapsed`, `ShouldCompletePanelOnCopy`, `ShowToast(message, [duration])`,
`AppState` (clsMap: id/title/tutorial), `CalculatePlotAxisIntervals`, `CallOp`.
`modVariants` exposes `WriteAnswerText(text) As Boolean` (the website's
`writeAnswerText`; False = clipboard unavailable), `ClearCompletionForAnswerVariant`,
`RefreshAnswerVariantPreviews`, `RandomizeAnswerOptions`, `PickerForCopyButton`,
`ClearCodeVariantCache`, `ActiveAnswerCopyButton`. `modDom` provides `Q`, `QA`,
`EscapeHtml`, timers. A large module may be split into `opX`, `opX2`, … if it nears
VBA's limits; keep names unique across the project.

## Working alongside other agents

Several agents edit `excel/vba` at the same time, each owning its own files. Never edit a
file you don't own (`mod*`/`cls*` foundation files, other `op*` modules, other agents'
tests); if you need a change there, work around it locally and describe the change in your
report. Always compile/run with `--only` listing your files plus the foundation set:

```
--only modJs.bas,clsList.cls,clsMap.cls,modRegex.bas,clsRegex.cls,clsMatch.cls,modDom.bas,clsNode.cls,clsEvent.cls,clsDomHash.cls,clsDomSelector.cls,modShared.bas,modVariants.bas,modClipboard.bas,testUtil.bas,testDriver.bas,<your files> --no-form
```

`modVariants` may still be the stub (copy = canonical text) while its port is in
progress; keep calling its API. VBA reserved words can't be identifiers (`Currency`,
`Date`, `Open`, `Close`, `Error`, `Input`, `Get`, `Put`, `Print`, `Option`, `Type`, `Step`,
`Stop`, `Len`, `Mid`, `Name`…) — rename and keep going.

## Calls between modules

- Call `modShared`, `modVariants`, `modDom`, `modJs`, `modRegex`, `modClipboard` directly.
- Call another `op*` module or `modApp` through `CallOp "RenderPcaCode", arg1, …`
  (Application.Run by name), so each module compiles on its own.

## Startup and events

Each module has `Public Sub Init<Module>()` (e.g. `InitOpStats`) that resets its state
and performs the website's top-level statements for its range **in the website's
order**: `addEventListener` registrations (handler = a Public Sub in the module taking
`ByVal e As clsEvent`), initial renders, sample select setup, etc. `modApp` calls the Init
procedures in the order the website's code runs them.

## Translating JavaScript

- Every Public procedure the DOM invokes (event handlers, `SetTimeout`/microtask callbacks,
  the mutation hook) must trap its own errors: `On Error GoTo Failed` ... `Exit Sub` /
  `Failed: ReportError "Module.Proc"`. The website's `try/catch` blocks become `On Error`
  blocks too.
- `NewEvent("input")` does not bubble; use `NewEvent("input", True)` for the website's
  `new Event("input", { bubbles: true })`.

- Closures/arrow callbacks become loops or named procedures; state captured by a closure
  becomes module-level variables. `setTimeout(fn, ms)` → `SetTimeout "Module.Proc", ms[, arg]`.
- `async`/`await` (clipboard) becomes synchronous: `try { await writeAnswerText(t);
  showToast("Code copied") } catch { showToast("Clipboard access is unavailable") }` →
  `If WriteAnswerText(t) Then ShowToast "Code copied" Else ShowToast "Clipboard access is unavailable"`.
- `navigator.clipboard.readText()` → `ClipboardRead()`.
- Template literals and `innerHTML` assignments: build the same HTML string and assign
  `node.InnerHtml = html` — the DOM parses it exactly like the browser markup, so keep
  classes, ids and data attributes identical. Escape with `EscapeHtml`.
- `null`/`undefined` results → `Null`/`Nothing`/`Empty` consistently; document per function.
- Objects/arrays → `clsMap`/`clsList` (or small classes named `cls<Module><Thing>`).
- Regular expressions → `Rx(pattern, flags)` with identical patterns; `String.replace`
  with a callback → loop over `MatchAll` and rebuild the string.
- Number → string conversions must use `JsNumberToString`/`JsToFixed`/`JsToPrecision`
  /`JsToLocaleString` — never `CStr`, `Format` or `Str` (locale-dependent).
- Unicode text (H₀, ≤, —, “ ”) via `ChrW$` or the symbol constants in `modJs`.
- Keep every user-visible string exactly as the website prints it.

## Verification (required)

Use the shared step-script tools: write scripts under `excel/test/steps/<module>/` (format
in the header of `test/ref-run.js`: open/input/change/select/click/key/canon/read/val/
class/attr/checked/disabled/count/clip/toast) and run

```
node test/compare.js test/steps/opStats/*.txt [--boot InitApp] [--only ...]
```

It runs each script against the website in Node and against the VBA port (one Excel
session for all files) and prints differing records. `--boot` names the VBA procedure
that initialises the app before the steps (default `InitApp`, the shell's full startup);
until `modApp` exists, write your own boot procedure in `test<Module>.bas` that calls
`InitShared`, `InitVariants` and your `Init*`, and builds whatever the steps need (e.g.
the sidebar entries `open` clicks, or make `open` work by calling your render functions
the way the website's `setActive` does). Write step files with literal `\n` escapes
(don't let the shell expand them).

Under the hood:

1. Node: `pdf/test/harness.js` `bootNodeApp()` runs the website on minidom. For each case,
   open the operation (click its sidebar entry), set inputs (`value` + `input`/`change`
   events, clicks), call `browser.drain()`, then force every `.answer-variant-select` to
   `"1"` with a `change` event (Option 1 = canonical) and record the text of the output
   elements (`textContent`), relevant classes/visibility and clipboard writes.
2. VBA: a test function in `test<Module>.bas` loads `excel/.build/body.html`
   (`testUtil.ReadUtf8File`) with `LoadDocument`, runs `InitShared`, `InitVariants` and
   your `Init*`, replays the same steps through the DOM (set `Value`, `DispatchEvent
   NewEvent("input")`, `Click`, `Settle`) and returns the same records.
3. Compare; iterate to zero differences. Cover every sample, valid and invalid inputs,
   every option/toggle, and edge cases named in `../AGENTS.md` for your operations.

`node test/body.js` writes `excel/.build/body.html` (the website's body markup).
