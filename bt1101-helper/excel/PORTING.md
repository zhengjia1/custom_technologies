# BT1101 Prose Lab — Excel (pure VBA) edition: porting guide

The Excel edition is a hand port of `../index.html` to VBA so the same `.xlsm` works in
desktop Excel on **Windows and Mac**. There is no JavaScript engine: every parser,
generator and answer-variant rule is re-implemented in VBA and checked against the
website's own output.

## Rules for every VBA source (Mac compatibility)

- Only the VBA language, the Excel object model and MSForms. No `CreateObject`/`GetObject`
  (so no `Scripting.Dictionary`, `VBScript.RegExp`, `MSXML`, `ADODB`, `WScript`), no
  `SendKeys`, no file-system paths with `\`.
- `Declare` statements only inside `#If Mac Then ... #Else ... #End If` blocks, for
  Windows-only conveniences in the framework (clipboard, mouse wheel). Everything must
  also compile and work with those blocks absent.
- 32- and 64-bit safe. `Option Explicit` in every module.
- Source files are **plain ASCII** (the toolchain rejects anything else). Build other
  characters with `ChrW$`, e.g. `"H" & ChrW$(&H2080)` for H₀; modJs provides constants.
- Keep each procedure well under VBA's 64 KB compiled limit (aim < 400 lines), at most
  24 line continuations per statement, physical lines < 1000 characters.
- Strings are UTF-16 in both languages, so indices line up. Port JS 0-based string
  arithmetic with the 0-based helpers in modJs (`JsSlice`, `JsIndexOf`, …) rather than
  re-deriving 1-based `Mid$` offsets by hand.
- No randomness except through `modJs.JsRandom` (so tests can seed it).

## File layout

- `vba/*.bas`, `vba/*.cls` — modules and classes (imported).
- `vba/frmApp.code.vba` — the app window's code (the form is created by the build).
- `vba/ThisWorkbook.code.vba` — workbook events.
- Module prefixes: `modRegex`/`clsRegex`/`clsMatch` (regex), `modJs`, `clsList`, `clsMap`
  (JS semantics), `modUi`/`clsUi`/`clsNode` (framework), `modVariants` (answer options),
  `op*` (one module per website operation group), `test*` (test-only, excluded from
  the release build).

## Toolchain (Windows, Excel installed)

```
node tools/vba.js check [--only a.bas,b.cls] [--no-form]
node tools/vba.js run <Module.Function> [args...] [--argfile file] [--only ...] [--no-form]
```

`check` builds a scratch workbook from the sources and compiles it, printing the exact
module, line and message of any compile error. `run` also calls a public function (up to
four String arguments) and prints its String result; run-time error dialogs are captured
and printed. Each call takes ~20 s, so batch many test cases into one function call that
returns a report. Several agents may run the toolchain at once.

Reference outputs: `pdf/test/harness.js` (`bootNodeApp()`) runs the real website on
minidom in Node — set input values, dispatch `input`/`change`/`click` events, drain, and
read the output elements. Use it to produce expected strings for differential tests.

## Core library API

### Regular expressions (`modRegex`, `clsRegex`, `clsMatch`)

JavaScript regex semantics for the features the website uses: literals, `.`, classes
(`[...]`, negated, ranges, `\d \D \s \S \w \W` inside and outside classes), escapes
(`\n \r \t \f \v \0 \xHH \uHHHH`, escaped metacharacters), anchors `^ $` (with `m`),
`\b \B`, groups (capturing, `(?:...)`), alternation, quantifiers `* + ? {n} {n,} {n,m}`
greedy and lazy, lookahead `(?=...) (?!...)`, flags `g i m u` (with `\p{L}` under `u`).
JS `\s` is the full ECMAScript whitespace set (includes U+00A0, U+FEFF, U+2028/9, …).

```vb
Set re = Rx(pattern, flags)            ' compiled once and cached by pattern+flags
re.Test(text) As Boolean
re.Exec(text, Optional start As Long = 0) As clsMatch   ' Nothing when no match; start is 0-based
re.MatchAll(text) As Collection        ' of clsMatch, like [...text.matchAll(re)] (non-overlapping)
re.Replace(text, replacement) As String ' JS String.prototype.replace: all matches with g, else first;
                                        ' $1..$99, $&, $$, $`, $' in the replacement
re.Split(text) As clsList              ' JS String.prototype.split(re) incl. captured groups
re.Global, re.IgnoreCase, re.Multiline  ' read-only flags
m.Index (0-based), m.Value, m.EndIndex (0-based, exclusive), m.GroupCount
m.Group(n) As String                    ' "" when the group did not participate
m.Matched(n) As Boolean                 ' False when the group did not participate (JS undefined)
m.Input
```

Convenience in `modRegex`: `RxTest(text, pattern, flags)`, `RxReplace(text, pattern,
replacement, flags)`, `RxMatch(text, pattern, flags) As clsMatch` (first match or Nothing),
`RxEscape(text)` (like the website's escapeRegExp).

### JavaScript semantics (`modJs`, `clsList`, `clsMap`)

Numbers are `Double`. NaN and ±Infinity are real IEEE values (`JsNaN()`,
`JsInfinity(sign)`), but **VBA raises error 6 when a NaN Double is compared** (`<`, `=`,
…) and errors where JS yields NaN/Infinity (0/0 → 6, x/0 → 11, `Sqr(-1)`/`Log(0)` → 5,
overflow → 6). Wherever a value may be NaN/Infinity or an operation may overflow or divide
by zero, use `JsIsNaN`/`JsIsFinite` first or the safe helpers `JsAdd/JsSub/JsMul/JsDiv/
JsMod/JsPow`, `JsSqrt/JsLog/JsLog10/JsExp/JsFloor/JsCeil/JsAbs…`, `JsMax/JsMin`, and the
comparisons `JsEq/JsNe/JsLt/JsLe/JsGt/JsGe`. Use `JsMod` for `%` (VBA `Mod` rounds to
integers). VBA evaluates expressions in x87 extended precision; where bit-exact agreement
with the website matters, store each intermediate step in a `Double` variable.
`JsPow(10, k)` reproduces V8's results exactly (the website computes `10 ** -decimals`).
Full API list: header of `vba/modJs.bas`.

```vb
JsNumberToString(x) As String     ' String(x) / x.toString(): shortest round-trip, JS exponent rules
JsToFixed(x, digits) As String    ' exact JS toFixed (binary value, ties away from zero on exact ties)
JsToPrecision(x, p) As String     ' exact JS toPrecision
JsToExponential(x, digits) As String
JsToLocaleString(x, minFrac, maxFrac, useGrouping) As String   ' en-US
JsNumber(text) As Double          ' Number(text) (NaN when invalid; "" -> 0)
JsParseFloat(text) As Double      ' parseFloat
JsParseInt(text, Optional radix) As Double
JsMathRound(x) As Double          ' Math.round (half toward +Infinity)
JsRandom() As Double              ' Math.random; JsSeedRandom seed for tests
JsTrim/JsTrimStart/JsTrimEnd(text) ' ECMAScript whitespace set
JsSlice(text, start, Optional end)  ' negative indices count from the end
JsSubstring(text, start, Optional end)
JsIndexOf(text, search, Optional from) / JsLastIndexOf  ' 0-based, -1 when absent
JsIncludes, JsStartsWith, JsEndsWith, JsPadStart, JsPadEnd, JsRepeat
JsCharAt(text, i), JsCharCodeAt(text, i)
JsSplit(text, separator, Optional limit) As clsList        ' string separator
JsLower(text), JsUpper(text)
```

`clsList` — a growable 0-based list of Variants (values or objects): `Add`/`Push`, `Pop`,
`Shift`, `Unshift`, `Item(i)` (default member, get/let/set), `Count`, `Slice(start, [end])`,
`Splice`, `IndexOf(value)`, `Includes`, `Join(separator)`, `Reverse`, `Concat(other)`,
`SortStrings`, `SortNumbers`, `SortBy(keys As clsList)` (stable), `ToArray()`, `Clone`,
`Clear`, `First`, `Last`; `For Each` support through `NewEnum`.
`clsMap` — insertion-ordered string-keyed map: `Exists`, `Item(key)` (default; get/let/set,
missing key reads as Empty), `Add`, `Remove`, `Keys`/`Values` (clsList), `Count`, `Clear`.
`NewList(ParamArray items)` and `NewMap()` in `modJs` construct them.

## Framework API

See `FRAMEWORK.md` (written alongside `clsUi`).
