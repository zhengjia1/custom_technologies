// Headless profile of the DOM side of every user action (vba/testPerf.bas), optionally
// with a per-phase breakdown. With --phases, the sources are copied to a private temp
// tree and instrumented there (wrappers around the hot procedures + a generated modProf
// module), so the real excel/vba sources are never touched.
//
//   node test/perf.js [--phases] [--passes N] [--out file] [--src folder]
//
// Each output row: ms, operation, action, element[, top phases "name ms/calls"].
// Phase times are inclusive (a phase nested in another counts in both); recursion into
// the same phase is counted once.

const fs = require("fs");
const os = require("os");
const path = require("path");
const { spawnSync } = require("child_process");

const ROOT = path.join(__dirname, "..");
const MODULES = ["modJs.bas", "clsList.cls", "clsMap.cls", "modRegex.bas", "clsRegex.cls", "clsMatch.cls", "modDom.bas",
  "clsNode.cls", "clsEvent.cls", "clsDomHash.cls", "clsDomSelector.cls", "modShared.bas", "modClipboard.bas", "testUtil.bas",
  "modVariants.bas", "modVariants2.bas", "modVariants3.bas", "modApp.bas", "modApp2.bas", "modApp3.bas", "opStats.bas",
  "opAnova.bas", "opInference.bas", "opInference2.bas", "opRegression.bas", "opRegression2.bas", "opMining.bas", "opPlot.bas",
  "opPlot2.bas", "opLp.bas", "opLp2.bas", "testPerf.bas"];

// [file, procedure kind, name, phase]: wrapped procedures.
const WRAP = [
  ["modDom.bas", "Function", "DomQuery", "query"],
  ["modDom.bas", "Function", "ParseFragment", "parseHtml"],
  ["modRegex.bas", "Function", "RxCompile", "rxCompile"],
  ["clsRegex.cls", "Function", "Test", "rxExec"],
  ["clsRegex.cls", "Function", "Exec", "rxExec"],
  ["clsRegex.cls", "Function", "MatchAll", "rxExec"],
  ["clsRegex.cls", "Function", "Replace", "rxExec"],
  ["clsRegex.cls", "Function", "Split", "rxExec"],
  ["modJs.bas", "Function", "JsNumberToString", "numFormat"],
  ["modJs.bas", "Function", "JsToFixed", "numFormat"],
  ["modJs.bas", "Function", "JsToPrecision", "numFormat"],
  ["modVariants2.bas", "Function", "VarCodeAnswerVariants", "codeVariants"],
  ["modVariants3.bas", "Function", "VarProseAnswerVariants", "proseVariants"],
  ["modVariants2.bas", "Sub", "TokenizeText", "cv.TokenizeText"],
  ["modVariants2.bas", "Function", "AssignedObjects", "cv.AssignedObjects"],
  ["modVariants2.bas", "Function", "VarFreshRName", "cv.VarFreshRName"],
  ["modVariants2.bas", "Function", "ReplaceSymbols", "cv.ReplaceSymbols"],
  ["modVariants2.bas", "Function", "ApplyEdits", "cv.ApplyEdits"],
  ["modVariants2.bas", "Sub", "ArgumentCalls", "cv.ArgumentCalls"],
  ["modVariants2.bas", "Function", "SafeArguments", "cv.SafeArguments"],
  ["modVariants2.bas", "Function", "ObjectNameType", "cv.ObjectNameType"],
  ["modVariants2.bas", "Sub", "Tokenize", "cv.Tokenize"],
  ["modVariants2.bas", "Sub", "AddIdentifiers", "cv.AddIdentifiers"],
  ["modVariants2.bas", "Function", "ProtectedRNames", "cv.ProtectedRNames"],
  ["modVariants2.bas", "Function", "RenameObjects", "cv.RenameObjects"],
  ["modVariants2.bas", "Function", "StoreReports", "cv.StoreReports"],
  ["modVariants2.bas", "Function", "VaryStructure", "cv.VaryStructure"],
  ["modVariants2.bas", "Function", "ChangePlotStyle", "cv.ChangePlotStyle"],
  ["modVariants2.bas", "Function", "ShuffleNamedArguments", "cv.ShuffleNamedArguments"],
  ["modVariants2.bas", "Function", "FallbackResult", "cv.FallbackResult"],
  ["modVariants2.bas", "Function", "FallbackIntermediate", "cv.FallbackIntermediate"],
  ["modVariants2.bas", "Function", "VarRStatements", "cv.VarRStatements"],
  ["modVariants2.bas", "Sub", "RandomPlotStylePlans", "cv.RandomPlotStylePlans"],
  ["modVariants.bas", "Sub", "ApplyAnswerVariantPreview", "applyPreview"],
  ["modVariants.bas", "Function", "GroupOf", "av.GroupOf"],
  ["modVariants.bas", "Function", "OwnerOf", "av.OwnerOf"],
  ["modVariants.bas", "Function", "PickerForCopyButton", "av.PickerForCopyButton"],
  ["modVariants.bas", "Function", "OutputNodes", "av.OutputNodes"],
  ["modVariants.bas", "Function", "KindOf", "av.KindOf"],
  ["modVariants.bas", "Function", "CachedVariants", "av.CachedVariants"],
  ["modVariants.bas", "Function", "Key", "av.Key"],
  ["modVariants.bas", "Sub", "GroupedEntries", "av.GroupedEntries"],
  ["modVariants.bas", "Function", "IsAnswerCopyButton", "av.IsAnswerCopyButton"],
  ["modVariants.bas", "Sub", "EnhanceAnswerVariantControls", "enhanceControls"],
  ["modVariants.bas", "Sub", "RandomizeAnswerOptions", "randomizeOptions"],
  ["modVariants.bas", "Sub", "PrepareAnswerOptionsForVisit", "prepareVisit"],
  ["modVariants.bas", "Function", "WriteAnswerText", "writeAnswer"],
  ["modApp2.bas", "Sub", "OnDomMutation", "mutationHook"],
  ["modDom.bas", "Sub", "DomStructBump", "structBump"],
  ["clsDomSelector.cls", "Sub", "Collect", "walk"],
  ["modDom.bas", "Sub", "BuildSnapshot", "snapBuild"],
  ["modDom.bas", "Sub", "CollectDocument", "collectDoc"],
  ["clsNode.cls", "Property Let", "InnerHtml", "innerHtmlSet"],
  ["clsNode.cls", "Property Let", "TextContent", "textSet"],
  ["clsNode.cls", "Property Get", "TextContent", "textGet"],
  ["clsNode.cls", "Function", "DispatchEvent", "dispatch"],
];

const PROF = `Attribute VB_Name = "modProf"
Option Explicit
#If Mac Then
#Else
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef counter As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef frequency As Currency) As Long
#End If
Private mFreq As Double
Private mNames(0 To 1023) As String
Private mTotal(0 To 1023) As Double
Private mCount(0 To 1023) As Long
Private mDepth(0 To 1023) As Long
Private mStart(0 To 1023) As Double
Private mGrand(0 To 1023) As Double
Private mGrandCount(0 To 1023) As Long
Private mN As Long
Private mIndex As clsDomHash

Private Function NowMs() As Double
#If Mac Then
    NowMs = Timer * 1000#
#Else
    Dim c As Currency, f As Currency
    If mFreq = 0 Then QueryPerformanceFrequency f: mFreq = CDbl(f)
    QueryPerformanceCounter c
    NowMs = CDbl(c) * 1000# / mFreq
#End If
End Function

Public Function ProfId(ByVal name As String) As Long
    Dim found As Boolean, v As Variant
    If mIndex Is Nothing Then Set mIndex = New clsDomHash
    v = mIndex.Lookup(name, found)
    If found Then ProfId = v: Exit Function
    If mN = 0 Then mN = 1
    If mN > 1023 Then ProfId = 1023: Exit Function
    mNames(mN) = name
    mIndex.Store name, mN
    ProfId = mN
    mN = mN + 1
End Function

Public Sub ProfEnter(ByVal id As Long)
    If mDepth(id) = 0 Then mStart(id) = NowMs()
    mDepth(id) = mDepth(id) + 1
End Sub

Public Sub ProfExit(ByVal id As Long)
    mDepth(id) = mDepth(id) - 1
    If mDepth(id) = 0 Then
        Dim d As Double
        d = NowMs() - mStart(id)
        mTotal(id) = mTotal(id) + d
        mCount(id) = mCount(id) + 1
        mGrand(id) = mGrand(id) + d
        mGrandCount(id) = mGrandCount(id) + 1
    End If
End Sub

Public Sub ProfReset()
    Dim i As Long
    For i = 0 To 1023
        mTotal(i) = 0: mCount(i) = 0: mDepth(i) = 0
    Next i
End Sub

' Whole-run totals (never reset), one per line, slowest first.
Public Function ProfGrand(Optional ByVal topN As Long = 60) As String
    Dim used(0 To 1023) As Boolean, k As Long, i As Long, best As Long, out As String
    For k = 1 To topN
        best = -1
        For i = 1 To mN - 1
            If Not used(i) And mGrand(i) > 0.05 Then
                If best < 0 Then
                    best = i
                ElseIf mGrand(i) > mGrand(best) Then
                    best = i
                End If
            End If
        Next i
        If best < 0 Then Exit For
        used(best) = True
        out = out & Format$(mGrand(best), "0.0") & vbTab & mGrandCount(best) & vbTab & mNames(best) & vbLf
    Next k
    ProfGrand = out
End Function

' Top phases since the last reset: "name ms/calls; ..."
Public Function ProfTake(Optional ByVal topN As Long = 8) As String
    Dim used(0 To 1023) As Boolean, k As Long, i As Long, best As Long, out As String
    For k = 1 To topN
        best = -1
        For i = 1 To mN - 1
            If Not used(i) And mTotal(i) > 0.05 Then
                If best < 0 Then
                    best = i
                ElseIf mTotal(i) > mTotal(best) Then
                    best = i
                End If
            End If
        Next i
        If best < 0 Then Exit For
        used(best) = True
        out = out & IIf(Len(out) > 0, "; ", "") & mNames(best) & " " & Format$(mTotal(best), "0.0") & "/" & mCount(best)
    Next k
    ProfTake = out
End Function
`;

function splitParams(text) {
  const out = []; let depth = 0, cur = "";
  for (const ch of text) {
    if (ch === "(") depth++;
    if (ch === ")") depth--;
    if (ch === "," && depth === 0) { out.push(cur); cur = ""; } else cur += ch;
  }
  if (cur.trim()) out.push(cur);
  return out.map(p => p.trim().replace(/^Optional\s+/i, "").replace(/^(ByVal|ByRef)\s+/i, "").replace(/^ParamArray\s+/i, "").split(/[\s(]/)[0]);
}

function wrap(source, kind, name, phase, file) {
  const lines = source.split(/\r?\n/);
  const header = new RegExp(`^(Public|Private|Friend)?\\s*${kind.replace(" ", "\\s+")}\\s+${name}\\s*\\((.*)\\)(\\s+As\\s+(\\w+))?\\s*$`);
  const at = lines.findIndex(line => header.test(line));
  if (at < 0) throw new Error(`perf.js: ${file} ${kind} ${name} not found`);
  const m = lines[at].match(header);
  const scope = m[1] || "Public", params = m[2], returnType = m[4];
  const endWord = kind.startsWith("Property") ? "Property" : kind;
  let end = at + 1;
  while (!new RegExp(`^End\\s+${endWord}\\b`).test(lines[end])) end++;
  // Original body becomes a private procedure (properties become Sub/Function).
  const innerKind = kind === "Property Let" ? "Sub" : kind === "Property Get" ? "Function" : kind;
  const inner = `${name}__${kind === "Property Let" ? "L" : kind === "Property Get" ? "G" : "P"}`;
  const body = lines.slice(at + 1, end).map(line => line
    .replace(new RegExp(`\\bExit\\s+${endWord}\\b`, "g"), `Exit ${innerKind}`)
    .replace(new RegExp(`(^|[^.\\w])${name}(\\s*=)`, "g"), `$1${inner}$2`)
    .replace(new RegExp(`(^|\\s)Set\\s+${name}(\\s*=)`, "g"), `$1Set ${inner}$2`));
  const isObject = returnType && !/^(String|Long|Double|Boolean|Integer|Variant|Byte|Single|Currency|Date)$/i.test(returnType);
  const names = splitParams(params).join(", ");
  const call = innerKind === "Sub" ? `${inner} ${names}` : `${isObject ? "Set " : ""}${name} = ${inner}(${names})`;
  const id = `modProf.ProfId("${phase}")`;
  const wrapper = [lines[at], `    Static profId__ As Long`, `    If profId__ = 0 Then profId__ = ${id}`, `    modProf.ProfEnter profId__`, `    ${call}`, `    modProf.ProfExit profId__`, `End ${endWord}`];
  const innerLines = [`Private ${innerKind} ${inner}(${params})${returnType ? " As " + returnType : ""}`, ...body, `End ${innerKind}`];
  if (returnType === "Variant" && innerKind === "Function") wrapper[4] = `    If IsObject(${inner}(${names})) Then Set ${name} = ${inner}(${names}) Else ${name} = ${inner}(${names})`;
  return [...lines.slice(0, at), ...wrapper, ...innerLines, ...lines.slice(end + 1)].join("\r\n");
}

function instrument(dir) {
  for (const [file, kind, name, phase] of WRAP) {
    const target = path.join(dir, "vba", file);
    fs.writeFileSync(target, wrap(fs.readFileSync(target, "utf8"), kind, name, phase, file));
  }
  {
    // Per-selector timing inside the DomQuery wrapper ("q doc|sub <selector>").
    const f = path.join(dir, "vba", "modDom.bas");
    let t = fs.readFileSync(f, "utf8");
    const a = (t.match(/    Set DomQuery = DomQuery__P\([^\r\n]*\)\r\n/) || ["<none>"])[0];
    if (!t.includes(a)) throw new Error("DomQuery wrapper");
    t = t.replace(a, "    Dim qid As Long: qid = modProf.ProfId(\"q \" & IIf(root.NodeType = 9, \"doc \", \"sub \") & selector)\r\n    modProf.ProfEnter qid\r\n" + a + "    modProf.ProfExit qid\r\n");
    fs.writeFileSync(f, t);
    // Tree walks per selector ("w <selector>"): cache misses and uncacheable selectors.
    const g = path.join(dir, "vba", "clsDomSelector.cls");
    let u = fs.readFileSync(g, "utf8");
    const b = "    Collect__P root, firstOnly, results\r\n";
    if (!u.includes(b)) throw new Error("Collect wrapper");
    u = u.replace(b, "    Dim wid As Long: wid = modProf.ProfId(\"w \" & mSource)\r\n    modProf.ProfEnter wid\r\n" + b + "    modProf.ProfExit wid\r\n");
    fs.writeFileSync(g, u);
  }
  // Per-handler timing: event listeners and timers/microtasks by procedure name.
  const domFile = path.join(dir, "vba", "modDom.bas");
  let dom = fs.readFileSync(domFile, "utf8");
  dom = dom.replace("    Application.Run handler, evt\r\n", "    Dim pid As Long: pid = modProf.ProfId(\"on \" & handler)\r\n    modProf.ProfEnter pid\r\n    Application.Run handler, evt\r\n    modProf.ProfExit pid\r\n");
  dom = dom.replace("    If hasArg Then Application.Run handler, arg Else Application.Run handler\r\n", "    Dim pid As Long: pid = modProf.ProfId(\"task \" & handler)\r\n    modProf.ProfEnter pid\r\n    If hasArg Then Application.Run handler, arg Else Application.Run handler\r\n    modProf.ProfExit pid\r\n");
  fs.writeFileSync(domFile, dom);
  fs.writeFileSync(path.join(dir, "vba", "modProf.bas"), PROF.replace(/\r?\n/g, "\r\n"));
}

function main() {
  const argv = process.argv.slice(2);
  const phases = argv.includes("--phases");
  const passes = argv.includes("--passes") ? argv[argv.indexOf("--passes") + 1] : "1";
  const outFile = argv.includes("--out") ? path.resolve(argv[argv.indexOf("--out") + 1]) : null;
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "bt1101-perf-"));
  try {
    fs.mkdirSync(path.join(dir, "vba"));
    fs.cpSync(path.join(ROOT, "tools"), path.join(dir, "tools"), { recursive: true });
    // --src <dir>: take the modules (except testPerf) from another source folder, e.g. an older copy.
    const src = argv.includes("--src") ? path.resolve(argv[argv.indexOf("--src") + 1]) : path.join(ROOT, "vba");
    for (const name of MODULES) fs.copyFileSync(path.join(name === "testPerf.bas" ? path.join(ROOT, "vba") : src, name), path.join(dir, "vba", name));
    const only = [...MODULES];
    if (phases) { instrument(dir); only.push("modProf.bas"); }
    const body = path.join(ROOT, ".build", "body.html");
    if (!fs.existsSync(body)) spawnSync(process.execPath, [path.join(__dirname, "body.js")], { stdio: "inherit" });
    const run = spawnSync(process.execPath, [path.join(dir, "tools", "vba.js"), "run", "testPerf.ProfileAll", body, passes, "--only", only.join(","), "--no-form"], { encoding: "utf8", maxBuffer: 64 * 1024 * 1024 });
    const text = (run.stdout || "") + (run.stderr || "");
    if (outFile) fs.writeFileSync(outFile, text); else process.stdout.write(text);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true, maxRetries: 20, retryDelay: 250 });
  }
}

main();
