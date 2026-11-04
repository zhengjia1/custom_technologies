Attribute VB_Name = "testModApp"
' Test-only checks for the application shell (modApp, modApp2, modApp3) that the step
' scripts cannot express: startup diagnostics, update-version parsing, number steppers
' with browser semantics, search scoring and storage.
Option Explicit

Private mOut As String
Private mFailures As Long
Private mCount As Long

Private Sub Check(ByVal name As String, ByVal actual As Variant, ByVal expected As Variant)
    mCount = mCount + 1
    If CStr(actual) <> CStr(expected) Then
        mFailures = mFailures + 1
        mOut = mOut & "FAIL " & name & ": got [" & CStr(actual) & "] expected [" & CStr(expected) & "]" & vbLf
    End If
End Sub

Public Function RunAll(ByVal bodyPath As String) As String
    Dim v As clsMap, el As clsNode
    mOut = "": mFailures = 0: mCount = 0
    On Error GoTo Failed
    JsSeedRandom 1
    LoadDocument ReadUtf8File(bodyPath)
    InitApp
    Settle 700
    ' Release versions.
    Set v = ParseReleaseVersion("v021126_2300h")
    Check "version canonical", v("canonical"), "021126_2300"
    Check "version timestamp", JsNumberToString(v("timestamp")), "1793660400000"
    Check "invalid date", ParseReleaseVersion("310226_1200") Is Nothing, True
    Check "leap day", ParseReleaseVersion("290224_0000") Is Nothing, False
    Check "non-leap day", ParseReleaseVersion("290225_0000") Is Nothing, True
    Check "hour 24", ParseReleaseVersion("010126_2400") Is Nothing, True
    Check "minute 60", ParseReleaseVersion("010126_2360") Is Nothing, True
    Check "month 13", ParseReleaseVersion("011326_1200") Is Nothing, True
    Check "bad format", ParseReleaseVersion("21126_2300") Is Nothing, True
    Check "chronological", ParseReleaseVersion("010127_0000")("timestamp") > ParseReleaseVersion("311226_2359")("timestamp"), True
    Check "first line marker", ReleaseVersionFromFirstLine("<!-- BT1101 Prose Lab version: v031126_0900h -->" & vbCrLf & "<!doctype html>")("canonical"), "031126_0900"
    Check "marker not first line", ReleaseVersionFromFirstLine(vbLf & "<!-- BT1101 Prose Lab version: v031126_0900h -->") Is Nothing, True
    Check "display", DisplayReleaseVersion("021126_2300"), "v021126_2300h"
    Check "app version valid", ParseReleaseVersion(APP_VERSION) Is Nothing, False
    ' Standalone update HTML: the rocket-loader type is stripped and CF scripts dropped.
    Check "standalone", PrepareStandaloneUpdateHtml("<!doctype html><script type=""abc123-text/javascript"">const groups = [];</script>" & _
        "<script src=""/cdn-cgi/scripts/x/cloudflare-static/rocket-loader.min.js""></script><script>window.__CF$cv$params={};/cdn-cgi/challenge-platform/</script><p>x</p>"), _
        "<!doctype html><script>const groups = [];</script><p>x</p>"
    ' Search scoring.
    Check "normalise", NormaliseSearchText("  K-Means  (Clustering)! "), "k means clustering"
    Check "levenshtein", LevenshteinDistance("kitten", "sitting"), 3
    Check "fuzzy exact", FuzzySearchScore("shapiro", "Shapiro-Wilk"), 1
    Check "fuzzy typo", JsNumberToString(FuzzySearchScore("shapro", "Shapiro Wilk")), "0.8571428571428572"
    Check "exact word", IsExactSearchMatch("glm", "GLM / binary variable"), True
    Check "not exact", IsExactSearchMatch("gl", "GLM / binary variable"), False
    ' Number steppers (browser stepUp semantics; step="any" uses the fallback).
    Set el = Q("#oneSampleAlpha")
    el.Value = "0.05"
    StepNumberInput el, 1
    Check "fallback up", el.Value, "0.06"
    StepNumberInput el, -1
    StepNumberInput el, -1
    Check "fallback down", el.Value, "0.04"
    el.Value = "0.5"
    StepNumberInput el, 1
    Check "fallback max", el.Value, "0.5"
    Set el = Q("#confidenceLevel")
    el.Value = "95"
    StepNumberInput el, 1
    Check "fallback integer", el.Value, "96"
    Set el = Q("#kfoldK")
    el.Value = "2"
    StepNumberInput el, -1
    Check "stepDown at min", el.Value, "2"
    el.Value = "2.5"
    StepNumberInput el, 1
    Check "stepUp off grid", el.Value, "3"
    el.Value = ""
    StepNumberInput el, 1
    Check "stepUp empty", el.Value, "2"
    Set el = Q("#plotThreshold")
    el.Value = "100"
    StepNumberInput el, 1
    Check "stepUp at max", el.Value, "100"
    el.Value = "0"
    StepNumberInput el, -1
    Check "stepDown below min (unchanged, like the browser)", el.Value, "0"
    ' Storage (test mode keeps it in memory).
    StorageSetItem "bt1101-test", "x"
    Check "storage", StorageGetItem("bt1101-test"), "x"
    StorageRemoveItem "bt1101-test"
    Check "storage removed", IsNull(StorageGetItem("bt1101-test")), True
    Check "theme stored", IsNull(StorageGetItem("bt1101-theme")), True
    Q("#themeToggle").Click
    Settle 700
    Check "theme toggled", StorageGetItem("bt1101-theme"), "dark"
    Check "theme attr", DocumentElement.DataGet("theme"), "dark"
    ' Sidebar width (window 1280: limits 210..480).
    Check "sidebar clamp", SetSidebarWidth(1000), 480
    Check "sidebar stored", StorageGetItem("bt1101-sidebar-width"), "480"
    Check "sidebar style", DocumentElement.GetStyle("--sidebar-width"), "480px"
    Check "sidebar default", SetSidebarWidth(0, False), 310
    Check "chart detection", ContainsRChart("x <- 1"), False
    Check "chart detection hist", ContainsRChart("h <- hist(x, plot = FALSE)"), True
    Check "chart detection ggplot", ContainsRChart("ggplot(d, aes(x)) + geom_bar()"), True
    Check "chart detection name only", ContainsRChart("plot <- 1" & vbLf & "print(plot)"), False
    Check "errors", LastErrors, ""
    RunAll = mCount & " checks, " & mFailures & " failures" & vbLf & mOut
    Exit Function
Failed:
    RunAll = "ERROR " & Err.Number & " " & Err.Description & vbLf & mOut & vbLf & LastErrors
End Function
