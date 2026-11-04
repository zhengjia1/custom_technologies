Attribute VB_Name = "testRender"
' Test-only: layout snapshots for test/layout-compare.js, timings and window screenshots.
Option Explicit

#If Mac Then
#Else
Private Declare PtrSafe Function GetFocusWnd Lib "user32" Alias "GetFocus" () As LongPtr
Private Declare PtrSafe Function PostMessageW Lib "user32" (ByVal hwnd As LongPtr, ByVal msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
#End If

Private Const BENCH_LIST As String = "standardBench summarizeBench shapeBench wideTableBench intervalBench oneSampleBench twoSampleBench anovaBench corrMultipleBench simpleRegressionBench advancedRegressionBench stepwiseBench kfoldBench pcaBench kmeansBench lpBench ggplotBench"


' Loads the rule table, the fonts (base64 files font-400.b64... in fontDir) and the body.
Public Sub Prepare(ByVal cssPath As String, ByVal fontDir As String)
    Dim w As Variant, path As String, sep As String
    sep = Application.PathSeparator
    modRender.LoadStyleText ReadUtf8File(cssPath)
    If Len(fontDir) > 0 Then
        For Each w In Array("400", "600", "700")
            path = fontDir & sep & "font-" & w & ".b64"
            If Len(Dir$(path)) > 0 Then modRender.UseFontData CStr(w), ReadUtf8File(path)
        Next w
    End If
End Sub

Private Sub ShowBench(ByVal bench As String)
    Dim list As clsList, i As Long
    Set list = QA("main > section")
    For i = 0 To list.Count - 1
        list(i).ToggleClass "hidden", (list(i).Id <> bench)
    Next i
End Sub

' Snapshots of every bench in both themes at 1280/960/800/375 px, written to outPath.
Public Function RunLayoutSuite(ByVal cssPath As String, ByVal bodyPath As String, ByVal fontDir As String, ByVal outPath As String) As String
    Dim benches() As String, themes As Variant, widths As Variant, b As Long, t As Variant, w As Variant
    Dim out As String, t0 As Double, snap As String, count As Long, slowest As Double, dt As Double, slowKey As String
    On Error GoTo Failed
    Prepare cssPath, fontDir
    LoadDocument ReadUtf8File(bodyPath)
    benches = Split(BENCH_LIST, " ")
    themes = Array("light", "dark")
    widths = Array(1280, 960, 800, 375)
    For b = 0 To UBound(benches)
        ShowBench benches(b)
        For Each t In themes
            DocumentElement.SetAttribute "data-theme", CStr(t)
            For Each w In widths
                t0 = Timer
                snap = modRender.LayoutSnapshot(CDbl(w), 860)
                dt = (Timer - t0) * 1000
                If dt > slowest Then slowest = dt: slowKey = benches(b) & " " & t & " " & w
                out = out & "=== " & benches(b) & " " & t & " " & w & vbLf & snap & vbLf
                count = count + 1
            Next w
        Next t
    Next b
    WriteUtf8 outPath, out
    RunLayoutSuite = count & " snapshots; slowest " & Format$(slowest, "0") & " ms (" & slowKey & "); " & modStyle.StyleStats() & "; fonts " & modRender.FontNames() & vbLf & LastErrors
    Exit Function
Failed:
    RunLayoutSuite = "ERROR " & Err.Description & vbLf & LastErrors
End Function

' One snapshot (bench, theme, width) for quick checks.
Public Function OneSnapshot(ByVal cssPath As String, ByVal bodyPath As String, ByVal fontDir As String, ByVal key As String) As String
    Dim parts() As String
    On Error GoTo Failed
    Prepare cssPath, fontDir
    LoadDocument ReadUtf8File(bodyPath)
    parts = Split(key, " ")
    ShowBench parts(0)
    DocumentElement.SetAttribute "data-theme", parts(1)
    OneSnapshot = modRender.LayoutSnapshot(CDbl(parts(2)), 860) & vbLf & LastErrors
    Exit Function
Failed:
    OneSnapshot = "ERROR " & Err.Description & vbLf & LastErrors
End Function

Public Sub WriteUtf8(ByVal path As String, ByVal text As String)
    Dim bytes() As Byte, n As Long, i As Long, c As Long, handle As Integer, o As Long
    n = Len(text)
    ReDim bytes(0 To n * 3 + 1)
    For i = 1 To n
        c = AscW(Mid$(text, i, 1)) And &HFFFF&
        If c >= &HD800& And c <= &HDBFF& And i < n Then
            c = &H10000 + (c - &HD800&) * 1024 + ((AscW(Mid$(text, i + 1, 1)) And &HFFFF&) - &HDC00&)
            i = i + 1
        End If
        If c < &H80 Then
            bytes(o) = c: o = o + 1
        ElseIf c < &H800 Then
            bytes(o) = &HC0 Or (c \ 64): bytes(o + 1) = &H80 Or (c And 63): o = o + 2
        ElseIf c < &H10000 Then
            bytes(o) = &HE0 Or (c \ 4096): bytes(o + 1) = &H80 Or ((c \ 64) And 63): bytes(o + 2) = &H80 Or (c And 63): o = o + 3
        Else
            If o + 4 > UBound(bytes) Then ReDim Preserve bytes(0 To o + 8)
            bytes(o) = &HF0 Or (c \ 262144): bytes(o + 1) = &H80 Or ((c \ 4096) And 63)
            bytes(o + 2) = &H80 Or ((c \ 64) And 63): bytes(o + 3) = &H80 Or (c And 63): o = o + 4
        End If
    Next i
    If Len(Dir$(path)) > 0 Then Kill path
    handle = FreeFile
    Open path For Binary Access Write As #handle
    If o > 0 Then
        ReDim Preserve bytes(0 To o - 1)
        Put #handle, , bytes
    End If
    Close #handle
End Sub

Public Function DebugSteps(ByVal cssPath As String, ByVal bodyPath As String, ByVal fontDir As String) As String
    Dim stepName As String, out As String, lay As clsRenderLayout, pt As clsRenderPaint
    On Error GoTo Failed
    stepName = "prepare": Prepare cssPath, fontDir
    stepName = "load": LoadDocument ReadUtf8File(bodyPath)
    stepName = "bench": ShowBench "standardBench"
    stepName = "stylebegin": modStyle.StyleBegin 1280, 860
    stepName = "dump": out = out & modStyle.StyleDump(Q("main")) & vbLf
    stepName = "measure": out = out & "w(Hello)=" & modRender.MeasureText("Hello", 13, 400, False) & " lh=" & modRender.NaturalLineHeight(13, 400) & vbLf
    stepName = "layout": Set lay = New clsRenderLayout: lay.Run 1280, 860
    stepName = "paint": Set pt = New clsRenderPaint: Set pt.Engine = lay: pt.PathIds = True: pt.Paint
    out = out & "regions " & pt.Regions().Count
    DebugSteps = out
    Exit Function
Failed:
    DebugSteps = "ERROR at " & stepName & ": " & Err.Description & " (" & Err.Source & ")" & vbLf & out
End Function

' ---------------------------------------------------------------- the real window

' Boots the whole app the way modBoot.StartApp does (rules and fonts from files).
Private Sub BootApp(ByVal buildDir As String)
    Dim sep As String
    sep = Application.PathSeparator
    Prepare buildDir & sep & "css-rules.txt", buildDir & sep & "fonts"
    LoadDocument ReadUtf8File(buildDir & sep & "body.html")
    Application.Run "modApp.SetAppTestMode", True
    Application.Run "InitApp"
    Settle 700
End Sub

Private Sub OpenOperation(ByVal id As String)
    Dim el As clsNode
    If id = "universal" Then Set el = Q(".universal-entry") Else Set el = Q(".operation[data-id=""" & id & """]")
    If Not el Is Nothing Then el.Click
    Settle 700
End Sub

Private Sub SetTheme(ByVal theme As String)
    Dim cur As Variant
    cur = DocumentElement.GetAttribute("data-theme")
    If IsNull(cur) Then cur = "light"
    If CStr(cur) <> theme Then Q("#themeToggle").Click
    Settle 700
End Sub

Private Function Capture(ByVal scriptPath As String, ByVal outPath As String) As Boolean
    Dim h As LongPtr, t As Double, i As Long
    On Error Resume Next
    modRender.AppForm().Repaint
    On Error GoTo 0
    t = Timer
    Do While Timer - t < 0.8
        DoEvents
    Loop
    h = modRender.AppWindowHandle()
    If h = 0 Then Exit Function
    If Len(Dir$(outPath)) > 0 Then Kill outPath
    Shell "powershell -NoProfile -ExecutionPolicy Bypass -File """ & scriptPath & """ -Hwnd " & CStr(h) & " -Out """ & outPath & """", vbHide
    t = Timer
    Do While Len(Dir$(outPath)) = 0 And Timer - t < 30
        DoEvents
    Loop
    Capture = (Len(Dir$(outPath)) > 0)
End Function

' shots: "op theme width[;...]"; writes outDir\<op>-<theme>-<width>.png. Also reports
' render timings (first render, and keystrokes in the universal input).
Public Function Shots(ByVal buildDir As String, ByVal scriptPath As String, ByVal shotList As String, ByVal outDir As String) As String
    Dim specs() As String, i As Long, p() As String, out As String, file As String, sep As String, ctl As Object, k As Long
    Dim t0 As Double, total As Double, ok As Boolean
    On Error GoTo Failed
    sep = Application.PathSeparator
    BootApp buildDir
    t0 = Timer
    modRender.OpenAppWindow
    out = "open " & Format$((Timer - t0) * 1000, "0") & " ms (render " & Format$(modRender.LastRenderMs, "0") & " ms: layout " & Format$(modRender.LastLayoutMs, "0") & _
        ", paint " & Format$(modRender.LastPaintMs, "0") & ", controls " & Format$(modRender.LastApplyMs, "0") & ", " & modRender.LastControlsCreated & " controls created)" & vbLf
    ' Keystrokes in the universal input (each one regenerates the prose).
    Set ctl = modRender.ControlFor(Q("#inputText"))
    If Not ctl Is Nothing Then
        For k = 1 To 6
            t0 = Timer
            ctl.Text = ctl.Text & IIf(k Mod 2 = 1, "x", "")
            If k > 1 Then total = total + (Timer - t0) * 1000
            out = out & "key " & k & ": " & Format$((Timer - t0) * 1000, "0") & " ms (layout " & Format$(modRender.LastLayoutMs, "0") & ", paint " & _
                Format$(modRender.LastPaintMs, "0") & ", controls " & Format$(modRender.LastApplyMs, "0") & ", created " & modRender.LastControlsCreated & ", items " & modRender.LastItemCount & ")" & vbLf
        Next k
        ctl.Text = ""
    End If
    specs = Split(shotList, ";")
    For i = 0 To UBound(specs)
        p = Split(Trim$(specs(i)), " ")
        If UBound(p) >= 2 Then
            OpenOperation p(0)
            SetTheme p(1)
            ' Optional extra tokens: clicks on elements (e.g. a sample button) before the shot.
            For k = 3 To UBound(p)
                If Not Q(p(k)) Is Nothing Then Q(p(k)).Click
                Settle 700
            Next k
            WriteUtf8 outDir & sep & "shots-log.txt", out & LastErrors & "before render " & specs(i)
            modRender.SetClientSize CDbl(p(2)), 860
            modRender.ResizeNow
            modRender.RenderNow
            t0 = Timer
            modRender.RenderNow
            file = outDir & sep & p(0) & "-" & p(1) & "-" & p(2) & IIf(UBound(p) >= 3, "-sample", "") & ".png"
            WriteUtf8 outDir & sep & "shots-log.txt", out & LastErrors & "capturing " & file
            ok = Capture(scriptPath, file)
            out = out & modRender.DescribeWindow()
            WriteUtf8 outDir & sep & "shots-log.txt", out & LastErrors
            out = out & specs(i) & ": " & IIf(ok, "captured", "NOT captured") & " (render " & Format$(modRender.LastRenderMs, "0") & " ms, " & modRender.LastItemCount & " items)" & vbLf
        End If
    Next i
    modRender.CloseAppWindow
    Shots = out & LastErrors
    Exit Function
Failed:
    Shots = "ERROR " & Err.Description & vbLf & out & LastErrors
    On Error Resume Next
    modRender.CloseAppWindow
End Function
' The app window's child windows after opening an operation (diagnostics).
Public Function WindowTree(ByVal buildDir As String, ByVal op As String) As String
    Dim out As String
    On Error GoTo Failed
    BootApp buildDir
    modRender.OpenAppWindow
    If Len(op) > 0 Then OpenOperation op
    modRender.RenderNow
    out = modRender.DescribeChildWindows() & vbLf & modRender.DescribeWindow()
    modRender.CloseAppWindow
    WindowTree = out & LastErrors
    Exit Function
Failed:
    WindowTree = "ERROR " & Err.Description & vbLf & out & LastErrors
    On Error Resume Next
    modRender.CloseAppWindow
End Function

' ---------------------------------------------------------------- incremental layout

' Compares the incremental layout (reusing cached fragments) with a full layout after
' every step of a walk through the operations: open, sample, typing, theme, viewport.
' Returns the steps checked and the first differences.
Public Function VerifyIncremental(ByVal buildDir As String, ByVal ops As String) As String
    Dim out As String, ids As clsList, list As clsList, i As Long, id As String, el As clsNode, bench As clsNode
    Dim sample As clsNode, ta As clsNode, fails As Long, checks As Long, parts() As String, k As Long, sel As clsNode
    On Error GoTo Failed
    BootApp buildDir
    Set ids = New clsList
    If Len(Trim$(ops)) > 0 Then
        parts = Split(Trim$(ops), " ")
        For i = 0 To UBound(parts)
            ids.Add parts(i)
        Next i
    Else
        ids.Add "universal"
        Set list = QA(".operation[data-id]")
        For i = 0 To list.Count - 1
            ids.Add CStr(list(i).GetAttribute("data-id"))
        Next i
    End If
    out = out & CompareLayouts("boot", 1280, checks, fails)
    out = out & CompareLayouts("boot again", 1280, checks, fails)
    For i = 0 To ids.Count - 1
        id = CStr(ids(i))
        If id = "universal" Then Set el = Q(".universal-entry") Else Set el = Q(".operation[data-id=""" & id & """]")
        If Not el Is Nothing Then
            el.Click
            Settle 700
            out = out & CompareLayouts("open " & id, 1280, checks, fails)
            Set bench = ActiveBench()
            If id = "universal" Then
                Set sel = Q("#universalSampleSelect")
                sel.SelectedIndex = 1
                sel.DispatchEvent NewEvent("change", True)
                Settle 700
                out = out & CompareLayouts("sample " & id, 1280, checks, fails)
            ElseIf Not bench Is Nothing Then
                Set list = bench.QuerySelectorAll(".sample-btn")
                For k = 0 To list.Count - 1
                    If Not list(k).HasClass("hidden") Then
                        list(k).Click
                        Settle 700
                        out = out & CompareLayouts("sample " & id, 1280, checks, fails)
                        Exit For
                    End If
                Next k
            End If
            Set bench = ActiveBench()
            If Not bench Is Nothing Then
                Set ta = bench.QuerySelector("textarea")
                If Not ta Is Nothing Then
                    ta.Value = ta.Value & " 12"
                    ta.DispatchEvent NewEvent("input", True)
                    Settle 700
                    out = out & CompareLayouts("type " & id, 1280, checks, fails)
                End If
            End If
        End If
        If fails > 6 Then Exit For
    Next i
    Q("#themeToggle").Click
    Settle 700
    out = out & CompareLayouts("theme", 1280, checks, fails)
    out = out & CompareLayouts("narrow", 700, checks, fails)
    Q("#themeToggle").Click
    Settle 700
    out = out & CompareLayouts("theme back", 1280, checks, fails)
    VerifyIncremental = checks & " checks, " & fails & " differ" & vbLf & out & LastErrors
    Exit Function
Failed:
    VerifyIncremental = "ERROR " & Err.Description & vbLf & out & LastErrors
End Function

Private Function CompareLayouts(ByVal label As String, ByVal width As Double, ByRef checks As Long, ByRef fails As Long) As String
    Dim a As String, b As String, la() As String, lb() As String, i As Long, n As Long, stats As String, t0 As Double, ta As Double, tb As Double
    t0 = modRender.ClockMs()
    a = modRender.LayoutSnapshot(width, 860)
    ta = modRender.ClockMs() - t0
    stats = modRender.LayoutCacheStats()
    modRender.ResetLayoutCache
    t0 = modRender.ClockMs()
    b = modRender.LayoutSnapshot(width, 860)
    tb = modRender.ClockMs() - t0
    ' Leave the cache warm for the next step: lay out once more incrementally.
    checks = checks + 1
    If a = b Then
        CompareLayouts = "same  " & label & " (incremental " & Format$(ta, "0") & " ms: " & stats & "; full " & Format$(tb, "0") & " ms)" & vbLf
        Exit Function
    End If
    fails = fails + 1
    la = Split(a, vbLf): lb = Split(b, vbLf)
    CompareLayouts = "DIFF  " & label & " (" & (UBound(la) + 1) & " vs " & (UBound(lb) + 1) & " lines; " & stats & ")" & vbLf
    For i = 0 To UBound(la)
        If i > UBound(lb) Then Exit For
        If la(i) <> lb(i) Then
            CompareLayouts = CompareLayouts & "   inc  " & la(i) & vbLf & "   full " & lb(i) & vbLf
            n = n + 1
            If n >= 4 Then Exit For
        End If
    Next i
End Function

' ---------------------------------------------------------------- profiling

' One timing line for the last window action: total, DOM handler, Settle, render (layout /
' paint / controls), the time Windows then needs to repaint (flush), controls created /
' updated, items and text measurements.
Private Function TimingLine(ByVal label As String, ByVal total As Double) As String
    Dim flush As Double, t0 As Double
    t0 = modRender.ClockMs()
    DoEvents
    flush = modRender.ClockMs() - t0
    TimingLine = label & vbTab & Format$(total, "0") & vbTab & Format$(modRender.LastHandlerMs, "0") & vbTab & Format$(modRender.LastSettleMs, "0") & vbTab & _
        Format$(modRender.LastRenderMs, "0") & vbTab & Format$(modRender.LastLayoutMs, "0") & vbTab & Format$(modRender.LastPaintMs, "0") & vbTab & _
        Format$(modRender.LastApplyMs, "0") & vbTab & Format$(flush, "0") & vbTab & modRender.LastControlsCreated & vbTab & modRender.LastControlsUpdated & vbTab & _
        modRender.LastItemCount & vbTab & modRender.LastMeasureCount & "/" & Format$(modRender.LastMeasureMs, "0") & vbTab & modRender.LayoutCacheStats() & vbLf
End Function

Private Function ActiveBench() As clsNode
    Dim list As clsList, i As Long
    Set list = QA("main > section")
    For i = 0 To list.Count - 1
        If Not list(i).HasClass("hidden") Then Set ActiveBench = list(i): Exit Function
    Next i
End Function

Private Function VisibleSample(ByVal bench As clsNode) As clsNode
    Dim list As clsList, i As Long
    If bench Is Nothing Then Exit Function
    Set list = bench.QuerySelectorAll(".sample-btn")
    For i = 0 To list.Count - 1
        If Not list(i).HasClass("hidden") And Not modRender.ControlFor(list(i)) Is Nothing Then Set VisibleSample = list(i): Exit Function
    Next i
End Function

Private Function TypingTarget(ByVal bench As clsNode) As Object
    Dim list As clsList, i As Long, c As Object
    If bench Is Nothing Then Exit Function
    Set list = bench.QuerySelectorAll("textarea, input")
    For i = 0 To list.Count - 1
        Set c = modRender.ControlFor(list(i))
        If Not c Is Nothing Then
            If TypeName(c) = "TextBox" Then
                If c.Enabled And Not c.Locked Then Set TypingTarget = c: Exit Function
            End If
        End If
    Next i
End Function

' Times the real window: opening it, every operation (click its sidebar entry), its
' sample, keystrokes in its first text field, theme switches, resizes and scrolling.
' ops: space-separated operation ids ("" = all). Writes a tab-separated table to outPath.
Public Function Profile(ByVal buildDir As String, ByVal outPath As String, ByVal ops As String) As String
    Dim out As String, t0 As Double, ids As clsList, list As clsList, i As Long, id As String, el As clsNode, bench As clsNode
    Dim sample As clsNode, ctl As Object, k As Long, base As String, fr As Object, sep As String, keyTotal As Double, keyMax As Double, dt As Double
    Dim sel As Object
    On Error GoTo Failed
    sep = Application.PathSeparator
    t0 = modRender.ClockMs()
    BootApp buildDir
    out = "boot" & vbTab & Format$(modRender.ClockMs() - t0, "0") & vbLf
    out = out & "scenario" & vbTab & "total" & vbTab & "handler" & vbTab & "settle" & vbTab & "render" & vbTab & "layout" & vbTab & "paint" & vbTab & "apply" & vbTab & _
        "flush" & vbTab & "created" & vbTab & "updated" & vbTab & "items" & vbTab & "measured/ms" & vbLf
    t0 = modRender.ClockMs()
    modRender.OpenAppWindow
    out = out & TimingLine("open window", modRender.ClockMs() - t0)
    modRender.SetClientSize 1280, 860
    modRender.ResizeNow
    t0 = modRender.ClockMs()
    modRender.RenderNow
    out = out & TimingLine("re-render unchanged", modRender.ClockMs() - t0)
    Set ids = New clsList
    If Len(Trim$(ops)) > 0 Then
        Dim parts() As String
        parts = Split(Trim$(ops), " ")
        For i = 0 To UBound(parts)
            ids.Add parts(i)
        Next i
    Else
        ids.Add "universal"
        Set list = QA(".operation[data-id]")
        For i = 0 To list.Count - 1
            ids.Add CStr(list(i).GetAttribute("data-id"))
        Next i
    End If
    For i = 0 To ids.Count - 1
        id = CStr(ids(i))
        If id = "universal" Then Set el = Q(".universal-entry") Else Set el = Q(".operation[data-id=""" & id & """]")
        If el Is Nothing Then
            out = out & "open " & id & vbTab & "no entry" & vbLf
        Else
            t0 = modRender.ClockMs()
            If Not modRender.SimulateClick(el) Then
                out = out & "open " & id & vbTab & "no control" & vbLf
            Else
                out = out & TimingLine("open " & id, modRender.ClockMs() - t0)
                Set bench = ActiveBench()
                ' The sample (Universal Detector: its sample list).
                If id = "universal" Then
                    Set sel = modRender.ControlFor(Q("#universalSampleSelect"))
                    If Not sel Is Nothing Then
                        t0 = modRender.ClockMs()
                        sel.ListIndex = 1
                        out = out & TimingLine("sample " & id, modRender.ClockMs() - t0)
                    End If
                Else
                    Set sample = VisibleSample(bench)
                    If Not sample Is Nothing Then
                        t0 = modRender.ClockMs()
                        modRender.SimulateClick sample
                        out = out & TimingLine("sample " & id, modRender.ClockMs() - t0)
                    End If
                End If
                ' Keystrokes in the first text field.
                Set ctl = TypingTarget(ActiveBench())
                If Not ctl Is Nothing Then
                    base = ctl.Text
                    keyTotal = 0: keyMax = 0
                    For k = 1 To 4
                        t0 = modRender.ClockMs()
                        If k Mod 2 = 1 Then ctl.Text = base & "7" Else ctl.Text = base
                        dt = modRender.ClockMs() - t0
                        keyTotal = keyTotal + dt
                        If dt > keyMax Then keyMax = dt
                        If k = 1 Then out = out & TimingLine("key " & id, dt)
                    Next k
                    out = out & "keys " & id & vbTab & "avg " & Format$(keyTotal / 4, "0") & " max " & Format$(keyMax, "0") & vbLf
                End If
                ' Revisit the operation after another one (pool reuse).
            End If
        End If
        WriteUtf8 outPath, out & LastErrors
    Next i
    ' Theme switches.
    For k = 1 To 2
        t0 = modRender.ClockMs()
        modRender.SimulateClick Q("#themeToggle")
        out = out & TimingLine("theme " & k, modRender.ClockMs() - t0)
    Next k
    ' Resizes.
    Dim widths As Variant, w As Variant
    widths = Array(960, 800, 700, 1280)
    For Each w In widths
        modRender.SetClientSize CDbl(w), 860
        t0 = modRender.ClockMs()
        modRender.ResizeNow
        out = out & TimingLine("resize " & w, modRender.ClockMs() - t0)
    Next w
    ' Scrolling the page.
    Set fr = modRender.RegionFrame("main")
    If Not fr Is Nothing Then
        For k = 1 To 4
            t0 = modRender.ClockMs()
            fr.ScrollTop = IIf(k Mod 2 = 1, 400, 0)
            fr.Repaint
            DoEvents
            out = out & "scroll " & k & vbTab & Format$(modRender.ClockMs() - t0, "0") & vbLf
        Next k
    End If
    modRender.CloseAppWindow
    WriteUtf8 outPath, out & LastErrors
    Profile = out & LastErrors
    Exit Function
Failed:
    Profile = "ERROR " & Err.Description & vbLf & out & LastErrors
    On Error Resume Next
    WriteUtf8 outPath, Profile
    modRender.CloseAppWindow
End Function

Private Function Check(ByVal label As String,ByVal ok As Boolean, Optional ByVal detail As String = "") As String
    Check = IIf(ok, "PASS ", "FAIL ") & label & IIf(Len(detail) > 0, " (" & detail & ")", "") & vbLf
End Function

' Drives the window's event paths and checks the DOM reacts; captures the result.
Public Function Interactions(ByVal buildDir As String, ByVal scriptPath As String, ByVal outDir As String) As String
    Dim out As String, ctl As Object, cancel As Boolean, v As String, sep As String, t0 As Double, f As clsNode
    On Error GoTo Failed
    sep = Application.PathSeparator
    BootApp buildDir
    modRender.OpenAppWindow
    modRender.SetClientSize 1280, 860
    modRender.ResizeNow
    out = out & Check("innerWidth follows the window", modApp.WindowInnerWidth = 1280, CStr(modApp.WindowInnerWidth))
    ' A letter typed on a button starts a search.
    cancel = modRender.SimulateKey(Q("#pasteBtn"), 80, 0)
    Set f = modRender.FocusedNode()
    out = out & Check("letter outside a field types into search", Q("#searchInput").Value = "p" And cancel, "value '" & Q("#searchInput").Value & "', cancelled " & cancel)
    out = out & Check("search field has the window focus", f Is Q("#searchInput"), IIf(f Is Nothing, "none", "other"))
    ' Escape clears nothing here but must not fail; clear the search with the button.
    Q("#clearSearch").Click
    Settle 700
    modRender.RenderNow
    out = out & Check("clear search", Q("#searchInput").Value = "")
    ' Choosing a sample from the list loads it and generates prose.
    Set ctl = modRender.ControlFor(Q("#universalSampleSelect"))
    If Not ctl Is Nothing Then
        t0 = Timer
        ctl.ListIndex = 1
        out = out & Check("sample list loads a sample", Len(Q("#inputText").Value) > 0, Format$((Timer - t0) * 1000, "0") & " ms")
        out = out & Check("prose generated", InStr(1, Q("#outputBody").TextContent, "H") > 0, Left$(Replace(Q("#outputBody").TextContent, vbLf, " "), 60))
        out = out & Check("a timer is pending (toast)", modDom.NextTimerDelay() >= 0, CStr(modDom.NextTimerDelay()))
    Else
        out = out & Check("sample list control exists", False)
    End If
    out = out & Check("toast shown as an overlay", Q("#toast").HasClass("show") And InStr(1, modRender.DescribeWindow(), "ov:") > 0)
    Capture scriptPath, outDir & sep & "universal-sample-light-1280.png"
    ' Sidebar drag.
    modRender.SimulateResizerDrag 310, 400
    v = DocumentElement.GetStyle("--sidebar-width")
    out = out & Check("sidebar drag resizes the library", v = "400px", v)
    Capture scriptPath, outDir & sep & "universal-dragged-light-1280.png"
    ' Narrow window: the mobile layout.
    modRender.SetClientSize 700, 860
    modRender.ResizeNow
    out = out & Check("innerWidth after resize", modApp.WindowInnerWidth = 700, CStr(modApp.WindowInnerWidth))
    Capture scriptPath, outDir & sep & "universal-sample-light-700.png"
    modRender.SetClientSize 1280, 860
    modRender.ResizeNow
    ' K-means scaling warning (a fixed overlay modal).
    OpenOperation "kmeans"
    modRender.RenderNow
    Application.Run "modApp.ShowKmeansScaleAlert"
    Settle 700
    modRender.RenderNow
    Capture scriptPath, outDir & sep & "kmeans-modal-light-1280.png"
    out = out & Check("modal focuses Keep scaling", modRender.FocusedNode() Is Q("#keepKmeansScalingBtn"))
    cancel = modRender.SimulateKey(Q("#keepKmeansScalingBtn"), 27, 0)
    out = out & Check("Escape is swallowed by the decision modal", cancel And Not Q("#kmeansScaleAlert").HasClass("hidden"))
    cancel = modRender.SimulateKey(Q("#keepKmeansScalingBtn"), 9, 0)
    out = out & Check("Tab cycles to Disable scaling", cancel And modRender.FocusedNode() Is Q("#disableKmeansScalingBtn"))
    cancel = modRender.SimulateKey(Q("#keepKmeansScalingBtn"), 13, 0)
    out = out & Check("Enter on a button presses it", Q("#kmeansScaleAlert").HasClass("hidden") And Q("#kmeansScale").Checked)
    modRender.CloseAppWindow
    Interactions = out & LastErrors
    Exit Function
Failed:
    Interactions = "ERROR " & Err.Description & vbLf & out & LastErrors
    On Error Resume Next
    modRender.CloseAppWindow
End Function


' ---------------------------------------------------------------- visual checks (drawn controls)

Private Sub Pump(ByVal seconds As Double)
    Dim t As Double
    t = Timer
    Do While Timer - t < seconds
        DoEvents
    Loop
End Sub

' Posts a key press to the window that has the keyboard focus (Windows; no other window is
' involved). Returns False when Excel's thread has no focus window.
Private Function PostKey(ByVal vk As Long) As Boolean
    #If Mac Then
    #Else
        Dim h As LongPtr
        h = GetFocusWnd()
        If h = 0 Then Exit Function
        PostMessageW h, &H100, vk, 1
        PostMessageW h, &H101, vk, &HC0000001
        Pump 0.3
        PostKey = True
    #End If
End Function

Private Function NodeName(ByVal n As clsNode) As String
    If n Is Nothing Then NodeName = "(none)": Exit Function
    NodeName = n.LocalName
    If Len(n.Id) > 0 Then NodeName = NodeName & "#" & n.Id
    If Len(n.ClassName) > 0 Then NodeName = NodeName & "." & Replace(n.ClassName, " ", ".")
End Function

' Drawn buttons and switches, keyboard focus, sticky library button, fieldset legend,
' short pages: captures in outDir and PASS/FAIL lines.
Public Function VisualCheck(ByVal buildDir As String, ByVal scriptPath As String, ByVal outDir As String) As String
    Dim out As String, sep As String, n As clsNode, f As clsNode, fr As Object, ctl As Object, before As Boolean, ok As Boolean, i As Long
    Dim stage As String, sw As clsNode
    On Error GoTo Failed
    sep = Application.PathSeparator
    stage = "boot": BootApp buildDir
    modRender.OpenAppWindow
    modRender.SetClientSize 1280, 860
    modRender.ResizeNow
    Capture scriptPath, outDir & sep & "v-universal-light-1280.png"
    ' Buttons are drawn: their control is a zero-size proxy.
    Set ctl = modRender.ControlFor(Q("#pasteBtn"))
    out = out & Check("Paste button has a proxy control", Not ctl Is Nothing)
    If Not ctl Is Nothing Then out = out & Check("proxy has no area", ctl.Width = 0 And ctl.Height = 0, ctl.Width & "x" & ctl.Height)
    ' A click on the drawn button (its labels) presses it and focuses its proxy.
    stage = "click": ok = modRender.SimulateClick(Q("#clearSearch"))
    out = out & Check("click on drawn clear-search button", ok)
    ' Keyboard: focus the Paste button's proxy, then Tab with a posted key.
    stage = "keys"
    modRender.SimulateKey Q("#pasteBtn"), 16, 0
    modRender.RenderNow
    Set f = modRender.FocusedNode()
    out = out & Check("proxy takes the window focus", f Is Q("#pasteBtn"), NodeName(f))
    Capture scriptPath, outDir & sep & "v-focus-paste-1280.png"
    If PostKey(9) Then
        Set f = modRender.FocusedNode()
        out = out & Check("Tab moves to the next control", Not f Is Q("#pasteBtn"), NodeName(f) & " / dom " & NodeName(modDom.ActiveElement()))
        modRender.RenderNow
        Capture scriptPath, outDir & sep & "v-focus-tab-1280.png"
    Else
        out = out & "SKIP Tab key (no focus window)" & vbLf
    End If
    ' ANOVA: drawn collapse chevrons and the normality switch.
    stage = "anova": OpenOperation "anova-many"
    modRender.RenderNow
    Capture scriptPath, outDir & sep & "v-anova-light-1280.png"
    Set sw = Q("#anovaNormalityAssumed")
    before = sw.Checked
    modRender.SimulateClick sw
    out = out & Check("switch click toggles", sw.Checked <> before)
    ' Space on the focused switch proxy.
    modRender.SimulateKey sw, 16, 0
    modRender.RenderNow
    before = sw.Checked
    If PostKey(32) Then
        out = out & Check("Space toggles the focused switch", sw.Checked <> before, CStr(sw.Checked))
        modRender.RenderNow
        Capture scriptPath, outDir & sep & "v-anova-switch-focus-1280.png"
    Else
        out = out & "SKIP Space key (no focus window)" & vbLf
    End If
    If sw.Checked <> True Then modRender.SimulateClick sw
    ' Summarize: legend on the line, switches.
    stage = "summarize": OpenOperation "summarize"
    modRender.RenderNow
    Capture scriptPath, outDir & sep & "v-summarize-light-1280.png"
    ' Narrow: sticky library button.
    stage = "narrow": OpenOperation "anova-many"
    modRender.SetClientSize 375, 860
    modRender.ResizeNow
    modRender.RenderNow
    Capture scriptPath, outDir & sep & "v-anova-light-375-top.png"
    Set fr = modRender.RegionFrame("main")
    If Not fr Is Nothing Then
        fr.ScrollTop = 600
        Pump 0.3
        modRender.RenderNow
        Capture scriptPath, outDir & sep & "v-anova-light-375-scrolled.png"
        out = out & Check("scroll event seen", True, "top " & modRender.LastScrollEventTop & " dy " & modRender.LastScrollEventDy)
        fr.ScrollTop = 0
        Pump 0.3
    End If
    out = out & modRender.DescribeWindow()
    ' 800: library width; dark kmeans.
    stage = "800": modRender.SetClientSize 800, 860
    modRender.ResizeNow
    SetTheme "dark"
    OpenOperation "kmeans"
    modRender.RenderNow
    Capture scriptPath, outDir & sep & "v-kmeans-dark-800.png"
    out = out & modRender.DescribeWindow()
    ' Up/Down step a number field; Page Down scrolls the page.
    stage = "keys2"
    Set n = Q("#kmeansK")
    before = (n.Value = "3")
    modRender.SimulateKey n, 38, 0
    out = out & Check("ArrowUp steps a number field", n.Value = "4", n.Value)
    modRender.SimulateKey n, 40, 0
    out = out & Check("ArrowDown steps back", n.Value = "3", n.Value)
    Set fr = modRender.RegionFrame("main")
    fr.ScrollTop = 0
    modRender.SimulateKey Q("#loadKmeansSampleBtn"), 34, 0
    out = out & Check("Page Down scrolls the page", fr.ScrollTop > 0, CStr(fr.ScrollTop))
    modRender.SimulateKey Q("#loadKmeansSampleBtn"), 36, 0
    out = out & Check("Home scrolls back", fr.ScrollTop = 0, CStr(fr.ScrollTop))
    SetTheme "light"
    modRender.CloseAppWindow
    VisualCheck = out & LastErrors
    Exit Function
Failed:
    VisualCheck = "ERROR at " & stage & ": " & Err.Description & vbLf & out & LastErrors
    On Error Resume Next
    modRender.CloseAppWindow
End Function

' Every operation with its sample: captures at the top, scrolled half way and at the end of
' the page (blank regions, scrolling past the end). Reports each page's scroll range.
' spec: "theme width [op ...]" (no ops = every operation).
Public Function WalkCaptures(ByVal buildDir As String, ByVal scriptPath As String, ByVal outDir As String, ByVal spec As String) As String
    Dim out As String, sep As String, ids As clsList, list As clsList, i As Long, id As String, sample As clsNode, fr As Object, parts() As String
    Dim maxTop As Double, base As String, el As clsNode, theme As String, width As Double, ops As String, sp() As String
    On Error GoTo Failed
    sep = Application.PathSeparator
    sp = Split(Trim$(spec), " ")
    theme = sp(0): width = Val(sp(1))
    For i = 2 To UBound(sp)
        ops = ops & " " & sp(i)
    Next i
    BootApp buildDir
    modRender.OpenAppWindow
    modRender.SetClientSize width, 860
    modRender.ResizeNow
    SetTheme theme
    Set ids = New clsList
    If Len(Trim$(ops)) > 0 Then
        parts = Split(Trim$(ops), " ")
        For i = 0 To UBound(parts)
            ids.Add parts(i)
        Next i
    Else
        ids.Add "universal"
        Set list = QA(".operation[data-id]")
        For i = 0 To list.Count - 1
            ids.Add CStr(list(i).GetAttribute("data-id"))
        Next i
    End If
    For i = 0 To ids.Count - 1
        id = CStr(ids(i))
        If id = "universal" Then Set el = Q(".universal-entry") Else Set el = Q(".operation[data-id=""" & id & """]")
        If Not el Is Nothing Then
            modRender.SimulateClick el
            If id = "universal" Then
                Set fr = modRender.ControlFor(Q("#universalSampleSelect"))
                If Not fr Is Nothing Then fr.ListIndex = 1
            Else
                Set sample = VisibleSample(ActiveBench())
                If Not sample Is Nothing Then modRender.SimulateClick sample
            End If
            modRender.RenderNow
            base = outDir & sep & id & "-" & theme & "-" & CStr(width)
            Set fr = modRender.RegionFrame("main")
            If Not fr Is Nothing Then
                fr.ScrollTop = 0
                Pump 0.2
                Capture scriptPath, base & "-a.png"
                maxTop = fr.ScrollHeight - fr.InsideHeight
                out = out & id & vbTab & "scrollHeight " & Format$(fr.ScrollHeight / 0.75, "0") & " inside " & Format$(fr.InsideHeight / 0.75, "0") & " bars " & fr.ScrollBars & vbLf
                If maxTop > 1 Then
                    fr.ScrollTop = maxTop / 2
                    Pump 0.2
                    modRender.RenderNow
                    Capture scriptPath, base & "-b.png"
                    fr.ScrollTop = maxTop
                    Pump 0.2
                    Capture scriptPath, base & "-c.png"
                    fr.ScrollTop = 0
                End If
            End If
        End If
    Next i
    SetTheme "light"
    modRender.CloseAppWindow
    WalkCaptures = out & LastErrors
    Exit Function
Failed:
    WalkCaptures = "ERROR " & Err.Description & vbLf & out & LastErrors
    On Error Resume Next
    modRender.CloseAppWindow
End Function
