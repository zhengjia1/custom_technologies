Attribute VB_Name = "testE2E"
' Test-only: end-to-end tests through the real app window. Every action goes through the
' MSForms controls modRender draws (typing sets a TextBox's Text, picking sets a ComboBox's
' ListIndex, ticking flips a CheckBox, clicks/keys use modRender's event-path hooks), then
' the DOM state, the painted controls and the real Windows clipboard are checked.
'
'   RunSuite(buildDir, outPath, scenarios)
'     buildDir   folder with css-rules.txt, body.html and fonts\font-*.b64
'     outPath    UTF-8 result file, one record per line: STATUS<TAB>scenario<TAB>check<TAB>detail
'                (STATUS = PASS, FAIL or WARN; rewritten after every record)
'     scenarios  space-separated names (see RunScenario), "all" or "ops" (every operation)
'
' Run with test/e2e/run.js. The user's clipboard text is saved first and restored at the end.
Option Explicit

#If Mac Then
#Else
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal format As Long) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal memory As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal memory As LongPtr) As Long
Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal text As LongPtr) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal destination As LongPtr, ByVal source As LongPtr, ByVal bytes As LongPtr)
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)
Private Declare PtrSafe Function GetClipboardSequenceNumber Lib "user32" () As Long
#End If

Private Const SENTINEL As String = "BT1101-E2E-SENTINEL"
Private Const ALERT_IDS As String = "plotAxisRangeAlert variableRenameAlert regressionUnitsAlert kmeansScaleAlert knownPopulationSdAlert"

Private mOut As String
Private mOutPath As String
Private mScenario As String
Private mPass As Long, mFail As Long, mWarn As Long
Private mSavedClip As String
Private mEscapeTurn As Long
Private mStarted As Double
Private mMissingCell As String

' ================================================================ entry point

Public Function RunSuite(ByVal buildDir As String, ByVal outPath As String, ByVal scenarios As String) As String
    Dim names() As String, i As Long, k As Long, list As clsList
    On Error GoTo Failed
    mOutPath = outPath
    mOut = "": mPass = 0: mFail = 0: mWarn = 0
    mStarted = Timer
    mSavedClip = RawClipboardRead()
    Boot buildDir
    Set list = NewList()
    names = Split(Trim$(scenarios), " ")
    For i = 0 To UBound(names)
        Select Case LCase$(Trim$(names(i)))
            Case ""
            Case "all"
                list.Push "universal": list.Push "search": list.Push "letter": list.Push "theme": list.Push "resizer"
                list.Push "plotsearch": list.Push "stats": list.Push "axis": list.Push "kmeans": list.Push "knownsd"
                list.Push "lprows": list.Push "lpsens": list.Push "anova": list.Push "units": list.Push "revisit": list.Push "resize"
                list.Push "op:universal"
                For k = 0 To modApp.OpCount - 1
                    list.Push "op:" & modApp.OpIds(k)
                Next k
            Case "ops"
                list.Push "op:universal"
                For k = 0 To modApp.OpCount - 1
                    list.Push "op:" & modApp.OpIds(k)
                Next k
            Case Else
                list.Push Trim$(names(i))
        End Select
    Next i
    For i = 0 To list.Count - 1
        RunScenario CStr(list(i))
    Next i
    mScenario = "suite"
    Finish
    RunSuite = Summary()
    Exit Function
Failed:
    Rec "FAIL", "suite crashed", Err.Description
    On Error Resume Next
    Finish
    RunSuite = Summary()
End Function

Private Sub Finish()
    On Error Resume Next
    modRender.CloseAppWindow
    If Len(mSavedClip) > 0 Then modClipboard.ClipboardWrite mSavedClip
    Flush
End Sub

Private Function Summary() As String
    Summary = mPass & " passed, " & mFail & " failed, " & mWarn & " warnings (" & Format$(Timer - mStarted, "0") & " s); results in " & mOutPath
End Function

Private Sub Boot(ByVal buildDir As String)
    Dim sep As String, w As Variant, path As String
    sep = Application.PathSeparator
    modRender.LoadStyleText ReadUtf8File(buildDir & sep & "css-rules.txt")
    For Each w In Array("400", "600", "700")
        path = buildDir & sep & "fonts" & sep & "font-" & w & ".b64"
        If Len(Dir$(path)) > 0 Then modRender.UseFontData CStr(w), ReadUtf8File(path)
    Next w
    JsSeedRandom 11
    LoadDocument ReadUtf8File(buildDir & sep & "body.html")
    Application.Run "modApp.SetAppTestMode", True
    Application.Run "InitApp"
    Settle 700
    modRender.OpenAppWindow
    modRender.SetClientSize 1280, 860
    modRender.ResizeNow
    mScenario = "boot"
    Ok "app window opens", modRender.AppWindowOpen()
    Ok "window handle found", modRender.AppWindowHandle() <> 0
    Ok "innerWidth follows the window", modApp.WindowInnerWidth = 1280, CStr(modApp.WindowInnerWidth)
    CheckErrors
End Sub

Private Sub RunScenario(ByVal name As String)
    On Error GoTo Failed
    mScenario = name
    LastErrors = ""
    Select Case True
        Case Left$(name, 3) = "op:": OpSweep Mid$(name, 4)
        Case name = "universal": ScenarioUniversal
        Case name = "search": ScenarioSearch
        Case name = "letter": ScenarioLetter
        Case name = "theme": ScenarioTheme
        Case name = "resizer": ScenarioResizer
        Case name = "resize": ScenarioResize
        Case name = "plotsearch": ScenarioPlotSearch
        Case name = "stats": ScenarioStats
        Case name = "axis": ScenarioAxis
        Case name = "kmeans": ScenarioKmeans
        Case name = "knownsd": ScenarioKnownSd
        Case name = "lprows": ScenarioLpRows
        Case name = "lpsens": ScenarioLpSensitivity
        Case name = "anova": ScenarioAnova
        Case name = "units": ScenarioUnits
        Case name = "revisit": ScenarioRevisit
        Case Else: Rec "FAIL", "unknown scenario", name
    End Select
    mScenario = name
    CheckErrors
    CloseAlerts
    Exit Sub
Failed:
    Rec "FAIL", "scenario crashed", Err.Description & " (" & Err.Source & ")"
    CheckErrors
    On Error Resume Next
    CloseAlerts
End Sub

' ================================================================ recording

Private Sub Rec(ByVal status As String, ByVal check As String, Optional ByVal detail As String = "")
    Select Case status
        Case "PASS": mPass = mPass + 1
        Case "FAIL": mFail = mFail + 1
        Case Else: mWarn = mWarn + 1
    End Select
    mOut = mOut & status & vbTab & mScenario & vbTab & check & vbTab & OneLine(detail) & vbLf
    Flush
End Sub

Private Sub Flush()
    On Error Resume Next
    If Len(mOutPath) > 0 Then WriteUtf8File mOutPath, mOut
End Sub

Private Function OneLine(ByVal text As String) As String
    text = Replace(Replace(Replace(text, vbCrLf, "\n"), vbLf, "\n"), vbCr, "\r")
    text = Replace(text, vbTab, "\t")
    If Len(text) > 400 Then text = Left$(text, 400) & "..."
    OneLine = text
End Function

Private Sub Ok(ByVal check As String, ByVal condition As Boolean, Optional ByVal detail As String = "")
    Rec IIf(condition, "PASS", "FAIL"), check, detail
End Sub

Private Sub Soft(ByVal check As String, ByVal condition As Boolean, Optional ByVal detail As String = "")
    Rec IIf(condition, "PASS", "WARN"), check, detail
End Sub

Private Sub CheckErrors()
    If Len(LastErrors) > 0 Then
        Rec "FAIL", "handler errors reported", LastErrors
        LastErrors = ""
    End If
End Sub

' ================================================================ driving the controls

Private Function CtlOf(ByVal node As clsNode) As Object
    If node Is Nothing Then Exit Function
    Set CtlOf = modRender.ControlFor(node)
End Function

Private Function Drawn(ByVal node As clsNode) As Boolean
    Dim c As Object
    Set c = CtlOf(node)
    If c Is Nothing Then Exit Function
    Drawn = ShownControl(c)
End Function

' A control and every frame around it are visible.
Private Function ShownControl(ByVal c As Object) As Boolean
    Dim p As Object
    On Error GoTo Done
    If Not c.Visible Then Exit Function
    Set p = c.Parent
    Do While TypeName(p) = "Frame"
        If Not p.Visible Then Exit Function
        Set p = p.Parent
    Loop
    ShownControl = True
Done:
End Function

Private Sub FocusCtl(ByVal c As Object)
    If c Is Nothing Then Exit Sub
    On Error Resume Next
    c.SetFocus
    Err.Clear
End Sub

' A mouse click on the element's control (or the label/hit area that activates it).
Private Function Press(ByVal node As clsNode) As Boolean
    If node Is Nothing Then Exit Function
    FocusCtl CtlOf(node)
    Press = modRender.SimulateClick(node)
End Function

' Types text into the element's TextBox (the whole value at once, like a paste).
Private Function TypeText(ByVal node As clsNode, ByVal text As String) As Boolean
    Dim c As Object
    Set c = CtlOf(node)
    If c Is Nothing Then Exit Function
    If TypeName(c) <> "TextBox" Then Exit Function
    FocusCtl c
    If c.MultiLine Then text = Replace(Replace(text, vbCrLf, vbLf), vbLf, vbCrLf)
    c.Text = text
    TypeText = True
End Function

' Types text one character at a time (one Change event per keystroke).
Private Function TypeChars(ByVal node As clsNode, ByVal text As String) As Boolean
    Dim c As Object, i As Long
    For i = 1 To Len(text)
        Set c = CtlOf(node)
        If c Is Nothing Then Exit Function
        If TypeName(c) <> "TextBox" Then Exit Function
        FocusCtl c
        c.Text = c.Text & Mid$(text, i, 1)
    Next i
    TypeChars = True
End Function

Private Function Pick(ByVal node As clsNode, ByVal index As Long) As Boolean
    Dim c As Object
    Set c = CtlOf(node)
    If c Is Nothing Then Exit Function
    If TypeName(c) <> "ComboBox" Then Exit Function
    If index >= c.ListCount Then Exit Function
    FocusCtl c
    If c.ListIndex <> index Then c.ListIndex = index
    Pick = True
End Function

Private Function Tick(ByVal node As clsNode) As Boolean
    Dim c As Object
    Set c = CtlOf(node)
    If c Is Nothing Then Exit Function
    If TypeName(c) <> "CheckBox" Then Exit Function
    FocusCtl c
    c.Value = Not CBool(c.Value)
    Tick = True
End Function

Private Function KeyOn(ByVal node As clsNode, ByVal keyCode As Integer, Optional ByVal shift As Integer = 0) As Boolean
    KeyOn = modRender.SimulateKey(node, keyCode, shift)
End Function

Private Function FocusName() As String
    Dim f As clsNode
    Set f = modRender.FocusedNode()
    If f Is Nothing Then FocusName = "(none)" Else FocusName = NodeName(f)
End Function

Private Function NodeName(ByVal n As clsNode) As String
    If n Is Nothing Then NodeName = "(nothing)": Exit Function
    If Len(n.Id) > 0 Then
        NodeName = "#" & n.Id
    Else
        NodeName = n.LocalName & "." & Replace(n.ClassName, " ", ".") & " '" & Left$(Trim$(n.TextContent), 24) & "'"
    End If
End Function

Private Function ActiveBench() As clsNode
    Dim list As clsList, i As Long
    Set list = QA("main > section")
    For i = 0 To list.Count - 1
        If Not list(i).HasClass("hidden") Then Set ActiveBench = list(i): Exit Function
    Next i
End Function

Private Function EntryFor(ByVal id As String) As clsNode
    If id = "universal" Then
        Set EntryFor = Q(".universal-entry")
    Else
        Set EntryFor = Q("#tutorialList .operation[data-id=""" & id & """]")
    End If
End Function

Private Function OpenOp(ByVal id As String) As Boolean
    Dim entry As clsNode
    If CStr(AppState("id")) = id Then OpenOp = True: Exit Function
    If Len(Q("#searchInput").Value) > 0 Then Press Q("#clearSearch")
    Set entry = EntryFor(id)
    If entry Is Nothing Then Exit Function
    Press entry
    OpenOp = (CStr(AppState("id")) = id)
End Function

Private Function Norm(ByVal text As String) As String
    Norm = Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf)
End Function

Private Function ToastShown() As Boolean
    ToastShown = Q("#toast").HasClass("show")
End Function

Private Function ToastText() As String
    ToastText = Q("#toast").TextContent
End Function

' Elements inside the scope are shown by the DOM (no hidden/collapsed ancestor).
Private Function DomShown(ByVal node As clsNode, Optional ByVal ignoreCollapse As Boolean = False) As Boolean
    Dim p As clsNode, head As clsNode
    Set p = node
    Do While Not p Is Nothing
        If p.NodeType <> 1 Then Exit Do
        If p.HasClass("hidden") Or p.Hidden Then Exit Function
        If p.GetStyle("display") = "none" Then Exit Function
        If p.HasClass("lp-collapsible") And p.HasClass("is-collapsed") And Not p Is node And Not ignoreCollapse Then
            Set head = node.Closest(".panel-head")
            If head Is Nothing Then Exit Function
            If Not head.ParentElement Is p Then Exit Function
        End If
        If p.LocalName = "body" Then Exit Do
        Set p = p.ParentElement
    Loop
    DomShown = True
End Function

' Expands collapsed cards around the element (through their toggle buttons) until it is drawn.
Private Function EnsureShown(ByVal node As clsNode) As Boolean
    Dim panel As clsNode, toggle As clsNode, guard As Long
    If Drawn(node) Then EnsureShown = True: Exit Function
    Do
        guard = guard + 1
        If guard > 6 Then Exit Do
        Set panel = CollapsedAncestor(node)
        If panel Is Nothing Then Exit Do
        Set toggle = PanelToggle(panel)
        If toggle Is Nothing Then Exit Do
        If Not Press(toggle) Then Exit Do
    Loop
    EnsureShown = Drawn(node)
End Function

Private Function CollapsedAncestor(ByVal node As clsNode) As clsNode
    Dim p As clsNode, head As clsNode
    Set p = node.ParentElement
    Do While Not p Is Nothing
        If p.NodeType <> 1 Then Exit Do
        If p.HasClass("lp-collapsible") And p.HasClass("is-collapsed") Then
            Set head = node.Closest(".panel-head")
            If head Is Nothing Then Set CollapsedAncestor = p: Exit Function
            If Not head.ParentElement Is p Then Set CollapsedAncestor = p: Exit Function
        End If
        Set p = p.ParentElement
    Loop
End Function

Private Function PanelToggle(ByVal panel As clsNode) As clsNode
    Dim kids As clsList, i As Long
    Set kids = panel.Children
    For i = 0 To kids.Count - 1
        If kids(i).HasClass("panel-head") Then
            Set PanelToggle = kids(i).QuerySelector(".lp-collapse-toggle")
            Exit Function
        End If
    Next i
End Function

Private Function DrawnCount(ByVal scope As clsNode) As Long
    Dim list As clsList, i As Long
    If scope Is Nothing Then Exit Function
    Set list = scope.QuerySelectorAll("input, textarea, select, button")
    For i = 0 To list.Count - 1
        If Drawn(list(i)) Then DrawnCount = DrawnCount + 1
    Next i
End Function

' Every visible caption/text in the window, joined with line breaks.
Private Function PaintedText() As String
    Dim f As Object, c As Object, sb As String
    Set f = modRender.AppForm()
    If f Is Nothing Then Exit Function
    On Error Resume Next
    For Each c In f.Controls
        If ShownControl(c) Then
            Select Case TypeName(c)
                Case "Label", "CommandButton", "CheckBox", "OptionButton": sb = sb & c.Caption & vbLf
                Case "TextBox", "ComboBox": sb = sb & c.Text & vbLf
            End Select
        End If
    Next c
    PaintedText = sb
End Function

Private Function OpLabel(ByVal id As String) As String
    Dim i As Long
    If id = "universal" Then OpLabel = "Universal detector": Exit Function
    For i = 0 To modApp.OpCount - 1
        If modApp.OpIds(i) = id Then OpLabel = modApp.OpLabels(i): Exit Function
    Next i
End Function

Private Function FirstBlockText(ByVal el As clsNode) As String
    If el.FirstElementChild Is Nothing Then FirstBlockText = el.TextContent Else FirstBlockText = el.FirstElementChild.TextContent
End Function

' A disabled button: not a tab stop, and pressing it copies nothing.
Private Sub DisabledCopy(ByVal btn As clsNode, ByVal label As String)
    Dim c As Object
    Set c = CtlOf(btn)
    Ok label & ": button drawn", Not c Is Nothing
    If c Is Nothing Then Exit Sub
    Ok label & ": disabled button is not a tab stop", Not c.TabStop
    modClipboard.ClipboardWrite SENTINEL
    modClipboard.LastClipboardText = ""
    Press btn
    Ok label & ": pressing it copies nothing", RawClipboardRead() = SENTINEL And Len(modClipboard.LastClipboardText) = 0
End Sub

Private Function Painted(ByVal text As String) As Boolean
    text = Trim$(text)
    If Len(text) = 0 Then Exit Function
    Painted = InStr(1, PaintedText(), text, vbBinaryCompare) > 0
End Function

' A distinctive fragment of the text (the longest space-free ASCII token, at most 16 chars).
Private Function Fragment(ByVal text As String) As String
    Dim words() As String, i As Long, best As String, w As String, j As Long, okAscii As Boolean
    words = Split(Replace(Replace(text, vbLf, " "), vbTab, " "), " ")
    For i = 0 To UBound(words)
        w = words(i)
        okAscii = True
        For j = 1 To Len(w)
            If AscW(Mid$(w, j, 1)) > 126 Or AscW(Mid$(w, j, 1)) < 33 Then okAscii = False: Exit For
        Next j
        If okAscii And Len(w) > Len(best) Then best = w
    Next i
    If Len(best) > 16 Then best = Left$(best, 16)
    Fragment = best
End Function

' ================================================================ clipboard

Private Function RawClipboardRead() As String
#If Mac Then
    RawClipboardRead = modClipboard.ClipboardRead()
#Else
    Dim memory As LongPtr, pointer As LongPtr, length As Long, attempt As Long, opened As Boolean
    On Error GoTo Done
    For attempt = 1 To 20
        If OpenClipboard(0) <> 0 Then opened = True: Exit For
        Sleep 20
    Next attempt
    If Not opened Then Exit Function
    memory = GetClipboardData(13)
    If memory <> 0 Then
        pointer = GlobalLock(memory)
        If pointer <> 0 Then
            length = lstrlenW(pointer)
            RawClipboardRead = String$(length, vbNullChar)
            If length > 0 Then CopyMemory StrPtr(RawClipboardRead), pointer, CLngPtr(length) * 2
            GlobalUnlock memory
        End If
    End If
    CloseClipboard
Done:
#End If
End Function

Private Function ClipboardSeq() As Long
#If Mac Then
#Else
    ClipboardSeq = GetClipboardSequenceNumber()
#End If
End Function

' Writes the text with the app's clipboard routine and reads it straight back (up to five
' tries); returns the raw clipboard text when it matched, "" otherwise.
Private Function VerifyClipboardWrite(ByVal text As String) As String
    Dim attempt As Long, raw As String
    For attempt = 1 To 5
        If modClipboard.ClipboardWrite(text) Then
            raw = RawClipboardRead()
            If Norm(raw) = Norm(text) Then VerifyClipboardWrite = raw: Exit Function
        End If
    Next attempt
End Function

Private Function BareLineFeeds(ByVal raw As String) As Long
    Dim i As Long
    For i = 1 To Len(raw)
        If Mid$(raw, i, 1) = vbLf Then
            If i = 1 Then
                BareLineFeeds = BareLineFeeds + 1
            ElseIf Mid$(raw, i - 1, 1) <> vbCr Then
                BareLineFeeds = BareLineFeeds + 1
            End If
        End If
    Next i
End Function

' Where the copied text is shown: "exact <element>", "lines" (every line is in the
' workbench's text) or "" (not shown).
Private Function DisplayedMatch(ByVal text As String, ByVal scope As clsNode) As String
    Dim list As clsList, i As Long, el As clsNode, lines() As String, all As String, missing As Long, t As String, hiddenHit As String
    If scope Is Nothing Then Exit Function
    t = Trim$(text)
    Set list = scope.QuerySelectorAll("*")
    For i = 0 To list.Count - 1
        Set el = list(i)
        Select Case el.LocalName
            Case "textarea", "input"
                If Trim$(Norm(el.Value)) = t And Len(t) > 0 Then DisplayedMatch = "exact " & NodeName(el): Exit Function
            Case "script", "style", "option", "select"
            Case Else
                If Trim$(Norm(el.TextContent)) = t And Len(t) > 0 Then
                    If DomShown(el, True) Then DisplayedMatch = "exact " & NodeName(el): Exit Function
                    If Len(hiddenHit) = 0 Then hiddenHit = NodeName(el)
                End If
        End Select
    Next i
    all = Norm(scope.TextContent)
    lines = Split(text, vbLf)
    For i = 0 To UBound(lines)
        If Len(Trim$(lines(i))) > 0 Then
            If InStr(1, all, Trim$(lines(i)), vbBinaryCompare) = 0 Then missing = missing + 1
        End If
    Next i
    If missing = 0 And Len(t) > 0 Then
        DisplayedMatch = "lines"
        If Len(hiddenHit) > 0 Then DisplayedMatch = "lines (also equals hidden " & hiddenHit & ")"
    ElseIf Len(hiddenHit) > 0 Then
        DisplayedMatch = "exact (hidden) " & hiddenHit
    End If
End Function

' ================================================================ copy, answer options, alerts

Private Function IsProseButton(ByVal btn As clsNode) As Boolean
    Dim cap As String
    cap = LCase$(Trim$(btn.TextContent) & " " & NullText(btn.GetAttribute("title")))
    IsProseButton = (InStr(cap, "prose") > 0) Or btn.Id = "copyBtn" Or Not IsNull(btn.GetAttribute("data-sensitivity-copy")) _
        Or Not IsNull(btn.GetAttribute("data-copy-multi-regression")) Or Not IsNull(btn.GetAttribute("data-copy-multi-model")) _
        Or btn.Id = "copyLpAsciiBtn"
End Function

Private Function NullText(ByVal v As Variant) As String
    If Not IsNull(v) Then NullText = CStr(v)
End Function

Private Function IsStagedButton(ByVal btn As clsNode) As Boolean
    Select Case btn.Id
        Case "copyAnovaVarianceCodeBtn", "copyAnovaVarianceProseBtn", "copyAnovaStandardCodeBtn", "copyAnovaStandardProseBtn", _
             "copyAnovaWelchCodeBtn", "copyAnovaWelchProseBtn"
            IsStagedButton = True
    End Select
End Function

Private Function ButtonLabel(ByVal btn As clsNode) As String
    Dim owner As clsNode
    If Len(btn.Id) > 0 Then ButtonLabel = "#" & btn.Id: Exit Function
    Set owner = btn.ParentElement
    Do While Not owner Is Nothing
        If Len(owner.Id) > 0 Then Exit Do
        Set owner = owner.ParentElement
    Loop
    ButtonLabel = "'" & Trim$(btn.TextContent) & "'"
    If Not IsNull(btn.GetAttribute("data-sensitivity-copy")) Then ButtonLabel = ButtonLabel & "[" & btn.GetAttribute("data-sensitivity-copy") & "]"
    If Not IsNull(btn.GetAttribute("data-copy-multi-regression")) Then ButtonLabel = ButtonLabel & "[reg " & btn.GetAttribute("data-copy-multi-regression") & "]"
    If Not IsNull(btn.GetAttribute("data-copy-multi-model")) Then ButtonLabel = ButtonLabel & "[model " & btn.GetAttribute("data-copy-multi-model") & "]"
    If Not owner Is Nothing Then ButtonLabel = ButtonLabel & " in #" & owner.Id
End Function

Private Function CopyButtons(ByVal scope As clsNode) As clsList
    Dim list As clsList, i As Long, b As clsNode, cap As String
    Set CopyButtons = NewList()
    Set list = scope.QuerySelectorAll("button")
    For i = 0 To list.Count - 1
        Set b = list(i)
        cap = Trim$(b.TextContent)
        If b.Id = "copyBtn" Or (LCase$(Left$(cap, 4)) = "copy" And Not b.HasClass("lp-collapse-toggle")) Then CopyButtons.Push b
    Next i
End Function

' Presses a copy button and checks the real clipboard, the displayed text, the toast, any
' warning modal and the completion state. Returns the copied text (LF line endings).
Private Function CopyOnce(ByVal btn As clsNode, ByVal label As String) As String
    Dim raw As String, clip As String, scope As clsNode, panel As clsNode, expectComplete As Boolean, m As String
    Dim frag As String, painted As String, wasShown As Boolean, app As String, retried As String, seq0 As Long, seq1 As Long
    Set scope = ActiveBench()
    If Not EnsureShown(btn) Then
        Ok label & ": copy button drawn", False, "no control (DOM shown " & DomShown(btn) & ")"
        Exit Function
    End If
    modClipboard.ClipboardWrite SENTINEL
    modClipboard.LastClipboardText = ""
    seq0 = ClipboardSeq()
    Set panel = btn.Closest(".lp-collapsible")
    If Not panel Is Nothing Then expectComplete = ShouldCompletePanelOnCopy(panel, btn) And Not IsStagedButton(btn) And Not btn.Disabled
    If Not Press(btn) Then
        Ok label & ": copy pressed through its control", False
        Exit Function
    End If
    raw = RawClipboardRead()
    seq1 = ClipboardSeq()
    app = Norm(modClipboard.LastClipboardText)
    Ok label & ": app wrote the answer to the clipboard", Len(app) > 0, Left$(app, 80)
    If Norm(raw) = app Then
        Ok label & ": real clipboard = text the app wrote", True, Len(app) & " chars"
    Else
        ' The Windows clipboard is shared: another process (e.g. a concurrent test run) may
        ' have written it while the window re-rendered. Re-run the app's own clipboard write
        ' and read back at once to tell interference from a real defect.
        retried = VerifyClipboardWrite(app)
        If Len(retried) > 0 Then
            Soft label & ": real clipboard = text the app wrote (another process wrote the clipboard during the render; verified on an immediate rewrite)", True, "sequence +" & (seq1 - seq0)
            raw = retried
        Else
            Ok label & ": real clipboard = text the app wrote", False, "app wrote " & Len(app) & " chars, clipboard has " & Len(Norm(raw)) & " ('" & Left$(Norm(raw), 40) & "'), sequence +" & (seq1 - seq0)
        End If
    End If
    clip = app
    If InStr(raw, vbLf) > 0 Then Ok label & ": CRLF line endings on the clipboard", BareLineFeeds(raw) = 0, BareLineFeeds(raw) & " bare LF"
    If btn.Id = "copyLpAsciiBtn" Then
        ' The model table is shown as an HTML table; the copy is its ASCII rendering.
        Ok label & ": ASCII rendering of the displayed model table", Left$(clip, 1) = "+" And InStr(clip, "|") > 0 And AsciiCellsShown(clip, scope), "missing cell '" & mMissingCell & "'"
    Else
        m = DisplayedMatch(clip, scope)
        Ok label & ": clipboard matches the displayed output", Len(m) > 0 And InStr(m, "(hidden)") = 0, m
    End If
    Ok label & ": toast confirms the copy", ToastShown() And InStr(1, ToastText(), "cop", vbTextCompare) > 0, ToastText()
    HandleAlerts btn, label, clip
    If expectComplete Then
        Ok label & ": card completes (Done! + collapsed)", panel.HasClass("is-complete") And panel.HasClass("is-collapsed") And BadgeText(panel) = "Done!", _
            "class '" & panel.ClassName & "' badge '" & BadgeText(panel) & "'"
        If panel.HasClass("is-collapsed") Then
            Soft label & ": Done! badge painted", InStr(PaintedText(), "Done!") > 0
        End If
    End If
    ' The copied text is painted in the window (unless its card just collapsed).
    frag = Fragment(clip)
    wasShown = True
    If Not panel Is Nothing Then wasShown = Not panel.HasClass("is-collapsed")
    If Len(frag) >= 4 And wasShown Then
        painted = PaintedText()
        Soft label & ": copied text painted in the window", InStr(1, painted, frag, vbBinaryCompare) > 0, "fragment '" & frag & "'"
    End If
    CopyOnce = clip
End Function

' Every cell of an ASCII table appears in the scope's text.
Private Function AsciiCellsShown(ByVal table As String, ByVal scope As clsNode) As Boolean
    Dim lines() As String, cells() As String, i As Long, j As Long, all As String, cell As String
    all = Replace(Replace(Replace(Norm(scope.TextContent), " ", ""), ChrW$(&H2264), "<="), ChrW$(&H2265), ">=")
    lines = Split(table, vbLf)
    For i = 0 To UBound(lines)
        If Left$(lines(i), 1) = "|" Then
            cells = Split(lines(i), "|")
            For j = 0 To UBound(cells)
                cell = Replace(Trim$(cells(j)), " ", "")
                If Len(cell) > 0 Then
                    If InStr(1, all, cell, vbBinaryCompare) = 0 Then mMissingCell = cells(j): Exit Function
                End If
            Next j
        End If
    Next i
    AsciiCellsShown = True
End Function

Private Function BadgeText(ByVal panel As clsNode) As String
    Dim b As clsNode
    Set b = panel.QuerySelector(".completion-badge")
    If Not b Is Nothing Then BadgeText = Trim$(b.TextContent)
End Function

Private Function OpenAlertIds() As String
    Dim ids() As String, i As Long
    ids = Split(ALERT_IDS, " ")
    For i = 0 To UBound(ids)
        If Not ById(ids(i)) Is Nothing Then
            If Not ById(ids(i)).HasClass("hidden") Then OpenAlertIds = OpenAlertIds & ids(i) & " "
        End If
    Next i
    OpenAlertIds = Trim$(OpenAlertIds)
End Function

Private Function AlertButton(ByVal alertId As String) As String
    Select Case alertId
        Case "plotAxisRangeAlert": AlertButton = "dismissPlotAxisRangeAlert"
        Case "variableRenameAlert": AlertButton = "dismissVariableRenameAlert"
        Case "regressionUnitsAlert": AlertButton = "dismissRegressionUnitsAlert"
        Case "kmeansScaleAlert": AlertButton = "keepKmeansScalingBtn"
        Case "knownPopulationSdAlert": AlertButton = "useUnknownSdBtn"
    End Select
End Function

Private Function IsChartCode(ByVal code As String) As Boolean
    On Error Resume Next
    IsChartCode = CBool(Application.Run("ContainsRChart", code))
End Function

Private Sub HandleAlerts(ByVal btn As clsNode, ByVal label As String, ByVal clip As String)
    Dim expectPlot As Boolean, openIds As String, expectRename As Boolean, expectUnits As Boolean
    expectPlot = Not modShared.CalculatePlotAxisIntervals And Not IsProseButton(btn) And IsChartCode(clip)
    expectRename = (btn.Id = "copySensitivityProseBtn") Or (NullText(btn.GetAttribute("data-sensitivity-copy")) = "solution")
    expectUnits = (btn.Id = "copyMultiRegressionProseBtn") Or Not IsNull(btn.GetAttribute("data-copy-multi-regression"))
    openIds = OpenAlertIds()
    Ok label & ": chart axis-range warning " & IIf(expectPlot, "shown", "not shown"), (InStr(openIds, "plotAxisRangeAlert") > 0) = expectPlot, "open: '" & openIds & "'"
    Ok label & ": placeholder-name warning " & IIf(expectRename, "shown", "not shown"), (InStr(openIds, "variableRenameAlert") > 0) = expectRename, "open: '" & openIds & "'"
    Ok label & ": regression-units warning " & IIf(expectUnits, "shown", "not shown"), (InStr(openIds, "regressionUnitsAlert") > 0) = expectUnits, "open: '" & openIds & "'"
    If InStr(openIds, "plotAxisRangeAlert") > 0 Then DismissAlert "plotAxisRangeAlert", btn, label
    If InStr(openIds, "variableRenameAlert") > 0 Then DismissAlert "variableRenameAlert", btn, label
    If InStr(openIds, "regressionUnitsAlert") > 0 Then DismissAlert "regressionUnitsAlert", btn, label
End Sub

' Checks a warning modal (drawn, focused, Tab trapped, Escape) and dismisses it with
' Escape (axis-range warning, every other time) or its button; focus must return.
Private Sub DismissAlert(ByVal alertId As String, ByVal returnTo As clsNode, ByVal label As String)
    Dim al As clsNode, dismiss As clsNode, cancel As Boolean, viaEscape As Boolean, tag As String
    Set al = ById(alertId)
    Set dismiss = ById(AlertButton(alertId))
    tag = label & ": " & alertId
    Ok tag & " drawn over the page", Drawn(dismiss), DescribeOverlay()
    Ok tag & " focuses its button", modRender.FocusedNode() Is dismiss, "focus " & FocusName()
    Ok tag & " body marked modal-open", Body.HasClass("variable-rename-alert-open")
    If alertId = "plotAxisRangeAlert" Then
        Ok tag & " footnote", InStr(al.TextContent, "Press Escape to dismiss this warning.") > 0
        Soft tag & " footnote painted", InStr(PaintedText(), "Press Escape to dismiss this warning.") > 0
    End If
    cancel = KeyOn(dismiss, 9, 0)
    Ok tag & " Tab stays inside the modal", cancel And modRender.FocusedNode() Is dismiss, "cancelled " & cancel & ", focus " & FocusName()
    If alertId = "plotAxisRangeAlert" Then
        mEscapeTurn = mEscapeTurn + 1
        viaEscape = (mEscapeTurn Mod 2 = 1)
    End If
    If viaEscape Then
        cancel = KeyOn(dismiss, 27, 0)
        Ok tag & " Escape dismisses", al.HasClass("hidden") And cancel, "cancelled " & cancel
    Else
        If alertId <> "plotAxisRangeAlert" Then
            KeyOn dismiss, 27, 0
            Ok tag & " Escape keeps the decision modal open (as the website)", Not al.HasClass("hidden")
        End If
        Ok tag & " dismiss button pressed", Press(dismiss)
        Ok tag & " closed by its button", al.HasClass("hidden")
    End If
    Ok tag & " modal-open class cleared", Not Body.HasClass("variable-rename-alert-open")
    If returnTo.Disabled Then
        Soft tag & " focus target disabled after copying (website: focus falls to the page)", True, "focus " & FocusName()
    ElseIf Drawn(returnTo) Then
        Ok tag & " focus returns to " & NodeName(returnTo), modRender.FocusedNode() Is returnTo, "focus " & FocusName()
    Else
        Soft tag & " focus target still drawn", False, NodeName(returnTo) & IIf(CollapsedAncestor(returnTo) Is Nothing, " (not in a collapsed card)", " (inside collapsed #" & CollapsedAncestor(returnTo).Id & ")") & "; window focus " & FocusName() & ", DOM focus " & NodeName(modDom.ActiveElement())
    End If
End Sub

Private Function DescribeOverlay() As String
    Dim d As String
    d = modRender.DescribeWindow()
    If InStr(d, "ov") > 0 Then DescribeOverlay = "overlay frame present" Else DescribeOverlay = "no overlay frame listed"
End Function

Private Sub CloseAlerts()
    Dim ids() As String, i As Long, guard As Long
    ids = Split(ALERT_IDS, " ")
    For i = 0 To UBound(ids)
        guard = 0
        Do While Not ById(ids(i)).HasClass("hidden") And guard < 2
            Press ById(AlertButton(ids(i)))
            guard = guard + 1
        Loop
    Next i
End Sub

' Picks Option n on the button's answer dropdown through the ComboBox.
Private Function PickOption(ByVal sel As clsNode, ByVal index As Long, ByVal label As String) As Boolean
    Dim panel As clsNode, before As String, after As String, scope As clsNode, c As Object
    If Not EnsureShown(sel) Then
        Ok label & ": answer dropdown drawn", False
        Exit Function
    End If
    Set scope = ActiveBench()
    before = Norm(scope.TextContent)
    Set panel = sel.Closest(".lp-collapsible")
    If Not Pick(sel, index) Then
        Ok label & ": pick Option " & (index + 1), False
        Exit Function
    End If
    Set c = CtlOf(sel)
    Ok label & ": Option " & (index + 1) & " selected in DOM and window", sel.Value = CStr(index + 1) And c.ListIndex = index, "DOM " & sel.Value & ", list " & c.ListIndex
    after = Norm(scope.TextContent)
    If Not panel Is Nothing Then Ok label & ": option change clears completion", Not panel.HasClass("is-complete")
    PickOption = (before <> after)
End Function

Private Sub OptionCycle(ByVal btn As clsNode, ByVal picker As clsNode, ByVal label As String)
    Dim sel As clsNode, c As Object, t1 As String, t2 As String, t2r As String, t3 As String, t2b As String, t1b As String, changed As Boolean
    Set sel = picker.QuerySelector("select")
    If sel Is Nothing Then Ok label & ": dropdown has a select", False: Exit Sub
    If Not EnsureShown(sel) Then Ok label & ": answer dropdown drawn", False, "DOM shown " & DomShown(sel): Exit Sub
    Set c = CtlOf(sel)
    Ok label & ": dropdown is a ComboBox with six options", TypeName(c) = "ComboBox" And c.ListCount = 6 And sel.Options.Count = 6, TypeName(c) & " " & c.ListCount
    Ok label & ": dropdown shows the DOM selection", c.ListIndex = sel.SelectedIndex, c.ListIndex & " vs " & sel.SelectedIndex
    PickOption sel, 0, label
    t1 = CopyOnce(btn, label & " opt1")
    changed = PickOption(sel, 1, label)
    t2 = CopyOnce(btn, label & " opt2")
    Soft label & ": Option 2 changes the visible output", changed And t2 <> t1
    t2r = CopyOnce(btn, label & " opt2 again")
    Ok label & ": repeated copy keeps the exact text", t2r = t2
    PickOption sel, 2, label
    t3 = CopyOnce(btn, label & " opt3")
    Soft label & ": Option 3 differs from Option 2", t3 <> t2
    PickOption sel, 1, label
    t2b = CopyOnce(btn, label & " opt2 reselected")
    Soft label & ": reselecting Option 2 gives a fresh variation", t2b <> t2
    PickOption sel, 0, label
    t1b = CopyOnce(btn, label & " opt1 again")
    Ok label & ": Option 1 restores the canonical answer", t1b = t1, IIf(t1b = t1, "", "first diff at " & FirstDiff(t1, t1b))
End Sub

Private Function FirstDiff(ByVal a As String, ByVal b As String) As String
    Dim i As Long
    For i = 1 To IIf(Len(a) < Len(b), Len(a), Len(b))
        If Mid$(a, i, 1) <> Mid$(b, i, 1) Then Exit For
    Next i
    FirstDiff = i & ": '" & Mid$(a, IIf(i > 10, i - 10, 1), 30) & "' vs '" & Mid$(b, IIf(i > 10, i - 10, 1), 30) & "'"
End Function

' Sets every drawn answer dropdown of the workbench to Option 1 through its ComboBox.
Private Sub CanonAll(ByVal scope As clsNode)
    Dim list As clsList, i As Long
    Set list = scope.QuerySelectorAll(".answer-variant-select")
    For i = 0 To list.Count - 1
        If Drawn(list(i)) Then
            If list(i).Value <> "1" Then Pick list(i), 0
        End If
    Next i
End Sub

' ================================================================ per-operation sweep

Private Sub OpSweep(ByVal id As String)
    Dim entry As clsNode, bench As clsNode, list As clsList, i As Long, seen As clsMap, picker As clsNode, btn As clsNode, label As String
    mScenario = "op " & id
    If Len(Q("#searchInput").Value) > 0 Then Press Q("#clearSearch")
    Set entry = EntryFor(id)
    Ok "sidebar entry exists", Not entry Is Nothing
    If entry Is Nothing Then Exit Sub
    Ok "sidebar entry opens the operation", Press(entry) And CStr(AppState("id")) = id, "state " & AppState("id")
    Ok "sidebar entry marked active", entry.HasClass("active")
    Soft "sidebar entry painted", Painted(OpLabel(id)), OpLabel(id)
    If id = "universal" Then Soft "sidebar entries keyboard-focusable (drawn as a control, not a label + hit area)", Drawn(entry)
    Set bench = ActiveBench()
    Ok "workbench drawn", DrawnCount(bench) > 0, DrawnCount(bench) & " controls in #" & bench.Id
    Ok "title shows the operation", Len(Trim$(Q("#activeTitle").TextContent)) > 0, Q("#activeTitle").TextContent
    SampleSweep id, bench
    If id = "summarize" Then
        ' No sample: the column fields are typed (the first keystroke appends a new field).
        Ok "typed a column name", TypeChars(QA("#columnFields .column-row input")(0), "Amount")
        Ok "column field appended while typing", QA("#columnFields .column-row").Count >= 2
    End If
    TypingSweep bench
    Set seen = NewMap()
    Set list = CopyButtons(bench)
    For i = 0 To list.Count - 1
        Set btn = list(i)
        label = ButtonLabel(btn)
        If Not DomShown(btn, True) Then
            ' Hidden by the operation (another mode): not drawn, nothing to press.
            Ok label & ": hidden button not drawn", Not Drawn(btn)
        ElseIf btn.Disabled Then
            If Not btn.Closest(".is-unavailable") Is Nothing Then
                Ok label & ": disabled on the unavailable route", True
            ElseIf Trim$(btn.TextContent) = "Copied" Then
                Ok label & ": section already marked Copied (disabled)", True
            Else
                Soft label & ": enabled after loading the samples", False, "disabled"
            End If
        Else
            Set picker = modVariants.PickerForCopyButton(btn)
            If picker Is Nothing Then
                CopyOnce btn, label
            ElseIf seen.Exists(CStr(picker.Uid)) Then
                CopyOnce btn, label
            Else
                seen(CStr(picker.Uid)) = True
                OptionCycle btn, picker, label
            End If
        End If
        CloseAlerts
    Next i
    Ok "copy buttons found", list.Count > 0, list.Count & " buttons, " & seen.Count & " answer dropdowns"
    CheckErrors
End Sub

Private Sub SampleSweep(ByVal id As String, ByVal bench As clsNode)
    Dim list As clsList, i As Long, b As clsNode, before As String, label As String, cap As String, n As Long
    If id = "universal" Then
        Ok "sample list drawn", Pick(Q("#universalSampleSelect"), 1)
        Ok "sample list loads the first sample", Len(Q("#inputText").Value) > 0 And Len(Trim$(Q("#outputBody").TextContent)) > 0
        If CStr(AppState("id")) <> "universal" Then OpenOp "universal"
        Exit Sub
    End If
    If bench.Id = "standardBench" Then
        Ok "Use sample button drawn", Drawn(Q("#singleSampleBtn"))
        Ok "Use sample pressed", Press(Q("#singleSampleBtn"))
        Ok "sample fills the input", Len(Q("#inputText").Value) > 0, Left$(Q("#inputText").Value, 40)
        Ok "sample generates prose", InStr(Q("#outputBody").TextContent, "H") > 0, Left$(Q("#outputBody").TextContent, 60)
        Soft "input TextBox shows the sample", Norm(CtlOf(Q("#inputText")).Text) = Norm(Q("#inputText").Value)
        Exit Sub
    End If
    Set list = bench.QuerySelectorAll(".sample-btn")
    For i = 0 To list.Count - 1
        Set b = list(i)
        cap = Trim$(b.TextContent)
        If LCase$(Left$(cap, 10)) = "use sample" Or LCase$(Left$(cap, 10)) = "use parser" Then
            label = "sample " & ButtonLabel(b)
            If Not DomShown(b, True) Then
                Ok label & " hidden and not drawn", Not Drawn(b)
            ElseIf b.Disabled Then
                If Not b.Closest(".is-unavailable") Is Nothing Then Ok label & " disabled on the unavailable route", True Else Soft label & " enabled", False, "disabled"
            ElseIf Not EnsureShown(b) Then
                Ok label & " drawn", False, "DOM shown " & DomShown(b)
            Else
                before = Norm(bench.TextContent) & Values(bench)
                Ok label & " pressed", Press(b)
                Ok label & " loads (output or inputs changed, toast)", before <> Norm(bench.TextContent) & Values(bench) Or InStr(1, ToastText(), "loaded", vbTextCompare) > 0, ToastText()
                n = n + 1
            End If
        End If
    Next i
    If id <> "summarize" Then Soft "operation has a sample button", n > 0, n & " pressed"
End Sub

Private Function Values(ByVal scope As clsNode) As String
    Dim list As clsList, i As Long
    Set list = scope.QuerySelectorAll("input, textarea, select")
    For i = 0 To list.Count - 1
        Values = Values & "|" & list(i).Value
    Next i
End Function

' Types into the operation's fields through their TextBoxes and checks the output follows.
Private Sub TypingSweep(ByVal bench As clsNode)
    Dim list As clsList, i As Long, el As clsNode, c As Object, v As String, outText As String, target As clsNode
    Dim newValue As String, ta As clsNode, cleared As String, restored As String, original As String
    outText = Norm(bench.TextContent)
    Set list = bench.QuerySelectorAll("input")
    For i = 0 To list.Count - 1
        Set el = list(i)
        Select Case LCase$(el.TypeName)
            Case "", "text", "search"
                v = Trim$(el.Value)
                If Len(v) >= 3 And Not el.Disabled And Drawn(el) And NullText(el.GetAttribute("role")) <> "combobox" Then
                    If InStr(1, outText, v, vbBinaryCompare) > 0 Then Set target = el: Exit For
                End If
        End Select
    Next i
    If target Is Nothing Then
        ' A number field instead (Skewness & kurtosis has only numbers).
        For i = 0 To list.Count - 1
            Set el = list(i)
            If LCase$(el.TypeName) = "number" And Not el.Disabled And Drawn(el) Then Set target = el: Exit For
        Next i
        If target Is Nothing Then
            If bench.QuerySelector("textarea") Is Nothing Then Soft "a text field whose value appears in the output", False
        Else
            original = target.Value
            newValue = "1.37"
            If original = newValue Then newValue = "-0.82"
            Ok "typed a number into " & NodeName(target), TypeText(target, newValue)
            Ok "DOM value follows the TextBox", target.Value = newValue, target.Value
            Ok "output shows the typed number", InStr(1, Norm(bench.TextContent), newValue, vbBinaryCompare) > 0
            TypeText target, original
            Set target = Nothing
        End If
    Else
        original = target.Value
        newValue = Trim$(original) & "Zq"
        Ok "typed into " & NodeName(target), TypeText(target, newValue)
        Ok "DOM value follows the TextBox", target.Value = newValue, target.Value
        Ok "output shows the typed text", InStr(1, Norm(bench.TextContent), newValue, vbBinaryCompare) > 0
        Ok "TextBox keeps the typed text after re-render", CtlOf(target).Text = newValue, CtlOf(target).Text
        TypeText target, original
        Ok "retyping the original removes it from the output", InStr(1, Norm(bench.TextContent), newValue, vbBinaryCompare) = 0
    End If
    ' The main paste area (the first drawn textarea holding text).
    Set list = bench.QuerySelectorAll("textarea")
    For i = 0 To list.Count - 1
        If Len(Trim$(list(i).Value)) > 0 And Drawn(list(i)) And Not list(i).ReadOnly Then Set ta = list(i): Exit For
    Next i
    If ta Is Nothing Then Exit Sub
    original = ta.Value
    outText = Norm(bench.TextContent)
    Ok "cleared " & NodeName(ta) & " by typing", TypeText(ta, "")
    cleared = Norm(bench.TextContent)
    Ok "clearing the paste area updates the output", cleared <> outText
    Ok "typed the output back into " & NodeName(ta), TypeText(ta, original)
    Ok "DOM keeps the pasted text (line breaks normalised)", Norm(ta.Value) = Norm(original), Len(ta.Value) & " vs " & Len(original)
    restored = Norm(bench.TextContent)
    Ok "typing the output regenerates the answer", restored <> cleared
End Sub

' ================================================================ targeted scenarios

Private Sub ScenarioUniversal()
    Dim sel As clsNode, c As Object, i As Long, route As String, selfContained As Boolean, fr As Object, frag As String
    OpenOp "universal"
    Set sel = Q("#universalSampleSelect")
    Set c = CtlOf(sel)
    Ok "sample dropdown drawn", Not c Is Nothing
    If c Is Nothing Then Exit Sub
    Ok "one entry per trained sample (plus prompt)", c.ListCount = modApp.UniversalSampleCount + 1, c.ListCount & " entries"
    For i = 1 To c.ListCount - 1
        OpenOp "universal"
        route = sel.Options(i).Value
        Ok "pick sample " & route, Pick(sel, i)
        Ok route & ": sample text in the input", Len(Q("#inputText").Value) > 0
        Select Case route
            ' The website routes interval output to the interval operation (AGENTS.md lists it as
            ' self-contained); the port follows the website.
            Case "shapiro", "pearson", "one-sample", "two-sample": selfContained = True
            Case Else: selfContained = False
        End Select
        If selfContained Then
            Ok route & ": stays on the detector", CStr(AppState("id")) = "universal", CStr(AppState("id"))
            Ok route & ": prose generated", Len(Trim$(Q("#outputBody").TextContent)) > 20, Left$(Q("#outputBody").TextContent, 60)
            frag = Fragment(FirstBlockText(Q("#outputBody")))
            If Len(frag) < 4 Then frag = Fragment(Q("#outputBody").TextContent)
            Soft route & ": prose painted", Painted(frag), frag
        Else
            Ok route & ": navigates to the operation", CStr(AppState("id")) = route, CStr(AppState("id"))
            Set fr = modRender.RegionFrame("main")
            If route <> "lpsolve" And route <> "glm" And route <> "confusion" Then
                If Not fr Is Nothing Then Soft route & ": page scrolled to the interpretation block", fr.ScrollTop > 0, "ScrollTop " & fr.ScrollTop
            End If
        End If
    Next i
End Sub

Private Sub ScenarioSearch()
    Dim s As clsNode, first As clsNode, cancel As Boolean
    OpenOp "universal"
    Set s = Q("#searchInput")
    Ok "search box drawn", Drawn(s)
    Ok "typed 'anova' into search", TypeChars(s, "anova")
    Ok "search value", s.Value = "anova", s.Value
    Ok "ranked results shown", Not Q("#rankedSearchResults").HasClass("hidden") And Q("#tutorialList").HasClass("hidden")
    Set first = Q("#rankedSearchResults .ranked-search-result")
    Ok "first result painted", Painted(first.TextContent), NodeName(first)
    Ok "first result is the exact match", NullText(first.GetAttribute("data-search-exact")) = "true", first.TextContent
    cancel = KeyOn(s, 13)
    Ok "Enter opens the exact match", CStr(AppState("id")) = "anova-many", CStr(AppState("id"))
    Press Q("#clearSearch")
    Ok "clear button empties search", s.Value = "" And Not Q("#tutorialList").HasClass("hidden")
    Ok "clear button focuses search", modRender.FocusedNode() Is s, FocusName()
    Ok "typed 'regresion' (fuzzy)", TypeText(s, "regresion")
    Set first = Q("#rankedSearchResults .ranked-search-result")
    Ok "fuzzy results listed", Not first Is Nothing
    If Not first Is Nothing Then
        Ok "result painted", Painted(first.TextContent), first.TextContent
        Ok "pressing a result opens it", Press(first) And CStr(AppState("id")) = NullText(first.GetAttribute("data-id")), CStr(AppState("id"))
    End If
    TypeText s, "zzzzqqq"
    Ok "no-match notice", Not Q("#noResults").HasClass("hidden") And InStr(Q("#resultCount").TextContent, "closest") > 0, Q("#resultCount").TextContent
    Soft "no-match notice painted", InStr(PaintedText(), "No matches") > 0
    TypeText s, "cor.test"
    Ok "search by R function finds Pearson", InStr(Q("#rankedSearchResults").TextContent, "Pearson") > 0, Left$(Q("#rankedSearchResults").TextContent, 80)
    TypeText s, ""
    Ok "empty search restores the list", Not Q("#tutorialList").HasClass("hidden") And Q("#rankedSearchResults").HasClass("hidden")
End Sub

Private Sub ScenarioLetter()
    Dim cancel As Boolean
    OpenOp "summarize"
    If Len(Q("#searchInput").Value) > 0 Then Press Q("#clearSearch")
    cancel = KeyOn(Q("#themeToggle"), 75, 0)
    Ok "letter on a button types into search", Q("#searchInput").Value = "k" And cancel, "value '" & Q("#searchInput").Value & "', cancelled " & cancel
    Ok "search has the window focus", modRender.FocusedNode() Is Q("#searchInput"), FocusName()
    Ok "search TextBox shows the letter", CtlOf(Q("#searchInput")).Text = "k", CtlOf(Q("#searchInput")).Text
    cancel = KeyOn(Q("#copySummaryCodeBtn"), 77, 1)
    Ok "shifted letter appends to search", Q("#searchInput").Value = "kM" Or Q("#searchInput").Value = "km", Q("#searchInput").Value
    Press Q("#clearSearch")
    cancel = KeyOn(Q("#summaryDataset"), 75, 0)
    Ok "letter in a text field stays in the field", Q("#searchInput").Value = "", "search '" & Q("#searchInput").Value & "'"
    cancel = KeyOn(Q(".universal-entry"), 75, 2)
    Ok "Ctrl+letter does not search", Q("#searchInput").Value = "", Q("#searchInput").Value
End Sub

Private Sub ScenarioTheme()
    Dim before As Variant, fr As Object, bg1 As Long, bg2 As Long
    before = DocumentElement.GetAttribute("data-theme")
    Set fr = modRender.RegionFrame("main")
    If Not fr Is Nothing Then bg1 = fr.BackColor
    Ok "theme toggle drawn", Drawn(Q("#themeToggle"))
    Press Q("#themeToggle")
    Ok "theme toggles to dark", NullText(DocumentElement.GetAttribute("data-theme")) = "dark", NullText(DocumentElement.GetAttribute("data-theme"))
    Ok "toggle pressed state", NullText(Q("#themeToggle").GetAttribute("aria-pressed")) = "true"
    Set fr = modRender.RegionFrame("main")
    If Not fr Is Nothing Then bg2 = fr.BackColor
    Ok "window repainted in dark colours", bg1 <> bg2, Hex$(bg1) & " -> " & Hex$(bg2)
    Ok "theme persisted (test store)", NullText(modApp.StorageGetItem("bt1101-theme")) = "dark", NullText(modApp.StorageGetItem("bt1101-theme"))
    Ok "toggle icon switches to the sun", Q("#themeToggle .theme-toggle-icon").TextContent = ChrW$(&H2600)
    Soft "sun icon painted", Painted(ChrW$(&H2600))
    Press Q("#themeToggle")
    Ok "theme toggles back to light", NullText(DocumentElement.GetAttribute("data-theme")) = "light"
    Set fr = modRender.RegionFrame("main")
    If Not fr Is Nothing Then Ok "window repainted in light colours", fr.BackColor = bg1, Hex$(fr.BackColor)
End Sub

Private Sub ScenarioResizer()
    Dim v As String, lib As Object, w1 As Double
    modRender.SetClientSize 1280, 860
    modRender.ResizeNow
    Set lib = modRender.RegionFrame("library")
    If Not lib Is Nothing Then w1 = lib.Width
    modRender.SimulateResizerDrag 310, 400
    v = DocumentElement.GetStyle("--sidebar-width")
    Ok "drag sets the sidebar width", v = "400px", v
    Set lib = modRender.RegionFrame("library")
    If Not lib Is Nothing Then Ok "library frame widened", lib.Width > w1 + 30, w1 & " -> " & lib.Width
    Ok "width persisted (test store)", NullText(modApp.StorageGetItem("bt1101-sidebar-width")) <> "", NullText(modApp.StorageGetItem("bt1101-sidebar-width"))
    modRender.SimulateResizerDrag 400, 120
    v = DocumentElement.GetStyle("--sidebar-width")
    Ok "drag clamps to the minimum width", Val(v) > 120, v
    modRender.SimulateResizerDrag Val(v), 310
    v = DocumentElement.GetStyle("--sidebar-width")
    Ok "drag back", Abs(Val(v) - 310) <= 4, v
End Sub

Private Sub ScenarioResize()
    Dim fr As Object, w As Variant, btn As clsNode, entry As clsNode, cancel As Boolean
    For Each w In Array(960, 800)
        modRender.SetClientSize CDbl(w), 860
        modRender.ResizeNow
        Ok "innerWidth " & w, modApp.WindowInnerWidth = w, CStr(modApp.WindowInnerWidth)
        Set fr = modRender.RegionFrame("main")
        If Not fr Is Nothing Then Ok "no horizontal page scroll at " & w, (fr.ScrollBars And 1) = 0 Or fr.ScrollWidth <= fr.InsideWidth + 2, "bars " & fr.ScrollBars & " sw " & fr.ScrollWidth & " iw " & fr.InsideWidth
        Ok "sidebar still drawn at " & w, Drawn(Q("#searchInput"))
        Ok "mobile library button hidden at " & w, Not Painted("Browse operations")
    Next w
    modRender.SetClientSize 375, 760
    modRender.ResizeNow
    Ok "innerWidth 375", modApp.WindowInnerWidth = 375, CStr(modApp.WindowInnerWidth)
    Set btn = Q("#mobileLibraryBtn")
    Ok "mobile library button painted at 375", Painted("Browse operations")
    Soft "mobile library button keyboard-focusable (a control, not a label)", Drawn(btn)
    Set fr = modRender.RegionFrame("main")
    If Not fr Is Nothing Then Ok "no horizontal page scroll at 375", (fr.ScrollBars And 1) = 0 Or fr.ScrollWidth <= fr.InsideWidth + 2, "bars " & fr.ScrollBars & " sw " & fr.ScrollWidth & " iw " & fr.InsideWidth
    Ok "library closed at first", Not Q("#library").HasClass("open")
    Ok "library button opens the library", Press(btn) And Q("#library").HasClass("open")
    Ok "aria-expanded true", NullText(btn.GetAttribute("aria-expanded")) = "true"
    Set entry = EntryFor("pearson")
    Ok "library entries painted when open", Painted(OpLabel("pearson"))
    If CStr(AppState("id")) <> "pearson" Then Press entry
    Ok "choosing an operation closes the library", CStr(AppState("id")) = "pearson" And Not Q("#library").HasClass("open")
    cancel = KeyOn(Q("#singleSampleBtn"), 83, 0)
    Ok "letter at phone width opens the library and searches", Q("#library").HasClass("open") And Q("#searchInput").Value = "s", "open " & Q("#library").HasClass("open") & ", search '" & Q("#searchInput").Value & "'"
    Press Q("#clearSearch")
    If Q("#library").HasClass("open") Then Press btn
    Ok "library button closes the library", Not Q("#library").HasClass("open")
    ' The window's own resize path: UserForm_Resize schedules ResizeTick (Application.OnTime).
    modRender.SetClientSize 960, 860
    modRender.ResizeTick
    Ok "resize event path (ResizeTick) updates innerWidth", modApp.WindowInnerWidth = 960, CStr(modApp.WindowInnerWidth)
    Set fr = modRender.RegionFrame("library")
    If Not fr Is Nothing Then
        Ok "library at most 30vw (288px) at 960", fr.Width / (modRender.AppForm().InsideWidth / 960) <= 290, Format$(fr.Width / (modRender.AppForm().InsideWidth / 960), "0") & " px"
    End If
    modRender.SetClientSize 1280, 860
    modRender.ResizeNow
    Ok "back to 1280", modApp.WindowInnerWidth = 1280, CStr(modApp.WindowInnerWidth)
End Sub

Private Sub ScenarioPlotSearch()
    Dim s As clsNode, wf As clsNode, cancel As Boolean, menu As clsNode, opt As clsNode, active As String
    OpenOp "ggplot"
    Set s = Q("#plotWorkflowSearch")
    Set wf = Q("#plotWorkflow")
    Set menu = Q("#plotWorkflowMenu")
    Ok "plot search drawn", Drawn(s)
    TypeText s, "histog"
    cancel = KeyOn(s, 9)
    Ok "one result: Tab commits", s.Value = "histogram" And wf.Value = "histogram", s.Value & " / " & wf.Value
    Ok "one result: Tab advances naturally (not cancelled)", Not cancel
    Ok "menu closed after commit", menu.HasClass("hidden")
    Press Q("#loadGgplotSampleBtn")
    Ok "code follows the committed plot", InStr(Q("#ggplotCode").TextContent, "hist(") > 0, Left$(Q("#ggplotCode").TextContent, 60)
    TypeText s, "bar"
    cancel = KeyOn(s, 9)
    Ok "several results: Tab stays in search", cancel And modRender.FocusedNode() Is s, "cancelled " & cancel & ", focus " & FocusName()
    Ok "several results: menu open", Not menu.HasClass("hidden")
    active = NullText(s.GetAttribute("aria-activedescendant"))
    Ok "first result highlighted", Len(active) > 0 And Not menu.QuerySelector(".plot-workflow-option") Is Nothing, active
    If Len(active) > 0 Then Ok "highlight is the first option", menu.QuerySelector(".plot-workflow-option").Id = active, active
    Set opt = menu.QuerySelector(".plot-workflow-option")
    Ok "menu options painted", Painted(opt.TextContent), opt.TextContent
    Soft "headings painted", InStr(1, PaintedText(), "Summaries", vbTextCompare) > 0
    cancel = KeyOn(s, 40)
    Ok "Down moves the highlight", NullText(s.GetAttribute("aria-activedescendant")) <> active And cancel, NullText(s.GetAttribute("aria-activedescendant"))
    active = NullText(s.GetAttribute("aria-activedescendant"))
    cancel = KeyOn(s, 38)
    cancel = KeyOn(s, 40)
    Ok "Up then Down returns", NullText(s.GetAttribute("aria-activedescendant")) = active
    cancel = KeyOn(s, 13)
    Ok "Enter commits the highlighted result", Len(wf.Value) > 0 And menu.HasClass("hidden") And "plotWorkflowOption-" & NullText(wf.GetAttribute("data-committed-choice")) = active, s.Value & " / " & wf.Value & " / " & active
    TypeText s, "Stacked bar plot"
    cancel = KeyOn(s, 9)
    Ok "stacked bar plot routes to contingency + Stacked", wf.Value = "contingency" And Q("#plotStyle").Value = "stacked", wf.Value & " / " & Q("#plotStyle").Value
    Ok "arrangement dropdown shows Stacked", CtlOf(Q("#plotStyle")).ListIndex = 1
    TypeText s, ""
    cancel = KeyOn(s, 9)
    Ok "unfiltered list: Tab highlights the first", cancel And Not menu.HasClass("hidden") And Len(NullText(s.GetAttribute("aria-activedescendant"))) > 0
    Ok "Grouped barplot listed", Not Q("#plotWorkflowOption-grouped-summary") Is Nothing
    cancel = KeyOn(s, 27)
    Ok "Escape closes the menu", menu.HasClass("hidden")
    TypeText s, "grouped barplot"
    Set opt = Q("#plotWorkflowOption-grouped-summary")
    If Not opt Is Nothing Then
        Ok "pressing a menu option commits it", Press(opt) And wf.Value = "grouped-summary", wf.Value
    Else
        Ok "Grouped barplot option rendered in the menu", False
    End If
    TypeText s, "histogram"
    KeyOn s, 9
End Sub

Private Sub ScenarioStats()
    Dim rows As clsList, inp As clsNode, n As Long, code As String, cb As clsNode, keys As Variant, k As Variant
    OpenOp "summarize"
    Set rows = QA("#columnFields .column-row")
    ' Back to one empty row.
    Do While rows.Count > 1
        Press rows(rows.Count - 1).QuerySelector("button")
        Set rows = QA("#columnFields .column-row")
        n = n + 1
        If n > 10 Then Exit Do
    Loop
    Set inp = QA("#columnFields .column-row")(0).QuerySelector("input")
    TypeText inp, ""
    Ok "waiting state", InStr(Q("#summaryCode").TextContent, "Enter a column name") > 0 And Q("#copySummaryCodeBtn").Disabled
    DisabledCopy Q("#copySummaryCodeBtn"), "copy (no column)"
    Ok "typed a column name key by key", TypeChars(inp, "Amount")
    Set rows = QA("#columnFields .column-row")
    Ok "a new column field appears while typing", rows.Count = 2, rows.Count & " rows"
    Ok "the new field is drawn", Drawn(rows(rows.Count - 1).QuerySelector("input"))
    code = Q("#summaryCode").TextContent
    Ok "single-column code", InStr(code, "Descriptive Statistics Table for Amount") > 0
    Ok "typed a second column", TypeText(rows(1).QuerySelector("input"), "Price")
    Ok "multi-column toast", InStr(ToastText(), "multi-column") > 0 And ToastShown(), ToastText()
    Ok "multi-column code", InStr(Q("#summaryCode").TextContent, "Descriptive Statistics Table""") > 0 Or InStr(Q("#summaryCode").TextContent, "caption = ""Descriptive Statistics Table""") > 0
    Set cb = Q("#summaryStatisticFields input[value=""min""]")
    Ok "statistic toggle drawn", Drawn(cb)
    Ok "ticked Minimum", Tick(cb) And cb.Checked
    Ok "Minimum added to the code", InStr(Q("#summaryCode").TextContent, "min = c(") > 0
    Ok "CheckBox shows ticked", CBool(CtlOf(cb).Value)
    keys = Array("n", "mean", "sd", "median", "skew", "kurtosis", "min")
    For Each k In keys
        Set cb = Q("#summaryStatisticFields input[value=""" & k & """]")
        If cb.Checked Then Tick cb
    Next k
    Ok "no statistic selected", InStr(Q("#summaryCode").TextContent, "Select at least one statistic") > 0 And Q("#copySummaryCodeBtn").Disabled
    DisabledCopy Q("#copySummaryCodeBtn"), "copy (no statistic)"
    For Each k In Array("n", "mean", "sd", "median", "skew", "kurtosis")
        Tick Q("#summaryStatisticFields input[value=""" & k & """]")
    Next k
    code = Q("#summaryCode").TextContent
    Ok "defaults back (vars n mean sd median skew kurtosis)", InStr(code, "kurtosis = c(") > 0 And InStr(code, "min = c(") = 0
    Press QA("#columnFields .column-row")(1).QuerySelector("button")
    Ok "removing a column returns to single-column", InStr(ToastText(), "single-column") > 0, ToastText()
End Sub

Private Sub ScenarioAxis()
    Dim tg As clsNode, list As clsList, i As Long, clip As String, allOn As Boolean
    OpenOp "ggplot"
    Press Q("#loadGgplotSampleBtn")
    Set tg = Q("#ggplotBench .plot-axis-interval-toggle")
    Ok "axis-interval toggle drawn", Drawn(tg)
    Ok "off by default", Not tg.Checked And Not modShared.CalculatePlotAxisIntervals
    clip = CopyOnce(Q("#copyGgplotCodeBtn"), "axis off copy")
    Ok "ticked Calculate axis intervals", Tick(tg) And tg.Checked
    Set list = QA(".plot-axis-interval-toggle")
    allOn = True
    For i = 0 To list.Count - 1
        If Not list(i).Checked Then allOn = False
    Next i
    Ok "all three toggles synchronised", allOn And modShared.CalculatePlotAxisIntervals, list.Count & " toggles"
    Ok "code regenerated with intervals", Q("#ggplotCode").TextContent <> clip
    clip = CopyOnce(Q("#copyGgplotCodeBtn"), "axis on copy")
    OpenOp "lm-predict"
    Ok "lm-predict toggle shows ticked", CBool(CtlOf(Q("#simpleRegressionBench .plot-axis-interval-toggle")).Value)
    Ok "unticked from lm-predict", Tick(Q("#simpleRegressionBench .plot-axis-interval-toggle"))
    Ok "toggle off everywhere", Not modShared.CalculatePlotAxisIntervals And Not Q("#ggplotBench .plot-axis-interval-toggle").Checked
End Sub

Private Sub ScenarioKmeans()
    Dim sc As clsNode, cancel As Boolean
    OpenOp "kmeans"
    Press Q("#loadKmeansSampleBtn")
    CanonAll ActiveBench()
    Set sc = Q("#kmeansScale")
    Ok "scaling toggle drawn and ticked", Drawn(sc) And sc.Checked
    Ok "scaled code", InStr(Q("#kmeansCode").TextContent, "X_scaled <- scale(X)") > 0
    Ok "unticked scaling", Tick(sc)
    Ok "warning modal opens", Not Q("#kmeansScaleAlert").HasClass("hidden")
    Ok "toggle stays ticked until confirmed", sc.Checked And CBool(CtlOf(sc).Value), "DOM " & sc.Checked & ", window " & CtlOf(sc).Value
    Ok "focus on Keep scaling", modRender.FocusedNode() Is Q("#keepKmeansScalingBtn"), FocusName()
    Ok "modal buttons drawn", Drawn(Q("#keepKmeansScalingBtn")) And Drawn(Q("#disableKmeansScalingBtn"))
    cancel = KeyOn(Q("#keepKmeansScalingBtn"), 27)
    Ok "Escape keeps the decision modal open", Not Q("#kmeansScaleAlert").HasClass("hidden")
    cancel = KeyOn(Q("#keepKmeansScalingBtn"), 9)
    Ok "Tab moves to Disable scaling", modRender.FocusedNode() Is Q("#disableKmeansScalingBtn"), FocusName()
    cancel = KeyOn(Q("#disableKmeansScalingBtn"), 9, 1)
    Ok "Shift+Tab back to Keep scaling", modRender.FocusedNode() Is Q("#keepKmeansScalingBtn"), FocusName()
    Ok "Keep scaling pressed", Press(Q("#keepKmeansScalingBtn"))
    Ok "kept: modal closed, still scaled", Q("#kmeansScaleAlert").HasClass("hidden") And sc.Checked And InStr(Q("#kmeansCode").TextContent, "X_scaled <- scale(X)") > 0
    Ok "focus returns to the toggle", modRender.FocusedNode() Is sc, FocusName()
    Tick sc
    Ok "Disable scaling pressed", Press(Q("#disableKmeansScalingBtn"))
    Ok "disabled: modal closed, toggle unticked in DOM and window", Q("#kmeansScaleAlert").HasClass("hidden") And Not sc.Checked And Not CBool(CtlOf(sc).Value)
    CanonAll ActiveBench()
    Ok "unscaled code uses X everywhere", InStr(Q("#kmeansCode").TextContent, "scale(X)") = 0 And InStr(Q("#kmeansCode").TextContent, "X_scaled") = 0
    Ok "re-ticking scales without a modal", Tick(sc) And sc.Checked And Q("#kmeansScaleAlert").HasClass("hidden")
    CanonAll ActiveBench()
    Ok "scaled code back", InStr(Q("#kmeansCode").TextContent, "X_scaled <- scale(X)") > 0
End Sub

Private Sub ScenarioKnownSd()
    Dim sd As clsNode, cancel As Boolean
    OpenOp "interval"
    Pick Q("#intervalCalcType"), 0
    Pick Q("#intervalParameter"), 0
    TypeText Q("#intervalDataset"), "dataset"
    TypeText Q("#intervalLabel"), "Amount"
    Set sd = Q("#intervalStdev")
    Ok "standard-deviation dropdown drawn", Drawn(sd)
    Ok "picked known standard deviation", Pick(sd, 1)
    Ok "confirmation modal opens", Not Q("#knownPopulationSdAlert").HasClass("hidden")
    Ok "value stays unknown until confirmed", sd.Value = "unknown" And CtlOf(sd).ListIndex = 0, sd.Value & " / " & CtlOf(sd).ListIndex
    Ok "focus on Use unknown", modRender.FocusedNode() Is Q("#useUnknownSdBtn"), FocusName()
    cancel = KeyOn(Q("#useUnknownSdBtn"), 27)
    Ok "Escape keeps the modal open", Not Q("#knownPopulationSdAlert").HasClass("hidden")
    cancel = KeyOn(Q("#useUnknownSdBtn"), 9)
    Ok "Tab cycles to the confirm button", modRender.FocusedNode() Is Q("#confirmKnownSdBtn"), FocusName()
    Ok "pressed Use unknown", Press(Q("#useUnknownSdBtn"))
    Ok "unknown kept", Q("#knownPopulationSdAlert").HasClass("hidden") And sd.Value = "unknown"
    Ok "focus returns to the dropdown", modRender.FocusedNode() Is sd, FocusName()
    Pick sd, 1
    Ok "pressed Yes, population SD is known", Press(Q("#confirmKnownSdBtn"))
    Ok "known selected in DOM and window", sd.Value = "known" And CtlOf(sd).ListIndex = 1, sd.Value & " / " & CtlOf(sd).ListIndex
    Ok "population SD field drawn", EnsureShown(Q("#intervalSd"))
    TypeText Q("#intervalSd"), "4"
    Ok "code uses qnorm for known sigma", InStr(Q("#intervalOutput").TextContent, "qnorm") > 0, Left$(Q("#intervalOutput").TextContent, 120)
    Pick Q("#intervalParameter"), 1
    Ok "proportion disables the SD choice", sd.Disabled Or sd.Value = "na", sd.Value
    Pick Q("#intervalParameter"), 0
    Pick sd, 0
End Sub

Private Sub ScenarioLpRows()
    Dim rows As clsList, n As Long, last As clsNode, nm As clsNode, codeBefore As String, rm As clsNode, c As Object
    OpenOp "lpsolve"
    Ok "Reset all pressed", Press(Q("#resetLpBtn"))
    Set rows = QA(".lp-constraint-row")
    n = rows.Count
    Ok "reset leaves one blank row", n = 1, n & " rows"
    Ok "decision variables typed", TypeText(Q("#lpVariables"), "X1, X2, X3")
    Ok "objective typed", TypeText(Q("#lpObjective"), "5, 4, 3")
    Set last = rows(n - 1)
    Ok "constraint name typed key by key", TypeChars(last.QuerySelector(".lp-name"), "Budget")
    Set rows = QA(".lp-constraint-row")
    Ok "a trailing blank row appended while typing", rows.Count = n + 1, rows.Count & " rows"
    Ok "the new row is drawn", Drawn(rows(rows.Count - 1).QuerySelector(".lp-name"))
    Ok "only remove controls (no add button / steppers)", Q("#lpBench .lp-add-constraint") Is Nothing
    TypeText last.QuerySelector(".lp-coefficients"), "1, 2"
    TypeText last.QuerySelector(".lp-rhs"), "10"
    Ok "wrong coefficient count highlighted", last.QuerySelector(".lp-coefficients").HasClass("input-error"), Q("#lpConstraintError").TextContent
    Ok "copy disabled while invalid", Q("#copyLpCodeBtn").Disabled
    DisabledCopy Q("#copyLpCodeBtn"), "copy (invalid row)"
    TypeText last.QuerySelector(".lp-coefficients"), "1, 2, 3"
    Ok "valid row accepted", Not Q("#copyLpCodeBtn").Disabled
    Ok "constraint on its own code line", InStr(Q("#lpBench").TextContent, "1, 2, 3") > 0
    Pick last.QuerySelector(".lp-operator"), 1
    Ok "operator picked through the dropdown", last.QuerySelector(".lp-operator").Value = ">=", last.QuerySelector(".lp-operator").Value
    TypeChars rows(rows.Count - 1).QuerySelector(".lp-name"), "Z"
    Set rows = QA(".lp-constraint-row")
    n = rows.Count
    Set rm = rows(n - 2).QuerySelector(".lp-remove-constraint")
    Ok "remove control drawn", Drawn(rm)
    Ok "remove pressed", Press(rm)
    Ok "row removed", QA(".lp-constraint-row").Count = n - 1, QA(".lp-constraint-row").Count & " rows"
    Press Q("#loadLpSampleBtn")
    Ok "sample loads several rows", QA(".lp-constraint-row").Count > 2
    CopyOnce Q("#copyLpCodeBtn"), "lp code"
    Ok "code card done", Not Q("#copyLpCodeBtn").Closest(".lp-collapsible") Is Nothing
    TypeText Q("#lpObjective"), "0.25, 0.40, 0.60"
    Ok "model change clears every Done!", Q("#lpBench .completion-badge") Is Nothing And QA("#lpBench .lp-collapsible.is-collapsed").Count = 0, QA("#lpBench .completion-badge").Count & " badges, " & QA("#lpBench .lp-collapsible.is-collapsed").Count & " collapsed"
End Sub

Private Sub ScenarioLpSensitivity()
    Dim list As clsList, i As Long, b As clsNode, panel As clsNode, copied As Long
    OpenOp "lpsolve"
    Press Q("#loadLpSampleBtn")
    If Q("#lpSensitivity").Checked = False Then Tick Q("#lpSensitivity")
    Ok "sensitivity sample pressed", Press(Q("#loadSensitivitySampleBtn"))
    Set list = QA("#sensitivityInterpretation [data-sensitivity-copy]")
    Ok "four section copy buttons", list.Count >= 4, list.Count & " sections"
    Set panel = Q("#sensitivityInterpretationPanel")
    For i = 0 To list.Count - 1
        Set b = list(i)
        CopyOnce b, "section " & NullText(b.GetAttribute("data-sensitivity-copy"))
        Ok "section " & NullText(b.GetAttribute("data-sensitivity-copy")) & " marked Copied", Trim$(b.TextContent) = "Copied" And b.Closest(".sensitivity-prose-card").HasClass("is-copy-complete"), Trim$(b.TextContent)
        If Not panel Is Nothing And i < list.Count - 1 Then Ok "card not complete after one section", Not panel.HasClass("is-complete")
    Next i
    If Not panel Is Nothing Then Ok "card completes after every section", panel.HasClass("is-complete") And panel.HasClass("is-collapsed")
    EnsureShown Q("#sensDuals")
    TypeText Q("#sensDuals"), Q("#sensDuals").Value & " "
    Ok "editing sensitivity resets section tracking", QA("#sensitivityInterpretation .is-copy-complete").Count = 0 And Not panel.HasClass("is-complete")
    CopyOnce Q("#copySensitivityProseBtn"), "copy prose (all sections)"
    Ok "Copy prose completes the card", panel.HasClass("is-complete")
End Sub

Private Sub ScenarioAnova()
    Dim vp As clsNode, sp As clsNode, wp As clsNode, route As String
    OpenOp "anova-many"
    Ok "sample pressed", Press(Q("#loadAnovaSampleBtn"))
    Set vp = Q("#anovaVariancePanel")
    Set sp = Q("#anovaStandardPanel")
    Set wp = Q("#anovaWelchPanel")
    Ok "variance output sample pressed", EnsureShown(Q("#loadAnovaVarianceOutputBtn")) And Press(Q("#loadAnovaVarianceOutputBtn"))
    CopyOnce Q("#copyAnovaVarianceCodeBtn"), "step 2 code"
    Ok "step 2: code box done, card not yet complete", Q("#anovaVarianceCodeBox").HasClass("is-copy-complete") And Not vp.HasClass("is-complete")
    CopyOnce Q("#copyAnovaVarianceProseBtn"), "step 2 prose"
    Ok "step 2: card complete + collapsed after code and prose", vp.HasClass("is-complete") And vp.HasClass("is-collapsed") And BadgeText(vp) = "Done!"
    If sp.HasClass("is-unavailable") Then route = "welch" Else route = "standard"
    Ok "one route greyed", sp.HasClass("is-unavailable") Xor wp.HasClass("is-unavailable"), "standard " & sp.HasClass("is-unavailable") & ", welch " & wp.HasClass("is-unavailable")
    If route = "standard" Then
        Ok "step 3a sample pressed", EnsureShown(Q("#loadAnovaStandardOutputBtn")) And Press(Q("#loadAnovaStandardOutputBtn"))
        CopyOnce Q("#copyAnovaStandardProseBtn"), "step 3a prose"
        Ok "step 3a: prose box done, card not complete", Q("#anovaStandardProseBox").HasClass("is-copy-complete") And Not sp.HasClass("is-complete")
        CopyOnce Q("#copyAnovaStandardCodeBtn"), "step 3a code"
        Ok "step 3a: complete after both", sp.HasClass("is-complete") And sp.HasClass("is-collapsed")
    Else
        Ok "step 3b sample pressed", EnsureShown(Q("#loadAnovaWelchOutputsBtn")) And Press(Q("#loadAnovaWelchOutputsBtn"))
        CopyOnce Q("#copyAnovaWelchProseBtn"), "step 3b prose"
        Ok "step 3b: prose box done, card not complete", Q("#anovaWelchProseBox").HasClass("is-copy-complete") And Not wp.HasClass("is-complete")
        CopyOnce Q("#copyAnovaWelchCodeBtn"), "step 3b code"
        Ok "step 3b: complete after both", wp.HasClass("is-complete") And wp.HasClass("is-collapsed")
    End If
    EnsureShown Q("#anovaResponse")
    TypeText Q("#anovaResponse"), "Sales"
    Ok "editing Step 0 clears completion", Not vp.HasClass("is-complete") And Not Q("#anovaVarianceCodeBox").HasClass("is-copy-complete"), "variance '" & vp.ClassName & "'"
    Ok "Step 0 name reaches the code", InStr(Q("#anovaBench").TextContent, "Sales") > 0
    TypeText Q("#anovaResponse"), "Amount"
    Ok "normality toggle drawn", EnsureShown(Q("#anovaNormalityAssumed"))
    Tick Q("#anovaNormalityAssumed")
    Ok "unticking normality enables Step 1", Not Q("#anovaNormalityAssumed").Checked And Not Q("#anovaNormalityPanel").HasClass("is-unavailable"), NullText(Q("#anovaNormalityPanel").GetAttribute("class"))
    Tick Q("#anovaNormalityAssumed")
    Ok "ticking normality greys Step 1", Q("#anovaNormalityAssumed").Checked
End Sub

Private Sub ScenarioUnits()
    OpenOp "multi-regression"
    Ok "regression output sample pressed", EnsureShown(Q("#loadMultiRegressionOutputBtn")) And Press(Q("#loadMultiRegressionOutputBtn"))
    Ok "copy all prose enabled", Not Q("#copyMultiRegressionProseBtn").Disabled
    CopyOnce Q("#copyMultiRegressionProseBtn"), "regression all prose"
    If Not Q("[data-copy-multi-regression]") Is Nothing Then CopyOnce Q("[data-copy-multi-regression]"), "regression predictor prose"
    If Not Q("[data-copy-multi-model]") Is Nothing Then CopyOnce Q("[data-copy-multi-model]"), "regression model prose"
End Sub

' Answer selections on revisits (untouched: re-randomised; edited: kept exactly) and the
' toast expiring when the window's timer fires.
Private Sub ScenarioRevisit()
    Dim sel As clsNode, v1 As String, c1 As String, v2 As String, c2 As String, changed As Long, k As Long, i As Long
    For k = 1 To 4
        OpenOp "kfold"
        Set sel = modVariants.PickerForCopyButton(Q("#copyKfoldCodeBtn")).QuerySelector("select")
        If k > 1 Then
            If sel.Value <> v1 Or Q("#kfoldCode").TextContent <> c1 Then changed = changed + 1
        End If
        v1 = sel.Value: c1 = Q("#kfoldCode").TextContent
        OpenOp "shapiro"
    Next k
    Soft "untouched revisits re-randomise the answer (3 revisits)", changed > 0, changed & " of 3 changed"
    OpenOp "kfold"
    Ok "sample pressed", Press(Q("#loadKfoldSampleBtn"))
    Ok "typed into the dataset field", TypeText(Q("#kfoldDataset"), "mydata")
    Set sel = modVariants.PickerForCopyButton(Q("#copyKfoldCodeBtn")).QuerySelector("select")
    Ok "picked Option 4", Pick(sel, 3)
    v2 = sel.Value: c2 = Q("#kfoldCode").TextContent
    OpenOp "shapiro"
    OpenOp "pearson"
    OpenOp "kfold"
    Set sel = modVariants.PickerForCopyButton(Q("#copyKfoldCodeBtn")).QuerySelector("select")
    Ok "edited operation keeps its option on return", sel.Value = v2 And CtlOf(sel).ListIndex = CLng(v2) - 1, sel.Value & " vs " & v2
    Ok "edited operation keeps its exact code on return", Q("#kfoldCode").TextContent = c2
    Press Q("#copyKfoldCodeBtn")
    Ok "toast shown after copying", ToastShown()
    For i = 1 To 6
        If Not ToastShown() Then Exit For
        modRender.TimerTick
    Next i
    Ok "toast expires when the window timer fires", Not ToastShown(), i & " ticks"
    TypeText Q("#kfoldDataset"), "dataset"
    ' Number steppers (the arrow buttons beside number fields) and arrow keys in the field.
    Dim k0 As String, steps As clsList
    k0 = Q("#kfoldK").Value
    Set steps = Q("#kfoldK").ParentElement.QuerySelectorAll(".number-step-button")
    Ok "number field has two stepper buttons", steps.Count = 2, steps.Count & " buttons"
    If steps.Count = 2 Then
        Ok "stepper up pressed", Press(steps(1))
        Ok "stepper up increments", Q("#kfoldK").Value = CStr(CLng(k0) + 1), Q("#kfoldK").Value
        Ok "TextBox shows the stepped value", CtlOf(Q("#kfoldK")).Text = Q("#kfoldK").Value, CtlOf(Q("#kfoldK")).Text
        Ok "code follows the stepped value", InStr(Q("#kfoldCode").TextContent, CStr(CLng(k0) + 1)) > 0
        Press steps(0)
        Ok "stepper down decrements", Q("#kfoldK").Value = k0, Q("#kfoldK").Value
    End If
    KeyOn Q("#kfoldK"), 38
    Soft "ArrowUp in a number field increments (browser behaviour)", Q("#kfoldK").Value = CStr(CLng(k0) + 1), Q("#kfoldK").Value
    If Q("#kfoldK").Value <> k0 Then TypeText Q("#kfoldK"), k0
End Sub
