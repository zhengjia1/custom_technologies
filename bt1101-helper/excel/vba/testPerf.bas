Attribute VB_Name = "testPerf"
' Test-only: headless profile of the DOM side of every user action (no window). Boots the
' whole app like testDriver, then for every operation times: opening it, each "Use sample"
' button / sample option, typing one character into each visible text field, toggling
' each visible checkbox/select, changing each answer option and clicking each copy
' button -- each measured from the dispatch to the end of modDom.Settle 700.
'
'   node tools/vba.js run testPerf.ProfileAll <path to .build/body.html> [passes] --only ... --no-form
'
' Returns one line per action (ms, operation, action, element) and a summary of the
' slowest actions per category.
Option Explicit

#If Mac Then
#Else
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef counter As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef frequency As Currency) As Long
#End If

Private mFreq As Currency
Private mRows As clsList
Private mOp As String
Private mMaxPerKind As Long
Private mStage As String

' Milliseconds from a high-resolution clock (Timer on Mac).
Public Function PerfNowMs() As Double
#If Mac Then
    PerfNowMs = Timer * 1000#
#Else
    Dim c As Currency
    If mFreq = 0 Then QueryPerformanceFrequency mFreq
    QueryPerformanceCounter c
    PerfNowMs = CDbl(c) * 1000# / CDbl(mFreq)
#End If
End Function

Public Function ProfileAll(ByVal bodyPath As String, Optional ByVal passes As String = "1") As String
    Dim t0 As Double, out As String, pass As Long, nPass As Long
    On Error GoTo Failed
    nPass = CLng(Val(passes))
    If nPass < 1 Then nPass = 1
    mMaxPerKind = 12
    JsSeedRandom 1
    LastClipboardText = ""
    LastErrors = ""
    t0 = PerfNowMs()
    LoadDocument ReadUtf8File(bodyPath)
    out = "boot.load" & vbTab & Format$(PerfNowMs() - t0, "0.0") & vbLf
    t0 = StartTiming()
    Application.Run "InitApp"
    Settle 700
    out = out & "boot.init" & vbTab & Format$(PerfNowMs() - t0, "0.0") & vbTab & PhaseText() & vbLf
    For pass = 1 To nPass
        Set mRows = New clsList
        ProfilePass
        out = out & "===== pass " & pass & vbLf & Report()
    Next pass
    If Len(LastErrors) > 0 Then out = out & "VBAERROR" & vbTab & LastErrors & vbLf
    out = out & "===== phases (whole run)" & vbLf & GrandPhases()
    ProfileAll = out
    Exit Function
Failed:
    ProfileAll = out & "FAILED: " & Err.Description & " (" & mOp & " / " & mStage & ")"
End Function

Private Sub ProfilePass()
    Dim ops As clsList, i As Long, id As String
    Set ops = New clsList
    ops.Add "universal"
    For i = 0 To modApp.OpCount - 1
        ops.Add modApp.OpIds(i)
    Next i
    For i = 0 To ops.Count - 1
        id = ops(i)
        ProfileOperation id
    Next i
End Sub

Private Function EntryFor(ByVal id As String) As clsNode
    If id = "universal" Then
        Set EntryFor = Q(".universal-entry")
    Else
        Set EntryFor = Q("#tutorialList .operation[data-id=""" & id & """]")
    End If
End Function

Private Function TimeClick(ByVal el As clsNode) As Double
    Dim t As Double
    t = StartTiming()
    el.Click
    Settle 700
    TimeClick = PerfNowMs() - t
End Function

Private Function TimeEvent(ByVal el As clsNode, ByVal evType As String) As Double
    Dim t As Double
    t = StartTiming()
    el.DispatchEvent NewEvent(evType, True)
    Settle 700
    TimeEvent = PerfNowMs() - t
End Function

Private Sub Record(ByVal kind As String, ByVal what As String, ByVal ms As Double)
    mRows.Add Format$(ms, "0.0") & vbTab & mOp & vbTab & kind & vbTab & what & vbTab & PhaseText()
End Sub

Private Function Describe(ByVal el As clsNode) As String
    If Len(el.Id) > 0 Then
        Describe = "#" & el.Id
    Else
        Describe = el.LocalName & "." & Replace(el.ClassName, " ", ".")
        If Not el.ParentElement Is Nothing Then
            If Len(el.ParentElement.Id) > 0 Then Describe = "#" & el.ParentElement.Id & ">" & Describe
        End If
    End If
End Function

Private Function Shown(ByVal el As clsNode) As Boolean
    Dim n As clsNode, prev As clsNode
    Set n = el
    Do While Not n Is Nothing
        If n.NodeType = 1 Then
            If n.Hidden Then Exit Function
            If n.HasClass("hidden") Then Exit Function
            If n.GetStyle("display") = "none" Then Exit Function
            If n.LocalName = "details" And Not prev Is Nothing Then
                If Not n.IsOpen And prev.LocalName <> "summary" Then Exit Function
            End If
        End If
        Set prev = n
        Set n = n.ParentNode
    Loop
    Shown = True
End Function

Private Function ShownIn(ByVal selector As String) As clsList
    Dim all As clsList, i As Long, el As clsNode, res As New clsList, main As clsNode
    Set main = Q("main")
    If main Is Nothing Then Set main = Doc
    Set all = main.QuerySelectorAll(selector)
    For i = 0 To all.Count - 1
        Set el = all(i)
        If Shown(el) And Not el.Disabled Then res.Add el
    Next i
    Set ShownIn = res
End Function

Private Sub ProfileOperation(ByVal id As String)
    Dim entry As clsNode, list As clsList, i As Long, el As clsNode, ms As Double, old As String, k As Long, n As Long
    Dim txt As String
    mOp = id
    Set entry = EntryFor(id)
    If entry Is Nothing Then Record "open", "MISSING", 0: Exit Sub
    mStage = "open"
    Record "open", "", TimeClick(entry)
    mStage = "samples"
    ' sample buttons and the detector's sample select
    Set list = ShownIn("button.sample-btn")
    n = 0
    For i = 0 To list.Count - 1
        If n >= mMaxPerKind Then Exit For
        Set el = list(i)
        If Shown(el) Then
            Record "sample", Describe(el), TimeClick(el)
            n = n + 1
        End If
    Next i
    If id = "universal" Then
        Set el = Q("#universalSampleSelect")
        n = el.Options.Count
        For k = 1 To n - 1
            mOp = "universal"
            EntryFor("universal").Click
            Settle 700
            Set el = Q("#universalSampleSelect")
            If k >= el.Options.Count Then Exit For
            el.SelectedIndex = k
            ms = StartTiming()
            el.DispatchEvent NewEvent("change", True)
            Settle 700
            mStage = "usample" & k
            Record "sample", "#universalSampleSelect[" & k & "]", PerfNowMs() - ms
        Next k
        EntryFor("universal").Click
        Settle 700
        Q("#universalSampleSelect").SelectedIndex = 1
        Q("#universalSampleSelect").DispatchEvent NewEvent("change", True)
        Settle 700
    End If
    mStage = "type"
    ' typing one character into each visible text field
    Set list = ShownIn("textarea, input")
    n = 0
    For i = 0 To list.Count - 1
        If n >= mMaxPerKind Then Exit For
        Set el = list(i)
        Select Case LCase$(el.TypeName)
        Case "checkbox", "radio", "button", "submit", "hidden", "range"
        Case Else
            If el.ReadOnly Or Not Shown(el) Then GoTo NextField
            If el.Id = "librarySearch" Or el.Id = "operationSearch" Then GoTo NextField
            old = el.Value
            If LCase$(el.TypeName) = "number" Then txt = old & "1" Else txt = old & "x"
            el.Value = txt
            Record "type", Describe(el), TimeEvent(el, "input")
            el.Value = old
            el.DispatchEvent NewEvent("input", True)
            Settle 700
            n = n + 1
        End Select
NextField:
    Next i
    mStage = "toggle"
    ' checkboxes
    Set list = ShownIn("input[type=checkbox]")
    n = 0
    For i = 0 To list.Count - 1
        If n >= mMaxPerKind Then Exit For
        Set el = list(i)
        If Shown(el) Then
            Record "toggle", Describe(el), TimeClick(el)
            el.Click
            Settle 700
            n = n + 1
        End If
    Next i
    mStage = "select"
    ' selects (not answer options / samples)
    Set list = ShownIn("select")
    n = 0
    For i = 0 To list.Count - 1
        If n >= mMaxPerKind Then Exit For
        Set el = list(i)
        If el.HasClass("answer-variant-select") Or el.Id = "universalSampleSelect" Then GoTo NextSelect
        If Not Shown(el) Then GoTo NextSelect
        If el.Options.Count < 2 Then GoTo NextSelect
        k = el.SelectedIndex
        el.SelectedIndex = (k + 1) Mod el.Options.Count
        ms = StartTiming()
        el.DispatchEvent NewEvent("input", True)
        el.DispatchEvent NewEvent("change", True)
        Settle 700
        Record "select", Describe(el), PerfNowMs() - ms
        el.SelectedIndex = k
        el.DispatchEvent NewEvent("input", True)
        el.DispatchEvent NewEvent("change", True)
        Settle 700
        n = n + 1
NextSelect:
    Next i
    mStage = "option"
    ' answer options: two changes per control (to the next option, then option 2)
    Set list = ShownIn(".answer-variant-select")
    n = 0
    For i = 0 To list.Count - 1
        If n >= mMaxPerKind Then Exit For
        Set el = list(i)
        If Shown(el) Then
            k = CLng(Val(el.Value))
            el.Value = CStr((k Mod 6) + 1)
            Record "option", Describe(el) & " ->" & el.Value, TimeEvent(el, "change")
            el.Value = "1"
            Record "option", Describe(el) & " ->1", TimeEvent(el, "change")
            n = n + 1
        End If
    Next i
    mStage = "copy"
    ' copy buttons
    Set list = ShownIn("button")
    n = 0
    For i = 0 To list.Count - 1
        If n >= mMaxPerKind Then Exit For
        Set el = list(i)
        If Left$(DomTrim(el.TextContent), 4) = "Copy" And Shown(el) Then
            Record "copy", Describe(el), TimeClick(el)
            DismissModals
            n = n + 1
        End If
    Next i
    DismissModals
End Sub

Private Sub DismissModals()
    Dim list As clsList, i As Long, el As clsNode
    Set list = QA(".variable-rename-dismiss")
    For i = 0 To list.Count - 1
        Set el = list(i)
        If Shown(el) Then el.Click: Settle 700
    Next i
End Sub

' Phase breakdown from modProf when the sources were instrumented (test/perf.js --phases).
Private Function StartTiming() As Double
    On Error Resume Next
    Application.Run "modProf.ProfReset"
    On Error GoTo 0
    StartTiming = PerfNowMs()
End Function

Private Function PhaseText() As String
    On Error Resume Next
    PhaseText = Application.Run("modProf.ProfTake", 16)
End Function

Private Function GrandPhases() As String
    On Error Resume Next
    GrandPhases = Application.Run("modProf.ProfGrand", 400)
End Function

' All rows, then the slowest per kind.
Private Function Report() As String
    Dim out As String, i As Long, kinds As Variant, kind As Variant, parts() As String
    Dim ms() As Double, idx() As Long, n As Long, j As Long, tmp As Double, ti As Long, total As Double
    n = mRows.Count
    If n = 0 Then Exit Function
    ReDim ms(0 To n - 1): ReDim idx(0 To n - 1)
    For i = 0 To n - 1
        parts = Split(mRows(i), vbTab)
        ms(i) = Val(parts(0))
        idx(i) = i
        total = total + ms(i)
        out = out & mRows(i) & vbLf
    Next i
    ' sort indices by ms descending (insertion sort; a few hundred rows)
    For i = 1 To n - 1
        ti = idx(i)
        j = i - 1
        Do While j >= 0
            If ms(idx(j)) >= ms(ti) Then Exit Do
            idx(j + 1) = idx(j)
            j = j - 1
        Loop
        idx(j + 1) = ti
    Next i
    out = out & "----- total " & Format$(total, "0") & " ms over " & n & " actions" & vbLf
    kinds = Array("open", "sample", "type", "toggle", "select", "option", "copy")
    For Each kind In kinds
        out = out & "----- slowest " & kind & vbLf
        j = 0
        For i = 0 To n - 1
            parts = Split(mRows(idx(i)), vbTab)
            If parts(2) = kind Then
                out = out & mRows(idx(i)) & vbLf
                j = j + 1
                If j >= 8 Then Exit For
            End If
        Next i
    Next kind
    Report = out
End Function

