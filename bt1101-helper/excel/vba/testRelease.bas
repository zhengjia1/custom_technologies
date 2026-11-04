Attribute VB_Name = "testRelease"
' Test-only: walks the release workbook as a user would. The app is started by
' modBoot.StartApp (stored markup, CSS and fonts), then every operation is opened, its
' first sample loaded, and the app window alone captured with PrintWindow.
Option Explicit

Public Function WalkAll(ByVal scriptPath As String, ByVal outDir As String, ByVal width As Double, ByVal theme As String) As String
    Dim entries As clsList, ids As clsList, i As Long, id As String, out As String, sample As clsNode, file As String, stage As String, sampleNote As String
    On Error GoTo Failed
    stage = "start": If Not AppRunning() Then StartApp
    stage = "window": SetWindow width
    stage = "theme": SetTheme theme
    stage = "entries"
    Set ids = NewList()
    ids.Push "universal"
    Set entries = QA("#tutorialList .operation[data-id]")
    For i = 0 To entries.Count - 1
        ids.Push CStr(entries(i).DataGet("id"))
    Next i
    For i = 0 To ids.Count - 1
        id = ids(i)
        LastErrors = ""
        stage = "open " & id: Open1 id
        stage = "sample " & id: Set sample = FirstSampleButton()
        If Not sample Is Nothing Then sample.Click: Settle 700
        stage = "render " & id: modRender.RenderNow
        file = outDir & Application.PathSeparator & id & "-" & theme & "-" & CStr(width) & ".png"
        If sample Is Nothing Then sampleNote = " (no sample button)" Else sampleNote = " (sample " & sample.Id & ")"
        out = out & id & ": " & IIf(Capture(scriptPath, file), "captured", "NOT captured") & sampleNote & _
            " render " & Format$(modRender.LastRenderMs, "0") & " ms" & IIf(Len(LastErrors) > 0, " ERRORS: " & Replace(LastErrors, vbLf, " | "), "") & vbLf
    Next i
    WalkAll = out
    Exit Function
Failed:
    WalkAll = "ERROR at " & stage & ": " & Err.Description & vbLf & out & LastErrors
End Function

Private Sub Open1(ByVal id As String)
    Dim el As clsNode
    If id = "universal" Then Set el = Q(".universal-entry") Else Set el = Q("#tutorialList .operation[data-id=""" & id & """]")
    If Not el Is Nothing Then el.Click
    Settle 700
    modRender.RenderNow
End Sub

' The visible bench's first "Use sample"-style button (or the detector's sample list).
Private Function FirstSampleButton() As clsNode
    Dim list As clsList, i As Long, b As clsNode, bid As String, sel As clsNode
    Set list = QA("main button")
    For i = 0 To list.Count - 1
        Set b = list(i)
        bid = b.Id
        If (bid = "singleSampleBtn" Or (Left$(bid, 4) = "load" And InStr(bid, "Sample") > 0)) And b.Closest(".hidden") Is Nothing And Not b.Disabled Then
            Set FirstSampleButton = b
            Exit Function
        End If
    Next i
    Set sel = Q("#universalSampleSelect")
    If AppState("id") = "universal" And Not sel Is Nothing Then
        sel.SelectedIndex = 1
        sel.DispatchEvent NewEvent("change", True)
        Settle 700
    End If
End Function

Private Sub SetWindow(ByVal width As Double)
    modRender.SetClientSize width, 860
    modRender.ResizeNow
    modRender.RenderNow
End Sub

Private Sub SetTheme(ByVal theme As String)
    Dim cur As Variant
    cur = DocumentElement.GetAttribute("data-theme")
    If IsNull(cur) Then cur = "light"
    If CStr(cur) <> theme Then Q("#themeToggle").Click
    Settle 700
    modRender.RenderNow
End Sub

Private Function Capture(ByVal scriptPath As String, ByVal outPath As String) As Boolean
    Dim h As LongPtr, t As Double
    t = Timer
    Do While Timer - t < 0.6
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
