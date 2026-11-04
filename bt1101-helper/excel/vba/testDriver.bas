Attribute VB_Name = "testDriver"
' Test-only: runs a step script (see test/ref-run.js) against the VBA port and returns
' the same records the Node runner prints, for test/compare.js.
' Extra header steps: "load <body.html path>" and "boot <Module.Procedure>" (default InitApp).
Option Explicit

Public Function RunStepsFile(ByVal path As String) As String
    RunStepsFile = RunSteps(ReadUtf8File(path))
End Function

' Runs each steps file named in the list file (one path per line) on a fresh document.
Public Function RunStepsList(ByVal listPath As String) As String
    Dim paths() As String, i As Long, out As String
    paths = Split(Replace(ReadUtf8File(listPath), vbCrLf, vbLf), vbLf)
    For i = 0 To UBound(paths)
        If Len(Trim$(paths(i))) > 0 Then out = out & vbLf & "=====FILE=====" & vbLf & RunStepsFile(Trim$(paths(i)))
    Next i
    RunStepsList = out
End Function

Public Function RunSteps(ByVal stepsText As String) As String
    Dim lines() As String, i As Long, out As clsList, line As String, parts() As String
    Dim command As String, args As clsList, j As Long, booted As Boolean, bodyPath As String, bootProc As String
    Set out = NewList()
    bootProc = "InitApp"
    lines = Split(Replace(stepsText, vbCrLf, vbLf), vbLf)
    For i = 0 To UBound(lines)
        line = lines(i)
        If Len(Trim$(line)) = 0 Or Left$(line, 1) = "#" Then GoTo NextLine
        parts = Split(line, vbTab)
        command = parts(0)
        Set args = NewList()
        For j = 1 To UBound(parts)
            args.Push Unescape(parts(j))
        Next j
        If command = "load" Then
            bodyPath = args(0)
            GoTo NextLine
        ElseIf command = "boot" Then
            bootProc = args(0)
            GoTo NextLine
        End If
        If Not booted Then
            JsSeedRandom 1
            LastClipboardText = ""
            LastErrors = ""
            LoadDocument ReadUtf8File(bodyPath)
            Application.Run bootProc
            Settle 700
            booted = True
        End If
        On Error GoTo StepFailed
        RunStep command, args, line, out
        On Error GoTo 0
NextLine:
    Next i
    ' Errors that handlers trapped and reported would otherwise pass silently.
    If Len(LastErrors) > 0 Then out.Push "VBAERROR" & vbTab & EscapeStep(LastErrors)
    RunSteps = out.Join(vbLf)
    Exit Function
StepFailed:
    out.Push line & vbTab & EscapeStep("ERROR line " & (i + 1) & " (" & command & "): " & Err.Description)
    Resume NextLine
End Function

Private Sub RunStep(ByVal command As String, ByVal args As clsList, ByVal line As String, ByVal out As clsList)
    Dim el As clsNode, entry As clsNode, list As clsList, k As Long, evt As clsEvent, mods As String, toast As clsNode
    Select Case command
        Case "open"
            If args(0) = "universal" Then Set entry = Q(".universal-entry") Else Set entry = Q("#tutorialList .operation[data-id=""" & args(0) & """]")
            If entry Is Nothing Then Err.Raise vbObjectError + 1, , "No operation " & args(0)
            entry.Click
        Case "value"
            One(args(0)).Value = args(1)
        Case "input"
            Set el = One(args(0))
            If args.Count > 1 Then el.Value = args(1)
            el.DispatchEvent NewEvent("input", True)
        Case "change"
            Set el = One(args(0))
            If args.Count > 1 Then el.Value = args(1)
            el.DispatchEvent NewEvent("change", True)
        Case "select"
            Set el = One(args(0))
            el.SelectedIndex = CLng(args(1))
            el.DispatchEvent NewEvent("input", True)
            el.DispatchEvent NewEvent("change", True)
        Case "click"
            One(args(0)).Click
        Case "key"
            If args.Count > 2 Then mods = args(2)
            Set evt = NewKeyEvent("keydown", args(1), InStr(mods, "s") > 0, InStr(mods, "c") > 0, InStr(mods, "a") > 0)
            One(args(0)).DispatchEvent evt
        Case "canon"
            Set list = QA(".answer-variant-select")
            For k = 0 To list.Count - 1
                Set el = list(k)
                If el.Value <> "1" Then
                    el.Value = "1"
                    el.DispatchEvent NewEvent("change", True)
                End If
            Next k
        Case "read": out.Push line & vbTab & EscapeStep(One(args(0)).TextContent)
        Case "val": out.Push line & vbTab & EscapeStep(One(args(0)).Value)
        Case "class": out.Push line & vbTab & EscapeStep(NullToText(One(args(0)).GetAttribute("class"), ""))
        Case "attr": out.Push line & vbTab & EscapeStep(NullToText(One(args(0)).GetAttribute(args(1)), "<null>"))
        Case "checked": out.Push line & vbTab & IIf(One(args(0)).Checked, "1", "0")
        Case "disabled": out.Push line & vbTab & IIf(One(args(0)).Disabled, "1", "0")
        Case "count": out.Push line & vbTab & CStr(QA(args(0)).Count)
        Case "clip": out.Push line & vbTab & EscapeStep(IIf(Len(LastClipboardText) > 0, LastClipboardText, "<none>"))
        Case "toast"
            Set toast = ById("toast")
            out.Push line & vbTab & EscapeStep(IIf(toast.HasClass("show"), "shown: ", "hidden: ") & toast.TextContent)
        Case Else
            Err.Raise vbObjectError + 2, , "Unknown step"
    End Select
    Select Case command
        Case "read", "val", "class", "attr", "checked", "disabled", "count", "clip", "toast"
        Case Else: Settle 700
    End Select
End Sub

Private Function One(ByVal selector As String) As clsNode
    Set One = Q(selector)
    If One Is Nothing Then Err.Raise vbObjectError + 3, , "No element matches " & selector
End Function

Private Function NullToText(ByVal value As Variant, ByVal fallback As String) As String
    If IsNull(value) Then NullToText = fallback Else NullToText = CStr(value)
End Function

Private Function Unescape(ByVal text As String) As String
    Dim i As Long, ch As String, out As String
    i = 1
    Do While i <= Len(text)
        ch = Mid$(text, i, 1)
        If ch = "\" And i < Len(text) Then
            Select Case Mid$(text, i + 1, 1)
                Case "n": out = out & vbLf
                Case "t": out = out & vbTab
                Case "r": out = out & vbCr
                Case Else: out = out & Mid$(text, i + 1, 1)
            End Select
            i = i + 2
        Else
            out = out & ch
            i = i + 1
        End If
    Loop
    Unescape = out
End Function

Public Function EscapeStep(ByVal text As String) As String
    EscapeStep = Replace(Replace(Replace(Replace(text, "\", "\\"), vbLf, "\n"), vbTab, "\t"), vbCr, "\r")
End Function
