Attribute VB_Name = "testOpInference"
' Test-only: boots the Tutorial 5 inference operations without the shell (modApp), and
' evaluates the distribution functions for test/opInference-numeric.js.
Option Explicit

Private mReturnFocus As clsNode

' Like the website's startup for these operations: the known-SD modal (shell markup and
' listeners), sidebar entries, the document click listener that does setActive's work for
' one-sample/two-sample/interval, then InitOpInference.
Public Sub Boot()
    Dim items As Variant, i As Long, button As clsNode, list As clsNode
    InitShared
    InitVariants
    Body.InsertAdjacentHtml "beforeend", JsU( _
        "<div class=""variable-rename-alert hidden"" id=""knownPopulationSdAlert"" role=""dialog"" aria-modal=""true"">" & _
        "<div class=""variable-rename-dialog""><h2 id=""knownPopulationSdTitle"">Are you sure?</h2>" & _
        "<div class=""variable-rename-actions""><button class=""sample-btn variable-rename-dismiss"" id=""useUnknownSdBtn"" type=""button"">Use unknown standard deviation</button>" & _
        "<button class=""primary-btn danger-btn variable-rename-dismiss"" id=""confirmKnownSdBtn"" type=""button"">Yes, population SD is known</button></div></div></div>")
    Q("#useUnknownSdBtn").AddEventListener "click", "testOpInference.OnUseUnknownSd"
    Q("#confirmKnownSdBtn").AddEventListener "click", "testOpInference.OnConfirmKnownSd"
    items = Array("interval", "Confidence / prediction intervals", "one-sample", "One-sample mean / proportion", "two-sample", "Two-sample mean / variance")
    Set list = Q("#tutorialList")
    For i = 0 To UBound(items) Step 2
        Set button = CreateElement("button")
        button.TypeName = "button"
        button.ClassName = "operation"
        button.DataSet "id", CStr(items(i))
        button.DataSet "tutorial", "Statistical Inference"
        button.TextContent = CStr(items(i + 1))
        list.AppendChild button
    Next i
    Doc.AddEventListener "click", "testOpInference.OnDocumentClick"
    InitOpInference
    KnownSdAlertProcedure = "testOpInference.ShowKnownSdAlert"
    TestPasteUnavailable = True
End Sub

Public Sub OnDocumentClick(ByVal e As clsEvent)
    Dim op As clsNode, id As String, list As clsList, i As Long, el As clsNode
    On Error GoTo Failed
    If e.Target Is Nothing Then Exit Sub
    If e.Target.NodeType <> 1 Then Exit Sub
    Set op = e.Target.Closest(".operation, .universal-entry")
    If op Is Nothing Then Exit Sub
    id = CStr(op.DataGet("id"))
    Set list = QA(".operation, .universal-entry")
    For i = 0 To list.Count - 1
        Set el = list(i)
        el.ToggleClass "active", CStr(el.DataGet("id")) = id
    Next i
    Q("#activeTitle").TextContent = op.TextContent
    Q("#standardBench").ToggleClass "hidden", id <> "universal"
    Q("#intervalBench").ToggleClass "hidden", id <> "interval"
    Q("#oneSampleBench").ToggleClass "hidden", id <> "one-sample"
    Q("#twoSampleBench").ToggleClass "hidden", id <> "two-sample"
    Select Case id
        Case "one-sample"
            SyncOneSampleForm
            RequestAnimationFrame "testOpInference.FocusOneSample"
        Case "two-sample"
            SyncTwoSampleForm
            RequestAnimationFrame "testOpInference.FocusTwoSample"
        Case "interval"
            Set list = Q("#intervalCalcType").QuerySelectorAll("option")
            For i = 0 To list.Count - 1
                list(i).Disabled = False
            Next i
            SyncIntervalForm
            RequestAnimationFrame "testOpInference.FocusInterval"
    End Select
    Exit Sub
Failed:
    ReportError "testOpInference.OnDocumentClick"
End Sub

Public Sub FocusOneSample()
    Q("#oneSampleDataset").Focus
End Sub

Public Sub FocusTwoSample()
    Q("#twoSampleDataset").Focus
End Sub

Public Sub FocusInterval()
    Q("#intervalLabel").Focus
End Sub

Public Sub FocusUnknownSd()
    Q("#useUnknownSdBtn").Focus
End Sub

' The shell's showKnownPopulationSdAlert / closeKnownPopulationSdAlert.
Public Sub ShowKnownSdAlert()
    Set mReturnFocus = ActiveElement()
    Q("#knownPopulationSdAlert").RemoveClass "hidden"
    Body.AddClass "variable-rename-alert-open"
    RequestAnimationFrame "testOpInference.FocusUnknownSd"
End Sub

Private Sub CloseKnownSdAlert(ByVal useKnown As Boolean)
    Q("#knownPopulationSdAlert").AddClass "hidden"
    Body.RemoveClass "variable-rename-alert-open"
    Q("#intervalStdev").Value = IIf(useKnown, "known", "unknown")
    Q("#intervalStdev").DataSet "previous", Q("#intervalStdev").Value
    SyncIntervalForm
    If Not mReturnFocus Is Nothing Then mReturnFocus.Focus
End Sub

Public Sub OnUseUnknownSd(ByVal e As clsEvent)
    On Error GoTo Failed
    CloseKnownSdAlert False
    Exit Sub
Failed:
    ReportError "testOpInference.OnUseUnknownSd"
End Sub

Public Sub OnConfirmKnownSd(ByVal e As clsEvent)
    On Error GoTo Failed
    CloseKnownSdAlert True
    Exit Sub
Failed:
    ReportError "testOpInference.OnConfirmKnownSd"
End Sub

' Numeric cases: each line "fn<TAB>hex1[<TAB>hex2]" (arguments as IEEE bit patterns).
' Returns one line per case: hex of the result, toFixed(2), toFixed(6), formatFinalNumber.
Public Function NumericCases(ByVal text As String) As String
    Dim lines() As String, i As Long, parts() As String, r As Double, out As clsList, a As Double, b As Double
    Set out = NewList()
    lines = Split(Replace(text, vbCrLf, vbLf), vbLf)
    For i = 0 To UBound(lines)
        If Len(lines(i)) > 0 Then
            parts = Split(lines(i), vbTab)
            a = JsHexToDouble(parts(1))
            If UBound(parts) >= 2 Then b = JsHexToDouble(parts(2))
            On Error GoTo CaseFailed
            Select Case parts(0)
                Case "inverseNormal": r = InverseNormal(a)
                Case "logGamma": r = LogGamma(a)
                Case "exp": r = FdExp(a)
                Case "betaFraction": r = BetaFraction(a, b, JsHexToDouble(parts(3)))
                Case "regularizedBeta": r = RegularizedBeta(a, b, JsHexToDouble(parts(3)))
                Case "studentTCdf": r = StudentTCdf(a, b)
                Case "inverseStudentT": r = InverseStudentT(a, b)
            End Select
            On Error GoTo 0
            out.Push JsDoubleToHex(r) & vbTab & JsToFixed(r, 2) & vbTab & JsToFixed(r, 6) & vbTab & FormatFinalNumber(r)
        End If
NextCase:
    Next i
    NumericCases = out.Join(vbLf)
    Exit Function
CaseFailed:
    out.Push "ERROR " & Err.Description
    Resume NextCase
End Function
