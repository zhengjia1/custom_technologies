Attribute VB_Name = "testOpMining"
' Test-only boot for the opMining step scripts (test/steps/opMining): loads the module on
' its own and stands in for the shell pieces the steps need (sidebar entries and the
' parts of setActive for stepwise/kfold/pca/kmeans, the K-means scaling modal and the
' "Calculate axis intervals" toggle).
Option Explicit

Public Sub Boot()
    Dim control As clsNode
    ' Each steps file starts with an empty clipboard, like the website runner.
    LastClipboardText = ""
    InitShared
    InitVariants
    InitOpMining
    KmeansScaleAlertProc = "testOpMining.ShowKmeansScaleAlertStub"

    Body.InsertAdjacentHtml "beforeend", _
        "<div class=""variable-rename-alert hidden"" id=""kmeansScaleAlert"" role=""dialog"" aria-modal=""true"" aria-labelledby=""kmeansScaleAlertTitle"" aria-describedby=""kmeansScaleAlertDescription"">" & _
        "<div class=""variable-rename-dialog""><div class=""variable-rename-kicker"">K-means scaling warning</div>" & _
        "<h2 id=""kmeansScaleAlertTitle"">Are you sure to disable scaling?</h2>" & _
        "<p id=""kmeansScaleAlertDescription"">Disable scaling only when the variables are already directly comparable; otherwise, scale them for distance-based methods such as k-means.</p>" & _
        "<div class=""variable-rename-actions""><button class=""sample-btn variable-rename-dismiss"" id=""keepKmeansScalingBtn"" type=""button"">Keep scaling</button>" & _
        "<button class=""primary-btn danger-btn variable-rename-dismiss"" id=""disableKmeansScalingBtn"" type=""button"">Disable scaling</button></div></div></div>"
    Q("#keepKmeansScalingBtn").AddEventListener "click", "testOpMining.OnKeepScaling"
    Q("#disableKmeansScalingBtn").AddEventListener "click", "testOpMining.OnDisableScaling"

    AddEntry "pca", "PCA & component significance"
    AddEntry "stepwise", "Forward/Backward Stepwise Selection"
    AddEntry "kmeans", "K-means clustering"
    AddEntry "kfold", "K-fold cross-validation"
    Doc.AddEventListener "click", "testOpMining.OnDocumentClick"

    Set control = CreateElement("label")
    control.ClassName = "kmeans-scale-toggle"
    control.InnerHtml = "<input type=""checkbox"" role=""switch"" class=""plot-axis-interval-toggle""><span>Calculate axis intervals</span><small>Off: basic plotting parameters and automatic axes. Applies to all plots.</small>"
    Q("#kmeansCode").Before control
    control.QuerySelector("input").AddEventListener "change", "testOpMining.OnAxisToggleChange"
End Sub

Private Sub AddEntry(ByVal id As String, ByVal label As String)
    Dim button As clsNode
    Set button = CreateElement("button")
    button.TypeName = "button"
    button.ClassName = "operation"
    button.DataSet "id", id
    button.DataSet "tutorial", "T7"
    button.TextContent = label
    Q("#tutorialList").AppendChild button
End Sub

' The parts of the website's setActive that concern the data-mining operations.
Public Sub OnDocumentClick(ByVal e As clsEvent)
    Dim op As clsNode, id As String, target As clsNode
    On Error GoTo Failed
    Set target = e.Target
    If target Is Nothing Then Exit Sub
    If target.NodeType <> 1 Then Set target = target.ParentElement
    Set op = target.Closest(".operation, .universal-entry")
    If op Is Nothing Then Exit Sub
    id = CStr(op.DataGet("id"))
    AppState("id") = id
    Q("#stepwiseBench").ToggleClass "hidden", id <> "stepwise"
    Q("#kfoldBench").ToggleClass "hidden", id <> "kfold"
    Q("#pcaBench").ToggleClass "hidden", id <> "pca"
    Q("#kmeansBench").ToggleClass "hidden", id <> "kmeans"
    Select Case id
        Case "stepwise": RenderStepwiseCode
        Case "kfold": RenderKfoldCode
        Case "pca": RenderPcaCode
        Case "kmeans": RenderKmeansCode
    End Select
    Exit Sub
Failed:
    ReportError "testOpMining.OnDocumentClick"
End Sub

Public Sub ShowKmeansScaleAlertStub()
    On Error GoTo Failed
    Q("#kmeansScaleAlert").RemoveClass "hidden"
    Body.AddClass "variable-rename-alert-open"
    Exit Sub
Failed:
    ReportError "testOpMining.ShowKmeansScaleAlertStub"
End Sub

Private Sub CloseKmeansScaleAlert(ByVal disableScaling As Boolean)
    Q("#kmeansScaleAlert").AddClass "hidden"
    Body.RemoveClass "variable-rename-alert-open"
    Q("#kmeansScale").Checked = Not disableScaling
    RenderKmeansCode
End Sub

Public Sub OnKeepScaling(ByVal e As clsEvent)
    On Error GoTo Failed
    CloseKmeansScaleAlert False
    Exit Sub
Failed:
    ReportError "testOpMining.OnKeepScaling"
End Sub

Public Sub OnDisableScaling(ByVal e As clsEvent)
    On Error GoTo Failed
    CloseKmeansScaleAlert True
    Exit Sub
Failed:
    ReportError "testOpMining.OnDisableScaling"
End Sub

Public Sub OnAxisToggleChange(ByVal e As clsEvent)
    Dim list As clsList, i As Long
    On Error GoTo Failed
    CalculatePlotAxisIntervals = e.Target.Checked
    Set list = QA(".plot-axis-interval-toggle")
    For i = 0 To list.Count - 1
        list(i).Checked = CalculatePlotAxisIntervals
    Next i
    ClearCompletionForAnswerVariant Q("#copyKmeansCodeBtn")
    RenderKmeansCode
    RefreshAnswerVariantPreviews
    Exit Sub
Failed:
    ReportError "testOpMining.OnAxisToggleChange"
End Sub

' Errors reported by handlers since boot (for debugging step scripts).
Public Function Errors() As String
    Errors = LastErrors
End Function
