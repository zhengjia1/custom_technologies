Attribute VB_Name = "testOpPlot"
' Test-only boot for the Plotting & tables port (opPlot, opPlot2): stands in for the parts
' of the shell (modApp) that the step scripts use - the sidebar entry for "open ggplot"
' (what setActive does for it) and the "Calculate axis intervals" toggle before #ggplotCode.
Option Explicit

Public Sub Boot()
    Dim button As clsNode, control As clsNode
    InitShared
    InitVariants
    Set button = CreateElement("button")
    button.SetAttribute "type", "button"
    button.ClassName = "operation"
    button.DataSet "id", "ggplot"
    button.DataSet "tutorial", "T2"
    button.TextContent = "Plotting & tables"
    Q("#tutorialList").AppendChild button
    button.AddEventListener "click", "testOpPlot.OnOpen"
    InitOpPlot
    ' plotAxisControls (website lines 4799-4818), for the plotting bench only.
    Set control = CreateElement("label")
    control.ClassName = "kmeans-scale-toggle"
    control.SetAttribute "style", "display:flex;gap:10px;padding:12px 16px;align-items:center;flex-wrap:wrap"
    control.InnerHtml = "<input type=""checkbox"" role=""switch"" class=""plot-axis-interval-toggle""><span>Calculate axis intervals</span><small>Off: basic plotting parameters and automatic axes. Applies to all plots.</small>"
    Q("#ggplotCode").Before control
    control.QuerySelector("input").AddEventListener "change", "testOpPlot.OnAxisToggle"
End Sub

Public Sub OnOpen(ByVal e As clsEvent)
    Dim list As clsList, i As Long
    On Error GoTo Failed
    AppState("id") = "ggplot"
    AppState("title") = "Plotting & tables"
    Set list = QA(".operation, .universal-entry")
    For i = 0 To list.Count - 1
        list(i).ToggleClass "active", (NzText(list(i).DataGet("id")) = "ggplot")
    Next i
    Q("#standardBench").ToggleClass "hidden", True
    Q("#ggplotBench").ToggleClass "hidden", False
    RenderGgplotCode
    Exit Sub
Failed:
    ReportError "testOpPlot.OnOpen"
End Sub

Public Sub OnAxisToggle(ByVal e As clsEvent)
    Dim list As clsList, i As Long
    On Error GoTo Failed
    CalculatePlotAxisIntervals = e.Target.Checked
    Set list = QA(".plot-axis-interval-toggle")
    For i = 0 To list.Count - 1
        list(i).Checked = CalculatePlotAxisIntervals
    Next i
    ClearCompletionForAnswerVariant Q("#copyGgplotCodeBtn")
    RenderGgplotCode
    RefreshAnswerVariantPreviews
    Exit Sub
Failed:
    ReportError "testOpPlot.OnAxisToggle"
End Sub

Private Function NzText(ByVal value As Variant) As String
    If Not IsNull(value) Then NzText = CStr(value)
End Function
