Attribute VB_Name = "testOpRegression"
' Test-only boot for the regression operations: builds the sidebar entries and the
' "Calculate axis intervals" toggle the shell (modApp) normally provides, and a click
' listener doing what the website's setActive does for the regression operations.
Option Explicit

Public Sub Boot()
    Dim list As clsNode
    InitShared
    InitVariants
    InitOpRegression
    RegressionUnitsAlertHandler = "testOpRegression.StubUnitsAlert"
    Set list = Q("#tutorialList")
    AddOperation list, "lm-predict", "One-variable linear regression"
    AddOperation list, "multi-regression", "Multi-variable linear regression"
    AddOperation list, "model-significance", "Model significance"
    AddOperation list, "model-comparison", "Pairwise model comparison"
    AddOperation list, "dummy", "Dummy variables / reference groups"
    AddOperation list, "interaction", "Interaction terms"
    AddOperation list, "glm", "GLM / binary variable"
    AddOperation list, "confusion", "Confusion matrix"
    Doc.AddEventListener "click", "testOpRegression.OnDocumentClick"
    AddAxisToggle "regressionCode"
End Sub

Private Sub AddOperation(ByVal list As clsNode, ByVal id As String, ByVal label As String)
    Dim button As clsNode
    Set button = CreateElement("button")
    button.SetAttribute "type", "button"
    button.ClassName = "operation"
    button.DataSet "id", id
    button.DataSet "tutorial", "Linear Regression"
    button.TextContent = label
    list.AppendChild button
End Sub

Private Sub AddAxisToggle(ByVal codeId As String)
    Dim control As clsNode
    Set control = CreateElement("label")
    control.ClassName = "kmeans-scale-toggle"
    control.SetAttribute "style", "display:flex;gap:10px;padding:12px 16px;align-items:center;flex-wrap:wrap"
    control.InnerHtml = "<input type=""checkbox"" role=""switch"" class=""plot-axis-interval-toggle""><span>Calculate axis intervals</span><small>Off: basic plotting parameters and automatic axes. Applies to all plots.</small>"
    ById(codeId).Before control
    control.QuerySelector("input").AddEventListener "change", "testOpRegression.OnAxisToggleChange"
End Sub

Public Sub StubUnitsAlert()
End Sub

Public Sub OnAxisToggleChange(ByVal e As clsEvent)
    Dim toggles As clsList, i As Long, node As clsNode
    On Error GoTo Failed
    CalculatePlotAxisIntervals = e.Target.Checked
    Set toggles = QA(".plot-axis-interval-toggle")
    For i = 0 To toggles.Count - 1
        Set node = toggles(i)
        node.Checked = CalculatePlotAxisIntervals
    Next i
    ClearCompletionForAnswerVariant ById("copyRegressionCodeBtn")
    RenderRegressionCode
    RefreshAnswerVariantPreviews
    Exit Sub
Failed:
    ReportError "testOpRegression.OnAxisToggleChange"
End Sub

Public Sub OnDocumentClick(ByVal e As clsEvent)
    Dim op As clsNode
    On Error GoTo Failed
    Set op = e.Target.Closest(".operation, .universal-entry")
    If op Is Nothing Then Exit Sub
    SetActiveForTest TextOf(op.DataGet("id")), op.TextContent
    Exit Sub
Failed:
    ReportError "testOpRegression.OnDocumentClick"
End Sub

Private Sub SetActiveForTest(ByVal id As String, ByVal title As String)
    Dim ops As clsList, i As Long, node As clsNode, isSimple As Boolean, isAdvanced As Boolean
    AppState("id") = id
    AppState("title") = title
    AppState("tutorial") = "Linear Regression"
    Set ops = QA(".operation, .universal-entry")
    For i = 0 To ops.Count - 1
        Set node = ops(i)
        node.ToggleClass "active", TextOf(node.DataGet("id")) = id
    Next i
    Q("#activeTitle").TextContent = title
    isSimple = (id = "lm-predict")
    Select Case id
        Case "multi-regression", "model-significance", "model-comparison", "dummy", "interaction", "glm", "confusion": isAdvanced = True
    End Select
    Q("#standardBench").ToggleClass "hidden", isSimple Or isAdvanced
    Q("#simpleRegressionBench").ToggleClass "hidden", Not isSimple
    Q("#advancedRegressionBench").ToggleClass "hidden", Not isAdvanced
    If isSimple Then
        RenderRegressionCode
    ElseIf isAdvanced Then
        RenderAdvancedRegressionForm id
    End If
End Sub

Private Function TextOf(ByVal value As Variant) As String
    If Not IsNull(value) Then TextOf = CStr(value)
End Function
