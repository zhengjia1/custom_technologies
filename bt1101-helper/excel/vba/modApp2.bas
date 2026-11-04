Attribute VB_Name = "modApp2"
' Application shell, part 2: answer-option document listeners, collapsible cards, the
' sidebar operation list, setActive, the Universal Detector (detectionRules/detect), the
' self-contained parsers (Shapiro-Wilk, Pearson, interval, one-/two-sample), generate,
' textarea fitting and the standard workbench handlers. See modApp for the startup order.
Option Explicit

Private mGenerationTimer As Long
Private Const COMPACT_TEXTAREA_SELECTOR As String = ".pca-variance-input, .test-output-input, .lp-smart-input"

' ================================================================ small DOM helpers

Public Function ShellClosest(ByVal node As clsNode, ByVal selector As String) As clsNode
    If node Is Nothing Then Exit Function
    If node.NodeType <> 1 Then Exit Function
    Set ShellClosest = node.Closest(selector)
End Function

Public Function ShellMatches(ByVal node As clsNode, ByVal selector As String) As Boolean
    If node Is Nothing Then Exit Function
    If node.NodeType <> 1 Then Exit Function
    ShellMatches = node.Matches(selector)
End Function

Public Function ShellDataText(ByVal node As clsNode, ByVal key As String) As String
    Dim value As Variant
    If node Is Nothing Then Exit Function
    value = node.DataGet(key)
    If Not IsNull(value) Then ShellDataText = CStr(value)
End Function

Private Function StateId() As String
    StateId = CStr(AppState("id"))
End Function

' requestAnimationFrame(() => $(selector)?.focus())
Private Sub FocusLater(ByVal selector As String)
    SetTimeout "modApp2.FocusSelector", 16, selector
End Sub

Public Sub FocusSelector(ByVal selector As String)
    Dim el As clsNode
    On Error GoTo Failed
    Set el = Q(selector)
    If Not el Is Nothing Then el.Focus
    Exit Sub
Failed:
    ReportError "modApp2.FocusSelector"
End Sub

' requestAnimationFrame(() => node.scrollIntoView(...)): recorded for the renderer.
Private Sub ScrollIntoViewLater(ByVal node As clsNode)
    If node Is Nothing Then Exit Sub
    SetTimeout "modApp2.ScrollNodeIntoView", 16, node
End Sub

Public Sub ScrollNodeIntoView(ByVal node As clsNode)
    On Error GoTo Failed
    Set modApp.ScrollRequestTarget = node
    Exit Sub
Failed:
    ReportError "modApp2.ScrollNodeIntoView"
End Sub

Private Sub SetValueIfPresent(ByVal selector As String, ByVal value As String)
    Dim el As clsNode
    Set el = Q(selector)
    If Not el Is Nothing Then el.Value = value
End Sub

' ================================================================ startup parts

' Website lines 963-1590 (top-level statements in order).
Public Sub InitShellPartA()
    ' Answer-option bookkeeping listeners (lines 974-993).
    Doc.AddEventListener "click", "modApp2.OnDocClickAnswerVariants", True
    Doc.AddEventListener "paste", "modApp2.OnDocPaste", True
    Doc.AddEventListener "input", "modApp2.OnDocInputAnswerVariants"
    Doc.AddEventListener "change", "modApp2.OnDocChangeAnswerVariants"
    ' answerVariantObserver (and, from part C, the number-input observer).
    SetMutationHook "modApp2.OnDomMutation"
    modApp3.NumberObserverActive = False
    ' Update check.
    Q("#checkForUpdatesBtn").AddEventListener "click", "modApp.HandleUpdateCheckClick"
    CheckForUpdates
    ' Objective-value field and the warning modals.
    InsertShellMarkup
    RegisterAlertListeners
    ConfigureSampleSelect "universal", "Universal detector"
    ' Collapsible completion tracking.
    Doc.AddEventListener "click", "modApp2.OnDocClickCollapsible", True
    Doc.AddEventListener "input", "modApp2.OnDocInputCollapsible"
    Doc.AddEventListener "change", "modApp2.OnDocChangeCollapsible"
    InitialiseLpCollapsibles
    EnhanceAnswerVariantControls
    BuildOperationList
    Doc.AddEventListener "click", "modApp2.OnDocClickOperation"
End Sub

' Website lines 3889-4258.
Public Sub InitShellPartB()
    Doc.AddEventListener "input", "modApp2.OnDocInputFitTextarea"
    RequestAnimationFrame "modApp2.FitAllCompactTextareas"
    Q("#inputText").AddEventListener "input", "modApp2.OnInputTextInput"
    Q("#resetBtn").AddEventListener "click", "modApp2.OnResetClick"
    Q("#universalSampleSelect").AddEventListener "change", "modApp2.OnUniversalSampleChange"
    Q("#singleSampleBtn").AddEventListener "click", "modApp2.OnSingleSampleClick"
    Q("#pasteBtn").AddEventListener "click", "modApp2.OnPasteClick"
    Q("#copyBtn").AddEventListener "click", "modApp2.OnCopyClick"
End Sub

' ================================================================ answer-option listeners

Private Sub RememberTrustedOperationEdit(ByVal e As clsEvent)
    Dim bench As clsNode, operationButton As clsNode, id As String
    If ShellMatches(e.Target, ".answer-variant-select") Then Exit Sub
    Set bench = ShellClosest(e.Target, ".bench")
    If bench Is Nothing Then Exit Sub
    If bench.HasClass("hidden") Then Exit Sub
    id = StateId()
    UserEditedOperations(id) = True
    Set operationButton = Q("#tutorialList .operation[data-id=""" & CssEscape(id) & """]")
    If Not operationButton Is Nothing Then operationButton.DataSet "userEdited", "true"
    RememberOperationAnswerSelections id, bench
    Set LockedOperationAnswerSelections(id) = CollectAnswerSelections(bench)
End Sub

Public Sub OnDocClickAnswerVariants(ByVal e As clsEvent)
    Dim button As clsNode, scope As clsNode
    On Error GoTo Failed
    Set button = ShellClosest(e.Target, "button")
    If Not button Is Nothing Then
        Set modVariants.ActiveAnswerCopyButton = button
        If RxTest(button.Id, "^paste", "i") Then
            Set scope = button.Closest(".bench")
            If scope Is Nothing Then Set scope = Doc
            SetTimeout "modApp2.RandomizeAnswerOptionsAfterPaste", 500, scope
        End If
    End If
    ScheduleAnswerVariantRefresh e.Target
    Exit Sub
Failed:
    ReportError "modApp2.OnDocClickAnswerVariants"
End Sub

Public Sub OnDocPaste(ByVal e As clsEvent)
    Dim scope As clsNode
    On Error GoTo Failed
    Set scope = ShellClosest(e.Target, ".bench")
    If scope Is Nothing Then Set scope = Doc
    SetTimeout "modApp2.RandomizeAnswerOptionsAfterPaste", 0, scope
    Exit Sub
Failed:
    ReportError "modApp2.OnDocPaste"
End Sub

Public Sub RandomizeAnswerOptionsAfterPaste(ByVal scope As clsNode)
    On Error GoTo Failed
    RandomizeAnswerOptions scope, True
    Exit Sub
Failed:
    ReportError "modApp2.RandomizeAnswerOptionsAfterPaste"
End Sub

Public Sub OnDocInputAnswerVariants(ByVal e As clsEvent)
    On Error GoTo Failed
    RememberTrustedOperationEdit e
    ScheduleAnswerVariantRefresh e.Target
    Exit Sub
Failed:
    ReportError "modApp2.OnDocInputAnswerVariants"
End Sub

Public Sub OnDocChangeAnswerVariants(ByVal e As clsEvent)
    Dim id As String
    On Error GoTo Failed
    RememberTrustedOperationEdit e
    If ShellMatches(e.Target, ".answer-variant-select") Then
        id = StateId()
        RememberOperationAnswerSelections id
        If UserEditedOperations.Exists(id) Then Set LockedOperationAnswerSelections(id) = CollectAnswerSelections()
    Else
        ScheduleAnswerVariantRefresh e.Target
    End If
    Exit Sub
Failed:
    ReportError "modApp2.OnDocChangeAnswerVariants"
End Sub

' MutationObserver callbacks: answer-option controls for every added element, then (once
' part C has run) number steppers, queued so all records reach the first observer first.
Public Sub OnDomMutation(ByVal parent As clsNode, ByVal added As clsList)
    Dim i As Long, node As clsNode
    On Error GoTo Failed
    For i = 0 To added.Count - 1
        Set node = added(i)
        If node.NodeType = 1 Then
            EnhanceAnswerVariantControls node
            If modApp3.NumberObserverActive Then QueueMicrotask "modApp3.EnhanceNumberInputsTask", node
        End If
    Next i
    Exit Sub
Failed:
    ReportError "modApp2.OnDomMutation"
End Sub

' ================================================================ sample select

Public Sub ConfigureSampleSelect(ByVal id As String, ByVal title As String)
    Dim sel As clsNode, singleBtn As clsNode, i As Long, hasSingleSample As Boolean
    Set sel = Q("#universalSampleSelect")
    Set singleBtn = Q("#singleSampleBtn")
    sel.InnerHtml = ""
    sel.AppendChild NewOption(JsU("Use sample\u2026"), "")
    If id = "universal" Then
        For i = 0 To UniversalSampleCount - 1
            sel.AppendChild NewOption(UniversalLabels(i), UniversalRoutes(i))
        Next i
    End If
    sel.Value = ""
    hasSingleSample = (id <> "universal" And SampleIndex(id) >= 0)
    sel.ToggleClass "hidden", id <> "universal"
    sel.Disabled = (id <> "universal")
    sel.Title = "Choose a detector sample to load immediately"
    singleBtn.ToggleClass "hidden", Not hasSingleSample
    singleBtn.Disabled = Not hasSingleSample
    If hasSingleSample Then singleBtn.Title = "Load the " & title & " sample" Else singleBtn.Title = ""
End Sub

' new Option(text, value)
Private Function NewOption(ByVal text As String, ByVal value As String) As clsNode
    Dim opt As clsNode
    Set opt = CreateElement("option")
    opt.Text = text
    opt.Value = value
    Set NewOption = opt
End Function

' ================================================================ collapsible cards

' The website's resetLpWorkflowAfterModelChange (its copiedSensitivitySections set lives
' in opLp: cleared through opLp's ClearCopiedSensitivitySections).
Public Sub ResetLpWorkflowAfterModelChange()
    Dim panels As clsList, i As Long, panel As clsNode, done As clsList, k As Long, buttons As clsList
    Set panels = QA("#lpBench .lp-collapsible")
    For i = 0 To panels.Count - 1
        Set panel = panels(i)
        SetPanelCompletion panel, False
        SetLpPanelCollapsed panel, False
        Set done = panel.QuerySelectorAll(".is-copy-complete")
        For k = 0 To done.Count - 1
            done(k).RemoveClass "is-copy-complete"
        Next k
    Next i
    TryCallOp "ClearCopiedSensitivitySections"
    Set buttons = QA("#sensitivityInterpretation [data-sensitivity-copy]")
    For i = 0 To buttons.Count - 1
        buttons(i).TextContent = "Copy"
        buttons(i).Disabled = False
    Next i
End Sub

Public Sub OnDocClickCollapsible(ByVal e As clsEvent)
    Dim button As clsNode, panel As clsNode
    On Error GoTo Failed
    Set button = ShellClosest(e.Target, "button")
    If button Is Nothing Then Exit Sub
    Set panel = button.Closest(".lp-collapsible")
    If panel Is Nothing Or button.Disabled Then Exit Sub
    If RxTest(JsTrim(button.TextContent), "^Copy\b", "i") Then
        If ShouldCompletePanelOnCopy(panel, button) Then
            Set PendingCopyPanel = panel
            Set PendingCopyButton = button
            PendingCopyTime = Now
        Else
            Set PendingCopyPanel = Nothing
            Set PendingCopyButton = Nothing
        End If
    ElseIf Not button.HasClass("lp-collapse-toggle") Then
        SetPanelCompletion panel, False
    End If
    Exit Sub
Failed:
    ReportError "modApp2.OnDocClickCollapsible"
End Sub

Public Sub OnDocInputCollapsible(ByVal e As clsEvent)
    On Error GoTo Failed
    SetPanelCompletion ShellClosest(e.Target, ".lp-collapsible"), False
    Exit Sub
Failed:
    ReportError "modApp2.OnDocInputCollapsible"
End Sub

Public Sub OnDocChangeCollapsible(ByVal e As clsEvent)
    On Error GoTo Failed
    SetPanelCompletion ShellClosest(e.Target, ".lp-collapsible"), False
    Exit Sub
Failed:
    ReportError "modApp2.OnDocChangeCollapsible"
End Sub

Public Sub InitialiseLpCollapsibles()
    Dim panels As clsList, i As Long, panel As clsNode, head As clsNode, actions As clsNode
    Dim children As clsList, k As Long, titleContainer As clsNode, strong As clsNode, title As String, button As clsNode
    Set panels = QA("#lpBench .panel, #pcaBench > .panel, #advancedRegressionBench > .panel, #anovaBench > .panel, #ggplotBench > .axis-transform-panel")
    For i = 0 To panels.Count - 1
        Set panel = panels(i)
        Set head = panel.QuerySelector(":scope > .panel-head")
        If head Is Nothing Then GoTo NextPanel
        If Len(ShellDataText(panel, "collapsibleReady")) > 0 Then GoTo NextPanel
        panel.DataSet "collapsibleReady", "true"
        panel.AddClass "lp-collapsible"
        Set actions = head.QuerySelector(":scope > .panel-actions")
        If actions Is Nothing Then
            Set actions = CreateElement("div")
            actions.ClassName = "panel-actions"
            Set children = head.Children
            For k = 0 To children.Count - 1
                If children(k).TagName = "BUTTON" Then actions.AppendChild children(k)
            Next k
            head.AppendChild actions
        End If
        Set titleContainer = head.QuerySelector(".panel-title")
        title = "section"
        If Not titleContainer Is Nothing Then
            Set strong = titleContainer.QuerySelector("strong")
            If Not strong Is Nothing Then
                If Len(JsTrim(strong.TextContent)) > 0 Then title = JsTrim(strong.TextContent)
            End If
        End If
        Set button = CreateElement("button")
        button.TypeName = "button"
        button.ClassName = "lp-collapse-toggle"
        button.SetAttribute "aria-expanded", "true"
        button.SetAttribute "aria-label", "Collapse " & title
        button.AddEventListener "click", "modApp2.OnCollapseToggleClick"
        If Not titleContainer Is Nothing Then titleContainer.InsertBefore button, titleContainer.FirstChild
NextPanel:
    Next i
End Sub

Public Sub OnCollapseToggleClick(ByVal e As clsEvent)
    Dim panel As clsNode
    On Error GoTo Failed
    Set panel = e.CurrentTarget.Closest(".lp-collapsible")
    If Not panel Is Nothing Then SetLpPanelCollapsed panel, Not panel.HasClass("is-collapsed")
    Exit Sub
Failed:
    ReportError "modApp2.OnCollapseToggleClick"
End Sub

' ================================================================ sidebar and setActive

Private Sub BuildOperationList()
    Dim list As clsNode, i As Long, button As clsNode
    Set list = Q("#tutorialList")
    For i = 0 To OpCount - 1
        Set button = CreateElement("button")
        button.TypeName = "button"
        button.ClassName = "operation"
        button.DataSet "id", OpIds(i)
        button.DataSet "tutorial", GroupShorts(OpGroup(i))
        button.DataSet "search", JsLower(OpLabels(i) & " " & OpTerms(i) & " " & GroupTitles(OpGroup(i)))
        button.TextContent = OpLabels(i)
        list.AppendChild button
    Next i
End Sub

Public Sub OnDocClickOperation(ByVal e As clsEvent)
    Dim op As clsNode, titleEl As clsNode, title As String, tutorial As String
    On Error GoTo Failed
    Set op = ShellClosest(e.Target, ".operation, .universal-entry")
    If op Is Nothing Then Exit Sub
    Set titleEl = op.QuerySelector(".entry-title")
    If Not titleEl Is Nothing Then title = titleEl.TextContent
    If Len(title) = 0 Then title = op.TextContent
    tutorial = ShellDataText(op, "tutorial")
    If Len(tutorial) = 0 Then tutorial = "Smart entry point"
    SetActive ShellDataText(op, "id"), title, tutorial
    Exit Sub
Failed:
    ReportError "modApp2.OnDocClickOperation"
End Sub

Private Sub ToggleBench(ByVal selector As String, ByVal hidden As Boolean)
    Q(selector).ToggleClass "hidden", hidden
End Sub

Public Sub SetActive(ByVal id As String, ByVal title As String, ByVal tutorial As String)
    Dim previousId As String, isNewVisit As Boolean, entries As clsList, i As Long, el As clsNode
    Dim isSpecial As Boolean, isAdvancedRegression As Boolean, isLp As Boolean, sampleAt As Long
    Dim options As clsList, outputTitle As String
    previousId = StateId()
    isNewVisit = (id <> previousId)
    If isNewVisit Then RememberOperationAnswerSelections previousId
    AppState("id") = id: AppState("title") = title: AppState("tutorial") = tutorial
    Set entries = QA(".operation, .universal-entry")
    For i = 0 To entries.Count - 1
        Set el = entries(i)
        el.ToggleClass "active", ShellDataText(el, "id") = id And Not IsNull(el.DataGet("id"))
    Next i
    Q("#activeTitle").TextContent = title
    Select Case id
        Case "multi-regression", "model-significance", "model-comparison", "dummy", "interaction", "glm", "confusion"
            isAdvancedRegression = True
    End Select
    isLp = (id = "lpsolve" Or id = "sensitivity")
    Select Case id
        Case "summarize", "shape", "wide-table", "interval", "one-sample", "two-sample", "anova-many", "corr-multiple", _
             "lm-predict", "stepwise", "kfold", "pca", "kmeans", "ggplot"
            isSpecial = True
    End Select
    isSpecial = isSpecial Or isAdvancedRegression Or isLp
    ToggleBench "#standardBench", isSpecial
    ToggleBench "#summarizeBench", id <> "summarize"
    ToggleBench "#shapeBench", id <> "shape"
    ToggleBench "#wideTableBench", id <> "wide-table"
    ToggleBench "#intervalBench", id <> "interval"
    ToggleBench "#oneSampleBench", id <> "one-sample"
    ToggleBench "#twoSampleBench", id <> "two-sample"
    ToggleBench "#anovaBench", id <> "anova-many"
    ToggleBench "#corrMultipleBench", id <> "corr-multiple"
    ToggleBench "#simpleRegressionBench", id <> "lm-predict"
    ToggleBench "#advancedRegressionBench", Not isAdvancedRegression
    ToggleBench "#stepwiseBench", id <> "stepwise"
    ToggleBench "#kfoldBench", id <> "kfold"
    ToggleBench "#pcaBench", id <> "pca"
    ToggleBench "#kmeansBench", id <> "kmeans"
    ToggleBench "#lpBench", Not isLp
    ToggleBench "#lpSensitivityAnalysis", Not isLp
    ToggleBench "#ggplotBench", id <> "ggplot"
    ToggleBench "#referencePanel", isSpecial Or id = "universal"
    If id = "summarize" Then
        TryCallOp "RenderSummaryCode": FocusLater "#columnFields .column-input"
    ElseIf id = "shape" Then
        TryCallOp "RenderShapeInterpretation": FocusLater "#skewnessInput"
    ElseIf id = "wide-table" Then
        TryCallOp "RenderWideTableCode": FocusLater "#wideDatasetInput"
    ElseIf id = "one-sample" Then
        TryCallOp "SyncOneSampleForm": FocusLater "#oneSampleDataset"
    ElseIf id = "two-sample" Then
        TryCallOp "SyncTwoSampleForm": FocusLater "#twoSampleDataset"
    ElseIf id = "anova-many" Then
        TryCallOp "RenderAnovaBlocks": FocusLater "#anovaDataset"
    ElseIf id = "corr-multiple" Then
        TryCallOp "RenderCorrMultipleCode": FocusLater "#corrMultipleColumns"
    ElseIf id = "lm-predict" Then
        TryCallOp "RenderRegressionCode": FocusLater "#regressionY"
    ElseIf isAdvancedRegression Then
        TryCallOp "RenderAdvancedRegressionForm", id: FocusLater "#advancedRegressionForm input"
    ElseIf id = "stepwise" Then
        TryCallOp "RenderStepwiseCode": FocusLater "#stepwiseOutcome"
    ElseIf id = "kfold" Then
        TryCallOp "RenderKfoldCode": FocusLater "#kfoldOutcome"
    ElseIf id = "pca" Then
        TryCallOp "RenderPcaCode": FocusLater "#pcaVariables"
    ElseIf id = "kmeans" Then
        TryCallOp "RenderKmeansCode": FocusLater "#kmeansFeatures"
    ElseIf isLp Then
        If id = "sensitivity" And Q("#lpVariableType").Value = "continuous" Then Q("#lpSensitivity").Checked = True
        TryCallOp "RenderLpModel"
        TryCallOp "RenderSensitivityInterpretation"
        FocusLater "#lpVariables"
    ElseIf id = "ggplot" Then
        TryCallOp "RenderGgplotCode": FocusLater "#plotX"
    ElseIf id = "interval" Then
        Set options = Q("#intervalCalcType").QuerySelectorAll("option")
        For i = 0 To options.Count - 1
            options(i).Disabled = False
        Next i
        TryCallOp "SyncIntervalForm"
        FocusLater "#intervalLabel"
    End If
    sampleAt = SampleIndex(id)
    ConfigureSampleSelect id, title
    If sampleAt >= 0 Then
        Q("#sampleCode").TextContent = SampleCodes(sampleAt)
        Q("#sampleInput").TextContent = SampleInputs(sampleAt)
        Q("#sampleProse").TextContent = SampleProse(sampleAt)
        Q("#inputText").Placeholder = SampleInputs(sampleAt)
    Else
        Q("#sampleCode").TextContent = "No separate sample code is required for " & title & "."
        Q("#sampleInput").TextContent = "Use the active workspace to enter " & title & " values or output."
        Q("#sampleProse").TextContent = "The generated " & title & " result appears in the workspace above."
        Q("#inputText").Placeholder = JsU("Paste your RStudio console output here\u2026") & vbLf & vbLf & "Extra or missing spaces (up to 3) are supported."
    End If
    If id = "universal" Then outputTitle = "Your interpretation will appear here" Else outputTitle = title & " selected"
    Q("#outputBody").InnerHtml = "<div class=""empty-state""><div class=""empty-state-inner""><span class=""empty-glyph"">" & ChrW$(&HB6) & "</span><strong>" & outputTitle & _
        "</strong><p>Paste the matching RStudio output, then choose Identify &amp; generate.</p></div></div>"
    If isNewVisit Then PrepareAnswerOptionsForVisit id
    If WindowInnerWidth <= 720 Then
        Q("#library").RemoveClass "open"
        Q("#mobileLibraryBtn").SetAttribute "aria-expanded", "false"
    End If
End Sub

' ================================================================ detector

' { id, title, tutorial } for the group entry of an operation id.
Private Function RouteFor(ByVal id As String, ByVal fallbackTitle As String, ByVal fallbackTutorial As String) As clsMap
    Dim at As Long, route As clsMap
    Set route = NewMap()
    at = OpIndex(id)
    route("id") = id
    If at >= 0 Then
        route("title") = OpLabels(at)
        route("tutorial") = GroupShorts(OpGroup(at))
    Else
        route("title") = fallbackTitle
        route("tutorial") = fallbackTutorial
    End If
    Set RouteFor = route
End Function

Private Function IsGamesHowellOutputLocal(ByVal text As String) As Boolean
    ' Local copy of opAnova's isGamesHowellOutput (two regular expressions).
    IsGamesHowellOutputLocal = RxTest(text, "games[\s_\u2013-]*howell", "i") Or _
        RxTest(RxReplace(text, "[\s|]", "", "g"), "\.y\.group1group2estimateconf\.lowconf\.highp\.adj", "i")
End Function

Private Function MapFlag(ByVal map As Object, ByVal key As String) As Boolean
    On Error GoTo Failed
    If map Is Nothing Then Exit Function
    If TypeOf map Is clsMap Then
        Dim m As clsMap
        Set m = map
        If m.Exists(key) Then
            If Not IsObject(m(key)) Then
                If VarType(m(key)) = vbBoolean Then MapFlag = m(key)
            End If
        End If
    End If
    Exit Function
Failed:
    Err.Clear
End Function

' detectionRules: the first rule whose pattern matches, or "".
Public Function DetectionRuleId(ByVal text As String) As String
    Dim ids As Variant, patterns(0 To 9) As String, i As Long
    ids = Array("interval", "one-sample", "two-sample", "shapiro", "pearson", "glm", "multi-regression", "confusion", "pca", "lpsolve")
    patterns(0) = "\b[ul](?:CI|PI)[A-Za-z0-9.]+\b|tail_probability\s*<-\s*[0-9.]+[\s\S]*(?:qt|qnorm)\s*\("
    patterns(1) = "one\s+sample\s+t-test|1-sample\s+proportions|\bz\s*=\s*[-+0-9.e]+[\s\S]*p-value"
    patterns(2) = "two\s+sample\s+t-test|F\s+test\s+to\s+compare\s+two\s+variances"
    patterns(3) = "shapiro|\bW\s*=|normality test"
    patterns(4) = "pearson|cor\.test|sample estimates[\s\S]*\bcor\b"
    patterns(5) = "glm\s*\(|binomial|null deviance|logit"
    patterns(6) = "call:\s*lm|residual standard error|coefficients:"
    patterns(7) = "confusion matrix|sensitivity|specificity"
    patterns(8) = "principal components|proportion of variance|rotation|prcomp"
    patterns(9) = "lpsolve|objective function|solution status|(?:minimi[sz]e|maximi[sz]e)[\s\S]*(?:subject\s+to|\\(?:leq|geq|le|ge)|[\u2264\u2265]|<=|>=)"
    For i = 0 To 9
        If RxTest(text, patterns(i), "i") Then DetectionRuleId = ids(i): Exit Function
    Next i
End Function

' detect(text): clsMap { id, title, tutorial[, gamesHowell][, parsedModel] } or Nothing.
Public Function Detect(ByVal text As String) As clsMap
    Dim route As clsMap, parsedLpModel As Object, parsedRegression As Object, coefficients As Object, id As String
    If IsGamesHowellOutputLocal(text) Or RxTest(text, "Welch\s+ANOVA|welch_anova_test\s*\(", "i") Then
        Set route = RouteFor("anova-many", "ANOVA", "Statistical inference")
        route("gamesHowell") = IsGamesHowellOutputLocal(text)
        Set Detect = route
        Exit Function
    End If
    Set parsedLpModel = TryCallOpObject("ParseLpModelText", text)
    If MapFlag(parsedLpModel, "ok") Then
        Set route = RouteFor("lpsolve", "lpSolve", "Linear Optimization")
        Set route("parsedModel") = parsedLpModel
        Set Detect = route
        Exit Function
    End If
    Set parsedRegression = TryCallOpObject("ParseMultiRegressionOutput", text)
    If MapFlag(parsedRegression, "ok") Then
        On Error Resume Next
        Set coefficients = parsedRegression("coefficients")
        On Error GoTo 0
        If Not coefficients Is Nothing Then
            If coefficients.Count > 1 Then
                Set Detect = RouteFor("multi-regression", "Multi-variable linear regression", "Linear Regression")
                Exit Function
            End If
        End If
    End If
    id = DetectionRuleId(text)
    If Len(id) = 0 Then Exit Function
    If OpIndex(id) < 0 Then Exit Function
    Set Detect = RouteFor(id, "", "")
End Function

' ================================================================ parsers

Private Function ErrorResult(ByVal message As String) As clsMap
    Dim result As clsMap
    Set result = NewMap()
    result("error") = message
    Set ErrorResult = result
End Function

Private Function DisplayedP(ByVal operator As String, ByVal pValue As Double) As String
    If operator = "=" Then DisplayedP = FormatPValue(pValue) Else DisplayedP = operator & " " & FormatPValue(pValue)
End Function

Private Function GroupOrNull(ByVal m As clsMatch, ByVal n As Long) As Variant
    GroupOrNull = Null
    If m Is Nothing Then Exit Function
    If m.Matched(n) Then GroupOrNull = m.Group(n)
End Function

Private Function NumberPattern() As String
    NumberPattern = "([+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?)"
End Function

Public Function ParseShapiro(ByVal text As String) As clsMap
    Dim variableMatch As clsMatch, wMatch As clsMatch, pMatch As clsMatch, operator As String, pValue As Double
    Dim reject As Variant, variable As String, shownP As String, decision As String, evidence As String, result As clsMap
    Set variableMatch = RxMatch(text, "(?:#{1,3}\s{0,3})?data\s{0,3}:\s{0,3}([^\r\n]+)", "i")
    Set wMatch = RxMatch(text, "\bW\s{0,3}=\s{0,3}" & NumberPattern(), "i")
    Set pMatch = RxMatch(text, "p(?:-|\s{0,3})value\s{0,3}(<=|>=|<|>|=)\s{0,3}" & NumberPattern(), "i")
    If pMatch Is Nothing Then
        Set ParseShapiro = ErrorResult(JsU("I could not find a p-value. Include a line such as \u2018p-value = 0.032\u2019 or \u2018p-value < 2.2e-16\u2019."))
        Exit Function
    End If
    operator = pMatch.Group(1)
    pValue = JsNumber(pMatch.Group(2))
    If Not JsIsFinite(pValue) Then Set ParseShapiro = ErrorResult("The detected p-value is not a valid number."): Exit Function
    reject = PValueDecision(operator, pValue)
    If IsNull(reject) Then Set ParseShapiro = ErrorResult(PValueDecisionError(operator, pValue)): Exit Function
    If variableMatch Is Nothing Then variable = "the variable" Else variable = variableMatch.Group(1)
    If Len(variable) = 0 Then variable = "the variable"
    variable = RxReplace(variable, "^\s{0,3}|\s{0,3}$", "", "g")
    shownP = DisplayedP(operator, pValue)
    If reject Then decision = "reject": evidence = "sufficient" Else decision = "fail to reject": evidence = "insufficient"
    Set result = NewMap()
    result("variable") = variable
    result("w") = GroupOrNull(wMatch, 1)
    result("pValue") = pValue
    result("displayedP") = shownP
    result("reject") = CBool(reject)
    result("prose") = SYM_H0() & ": " & variable & " is normally distributed." & vbLf & vbLf & SYM_H1() & ": " & variable & " is not normally distributed." & vbLf & vbLf & _
        "Since the p-value is " & shownP & ", which is " & PValueComparison(operator, pValue) & " 0.05, we " & decision & _
        " the null hypothesis at the 5% significance level. There is " & evidence & " evidence to conclude that " & variable & " is not normally distributed."
    Set ParseShapiro = result
End Function

Private Function FriendlyName(ByVal value As String) As String
    FriendlyName = Replace(RxReplace(JsTrim(value), "^as\.numeric\(([\s\S]*)\)$", "$1", "i"), "`", "")
End Function

Public Function ParsePearson(ByVal text As String) As clsMap
    Dim cleaned As String, dataMatch As clsMatch, flat As String, dataText As String, dataParts As clsList
    Dim variable1 As String, variable2 As String, pMatch As clsMatch, corMatch As clsMatch, tMatch As clsMatch
    Dim dfMatch As clsMatch, ciMatch As clsMatch, operator As String, pValue As Double, correlation As Double
    Dim reject As Variant, shownP As String, direction As String, magnitude As Double, strength As String
    Dim decision As String, statisticText As String, conclusion As String, result As clsMap, rest As String, i As Long
    cleaned = RxReplace(text, "^\s{0,3}#{1,3}\s{0,3}", "", "gm")
    cleaned = Replace(cleaned, vbCr, "")
    Set dataMatch = RxMatch(cleaned, "data\s{0,3}:\s{0,3}([\s\S]*?)(?=\n\s{0,3}t\s{0,3}=)", "i")
    flat = JsTrim(RxReplace(cleaned, "\s+", " ", "g"))
    If dataMatch Is Nothing Then dataText = "variable 1 and variable 2" Else dataText = dataMatch.Group(1)
    If Len(dataText) = 0 Then dataText = "variable 1 and variable 2"
    Set dataParts = Rx("\s{0,3}\band\b\s{0,3}", "i").Split(JsTrim(RxReplace(dataText, "\s+", " ", "g")))
    If dataParts.Count > 0 Then variable1 = CStr(dataParts(0))
    If Len(variable1) = 0 Then variable1 = "variable 1"
    variable1 = FriendlyName(variable1)
    For i = 1 To dataParts.Count - 1
        If i > 1 Then rest = rest & " and "
        rest = rest & CStr(dataParts(i))
    Next i
    If Len(rest) = 0 Then rest = "variable 2"
    variable2 = FriendlyName(rest)
    Set pMatch = RxMatch(flat, "p(?:-|\s{0,3})value\s{0,3}(<=|>=|<|>|=)\s{0,3}" & NumberPattern(), "i")
    Set corMatch = RxMatch(flat, "sample\s{0,3}estimates\s{0,3}:\s{0,3}(?:cor\s{0,3})?" & NumberPattern(), "i")
    Set tMatch = RxMatch(flat, "\bt\s{0,3}=\s{0,3}" & NumberPattern(), "i")
    Set dfMatch = RxMatch(flat, "\bdf\s{0,3}=\s{0,3}" & NumberPattern(), "i")
    Set ciMatch = RxMatch(flat, "95\s{0,3}percent\s{0,3}confidence\s{0,3}interval\s{0,3}:\s{0,3}" & NumberPattern() & "\s{1,3}" & NumberPattern(), "i")
    If pMatch Is Nothing Then Set ParsePearson = ErrorResult(JsU("I could not find a p-value. Include the \u2018t = ..., df = ..., p-value = ...\u2019 line.")): Exit Function
    If corMatch Is Nothing Then Set ParsePearson = ErrorResult(JsU("I could not find the correlation coefficient. Include the \u2018sample estimates: cor\u2019 section.")): Exit Function
    operator = pMatch.Group(1)
    pValue = JsNumber(pMatch.Group(2))
    correlation = JsNumber(corMatch.Group(1))
    If Not JsIsFinite(pValue) Or Not JsIsFinite(correlation) Then Set ParsePearson = ErrorResult("The detected p-value or correlation coefficient is not a valid number."): Exit Function
    reject = PValueDecision(operator, pValue)
    If IsNull(reject) Then Set ParsePearson = ErrorResult(PValueDecisionError(operator, pValue)): Exit Function
    shownP = DisplayedP(operator, pValue)
    If correlation > 0 Then
        direction = "positive"
    ElseIf correlation < 0 Then
        direction = "negative"
    Else
        direction = "neither positive nor negative"
    End If
    magnitude = Abs(correlation)
    If magnitude < 0.3 Then
        strength = "weak"
    ElseIf magnitude < 0.7 Then
        strength = "mild"
    Else
        strength = "strong"
    End If
    If reject Then decision = "reject" Else decision = "fail to reject"
    If tMatch Is Nothing Then
        statisticText = " The p-value is " & shownP & "."
    Else
        statisticText = " The test statistic is t = " & FormatFinalNumber(JsNumber(tMatch.Group(1))) & " with a p-value of " & shownP & "."
    End If
    If reject Then
        conclusion = "With r = " & FormatFinalNumber(correlation) & ", there is a statistically significant " & strength & " " & direction & " correlation between the variables."
    Else
        conclusion = "There is no statistically significant correlation between the two variables."
    End If
    Set result = NewMap()
    result("variable1") = variable1
    result("variable2") = variable2
    result("correlation") = correlation
    result("direction") = direction
    result("strength") = strength
    result("pValue") = pValue
    result("displayedP") = shownP
    result("reject") = CBool(reject)
    result("t") = GroupOrNull(tMatch, 1)
    result("df") = GroupOrNull(dfMatch, 1)
    If ciMatch Is Nothing Then result("confidenceInterval") = Null Else Set result("confidenceInterval") = NewList(ciMatch.Group(1), ciMatch.Group(2))
    result("prose") = SYM_H0() & ": There is no correlation between " & variable1 & " and " & variable2 & " (" & SYM_RHO() & " = 0)." & vbLf & vbLf & _
        SYM_H1() & ": There is a correlation between " & variable1 & " and " & variable2 & " (" & SYM_RHO() & " " & SYM_NE() & " 0)." & vbLf & vbLf & _
        "A correlation test was conducted between " & variable1 & " and " & variable2 & "." & statisticText & " Since the p-value is " & _
        PValueComparison(operator, pValue) & " 0.05, we " & decision & " the null hypothesis at the 5% significance level. " & conclusion
    Set ParsePearson = result
End Function

' Local port of opInference's parsePrintedInterval: clsMap { values, lower, upper } or Nothing.
Private Function ParsePrintedIntervalLocal(ByVal text As String, ByVal minimumValues As Long) As clsMap
    Dim lines As clsList, i As Long, line As String, hasRowPrefix As Boolean, cleaned As String
    Dim found As Collection, m As clsMatch, values As clsList, v As Double, result As clsMap
    Set lines = Rx("\r?\n").Split(text)
    For i = 0 To lines.Count - 1
        line = CStr(lines(i))
        hasRowPrefix = RxTest(line, "^\s*\[\s*\d+\s*,?\s*\]\s*")
        cleaned = RxReplace(line, "^\s*\[\s*\d+\s*,?\s*\]\s*", "")
        If Not hasRowPrefix And Not RxTest(cleaned, "^\s*[-+]?(?:\d|\.)") Then GoTo NextLine
        Set values = NewList()
        Set found = Rx("[-+]?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?", "gi").MatchAll(cleaned)
        For Each m In found
            v = JsNumber(m.Value)
            If JsIsFinite(v) Then values.Push v
        Next m
        If values.Count >= minimumValues Then
            Set result = NewMap()
            Set result("values") = values
            If values.Count >= 2 Then
                result("lower") = JsMin(values(0), values(1))
                result("upper") = JsMax(values(0), values(1))
            Else
                result("lower") = Null
                result("upper") = Null
            End If
            Set ParsePrintedIntervalLocal = result
            Exit Function
        End If
NextLine:
    Next i
End Function

Public Function ParseUniversalInterval(ByVal text As String) As clsMap
    Dim parsed As clsMap, header As clsMatch, prediction As Boolean, confidence As Double, label As String
    Dim bounds As String, proportion As Boolean, prose As String, c As String, result As clsMap, headerConfidence As Double
    Set parsed = ParsePrintedIntervalLocal(text, 2)
    If parsed Is Nothing Then Set ParseUniversalInterval = ErrorResult("I could not find the two printed interval limits."): Exit Function
    Set header = RxMatch(text, "\b[ul](CI|PI)([A-Za-z][A-Za-z0-9]*?)(\d{2})\b", "i")
    If Not header Is Nothing Then prediction = (JsUpper(header.Group(1)) = "PI")
    If Not prediction Then prediction = RxTest(text, "\b(?:prediction|predictive)\s+interval\b", "i")
    confidence = 95
    If Not header Is Nothing Then
        headerConfidence = JsNumber(header.Group(3))
        If headerConfidence >= 80 And headerConfidence < 100 Then confidence = headerConfidence
    End If
    If Not header Is Nothing Then label = header.Group(2)
    If Len(label) = 0 Then label = "the parameter"
    bounds = "[" & FormatFinalNumber(parsed("lower")) & ", " & FormatFinalNumber(parsed("upper")) & "]"
    proportion = RxTest(text, "proportion|sample_proportion|success(?:es|_value)", "i")
    c = JsNumberToString(confidence)
    If prediction Then
        prose = "For the " & c & "% Prediction Interval: " & bounds & ", given the observed " & label & ", the " & label & " for a new observation will lie within this interval with a " & c & _
            "% level of confidence. More generally, with repeated random sampling from the same population, " & c & "% of such constructed predictive intervals would contain the new " & label & "."
    ElseIf proportion Then
        prose = "The " & c & "% confidence interval for the population proportion of " & label & " is " & bounds & ". We are " & c & "% confident that this interval contains the true population proportion. " & _
            "More generally, with repeated random sampling from the same population, " & c & "% of such constructed confidence intervals would contain the true population proportion."
    Else
        prose = "The " & c & "% confidence interval for the population mean of " & label & " is " & bounds & ". We are " & c & "% confident that this interval contains the true population mean of " & label & _
            ". More generally, with repeated random sampling from the same population, " & c & "% of such constructed confidence intervals would contain the true population mean."
    End If
    Set result = NewMap()
    result("prose") = prose
    result("type") = IIf(prediction, "prediction", "confidence")
    result("parameter") = IIf(proportion, "proportion", "mean")
    result("confidence") = confidence
    result("label") = label
    Set ParseUniversalInterval = result
End Function

' Local port of opInference's parseOneSampleOutput / parseTwoSampleOutput: clsMap
' { operator, pValue, label, value (Null when absent) } or { error }.
Private Function ParseTestOutputLocal(ByVal text As String, ByVal oneSample As Boolean) As clsMap
    Dim flat As String, pMatch As clsMatch, statMatch As clsMatch, result As clsMap, pValue As Double
    flat = JsTrim(RxReplace(RxReplace(text, "^\s{0,3}#{1,3}\s{0,3}", "", "gm"), "\s+", " ", "g"))
    If oneSample Then
        Set pMatch = RxMatch(flat, "p(?:-|_|\s{0,3})value\s{0,3}(<=|>=|<|>|=)\s{0,3}" & NumberPattern(), "i")
        If pMatch Is Nothing Then Set ParseTestOutputLocal = ErrorResult("I could not find a p-value in the pasted test output."): Exit Function
        pValue = JsNumber(pMatch.Group(2))
        Set statMatch = RxMatch(flat, "\b(t|z|X-squared)\s{0,3}=\s{0,3}" & NumberPattern(), "i")
        If Not JsIsFinite(pValue) Then Set ParseTestOutputLocal = ErrorResult("The detected p-value is not valid."): Exit Function
    Else
        Set pMatch = RxMatch(flat, "p(?:-|\s{0,3})value\s{0,3}(<=|>=|<|>|=)\s{0,3}" & NumberPattern(), "i")
        Set statMatch = RxMatch(flat, "\b(t|F)\s{0,3}=\s{0,3}" & NumberPattern(), "i")
        If pMatch Is Nothing Then Set ParseTestOutputLocal = ErrorResult("I could not find a p-value in the pasted output."): Exit Function
        pValue = JsNumber(pMatch.Group(2))
    End If
    Set result = NewMap()
    result("operator") = pMatch.Group(1)
    result("pValue") = pValue
    result("label") = GroupOrNull(statMatch, 1)
    result("value") = GroupOrNull(statMatch, 2)
    Set ParseTestOutputLocal = result
End Function

Private Function NullRelation(ByVal relation As String) As String
    If relation = "greater than" Then
        NullRelation = "less than or equal to"
    ElseIf relation = "less than" Then
        NullRelation = "greater than or equal to"
    Else
        NullRelation = "equal to"
    End If
End Function

Private Function DataLine(ByVal text As String, ByVal fallback As String) As String
    Dim m As clsMatch
    Set m = RxMatch(text, "^\s*data\s*:\s*(.+)$", "mi")
    If Not m Is Nothing Then DataLine = JsTrim(m.Group(1))
    If Len(DataLine) = 0 Then DataLine = fallback
End Function

Public Function ParseUniversalOneSample(ByVal text As String) As clsMap
    Dim r As clsMap, flat As String, data As String, alternative As clsMatch, nullValue As String, m As clsMatch
    Dim parameter As String, relation As String, reject As Variant, operator As String, pValue As Double
    Dim pDisplay As String, resultLead As String, hypotheses As String, conclusion As String, res As clsMap
    Set r = ParseTestOutputLocal(text, True)
    If r.Exists("error") Then Set ParseUniversalOneSample = r: Exit Function
    flat = JsTrim(RxReplace(text, "\s+", " ", "g"))
    data = DataLine(text, "the tested population parameter")
    Set alternative = RxMatch(flat, "alternative hypothesis:\s*true\s+(mean|p|proportion)\s+is\s+(not equal to|greater than|less than)\s+([-+0-9.e]+)", "i")
    If Not alternative Is Nothing Then nullValue = alternative.Group(3)
    If Len(nullValue) = 0 Then
        Set m = RxMatch(flat, "null probability\s+([-+0-9.e]+)", "i")
        If Not m Is Nothing Then nullValue = m.Group(1)
    End If
    If RxTest(text, "proportion|X-squared|\bz\s*=", "i") Then parameter = "population proportion" Else parameter = "population mean"
    If Not alternative Is Nothing Then relation = alternative.Group(2) Else relation = "different from"
    operator = r("operator"): pValue = r("pValue")
    reject = PValueDecision(operator, pValue)
    If IsNull(reject) Then Set ParseUniversalOneSample = ErrorResult(PValueDecisionError(operator, pValue)): Exit Function
    pDisplay = DisplayedP(operator, pValue)
    If IsNull(r("value")) Then
        resultLead = "The p-value is " & pDisplay & "."
    Else
        resultLead = "The test statistic is " & r("label") & " = " & FormatFinalNumber(JsNumber(r("value"))) & " and the p-value is " & pDisplay & "."
    End If
    If Len(nullValue) > 0 Then
        hypotheses = SYM_H0() & ": The " & parameter & " is " & NullRelation(relation) & " " & nullValue & "." & vbLf & vbLf & _
            SYM_H1() & ": The " & parameter & " is " & relation & " " & nullValue & "." & vbLf & vbLf
    End If
    If reject Then
        conclusion = "There is sufficient evidence to conclude that the " & parameter & " for " & data & " is " & relation & " the hypothesized value."
    Else
        conclusion = "There is insufficient evidence to conclude that the " & parameter & " for " & data & " is " & relation & " the hypothesized value."
    End If
    Set res = NewMap()
    Set ParseUniversalOneSample = res
    res("prose") = hypotheses & resultLead & " Since the p-value is " & PValueComparison(operator, pValue) & " 0.05, we " & _
        IIf(reject, "reject", "fail to reject") & " the null hypothesis at the 5% significance level. " & conclusion
End Function

Public Function ParseUniversalTwoSample(ByVal text As String) As clsMap
    Dim r As clsMap, data As String, variance As Boolean, reject As Variant, operator As String, pValue As Double
    Dim pDisplay As String, resultLead As String, alternative As clsMatch, parameter As String, nullValue As String, relation As String, res As clsMap
    Set r = ParseTestOutputLocal(text, False)
    If r.Exists("error") Then Set ParseUniversalTwoSample = r: Exit Function
    data = DataLine(text, "the two groups")
    variance = RxTest(text, "F\s+test\s+to\s+compare\s+two\s+variances", "i")
    operator = r("operator"): pValue = r("pValue")
    reject = PValueDecision(operator, pValue)
    If IsNull(reject) Then Set ParseUniversalTwoSample = ErrorResult(PValueDecisionError(operator, pValue)): Exit Function
    pDisplay = DisplayedP(operator, pValue)
    If IsNull(r("value")) Then
        resultLead = "the p-value is " & pDisplay
    Else
        resultLead = "the test statistic is " & r("label") & " = " & FormatFinalNumber(JsNumber(r("value"))) & " and the p-value is " & pDisplay
    End If
    Set alternative = RxMatch(RxReplace(text, "\s+", " ", "g"), "alternative hypothesis:\s*true\s+(difference in means|mean difference|ratio of variances)\s+is\s+(not equal to|greater than|less than)\s+([-+0-9.e]+)", "i")
    If variance Then parameter = "population variance ratio" Else parameter = "population mean difference"
    If Not alternative Is Nothing Then nullValue = alternative.Group(3)
    If Len(nullValue) = 0 Then nullValue = IIf(variance, "1", "0")
    If Not alternative Is Nothing Then relation = alternative.Group(2) Else relation = "different from"
    Set res = NewMap()
    Set ParseUniversalTwoSample = res
    res("prose") = SYM_H0() & ": The " & parameter & " is " & NullRelation(relation) & " " & nullValue & "." & vbLf & vbLf & _
        SYM_H1() & ": The " & parameter & " is " & relation & " " & nullValue & "." & vbLf & vbLf & "For " & data & ", " & resultLead & _
        ". Since the p-value is " & PValueComparison(operator, pValue) & " 0.05, we " & IIf(reject, "reject", "fail to reject") & _
        " the null hypothesis at the 5% significance level. There is " & IIf(reject, "sufficient", "insufficient") & " evidence to conclude that the " & _
        parameter & " is " & relation & " " & nullValue & "."
End Function

' ================================================================ generate

' Timer entry (the website's setTimeout(generate, 320)).
Public Sub GenerateTimer()
    On Error GoTo Failed
    mGenerationTimer = 0
    Generate
    Exit Sub
Failed:
    ReportError "modApp2.GenerateTimer"
End Sub

Public Sub Generate()
    Dim text As String, routed As clsMap, routedId As String, universal As Boolean
    text = JsTrim(Q("#inputText").Value)
    If Len(text) = 0 Then
        ShowToast "Paste an RStudio output first"
        Q("#inputText").Focus
        Exit Sub
    End If
    universal = (StateId() = "universal")
    If universal Then
        Set routed = Detect(text)
    Else
        Set routed = NewMap()
        routed("id") = AppState("id"): routed("title") = AppState("title"): routed("tutorial") = AppState("tutorial")
    End If
    If Not routed Is Nothing Then routedId = CStr(routed("id"))
    If universal And Len(routedId) > 0 Then
        If GenerateRouted(text, routed) Then Exit Sub
    End If
    RenderGeneratedProse text, routed, universal
End Sub

' The detector's navigating routes; True when the output was handed to an operation.
Private Function GenerateRouted(ByVal text As String, ByVal routed As clsMap) As Boolean
    Dim id As String, title As String, tutorial As String, interval As clsMap, targetPanel As clsNode
    id = routed("id"): title = routed("title"): tutorial = routed("tutorial")
    GenerateRouted = True
    Select Case id
        Case "anova-many"
            SetActive "anova-many", title, tutorial
            TryCallOp "ResetAnovaBranchCopyState", "welch", "all"
            Q("#anovaWelchOutput").Value = IIf(RxTest(text, "Welch\s+ANOVA", "i"), text, "")
            Q("#anovaGamesHowellOutput").Value = IIf(CBool(routed("gamesHowell")), text, "")
            TryCallOp "RenderAnovaWelchInterpretation"
            SetLpPanelCollapsed Q("#anovaWelchPanel"), False
            ScrollIntoViewLater Q("#anovaWelchPanel")
            ShowToast "ANOVA output detected and loaded"
        Case "lpsolve"
            If Not routed.Exists("parsedModel") Then GenerateRouted = False: Exit Function
            SetActive "lpsolve", title, tutorial
            Q("#lpSmartInput").Value = text
            TryCallOp "ApplyParsedLpModel", routed("parsedModel")
            ShowToast "Optimisation model detected and loaded"
        Case "interval"
            Set interval = ParseUniversalInterval(text)
            SetActive "interval", title, tutorial
            If Not interval.Exists("error") Then
                Q("#intervalCalcType").Value = interval("type")
                Q("#intervalParameter").Value = interval("parameter")
                Q("#confidenceLevel").Value = JsNumberToString(interval("confidence"))
                Q("#intervalLabel").Value = interval("label")
            End If
            Q("#intervalResultInput").Value = text
            TryCallOp "SyncIntervalForm"
            TryCallOp "RenderIntervalPastedResult"
            ScrollIntoViewLater Q("#intervalInterpretationPanel")
            ShowToast "Interval output detected and loaded"
        Case "multi-regression"
            SetActive "multi-regression", title, tutorial
            Q("#multiRegressionOutput").Value = text
            TryCallOp "RenderMultiRegressionInterpretation"
            ScrollIntoViewLater Q("#multiRegressionInterpretationPanel")
            ShowToast "Multi-variable regression output detected and loaded"
        Case "confusion"
            SetActive "confusion", title, tutorial
            Q("#confusionOutput").Value = text
            TryCallOp "RenderConfusionInterpretation"
            SetLpPanelCollapsed Q("#confusionInterpretationPanel"), False
            ScrollIntoViewLater Q("#confusionInterpretationPanel")
            ShowToast "Confusion matrix detected and loaded"
        Case "pca"
            SetActive "pca", title, tutorial
            Set targetPanel = Q("#pcaVariancePanel")
            If RxTest(text, "proportion of variance|cumulative proportion", "i") Then
                Q("#pcaSummaryOutput").Value = text
                TryCallOp "RenderPcaVariance"
            ElseIf RxTest(text, "\bPC1\b[\s\S]*\bPC2\b", "i") Then
                Q("#pcaLoadingsOutput").Value = text
                TryCallOp "RenderPcaLoadings"
                Set targetPanel = Q("#pcaLoadingsPanel")
            End If
            SetLpPanelCollapsed targetPanel, False
            ScrollIntoViewLater targetPanel
            ShowToast "PCA output detected and loaded"
        Case "kmeans"
            SetActive "kmeans", title, tutorial
            ShowToast title & " detected and opened"
        Case "glm"
            GenerateGlmRoute text, title, tutorial
        Case Else
            GenerateRouted = False
    End Select
End Function

Private Sub GenerateGlmRoute(ByVal text As String, ByVal title As String, ByVal tutorial As String)
    Dim flattened As String, formulaMatch As clsMatch, datasetMatch As clsMatch, parts As clsList, i As Long
    Dim predictors As String, item As String, code As Variant
    flattened = RxReplace(text, "\s+", " ", "g")
    Set formulaMatch = RxMatch(flattened, "glm\s*\(\s*formula\s*=\s*([^~,(]+?)\s*~\s*([^,]+?)(?=,\s*(?:family|data)\s*=)", "i")
    Set datasetMatch = RxMatch(flattened, "\bdata\s*=\s*([A-Za-z.][A-Za-z0-9._]*)", "i")
    SetActive "glm", title, tutorial
    If Not formulaMatch Is Nothing Then
        SetValueIfPresent "#advResponse", Replace(JsTrim(formulaMatch.Group(1)), "`", "")
        Set parts = JsSplit(formulaMatch.Group(2), "+")
        For i = 0 To parts.Count - 1
            item = Replace(JsTrim(CStr(parts(i))), "`", "")
            If Len(item) > 0 Then
                If Len(predictors) > 0 Then predictors = predictors & ", "
                predictors = predictors & item
            End If
        Next i
        SetValueIfPresent "#advPredictors", predictors
    End If
    If Not datasetMatch Is Nothing Then SetValueIfPresent "#advDataset", datasetMatch.Group(1)
    code = TryCallOp("AdvancedRegressionCode", "glm")
    If Not IsEmpty(code) Then Q("#advancedRegressionCode").TextContent = CStr(code)
    Q("#glmOutput").Value = text
    TryCallOp "RenderGlmInterpretation"
    SetLpPanelCollapsed Q("#glmInterpretationPanel"), False
    ScrollIntoViewLater Q("#glmInterpretationPanel")
    ShowToast "Binary GLM output detected and loaded"
End Sub

' Self-contained parsers (and the "open the operation" message) in #outputBody.
Private Sub RenderGeneratedProse(ByVal text As String, ByVal routed As clsMap, ByVal universal As Boolean)
    Dim routeTitle As String, prose As String, isComplete As Boolean, errText As String, result As clsMap
    Dim routedId As String, message As String, routeBanner As String, heading As String
    If Not routed Is Nothing Then routedId = CStr(routed("id")): routeTitle = CStr(routed("title"))
    If Len(routeTitle) = 0 Then routeTitle = "Unrecognised output"
    Select Case routedId
        Case "shapiro": Set result = ParseShapiro(text)
        Case "pearson": Set result = ParsePearson(text)
        Case "interval": Set result = ParseUniversalInterval(text)
        Case "one-sample": Set result = ParseUniversalOneSample(text)
        Case "two-sample": Set result = ParseUniversalTwoSample(text)
    End Select
    If Not result Is Nothing Then
        If result.Exists("error") Then errText = CStr(result("error")) Else prose = CStr(result("prose"))
        isComplete = (Len(errText) = 0)
    ElseIf Not routed Is Nothing Then
        prose = "This output matches " & routeTitle & ". Open that operation from the library to use its dedicated workflow."
    Else
        prose = "No confident match was found. Select the relevant operation from the library to continue."
    End If
    If Len(errText) > 0 Then message = errText Else message = prose
    If universal Then routeBanner = "<div class=""route-card""><strong>" & EscapeHtml(routeTitle) & "</strong></div>"
    If Len(errText) > 0 Then heading = "Check the pasted output" Else heading = "Generated interpretation"
    Q("#outputBody").InnerHtml = routeBanner & "<div class=""prose-card""><h3>" & heading & "</h3><p id=""generatedProse"" class=""" & _
        IIf(isComplete, "", "waiting-note") & """>" & EscapeHtml(message) & "</p></div>"
End Sub

' ================================================================ textarea fitting


' The browser measures scrollHeight; here the content height is estimated from its lines.
Public Sub FitTextareaToContent(ByVal textarea As clsNode)
    Dim lines As Long, contentHeight As Double, heightPx As Double
    If Not ShellMatches(textarea, COMPACT_TEXTAREA_SELECTOR) Then Exit Sub
    textarea.SetStyle "height", "0px"
    lines = Len(textarea.Value) - Len(Replace(textarea.Value, vbLf, "")) + 1
    contentHeight = lines * 21 + 22 + 2
    heightPx = contentHeight
    If heightPx > 320 Then heightPx = 320
    If heightPx < 88 Then heightPx = 88
    textarea.SetStyle "height", JsNumberToString(heightPx) & "px"
    textarea.SetStyle "overflow-y", "auto"
End Sub

Public Sub FitAllCompactTextareas()
    Dim list As clsList, i As Long
    On Error GoTo Failed
    Set list = QA(COMPACT_TEXTAREA_SELECTOR)
    For i = 0 To list.Count - 1
        FitTextareaToContent list(i)
    Next i
    Exit Sub
Failed:
    ReportError "modApp2.FitAllCompactTextareas"
End Sub

Public Sub OnDocInputFitTextarea(ByVal e As clsEvent)
    On Error GoTo Failed
    FitTextareaToContent e.Target
    Exit Sub
Failed:
    ReportError "modApp2.OnDocInputFitTextarea"
End Sub

' ================================================================ workbench handlers

Public Sub CopyProse()
    Dim proseEl As clsNode, prose As String
    Set proseEl = Q("#generatedProse")
    If Not proseEl Is Nothing Then prose = proseEl.TextContent
    If Len(prose) = 0 Then
        ShowToast "Nothing to copy yet"
        Exit Sub
    End If
    If WriteAnswerText(prose) Then ShowToast "Prose copied" Else ShowToast "Clipboard access is unavailable"
End Sub

Public Sub OnInputTextInput(ByVal e As clsEvent)
    On Error GoTo Failed
    If mGenerationTimer <> 0 Then ClearTimeout mGenerationTimer
    mGenerationTimer = 0
    If Len(JsTrim(e.Target.Value)) = 0 Then
        SetActive AppState("id"), AppState("title"), AppState("tutorial")
        Exit Sub
    End If
    mGenerationTimer = SetTimeout("modApp2.GenerateTimer", 320)
    Exit Sub
Failed:
    ReportError "modApp2.OnInputTextInput"
End Sub

Public Sub OnResetClick(ByVal e As clsEvent)
    On Error GoTo Failed
    Q("#inputText").Value = ""
    SetActive AppState("id"), AppState("title"), AppState("tutorial")
    Q("#inputText").Focus
    Exit Sub
Failed:
    ReportError "modApp2.OnResetClick"
End Sub

Private Sub LoadWorkbenchSample(ByVal inputText As String, ByVal label As String)
    Q("#universalSampleSelect").Value = ""
    Q("#inputText").Value = inputText
    Q("#inputText").DispatchEvent NewEvent("input")
    Q("#inputText").Focus
    If StateId() = "universal" Then ShowToast label & " sample loaded" Else ShowToast "Sample loaded"
End Sub

Public Sub OnUniversalSampleChange(ByVal e As clsEvent)
    Dim value As String, i As Long
    On Error GoTo Failed
    value = Q("#universalSampleSelect").Value
    If Len(value) = 0 Then Exit Sub
    For i = 0 To UniversalSampleCount - 1
        If UniversalRoutes(i) = value Then
            LoadWorkbenchSample UniversalInputs(i), UniversalLabels(i)
            Exit Sub
        End If
    Next i
    Exit Sub
Failed:
    ReportError "modApp2.OnUniversalSampleChange"
End Sub

Public Sub OnSingleSampleClick(ByVal e As clsEvent)
    Dim at As Long
    On Error GoTo Failed
    at = SampleIndex(StateId())
    If at >= 0 Then LoadWorkbenchSample SampleInputs(at), ""
    Exit Sub
Failed:
    ReportError "modApp2.OnSingleSampleClick"
End Sub

' navigator.clipboard.readText(): unavailable in test mode (like the reference harness).
Public Sub OnPasteClick(ByVal e As clsEvent)
    Dim text As String
    On Error GoTo Failed
    If AppTestMode() Then
        ShowToast "Use Ctrl/Cmd + V to paste"
        Exit Sub
    End If
    text = ClipboardRead()
    Q("#inputText").Value = text
    Q("#inputText").DispatchEvent NewEvent("input")
    ShowToast "Clipboard pasted"
    Exit Sub
Failed:
    ReportError "modApp2.OnPasteClick"
End Sub

Public Sub OnCopyClick(ByVal e As clsEvent)
    On Error GoTo Failed
    CopyProse
    Exit Sub
Failed:
    ReportError "modApp2.OnCopyClick"
End Sub
