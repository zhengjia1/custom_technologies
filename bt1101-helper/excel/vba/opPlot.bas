Attribute VB_Name = "opPlot"
' Plotting & tables (website app-src.js lines 3198-3452 and 4259-4373): the form side of
' renderGgplotCode, the categorised searchable plot-type combobox, contingency style label
' sync, samples and copying. The R code recipes themselves live in opPlot2.
Option Explicit

' visiblePlotWorkflowOptions / activePlotWorkflowIndex
Private mVisibleOptions As clsList
Private mActiveIndex As Long

Public Sub InitOpPlot()
    Dim initial As clsNode, opts As clsList, i As Long, searchText As String, committed As String
    Set mVisibleOptions = NewList()
    mActiveIndex = -1
    searchText = JsTrim(Q("#plotWorkflowSearch").Value)
    Set opts = Q("#plotWorkflow").Options
    For i = 0 To opts.Count - 1
        If JsTrim(opts(i).TextContent) = searchText Then Set initial = opts(i): Exit For
    Next i
    If initial Is Nothing Then committed = Q("#plotWorkflow").Value Else committed = initial.Value
    If Len(committed) = 0 Then committed = Q("#plotWorkflow").Value
    Q("#plotWorkflow").DataSet "committedChoice", committed

    With Q("#plotWorkflowSearch")
        .AddEventListener "focus", "opPlot.OnSearchFocus"
        .AddEventListener "click", "opPlot.OnSearchFocus"
        .AddEventListener "input", "opPlot.OnSearchInput"
        .AddEventListener "change", "opPlot.OnSearchChange"
        .AddEventListener "keydown", "opPlot.OnSearchKeydown"
    End With
    Q("#plotWorkflowMenu").AddEventListener "click", "opPlot.OnMenuClick"
    Doc.AddEventListener "click", "opPlot.OnDocumentClick"
    Q("#plotStyle").AddEventListener "change", "opPlot.OnPlotStyleChange"
    Q("#plotStyle").AddEventListener "change", "opPlot.OnPlotStyleRandomize"
    Q("#plotXTransform").AddEventListener "change", "opPlot.OnRenderEvent"
    Q("#plotYTransform").AddEventListener "change", "opPlot.OnRenderEvent"
    Dim inputs As clsList
    Set inputs = QA(".plotting-input")
    For i = 0 To inputs.Count - 1
        inputs(i).AddEventListener "input", "opPlot.OnRenderEvent"
    Next i
    Q("#plotWorkflow").AddEventListener "change", "opPlot.OnRenderEvent"
    Q("#plotStyle").AddEventListener "change", "opPlot.OnRenderEvent"
    Q("#loadGgplotSampleBtn").AddEventListener "click", "opPlot.OnLoadSample"
    Q("#copyGgplotCodeBtn").AddEventListener "click", "opPlot.OnCopyCode"
    Q("#loadGgplotSampleBtn").AddEventListener "click", "opPlot.OnLoadSampleStyle"
End Sub

' ---------------------------------------------------------------- renderGgplotCode

' One plot type's form settings (the website's `settings` entry); Nothing when unknown.
Private Function PlotSetting(ByVal workflow As String) As clsMap
    Select Case workflow
    Case "summary": Set PlotSetting = MakeSetting("Grouping variable", "Numeric variables, separated by commas", False, "", False, False, False, _
        "Compare means across categories", "Select a categorical grouping variable and at least one numeric variable. The data are summarized before plotting.", "dplyr + tidyr + ggplot2")
    Case "grouped-summary": Set PlotSetting = MakeSetting("Numeric variable", "", True, "Categorical grouping variable", False, False, False, _
        "Compare mean and standard deviation", "Calculate the mean and standard deviation for each category, print the summary table, then compare the categories within Mean and Standard deviation.", "dplyr + knitr + base R")
    Case "frequency": Set PlotSetting = MakeSetting("Categorical variable", "", False, "", False, False, False, _
        "Categorical frequency", "Create the frequency table first, then show every category and count in a horizontal bar plot.", "dplyr + knitr + base R")
    Case "pie": Set PlotSetting = MakeSetting("Categorical variable (5 groups or fewer)", "", False, "", False, False, False, _
        "Parts of a whole", "Use a pie chart only for a small number of categories. Percentage labels are calculated explicitly.", "dplyr + base R")
    Case "contingency": Set PlotSetting = MakeSetting("First categorical variable", "", True, "Second categorical variable", True, False, False, _
        "Two-way categorical relationship", "Build a contingency table before choosing a side-by-side, stacked, or mosaic display.", "knitr + base R")
    Case "grouped-means": Set PlotSetting = MakeSetting("First categorical variable", "Numeric variable", True, "Second categorical variable", False, False, False, _
        "Mean across two categories", "Summarize the numeric response for every combination of the two categorical variables, then plot those means.", "dplyr + knitr + base R")
    Case "histogram": Set PlotSetting = MakeSetting("Continuous variable", "", False, "", False, True, False, _
        "Distribution shape", "Use a histogram for one continuous variable. The generated code also converts its bins into a frequency table.", "knitr + base R")
    Case "boxplot": Set PlotSetting = MakeSetting("Numerical variable", "", False, "", False, False, False, _
        "Spread and outliers", "Generate both standard (1.5 IQR) and extreme (3 IQR) outlier checks, then isolate mild-only outliers.", "base R")
    Case "grouped-boxplot": Set PlotSetting = MakeSetting("Categorical grouping variable", "Numerical variable", False, "", False, False, False, _
        "Numerical values across groups", "Compare the numerical distribution across categories and add each group's mean as a red point.", "base R")
    Case "scatter": Set PlotSetting = MakeSetting("X numerical variable", "Y numerical variable", False, "", False, False, False, _
        "Two numerical variables", "Plot the relationship between two variables. Use the correlation or regression operation for a statistical test or fitted model.", "base R")
    Case "line": Set PlotSetting = MakeSetting("Ordered X variable (for example, year)", "Y numerical variable", False, "", False, False, False, _
        "Change over an ordered axis", "Use a line plot only when the X values have a meaningful order. Numeric axes use round ticks; date/time axes use calendar labels.", "base R")
    Case "pareto": Set PlotSetting = MakeSetting("Non-negative value column", "", False, "", False, False, True, _
        "Pareto contribution analysis", "Find the smallest number of highest-value observations that reach the target share of the total, then print the result as one sentence.", "dplyr")
    Case Else: Set PlotSetting = Nothing
    End Select
End Function

Private Function MakeSetting(ByVal x As String, ByVal y As String, ByVal hasGroup As Boolean, ByVal groupLabel As String, _
        ByVal hasStyle As Boolean, ByVal hasBins As Boolean, ByVal hasThreshold As Boolean, _
        ByVal title As String, ByVal guide As String, ByVal kind As String) As clsMap
    Dim s As clsMap
    Set s = NewMap()
    s("x") = x: s("y") = y: s("group") = hasGroup: s("groupLabel") = groupLabel
    s("style") = hasStyle: s("bins") = hasBins: s("threshold") = hasThreshold
    s("title") = title: s("guide") = guide: s("type") = kind
    Set MakeSetting = s
End Function

' transformAxes: whether the X and Y axes accept a transformation.
Private Sub TransformAxes(ByVal workflow As String, ByRef canX As Boolean, ByRef canY As Boolean)
    Select Case workflow
        Case "summary", "grouped-means", "grouped-boxplot": canX = False: canY = True
        Case "histogram", "boxplot": canX = True: canY = False
        Case "scatter", "line": canX = True: canY = True
        Case Else: canX = False: canY = False
    End Select
End Sub

Public Sub RenderGgplotCode()
    Dim dataset As String, workflow As String, x As String, yText As String, group As String, style As String
    Dim xTransform As String, yTransform As String, binsText As String, threshold As Double
    Dim setting As clsMap, canX As Boolean, canY As Boolean, activeX As String, activeY As String
    Dim plotKind As String, shownX As String, shownY As String, shownGroup As String
    Dim suggestedTitle As String, title As String, tableCaption As String, code As String, ready As Boolean
    Dim workflowSelect As clsNode, label As Variant, xFallback As String, yFallback As String, groupFallback As String

    dataset = JsTrim(Q("#plotDataset").Value)
    If dataset = "" Then dataset = "dataset"
    Set workflowSelect = Q("#plotWorkflow")
    workflow = workflowSelect.Value
    x = JsTrim(Q("#plotX").Value)
    yText = JsTrim(Q("#plotY").Value)
    group = JsTrim(Q("#plotGroup").Value)
    style = Q("#plotStyle").Value
    xTransform = Q("#plotXTransform").Value
    yTransform = Q("#plotYTransform").Value
    binsText = JsTrim(Q("#plotBins").Value)
    If binsText = "" Then binsText = "Sturges"
    threshold = JsNumber(Q("#plotThreshold").Value)

    Set setting = PlotSetting(workflow)
    If setting Is Nothing Then
        Q("#plotTitle").Placeholder = "Select a plot type"
        Q("#ggplotCode").TextContent = "Choose a plot type from the search results."
        Q("#copyGgplotCodeBtn").Disabled = True
        Q("#loadGgplotSampleBtn").Disabled = True
        Exit Sub
    End If
    Q("#loadGgplotSampleBtn").Disabled = False
    TransformAxes workflow, canX, canY
    Q("#plotXTransform").Disabled = Not canX
    Q("#plotYTransform").Disabled = Not canY
    Q("#plotXTransformField").ToggleClass "inapplicable", Not canX
    Q("#plotYTransformField").ToggleClass "inapplicable", Not canY
    Q("#plotXLabel").TextContent = setting("x")
    Q("#plotYField").ToggleClass "hidden", Len(setting("y")) = 0
    Q("#plotAgainstText").ToggleClass "hidden", Len(setting("y")) = 0
    If Len(setting("y")) > 0 Then Q("#plotYLabel").TextContent = setting("y") Else Q("#plotYLabel").TextContent = "Numeric variable"
    Q("#plotGroupField").ToggleClass "hidden", Not setting("group")
    Q("#plotGroupLead").ToggleClass "hidden", Not setting("group")
    If workflow = "contingency" Or workflow = "grouped-summary" Then Q("#plotGroupLead").TextContent = "against" Else Q("#plotGroupLead").TextContent = ", grouped by"
    If Len(setting("groupLabel")) > 0 Then Q("#plotGroupLabel").TextContent = setting("groupLabel") Else Q("#plotGroupLabel").TextContent = "Second categorical variable"
    Q("#plotStyleField").ToggleClass "hidden", Not setting("style")
    Q("#plotBinsField").ToggleClass "hidden", Not setting("bins")
    Q("#plotThresholdField").ToggleClass "hidden", Not setting("threshold")
    Q("#plotGuidanceTitle").TextContent = setting("title")
    Q("#plotGuidance").TextContent = setting("guide")
    Q("#plotCodeType").TextContent = setting("type")

    If canX Then activeX = xTransform Else activeX = "none"
    If canY Then activeY = yTransform Else activeY = "none"
    label = workflowSelect.DataGet("plotLabel")
    If IsNull(label) Then plotKind = "" Else plotKind = CStr(label)
    If plotKind = "" Then plotKind = JsTrim(workflowSelect.SelectedOption.TextContent)
    If x = "" Then xFallback = Q("#plotX").Placeholder Else xFallback = x
    If yText = "" Then yFallback = Q("#plotY").Placeholder Else yFallback = yText
    If group = "" Then groupFallback = Q("#plotGroup").Placeholder Else groupFallback = group
    shownX = opPlot2.TransformLabel(xFallback, activeX)
    shownY = opPlot2.TransformLabel(yFallback, activeY)
    shownGroup = ChartDisplayLabel(groupFallback)
    If Len(setting("y")) > 0 Then
        suggestedTitle = plotKind & " of " & shownY & " against " & shownX
    ElseIf workflow = "contingency" Or workflow = "grouped-summary" Then
        suggestedTitle = plotKind & " of " & shownX & " against " & shownGroup
    Else
        suggestedTitle = plotKind & " of " & shownX
    End If
    suggestedTitle = JsUpper(JsCharAt(suggestedTitle, 0)) & JsSlice(suggestedTitle, 1)
    Q("#plotTitle").Placeholder = suggestedTitle
    title = JsTrim(Q("#plotTitle").Value)
    If title = "" Then title = suggestedTitle
    ' Tables have their own captions, independent of the chart title and arrangement.
    Select Case workflow
        Case "contingency": tableCaption = "Frequency table of " & x & " by " & group
        Case "grouped-means": tableCaption = "Table of mean " & yText & " by " & x & " and " & group
        Case "grouped-summary": tableCaption = "Descriptive statistics table for " & x & " by " & group
        Case Else: tableCaption = "Frequency table of " & opPlot2.TransformLabel(x, activeX)
    End Select

    code = opPlot2.PlotRecipeCode(workflow, dataset, x, yText, group, style, activeX, activeY, binsText, threshold, title, tableCaption, ready)
    If ready Then Q("#ggplotCode").TextContent = code Else Q("#ggplotCode").TextContent = "Complete the required plotting fields."
    Q("#copyGgplotCodeBtn").Disabled = Not ready
End Sub

' ---------------------------------------------------------------- plot-type combobox

Private Sub ClosePlotWorkflowMenu()
    Q("#plotWorkflowMenu").AddClass "hidden"
    Q("#plotWorkflowSearch").SetAttribute "aria-expanded", "false"
    Q("#plotWorkflowSearch").RemoveAttribute "aria-activedescendant"
    mActiveIndex = -1
End Sub

Private Sub RenderPlotWorkflowMenu(Optional ByVal query As String = "")
    Dim normalized As String, opts As clsList, i As Long, opt As clsNode, category As String
    Dim groups As clsMap, keys As clsList, members As clsList, html As String, k As Long, part As String, v As Variant
    normalized = JsLower(JsTrim(query))
    Set mVisibleOptions = NewList()
    Set opts = Q("#plotWorkflow").Options
    For i = 0 To opts.Count - 1
        Set opt = opts(i)
        If normalized = "" Then
            mVisibleOptions.Push opt
        ElseIf JsIncludes(JsLower(JsTrim(opt.TextContent)), normalized) Then
            mVisibleOptions.Push opt
        End If
    Next i
    Set groups = NewMap()
    For i = 0 To mVisibleOptions.Count - 1
        Set opt = mVisibleOptions(i)
        v = opt.DataGet("category")
        If IsNull(v) Then category = "Other" Else category = CStr(v)
        If category = "" Then category = "Other"
        If Not groups.Exists(category) Then groups.Add category, NewList()
        groups(category).Push opt
    Next i
    If mVisibleOptions.Count > 0 Then
        Set keys = groups.Keys
        For k = 0 To keys.Count - 1
            Set members = groups(CStr(keys(k)))
            part = "<div class=""plot-workflow-group""><div class=""plot-workflow-heading"">" & EscapeHtml(CStr(keys(k))) & "</div>"
            For i = 0 To members.Count - 1
                Set opt = members(i)
                part = part & "<button class=""plot-workflow-option"" id=""plotWorkflowOption-" & EscapeHtml(opt.Value) & """ type=""button"" role=""option"" data-option-value=""" & _
                    EscapeHtml(opt.Value) & """>" & EscapeHtml(JsTrim(opt.TextContent)) & "</button>"
            Next i
            html = html & part & "</div>"
        Next k
    Else
        html = "<div class=""plot-workflow-empty"">No matching plot types</div>"
    End If
    Q("#plotWorkflowMenu").InnerHtml = html
    Q("#plotWorkflowMenu").RemoveClass "hidden"
    Q("#plotWorkflowSearch").SetAttribute "aria-expanded", "true"
    mActiveIndex = -1
End Sub

Private Sub SetActivePlotWorkflow(ByVal index As Long)
    Dim n As Long, buttons As clsList, i As Long
    n = mVisibleOptions.Count
    If n = 0 Then Exit Sub
    mActiveIndex = (index + n) Mod n
    If mActiveIndex < 0 Then mActiveIndex = mActiveIndex + n
    Set buttons = QA("#plotWorkflowMenu .plot-workflow-option")
    For i = 0 To buttons.Count - 1
        buttons(i).ToggleClass "is-active", (i = mActiveIndex)
    Next i
    Q("#plotWorkflowSearch").SetAttribute "aria-activedescendant", "plotWorkflowOption-" & mVisibleOptions(mActiveIndex).Value
End Sub

Private Function ContingencyAlias(ByVal value As String) As String
    Select Case value
        Case "contingency-grouped": ContingencyAlias = "grouped"
        Case "contingency-stacked": ContingencyAlias = "stacked"
        Case "contingency-mosaic": ContingencyAlias = "mosaic"
    End Select
End Function

Private Sub SyncPlotWorkflowSearch()
    Dim search As clsNode, sel As clsNode, query As String, opts As clsList, i As Long, match As clsNode
    Dim previousChoice As String, v As Variant, alias As String
    Set search = Q("#plotWorkflowSearch")
    Set sel = Q("#plotWorkflow")
    query = JsLower(JsTrim(search.Value))
    Set opts = sel.Options
    For i = 0 To opts.Count - 1
        If JsLower(JsTrim(opts(i).TextContent)) = query Then Set match = opts(i): Exit For
    Next i
    v = sel.DataGet("committedChoice")
    If Not IsNull(v) Then previousChoice = CStr(v)
    If Not match Is Nothing Then alias = ContingencyAlias(match.Value)
    If Len(alias) > 0 Then
        sel.Value = "contingency"
        Q("#plotStyle").Value = alias
    ElseIf Not match Is Nothing Then
        sel.Value = match.Value
    Else
        sel.Value = ""
    End If
    If match Is Nothing Then sel.DataSet "plotLabel", "" Else sel.DataSet "plotLabel", JsTrim(match.TextContent)
    search.SetAttribute "aria-invalid", IIf(Len(query) > 0 And match Is Nothing, "true", "false")
    sel.DispatchEvent NewEvent("change")
    If Not match Is Nothing Then
        sel.DataSet "committedChoice", match.Value
        If Len(previousChoice) > 0 And previousChoice <> match.Value Then
            ClearCodeVariantCache Q("#ggplotBench")
            RandomizeAnswerOptions Q("#ggplotBench"), True
        End If
    End If
End Sub

' Commits an option's label to the search box (Enter, Tab with one result, menu click).
Private Sub CommitPlotWorkflow(ByVal opt As clsNode)
    Q("#plotWorkflowSearch").Value = JsTrim(opt.TextContent)
    SyncPlotWorkflowSearch
    ClosePlotWorkflowMenu
End Sub

Private Sub CompleteSinglePlotWorkflow(ByVal e As clsEvent)
    Dim search As clsNode, query As String, opts As clsList, matches As clsList, i As Long, delta As Long
    Set search = Q("#plotWorkflowSearch")
    If e.Key = "ArrowDown" Or e.Key = "ArrowUp" Then
        e.PreventDefault
        If Q("#plotWorkflowMenu").HasClass("hidden") Then RenderPlotWorkflowMenu search.Value
        If e.Key = "ArrowDown" Then delta = 1 Else delta = -1
        SetActivePlotWorkflow mActiveIndex + delta
        Exit Sub
    End If
    If e.Key = "Enter" And mActiveIndex >= 0 Then
        e.PreventDefault
        CommitPlotWorkflow mVisibleOptions(mActiveIndex)
        Exit Sub
    End If
    If e.Key = "Escape" Then
        ClosePlotWorkflowMenu
        Exit Sub
    End If
    If e.Key <> "Tab" Then Exit Sub
    query = JsLower(JsTrim(search.Value))
    Set matches = NewList()
    Set opts = Q("#plotWorkflow").Options
    For i = 0 To opts.Count - 1
        If query = "" Then
            matches.Push opts(i)
        ElseIf JsIncludes(JsLower(JsTrim(opts(i).TextContent)), query) Then
            matches.Push opts(i)
        End If
    Next i
    If matches.Count = 1 Then
        CommitPlotWorkflow matches(0)
        Exit Sub
    End If
    If matches.Count >= 2 Then
        e.PreventDefault
        RenderPlotWorkflowMenu query
        SetActivePlotWorkflow 0
        Exit Sub
    End If
    ClosePlotWorkflowMenu
End Sub

Private Sub SyncContingencyStyleLabel()
    Dim label As String
    If Q("#plotWorkflow").Value <> "contingency" Then Exit Sub
    Select Case Q("#plotStyle").Value
        Case "grouped": label = "side-by-side bar plot"
        Case "stacked": label = "stacked bar plot"
        Case "mosaic": label = "mosaic plot"
    End Select
    Q("#plotWorkflowSearch").Value = label
    Q("#plotWorkflow").DataSet "plotLabel", label
    RenderGgplotCode
End Sub

' ---------------------------------------------------------------- event handlers

Public Sub OnSearchFocus(ByVal e As clsEvent)
    On Error GoTo Failed
    RenderPlotWorkflowMenu
    Exit Sub
Failed:
    ReportError "opPlot.OnSearchFocus"
End Sub

Public Sub OnSearchInput(ByVal e As clsEvent)
    On Error GoTo Failed
    SyncPlotWorkflowSearch
    RenderPlotWorkflowMenu Q("#plotWorkflowSearch").Value
    Exit Sub
Failed:
    ReportError "opPlot.OnSearchInput"
End Sub

Public Sub OnSearchChange(ByVal e As clsEvent)
    On Error GoTo Failed
    SyncPlotWorkflowSearch
    Exit Sub
Failed:
    ReportError "opPlot.OnSearchChange"
End Sub

Public Sub OnSearchKeydown(ByVal e As clsEvent)
    On Error GoTo Failed
    CompleteSinglePlotWorkflow e
    Exit Sub
Failed:
    ReportError "opPlot.OnSearchKeydown"
End Sub

Public Sub OnMenuClick(ByVal e As clsEvent)
    Dim button As clsNode, opts As clsList, i As Long, wanted As String, v As Variant
    On Error GoTo Failed
    If e.Target Is Nothing Then Exit Sub
    If e.Target.NodeType <> 1 Then Exit Sub
    Set button = e.Target.Closest("[data-option-value]")
    If button Is Nothing Then Exit Sub
    v = button.DataGet("optionValue")
    If Not IsNull(v) Then wanted = CStr(v)
    Set opts = Q("#plotWorkflow").Options
    For i = 0 To opts.Count - 1
        If opts(i).Value = wanted Then
            CommitPlotWorkflow opts(i)
            Exit Sub
        End If
    Next i
    Exit Sub
Failed:
    ReportError "opPlot.OnMenuClick"
End Sub

Public Sub OnDocumentClick(ByVal e As clsEvent)
    On Error GoTo Failed
    If e.Target Is Nothing Then Exit Sub
    If e.Target.NodeType <> 1 Then Exit Sub
    If e.Target.Closest(".plot-type-search") Is Nothing Then ClosePlotWorkflowMenu
    Exit Sub
Failed:
    ReportError "opPlot.OnDocumentClick"
End Sub

Public Sub OnPlotStyleChange(ByVal e As clsEvent)
    On Error GoTo Failed
    SyncContingencyStyleLabel
    Exit Sub
Failed:
    ReportError "opPlot.OnPlotStyleChange"
End Sub

Public Sub OnPlotStyleRandomize(ByVal e As clsEvent)
    On Error GoTo Failed
    RandomizeAnswerOptions Q("#ggplotBench"), True
    Exit Sub
Failed:
    ReportError "opPlot.OnPlotStyleRandomize"
End Sub

Public Sub OnRenderEvent(ByVal e As clsEvent)
    On Error GoTo Failed
    RenderGgplotCode
    Exit Sub
Failed:
    ReportError "opPlot.OnRenderEvent"
End Sub

' The first "Use sample" listener: fills the sentence for the current plot type.
Public Sub OnLoadSample(ByVal e As clsEvent)
    Dim workflow As String, sample As Variant
    On Error GoTo Failed
    workflow = Q("#plotWorkflow").Value
    Select Case workflow
        Case "summary": sample = Array("column1", "column2, column3, column4", "", "Mean values by group")
        Case "grouped-summary": sample = Array("Amount", "", "Product", "Mean and SD of Amount across Products")
        Case "frequency": sample = Array("column1", "", "", "Frequency by category")
        Case "pie": sample = Array("column1", "", "", "Percentage by category")
        Case "contingency": sample = Array("column1", "", "column2", "Relationship between categories")
        Case "grouped-means": sample = Array("column1", "column3", "column2", "Average column3 by group")
        Case "histogram": sample = Array("column1", "", "", "Distribution of column1")
        Case "boxplot": sample = Array("column1", "", "", "Boxplot of column1")
        Case "grouped-boxplot": sample = Array("column1", "column2", "", "Column2 across groups")
        Case "scatter": sample = Array("column1", "column2", "", "Scatterplot of column2 against column1")
        Case "line": sample = Array("column1", "column2", "", "Column2 over column1")
        Case "pareto": sample = Array("column1", "", "", "Pareto analysis of column1")
        Case Else
            ' The website destructures an undefined sample and throws here.
            Err.Raise vbObjectError + 513, , "sample is undefined"
    End Select
    If workflow = "grouped-summary" Then Q("#plotDataset").Value = "ST" Else Q("#plotDataset").Value = "dataset"
    Q("#plotX").Value = sample(0)
    Q("#plotY").Value = sample(1)
    Q("#plotGroup").Value = sample(2)
    Q("#plotTitle").Value = sample(3)
    Q("#plotStyle").Value = "grouped"
    Q("#plotBins").Value = "Sturges"
    Q("#plotThreshold").Value = "0.80"
    RenderGgplotCode
    ShowToast "Sample loaded"
    Exit Sub
Failed:
    ReportError "opPlot.OnLoadSample"
End Sub

' The second "Use sample" listener: keeps a contingency alias's arrangement and resets
' the target and transformations.
Public Sub OnLoadSampleStyle(ByVal e As clsEvent)
    Dim v As Variant, intended As String
    On Error GoTo Failed
    v = Q("#plotWorkflow").DataGet("plotLabel")
    If Not IsNull(v) Then
        Select Case CStr(v)
            Case "side-by-side bar plot": intended = "grouped"
            Case "stacked bar plot": intended = "stacked"
            Case "mosaic plot": intended = "mosaic"
        End Select
    End If
    If Q("#plotWorkflow").Value = "contingency" And Len(intended) > 0 Then Q("#plotStyle").Value = intended
    Q("#plotThreshold").Value = "80"
    Q("#plotXTransform").Value = "none"
    Q("#plotYTransform").Value = "none"
    SyncContingencyStyleLabel
    RenderGgplotCode
    Exit Sub
Failed:
    ReportError "opPlot.OnLoadSampleStyle"
End Sub

Public Sub OnCopyCode(ByVal e As clsEvent)
    On Error GoTo Failed
    If Q("#copyGgplotCodeBtn").Disabled Then Exit Sub
    If WriteAnswerText(Q("#ggplotCode").TextContent) Then
        ShowToast "R code copied"
    Else
        ShowToast "Clipboard access is unavailable"
    End If
    Exit Sub
Failed:
    ReportError "opPlot.OnCopyCode"
End Sub
