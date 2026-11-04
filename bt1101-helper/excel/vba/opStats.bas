Attribute VB_Name = "opStats"
' Tutorial 4 operations: Summarize for stats (dynamic column fields, statistic toggles,
' single/multi-column transition toasts), Skewness & kurtosis (with the kurtosis /
' excess-kurtosis toggle) and the narrow-to-wide table generator.
' Ported from app-src.js lines 1591-1695, 3782-3888 and 4704-4742.
Option Explicit

' summaryWasMultiColumn (website line 1567, used only here) and shapeMetric.
Private mSummaryWasMultiColumn As Boolean
Private mShapeMetric As String

' The summaryStatistics table: key, label, selected by default, R expression ({s} = source).
Private mStatKeys As clsList
Private mStatLabels As clsList
Private mStatSelected As clsList
Private mStatExpressions As clsList

' ---------------------------------------------------------------- startup

' Website lines 1591 (summaryStatistics), 1684-1695 (statistic switches, first column
' field, dataset listener, first render) and 3822 (shapeMetric).
Public Sub InitOpStats()
    Dim html As String, i As Long
    mSummaryWasMultiColumn = False
    mShapeMetric = "kurtosis"
    DefineSummaryStatistics
    For i = 0 To mStatKeys.Count - 1
        html = html & "<label class=""kmeans-scale-toggle""><input type=""checkbox"" role=""switch"" value=""" & mStatKeys(i) & """" & _
            IIf(mStatSelected(i), " checked", "") & "><span>" & mStatLabels(i) & "</span></label>"
    Next i
    Q("#summaryStatisticFields").InnerHtml = html
    Q("#summaryStatisticFields").AddEventListener "change", "opStats.OnSummaryStatisticsChange"
    AddSummaryColumnField
    Q("#summaryDataset").AddEventListener "input", "opStats.OnSummaryDatasetInput"
    RenderSummaryCode
End Sub

' Website lines 4704-4742: wide-table, skewness/kurtosis and summary copy listeners.
Public Sub InitOpStatsEvents()
    Dim list As clsList, i As Long
    Set list = QA(".wide-input")
    For i = 0 To list.Count - 1
        list(i).AddEventListener "input", "opStats.OnWideInput"
    Next i
    Q("#loadWideSampleBtn").AddEventListener "click", "opStats.OnLoadWideSample"
    Q("#copyWideCodeBtn").AddEventListener "click", "opStats.OnCopyWideCode"
    Set list = QA(".metric-option")
    For i = 0 To list.Count - 1
        list(i).AddEventListener "click", "opStats.OnMetricOptionClick"
    Next i
    Q("#skewnessInput").AddEventListener "input", "opStats.OnShapeInput"
    Q("#kurtosisInput").AddEventListener "input", "opStats.OnShapeInput"
    Q("#loadShapeSampleBtn").AddEventListener "click", "opStats.OnLoadShapeSample"
    Q("#copyShapeProseBtn").AddEventListener "click", "opStats.OnCopyShapeProse"
    Q("#copySummaryCodeBtn").AddEventListener "click", "opStats.OnCopySummaryCode"
End Sub

Private Sub DefineSummaryStatistic(ByVal key As String, ByVal label As String, ByVal selected As Boolean, ByVal expression As String)
    mStatKeys.Push key
    mStatLabels.Push label
    mStatSelected.Push selected
    mStatExpressions.Push expression
End Sub

Private Sub DefineSummaryStatistics()
    Set mStatKeys = NewList()
    Set mStatLabels = NewList()
    Set mStatSelected = NewList()
    Set mStatExpressions = NewList()
    DefineSummaryStatistic "n", "Non-missing count (n)", True, "sum(!is.na({s}))"
    DefineSummaryStatistic "mean", "Mean", True, "mean({s}, na.rm = TRUE)"
    DefineSummaryStatistic "sd", "Standard deviation", True, "sd({s}, na.rm = TRUE)"
    DefineSummaryStatistic "median", "Median", True, "median({s}, na.rm = TRUE)"
    DefineSummaryStatistic "skew", "Skewness", True, "psych::skew({s}, na.rm = TRUE)"
    DefineSummaryStatistic "kurtosis", "Excess kurtosis", True, "psych::kurtosi({s}, na.rm = TRUE)"
    DefineSummaryStatistic "min", "Minimum", False, "min({s}, na.rm = TRUE)"
    DefineSummaryStatistic "max", "Maximum", False, "max({s}, na.rm = TRUE)"
    DefineSummaryStatistic "rows", "Total rows", False, "length({s})"
    DefineSummaryStatistic "missing", "Missing count", False, "sum(is.na({s}))"
    DefineSummaryStatistic "variance", "Variance", False, "var({s}, na.rm = TRUE)"
    DefineSummaryStatistic "q1", "First quartile (Q1)", False, "unname(quantile({s}, probs = 0.25, na.rm = TRUE))"
    DefineSummaryStatistic "q3", "Third quartile (Q3)", False, "unname(quantile({s}, probs = 0.75, na.rm = TRUE))"
    DefineSummaryStatistic "iqr", "Interquartile range", False, "IQR({s}, na.rm = TRUE)"
    DefineSummaryStatistic "sum", "Sum", False, "sum({s}, na.rm = TRUE)"
End Sub

' ---------------------------------------------------------------- summarize for stats

' summaryColumns (website line 1585): trimmed, non-empty values of the column fields.
Private Function SummaryColumns() As clsList
    Dim inputs As clsList, out As clsList, i As Long, value As String
    Set inputs = QA("#columnFields .column-input")
    Set out = NewList()
    For i = 0 To inputs.Count - 1
        value = JsTrim(inputs(i).Value)
        If Len(value) > 0 Then out.Push value
    Next i
    Set SummaryColumns = out
End Function

' Indices (into the statistic table) of the checked statistics, in table order.
Public Function SelectedSummaryStatistics() As clsList
    Dim checked As clsList, selected As clsList, out As clsList, i As Long
    Set checked = QA("#summaryStatisticFields input:checked")
    Set selected = NewList()
    For i = 0 To checked.Count - 1
        selected.Push checked(i).Value
    Next i
    Set out = NewList()
    For i = 0 To mStatKeys.Count - 1
        If selected.Includes(mStatKeys(i)) Then out.Push i
    Next i
    Set SelectedSummaryStatistics = out
End Function

Public Function BuildSummaryStatisticsCode(ByVal columns As clsList, ByVal datasetName As String, ByVal statistics As clsList) As String
    Dim dataset As String, sources As clsList, multiple As Boolean, vars As clsList, fields As clsList
    Dim i As Long, k As Long, values As clsList, expression As String, caption As String, joined As String
    dataset = RColumnName(datasetName)
    Set sources = NewList()
    Set vars = NewList()
    For i = 0 To columns.Count - 1
        sources.Push dataset & "$" & RColumnName(columns(i))
        vars.Push """" & RString(columns(i)) & """"
    Next i
    multiple = columns.Count > 1
    Set fields = NewList()
    If multiple Then fields.Push "  vars = c(" & vars.Join(", ") & ")" Else fields.Push "  vars = " & vars.Join(", ")
    For k = 0 To statistics.Count - 1
        expression = mStatExpressions(statistics(k))
        Set values = NewList()
        For i = 0 To sources.Count - 1
            values.Push Replace(expression, "{s}", sources(i))
        Next i
        If multiple Then
            joined = ""
            For i = 0 To values.Count - 1
                If i > 0 Then joined = joined & "," & vbLf
                joined = joined & "    " & values(i)
            Next i
            fields.Push "  " & mStatKeys(statistics(k)) & " = c(" & vbLf & joined & vbLf & "  )"
        Else
            fields.Push "  " & mStatKeys(statistics(k)) & " = " & values(0)
        End If
    Next k
    If multiple Then caption = "Descriptive Statistics Table" Else caption = "Descriptive Statistics Table for " & columns(0)
    BuildSummaryStatisticsCode = "tab.summary <- data.frame(" & vbLf & fields.Join("," & vbLf) & vbLf & ")" & vbLf & vbLf & _
        "print(tab.summary)" & vbLf & vbLf & "knitr::kable(" & vbLf & "  tab.summary," & vbLf & "  row.names = FALSE," & vbLf & _
        "  digits = 2," & vbLf & "  caption = """ & RString(caption) & """" & vbLf & ")"
End Function

Public Function BuildSingleSummaryCode(ByVal columnName As String, ByVal datasetName As String, Optional ByVal statistics As clsList = Nothing) As String
    If statistics Is Nothing Then Set statistics = SelectedSummaryStatistics()
    BuildSingleSummaryCode = BuildSummaryStatisticsCode(NewList(columnName), datasetName, statistics)
End Function

Public Function BuildMultiSummaryCode(ByVal columns As clsList, ByVal datasetName As String, Optional ByVal statistics As clsList = Nothing) As String
    If statistics Is Nothing Then Set statistics = SelectedSummaryStatistics()
    BuildMultiSummaryCode = BuildSummaryStatisticsCode(columns, datasetName, statistics)
End Function

Private Sub RefreshSummaryFields()
    Dim rows As clsList, i As Long, row As clsNode
    Set rows = QA("#columnFields .column-row")
    For i = 0 To rows.Count - 1
        Set row = rows(i)
        row.QuerySelector("input").Placeholder = "column" & (i + 1)
        row.QuerySelector("button").Disabled = (rows.Count = 1)
    Next i
End Sub

Private Sub AddSummaryColumnField(Optional ByVal value As String = "")
    Dim row As clsNode
    Set row = CreateElement("div")
    row.ClassName = "column-row"
    row.InnerHtml = "<input class=""column-input"" type=""text"" value=""" & EscapeHtml(value) & """ aria-label=""Column name"">" & _
        "<button class=""remove-column"" type=""button"" title=""Remove column"" aria-label=""Remove column"">" & ChrW$(&HD7) & "</button>"
    row.QuerySelector("input").AddEventListener "input", "opStats.OnSummaryColumnInput"
    row.QuerySelector("button").AddEventListener "click", "opStats.OnRemoveSummaryColumn"
    Q("#columnFields").AppendChild row
    RefreshSummaryFields
End Sub

Public Sub RenderSummaryCode()
    Dim columns As clsList, statistics As clsList, dataset As String, isMultiColumn As Boolean, code As String
    Set columns = SummaryColumns()
    Set statistics = SelectedSummaryStatistics()
    dataset = JsTrim(Q("#summaryDataset").Value)
    If Len(dataset) = 0 Then dataset = "datafile"
    isMultiColumn = columns.Count > 1
    If isMultiColumn And Not mSummaryWasMultiColumn Then ShowToast "Changed to multi-column summary code"
    If Not isMultiColumn And columns.Count = 1 And mSummaryWasMultiColumn Then ShowToast "Changed to single-column summary code"
    mSummaryWasMultiColumn = isMultiColumn
    If columns.Count = 0 Then
        Q("#summaryModeLabel").TextContent = "waiting for a column"
    ElseIf isMultiColumn Then
        Q("#summaryModeLabel").TextContent = columns.Count & " columns " & ChrW$(&HB7) & " multi-column template"
    Else
        Q("#summaryModeLabel").TextContent = "1 column " & ChrW$(&HB7) & " single-column template"
    End If
    If columns.Count = 0 Then
        code = "Enter a column name to generate the single-column summary code."
    ElseIf statistics.Count = 0 Then
        code = "Select at least one statistic to generate the summary code."
    ElseIf isMultiColumn Then
        code = BuildMultiSummaryCode(columns, dataset, statistics)
    Else
        code = BuildSingleSummaryCode(columns(0), dataset, statistics)
    End If
    Q("#summaryCode").TextContent = code
    Q("#copySummaryCodeBtn").Disabled = (columns.Count = 0 Or statistics.Count = 0)
End Sub

Public Sub OnSummaryStatisticsChange(ByVal e As clsEvent)
    On Error GoTo Failed
    ClearCompletionForAnswerVariant Q("#copySummaryCodeBtn")
    RenderSummaryCode
    RefreshAnswerVariantPreviews
    Exit Sub
Failed:
    ReportError "opStats.OnSummaryStatisticsChange"
End Sub

Public Sub OnSummaryDatasetInput(ByVal e As clsEvent)
    On Error GoTo Failed
    RenderSummaryCode
    Exit Sub
Failed:
    ReportError "opStats.OnSummaryDatasetInput"
End Sub

Public Sub OnSummaryColumnInput(ByVal e As clsEvent)
    Dim inputNode As clsNode, row As clsNode, rows As clsList
    On Error GoTo Failed
    Set inputNode = e.CurrentTarget
    Set row = inputNode.ParentNode
    Set rows = QA("#columnFields .column-row")
    If rows.Count > 0 Then
        If rows(rows.Count - 1) Is row And Len(JsTrim(inputNode.Value)) > 0 Then AddSummaryColumnField
    End If
    RenderSummaryCode
    Exit Sub
Failed:
    ReportError "opStats.OnSummaryColumnInput"
End Sub

Public Sub OnRemoveSummaryColumn(ByVal e As clsEvent)
    Dim row As clsNode
    On Error GoTo Failed
    Set row = e.CurrentTarget.ParentNode
    If Not row Is Nothing Then row.Remove
    If Q("#columnFields .column-row") Is Nothing Then AddSummaryColumnField
    RefreshSummaryFields
    RenderSummaryCode
    Exit Sub
Failed:
    ReportError "opStats.OnRemoveSummaryColumn"
End Sub

Public Sub OnCopySummaryCode(ByVal e As clsEvent)
    On Error GoTo Failed
    If SummaryColumns().Count = 0 Then Exit Sub
    If WriteAnswerText(Q("#summaryCode").TextContent) Then ShowToast "R code copied" Else ShowToast "Clipboard access is unavailable"
    Exit Sub
Failed:
    ReportError "opStats.OnCopySummaryCode"
End Sub

' ---------------------------------------------------------------- wide table

Public Function BuildWideTableCode(ByVal dataset As String, ByVal rowName As String, ByVal columnName As String, ByVal valueName As String, ByVal caption As String) As String
    Dim data As String, row As String, column As String, value As String
    data = RColumnName(dataset)
    row = RColumnName(rowName)
    column = RColumnName(columnName)
    value = RColumnName(valueName)
    If Len(caption) = 0 Then caption = "Descriptive Statistics Table"
    BuildWideTableCode = "# Convert a narrow table to wide format" & vbLf & vbLf & _
        "wide_table <- " & data & " %>%" & vbLf & _
        "  group_by(" & row & ", " & column & ") %>%" & vbLf & _
        "  summarise(" & vbLf & _
        "    mean_value = mean(" & value & ", na.rm = TRUE)," & vbLf & _
        "    .groups = ""drop""" & vbLf & _
        "  ) %>%" & vbLf & _
        "  tidyr::complete(" & row & ", " & column & ") %>%" & vbLf & _
        "  tidyr::spread(key = " & column & ", value = mean_value)" & vbLf & vbLf & _
        "knitr::kable(" & vbLf & _
        "  wide_table," & vbLf & _
        "  digits = 2," & vbLf & _
        "  caption = """ & RString(caption) & """" & vbLf & ")"
End Function

Public Sub RenderWideTableCode()
    Dim dataset As String, row As String, column As String, value As String, caption As String, ready As Boolean, indent As String
    dataset = JsTrim(Q("#wideDatasetInput").Value)
    row = JsTrim(Q("#wideRowInput").Value)
    column = JsTrim(Q("#wideColumnInput").Value)
    value = JsTrim(Q("#wideValueInput").Value)
    caption = JsTrim(Q("#wideCaptionInput").Value)
    If Len(caption) = 0 Then caption = "Descriptive Statistics Table"
    ready = Len(dataset) > 0 And Len(row) > 0 And Len(column) > 0 And Len(value) > 0
    indent = vbLf & "        "
    Q("#wideMapping").InnerHtml = indent & "<span><strong>Rows</strong>" & EscapeHtml(IIf(Len(row) > 0, row, "column1")) & "</span>" & _
        indent & "<span><strong>Columns</strong>unique values of " & EscapeHtml(IIf(Len(column) > 0, column, "column2")) & "</span>" & _
        indent & "<span><strong>Cells</strong>mean of " & EscapeHtml(IIf(Len(value) > 0, value, "column3")) & "</span>"
    If ready Then
        Q("#wideTableCode").TextContent = BuildWideTableCode(dataset, row, column, value, caption)
    Else
        Q("#wideTableCode").TextContent = "Complete the dataset, row, column-header, and value fields to generate the wide-table code."
    End If
    Q("#copyWideCodeBtn").Disabled = Not ready
End Sub

Public Sub OnWideInput(ByVal e As clsEvent)
    On Error GoTo Failed
    RenderWideTableCode
    Exit Sub
Failed:
    ReportError "opStats.OnWideInput"
End Sub

Public Sub OnLoadWideSample(ByVal e As clsEvent)
    On Error GoTo Failed
    Q("#wideDatasetInput").Value = "dataset"
    Q("#wideRowInput").Value = "column1"
    Q("#wideColumnInput").Value = "column2"
    Q("#wideValueInput").Value = "column3"
    Q("#wideCaptionInput").Value = "Mean by Row and Column"
    RenderWideTableCode
    ShowToast "Sample loaded"
    Exit Sub
Failed:
    ReportError "opStats.OnLoadWideSample"
End Sub

Public Sub OnCopyWideCode(ByVal e As clsEvent)
    On Error GoTo Failed
    If Q("#copyWideCodeBtn").Disabled Then Exit Sub
    If WriteAnswerText(Q("#wideTableCode").TextContent) Then ShowToast "R code copied" Else ShowToast "Clipboard access is unavailable"
    Exit Sub
Failed:
    ReportError "opStats.OnCopyWideCode"
End Sub

' ---------------------------------------------------------------- skewness and kurtosis

Public Function SkewnessDescription(ByVal value As Double) As String
    Dim magnitude As Double, direction As String
    magnitude = Abs(value)
    If magnitude = 0 Then
        SkewnessDescription = "perfectly symmetrical"
        Exit Function
    End If
    If value < 0 Then direction = "left-skewed" Else direction = "right-skewed"
    If magnitude < 0.5 Then
        SkewnessDescription = "slightly " & direction
    ElseIf magnitude <= 1 Then
        SkewnessDescription = "moderately " & direction
    Else
        SkewnessDescription = "highly " & direction
    End If
End Function

Public Sub RenderShapeInterpretation()
    Dim skewText As String, kurtosisText As String, paragraphs As clsList, errors As clsList
    Dim skewness As Double, value As Double, baseline As Double, label As String, difference As Double
    Dim output As clsNode, prose As String, heading As String, noteClass As String
    skewText = JsTrim(Q("#skewnessInput").Value)
    kurtosisText = JsTrim(Q("#kurtosisInput").Value)
    Set paragraphs = NewList()
    Set errors = NewList()

    If Len(skewText) > 0 Then
        skewness = JsNumber(skewText)
        If JsIsFinite(skewness) Then
            paragraphs.Push "The coefficient of skewness is " & FormatFinalNumber(skewness) & ", indicating a " & SkewnessDescription(skewness) & " distribution."
        Else
            errors.Push "Enter a valid coefficient of skewness."
        End If
    End If

    If Len(kurtosisText) > 0 Then
        value = JsNumber(kurtosisText)
        If JsIsFinite(value) Then
            If mShapeMetric = "kurtosis" Then
                baseline = 3
                label = "kurtosis"
            Else
                baseline = 0
                label = "excess kurtosis"
            End If
            difference = value - baseline
            If Abs(difference) < 0.000000000001 Then
                paragraphs.Push "The " & label & " is " & FormatFinalNumber(value) & ", equal to the mesokurtic reference value of " & FormatFinalNumber(baseline) & _
                    ". The distribution is mesokurtic, with tail weight comparable to a normal distribution."
            ElseIf difference < 0 Then
                paragraphs.Push "The " & label & " is " & FormatFinalNumber(value) & ", which is below " & FormatFinalNumber(baseline) & _
                    ". The distribution is platykurtic, with lighter tails and fewer extreme values."
            Else
                paragraphs.Push "The " & label & " is " & FormatFinalNumber(value) & ", which is above " & FormatFinalNumber(baseline) & _
                    ". The distribution is leptokurtic, with heavier tails and more extreme values."
            End If
        Else
            errors.Push "Enter a valid " & IIf(mShapeMetric = "kurtosis", "kurtosis", "excess kurtosis") & " value."
        End If
    End If

    Set output = Q("#shapeOutput")
    If Len(skewText) = 0 And Len(kurtosisText) = 0 Then
        output.InnerHtml = "<div class=""empty-state compact""><div class=""empty-state-inner""><strong>Enter one or both coefficients</strong><p>The interpretation updates automatically.</p></div></div>"
        Q("#copyShapeProseBtn").Disabled = True
        Exit Sub
    End If
    prose = errors.Concat(paragraphs).Join(vbLf & vbLf)
    If errors.Count > 0 Then
        heading = "Check the entered values"
        noteClass = "waiting-note"
    Else
        heading = "Generated interpretation"
    End If
    output.InnerHtml = "<div class=""prose-card""><h3>" & heading & "</h3><p id=""shapeGeneratedProse"" class=""" & noteClass & """>" & EscapeHtml(prose) & "</p></div>"
    Q("#copyShapeProseBtn").Disabled = (paragraphs.Count = 0)
End Sub

Public Sub SetShapeMetric(ByVal metric As String)
    Dim buttons As clsList, i As Long, button As clsNode, active As Boolean, isExcess As Boolean
    mShapeMetric = metric
    Set buttons = QA(".metric-option")
    For i = 0 To buttons.Count - 1
        Set button = buttons(i)
        active = (JsStr(button.DataGet("metric")) = metric)
        button.ToggleClass "active", active
        button.SetAttribute "aria-pressed", IIf(active, "true", "false")
    Next i
    isExcess = (metric = "excess")
    Q("#kurtosisInputLabel").TextContent = IIf(isExcess, "Excess kurtosis value", "Kurtosis value")
    Q("#kurtosisInput").Placeholder = IIf(isExcess, "e.g. -0.600", "e.g. 2.40")
    If isExcess Then
        Q("#kurtosisGuideText").TextContent = "Excess kurtosis: 0 = mesokurtic; below 0 = platykurtic; above 0 = leptokurtic."
    Else
        Q("#kurtosisGuideText").TextContent = "Kurtosis: 3 = mesokurtic; below 3 = platykurtic; above 3 = leptokurtic."
    End If
    RenderShapeInterpretation
End Sub

Public Sub OnMetricOptionClick(ByVal e As clsEvent)
    On Error GoTo Failed
    SetShapeMetric JsStr(e.CurrentTarget.DataGet("metric"))
    Exit Sub
Failed:
    ReportError "opStats.OnMetricOptionClick"
End Sub

Public Sub OnShapeInput(ByVal e As clsEvent)
    On Error GoTo Failed
    RenderShapeInterpretation
    Exit Sub
Failed:
    ReportError "opStats.OnShapeInput"
End Sub

Public Sub OnLoadShapeSample(ByVal e As clsEvent)
    On Error GoTo Failed
    Q("#skewnessInput").Value = "0.75"
    Q("#kurtosisInput").Value = IIf(mShapeMetric = "kurtosis", "2.4", "-0.6")
    RenderShapeInterpretation
    ShowToast "Sample loaded"
    Exit Sub
Failed:
    ReportError "opStats.OnLoadShapeSample"
End Sub

Public Sub OnCopyShapeProse(ByVal e As clsEvent)
    Dim node As clsNode, prose As String
    On Error GoTo Failed
    Set node = Q("#shapeGeneratedProse")
    If node Is Nothing Then Exit Sub
    prose = node.TextContent
    If Len(prose) = 0 Then Exit Sub
    If WriteAnswerText(prose) Then ShowToast "Prose copied" Else ShowToast "Clipboard access is unavailable"
    Exit Sub
Failed:
    ReportError "opStats.OnCopyShapeProse"
End Sub
