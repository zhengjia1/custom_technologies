Attribute VB_Name = "opPlot2"
' Plotting & tables: the R code recipes of the website's renderGgplotCode (app-src.js
' lines 3198-3452), in both "Calculate axis intervals" modes. opPlot.RenderGgplotCode
' reads the form and calls PlotRecipeCode; this module only builds the code text.
Option Explicit

' The form values of the current render (renderGgplotCode's local constants).
Private mWorkflow As String, mDataset As String, mX As String, mYText As String, mY As String
Private mYs As clsList, mGroup As String, mStyle As String, mXT As String, mYT As String
Private mBinsText As String, mThreshold As Double, mTitle As String, mCaption As String
Private mXName As String, mYName As String, mGroupName As String

' ---------------------------------------------------------------- shared helpers

Public Function TransformExpression(ByVal expression As String, ByVal transform As String) As String
    Select Case transform
        Case "log": TransformExpression = "log(" & expression & ")"
        Case "log10": TransformExpression = "log10(" & expression & ")"
        Case "log2": TransformExpression = "log2(" & expression & ")"
        Case "sqrt": TransformExpression = "sqrt(" & expression & ")"
        Case "inverse": TransformExpression = "1 / (" & expression & ")"
        Case Else: TransformExpression = expression
    End Select
End Function

Public Function TransformLabel(ByVal name As String, ByVal transform As String) As String
    name = ChartDisplayLabel(name)
    Select Case transform
        Case "log": TransformLabel = "ln(" & name & ")"
        Case "log10": TransformLabel = "log10(" & name & ")"
        Case "log2": TransformLabel = "log2(" & name & ")"
        Case "sqrt": TransformLabel = "sqrt(" & name & ")"
        Case "inverse": TransformLabel = "1/" & name
        Case Else: TransformLabel = name
    End Select
End Function

' Joins its arguments with line feeds (a template literal's lines).
Private Function Lines(ParamArray parts() As Variant) As String
    Dim i As Long, out As String
    For i = LBound(parts) To UBound(parts)
        If i > LBound(parts) Then out = out & vbLf
        out = out & CStr(parts(i))
    Next i
    Lines = out
End Function

Private Function Quoted(ByVal value As String) As String
    Quoted = """" & RString(value) & """"
End Function

Private Function RefOf(ByVal name As String) As String
    RefOf = RColumnName(mDataset) & "$" & RColumnName(name)
End Function

Private Function DS() As String
    DS = RColumnName(mDataset)
End Function

Private Function OrientationSetup(ByVal rangeCode As String, ByVal categoryLabel As String, ByVal numericLabel As String, Optional ByVal isBoxplot As Boolean = False) As String
    Dim limits As String
    If Not isBoxplot Then limits = "axis_limits <- list(NULL, " & rangeCode & ")" & vbLf
    OrientationSetup = Lines("layout_index <- 1 + as.integer(use_horizontal)", _
        "left_margin <- min(16, max(8, max_label_chars * 0.6))", _
        limits & "axis_labels <- c(" & Quoted(ChartDisplayLabel(categoryLabel)) & ", " & Quoted(ChartDisplayLabel(numericLabel)) & ")", _
        "old_par <- par(mar = list(c(8, 4, 4, 2), c(5, left_margin, 4, 2))[[layout_index]] + 0.1)")
End Function

' bxp() swaps its data and category limits internally for horizontal boxplots.
Private Function OrientationArguments(Optional ByVal isBoxplot As Boolean = False) As String
    Dim limits As String
    If isBoxplot Then limits = "  ylim = value_range," Else limits = "  xlim = axis_limits[[layout_index]], ylim = axis_limits[[3 - layout_index]],"
    OrientationArguments = Lines("  xlab = axis_labels[layout_index], ylab = axis_labels[3 - layout_index],", limits, "  las = 3 - layout_index")
End Function

Private Function HistogramBins() As String
    If RxTest(mBinsText, "^\d+$") Then HistogramBins = mBinsText Else HistogramBins = Quoted(mBinsText)
End Function

' code.slice(0, code.indexOf(marker))
Private Function PrefixBefore(ByVal code As String, ByVal marker As String) As String
    PrefixBefore = JsSlice(code, 0, JsIndexOf(code, marker))
End Function

' ---------------------------------------------------------------- entry point

' Returns the generated code for the workflow and sets ready (renderGgplotCode's `code`
' and `ready`). The caller shows "Complete the required plotting fields." when not ready.
Public Function PlotRecipeCode(ByVal workflow As String, ByVal dataset As String, ByVal x As String, ByVal yText As String, _
        ByVal group As String, ByVal style As String, ByVal xTransform As String, ByVal yTransform As String, _
        ByVal binsText As String, ByVal threshold As Double, ByVal title As String, ByVal tableCaption As String, _
        ByRef ready As Boolean) As String
    Dim code As String
    mWorkflow = workflow: mDataset = dataset: mX = x: mYText = yText
    Set mYs = PredictorList(yText)
    mGroup = group: mStyle = style: mXT = xTransform: mYT = yTransform
    mBinsText = binsText: mThreshold = threshold: mTitle = title: mCaption = tableCaption
    ready = (Len(x) > 0)
    If x = "" Then mXName = RColumnName("column1") Else mXName = RColumnName(x)
    If mYs.Count > 0 Then mY = CStr(mYs(0)) Else mY = ""
    If mY = "" Then mYName = RColumnName("column2") Else mYName = RColumnName(mY)
    If group = "" Then mGroupName = RColumnName("column3") Else mGroupName = RColumnName(group)

    Select Case workflow
        Case "summary"
            ready = ready And mYs.Count > 0
            code = SummaryCode()
        Case "grouped-summary"
            ready = ready And Len(group) > 0
            code = GroupedSummaryCode()
        Case "frequency"
            code = FrequencyCode()
        Case "pie"
            code = PieCode()
        Case "contingency"
            ready = ready And Len(group) > 0
            code = ContingencyCode()
        Case "grouped-means"
            ready = ready And Len(mY) > 0 And Len(group) > 0
            code = GroupedMeansCode()
        Case "histogram"
            code = HistogramCode()
        Case "boxplot"
            code = BoxplotCode()
        Case "grouped-boxplot"
            ready = ready And Len(mY) > 0
            code = GroupedBoxplotCode()
        Case "scatter", "line"
            ready = ready And Len(mY) > 0
        Case "pareto"
            If ready Then ready = JsIsFinite(threshold)
            If ready Then ready = (threshold >= 1 And threshold <= 100)
            code = ParetoCode()
    End Select

    If ready Then code = AxisAdjustedCode(code)
    If ready And Not CalculatePlotAxisIntervals Then code = BasicPlotCode(code)
    PlotRecipeCode = code
End Function

' ---------------------------------------------------------------- recipes (axis-interval layout)

Private Function SummaryCode() As String
    Dim summaries As String, i As Long, name As String
    For i = 0 To mYs.Count - 1
        name = CStr(mYs(i))
        If i > 0 Then summaries = summaries & "," & vbLf
        summaries = summaries & "    mean_" & Replace(RColumnName(name), "`", "") & " = mean(" & RColumnName(name) & ", na.rm = TRUE)"
    Next i
    SummaryCode = Lines("library(dplyr)", "library(tidyr)", "library(ggplot2)", "", _
        "result <- " & DS() & " %>%", _
        "  group_by(" & mXName & ") %>%", _
        "  summarise(", _
        summaries & ",", _
        "    .groups = ""drop""", _
        "  )", "", _
        "result_long <- result %>%", _
        "  pivot_longer(", _
        "    cols = starts_with(""mean_""),", _
        "    names_to = ""Variable"",", _
        "    values_to = ""Mean""", _
        "  )", "", _
        "p <- ggplot(result_long, aes(x = " & mXName & ", y = " & TransformExpression("Mean", mYT) & ", fill = Variable)) +", _
        "  geom_bar(stat = ""identity"", position = ""dodge"") +", _
        "  labs(title = " & Quoted(mTitle) & ", x = " & Quoted(ChartDisplayLabel(mX)) & ", y = " & Quoted(TransformLabel("Mean", mYT)) & ") +", _
        "  theme_minimal() +", _
        "  theme(axis.text.x = element_text(angle = 45, hjust = 1))", "", _
        "p")
End Function

Private Function GroupedSummaryCode() As String
    Dim head As String
    head = Lines("library(dplyr)", "library(knitr)", "", _
        "summary_table <- " & DS() & " %>%", _
        "  group_by(" & mGroupName & ") %>%", _
        "  summarise(", _
        "    Mean = mean(" & mXName & ", na.rm = TRUE),", _
        "    SD = sd(" & mXName & ", na.rm = TRUE),", _
        "    .groups = ""drop""", _
        "  )", "", _
        "knitr::kable(", _
        "  summary_table, row.names = FALSE, digits = 2,", _
        "  caption = " & Quoted(mCaption), _
        ")", "")
    GroupedSummaryCode = Lines(head, _
        "bar_values <- as.matrix(summary_table[, c(""Mean"", ""SD"")])", _
        "group_labels <- as.character(summary_table[[" & Quoted(mGroup) & "]])", _
        "bar_colours <- rainbow(nrow(bar_values))", _
        BoundedTickCode("value", "c(0, bar_values)", True), _
        "old_par <- par(mar = c(6, 4, 4, 2) + 0.1)", "", _
        "bar_positions <- barplot(", _
        "  bar_values, beside = TRUE, col = bar_colours,", _
        "  names.arg = c(""Mean"", ""Standard deviation""),", _
        "  main = " & Quoted(mTitle) & ", xlab = ""Statistic"", ylab = " & Quoted(ChartDisplayLabel(mX)) & ",", _
        "  " & NumericAxisArguments("value", "y", "grDevices::extendrange(c(0, bar_values), f = 0.15)"), _
        "  las = 1, cex.names = 0.8", _
        ")", _
        NumericAxisCode(), _
        "text(", _
        "  x = bar_positions, y = bar_values,", _
        "  labels = format(round(bar_values, 2), trim = TRUE),", _
        "  pos = 3, cex = 0.7", _
        ")", _
        "legend(", _
        "  ""topright"", legend = group_labels, fill = bar_colours,", _
        "  title = " & Quoted(ChartDisplayLabel(mGroup)) & ", cex = 0.8", _
        ")", _
        "par(old_par)")
End Function

Private Function FrequencyCode() As String
    FrequencyCode = Lines("library(dplyr)", "library(knitr)", "", _
        "FreqTable <- " & DS() & " %>% dplyr::count(" & mXName & ", sort = TRUE)", _
        "kable(FreqTable, caption = " & Quoted(mCaption) & ", col.names = c(" & Quoted(mX) & ", ""Frequency""))", "", _
        "category_labels <- as.character(FreqTable$" & mXName & ")", _
        "max_label_chars <- max(nchar(category_labels), na.rm = TRUE)", _
        "left_margin <- min(18, max(8, max_label_chars * 0.7))", _
        BoundedTickCode("frequency", "c(0, FreqTable$n)", True, True), _
        "old_par <- par(mar = c(5, left_margin, 4, 2) + 0.1)", "", _
        "bp <- barplot(", _
        "  FreqTable$n,", _
        "  names.arg = category_labels,", _
        "  main = " & Quoted(mTitle) & ",", _
        "  xlab = ""Frequency"", ylab = " & Quoted(ChartDisplayLabel(mX)) & ",", _
        "  col = rainbow(nrow(FreqTable)),", _
        "  " & NumericAxisArguments("frequency", "x", "grDevices::extendrange(c(0, FreqTable$n), f = c(0, 0.15))"), _
        "  horiz = TRUE, las = 1, cex.names = 0.8", _
        ")", _
        NumericAxisCode(), _
        "text(x = FreqTable$n, y = bp, labels = FreqTable$n, pos = 4, offset = 0.3, cex = 0.7)", _
        "par(old_par)")
End Function

Private Function PieCode() As String
    PieCode = Lines("library(dplyr)", "library(glue)", "", _
        "FreqTable <- " & DS() & " %>% dplyr::count(" & mXName & ", sort = TRUE)", _
        "x.values <- FreqTable$n", _
        "x.perc <- round(100 * x.values / sum(x.values), 1)", _
        "percent.label <- glue::glue(""{FreqTable$" & mXName & "}, {x.perc}%"")", _
        "label_cex <- 0.8", "", _
        "pie(", _
        "  x.values, labels = percent.label,", _
        "  main = " & Quoted(mTitle) & ", radius = 0.75, cex = label_cex,", _
        "  col = rainbow(length(x.values))", _
        ")")
End Function

Private Function ContingencyCode() As String
    Dim chart As String, totals As String, rangeCode As String, beside As String
    If mStyle = "mosaic" Then
        chart = Lines("old_par <- par(mar = c(6, 6, 4, 2) + 0.1)", _
            "mosaicplot(", _
            "  ContTable, main = " & Quoted(mTitle) & ", xlab = " & Quoted(ChartDisplayLabel(mX)) & ", ylab = " & Quoted(ChartDisplayLabel(mGroup)) & ",", _
            "  col = rainbow(nrow(ContTable)), las = 1, cex.axis = 0.8", _
            ")", _
            "par(old_par)")
    Else
        If mStyle = "grouped" Then totals = "ContTable": beside = "TRUE" Else totals = "colSums(ContTable)": beside = "FALSE"
        If CalculatePlotAxisIntervals Then rangeCode = "frequency_range" Else rangeCode = "grDevices::extendrange(c(0, " & totals & "), f = c(0, 0.15))"
        chart = Lines(BoundedTickCode("frequency", "c(0, " & totals & ")", True, True), _
            "category_labels <- colnames(ContTable)", _
            "max_label_chars <- max(nchar(category_labels), na.rm = TRUE)", _
            "use_horizontal <- length(category_labels) > 6 || max_label_chars > 12", _
            OrientationSetup(rangeCode, mGroup, "Frequency"), "", _
            "barplot(", _
            "  ContTable, beside = " & beside & ", horiz = use_horizontal,", _
            "  names.arg = category_labels, main = " & Quoted(mTitle) & ",", _
            "  col = rainbow(nrow(ContTable)), legend = TRUE,", _
            OrientationArguments() & ", cex.names = 0.8", _
            ")", _
            NumericAxisCode(), _
            "par(old_par)")
    End If
    ContingencyCode = Lines("library(knitr)", "", _
        "ContTable <- table(" & RefOf(mX) & ", " & RefOf(mGroup) & ")", _
        "ContTable", _
        "kable(ContTable, caption = " & Quoted(mCaption) & ")", "", _
        chart)
End Function

Private Function GroupedMeansCode() As String
    Dim rangeCode As String
    If CalculatePlotAxisIntervals Then rangeCode = "value_range" Else rangeCode = "grDevices::extendrange(c(0, plot_values), f = 0.15)"
    GroupedMeansCode = Lines("library(dplyr)", "library(knitr)", "", _
        "AvgTable_long <- " & DS() & " %>%", _
        "  group_by(" & mXName & ", " & mGroupName & ") %>%", _
        "  summarise(avg_" & mYName & " = mean(" & mYName & ", na.rm = TRUE), .groups = ""drop"")", "", _
        "kable(AvgTable_long, caption = " & Quoted(mCaption) & ")", _
        "AvgTable <- xtabs(avg_" & mYName & " ~ " & mXName & " + " & mGroupName & ", data = AvgTable_long)", _
        "plot_values <- " & TransformExpression("AvgTable", mYT), _
        BoundedTickCode("value", "c(0, plot_values)", True), _
        "bar_cols <- rainbow(nrow(AvgTable))", _
        "category_labels <- colnames(AvgTable)", _
        "max_label_chars <- max(nchar(category_labels), na.rm = TRUE)", _
        "use_horizontal <- length(category_labels) > 6 || max_label_chars > 12", _
        OrientationSetup(rangeCode, mGroup, TransformLabel("Average " & mY, mYT)), "", _
        "barplot(", _
        "  plot_values, beside = TRUE, horiz = use_horizontal, col = bar_cols,", _
        "  names.arg = category_labels, main = " & Quoted(mTitle) & ",", _
        OrientationArguments() & ", cex.names = 0.8", _
        ")", _
        NumericAxisCode(), _
        "legend(""topright"", legend = rownames(AvgTable), fill = bar_cols, title = " & Quoted(ChartDisplayLabel(mX)) & ", cex = 0.7)", _
        "par(old_par)")
End Function

Private Function HistogramCode() As String
    Dim values As String, limits As String
    values = TransformExpression(RefOf(mX), mXT)
    If CalculatePlotAxisIntervals Then limits = "xlim = x_range, ylim = y_range" Else limits = "ylim = grDevices::extendrange(c(0, h$counts), f = c(0, 0.15))"
    HistogramCode = Lines("library(knitr)", "", _
        "h <- hist(", _
        "  " & values & ", breaks = " & HistogramBins() & ", plot = FALSE", _
        ")", _
        BoundedTickCode("x", "h$breaks"), _
        BoundedTickCode("y", "c(0, h$counts)", True, True), _
        "plot(", _
        "  h, main = " & Quoted(mTitle) & ",", _
        "  xlab = " & Quoted(TransformLabel(mX, mXT)) & ", ylab = ""Frequency"",", _
        "  col = ""lightgreen"", labels = FALSE, freq = TRUE,", _
        "  " & limits, _
        ")", _
        NumericAxisCode(), _
        NumericAxisCode(), _
        "text(x = h$mids, y = h$counts, pos = 3, labels = h$counts, cex = 0.7)", "", _
        "Groups <- cut(" & values & ", breaks = h$breaks, include.lowest = TRUE, dig.lab = 5)", _
        "FreqTable <- table(Groups)", _
        "kable(FreqTable, caption = " & Quoted(mCaption) & ", col.names = c(" & Quoted(TransformLabel(mX, mXT) & " interval") & ", ""Frequency""))")
End Function

Private Function BoxplotCode() As String
    Dim values As String, label As String
    values = TransformExpression(RefOf(mX), mXT)
    label = Quoted(TransformLabel(mX, mXT))
    BoxplotCode = Lines("output.1.5 <- boxplot(", _
        "  " & values & ", range = 1.5, horizontal = TRUE, col = ""orange"",", _
        "  xlab = " & label & ",", _
        "  main = " & Quoted(mTitle & " (standard outliers)"), _
        ")", _
        "output.1.5$out", "", _
        "output.3 <- boxplot(", _
        "  " & values & ", range = 3, horizontal = TRUE, col = ""orange"",", _
        "  xlab = " & label & ",", _
        "  main = " & Quoted(mTitle & " (extreme outliers)"), _
        ")", _
        "output.3$out", "", _
        "# Mild-only outliers: between 1.5 and 3 IQR", _
        "setdiff(output.1.5$out, output.3$out)")
End Function

Private Function GroupedBoxplotCode() As String
    GroupedBoxplotCode = Lines("group_labels <- levels(factor(" & RefOf(mX) & "))", _
        "max_label_chars <- max(nchar(group_labels), na.rm = TRUE)", _
        "use_horizontal <- length(group_labels) > 6 || max_label_chars > 12", _
        "numeric_values <- " & TransformExpression(RefOf(mY), mYT), _
        "means <- tapply(numeric_values, " & RefOf(mX) & ", mean, na.rm = TRUE)", _
        BoundedTickCode("value", "numeric_values"), _
        OrientationSetup("value_range", mX, TransformLabel(mY, mYT), True), "", _
        "boxplot(", _
        "  " & TransformExpression(mYName, mYT) & " ~ " & mXName & ", data = " & DS() & ", horizontal = use_horizontal,", _
        "  main = " & Quoted(mTitle) & ", col = ""lightblue"",", _
        OrientationArguments(True) & ", cex.axis = 0.8", _
        ")", _
        NumericAxisCode(), _
        "mean_coordinates <- list(seq_along(means), means)", _
        "points(mean_coordinates[[layout_index]], mean_coordinates[[3 - layout_index]], col = ""red"", pch = 19)", _
        "legend(""topright"", legend = ""Group mean"", col = ""red"", pch = 19, cex = 0.8)", _
        "par(old_par)")
End Function

Private Function ParetoCode() As String
    Dim target As String
    If JsIsFinite(mThreshold) Then target = JsNumberToString(mThreshold) Else target = "80"
    ParetoCode = Lines("library(dplyr)", "", _
        "VALUE_COL <- " & Quoted(mX), _
        "TARGET_PERCENT <- " & target, _
        "original_n <- nrow(" & DS() & ")", "", _
        "pareto.df <- " & DS() & " %>%", _
        "  transmute(Value = .data[[VALUE_COL]]) %>%", _
        "  filter(!is.na(Value))", "", _
        "total_value <- sum(pareto.df$Value)", "", _
        "pareto.df <- pareto.df %>%", _
        "  arrange(desc(Value)) %>%", _
        "  mutate(CumulativePercent = 100 * cumsum(Value) / total_value)", "", _
        "number_contributors <- which(pareto.df$CumulativePercent + 1e-10 >= TARGET_PERCENT)[1]", _
        "percentage_of_dataset <- 100 * number_contributors / original_n", _
        "actual_contribution <- pareto.df$CumulativePercent[number_contributors]", "", _
        "cat(", _
        "  number_contributors, "" observations ("",", _
        "  format(round(percentage_of_dataset, 2), nsmall = 2),", _
        "  ""% of the dataset) account for "",", _
        "  format(round(actual_contribution, 2), nsmall = 2),", _
        "  ""% of total "", VALUE_COL, "".\n"", sep = """"", _
        ")")
End Function

' ---------------------------------------------------------------- ready-state adjustments

' The website's post-processing for summary, boxplot, scatter and line once ready.
Private Function AxisAdjustedCode(ByVal code As String) As String
    Dim meanRange As String, values As String, scaleLine As String, xValues As String, yValues As String
    Dim lineType As String, axisArgs As String, replacement As String
    Select Case mWorkflow
        Case "summary"
            meanRange = TransformExpression("result_long$Mean", mYT)
            code = Replace(code, vbLf & "p <- ggplot", vbLf & BoundedTickCode("y", "c(0, " & meanRange & ")", True) & vbLf & "p <- ggplot", 1, 1)
            If CalculatePlotAxisIntervals Then scaleLine = "  scale_y_continuous(breaks = y_ticks, limits = y_range, expand = expansion(mult = 0)) +" & vbLf
            code = Replace(code, "  geom_bar(stat = ""identity"", position = ""dodge"") +", _
                "  geom_bar(stat = ""identity"", position = ""dodge"") +" & vbLf & scaleLine & "  scale_fill_brewer(palette = ""Set2"") +", 1, 1)
        Case "boxplot"
            values = TransformExpression(RefOf(mX), mXT)
            code = "plot_values <- " & values & vbLf & BoundedTickCode("value", "plot_values") & vbLf & vbLf & Replace(code, values, "plot_values")
            If CalculatePlotAxisIntervals Then
                replacement = "$1," & vbLf & "  ylim = value_range" & vbLf & ")"
            Else
                replacement = "$1" & vbLf & ")"
            End If
            code = Rx("(main = [^\n]+)\n\)", "g").Replace(code, replacement)
        Case "scatter", "line"
            xValues = TransformExpression(RefOf(mX), mXT)
            yValues = TransformExpression(RefOf(mY), mYT)
            If mWorkflow = "line" Then lineType = " type = ""l"","
            If CalculatePlotAxisIntervals Then axisArgs = "," & vbLf & "  " & CoordinateAxisArguments()
            code = Lines(CoordinateTickCode("x", xValues), CoordinateTickCode("y", yValues), "", _
                "plot(x_values, y_values," & lineType, _
                "  main = " & Quoted(mTitle) & ", xlab = " & Quoted(TransformLabel(mX, mXT)) & ",", _
                "  ylab = " & Quoted(TransformLabel(mY, mYT)) & ", col = ""blue""" & axisArgs & ")")
    End Select
    AxisAdjustedCode = code
End Function

' ---------------------------------------------------------------- basic mode

' Basic mode follows the supplied plotting reference. Keep analysis/table calculations,
' but use ordinary plot defaults instead of custom layout.
Private Function BasicPlotCode(ByVal code As String) As String
    Select Case mWorkflow
        Case "summary"
            code = PrefixBefore(code, "p <- ggplot") & Lines( _
                "ggplot(result_long, aes(x = " & mXName & ", y = " & TransformExpression("Mean", mYT) & ", fill = Variable)) +", _
                "  geom_bar(stat = ""identity"", position = ""dodge"") +", _
                "  labs(title = " & Quoted(mTitle) & ", x = " & Quoted(ChartDisplayLabel(mX)) & ", y = " & Quoted(TransformLabel("Mean", mYT)) & ")")
        Case "grouped-summary"
            code = PrefixBefore(code, "bar_values <-") & Lines( _
                "bar_values <- as.matrix(summary_table[, c(""Mean"", ""SD"")])", _
                "bar_colours <- rainbow(nrow(bar_values))", _
                "bar_positions <- barplot(", _
                "  bar_values, beside = TRUE, col = bar_colours,", _
                "  names.arg = c(""Mean"", ""Standard deviation""),", _
                "  main = " & Quoted(mTitle) & ", xlab = ""Statistic"", ylab = " & Quoted(ChartDisplayLabel(mX)) & ",", _
                "  ylim = c(min(0, bar_values), max(0, bar_values) + 1)", _
                ")", _
                "text(x = bar_positions, y = bar_values, labels = round(bar_values, 2), pos = 3, cex = 0.7)", _
                "legend(""topright"", legend = summary_table[[" & Quoted(mGroup) & "]], fill = bar_colours, title = " & Quoted(ChartDisplayLabel(mGroup)) & ", cex = 0.8)")
        Case "frequency"
            code = Lines("library(dplyr)", "library(knitr)", "", _
                "FreqTable <- " & DS() & " %>% dplyr::count(" & mXName & ")", _
                "kable(FreqTable, caption = " & Quoted(mCaption) & ", col.names = c(" & Quoted(mX) & ", ""Frequency""))", "", _
                "barplot(FreqTable$n, names.arg = FreqTable$" & mXName & ",", _
                "  main = " & Quoted(mTitle) & ", xlab = ""Frequency"", ylab = " & Quoted(ChartDisplayLabel(mX)) & ",", _
                "  col = ""blue"", horiz = TRUE, las = 1, cex.names = 0.8)")
        Case "pie"
            code = Lines("library(dplyr)", "", _
                "FreqTable <- " & DS() & " %>% dplyr::count(" & mXName & ")", _
                "x.perc <- round(100 * FreqTable$n / sum(FreqTable$n), 1)", _
                "percent.label <- paste(FreqTable$" & mXName & ", "", "", x.perc, ""%"", sep = """")", _
                "pie(FreqTable$n, labels = percent.label, main = " & Quoted(mTitle) & ",", _
                "  radius = 0.8, cex = 0.8, col = RColorBrewer::brewer.pal(8, ""Set2""))")
        Case "contingency"
            code = BasicContingencyCode()
        Case "grouped-means"
            code = PrefixBefore(code, "plot_values <-") & Lines( _
                "plot_values <- " & TransformExpression("AvgTable", mYT), _
                "bar_cols <- rainbow(nrow(AvgTable))", _
                "barplot(plot_values, beside = TRUE, col = bar_cols,", _
                "  main = " & Quoted(mTitle) & ", xlab = " & Quoted(ChartDisplayLabel(mGroup)) & ", ylab = " & Quoted(TransformLabel("Average " & mY, mYT)) & ",", _
                "  ylim = c(min(0, plot_values), max(0, plot_values) + 1))", _
                "legend(""topright"", legend = rownames(AvgTable), fill = bar_cols, title = " & Quoted(ChartDisplayLabel(mX)) & ", cex = 0.7)")
        Case "histogram"
            code = BasicHistogramCode()
        Case "grouped-boxplot"
            code = Lines("boxplot(" & TransformExpression(mYName, mYT) & " ~ " & mXName & ", data = " & DS() & ",", _
                "  main = " & Quoted(mTitle) & ", xlab = " & Quoted(ChartDisplayLabel(mX)) & ", ylab = " & Quoted(TransformLabel(mY, mYT)) & ", col = ""lightblue"")", _
                "means <- tapply(" & TransformExpression(RefOf(mY), mYT) & ", " & RefOf(mX) & ", mean, na.rm = TRUE)", _
                "points(seq_along(means), means, col = ""red"", pch = 19)", _
                "legend(""topright"", legend = ""Group mean"", col = ""red"", pch = 19, cex = 0.8)")
        Case "scatter", "line"
            Dim lineType As String
            If mWorkflow = "line" Then lineType = " type = ""l"","
            code = Lines("plot(" & TransformExpression(RefOf(mX), mXT) & ", " & TransformExpression(RefOf(mY), mYT) & "," & lineType, _
                "  main = " & Quoted(mTitle) & ", xlab = " & Quoted(TransformLabel(mX, mXT)) & ",", _
                "  ylab = " & Quoted(TransformLabel(mY, mYT)) & ", col = ""blue"")")
    End Select
    BasicPlotCode = code
End Function

Private Function BasicContingencyCode() As String
    Dim beside As String
    If mStyle = "mosaic" Then
        BasicContingencyCode = Lines("library(knitr)", "", _
            "ContTable <- table(" & RefOf(mX) & ", " & RefOf(mGroup) & ")", _
            "ContTable", _
            "kable(ContTable, caption = " & Quoted(mCaption) & ")", _
            "mosaicplot(ContTable, main = " & Quoted(mTitle) & ", xlab = " & Quoted(ChartDisplayLabel(mX)) & ", ylab = " & Quoted(ChartDisplayLabel(mGroup)) & ", col = RColorBrewer::brewer.pal(8, ""Set2""))")
        Exit Function
    End If
    If mStyle = "grouped" Then beside = "TRUE" Else beside = "FALSE"
    BasicContingencyCode = Lines("library(dplyr)", "library(tidyr)", "library(knitr)", "", _
        "ContLong <- " & DS() & " %>% dplyr::count(" & mXName & ", " & mGroupName & ")", _
        "ContWide <- ContLong %>% spread(key = " & mGroupName & ", value = n, fill = 0)", _
        "kable(ContWide, caption = " & Quoted(mCaption) & ")", _
        "ContTable <- as.matrix(ContWide[-1])", _
        "bar_colours <- RColorBrewer::brewer.pal(8, ""Set2"")", _
        "barplot(ContTable, beside = " & beside & ",", _
        "  names.arg = names(ContWide)[-1], col = bar_colours,", _
        "  main = " & Quoted(mTitle) & ", xlab = " & Quoted(ChartDisplayLabel(mGroup)) & ", ylab = ""Frequency"",", _
        "  las = 1, cex.names = 0.8)", _
        "legend(""topright"", legend = ContWide$" & mXName & ", fill = bar_colours,", _
        "  title = " & Quoted(ChartDisplayLabel(mX)) & ", cex = 0.6)")
End Function

Private Function BasicHistogramCode() As String
    Dim values As String
    values = TransformExpression(RefOf(mX), mXT)
    BasicHistogramCode = Lines("library(knitr)", "", _
        "h <- hist(" & values & ", breaks = " & HistogramBins() & ",", _
        "  main = " & Quoted(mTitle) & ", xlab = " & Quoted(TransformLabel(mX, mXT)) & ",", _
        "  ylab = ""Frequency"", col = ""lightgreen"", labels = TRUE)", "", _
        "Groups <- cut(" & values & ", breaks = h$breaks, include.lowest = TRUE, dig.lab = 5)", _
        "FreqTable <- table(Groups)", _
        "kable(FreqTable, caption = " & Quoted(mCaption) & ", col.names = c(" & Quoted(TransformLabel(mX, mXT) & " interval") & ", ""Frequency""))")
End Function
