Attribute VB_Name = "opRegression"
' Regression and classification operations (website app-src.js lines 2058-2540 and the
' event wiring at 4495-4515): one-variable regression code, the advanced regression
' form and its generated code. The interpretations (model comparison, multi-variable
' regression, GLM, coefficient intervals, confusion matrix) are in opRegression2.
' boundedTickCode/coordinate axis helpers and predictorList live in modShared.
Option Explicit

' Procedure the copy handlers call for the website's showRegressionUnitsAlert (modApp).
' Tests may point it at a stub.
Public RegressionUnitsAlertHandler As String

' ---------------------------------------------------------------- startup

' Resets state and performs the website's top-level statements for this range in order.
Public Sub InitOpRegression()
    Dim inputs As clsList, i As Long, node As clsNode
    RegressionUnitsAlertHandler = "ShowRegressionUnitsAlert"
    ResetRegressionInterpretationState
    Q("#loadAdvancedRegressionSampleBtn").AddEventListener "click", "opRegression.OnLoadAdvancedRegressionSample"
    Q("#copyAdvancedRegressionCodeBtn").AddEventListener "click", "opRegression.OnCopyAdvancedRegressionCode"
    Q("#copyAdvancedRegressionPredictionCodeBtn").AddEventListener "click", "opRegression.OnCopyAdvancedRegressionPredictionCode"
    Q("#copyAdvancedRegressionConfidenceCodeBtn").AddEventListener "click", "opRegression.OnCopyAdvancedRegressionConfidenceCode"
    Q("#glmOutput").AddEventListener "input", "opRegression2.OnGlmOutputInput"
    Q("#loadGlmOutputBtn").AddEventListener "click", "opRegression2.OnLoadGlmOutput"
    Q("#copyGlmProseBtn").AddEventListener "click", "opRegression2.OnCopyGlmProse"
    Q("#glmProse").AddEventListener "click", "opRegression2.OnGlmProseClick"
    Q("#modelComparisonOutput").AddEventListener "input", "opRegression2.OnModelComparisonOutputInput"
    Q("#loadModelComparisonOutputBtn").AddEventListener "click", "opRegression2.OnLoadModelComparisonOutput"
    Q("#copyModelComparisonProseBtn").AddEventListener "click", "opRegression2.OnCopyModelComparisonProse"
    Q("#multiRegressionOutput").AddEventListener "input", "opRegression2.OnMultiRegressionOutputInput"
    Q("#loadMultiRegressionOutputBtn").AddEventListener "click", "opRegression2.OnLoadMultiRegressionOutput"
    Q("#copyMultiRegressionProseBtn").AddEventListener "click", "opRegression2.OnCopyMultiRegressionProse"
    Q("#multiRegressionProse").AddEventListener "click", "opRegression2.OnMultiRegressionProseClick"
    Q("#confusionOutput").AddEventListener "input", "opRegression2.OnConfusionOutputInput"
    Q("#loadConfusionOutputBtn").AddEventListener "click", "opRegression2.OnLoadConfusionOutput"
    Q("#copyConfusionProseBtn").AddEventListener "click", "opRegression2.OnCopyConfusionProse"
    Q("#confusionProse").AddEventListener "click", "opRegression2.OnConfusionProseClick"
    Q("#coefficientIntervalOutput").AddEventListener "input", "opRegression2.OnCoefficientIntervalOutputInput"
    Q("#loadCoefficientIntervalBtn").AddEventListener "click", "opRegression2.OnLoadCoefficientInterval"
    Q("#copyCoefficientIntervalBtn").AddEventListener "click", "opRegression2.OnCopyCoefficientInterval"
    Q("#coefficientIntervalProse").AddEventListener "click", "opRegression2.OnCoefficientIntervalProseClick"
    Set inputs = QA(".regression-input")
    For i = 0 To inputs.Count - 1
        Set node = inputs(i)
        node.AddEventListener "input", "opRegression.OnRegressionInput"
    Next i
    Q("#regressionTransform").AddEventListener "change", "opRegression.OnRegressionInput"
    Q("#loadRegressionSampleBtn").AddEventListener "click", "opRegression.OnLoadRegressionSample"
    Q("#copyRegressionCodeBtn").AddEventListener "click", "opRegression.OnCopyRegressionCode"
End Sub

' Calls the shell's showRegressionUnitsAlert.
Public Sub ShowRegressionUnitsAlertFromRegression()
    If Len(RegressionUnitsAlertHandler) > 0 Then CallOp RegressionUnitsAlertHandler
End Sub

' ---------------------------------------------------------------- one-variable regression

Private Function TransformExpression(ByVal transform As String, ByVal xRef As String) As String
    Select Case transform
        Case "sqrt": TransformExpression = "sqrt(" & xRef & ")"
        Case "inverse": TransformExpression = "1 / " & xRef
        Case "log10": TransformExpression = "log10(" & xRef & ")"
        Case Else: TransformExpression = "log(" & xRef & ")"
    End Select
End Function

Public Function BuildBasicRegressionCode(ByVal datasetName As String, ByVal xName As String, ByVal yName As String, ByVal transform As String) As String
    Dim dataset As String, x As String, y As String, xRef As String, yRef As String
    Dim transformed As Boolean, expression As String, label As String, predictor As String, out As String
    Dim yLabel As String, xLabel As String
    dataset = RColumnName(datasetName): x = RColumnName(xName): y = RColumnName(yName)
    xRef = dataset & "$" & x: yRef = dataset & "$" & y
    transformed = (transform <> "none")
    expression = TransformExpression(transform, xRef)
    If Not transformed Then
        label = xName
    ElseIf transform = "inverse" Then
        label = "1/" & xName
    Else
        label = transform & "(" & xName & ")"
    End If
    If transformed Then predictor = "transformed_X" Else predictor = x
    yLabel = RString(ChartDisplayLabel(yName)): xLabel = RString(ChartDisplayLabel(label))
    If transformed Then out = dataset & "$transformed_X <- " & expression & vbLf & vbLf
    out = out & "plot(" & vbLf & _
        "  " & dataset & "$" & predictor & ", " & yRef & "," & vbLf & _
        "  main = ""Scatterplot of " & yLabel & " against " & xLabel & """," & vbLf & _
        "  xlab = """ & xLabel & """, ylab = """ & yLabel & """, col = ""blue""" & vbLf & _
        ")" & vbLf & vbLf & _
        "regression_data <- " & dataset & vbLf & _
        "regression_data$" & predictor & " <- as.numeric(" & dataset & "$" & predictor & ")" & vbLf & _
        "regression_data$" & y & " <- as.numeric(" & yRef & ")" & vbLf & _
        "linreg <- lm(" & y & " ~ " & predictor & ", data = regression_data)" & vbLf & _
        "summary(linreg)" & vbLf & _
        "summary(linreg)$r.squared" & vbLf & _
        "summary(linreg)$adj.r.squared" & vbLf & _
        "sigma(linreg)" & vbLf & vbLf & _
        "plot(linreg, 1)" & vbLf & _
        "plot(linreg, 2)"
    BuildBasicRegressionCode = out
End Function

Private Function RegressionAxisCode(ByVal predictor As String, ByVal yRef As String) As String
    RegressionAxisCode = CoordinateTickCode("x", predictor) & vbLf & CoordinateTickCode("y", yRef) & vbLf & CoordinatePlotMarginCode() & vbLf
End Function

Public Function BuildRegressionCode(ByVal datasetName As String, ByVal xName As String, ByVal yName As String, ByVal transform As String) As String
    Dim dataset As String, x As String, y As String, xRef As String, yRef As String
    Dim axisArguments As String, axes As String, label As String, expression As String
    Dim yLabel As String, xLabel As String
    If Not CalculatePlotAxisIntervals Then
        BuildRegressionCode = BuildBasicRegressionCode(datasetName, xName, yName, transform)
        Exit Function
    End If
    dataset = RColumnName(datasetName): x = RColumnName(xName): y = RColumnName(yName)
    xRef = dataset & "$" & x: yRef = dataset & "$" & y
    axisArguments = CoordinateAxisArguments()
    axes = ""
    yLabel = RString(ChartDisplayLabel(yName))
    If transform = "none" Then
        xLabel = RString(ChartDisplayLabel(xName))
        BuildRegressionCode = "# Scatterplot of Y against X" & vbLf & RegressionAxisCode(xRef, yRef) & _
            "plot(x_values, y_values, xlab = """ & xLabel & """, ylab = """ & yLabel & """, main = ""Scatterplot of " & yLabel & " against " & xLabel & """, " & axisArguments & ")" & vbLf & axes & vbLf & vbLf & _
            "# Fit the one-variable linear regression" & vbLf & _
            "regression_data <- " & dataset & vbLf & _
            "regression_data$" & x & " <- as.numeric(x_values)" & vbLf & _
            "regression_data$" & y & " <- as.numeric(y_values)" & vbLf & _
            "linreg <- lm(" & y & " ~ " & x & ", data = regression_data)" & vbLf & _
            "summary(linreg)" & vbLf & vbLf & _
            "# Model-fit measures" & vbLf & _
            "summary(linreg)$r.squared" & Space$(10) & "# proportion of variation explained" & vbLf & _
            "summary(linreg)$adj.r.squared" & Space$(6) & "# adjusted for the number of predictors" & vbLf & _
            "sigma(linreg)" & Space$(22) & "# residual standard error; smaller is better" & vbLf & vbLf & _
            "# Diagnostic plots" & vbLf & _
            "plot(linreg, 1) # residuals versus fitted" & vbLf & _
            "plot(linreg, 2) # normal Q-Q"
        Exit Function
    End If
    expression = TransformExpression(transform, xRef)
    Select Case transform
        Case "sqrt": label = "sqrt(" & xName & ")"
        Case "inverse": label = "1/" & xName
        Case "log10": label = "log10(" & xName & ")"
        Case Else: label = "log(" & xName & ")"
    End Select
    xLabel = RString(ChartDisplayLabel(label))
    BuildRegressionCode = "# Create the transformed predictor" & vbLf & _
        dataset & "$transformed_X <- " & expression & vbLf & vbLf & _
        "# Plot the transformed predictor against Y" & vbLf & RegressionAxisCode(dataset & "$transformed_X", yRef) & _
        "plot(x_values, y_values, xlab = """ & xLabel & """, ylab = """ & yLabel & """, main = ""Scatterplot of " & yLabel & " against " & xLabel & """, " & axisArguments & ")" & vbLf & axes & vbLf & vbLf & _
        "# Fit the transformed one-variable linear regression" & vbLf & _
        "regression_data <- " & dataset & vbLf & _
        "regression_data$transformed_X <- as.numeric(x_values)" & vbLf & _
        "regression_data$" & y & " <- as.numeric(y_values)" & vbLf & _
        "linreg <- lm(" & y & " ~ transformed_X, data = regression_data)" & vbLf & _
        "summary(linreg)" & vbLf & vbLf & _
        "# Model-fit measures" & vbLf & _
        "summary(linreg)$r.squared" & vbLf & _
        "summary(linreg)$adj.r.squared" & vbLf & _
        "sigma(linreg)" & vbLf & vbLf & _
        "# Diagnostic plots" & vbLf & _
        "plot(linreg, 1)" & vbLf & _
        "plot(linreg, 2)"
End Function

Public Sub RenderRegressionCode()
    Dim dataset As String, x As String, y As String, transform As String, ready As Boolean
    Dim xLabel As String, xShown As String, yShown As String
    dataset = JsTrim(Q("#regressionDataset").Value)
    If Len(dataset) = 0 Then dataset = "dataset"
    x = JsTrim(Q("#regressionX").Value)
    y = JsTrim(Q("#regressionY").Value)
    transform = Q("#regressionTransform").Value
    ready = (Len(x) > 0 And Len(y) > 0)
    If Len(x) > 0 Then xShown = x Else xShown = "column1"
    If Len(y) > 0 Then yShown = y Else yShown = "column2"
    If transform = "none" Then xLabel = xShown Else xLabel = transform & "(" & xShown & ")"
    Q("#regressionMapping").InnerHtml = "<span><strong>Model</strong>" & EscapeHtml(yShown) & " ~ " & EscapeHtml(xLabel) & _
        "</span><span><strong>Checks</strong>Residuals vs fitted; Normal Q" & ChrW$(&H2013) & "Q</span>"
    If ready Then
        Q("#regressionCode").TextContent = BuildRegressionCode(dataset, x, y, transform)
    Else
        Q("#regressionCode").TextContent = "Complete X and Y to generate the regression workflow."
    End If
    Q("#copyRegressionCodeBtn").Disabled = Not ready
End Sub

' ---------------------------------------------------------------- advanced regression code

Public Function AdvancedInput(ByVal label As String, ByVal id As String, ByVal value As String, Optional ByVal help As String = "") As String
    AdvancedInput = "<label>" & label & "<input class=""column-input advanced-regression-input"" id=""" & id & """ value=""" & EscapeHtml(value) & """>" & _
        IIf(Len(help) > 0, "<small>" & help & "</small>", "") & "</label>"
End Function

' $("#"+id)?.value.trim() || ""
Private Function AdvValue(ByVal id As String) As String
    Dim node As clsNode
    Set node = Q("#" & id)
    If Not node Is Nothing Then AdvValue = JsTrim(node.Value)
End Function

Private Function Fallback(ByVal value As String, ByVal defaultValue As String) As String
    If Len(value) > 0 Then Fallback = value Else Fallback = defaultValue
End Function

Private Function MapColumnNames(ByVal names As clsList) As clsList
    Dim out As clsList, i As Long
    Set out = NewList()
    For i = 0 To names.Count - 1
        out.Push RColumnName(CStr(names(i)))
    Next i
    Set MapColumnNames = out
End Function

Private Function NewDataLines(ByVal predictors As clsList) As String
    Dim parts As clsList, i As Long
    Set parts = NewList()
    For i = 0 To predictors.Count - 1
        parts.Push "  " & RColumnName(CStr(predictors(i))) & " = c(your_value)"
    Next i
    NewDataLines = parts.Join("," & vbLf)
End Function

Public Function AdvancedRegressionCode(ByVal mode As String) As String
    Dim dataset As String, y As String, predictors As clsList, terms As String, selected As clsList
    Dim restricted As String, full As String, categorical As String, levels As clsList, reference As String
    Dim other As clsList, rhs As String, quotedLevels As clsList, i As Long, x As String, moderator As String
    Dim response As String, positive As String, outcome As String, model As String, threshold As String
    dataset = RColumnName(Fallback(AdvValue("advDataset"), "dataset"))
    y = RColumnName(Fallback(AdvValue("advY"), "column1"))
    Set predictors = PredictorList(Fallback(AdvValue("advPredictors"), "column2, column3"))
    terms = MapColumnNames(predictors).Join(" + ")
    Select Case mode
    Case "multi-regression"
        Set selected = NewList(y)
        Set selected = selected.Concat(MapColumnNames(predictors))
        AdvancedRegressionCode = "# Select variables and fit the multivariable model" & vbLf & _
            "clean.data <- " & dataset & " %>% dplyr::select(" & selected.Join(", ") & ")" & vbLf & _
            "final_model <- lm(" & y & " ~ " & terms & ", data = clean.data)" & vbLf & _
            "summary(final_model)" & vbLf & _
            "summary(final_model)$r.squared" & vbLf & _
            "summary(final_model)$adj.r.squared" & vbLf & _
            "plot(final_model, 1, col = ""#3B4CC0"")" & vbLf & _
            "plot(final_model, 2, col = ""#3B4CC0"")"
    Case "model-significance"
        AdvancedRegressionCode = "fit <- lm(" & y & " ~ " & terms & ", data = " & dataset & ")" & vbLf & _
            "model_summary <- summary(fit)" & vbLf & _
            "model_summary" & vbLf & _
            "model_summary$fstatistic # use the F statistic and its p-value to assess overall significance"
    Case "model-comparison"
        restricted = MapColumnNames(PredictorList(Fallback(AdvValue("advRestricted"), "column2"))).Join(" + ")
        full = MapColumnNames(PredictorList(Fallback(AdvValue("advFull"), "column2, column3"))).Join(" + ")
        AdvancedRegressionCode = "m_restricted <- lm(" & y & " ~ " & restricted & ", data = " & dataset & ")" & vbLf & _
            "m_full <- lm(" & y & " ~ " & full & ", data = " & dataset & ")" & vbLf & _
            "anova(m_restricted, m_full)" & vbLf & _
            "# Pr(>F) < 0.05: reject H0; the full model is significantly better."
    Case "dummy"
        categorical = RColumnName(Fallback(AdvValue("advCategorical"), "column2"))
        Set levels = PredictorList(Fallback(AdvValue("advLevels"), "Young, Middle, Old"))
        reference = AdvValue("advReference")
        If Len(reference) = 0 Then
            ' levels[0] is undefined for an empty list, and rString(undefined) throws.
            If levels.Count = 0 Then Err.Raise vbObjectError + 513, "AdvancedRegressionCode", "Cannot read properties of undefined (reading 'replace')"
            reference = CStr(levels(0))
        End If
        Set other = MapColumnNames(PredictorList(Fallback(AdvValue("advOther"), "column3")))
        rhs = NewList(categorical).Concat(other).Join(" + ")
        Set quotedLevels = NewList()
        For i = 0 To levels.Count - 1
            quotedLevels.Push """" & RString(CStr(levels(i))) & """"
        Next i
        AdvancedRegressionCode = dataset & "$" & categorical & " <- factor(" & dataset & "$" & categorical & ", levels = c(" & quotedLevels.Join(", ") & "))" & vbLf & _
            dataset & "$" & categorical & " <- relevel(" & dataset & "$" & categorical & ", ref = """ & RString(reference) & """)" & vbLf & _
            "contrasts(" & dataset & "$" & categorical & ")" & vbLf & _
            "model <- lm(" & y & " ~ " & rhs & ", data = " & dataset & ")" & vbLf & _
            "summary(model)"
    Case "interaction"
        x = RColumnName(Fallback(AdvValue("advX"), "column2"))
        moderator = RColumnName(Fallback(AdvValue("advModerator"), "column3"))
        AdvancedRegressionCode = "interaction_model <- lm(" & y & " ~ " & moderator & " + " & x & " + " & moderator & " * " & x & ", data = " & dataset & ")" & vbLf & _
            "summary(interaction_model)" & vbLf & _
            "# The " & moderator & ":" & x & " coefficient shows how the effect of " & x & " changes for each one-unit increase in " & moderator & "."
    Case "glm"
        response = RColumnName(Fallback(AdvValue("advResponse"), "column1"))
        positive = Fallback(AdvValue("advPositive"), "Yes")
        AdvancedRegressionCode = dataset & "$" & response & " <- ifelse(" & dataset & "$" & response & " == """ & RString(positive) & """, 1, 0)" & vbLf & _
            dataset & "$" & response & " <- factor(" & dataset & "$" & response & ")" & vbLf & vbLf & _
            "logit_model <- glm(" & response & " ~ " & terms & ", data = " & dataset & ", family = ""binomial"")" & vbLf & _
            "summary(logit_model)" & vbLf & _
            "exp(coef(logit_model)) # odds ratios" & vbLf & vbLf & _
            "new_data <- data.frame(" & vbLf & NewDataLines(predictors) & vbLf & ")" & vbLf & _
            "predict(logit_model, newdata = new_data, type = ""response"")"
    Case Else
        outcome = RColumnName(Fallback(AdvValue("advOutcome"), "column1"))
        model = Fallback(AdvValue("advModel"), "logit_model")
        threshold = Fallback(AdvValue("advThreshold"), "0.30")
        AdvancedRegressionCode = "predicted_probs <- predict(" & model & ", type = ""response"")" & vbLf & _
            "predicted_status <- ifelse(predicted_probs >= " & threshold & ", 1, 0)" & vbLf & _
            "actual <- factor(" & dataset & "$" & outcome & ", levels = c(0, 1))" & vbLf & _
            "predicted <- factor(predicted_status, levels = c(0, 1))" & vbLf & _
            "caret::confusionMatrix(predicted, actual, positive = ""1"")"
    End Select
End Function

Public Function AdvancedRegressionPredictionCode() As String
    Dim predictors As clsList
    Set predictors = PredictorList(Fallback(AdvValue("advPredictors"), "column2, column3"))
    AdvancedRegressionPredictionCode = "# Predict a new observation" & vbLf & _
        "analysis_data <- data.frame(" & vbLf & NewDataLines(predictors) & vbLf & ")" & vbLf & _
        "predict(final_model, newdata = analysis_data, type = ""response"")" & vbLf & _
        "predict(final_model, newdata = analysis_data, interval = ""confidence"", level = 0.95)" & vbLf & _
        "predict(final_model, newdata = analysis_data, interval = ""prediction"", level = 0.95)"
End Function

Public Function AdvancedRegressionConfidenceCode() As String
    Dim predictors As clsList, parameter As String
    Set predictors = PredictorList(Fallback(AdvValue("advPredictors"), "column2, column3"))
    If predictors.Count > 0 Then parameter = RColumnName(CStr(predictors(0))) Else parameter = RColumnName("column2")
    AdvancedRegressionConfidenceCode = "# Confidence interval for a regression coefficient" & vbLf & _
        "confint(final_model, """ & RString(parameter) & """, level = 0.95)"
End Function

Public Sub RenderAdvancedRegressionCodeOutputs(ByVal mode As String)
    Dim isMultiRegression As Boolean
    isMultiRegression = (mode = "multi-regression")
    Q("#advancedRegressionCodeTitle").TextContent = IIf(isMultiRegression, "Model summary R code", "Generated R code")
    Q("#advancedRegressionCode").TextContent = AdvancedRegressionCode(mode)
    Q("#advancedRegressionPredictionPanel").ToggleClass "hidden", Not isMultiRegression
    Q("#advancedRegressionConfidencePanel").ToggleClass "hidden", Not isMultiRegression
    If isMultiRegression Then
        Q("#advancedRegressionPredictionCode").TextContent = AdvancedRegressionPredictionCode()
        Q("#advancedRegressionConfidenceCode").TextContent = AdvancedRegressionConfidenceCode()
    End If
End Sub

Public Sub RenderAdvancedRegressionForm(ByVal mode As String)
    Dim title As String, subtitle As String, fields As String, note As String, inputs As clsList, i As Long, node As clsNode
    Select Case mode
        Case "multi-regression": title = "Multi-variable linear regression": subtitle = "model summary, prediction, and confidence interval"
        Case "model-significance": title = "Statistical significance of model": subtitle = "overall F-test"
        Case "model-comparison": title = "Pairwise model comparison": subtitle = "restricted versus full"
        Case "dummy": title = "Dummy variables / reference groups": subtitle = "factor coding"
        Case "interaction": title = "Interaction terms": subtitle = "moderation"
        Case "glm": title = "GLM / binary variable": subtitle = "logistic regression"
        Case "confusion": title = "Confusion matrix": subtitle = "classification performance"
        Case Else: Err.Raise vbObjectError + 514, "RenderAdvancedRegressionForm", "Cannot read properties of undefined (reading '0')"
    End Select
    Q("#advancedRegressionTitle").TextContent = title
    Q("#advancedRegressionSubtitle").TextContent = subtitle
    fields = AdvancedInput("Dataset name", "advDataset", "dataset")
    If mode = "multi-regression" Or mode = "model-significance" Then
        fields = fields & AdvancedInput("Dependent variable (Y)", "advY", "column1") & AdvancedInput("Predictors, separated by commas", "advPredictors", "column2, column3")
        If mode = "multi-regression" Then
            note = "Each coefficient is the average change in Y for a one-unit increase in that predictor, holding all other predictors constant. A coefficient CI containing 0 is not significant at 5%."
        Else
            note = "Use the model F-test p-value: p < 0.05 indicates that the predictors jointly explain a statistically significant amount of variation in Y."
        End If
    ElseIf mode = "model-comparison" Then
        fields = AdvancedInput("Dataset name", "advDataset", "df1") & AdvancedInput("Dependent variable (Y)", "advY", "Salary") & _
            AdvancedInput("Restricted-model predictors", "advRestricted", "Years, CAtBat, CHits, CHmRun") & _
            AdvancedInput("Full-model predictors", "advFull", "Years, CAtBat, CHits, CHmRun, CRuns, CRBI, CWalks")
        note = SYM_H0() & ": The added predictors do not help explain Y. Pr(>F) < 0.05 means the full model is significantly better."
    ElseIf mode = "dummy" Then
        fields = fields & AdvancedInput("Dependent variable (Y)", "advY", "column1") & AdvancedInput("Categorical variable", "advCategorical", "column2") & _
            AdvancedInput("Levels, separated by commas", "advLevels", "Young, Middle, Old") & AdvancedInput("Reference level", "advReference", "Young") & _
            AdvancedInput("Other predictors", "advOther", "column3")
        note = "A categorical variable with n levels uses n" & SYM_MINUS() & "1 dummy coefficients. Each coefficient compares that level with the reference group."
    ElseIf mode = "interaction" Then
        fields = fields & AdvancedInput("Dependent variable (Y)", "advY", "column1") & AdvancedInput("Predictor (X)", "advX", "column2") & AdvancedInput("Moderator (M)", "advModerator", "column3")
        note = "A positive interaction strengthens the effect of X as M increases; a negative interaction weakens it. Keep the interaction in marginal-effect calculations even when its p-value exceeds 0.05."
    ElseIf mode = "glm" Then
        fields = fields & AdvancedInput("Binary response column", "advResponse", "column1") & AdvancedInput("Value representing 1", "advPositive", "Yes") & _
            AdvancedInput("Predictors, separated by commas", "advPredictors", "column2, column3")
        note = "Use GLM for binary outcomes. Exponentiate coefficients to obtain odds ratios, and check the reference group before interpreting categorical predictors."
    Else
        fields = fields & AdvancedInput("Outcome column (0/1)", "advOutcome", "column1") & AdvancedInput("Fitted model object", "advModel", "logit_model") & _
            AdvancedInput("Classification threshold", "advThreshold", "0.30")
        note = "Accuracy covers all classifications; precision focuses on predicted positives; sensitivity on actual positives; specificity on actual negatives."
    End If
    Q("#advancedRegressionForm").InnerHtml = "<div class=""interval-fields"">" & fields & "</div><div class=""shape-guide""><strong>Interpretation guide</strong><span>" & EscapeHtml(note) & "</span></div>"
    Set inputs = QA(".advanced-regression-input")
    For i = 0 To inputs.Count - 1
        Set node = inputs(i)
        node.AddEventListener "input", "opRegression.OnAdvancedRegressionInput"
    Next i
    RenderAdvancedRegressionCodeOutputs mode
    Q("#modelComparisonInterpretationPanel").ToggleClass "hidden", mode <> "model-comparison"
    Q("#multiRegressionInterpretationPanel").ToggleClass "hidden", mode <> "multi-regression"
    Q("#coefficientIntervalPanel").ToggleClass "hidden", mode <> "multi-regression"
    Q("#glmInterpretationPanel").ToggleClass "hidden", mode <> "glm"
    Q("#confusionInterpretationPanel").ToggleClass "hidden", mode <> "confusion"
    If mode = "multi-regression" Then
        RenderMultiRegressionInterpretation
        RenderCoefficientIntervals
    ElseIf mode = "glm" Then
        RenderGlmInterpretation
    ElseIf mode = "confusion" Then
        RenderConfusionInterpretation
    End If
End Sub

' ---------------------------------------------------------------- event handlers

Private Function ActiveId() As String
    ActiveId = CStr(AppState("id"))
End Function

Private Sub CopyText(ByVal text As String, ByVal message As String)
    If WriteAnswerText(text) Then ShowToast message Else ShowToast "Clipboard access is unavailable"
End Sub

Public Sub OnAdvancedRegressionInput(ByVal e As clsEvent)
    On Error GoTo Failed
    RenderAdvancedRegressionCodeOutputs ActiveId()
    If ActiveId() = "multi-regression" Then RenderCoefficientIntervals
    Exit Sub
Failed:
    ReportError "opRegression.OnAdvancedRegressionInput"
End Sub

Public Sub OnLoadAdvancedRegressionSample(ByVal e As clsEvent)
    On Error GoTo Failed
    RenderAdvancedRegressionForm ActiveId()
    ShowToast "Sample loaded"
    Exit Sub
Failed:
    ReportError "opRegression.OnLoadAdvancedRegressionSample"
End Sub

Public Sub OnCopyAdvancedRegressionCode(ByVal e As clsEvent)
    On Error GoTo Failed
    CopyText Q("#advancedRegressionCode").TextContent, "R code copied"
    Exit Sub
Failed:
    ReportError "opRegression.OnCopyAdvancedRegressionCode"
End Sub

Public Sub OnCopyAdvancedRegressionPredictionCode(ByVal e As clsEvent)
    On Error GoTo Failed
    CopyText Q("#advancedRegressionPredictionCode").TextContent, "Prediction R code copied"
    Exit Sub
Failed:
    ReportError "opRegression.OnCopyAdvancedRegressionPredictionCode"
End Sub

Public Sub OnCopyAdvancedRegressionConfidenceCode(ByVal e As clsEvent)
    On Error GoTo Failed
    CopyText Q("#advancedRegressionConfidenceCode").TextContent, "Confidence interval R code copied"
    Exit Sub
Failed:
    ReportError "opRegression.OnCopyAdvancedRegressionConfidenceCode"
End Sub

Public Sub OnRegressionInput(ByVal e As clsEvent)
    On Error GoTo Failed
    RenderRegressionCode
    Exit Sub
Failed:
    ReportError "opRegression.OnRegressionInput"
End Sub

Public Sub OnLoadRegressionSample(ByVal e As clsEvent)
    On Error GoTo Failed
    Q("#regressionDataset").Value = "dataset"
    Q("#regressionX").Value = "column1"
    Q("#regressionY").Value = "column2"
    Q("#regressionTransform").Value = "none"
    RenderRegressionCode
    ShowToast "Sample loaded"
    Exit Sub
Failed:
    ReportError "opRegression.OnLoadRegressionSample"
End Sub

Public Sub OnCopyRegressionCode(ByVal e As clsEvent)
    On Error GoTo Failed
    If Q("#copyRegressionCodeBtn").Disabled Then Exit Sub
    CopyText Q("#regressionCode").TextContent, "R code copied"
    Exit Sub
Failed:
    ReportError "opRegression.OnCopyRegressionCode"
End Sub
