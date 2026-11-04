Attribute VB_Name = "opRegression2"
' Regression interpretations (website app-src.js lines 2153-2525): pairwise model
' comparison, multi-variable regression output, binary GLM, coefficient confidence
' intervals and the confusion matrix, with their copy handlers.
'
' Parse results are clsMaps: "ok" (Boolean) and "message" on failure; on success the
' website's object fields. JS null becomes Null (numbers) or Nothing (objects).
Option Explicit

Private Const NUMBER_PATTERN As String = "[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?"

Public ModelComparisonProse As String
Public MultiRegressionProseSections As clsList   ' of clsMap {predictor, text}
Public MultiRegressionModelSections As clsList   ' of clsMap {title, text}
Public GlmProseSections As clsList               ' of clsMap {predictor, text}
Public CoefficientIntervalSections As clsList    ' of clsMap {parameter, text}
Public ConfusionProseSections As clsList         ' of clsMap {title, text}

Public Sub ResetRegressionInterpretationState()
    ModelComparisonProse = ""
    Set MultiRegressionProseSections = NewList()
    Set MultiRegressionModelSections = NewList()
    Set GlmProseSections = NewList()
    Set CoefficientIntervalSections = NewList()
    Set ConfusionProseSections = NewList()
End Sub

' ---------------------------------------------------------------- samples

Private Function JoinLines(ParamArray lines() As Variant) As String
    Dim i As Long, out As String
    For i = LBound(lines) To UBound(lines)
        If i > LBound(lines) Then out = out & vbLf
        out = out & lines(i)
    Next i
    JoinLines = out
End Function

Public Function ModelComparisonSample() As String
    ModelComparisonSample = JoinLines("Analysis of Variance Table", "", _
        "Model 1: Salary ~ Years + CAtBat + CHits + CHmRun", _
        "Model 2: Salary ~ Years + CAtBat + CHits + CHmRun + CRuns + CRBI + CWalks", _
        "  Res.Df      RSS Df Sum of Sq      F Pr(>F)", _
        "1    258 31893223", _
        "2    255 31469243  3    423981 1.1452 0.3314")
End Function

Public Function MultiRegressionSample() As String
    Dim lq As String, rq As String
    lq = ChrW$(&H2018): rq = ChrW$(&H2019)
    MultiRegressionSample = JoinLines("Call:", _
        "lm(formula = sleep ~ age + educ + income + totwrk, data = sleepdata)", "", _
        "Residuals:", _
        "     Min       1Q   Median       3Q      Max", _
        "-2407.90  -234.26    15.11   248.27  1302.72", "", _
        "Coefficients:", _
        "              Estimate Std. Error t value Pr(>|t|)", _
        "(Intercept)  3.624e+03  1.139e+02  31.813   <2e-16 ***", _
        "age          2.377e+00  1.465e+00   1.623    0.105", _
        "educ        -9.742e+00  6.162e+00  -1.581    0.114", _
        "income      -1.368e-03  1.791e-03  -0.764    0.445", _
        "totwrk      -1.468e-01  1.682e-02  -8.731   <2e-16 ***", _
        "---", _
        "Signif. codes:  0 " & lq & "***" & rq & " 0.001 " & lq & "**" & rq & " 0.01 " & lq & "*" & rq & " 0.05 " & lq & "." & rq & " 0.1 " & lq & " " & rq & " 1", "", _
        "Residual standard error: 419.5 on 701 degrees of freedom", _
        "Multiple R-squared:  0.1141, Adjusted R-squared:  0.109", _
        "F-statistic: 22.57 on 4 and 701 DF, p-value: < 2.2e-16")
End Function

Public Function GlmSampleOutput() As String
    GlmSampleOutput = JoinLines("## ", "## Call:", _
        "## glm(formula = HIV_Status ~ HAART_A + HAART_B + VLoad + Age, family = ""binomial"", ", _
        "##     data = dataset)", "## ", "## Coefficients:", _
        "##               Estimate Std. Error z value Pr(>|z|)    ", _
        "## (Intercept) -3.091e+00  6.639e-01  -4.655 3.24e-06 ***", _
        "## HAART_A     -1.179e+00  3.302e-01  -3.570 0.000357 ***", _
        "## HAART_B     -2.427e+00  4.435e-01  -5.471 4.48e-08 ***", _
        "## VLoad        4.827e-04  5.098e-05   9.469  < 2e-16 ***", _
        "## Age         -8.136e-03  1.819e-02  -0.447 0.654743    ", _
        "## ---", _
        "## Signif. codes:  0 '***' 0.001 '**' 0.01 '*' 0.05 '.' 0.1 ' ' 1", "## ", _
        "## (Dispersion parameter for binomial family taken to be 1)", "## ", _
        "##     Null deviance: 501.28  on 489  degrees of freedom", _
        "## Residual deviance: 314.90  on 485  degrees of freedom", _
        "## AIC: 324.9", "## ", _
        "## Number of Fisher Scoring iterations: 6")
End Function

Public Function CoefficientIntervalSample() As String
    CoefficientIntervalSample = JoinLines("            2.5 %     97.5 %", "totwrk -0.179863 -0.1138225")
End Function

Public Function ConfusionSampleOutput() As String
    ConfusionSampleOutput = JoinLines("Confusion Matrix and Statistics", "", _
        "          Reference", _
        "Prediction  No Yes", _
        "       No   42   6", _
        "       Yes   8  44", "", _
        "Sensitivity : 0.880", _
        "Specificity : 0.840", _
        "Accuracy : 0.860")
End Function

' ---------------------------------------------------------------- small helpers

Private Function Fail(ByVal message As String) As clsMap
    Dim result As clsMap
    Set result = NewMap()
    result("ok") = False
    result("message") = message
    Set Fail = result
End Function

Private Function Section(ByVal keyName As String, ByVal keyValue As String, ByVal text As String) As clsMap
    Dim result As clsMap
    Set result = NewMap()
    result(keyName) = keyValue
    result("text") = text
    Set Section = result
End Function

Private Function StripBackticks(ByVal text As String) As String
    StripBackticks = RxReplace(text, "^`|`$", "", "g")
End Function

Private Function SplitLines(ByVal text As String) As clsList
    Set SplitLines = JsSplit(text, vbLf)
End Function

Private Function JoinSectionTexts(ByVal sections As clsList) As String
    Dim parts As clsList, i As Long, item As clsMap
    Set parts = NewList()
    For i = 0 To sections.Count - 1
        Set item = sections(i)
        parts.Push item("text")
    Next i
    JoinSectionTexts = parts.Join(vbLf & vbLf)
End Function

Private Sub CopyText(ByVal text As String, ByVal message As String, Optional ByVal unitsAlert As Boolean = False)
    If WriteAnswerText(text) Then
        ShowToast message
        If unitsAlert Then ShowRegressionUnitsAlertFromRegression
    Else
        ShowToast "Clipboard access is unavailable"
    End If
End Sub

' Index from a data-* attribute value (Number(value)); -1 when it is not a valid index.
Private Function SectionIndex(ByVal raw As Variant, ByVal sections As clsList) As Long
    Dim value As Double
    SectionIndex = -1
    value = JsToNumber(raw)
    If Not JsIsFinite(value) Then Exit Function
    If Not JsIsInteger(value) Then Exit Function
    If value < 0 Or value >= sections.Count Then Exit Function
    SectionIndex = CLng(value)
End Function

Private Function ProseCard(ByVal heading As String, ByVal attrName As String, ByVal index As Long, ByVal text As String, Optional ByVal headingAttributes As String = "") As String
    ProseCard = "<article class=""sensitivity-prose-card""><div class=""sensitivity-prose-head""><h3" & headingAttributes & ">" & heading & _
        "</h3><button class=""sample-btn"" type=""button"" " & attrName & "=""" & index & """>Copy</button></div><p>" & EscapeHtml(text) & "</p></article>"
End Function

' ---------------------------------------------------------------- model comparison

Public Function FormatModelComparisonNumber(ByVal value As Double) As String
    Dim absolute As Double
    absolute = Abs(value)
    If absolute = 0 Then
        FormatModelComparisonNumber = "0"
    ElseIf absolute >= 1 Then
        FormatModelComparisonNumber = JsToFixed(value, 2)
    Else
        FormatModelComparisonNumber = JsToPrecision(value, 3)
    End If
End Function

Public Function ExtractModelComparisonFormula(ByVal output As String, ByVal modelNumber As Long) As String
    Dim lines As clsList, start As Long, i As Long, formula As String, continuation As String, startRx As clsRegex
    Set lines = SplitLines(output)
    Set startRx = Rx("^\s*Model\s*" & modelNumber & "\s*:", "i")
    start = -1
    For i = 0 To lines.Count - 1
        If startRx.Test(CStr(lines(i))) Then start = i: Exit For
    Next i
    If start < 0 Then Exit Function
    formula = JsTrim(Rx("^\s*Model\s*" & modelNumber & "\s*:\s*", "i").Replace(CStr(lines(start)), ""))
    For i = start + 1 To lines.Count - 1
        continuation = JsTrim(CStr(lines(i)))
        If Len(continuation) = 0 Then Exit For
        If RxTest(continuation, "^Model\s*\d+\s*:", "i") Or RxTest(continuation, "^(?:Res\.Df|Df\b|RSS\b|Sum of Sq\b|Signif\.|---)", "i") Or RxTest(continuation, "^\d+\s+[-+]?\d") Then Exit For
        If Not JsEndsWith(formula, "+") And Not JsStartsWith(continuation, "+") Then Exit For
        formula = formula & " " & continuation
    Next i
    ExtractModelComparisonFormula = JsTrim(RxReplace(formula, "\s+", " ", "g"))
End Function

Private Sub SetModelComparisonMessage(ByVal message As String, Optional ByVal isError As Boolean = False)
    Dim prose As clsNode
    ModelComparisonProse = ""
    Set prose = Q("#modelComparisonProse")
    prose.TextContent = message
    prose.ToggleClass "is-error", isError
    Q("#copyModelComparisonProseBtn").Disabled = True
End Sub

Private Function ParseFormulaParts(ByVal formula As String) As clsMap
    Dim parts As clsList, result As clsMap, predictors As clsList, items As clsList, i As Long, item As String, rhs As String
    Set parts = JsSplit(formula, "~")
    Set result = NewMap()
    result("outcome") = StripBackticks(JsTrim(CStr(parts(0))))
    If parts.Count > 1 Then rhs = CStr(parts(1))
    Set items = JsSplit(rhs, "+")
    Set predictors = NewList()
    For i = 0 To items.Count - 1
        item = StripBackticks(JsTrim(CStr(items(i))))
        If Len(item) > 0 Then predictors.Push item
    Next i
    Set result("predictors") = predictors
    Set ParseFormulaParts = result
End Function

Public Sub RenderModelComparisonInterpretation()
    Dim output As String, model1 As String, model2 As String, restricted As clsMap, full As clsMap
    Dim restrictedPredictors As clsList, fullPredictors As clsList, excluded As clsList, i As Long, name As String
    Dim lines As clsList, headerIndex As Long, matches As Collection, m As clsMatch, values As clsList, allFinite As Boolean
    Dim rowLine As String, rowValues As clsList, found As Boolean, fValue As Double, pValue As Double, pOperator As String
    Dim pm As clsMatch, significant As Variant, relation As String, variableLabel As String, improvementPhrase As String
    Dim betaTerms As clsList, nullHypothesis As String, alternativeHypothesis As String, prose As clsNode, number As Double
    output = Replace(Q("#modelComparisonOutput").Value, vbCr, "")
    If Len(JsTrim(output)) = 0 Then SetModelComparisonMessage "Paste the pairwise ANOVA output to compare the nested models.": Exit Sub
    model1 = ExtractModelComparisonFormula(output, 1)
    model2 = ExtractModelComparisonFormula(output, 2)
    If Len(model1) = 0 Or Len(model2) = 0 Then SetModelComparisonMessage "I could not find both the Model 1 and Model 2 formulas.", True: Exit Sub
    Set restricted = ParseFormulaParts(model1)
    Set full = ParseFormulaParts(model2)
    Set restrictedPredictors = restricted("predictors")
    Set fullPredictors = full("predictors")
    Set excluded = NewList()
    For i = 0 To fullPredictors.Count - 1
        name = CStr(fullPredictors(i))
        If Not restrictedPredictors.Includes(name) Then excluded.Push name
    Next i
    If Len(restricted("outcome")) = 0 Or excluded.Count = 0 Then SetModelComparisonMessage "I could not identify the dependent variable and predictors excluded from Model 1.", True: Exit Sub
    Set lines = SplitLines(output)
    headerIndex = -1
    For i = 0 To lines.Count - 1
        If RxTest(CStr(lines(i)), "Pr\s*\(\s*>\s*F\s*\)", "i") Then headerIndex = i: Exit For
    Next i
    If headerIndex < 0 Then SetModelComparisonMessage "I could not find the F-statistic and Pr(>F) columns.", True: Exit Sub
    For i = headerIndex + 1 To lines.Count - 1
        Set matches = Rx("[-+]?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?", "gi").MatchAll(CStr(lines(i)))
        Set values = NewList()
        allFinite = True
        For Each m In matches
            number = JsNumber(m.Value)
            If Not JsIsFinite(number) Then allFinite = False
            values.Push number
        Next m
        If values.Count >= 7 And allFinite Then
            found = True
            rowLine = CStr(lines(i))
            Set rowValues = values
        End If
    Next i
    If Not found Then SetModelComparisonMessage "I found the ANOVA headings but could not read the comparison row.", True: Exit Sub
    fValue = rowValues(rowValues.Count - 2)
    pValue = rowValues(rowValues.Count - 1)
    Set pm = RxMatch(rowLine, "(<=|>=|<|>)\s*[-+]?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?\s*[* .]*$", "i")
    pOperator = "="
    If Not pm Is Nothing Then
        If Len(pm.Group(1)) > 0 Then pOperator = pm.Group(1)
    End If
    significant = PValueDecision(pOperator, pValue)
    relation = PValueComparison(pOperator, pValue)
    If excluded.Count = 1 Then variableLabel = "variable" Else variableLabel = "variables"
    If IsNull(significant) Then SetModelComparisonMessage PValueDecisionError(pOperator, pValue), True: Exit Sub
    If significant Then
        If excluded.Count = 1 Then improvementPhrase = "significantly improves" Else improvementPhrase = "significantly improve"
    Else
        If excluded.Count = 1 Then improvementPhrase = "does not significantly improve" Else improvementPhrase = "do not significantly improve"
    End If
    Set betaTerms = NewList()
    For i = 0 To excluded.Count - 1
        betaTerms.Push SYM_BETA() & excluded(i)
    Next i
    nullHypothesis = SYM_H0() & ": " & betaTerms.Join(" = ") & " = 0"
    If excluded.Count = 1 Then
        alternativeHypothesis = SYM_H1() & ": " & betaTerms(0) & " is not equal to 0"
    Else
        alternativeHypothesis = SYM_H1() & ": At least one of " & betaTerms.Join(", ") & " is not equal to 0"
    End If
    ModelComparisonProse = nullHypothesis & vbLf & alternativeHypothesis & vbLf & vbLf & _
        "F-statistic = " & FormatModelComparisonNumber(fValue) & ", P-value " & pOperator & " " & FormatPValue(pValue) & ", which is " & relation & _
        " 0.05, hence we " & IIf(significant, "have", "do not have") & " sufficient evidence to reject the null hypothesis at 5% level of significance." & vbLf & vbLf & _
        "The excluded " & variableLabel & " (" & excluded.Join(", ") & ") " & improvementPhrase & " the model's ability to explain " & restricted("outcome") & "."
    Set prose = Q("#modelComparisonProse")
    prose.RemoveClass "is-error"
    prose.TextContent = ModelComparisonProse
    Q("#copyModelComparisonProseBtn").Disabled = False
End Sub

' ---------------------------------------------------------------- shared regression wording

Public Function NormaliseScientificToken(ByVal raw As String) As String
    Dim token As String, m As clsMatch, exponent As Double, sign As String
    token = JsTrim(raw)
    Set m = RxMatch(token, "^([-+]?(?:\d+\.?\d*|\.\d+))[eE]([-+]?)(\d+)$")
    If m Is Nothing Then NormaliseScientificToken = token: Exit Function
    sign = m.Group(2)
    If Len(sign) = 0 Then sign = "+"
    exponent = JsNumber(sign & m.Group(3))
    If exponent = 0 Then
        NormaliseScientificToken = JsNumberToString(JsNumber(m.Group(1)))
    Else
        NormaliseScientificToken = m.Group(1) & "e" & IIf(exponent >= 0, "+", "") & JsNumberToString(exponent)
    End If
End Function

Public Function JoinRegressionNames(ByVal names As clsList) As String
    If names.Count = 0 Then
        JoinRegressionNames = "all other predictors"
    ElseIf names.Count = 1 Then
        JoinRegressionNames = names(0)
    Else
        JoinRegressionNames = names.Slice(0, names.Count - 1).Join(", ") & ", and " & names(names.Count - 1)
    End If
End Function

Public Function RegressionDisplayName(ByVal name As String) As String
    Dim cleaned As String
    cleaned = StripBackticks(JsTrim(name))
    Select Case JsLower(cleaned)
        Case "sleep": RegressionDisplayName = "weekly sleep duration"
        Case "educ": RegressionDisplayName = "education"
        Case "totwrk": RegressionDisplayName = "total minutes worked"
        Case Else
            RegressionDisplayName = JsLower(RxReplace(RxReplace(cleaned, "[._]+", " ", "g"), "([a-z])([A-Z])", "$1 $2", "g"))
    End Select
End Function

' ---------------------------------------------------------------- GLM

Public Function ParseGlmOutput(ByVal text As String) As clsMap
    Dim output As String, callMatch As clsMatch, outcome As String, lines As clsList, coefficientIndex As Long, headerIndex As Long
    Dim i As Long, line As String, rowRx As clsRegex, m As clsMatch, name As String, estimate As Double, pValue As Double
    Dim coefficients As clsList, c As clsMap, result As clsMap, pOperator As String
    output = Replace(RxReplace(text, "^\s{0,3}#{1,3}\s{0,3}", "", "gm"), vbCr, "")
    Set callMatch = RxMatch(output, "glm\s*\(\s*(?:formula\s*=\s*)?([^~,(]+?)\s*~\s*([\s\S]*?)(?=,\s*(?:family|data)\s*=)", "i")
    If callMatch Is Nothing Then Set ParseGlmOutput = Fail("I could not find the glm() formula in the Call section."): Exit Function
    If Not RxTest(output, "family\s*=\s*[""']?binomial", "i") Then Set ParseGlmOutput = Fail("This interpreter requires a binomial GLM summary."): Exit Function
    outcome = StripBackticks(JsTrim(callMatch.Group(1)))
    Set lines = SplitLines(output)
    coefficientIndex = -1: headerIndex = -1
    For i = 0 To lines.Count - 1
        If RxTest(CStr(lines(i)), "^\s*Coefficients\s*:\s*$", "i") Then coefficientIndex = i: Exit For
    Next i
    For i = 0 To lines.Count - 1
        line = lines(i)
        If i > coefficientIndex And RxTest(line, "Estimate", "i") And RxTest(line, "Std\.\s*Error", "i") And RxTest(line, "z\s*value", "i") And RxTest(line, "Pr\s*\(\s*>\s*\|z\|\s*\)", "i") Then headerIndex = i: Exit For
    Next i
    If coefficientIndex < 0 Or headerIndex < 0 Then Set ParseGlmOutput = Fail("I could not find the Coefficients table with Estimate, z value, and Pr(>|z|)."): Exit Function
    Set rowRx = Rx("^\s*(.+?)\s+(" & NUMBER_PATTERN & ")\s+(" & NUMBER_PATTERN & ")\s+(" & NUMBER_PATTERN & ")\s+(<=|>=|<|>)?\s*(" & NUMBER_PATTERN & ")(?:\s+\S+)?\s*$", "i")
    Set coefficients = NewList()
    For i = headerIndex + 1 To lines.Count - 1
        line = lines(i)
        If RxTest(line, "^\s*(?:---|Signif\.|Dispersion|Null deviance|Residual deviance|AIC|Number of Fisher)", "i") Then Exit For
        If Len(JsTrim(line)) = 0 Then
            If coefficients.Count > 0 Then Exit For
            GoTo NextLine
        End If
        Set m = rowRx.Exec(line)
        If m Is Nothing Then GoTo NextLine
        name = StripBackticks(JsTrim(m.Group(1)))
        If name = "(Intercept)" Then GoTo NextLine
        estimate = JsNumber(m.Group(2)): pValue = JsNumber(m.Group(6))
        If Not JsIsFinite(estimate) Or Not JsIsFinite(pValue) Then GoTo NextLine
        pOperator = m.Group(5)
        If Len(pOperator) = 0 Then pOperator = "="
        Set c = NewMap()
        c("name") = name: c("estimateRaw") = m.Group(2): c("estimate") = estimate
        c("zValue") = JsNumber(m.Group(4)): c("pOperator") = pOperator: c("pRaw") = m.Group(6): c("pValue") = pValue
        coefficients.Push c
NextLine:
    Next i
    For i = 0 To coefficients.Count - 1
        Set c = coefficients(i)
        If IsNull(PValueDecision(c("pOperator"), c("pValue"))) Then Set ParseGlmOutput = Fail("A predictor p-value does not establish significance at 0.05. Paste a more precise p-value."): Exit Function
    Next i
    If coefficients.Count = 0 Then Set ParseGlmOutput = Fail("I found the GLM headings but could not read any predictor rows."): Exit Function
    Set result = NewMap()
    result("ok") = True
    result("outcome") = outcome
    Set result("coefficients") = coefficients
    Set ParseGlmOutput = result
End Function

Public Function GlmDisplayOutcome(ByVal name As String) As String
    Dim withoutStatus As String
    withoutStatus = Rx("(?:[._\s]?status)$", "i").Replace(name, "")
    If Len(withoutStatus) = 0 Then withoutStatus = name
    GlmDisplayOutcome = RxReplace(RxReplace(withoutStatus, "[._]+", " ", "g"), "([a-z])([A-Z])", "$1 $2", "g")
End Function

Public Function FormatGlmEstimate(ByVal coefficient As clsMap) As String
    Dim absolute As Double
    absolute = Abs(CDbl(coefficient("estimate")))
    If absolute >= 0.001 And absolute < 1000000# Then
        FormatGlmEstimate = JsNumberToString(coefficient("estimate"))
    Else
        FormatGlmEstimate = NormaliseScientificToken(coefficient("estimateRaw"))
    End If
End Function

Public Function BuildGlmProseSections(ByVal parsed As clsMap) As clsList
    Dim outcome As String, coefficients As clsList, out As clsList, i As Long, c As clsMap, name As String
    Dim oddsRatio As Double, change As Double, increase As Boolean, significant As Variant, pDisplay As String, pClause As String
    Dim subject As String, association As String, changeText As String, significance As String, est As String, text As String
    outcome = GlmDisplayOutcome(parsed("outcome"))
    Set coefficients = parsed("coefficients")
    Set out = NewList()
    For i = 0 To coefficients.Count - 1
        Set c = coefficients(i)
        name = c("name")
        oddsRatio = JsExp(c("estimate"))
        change = JsSub(oddsRatio, 1)
        change = JsAbs(change)
        change = JsMul(change, 100)
        increase = oddsRatio > 1
        significant = PValueDecision(c("pOperator"), c("pValue"))
        pDisplay = NormaliseScientificToken(c("pRaw"))
        If c("pOperator") = "=" Then pClause = pDisplay Else pClause = c("pOperator") & " " & pDisplay
        If RxTest(name, "^HAART(?:[_\s]|$)", "i") Then subject = name & " regimen" Else subject = "a one-unit increase in " & name
        association = FormatFinalNumber(oddsRatio) & " times the odds of " & outcome
        If oddsRatio = 1 Then
            changeText = "no change in the odds"
        Else
            changeText = FormatFinalNumber(change) & "% " & IIf(increase, "increase", "decrease") & " in the odds"
        End If
        If significant Then significance = "statistically significant" Else significance = "not statistically significant"
        est = FormatGlmEstimate(c)
        text = "The coefficient for " & name & " is " & est & ". The odds ratio is exp(" & est & ") = " & FormatFinalNumber(oddsRatio) & _
            ", meaning " & subject & " is associated with " & association & ", holding other variables constant. This corresponds to " & _
            IIf(oddsRatio = 1, "", "a ") & changeText & ". Since the p-value is " & pClause & ", which is " & PValueComparison(c("pOperator"), c("pValue")) & _
            " 0.05, this effect is " & significance & " at the 5% level."
        out.Push Section("predictor", name, text)
    Next i
    Set BuildGlmProseSections = out
End Function

Public Sub RenderGlmInterpretation()
    Dim inputText As String, container As clsNode, copyAll As clsNode, parsed As clsMap, html As String, i As Long, s As clsMap
    inputText = Q("#glmOutput").Value
    Set container = Q("#glmProse")
    Set copyAll = Q("#copyGlmProseBtn")
    Set parsed = ParseGlmOutput(inputText)
    Set GlmProseSections = NewList()
    If Len(JsTrim(inputText)) = 0 Then
        container.InnerHtml = "<div class=""pca-variance-prose"">Paste a binomial GLM summary to interpret every predictor as an odds ratio.</div>"
        copyAll.Disabled = True
        Exit Sub
    End If
    If Not parsed("ok") Then
        container.InnerHtml = "<div class=""pca-variance-prose is-error"">" & EscapeHtml(parsed("message")) & "</div>"
        copyAll.Disabled = True
        Exit Sub
    End If
    Set GlmProseSections = BuildGlmProseSections(parsed)
    For i = 0 To GlmProseSections.Count - 1
        Set s = GlmProseSections(i)
        html = html & ProseCard(EscapeHtml(s("predictor")), "data-copy-glm", i, s("text"))
    Next i
    container.InnerHtml = html
    copyAll.Disabled = False
End Sub

' ---------------------------------------------------------------- multi-variable regression

' Also used by the universal detector (modApp): returns {ok, message} or
' {ok, outcome, coefficients (clsList of clsMap), rSquared (Double or Null), fTest (clsMap or Nothing)}.
Public Function ParseMultiRegressionOutput(ByVal text As String) As clsMap
    Dim output As String, callMatch As clsMatch, formula As String, tilde As Long, outcome As String
    Dim lines As clsList, coefficientIndex As Long, headerIndex As Long, i As Long, line As String, rowRx As clsRegex, m As clsMatch
    Dim coefficients As clsList, c As clsMap, name As String, pOperator As String, rMatch As clsMatch, fMatch As clsMatch
    Dim rSquared As Variant, fTest As clsMap, result As clsMap, invalid As Boolean
    output = Replace(text, vbCr, "")
    Set callMatch = RxMatch(output, "lm\s*\(\s*(?:formula\s*=\s*)?([\s\S]*?)\s*,\s*data\s*=", "i")
    If callMatch Is Nothing Then Set ParseMultiRegressionOutput = Fail("I could not find the lm() formula and dependent variable in the Call section."): Exit Function
    formula = JsTrim(RxReplace(callMatch.Group(1), "\s+", " ", "g"))
    tilde = JsIndexOf(formula, "~")
    If tilde >= 0 Then outcome = StripBackticks(JsTrim(JsSlice(formula, 0, tilde)))
    If Len(outcome) = 0 Then Set ParseMultiRegressionOutput = Fail("I found lm(), but could not identify the dependent variable before ~."): Exit Function
    Set lines = SplitLines(output)
    coefficientIndex = -1: headerIndex = -1
    For i = 0 To lines.Count - 1
        If RxTest(CStr(lines(i)), "^\s*Coefficients\s*:\s*$", "i") Then coefficientIndex = i: Exit For
    Next i
    For i = 0 To lines.Count - 1
        line = lines(i)
        If i > coefficientIndex And RxTest(line, "Estimate", "i") And RxTest(line, "Std\.\s*Error", "i") And RxTest(line, "t\s*value", "i") And RxTest(line, "Pr\s*\(\s*>\s*\|t\|\s*\)", "i") Then headerIndex = i: Exit For
    Next i
    If coefficientIndex < 0 Or headerIndex < 0 Then Set ParseMultiRegressionOutput = Fail("I could not find the Coefficients table with Estimate, t value, and Pr(>|t|)."): Exit Function
    Set rowRx = Rx("^\s*(.+?)\s+(" & NUMBER_PATTERN & ")\s+(" & NUMBER_PATTERN & ")\s+(" & NUMBER_PATTERN & ")\s+(<=|>=|<|>)?\s*(" & NUMBER_PATTERN & ")(?:\s+\S+)?\s*$", "i")
    Set coefficients = NewList()
    For i = headerIndex + 1 To lines.Count - 1
        line = lines(i)
        If RxTest(line, "^\s*(?:---|Signif\.|Residual standard error|Multiple R-squared|F-statistic)", "i") Then Exit For
        If Len(JsTrim(line)) = 0 Then
            If coefficients.Count > 0 Then Exit For
            GoTo NextLine
        End If
        Set m = rowRx.Exec(line)
        If m Is Nothing Then GoTo NextLine
        name = StripBackticks(JsTrim(m.Group(1)))
        If name = "(Intercept)" Then GoTo NextLine
        pOperator = m.Group(5)
        If Len(pOperator) = 0 Then pOperator = "="
        Set c = NewMap()
        c("name") = name: c("estimateRaw") = m.Group(2): c("estimate") = JsNumber(m.Group(2))
        c("tRaw") = m.Group(4): c("tValue") = JsNumber(m.Group(4))
        c("pOperator") = pOperator: c("pRaw") = m.Group(6): c("pValue") = JsNumber(m.Group(6))
        coefficients.Push c
NextLine:
    Next i
    If coefficients.Count = 0 Then Set ParseMultiRegressionOutput = Fail("I found the coefficient headings but could not read any predictor rows. Include the complete table from summary(lm(...))."): Exit Function
    Set rMatch = RxMatch(output, "Multiple\s+R-squared\s*:\s*(" & NUMBER_PATTERN & ")", "i")
    Set fMatch = RxMatch(output, "F-statistic\s*:\s*(" & NUMBER_PATTERN & ")\s+on\s+(" & NUMBER_PATTERN & ")\s+and\s+(" & NUMBER_PATTERN & ")\s+DF\s*,?\s*p-value\s*:\s*(<=|>=|<|>|=)?\s*(" & NUMBER_PATTERN & ")", "i")
    rSquared = Null
    If Not rMatch Is Nothing Then rSquared = JsNumber(rMatch.Group(1))
    If Not fMatch Is Nothing Then
        Set fTest = NewMap()
        fTest("statistic") = JsNumber(fMatch.Group(1)): fTest("df1") = JsNumber(fMatch.Group(2)): fTest("df2") = JsNumber(fMatch.Group(3))
        pOperator = fMatch.Group(4)
        If Len(pOperator) = 0 Then pOperator = "="
        fTest("pOperator") = pOperator: fTest("pRaw") = fMatch.Group(5): fTest("pValue") = JsNumber(fMatch.Group(5))
    End If
    For i = 0 To coefficients.Count - 1
        Set c = coefficients(i)
        If IsNull(PValueDecision(c("pOperator"), c("pValue"))) Then invalid = True
    Next i
    If Not fTest Is Nothing Then
        If IsNull(PValueDecision(fTest("pOperator"), fTest("pValue"))) Then invalid = True
    End If
    If invalid Then Set ParseMultiRegressionOutput = Fail("A coefficient or overall F-test p-value does not establish significance at 0.05. Paste a more precise p-value."): Exit Function
    Set result = NewMap()
    result("ok") = True
    result("outcome") = outcome
    Set result("coefficients") = coefficients
    If Not IsNull(rSquared) Then
        If Not JsIsFinite(CDbl(rSquared)) Then rSquared = Null
    End If
    result("rSquared") = rSquared
    If Not fTest Is Nothing Then
        If Not (JsIsFinite(fTest("statistic")) And JsIsFinite(fTest("df1")) And JsIsFinite(fTest("df2")) And JsIsFinite(fTest("pValue"))) Then Set fTest = Nothing
    End If
    Set result("fTest") = fTest
    Set ParseMultiRegressionOutput = result
End Function

Private Function DegreesOfFreedom(ByVal value As Double) As String
    If JsIsInteger(value) Then DegreesOfFreedom = JsNumberToString(value) Else DegreesOfFreedom = FormatFinalNumber(value)
End Function

Private Function MultiRegressionCoefficientText(ByVal c As clsMap, ByVal names As clsList, ByVal outcome As String) As String
    Dim others As clsList, controlLabels As clsList, i As Long, direction As String, significant As Variant, pDisplay As String
    Dim pRelation As String, effect As String, change As String, parameterLabel As String, estimate As Double
    Set others = NewList(): Set controlLabels = NewList()
    For i = 0 To names.Count - 1
        If names(i) <> c("name") Then
            others.Push names(i)
            controlLabels.Push RegressionDisplayName(names(i))
        End If
    Next i
    estimate = c("estimate")
    If estimate < 0 Then
        direction = "decreases"
    ElseIf estimate > 0 Then
        direction = "increases"
    Else
        direction = "does not change"
    End If
    significant = PValueDecision(c("pOperator"), c("pValue"))
    pDisplay = NormaliseScientificToken(c("pRaw"))
    pRelation = PValueComparison(c("pOperator"), c("pValue"))
    effect = FormatFinalNumber(JsAbs(estimate))
    If direction = "does not change" Then change = direction & " (estimated change = 0 units)" Else change = direction & " by " & effect & " units"
    parameterLabel = RegressionDisplayName(c("name"))
    MultiRegressionCoefficientText = "The regression coefficient estimate of " & c("name") & " was " & NormaliseScientificToken(c("estimateRaw")) & _
        ". This implies that, holding " & JoinRegressionNames(others) & " constant, for every 1-unit increase in " & c("name") & ", " & outcome & " " & change & _
        ". Population parameter " & parameterLabel & " has a t-value = " & FormatFinalNumber(c("tValue")) & " and p-value " & c("pOperator") & " " & pDisplay & _
        ", which is " & pRelation & " 0.05. Hence, we " & IIf(significant, "reject", "fail to reject") & " the null hypothesis at the 5% significance level. There is " & _
        IIf(significant, "sufficient", "insufficient") & " evidence to conclude that " & parameterLabel & " has a statistically significant effect on " & _
        RegressionDisplayName(outcome) & " after controlling for " & JoinRegressionNames(controlLabels) & "."
End Function

Public Sub RenderMultiRegressionInterpretation()
    Dim container As clsNode, copyAll As clsNode, parsed As clsMap, inputText As String, names As clsList, coefficients As clsList
    Dim i As Long, c As clsMap, outcome As String, rSquared As Double, explained As Double, unexplained As Double
    Dim f As clsMap, significant As Variant, pDisplay As String, pRelation As String, betas As clsList, decision As String
    Dim html As String, s As clsMap, missingR2 As String, missingF As String
    Set container = Q("#multiRegressionProse")
    Set copyAll = Q("#copyMultiRegressionProseBtn")
    inputText = Q("#multiRegressionOutput").Value
    Set parsed = ParseMultiRegressionOutput(inputText)
    Set MultiRegressionProseSections = NewList()
    Set MultiRegressionModelSections = NewList()
    If Len(JsTrim(inputText)) = 0 Then
        container.InnerHtml = "<div class=""pca-variance-prose"">Paste a multi-variable linear-regression summary to interpret every predictor.</div>"
        copyAll.Disabled = True
        Exit Sub
    End If
    If Not parsed("ok") Then
        container.InnerHtml = "<div class=""pca-variance-prose is-error"">" & EscapeHtml(parsed("message")) & "</div>"
        copyAll.Disabled = True
        Exit Sub
    End If
    outcome = parsed("outcome")
    Set coefficients = parsed("coefficients")
    Set names = NewList()
    For i = 0 To coefficients.Count - 1
        Set c = coefficients(i)
        names.Push c("name")
    Next i
    If Not IsNull(parsed("rSquared")) Then
        rSquared = parsed("rSquared")
        explained = rSquared * 100
        explained = JsMax(0, JsMin(100, explained))
        unexplained = 100 - explained
        MultiRegressionModelSections.Push Section("title", "Multiple R-squared", "The Multiple R-squared value is " & FormatFinalNumber(rSquared) & _
            ". This shows that approximately " & JsToFixed(explained, 1) & "% of the variation in " & outcome & " is explained by the predictors " & _
            JoinRegressionNames(names) & ". The remaining " & JsToFixed(unexplained, 1) & "% is left unexplained by this model.")
    End If
    Set f = parsed("fTest")
    If Not f Is Nothing Then
        significant = PValueDecision(f("pOperator"), f("pValue"))
        pDisplay = NormaliseScientificToken(f("pRaw"))
        pRelation = PValueComparison(f("pOperator"), f("pValue"))
        Set betas = NewList()
        For i = 0 To names.Count - 1
            betas.Push SYM_BETA() & names(i)
        Next i
        If significant Then
            decision = "we reject " & SYM_H0() & " and conclude that at least one predictor significantly explains variation in " & outcome & "."
        Else
            decision = "we fail to reject " & SYM_H0() & " and conclude that there is insufficient evidence that at least one predictor significantly explains variation in " & outcome & "."
        End If
        MultiRegressionModelSections.Push Section("title", "Overall model F-test", SYM_H0() & ": " & betas.Join(" = ") & " = 0." & vbLf & vbLf & _
            SYM_H1() & ": At least one of " & JoinRegressionNames(betas) & " " & SYM_NE() & " 0." & vbLf & vbLf & _
            "F-test result: F-statistic = " & FormatFinalNumber(f("statistic")) & " on " & DegreesOfFreedom(f("df1")) & " and " & DegreesOfFreedom(f("df2")) & _
            " degrees of freedom, p-value " & f("pOperator") & " " & pDisplay & ". Since the p-value " & f("pOperator") & " " & pDisplay & ", which is " & pRelation & " 0.05, " & decision)
    End If
    For i = 0 To coefficients.Count - 1
        Set c = coefficients(i)
        MultiRegressionProseSections.Push Section("predictor", c("name"), MultiRegressionCoefficientText(c, names, outcome))
    Next i
    If IsNull(parsed("rSquared")) Then missingR2 = "<article class=""sensitivity-prose-card""><div class=""sensitivity-prose-head""><h3>Multiple R-squared</h3></div><p class=""pca-variance-prose is-error"">Multiple R-squared was not found in the pasted output.</p></article>"
    If f Is Nothing Then missingF = "<article class=""sensitivity-prose-card""><div class=""sensitivity-prose-head""><h3>Overall model F-test</h3></div><p class=""pca-variance-prose is-error"">The F-statistic, degrees of freedom, or model p-value was not found.</p></article>"
    For i = 0 To MultiRegressionModelSections.Count - 1
        Set s = MultiRegressionModelSections(i)
        html = html & ProseCard(EscapeHtml(s("title")), "data-copy-multi-model", i, s("text"))
    Next i
    html = html & missingR2 & missingF
    For i = 0 To MultiRegressionProseSections.Count - 1
        Set s = MultiRegressionProseSections(i)
        html = html & ProseCard(EscapeHtml(s("predictor")), "data-copy-multi-regression", i, s("text"))
    Next i
    container.InnerHtml = html
    copyAll.Disabled = False
End Sub

' ---------------------------------------------------------------- coefficient intervals

Public Function ParseCoefficientIntervals(ByVal text As String) As clsMap
    Dim lines As clsList, headerIndex As Long, i As Long, percentages As clsList, matches As Collection, m As clsMatch
    Dim confidence As Double, rowRx As clsRegex, rows As clsList, line As String, parameter As String, first As Double, second As Double
    Dim row As clsMap, result As clsMap
    Set lines = SplitLines(Replace(text, vbCr, ""))
    headerIndex = -1
    For i = 0 To lines.Count - 1
        If RxTest(CStr(lines(i)), "(?:\d+(?:\.\d+)?)\s*%[\s\S]*(?:\d+(?:\.\d+)?)\s*%") Then headerIndex = i: Exit For
    Next i
    If headerIndex < 0 Then Set ParseCoefficientIntervals = Fail("I could not find two percentage headings such as 2.5% and 97.5%."): Exit Function
    Set percentages = NewList()
    Set matches = Rx("\d+(?:\.\d+)?(?=\s*%)", "g").MatchAll(CStr(lines(headerIndex)))
    For Each m In matches
        percentages.Push JsNumber(m.Value)
    Next m
    If percentages.Count < 2 Then Set ParseCoefficientIntervals = Fail("I found the interval heading but could not read both endpoint percentages."): Exit Function
    confidence = percentages(1) - percentages(0)
    confidence = Abs(confidence)
    Set rowRx = Rx("^\s*(.+?)\s+(" & NUMBER_PATTERN & ")\s+(" & NUMBER_PATTERN & ")\s*$", "i")
    Set rows = NewList()
    For i = headerIndex + 1 To lines.Count - 1
        line = lines(i)
        If Len(JsTrim(line)) = 0 Then GoTo NextLine
        Set m = rowRx.Exec(line)
        If m Is Nothing Then GoTo NextLine
        parameter = StripBackticks(JsTrim(m.Group(1)))
        If RxTest(parameter, "^\(?Intercept\)?$", "i") Then GoTo NextLine
        first = JsNumber(m.Group(2)): second = JsNumber(m.Group(3))
        If Not JsIsFinite(first) Or Not JsIsFinite(second) Then GoTo NextLine
        Set row = NewMap()
        row("parameter") = parameter
        If first < second Then
            row("lower") = first: row("upper") = second
        Else
            row("lower") = second: row("upper") = first
        End If
        rows.Push row
NextLine:
    Next i
    If rows.Count = 0 Then Set ParseCoefficientIntervals = Fail("I found the confidence-level headings but could not read any non-intercept parameter rows."): Exit Function
    Set result = NewMap()
    result("ok") = True
    result("confidence") = confidence
    Set result("rows") = rows
    Set ParseCoefficientIntervals = result
End Function

Public Sub RenderCoefficientIntervals()
    Dim inputText As String, container As clsNode, copyAll As clsNode, parsed As clsMap, dependent As String, advY As clsNode
    Dim confidenceLabel As String, rows As clsList, i As Long, row As clsMap, lower As String, upper As String
    Dim above As Boolean, below As Boolean, position As String, evidence As String, parameter As String, html As String, s As clsMap
    inputText = Q("#coefficientIntervalOutput").Value
    Set container = Q("#coefficientIntervalProse")
    Set copyAll = Q("#copyCoefficientIntervalBtn")
    Set parsed = ParseCoefficientIntervals(inputText)
    Set advY = Q("#advY")
    If Not advY Is Nothing Then dependent = JsTrim(advY.Value)
    If Len(dependent) = 0 Then dependent = "the dependent variable"
    Set CoefficientIntervalSections = NewList()
    If Len(JsTrim(inputText)) = 0 Then
        container.InnerHtml = "<div class=""pca-variance-prose"">Paste the coefficient confidence intervals to assess which parameters differ significantly from 0.</div>"
        copyAll.Disabled = True
        Exit Sub
    End If
    If Not parsed("ok") Then
        container.InnerHtml = "<div class=""pca-variance-prose is-error"">" & EscapeHtml(parsed("message")) & "</div>"
        copyAll.Disabled = True
        Exit Sub
    End If
    If JsIsInteger(parsed("confidence")) Then confidenceLabel = JsNumberToString(parsed("confidence")) Else confidenceLabel = FormatFinalNumber(parsed("confidence"))
    Set rows = parsed("rows")
    For i = 0 To rows.Count - 1
        Set row = rows(i)
        parameter = row("parameter")
        lower = FormatFinalNumber(row("lower")): upper = FormatFinalNumber(row("upper"))
        above = row("lower") > 0: below = row("upper") < 0
        If above Then
            position = "above"
        ElseIf below Then
            position = "below"
        Else
            position = "across"
        End If
        If above Or below Then
            evidence = "The relationship between " & parameter & " and " & dependent & " is statistically significant because the entire confidence interval lies " & position & _
                " 0; thus, we have sufficient evidence to reject the null hypothesis that " & parameter & " has no effect on " & dependent & ", assuming all other variables are kept constant."
        Else
            evidence = "The relationship between " & parameter & " and " & dependent & " is not statistically significant because the confidence interval contains 0; thus, we do not have sufficient evidence to reject the null hypothesis that " & _
                parameter & " has no effect on " & dependent & ", assuming all other variables are kept constant."
        End If
        CoefficientIntervalSections.Push Section("parameter", parameter, "The " & confidenceLabel & "% confidence interval for the " & parameter & " parameter is (" & lower & ", " & upper & _
            "). This means that we are " & confidenceLabel & "% confident that the true value of " & parameter & " lies between " & lower & " and " & upper & ". " & evidence)
    Next i
    For i = 0 To CoefficientIntervalSections.Count - 1
        Set s = CoefficientIntervalSections(i)
        html = html & ProseCard(EscapeHtml(s("parameter")), "data-copy-coefficient-interval", i, s("text"))
    Next i
    container.InnerHtml = html
    copyAll.Disabled = False
End Sub

' ---------------------------------------------------------------- confusion matrix

Private Function ConfusionDefinition(ByVal title As String) As String
    Select Case title
        Case "Classification counts": ConfusionDefinition = "True negatives and true positives are correct predictions. False negatives and false positives are classification errors."
        Case "Accuracy": ConfusionDefinition = "The proportion of all observations classified correctly: (TP + TN) / total."
        Case "Sensitivity": ConfusionDefinition = "The proportion of actual positive cases correctly predicted as positive: TP / (TP + FN). It is also called recall or the true-positive rate."
        Case "Specificity": ConfusionDefinition = "The proportion of actual negative cases correctly predicted as negative: TN / (TN + FP), also called the true-negative rate."
        Case "Precision": ConfusionDefinition = "The proportion of predicted positive cases that are actually positive: TP / (TP + FP)."
    End Select
End Function

' readMetric: Nothing when absent, else {value, raw, decimals}.
Private Function ReadConfusionMetric(ByVal text As String, ByVal name As String) As clsMap
    Dim m As clsMatch, raw As String, result As clsMap
    Set m = RxMatch(text, name & "\s{0,3}:\s{0,3}([-+]?(?:\d+\.?\d*|\.\d+))", "i")
    If m Is Nothing Then Exit Function
    raw = m.Group(1)
    Set result = NewMap()
    result("value") = JsNumber(raw)
    result("raw") = raw
    If JsIndexOf(raw, ".") >= 0 Then result("decimals") = Len(CStr(JsSplit(raw, ".")(1))) Else result("decimals") = 0&
    Set ReadConfusionMetric = result
End Function

Private Function ConfusionFormula(ByVal name As String, ByVal tn As Double, ByVal fn As Double, ByVal fp As Double, ByVal tp As Double, ByVal total As Double) As String
    Select Case name
        Case "Sensitivity": ConfusionFormula = "TP / (TP + FN) = " & JsNumberToString(tp) & " / (" & JsNumberToString(tp) & " + " & JsNumberToString(fn) & ")"
        Case "Specificity": ConfusionFormula = "TN / (TN + FP) = " & JsNumberToString(tn) & " / (" & JsNumberToString(tn) & " + " & JsNumberToString(fp) & ")"
        Case Else: ConfusionFormula = "(TP + TN) / total = (" & JsNumberToString(tp) & " + " & JsNumberToString(tn) & ") / " & JsNumberToString(total)
    End Select
End Function

' a / b, or Null when b is not positive (the website's "denominator > 0 ? a / b : null").
Private Function RatioOrNull(ByVal numerator As Double, ByVal denominator As Double) As Variant
    Dim quotient As Double
    If denominator > 0 Then
        quotient = numerator / denominator
        RatioOrNull = quotient
    Else
        RatioOrNull = Null
    End If
End Function

Public Function ParseConfusionOutput(ByVal text As String) As clsMap
    Dim lines As clsList, headerIndex As Long, i As Long, headings As clsList, negativeLabel As String, positiveLabel As String
    Dim rowRx As clsRegex, rows As clsList, m As clsMatch, row As clsMap, tn As Double, fn As Double, fp As Double, tp As Double, total As Double
    Dim names As Variant, name As Variant, stated As clsMap, expected As Double, tolerance As Double, difference As Double, power As Double
    Dim reported As clsMap, calculated As clsMap, result As clsMap, value As Variant, denominator As Double
    Set lines = SplitLines(Replace(text, vbCr, ""))
    headerIndex = -1
    For i = 0 To lines.Count - 1
        If RxTest(CStr(lines(i)), "^\s*Prediction\s+\S+\s+\S+\s*$", "i") Then headerIndex = i: Exit For
    Next i
    If headerIndex < 0 Then Set ParseConfusionOutput = Fail("I could not find the " & SYM_LSQUO() & "Prediction" & SYM_RSQUO() & " row with two reference classes."): Exit Function
    Set headings = Rx("\s+").Split(JsTrim(CStr(lines(headerIndex))))
    negativeLabel = headings(1): positiveLabel = headings(2)
    Set rowRx = Rx("^\s*(\S+)\s+([-+]?(?:\d+\.?\d*|\.\d+))\s+([-+]?(?:\d+\.?\d*|\.\d+))\s*$")
    Set rows = NewList()
    For i = headerIndex + 1 To lines.Count - 1
        Set m = rowRx.Exec(CStr(lines(i)))
        If Not m Is Nothing Then
            Set row = NewMap()
            row("label") = m.Group(1): row("first") = JsNumber(m.Group(2)): row("second") = JsNumber(m.Group(3))
            rows.Push row
        End If
        If rows.Count = 2 Then Exit For
    Next i
    If rows.Count < 2 Then Set ParseConfusionOutput = Fail("I found the class headings but could not read both prediction rows."): Exit Function
    Set row = rows(0): tn = row("first"): fn = row("second")
    Set row = rows(1): fp = row("first"): tp = row("second")
    If Not (JsIsFinite(tn) And JsIsFinite(fn) And JsIsFinite(fp) And JsIsFinite(tp)) Then Set ParseConfusionOutput = Fail("I found the class headings but could not read both prediction rows."): Exit Function
    total = tn + fn
    total = total + fp
    total = total + tp
    If tn < 0 Or Not JsIsInteger(tn) Or fn < 0 Or Not JsIsInteger(fn) Or fp < 0 Or Not JsIsInteger(fp) Or tp < 0 Or Not JsIsInteger(tp) Then
        Set ParseConfusionOutput = Fail("Every confusion-matrix cell must be a non-negative whole-number count."): Exit Function
    End If
    If Not (total > 0) Then Set ParseConfusionOutput = Fail("The four confusion-matrix counts must add to a positive total."): Exit Function
    Set reported = NewMap()
    Set reported("Sensitivity") = ReadConfusionMetric(text, "Sensitivity")
    Set reported("Specificity") = ReadConfusionMetric(text, "Specificity")
    Set reported("Accuracy") = ReadConfusionMetric(text, "Accuracy")
    Set calculated = NewMap()
    denominator = tp + fn: calculated("Sensitivity") = RatioOrNull(tp, denominator)
    denominator = tn + fp: calculated("Specificity") = RatioOrNull(tn, denominator)
    denominator = tp + tn: calculated("Accuracy") = RatioOrNull(denominator, total)
    names = Array("Sensitivity", "Specificity", "Accuracy")
    For Each name In names
        Set stated = reported(name)
        If Not stated Is Nothing Then
            If Not JsIsFinite(stated("value")) Or stated("value") < 0 Or stated("value") > 1 Then Set ParseConfusionOutput = Fail("Reported " & name & " must be a number from 0 to 1."): Exit Function
            If IsNull(calculated(name)) Then Set ParseConfusionOutput = Fail(name & " is undefined because its denominator is zero."): Exit Function
            expected = calculated(name)
            power = JsPow(10, -CDbl(stated("decimals")))
            tolerance = 0.5 * power
            tolerance = tolerance + JsEpsilon()
            difference = stated("value") - expected
            difference = Abs(difference)
            If difference > tolerance Then
                Set ParseConfusionOutput = Fail("Reported " & name & " = " & stated("raw") & " does not match the grid. Using " & ConfusionFormula(name, tn, fn, fp, tp, total) & _
                    ", " & name & " should be " & JsToFixed(expected, stated("decimals")) & ". Correct the matrix or the reported statistic.")
                Exit Function
            End If
        End If
    Next name
    Set result = NewMap()
    result("ok") = True
    result("negativeLabel") = negativeLabel: result("positiveLabel") = positiveLabel
    result("tn") = tn: result("fn") = fn: result("fp") = fp: result("tp") = tp: result("total") = total
    For Each name In names
        Set stated = reported(name)
        If stated Is Nothing Then value = calculated(name) Else value = stated("value")
        result(JsLower(name)) = value
    Next name
    denominator = tp + fp
    result("precision") = RatioOrNull(tp, denominator)
    Set ParseConfusionOutput = result
End Function

Public Function ConfusionInfo(ByVal title As String) As String
    ConfusionInfo = "<button class=""metric-info"" type=""button"" aria-label=""What is " & EscapeHtml(title) & "?"">i<span class=""metric-tooltip"" role=""tooltip"">" & _
        EscapeHtml(ConfusionDefinition(title)) & "</span></button>"
End Function

' formatFinalNumber / pct for values that may be JS null.
Private Function FinalOrNull(ByVal value As Variant) As String
    If IsNull(value) Then FinalOrNull = "null" Else FinalOrNull = FormatFinalNumber(value)
End Function

Private Function Pct(ByVal value As Variant) As String
    Dim scaled As Double
    If IsNull(value) Then Pct = "0.0%": Exit Function
    scaled = value * 100
    Pct = JsToFixed(scaled, 1) & "%"
End Function

Public Sub RenderConfusionInterpretation()
    Dim inputText As String, container As clsNode, copyAll As clsNode, p As clsMap, html As String, i As Long, s As clsMap
    Dim negative As String, positive As String
    inputText = Q("#confusionOutput").Value
    Set container = Q("#confusionProse")
    Set copyAll = Q("#copyConfusionProseBtn")
    Set p = ParseConfusionOutput(inputText)
    Set ConfusionProseSections = NewList()
    If Len(JsTrim(inputText)) = 0 Then
        container.InnerHtml = "<div class=""pca-variance-prose"">Paste a confusion matrix to interpret its classification counts and performance measures.</div>"
        copyAll.Disabled = True
        Exit Sub
    End If
    If Not p("ok") Then
        container.InnerHtml = "<div class=""pca-variance-prose is-error"">" & EscapeHtml(p("message")) & "</div>"
        copyAll.Disabled = True
        Exit Sub
    End If
    negative = p("negativeLabel"): positive = p("positiveLabel")
    ConfusionProseSections.Push Section("title", "Classification counts", "The model correctly classified " & JsNumberToString(p("tn")) & " " & negative & " observations and " & _
        JsNumberToString(p("tp")) & " " & positive & " observations. It incorrectly classified " & JsNumberToString(p("fp")) & " " & negative & " observations as " & positive & _
        " (false positives) and " & JsNumberToString(p("fn")) & " " & positive & " observations as " & negative & " (false negatives).")
    ConfusionProseSections.Push Section("title", "Accuracy", "The accuracy is " & FinalOrNull(p("accuracy")) & ", meaning that the model correctly classified " & Pct(p("accuracy")) & _
        " of all " & JsNumberToString(p("total")) & " observations.")
    ConfusionProseSections.Push Section("title", "Sensitivity", "The sensitivity is " & FinalOrNull(p("sensitivity")) & ", meaning that the model correctly identified " & Pct(p("sensitivity")) & _
        " of the actual " & positive & " observations.")
    ConfusionProseSections.Push Section("title", "Specificity", "The specificity is " & FinalOrNull(p("specificity")) & ", meaning that the model correctly identified " & Pct(p("specificity")) & _
        " of the actual " & negative & " observations.")
    If Not IsNull(p("precision")) Then
        ConfusionProseSections.Push Section("title", "Precision", "The precision is " & FormatFinalNumber(p("precision")) & ", meaning that " & Pct(p("precision")) & _
            " of observations predicted as " & positive & " were actually " & positive & ".")
    End If
    For i = 0 To ConfusionProseSections.Count - 1
        Set s = ConfusionProseSections(i)
        html = html & ProseCard(EscapeHtml(s("title")) & " " & ConfusionInfo(s("title")), "data-copy-confusion", i, s("text"), " class=""metric-title""")
    Next i
    container.InnerHtml = html
    copyAll.Disabled = False
End Sub

' ---------------------------------------------------------------- event handlers

Private Function ActiveId() As String
    ActiveId = CStr(AppState("id"))
End Function

Public Sub OnGlmOutputInput(ByVal e As clsEvent)
    On Error GoTo Failed
    RenderGlmInterpretation
    Exit Sub
Failed:
    ReportError "opRegression2.OnGlmOutputInput"
End Sub

Public Sub OnLoadGlmOutput(ByVal e As clsEvent)
    On Error GoTo Failed
    Q("#advDataset").Value = "dataset"
    Q("#advResponse").Value = "HIV_Status"
    Q("#advPredictors").Value = "HAART_A, HAART_B, VLoad, Age"
    Q("#advancedRegressionCode").TextContent = AdvancedRegressionCode("glm")
    Q("#glmOutput").Value = GlmSampleOutput()
    RenderGlmInterpretation
    ShowToast "GLM sample output loaded"
    Exit Sub
Failed:
    ReportError "opRegression2.OnLoadGlmOutput"
End Sub

Public Sub OnCopyGlmProse(ByVal e As clsEvent)
    On Error GoTo Failed
    If GlmProseSections.Count = 0 Then Exit Sub
    CopyText JoinSectionTexts(GlmProseSections), "All GLM interpretations copied"
    Exit Sub
Failed:
    ReportError "opRegression2.OnCopyGlmProse"
End Sub

Public Sub OnGlmProseClick(ByVal e As clsEvent)
    Dim button As clsNode, index As Long, s As clsMap
    On Error GoTo Failed
    Set button = e.Target.Closest("[data-copy-glm]")
    If button Is Nothing Then Exit Sub
    index = SectionIndex(button.DataGet("copyGlm"), GlmProseSections)
    If index < 0 Then Exit Sub
    Set s = GlmProseSections(index)
    CopyText s("text"), s("predictor") & " interpretation copied"
    Exit Sub
Failed:
    ReportError "opRegression2.OnGlmProseClick"
End Sub

Public Sub OnModelComparisonOutputInput(ByVal e As clsEvent)
    On Error GoTo Failed
    RenderModelComparisonInterpretation
    Exit Sub
Failed:
    ReportError "opRegression2.OnModelComparisonOutputInput"
End Sub

Public Sub OnLoadModelComparisonOutput(ByVal e As clsEvent)
    On Error GoTo Failed
    Q("#modelComparisonOutput").Value = ModelComparisonSample()
    RenderModelComparisonInterpretation
    ShowToast "Model-comparison sample loaded"
    Exit Sub
Failed:
    ReportError "opRegression2.OnLoadModelComparisonOutput"
End Sub

Public Sub OnCopyModelComparisonProse(ByVal e As clsEvent)
    On Error GoTo Failed
    If Len(ModelComparisonProse) = 0 Then Exit Sub
    CopyText ModelComparisonProse, "Model-comparison prose copied"
    Exit Sub
Failed:
    ReportError "opRegression2.OnCopyModelComparisonProse"
End Sub

Public Sub OnMultiRegressionOutputInput(ByVal e As clsEvent)
    On Error GoTo Failed
    RenderMultiRegressionInterpretation
    Exit Sub
Failed:
    ReportError "opRegression2.OnMultiRegressionOutputInput"
End Sub

Public Sub OnLoadMultiRegressionOutput(ByVal e As clsEvent)
    On Error GoTo Failed
    Q("#multiRegressionOutput").Value = MultiRegressionSample()
    If ActiveId() = "multi-regression" Then
        Q("#advDataset").Value = "sleepdata"
        Q("#advY").Value = "sleep"
        Q("#advPredictors").Value = "age, educ, income, totwrk"
        RenderAdvancedRegressionCodeOutputs "multi-regression"
    End If
    RenderMultiRegressionInterpretation
    ShowToast "Multi-variable regression sample loaded"
    Exit Sub
Failed:
    ReportError "opRegression2.OnLoadMultiRegressionOutput"
End Sub

Public Sub OnCopyMultiRegressionProse(ByVal e As clsEvent)
    Dim sections As clsList
    On Error GoTo Failed
    Set sections = MultiRegressionModelSections.Concat(MultiRegressionProseSections)
    If sections.Count = 0 Then Exit Sub
    CopyText JoinSectionTexts(sections), "All regression interpretations copied", True
    Exit Sub
Failed:
    ReportError "opRegression2.OnCopyMultiRegressionProse"
End Sub

Public Sub OnMultiRegressionProseClick(ByVal e As clsEvent)
    Dim button As clsNode, index As Long, s As clsMap
    On Error GoTo Failed
    Set button = e.Target.Closest("[data-copy-multi-model]")
    If Not button Is Nothing Then
        index = SectionIndex(button.DataGet("copyMultiModel"), MultiRegressionModelSections)
        If index < 0 Then Exit Sub
        Set s = MultiRegressionModelSections(index)
        CopyText s("text"), s("title") & " interpretation copied"
        Exit Sub
    End If
    Set button = e.Target.Closest("[data-copy-multi-regression]")
    If button Is Nothing Then Exit Sub
    index = SectionIndex(button.DataGet("copyMultiRegression"), MultiRegressionProseSections)
    If index < 0 Then Exit Sub
    Set s = MultiRegressionProseSections(index)
    CopyText s("text"), s("predictor") & " interpretation copied", True
    Exit Sub
Failed:
    ReportError "opRegression2.OnMultiRegressionProseClick"
End Sub

Public Sub OnConfusionOutputInput(ByVal e As clsEvent)
    On Error GoTo Failed
    RenderConfusionInterpretation
    Exit Sub
Failed:
    ReportError "opRegression2.OnConfusionOutputInput"
End Sub

Public Sub OnLoadConfusionOutput(ByVal e As clsEvent)
    On Error GoTo Failed
    Q("#confusionOutput").Value = ConfusionSampleOutput()
    RenderConfusionInterpretation
    ShowToast "Confusion-matrix sample loaded"
    Exit Sub
Failed:
    ReportError "opRegression2.OnLoadConfusionOutput"
End Sub

Public Sub OnCopyConfusionProse(ByVal e As clsEvent)
    On Error GoTo Failed
    If ConfusionProseSections.Count = 0 Then Exit Sub
    CopyText JoinSectionTexts(ConfusionProseSections), "All confusion-matrix interpretations copied"
    Exit Sub
Failed:
    ReportError "opRegression2.OnCopyConfusionProse"
End Sub

Public Sub OnConfusionProseClick(ByVal e As clsEvent)
    Dim button As clsNode, index As Long, s As clsMap
    On Error GoTo Failed
    Set button = e.Target.Closest("[data-copy-confusion]")
    If button Is Nothing Then Exit Sub
    index = SectionIndex(button.DataGet("copyConfusion"), ConfusionProseSections)
    If index < 0 Then Exit Sub
    Set s = ConfusionProseSections(index)
    CopyText s("text"), s("title") & " interpretation copied"
    Exit Sub
Failed:
    ReportError "opRegression2.OnConfusionProseClick"
End Sub

Public Sub OnCoefficientIntervalOutputInput(ByVal e As clsEvent)
    On Error GoTo Failed
    RenderCoefficientIntervals
    Exit Sub
Failed:
    ReportError "opRegression2.OnCoefficientIntervalOutputInput"
End Sub

Public Sub OnLoadCoefficientInterval(ByVal e As clsEvent)
    On Error GoTo Failed
    Q("#coefficientIntervalOutput").Value = CoefficientIntervalSample()
    If ActiveId() = "multi-regression" Then
        If Not Q("#advY") Is Nothing Then Q("#advY").Value = "sleep"
    End If
    RenderCoefficientIntervals
    ShowToast "Coefficient-interval sample loaded"
    Exit Sub
Failed:
    ReportError "opRegression2.OnLoadCoefficientInterval"
End Sub

Public Sub OnCopyCoefficientInterval(ByVal e As clsEvent)
    On Error GoTo Failed
    If CoefficientIntervalSections.Count = 0 Then Exit Sub
    CopyText JoinSectionTexts(CoefficientIntervalSections), "All interval interpretations copied"
    Exit Sub
Failed:
    ReportError "opRegression2.OnCopyCoefficientInterval"
End Sub

Public Sub OnCoefficientIntervalProseClick(ByVal e As clsEvent)
    Dim button As clsNode, index As Long, s As clsMap
    On Error GoTo Failed
    Set button = e.Target.Closest("[data-copy-coefficient-interval]")
    If button Is Nothing Then Exit Sub
    index = SectionIndex(button.DataGet("copyCoefficientInterval"), CoefficientIntervalSections)
    If index < 0 Then Exit Sub
    Set s = CoefficientIntervalSections(index)
    CopyText s("text"), s("parameter") & " interval interpretation copied"
    Exit Sub
Failed:
    ReportError "opRegression2.OnCoefficientIntervalProseClick"
End Sub
