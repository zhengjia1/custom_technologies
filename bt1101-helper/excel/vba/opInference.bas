Attribute VB_Name = "opInference"
' Tutorial 5: one-sample mean/proportion tests and two-sample mean/variance tests (code
' generators, output parsers, interpretations), plus the event wiring of all Tutorial 5
' inference operations. Ported from app-src.js lines 1696-1863 and 4563-4703; the
' intervals and distribution functions (3453-3781) are in opInference2.
Option Explicit

Private Type TOneSample
    parameter As String
    alternative As String
    dataset As String
    column As String
    nullValue As Double
    alpha As Double
    filter As String
    excludeMissing As Boolean
    population As String
    condition As String
    success As String
End Type

Private Type TTwoSample
    parameter As String
    design As String
    alternative As String
    dataset As String
    response1 As String
    response2 As String
    alpha As Double
End Type

' Prose captured by the dynamically created "Copy prose" buttons (the website's closures).
Private mOneSampleProse As String
Private mTwoSampleProse As String
' Used when the shell (modApp) is absent: tests set it like the shell's test mode.
Public TestPasteUnavailable As Boolean

' Performs the website's top-level statements for lines 3616 and 4563-4703 in order.
Public Sub InitOpInference()
    Dim list As clsList, i As Long
    mOneSampleProse = ""
    mTwoSampleProse = ""
    ResetIntervalState
    KnownSdAlertProcedure = "ShowKnownPopulationSdAlert"
    Q("#twoSampleParameter").AddEventListener "change", "opInference.OnTwoSampleSync"
    Q("#twoSamplePaired").AddEventListener "change", "opInference.OnTwoSampleSync"
    Q("#twoSampleAlternative").AddEventListener "change", "opInference.OnTwoSampleSync"
    Set list = QA(".two-sample-input")
    For i = 0 To list.Count - 1
        list(i).AddEventListener "input", "opInference.OnTwoSampleSync"
    Next i
    Q("#twoSampleOutputText").AddEventListener "input", "opInference.OnTwoSampleOutputInput"
    Q("#loadTwoSampleOutputBtn").AddEventListener "click", "opInference.OnLoadTwoSampleOutput"
    Q("#pasteTwoSampleBtn").AddEventListener "click", "opInference.OnPasteTwoSample"
    Q("#copyTwoSampleCodeBtn").AddEventListener "click", "opInference.OnCopyTwoSampleCode"
    Q("#oneSampleAlternative").AddEventListener "change", "opInference.OnOneSampleSync"
    Q("#oneSampleParameter").AddEventListener "change", "opInference.OnOneSampleParameterChange"
    Set list = QA(".one-sample-input")
    For i = 0 To list.Count - 1
        list(i).AddEventListener "input", "opInference.OnOneSampleSync"
    Next i
    Q("#oneSampleExcludeMissing").AddEventListener "change", "opInference.OnOneSampleSync"
    Q("#oneSampleOutputText").AddEventListener "input", "opInference.OnOneSampleOutputInput"
    Q("#loadOneSampleSetupBtn").AddEventListener "click", "opInference.OnLoadOneSampleSetup"
    Q("#loadOneSampleOutputBtn").AddEventListener "click", "opInference.OnLoadOneSampleOutput"
    Q("#pasteOneSampleBtn").AddEventListener "click", "opInference.OnPasteOneSample"
    Q("#copyOneSampleCodeBtn").AddEventListener "click", "opInference.OnCopyOneSampleCode"
    Q("#intervalCalcType").AddEventListener "change", "opInference2.OnIntervalSync"
    Q("#intervalParameter").AddEventListener "change", "opInference2.OnIntervalSync"
    Q("#intervalProportionMode").AddEventListener "change", "opInference2.OnIntervalSync"
    Q("#intervalConditionOperator").AddEventListener "change", "opInference2.OnIntervalRender"
    Q("#intervalExcludeMissing").AddEventListener "change", "opInference2.OnIntervalRender"
    Q("#intervalStdev").AddEventListener "change", "opInference2.OnIntervalStdevChange"
    Set list = QA(".interval-input")
    For i = 0 To list.Count - 1
        list(i).AddEventListener "input", "opInference2.OnIntervalRender"
    Next i
    Q("#loadIntervalSampleBtn").AddEventListener "click", "opInference2.OnLoadIntervalSample"
    Q("#intervalResultInput").AddEventListener "input", "opInference2.OnIntervalPastedInput"
    Q("#loadIntervalResultSampleBtn").AddEventListener "click", "opInference2.OnLoadIntervalResultSample"
    Q("#copyIntervalPastedProseBtn").AddEventListener "click", "opInference2.OnCopyIntervalPastedProse"
    Q("#copyIntervalProseBtn").AddEventListener "click", "opInference2.OnCopyIntervalProse"
    Q("#copyIntervalCodeBtn").AddEventListener "click", "opInference2.OnCopyIntervalCode"
End Sub

' ---------------------------------------------------------------- shared helpers

Private Function FieldText(ByVal selector As String) As String
    FieldText = JsTrim(Q(selector).Value)
End Function

Private Function OrDefault(ByVal text As String, ByVal fallback As String) As String
    If Len(text) > 0 Then OrDefault = text Else OrDefault = fallback
End Function

' pValueDecision, guarded against a NaN alpha (the website returns null for it).
Private Function SafeDecision(ByVal operator As String, ByVal pValue As Double, ByVal alpha As Double) As Variant
    If JsIsNaN(alpha) Or JsIsNaN(pValue) Then SafeDecision = Null Else SafeDecision = PValueDecision(operator, pValue, alpha)
End Function

Private Function PercentLabel(ByVal alpha As Double) As String
    PercentLabel = JsToLocaleString(JsMul(alpha, 100))
End Function

Private Function HypothesisHtml(ByVal h0 As String, ByVal h1 As String) As String
    HypothesisHtml = "<strong>Hypotheses</strong><span>" & EscapeHtml(h0) & "</span><span>" & EscapeHtml(h1) & "</span>"
End Function

Private Function EmptyInterpretationHtml() As String
    EmptyInterpretationHtml = "<div class=""empty-state compact""><div class=""empty-state-inner""><strong>Run the generated code</strong><p>Paste its RStudio output above to complete the decision.</p></div></div>"
End Function

Private Function NumberPattern() As String
    NumberPattern = "([+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?)"
End Function

' ---------------------------------------------------------------- one-sample tests

Private Function OneSampleSettings() As TOneSample
    Dim s As TOneSample
    s.parameter = Q("#oneSampleParameter").Value
    s.alternative = Q("#oneSampleAlternative").Value
    s.dataset = OrDefault(FieldText("#oneSampleDataset"), "dataset")
    s.column = OrDefault(FieldText("#oneSampleColumn"), "column1")
    s.nullValue = JsNumber(Q("#oneSampleNull").Value)
    s.alpha = JsNumber(Q("#oneSampleAlpha").Value)
    s.filter = FieldText("#oneSampleFilter")
    s.excludeMissing = Q("#oneSampleExcludeMissing").Checked
    s.population = OrDefault(FieldText("#oneSamplePopulation"), "observations")
    s.condition = FieldText("#oneSampleCondition")
    s.success = OrDefault(FieldText("#oneSampleSuccess"), "yes")
    OneSampleSettings = s
End Function

' Returns H0 and H1 through h0/h1.
Private Sub OneSampleHypotheses(ByRef s As TOneSample, ByRef h0 As String, ByRef h1 As String)
    Dim symbol As String, value As String
    If s.parameter = "mean" Then symbol = SYM_MU() Else symbol = "p"
    value = JsNumberToString(s.nullValue)
    If s.alternative = "greater" Then
        h0 = SYM_H0() & ": " & symbol & " " & SYM_LE() & " " & value
        h1 = SYM_H1() & ": " & symbol & " > " & value
    ElseIf s.alternative = "less" Then
        h0 = SYM_H0() & ": " & symbol & " " & SYM_GE() & " " & value
        h1 = SYM_H1() & ": " & symbol & " < " & value
    Else
        h0 = SYM_H0() & ": " & symbol & " = " & value
        h1 = SYM_H1() & ": " & symbol & " " & SYM_NE() & " " & value
    End If
End Sub

Private Function BuildOneSampleCode(ByRef s As TOneSample) As String
    Dim data As String, column As String, confidence As String, preparation As String, inner As String, pCode As String, nullText As String
    data = RColumnName(s.dataset)
    column = RColumnName(s.column)
    confidence = JsNumberToString(JsSub(1, s.alpha))
    nullText = JsNumberToString(s.nullValue)
    If s.excludeMissing Then
        If Len(s.filter) > 0 Then
            inner = vbLf & "    " & s.filter & "," & vbLf & "    !is.na(" & column & ")" & vbLf & "  "
        Else
            inner = "!is.na(" & column & ")"
        End If
        preparation = "analysis_data <- " & data & " %>%" & vbLf & "  dplyr::filter(" & inner & ")"
    ElseIf Len(s.filter) > 0 Then
        preparation = "analysis_data <- " & data & " %>% dplyr::filter(" & s.filter & ")"
    Else
        preparation = "analysis_data <- " & data
    End If
    If s.parameter = "mean" Then
        BuildOneSampleCode = "# Prepare the population to analyse" & vbLf & preparation & vbLf & vbLf & _
            "# One-sample t-test for a population mean" & vbLf & "t.test(analysis_data$" & column & ", mu = " & nullText & _
            ", alternative = """ & s.alternative & """, conf.level = " & confidence & ")"
        Exit Function
    End If
    If Len(s.condition) > 0 Then
        If s.alternative = "less" Then
            pCode = "pnorm(z_statistic)"
        ElseIf s.alternative = "greater" Then
            pCode = "1 - pnorm(z_statistic)"
        Else
            pCode = "2 * pnorm(-abs(z_statistic))"
        End If
        BuildOneSampleCode = "# Prepare the population to analyse" & vbLf & preparation & vbLf & vbLf & _
            "# Define a success and calculate the sample proportion" & vbLf & _
            "successes <- sum(with(analysis_data, " & s.condition & "), na.rm = TRUE)" & vbLf & _
            "n <- nrow(analysis_data)" & vbLf & "sample_proportion <- successes / n" & vbLf & vbLf & _
            "# One-sample z-test for a population proportion" & vbLf & "null_proportion <- " & nullText & vbLf & _
            "z_statistic <- (sample_proportion - null_proportion) / sqrt(null_proportion * (1 - null_proportion) / n)" & vbLf & _
            "p_value <- " & pCode & vbLf & "cat(""z ="", z_statistic, ""\n"")" & vbLf & "cat(""p-value ="", p_value, ""\n"")"
        Exit Function
    End If
    BuildOneSampleCode = "# Prepare the population to analyse" & vbLf & preparation & vbLf & vbLf & _
        "# One-sample proportion test" & vbLf & "prop.test(" & vbLf & _
        "  sum(analysis_data$" & column & " == """ & RString(s.success) & """, na.rm = TRUE)," & vbLf & _
        "  sum(!is.na(analysis_data$" & column & "))," & vbLf & _
        "  p = " & nullText & "," & vbLf & _
        "  alternative = """ & s.alternative & """," & vbLf & _
        "  conf.level = " & confidence & vbLf & ")"
End Function

Private Function OneSampleSubject(ByRef s As TOneSample) As String
    Dim m As clsMatch, words As String
    If s.parameter <> "proportion" Then OneSampleSubject = s.column: Exit Function
    If Len(s.condition) > 0 Then
        Set m = RxMatch(s.condition, "^\s*([A-Za-z.][A-Za-z0-9._]*)\s*(>=|<=|==|!=|>|<)\s*([-+]?(?:\d+\.?\d*|\.\d+))\s*$")
        If Not m Is Nothing Then
            Select Case m.Group(2)
                Case ">": words = "greater than"
                Case ">=": words = "at least"
                Case "<": words = "less than"
                Case "<=": words = "at most"
                Case "==": words = "equal to"
                Case "!=": words = "not equal to"
            End Select
            OneSampleSubject = s.population & " with " & m.Group(1) & " " & words & " " & FormatFinalNumber(JsNumber(m.Group(3)))
            Exit Function
        End If
        OneSampleSubject = s.population & " satisfying " & s.condition
        Exit Function
    End If
    OneSampleSubject = s.population & " where " & s.column & " equals " & s.success
End Function

Public Sub SyncOneSampleForm()
    Dim s As TOneSample, proportion As Boolean, h0 As String, h1 As String
    s = OneSampleSettings()
    proportion = (s.parameter = "proportion")
    Q("#oneSampleSuccessField").ToggleClass "hidden", Not proportion
    Q("#oneSampleConditionField").ToggleClass "hidden", Not proportion
    Q("#oneSamplePopulationField").ToggleClass "hidden", Not proportion
    Q("#oneSampleSuccessField").ToggleClass "inapplicable", proportion And Len(s.condition) > 0
    Q("#oneSampleSuccess").Disabled = proportion And Len(s.condition) > 0
    Q("#oneSampleNull").Placeholder = IIf(proportion, "0.5", "50")
    OneSampleHypotheses s, h0, h1
    Q("#oneSampleHypotheses").InnerHtml = HypothesisHtml(h0, h1)
    Q("#oneSampleCode").TextContent = BuildOneSampleCode(s)
    RenderOneSampleInterpretation
End Sub

' parseOneSampleOutput(text): a clsMap with "error", or with "operator", "pValue" (Double),
' "statisticLabel" and "statisticValue" (Strings, Null when absent).
Public Function ParseOneSampleOutput(ByVal text As String) As clsMap
    Dim flat As String, pMatch As clsMatch, statMatch As clsMatch, pValue As Double, result As clsMap
    Set result = NewMap()
    flat = RxReplace(text, "^\s{0,3}#{1,3}\s{0,3}", "", "gm")
    flat = JsTrim(RxReplace(flat, "\s+", " ", "g"))
    Set pMatch = RxMatch(flat, "p(?:-|_|\s{0,3})value\s{0,3}(<=|>=|<|>|=)\s{0,3}" & NumberPattern(), "i")
    If pMatch Is Nothing Then
        result("error") = "I could not find a p-value in the pasted test output."
        Set ParseOneSampleOutput = result
        Exit Function
    End If
    pValue = JsNumber(pMatch.Group(2))
    Set statMatch = RxMatch(flat, "\b(t|z|X-squared)\s{0,3}=\s{0,3}" & NumberPattern(), "i")
    If JsIsFinite(pValue) Then
        result("operator") = pMatch.Group(1)
        result("pValue") = pValue
        If statMatch Is Nothing Then
            result("statisticLabel") = Null
            result("statisticValue") = Null
        Else
            result("statisticLabel") = statMatch.Group(1)
            result("statisticValue") = statMatch.Group(2)
        End If
    Else
        result("error") = "The detected p-value is not valid."
    End If
    Set ParseOneSampleOutput = result
End Function

Private Function PDisplay(ByVal operator As String, ByVal pValue As Double, ByVal alpha As Double) As String
    If operator = "=" Then PDisplay = FormatPValue(pValue, alpha) Else PDisplay = operator & " " & FormatPValue(pValue, alpha)
End Function

Private Function StatisticPhrase(ByVal label As Variant, ByVal value As Variant) As String
    If IsNull(label) Or IsNull(value) Then Exit Function
    If Len(label) = 0 Or Len(value) = 0 Then Exit Function
    StatisticPhrase = label & " = " & FormatFinalNumber(JsNumber(CStr(value))) & " and "
End Function

Public Sub RenderOneSampleInterpretation()
    Dim text As String, result As clsMap, s As TOneSample, reject As Variant, box As clsNode
    Dim operator As String, pValue As Double, comparison As String, decision As String, evidence As String
    Dim parameter As String, conclusion As String, h0 As String, h1 As String, nullText As String, prose As String
    Set box = Q("#oneSampleInterpretation")
    text = JsTrim(Q("#oneSampleOutputText").Value)
    If Len(text) = 0 Then
        box.InnerHtml = EmptyInterpretationHtml()
        Exit Sub
    End If
    Set result = ParseOneSampleOutput(text)
    If result.Exists("error") Then
        box.InnerHtml = "<div class=""prose-card""><h3>Check the pasted output</h3><p class=""waiting-note"">" & EscapeHtml(result("error")) & "</p></div>"
        Exit Sub
    End If
    s = OneSampleSettings()
    operator = result("operator")
    pValue = result("pValue")
    reject = SafeDecision(operator, pValue, s.alpha)
    If IsNull(reject) Then
        box.InnerHtml = "<div class=""prose-card""><p class=""waiting-note"">" & EscapeHtml(PValueDecisionError(operator, pValue, s.alpha)) & "</p></div>"
        Exit Sub
    End If
    comparison = PValueComparison(operator, pValue, s.alpha)
    If reject Then decision = "reject": evidence = "sufficient" Else decision = "fail to reject": evidence = "insufficient"
    If s.parameter = "mean" Then parameter = "population mean" Else parameter = "population proportion"
    nullText = JsNumberToString(s.nullValue)
    If s.alternative = "greater" Then
        conclusion = "is greater than " & nullText
    ElseIf s.alternative = "less" Then
        conclusion = "is less than " & nullText
    Else
        conclusion = "is different from " & nullText
    End If
    OneSampleHypotheses s, h0, h1
    prose = h0 & vbLf & vbLf & h1 & vbLf & vbLf & "Since " & StatisticPhrase(result("statisticLabel"), result("statisticValue")) & _
        "the p-value is " & PDisplay(operator, pValue, s.alpha) & ", which is " & comparison & " " & JsNumberToString(s.alpha) & _
        ", we " & decision & " the null hypothesis at the " & PercentLabel(s.alpha) & "% significance level. There is " & evidence & _
        " evidence to conclude that the " & parameter & " of " & OneSampleSubject(s) & " " & conclusion & "."
    mOneSampleProse = prose
    box.InnerHtml = "<div class=""prose-card""><h3>Generated interpretation</h3><p id=""oneSampleGeneratedProse"">" & EscapeHtml(prose) & _
        "</p><button class=""sample-btn inline-copy"" id=""copyOneSampleProseBtn"" type=""button"">Copy prose</button></div>"
    Q("#copyOneSampleProseBtn").AddEventListener "click", "opInference.OnCopyOneSampleProse"
End Sub

' ---------------------------------------------------------------- two-sample tests

Private Function TwoSampleSettings() As TTwoSample
    Dim s As TTwoSample
    s.parameter = Q("#twoSampleParameter").Value
    If s.parameter = "mean" And Q("#twoSamplePaired").Checked Then s.design = "paired" Else s.design = "independent"
    If s.parameter = "mean" Then s.alternative = Q("#twoSampleAlternative").Value Else s.alternative = "two.sided"
    s.dataset = OrDefault(FieldText("#twoSampleDataset"), "dataset")
    s.response1 = OrDefault(FieldText("#twoSampleResponse1"), "column1")
    s.response2 = OrDefault(FieldText("#twoSampleResponse2"), "column2")
    s.alpha = 0.05
    TwoSampleSettings = s
End Function

Private Sub TwoSampleHypotheses(ByRef s As TTwoSample, ByRef h0 As String, ByRef h1 As String)
    Dim symbol As String, one As String, two As String
    If s.parameter = "mean" Then symbol = SYM_MU() Else symbol = SYM_SIGMA() & SYM_SQUARED()
    If s.design = "paired" And s.parameter = "mean" Then
        If s.alternative = "greater" Then
            h0 = SYM_H0() & ": " & SYM_MU() & "d " & SYM_LE() & " 0": h1 = SYM_H1() & ": " & SYM_MU() & "d > 0"
        ElseIf s.alternative = "less" Then
            h0 = SYM_H0() & ": " & SYM_MU() & "d " & SYM_GE() & " 0": h1 = SYM_H1() & ": " & SYM_MU() & "d < 0"
        Else
            h0 = SYM_H0() & ": " & SYM_MU() & "d = 0": h1 = SYM_H1() & ": " & SYM_MU() & "d " & SYM_NE() & " 0"
        End If
        Exit Sub
    End If
    one = symbol & SYM_SUB1()
    two = symbol & SYM_SUB2()
    If s.alternative = "greater" Then
        h0 = SYM_H0() & ": " & one & " " & SYM_LE() & " " & two: h1 = SYM_H1() & ": " & one & " > " & two
    ElseIf s.alternative = "less" Then
        h0 = SYM_H0() & ": " & one & " " & SYM_GE() & " " & two: h1 = SYM_H1() & ": " & one & " < " & two
    Else
        h0 = SYM_H0() & ": " & one & " = " & two: h1 = SYM_H1() & ": " & one & " " & SYM_NE() & " " & two
    End If
End Sub

Private Function BuildTwoSampleCode(ByRef s As TTwoSample) As String
    Dim data As String, r1 As String, r2 As String, columns As String
    data = RColumnName(s.dataset)
    r1 = RColumnName(s.response1)
    r2 = RColumnName(s.response2)
    If s.parameter = "variance" Then
        BuildTwoSampleCode = "var.test(" & r1 & " ~ " & r2 & ", data = " & data & ")"
        Exit Function
    End If
    If s.design = "paired" Then columns = data & "$" & r1 & ", " & data & "$" & r2 Else columns = data & "$" & r1 & " ~ " & data & "$" & r2
    BuildTwoSampleCode = "t.test(" & columns & "," & vbLf & "       alternative = """ & s.alternative & """," & _
        IIf(s.design = "paired", vbLf & "       paired = TRUE,", "") & vbLf & "       var.equal = FALSE," & vbLf & "       conf.level = 0.95)"
End Function

Public Sub SyncTwoSampleForm()
    Dim s As TTwoSample, variance As Boolean, paired As Boolean, h0 As String, h1 As String
    s = TwoSampleSettings()
    variance = (s.parameter = "variance")
    paired = (s.design = "paired")
    Q("#twoSamplePairedField").ToggleClass "hidden", variance
    Q("#twoSampleAlternativeField").ToggleClass "hidden", variance
    If variance Then
        Q("#twoSampleTestTitle").TextContent = "Fisher's F test"
    ElseIf paired Then
        Q("#twoSampleTestTitle").TextContent = "Paired two-sample t-test"
    Else
        Q("#twoSampleTestTitle").TextContent = "Independent two-sample t-test"
    End If
    Q("#twoResponse1Label").FirstChild.TextContent = IIf(paired, "Column 1 (first response)", "Column 1 (response)")
    Q("#twoResponse2Field").FirstChild.TextContent = IIf(paired, "Column 2 (second response)", "Column 2 (grouping)")
    TwoSampleHypotheses s, h0, h1
    Q("#twoSampleHypotheses").InnerHtml = HypothesisHtml(h0, h1)
    Q("#twoSampleCode").TextContent = BuildTwoSampleCode(s)
    RenderTwoSampleInterpretation
End Sub

' parseTwoSampleOutput(text): a clsMap with "error", or with "operator", "pValue" (Double,
' possibly not finite), "statLabel" and "statValue" (Strings, Null when absent).
Public Function ParseTwoSampleOutput(ByVal text As String) As clsMap
    Dim flat As String, p As clsMatch, stat As clsMatch, result As clsMap
    Set result = NewMap()
    flat = RxReplace(text, "^\s{0,3}#{1,3}\s{0,3}", "", "gm")
    flat = JsTrim(RxReplace(flat, "\s+", " ", "g"))
    Set p = RxMatch(flat, "p(?:-|\s{0,3})value\s{0,3}(<=|>=|<|>|=)\s{0,3}" & NumberPattern(), "i")
    Set stat = RxMatch(flat, "\b(t|F)\s{0,3}=\s{0,3}" & NumberPattern(), "i")
    If p Is Nothing Then
        result("error") = "I could not find a p-value in the pasted output."
        Set ParseTwoSampleOutput = result
        Exit Function
    End If
    result("operator") = p.Group(1)
    result("pValue") = JsNumber(p.Group(2))
    If stat Is Nothing Then
        result("statLabel") = Null
        result("statValue") = Null
    Else
        result("statLabel") = stat.Group(1)
        result("statValue") = stat.Group(2)
    End If
    Set ParseTwoSampleOutput = result
End Function

Public Sub RenderTwoSampleInterpretation()
    Dim text As String, box As clsNode, result As clsMap, s As TTwoSample, reject As Variant, operator As String, pValue As Double
    Dim decision As String, evidence As String, comparison As String, h0 As String, h1 As String, relation As String, claim As String, prose As String
    text = JsTrim(Q("#twoSampleOutputText").Value)
    Set box = Q("#twoSampleInterpretation")
    If Len(text) = 0 Then box.InnerHtml = EmptyInterpretationHtml(): Exit Sub
    Set result = ParseTwoSampleOutput(text)
    If result.Exists("error") Then
        box.InnerHtml = "<div class=""prose-card""><h3>Check the pasted output</h3><p class=""waiting-note"">" & EscapeHtml(result("error")) & "</p></div>"
        Exit Sub
    End If
    s = TwoSampleSettings()
    operator = result("operator")
    pValue = result("pValue")
    reject = SafeDecision(operator, pValue, s.alpha)
    If IsNull(reject) Then
        box.InnerHtml = "<div class=""prose-card""><p class=""waiting-note"">" & EscapeHtml(PValueDecisionError(operator, pValue, s.alpha)) & "</p></div>"
        Exit Sub
    End If
    If reject Then decision = "reject": evidence = "sufficient" Else decision = "fail to reject": evidence = "insufficient"
    comparison = PValueComparison(operator, pValue, s.alpha)
    TwoSampleHypotheses s, h0, h1
    If s.alternative = "greater" Then
        relation = "greater than"
    ElseIf s.alternative = "less" Then
        relation = "less than"
    Else
        relation = "different from"
    End If
    If s.parameter = "variance" Then
        claim = "the population variances of " & s.response1 & " differ between the two groups in " & s.response2
    ElseIf s.design = "paired" Then
        claim = "the population mean paired difference (" & s.response1 & " minus " & s.response2 & ") is " & relation & " 0"
    Else
        claim = "the population mean of " & s.response1 & " for group 1 is " & relation & " that for group 2 in " & s.response2
    End If
    prose = h0 & vbLf & vbLf & h1 & vbLf & vbLf & "Since " & StatisticPhrase(result("statLabel"), result("statValue")) & _
        "the p-value is " & PDisplay(operator, pValue, s.alpha) & ", which is " & comparison & " " & JsNumberToString(s.alpha) & _
        ", we " & decision & " the null hypothesis at the " & PercentLabel(s.alpha) & "% significance level. There is " & evidence & _
        " evidence to conclude that " & claim & "."
    mTwoSampleProse = prose
    box.InnerHtml = "<div class=""prose-card""><h3>Generated interpretation</h3><p id=""twoSampleGeneratedProse"">" & EscapeHtml(prose) & _
        "</p><button class=""sample-btn inline-copy"" id=""copyTwoSampleProseBtn"">Copy prose</button></div>"
    Q("#copyTwoSampleProseBtn").AddEventListener "click", "opInference.OnCopyTwoSampleProse"
End Sub

' ---------------------------------------------------------------- event handlers

Private Sub CopyWithToast(ByVal text As String, ByVal toast As String)
    If WriteAnswerText(text) Then ShowToast toast Else ShowToast "Clipboard access is unavailable"
End Sub

Public Sub OnTwoSampleSync(ByVal e As clsEvent)
    On Error GoTo Failed
    SyncTwoSampleForm
    Exit Sub
Failed:
    ReportError "opInference.OnTwoSampleSync"
End Sub

Public Sub OnTwoSampleOutputInput(ByVal e As clsEvent)
    On Error GoTo Failed
    RenderTwoSampleInterpretation
    Exit Sub
Failed:
    ReportError "opInference.OnTwoSampleOutputInput"
End Sub

Public Sub OnLoadTwoSampleOutput(ByVal e As clsEvent)
    Dim s As TTwoSample, variance As Boolean, paired As Boolean, heading As String, columns As String, meanPValue As String, statistics As String
    On Error GoTo Failed
    s = TwoSampleSettings()
    variance = (s.parameter = "variance")
    paired = (s.design = "paired")
    If variance Then
        heading = "F test to compare two variances"
    ElseIf paired Then
        heading = "Paired t-test"
    Else
        heading = "Welch Two Sample t-test"
    End If
    If paired Then columns = s.response1 & " and " & s.response2 Else columns = s.response1 & " by " & s.response2
    If s.alternative = "greater" Then
        meanPValue = "3.845e-09"
    ElseIf s.alternative = "less" Then
        meanPValue = "1"
    Else
        meanPValue = "7.69e-09"
    End If
    If variance Then statistics = "F = 1.42, num df = 29, denom df = 31, p-value = 0.326" Else statistics = "t = 8.03, df = 29, p-value = " & meanPValue
    Q("#twoSampleOutputText").Value = heading & vbLf & vbLf & "data: " & columns & vbLf & statistics
    RenderTwoSampleInterpretation
    ShowToast "Sample output loaded"
    Exit Sub
Failed:
    ReportError "opInference.OnLoadTwoSampleOutput"
End Sub

' navigator.clipboard.readText() is unavailable in the shell's test mode (like the reference
' harness); otherwise ClipboardRead cannot fail (it returns "" instead).
Private Function ClipboardReadUnavailable() As Boolean
    On Error GoTo NoShell
    ClipboardReadUnavailable = CBool(CallOp("AppTestMode"))
    Exit Function
NoShell:
    ClipboardReadUnavailable = TestPasteUnavailable
End Function

Public Sub OnPasteTwoSample(ByVal e As clsEvent)
    On Error GoTo Failed
    If ClipboardReadUnavailable() Then ShowToast "Use Ctrl/Cmd + V to paste": Exit Sub
    Q("#twoSampleOutputText").Value = ClipboardRead()
    RenderTwoSampleInterpretation
    ShowToast "Clipboard pasted"
    Exit Sub
Failed:
    ReportError "opInference.OnPasteTwoSample"
    ShowToast "Use Ctrl/Cmd + V to paste"
End Sub

Public Sub OnCopyTwoSampleCode(ByVal e As clsEvent)
    On Error GoTo Failed
    CopyWithToast Q("#twoSampleCode").TextContent, "R code copied"
    Exit Sub
Failed:
    ReportError "opInference.OnCopyTwoSampleCode"
End Sub

Public Sub OnCopyTwoSampleProse(ByVal e As clsEvent)
    On Error GoTo Failed
    CopyWithToast mTwoSampleProse, "Prose copied"
    Exit Sub
Failed:
    ReportError "opInference.OnCopyTwoSampleProse"
End Sub

Public Sub OnOneSampleSync(ByVal e As clsEvent)
    On Error GoTo Failed
    SyncOneSampleForm
    Exit Sub
Failed:
    ReportError "opInference.OnOneSampleSync"
End Sub

Public Sub OnOneSampleParameterChange(ByVal e As clsEvent)
    Dim current As String
    On Error GoTo Failed
    current = Q("#oneSampleNull").Value
    If Q("#oneSampleParameter").Value = "proportion" And current = "50" Then Q("#oneSampleNull").Value = "0.5"
    If Q("#oneSampleParameter").Value = "mean" And current = "0.5" Then Q("#oneSampleNull").Value = "50"
    SyncOneSampleForm
    Exit Sub
Failed:
    ReportError "opInference.OnOneSampleParameterChange"
End Sub

Public Sub OnOneSampleOutputInput(ByVal e As clsEvent)
    On Error GoTo Failed
    RenderOneSampleInterpretation
    Exit Sub
Failed:
    ReportError "opInference.OnOneSampleOutputInput"
End Sub

Public Sub OnLoadOneSampleSetup(ByVal e As clsEvent)
    Dim proportion As Boolean
    On Error GoTo Failed
    proportion = (Q("#oneSampleParameter").Value = "proportion")
    Q("#oneSampleDataset").Value = "dataset"
    Q("#oneSampleColumn").Value = "column1"
    Q("#oneSampleNull").Value = IIf(proportion, "0.5", "50")
    Q("#oneSampleAlpha").Value = "0.05"
    Q("#oneSampleFilter").Value = IIf(proportion, "Product == ""Book""", "")
    Q("#oneSampleExcludeMissing").Checked = False
    Q("#oneSamplePopulation").Value = IIf(proportion, "book sales transactions", "observations")
    Q("#oneSampleCondition").Value = IIf(proportion, "Amount > 50", "")
    Q("#oneSampleSuccess").Value = "yes"
    If proportion Then
        Q("#oneSampleDataset").Value = "ST"
        Q("#oneSampleColumn").Value = "Amount"
        Q("#oneSampleNull").Value = "0.25"
        Q("#oneSampleAlternative").Value = "less"
    End If
    SyncOneSampleForm
    ShowToast "Sample setup loaded"
    Exit Sub
Failed:
    ReportError "opInference.OnLoadOneSampleSetup"
End Sub

Public Sub OnLoadOneSampleOutput(ByVal e As clsEvent)
    Dim s As TOneSample, p As String, relation As String, nullText As String
    On Error GoTo Failed
    s = OneSampleSettings()
    If s.alternative = "two.sided" Then p = "0.0716" Else p = "0.018"
    If s.alternative = "two.sided" Then relation = "not equal to" Else relation = s.alternative
    nullText = JsNumberToString(s.nullValue)
    If s.parameter = "mean" Then
        Q("#oneSampleOutputText").Value = "One Sample t-test" & vbLf & vbLf & "data: " & s.dataset & "$" & s.column & vbLf & _
            "t = -1.85, df = 39, p-value = " & p & vbLf & "alternative hypothesis: true mean is " & relation & " " & nullText & vbLf & _
            "95 percent confidence interval:" & vbLf & "48.4 50.2" & vbLf & "sample estimates:" & vbLf & "mean of x" & vbLf & "49.3"
    ElseIf Len(s.condition) > 0 Then
        Q("#oneSampleOutputText").Value = "z = -1.75" & vbLf & "p-value = 0.04"
    Else
        Q("#oneSampleOutputText").Value = "1-sample proportions test with continuity correction" & vbLf & vbLf & "data: 25 out of 40, null probability " & nullText & vbLf & _
            "X-squared = 5.12, df = 1, p-value = " & p & vbLf & "alternative hypothesis: true p is " & relation & " " & nullText & vbLf & _
            "95 percent confidence interval:" & vbLf & "0.46 0.77" & vbLf & "sample estimates:" & vbLf & "p" & vbLf & "0.625"
    End If
    RenderOneSampleInterpretation
    ShowToast "Sample output loaded"
    Exit Sub
Failed:
    ReportError "opInference.OnLoadOneSampleOutput"
End Sub

Public Sub OnPasteOneSample(ByVal e As clsEvent)
    On Error GoTo Failed
    If ClipboardReadUnavailable() Then ShowToast "Use Ctrl/Cmd + V to paste": Exit Sub
    Q("#oneSampleOutputText").Value = ClipboardRead()
    RenderOneSampleInterpretation
    ShowToast "Clipboard pasted"
    Exit Sub
Failed:
    ReportError "opInference.OnPasteOneSample"
    ShowToast "Use Ctrl/Cmd + V to paste"
End Sub

Public Sub OnCopyOneSampleCode(ByVal e As clsEvent)
    On Error GoTo Failed
    CopyWithToast Q("#oneSampleCode").TextContent, "R code copied"
    Exit Sub
Failed:
    ReportError "opInference.OnCopyOneSampleCode"
End Sub

Public Sub OnCopyOneSampleProse(ByVal e As clsEvent)
    On Error GoTo Failed
    CopyWithToast mOneSampleProse, "Prose copied"
    Exit Sub
Failed:
    ReportError "opInference.OnCopyOneSampleProse"
End Sub
