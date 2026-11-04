Attribute VB_Name = "opAnova"
' The staged ANOVA workflow (Step 0 setup, Step 1 normality, Step 2 variance test and
' routing, Step 3a standard ANOVA + Tukey HSD, Step 3b Welch ANOVA + Games-Howell, staged
' copy completion) and the multiple-correlations code generator.
' Ported from app-src.js lines 1864-2057 and 4516-4562.
Option Explicit

Private Type TAnovaSettings
    dataset As String
    response As String
    group As String
    groupCount As Double   ' may be NaN
    alpha As Double        ' may be NaN
    normalityAssumed As Boolean
End Type

' anovaNormalityWasAssumed: Null (not yet rendered) / True / False.
Private mNormalityWasAssumed As Variant
' anovaVarianceRoute: "" (null), "standard" or "welch".
Private mVarianceRoute As String
Private mVarianceProse As String, mStandardProse As String, mWelchProse As String
Private mLastVarianceCode As String, mLastStandardCode As String, mLastWelchCode As String
Private mVarianceCopyCode As Boolean, mVarianceCopyProse As Boolean
Private mStandardCopyCode As Boolean, mStandardCopyProse As Boolean
Private mWelchCopyCode As Boolean, mWelchCopyProse As Boolean

Private Const NUMBER_PATTERN As String = "(?:<=|>=|<|>)?\s*-?[0-9]*\.?[0-9]+(?:e[+-]?\d+)?"

' ---------------------------------------------------------------- startup

' Website lines 1869-1871 (state) and 4516-4562 (listeners). Call after the shell has
' made the #anovaBench panels collapsible (initialiseLpCollapsibles).
Public Sub InitOpAnova()
    Dim list As clsList, i As Long
    mNormalityWasAssumed = Null
    mVarianceRoute = ""
    mVarianceProse = "": mStandardProse = "": mWelchProse = ""
    mLastVarianceCode = "": mLastStandardCode = "": mLastWelchCode = ""
    mVarianceCopyCode = False: mVarianceCopyProse = False
    mStandardCopyCode = False: mStandardCopyProse = False
    mWelchCopyCode = False: mWelchCopyProse = False

    Set list = QA(".corr-multiple-input")
    For i = 0 To list.Count - 1
        list(i).AddEventListener "input", "opAnova.OnCorrMultipleInput"
    Next i
    Q("#loadCorrMultipleSampleBtn").AddEventListener "click", "opAnova.OnLoadCorrMultipleSample"
    Q("#copyCorrMultipleCodeBtn").AddEventListener "click", "opAnova.OnCopyCorrMultipleCode"

    Q("#anovaNormalityAssumed").AddEventListener "change", "opAnova.OnAnovaSettingsChange"
    Set list = QA(".anova-input")
    For i = 0 To list.Count - 1
        list(i).AddEventListener "input", "opAnova.OnAnovaSettingsChange"
    Next i
    Q("#anovaVarianceOutput").AddEventListener "input", "opAnova.OnAnovaVarianceOutputInput"
    Q("#loadAnovaSampleBtn").AddEventListener "click", "opAnova.OnLoadAnovaSample"
    Q("#loadAnovaVarianceOutputBtn").AddEventListener "click", "opAnova.OnLoadAnovaVarianceOutput"
    Q("#copyAnovaVarianceProseBtn").AddEventListener "click", "opAnova.OnCopyAnovaVarianceProse"
    Q("#copyAnovaNormalityCodeBtn").AddEventListener "click", "opAnova.OnCopyAnovaCode"
    Q("#copyAnovaVarianceCodeBtn").AddEventListener "click", "opAnova.OnCopyAnovaCode"
    Q("#copyAnovaStandardCodeBtn").AddEventListener "click", "opAnova.OnCopyAnovaCode"
    Q("#copyAnovaWelchCodeBtn").AddEventListener "click", "opAnova.OnCopyAnovaCode"
    Q("#anovaVarianceOutput").AddEventListener "input", "opAnova.OnAnovaVarianceOutputReset"
    Q("#loadAnovaVarianceOutputBtn").AddEventListener "click", "opAnova.OnAnovaVarianceOutputReset"
    Q("#anovaStandardOutput").AddEventListener "input", "opAnova.OnAnovaStandardOutputInput"
    Q("#loadAnovaStandardOutputBtn").AddEventListener "click", "opAnova.OnLoadAnovaStandardOutput"
    Q("#copyAnovaStandardProseBtn").AddEventListener "click", "opAnova.OnCopyAnovaStandardProse"
    Q("#anovaWelchOutput").AddEventListener "input", "opAnova.OnAnovaWelchOutputInput"
    Q("#anovaGamesHowellOutput").AddEventListener "input", "opAnova.OnAnovaWelchOutputInput"
    Q("#loadAnovaWelchOutputsBtn").AddEventListener "click", "opAnova.OnLoadAnovaWelchOutputs"
    Q("#copyAnovaWelchProseBtn").AddEventListener "click", "opAnova.OnCopyAnovaWelchProse"
End Sub

' ---------------------------------------------------------------- helpers

Private Function U(ByVal text As String) As String
    U = JsU(text)
End Function

Private Function IsValidAlpha(ByVal alpha As Double) As Boolean
    If JsIsFinite(alpha) Then IsValidAlpha = (alpha > 0 And alpha < 1)
End Function

Private Function IsValidGroupCount(ByVal count As Double) As Boolean
    If JsIsInteger(count) Then IsValidGroupCount = (count >= 2)
End Function

Private Function IsTwoGroups(ByVal count As Double) As Boolean
    IsTwoGroups = JsEq(count, 2)
End Function

Private Function ValueOrDefault(ByVal selector As String, ByVal fallback As String) As String
    ValueOrDefault = JsTrim(Q(selector).Value)
    If Len(ValueOrDefault) = 0 Then ValueOrDefault = fallback
End Function

Private Function AnovaBlockSettings() As TAnovaSettings
    Dim s As TAnovaSettings
    s.dataset = ValueOrDefault("#anovaDataset", "dataset")
    s.response = ValueOrDefault("#anovaResponse", "column1")
    s.group = ValueOrDefault("#anovaGroup", "column2")
    s.groupCount = JsNumber(Q("#anovaGroupCount").Value)
    s.alpha = JsNumber(Q("#anovaAlpha").Value)
    s.normalityAssumed = Q("#anovaNormalityAssumed").Checked
    AnovaBlockSettings = s
End Function

' Values of every match of pattern in text (String.prototype.match with the g flag).
Private Function MatchValues(ByVal text As String, ByVal pattern As String, ByVal flags As String) As clsList
    Dim out As clsList, m As Variant, found As Collection
    Set out = NewList()
    Set found = Rx(pattern, flags).MatchAll(text)
    For Each m In found
        out.Push m.Value
    Next m
    Set MatchValues = out
End Function

' text.split(/\r?\n/).map(line => line.replace(/^\s*##\s?/, "").trim())
Private Function AnovaOutputRows(ByVal text As String) As clsList
    Dim lines As clsList, out As clsList, i As Long
    Set lines = Rx("\r?\n").Split(text)
    Set out = NewList()
    For i = 0 To lines.Count - 1
        out.Push JsTrim(Rx("^\s*##\s?").Replace(lines(i), ""))
    Next i
    Set AnovaOutputRows = out
End Function

' ---------------------------------------------------------------- code builders

Private Function BuildAnovaNormalityCode(ByRef s As TAnovaSettings) As String
    BuildAnovaNormalityCode = RColumnName(s.dataset) & " %>%" & vbLf & "  group_by(" & RColumnName(s.group) & ") %>%" & vbLf & _
        "  shapiro_test(" & RColumnName(s.response) & ")"
End Function

Private Function BuildAnovaVarianceCode(ByRef s As TAnovaSettings) As String
    Dim data As String, response As String, group As String
    data = RColumnName(s.dataset): response = RColumnName(s.response): group = RColumnName(s.group)
    If IsTwoGroups(s.groupCount) Then
        BuildAnovaVarianceCode = "var.test(" & data & "$" & response & " ~ " & data & "$" & group & ")"
    Else
        BuildAnovaVarianceCode = "bartlett.test(" & response & " ~ " & group & ", data = " & data & ")"
    End If
End Function

Private Function BuildStandardAnovaCode(ByRef s As TAnovaSettings) As String
    BuildStandardAnovaCode = "model <- aov(" & RColumnName(s.response) & " ~ " & RColumnName(s.group) & ", data = " & RColumnName(s.dataset) & ")" & vbLf & _
        "summary(model)" & vbLf & "TukeyHSD(model)"
End Function

Private Function BuildWelchAnovaCode(ByRef s As TAnovaSettings) As String
    Dim data As String, response As String, group As String
    data = RColumnName(s.dataset): response = RColumnName(s.response): group = RColumnName(s.group)
    BuildWelchAnovaCode = data & " %>% welch_anova_test(" & response & " ~ " & group & ")" & vbLf & _
        data & " %>% games_howell_test(" & response & " ~ " & group & ") %>% knitr::kable()"
End Function

' ---------------------------------------------------------------- staged copy state

Private Sub RemoveCopyComplete(ByVal selector As String)
    Dim node As clsNode
    Set node = Q(selector)
    If Not node Is Nothing Then node.RemoveClass "is-copy-complete"
End Sub

Public Sub ResetAnovaVarianceCopyState(Optional ByVal part As String = "all")
    If part = "all" Or part = "code" Then
        mVarianceCopyCode = False
        RemoveCopyComplete "#anovaVarianceCodeBox"
    End If
    If part = "all" Or part = "prose" Then
        mVarianceCopyProse = False
        RemoveCopyComplete "#anovaVarianceProseBox"
    End If
    SetPanelCompletion Q("#anovaVariancePanel"), False
End Sub

Public Sub ResetAnovaBranchCopyState(ByVal branch As String, Optional ByVal part As String = "all")
    Dim prefix As String
    If branch = "standard" Then
        prefix = "#anovaStandard"
        If part = "all" Or part = "code" Then mStandardCopyCode = False
        If part = "all" Or part = "prose" Then mStandardCopyProse = False
    ElseIf branch = "welch" Then
        prefix = "#anovaWelch"
        If part = "all" Or part = "code" Then mWelchCopyCode = False
        If part = "all" Or part = "prose" Then mWelchCopyProse = False
    Else
        Exit Sub
    End If
    If part = "all" Or part = "code" Then RemoveCopyComplete prefix & "CodeBox"
    If part = "all" Or part = "prose" Then RemoveCopyComplete prefix & "ProseBox"
    SetPanelCompletion Q(prefix & "Panel"), False
End Sub

' completeAnovaVarianceCopy: records a staged copy. True when the button belongs to a
' staged ANOVA step (the caller must then not complete the panel itself).
Public Function CompleteAnovaVarianceCopy(ByVal button As Variant) As Boolean
    Dim id As String, part As String, stage As String, complete As Boolean, panel As clsNode
    If Not IsObject(button) Then Exit Function
    If button Is Nothing Then Exit Function
    id = button.Id
    Select Case id
        Case "copyAnovaVarianceCodeBtn": stage = "Variance": part = "Code"
        Case "copyAnovaVarianceProseBtn": stage = "Variance": part = "Prose"
        Case "copyAnovaStandardCodeBtn": stage = "Standard": part = "Code"
        Case "copyAnovaStandardProseBtn": stage = "Standard": part = "Prose"
        Case "copyAnovaWelchCodeBtn": stage = "Welch": part = "Code"
        Case "copyAnovaWelchProseBtn": stage = "Welch": part = "Prose"
        Case Else: Exit Function
    End Select
    Select Case stage & part
        Case "VarianceCode": mVarianceCopyCode = True
        Case "VarianceProse": mVarianceCopyProse = True
        Case "StandardCode": mStandardCopyCode = True
        Case "StandardProse": mStandardCopyProse = True
        Case "WelchCode": mWelchCopyCode = True
        Case "WelchProse": mWelchCopyProse = True
    End Select
    If Not Q("#anova" & stage & part & "Box") Is Nothing Then Q("#anova" & stage & part & "Box").AddClass "is-copy-complete"
    Select Case stage
        Case "Variance": complete = mVarianceCopyCode And mVarianceCopyProse
        Case "Standard": complete = mStandardCopyCode And mStandardCopyProse
        Case "Welch": complete = mWelchCopyCode And mWelchCopyProse
    End Select
    If complete Then
        Set panel = Q("#anova" & stage & "Panel")
        SetPanelCompletion panel, True
        SetLpPanelCollapsed panel, True
    End If
    CompleteAnovaVarianceCopy = True
End Function

' ---------------------------------------------------------------- Step 2: variance test

Private Sub ApplyAnovaVarianceRoute(ByVal route As String, ByVal validAlpha As Boolean)
    Dim k As Long, branchRoute As String, panel As clsNode, button As clsNode, toggle As clsNode, unavailable As Boolean
    For k = 0 To 1
        If k = 0 Then
            branchRoute = "standard": Set panel = Q("#anovaStandardPanel"): Set button = Q("#copyAnovaStandardCodeBtn")
        Else
            branchRoute = "welch": Set panel = Q("#anovaWelchPanel"): Set button = Q("#copyAnovaWelchCodeBtn")
        End If
        unavailable = (Len(route) > 0 And branchRoute <> route)
        Set toggle = panel.QuerySelector(".lp-collapse-toggle")
        panel.ToggleClass "is-unavailable", unavailable
        panel.SetAttribute "aria-disabled", IIf(unavailable, "true", "false")
        If Not toggle Is Nothing Then toggle.Disabled = unavailable
        button.Disabled = (Not validAlpha) Or unavailable
        If branchRoute = "standard" Then
            Q("#loadAnovaStandardOutputBtn").Disabled = unavailable
            Q("#copyAnovaStandardProseBtn").Disabled = unavailable Or Len(mStandardProse) = 0
        Else
            Q("#loadAnovaWelchOutputsBtn").Disabled = unavailable
            Q("#copyAnovaWelchProseBtn").Disabled = unavailable Or Len(mWelchProse) = 0
        End If
        If unavailable Then
            ResetAnovaBranchCopyState branchRoute, "all"
            SetLpPanelCollapsed panel, True
        ElseIf Len(route) > 0 And route <> mVarianceRoute Then
            SetLpPanelCollapsed panel, False
        End If
    Next k
    mVarianceRoute = route
End Sub

Private Sub ShowBoxError(ByVal box As clsNode, ByVal message As String)
    box.AddClass "is-error"
    box.TextContent = message
End Sub

Public Sub ParseAnovaVarianceOutput(ByVal text As String)
    Dim box As clsNode, copy As clsNode, s As TAnovaSettings, alpha As Double, validAlpha As Boolean, validGroups As Boolean
    Dim pMatch As clsMatch, operator As String, raw As String, p As Double, isBartlett As Boolean, isFTest As Boolean
    Dim reject As Variant, testName As String, scope As String, relation As String, decision As String, evidence As String
    Set box = Q("#anovaVarianceInterpretation")
    Set copy = Q("#copyAnovaVarianceProseBtn")
    s = AnovaBlockSettings()
    alpha = s.alpha
    validAlpha = IsValidAlpha(alpha)
    validGroups = IsValidGroupCount(s.groupCount)
    mVarianceProse = ""
    copy.Disabled = True
    If Not validGroups Then
        ShowBoxError box, "Enter a whole number of groups of at least 2."
        ApplyAnovaVarianceRoute "", validAlpha
        Exit Sub
    End If
    If Len(JsTrim(text)) = 0 Then
        box.RemoveClass "is-error"
        box.TextContent = "Paste the variance-test output to choose Step 3a or Step 3b."
        ApplyAnovaVarianceRoute "", validAlpha
        Exit Sub
    End If
    Set pMatch = Rx("p\s*-?\s*value\s*(<=|>=|<|>|=)\s*([0-9]*\.?[0-9]+(?:e[+-]?\d+)?)", "i").Exec(text)
    If pMatch Is Nothing Then
        ShowBoxError box, U("I could not find the p-value. Include the line containing \u2018p-value =\u2019 or \u2018p-value <\u2019.")
        ApplyAnovaVarianceRoute "", validAlpha
        Exit Sub
    End If
    If Not validAlpha Then
        ShowBoxError box, "Enter a significance level between 0 and 1 before interpreting this result."
        ApplyAnovaVarianceRoute "", False
        Exit Sub
    End If
    operator = pMatch.Group(1)
    raw = pMatch.Group(2)
    p = JsNumber(raw)
    isBartlett = RxTest(text, "bartlett", "i")
    isFTest = RxTest(text, "F\s+test\s+to\s+compare\s+two\s+variances|ratio\s+of\s+variances", "i")
    If Not JsIsFinite(p) Or (Not isBartlett And Not isFTest) Then
        ShowBoxError box, U("The p-value was found, but the output was not recognised as Bartlett\u2019s test or an F test comparing two variances.")
        ApplyAnovaVarianceRoute "", True
        Exit Sub
    End If
    If (IsTwoGroups(s.groupCount) And Not isFTest) Or (s.groupCount > 2 And Not isBartlett) Then
        If IsTwoGroups(s.groupCount) Then
            ShowBoxError box, "Two groups require output from var.test(). Paste the F test to compare two variances."
        Else
            ShowBoxError box, "Three or more groups require output from bartlett.test(). Paste the Bartlett test result."
        End If
        ApplyAnovaVarianceRoute "", True
        Exit Sub
    End If
    reject = PValueDecision(operator, p, alpha)
    If IsNull(reject) Then
        ShowBoxError box, "The printed bound p-value " & operator & " " & raw & " does not establish whether p is above or below " & JsNumberToString(alpha) & ". Paste a more precise p-value."
        ApplyAnovaVarianceRoute "", True
        Exit Sub
    End If
    If isBartlett Then
        testName = U("Bartlett\u2019s test")
        scope = "across the groups of " & s.group
    Else
        testName = "An F test"
        scope = "between the two groups of " & s.group
    End If
    relation = PValueComparison(operator, p, alpha)
    If reject Then
        decision = "reject"
        evidence = "There is evidence that the variances are not equal across groups, so the equal variance assumption is violated."
    Else
        decision = "fail to reject"
        evidence = "There is insufficient evidence that the variances differ across groups, so the equal variance assumption is not violated."
    End If
    mVarianceProse = testName & " was conducted to assess homogeneity of variances " & scope & "." & vbLf & _
        SYM_H0() & ": The population variances are equal across all groups." & vbLf & _
        SYM_H1() & ": At least one group has a different population variance." & vbLf & _
        "Since the p-value (" & operator & " " & raw & ") is " & relation & " " & JsNumberToString(alpha) & ", we " & decision & " " & SYM_H0() & ". " & evidence
    box.RemoveClass "is-error"
    box.TextContent = mVarianceProse
    copy.Disabled = False
    ApplyAnovaVarianceRoute IIf(reject, "welch", "standard"), True
End Sub

' ---------------------------------------------------------------- printed p-values

' parseAnovaPrintedP: clsMap {operator, raw, value} or Nothing.
Public Function ParseAnovaPrintedP(ByVal token As Variant) As clsMap
    Dim text As String, m As clsMatch, result As clsMap
    If Not IsEmpty(token) And Not IsNull(token) Then text = CStr(token)
    Set m = Rx("^(<=|>=|<|>|=)?\s*([0-9]*\.?[0-9]+(?:e[+-]?\d+)?)$", "i").Exec(JsTrim(text))
    If m Is Nothing Then Exit Function
    Set result = NewMap()
    If m.Matched(1) And Len(m.Group(1)) > 0 Then result("operator") = m.Group(1) Else result("operator") = "="
    result("raw") = m.Group(2)
    result("value") = JsNumber(m.Group(2))
    Set ParseAnovaPrintedP = result
End Function

Private Function AnovaPDecision(ByVal p As clsMap, ByVal alpha As Double) As Variant
    AnovaPDecision = PValueDecision(p("operator"), p("value"), alpha)
End Function

Private Function PDisplay(ByVal p As clsMap) As String
    PDisplay = p("operator") & " " & p("raw")
End Function

Private Function PercentText(ByVal alpha As Double) As String
    PercentText = JsToLocaleString(alpha * 100)
End Function

Private Function DfDisplay(ByVal value As Double) As String
    If JsIsInteger(value) Then DfDisplay = JsNumberToString(value) Else DfDisplay = FormatFinalNumber(value)
End Function

' ---------------------------------------------------------------- Step 3a: standard ANOVA

Public Sub RenderAnovaStandardInterpretation()
    Dim box As clsNode, copy As clsNode, text As String, s As TAnovaSettings, alpha As Double
    Dim rows As clsList, i As Long, line As String, resultLine As String, found As Boolean, numbers As clsList
    Dim residualLine As String, residualNumbers As clsList, df1 As Double, df2 As Double, f As Double, p As clsMap
    Dim reject As Variant, percent As String, base As String, tukeyHeader As Long, followUp As String
    Set box = Q("#anovaStandardInterpretation")
    Set copy = Q("#copyAnovaStandardProseBtn")
    text = Q("#anovaStandardOutput").Value
    s = AnovaBlockSettings()
    alpha = s.alpha
    mStandardProse = ""
    copy.Disabled = True
    If Len(JsTrim(text)) = 0 Then
        box.RemoveClass "is-error"
        box.TextContent = "Paste the standard ANOVA output to interpret the omnibus test and Tukey comparisons."
        Exit Sub
    End If
    If Not IsValidAlpha(alpha) Then
        ShowBoxError box, "Enter a significance level between 0 and 1 before interpreting the ANOVA result."
        Exit Sub
    End If
    If Not RxTest(text, "Pr\s*\(>\s*F\)", "i") Then
        ShowBoxError box, "Paste the ANOVA table from summary(model), including the F value and Pr(>F)."
        Exit Sub
    End If
    Set rows = AnovaOutputRows(text)
    For i = 0 To rows.Count - 1
        line = rows(i)
        If Len(line) > 0 Then
            If Not RxTest(line, "^(?:Df\b|Residuals\b|Signif\.|---|Tukey)", "i") Then
                If MatchValues(line, NUMBER_PATTERN, "gi").Count >= 5 Then
                    resultLine = line
                    found = True
                    Exit For
                End If
            End If
        End If
    Next i
    If Not found Then
        ShowBoxError box, "I found the ANOVA heading but could not read the model result row."
        Exit Sub
    End If
    Set numbers = MatchValues(resultLine, NUMBER_PATTERN, "gi")
    found = False
    For i = 0 To rows.Count - 1
        If RxTest(rows(i), "^Residuals\b", "i") Then
            residualLine = rows(i)
            found = True
            Exit For
        End If
    Next i
    If found Then Set residualNumbers = MatchValues(residualLine, "-?[0-9]*\.?[0-9]+(?:e[+-]?\d+)?", "gi") Else Set residualNumbers = NewList()
    df1 = JsNumber(numbers(numbers.Count - 5))
    If residualNumbers.Count > 0 Then df2 = JsNumber(residualNumbers(0)) Else df2 = JsNaN()
    Set p = ParseAnovaPrintedP(numbers(numbers.Count - 1))
    f = JsNumber(numbers(numbers.Count - 2))
    If p Is Nothing Or Not JsIsFinite(df1) Or Not JsIsFinite(df2) Or Not JsIsFinite(f) Then
        ShowBoxError box, "The ANOVA degrees of freedom, F-statistic, or p-value could not be read."
        Exit Sub
    End If
    reject = AnovaPDecision(p, alpha)
    If IsNull(reject) Then
        ShowBoxError box, "The printed p-value " & PDisplay(p) & " does not establish whether p is above or below " & JsNumberToString(alpha) & "."
        Exit Sub
    End If
    percent = PercentText(alpha)
    base = "A one-way ANOVA was conducted to compare the mean " & s.response & " across the groups of " & s.group & "." & vbLf & _
        U("Let \u03bc\u2081, \u03bc\u2082, ..., \u03bc\u2096 denote the population mean ") & s.response & " for the k groups." & vbLf & _
        SYM_H0() & U(": \u03bc\u2081 = \u03bc\u2082 = ... = \u03bc\u2096.") & vbLf & _
        SYM_H1() & ": At least one population mean differs." & vbLf & _
        "The ANOVA result was " & IIf(reject, "", "not ") & "statistically significant, F(" & DfDisplay(df1) & ", " & DfDisplay(df2) & ") = " & _
        FormatFinalNumber(f) & ", p " & PDisplay(p) & ". Therefore, we " & IIf(reject, "reject", "fail to reject") & " " & SYM_H0() & _
        ". There is " & IIf(reject, "sufficient", "insufficient") & " evidence at the " & percent & "% significance level to conclude that the mean " & _
        s.response & " differs across the groups of " & s.group & "."
    tukeyHeader = -1
    For i = 0 To rows.Count - 1
        If RxTest(rows(i), "\bdiff\b.*\blwr\b.*\bupr\b.*p\s*adj", "i") Then
            tukeyHeader = i
            Exit For
        End If
    Next i
    If reject And tukeyHeader >= 0 Then
        If Not TukeyFollowUp(rows, tukeyHeader, alpha, s, percent, followUp) Then
            ShowBoxError box, base & vbLf & "A Tukey adjusted p-value bound is inconclusive. Paste more precise pairwise p-values."
            Exit Sub
        End If
    ElseIf reject Then
        followUp = " Tukey HSD should be used to identify which pairs of group means differ."
    End If
    mStandardProse = base & followUp
    box.RemoveClass "is-error"
    box.TextContent = mStandardProse
    copy.Disabled = (mVarianceRoute = "welch")
End Sub

' The Tukey comparisons after the header row. False when an adjusted p-value bound is
' inconclusive; otherwise sets followUp (possibly "").
Private Function TukeyFollowUp(ByVal rows As clsList, ByVal header As Long, ByVal alpha As Double, ByRef s As TAnovaSettings, ByVal percent As String, ByRef followUp As String) As Boolean
    Dim i As Long, line As String, vals As clsList, labelMatch As clsMatch, label As String, pValue As clsMap
    Dim comparisons As Long, significant As clsList, decision As Variant, split As Long, group1 As String, group2 As String
    Dim descriptions As clsList, k As Long, joined As String
    Set significant = NewList()
    For i = header + 1 To rows.Count - 1
        line = rows(i)
        Set vals = MatchValues(line, NUMBER_PATTERN, "gi")
        Set labelMatch = Rx("^(.+?)\s+(?=-?(?:\d|\.\d))").Exec(line)
        label = ""
        If Not labelMatch Is Nothing Then label = JsTrim(labelMatch.Group(1))
        Set pValue = Nothing
        If vals.Count >= 4 Then Set pValue = ParseAnovaPrintedP(vals(vals.Count - 1))
        If Len(label) > 0 And Not pValue Is Nothing Then
            comparisons = comparisons + 1
            decision = AnovaPDecision(pValue, alpha)
            If IsNull(decision) Then Exit Function
            If decision Then
                split = JsLastIndexOf(label, "-")
                If split > 0 Then
                    group1 = JsSlice(label, 0, split)
                    group2 = JsSlice(label, split + 1)
                Else
                    group1 = label
                    group2 = "the comparison group"
                End If
                significant.Push "between " & group1 & " and " & group2 & " (adjusted p " & PDisplay(pValue) & ")"
            End If
        End If
    Next i
    If significant.Count > 0 Then
        Set descriptions = significant
        If descriptions.Count = 1 Then
            joined = descriptions(0)
        Else
            For k = 0 To descriptions.Count - 2
                If k > 0 Then joined = joined & ", "
                joined = joined & descriptions(k)
            Next k
            joined = joined & ", and " & descriptions(descriptions.Count - 1)
        End If
        followUp = vbLf & U("Post-hoc pairwise comparisons were conducted using Tukey\u2019s HSD test. The results showed that the mean ") & s.response & _
            " differs significantly " & joined & ". " & _
            IIf(significant.Count < comparisons, "The remaining pairwise comparisons were not statistically significant.", "All reported pairwise comparisons were statistically significant.")
    ElseIf comparisons > 0 Then
        followUp = vbLf & U("Post-hoc pairwise comparisons were conducted using Tukey\u2019s HSD test. Although the overall ANOVA was significant, Tukey\u2019s HSD did not identify any pairwise differences as statistically significant at the ") & percent & "% level."
    End If
    TukeyFollowUp = True
End Function

' ---------------------------------------------------------------- Step 3b: Welch + Games-Howell

' parseWelchAnovaResult: clsMap with "error", or "p" (clsMap), "statistic", "dfn", "dfd".
Public Function ParseWelchAnovaResult(ByVal text As String) As clsMap
    Dim result As clsMap, rows As clsList, i As Long, row As String, found As Boolean, parts As clsList, methodIndex As Long
    Dim p As clsMap, statistic As Double, dfn As Double, dfd As Double
    Set result = NewMap()
    Set ParseWelchAnovaResult = result
    If Not RxTest(text, "Welch\s+ANOVA", "i") Then
        result("error") = "Paste the tibble produced by welch_anova_test()."
        Exit Function
    End If
    Set rows = AnovaOutputRows(text)
    For i = 0 To rows.Count - 1
        If RxTest(rows(i), "^\d+\s+") And RxTest(rows(i), "Welch\s+ANOVA\s*$", "i") Then
            row = rows(i)
            found = True
            Exit For
        End If
    Next i
    If Not found Then
        result("error") = "I found the Welch ANOVA heading but could not read its result row."
        Exit Function
    End If
    row = Rx("e([+-])\s+(\d+)", "gi").Replace(row, "e$1$2")
    row = Rx("(<=|>=|<|>)\s+(?=[0-9.])", "g").Replace(row, "$1")
    Set parts = Rx("\s+").Split(row)
    methodIndex = -1
    For i = 0 To parts.Count - 1
        If RxTest(parts(i), "^Welch$", "i") Then
            methodIndex = i
            Exit For
        End If
    Next i
    If methodIndex < 5 Then
        result("error") = "The Welch result row is missing its statistic, degrees of freedom, or p-value."
        Exit Function
    End If
    Set p = ParseAnovaPrintedP(parts(methodIndex - 1))
    statistic = JsNumber(parts(methodIndex - 4))
    dfn = JsNumber(parts(methodIndex - 3))
    dfd = JsNumber(parts(methodIndex - 2))
    If p Is Nothing Then
        result("error") = "The Welch statistic, degrees of freedom, or p-value could not be read."
        Exit Function
    End If
    If Not (JsIsFinite(statistic) And JsIsFinite(dfn) And JsIsFinite(dfd) And JsIsFinite(p("value"))) Then
        result("error") = "The Welch statistic, degrees of freedom, or p-value could not be read."
        Exit Function
    End If
    Set result("p") = p
    result("statistic") = statistic
    result("dfn") = dfn
    result("dfd") = dfd
End Function

Public Function IsGamesHowellOutput(ByVal text As String) As Boolean
    If RxTest(text, "games[\s_\u2013\-]*howell", "i") Then
        IsGamesHowellOutput = True
    Else
        IsGamesHowellOutput = RxTest(Rx("[\s|]", "g").Replace(text, ""), "\.y\.group1group2estimateconf\.lowconf\.highp\.adj", "i")
    End If
End Function

Private Function IsFiniteNumberText(ByVal value As String) As Boolean
    If Len(value) = 0 Then Exit Function
    IsFiniteNumberText = JsIsFinite(JsNumber(value))
End Function

' parseGamesHowellResult: clsMap with "error", or "rows" (clsList of clsMap response,
' group1, group2, p).
Public Function ParseGamesHowellResult(ByVal text As String) As clsMap
    Dim result As clsMap, rows As clsList, invalid As Boolean, lines As clsList, i As Long, k As Long, clean As String
    Dim delimited As Boolean, parts As clsList, p As clsMap, ok As Boolean, row As clsMap, raw As clsList
    Set result = NewMap()
    Set ParseGamesHowellResult = result
    If Not IsGamesHowellOutput(text) Then
        result("error") = U("Paste the Games\u2013Howell table, including .y., group1, group2, and p.adj.")
        Exit Function
    End If
    Set rows = NewList()
    Set lines = Rx("\r?\n").Split(text)
    For i = 0 To lines.Count - 1
        clean = Rx("^\s*(?:>{1,3}\s*)?(?:#{1,3}\s*)?").Replace(lines(i), "")
        clean = Rx("\\([_*])", "g").Replace(clean, "$1")
        clean = Rx("<br\s*\/?\s*>", "gi").Replace(clean, "")
        clean = Rx("e([+-])\s+(\d+)", "gi").Replace(clean, "e$1$2")
        clean = JsTrim(Rx("(<=|>=|<|>)\s+(?=[0-9.])", "g").Replace(clean, "$1"))
        If Len(clean) = 0 Then GoTo NextLine
        If RxTest(clean, "group1|Welch|\bDFn\b|^#|^[-:|\s]+$|<chr>|<dbl>|<int>|^\.\.\.|^\u2139") Then GoTo NextLine
        delimited = InStr(clean, "|") > 0 Or InStr(clean, vbTab) > 0
        If delimited Then
            Set raw = JsSplit(Rx("^\||\|$", "g").Replace(clean, ""), IIf(InStr(clean, "|") > 0, "|", vbTab))
            Set parts = NewList()
            For k = 0 To raw.Count - 1
                parts.Push JsTrim(raw(k))
            Next k
        Else
            Set parts = Rx("\s+").Split(clean)
        End If
        If parts.Count >= 8 Then
            If RxTest(parts(0), "^\d+$") Then Set parts = parts.Slice(1)
        End If
        If parts.Count < 7 Then GoTo NextLine
        Set p = ParseAnovaPrintedP(parts(6))
        ok = Not p Is Nothing
        If ok Then ok = JsIsFinite(p("value"))
        If ok Then ok = (p("value") >= 0 And p("value") <= 1)
        If ok Then ok = (Len(parts(1)) > 0 And Len(parts(2)) > 0)
        If ok Then ok = IsFiniteNumberText(parts(3)) And IsFiniteNumberText(parts(4)) And IsFiniteNumberText(parts(5))
        If Not ok Then
            invalid = True
            GoTo NextLine
        End If
        Set row = NewMap()
        row("response") = parts(0)
        row("group1") = parts(1)
        row("group2") = parts(2)
        Set row("p") = p
        rows.Push row
NextLine:
    Next i
    If invalid Then
        result("error") = U("A Games\u2013Howell row contains an unreadable estimate, interval, or adjusted p-value. Paste the complete kable() table.")
    ElseIf rows.Count > 0 Then
        Set result("rows") = rows
    Else
        result("error") = U("No readable Games\u2013Howell comparison rows were found.")
    End If
End Function

' gamesHowellProse: clsMap with "error" or "prose".
Public Function GamesHowellProse(ByVal games As clsMap, ByVal alpha As Double) As clsMap
    Dim result As clsMap, rows As clsList, i As Long, significant As clsList, row As clsMap, summary As String, decision As Variant
    Set result = NewMap()
    Set GamesHowellProse = result
    Set rows = games("rows")
    For i = 0 To rows.Count - 1
        If IsNull(AnovaPDecision(rows(i)("p"), alpha)) Then
            result("error") = U("A Games\u2013Howell adjusted p-value bound is inconclusive. Paste more precise pairwise p-values.")
            Exit Function
        End If
    Next i
    Set significant = NewList()
    For i = 0 To rows.Count - 1
        Set row = rows(i)
        decision = AnovaPDecision(row("p"), alpha)
        If decision Then significant.Push row("group1") & " vs " & row("group2") & " (adjusted p " & PDisplay(row("p")) & ")"
    Next i
    summary = U("Games\u2013Howell comparisons for ") & rows(0)("response") & ": " & significant.Count & " of " & rows.Count & _
        " reported pairwise comparisons were statistically significant at the " & PercentText(alpha) & "% significance level."
    If significant.Count > 0 Then
        result("prose") = summary & vbLf & "Significant differences: " & significant.Join("; ") & "."
    Else
        result("prose") = summary & " There is insufficient evidence of a difference for any reported pair."
    End If
End Function

Public Sub RenderAnovaWelchInterpretation()
    Dim box As clsNode, copy As clsNode, welchText As String, gamesText As String, s As TAnovaSettings, alpha As Double
    Dim games As clsMap, result As clsMap, welch As clsMap, reject As Variant, percent As String, base As String, postHoc As String, i As Long, rows As clsList
    Set box = Q("#anovaWelchInterpretation")
    Set copy = Q("#copyAnovaWelchProseBtn")
    welchText = Q("#anovaWelchOutput").Value
    gamesText = Q("#anovaGamesHowellOutput").Value
    s = AnovaBlockSettings()
    alpha = s.alpha
    mWelchProse = ""
    copy.Disabled = True
    If Not IsValidAlpha(alpha) Then
        ShowBoxError box, "Enter a significance level between 0 and 1 before interpreting the Welch result."
        Exit Sub
    End If
    If Len(JsTrim(welchText)) = 0 Then
        If Len(JsTrim(gamesText)) = 0 Then
            box.RemoveClass "is-error"
            box.TextContent = U("Paste the Welch ANOVA result or the Games\u2013Howell table.")
            Exit Sub
        End If
        Set games = ParseGamesHowellResult(gamesText)
        If games.Exists("error") Then Set result = games Else Set result = GamesHowellProse(games, alpha)
        box.ToggleClass "is-error", result.Exists("error")
        If result.Exists("error") Then
            box.TextContent = result("error")
            Exit Sub
        End If
        mWelchProse = result("prose") & vbLf & "The omnibus Welch ANOVA result was not supplied, so no overall ANOVA decision is made."
        box.TextContent = mWelchProse
        copy.Disabled = (mVarianceRoute = "standard")
        Exit Sub
    End If
    Set welch = ParseWelchAnovaResult(welchText)
    If welch.Exists("error") Then
        ShowBoxError box, welch("error")
        Exit Sub
    End If
    reject = AnovaPDecision(welch("p"), alpha)
    If IsNull(reject) Then
        ShowBoxError box, "The printed p-value " & PDisplay(welch("p")) & " does not establish whether p is above or below " & JsNumberToString(alpha) & "."
        Exit Sub
    End If
    percent = PercentText(alpha)
    base = U("Since the equal variance assumption was violated, Welch\u2019s ANOVA was conducted to compare the mean ") & s.response & " across the groups of " & s.group & "." & vbLf & _
        U("Let \u03bc\u2081, \u03bc\u2082, \u2026, \u03bc\u2096 denote the population mean ") & s.response & " for the " & JsNumberToString(s.groupCount) & " groups. " & _
        SYM_H0() & U(": \u03bc\u2081 = \u03bc\u2082 = \u2026 = \u03bc\u2096. ") & SYM_H1() & ": At least one population mean differs." & vbLf & _
        "The Welch ANOVA result was " & IIf(reject, "", "not ") & "statistically significant, p " & PDisplay(welch("p")) & ". Therefore, we " & _
        IIf(reject, "reject", "fail to reject") & " " & SYM_H0() & ". There is " & IIf(reject, "sufficient", "insufficient") & " evidence at the " & percent & _
        "% significance level to conclude that the mean " & s.response & " differs across the groups of " & s.group & "."
    If reject Then
        If Len(JsTrim(gamesText)) = 0 Then
            ShowBoxError box, base & vbLf & U("Paste the Games\u2013Howell output to identify which groups differ.")
            Exit Sub
        End If
        Set games = ParseGamesHowellResult(gamesText)
        If games.Exists("error") Then
            ShowBoxError box, base & vbLf & games("error")
            Exit Sub
        End If
        Set rows = games("rows")
        For i = 0 To rows.Count - 1
            If IsNull(AnovaPDecision(rows(i)("p"), alpha)) Then
                ShowBoxError box, base & vbLf & U("A Games\u2013Howell adjusted p-value bound is inconclusive. Paste more precise pairwise p-values.")
                Exit Sub
            End If
        Next i
        postHoc = GamesHowellProse(games, alpha)("prose")
    Else
        postHoc = "Post-hoc pairwise comparisons are not required because the Welch ANOVA was not statistically significant."
    End If
    mWelchProse = base & vbLf & postHoc
    box.RemoveClass "is-error"
    box.TextContent = mWelchProse
    copy.Disabled = (mVarianceRoute = "standard")
End Sub

' ---------------------------------------------------------------- all steps

Public Sub RenderAnovaBlocks()
    Dim s As TAnovaSettings, settings As TAnovaSettings, validAlpha As Boolean, validGroups As Boolean
    Dim varianceCode As String, standardCode As String, welchCode As String, normalityPanel As clsNode, normalityToggle As clsNode
    s = AnovaBlockSettings()
    validAlpha = IsValidAlpha(s.alpha)
    validGroups = IsValidGroupCount(s.groupCount)
    settings = s
    If Not validGroups Then settings.groupCount = 3
    If Not validAlpha Then settings.alpha = 0.05
    Q("#anovaNormalityCode").TextContent = BuildAnovaNormalityCode(settings)
    varianceCode = BuildAnovaVarianceCode(settings)
    If Len(mLastVarianceCode) > 0 And mLastVarianceCode <> varianceCode Then ResetAnovaVarianceCopyState "all"
    mLastVarianceCode = varianceCode
    Q("#anovaVarianceCode").TextContent = varianceCode
    standardCode = BuildStandardAnovaCode(settings)
    welchCode = BuildWelchAnovaCode(settings)
    If Len(mLastStandardCode) > 0 And mLastStandardCode <> standardCode Then ResetAnovaBranchCopyState "standard", "code"
    If Len(mLastWelchCode) > 0 And mLastWelchCode <> welchCode Then ResetAnovaBranchCopyState "welch", "code"
    mLastStandardCode = standardCode
    mLastWelchCode = welchCode
    Q("#anovaStandardCode").TextContent = standardCode
    Q("#anovaWelchCode").TextContent = welchCode
    Set normalityPanel = Q("#anovaNormalityPanel")
    Set normalityToggle = normalityPanel.QuerySelector(".lp-collapse-toggle")
    normalityPanel.ToggleClass "is-unavailable", s.normalityAssumed
    normalityPanel.SetAttribute "aria-disabled", IIf(s.normalityAssumed, "true", "false")
    If Not normalityToggle Is Nothing Then normalityToggle.Disabled = s.normalityAssumed
    Q("#copyAnovaNormalityCodeBtn").Disabled = s.normalityAssumed Or Not validAlpha
    If s.normalityAssumed Then
        SetPanelCompletion normalityPanel, False
        SetLpPanelCollapsed normalityPanel, True
    ElseIf IsNull(mNormalityWasAssumed) Then
        SetLpPanelCollapsed normalityPanel, False
    ElseIf mNormalityWasAssumed Then
        SetLpPanelCollapsed normalityPanel, False
    End If
    mNormalityWasAssumed = s.normalityAssumed
    Q("#copyAnovaVarianceCodeBtn").Disabled = Not validAlpha Or Not validGroups
    ParseAnovaVarianceOutput Q("#anovaVarianceOutput").Value
    RenderAnovaStandardInterpretation
    RenderAnovaWelchInterpretation
End Sub

' ---------------------------------------------------------------- event handlers

Public Sub OnAnovaSettingsChange(ByVal e As clsEvent)
    On Error GoTo Failed
    RenderAnovaBlocks
    Exit Sub
Failed:
    ReportError "opAnova.OnAnovaSettingsChange"
End Sub

Public Sub OnAnovaVarianceOutputInput(ByVal e As clsEvent)
    On Error GoTo Failed
    ParseAnovaVarianceOutput e.CurrentTarget.Value
    Exit Sub
Failed:
    ReportError "opAnova.OnAnovaVarianceOutputInput"
End Sub

Public Sub OnAnovaVarianceOutputReset(ByVal e As clsEvent)
    On Error GoTo Failed
    ResetAnovaVarianceCopyState "prose"
    Exit Sub
Failed:
    ReportError "opAnova.OnAnovaVarianceOutputReset"
End Sub

Public Sub OnLoadAnovaSample(ByVal e As clsEvent)
    On Error GoTo Failed
    Q("#anovaDataset").Value = "dvd"
    Q("#anovaResponse").Value = "Amount"
    Q("#anovaGroup").Value = "Region"
    Q("#anovaGroupCount").Value = "4"
    Q("#anovaAlpha").Value = "0.05"
    Q("#anovaNormalityAssumed").Checked = True
    Q("#anovaVarianceOutput").Value = ""
    RenderAnovaBlocks
    ShowToast "DVD-region ANOVA setup loaded"
    Exit Sub
Failed:
    ReportError "opAnova.OnLoadAnovaSample"
End Sub

Public Sub OnLoadAnovaVarianceOutput(ByVal e As clsEvent)
    Dim s As TAnovaSettings, two As Boolean, data As String, response As String, group As String
    On Error GoTo Failed
    s = AnovaBlockSettings()
    two = IsTwoGroups(s.groupCount)
    data = RColumnName(s.dataset): response = RColumnName(s.response): group = RColumnName(s.group)
    If two Then
        Q("#anovaVarianceOutput").Value = "##" & vbLf & "##  F test to compare two variances" & vbLf & "##" & vbLf & _
            "## data:  " & data & "$" & response & " by " & data & "$" & group & vbLf & _
            "## F = 1.0681, num df = 116, denom df = 73, p-value = 0.7686" & vbLf & _
            "## alternative hypothesis: true ratio of variances is not equal to 1" & vbLf & _
            "## 95 percent confidence interval:" & vbLf & "##  0.6970127 1.6043795" & vbLf & _
            "## sample estimates:" & vbLf & "## ratio of variances" & vbLf & "##           1.068108"
    Else
        Q("#anovaVarianceOutput").Value = "##" & vbLf & "##  Bartlett test of homogeneity of variances" & vbLf & "##" & vbLf & _
            "## data:  " & response & " by " & group & vbLf & _
            "## Bartlett's K-squared = 877.66, df = 3, p-value < 2.2e-16"
    End If
    ParseAnovaVarianceOutput Q("#anovaVarianceOutput").Value
    ShowToast IIf(two, "F-test output loaded", "Bartlett output loaded")
    Exit Sub
Failed:
    ReportError "opAnova.OnLoadAnovaVarianceOutput"
End Sub

Public Sub OnCopyAnovaVarianceProse(ByVal e As clsEvent)
    On Error GoTo Failed
    If Len(mVarianceProse) = 0 Then Exit Sub
    If WriteAnswerText(mVarianceProse) Then ShowToast "Variance-test prose copied" Else ShowToast "Clipboard access is unavailable"
    Exit Sub
Failed:
    ReportError "opAnova.OnCopyAnovaVarianceProse"
End Sub

' The four "Copy code" buttons (normality, variance, standard, Welch).
Public Sub OnCopyAnovaCode(ByVal e As clsEvent)
    Dim button As clsNode, codeSelector As String
    On Error GoTo Failed
    Set button = e.CurrentTarget
    Select Case button.Id
        Case "copyAnovaNormalityCodeBtn": codeSelector = "#anovaNormalityCode"
        Case "copyAnovaVarianceCodeBtn": codeSelector = "#anovaVarianceCode"
        Case "copyAnovaStandardCodeBtn": codeSelector = "#anovaStandardCode"
        Case "copyAnovaWelchCodeBtn": codeSelector = "#anovaWelchCode"
        Case Else: Exit Sub
    End Select
    If button.Disabled Then Exit Sub
    If WriteAnswerText(Q(codeSelector).TextContent) Then ShowToast "R code copied" Else ShowToast "Clipboard access is unavailable"
    Exit Sub
Failed:
    ReportError "opAnova.OnCopyAnovaCode"
End Sub

Public Sub OnAnovaStandardOutputInput(ByVal e As clsEvent)
    On Error GoTo Failed
    ResetAnovaBranchCopyState "standard", "prose"
    RenderAnovaStandardInterpretation
    Exit Sub
Failed:
    ReportError "opAnova.OnAnovaStandardOutputInput"
End Sub

Public Sub OnLoadAnovaStandardOutput(ByVal e As clsEvent)
    Dim s As TAnovaSettings, data As String, response As String, group As String, factor As String
    On Error GoTo Failed
    s = AnovaBlockSettings()
    data = RColumnName(s.dataset): response = RColumnName(s.response): group = RColumnName(s.group)
    factor = "as.factor(" & data & "$" & group & ")"
    Q("#anovaStandardOutput").Value = Join(Array( _
        "##                        Df Sum Sq Mean Sq F value Pr(>F)", _
        "## " & factor & "   3   20.7   6.898   0.866   0.46", _
        "## Residuals             207 1648.8   7.965", _
        "", _
        "##   Tukey multiple comparisons of means", _
        "##     95% family-wise confidence level", _
        "##", _
        "## Fit: aov(formula = " & data & "$" & response & " ~ " & factor & ")", _
        "##", _
        "## $`" & factor & "`", _
        "##                      diff        lwr      upr     p adj", _
        "## North-East   0.8064285714 -0.7887125 2.401570 0.5579913", _
        "## South-East   0.7760617761 -0.8720884 2.424212 0.6151432", _
        "## West-East    0.7763650794 -0.5896321 2.142362 0.4562708", _
        "## South-North -0.0303667954 -1.6785170 1.617783 0.9999609", _
        "## West-North  -0.0300634921 -1.3960607 1.335934 0.9999333", _
        "## West-South   0.0003033033 -1.4272374 1.427844 1.0000000"), vbLf)
    ResetAnovaBranchCopyState "standard", "all"
    RenderAnovaBlocks
    ShowToast "Standard ANOVA and Tukey output loaded"
    Exit Sub
Failed:
    ReportError "opAnova.OnLoadAnovaStandardOutput"
End Sub

Public Sub OnCopyAnovaStandardProse(ByVal e As clsEvent)
    On Error GoTo Failed
    If Len(mStandardProse) = 0 Then Exit Sub
    If WriteAnswerText(mStandardProse) Then ShowToast "Standard ANOVA prose copied" Else ShowToast "Clipboard access is unavailable"
    Exit Sub
Failed:
    ReportError "opAnova.OnCopyAnovaStandardProse"
End Sub

Public Sub OnAnovaWelchOutputInput(ByVal e As clsEvent)
    On Error GoTo Failed
    ResetAnovaBranchCopyState "welch", "prose"
    RenderAnovaWelchInterpretation
    Exit Sub
Failed:
    ReportError "opAnova.OnAnovaWelchOutputInput"
End Sub

Public Sub OnLoadAnovaWelchOutputs(ByVal e As clsEvent)
    Dim s As TAnovaSettings, response As String, times As String
    On Error GoTo Failed
    s = AnovaBlockSettings()
    response = RColumnName(s.response)
    times = ChrW$(&HD7)
    Q("#anovaWelchOutput").Value = Join(Array( _
        "## # A tibble: 1 " & times & " 7", _
        "##   .y.                     n statistic   DFn   DFd     p method", _
        "## * <chr>               <int>     <dbl> <dbl> <dbl> <dbl> <chr>", _
        "## 1 " & response & "  7500     2113.     3 3452.     0 Welch ANOVA"), vbLf)
    Q("#anovaGamesHowellOutput").Value = Join(Array( _
        "## # A tibble: 6 " & times & " 8", _
        "##   .y.            group1 group2 estimate conf.low conf.high    p.adj p.adj.signif", _
        "## * <chr>          <chr>  <chr>     <dbl>    <dbl>     <dbl>    <dbl> <chr>", _
        "## 1 " & response & " Mild   Moderate    2.80    2.64       2.96  1.54e-8 ****", _
        "## 2 " & response & " Mild   None       -1.06   -1.22      -0.913 6.52e-11 ****", _
        "## 3 " & response & " Mild   Severe      3.07    2.89       3.24  1.24e-8 ****", _
        "## 4 " & response & " Moderate None     -3.86   -4.02      -3.71  0        ****", _
        "## 5 " & response & " Moderate Severe    0.269   0.0917     0.446 5.68e-4 ***", _
        "## 6 " & response & " None   Severe      4.13    3.96       4.30  0        ****"), vbLf)
    ResetAnovaBranchCopyState "welch", "prose"
    RenderAnovaBlocks
    ShowToast U("Welch and Games\u2013Howell outputs loaded")
    Exit Sub
Failed:
    ReportError "opAnova.OnLoadAnovaWelchOutputs"
End Sub

Public Sub OnCopyAnovaWelchProse(ByVal e As clsEvent)
    On Error GoTo Failed
    If Len(mWelchProse) = 0 Then Exit Sub
    If WriteAnswerText(mWelchProse) Then ShowToast "Welch ANOVA prose copied" Else ShowToast "Clipboard access is unavailable"
    Exit Sub
Failed:
    ReportError "opAnova.OnCopyAnovaWelchProse"
End Sub

' ---------------------------------------------------------------- multiple correlations

Public Sub RenderCorrMultipleCode()
    Dim dataset As String, parts As clsList, columns As clsList, i As Long, item As String, ready As Boolean
    dataset = JsTrim(Q("#corrMultipleDataset").Value)
    If Len(dataset) = 0 Then dataset = "dataset"
    Set parts = JsSplit(Q("#corrMultipleColumns").Value, ",")
    Set columns = NewList()
    For i = 0 To parts.Count - 1
        item = JsTrim(parts(i))
        If Len(item) > 0 Then columns.Push RColumnName(item)
    Next i
    ready = columns.Count >= 2
    If ready Then
        Q("#corrMultipleCode").TextContent = "# Correlation matrix for selected variables" & vbLf & "psych::corr.test(" & vbLf & "  " & _
            RColumnName(dataset) & " %>% dplyr::select(" & columns.Join(", ") & ")" & vbLf & ")"
    Else
        Q("#corrMultipleCode").TextContent = "Enter at least two columns to generate corr.test() code."
    End If
    Q("#copyCorrMultipleCodeBtn").Disabled = Not ready
End Sub

Public Sub OnCorrMultipleInput(ByVal e As clsEvent)
    On Error GoTo Failed
    RenderCorrMultipleCode
    Exit Sub
Failed:
    ReportError "opAnova.OnCorrMultipleInput"
End Sub

Public Sub OnLoadCorrMultipleSample(ByVal e As clsEvent)
    On Error GoTo Failed
    Q("#corrMultipleDataset").Value = "dataset"
    Q("#corrMultipleColumns").Value = "column1, column2, column3"
    RenderCorrMultipleCode
    ShowToast "Sample loaded"
    Exit Sub
Failed:
    ReportError "opAnova.OnLoadCorrMultipleSample"
End Sub

Public Sub OnCopyCorrMultipleCode(ByVal e As clsEvent)
    On Error GoTo Failed
    If Q("#copyCorrMultipleCodeBtn").Disabled Then Exit Sub
    If WriteAnswerText(Q("#corrMultipleCode").TextContent) Then ShowToast "R code copied" Else ShowToast "Clipboard access is unavailable"
    Exit Sub
Failed:
    ReportError "opAnova.OnCopyCorrMultipleCode"
End Sub
