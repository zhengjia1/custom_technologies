Attribute VB_Name = "opLp2"
' Linear optimisation (lpSolve), part 2: sensitivity-output parsing, the Solution /
' Shadow Price / Reduced Cost / Slack interpretation with per-section copy tracking,
' and the event wiring (InitOpLp).
' Ported from index.html (app-src.js lines 3048-3196 and 4374-4467).
Option Explicit

' copiedSensitivitySections: section key -> True (a Set).
Public CopiedSensitivitySections As clsMap

' sensitivityProseSections: key (solution/shadow/reduced/slack) -> prose, in that order.
Private mSensitivityProseSections As clsMap
Private mSensitivityOutputTimer As Long
Private mLpParserTimer As Long

Private Const SENS_STATUS_HINT As String = "The values will be identified and placed into the fields below automatically."

' ---------------------------------------------------------------- startup

Public Sub InitOpLp()
    Dim inputs As clsList, i As Long, node As clsNode
    opLp.LpAsciiTableText = ""
    Set mSensitivityProseSections = NewMap()
    Set CopiedSensitivitySections = NewMap()
    mSensitivityOutputTimer = 0
    mLpParserTimer = 0
    Q("#sensRawOutput").AddEventListener "input", "opLp2.OnSensRawOutputInput"
    Set inputs = QA(".sensitivity-input")
    For i = 0 To inputs.Count - 1
        Set node = inputs(i)
        node.AddEventListener "input", "opLp2.OnSensitivityInput"
    Next i
    Q("#loadSensitivitySampleBtn").AddEventListener "click", "opLp2.OnLoadSensitivitySample"
    Q("#sensitivityInterpretation").AddEventListener "click", "opLp2.OnSensitivityInterpretationClick"
    Q("#copySensitivityProseBtn").AddEventListener "click", "opLp2.OnCopySensitivityProse"
    Q("#lpBench").AddEventListener "input", "opLp2.OnLpBenchInput"
    Q("#lpBench").AddEventListener "change", "opLp2.OnLpBenchInput"
    Q("#lpBench").AddEventListener "click", "opLp2.OnLpBenchClick"
    Q("#lpSmartInput").AddEventListener "input", "opLp2.OnLpSmartInput"
    Q("#resetLpBtn").AddEventListener "click", "opLp2.OnResetLp"
    Q("#loadLpParserSampleBtn").AddEventListener "click", "opLp2.OnLoadLpParserSample"
    Set inputs = QA(".lp-setting")
    For i = 0 To inputs.Count - 1
        Set node = inputs(i)
        node.AddEventListener "input", "opLp2.OnLpSettingRender"
        node.AddEventListener "change", "opLp2.OnLpSettingRender"
    Next i
    Q("#loadLpSampleBtn").AddEventListener "click", "opLp2.OnLoadLpSample"
    Q("#copyLpCodeBtn").AddEventListener "click", "opLp2.OnCopyLpCode"
    opLp.AddLpConstraint
    opLp.RenderLpModel
    Set inputs = QA(".lp-setting")
    For i = 0 To inputs.Count - 1
        Set node = inputs(i)
        node.AddEventListener "input", "opLp2.OnLpSettingMerged"
        node.AddEventListener "change", "opLp2.OnLpSettingMerged"
    Next i
    Q("#loadLpSampleBtn").AddEventListener "click", "opLp2.OnLoadLpSampleMerged"
    Q("#copyLpAsciiBtn").AddEventListener "click", "opLp2.OnCopyLpAscii"
End Sub

' ---------------------------------------------------------------- copied-section tracking

' The workflow reset after model changes is the shell's ResetLpWorkflowAfterModelChange
' (modApp2); it clears this module's set through ClearCopiedSensitivitySections.

' For the answer-options engine (clearCompletionForAnswerVariant).
Public Sub ForgetCopiedSensitivitySection(ByVal key As String)
    If CopiedSensitivitySections Is Nothing Then Exit Sub
    If CopiedSensitivitySections.Exists(key) Then CopiedSensitivitySections.Remove key
End Sub

Public Sub ClearCopiedSensitivitySections()
    If CopiedSensitivitySections Is Nothing Then Set CopiedSensitivitySections = NewMap()
    CopiedSensitivitySections.Clear
End Sub

' ---------------------------------------------------------------- sensitivity output parser

Private Function SensitivityNumbers(ByVal value As String) As clsList
    Dim out As clsList, matches As Collection, m As clsMatch
    Set out = NewList()
    Set matches = Rx("[-+]?(?:\d*\.?\d+)(?:e[-+]?\d+)?", "gi").MatchAll(value)
    For Each m In matches
        out.Push JsNumber(m.Value)
    Next m
    Set SensitivityNumbers = out
End Function

Private Function FirstValueText(ByVal block As clsMap) As String
    Dim values As clsList
    Set values = block("values")
    If values.Count = 0 Then FirstValueText = "undefined" Else FirstValueText = JsNumberToString(CDbl(values(0)))
End Function

Private Sub PushAll(ByVal target As clsList, ByVal source As clsList)
    Dim i As Long
    For i = 0 To source.Count - 1
        target.Push source(i)
    Next i
End Sub

Public Function ParseSensitivityOutput(ByVal raw As String) As Boolean
    Dim status As clsNode, variables As clsList, constraintCount As Long, rawLines As clsList, lines As clsList
    Dim i As Long, line As String, solution As clsList, names As clsList, namesMatch As Boolean, k As Long, candidate As clsList
    Dim blocks As clsList, current As clsMap, block As clsMap, m As clsMatch, start As Double
    Dim duals As clsList, coefFrom As clsList, coefTo As clsList, rhsFrom As clsList, rhsTo As clsList
    Dim rhsUnavailable As Boolean, sizesValid As Boolean, everything As clsList, objectiveText As String

    Set status = Q("#sensParseStatus")
    Set variables = PredictorList(Q("#lpVariables").Value)
    constraintCount = opLp.LpRows().Count
    If Len(JsTrim(raw)) = 0 Then
        status.ClassName = "lp-parser-status"
        status.TextContent = SENS_STATUS_HINT
        Exit Function
    End If
    If variables.Count = 0 Or constraintCount = 0 Then
        status.ClassName = "lp-parser-status error"
        status.TextContent = "Complete the optimisation model before pasting its sensitivity output."
        Exit Function
    End If
    Set rawLines = JsSplit(RxReplace(raw, "\r", "", "g"), vbLf)
    Set lines = NewList()
    For i = 0 To rawLines.Count - 1
        line = JsTrim(RxReplace(RxReplace(CStr(rawLines(i)), "^\s*##\s?", ""), "^\s*>\s?", ""))
        If Len(line) > 0 Then lines.Push line
    Next i
    Set solution = NewList()
    For i = 0 To lines.Count - 2
        Set names = Rx("\s+").Split(RxReplace(CStr(lines(i)), "[""',]", "", "g"))
        namesMatch = (names.Count = variables.Count)
        If namesMatch Then
            For k = 0 To names.Count - 1
                If JsLower(CStr(names(k))) <> JsLower(CStr(variables(k))) Then namesMatch = False: Exit For
            Next k
        End If
        If namesMatch Then
            Set candidate = SensitivityNumbers(CStr(lines(i + 1)))
            If candidate.Count = variables.Count Then
                Set solution = candidate
                Exit For
            End If
        End If
    Next i
    Set blocks = NewList()
    Set current = Nothing
    For i = 0 To lines.Count - 1
        line = lines(i)
        If RxTest(line, "^NULL$", "i") Then
            Set block = NewMap()
            Set block("values") = NewList()
            block("isNull") = True
            blocks.Push block
            Set current = Nothing
        Else
            Set m = RxMatch(line, "^\[(\d+)\]\s*(.*)$")
            If Not m Is Nothing Then
                start = JsNumber(m.Group(1))
                If JsEq(start, 1) Or current Is Nothing Then
                    Set current = NewMap()
                    Set current("values") = NewList()
                    current("isNull") = False
                    blocks.Push current
                End If
                PushAll current("values"), SensitivityNumbers(m.Group(2))
            End If
        End If
    Next i
    If solution.Count <> variables.Count Or blocks.Count < 5 Then
        status.ClassName = "lp-parser-status error"
        status.TextContent = "Could not identify the complete solution, objective value, status, duals, and objective-coefficient ranges."
        Exit Function
    End If
    objectiveText = FirstValueText(blocks(0))
    Set duals = blocks(2)("values")
    Set coefFrom = blocks(3)("values")
    Set coefTo = blocks(4)("values")
    rhsUnavailable = (blocks.Count < 7)
    If Not rhsUnavailable Then rhsUnavailable = CBool(blocks(5)("isNull")) Or CBool(blocks(6)("isNull"))
    If rhsUnavailable Then
        Set rhsFrom = NewList()
        Set rhsTo = NewList()
    Else
        Set rhsFrom = blocks(5)("values")
        Set rhsTo = blocks(6)("values")
    End If
    sizesValid = duals.Count = constraintCount + variables.Count And coefFrom.Count = variables.Count And coefTo.Count = variables.Count
    If sizesValid And Not rhsUnavailable Then sizesValid = (rhsFrom.Count = constraintCount And rhsTo.Count = constraintCount)
    Set everything = NewList()
    PushAll everything, solution
    PushAll everything, duals
    PushAll everything, coefFrom
    PushAll everything, coefTo
    PushAll everything, rhsFrom
    PushAll everything, rhsTo
    If Not sizesValid Or Not opLp.AllFinite(everything) Then
        status.ClassName = "lp-parser-status error"
        status.TextContent = "The pasted vector lengths do not match the current model. Expected " & variables.Count & " solution values, " & _
            (constraintCount + variables.Count) & " dual values, and " & constraintCount & " values for each available RHS range."
        Exit Function
    End If
    Q("#sensSolution").Value = solution.Join(", ")
    Q("#sensObjectiveValue").Value = objectiveText
    Q("#sensDuals").Value = duals.Join(", ")
    Q("#sensCoefFrom").Value = coefFrom.Join(", ")
    Q("#sensCoefTo").Value = coefTo.Join(", ")
    Q("#sensRhsFrom").Value = rhsFrom.Join(", ")
    Q("#sensRhsTo").Value = rhsTo.Join(", ")
    status.ClassName = "lp-parser-status success"
    status.TextContent = "Parsed solution, objective value " & objectiveText & ", status " & FirstValueText(blocks(1)) & ", duals, and coefficient ranges." & _
        IIf(rhsUnavailable, " RHS sensitivity ranges were returned as NULL.", " RHS sensitivity ranges were also parsed.")
    RenderSensitivityInterpretation
    ParseSensitivityOutput = True
End Function

Public Function BuildSensitivitySampleOutput() As String
    Dim variables As clsList, constraintCount As Long, solution As clsList, duals As clsList, coefFrom As clsList
    Dim coefTo As clsList, rhsFrom As clsList, rhsTo As clsList, i As Long
    Set variables = PredictorList(Q("#lpVariables").Value)
    constraintCount = opLp.LpRows().Count
    If variables.Count = 0 Or constraintCount = 0 Then Exit Function
    Set solution = NewList(): Set duals = NewList(): Set coefFrom = NewList(): Set coefTo = NewList()
    Set rhsFrom = NewList(): Set rhsTo = NewList()
    For i = 0 To variables.Count - 1
        If i = 0 Then
            solution.Push 7#
        ElseIf i = 1 Then
            solution.Push 2.2
        Else
            solution.Push 0#
        End If
        coefFrom.Push CDbl(i + 1) / 10
        coefTo.Push CDbl(i + 3) / 10
    Next i
    For i = 0 To constraintCount + variables.Count - 1
        If i < constraintCount Then
            duals.Push CDbl(i + 1) / 100
        ElseIf i = constraintCount Then
            duals.Push -0.1
        Else
            duals.Push 0#
        End If
    Next i
    For i = 0 To constraintCount - 1
        rhsFrom.Push CDbl(i + 1) * 50
        rhsTo.Push CDbl(i + 1) * 100
    Next i
    BuildSensitivitySampleOutput = variables.Join("       ") & vbLf & solution.Join("   ") & vbLf & _
        "[1] 100" & vbLf & "[1] 0" & vbLf & "[1] " & duals.Join(" ") & vbLf & "[1] " & coefFrom.Join(" ") & _
        vbLf & "[1] " & coefTo.Join(" ") & vbLf & "[1] " & rhsFrom.Join(" ") & vbLf & "[1] " & rhsTo.Join(" ")
End Function

' ---------------------------------------------------------------- interpretation

Public Function NaturalList(ByVal items As clsList) As String
    If items.Count < 2 Then
        If items.Count = 1 Then NaturalList = items(0)
        Exit Function
    End If
    If items.Count = 2 Then
        NaturalList = items(0) & " and " & items(1)
        Exit Function
    End If
    NaturalList = items.Slice(0, -1).Join(", ") & ", and " & items(items.Count - 1)
End Function

Private Function FormatOptimalQuantity(ByVal value As Double) As String
    If JsLt(JsAbs(JsSub(value, JsMathRound(value))), 0.000000001) Then
        FormatOptimalQuantity = JsNumberToString(JsMathRound(value))
    Else
        FormatOptimalQuantity = FormatFinalNumber(value)
    End If
End Function

Private Function IsNonNegativityConstraint(ByVal name As String) As Boolean
    IsNonNegativityConstraint = RxTest(JsTrim(name), "^non[\s-]*negativity\b", "i")
End Function

Private Function RoundedSlackQuantity(ByVal value As Double) As Double
    Dim rounded As Double
    rounded = JsDiv(JsMathRound(JsMul(value, 100)), 100)
    If JsIsNegativeZero(rounded) Then rounded = 0
    RoundedSlackQuantity = rounded
End Function

Private Function FormatSlackQuantity(ByVal value As Double) As String
    FormatSlackQuantity = FormatFinalNumber(RoundedSlackQuantity(value))
End Function

Private Function IsBindingSlack(ByVal value As Double) As Boolean
    IsBindingSlack = JsEq(RoundedSlackQuantity(value), 0)
End Function

Private Function ListNumber(ByVal values As clsList, ByVal index As Long) As Double
    If index < values.Count Then ListNumber = CDbl(values(index)) Else ListNumber = JsNaN()
End Function

Public Sub RenderSensitivityInterpretation()
    Dim variables As clsList, objective As clsList, rows As clsList, solution As clsList, objectiveValueText As String
    Dim objectiveValue As Double, duals As clsList, coefFrom As clsList, coefTo As clsList, rhsFrom As clsList, rhsTo As clsList
    Dim constraintCount As Long, rhsUnavailable As Boolean, rhsValid As Boolean, valid As Boolean, box As clsNode, everything As clsList
    Dim direction As String, slacks As clsList, i As Long, k As Long, row As clsMap, coefficients As clsList, lhs As Double
    Dim rhs As Double, slack As Double, parts As clsList, solutionText As String, shadowParagraphs As clsList
    Dim shadow As Double, zero As Boolean, binding As Boolean, paragraph As String, bindingNames As clsList
    Dim reducedParagraphs As clsList, cost As Double, improvement As Double, improved As Double, slackText As String
    Dim html As String, key As Variant, title As String, objectiveCoefficient As Double, ks As clsList

    CopiedSensitivitySections.Clear
    SetPanelCompletion Q("#sensitivityInterpretationPanel"), False
    Set variables = PredictorList(Q("#lpVariables").Value)
    Set objective = opLp.NumberList(Q("#lpObjective").Value)
    Set rows = opLp.LpRows()
    Set solution = opLp.NumberList(Q("#sensSolution").Value)
    objectiveValueText = JsTrim(Q("#sensObjectiveValue").Value)
    objectiveValue = JsNumber(objectiveValueText)
    Set duals = opLp.NumberList(Q("#sensDuals").Value)
    Set coefFrom = opLp.NumberList(Q("#sensCoefFrom").Value)
    Set coefTo = opLp.NumberList(Q("#sensCoefTo").Value)
    Set rhsFrom = opLp.NumberList(Q("#sensRhsFrom").Value)
    Set rhsTo = opLp.NumberList(Q("#sensRhsTo").Value)
    constraintCount = rows.Count
    rhsUnavailable = (rhsFrom.Count = 0 And rhsTo.Count = 0)
    rhsValid = rhsUnavailable Or (rhsFrom.Count = constraintCount And rhsTo.Count = constraintCount)
    Set everything = NewList()
    PushAll everything, objective
    PushAll everything, solution
    PushAll everything, duals
    PushAll everything, coefFrom
    PushAll everything, coefTo
    PushAll everything, rhsFrom
    PushAll everything, rhsTo
    valid = variables.Count > 0 And rows.Count > 0 And objective.Count = variables.Count And solution.Count = variables.Count And _
        Len(objectiveValueText) > 0 And JsIsFinite(objectiveValue) And duals.Count = constraintCount + variables.Count And _
        coefFrom.Count = variables.Count And coefTo.Count = variables.Count And rhsValid
    If valid Then valid = opLp.AllFinite(everything)
    Set box = Q("#sensitivityInterpretation")
    If Not valid Then
        Set mSensitivityProseSections = NewMap()
        box.InnerHtml = "<div class=""empty-state compact""><div class=""empty-state-inner""><strong>Complete the sensitivity results</strong><p>Include the optimal solution and objective value, followed by " & _
            constraintCount & " shadow prices and " & variables.Count & " reduced costs. RHS ranges may both be blank when lpSolve returns NULL.</p></div></div>"
        Q("#copySensitivityProseBtn").Disabled = True
        Exit Sub
    End If
    direction = Q("#lpDirection").Value

    Set slacks = NewList()
    For i = 0 To rows.Count - 1
        Set row = rows(i)
        Set coefficients = row("coefficients")
        lhs = 0
        For k = 0 To coefficients.Count - 1
            lhs = JsAdd(lhs, JsMul(CDbl(coefficients(k)), ListNumber(solution, k)))
        Next k
        rhs = row("rhs")
        If row("operator") = "<=" Then
            slack = JsSub(rhs, lhs)
        ElseIf row("operator") = ">=" Then
            slack = JsSub(lhs, rhs)
        Else
            slack = JsAbs(JsSub(rhs, lhs))
        End If
        slacks.Push slack
    Next i

    Set parts = NewList()
    For i = 0 To variables.Count - 1
        parts.Push FormatOptimalQuantity(CDbl(solution(i))) & " " & variables(i)
    Next i
    solutionText = "Optimal production is " & NaturalList(parts) & ", with a total objective value of $" & FormatFinalNumber(objectiveValue) & "."

    Set shadowParagraphs = NewList()
    Set bindingNames = NewList()
    For i = 0 To rows.Count - 1
        Set row = rows(i)
        If Not IsNonNegativityConstraint(CStr(row("name"))) Then
            shadow = duals(i)
            zero = Abs(shadow) < 0.000000001
            binding = IsBindingSlack(CDbl(slacks(i)))
            paragraph = "The shadow price of the " & row("name") & " constraint is $" & JsToFixed(shadow, 2) & ". This means that increasing the available " & _
                row("name") & " resource by 1 unit will " & IIf(zero, "not change", "change") & " the optimal objective value" & _
                IIf(zero, ".", " by $" & JsToFixed(Abs(shadow), 2) & ".")
            If rhsUnavailable Then
                paragraph = paragraph & " The allowable RHS sensitivity range was not returned by lpSolve."
            Else
                paragraph = paragraph & " This applies provided the available resource remains within the range [" & FormatFinalNumber(CDbl(rhsFrom(i))) & ", " & _
                    FormatFinalNumber(CDbl(rhsTo(i))) & "] and all other constraints remain unchanged."
            End If
            If binding Then
                paragraph = paragraph & " The constraint is binding and the resource is fully utilised at the optimum."
            Else
                paragraph = paragraph & " The constraint is non-binding" & ChrW$(&H2014) & "the resource is not fully utilised at the optimum."
            End If
            shadowParagraphs.Push paragraph
            If binding Then bindingNames.Push row("name")
        End If
    Next i
    If bindingNames.Count > 0 Then
        shadowParagraphs.Push "The binding constraints are " & NaturalList(bindingNames) & " because their slack values round to 0. These resources are fully utilised at the optimal solution."
    Else
        shadowParagraphs.Push "There are no binding resource constraints in the supplied solution."
    End If

    Set reducedParagraphs = NewList()
    For i = 0 To variables.Count - 1
        cost = duals(constraintCount + i)
        zero = Abs(cost) < 0.000000001
        improvement = Abs(cost)
        objectiveCoefficient = objective(i)
        If direction = "max" Then improved = JsAdd(objectiveCoefficient, improvement) Else improved = JsSub(objectiveCoefficient, improvement)
        paragraph = "The reduced cost of " & variables(i) & " is $" & JsToFixed(cost, 2) & ". "
        If zero Then
            paragraph = paragraph & variables(i) & " is already produced at its optimal quantity in the current solution."
        Else
            paragraph = paragraph & variables(i) & " is not produced in the optimal solution. Its objective coefficient would need to improve by $" & JsToFixed(improvement, 2) & _
                ChrW$(&H2014) & "from $" & JsToFixed(objectiveCoefficient, 2) & " to $" & JsToFixed(improved, 2) & ChrW$(&H2014) & "before it becomes worthwhile to produce."
        End If
        paragraph = paragraph & " The objective coefficient can vary between " & FormatFinalNumber(CDbl(coefFrom(i))) & " and " & FormatFinalNumber(CDbl(coefTo(i))) & _
            " without changing the current optimal basis."
        reducedParagraphs.Push paragraph
    Next i

    Set parts = NewList()
    For i = 0 To rows.Count - 1
        Set row = rows(i)
        parts.Push row("name") & " = " & FormatSlackQuantity(CDbl(slacks(i)))
    Next i
    slackText = "Slack values (unused resource) are: " & parts.Join(", ") & ". A slack of 0 indicates a binding constraint."

    Set mSensitivityProseSections = NewMap()
    mSensitivityProseSections("solution") = solutionText
    mSensitivityProseSections("shadow") = shadowParagraphs.Join(vbLf & vbLf)
    mSensitivityProseSections("reduced") = reducedParagraphs.Join(vbLf & vbLf)
    mSensitivityProseSections("slack") = slackText
    html = "<div class=""sensitivity-prose-grid"">"
    Set ks = mSensitivityProseSections.Keys
    For Each key In ks
        Select Case key
            Case "solution": title = "Solution"
            Case "shadow": title = "Shadow Price"
            Case "reduced": title = "Reduced Cost"
            Case Else: title = "Slack"
        End Select
        html = html & "<section class=""sensitivity-prose-card""><div class=""sensitivity-prose-head""><h3>" & title & _
            "</h3><button class=""sample-btn"" type=""button"" data-sensitivity-copy=""" & key & """>Copy</button></div><p>" & _
            EscapeHtml(CStr(mSensitivityProseSections(key))) & "</p></section>"
    Next key
    box.InnerHtml = html & "</div>"
    Q("#copySensitivityProseBtn").Disabled = False
End Sub

Public Function SensitivityProseText(ByVal key As String) As String
    If mSensitivityProseSections.Exists(key) Then SensitivityProseText = mSensitivityProseSections(key)
End Function

Private Sub ShowRenameAlert()
    CallOp "ShowVariableRenameAlert"
End Sub

Public Sub MarkSensitivitySectionCopied(ByVal key As String, ByVal button As clsNode)
    Dim card As clsNode, allCopied As Boolean, section As Variant, panel As clsNode, ks As clsList
    CopiedSensitivitySections(key) = True
    button.TextContent = "Copied"
    button.Disabled = True
    Set card = button.Closest(".sensitivity-prose-card")
    If Not card Is Nothing Then card.AddClass "is-copy-complete"
    allCopied = (mSensitivityProseSections.Count > 0)
    Set ks = mSensitivityProseSections.Keys
    For Each section In ks
        If Not CopiedSensitivitySections.Exists(CStr(section)) Then allCopied = False
    Next section
    If allCopied Then
        Set panel = Q("#sensitivityInterpretationPanel")
        SetPanelCompletion panel, True
        SetLpPanelCollapsed panel, True
    End If
End Sub

' ---------------------------------------------------------------- event handlers

Public Sub OnSensRawOutputInput(ByVal e As clsEvent)
    On Error GoTo Failed
    If mSensitivityOutputTimer <> 0 Then ClearTimeout mSensitivityOutputTimer
    mSensitivityOutputTimer = SetTimeout("opLp2.SensitivityOutputTimerFired", 180)
    Exit Sub
Failed:
    ReportError "opLp2.OnSensRawOutputInput"
End Sub

Public Sub SensitivityOutputTimerFired()
    On Error GoTo Failed
    mSensitivityOutputTimer = 0
    ParseSensitivityOutput Q("#sensRawOutput").Value
    Exit Sub
Failed:
    ReportError "opLp2.SensitivityOutputTimerFired"
End Sub

Public Sub OnSensitivityInput(ByVal e As clsEvent)
    On Error GoTo Failed
    RenderSensitivityInterpretation
    Exit Sub
Failed:
    ReportError "opLp2.OnSensitivityInput"
End Sub

Public Sub OnLoadSensitivitySample(ByVal e As clsEvent)
    Dim sample As String
    On Error GoTo Failed
    sample = BuildSensitivitySampleOutput()
    If Len(sample) = 0 Then
        ShowToast "Complete the optimisation model first"
        Exit Sub
    End If
    Q("#sensRawOutput").Value = sample
    ParseSensitivityOutput sample
    Q("#sensRawOutput").Focus
    ShowToast "Sensitivity output sample loaded"
    Exit Sub
Failed:
    ReportError "opLp2.OnLoadSensitivitySample"
End Sub

Public Sub OnSensitivityInterpretationClick(ByVal e As clsEvent)
    Dim button As clsNode, key As String, text As String
    On Error GoTo Failed
    If e.Target Is Nothing Then Exit Sub
    If e.Target.NodeType <> 1 Then Exit Sub
    Set button = e.Target.Closest("[data-sensitivity-copy]")
    If button Is Nothing Then Exit Sub
    key = JsStr(button.DataGet("sensitivityCopy"))
    text = SensitivityProseText(key)
    If Len(text) = 0 Then Exit Sub
    If WriteAnswerText(text) Then
        MarkSensitivitySectionCopied key, button
        If key = "solution" Then
            ShowToast "Solution copied"
            ShowRenameAlert
        Else
            ShowToast button.Closest(".sensitivity-prose-card").QuerySelector("h3").TextContent & " copied"
        End If
    Else
        ShowToast "Clipboard access is unavailable"
    End If
    Exit Sub
Failed:
    ReportError "opLp2.OnSensitivityInterpretationClick"
End Sub

Public Sub OnCopySensitivityProse(ByVal e As clsEvent)
    Dim prose As String, buttons As clsList, i As Long, button As clsNode, values As clsList
    On Error GoTo Failed
    Set values = mSensitivityProseSections.Values
    prose = values.Join(vbLf & vbLf)
    If Len(prose) = 0 Then Exit Sub
    If WriteAnswerText(prose) Then
        Set buttons = Q("#sensitivityInterpretation").QuerySelectorAll("[data-sensitivity-copy]")
        For i = 0 To buttons.Count - 1
            Set button = buttons(i)
            MarkSensitivitySectionCopied JsStr(button.DataGet("sensitivityCopy")), button
        Next i
        ShowToast "All prose copied"
        ShowRenameAlert
    Else
        ShowToast "Clipboard access is unavailable"
    End If
    Exit Sub
Failed:
    ReportError "opLp2.OnCopySensitivityProse"
End Sub

' #lpBench input and change: any model setting or constraint edit resets the workflow.
Public Sub OnLpBenchInput(ByVal e As clsEvent)
    On Error GoTo Failed
    If e.Target Is Nothing Then Exit Sub
    If e.Target.NodeType <> 1 Then Exit Sub
    If e.Target.Matches(".lp-setting, .lp-row-input") Then CallOp "ResetLpWorkflowAfterModelChange"
    Exit Sub
Failed:
    ReportError "opLp2.OnLpBenchInput"
End Sub

Public Sub OnLpBenchClick(ByVal e As clsEvent)
    On Error GoTo Failed
    If e.Target Is Nothing Then Exit Sub
    If e.Target.NodeType <> 1 Then Exit Sub
    If Not e.Target.Closest("#resetLpBtn, #loadLpSampleBtn, #loadLpParserSampleBtn, .lp-remove-constraint") Is Nothing Then CallOp "ResetLpWorkflowAfterModelChange"
    Exit Sub
Failed:
    ReportError "opLp2.OnLpBenchClick"
End Sub

Public Sub OnLpSmartInput(ByVal e As clsEvent)
    On Error GoTo Failed
    If mLpParserTimer <> 0 Then ClearTimeout mLpParserTimer
    mLpParserTimer = SetTimeout("opLp2.LpParserTimerFired", 260)
    Exit Sub
Failed:
    ReportError "opLp2.OnLpSmartInput"
End Sub

Public Sub LpParserTimerFired()
    On Error GoTo Failed
    mLpParserTimer = 0
    opLp.RunLpSmartParser
    Exit Sub
Failed:
    ReportError "opLp2.LpParserTimerFired"
End Sub

Public Sub OnResetLp(ByVal e As clsEvent)
    On Error GoTo Failed
    opLp.ResetLpModel
    Exit Sub
Failed:
    ReportError "opLp2.OnResetLp"
End Sub

Public Function LpParserSampleText() As String
    LpParserSampleText = Join(Array( _
        "| Minimize total effort using | Effort = 6 $X_1$ + 7$X_2$ + |", _
        "| decision variables $X_1$, $X_2$, $X_3$, | 9 $X_3$ + 8 $X_4$ + |", _
        "| $X_4$, $X_5$, $X_6$ | 10 $X_5$ + 20 $X_6$ |", _
        "|---|---|", _
        "| Subject to | |", _
        "| Credit Requirement | $2X_1+4X_2+3X_3+3X_4+4X_5+5X_6\geq12$ |", _
        "| Weekly effort | $6X_1+7X_2+9X_3+8X_4+10X_5+20X_6\leq30$ |", _
        "| CS2 Prereq | $X_3-X_4\geq0$ |", _
        "| BT2 prereq | $(X_1+X_5)-X_2\geq0$ |", _
        "| DT prereq | $(X_1+X_3)-2X_6\geq0$ |", _
        "| Binary, Integer, Non-negativity Constraints | X1, X2, X3 all binary, integers |"), vbLf)
End Function

Public Sub OnLoadLpParserSample(ByVal e As clsEvent)
    On Error GoTo Failed
    Q("#lpSmartInput").Value = LpParserSampleText()
    opLp.RunLpSmartParser
    Q("#lpSmartInput").Focus
    ShowToast "Parser sample loaded"
    Exit Sub
Failed:
    ReportError "opLp2.OnLoadLpParserSample"
End Sub

Public Sub OnLpSettingRender(ByVal e As clsEvent)
    On Error GoTo Failed
    opLp.RenderLpModel
    If AppState("id") = "sensitivity" Then RenderSensitivityInterpretation
    Exit Sub
Failed:
    ReportError "opLp2.OnLpSettingRender"
End Sub

Public Sub OnLoadLpSample(ByVal e As clsEvent)
    Dim values As clsMap
    On Error GoTo Failed
    Q("#lpSmartInput").Value = ""
    Q("#lpParseStatus").ClassName = "lp-parser-status"
    Q("#lpParseStatus").TextContent = "Manual sample loaded. Paste a model above whenever you want to replace it automatically."
    Q("#lpDirection").Value = "max"
    Q("#lpVariableType").Value = "continuous"
    Q("#lpVariables").Value = "X1, X2, X3"
    Q("#lpObjective").Value = "0.15, 0.40, 0.60"
    Q("#lpSensitivity").Checked = True
    Q("#lpConstraintRows").InnerHtml = ""
    opLp.AddLpConstraint ConstraintValues("Budget", "15, 30, 40", "<=", "200")
    opLp.AddLpConstraint ConstraintValues("Storage", "7, 5, 8", "<=", "60")
    opLp.AddLpConstraint ConstraintValues("Diet", "1, 0, 0", "<=", "7")
    opLp.RenderLpModel
    ShowToast "Sample loaded"
    Exit Sub
Failed:
    ReportError "opLp2.OnLoadLpSample"
End Sub

Private Function ConstraintValues(ByVal name As String, ByVal coefficients As String, ByVal operator As String, ByVal rhs As String) As clsMap
    Dim result As clsMap
    Set result = NewMap()
    result("name") = name
    result("coefficients") = coefficients
    result("operator") = operator
    result("rhs") = rhs
    Set ConstraintValues = result
End Function

Public Sub OnCopyLpCode(ByVal e As clsEvent)
    On Error GoTo Failed
    If Q("#copyLpCodeBtn").Disabled Then Exit Sub
    If WriteAnswerText(Q("#lpCode").TextContent) Then ShowToast "R code copied" Else ShowToast "Clipboard access is unavailable"
    Exit Sub
Failed:
    ReportError "opLp2.OnCopyLpCode"
End Sub

Public Sub OnLpSettingMerged(ByVal e As clsEvent)
    On Error GoTo Failed
    If AppState("id") = "lpsolve" Then RenderSensitivityInterpretation
    Exit Sub
Failed:
    ReportError "opLp2.OnLpSettingMerged"
End Sub

Public Sub OnLoadLpSampleMerged(ByVal e As clsEvent)
    On Error GoTo Failed
    If AppState("id") = "lpsolve" Then RenderSensitivityInterpretation
    Exit Sub
Failed:
    ReportError "opLp2.OnLoadLpSampleMerged"
End Sub

Public Sub OnCopyLpAscii(ByVal e As clsEvent)
    On Error GoTo Failed
    If Q("#copyLpAsciiBtn").Disabled Or Len(opLp.LpAsciiTableText) = 0 Then Exit Sub
    If WriteAnswerText(opLp.LpAsciiTableText) Then ShowToast "ASCII table copied" Else ShowToast "Clipboard access is unavailable"
    Exit Sub
Failed:
    ReportError "opLp2.OnCopyLpAscii"
End Sub
