Attribute VB_Name = "opLp"
' Linear optimisation (lpSolve): model-text normalisation and smart parser, the manual
' model builder with auto-appending constraint rows, validation, the ASCII model table
' and the generated R code. Ported from index.html (app-src.js lines 2765-3047).
' Sensitivity analysis, event wiring and InitOpLp live in opLp2.
Option Explicit

' lpAsciiTableText: the ASCII model table for the last valid model ("" when invalid).
Public LpAsciiTableText As String

Private Const LP_PARSER_HINT As String = "Recognises min/max objectives, named constraints, X_1 variables, \u2264/\u2265 or \leq/\geq, and binary/integer declarations."

Public Function LpParserHint() As String
    LpParserHint = JsU(LP_PARSER_HINT)
End Function

' ---------------------------------------------------------------- model text parser

Public Function NormaliseLpSource(ByVal text As String) As String
    text = RxReplace(text, "\r", "", "g")
    text = RxReplace(text, "\\text\{([^}]*)\}", "$1", "gi")
    text = RxReplace(text, "\\(?:geq|ge)(?:\{\})?", ">=", "gi")
    text = RxReplace(text, "\\(?:leq|le)(?:\{\})?", "<=", "gi")
    text = RxReplace(text, "[\u2265\u2A7E]", ">=", "g")
    text = RxReplace(text, "[\u2264\u2A7D]", "<=", "g")
    text = RxReplace(text, "\\(?:times|cdot)", "*", "gi")
    text = RxReplace(text, "\\(?:quad|qquad|,|;|!)", " ", "gi")
    text = RxReplace(text, "[\u2212\u2013\u2014]", "-", "g")
    text = RxReplace(text, "([A-Za-z][A-Za-z0-9.]*)_\{?([A-Za-z0-9]+)\}?", "$1$2", "g")
    NormaliseLpSource = RxReplace(text, "\$+", "", "g")
End Function

' Every name-like token that is not a model keyword (duplicates kept, like the website).
Public Function LpExpressionVariables(ByVal expression As String) As clsList
    Dim out As clsList, matches As Collection, m As clsMatch, name As String
    Set out = NewList()
    Set matches = Rx("[A-Za-z][A-Za-z0-9.]*", "g").MatchAll(expression)
    For Each m In matches
        name = m.Value
        Select Case JsLower(name)
            Case "min", "max", "minimize", "minimise", "maximize", "maximise", "subject", "to", "all", _
                 "binary", "integer", "integers", "non", "negativity"
            Case Else
                out.Push name
        End Select
    Next m
    Set LpExpressionVariables = out
End Function

' { coefficients: clsList of Double, found: clsMap of names }
Public Function LpCoefficients(ByVal expression As String, ByVal variables As clsList) As clsMap
    Dim compact As String, coefficients As clsList, found As clsMap, i As Long, matches As Collection, m As clsMatch
    Dim index As Long, magnitude As Double, result As clsMap
    compact = RxReplace(expression, "[()\s*]", "", "g")
    Set coefficients = NewList()
    For i = 0 To variables.Count - 1
        coefficients.Push 0#
    Next i
    Set found = NewMap()
    Set matches = Rx("([+-]?)(?:(\d+(?:\.\d+)?|\.\d+))?([A-Za-z][A-Za-z0-9.]*)", "g").MatchAll(compact)
    For Each m In matches
        index = variables.IndexOf(m.Group(3))
        If index >= 0 Then
            If m.Matched(2) Then magnitude = JsNumber(m.Group(2)) Else magnitude = 1
            If m.Group(1) = "-" Then magnitude = -magnitude
            coefficients(index) = JsAdd(CDbl(coefficients(index)), magnitude)
            found(m.Group(3)) = True
        End If
    Next m
    Set result = NewMap()
    Set result("coefficients") = coefficients
    Set result("found") = found
    Set LpCoefficients = result
End Function

Private Function LpObjectiveSideScore(ByVal side As String) As Long
    Dim numericTerms As Long
    numericTerms = Rx("(?:^|[+\-\s])\d+(?:\.\d+)?\s*\*?\s*[A-Za-z]", "g").MatchAll(side).Count
    LpObjectiveSideScore = LpExpressionVariables(side).Count * 10 + numericTerms * 2 + IIf(RxTest(side, "[+\-]"), 1, 0)
End Function

Public Function LpObjectiveExpression(ByVal block As String) As String
    Dim equalsIndex As Long, sideA As String, sideB As String, chosen As String
    equalsIndex = JsIndexOf(block, "=")
    If equalsIndex < 0 Then
        LpObjectiveExpression = JsTrim(RxReplace(block, "\+\s*$", ""))
        Exit Function
    End If
    sideA = JsSlice(block, 0, equalsIndex)
    sideB = JsSlice(block, equalsIndex + 1)
    If LpObjectiveSideScore(sideB) > LpObjectiveSideScore(sideA) Then chosen = sideB Else chosen = sideA
    LpObjectiveExpression = JsTrim(RxReplace(chosen, "\+\s*$", ""))
End Function

Public Function IsLpNonNegativityDeclaration(ByVal name As String, ByVal left As String, ByVal operator As String, ByVal rhs As Double) As Boolean
    Dim variables As clsList, coefficients As clsList, i As Long, value As Double, ones As Long, allUnit As Boolean
    If operator <> ">=" Or JsNe(rhs, 0) Then Exit Function
    Set variables = LpExpressionVariables(left)
    If variables.Count = 0 Then Exit Function
    Set coefficients = LpCoefficients(left, variables)("coefficients")
    allUnit = True
    For i = 0 To coefficients.Count - 1
        value = coefficients(i)
        If JsLt(JsAbs(JsSub(value, 1)), 0.000000001) Then ones = ones + 1
        If Not (JsLt(JsAbs(value), 0.000000001) Or JsLt(JsAbs(JsSub(value, 1)), 0.000000001)) Then allUnit = False
    Next i
    IsLpNonNegativityDeclaration = (ones = 1 And allUnit) And (RxTest(name, "non[-\s]?negativity", "i") Or variables.Count = 1)
End Function

Private Function SplitLpTableRow(ByVal line As String) As clsList
    Dim cells As clsList, out As clsList, i As Long
    Set cells = JsSplit(line, "|")
    If Len(JsTrim(CStr(cells(0)))) = 0 Then cells.Shift
    If cells.Count > 0 Then
        If Len(JsTrim(CStr(cells(cells.Count - 1)))) = 0 Then cells.Pop
    End If
    Set out = NewList()
    For i = 0 To cells.Count - 1
        out.Push JsTrim(CStr(cells(i)))
    Next i
    Set SplitLpTableRow = out
End Function

Private Function IsLpSeparatorRow(ByVal cells As clsList) As Boolean
    Dim i As Long
    For i = 0 To cells.Count - 1
        If Not RxTest(RxReplace(CStr(cells(i)), "\s", "", "g"), "^:?-{3,}:?$") Then Exit Function
    Next i
    IsLpSeparatorRow = True
End Function

Private Function LpFailure(ByVal message As String) As clsMap
    Dim result As clsMap
    Set result = NewMap()
    result("ok") = False
    result("message") = message
    Set LpFailure = result
End Function

' parseLpModelText: { ok, message } or { ok, direction, variables, objective, constraints,
' type, binaryVariables, binaryPositions }. constraints: clsList of clsMap
' { name, left, operator, rhs, coefficients }.
Public Function ParseLpModelText(ByVal rawText As String) As clsMap
    Dim text As String, rawLines As clsList, lines As clsList, i As Long, line As String, cells As clsList
    Dim tableRows As clsList, subjectIndex As Long, objectiveRowIndex As Long, direction As String
    Dim objectiveBlock As String, parts As clsList, objectiveExpression As String, candidateRows As clsList
    Dim colon As Long, start As Long, parsedConstraints As clsList, declarationRows As clsList, joined As String
    Dim expression As String, operatorMatch As clsMatch, left As String, rhsMatch As clsMatch, rawName As String
    Dim name As String, operator As String, rhs As Double, constraint As clsMap, variables As clsList
    Dim objectiveParsed As clsMap, constraints As clsList, declarationText As String, declaredVariables As clsList
    Dim binaryVariables As clsList, modelType As String, binaryPositions As clsList, model As clsMap, names As clsList, j As Long

    text = NormaliseLpSource(rawText)
    Set rawLines = JsSplit(text, vbLf)
    Set lines = NewList()
    For i = 0 To rawLines.Count - 1
        line = JsTrim(CStr(rawLines(i)))
        If Len(line) > 0 Then lines.Push line
    Next i
    Set tableRows = NewList()
    For i = 0 To lines.Count - 1
        If JsIncludes(CStr(lines(i)), "|") Then
            Set cells = SplitLpTableRow(CStr(lines(i)))
            If cells.Count >= 2 Then
                If Not IsLpSeparatorRow(cells) Then tableRows.Push cells
            End If
        End If
    Next i
    subjectIndex = -1
    objectiveRowIndex = -1
    For i = 0 To tableRows.Count - 1
        Set cells = tableRows(i)
        If subjectIndex < 0 Then
            If RxTest(cells.Join(" "), "subject\s+to", "i") Then subjectIndex = i
        End If
        If objectiveRowIndex < 0 Then
            If RxTest(CStr(cells(0)), "\b(?:minimi[sz]e|maximi[sz]e)\b.*\bobjective\b", "i") Then objectiveRowIndex = i
        End If
    Next i
    If RxTest(text, "\bminimi[sz]e\b", "i") Then
        direction = "min"
    ElseIf RxTest(text, "\bmaximi[sz]e\b", "i") Then
        direction = "max"
    End If
    If objectiveRowIndex >= 0 Then
        Set cells = tableRows(objectiveRowIndex)
        objectiveBlock = cells.Slice(1).Join(" ")
    ElseIf subjectIndex >= 0 Then
        Set parts = NewList()
        For i = 0 To subjectIndex - 1
            Set cells = tableRows(i)
            parts.Push cells.Slice(1).Join(" ")
        Next i
        objectiveBlock = parts.Join(" ")
    End If
    If Len(objectiveBlock) = 0 Then
        For i = 0 To lines.Count - 1
            line = lines(i)
            If JsIncludes(line, "=") Then
                If LpExpressionVariables(JsSplit(line, "=").Slice(1).Join("=")).Count > 0 Then
                    objectiveBlock = line
                    Exit For
                End If
            End If
        Next i
        If Len(objectiveBlock) = 0 Then objectiveBlock = Rx("subject\s+to", "i").Split(text)(0)
    End If
    objectiveExpression = LpObjectiveExpression(objectiveBlock)

    If subjectIndex >= 0 Then
        Set candidateRows = tableRows.Slice(subjectIndex + 1)
    ElseIf objectiveRowIndex >= 0 Then
        Set candidateRows = tableRows.Slice(objectiveRowIndex + 1)
    Else
        start = -1
        For i = 0 To lines.Count - 1
            If RxTest(CStr(lines(i)), "subject\s+to", "i") Then start = i: Exit For
        Next i
        start = start + 1
        If start < 0 Then start = 0
        Set candidateRows = NewList()
        For i = start To lines.Count - 1
            line = lines(i)
            colon = JsIndexOf(line, ":")
            If colon > 0 Then
                candidateRows.Push NewList(JsSlice(line, 0, colon), JsSlice(line, colon + 1))
            Else
                candidateRows.Push NewList("", line)
            End If
        Next i
    End If

    Set parsedConstraints = NewList()
    Set declarationRows = NewList()
    For i = 0 To candidateRows.Count - 1
        Set cells = candidateRows(i)
        joined = cells.Join(" ")
        If RxTest(joined, "\b(binary|integers?|non[- ]?negativity)\b", "i") And Not RxTest(cells.Slice(1).Join(" "), "(<=|>=|=)") Then declarationRows.Push joined
        expression = cells.Slice(1).Join(" ")
        If Len(expression) = 0 Then expression = CStr(cells(0))
        Set operatorMatch = RxMatch(expression, "(<=|>=|=)")
        If Not operatorMatch Is Nothing Then
            left = JsTrim(JsSlice(expression, 0, operatorMatch.Index))
            Set rhsMatch = RxMatch(JsSlice(expression, operatorMatch.Index + Len(operatorMatch.Value)), "[-+]?(?:\d+(?:\.\d+)?|\.\d+)")
            If Len(left) > 0 And Not rhsMatch Is Nothing Then
                If cells.Count > 1 Then rawName = CStr(cells(0)) Else rawName = ""
                name = JsTrim(RxReplace(RxReplace(rawName, "[*_`]", " ", "g"), "\s+", " ", "g"))
                If Len(name) = 0 Then name = "Constraint " & (parsedConstraints.Count + 1)
                operator = operatorMatch.Value
                rhs = JsNumber(rhsMatch.Value)
                If IsLpNonNegativityDeclaration(name, left, operator, rhs) Then
                    declarationRows.Push joined
                Else
                    Set constraint = NewMap()
                    constraint("name") = name
                    constraint("left") = left
                    constraint("operator") = operator
                    constraint("rhs") = rhs
                    parsedConstraints.Push constraint
                End If
            End If
        End If
    Next i

    Set variables = NewList()
    AddLpVariables variables, objectiveExpression
    For i = 0 To parsedConstraints.Count - 1
        Set constraint = parsedConstraints(i)
        AddLpVariables variables, CStr(constraint("left"))
    Next i
    If Len(direction) = 0 Or variables.Count = 0 Or parsedConstraints.Count = 0 Then
        Set ParseLpModelText = LpFailure("I need a min/max objective and at least one constraint with a numeric right-hand side.")
        Exit Function
    End If
    Set objectiveParsed = LpCoefficients(objectiveExpression, variables)
    If objectiveParsed("found").Count = 0 Then
        Set ParseLpModelText = LpFailure("I found the model direction, but not the objective-function coefficients.")
        Exit Function
    End If
    Set constraints = NewList()
    For i = 0 To parsedConstraints.Count - 1
        Set constraint = parsedConstraints(i)
        Set constraint("coefficients") = LpCoefficients(CStr(constraint("left")), variables)("coefficients")
        constraints.Push constraint
    Next i

    declarationText = declarationRows.Join(" ")
    Set names = LpExpressionVariables(declarationText)
    Set declaredVariables = NewList()
    For i = 0 To names.Count - 1
        If variables.Includes(names(i)) Then declaredVariables.Push names(i)
    Next i
    Set binaryVariables = NewList()
    If RxTest(declarationText, "\bbinary\b", "i") Then
        If declaredVariables.Count > 0 Then
            Set binaryVariables = declaredVariables
        ElseIf RxTest(declarationText, "\ball\b[^.]*\bbinary\b|\bbinary\b[^.]*\ball\b", "i") Then
            Set binaryVariables = variables.Clone
        End If
    End If
    If binaryVariables.Count > 0 Then
        modelType = "binary"
    ElseIf RxTest(declarationText, "\bintegers?\b", "i") Then
        modelType = "integer"
    Else
        modelType = "continuous"
    End If
    Set binaryPositions = NewList()
    For i = 0 To binaryVariables.Count - 1
        j = variables.IndexOf(binaryVariables(i)) + 1
        If j > 0 Then binaryPositions.Push j
    Next i
    Set model = NewMap()
    model("ok") = True
    model("direction") = direction
    Set model("variables") = variables
    Set model("objective") = objectiveParsed("coefficients")
    Set model("constraints") = constraints
    model("type") = modelType
    Set model("binaryVariables") = binaryVariables
    Set model("binaryPositions") = binaryPositions
    Set ParseLpModelText = model
End Function

Private Sub AddLpVariables(ByVal variables As clsList, ByVal expression As String)
    Dim names As clsList, i As Long
    Set names = LpExpressionVariables(expression)
    For i = 0 To names.Count - 1
        If Not variables.Includes(names(i)) Then variables.Push names(i)
    Next i
End Sub

Public Sub ApplyParsedLpModel(ByVal model As clsMap)
    Dim constraints As clsList, constraint As clsMap, values As clsMap, i As Long, typeDescription As String
    Dim status As clsNode, modelType As String, coefficients As clsList
    CallOp "ResetLpWorkflowAfterModelChange"
    modelType = model("type")
    Q("#lpDirection").Value = model("direction")
    Q("#lpVariables").Value = model("variables").Join(", ")
    Q("#lpObjective").Value = model("objective").Join(", ")
    Q("#lpVariableType").Value = modelType
    Q("#lpBinary").Value = model("binaryPositions").Join(", ")
    Q("#lpSensitivity").Checked = (modelType = "continuous")
    Q("#lpConstraintRows").InnerHtml = ""
    Set constraints = model("constraints")
    For i = 0 To constraints.Count - 1
        Set constraint = constraints(i)
        Set coefficients = constraint("coefficients")
        Set values = NewMap()
        values("name") = constraint("name")
        values("coefficients") = coefficients.Join(", ")
        values("operator") = constraint("operator")
        values("rhs") = JsNumberToString(CDbl(constraint("rhs")))
        AddLpConstraint values
    Next i
    RenderLpModel
    If AppState("id") = "lpsolve" Or AppState("id") = "sensitivity" Then opLp2.RenderSensitivityInterpretation
    If modelType = "binary" Then
        typeDescription = model("binaryVariables").Join(", ") & " binary"
    ElseIf modelType = "integer" Then
        typeDescription = "all variables integer"
    Else
        typeDescription = "continuous variables"
    End If
    Set status = Q("#lpParseStatus")
    status.ClassName = "lp-parser-status success"
    status.TextContent = "Detected " & IIf(model("direction") = "min", "minimisation", "maximisation") & ", " & model("variables").Count & _
        " variables, " & constraints.Count & " constraints, and " & typeDescription & "."
End Sub

Public Function RunLpSmartParser() As Boolean
    Dim raw As String, status As clsNode, model As clsMap
    raw = JsTrim(Q("#lpSmartInput").Value)
    Set status = Q("#lpParseStatus")
    If Len(raw) = 0 Then
        status.ClassName = "lp-parser-status"
        status.TextContent = LpParserHint()
        Exit Function
    End If
    Set model = ParseLpModelText(raw)
    If Not CBool(model("ok")) Then
        status.ClassName = "lp-parser-status error"
        status.TextContent = model("message")
        Exit Function
    End If
    ApplyParsedLpModel model
    RunLpSmartParser = True
End Function

' ---------------------------------------------------------------- ASCII table

' tableRows: clsList of two-item clsLists.
Public Function BuildAsciiTable(ByVal tableRows As clsList) As String
    Dim cleanRows As clsList, row As clsList, clean As clsList, i As Long, c As Long, widths(0 To 1) As Long
    Dim border As String, lines As clsList
    Set cleanRows = NewList()
    For i = 0 To tableRows.Count - 1
        Set row = tableRows(i)
        Set clean = NewList()
        For c = 0 To row.Count - 1
            clean.Push JsTrim(RxReplace(JsStr(row(c)), "\s+", " ", "g"))
        Next c
        cleanRows.Push clean
    Next i
    For c = 0 To 1
        widths(c) = 0
        For i = 0 To cleanRows.Count - 1
            Set row = cleanRows(i)
            If Len(row(c)) > widths(c) Then widths(c) = Len(row(c))
        Next i
    Next c
    border = "+" & String$(widths(0) + 2, "-") & "+" & String$(widths(1) + 2, "-") & "+"
    Set lines = NewList(border)
    For i = 0 To cleanRows.Count - 1
        Set row = cleanRows(i)
        lines.Push "| " & JsPadEnd(CStr(row(0)), widths(0)) & " | " & JsPadEnd(CStr(row(1)), widths(1)) & " |"
        lines.Push border
    Next i
    BuildAsciiTable = lines.Join(vbLf)
End Function

' ---------------------------------------------------------------- manual model builder

Public Sub ResetLpModel()
    Dim inputs As clsList, i As Long, node As clsNode
    Q("#lpSmartInput").Value = ""
    Q("#sensRawOutput").Value = ""
    Q("#sensParseStatus").ClassName = "lp-parser-status"
    Q("#sensParseStatus").TextContent = "The values will be identified and placed into the fields below automatically."
    Q("#lpParseStatus").ClassName = "lp-parser-status"
    Q("#lpParseStatus").TextContent = LpParserHint()
    Q("#lpDirection").Value = "max"
    Q("#lpVariableType").Value = "continuous"
    Q("#lpVariables").Value = ""
    Q("#lpObjective").Value = ""
    Q("#lpBinary").Value = ""
    Q("#lpSensitivity").Checked = False
    Set inputs = QA(".sensitivity-input")
    For i = 0 To inputs.Count - 1
        Set node = inputs(i)
        node.Value = ""
    Next i
    Q("#lpConstraintRows").InnerHtml = ""
    AddLpConstraint
    RenderLpModel
    If AppState("id") = "lpsolve" Or AppState("id") = "sensitivity" Then opLp2.RenderSensitivityInterpretation
    Q("#lpSmartInput").Focus
    ShowToast "Optimisation model reset"
End Sub

Public Function LpRowHasContent(ByVal row As clsNode) As Boolean
    Dim inputs As clsList, i As Long, node As clsNode
    Set inputs = row.QuerySelectorAll(".lp-name, .lp-coefficients, .lp-rhs")
    For i = 0 To inputs.Count - 1
        Set node = inputs(i)
        If Len(JsTrim(node.Value)) > 0 Then LpRowHasContent = True: Exit Function
    Next i
End Function

' document.querySelectorAll(".lp-constraint-row"): rows are only ever created by
' AddLpConstraint inside #lpConstraintRows, so its subtree holds them all (and a subtree
' query avoids a document walk after every row change).
Private Function LpRowElements() As clsList
    Dim container As clsNode
    Set container = Q("#lpConstraintRows")
    If container Is Nothing Then Set LpRowElements = QA(".lp-constraint-row") Else Set LpRowElements = container.QuerySelectorAll(".lp-constraint-row")
End Function

Public Sub SyncLpConstraintButtons()
    Dim rows As clsList, onlyOne As Boolean, i As Long, row As clsNode
    Set rows = LpRowElements()
    onlyOne = (rows.Count <= 1)
    For i = 0 To rows.Count - 1
        Set row = rows(i)
        row.QuerySelector(".lp-remove-constraint").Disabled = onlyOne Or Not LpRowHasContent(row)
    Next i
End Sub

Public Sub RefreshLpAfterRowChange()
    SyncLpConstraintButtons
    RenderLpModel
    If AppState("id") = "lpsolve" Or AppState("id") = "sensitivity" Then opLp2.RenderSensitivityInterpretation
End Sub

Public Sub EnsureLpTrailingBlankRow()
    Dim rows As clsList
    Set rows = LpRowElements()
    If rows.Count = 0 Then
        AddLpConstraint
    ElseIf LpRowHasContent(rows(rows.Count - 1)) Then
        AddLpConstraint
    End If
End Sub

Private Function MapText(ByVal values As clsMap, ByVal key As String) As String
    If values Is Nothing Then Exit Function
    If values.Exists(key) Then MapText = JsStr(values(key))
End Function

' values: clsMap { name, coefficients, operator, rhs } (strings) or Nothing.
Public Function AddLpConstraint(Optional ByVal values As clsMap = Nothing, Optional ByVal afterRow As clsNode = Nothing) As clsNode
    Dim row As clsNode, inputs As clsList, i As Long, node As clsNode, operator As String
    Set row = CreateElement("div")
    row.ClassName = "lp-constraint-row"
    row.InnerHtml = "<input class=""column-input lp-row-input lp-name"" aria-label=""Constraint name"" placeholder=""Constraint name"" value=""" & EscapeHtml(MapText(values, "name")) & """>" & _
        "<input class=""column-input lp-row-input lp-coefficients"" aria-label=""Constraint coefficients"" placeholder=""15, 30, 40"" value=""" & EscapeHtml(MapText(values, "coefficients")) & """>" & _
        "<select class=""column-input lp-row-input lp-operator"" aria-label=""Constraint operator""><option value=""&lt;="">" & ChrW$(&H2264) & "</option><option value=""&gt;="">" & ChrW$(&H2265) & "</option><option value=""="">=</option></select>" & _
        "<input class=""column-input lp-row-input lp-rhs"" aria-label=""Constraint right-hand side"" type=""number"" step=""any"" placeholder=""200"" value=""" & EscapeHtml(MapText(values, "rhs")) & """>" & _
        "<button class=""sample-btn constraint-step-btn lp-remove-constraint"" type=""button"" aria-label=""Remove this constraint"" title=""Remove this constraint"">" & ChrW$(&H2212) & "</button>"
    operator = MapText(values, "operator")
    If Len(operator) = 0 Then operator = "<="
    row.QuerySelector(".lp-operator").Value = operator
    Set inputs = row.QuerySelectorAll(".lp-row-input")
    For i = 0 To inputs.Count - 1
        Set node = inputs(i)
        node.AddEventListener "input", "opLp.OnLpRowRender"
    Next i
    row.QuerySelector("select").AddEventListener "change", "opLp.OnLpRowRender"
    row.QuerySelector(".lp-remove-constraint").AddEventListener "click", "opLp.OnLpRemoveConstraint"
    If Not afterRow Is Nothing Then
        afterRow.InsertAdjacentElement "afterend", row
    Else
        Q("#lpConstraintRows").AppendChild row
    End If
    SyncLpConstraintButtons
    Set AddLpConstraint = row
End Function

Public Sub OnLpRowRender(ByVal e As clsEvent)
    On Error GoTo Failed
    RenderLpModel
    Exit Sub
Failed:
    ReportError "opLp.OnLpRowRender"
End Sub

Public Sub OnLpRemoveConstraint(ByVal e As clsEvent)
    Dim row As clsNode, focusRow As clsNode, nameInput As clsNode
    On Error GoTo Failed
    Set row = e.CurrentTarget.Closest(".lp-constraint-row")
    Set focusRow = row.NextElementSibling
    If focusRow Is Nothing Then Set focusRow = row.PreviousElementSibling
    If LpRowElements().Count <= 1 Then Exit Sub
    row.Remove
    RefreshLpAfterRowChange
    If Not focusRow Is Nothing Then
        Set nameInput = focusRow.QuerySelector(".lp-name")
        If Not nameInput Is Nothing Then nameInput.Focus
    End If
    Exit Sub
Failed:
    ReportError "opLp.OnLpRemoveConstraint"
End Sub

Public Function NumberList(ByVal value As String) As clsList
    Dim items As clsList, out As clsList, i As Long
    Set items = PredictorList(value)
    Set out = NewList()
    For i = 0 To items.Count - 1
        out.Push JsNumber(CStr(items(i)))
    Next i
    Set NumberList = out
End Function

Public Function AllFinite(ByVal values As clsList) As Boolean
    Dim i As Long
    For i = 0 To values.Count - 1
        If Not JsIsFinite(CDbl(values(i))) Then Exit Function
    Next i
    AllFinite = True
End Function

' lpRows: clsList of clsMap { element, hasContent, name, coefficientText, coefficients,
' operator, rhs, rhsText } for rows with content.
Public Function LpRows() As clsList
    Dim elements As clsList, out As clsList, i As Long, element As clsNode, row As clsMap, name As String
    Set elements = LpRowElements()
    Set out = NewList()
    For i = 0 To elements.Count - 1
        Set element = elements(i)
        If LpRowHasContent(element) Then
            Set row = NewMap()
            Set row("element") = element
            row("hasContent") = True
            name = JsTrim(element.QuerySelector(".lp-name").Value)
            If Len(name) = 0 Then name = "Constraint"
            row("name") = name
            row("coefficientText") = JsTrim(element.QuerySelector(".lp-coefficients").Value)
            Set row("coefficients") = NumberList(CStr(row("coefficientText")))
            row("operator") = element.QuerySelector(".lp-operator").Value
            row("rhsText") = JsTrim(element.QuerySelector(".lp-rhs").Value)
            row("rhs") = JsNumber(CStr(row("rhsText")))
            out.Push row
        End If
    Next i
    Set LpRows = out
End Function

Private Function FormatLpNumber(ByVal value As Variant) As String
    Dim n As Double
    n = JsToNumber(value)
    If JsIsFinite(n) Then FormatLpNumber = JsNumberToString(n) Else FormatLpNumber = JsStr(value)
End Function

Private Function VariableAt(ByVal variables As clsList, ByVal index As Long) As String
    If index < variables.Count Then VariableAt = variables(index) Else VariableAt = "undefined"
End Function

Private Function LpEquation(ByVal coefficients As clsList, ByVal variables As clsList, ByVal operator As String, ByVal rhs As Variant, ByVal ascii As Boolean) As String
    Dim parts As clsList, i As Long, symbol As String
    Set parts = NewList()
    For i = 0 To coefficients.Count - 1
        parts.Push FormatLpNumber(coefficients(i)) & VariableAt(variables, i)
    Next i
    If ascii Then
        symbol = operator
    ElseIf operator = "<=" Then
        symbol = ChrW$(&H2264)
    ElseIf operator = ">=" Then
        symbol = ChrW$(&H2265)
    Else
        symbol = "="
    End If
    LpEquation = parts.Join(" + ") & " " & symbol & " " & FormatLpNumber(rhs)
End Function

Private Function RowRhsValid(ByVal row As clsMap) As Boolean
    RowRhsValid = Len(row("rhsText")) > 0 And JsIsFinite(CDbl(row("rhs")))
End Function

Private Function LpConstraintLines(ByVal rows As clsList, ByVal kind As String) As String
    Dim lines As clsList, i As Long, row As clsMap, value As String, coefficients As clsList, parts As clsList, k As Long, label As String
    Set lines = NewList()
    For i = 0 To rows.Count - 1
        Set row = rows(i)
        Select Case kind
            Case "matrix"
                Set coefficients = row("coefficients")
                Set parts = NewList()
                For k = 0 To coefficients.Count - 1
                    parts.Push FormatLpNumber(coefficients(k))
                Next k
                value = parts.Join(", ")
            Case "dirs"
                value = """" & row("operator") & """"
            Case Else
                value = FormatLpNumber(row("rhs"))
        End Select
        label = JsTrim(RxReplace(CStr(row("name")), "[\r\n#]+", " ", "g"))
        If Len(label) = 0 Then label = "Constraint " & (i + 1)
        lines.Push "    " & value & IIf(i < rows.Count - 1, ",", "") & " # " & label
    Next i
    LpConstraintLines = lines.Join(vbLf)
End Function

Public Sub RenderLpModel()
    Dim variables As clsList, objective As clsList, direction As String, modelType As String, rows As clsList
    Dim sensitivityAllowed As Boolean, sensitivity As Boolean, binary As clsList, constraintErrors As clsList
    Dim i As Long, row As clsMap, field As clsNode, coefficients As clsList, countCorrect As Boolean, numbersValid As Boolean
    Dim invalid As Boolean, n As Long, constraintsValid As Boolean, objectiveValid As Boolean, binaryValid As Boolean
    Dim valid As Boolean, asciiRows As clsList, table As String, message As String, allRhs As Boolean, cellText As String
    Dim typeOption As String, sensitivityLines As String, code As String, quoted As clsList, element As clsNode

    EnsureLpTrailingBlankRow
    Set variables = PredictorList(Q("#lpVariables").Value)
    Set objective = NumberList(Q("#lpObjective").Value)
    direction = Q("#lpDirection").Value
    modelType = Q("#lpVariableType").Value
    Set rows = LpRows()
    sensitivityAllowed = (modelType = "continuous")
    sensitivity = sensitivityAllowed And Q("#lpSensitivity").Checked
    Set binary = NumberList(Q("#lpBinary").Value)
    Q("#lpBinaryField").ToggleClass "hidden", modelType <> "binary"
    Q("#lpSensitivityField").ToggleClass "hidden", Not sensitivityAllowed
    Q("#lpSensitivity").Disabled = Not sensitivityAllowed
    If Not sensitivityAllowed Then Q("#lpSensitivity").Checked = False

    Set constraintErrors = NewList()
    allRhs = True
    For i = 0 To rows.Count - 1
        Set row = rows(i)
        Set element = row("element")
        Set field = element.QuerySelector(".lp-coefficients")
        Set coefficients = row("coefficients")
        countCorrect = (variables.Count > 0 And coefficients.Count = variables.Count)
        numbersValid = AllFinite(coefficients)
        invalid = Not countCorrect Or Not numbersValid
        field.ToggleClass "input-error", invalid
        field.SetAttribute "aria-invalid", IIf(invalid, "true", "false")
        If variables.Count > 0 And coefficients.Count <> variables.Count Then
            n = coefficients.Count
            constraintErrors.Push "Constraint " & (i + 1) & " (" & ChrW$(&H201C) & row("name") & ChrW$(&H201D) & ") has " & n & " coefficient" & IIf(n = 1, "", "s") & _
                "; exactly " & variables.Count & " are required for " & variables.Join(", ") & ". Use 0 when a variable is absent."
        ElseIf countCorrect And Not numbersValid Then
            constraintErrors.Push "Constraint " & (i + 1) & " (" & ChrW$(&H201C) & row("name") & ChrW$(&H201D) & ") contains a non-numeric coefficient. Enter exactly " & variables.Count & " numbers."
        End If
        If Not RowRhsValid(row) Then allRhs = False
    Next i
    If variables.Count = 0 Then constraintErrors.Unshift "Enter the decision-variable names before completing the constraints."
    Q("#lpConstraintError").TextContent = constraintErrors.Join(" ")
    constraintsValid = variables.Count > 0 And rows.Count > 0 And constraintErrors.Count = 0 And allRhs
    objectiveValid = variables.Count > 0 And objective.Count = variables.Count And AllFinite(objective)
    binaryValid = (modelType <> "binary" Or binary.Count > 0)
    valid = objectiveValid And constraintsValid And binaryValid

    If valid Then
        Set asciiRows = NewList()
        asciiRows.Push NewList("Model component", "Mathematical expression")
        asciiRows.Push NewList(IIf(direction = "max", "Maximise", "Minimise") & " objective", LpEquation(objective, variables, "=", "Z", True))
        For i = 0 To rows.Count - 1
            Set row = rows(i)
            asciiRows.Push NewList(row("name"), LpEquation(row("coefficients"), variables, CStr(row("operator")), row("rhs"), True))
        Next i
        For i = 0 To variables.Count - 1
            asciiRows.Push NewList("Non-negativity: " & variables(i), variables(i) & " >= 0")
        Next i
        LpAsciiTableText = BuildAsciiTable(asciiRows)
    Else
        LpAsciiTableText = ""
    End If
    Q("#copyLpAsciiBtn").Disabled = Not valid

    table = "<table class=""lp-preview""><thead><tr><th>Model component</th><th>Mathematical expression</th></tr></thead><tbody><tr><td>" & _
        IIf(direction = "max", "Maximise", "Minimise") & " objective</td><td>" & _
        EscapeHtml(RxReplace(LpEquation(objective, variables, "=", "Z", False), " = NaN$", "")) & "</td></tr>"
    For i = 0 To rows.Count - 1
        Set row = rows(i)
        Set coefficients = row("coefficients")
        If coefficients.Count = variables.Count And AllFinite(coefficients) And Len(row("rhsText")) > 0 Then
            cellText = LpEquation(coefficients, variables, CStr(row("operator")), row("rhs"), False)
        Else
            cellText = "Invalid constraint" & ChrW$(&H2014) & "correct the highlighted coefficient field"
        End If
        table = table & "<tr><td>" & EscapeHtml(CStr(row("name"))) & "</td><td>" & EscapeHtml(cellText) & "</td></tr>"
    Next i
    For i = 0 To variables.Count - 1
        table = table & "<tr><td>Non-negativity: " & EscapeHtml(CStr(variables(i))) & "</td><td>" & EscapeHtml(CStr(variables(i))) & " " & ChrW$(&H2265) & " 0</td></tr>"
    Next i
    table = table & "</tbody></table>"
    Q("#lpTablePreview").InnerHtml = table

    If Not valid Then
        If constraintErrors.Count > 0 Then
            message = constraintErrors.Join(vbLf)
        ElseIf Not objectiveValid Then
            message = "The objective needs exactly " & variables.Count & " numeric coefficient" & IIf(variables.Count = 1, "", "s") & ChrW$(&H2014) & "one for each decision variable."
        ElseIf rows.Count = 0 Then
            message = "Enter at least one constraint."
        ElseIf Not allRhs Then
            message = "Every constraint needs a numeric right-hand-side value."
        Else
            message = "Complete the binary-variable positions."
        End If
        Q("#lpCode").TextContent = "ERROR: " & message & vbLf & vbLf & "R code will not be generated until the model is valid."
        Q("#copyLpCodeBtn").Disabled = True
        Exit Sub
    End If

    If modelType = "integer" Then
        typeOption = "," & vbLf & "    all.int = TRUE"
    ElseIf modelType = "binary" Then
        typeOption = "," & vbLf & "    binary.vec = c(" & binary.Join(", ") & ")"
    End If
    If sensitivity Then
        sensitivityLines = vbLf & "range_objcoef <- cbind(lp_sol$sens.coef.from, lp_sol$sens.coef.to)" & _
            vbLf & "lp_sol$duals          # first constraints: shadow prices; then variables: reduced costs" & _
            vbLf & "lp_sol$sens.coef.from # lower objective-coefficient bounds" & _
            vbLf & "lp_sol$sens.coef.to   # upper objective-coefficient bounds" & _
            vbLf & "lp_sol$sens.rhs.from  # lower RHS bounds" & _
            vbLf & "lp_sol$sens.rhs.to    # upper RHS bounds" & _
            vbLf & "slack <- constraint_rhs - constraint_mat %*% lp_sol$solution # 0 means binding for <= constraints"
    End If
    Set quoted = NewList()
    For i = 0 To variables.Count - 1
        quoted.Push """" & RString(CStr(variables(i))) & """"
    Next i
    code = "library(lpSolve)" & vbLf & vbLf & _
        "# Decision variables: " & variables.Join(", ") & vbLf & _
        "objective_fun <- c(" & objective.Join(", ") & ")" & vbLf & vbLf & _
        "constraint_mat <- matrix(" & vbLf & "  c(" & vbLf & LpConstraintLines(rows, "matrix") & vbLf & "  )," & vbLf & _
        "  ncol = " & variables.Count & "," & vbLf & "  byrow = TRUE" & vbLf & ")" & vbLf & _
        "constraint_dir <- c(" & vbLf & LpConstraintLines(rows, "dirs") & vbLf & ")" & vbLf & _
        "constraint_rhs <- c(" & vbLf & LpConstraintLines(rows, "rhs") & vbLf & ")" & vbLf & vbLf & _
        "lp_sol <- lp(" & vbLf & "  direction = """ & direction & """," & vbLf & "  objective_fun," & vbLf & "  constraint_mat," & vbLf & _
        "  constraint_dir," & vbLf & "  constraint_rhs," & vbLf & "  compute.sens = " & IIf(sensitivity, "TRUE", "FALSE") & typeOption & vbLf & ")" & vbLf & vbLf & _
        "names(lp_sol$solution) <- c(" & quoted.Join(", ") & ")" & vbLf & _
        "lp_sol$solution" & vbLf & "lp_sol$objval" & vbLf & "lp_sol$status # 0: optimal, 2: infeasible, 3: unbounded" & sensitivityLines
    Q("#lpCode").TextContent = code
    Q("#copyLpCodeBtn").Disabled = False
End Sub
