Attribute VB_Name = "modShared"
' Helpers the website's operations share: number and p-value wording, R name quoting,
' axis-interval code, panel completion and toasts. Ported from index.html; each
' procedure keeps the website function's name in PascalCase.
Option Explicit

' state = { id, title, tutorial } and the "Calculate axis intervals" toggle.
Public AppState As clsMap
Public CalculatePlotAxisIntervals As Boolean
' A pending copy that should complete its panel once the "copied" toast appears
' (the website's pendingCollapsibleCopy; set by WriteAnswerText).
Public PendingCopyPanel As clsNode, PendingCopyButton As clsNode, PendingCopyTime As Date
' Recent errors reported by handlers (newest first), for diagnostics and tests.
Public LastErrors As String

Private mToastTimer As Long

Public Sub InitShared()
    Set AppState = NewMap()
    AppState("id") = "universal"
    AppState("title") = "Universal detector"
    AppState("tutorial") = "Smart entry point"
    CalculatePlotAxisIntervals = False
    mToastTimer = 0
End Sub

' Calls a procedure in another operation module by name, so modules compile on their
' own. Returns the procedure's result (Empty for a Sub).
Public Function CallOp(ByVal procedure As String, ParamArray args() As Variant) As Variant
    Select Case UBound(args) - LBound(args) + 1
        Case 0: CallOp = Application.Run(procedure)
        Case 1: CallOp = Application.Run(procedure, args(0))
        Case 2: CallOp = Application.Run(procedure, args(0), args(1))
        Case 3: CallOp = Application.Run(procedure, args(0), args(1), args(2))
        Case 4: CallOp = Application.Run(procedure, args(0), args(1), args(2), args(3))
        Case Else: CallOp = Application.Run(procedure, args(0), args(1), args(2), args(3), args(4))
    End Select
End Function

' Handlers run through Application.Run cannot pass errors to their caller, so each one
' traps its own and calls this. Errors are logged, never shown as VBA dialogs.
Public Sub ReportError(ByVal where As String)
    Dim message As String
    message = where & ": " & Err.Description & " (" & Err.Number & ")"
    Debug.Print "BT1101 Prose Lab error in " & message
    LastErrors = Left$(message & vbLf & LastErrors, 4000)
    Err.Clear
    ' Visible on every platform without a modal dialog.
    On Error Resume Next
    Application.StatusBar = "BT1101 Prose Lab: " & Left$(message, 200)
    Err.Clear
End Sub

' ---------------------------------------------------------------- numbers and p-values

Public Function FormatFinalNumber(ByVal value As Double, Optional ByVal isCurrency As Boolean = False) As String
    Dim absolute As Double
    If Not JsIsFinite(value) Then
        FormatFinalNumber = JsNumberToString(value)
        Exit Function
    End If
    If isCurrency Then
        FormatFinalNumber = JsToFixed(value, 2)
        Exit Function
    End If
    absolute = Abs(value)
    If absolute = 0 Then
        FormatFinalNumber = "0"
    ElseIf absolute > 1 Then
        FormatFinalNumber = JsToFixed(value, 2)
    Else
        FormatFinalNumber = Replace(JsToPrecision(value, 3), "e+", "e")
    End If
End Function

' True: reject H0 (p < alpha). False: fail to reject. Null: the printed bound cannot decide.
Public Function PValueDecision(ByVal operator As String, ByVal pValue As Double, Optional ByVal alpha As Double = 0.05) As Variant
    PValueDecision = Null
    If Not JsIsFinite(pValue) Then Exit Function
    If pValue < 0 Or pValue > 1 Then Exit Function
    If JsIsNaN(alpha) Then Exit Function
    If Not (alpha > 0 And alpha < 1) Then Exit Function
    Select Case operator
        Case "="
            PValueDecision = (pValue < alpha)
        Case "<"
            If pValue > 0 And pValue <= alpha Then PValueDecision = True
        Case "<="
            If pValue < alpha Then PValueDecision = True
        Case ">", ">="
            If pValue >= alpha And Not (operator = ">" And pValue = 1) Then PValueDecision = False
    End Select
End Function

Public Function PValueComparison(ByVal operator As String, ByVal pValue As Double, Optional ByVal alpha As Double = 0.05) As String
    Dim decision As Variant
    decision = PValueDecision(operator, pValue, alpha)
    If IsNull(decision) Then
        PValueComparison = "not determined relative to"
    ElseIf decision Then
        PValueComparison = "less than"
    ElseIf operator = "=" And pValue = alpha Then
        PValueComparison = "equal to"
    ElseIf operator = ">=" And pValue = alpha Then
        PValueComparison = "greater than or equal to"
    Else
        PValueComparison = "greater than"
    End If
End Function

Public Function PValueDecisionError(ByVal operator As String, ByVal pValue As Double, Optional ByVal alpha As Double = 0.05) As String
    PValueDecisionError = "The reported p-value " & operator & " " & JsNumberToString(pValue) & " does not establish a valid decision at the " & _
        JsNumberToString(alpha) & " significance cutoff. Paste a more precise p-value."
End Function

Public Function FormatPValue(ByVal value As Double, Optional ByVal alpha As Double = 0.05) As String
    Dim formatted As String
    formatted = FormatFinalNumber(value)
    If JsEq(JsNumber(formatted), alpha) And JsNe(value, alpha) Then FormatPValue = JsNumberToString(value) Else FormatPValue = formatted
End Function

' ---------------------------------------------------------------- R names and labels

Public Function RColumnName(ByVal name As String) As String
    Dim trimmed As String
    trimmed = JsTrim(name)
    If RxTest(trimmed, "^[A-Za-z.][A-Za-z0-9._]*$") And Not RxTest(trimmed, "^\.\d") Then
        RColumnName = trimmed
    Else
        RColumnName = "`" & Replace(trimmed, "`", "\`") & "`"
    End If
End Function

Public Function RString(ByVal value As String) As String
    RString = Replace(Replace(value, "\", "\\"), """", "\""")
End Function

' Chart titles, axis labels and legends: underscores become spaces, each word capitalised.
Public Function ChartDisplayLabel(ByVal value As Variant) As String
    Dim text As String, out As String, i As Long, ch As String, previous As String
    If Not IsNull(value) And Not IsEmpty(value) Then text = CStr(value)
    text = RxReplace(text, "_+", " ", "g")
    text = JsTrim(RxReplace(text, "\s+", " ", "g"))
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If i = 1 Then
            If RxTest(ch, "^\p{L}$", "u") Then ch = UCase$(ch)
        ElseIf previous = " " Or previous = "," Or previous = "(" Or RxTest(previous, "^\s$") Then
            If RxTest(ch, "^\p{L}$", "u") Then ch = UCase$(ch)
        End If
        out = out & ch
        previous = Mid$(text, i, 1)
    Next i
    ChartDisplayLabel = out
End Function

Public Function PredictorList(ByVal value As String) As clsList
    Dim parts As clsList, out As clsList, i As Long, item As String
    Set parts = JsSplit(value, ",")
    Set out = NewList()
    For i = 0 To parts.Count - 1
        item = JsTrim(CStr(parts(i)))
        If Len(item) > 0 Then out.Push item
    Next i
    Set PredictorList = out
End Function

' ---------------------------------------------------------------- axis interval code

Public Function BoundedTickCode(ByVal stem As String, ByVal rangeExpression As String, Optional ByVal headroom As Boolean = False, Optional ByVal wholeNumbers As Boolean = False) As String
    Dim stepCode As String
    If Not CalculatePlotAxisIntervals Then Exit Function
    If wholeNumbers Then stepCode = "max(1, diff(" & stem & "_ticks)[1])" Else stepCode = "diff(" & stem & "_ticks)[1]"
    BoundedTickCode = stem & "_range <- range(" & rangeExpression & ", finite = TRUE)" & vbLf & _
        stem & "_ticks <- pretty(" & stem & "_range, n = 5)" & vbLf & _
        stem & "_step <- " & stepCode & vbLf & _
        stem & "_range <- c(floor(" & stem & "_range[1] / " & stem & "_step), " & IIf(headroom, "ceiling", "floor") & "(" & stem & "_range[2] / " & stem & "_step) + 1) * " & stem & "_step" & vbLf & _
        stem & "_ticks <- seq(" & stem & "_range[1], " & stem & "_range[2], by = " & stem & "_step)"
End Function

Public Function CoordinateTickCode(ByVal stem As String, ByVal expression As String) As String
    CoordinateTickCode = stem & "_values <- " & expression
    If CalculatePlotAxisIntervals Then CoordinateTickCode = CoordinateTickCode & vbLf & stem & "_range <- range(pretty(" & stem & "_values, n = 5))"
End Function

Public Function CoordinatePlotMarginCode() As String
    CoordinatePlotMarginCode = ""
End Function

Public Function CoordinateAxisCode() As String
    CoordinateAxisCode = ""
End Function

Public Function NumericAxisArguments(ByVal stem As String, ByVal axis As String, Optional ByVal automaticRange As String = "") As String
    If Not CalculatePlotAxisIntervals Then
        If Len(automaticRange) > 0 Then NumericAxisArguments = axis & "lim = " & automaticRange & ","
        Exit Function
    End If
    NumericAxisArguments = axis & "lim = " & stem & "_range,"
End Function

Public Function NumericAxisCode(Optional ByVal stem As String = "", Optional ByVal axis As String = "", Optional ByVal extra As Variant) As String
    NumericAxisCode = ""
End Function

Public Function CoordinateAxisArguments() As String
    If CalculatePlotAxisIntervals Then CoordinateAxisArguments = "xlim = x_range, ylim = y_range"
End Function

' ---------------------------------------------------------------- panels and toasts

Public Sub SetLpPanelCollapsed(ByVal panel As clsNode, ByVal collapsed As Boolean)
    Dim button As clsNode, strong As clsNode, title As String
    If panel Is Nothing Then Exit Sub
    panel.ToggleClass "is-collapsed", collapsed
    Set button = panel.QuerySelector(":scope > .panel-head .lp-collapse-toggle")
    Set strong = panel.QuerySelector(":scope > .panel-head .panel-title strong")
    title = "section"
    If Not strong Is Nothing Then
        If Len(JsTrim(strong.TextContent)) > 0 Then title = JsTrim(strong.TextContent)
    End If
    If Not button Is Nothing Then
        button.SetAttribute "aria-expanded", IIf(collapsed, "false", "true")
        button.SetAttribute "aria-label", IIf(collapsed, "Expand ", "Collapse ") & title
    End If
End Sub

Public Sub SetPanelCompletion(ByVal panel As clsNode, ByVal complete As Boolean)
    Dim actions As clsNode, badge As clsNode
    If panel Is Nothing Then Exit Sub
    If Not panel.HasClass("lp-collapsible") Then Exit Sub
    panel.ToggleClass "is-complete", complete
    Set actions = panel.QuerySelector(":scope > .panel-head .panel-actions")
    If Not actions Is Nothing Then Set badge = actions.QuerySelector(":scope > .completion-badge")
    If complete And Not actions Is Nothing And badge Is Nothing Then
        Set badge = CreateElement("span")
        badge.ClassName = "completion-badge"
        badge.TextContent = "Done!"
        badge.SetAttribute "role", "status"
        badge.SetAttribute "aria-label", "Copy completed"
        actions.Prepend badge
    ElseIf Not complete Then
        If Not badge Is Nothing Then badge.Remove
    End If
End Sub

Public Function ShouldCompletePanelOnCopy(ByVal panel As clsNode, ByVal button As clsNode) As Boolean
    Dim required As String
    Select Case panel.Id
        Case "multiRegressionInterpretationPanel": required = "copyMultiRegressionProseBtn"
        Case "glmInterpretationPanel": required = "copyGlmProseBtn"
        Case "sensitivityInterpretationPanel": required = "copySensitivityProseBtn"
    End Select
    ShouldCompletePanelOnCopy = (Len(required) = 0 Or button.Id = required)
End Function

Public Sub ShowToast(ByVal message As String, Optional ByVal duration As Long = 1800)
    Dim toast As clsNode, completedAnova As Variant
    RequestAnimationFrame "modShared.FitCompactTextareasLater"
    If RxTest(message, "copied", "i") And Not PendingCopyPanel Is Nothing Then
        If (Now - PendingCopyTime) * 86400# < 10 Then
            completedAnova = CallOp("CompleteAnovaVarianceCopy", PendingCopyButton)
            If Not CBool(completedAnova) Then
                SetPanelCompletion PendingCopyPanel, True
                SetLpPanelCollapsed PendingCopyPanel, True
            End If
        End If
        Set PendingCopyPanel = Nothing
        Set PendingCopyButton = Nothing
    ElseIf RxTest(message, "clipboard access is unavailable", "i") Then
        Set PendingCopyPanel = Nothing
        Set PendingCopyButton = Nothing
    End If
    Set toast = Q("#toast")
    If toast Is Nothing Then Exit Sub
    toast.TextContent = message
    toast.AddClass "show"
    If mToastTimer <> 0 Then ClearTimeout mToastTimer
    mToastTimer = SetTimeout("modShared.HideToast", duration)
End Sub

' showToast's requestAnimationFrame(fitAllCompactTextareas); the shell owns the fitting.
Public Sub FitCompactTextareasLater()
    On Error GoTo Failed
    CallOp "FitAllCompactTextareas"
    Exit Sub
Failed:
    If Err.Number = 1004 Then Err.Clear Else ReportError "modShared.FitCompactTextareasLater"
End Sub

Public Sub HideToast()
    Dim toast As clsNode
    On Error GoTo Failed
    mToastTimer = 0
    Set toast = Q("#toast")
    If Not toast Is Nothing Then toast.RemoveClass "show"
    Exit Sub
Failed:
    ReportError "modShared.HideToast"
End Sub
