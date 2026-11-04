Attribute VB_Name = "modApp3"
' Application shell, part 3 (website app-src.js lines 4743-end): number-input steppers,
' the synchronised "Calculate axis intervals" toggles, the theme toggle, the resizable
' sidebar, the mobile library button, typing-to-search and the ranked recipe search.
Option Explicit

' True once the website's second MutationObserver (number steppers) is observing.
Public NumberObserverActive As Boolean

Private Const SIDEBAR_STORAGE_KEY As String = "bt1101-sidebar-width"
Private Const THEME_STORAGE_KEY As String = "bt1101-theme"

Private mResizingSidebar As Boolean

' ================================================================ startup

Public Sub InitShellPartC()
    Dim resizer As clsNode, saved As Variant, savedWidth As Double
    mResizingSidebar = False
    ' plotAxisControls
    AddPlotAxisControl "ggplotCode"
    AddPlotAxisControl "regressionCode"
    AddPlotAxisControl "kmeansCode"
    ' enhanceNumberInputs() and the number-input MutationObserver.
    EnhanceNumberInputs Doc
    NumberObserverActive = True
    ' Theme.
    ApplyTheme ThemeAttribute()
    Q("#themeToggle").AddEventListener "click", "modApp3.OnThemeToggleClick"
    ' Sidebar width.
    Set resizer = Q("#sidebarResizer")
    savedWidth = 0
    saved = StorageGetItem(SIDEBAR_STORAGE_KEY)
    If Not IsNull(saved) Then savedWidth = JsNumber(CStr(saved))
    If JsIsNaN(savedWidth) Then savedWidth = 0
    If savedWidth = 0 Then savedWidth = 310
    SetSidebarWidth savedWidth, False
    resizer.AddEventListener "pointerdown", "modApp3.OnResizerPointerDown"
    Win.AddEventListener "pointermove", "modApp3.OnWindowPointerMove"
    Win.AddEventListener "pointerup", "modApp3.OnWindowPointerStop"
    Win.AddEventListener "pointercancel", "modApp3.OnWindowPointerStop"
    resizer.AddEventListener "keydown", "modApp3.OnResizerKeydown"
    resizer.AddEventListener "dblclick", "modApp3.OnResizerDblClick"
    Win.AddEventListener "resize", "modApp3.OnWindowResize"
    Q("#mobileLibraryBtn").AddEventListener "click", "modApp3.OnMobileLibraryClick"
    ' Typing a letter outside a text field starts a search.
    Win.AddEventListener "keydown", "modApp3.OnWindowKeydown"
    Q("#clearSearch").AddEventListener "click", "modApp3.OnClearSearchClick"
    Q("#searchInput").AddEventListener "input", "modApp3.OnSearchInput"
    Q("#searchInput").AddEventListener "keydown", "modApp3.OnSearchKeydown"
    Q("#searchInput").DispatchEvent NewEvent("input")
End Sub

' ================================================================ number steppers

' The browser's input.stepUp()/stepDown() for a number input. Returns False where the
' browser throws (step="any"), so the caller uses fallbackNumberStep.
Private Function BrowserStep(ByVal el As clsNode, ByVal direction As Long) As Boolean
    Dim stepText As String, stepValue As Double, hasMin As Boolean, hasMax As Boolean
    Dim minValue As Double, maxValue As Double, value As Double, before As Double, base As Double
    Dim k As Double, decimals As Long, attr As Variant
    attr = el.GetAttribute("step")
    stepValue = 1
    If Not IsNull(attr) Then
        stepText = JsLower(JsTrim(CStr(attr)))
        If stepText = "any" Then Exit Function
        If ValidFloat(stepText) Then
            If JsNumber(stepText) > 0 Then stepValue = JsNumber(stepText): decimals = DecimalPlaces(stepText)
        End If
    End If
    hasMin = AttributeNumber(el, "min", minValue)
    hasMax = AttributeNumber(el, "max", maxValue)
    BrowserStep = True
    If hasMin And hasMax Then
        If minValue > maxValue Then Exit Function
    End If
    If ValidFloat(el.Value) Then value = JsNumber(el.Value) Else value = 0
    decimals = JsMax(decimals, DecimalPlaces(el.Value))
    before = value
    If hasMin Then
        base = minValue
        decimals = JsMax(decimals, DecimalPlaces(CStr(el.GetAttribute("min"))))
    ElseIf AttributeNumber(el, "value", base) Then
        decimals = JsMax(decimals, DecimalPlaces(CStr(el.GetAttribute("value"))))
    Else
        base = 0
    End If
    If hasMin And hasMax Then
        If base + JsCeil(RoundTo((minValue - base) / stepValue, 9)) * stepValue > maxValue Then Exit Function
    End If
    k = RoundTo((value - base) / stepValue, 9)
    If k <> Int(k) Then
        If direction > 0 Then value = base + JsCeil(k) * stepValue Else value = base + JsFloor(k) * stepValue
    Else
        value = value + direction * stepValue
    End If
    If hasMin Then
        If value < minValue Then value = base + JsCeil(RoundTo((minValue - base) / stepValue, 9)) * stepValue
    End If
    If hasMax Then
        If value > maxValue Then value = base + JsFloor(RoundTo((maxValue - base) / stepValue, 9)) * stepValue
    End If
    If direction < 0 And value > before Then Exit Function
    If direction > 0 And value < before Then Exit Function
    If decimals > 12 Then decimals = 12
    el.Value = JsNumberToString(JsNumber(JsToFixed(value, decimals)))
End Function

Private Function RoundTo(ByVal x As Double, ByVal digits As Long) As Double
    RoundTo = JsNumber(JsToFixed(x, digits))
End Function

' A valid floating-point number string (the browser's number-input value syntax).
Private Function ValidFloat(ByVal text As String) As Boolean
    ValidFloat = RxTest(text, "^-?(?:\d+(?:\.\d+)?|\.\d+)(?:[eE][-+]?\d+)?$")
End Function

Private Function DecimalPlaces(ByVal text As String) As Long
    Dim m As clsMatch
    Set m = RxMatch(text, "\.(\d+)")
    If Not m Is Nothing Then DecimalPlaces = Len(m.Group(1))
End Function

Private Function AttributeNumber(ByVal el As clsNode, ByVal name As String, ByRef result As Double) As Boolean
    Dim attr As Variant
    attr = el.GetAttribute(name)
    If IsNull(attr) Then Exit Function
    If Not ValidFloat(CStr(attr)) Then Exit Function
    result = JsNumber(CStr(attr))
    AttributeNumber = True
End Function

Private Function AttributeText(ByVal el As clsNode, ByVal name As String) As String
    Dim attr As Variant
    attr = el.GetAttribute(name)
    If Not IsNull(attr) Then AttributeText = CStr(attr)
End Function

Public Sub FallbackNumberStep(ByVal el As clsNode, ByVal direction As Long)
    Dim current As Double, fraction As String, parts As clsList, explicitStep As Double, amount As Double
    Dim minimum As Double, maximum As Double, precision As Long, amountParts As clsList, nextValue As Double
    Dim minText As String, maxText As String
    current = JsNumber(el.Value)
    If JsIsNaN(current) Then current = 0
    Set parts = JsSplit(el.Value, ".")
    If parts.Count > 1 Then fraction = CStr(parts(1))
    explicitStep = JsNumber(AttributeText(el, "step"))
    If JsIsFinite(explicitStep) And JsGt(explicitStep, 0) Then
        amount = explicitStep
    ElseIf Len(fraction) > 0 Then
        amount = JsPow(10, -Len(fraction))
    Else
        amount = 1
    End If
    minText = AttributeText(el, "min")
    maxText = AttributeText(el, "max")
    If Len(minText) = 0 Then minimum = JsInfinity(-1) Else minimum = JsNumber(minText)
    If Len(maxText) = 0 Then maximum = JsInfinity(1) Else maximum = JsNumber(maxText)
    Set amountParts = JsSplit(JsNumberToString(amount), ".")
    precision = Len(fraction)
    If amountParts.Count > 1 Then
        If Len(CStr(amountParts(1))) > precision Then precision = Len(CStr(amountParts(1)))
    End If
    nextValue = JsMin(maximum, JsMax(minimum, JsAdd(current, JsMul(direction, amount))))
    If precision > 12 Then precision = 12
    el.Value = JsNumberToString(JsNumber(JsToFixed(nextValue, precision)))
End Sub

Public Sub StepNumberInput(ByVal el As clsNode, ByVal direction As Long)
    If el.Disabled Or el.ReadOnly Then Exit Sub
    If Not BrowserStep(el, direction) Then FallbackNumberStep el, direction
    el.DispatchEvent NewEvent("input", True)
    el.DispatchEvent NewEvent("change", True)
    el.Focus
End Sub

Public Sub EnhanceNumberInput(ByVal el As clsNode)
    Dim description As String, wrapper As clsNode, controls As clsNode, down As clsNode, up As clsNode, label As clsNode
    If el.Matches(".lp-rhs") Then Exit Sub
    If AttributeText(el, "data-horizontal-stepper") = "true" Then Exit Sub
    If Not el.Closest(".number-stepper") Is Nothing Then Exit Sub
    el.DataSet "horizontalStepper", "true"
    description = AttributeText(el, "aria-label")
    If Len(description) = 0 Then
        Set label = el.Closest("label")
        If Not label Is Nothing Then description = JsTrim(label.TextContent)
    End If
    If Len(description) = 0 Then description = "numeric value"
    Set wrapper = CreateElement("div")
    wrapper.ClassName = "number-stepper"
    Set controls = CreateElement("span")
    controls.ClassName = "number-step-buttons"
    Set down = CreateElement("button")
    down.TypeName = "button"
    down.ClassName = "number-step-button"
    down.TextContent = ChrW$(&H2193)
    down.Title = "Decrease " & description
    down.SetAttribute "aria-label", "Decrease " & description
    Set up = CreateElement("button")
    up.TypeName = "button"
    up.ClassName = "number-step-button"
    up.TextContent = ChrW$(&H2191)
    up.Title = "Increase " & description
    up.SetAttribute "aria-label", "Increase " & description
    el.Before wrapper
    wrapper.Append el, controls
    controls.Append down, up
    down.AddEventListener "click", "modApp3.OnNumberStepDown"
    up.AddEventListener "click", "modApp3.OnNumberStepUp"
End Sub

Public Sub EnhanceNumberInputs(ByVal root As clsNode)
    Dim list As clsList, i As Long
    If root Is Nothing Then Exit Sub
    If root.NodeType = 1 Then
        If root.Matches("input[type=""number""]") Then EnhanceNumberInput root
    End If
    If root.NodeType <> 1 And root.NodeType <> 9 And root.NodeType <> 11 Then Exit Sub
    Set list = root.QuerySelectorAll("input[type=""number""]")
    For i = 0 To list.Count - 1
        EnhanceNumberInput list(i)
    Next i
End Sub

' Microtask: the number-input MutationObserver's callback for one added node.
Public Sub EnhanceNumberInputsTask(ByVal node As clsNode)
    On Error GoTo Failed
    If node.NodeType = 1 Then EnhanceNumberInputs node
    Exit Sub
Failed:
    ReportError "modApp3.EnhanceNumberInputsTask"
End Sub

' The input a stepper button belongs to (the closure's captured input).
Private Function StepperInput(ByVal button As clsNode) As clsNode
    Dim wrapper As clsNode
    Set wrapper = button.Closest(".number-stepper")
    If wrapper Is Nothing Then Exit Function
    Set StepperInput = wrapper.FirstElementChild
End Function

Public Sub OnNumberStepDown(ByVal e As clsEvent)
    Dim el As clsNode
    On Error GoTo Failed
    Set el = StepperInput(e.CurrentTarget)
    If Not el Is Nothing Then StepNumberInput el, -1
    Exit Sub
Failed:
    ReportError "modApp3.OnNumberStepDown"
End Sub

Public Sub OnNumberStepUp(ByVal e As clsEvent)
    Dim el As clsNode
    On Error GoTo Failed
    Set el = StepperInput(e.CurrentTarget)
    If Not el Is Nothing Then StepNumberInput el, 1
    Exit Sub
Failed:
    ReportError "modApp3.OnNumberStepUp"
End Sub

' ================================================================ axis-interval toggles

Private Sub AddPlotAxisControl(ByVal codeId As String)
    Dim control As clsNode, target As clsNode
    Set target = ById(codeId)
    If target Is Nothing Then Exit Sub
    Set control = CreateElement("label")
    control.ClassName = "kmeans-scale-toggle"
    control.SetStyle "display", "flex"
    control.SetStyle "gap", "10px"
    control.SetStyle "padding", "12px 16px"
    control.SetStyle "align-items", "center"
    control.SetStyle "flex-wrap", "wrap"
    control.InnerHtml = "<input type=""checkbox"" role=""switch"" class=""plot-axis-interval-toggle""><span>Calculate axis intervals</span>" & _
        "<small>Off: basic plotting parameters and automatic axes. Applies to all plots.</small>"
    target.Before control
    control.QuerySelector("input").AddEventListener "change", "modApp3.OnPlotAxisToggleChange"
End Sub

Public Sub OnPlotAxisToggleChange(ByVal e As clsEvent)
    Dim list As clsList, i As Long
    On Error GoTo Failed
    CalculatePlotAxisIntervals = e.Target.Checked
    Set list = QA(".plot-axis-interval-toggle")
    For i = 0 To list.Count - 1
        list(i).Checked = CalculatePlotAxisIntervals
    Next i
    ClearCompletionForAnswerVariant ById("copyGgplotCodeBtn")
    TryCallOp "RenderGgplotCode"
    ClearCompletionForAnswerVariant ById("copyRegressionCodeBtn")
    TryCallOp "RenderRegressionCode"
    ClearCompletionForAnswerVariant ById("copyKmeansCodeBtn")
    TryCallOp "RenderKmeansCode"
    RefreshAnswerVariantPreviews
    Exit Sub
Failed:
    ReportError "modApp3.OnPlotAxisToggleChange"
End Sub

' ================================================================ theme

Private Function ThemeAttribute() As String
    Dim value As Variant
    value = DocumentElement.DataGet("theme")
    If Not IsNull(value) Then ThemeAttribute = CStr(value)
End Function

Public Sub ApplyTheme(ByVal theme As String, Optional ByVal persist As Boolean = False)
    Dim isDark As Boolean
    isDark = (theme = "dark")
    DocumentElement.DataSet "theme", IIf(isDark, "dark", "light")
    Q("#themeToggle").SetAttribute "aria-pressed", IIf(isDark, "true", "false")
    Q("#themeToggle .theme-toggle-icon").TextContent = IIf(isDark, ChrW$(&H2600), ChrW$(&H263E))
    If persist Then StorageSetItem THEME_STORAGE_KEY, IIf(isDark, "dark", "light")
End Sub

Public Sub OnThemeToggleClick(ByVal e As clsEvent)
    On Error GoTo Failed
    ApplyTheme IIf(ThemeAttribute() = "dark", "light", "dark"), True
    Exit Sub
Failed:
    ReportError "modApp3.OnThemeToggleClick"
End Sub

' ================================================================ sidebar width

Private Sub SidebarLimits(ByRef minimum As Long, ByRef maximum As Long)
    minimum = 210
    maximum = Int(WindowInnerWidth * 0.45)
    If maximum > 480 Then maximum = 480
    If maximum < 210 Then maximum = 210
End Sub

' setSidebarWidth(requested, persist): requested is a number (0/NaN = default 310).
Public Function SetSidebarWidth(ByVal requested As Double, Optional ByVal persist As Boolean = True) As Long
    Dim minimum As Long, maximum As Long, width As Double, resizer As clsNode
    SidebarLimits minimum, maximum
    If JsIsNaN(requested) Then requested = 0
    If requested = 0 Then requested = 310
    width = JsMathRound(JsMin(maximum, JsMax(minimum, requested)))
    DocumentElement.SetStyle "--sidebar-width", JsNumberToString(width) & "px"
    Set resizer = Q("#sidebarResizer")
    resizer.SetAttribute "aria-valuemin", CStr(minimum)
    resizer.SetAttribute "aria-valuemax", CStr(maximum)
    resizer.SetAttribute "aria-valuenow", JsNumberToString(width)
    If persist Then StorageSetItem SIDEBAR_STORAGE_KEY, JsNumberToString(width)
    SetSidebarWidth = CLng(width)
End Function

Private Function CurrentSidebarWidth() As Double
    CurrentSidebarWidth = JsNumber(AttributeText(Q("#sidebarResizer"), "aria-valuenow"))
    If JsIsNaN(CurrentSidebarWidth) Then CurrentSidebarWidth = 0
    If CurrentSidebarWidth = 0 Then CurrentSidebarWidth = 310
End Function

Public Sub OnResizerPointerDown(ByVal e As clsEvent)
    On Error GoTo Failed
    If WindowInnerWidth <= 720 Then Exit Sub
    mResizingSidebar = True
    Q("#sidebarResizer").AddClass "dragging"
    Body.AddClass "resizing-sidebar"
    e.PreventDefault
    Exit Sub
Failed:
    ReportError "modApp3.OnResizerPointerDown"
End Sub

Public Sub OnWindowPointerMove(ByVal e As clsEvent)
    On Error GoTo Failed
    If mResizingSidebar Then SetSidebarWidth e.ClientX
    Exit Sub
Failed:
    ReportError "modApp3.OnWindowPointerMove"
End Sub

Public Sub OnWindowPointerStop(ByVal e As clsEvent)
    On Error GoTo Failed
    mResizingSidebar = False
    Q("#sidebarResizer").RemoveClass "dragging"
    Body.RemoveClass "resizing-sidebar"
    Exit Sub
Failed:
    ReportError "modApp3.OnWindowPointerStop"
End Sub

Public Function SidebarResizing() As Boolean
    SidebarResizing = mResizingSidebar
End Function

Public Sub OnResizerKeydown(ByVal e As clsEvent)
    Dim current As Double, minimum As Long, maximum As Long, stepSize As Long, nextValue As Double, hasNext As Boolean
    On Error GoTo Failed
    current = CurrentSidebarWidth()
    SidebarLimits minimum, maximum
    If e.ShiftKey Then stepSize = 40 Else stepSize = 16
    Select Case e.Key
        Case "ArrowLeft": nextValue = current - stepSize: hasNext = True
        Case "ArrowRight": nextValue = current + stepSize: hasNext = True
        Case "Home": nextValue = minimum: hasNext = True
        Case "End": nextValue = maximum: hasNext = True
    End Select
    If hasNext Then
        e.PreventDefault
        SetSidebarWidth nextValue
    End If
    Exit Sub
Failed:
    ReportError "modApp3.OnResizerKeydown"
End Sub

Public Sub OnResizerDblClick(ByVal e As clsEvent)
    On Error GoTo Failed
    SetSidebarWidth 310
    Exit Sub
Failed:
    ReportError "modApp3.OnResizerDblClick"
End Sub

Public Sub OnWindowResize(ByVal e As clsEvent)
    On Error GoTo Failed
    SetSidebarWidth CurrentSidebarWidth(), False
    Exit Sub
Failed:
    ReportError "modApp3.OnWindowResize"
End Sub

Public Sub OnMobileLibraryClick(ByVal e As clsEvent)
    Dim isOpen As Boolean
    On Error GoTo Failed
    isOpen = Q("#library").ToggleClass("open")
    e.CurrentTarget.SetAttribute "aria-expanded", IIf(isOpen, "true", "false")
    Exit Sub
Failed:
    ReportError "modApp3.OnMobileLibraryClick"
End Sub

' ================================================================ typing-to-search

Public Function IsSearchTypingTarget(ByVal target As clsNode) As Boolean
    If target Is Nothing Then Exit Function
    If target.NodeType <> 1 Then Exit Function
    If target.Matches("textarea, select") Then IsSearchTypingTarget = True: Exit Function
    If target.TagName <> "INPUT" Then Exit Function
    Select Case target.TypeName
        Case "button", "checkbox", "color", "file", "radio", "range", "reset", "submit"
        Case Else: IsSearchTypingTarget = True
    End Select
End Function

Public Sub OnWindowKeydown(ByVal e As clsEvent)
    Dim search As clsNode
    On Error GoTo Failed
    If e.DefaultPrevented Or e.IsComposing Or e.CtrlKey Or e.MetaKey Or e.AltKey Then Exit Sub
    If Not RxTest(e.Key, "^[a-z]$", "i") Then Exit Sub
    If IsSearchTypingTarget(e.Target) Then Exit Sub
    If Body.HasClass("variable-rename-alert-open") Then Exit Sub
    Set search = Q("#searchInput")
    e.PreventDefault
    If WindowInnerWidth <= 720 And Not Q("#library").HasClass("open") Then Q("#mobileLibraryBtn").Click
    search.Value = search.Value & e.Key
    search.Focus
    search.DispatchEvent NewEvent("input", True)
    Exit Sub
Failed:
    ReportError "modApp3.OnWindowKeydown"
End Sub

Public Sub OnClearSearchClick(ByVal e As clsEvent)
    On Error GoTo Failed
    Q("#searchInput").Value = ""
    Q("#searchInput").DispatchEvent NewEvent("input")
    Q("#searchInput").Focus
    Exit Sub
Failed:
    ReportError "modApp3.OnClearSearchClick"
End Sub

' ================================================================ search ranking

Public Function NormaliseSearchText(ByVal value As String) As String
    NormaliseSearchText = RxReplace(JsTrim(RxReplace(JsLower(value), "[^a-z0-9]+", " ", "g")), "\s+", " ", "g")
End Function

Public Function LevenshteinDistance(ByVal leftText As String, ByVal rightText As String) As Long
    Dim previous() As Long, current() As Long, row As Long, column As Long, n As Long, m As Long
    Dim a As Long, b As Long, c As Long
    If leftText = rightText Then Exit Function
    n = Len(leftText): m = Len(rightText)
    If n = 0 Then LevenshteinDistance = m: Exit Function
    If m = 0 Then LevenshteinDistance = n: Exit Function
    ReDim previous(0 To m)
    ReDim current(0 To m)
    For column = 0 To m
        previous(column) = column
    Next column
    For row = 1 To n
        current(0) = row
        For column = 1 To m
            a = current(column - 1) + 1
            b = previous(column) + 1
            c = previous(column - 1)
            If Mid$(leftText, row, 1) <> Mid$(rightText, column, 1) Then c = c + 1
            If b < a Then a = b
            If c < a Then a = c
            current(column) = a
        Next column
        previous = current
    Next row
    LevenshteinDistance = previous(m)
End Function

Public Function FuzzySearchScore(ByVal query As String, ByVal candidate As String) As Double
    Dim needle As String, haystack As String, words() As String, tokens() As String, i As Long, j As Long
    Dim best As Double, score As Double, token As String, word As String, total As Double
    Dim shorter As Double, longer As Double, ratio As Double
    needle = NormaliseSearchText(query)
    haystack = NormaliseSearchText(candidate)
    If Len(needle) = 0 Then FuzzySearchScore = 1: Exit Function
    If Len(haystack) = 0 Then FuzzySearchScore = 0: Exit Function
    If InStr(1, haystack, needle, vbBinaryCompare) > 0 Or InStr(1, Replace(haystack, " ", ""), Replace(needle, " ", ""), vbBinaryCompare) > 0 Then
        FuzzySearchScore = 1
        Exit Function
    End If
    words = Split(haystack, " ")
    tokens = Split(needle, " ")
    total = 0
    For i = 0 To UBound(tokens)
        token = tokens(i)
        best = 0
        For j = 0 To UBound(words)
            word = words(j)
            If Len(token) < Len(word) Then shorter = Len(token): longer = Len(word) Else shorter = Len(word): longer = Len(token)
            If InStr(1, word, token, vbBinaryCompare) > 0 Or InStr(1, token, word, vbBinaryCompare) > 0 Then
                score = shorter / longer
            Else
                ratio = LevenshteinDistance(token, word) / longer
                score = 1 - ratio
            End If
            If score > best Then best = score
        Next j
        total = total + best
    Next i
    FuzzySearchScore = total / (UBound(tokens) + 1)
End Function

Public Function IsExactSearchMatch(ByVal query As String, ByVal label As String) As Boolean
    Dim needle As String, candidate As String, words() As String, i As Long
    needle = NormaliseSearchText(query)
    candidate = NormaliseSearchText(label)
    If Len(needle) = 0 Or Len(candidate) = 0 Then Exit Function
    If candidate = needle Then IsExactSearchMatch = True: Exit Function
    If InStr(1, needle, " ", vbBinaryCompare) > 0 Then
        IsExactSearchMatch = (InStr(1, candidate, needle, vbBinaryCompare) > 0)
    Else
        words = Split(candidate, " ")
        For i = 0 To UBound(words)
            If words(i) = needle Then IsExactSearchMatch = True: Exit Function
        Next i
    End If
End Function

' True when entry a sorts after entry b (labelScore desc, score desc, index asc).
Private Function SortsAfter(ByRef labelScores() As Double, ByRef scores() As Double, ByVal a As Long, ByVal b As Long) As Boolean
    If labelScores(a) <> labelScores(b) Then SortsAfter = labelScores(a) < labelScores(b): Exit Function
    If scores(a) <> scores(b) Then SortsAfter = scores(a) < scores(b): Exit Function
    SortsAfter = a > b
End Function

' Stable insertion sort of entry indices.
Private Sub SortEntries(ByRef order() As Long, ByVal count As Long, ByRef labelScores() As Double, ByRef scores() As Double)
    Dim i As Long, j As Long, item As Long
    For i = 1 To count - 1
        item = order(i)
        j = i - 1
        Do While j >= 0
            If Not SortsAfter(labelScores, scores, order(j), item) Then Exit Do
            order(j + 1) = order(j)
            j = j - 1
        Loop
        order(j + 1) = item
    Next i
End Sub

Public Sub OnSearchInput(ByVal e As clsEvent)
    Dim query As String, universal As clsNode, ops As clsList, elements() As clsNode, labels() As String
    Dim labelScores() As Double, scores() As Double, exact() As Boolean, count As Long, i As Long
    Dim titleEl As clsNode, minimumScore As Double, order() As Long, matchCount As Long, showingClosest As Boolean
    Dim ranked As clsNode, button As clsNode, el As clsNode, tutorial As String, rawQuery As String, k As Long
    On Error GoTo Failed
    rawQuery = e.Target.Value
    query = NormaliseSearchText(rawQuery)
    Set universal = Q(".universal-entry")
    Set ops = QA("#tutorialList .operation")
    count = ops.Count + 1
    ReDim elements(0 To count - 1): ReDim labels(0 To count - 1): ReDim labelScores(0 To count - 1)
    ReDim scores(0 To count - 1): ReDim exact(0 To count - 1): ReDim order(0 To count - 1)
    Set elements(0) = universal
    For i = 1 To count - 1
        Set elements(i) = ops(i - 1)
    Next i
    For i = 0 To count - 1
        Set el = elements(i)
        Set titleEl = el.QuerySelector(".entry-title")
        labels(i) = ""
        If Not titleEl Is Nothing Then labels(i) = JsTrim(titleEl.TextContent)
        If Len(labels(i)) = 0 Then labels(i) = JsTrim(el.TextContent)
        labelScores(i) = FuzzySearchScore(query, labels(i))
        scores(i) = FuzzySearchScore(query, labels(i) & " " & AttributeText(el, "data-search"))
        exact(i) = IsExactSearchMatch(query, labels(i))
    Next i
    If Len(query) <= 2 Then minimumScore = 0.85 Else minimumScore = 0.62
    matchCount = 0
    For i = 0 To count - 1
        If Len(query) = 0 Or scores(i) >= minimumScore Then order(matchCount) = i: matchCount = matchCount + 1
    Next i
    showingClosest = (Len(query) > 0 And matchCount = 0)
    If showingClosest Then
        For i = 0 To count - 1
            order(i) = i
        Next i
        SortEntries order, count, labelScores, scores
        matchCount = IIf(count < 5, count, 5)
    Else
        SortEntries order, matchCount, labelScores, scores
    End If
    Set ranked = Q("#rankedSearchResults")
    ranked.ReplaceChildren
    For k = 0 To matchCount - 1
        i = order(k)
        Set button = CreateElement("button")
        button.TypeName = "button"
        button.ClassName = "operation ranked-search-result" & IIf(elements(i).HasClass("active"), " active", "")
        button.DataSet "id", AttributeText(elements(i), "data-id")
        tutorial = AttributeText(elements(i), "data-tutorial")
        If Len(tutorial) = 0 Then tutorial = "Smart entry point"
        button.DataSet "tutorial", tutorial
        button.DataSet "searchExact", IIf(Not showingClosest And exact(i), "true", "false")
        button.TextContent = labels(i)
        ranked.AppendChild button
    Next k
    universal.ToggleClass "hidden", Len(query) > 0
    Q("#tutorialList").ToggleClass "hidden", Len(query) > 0
    ranked.ToggleClass "hidden", Len(query) = 0
    Q("#librarySectionLabel").TextContent = IIf(Len(query) > 0, "Search results", "Operations")
    If Len(query) = 0 Then
        universal.RemoveClass "hidden"
        For i = 0 To ops.Count - 1
            ops(i).RemoveClass "hidden"
        Next i
        ranked.ReplaceChildren
        matchCount = count
    End If
    If showingClosest Then
        Q("#resultCount").TextContent = "No matches " & ChrW$(&HB7) & " " & matchCount & " closest"
        Q("#noResults").TextContent = "No matches for " & ChrW$(&H201C) & JsTrim(rawQuery) & ChrW$(&H201D) & ". Showing the " & matchCount & _
            " closest operation" & IIf(matchCount = 1, "", "s") & "."
    Else
        Q("#resultCount").TextContent = matchCount & " operation" & IIf(matchCount = 1, "", "s")
        Q("#noResults").TextContent = "No matches."
    End If
    Q("#noResults").ToggleClass "hidden", Not showingClosest
    Exit Sub
Failed:
    ReportError "modApp3.OnSearchInput"
End Sub

Public Sub OnSearchKeydown(ByVal e As clsEvent)
    Dim firstResult As clsNode
    On Error GoTo Failed
    If e.Key <> "Enter" Then Exit Sub
    Set firstResult = Q("#rankedSearchResults .ranked-search-result")
    If firstResult Is Nothing Then Exit Sub
    If AttributeText(firstResult, "data-search-exact") <> "true" Then Exit Sub
    e.PreventDefault
    firstResult.Click
    Exit Sub
Failed:
    ReportError "modApp3.OnSearchKeydown"
End Sub
