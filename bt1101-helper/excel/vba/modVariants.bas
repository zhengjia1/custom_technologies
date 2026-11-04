Attribute VB_Name = "modVariants"
' Answer options: six versions of every copyable prose/code answer (Option 1 canonical).
' Port of the website's answer-variant engine (app-src.js lines 71-993): mega groups,
' the "Answer option" select inserted before copy buttons, per-node caches of the six
' versions, regeneration of Options 2-6 on reselection, previews, writeAnswerText
' (clipboard, plot-axis warning, completion tracking), per-operation selection memory and
' the document-level listeners. The generators live in modVariants2 (R code) and
' modVariants3 (prose). WeakMaps keyed by DOM nodes become clsMaps keyed by node Uid.
Option Explicit
Option Compare Binary

Public ActiveAnswerCopyButton As clsNode
Public UserEditedOperations As clsMap           ' operation id -> True
Public LockedOperationAnswerSelections As clsMap ' operation id -> clsMap(copy key -> option)

Private mOperationAnswerSelections As clsMap
Private mLastVisitAnswerSelections As clsMap
Private mCanonicalText As clsMap   ' "n<uid>" -> canonical text (answerVariantCanonicalText)
Private mRenderedText As clsMap    ' "n<uid>" -> last rendered text (answerVariantRenderedText)
Private mCodeCache As clsMap       ' "n<uid>" -> clsMap { canonical, variants, history }
Private mProseCache As clsMap
Private mSelectButton As clsMap    ' "s<uid>" -> the copy button its change listener closes over
Private mSelectGroup As clsMap     ' "s<uid>" -> mega-group index (-1 = none)
Private mNodes As clsMap           ' "n<uid>" -> node, for pruning entries of discarded nodes
Private mPruneAt As Long            ' prune when mNodes reaches this size
' answerVariantGroup results by button uid, valid while modDom.DomStructVersion is
' unchanged (group membership depends only on ids, classes, attributes and the tree).
Private mGroupMemo As clsMap
Private mGroupMemoVersion As Long
' data-copy-for -> first answer-option control, valid for one structural version.
Private mPickerMemo As clsMap
Private mPickerMemoVersion As Long

' megaAnswerGroups
Private mGroupCount As Long
Private mGroupParent() As String
Private mGroupScope() As String
Private mGroupChildren() As String
Private mGroupKind() As String
Private mGroupAggregate() As Boolean
Private mGroupResetTogether() As Boolean

Private Const SECTION_MARKER As String = "# __BT1101_SECTION_"

Public Sub InitVariants()
    Set ActiveAnswerCopyButton = Nothing
    Set UserEditedOperations = NewMap()
    Set LockedOperationAnswerSelections = NewMap()
    Set mOperationAnswerSelections = NewMap()
    Set mLastVisitAnswerSelections = NewMap()
    Set mCanonicalText = NewMap()
    Set mRenderedText = NewMap()
    Set mCodeCache = NewMap()
    Set mProseCache = NewMap()
    Set mSelectButton = NewMap()
    Set mSelectGroup = NewMap()
    Set mNodes = NewMap()
    mPruneAt = 400
    InitGroups
End Sub

Private Sub EnsureState()
    If mCodeCache Is Nothing Then InitVariants
End Sub

Private Sub InitGroups()
    mGroupCount = 0
    ReDim mGroupParent(0 To 11): ReDim mGroupScope(0 To 11): ReDim mGroupChildren(0 To 11)
    ReDim mGroupKind(0 To 11): ReDim mGroupAggregate(0 To 11): ReDim mGroupResetTogether(0 To 11)
    AddGroup "copyIntervalCodeBtn", "#intervalBench", "#copyIntervalCodeBtn, #copyIntervalProseBtn", "mixed", False, False
    AddGroup "copyAnovaVarianceCodeBtn", "#anovaVariancePanel", "#copyAnovaVarianceCodeBtn, #copyAnovaVarianceProseBtn", "mixed", False, True
    AddGroup "copyAnovaStandardCodeBtn", "#anovaStandardPanel", "#copyAnovaStandardCodeBtn, #copyAnovaStandardProseBtn", "mixed", False, True
    AddGroup "copyAnovaWelchCodeBtn", "#anovaWelchPanel", "#copyAnovaWelchCodeBtn, #copyAnovaWelchProseBtn", "mixed", False, True
    AddGroup "copyPcaCodeBtn", "#pcaCodeSections", "[data-copy-pca-section]", "code", True, False
    AddGroup "copyAdvancedRegressionCodeBtn", "#advancedRegressionBench", "#copyAdvancedRegressionCodeBtn, #copyAdvancedRegressionPredictionCodeBtn, #copyAdvancedRegressionConfidenceCodeBtn", "code", False, False
    AddGroup "copySensitivityProseBtn", "#sensitivityInterpretation", "[data-sensitivity-copy]", "prose", True, False
    AddGroup "copyPcaLinearBtn", "#pcaLinearProse", "[data-copy-pca-linear]", "prose", True, False
    AddGroup "copyGlmProseBtn", "#glmProse", "[data-copy-glm]", "prose", True, False
    AddGroup "copyMultiRegressionProseBtn", "#multiRegressionProse", "[data-copy-multi-model], [data-copy-multi-regression]", "prose", True, False
    AddGroup "copyConfusionProseBtn", "#confusionProse", "[data-copy-confusion]", "prose", True, False
    AddGroup "copyCoefficientIntervalBtn", "#coefficientIntervalProse", "[data-copy-coefficient-interval]", "prose", True, False
End Sub

Private Sub AddGroup(ByVal parentId As String, ByVal scope As String, ByVal children As String, ByVal kind As String, _
    ByVal aggregateCopy As Boolean, ByVal resetTogether As Boolean)
    mGroupParent(mGroupCount) = parentId
    mGroupScope(mGroupCount) = scope
    mGroupChildren(mGroupCount) = children
    mGroupKind(mGroupCount) = kind
    mGroupAggregate(mGroupCount) = aggregateCopy
    mGroupResetTogether(mGroupCount) = resetTogether
    mGroupCount = mGroupCount + 1
End Sub

' ================================================================ generators (public API)

' proseAnswerVariants: six prose versions (entry 0 = canonical).
Public Function ProseAnswerVariants(ByVal canonical As String, Optional ByVal history As clsList = Nothing) As clsList
    Set ProseAnswerVariants = VarProseAnswerVariants(canonical, history)
End Function

' codeAnswerVariants: six R-code versions (entry 0 = canonical); history = the previously
' displayed versions, whose argument orders a regenerated option avoids.
Public Function CodeAnswerVariants(ByVal canonical As String, Optional ByVal history As clsList = Nothing) As clsList
    Set CodeAnswerVariants = VarCodeAnswerVariants(canonical, history)
End Function

' rTokens(code): clsList of clsMap { text, start, end, parent[, open][, close] } (0-based).
Public Function RTokens(ByVal code As String) As clsList
    Set RTokens = VarTokenList(code)
End Function

' ================================================================ small helpers

' Cache key of a node (the website's WeakMaps are keyed by the node itself).
Private Function Key(ByVal node As clsNode) As String
    Dim k As String
    k = "n" & CStr(node.Uid)
    If Not mNodes.Exists(k) Then
        If mNodes.Count >= mPruneAt Then PruneDiscardedNodes
        Set mNodes(k) = node
    End If
    Key = k
End Function

' Drops the cache entries of nodes no longer in the document (re-rendered output nodes),
' which the website's WeakMaps release through garbage collection.
Private Sub PruneDiscardedNodes()
    Dim k As Variant, node As clsNode, ks As clsList
    Set ks = mNodes.Keys()
    For Each k In ks
        Set node = mNodes(CStr(k))
        If Not node.IsConnected Then
            mNodes.Remove CStr(k)
            mCanonicalText.Remove CStr(k)
            mRenderedText.Remove CStr(k)
            mCodeCache.Remove CStr(k)
            mProseCache.Remove CStr(k)
        End If
    Next k
    mPruneAt = 2 * mNodes.Count
    If mPruneAt < 400 Then mPruneAt = 400
End Sub

Private Function DataText(ByVal node As clsNode, ByVal camelKey As String) As String
    Dim value As Variant
    If node Is Nothing Then Exit Function
    If node.NodeType <> 1 Then Exit Function
    value = node.DataGet(camelKey)
    If Not IsNull(value) Then DataText = CStr(value)
End Function

Private Function AttrText(ByVal node As clsNode, ByVal nm As String) As String
    Dim value As Variant
    value = node.GetAttribute(nm)
    If Not IsNull(value) Then AttrText = CStr(value)
End Function

Private Function IsElement(ByVal node As clsNode) As Boolean
    If node Is Nothing Then Exit Function
    IsElement = (node.NodeType = 1)
End Function

Private Function IsButton(ByVal node As clsNode) As Boolean
    If Not IsElement(node) Then Exit Function
    IsButton = (node.LocalName = "button")
End Function

' Math.max(1, Math.min(6, Number(select?.value) || 1))
Private Function OptionOf(ByVal sel As clsNode) As Long
    Dim n As Double
    If sel Is Nothing Then OptionOf = 1: Exit Function
    n = JsNumber(sel.Value)
    If JsIsNaN(n) Then n = 1
    If n = 0 Then n = 1
    If n > 6 Then n = 6
    If n < 1 Then n = 1
    OptionOf = Int(n)
End Function

Private Function NumberOr1(ByVal text As String) As Long
    Dim n As Double
    n = JsNumber(text)
    If JsIsNaN(n) Then n = 1
    If n = 0 Then n = 1
    If n > 2147483647# Or n < -2147483647# Then n = 1
    NumberOr1 = Int(n)
End Function

' Runs a procedure of another module; False when it is absent or fails.
Private Function TryRun(ByVal procedure As String, ParamArray args() As Variant) As Boolean
    On Error GoTo Failed
    Select Case UBound(args) - LBound(args) + 1
        Case 0: Application.Run procedure
        Case 1: Application.Run procedure, args(0)
        Case Else: Application.Run procedure, args(0), args(1)
    End Select
    TryRun = True
    Exit Function
Failed:
    Err.Clear
End Function

' ================================================================ groups and owners

' answerVariantGroup(button): group index, -1 for none.
Private Function GroupOf(ByVal button As clsNode) As Long
    Dim g As Long, v As Long, k As String
    GroupOf = -1
    If Not IsElement(button) Then Exit Function
    v = DomStructVersion()
    If mGroupMemo Is Nothing Or mGroupMemoVersion <> v Then
        Set mGroupMemo = NewMap()
        mGroupMemoVersion = v
    End If
    k = CStr(button.Uid)
    If mGroupMemo.Exists(k) Then GroupOf = mGroupMemo(k): Exit Function
    g = GroupOfUncached(button)
    mGroupMemo(k) = g
    GroupOf = g
End Function

Private Function GroupOfUncached(ByVal button As clsNode) As Long
    Dim g As Long
    GroupOfUncached = -1
    For g = 0 To mGroupCount - 1
        If button.Id = mGroupParent(g) Then GroupOfUncached = g: Exit Function
        If button.Matches(mGroupChildren(g)) Then
            If Not button.Closest(mGroupScope(g)) Is Nothing Then GroupOfUncached = g: Exit Function
        End If
    Next g
End Function

' answerVariantOwner(button): the group's parent button, else the button itself.
Private Function OwnerOf(ByVal button As clsNode) As clsNode
    Dim g As Long
    g = GroupOf(button)
    If g >= 0 Then Set OwnerOf = ById(mGroupParent(g)) Else Set OwnerOf = button
End Function

Private Function IsAnswerCopyButton(ByVal button As clsNode) As Boolean
    Dim label As String
    If Not IsButton(button) Then Exit Function
    label = JsTrim(button.TextContent & " " & button.Title & " " & AttrText(button, "aria-label"))
    If button.Id = "copyLpAsciiBtn" Then Exit Function
    If RxTest(label, "\bCopy (?:ASCII )?table\b", "i") Then Exit Function
    If Not button.Closest("table") Is Nothing Then Exit Function
    IsAnswerCopyButton = RxTest(label, "^Copy\b", "i") Or RxTest(label, "\bCopy (?:prose|code|all)\b", "i")
End Function

' answerVariantKind(button, text): "code" or "prose".
Private Function KindOf(ByVal button As clsNode, Optional ByVal text As String = "") As String
    Dim identity As String
    KindOf = "code"
    If IsElement(button) Then
        identity = JsLower(button.Id & " " & button.TextContent)
        If InStr(1, identity, "code", vbBinaryCompare) > 0 Then Exit Function
        If button.Matches("[data-copy-pca-section]") Then Exit Function
        If Not button.Closest(".pca-code-block") Is Nothing Then Exit Function
    End If
    If RxTest(vbLf & text, "\n\s*(?:[A-Za-z.][A-Za-z0-9._]*\s*<-|library\(|[A-Za-z.][A-Za-z0-9._]*\s*\([^)]*(?:data|formula|method)\s*=)", "") Then Exit Function
    KindOf = "prose"
End Function

' answerVariantOutputNodes(button)
Private Function OutputNodes(ByVal button As clsNode) As clsList
    Dim section As clsNode, panel As clsNode
    If Not IsElement(button) Then Set OutputNodes = NewList(): Exit Function
    Set section = button.Closest(".pca-code-block")
    If Not section Is Nothing Then Set OutputNodes = section.QuerySelectorAll(".generated-code"): Exit Function
    Set section = button.Closest(".sensitivity-prose-card, .prose-card")
    If Not section Is Nothing Then Set OutputNodes = section.QuerySelectorAll("p"): Exit Function
    Set panel = button.Closest(".panel")
    If panel Is Nothing Then Set OutputNodes = NewList(): Exit Function
    If KindOf(button) = "code" Then
        Set OutputNodes = panel.QuerySelectorAll(".generated-code, .interval-code")
    Else
        Set OutputNodes = panel.QuerySelectorAll(".prose-card p, .pca-variance-prose:not(.is-error), .sensitivity-prose-card p")
    End If
End Function

' groupedAnswerVariantEntries(group): parallel lists of buttons and output nodes.
Private Sub GroupedEntries(ByVal g As Long, ByRef buttons As clsList, ByRef nodes As clsList)
    Dim scope As clsNode, list As clsList, i As Long, j As Long, button As clsNode, outs As clsList
    Set buttons = NewList()
    Set nodes = NewList()
    If g < 0 Then Exit Sub
    Set scope = Q(mGroupScope(g))
    If scope Is Nothing Then Exit Sub
    Set list = scope.QuerySelectorAll(mGroupChildren(g))
    For i = 0 To list.Count - 1
        Set button = list(i)
        If Not button.Disabled And button.Closest(".hidden") Is Nothing Then
            Set outs = OutputNodes(button)
            For j = 0 To outs.Count - 1
                buttons.Push button
                nodes.Push outs(j)
            Next j
        End If
    Next i
End Sub

Private Function GroupedNodes(ByVal g As Long) As clsList
    Dim buttons As clsList, nodes As clsList
    GroupedEntries g, buttons, nodes
    Set GroupedNodes = nodes
End Function

' ================================================================ caches

Private Function CachedVariants(ByVal cache As clsMap, ByVal node As clsNode, ByVal canonical As String, ByVal isCode As Boolean) As clsList
    Dim stored As clsMap, k As String, variants As clsList
    k = Key(node)
    If cache.Exists(k) Then
        Set stored = cache(k)
        If CStr(stored("canonical")) = canonical Then Set CachedVariants = stored("variants"): Exit Function
    End If
    If isCode Then Set variants = VarCodeAnswerVariants(canonical) Else Set variants = VarProseAnswerVariants(canonical)
    Set stored = NewMap()
    stored("canonical") = canonical
    Set stored("variants") = variants
    Set cache(k) = stored
    Set CachedVariants = variants
End Function

Private Function CachedCode(ByVal node As clsNode, ByVal canonical As String) As clsList
    Set CachedCode = CachedVariants(mCodeCache, node, canonical, True)
End Function

Private Function CachedProse(ByVal node As clsNode, ByVal canonical As String) As clsList
    Set CachedProse = CachedVariants(mProseCache, node, canonical, False)
End Function

' regenerateSelectedAnswerVariant(button, option, previousOption)
Private Sub RegenerateSelectedAnswerVariant(ByVal button As clsNode, ByVal choice As Long, ByVal previousOption As Long)
    Dim owner As clsNode, g As Long, keys As clsList, seen As clsMap, list As clsList, i As Long, node As clsNode, pass As Long
    If choice = 1 Then Exit Sub
    Set owner = OwnerOf(button)
    g = GroupOf(button)
    Set keys = NewList()
    Set seen = NewMap()
    If Not owner Is Nothing Then keys.Push owner: seen(Key(owner)) = True
    If g >= 0 Then Set list = GroupedNodes(g) Else Set list = OutputNodes(owner)
    For i = 0 To list.Count - 1
        Set node = list(i)
        If Not seen.Exists(Key(node)) Then keys.Push node: seen(Key(node)) = True
    Next i
    For i = 0 To keys.Count - 1
        For pass = 0 To 1
            If pass = 0 Then RegenerateEntry mCodeCache, keys(i), True, choice, previousOption Else RegenerateEntry mProseCache, keys(i), False, choice, previousOption
        Next pass
    Next i
End Sub

Private Sub RegenerateEntry(ByVal cache As clsMap, ByVal node As clsNode, ByVal isCode As Boolean, ByVal choice As Long, ByVal previousOption As Long)
    Dim stored As clsMap, history As clsList, previous As Variant, visible As Variant, variants As clsList
    Dim k As String, different As Long, i As Long, tmp As Variant, fresh As clsMap
    k = Key(node)
    If Not cache.Exists(k) Then Exit Sub
    Set stored = cache(k)
    ' Remember each slot's last wording across intervening selections.
    If stored.Exists("history") Then Set history = stored("history") Else Set history = stored("variants").Clone
    previous = ListItem(history, choice - 1)
    visible = ListItem(stored("variants"), previousOption - 1)
    If isCode Then
        Set variants = VarCodeAnswerVariants(CStr(stored("canonical")), history)
    Else
        Set variants = VarProseAnswerVariants(CStr(stored("canonical")), history)
    End If
    If SameText(variants(choice - 1), previous) Or SameText(variants(choice - 1), visible) Then
        different = -1
        For i = 1 To variants.Count - 1
            If Not SameText(variants(i), previous) And Not SameText(variants(i), visible) Then different = i: Exit For
        Next i
        If different > 0 Then
            tmp = variants(choice - 1)
            variants(choice - 1) = variants(different)
            variants(different) = tmp
        End If
    End If
    If choice - 1 < history.Count Then history(choice - 1) = variants(choice - 1)
    Set fresh = NewMap()
    fresh("canonical") = stored("canonical")
    Set fresh("variants") = variants
    Set fresh("history") = history
    Set cache(k) = fresh
End Sub

Private Function ListItem(ByVal list As clsList, ByVal index As Long) As Variant
    If index >= 0 And index < list.Count Then ListItem = list(index) Else ListItem = Empty
End Function

' JS === between a string and a possibly undefined value.
Private Function SameText(ByVal a As Variant, ByVal b As Variant) As Boolean
    If IsEmpty(a) Or IsEmpty(b) Then SameText = (IsEmpty(a) And IsEmpty(b)): Exit Function
    SameText = (CStr(a) = CStr(b))
End Function

' ================================================================ pickers and previews

Public Function PickerForCopyButton(ByVal button As clsNode) As clsNode
    Dim owner As clsNode, k As String, list As clsList, i As Long, node As clsNode, v As Variant
    EnsureState
    Set owner = OwnerOf(button)
    If owner Is Nothing Then Exit Function
    k = owner.Id
    If Len(k) = 0 Then k = DataText(owner, "answerVariantId")
    If Len(k) = 0 Then Exit Function
    If CssEscape(k) <> k Then
        Set PickerForCopyButton = Q(".answer-variant-control[data-copy-for=""" & CssEscape(k) & """]")
        Exit Function
    End If
    ' Same as the query above: the first control (document order) whose data-copy-for is
    ' k, from one pass over the (cached) list of every control, remembered by k until
    ' the next structural change.
    If mPickerMemo Is Nothing Or mPickerMemoVersion <> DomStructVersion() Then
        Set mPickerMemo = NewMap()
        mPickerMemoVersion = DomStructVersion()
        Set list = QAReadOnly(".answer-variant-control")
        For i = 0 To list.Count - 1
            Set node = list(i)
            v = node.DataGet("copyFor")
            If Not IsNull(v) Then
                If Not mPickerMemo.Exists(CStr(v)) Then Set mPickerMemo(CStr(v)) = node
            End If
        Next i
    End If
    If mPickerMemo.Exists(k) Then Set PickerForCopyButton = mPickerMemo(k)
End Function

' selectedAnswerVariant(text, button)
Private Function SelectedAnswerVariant(ByVal text As String, ByVal button As clsNode) As String
    Dim picker As clsNode, choice As Long, nodes As clsList, cacheKey As clsNode, variants As clsList
    Set picker = PickerForCopyButton(button)
    If picker Is Nothing Then SelectedAnswerVariant = text: Exit Function
    choice = OptionOf(picker.QuerySelector("select"))
    Set nodes = OutputNodes(button)
    If nodes.Count > 0 Then Set cacheKey = nodes(0) Else Set cacheKey = button
    If KindOf(button, text) = "code" Then Set variants = CachedCode(cacheKey, text) Else Set variants = CachedProse(cacheKey, text)
    SelectedAnswerVariant = variants(choice - 1)
End Function

' applyAnswerVariantPreview(button)
Private Sub ApplyAnswerVariantPreview(ByVal button As clsNode)
    Dim owner As clsNode, g As Long, picker As clsNode, sel As clsNode, buttons As clsList, nodes As clsList
    Dim canonicalValues As clsList, i As Long, node As clsNode, current As String, k As String, choice As Long
    Dim combined As String, parts As clsList, rendered As String, variants As clsList
    Set owner = OwnerOf(button)
    g = GroupOf(button)
    If g >= 0 And Not (button Is owner) Then Exit Sub
    Set picker = PickerForCopyButton(owner)
    If picker Is Nothing Then Exit Sub
    picker.ToggleClass "hidden", owner.HasClass("hidden")
    Set sel = picker.QuerySelector("select")
    If Not sel Is Nothing Then sel.Disabled = owner.Disabled
    If g >= 0 Then
        GroupedEntries g, buttons, nodes
    Else
        Set nodes = OutputNodes(owner)
    End If
    Set canonicalValues = NewList()
    For i = 0 To nodes.Count - 1
        Set node = nodes(i)
        current = node.TextContent
        k = Key(node)
        If Not mCanonicalText.Exists(k) Or Not mRenderedText.Exists(k) Then
            mCanonicalText(k) = current
        ElseIf current <> CStr(mRenderedText(k)) Then
            mCanonicalText(k) = current
        End If
        canonicalValues.Push CStr(mCanonicalText(k))
    Next i
    choice = OptionOf(sel)
    If g >= 0 Then
        If mGroupKind(g) = "code" And canonicalValues.Count > 1 And Not owner.Disabled Then
            For i = 0 To canonicalValues.Count - 1
                If i > 0 Then combined = combined & vbLf
                combined = combined & canonicalValues(i)
                If i < canonicalValues.Count - 1 Then combined = combined & vbLf & SECTION_MARKER & CStr(i + 1) & "__"
            Next i
            Set parts = Rx("\n# __BT1101_SECTION_\d+__\n", "").Split(CStr(CachedCode(owner, combined)(choice - 1)))
            For i = 0 To nodes.Count - 1
                Set node = nodes(i)
                If i < parts.Count Then rendered = parts(i) Else rendered = canonicalValues(i)
                If node.TextContent <> rendered Then node.TextContent = rendered
                mRenderedText(Key(node)) = rendered
            Next i
            Exit Sub
        End If
        If mGroupKind(g) = "mixed" Then
            For i = 0 To nodes.Count - 1
                Set node = nodes(i)
                If KindOf(buttons(i), CStr(canonicalValues(i))) = "code" Then
                    Set variants = CachedCode(node, canonicalValues(i))
                Else
                    Set variants = CachedProse(node, canonicalValues(i))
                End If
                rendered = variants(choice - 1)
                If node.TextContent <> rendered Then node.TextContent = rendered
                mRenderedText(Key(node)) = rendered
            Next i
            Exit Sub
        End If
    End If
    For i = 0 To nodes.Count - 1
        Set node = nodes(i)
        current = node.TextContent
        If owner.Disabled Then
            mRenderedText(Key(node)) = current
        Else
            If KindOf(owner, CStr(canonicalValues(i))) = "code" Then
                rendered = CachedCode(node, canonicalValues(i))(choice - 1)
            Else
                rendered = CachedProse(node, canonicalValues(i))(choice - 1)
            End If
            If current <> rendered Then node.TextContent = rendered
            mRenderedText(Key(node)) = rendered
        End If
    Next i
End Sub

Public Sub ApplyAnswerVariantPreviewTask(ByVal button As clsNode)
    On Error GoTo Failed
    ApplyAnswerVariantPreview button
    Exit Sub
Failed:
    ReportError "modVariants.ApplyAnswerVariantPreviewTask"
End Sub

Public Sub RefreshAnswerVariantPreviews(Optional ByVal scope As clsNode = Nothing)
    Dim buttons As clsList, list As clsList, i As Long
    EnsureState
    If scope Is Nothing Then Set scope = Doc
    Set buttons = NewList()
    If IsButton(scope) Then
        If Len(DataText(scope, "answerVariantReady")) > 0 Then buttons.Push scope
    End If
    Set list = scope.QuerySelectorAll("button[data-answer-variant-ready]")
    For i = 0 To list.Count - 1
        buttons.Push list(i)
    Next i
    For i = 0 To buttons.Count - 1
        ApplyAnswerVariantPreview buttons(i)
    Next i
End Sub

Public Sub RefreshAnswerVariantPreviewsTask(ByVal scope As clsNode)
    On Error GoTo Failed
    RefreshAnswerVariantPreviews scope
    Exit Sub
Failed:
    ReportError "modVariants.RefreshAnswerVariantPreviewsTask"
End Sub

Public Sub ScheduleAnswerVariantRefresh(Optional ByVal source As clsNode = Nothing)
    Dim scope As clsNode
    If IsElement(source) Then Set scope = source.Closest(".bench")
    If scope Is Nothing Then Set scope = Doc
    SetTimeout "modVariants.RefreshAnswerVariantPreviewsTask", 0, scope
End Sub

' ================================================================ copying

' writeAnswerText: copies the visibly selected option. Returns False when the clipboard
' is unavailable (the website's rejected promise).
Public Function WriteAnswerText(ByVal text As String) As Boolean
    Dim copyButton As clsNode, g As Long, owner As clsNode, isGroupOwner As Boolean, nodes As clsList
    Dim alreadyRendered As Boolean, panel As clsNode, trackCompletion As Boolean, groupedText As String
    Dim clipboardText As String, i As Long, completed As Boolean
    EnsureState
    Set copyButton = ActiveAnswerCopyButton
    g = GroupOf(copyButton)
    Set owner = OwnerOf(copyButton)
    If g >= 0 Then isGroupOwner = (copyButton Is owner) And mGroupAggregate(g)
    If isGroupOwner Then Set nodes = GroupedNodes(g) Else Set nodes = OutputNodes(copyButton)
    If nodes.Count = 1 Then
        If nodes(0).TextContent = text And mRenderedText.Exists(Key(nodes(0))) Then
            alreadyRendered = (CStr(mRenderedText(Key(nodes(0)))) = text)
        End If
    End If
    If IsElement(copyButton) Then Set panel = copyButton.Closest(".lp-collapsible")
    If Not panel Is Nothing Then
        If Not copyButton.Disabled Then trackCompletion = ShouldCompletePanelOnCopy(panel, copyButton)
    End If
    If trackCompletion Then
        Set PendingCopyPanel = panel
        Set PendingCopyButton = copyButton
        PendingCopyTime = Now
    End If
    If g >= 0 And nodes.Count > 0 Then
        For i = 0 To nodes.Count - 1
            If i > 0 Then groupedText = groupedText & vbLf & vbLf
            groupedText = groupedText & nodes(i).TextContent
        Next i
    End If
    If Len(groupedText) > 0 Then
        clipboardText = groupedText
    ElseIf alreadyRendered Then
        clipboardText = text
    Else
        clipboardText = SelectedAnswerVariant(text, copyButton)
    End If
    If Not CalculatePlotAxisIntervals Then
        If KindOf(copyButton, clipboardText) = "code" Then
            If ChartCode(clipboardText) Then TryRun "ShowPlotAxisRangeAlert", copyButton
        End If
    End If
    If Not ClipboardWrite(clipboardText) Then Exit Function
    If trackCompletion Then
        completed = False
        On Error Resume Next
        completed = CBool(Application.Run("CompleteAnovaVarianceCopy", copyButton))
        On Error GoTo 0
        If Not completed Then
            SetPanelCompletion panel, True
            SetLpPanelCollapsed panel, True
        End If
        Set PendingCopyPanel = Nothing
        Set PendingCopyButton = Nothing
    End If
    WriteAnswerText = True
End Function

' containsRChart (modApp), with a local fallback while the shell is absent.
Private Function ChartCode(ByVal code As String) As Boolean
    Dim result As Variant, tokens As clsList, i As Long
    On Error GoTo Fallback
    result = Application.Run("ContainsRChart", code)
    ChartCode = CBool(result)
    Exit Function
Fallback:
    Err.Clear
    On Error GoTo 0
    Set tokens = VarTokenList(code)
    For i = 0 To tokens.Count - 2
        Select Case CStr(tokens(i)("text"))
            Case "plot", "barplot", "hist", "pie", "boxplot", "mosaicplot", "ggplot", "qplot", "biplot", "pairs", _
                 "pairs.panels", "fviz_eig", "fviz_cluster", "screeplot", "qqnorm", "qqplot", "matplot"
                If CStr(tokens(i + 1)("text")) = "(" Then ChartCode = True: Exit Function
        End Select
    Next i
End Function

Public Sub ClearCompletionForAnswerVariant(ByVal button As clsNode)
    Dim panel As clsNode, copyCard As clsNode, id As String, list As clsList, i As Long, section As clsNode, card As clsNode
    If Not IsElement(button) Then Exit Sub
    Set panel = button.Closest(".lp-collapsible")
    If Not panel Is Nothing Then SetPanelCompletion panel, False
    Set copyCard = button.Closest(".sensitivity-prose-card, .copyable-stage-box")
    If Not copyCard Is Nothing Then copyCard.RemoveClass "is-copy-complete"
    If JsTrim(button.TextContent) = "Copied" Then button.TextContent = "Copy"
    id = button.Id
    Select Case id
        Case "copyAnovaVarianceCodeBtn": TryRun "ResetAnovaVarianceCopyState", "code"
        Case "copyAnovaVarianceProseBtn": TryRun "ResetAnovaVarianceCopyState", "prose"
        Case "copyAnovaStandardCodeBtn": TryRun "ResetAnovaBranchCopyState", "standard", "code"
        Case "copyAnovaStandardProseBtn": TryRun "ResetAnovaBranchCopyState", "standard", "prose"
        Case "copyAnovaWelchCodeBtn": TryRun "ResetAnovaBranchCopyState", "welch", "code"
        Case "copyAnovaWelchProseBtn": TryRun "ResetAnovaBranchCopyState", "welch", "prose"
    End Select
    If button.Matches("[data-sensitivity-copy]") Then TryRun "ForgetCopiedSensitivitySection", DataText(button, "sensitivityCopy")
    If id = "copySensitivityProseBtn" Then
        TryRun "ClearCopiedSensitivitySections"
        Set list = QA("#sensitivityInterpretation [data-sensitivity-copy]")
        For i = 0 To list.Count - 1
            Set section = list(i)
            Set card = section.Closest(".sensitivity-prose-card")
            If Not card Is Nothing Then card.RemoveClass "is-copy-complete"
            If JsTrim(section.TextContent) = "Copied" Then
                section.TextContent = "Copy"
                section.Disabled = False
            End If
        Next i
    End If
End Sub

' ================================================================ controls

Private Sub AddAnswerVariantControl(ByVal button As clsNode)
    Dim owner As clsNode, k As String, g As Long, kind As String, control As clsNode, sel As clsNode
    Dim label As String, html As String, i As Long, r As Double
    If Not IsAnswerCopyButton(button) Then Exit Sub
    If Len(DataText(button, "answerVariantReady")) > 0 Then Exit Sub
    Set owner = OwnerOf(button)
    If Not owner Is Nothing Then
        If Not (owner Is button) Then
            button.DataSet "answerVariantReady", "true"
            button.DataSet "answerVariantOwner", owner.Id
            QueueMicrotask "modVariants.ApplyAnswerVariantPreviewTask", owner
            Exit Sub
        End If
    End If
    If Len(button.Id) = 0 Then button.DataSet "answerVariantId", "dynamic-" & Mid$(VarRandomToBase36(JsRandom()), 3)
    k = button.Id
    If Len(k) = 0 Then k = DataText(button, "answerVariantId")
    g = GroupOf(button)
    If g >= 0 Then kind = mGroupKind(g) Else kind = KindOf(button)
    Set control = CreateElement("label")
    control.ClassName = "answer-variant-control"
    If button.HasClass("hidden") Then control.AddClass "hidden"
    control.DataSet "copyFor", k
    label = AttrText(button, "aria-label")
    If Len(label) = 0 Then label = JsTrim(button.TextContent)
    If Len(label) = 0 Then label = "copy"
    html = "<span>Answer option</span><select class=""answer-variant-select"" aria-label=""Answer option for " & EscapeHtml(label) & _
        """ title=""Option 1 is the original answer. Selecting Options 2" & ChrW$(&H2013) & "6 creates a fresh variation each time. Copy uses exactly the answer shown below."">"
    For i = 1 To 6
        html = html & "<option value=""" & CStr(i) & """>" & VariantLabel(kind, i) & "</option>"
    Next i
    control.InnerHtml = html & "</select>"
    button.Before control
    button.DataSet "answerVariantReady", "true"
    Set sel = control.QuerySelector("select")
    r = JsRandom() * 6
    sel.Value = CStr(1 + Int(r))
    sel.DataSet "option", sel.Value
    Set mSelectButton("s" & CStr(sel.Uid)) = button
    mSelectGroup("s" & CStr(sel.Uid)) = g
    sel.AddEventListener "change", "modVariants.OnAnswerSelectChange"
    QueueMicrotask "modVariants.ApplyAnswerVariantPreviewTask", button
End Sub

' answerVariantLabels[kind][index - 1]
Private Function VariantLabel(ByVal kind As String, ByVal index As Long) As String
    If index = 1 Then
        VariantLabel = "Option 1 " & ChrW$(&H2014) & " Original"
    ElseIf kind = "mixed" Then
        VariantLabel = "Option " & CStr(index) & " " & ChrW$(&H2014) & " Alternative " & CStr(index - 1)
    Else
        VariantLabel = "Option " & CStr(index) & " " & ChrW$(&H2014) & " Random variation"
    End If
End Function

' The select's change listener (a closure over its button and group in the website).
Public Sub OnAnswerSelectChange(ByVal e As clsEvent)
    Dim sel As clsNode, k As String, button As clsNode, g As Long, previousOption As Long, restore As Boolean
    Dim scope As clsNode, list As clsList, i As Long
    On Error GoTo Failed
    Set sel = e.CurrentTarget
    k = "s" & CStr(sel.Uid)
    If Not mSelectButton.Exists(k) Then Exit Sub
    Set button = mSelectButton(k)
    g = mSelectGroup(k)
    previousOption = NumberOr1(DataText(sel, "option"))
    sel.DataSet "option", sel.Value
    If IsObject(e.Detail) Then
        If TypeOf e.Detail Is clsMap Then restore = CBool(e.Detail.Exists("restoreAnswerSelection"))
    End If
    If Not restore Then
        If g >= 0 Then
            If mGroupResetTogether(g) Then
                Set scope = Q(mGroupScope(g))
                If Not scope Is Nothing Then
                    Set list = scope.QuerySelectorAll(mGroupChildren(g))
                    For i = 0 To list.Count - 1
                        ClearCompletionForAnswerVariant list(i)
                    Next i
                End If
            Else
                ClearCompletionForAnswerVariant button
            End If
        Else
            ClearCompletionForAnswerVariant button
        End If
        RegenerateSelectedAnswerVariant button, NumberOr1(sel.Value), previousOption
    End If
    ApplyAnswerVariantPreview button
    Exit Sub
Failed:
    ReportError "modVariants.OnAnswerSelectChange"
End Sub

Public Sub EnhanceAnswerVariantControls(Optional ByVal root As clsNode = Nothing)
    Dim list As clsList, i As Long
    EnsureState
    If root Is Nothing Then Set root = Doc
    If IsButton(root) Then AddAnswerVariantControl root
    Set list = root.QuerySelectorAll("button")
    For i = 0 To list.Count - 1
        AddAnswerVariantControl list(i)
    Next i
End Sub

Public Sub RandomizeAnswerOptions(Optional ByVal scope As clsNode = Nothing, Optional ByVal chooseDifferent As Boolean = False)
    Dim list As clsList, i As Long, sel As clsNode, current As Long, nxt As Long, r As Double
    EnsureState
    If scope Is Nothing Then Set scope = Doc
    EnhanceAnswerVariantControls scope
    Set list = scope.QuerySelectorAll(".answer-variant-select")
    For i = 0 To list.Count - 1
        Set sel = list(i)
        current = OptionOf(sel)
        If chooseDifferent Then
            r = JsRandom() * 5
            nxt = 1 + ((current - 1 + 1 + Int(r)) Mod 6)
        Else
            r = JsRandom() * 6
            nxt = 1 + Int(r)
        End If
        sel.Value = CStr(nxt)
        sel.DispatchEvent NewEvent("change", True)
    Next i
End Sub

' ================================================================ selection memory

Public Function VisibleOperationBench() As clsNode
    Set VisibleOperationBench = Q("main > .bench:not(.hidden), main > #lpBench:not(.hidden)")
End Function

Public Function CollectAnswerSelections(Optional ByVal scope As clsNode = Nothing) As clsMap
    Dim selections As clsMap, list As clsList, i As Long, sel As clsNode
    Set selections = NewMap()
    If scope Is Nothing Then Set scope = VisibleOperationBench()
    If Not scope Is Nothing Then
        Set list = scope.QuerySelectorAll(".answer-variant-control[data-copy-for] .answer-variant-select")
        For i = 0 To list.Count - 1
            Set sel = list(i)
            selections(DataText(sel.Closest(".answer-variant-control"), "copyFor")) = sel.Value
        Next i
    End If
    Set CollectAnswerSelections = selections
End Function

Public Sub ClearCodeVariantCache(ByVal scope As clsNode)
    Dim list As clsList, i As Long, j As Long, button As clsNode, owner As clsNode, nodes As clsList
    EnsureState
    If scope Is Nothing Then Exit Sub
    Set list = scope.QuerySelectorAll("button[data-answer-variant-ready]")
    For i = 0 To list.Count - 1
        Set button = list(i)
        Set owner = OwnerOf(button)
        If owner Is Nothing Then Set owner = button
        mCodeCache.Remove Key(owner)
        mProseCache.Remove Key(owner)
        Set nodes = OutputNodes(button)
        For j = 0 To nodes.Count - 1
            mCodeCache.Remove Key(nodes(j))
            mProseCache.Remove Key(nodes(j))
        Next j
    Next i
End Sub

Private Function OperationButton(ByVal operationId As String) As clsNode
    Set OperationButton = Q("#tutorialList .operation[data-id=""" & CssEscape(operationId) & """]")
End Function

Public Sub RememberOperationAnswerSelections(ByVal operationId As String, Optional ByVal scope As clsNode = Nothing)
    Dim selections As clsMap, button As clsNode
    EnsureState
    If Len(operationId) = 0 Then Exit Sub
    If scope Is Nothing Then Set scope = VisibleOperationBench()
    If scope Is Nothing Then Exit Sub
    Set selections = CollectAnswerSelections(scope)
    Set mOperationAnswerSelections(operationId) = selections
    Set button = OperationButton(operationId)
    If Not button Is Nothing Then button.DataSet "answerSelections", SelectionsJson(selections)
End Sub

' JSON.stringify(selections) for a map of strings.
Private Function SelectionsJson(ByVal selections As clsMap) As String
    Dim out As String, k As Variant, first As Boolean, ks As clsList
    out = "{"
    first = True
    Set ks = selections.Keys()
    For Each k In ks
        If Not first Then out = out & ","
        out = out & JsonString(CStr(k)) & ":" & JsonString(CStr(selections(CStr(k))))
        first = False
    Next k
    SelectionsJson = out & "}"
End Function

Private Function JsonString(ByVal text As String) As String
    Dim i As Long, c As Long, out As String
    For i = 1 To Len(text)
        c = AscW(Mid$(text, i, 1)) And &HFFFF&
        Select Case c
            Case 34: out = out & "\"""
            Case 92: out = out & "\\"
            Case 8: out = out & "\b"
            Case 12: out = out & "\f"
            Case 10: out = out & "\n"
            Case 13: out = out & "\r"
            Case 9: out = out & "\t"
            Case 0 To 31: out = out & "\u" & LCase$(Right$("000" & Hex$(c), 4))
            Case Else: out = out & Mid$(text, i, 1)
        End Select
    Next i
    JsonString = """" & out & """"
End Function

Public Sub PrepareAnswerOptionsForVisit(ByVal operationId As String)
    Dim scope As clsNode, button As clsNode, wasEdited As Boolean, saved As clsMap, list As clsList, i As Long
    Dim sel As clsNode, k As String, previous As Long, savedValue As String, evt As clsEvent, detail As clsMap
    Dim current As clsMap, r As Double
    EnsureState
    Set scope = VisibleOperationBench()
    If scope Is Nothing Then Exit Sub
    EnhanceAnswerVariantControls scope
    Set button = OperationButton(operationId)
    wasEdited = UserEditedOperations.Exists(operationId)
    If Not wasEdited Then wasEdited = (DataText(button, "userEdited") = "true")
    If wasEdited Then
        If LockedOperationAnswerSelections.Exists(operationId) Then Set saved = LockedOperationAnswerSelections(operationId)
    Else
        If mLastVisitAnswerSelections.Exists(operationId) Then Set saved = mLastVisitAnswerSelections(operationId)
    End If
    If saved Is Nothing Then Set saved = NewMap()
    Set list = scope.QuerySelectorAll(".answer-variant-control[data-copy-for] .answer-variant-select")
    If Not wasEdited Then
        ClearCodeVariantCache scope
        For i = 0 To list.Count - 1
            Set sel = list(i)
            k = DataText(sel.Closest(".answer-variant-control"), "copyFor")
            savedValue = ""
            If saved.Exists(k) Then savedValue = CStr(saved(k))
            If Len(savedValue) = 0 Then savedValue = sel.Value
            previous = OptionOfText(savedValue)
            r = JsRandom() * 5
            sel.Value = CStr(1 + ((previous - 1 + 1 + Int(r)) Mod 6))
            sel.DispatchEvent NewEvent("change", True)
        Next i
    Else
        For i = 0 To list.Count - 1
            Set sel = list(i)
            k = DataText(sel.Closest(".answer-variant-control"), "copyFor")
            savedValue = ""
            If saved.Exists(k) Then savedValue = CStr(saved(k))
            If Len(savedValue) > 0 Then
                sel.Value = savedValue
                Set evt = NewEvent("change", True)
                Set detail = NewMap()
                detail("restoreAnswerSelection") = True
                Set evt.Detail = detail
                sel.DispatchEvent evt
            End If
        Next i
    End If
    Set current = CollectAnswerSelections(scope)
    Set mLastVisitAnswerSelections(operationId) = current
    If wasEdited Then Set LockedOperationAnswerSelections(operationId) = current
    RememberOperationAnswerSelections operationId, scope
End Sub

Private Function OptionOfText(ByVal text As String) As Long
    Dim n As Double
    n = JsNumber(text)
    If JsIsNaN(n) Then n = 1
    If n = 0 Then n = 1
    If n > 6 Then n = 6
    If n < 1 Then n = 1
    OptionOfText = Int(n)
End Function

' ================================================================ document listeners

' The website's document-level answer-option listeners and MutationObserver. The shell
' (modApp2.InitShellPartA) registers its own copies of these; boots without the shell
' call this instead.
Public Sub InstallAnswerVariantListeners()
    Doc.AddEventListener "click", "modVariants.OnDocClick", True
    Doc.AddEventListener "paste", "modVariants.OnDocPaste", True
    Doc.AddEventListener "input", "modVariants.OnDocInput"
    Doc.AddEventListener "change", "modVariants.OnDocChange"
    SetMutationHook "modVariants.OnDomMutation"
End Sub

Private Function StateId() As String
    If Not AppState Is Nothing Then StateId = CStr(AppState("id"))
End Function

Private Sub RememberTrustedOperationEdit(ByVal e As clsEvent)
    Dim bench As clsNode, button As clsNode, id As String
    If Not IsElement(e.Target) Then Exit Sub
    If e.Target.Matches(".answer-variant-select") Then Exit Sub
    Set bench = e.Target.Closest(".bench")
    If bench Is Nothing Then Exit Sub
    If bench.HasClass("hidden") Then Exit Sub
    id = StateId()
    UserEditedOperations(id) = True
    Set button = OperationButton(id)
    If Not button Is Nothing Then button.DataSet "userEdited", "true"
    RememberOperationAnswerSelections id, bench
    Set LockedOperationAnswerSelections(id) = CollectAnswerSelections(bench)
End Sub

Public Sub OnDocClick(ByVal e As clsEvent)
    Dim button As clsNode, scope As clsNode
    On Error GoTo Failed
    If IsElement(e.Target) Then Set button = e.Target.Closest("button")
    If Not button Is Nothing Then
        Set ActiveAnswerCopyButton = button
        If RxTest(button.Id, "^paste", "i") Then
            Set scope = button.Closest(".bench")
            If scope Is Nothing Then Set scope = Doc
            SetTimeout "modVariants.RandomizeDifferentTask", 500, scope
        End If
    End If
    ScheduleAnswerVariantRefresh e.Target
    Exit Sub
Failed:
    ReportError "modVariants.OnDocClick"
End Sub

Public Sub OnDocPaste(ByVal e As clsEvent)
    Dim scope As clsNode
    On Error GoTo Failed
    If IsElement(e.Target) Then Set scope = e.Target.Closest(".bench")
    If scope Is Nothing Then Set scope = Doc
    SetTimeout "modVariants.RandomizeDifferentTask", 0, scope
    Exit Sub
Failed:
    ReportError "modVariants.OnDocPaste"
End Sub

Public Sub RandomizeDifferentTask(ByVal scope As clsNode)
    On Error GoTo Failed
    RandomizeAnswerOptions scope, True
    Exit Sub
Failed:
    ReportError "modVariants.RandomizeDifferentTask"
End Sub

Public Sub OnDocInput(ByVal e As clsEvent)
    On Error GoTo Failed
    RememberTrustedOperationEdit e
    ScheduleAnswerVariantRefresh e.Target
    Exit Sub
Failed:
    ReportError "modVariants.OnDocInput"
End Sub

Public Sub OnDocChange(ByVal e As clsEvent)
    Dim id As String, isSelect As Boolean
    On Error GoTo Failed
    RememberTrustedOperationEdit e
    If IsElement(e.Target) Then isSelect = e.Target.Matches(".answer-variant-select")
    If isSelect Then
        id = StateId()
        RememberOperationAnswerSelections id
        If UserEditedOperations.Exists(id) Then Set LockedOperationAnswerSelections(id) = CollectAnswerSelections()
    Else
        ScheduleAnswerVariantRefresh e.Target
    End If
    Exit Sub
Failed:
    ReportError "modVariants.OnDocChange"
End Sub

Public Sub OnDomMutation(ByVal parent As clsNode, ByVal added As clsList)
    Dim i As Long, node As clsNode
    On Error GoTo Failed
    For i = 0 To added.Count - 1
        Set node = added(i)
        If node.NodeType = 1 Then EnhanceAnswerVariantControls node
    Next i
    Exit Sub
Failed:
    ReportError "modVariants.OnDomMutation"
End Sub

' Diagnostics: number of nodes the caches currently track.
Public Function AnswerVariantCacheSize() As Long
    EnsureState
    AnswerVariantCacheSize = mNodes.Count
End Function
