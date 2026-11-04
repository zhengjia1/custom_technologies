Attribute VB_Name = "modDom"
Option Explicit

' The document, HTML parser/serializer, id index, selector cache, event helpers and the
' virtual-clock task queue. Port of pdf/src/minidom.js (+ the timer part of env.js).
' See FRAMEWORK.md for the public API. Procedures prefixed "Dom" are internal helpers
' shared with clsNode/clsDomSelector; operation code should not call them.

Private mDoc As clsNode
Private mWin As clsNode
Private mHtml As clsNode
Private mHead As clsNode
Private mBody As clsNode
Private mUid As Long
Private mVersion As Long
Private mActive As clsNode

' id index (like minidom's _ids / _idMisses / _idVersion)
Private mIds As clsDomHash
Private mMisses As clsDomHash
Private mIdVersion As Long

Private mSelectors As clsDomHash
Private mEntities As clsDomHash

' Query-result cache: querySelector/All results by (selector, root, first-only). Four
' counters track the changes that can alter which elements a selector matches: element
' insertion/removal (mTreeVer), class/id changes (mClsVer), other attributes and the
' dataset (mAttrVer), and form state - checked/selected, non-form disabled (mFormVer).
' Text, value, style and listener changes touch none of them. A cached result is valid
' while the counters its selector depends on (clsDomSelector.Dependencies) are unchanged,
' so the many repeated queries between re-renders are answered without walking the tree.
' Selectors whose result can change otherwise (:empty, :focus, :placeholder-shown, ...)
' are never cached (clsDomSelector.Cacheable).
Private mTreeVer As Long
Private mClsVer As Long
Private mAttrVer As Long
Private mFormVer As Long
Private mQCache As clsDomHash        ' key -> clsList
Private mQCacheVer As clsDomHash     ' key -> selector version when stored
Private Const QCACHE_LIMIT As Long = 3000

' Document snapshot: the connected elements in document order with their tag, padded
' class list and id, valid while mTreeVer and mClsVer are unchanged. Built on the second
' document-wide tree walk within one structural version; document-wide queries then
' test the arrays and fully match only the candidates (clsDomSelector.CollectFlat).
Private mSnapVersion As Long
Private mSnapValid As Boolean
Private mSnapWalkVersion As Long
Private mSnapN As Long
Private mSnapNodes() As clsNode
Private mSnapTags() As String
Private mSnapCls() As String
Private mSnapIds() As String

' timers
Private mTId() As Long
Private mTDue() As Double
Private mTHandler() As String
Private mTArg() As Variant
Private mTHasArg() As Boolean
Private mTCount As Long
Private mNow As Double
Private mSeq As Long

' microtasks (kind 0: handler, 1: mutation-hook delivery)
Private mMHandler() As String
Private mMArg() As Variant
Private mMHasArg() As Boolean
Private mMKind() As Long
Private mMHead As Long
Private mMCount As Long
Private mMCap As Long

' mutation hook
Private mHook As String
Private mRecParent() As clsNode
Private mRecAdded() As clsList
Private mRecCount As Long
Private mRecScheduled As Boolean
Private mSuppressHook As Boolean


Private Const SETTLE_LIMIT As Long = 20000

' ================================================================ document

Private Sub EnsureDocument()
    If mDoc Is Nothing Then InitDocument
End Sub

Private Sub InitDocument()
    DomStructBump 0
    Set mIds = New clsDomHash
    Set mMisses = New clsDomHash
    If mSelectors Is Nothing Then Set mSelectors = New clsDomHash
    Set mDoc = NewNode(9, "")
    Set mWin = NewNode(0, "")
    Set mHtml = CreateElement("html")
    Set mHead = CreateElement("head")
    Set mBody = CreateElement("body")
    mDoc.AppendChild mHtml
    mHtml.AppendChild mHead
    mHtml.AppendChild mBody
End Sub

Private Function NewNode(ByVal nodeType As Long, ByVal localName As String) As clsNode
    Dim node As New clsNode
    mUid = mUid + 1
    node.Init nodeType, localName, mUid
    Set NewNode = node
End Function

Public Property Get Doc() As clsNode
    EnsureDocument
    Set Doc = mDoc
End Property

Public Property Get Win() As clsNode
    EnsureDocument
    Set Win = mWin
End Property

Public Property Get Body() As clsNode
    EnsureDocument
    Set Body = mBody
End Property

Public Property Get Head() As clsNode
    EnsureDocument
    Set Head = mHead
End Property

Public Property Get DocumentElement() As clsNode
    EnsureDocument
    Set DocumentElement = mHtml
End Property

' Global mutation counter: increments on any change anywhere.
Public Function Version() As Long
    Version = mVersion
End Function

Public Function Q(ByVal selector As String) As clsNode
    Set Q = Doc.QuerySelector(selector)
End Function

Public Function QA(ByVal selector As String) As clsList
    Set QA = Doc.QuerySelectorAll(selector)
End Function

' document.querySelectorAll(selector) for callers that only read the list (it may be the
' cached result list itself: never change it).
Public Function QAReadOnly(ByVal selector As String) As clsList
    Set QAReadOnly = DomQuery(Doc, selector, False, True)
End Function

Public Function ById(ByVal elementId As String) As clsNode
    EnsureDocument
    Set ById = DomLookupId(elementId)
End Function

Public Function CreateElement(ByVal tagName As String) As clsNode
    Set CreateElement = NewNode(1, LCase$(tagName))
End Function

Public Function CreateTextNode(ByVal text As String) As clsNode
    Dim node As clsNode
    Set node = NewNode(3, "")
    node.SetDataRaw text
    Set CreateTextNode = node
End Function

Public Function CreateComment(ByVal text As String) As clsNode
    Dim node As clsNode
    Set node = NewNode(8, "")
    node.SetDataRaw text
    Set CreateComment = node
End Function

Public Function CreateFragment() As clsNode
    Set CreateFragment = NewNode(11, "")
End Function

' Parses markup into a new (disconnected) document fragment.
Public Function ParseFragment(ByVal html As String) As clsNode
    Dim frag As clsNode
    Set frag = CreateFragment()
    BuildTree frag, html
    Set ParseFragment = frag
End Function

' Replaces the document with one whose body is parsed from bodyHtml (the website's
' "<body ...>...</body>" markup; body attributes are kept).
Public Sub LoadDocument(ByVal bodyHtml As String)
    Dim p As Long, n As Long, closeAt As Long, attrText As String, frag As clsNode, tail As Long
    ResetDocument
    InitDocument
    n = Len(bodyHtml)
    p = 1
    Do While p <= n
        If Not DomIsSpace(AscW(Mid$(bodyHtml, p, 1))) Then Exit Do
        p = p + 1
    Loop
    If LCase$(Mid$(bodyHtml, p, 5)) = "<body" Then
        closeAt = InStr(p + 5, bodyHtml, ">")
        If closeAt > 0 Then
            attrText = Mid$(bodyHtml, p + 5, closeAt - p - 5)
            ApplyBodyAttributes attrText
            bodyHtml = Mid$(bodyHtml, closeAt + 1)
            ' strip a trailing </body> followed only by whitespace
            tail = Len(bodyHtml)
            Do While tail > 0
                If Not DomIsSpace(AscW(Mid$(bodyHtml, tail, 1))) Then Exit Do
                tail = tail - 1
            Loop
            If tail >= 7 Then
                If LCase$(Mid$(bodyHtml, tail - 6, 7)) = "</body>" Then bodyHtml = Left$(bodyHtml, tail - 7)
            End If
        End If
    End If
    Set frag = CreateFragment()
    BuildTree frag, bodyHtml
    mSuppressHook = True
    mBody.AppendChild frag
    mSuppressHook = False
End Sub

' /([A-Za-z-]+)\s*=\s*"([^"]*)"/g over the body tag's attribute text.
Private Sub ApplyBodyAttributes(ByVal text As String)
    Dim i As Long, j As Long, k As Long, n As Long, c As Long, nm As String, closeAt As Long
    n = Len(text)
    i = 1
    Do While i <= n
        j = i
        Do While j <= n
            c = AscW(Mid$(text, j, 1))
            If Not ((c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Or c = 45) Then Exit Do
            j = j + 1
        Loop
        If j > i Then
            nm = Mid$(text, i, j - i)
            k = j
            Do While k <= n
                If Not DomIsSpace(AscW(Mid$(text, k, 1))) Then Exit Do
                k = k + 1
            Loop
            If Mid$(text, k, 1) = "=" Then
                k = k + 1
                Do While k <= n
                    If Not DomIsSpace(AscW(Mid$(text, k, 1))) Then Exit Do
                    k = k + 1
                Loop
                If Mid$(text, k, 1) = """" Then
                    closeAt = InStr(k + 1, text, """")
                    If closeAt > 0 Then
                        mBody.SetAttribute nm, DecodeEntities(Mid$(text, k + 1, closeAt - k - 1))
                        i = closeAt + 1
                        GoTo NextScan
                    End If
                End If
            End If
        End If
        i = i + 1
NextScan:
    Loop
End Sub

' Breaks the document's parent/child reference cycles and clears all DOM state
' (timers, microtasks, hooks, id index). Call when the app closes.
Public Sub ResetDocument()
    If Not mDoc Is Nothing Then mDoc.Dismantle
    If Not mWin Is Nothing Then mWin.Dismantle
    Set mDoc = Nothing
    Set mWin = Nothing
    Set mHtml = Nothing
    Set mHead = Nothing
    Set mBody = Nothing
    Set mActive = Nothing
    Set mIds = Nothing
    Set mMisses = Nothing
    Set mSelectors = Nothing
    mIdVersion = 0
    mTCount = 0
    Erase mTId: Erase mTDue: Erase mTHandler: Erase mTArg: Erase mTHasArg
    mMHead = 0
    mMCount = 0
    mMCap = 0
    Erase mMHandler: Erase mMArg: Erase mMHasArg: Erase mMKind
    mRecCount = 0
    mRecScheduled = False
    Erase mRecParent: Erase mRecAdded
    mHook = ""
    mVersion = mVersion + 1
    DomStructBump 0
    Set mQCache = Nothing
    Set mQCacheVer = Nothing
    ClearSnapshot
End Sub

Private Sub ClearSnapshot()
    mSnapValid = False
    mSnapN = 0
    Erase mSnapNodes: Erase mSnapTags: Erase mSnapCls: Erase mSnapIds
End Sub

Private Sub BuildSnapshot()
    If mSnapN = 0 Then
        ReDim mSnapNodes(0 To 2047): ReDim mSnapTags(0 To 2047): ReDim mSnapCls(0 To 2047): ReDim mSnapIds(0 To 2047)
    End If
    mSnapN = 0
    SnapWalk mDoc
    mSnapVersion = mTreeVer + mClsVer
    mSnapValid = True
End Sub

Private Sub SnapWalk(ByVal node As clsNode)
    Dim i As Long, n As Long, kid As clsNode, cap As Long
    n = node.KidCount
    For i = 0 To n - 1
        Set kid = node.Kid(i)
        If kid.NodeType = 1 Then
            If mSnapN > UBound(mSnapNodes) Then
                cap = 2 * mSnapN
                ReDim Preserve mSnapNodes(0 To cap - 1): ReDim Preserve mSnapTags(0 To cap - 1)
                ReDim Preserve mSnapCls(0 To cap - 1): ReDim Preserve mSnapIds(0 To cap - 1)
            End If
            Set mSnapNodes(mSnapN) = kid
            mSnapTags(mSnapN) = kid.LocalName
            mSnapCls(mSnapN) = kid.ClassKey
            If kid.HasIdAttr Then mSnapIds(mSnapN) = kid.Id Else mSnapIds(mSnapN) = vbNullChar
            mSnapN = mSnapN + 1
            If kid.KidCount > 0 Then SnapWalk kid
        End If
    Next
End Sub

' Collects a document-wide query, through the snapshot when it pays off.
Private Sub CollectDocument(ByVal sel As clsDomSelector, ByVal firstOnly As Boolean, ByVal results As clsList)
    Dim v As Long
    v = mTreeVer + mClsVer
    If Not (mSnapValid And mSnapVersion = v) Then
        If mSnapWalkVersion <> v Then
            ' First document walk in this version: walk the tree.
            mSnapWalkVersion = v
            sel.Collect mDoc, firstOnly, results
            Exit Sub
        End If
        BuildSnapshot
    End If
    If mSnapN > 0 Then sel.CollectFlat mSnapNodes, mSnapTags, mSnapCls, mSnapIds, 0, mSnapN - 1, Nothing, firstOnly, results
End Sub

Public Function ActiveElement() As clsNode
    EnsureDocument
    If Not mActive Is Nothing Then
        If mActive.IsConnected Then Set ActiveElement = mActive: Exit Function
    End If
    Set ActiveElement = mBody
End Function

' ================================================================ events

Public Function NewEvent(ByVal evType As String, Optional ByVal bubbles As Boolean = False, Optional ByVal cancelable As Boolean = False) As clsEvent
    Dim evt As New clsEvent
    evt.EventType = evType
    evt.Bubbles = bubbles
    evt.Cancelable = cancelable
    Set NewEvent = evt
End Function

' Keyboard events from the framework bubble and are cancelable, like real key presses.
Public Function NewKeyEvent(ByVal evType As String, ByVal key As String, Optional ByVal shift As Boolean = False, Optional ByVal ctrl As Boolean = False, Optional ByVal alt As Boolean = False) As clsEvent
    Dim evt As clsEvent
    Set evt = NewEvent(evType, True, True)
    evt.Key = key
    evt.ShiftKey = shift
    evt.CtrlKey = ctrl
    evt.AltKey = alt
    evt.Detail = 0
    Set NewKeyEvent = evt
End Function

' Mouse events bubble and are cancelable (click() uses this too).
Public Function NewMouseEvent(ByVal evType As String, Optional ByVal clientX As Double = 0) As clsEvent
    Dim evt As clsEvent
    Set evt = NewEvent(evType, True, True)
    evt.ClientX = clientX
    evt.Detail = 0
    Set NewMouseEvent = evt
End Function

' Handlers run through Application.Run; an error a handler does not handle itself is
' reported by VBA in that handler (Application.Run does not pass it back to the caller).
Public Sub DomRunHandler(ByVal handler As String, ByVal evt As clsEvent)
    Application.Run handler, evt
End Sub

Private Sub InvokeTask(ByVal handler As String, ByVal hasArg As Boolean, ByRef arg As Variant)
    If hasArg Then Application.Run handler, arg Else Application.Run handler
End Sub

' ================================================================ timers (virtual clock)

Public Function SetTimeout(ByVal handler As String, Optional ByVal delayMs As Double = 0, Optional ByVal arg As Variant) As Long
    Dim cap As Long
    If delayMs < 0 Then delayMs = 0
    mSeq = mSeq + 1
    If mTCount = 0 Then
        cap = 8
        ReDim mTId(0 To cap - 1): ReDim mTDue(0 To cap - 1): ReDim mTHandler(0 To cap - 1)
        ReDim mTArg(0 To cap - 1): ReDim mTHasArg(0 To cap - 1)
    ElseIf mTCount > UBound(mTId) Then
        cap = mTCount * 2
        ReDim Preserve mTId(0 To cap - 1): ReDim Preserve mTDue(0 To cap - 1): ReDim Preserve mTHandler(0 To cap - 1)
        ReDim Preserve mTArg(0 To cap - 1): ReDim Preserve mTHasArg(0 To cap - 1)
    End If
    mTId(mTCount) = mSeq
    mTDue(mTCount) = mNow + delayMs
    mTHandler(mTCount) = handler
    mTHasArg(mTCount) = Not IsMissing(arg)
    If mTHasArg(mTCount) Then
        If IsObject(arg) Then Set mTArg(mTCount) = arg Else mTArg(mTCount) = arg
    Else
        mTArg(mTCount) = Empty
    End If
    mTCount = mTCount + 1
    SetTimeout = mSeq
End Function

Public Sub ClearTimeout(ByVal timerId As Long)
    Dim i As Long
    For i = 0 To mTCount - 1
        If mTId(i) = timerId Then RemoveTimerAt i: Exit Sub
    Next
End Sub

Private Sub RemoveTimerAt(ByVal i As Long)
    Dim j As Long
    For j = i To mTCount - 2
        mTId(j) = mTId(j + 1)
        mTDue(j) = mTDue(j + 1)
        mTHandler(j) = mTHandler(j + 1)
        If IsObject(mTArg(j + 1)) Then Set mTArg(j) = mTArg(j + 1) Else mTArg(j) = mTArg(j + 1)
        mTHasArg(j) = mTHasArg(j + 1)
    Next
    mTCount = mTCount - 1
    mTArg(mTCount) = Empty
End Sub

' Runs handler (no arguments) after 16 ms of virtual time.
Public Function RequestAnimationFrame(ByVal handler As String) As Long
    RequestAnimationFrame = SetTimeout(handler, 16)
End Function

Public Sub CancelAnimationFrame(ByVal timerId As Long)
    ClearTimeout timerId
End Sub

Public Sub QueueMicrotask(ByVal handler As String, Optional ByVal arg As Variant)
    If IsMissing(arg) Then
        PushMicrotask 0, handler, False, Empty
    Else
        PushMicrotask 0, handler, True, arg
    End If
End Sub

Private Sub PushMicrotask(ByVal kind As Long, ByVal handler As String, ByVal hasArg As Boolean, ByRef arg As Variant)
    Dim cap As Long
    If mMCount = 0 Then mMHead = 0
    If mMCount >= mMCap Then
        cap = mMCap * 2
        If cap < 8 Then cap = 8
        If mMCap = 0 Then
            ReDim mMHandler(0 To cap - 1): ReDim mMArg(0 To cap - 1): ReDim mMHasArg(0 To cap - 1): ReDim mMKind(0 To cap - 1)
        Else
            ReDim Preserve mMHandler(0 To cap - 1): ReDim Preserve mMArg(0 To cap - 1)
            ReDim Preserve mMHasArg(0 To cap - 1): ReDim Preserve mMKind(0 To cap - 1)
        End If
        mMCap = cap
    End If
    mMKind(mMCount) = kind
    mMHandler(mMCount) = handler
    mMHasArg(mMCount) = hasArg
    If IsObject(arg) Then Set mMArg(mMCount) = arg Else mMArg(mMCount) = arg
    mMCount = mMCount + 1
End Sub

Private Sub RunMicrotasks(ByRef steps As Long)
    Dim kind As Long, handler As String, hasArg As Boolean, arg As Variant
    Do While mMHead < mMCount
        kind = mMKind(mMHead)
        handler = mMHandler(mMHead)
        hasArg = mMHasArg(mMHead)
        If IsObject(mMArg(mMHead)) Then Set arg = mMArg(mMHead) Else arg = mMArg(mMHead)
        mMArg(mMHead) = Empty
        mMHead = mMHead + 1
        If mMHead = mMCount Then mMHead = 0: mMCount = 0
        If kind = 1 Then DeliverMutations Else InvokeTask handler, hasArg, arg
        steps = steps + 1
        If steps > SETTLE_LIMIT Then Err.Raise vbObjectError + 720, "modDom", "Task queue did not settle"
    Loop
End Sub

' Runs microtasks, then every timer due within horizonMs of the virtual clock (in due
' order, advancing the clock to each), repeating until nothing is left in the horizon.
Public Sub Settle(Optional ByVal horizonMs As Double = 700)
    Dim limit As Double, steps As Long, i As Long, best As Long
    Dim handler As String, hasArg As Boolean, arg As Variant
    limit = mNow + horizonMs
    Do
        RunMicrotasks steps
        best = -1
        For i = 0 To mTCount - 1
            If mTDue(i) <= limit Then
                If best < 0 Then
                    best = i
                ElseIf mTDue(i) < mTDue(best) Then
                    best = i
                End If
            End If
        Next
        If best < 0 Then Exit Do
        If mTDue(best) > mNow Then mNow = mTDue(best)
        handler = mTHandler(best)
        hasArg = mTHasArg(best)
        If IsObject(mTArg(best)) Then Set arg = mTArg(best) Else arg = mTArg(best)
        RemoveTimerAt best
        InvokeTask handler, hasArg, arg
        steps = steps + 1
        If steps > SETTLE_LIMIT Then Err.Raise vbObjectError + 720, "modDom", "Task queue did not settle"
    Loop
End Sub

' Milliseconds of virtual time until the next pending timer, or -1 when none.
Public Function NextTimerDelay() As Long
    Dim i As Long, best As Double, d As Double
    If mTCount = 0 Then NextTimerDelay = -1: Exit Function
    best = mTDue(0)
    For i = 1 To mTCount - 1
        If mTDue(i) < best Then best = mTDue(i)
    Next
    d = best - mNow
    If d < 0 Then d = 0
    NextTimerDelay = -Int(-d)
End Function

' Moves the virtual clock to the next pending timer's due time (call Settle after).
Public Sub AdvanceClock()
    Dim i As Long, best As Double
    If mTCount = 0 Then Exit Sub
    best = mTDue(0)
    For i = 1 To mTCount - 1
        If mTDue(i) < best Then best = mTDue(i)
    Next
    If best > mNow Then mNow = best
End Sub

Public Property Get ClockNow() As Double
    ClockNow = mNow
End Property

' ================================================================ mutation hook

' handler is called as Application.Run handler, parent, addedNodes (clsList) for every
' insertion into the connected document, delivered as a microtask (like MutationObserver).
Public Sub SetMutationHook(ByVal handler As String)
    mHook = handler
End Sub

Public Sub DomNoteInsert(ByVal parent As clsNode, ByRef ins() As clsNode, ByVal n As Long)
    Dim i As Long, added As clsList
    If mDoc Is Nothing Then Exit Sub
    If Not parent.IsConnected Then Exit Sub
    For i = 0 To n - 1
        DomRegisterTree ins(i)
    Next
    If Len(mHook) = 0 Or mSuppressHook Then Exit Sub
    Set added = New clsList
    For i = 0 To n - 1
        added.Add ins(i)
    Next
    If mRecCount = 0 Then
        ReDim mRecParent(0 To 7): ReDim mRecAdded(0 To 7)
    ElseIf mRecCount > UBound(mRecParent) Then
        ReDim Preserve mRecParent(0 To mRecCount * 2): ReDim Preserve mRecAdded(0 To mRecCount * 2)
    End If
    Set mRecParent(mRecCount) = parent
    Set mRecAdded(mRecCount) = added
    mRecCount = mRecCount + 1
    If Not mRecScheduled Then
        mRecScheduled = True
        PushMicrotask 1, "", False, Empty
    End If
End Sub

Private Sub DeliverMutations()
    Dim parents() As clsNode, added() As clsList, n As Long, i As Long
    mRecScheduled = False
    n = mRecCount
    If n = 0 Then Exit Sub
    parents = mRecParent
    added = mRecAdded
    mRecCount = 0
    Erase mRecParent: Erase mRecAdded
    If Len(mHook) = 0 Then Exit Sub
    For i = 0 To n - 1
        Application.Run mHook, parents(i), added(i)
    Next
End Sub

Public Sub DomNoteRemove(ByVal node As clsNode)
    If Not mActive Is Nothing Then
        If node.Contains(mActive) Then Set mActive = Nothing
    End If
End Sub

' ================================================================ internal helpers

Public Sub DomBump()
    mVersion = mVersion + 1
End Sub

' A change that can alter selector matching (see mTreeVer): kind 1 element children,
' 2 class/id, 3 other attributes/dataset, 4 form state, 0 everything.
Public Sub DomStructBump(Optional ByVal kind As Long = 0)
    Select Case kind
    Case 1: mTreeVer = mTreeVer + 1
    Case 2: mClsVer = mClsVer + 1
    Case 3: mAttrVer = mAttrVer + 1
    Case 4: mFormVer = mFormVer + 1
    Case Else
        mTreeVer = mTreeVer + 1
        mClsVer = mClsVer + 1
        mAttrVer = mAttrVer + 1
        mFormVer = mFormVer + 1
    End Select
End Sub

' Changes whenever anything selector-relevant changes (the counters only grow).
Public Function DomStructVersion() As Long
    DomStructVersion = mTreeVer + mClsVer + mAttrVer + mFormVer
End Function

' The version a cached result of sel depends on.
Private Function SelectorVersion(ByVal sel As clsDomSelector) As Long
    Dim deps As Long, v As Long
    deps = sel.Dependencies
    v = mTreeVer
    If deps And 2 Then v = v + mClsVer
    If deps And 4 Then v = v + mAttrVer
    If deps And 8 Then v = v + mFormVer
    SelectorVersion = v
End Function

Public Function DomNextUid() As Long
    mUid = mUid + 1
    DomNextUid = mUid
End Function

Public Sub DomSetActive(ByVal node As clsNode)
    Set mActive = node
End Sub

Public Function DomRawActive() As clsNode
    Set DomRawActive = mActive
End Function

' ---- id index

Public Sub DomRegisterId(ByVal el As clsNode)
    Dim cur As clsNode, found As Boolean, keep As Boolean, key As String
    If mIds Is Nothing Then Exit Sub
    key = el.Id
    If mIds.Exists(key) Then
        Set cur = mIds.Lookup(key, found)
        If Not cur Is el Then
            If cur.IsConnected Then
                If cur.HasIdAttr Then
                    If cur.Id = key Then keep = True
                End If
            End If
        End If
    End If
    If Not keep Then mIds.Store key, el
    mIdVersion = mIdVersion + 1
End Sub

Public Sub DomRegisterTree(ByVal node As clsNode)
    Dim i As Long
    If node.NodeType <> 1 Then Exit Sub
    If node.HasIdAttr Then DomRegisterId node
    For i = 0 To node.KidCount - 1
        DomRegisterTree node.Kid(i)
    Next
End Sub

Public Function DomLookupId(ByVal elementId As String) As clsNode
    Dim found As Boolean, v As Variant, el As clsNode
    EnsureDocument
    v = Empty
    If mIds.Exists(elementId) Then
        Set el = mIds.Lookup(elementId, found)
        If el.HasIdAttr Then
            If el.Id = elementId Then
                If el.IsConnected Then Set DomLookupId = el: Exit Function
            End If
        End If
    End If
    v = mMisses.Lookup(elementId, found)
    If found Then
        If v = mIdVersion Then Exit Function
    End If
    mIds.Clear
    IndexIds mDoc
    If mIds.Exists(elementId) Then
        Set DomLookupId = mIds.Lookup(elementId, found)
    Else
        mMisses.Store elementId, mIdVersion
    End If
End Function

Private Sub IndexIds(ByVal node As clsNode)
    Dim i As Long, kid As clsNode
    For i = 0 To node.KidCount - 1
        Set kid = node.Kid(i)
        If kid.NodeType = 1 Then
            If kid.HasIdAttr Then
                If Not mIds.Exists(kid.Id) Then mIds.Store kid.Id, kid
            End If
            If kid.KidCount > 0 Then IndexIds kid
        End If
    Next
End Sub

' ---- selectors

Public Function DomGetSelector(ByVal selector As String) As clsDomSelector
    Dim found As Boolean, sel As clsDomSelector
    If mSelectors Is Nothing Then Set mSelectors = New clsDomHash
    If mSelectors.Exists(selector) Then
        Set DomGetSelector = mSelectors.Lookup(selector, found)
        Exit Function
    End If
    Set sel = New clsDomSelector
    sel.Parse selector
    mSelectors.Store selector, sel
    Set DomGetSelector = sel
End Function

' "#id" with id matching [A-Za-z_][A-Za-z0-9_-]* (minidom's SIMPLE_ID fast path).
Private Function IsSimpleId(ByRef selector As String) As Boolean
    Dim i As Long, c As Long, n As Long
    n = Len(selector)
    If n < 2 Then Exit Function
    If AscW(selector) <> 35 Then Exit Function
    c = AscW(Mid$(selector, 2, 1))
    If Not ((c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Or c = 95) Then Exit Function
    For i = 3 To n
        c = AscW(Mid$(selector, i, 1))
        If Not ((c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Or (c >= 48 And c <= 57) Or c = 95 Or c = 45) Then Exit Function
    Next
    IsSimpleId = True
End Function

' readOnly: the caller promises not to change the returned list, so a cached result is
' returned without copying it.
Public Function DomQuery(ByVal root As clsNode, ByVal selector As String, ByVal firstOnly As Boolean, Optional ByVal readOnly As Boolean = False) As clsList
    Dim results As New clsList, found As clsNode, sel As clsDomSelector, key As String, hit As Boolean
    Dim list As clsList, ver As Long, stored As Variant
    EnsureDocument
    If IsSimpleId(selector) Then
        If root.NodeType = 9 Or root.IsConnected Then
            Set found = DomLookupId(Mid$(selector, 2))
            If Not found Is Nothing Then
                If Not found Is root Then
                    If root.Contains(found) Then results.Add found
                End If
            End If
            Set DomQuery = results
            Exit Function
        End If
    End If
    Set sel = DomGetSelector(selector)
    If Not sel.Cacheable Then
        If root Is mDoc Then CollectDocument sel, firstOnly, results Else sel.Collect root, firstOnly, results
        Set DomQuery = results
        Exit Function
    End If
    If mQCache Is Nothing Then
        Set mQCache = New clsDomHash
        Set mQCacheVer = New clsDomHash
    ElseIf mQCache.Count > QCACHE_LIMIT Then
        mQCache.Clear
        mQCacheVer.Clear
    End If
    If firstOnly Then key = "1" Else key = "0"
    key = key & CStr(root.Uid) & "|" & selector
    ver = SelectorVersion(sel)
    stored = mQCacheVer.Lookup(key, hit)
    If hit Then
        If stored = ver Then
            ' A copy, so callers may change the list they get.
            Set list = mQCache.Lookup(key, hit)
            If readOnly Then Set DomQuery = list Else Set DomQuery = list.Clone()
            Exit Function
        End If
    End If
    If root Is mDoc Then CollectDocument sel, firstOnly, results Else sel.Collect root, firstOnly, results
    ' Collect only reads the tree, so ver still describes the result.
    If readOnly Then mQCache.Store key, results Else mQCache.Store key, results.Clone()
    mQCacheVer.Store key, ver
    Set DomQuery = results
End Function

' ---- string helpers

' ECMAScript \s: TAB LF VT FF CR SP NBSP U+1680 U+2000-200A U+2028 U+2029 U+202F U+205F U+3000 U+FEFF
Public Function DomIsSpace(ByVal c As Long) As Boolean
    If c < 0 Then c = c + 65536
    If c <= 32 Then
        DomIsSpace = (c = 32 Or (c >= 9 And c <= 13))
    ElseIf c < 160 Then
        DomIsSpace = False
    Else
        Select Case c
        Case 160, &H1680, &H2000 To &H200A, &H2028, &H2029, &H202F, &H205F, &H3000, &HFEFF&: DomIsSpace = True
        End Select
    End If
End Function

Public Function DomTrim(ByVal text As String) As String
    Dim a As Long, b As Long, n As Long
    n = Len(text)
    a = 1
    Do While a <= n
        If Not DomIsSpace(AscW(Mid$(text, a, 1))) Then Exit Do
        a = a + 1
    Loop
    b = n
    Do While b >= a
        If Not DomIsSpace(AscW(Mid$(text, b, 1))) Then Exit Do
        b = b - 1
    Loop
    DomTrim = Mid$(text, a, b - a + 1)
End Function

' text.replace(/\s+/g, " ").trim()
Public Function DomCollapseSpace(ByVal text As String) As String
    Dim i As Long, n As Long, c As Long, out As String, used As Long, inSpace As Boolean, startAt As Long
    n = Len(text)
    i = 1
    Do While i <= n
        c = AscW(Mid$(text, i, 1))
        If DomIsSpace(c) Then
            Do While i <= n
                If Not DomIsSpace(AscW(Mid$(text, i, 1))) Then Exit Do
                i = i + 1
            Loop
            SbAdd out, used, " "
        Else
            startAt = i
            Do While i <= n
                If DomIsSpace(AscW(Mid$(text, i, 1))) Then Exit Do
                i = i + 1
            Loop
            SbAdd out, used, Mid$(text, startAt, i - startAt)
        End If
    Loop
    DomCollapseSpace = DomTrim(Left$(out, used))
End Function

' text.split(/\s+/).filter(Boolean) into tokens(); returns the count.
Public Function DomSplitSpace(ByVal text As String, ByRef tokens() As String) As Long
    Dim i As Long, n As Long, startAt As Long, count As Long
    n = Len(text)
    ReDim tokens(0 To 3)
    i = 1
    Do While i <= n
        Do While i <= n
            If Not DomIsSpace(AscW(Mid$(text, i, 1))) Then Exit Do
            i = i + 1
        Loop
        If i > n Then Exit Do
        startAt = i
        Do While i <= n
            If DomIsSpace(AscW(Mid$(text, i, 1))) Then Exit Do
            i = i + 1
        Loop
        If count > UBound(tokens) Then ReDim Preserve tokens(0 To count * 2)
        tokens(count) = Mid$(text, startAt, i - startAt)
        count = count + 1
    Loop
    DomSplitSpace = count
End Function

Public Function DomFromCodePoint(ByVal cp As Long) As String
    If cp < &H10000 And cp >= 0 Then
        DomFromCodePoint = ChrW$(cp)
    Else
        cp = cp - &H10000
        DomFromCodePoint = ChrW$(&HD800& + (cp \ 1024)) & ChrW$(&HDC00& + (cp Mod 1024))
    End If
End Function

' data-foo-bar -> fooBar
Public Function DomCamelData(ByVal attrName As String) As String
    DomCamelData = CamelDashes(Mid$(attrName, 6))
End Function

Private Function CamelDashes(ByVal text As String) As String
    Dim i As Long, n As Long, c As Long, out As String
    If InStr(1, text, "-") = 0 Then CamelDashes = text: Exit Function
    n = Len(text)
    i = 1
    Do While i <= n
        If Mid$(text, i, 1) = "-" And i < n Then
            c = AscW(Mid$(text, i + 1, 1))
            If c >= 97 And c <= 122 Then
                out = out & ChrW$(c - 32)
                i = i + 2
            Else
                out = out & "-"
                i = i + 1
            End If
        Else
            out = out & Mid$(text, i, 1)
            i = i + 1
        End If
    Loop
    CamelDashes = out
End Function

Private Function DashUppers(ByVal text As String) As String
    Dim i As Long, c As Long, out As String
    For i = 1 To Len(text)
        c = AscW(Mid$(text, i, 1))
        If c >= 65 And c <= 90 Then out = out & "-" & ChrW$(c + 32) Else out = out & ChrW$(c)
    Next
    DashUppers = out
End Function

' fooBar -> data-foo-bar
Public Function DomDashData(ByVal key As String) As String
    DomDashData = "data-" & DashUppers(key)
End Function

Public Function DomCamelCss(ByVal name As String) As String
    If Left$(name, 2) = "--" Then DomCamelCss = name Else DomCamelCss = CamelDashes(name)
End Function

Public Function DomDashCss(ByVal key As String) As String
    If Left$(key, 2) = "--" Then DomDashCss = key Else DomDashCss = DashUppers(key)
End Function

' Appends piece to a growable buffer (buf holds used characters; Left$(buf, used) is the text).
Public Sub SbAdd(ByRef buf As String, ByRef used As Long, ByRef piece As String)
    Dim n As Long, cap As Long
    n = Len(piece)
    If n = 0 Then Exit Sub
    cap = Len(buf)
    If used + n > cap Then
        cap = cap * 2
        If cap < used + n Then cap = used + n
        If cap < 256 Then cap = 256
        buf = buf & Space$(cap - Len(buf))
    End If
    Mid$(buf, used + 1, n) = piece
    used = used + n
End Sub

' The website's escapeHtml: & < > ' "
Public Function EscapeHtml(ByVal text As String) As String
    text = Replace(text, "&", "&amp;")
    text = Replace(text, "<", "&lt;")
    text = Replace(text, ">", "&gt;")
    text = Replace(text, "'", "&#39;")
    EscapeHtml = Replace(text, """", "&quot;")
End Function

' CSS.escape
Public Function CssEscape(ByVal value As String) As String
    Dim i As Long, n As Long, c As Long, out As String, ch As String
    n = Len(value)
    For i = 1 To n
        ch = Mid$(value, i, 1)
        c = AscW(ch)
        If c < 0 Then c = c + 65536
        If c = 0 Then
            out = out & ChrW$(&HFFFD&)
        ElseIf (c >= 1 And c <= 31) Or c = 127 Or (i = 1 And c >= 48 And c <= 57) Or (i = 2 And c >= 48 And c <= 57 And Left$(value, 1) = "-") Then
            out = out & "\" & LCase$(Hex$(c)) & " "
        ElseIf i = 1 And ch = "-" And n = 1 Then
            out = out & "\" & ch
        ElseIf c >= 128 Or ch = "-" Or ch = "_" Or (c >= 48 And c <= 57) Or (c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Then
            out = out & ch
        Else
            out = out & "\" & ch
        End If
    Next
    CssEscape = out
End Function

' ================================================================ entities

Private Sub E(ByVal name As String, ByVal code As Long)
    mEntities.Store name, ChrW$(code)
End Sub

Private Sub EnsureEntities()
    If Not mEntities Is Nothing Then Exit Sub
    Set mEntities = New clsDomHash
    E "amp", &H26
    E "lt", &H3C
    E "gt", &H3E
    E "quot", &H22
    E "apos", &H27
    E "nbsp", &HA0
    E "middot", &HB7
    E "times", &HD7
    E "hellip", &H2026
    E "ndash", &H2013
    E "mdash", &H2014
    E "rarr", &H2192
    E "larr", &H2190
    E "le", &H2264
    E "ge", &H2265
    E "ne", &H2260
    E "copy", &HA9
    E "deg", &HB0
    E "plusmn", &HB1
    E "minus", &H2212
    E "sup2", &HB2
    E "micro", &HB5
    E "sigma", &H3C3
    E "mu", &H3BC
    E "rho", &H3C1
    E "alpha", &H3B1
    E "beta", &H3B2
    E "ldquo", &H201C
    E "rdquo", &H201D
    E "lsquo", &H2018
    E "rsquo", &H2019
    E "bull", &H2022
    E "para", &HB6
End Sub

Private Function IsAsciiLetter(ByVal c As Long) As Boolean
    IsAsciiLetter = (c >= 65 And c <= 90) Or (c >= 97 And c <= 122)
End Function

Private Function IsAsciiDigit(ByVal c As Long) As Boolean
    IsAsciiDigit = (c >= 48 And c <= 57)
End Function

Private Function HexDigitValue(ByVal c As Long) As Long
    If c >= 48 And c <= 57 Then
        HexDigitValue = c - 48
    ElseIf c >= 65 And c <= 70 Then
        HexDigitValue = c - 55
    ElseIf c >= 97 And c <= 102 Then
        HexDigitValue = c - 87
    Else
        HexDigitValue = -1
    End If
End Function

Private Function CodeAt(ByRef text As String, ByVal i As Long) As Long
    If i < 1 Or i > Len(text) Then CodeAt = -1 Else CodeAt = AscW(Mid$(text, i, 1)) And &HFFFF&
End Function

' minidom decodeEntities: &#xHEX; &#DEC; and the named entities above.
Public Function DecodeEntities(ByVal text As String) As String
    Dim out As String, used As Long, i As Long, amp As Long, k As Long, c As Long, v As Double
    Dim found As Boolean, rep As Variant, ok As Boolean
    If InStr(1, text, "&") = 0 Then DecodeEntities = text: Exit Function
    EnsureEntities
    i = 1
    Do
        amp = InStr(i, text, "&")
        If amp = 0 Then SbAdd out, used, Mid$(text, i): Exit Do
        If amp > i Then SbAdd out, used, Mid$(text, i, amp - i)
        ok = False
        c = CodeAt(text, amp + 1)
        If c = 35 Then
            If CodeAt(text, amp + 2) = 120 Then
                k = amp + 3
                v = 0
                Do While HexDigitValue(CodeAt(text, k)) >= 0
                    v = v * 16 + HexDigitValue(CodeAt(text, k))
                    k = k + 1
                Loop
                If k > amp + 3 And CodeAt(text, k) = 59 And v <= &H10FFFF Then ok = True
            Else
                k = amp + 2
                v = 0
                Do While IsAsciiDigit(CodeAt(text, k))
                    v = v * 10 + (CodeAt(text, k) - 48)
                    k = k + 1
                Loop
                If k > amp + 2 And CodeAt(text, k) = 59 And v <= &H10FFFF Then ok = True
            End If
            If ok Then SbAdd out, used, DomFromCodePoint(CLng(v))
        ElseIf IsAsciiLetter(c) Then
            k = amp + 2
            Do While IsAsciiLetter(CodeAt(text, k)) Or IsAsciiDigit(CodeAt(text, k))
                k = k + 1
            Loop
            If CodeAt(text, k) = 59 Then
                rep = mEntities.Lookup(Mid$(text, amp + 1, k - amp - 1), found)
                If found Then SbAdd out, used, CStr(rep): ok = True
            End If
        End If
        If ok Then
            i = k + 1
        Else
            SbAdd out, used, "&"
            i = amp + 1
        End If
        If i > Len(text) Then Exit Do
    Loop
    DecodeEntities = Left$(out, used)
End Function

' ================================================================ HTML parser

Private Function IsVoidTag(ByRef nm As String) As Boolean
    Select Case nm
    Case "area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "source", "track", "wbr": IsVoidTag = True
    End Select
End Function

' AUTO_CLOSE[top][name]
Private Function AutoCloses(ByRef top As String, ByRef nm As String) As Boolean
    Select Case top
    Case "p"
        Select Case nm
        Case "p", "div", "section", "article", "h1", "h2", "h3", "h4", "ul", "ol", "pre", "table", "header", "footer": AutoCloses = True
        End Select
    Case "li": AutoCloses = (nm = "li")
    Case "option": AutoCloses = (nm = "option" Or nm = "optgroup")
    Case "optgroup": AutoCloses = (nm = "optgroup")
    Case "tr": AutoCloses = (nm = "tr")
    Case "td", "th": AutoCloses = (nm = "td" Or nm = "th" Or nm = "tr")
    End Select
End Function

Private Sub AddText(ByVal parent As clsNode, ByRef text As String)
    If Len(text) > 0 Then parent.AppendNew CreateTextNode(DecodeEntities(text))
End Sub

' Port of minidom buildTree: void, raw-text (script/style), RCDATA (textarea/title) and
' the implied end tags above. root is a fresh fragment.
Private Sub BuildTree(ByVal root As clsNode, ByRef html As String)
    Dim n As Long, codes() As Long, b() As Byte, i As Long
    Dim stack() As clsNode, top As Long, p As Long, openAt As Long, endAt As Long, c1 As Long
    Dim nm As String, q As Long, cur As Long, c As Long, el As clsNode, selfClosing As Boolean
    Dim an() As String, av() As String, ac As Long, attrName As String, attrValue As String, quote As Long, closeAt As Long
    Dim lowered As String, haveLower As Boolean, text As String, stopAt As Long
    n = Len(html)
    If n = 0 Then Exit Sub
    b = html
    ReDim codes(0 To n + 1)
    For i = 1 To n
        codes(i) = b(2 * i - 2) + 256& * b(2 * i - 1)
    Next
    codes(0) = -1
    codes(n + 1) = -1
    ReDim stack(0 To 63)
    Set stack(0) = root
    top = 0
    ReDim an(0 To 15): ReDim av(0 To 15)
    p = 1
    Do While p <= n
        openAt = InStr(p, html, "<", vbBinaryCompare)
        If openAt = 0 Then
            text = Mid$(html, p)
            AddText stack(top), text
            Exit Do
        End If
        If openAt > p Then
            text = Mid$(html, p, openAt - p)
            AddText stack(top), text
        End If
        c1 = codes(openAt + 1)
        If c1 = 33 And Mid$(html, openAt, 4) = "<!--" Then
            endAt = InStr(openAt + 4, html, "-->", vbBinaryCompare)
            If endAt = 0 Then stopAt = n + 1 Else stopAt = endAt
            stack(top).AppendNew CreateComment(Mid$(html, openAt + 4, stopAt - openAt - 4))
            If endAt = 0 Then p = n + 1 Else p = endAt + 3
        ElseIf c1 = 33 Or c1 = 63 Then
            endAt = InStr(openAt, html, ">", vbBinaryCompare)
            If endAt = 0 Then p = n + 1 Else p = endAt + 1
        ElseIf c1 = 47 Then
            endAt = InStr(openAt, html, ">", vbBinaryCompare)
            If endAt = 0 Then stopAt = n + 1 Else stopAt = endAt
            nm = LCase$(DomTrim(Mid$(html, openAt + 2, stopAt - openAt - 2)))
            p = stopAt + 1
            For i = top To 1 Step -1
                If stack(i).LocalName = nm Then top = i - 1: Exit For
            Next
        ElseIf IsAsciiLetter(c1) Then
            q = openAt + 2
            Do While q <= n And q - openAt < 64
                c = codes(q)
                If Not (IsAsciiLetter(c) Or IsAsciiDigit(c) Or c = 45) Then Exit Do
                q = q + 1
            Loop
            nm = LCase$(Mid$(html, openAt + 1, q - openAt - 1))
            cur = q
            ac = 0
            selfClosing = False
            Do While cur <= n
                Do While cur <= n
                    If Not DomIsSpace(codes(cur)) Then Exit Do
                    cur = cur + 1
                Loop
                If cur > n Then Exit Do
                c = codes(cur)
                If c = 62 Then cur = cur + 1: Exit Do
                If c = 47 And codes(cur + 1) = 62 Then selfClosing = True: cur = cur + 2: Exit Do
                q = cur
                Do While q <= n And q - cur < 256
                    c = codes(q)
                    If c = 34 Or c = 39 Or c = 62 Or c = 47 Or c = 61 Then Exit Do
                    If DomIsSpace(c) Then Exit Do
                    q = q + 1
                Loop
                If q = cur Then
                    cur = cur + 1
                Else
                    attrName = LCase$(Mid$(html, cur, q - cur))
                    cur = q
                    Do While cur <= n
                        If Not DomIsSpace(codes(cur)) Then Exit Do
                        cur = cur + 1
                    Loop
                    attrValue = ""
                    If codes(cur) = 61 Then
                        cur = cur + 1
                        Do While cur <= n
                            If Not DomIsSpace(codes(cur)) Then Exit Do
                            cur = cur + 1
                        Loop
                        quote = codes(cur)
                        If quote = 34 Or quote = 39 Then
                            closeAt = InStr(cur + 1, html, ChrW$(quote), vbBinaryCompare)
                            If closeAt = 0 Then stopAt = n + 1 Else stopAt = closeAt
                            attrValue = Mid$(html, cur + 1, stopAt - cur - 1)
                            cur = stopAt + 1
                        Else
                            q = cur
                            Do While q <= n And q - cur < 2048
                                c = codes(q)
                                If c = 62 Or DomIsSpace(c) Then Exit Do
                                q = q + 1
                            Loop
                            attrValue = Mid$(html, cur, q - cur)
                            cur = q
                        End If
                    End If
                    If ac > UBound(an) Then ReDim Preserve an(0 To ac * 2): ReDim Preserve av(0 To ac * 2)
                    an(ac) = attrName
                    av(ac) = DecodeEntities(attrValue)
                    ac = ac + 1
                End If
            Loop
            p = cur
            If top > 0 Then
                If AutoCloses(stack(top).LocalName, nm) Then top = top - 1
            End If
            If nm = "option" Or nm = "optgroup" Then
                If top > 0 Then
                    If stack(top).LocalName = "option" Then top = top - 1
                End If
            End If
            Set el = CreateElement(nm)
            For i = 0 To ac - 1
                If Not el.HasAttribute(an(i)) Then el.SetAttribute an(i), av(i)
            Next
            stack(top).AppendNew el
            If IsVoidTag(nm) Or selfClosing Then
                ' no children
            ElseIf nm = "script" Or nm = "style" Or nm = "textarea" Or nm = "title" Then
                If Not haveLower Then lowered = LCase$(html): haveLower = True
                closeAt = InStr(p, lowered, "</" & nm, vbBinaryCompare)
                If closeAt = 0 Then stopAt = n + 1 Else stopAt = closeAt
                text = Mid$(html, p, stopAt - p)
                If Len(text) > 0 Then
                    If nm = "textarea" Or nm = "title" Then
                        If Left$(text, 1) = vbLf Then text = Mid$(text, 2)
                        el.AppendNew CreateTextNode(DecodeEntities(text))
                    Else
                        el.AppendNew CreateTextNode(text)
                    End If
                End If
                If closeAt = 0 Then
                    p = n + 1
                Else
                    endAt = InStr(closeAt, html, ">", vbBinaryCompare)
                    If endAt = 0 Then p = n + 1 Else p = endAt + 1
                End If
            Else
                top = top + 1
                If top > UBound(stack) Then ReDim Preserve stack(0 To top * 2)
                Set stack(top) = el
            End If
        Else
            text = "<"
            AddText stack(top), text
            p = openAt + 1
        End If
    Loop
    mVersion = mVersion + 1
End Sub

' ================================================================ serializer

Private Function EscapeText(ByVal text As String) As String
    If InStr(1, text, "&") > 0 Then text = Replace(text, "&", "&amp;")
    If InStr(1, text, "<") > 0 Then text = Replace(text, "<", "&lt;")
    If InStr(1, text, ">") > 0 Then text = Replace(text, ">", "&gt;")
    If InStr(1, text, ChrW$(160)) > 0 Then text = Replace(text, ChrW$(160), "&nbsp;")
    EscapeText = text
End Function

Private Function EscapeAttribute(ByVal text As String) As String
    If InStr(1, text, "&") > 0 Then text = Replace(text, "&", "&amp;")
    If InStr(1, text, """") > 0 Then text = Replace(text, """", "&quot;")
    If InStr(1, text, ChrW$(160)) > 0 Then text = Replace(text, ChrW$(160), "&nbsp;")
    EscapeAttribute = text
End Function

Private Sub SerNode(ByVal node As clsNode, ByRef buf As String, ByRef used As Long)
    Dim i As Long, parent As clsNode, raw As Boolean
    Select Case node.NodeType
    Case 3
        Set parent = node.ParentNode
        If Not parent Is Nothing Then
            If parent.NodeType = 1 Then raw = (parent.LocalName = "script" Or parent.LocalName = "style")
        End If
        If raw Then SbAdd buf, used, node.Data Else SbAdd buf, used, EscapeText(node.Data)
    Case 8
        SbAdd buf, used, "<!--" & node.Data & "-->"
    Case 11
        For i = 0 To node.KidCount - 1
            SerNode node.Kid(i), buf, used
        Next
    Case 1
        SbAdd buf, used, "<" & node.LocalName
        For i = 0 To node.AttrCount - 1
            SbAdd buf, used, " " & node.AttrNameAt(i) & "=""" & EscapeAttribute(node.AttrValueAt(i)) & """"
        Next
        For i = 0 To node.DataCount - 1
            SbAdd buf, used, " " & DomDashData(node.DataKeyAt(i)) & "=""" & EscapeAttribute(node.DataValueAt(i)) & """"
        Next
        SbAdd buf, used, ">"
        If IsVoidTag(node.LocalName) Then Exit Sub
        For i = 0 To node.KidCount - 1
            SerNode node.Kid(i), buf, used
        Next
        SbAdd buf, used, "</" & node.LocalName & ">"
    End Select
End Sub

Public Function DomSerialize(ByVal node As clsNode) As String
    Dim buf As String, used As Long
    SerNode node, buf, used
    DomSerialize = Left$(buf, used)
End Function

Public Function DomSerializeChildren(ByVal node As clsNode) As String
    Dim buf As String, used As Long, i As Long
    For i = 0 To node.KidCount - 1
        SerNode node.Kid(i), buf, used
    Next
    DomSerializeChildren = Left$(buf, used)
End Function
