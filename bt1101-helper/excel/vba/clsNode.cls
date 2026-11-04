VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' A DOM node: element (1), text (3), comment (8), document (9), fragment (11), and the
' window-like event target (0, modDom.Win). Port of pdf/src/minidom.js; see FRAMEWORK.md.
' Create nodes through modDom (CreateElement, CreateTextNode, ...), never with New.

' ---- identity
Private mType As Long
Private mLocal As String
Private mUid As Long
Private mVersion As Long
Private mData As String
' ---- tree
Private mParent As clsNode
Private mIndex As Long
Private mKids() As clsNode
Private mKidCount As Long
Private mKidCap As Long
' ---- attributes (data-* live in the dataset arrays, like minidom)
Private mAN() As String
Private mAV() As String
Private mACount As Long
Private mDK() As String
Private mDV() As String
Private mDCount As Long
Private mHasId As Boolean
Private mIdText As String
Private mClsValid As Boolean
Private mCls() As String
Private mClsCount As Long
Private mClsKey As String     ' " a b c " (class tokens padded with spaces), with mCls
' ---- form state
Private mValueSet As Boolean
Private mValue As String
Private mCheckedSet As Boolean
Private mChecked As Boolean
Private mSel As Long
Private mDirty As Boolean
Private mExDisabled As Boolean
Private mExReadOnly As Boolean
' ---- inline style (camel-cased keys like CSSStyleDeclaration; custom properties verbatim)
Private mSN() As String
Private mSV() As String
Private mSCount As Long
' ---- listeners
Private mLId() As Long
Private mLType() As String
Private mLHandler() As String
Private mLCapture() As Boolean
Private mLOnce() As Boolean
Private mLCount As Long

' ================================================================ construction (modDom)

Friend Sub Init(ByVal nodeType As Long, ByVal localName As String, ByVal uid As Long)
    mType = nodeType
    mLocal = localName
    mUid = uid
    mSel = -1
End Sub

Friend Sub SetDataRaw(ByVal text As String)
    mData = text
End Sub

' ================================================================ identity

Public Property Get NodeType() As Long
    NodeType = mType
End Property

Public Property Get LocalName() As String
    LocalName = mLocal
End Property

Public Property Get TagName() As String
    If mType = 1 Then TagName = UCase$(mLocal) Else TagName = NodeName
End Property

Public Property Get NodeName() As String
    Select Case mType
    Case 1: NodeName = UCase$(mLocal)
    Case 3: NodeName = "#text"
    Case 8: NodeName = "#comment"
    Case 9: NodeName = "#document"
    Case 11: NodeName = "#document-fragment"
    Case Else: NodeName = "#window"
    End Select
End Property

Public Property Get Uid() As Long
    Uid = mUid
End Property

Public Property Get Version() As Long
    Version = mVersion
End Property

Friend Sub Touch()
    mVersion = mVersion + 1
    modDom.DomBump
End Sub

' Touch for changes that can alter which elements a selector matches; kind (see
' modDom.DomStructBump): 1 element children, 2 class/id, 3 other attributes/dataset,
' 4 form state (checked/selected, non-form disabled). Invalidates cached query results.
Friend Sub TouchStruct(ByVal kind As Long)
    mVersion = mVersion + 1
    modDom.DomBump
    modDom.DomStructBump kind
End Sub

' ================================================================ tree access

Friend Property Get KidCount() As Long
    KidCount = mKidCount
End Property

Friend Function Kid(ByVal i As Long) As clsNode
    Set Kid = mKids(i)
End Function

Friend Property Get IndexInParent() As Long
    IndexInParent = mIndex
End Property

Friend Sub SetParentRaw(ByVal parent As clsNode, ByVal index As Long)
    Set mParent = parent
    mIndex = index
End Sub

Public Property Get ParentNode() As clsNode
    Set ParentNode = mParent
End Property

Public Property Get ParentElement() As clsNode
    If Not mParent Is Nothing Then
        If mParent.NodeType = 1 Then Set ParentElement = mParent
    End If
End Property

Public Property Get ChildNodes() As clsList
    Dim list As New clsList, i As Long
    For i = 0 To mKidCount - 1
        list.Add mKids(i)
    Next
    Set ChildNodes = list
End Property

Public Property Get Children() As clsList
    Dim list As New clsList, i As Long
    For i = 0 To mKidCount - 1
        If mKids(i).NodeType = 1 Then list.Add mKids(i)
    Next
    Set Children = list
End Property

Public Property Get ChildElementCount() As Long
    Dim i As Long, n As Long
    For i = 0 To mKidCount - 1
        If mKids(i).NodeType = 1 Then n = n + 1
    Next
    ChildElementCount = n
End Property

Public Function HasChildNodes() As Boolean
    HasChildNodes = mKidCount > 0
End Function

Public Property Get FirstChild() As clsNode
    If mKidCount > 0 Then Set FirstChild = mKids(0)
End Property

Public Property Get LastChild() As clsNode
    If mKidCount > 0 Then Set LastChild = mKids(mKidCount - 1)
End Property

Public Property Get FirstElementChild() As clsNode
    Dim i As Long
    For i = 0 To mKidCount - 1
        If mKids(i).NodeType = 1 Then Set FirstElementChild = mKids(i): Exit Property
    Next
End Property

Public Property Get LastElementChild() As clsNode
    Dim i As Long
    For i = mKidCount - 1 To 0 Step -1
        If mKids(i).NodeType = 1 Then Set LastElementChild = mKids(i): Exit Property
    Next
End Property

Public Property Get NextSibling() As clsNode
    If mParent Is Nothing Then Exit Property
    If mIndex + 1 < mParent.KidCount Then Set NextSibling = mParent.Kid(mIndex + 1)
End Property

Public Property Get PreviousSibling() As clsNode
    If mParent Is Nothing Then Exit Property
    If mIndex > 0 Then Set PreviousSibling = mParent.Kid(mIndex - 1)
End Property

Public Property Get NextElementSibling() As clsNode
    Dim i As Long, n As Long, node As clsNode
    If mParent Is Nothing Then Exit Property
    n = mParent.KidCount
    For i = mIndex + 1 To n - 1
        Set node = mParent.Kid(i)
        If node.NodeType = 1 Then Set NextElementSibling = node: Exit Property
    Next
End Property

Public Property Get PreviousElementSibling() As clsNode
    Dim i As Long, node As clsNode
    If mParent Is Nothing Then Exit Property
    For i = mIndex - 1 To 0 Step -1
        Set node = mParent.Kid(i)
        If node.NodeType = 1 Then Set PreviousElementSibling = node: Exit Property
    Next
End Property

Public Property Get IsConnected() As Boolean
    Dim node As clsNode
    Set node = Me
    Do While Not node.ParentNode Is Nothing
        Set node = node.ParentNode
    Loop
    IsConnected = (node.NodeType = 9)
End Property

Public Function GetRootNode() As clsNode
    Dim node As clsNode
    Set node = Me
    Do While Not node.ParentNode Is Nothing
        Set node = node.ParentNode
    Loop
    Set GetRootNode = node
End Function

Public Function Contains(ByVal other As clsNode) As Boolean
    Do While Not other Is Nothing
        If other Is Me Then Contains = True: Exit Function
        Set other = other.ParentNode
    Loop
End Function

' ================================================================ mutation

Private Sub KidsReserve(ByVal n As Long)
    Dim cap As Long
    If n <= mKidCap Then Exit Sub
    cap = mKidCap
    If cap < 4 Then cap = 4
    Do While cap < n
        cap = cap * 2
    Loop
    If mKidCap = 0 Then ReDim mKids(0 To cap - 1) Else ReDim Preserve mKids(0 To cap - 1)
    mKidCap = cap
End Sub

' Inserts nodes (clsNode or String items of a 0-based Variant array) before ref
' (Nothing = append). Fragments contribute their children. Port of _insertNodes.
Friend Function InsertItems(ByRef items() As Variant, ByVal itemCount As Long, ByVal ref As clsNode) As Long
    Dim ins() As clsNode, n As Long, i As Long, j As Long, node As clsNode, frag As clsNode, k As Long, at As Long
    If itemCount = 0 Then Exit Function
    ReDim ins(0 To 7)
    For i = 0 To itemCount - 1
        If IsObject(items(i)) Then
            Set node = items(i)
        Else
            Set node = modDom.CreateTextNode(CStr(items(i)))
        End If
        If node.NodeType = 11 Then
            Set frag = node
            k = frag.KidCount
            For j = 0 To k - 1
                If n > UBound(ins) Then ReDim Preserve ins(0 To n * 2)
                Set ins(n) = frag.Kid(j)
                n = n + 1
            Next
            frag.DetachAllRaw
        Else
            If Not node.ParentNode Is Nothing Then node.ParentNode.RemoveChild node
            If n > UBound(ins) Then ReDim Preserve ins(0 To n * 2)
            Set ins(n) = node
            n = n + 1
        End If
    Next
    If ref Is Nothing Then
        at = mKidCount
    Else
        If Not ref.ParentNode Is Me Then Err.Raise vbObjectError + 701, "clsNode", "NotFoundError: reference node is not a child"
        at = ref.IndexInParent
    End If
    If n = 0 Then Exit Function
    KidsReserve mKidCount + n
    For i = mKidCount - 1 To at Step -1
        Set mKids(i + n) = mKids(i)
        mKids(i + n).SetParentRaw Me, i + n
    Next
    For i = 0 To n - 1
        Set mKids(at + i) = ins(i)
        ins(i).SetParentRaw Me, at + i
    Next
    mKidCount = mKidCount + n
    For i = 0 To n - 1
        If ins(i).NodeType = 1 Then modDom.DomStructBump 1: Exit For
    Next
    Touch
    modDom.DomNoteInsert Me, ins, n
    InsertItems = n
End Function

' Parser fast path: appends a new, parentless node to a disconnected container.
Friend Sub AppendNew(ByVal node As clsNode)
    KidsReserve mKidCount + 1
    Set mKids(mKidCount) = node
    node.SetParentRaw Me, mKidCount
    mKidCount = mKidCount + 1
    mVersion = mVersion + 1
    If node.NodeType = 1 Then modDom.DomStructBump 1
End Sub

' Detaches every child without events (fragment insertion moves them).
Friend Sub DetachAllRaw()
    Dim i As Long
    For i = 0 To mKidCount - 1
        mKids(i).SetParentRaw Nothing, 0
        Set mKids(i) = Nothing
    Next
    mKidCount = 0
    TouchStruct 1
End Sub

' Breaks parent/child reference cycles in this subtree (modDom.ResetDocument).
Friend Sub Dismantle()
    Dim i As Long
    For i = 0 To mKidCount - 1
        mKids(i).Dismantle
        mKids(i).SetParentRaw Nothing, 0
        Set mKids(i) = Nothing
    Next
    mKidCount = 0
    mLCount = 0
    Erase mKids
    mKidCap = 0
    modDom.DomStructBump 1
End Sub

Public Function AppendChild(ByVal node As clsNode) As clsNode
    Dim items(0 To 0) As Variant
    Set items(0) = node
    InsertItems items, 1, Nothing
    Set AppendChild = node
End Function

Public Function InsertBefore(ByVal node As clsNode, ByVal ref As clsNode) As clsNode
    Dim items(0 To 0) As Variant
    Set items(0) = node
    InsertItems items, 1, ref
    Set InsertBefore = node
End Function

Public Function RemoveChild(ByVal node As clsNode) As clsNode
    Dim at As Long, i As Long
    If Not node.ParentNode Is Me Then Err.Raise vbObjectError + 702, "clsNode", "NotFoundError: node is not a child"
    at = node.IndexInParent
    For i = at To mKidCount - 2
        Set mKids(i) = mKids(i + 1)
        mKids(i).SetParentRaw Me, i
    Next
    Set mKids(mKidCount - 1) = Nothing
    mKidCount = mKidCount - 1
    node.SetParentRaw Nothing, 0
    modDom.DomNoteRemove node
    If node.NodeType = 1 Then TouchStruct 1 Else Touch
    Set RemoveChild = node
End Function

Public Function ReplaceChild(ByVal node As clsNode, ByVal old As clsNode) As clsNode
    InsertBefore node, old
    RemoveChild old
    Set ReplaceChild = old
End Function

Public Sub Remove()
    If Not mParent Is Nothing Then mParent.RemoveChild Me
End Sub

Private Function ArgsToItems(ByRef args As Variant, ByRef items() As Variant) As Long
    Dim i As Long, n As Long
    If Not IsArray(args) Then Exit Function
    If UBound(args) < LBound(args) Then Exit Function
    ReDim items(0 To UBound(args) - LBound(args))
    For i = LBound(args) To UBound(args)
        If IsObject(args(i)) Then Set items(n) = args(i) Else items(n) = args(i)
        n = n + 1
    Next
    ArgsToItems = n
End Function

Public Sub Append(ParamArray nodesOrStrings() As Variant)
    Dim items() As Variant, n As Long, a As Variant
    a = nodesOrStrings
    n = ArgsToItems(a, items)
    If n > 0 Then InsertItems items, n, Nothing
End Sub

Public Sub Prepend(ParamArray nodesOrStrings() As Variant)
    Dim items() As Variant, n As Long, a As Variant
    a = nodesOrStrings
    n = ArgsToItems(a, items)
    If n > 0 Then InsertItems items, n, FirstChild
End Sub

Public Sub Before(ParamArray nodesOrStrings() As Variant)
    Dim items() As Variant, n As Long, a As Variant
    If mParent Is Nothing Then Exit Sub
    a = nodesOrStrings
    n = ArgsToItems(a, items)
    If n > 0 Then mParent.InsertItems items, n, Me
End Sub

Public Sub After(ParamArray nodesOrStrings() As Variant)
    Dim items() As Variant, n As Long, a As Variant
    If mParent Is Nothing Then Exit Sub
    a = nodesOrStrings
    n = ArgsToItems(a, items)
    If n > 0 Then mParent.InsertItems items, n, NextSibling
End Sub

Public Sub ReplaceWith(ParamArray nodesOrStrings() As Variant)
    Dim items() As Variant, n As Long, a As Variant, parent As clsNode
    Set parent = mParent
    If parent Is Nothing Then Exit Sub
    a = nodesOrStrings
    n = ArgsToItems(a, items)
    If n > 0 Then parent.InsertItems items, n, Me
    parent.RemoveChild Me
End Sub

Public Sub ReplaceChildren(ParamArray nodesOrStrings() As Variant)
    Dim items() As Variant, n As Long, a As Variant
    a = nodesOrStrings
    RemoveAllChildren
    n = ArgsToItems(a, items)
    If n > 0 Then InsertItems items, n, Nothing
End Sub

Friend Sub RemoveAllChildren()
    Dim i As Long, kid As clsNode, hadElement As Boolean
    If mKidCount = 0 Then Exit Sub
    For i = 0 To mKidCount - 1
        Set kid = mKids(i)
        Set mKids(i) = Nothing
        kid.SetParentRaw Nothing, 0
        If kid.NodeType = 1 Then hadElement = True
        modDom.DomNoteRemove kid
    Next
    mKidCount = 0
    If hadElement Then TouchStruct 1 Else Touch
End Sub

' ================================================================ text and markup

Public Property Get Data() As String
    Data = mData
End Property

Public Property Let Data(ByVal text As String)
    mData = text
    Touch
    If Not mParent Is Nothing Then mParent.Touch
End Property

Public Property Get NodeValue() As Variant
    If mType = 3 Or mType = 8 Then NodeValue = mData Else NodeValue = Null
End Property

Public Property Get Length() As Long
    Length = Len(mData)
End Property

Friend Sub AppendTextTo(ByRef buf As String, ByRef used As Long)
    Dim i As Long, kid As clsNode
    For i = 0 To mKidCount - 1
        Set kid = mKids(i)
        Select Case kid.NodeType
        Case 3: modDom.SbAdd buf, used, kid.Data
        Case 1, 11: kid.AppendTextTo buf, used
        End Select
    Next
End Sub

Public Property Get TextContent() As String
    Dim buf As String, used As Long
    Select Case mType
    Case 3, 8: TextContent = mData
    Case 9, 0: TextContent = ""
    Case Else
        If mKidCount = 1 Then
            If mKids(0).NodeType = 3 Then TextContent = mKids(0).Data: Exit Property
        End If
        AppendTextTo buf, used
        TextContent = Left$(buf, used)
    End Select
End Property

Public Property Let TextContent(ByVal text As String)
    If mType = 3 Or mType = 8 Then Data = text: Exit Property
    RemoveAllChildren
    If Len(text) > 0 Then AppendChild modDom.CreateTextNode(text)
End Property

Public Property Get InnerText() As String
    InnerText = TextContent
End Property

Public Property Let InnerText(ByVal text As String)
    TextContent = text
End Property

Public Property Get InnerHtml() As String
    InnerHtml = modDom.DomSerializeChildren(Me)
End Property

Public Property Let InnerHtml(ByVal html As String)
    RemoveAllChildren
    AppendChild modDom.ParseFragment(html)
End Property

Public Property Get OuterHtml() As String
    OuterHtml = modDom.DomSerialize(Me)
End Property

Public Property Let OuterHtml(ByVal html As String)
    ReplaceWith modDom.ParseFragment(html)
End Property

Public Sub InsertAdjacentHtml(ByVal position As String, ByVal html As String)
    InsertAdjacentElement position, modDom.ParseFragment(html)
End Sub

Public Sub InsertAdjacentText(ByVal position As String, ByVal text As String)
    InsertAdjacentElement position, modDom.CreateTextNode(text)
End Sub

Public Function InsertAdjacentElement(ByVal position As String, ByVal node As clsNode) As clsNode
    Select Case LCase$(position)
    Case "beforebegin": If Not mParent Is Nothing Then mParent.InsertBefore node, Me
    Case "afterbegin": InsertBefore node, FirstChild
    Case "beforeend": AppendChild node
    Case "afterend": If Not mParent Is Nothing Then mParent.InsertBefore node, NextSibling
    End Select
    Set InsertAdjacentElement = node
End Function

' ================================================================ attributes

Private Function FindAttr(ByRef name As String) As Long
    Dim i As Long
    For i = 0 To mACount - 1
        If mAN(i) = name Then FindAttr = i: Exit Function
    Next
    FindAttr = -1
End Function

Private Function FindData(ByRef key As String) As Long
    Dim i As Long
    For i = 0 To mDCount - 1
        If mDK(i) = key Then FindData = i: Exit Function
    Next
    FindData = -1
End Function

' Raw lookup for selectors: name must already be lower case.
Friend Function AttrLookup(ByRef name As String, ByRef value As String) As Boolean
    Dim i As Long
    If Left$(name, 5) = "data-" Then
        i = FindData(modDom.DomCamelData(name))
        If i >= 0 Then value = mDV(i): AttrLookup = True
    Else
        For i = 0 To mACount - 1
            If mAN(i) = name Then value = mAV(i): AttrLookup = True: Exit Function
        Next
    End If
End Function

Friend Function AttrRaw(ByVal name As String) As String
    Dim i As Long
    For i = 0 To mACount - 1
        If mAN(i) = name Then AttrRaw = mAV(i): Exit Function
    Next
End Function

Friend Function HasAttrRaw(ByVal name As String) As Boolean
    Dim i As Long
    For i = 0 To mACount - 1
        If mAN(i) = name Then HasAttrRaw = True: Exit Function
    Next
End Function

Public Function GetAttribute(ByVal name As String) As Variant
    Dim value As String
    If AttrLookup(LCase$(name), value) Then GetAttribute = value Else GetAttribute = Null
End Function

Public Function HasAttribute(ByVal name As String) As Boolean
    Dim value As String
    HasAttribute = AttrLookup(LCase$(name), value)
End Function

Public Sub SetAttribute(ByVal name As String, ByVal value As String)
    Dim i As Long, same As Boolean
    name = LCase$(name)
    If Left$(name, 5) = "data-" Then
        DataSet modDom.DomCamelData(name), value
        Exit Sub
    End If
    i = FindAttr(name)
    If i < 0 Then
        If mACount = 0 Then
            ReDim mAN(0 To 3): ReDim mAV(0 To 3)
        ElseIf mACount > UBound(mAN) Then
            ReDim Preserve mAN(0 To mACount * 2): ReDim Preserve mAV(0 To mACount * 2)
        End If
        mAN(mACount) = name
        mAV(mACount) = value
        mACount = mACount + 1
    Else
        same = (mAV(i) = value)
        mAV(i) = value
    End If
    AttrChanged name, Not same
End Sub

Public Sub RemoveAttribute(ByVal name As String)
    Dim i As Long, j As Long
    name = LCase$(name)
    If Left$(name, 5) = "data-" Then
        DataRemove modDom.DomCamelData(name)
        Exit Sub
    End If
    i = FindAttr(name)
    If i < 0 Then Exit Sub
    For j = i To mACount - 2
        mAN(j) = mAN(j + 1)
        mAV(j) = mAV(j + 1)
    Next
    mACount = mACount - 1
    AttrChanged name, True
End Sub

' structural: False when an existing attribute was set to its current value (the
' version still changes, but cached selector results stay valid).
Private Sub AttrChanged(ByRef name As String, ByVal structural As Boolean)
    If name = "class" Then
        mClsValid = False
    ElseIf name = "id" Then
        Dim i As Long
        i = FindAttr("id")
        mHasId = (i >= 0)
        If mHasId Then
            mIdText = mAV(i)
            If IsConnected Then modDom.DomRegisterId Me
        Else
            mIdText = ""
        End If
    End If
    If structural Then TouchStruct IIf(name = "class" Or name = "id", 2, 3) Else Touch
End Sub

Public Function ToggleAttribute(ByVal name As String, Optional ByVal force As Variant) As Boolean
    Dim has As Boolean, want As Boolean
    has = HasAttribute(name)
    If IsMissing(force) Then want = Not has Else want = CBool(force)
    If want And Not has Then
        SetAttribute name, ""
    ElseIf Not want And has Then
        RemoveAttribute name
    End If
    ToggleAttribute = want
End Function

Public Function AttributeNames() As clsList
    Dim list As New clsList, i As Long
    For i = 0 To mACount - 1
        list.Add mAN(i)
    Next
    For i = 0 To mDCount - 1
        list.Add modDom.DomDashData(mDK(i))
    Next
    Set AttributeNames = list
End Function

Friend Property Get AttrCount() As Long
    AttrCount = mACount
End Property

Friend Function AttrNameAt(ByVal i As Long) As String
    AttrNameAt = mAN(i)
End Function

Friend Function AttrValueAt(ByVal i As Long) As String
    AttrValueAt = mAV(i)
End Function

Friend Property Get DataCount() As Long
    DataCount = mDCount
End Property

Friend Function DataKeyAt(ByVal i As Long) As String
    DataKeyAt = mDK(i)
End Function

Friend Function DataValueAt(ByVal i As Long) As String
    DataValueAt = mDV(i)
End Function

' ---- dataset

Public Function DataGet(ByVal camelKey As String) As Variant
    Dim i As Long
    i = FindData(camelKey)
    If i >= 0 Then DataGet = mDV(i) Else DataGet = Null
End Function

Public Function DataHas(ByVal camelKey As String) As Boolean
    DataHas = FindData(camelKey) >= 0
End Function

Public Sub DataSet(ByVal camelKey As String, ByVal value As String)
    Dim i As Long, same As Boolean
    i = FindData(camelKey)
    If i < 0 Then
        If mDCount = 0 Then
            ReDim mDK(0 To 3): ReDim mDV(0 To 3)
        ElseIf mDCount > UBound(mDK) Then
            ReDim Preserve mDK(0 To mDCount * 2): ReDim Preserve mDV(0 To mDCount * 2)
        End If
        mDK(mDCount) = camelKey
        mDV(mDCount) = value
        mDCount = mDCount + 1
    Else
        same = (mDV(i) = value)
        mDV(i) = value
    End If
    If same Then Touch Else TouchStruct 3
End Sub

Public Sub DataRemove(ByVal camelKey As String)
    Dim i As Long, j As Long
    i = FindData(camelKey)
    If i < 0 Then Exit Sub
    For j = i To mDCount - 2
        mDK(j) = mDK(j + 1)
        mDV(j) = mDV(j + 1)
    Next
    mDCount = mDCount - 1
    TouchStruct 3
End Sub

' ---- reflected attributes

Private Function AttrOrEmpty(ByVal name As String) As String
    Dim i As Long
    i = FindAttr(name)
    If i >= 0 Then AttrOrEmpty = mAV(i)
End Function

Friend Property Get HasIdAttr() As Boolean
    HasIdAttr = mHasId
End Property

Public Property Get Id() As String
    Id = mIdText
End Property

Public Property Let Id(ByVal value As String)
    SetAttribute "id", value
End Property

Public Property Get ClassName() As String
    ClassName = AttrOrEmpty("class")
End Property

Public Property Let ClassName(ByVal value As String)
    SetAttribute "class", value
End Property

Public Property Get Title() As String
    Title = AttrOrEmpty("title")
End Property

Public Property Let Title(ByVal value As String)
    SetAttribute "title", value
End Property

Public Property Get Placeholder() As String
    Placeholder = AttrOrEmpty("placeholder")
End Property

Public Property Let Placeholder(ByVal value As String)
    SetAttribute "placeholder", value
End Property

Public Property Get Name() As String
    Name = AttrOrEmpty("name")
End Property

Public Property Let Name(ByVal value As String)
    SetAttribute "name", value
End Property

Public Property Get HtmlFor() As String
    HtmlFor = AttrOrEmpty("for")
End Property

Public Property Let HtmlFor(ByVal value As String)
    SetAttribute "for", value
End Property

Public Property Get Href() As String
    Href = AttrOrEmpty("href")
End Property

Public Property Let Href(ByVal value As String)
    SetAttribute "href", value
End Property

Public Property Get Role() As Variant
    Role = GetAttribute("role")
End Property

Public Property Get TabIndex() As Long
    Dim i As Long
    i = FindAttr("tabindex")
    If i < 0 Then TabIndex = -1 Else TabIndex = CLng(Val(mAV(i)))
End Property

Public Property Let TabIndex(ByVal value As Long)
    SetAttribute "tabindex", CStr(value)
End Property

' input type (lower case, default "text"); button submit/button/reset; textarea; select-one/-multiple.
Public Property Get TypeName() As String
    Dim t As String
    Select Case mLocal
    Case "input"
        t = LCase$(AttrOrEmpty("type"))
        If Len(t) = 0 Then t = "text"
        TypeName = t
    Case "button"
        t = LCase$(AttrOrEmpty("type"))
        If t = "button" Or t = "reset" Then TypeName = t Else TypeName = "submit"
    Case "textarea": TypeName = "textarea"
    Case "select": If HasAttrRaw("multiple") Then TypeName = "select-multiple" Else TypeName = "select-one"
    Case Else: TypeName = AttrOrEmpty("type")
    End Select
End Property

Public Property Let TypeName(ByVal value As String)
    SetAttribute "type", value
End Property

Friend Function IsFormElement() As Boolean
    Select Case mLocal
    Case "button", "input", "textarea", "select", "option", "optgroup": IsFormElement = (mType = 1)
    End Select
End Function

Friend Function HasValueProperty() As Boolean
    Select Case mLocal
    Case "button", "input", "textarea", "select", "option": HasValueProperty = (mType = 1)
    End Select
End Function

Public Property Get Disabled() As Boolean
    If IsFormElement() Then Disabled = HasAttrRaw("disabled") Else Disabled = mExDisabled
End Property

Public Property Let Disabled(ByVal value As Boolean)
    If IsFormElement() Then ToggleAttribute "disabled", value Else mExDisabled = value: TouchStruct 4
End Property

Public Property Get ReadOnly() As Boolean
    If mLocal = "input" Or mLocal = "textarea" Then ReadOnly = HasAttrRaw("readonly") Else ReadOnly = mExReadOnly
End Property

Public Property Let ReadOnly(ByVal value As Boolean)
    If mLocal = "input" Or mLocal = "textarea" Then ToggleAttribute "readonly", value Else mExReadOnly = value: Touch
End Property

Public Property Get Hidden() As Boolean
    Hidden = HasAttrRaw("hidden")
End Property

Public Property Let Hidden(ByVal value As Boolean)
    ToggleAttribute "hidden", value
End Property

Public Property Get IsOpen() As Boolean
    IsOpen = HasAttrRaw("open")
End Property

Public Property Let IsOpen(ByVal value As Boolean)
    ToggleAttribute "open", value
End Property

Public Property Get Multiple() As Boolean
    Multiple = HasAttrRaw("multiple")
End Property

Public Property Get Required() As Boolean
    Required = HasAttrRaw("required")
End Property

' ================================================================ classes

Private Sub EnsureClasses()
    Dim i As Long
    If mClsValid Then Exit Sub
    mClsCount = modDom.DomSplitSpace(AttrOrEmpty("class"), mCls)
    mClsKey = " "
    For i = 0 To mClsCount - 1
        mClsKey = mClsKey & mCls(i) & " "
    Next
    mClsValid = True
End Sub

' " cls1 cls2 " for selector prefiltering (modDom's document snapshot).
Friend Property Get ClassKey() As String
    If Not mClsValid Then EnsureClasses
    ClassKey = mClsKey
End Property

Friend Property Get ClassCount() As Long
    EnsureClasses
    ClassCount = mClsCount
End Property

Friend Function ClassAt(ByVal i As Long) As String
    ClassAt = mCls(i)
End Function

' Selector-matching probes (no string copies).
Friend Function HasClassRef(ByRef name As String) As Boolean
    Dim i As Long
    If Not mClsValid Then EnsureClasses
    For i = 0 To mClsCount - 1
        If mCls(i) = name Then HasClassRef = True: Exit Function
    Next
End Function

Friend Function TagIs(ByRef tag As String) As Boolean
    TagIs = (mLocal = tag)
End Function

Friend Function DataLookup(ByRef camelKey As String, ByRef value As String) As Boolean
    Dim i As Long
    For i = 0 To mDCount - 1
        If mDK(i) = camelKey Then value = mDV(i): DataLookup = True: Exit Function
    Next
End Function

Friend Function AttrLookupPlain(ByRef name As String, ByRef value As String) As Boolean
    Dim i As Long
    For i = 0 To mACount - 1
        If mAN(i) = name Then value = mAV(i): AttrLookupPlain = True: Exit Function
    Next
End Function

Public Function HasClass(ByVal name As String) As Boolean
    Dim i As Long
    If Not mClsValid Then EnsureClasses
    For i = 0 To mClsCount - 1
        If mCls(i) = name Then HasClass = True: Exit Function
    Next
End Function

Private Sub WriteClasses(ByRef tokens() As String, ByVal n As Long)
    Dim s As String, i As Long
    For i = 0 To n - 1
        If i > 0 Then s = s & " "
        s = s & tokens(i)
    Next
    SetAttribute "class", s
End Sub

Private Function CopyClasses(ByRef tokens() As String, ByVal extra As Long) As Long
    Dim i As Long
    EnsureClasses
    ReDim tokens(0 To mClsCount + extra)
    For i = 0 To mClsCount - 1
        tokens(i) = mCls(i)
    Next
    CopyClasses = mClsCount
End Function

Public Sub AddClass(ParamArray names() As Variant)
    Dim tokens() As String, n As Long, i As Long, j As Long, found As Boolean
    n = CopyClasses(tokens, UBound(names) - LBound(names) + 1)
    For i = LBound(names) To UBound(names)
        found = False
        For j = 0 To n - 1
            If tokens(j) = CStr(names(i)) Then found = True: Exit For
        Next
        If Not found Then tokens(n) = CStr(names(i)): n = n + 1
    Next
    WriteClasses tokens, n
End Sub

Public Sub RemoveClass(ParamArray names() As Variant)
    Dim tokens() As String, kept() As String, n As Long, k As Long, i As Long, j As Long, drop As Boolean
    n = CopyClasses(tokens, 0)
    ReDim kept(0 To n)
    For j = 0 To n - 1
        drop = False
        For i = LBound(names) To UBound(names)
            If tokens(j) = CStr(names(i)) Then drop = True: Exit For
        Next
        If Not drop Then kept(k) = tokens(j): k = k + 1
    Next
    If k <> n Or HasAttrRaw("class") Then WriteClasses kept, k
End Sub

Public Function ToggleClass(ByVal name As String, Optional ByVal force As Variant) As Boolean
    Dim has As Boolean, want As Boolean
    has = HasClass(name)
    If IsMissing(force) Then want = Not has Else want = CBool(force)
    If want And Not has Then
        AddClass name
    ElseIf Not want And has Then
        RemoveClass name
    End If
    ToggleClass = want
End Function

Public Function ReplaceClass(ByVal oldName As String, ByVal newName As String) As Boolean
    Dim tokens() As String, n As Long, i As Long
    n = CopyClasses(tokens, 0)
    For i = 0 To n - 1
        If tokens(i) = oldName Then
            tokens(i) = newName
            WriteClasses tokens, n
            ReplaceClass = True
            Exit Function
        End If
    Next
End Function

' ================================================================ inline style

Private Function FindStyle(ByRef key As String) As Long
    Dim i As Long
    For i = 0 To mSCount - 1
        If mSN(i) = key Then FindStyle = i: Exit Function
    Next
    FindStyle = -1
End Function

Public Sub SetStyle(ByVal name As String, ByVal value As String)
    Dim key As String, i As Long
    key = modDom.DomCamelCss(name)
    i = FindStyle(key)
    If i < 0 Then
        If mSCount = 0 Then
            ReDim mSN(0 To 3): ReDim mSV(0 To 3)
        ElseIf mSCount > UBound(mSN) Then
            ReDim Preserve mSN(0 To mSCount * 2): ReDim Preserve mSV(0 To mSCount * 2)
        End If
        mSN(mSCount) = key
        mSV(mSCount) = value
        mSCount = mSCount + 1
    Else
        mSV(i) = value
    End If
    Touch
End Sub

Public Function GetStyle(ByVal name As String) As String
    Dim i As Long
    i = FindStyle(modDom.DomCamelCss(name))
    If i >= 0 Then GetStyle = mSV(i)
End Function

Public Function RemoveStyle(ByVal name As String) As String
    Dim i As Long, j As Long
    i = FindStyle(modDom.DomCamelCss(name))
    If i < 0 Then Exit Function
    RemoveStyle = mSV(i)
    For j = i To mSCount - 2
        mSN(j) = mSN(j + 1)
        mSV(j) = mSV(j + 1)
    Next
    mSCount = mSCount - 1
    Touch
End Function

' Inline style property names in dash form (custom properties verbatim), in set order.
Public Function StyleNames() As clsList
    Dim list As New clsList, i As Long
    For i = 0 To mSCount - 1
        list.Add modDom.DomDashCss(mSN(i))
    Next
    Set StyleNames = list
End Function

' style.cssText = "a: b; c: d" (sets each declaration).
Public Property Let StyleText(ByVal cssText As String)
    Dim pieces() As String, k As Long, colon As Long, piece As String, propName As String, propValue As String
    pieces = Split(cssText, ";")
    For k = 0 To UBound(pieces)
        piece = pieces(k)
        colon = InStr(1, piece, ":")
        If colon > 1 Then
            propName = modDom.DomTrim(Left$(piece, colon - 1))
            propValue = modDom.DomTrim(Mid$(piece, colon + 1))
            SetStyle propName, propValue
        End If
    Next
End Property

' ================================================================ form controls

Public Property Get Value() As String
    Dim t As String, opt As clsNode
    If mType <> 1 Then Exit Property
    Select Case mLocal
    Case "input"
        t = TypeName
        If t = "checkbox" Or t = "radio" Then
            If HasAttrRaw("value") Then Value = AttrRaw("value") Else Value = "on"
        ElseIf mValueSet Then
            Value = mValue
        Else
            Value = AttrRaw("value")
        End If
    Case "textarea"
        If mValueSet Then Value = mValue Else Value = TextContent
    Case "select"
        Set opt = SelectedOptionNode()
        If Not opt Is Nothing Then Value = opt.Value
    Case "option"
        If HasAttrRaw("value") Then Value = AttrRaw("value") Else Value = Text
    Case "button"
        Value = AttrRaw("value")
    Case Else
        Value = mValue
    End Select
End Property

Public Property Let Value(ByVal text As String)
    Dim t As String, opts As clsList, i As Long, found As Boolean, opt As clsNode
    Select Case mLocal
    Case "input"
        t = TypeName
        If t = "checkbox" Or t = "radio" Then SetAttribute "value", text: Exit Property
        mValueSet = True
        mValue = text
    Case "textarea"
        mValueSet = True
        mValue = text
    Case "select"
        Set opts = Options
        For i = 0 To opts.Count - 1
            Set opt = opts.Item(i)
            If Not found And opt.Value = text Then
                opt.SetSelRaw 1
                found = True
            Else
                opt.SetSelRaw 0
            End If
        Next
        If Not mDirty Then modDom.DomStructBump 4
        mDirty = True
    Case "option", "button"
        SetAttribute "value", text
        Exit Property
    Case Else
        mValue = text
    End Select
    Touch
End Property

Public Property Get DefaultValue() As String
    If mLocal = "textarea" Then DefaultValue = TextContent Else DefaultValue = AttrRaw("value")
End Property

Public Property Let DefaultValue(ByVal text As String)
    If mLocal = "textarea" Then TextContent = text Else SetAttribute "value", text
End Property

Public Property Get Checked() As Boolean
    If mLocal = "input" And mType = 1 Then
        If mCheckedSet Then Checked = mChecked Else Checked = HasAttrRaw("checked")
    Else
        Checked = mChecked
    End If
End Property

Public Property Let Checked(ByVal value As Boolean)
    Dim root As clsNode, radios As clsList, i As Long, radio As clsNode, nm As String, was As Boolean
    was = Checked
    mCheckedSet = True
    mChecked = value
    If was = value Then Touch Else TouchStruct 4
    If mLocal <> "input" Then Exit Property
    If value And TypeName = "radio" Then
        nm = Name
        If Len(nm) = 0 Then Exit Property
        Set root = Closest("form")
        If root Is Nothing Then Set root = modDom.Doc
        Set radios = root.QuerySelectorAll("input[type=""radio""]")
        For i = 0 To radios.Count - 1
            Set radio = radios.Item(i)
            If Not radio Is Me Then
                If radio.Name = nm Then radio.SetCheckedRaw False
            End If
        Next
    End If
End Property

Friend Sub SetCheckedRaw(ByVal value As Boolean)
    Dim was As Boolean
    was = Checked
    mCheckedSet = True
    mChecked = value
    If was = value Then Touch Else TouchStruct 4
End Sub

Public Property Get DefaultChecked() As Boolean
    DefaultChecked = HasAttrRaw("checked")
End Property

Public Property Let DefaultChecked(ByVal value As Boolean)
    ToggleAttribute "checked", value
End Property

' ---- option

Friend Sub SetSelRaw(ByVal state As Long)
    Dim was As Long
    was = mSel
    mSel = state
    If was = state Then Touch Else TouchStruct 4
End Sub

Friend Property Get SelRaw() As Long
    SelRaw = mSel
End Property

Friend Sub SetDirtyRaw(ByVal value As Boolean)
    If mDirty <> value Then modDom.DomStructBump 4
    mDirty = value
End Sub

' The select that owns this option (through optgroups only), or Nothing.
Friend Function OwnerSelect() As clsNode
    Dim node As clsNode
    Set node = mParent
    Do While Not node Is Nothing
        If node.NodeType <> 1 Then Exit Function
        If node.LocalName = "select" Then Set OwnerSelect = node: Exit Function
        If node.LocalName <> "optgroup" Then Exit Function
        Set node = node.ParentNode
    Loop
End Function

Public Property Get Text() As String
    Text = modDom.DomCollapseSpace(TextContent)
End Property

Public Property Let Text(ByVal value As String)
    TextContent = value
End Property

Public Property Get Label() As String
    If HasAttrRaw("label") Then Label = AttrRaw("label") Else Label = Text
End Property

Public Property Get DefaultSelected() As Boolean
    DefaultSelected = HasAttrRaw("selected")
End Property

Public Property Let DefaultSelected(ByVal value As Boolean)
    ToggleAttribute "selected", value
End Property

Public Property Get Selected() As Boolean
    Dim sel As clsNode
    Set sel = OwnerSelect()
    If Not sel Is Nothing Then
        Selected = (sel.SelectedOptionNode() Is Me)
    ElseIf mSel >= 0 Then
        Selected = (mSel = 1)
    Else
        Selected = HasAttrRaw("selected")
    End If
End Property

Public Property Let Selected(ByVal value As Boolean)
    Dim sel As clsNode, opts As clsList, i As Long, opt As clsNode
    Set sel = OwnerSelect()
    If Not sel Is Nothing And value Then
        Set opts = sel.Options
        For i = 0 To opts.Count - 1
            Set opt = opts.Item(i)
            If opt Is Me Then opt.SetSelRaw 1 Else opt.SetSelRaw 0
        Next
    Else
        mSel = IIf(value, 1, 0)
        TouchStruct 4
    End If
    If Not sel Is Nothing Then sel.SetDirtyRaw True: sel.Touch
End Property

Public Property Get Index() As Long
    Dim sel As clsNode, opts As clsList, i As Long
    Set sel = OwnerSelect()
    If sel Is Nothing Then Exit Property
    Set opts = sel.Options
    For i = 0 To opts.Count - 1
        If opts.Item(i) Is Me Then Index = i: Exit Property
    Next
End Property

' ---- select

Private Sub CollectOptions(ByVal node As clsNode, ByVal list As clsList)
    Dim i As Long, kid As clsNode
    For i = 0 To node.KidCount - 1
        Set kid = node.Kid(i)
        If kid.NodeType = 1 Then
            If kid.LocalName = "option" Then
                If kid.OwnerSelect() Is Me Then list.Add kid
            End If
            CollectOptions kid, list
        End If
    Next
End Sub

Public Property Get Options() As clsList
    Dim list As New clsList
    CollectOptions Me, list
    Set Options = list
End Property

Friend Function SelectedOptionNode() As clsNode
    Dim opts As clsList, i As Long, opt As clsNode, s As Long, chosen As clsNode
    Set opts = Options
    For i = 0 To opts.Count - 1
        Set opt = opts.Item(i)
        s = opt.SelRaw
        If s < 0 Then
            If opt.HasAttrRaw("selected") Then Set chosen = opt
        ElseIf s = 1 Then
            Set chosen = opt
        End If
    Next
    If Not chosen Is Nothing Then Set SelectedOptionNode = chosen: Exit Function
    If mDirty Then Exit Function
    For i = 0 To opts.Count - 1
        Set opt = opts.Item(i)
        If Not opt.HasAttrRaw("disabled") Then Set SelectedOptionNode = opt: Exit Function
    Next
End Function

Public Property Get SelectedIndex() As Long
    Dim opt As clsNode, opts As clsList, i As Long
    Set opt = SelectedOptionNode()
    SelectedIndex = -1
    If opt Is Nothing Then Exit Property
    Set opts = Options
    For i = 0 To opts.Count - 1
        If opts.Item(i) Is opt Then SelectedIndex = i: Exit Property
    Next
End Property

Public Property Let SelectedIndex(ByVal index As Long)
    Dim opts As clsList, i As Long, opt As clsNode
    Set opts = Options
    For i = 0 To opts.Count - 1
        Set opt = opts.Item(i)
        If i = index Then opt.SetSelRaw 1 Else opt.SetSelRaw 0
    Next
    mDirty = True
    TouchStruct 4
End Property

Public Property Get SelectedOption() As clsNode
    Set SelectedOption = SelectedOptionNode()
End Property

' ================================================================ cloning

Friend Sub CopyStateFrom(ByVal source As clsNode)
    mValueSet = source.ValueSetRaw
    mValue = source.ValueRaw
    mCheckedSet = source.CheckedSetRaw
    mChecked = source.CheckedRaw
    mSel = source.SelRaw
    mDirty = source.DirtyRaw
    modDom.DomStructBump 4
End Sub

Friend Property Get ValueSetRaw() As Boolean
    ValueSetRaw = mValueSet
End Property

Friend Property Get ValueRaw() As String
    ValueRaw = mValue
End Property

Friend Property Get CheckedSetRaw() As Boolean
    CheckedSetRaw = mCheckedSet
End Property

Friend Property Get CheckedRaw() As Boolean
    CheckedRaw = mChecked
End Property

Friend Property Get DirtyRaw() As Boolean
    DirtyRaw = mDirty
End Property

Public Function CloneNode(Optional ByVal deep As Boolean = False) As clsNode
    Dim copy As clsNode, i As Long
    Select Case mType
    Case 3: Set CloneNode = modDom.CreateTextNode(mData): Exit Function
    Case 8: Set CloneNode = modDom.CreateComment(mData): Exit Function
    Case 11: Set copy = modDom.CreateFragment()
    Case 1
        Set copy = modDom.CreateElement(mLocal)
        For i = 0 To mACount - 1
            copy.SetAttribute mAN(i), mAV(i)
        Next
        For i = 0 To mDCount - 1
            copy.DataSet mDK(i), mDV(i)
        Next
        copy.CopyStateFrom Me
    Case Else
        Err.Raise vbObjectError + 703, "clsNode", "NotSupportedError: cannot clone this node"
    End Select
    If deep Then
        For i = 0 To mKidCount - 1
            copy.AppendChild mKids(i).CloneNode(True)
        Next
    End If
    Set CloneNode = copy
End Function

' ================================================================ queries

Public Function QuerySelector(ByVal selector As String) As clsNode
    Dim list As clsList
    Set list = modDom.DomQuery(Me, selector, True)
    If list.Count > 0 Then Set QuerySelector = list.Item(0)
End Function

Public Function QuerySelectorAll(ByVal selector As String) As clsList
    Set QuerySelectorAll = modDom.DomQuery(Me, selector, False)
End Function

Public Function GetElementsByTagName(ByVal tag As String) As clsList
    Set GetElementsByTagName = modDom.DomQuery(Me, tag, False)
End Function

Public Function Matches(ByVal selector As String) As Boolean
    If mType <> 1 Then Exit Function
    Matches = modDom.DomGetSelector(selector).MatchesNode(Me, Me)
End Function

Public Function Closest(ByVal selector As String) As clsNode
    Dim sel As clsDomSelector, node As clsNode
    Set sel = modDom.DomGetSelector(selector)
    Set node = Me
    Do While Not node Is Nothing
        If node.NodeType <> 1 Then Exit Function
        If sel.MatchesNode(node, node) Then Set Closest = node: Exit Function
        Set node = node.ParentNode
    Loop
End Function

Public Function GetElementById(ByVal elementId As String) As clsNode
    If mType = 9 Then
        Set GetElementById = modDom.DomLookupId(elementId)
    Else
        Set GetElementById = QuerySelector("#" & modDom.CssEscape(elementId))
    End If
End Function

' ================================================================ focus and activation

Public Sub Focus()
    If mType = 1 Then modDom.DomSetActive Me
End Sub

Public Sub Blur()
    If modDom.DomRawActive() Is Me Then modDom.DomSetActive Nothing
End Sub

Public Sub Click()
    Dim evt As clsEvent, control As clsNode, parent As clsNode
    If mType <> 1 Then Exit Sub
    If mLocal = "input" Then InputClick: Exit Sub
    If Disabled Then Exit Sub
    Set evt = modDom.NewMouseEvent("click")
    If Not DispatchEvent(evt) Then Exit Sub
    If mLocal = "summary" Then
        Set parent = mParent
        If Not parent Is Nothing Then
            If parent.NodeType = 1 And parent.LocalName = "details" Then parent.IsOpen = Not parent.IsOpen
        End If
    ElseIf mLocal = "label" Then
        Set control = LabelControl
        If Not control Is Nothing Then
            If Not control Is Me Then control.Click
        End If
    End If
End Sub

Public Property Get LabelControl() As clsNode
    Dim target As String
    target = HtmlFor
    If Len(target) > 0 Then
        Set LabelControl = modDom.DomLookupId(target)
    Else
        Set LabelControl = QuerySelector("input, select, textarea, button")
    End If
End Property

Private Sub InputClick()
    Dim t As String, toggles As Boolean, before As Boolean, ok As Boolean
    If Disabled Then Exit Sub
    t = TypeName
    toggles = (t = "checkbox" Or t = "radio")
    before = Checked
    If t = "checkbox" Then
        Checked = Not before
    ElseIf t = "radio" Then
        Checked = True
    End If
    ok = DispatchEvent(modDom.NewMouseEvent("click"))
    If Not toggles Then Exit Sub
    If Not ok Then Checked = before: Exit Sub
    If Checked <> before Then
        DispatchEvent modDom.NewEvent("input", True, False)
        DispatchEvent modDom.NewEvent("change", True, False)
    End If
End Sub

' ================================================================ events

Private Function FindListener(ByRef evType As String, ByRef handler As String, ByVal capture As Boolean) As Long
    Dim i As Long
    For i = 0 To mLCount - 1
        If mLType(i) = evType Then
            If mLHandler(i) = handler And mLCapture(i) = capture Then FindListener = i: Exit Function
        End If
    Next
    FindListener = -1
End Function

Public Sub AddEventListener(ByVal evType As String, ByVal handler As String, Optional ByVal capture As Boolean = False, Optional ByVal once As Boolean = False)
    If Len(handler) = 0 Then Exit Sub
    If FindListener(evType, handler, capture) >= 0 Then Exit Sub
    If mLCount = 0 Then
        ReDim mLId(0 To 3): ReDim mLType(0 To 3): ReDim mLHandler(0 To 3): ReDim mLCapture(0 To 3): ReDim mLOnce(0 To 3)
    ElseIf mLCount > UBound(mLId) Then
        ReDim Preserve mLId(0 To mLCount * 2): ReDim Preserve mLType(0 To mLCount * 2)
        ReDim Preserve mLHandler(0 To mLCount * 2): ReDim Preserve mLCapture(0 To mLCount * 2)
        ReDim Preserve mLOnce(0 To mLCount * 2)
    End If
    mLId(mLCount) = modDom.DomNextUid()
    mLType(mLCount) = evType
    mLHandler(mLCount) = handler
    mLCapture(mLCount) = capture
    mLOnce(mLCount) = once
    mLCount = mLCount + 1
End Sub

Public Sub RemoveEventListener(ByVal evType As String, ByVal handler As String, Optional ByVal capture As Boolean = False)
    Dim i As Long
    i = FindListener(evType, handler, capture)
    If i >= 0 Then RemoveListenerAt i
End Sub

Private Sub RemoveListenerAt(ByVal i As Long)
    Dim j As Long
    For j = i To mLCount - 2
        mLId(j) = mLId(j + 1)
        mLType(j) = mLType(j + 1)
        mLHandler(j) = mLHandler(j + 1)
        mLCapture(j) = mLCapture(j + 1)
        mLOnce(j) = mLOnce(j + 1)
    Next
    mLCount = mLCount - 1
End Sub

Private Function ListenerIndex(ByVal listenerId As Long) As Long
    Dim i As Long
    For i = 0 To mLCount - 1
        If mLId(i) = listenerId Then ListenerIndex = i: Exit Function
    Next
    ListenerIndex = -1
End Function

Public Function HasListeners(ByVal evType As String) As Boolean
    Dim i As Long
    For i = 0 To mLCount - 1
        If mLType(i) = evType Then HasListeners = True: Exit Function
    Next
End Function

' Runs this target's listeners for one phase (1 capture, 3 bubble) on a snapshot.
Friend Sub InvokeListeners(ByVal evt As clsEvent, ByVal phase As Long)
    Dim ids() As Long, n As Long, i As Long, k As Long, evType As String
    If mLCount = 0 Then Exit Sub
    evType = evt.EventType
    ReDim ids(0 To mLCount - 1)
    For i = 0 To mLCount - 1
        If mLType(i) = evType Then ids(n) = mLId(i): n = n + 1
    Next
    For i = 0 To n - 1
        k = ListenerIndex(ids(i))
        If k >= 0 Then
            If Not ((phase = 1 And Not mLCapture(k)) Or (phase = 3 And mLCapture(k))) Then
                Dim handler As String
                handler = mLHandler(k)
                If mLOnce(k) Then RemoveListenerAt k
                evt.SetCurrent Me
                modDom.DomRunHandler handler, evt
                If evt.StoppedNow Then Exit Sub
            End If
        End If
    Next
End Sub

Public Function DispatchEvent(ByVal evt As clsEvent) As Boolean
    Dim path() As clsNode, n As Long, node As clsNode, i As Long
    ReDim path(0 To 15)
    Set node = Me
    Do While Not node Is Nothing
        If n > UBound(path) Then ReDim Preserve path(0 To n * 2)
        Set path(n) = node
        n = n + 1
        If node.NodeType = 9 Then
            Set node = modDom.Win
        Else
            Set node = node.ParentNode
        End If
    Loop
    evt.BeginDispatch Me
    For i = n - 1 To 1 Step -1
        If evt.Stopped Then Exit For
        evt.SetPhase 1
        path(i).InvokeListeners evt, 1
    Next
    If Not evt.Stopped Then
        evt.SetPhase 2
        InvokeListeners evt, 1
        If Not evt.StoppedNow Then InvokeListeners evt, 3
    End If
    If evt.Bubbles Then
        For i = 1 To n - 1
            If evt.Stopped Then Exit For
            evt.SetPhase 3
            path(i).InvokeListeners evt, 3
        Next
    End If
    evt.SetPhase 0
    evt.SetCurrent Nothing
    DispatchEvent = Not evt.DefaultPrevented
End Function
