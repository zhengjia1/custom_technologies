VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDomSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' A parsed CSS selector list (port of minidom's parseSelector and matchers).
' Supports type/universal, #id, .class, [attr], [attr op value] with = ~= |= ^= $= *=,
' combinators (descendant, >, +, ~), selector lists, and the pseudo-classes
' :not() :is() :where() :has() :scope :checked :disabled :enabled :empty :first-child
' :last-child :only-child :nth-child() :first-of-type :focus :focus-visible :focus-within
' :placeholder-shown :hover :active :visited :link :any-link :root.
' Obtain instances through modDom.DomGetSelector (cached by selector text).

Private Type TAttr
    Name As String
    Op As String
    Value As String
    IsData As Boolean     ' data-* attributes live in the dataset (camel-cased key)
    DataKey As String
End Type

Private Type TPseudo
    Name As String
    Arg As String
    Inner As clsDomSelector
    NthKind As Long       ' 0 exact, 1 odd, 2 even, 3 an+b, 4 never
    NthA As Long
    NthB As Long
End Type

Private Type TCompound
    Tag As String
    ClsPadded() As String ' " cls " per class, for the snapshot prefilter
    IdCount As Long
    Ids() As String
    ClsCount As Long
    Classes() As String
    AttrCount As Long
    Attrs() As TAttr
    PsCount As Long
    Pseudos() As TPseudo
    Comb As String        ' combinator joining this compound to the next one leftwards
End Type

Private Type TComplex
    Count As Long
    Parts() As TCompound  ' right-to-left
End Type

Private mList() As TComplex
Private mCount As Long
Private mSource As String

' parse state
Private mT As String
Private mC() As Long
Private mN As Long
Private mP As Long

' query state
Private mFound As Boolean
' False when a pseudo-class depends on state outside modDom's structural version.
Private mCacheable As Boolean
' modDom change counters the matches depend on: 2 class/id, 4 attributes/dataset,
' 8 form state (element insertion/removal always counts).
Private mDeps As Long

Public Property Get Source() As String
    Source = mSource
End Property

' ================================================================ parsing

Friend Sub Parse(ByVal text As String)
    mSource = text
    ParseFrom text, 0
    SkipSpace
    If mP < mN Then Fail
End Sub

' True when query results depend only on what modDom.DomStructBump tracks (elements,
' attributes, dataset, checked/selected state), so they may be cached between changes.
Friend Property Get Cacheable() As Boolean
    Cacheable = mCacheable
End Property

Friend Property Get Dependencies() As Long
    Dependencies = mDeps
End Property

' Parses a selector list starting at 0-based pos; returns the position after it.
Friend Function ParseFrom(ByVal text As String, ByVal pos As Long) As Long
    Dim b() As Byte, i As Long
    mT = text
    mN = Len(text)
    If mN > 0 Then
        b = text
        ReDim mC(0 To mN - 1)
        For i = 0 To mN - 1
            mC(i) = b(2 * i) + 256& * b(2 * i + 1)
        Next
    End If
    mP = pos
    If Len(mSource) = 0 Then mSource = text
    mCount = 0
    mCacheable = True
    mDeps = 0
    Do
        SkipSpace
        If mCount = 0 Then
            ReDim mList(0 To 3)
        ElseIf mCount > UBound(mList) Then
            ReDim Preserve mList(0 To mCount * 2)
        End If
        ParseComplex mList(mCount)
        mCount = mCount + 1
        SkipSpace
        If Ch(mP) = 44 Then
            mP = mP + 1
        Else
            Exit Do
        End If
    Loop
    ParseFrom = mP
End Function

Private Sub Fail()
    Err.Raise vbObjectError + 710, "clsDomSelector", "SyntaxError: Unsupported selector: " & mT
End Sub

Private Function Ch(ByVal i As Long) As Long
    If i < mN And i >= 0 Then Ch = mC(i) Else Ch = -1
End Function

Private Sub SkipSpace()
    Do While mP < mN
        If Not modDom.DomIsSpace(mC(mP)) Then Exit Do
        mP = mP + 1
    Loop
End Sub

Private Function IsHex(ByVal c As Long) As Boolean
    IsHex = (c >= 48 And c <= 57) Or (c >= 65 And c <= 70) Or (c >= 97 And c <= 102)
End Function

Private Function HexValue(ByVal c As Long) As Long
    If c >= 48 And c <= 57 Then
        HexValue = c - 48
    ElseIf c >= 65 And c <= 70 Then
        HexValue = c - 55
    Else
        HexValue = c - 87
    End If
End Function

Private Function ReadIdentifier() As String
    Dim result As String, c As Long, k As Long, j As Long, cp As Long
    Do While mP < mN
        c = mC(mP)
        If c = 92 Then
            k = 0
            Do While k < 6
                If Not IsHex(Ch(mP + 1 + k)) Then Exit Do
                k = k + 1
            Loop
            If k > 0 Then
                cp = 0
                For j = 1 To k
                    cp = cp * 16 + HexValue(mC(mP + j))
                Next
                result = result & modDom.DomFromCodePoint(cp)
                mP = mP + 1 + k
                If Ch(mP) = 32 Then mP = mP + 1
            Else
                If mP + 1 < mN Then result = result & ChrW$(mC(mP + 1))
                mP = mP + 2
            End If
        ElseIf (c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Or (c >= 48 And c <= 57) Or c = 95 Or c = 45 Or c >= 128 Then
            result = result & ChrW$(c)
            mP = mP + 1
        Else
            Exit Do
        End If
    Loop
    ReadIdentifier = result
End Function

Private Sub ParseComplex(ByRef cx As TComplex)
    Dim comb As String, before As Long, c As Long, i As Long, n As Long
    Dim tmp() As TCompound
    ReDim tmp(0 To 3)
    Do
        SkipSpace
        If n > UBound(tmp) Then ReDim Preserve tmp(0 To n * 2)
        ParseCompound tmp(n)
        tmp(n).Comb = comb
        n = n + 1
        before = mP
        SkipSpace
        c = Ch(mP)
        If c = 62 Or c = 43 Or c = 126 Then
            comb = ChrW$(c)
            mP = mP + 1
        ElseIf mP < mN And c <> 44 And c <> 41 Then
            comb = " "
        Else
            Exit Do
        End If
    Loop
    cx.Count = n
    ReDim cx.Parts(0 To n - 1)
    For i = 0 To n - 1
        cx.Parts(i) = tmp(n - 1 - i)
    Next
End Sub

Private Sub ParseCompound(ByRef cp As TCompound)
    Dim c As Long, anything As Boolean, nm As String, op As String, value As String, q As Long
    Dim closeAt As Long, inner As clsDomSelector
    c = Ch(mP)
    If c = 42 Then
        mP = mP + 1
        anything = True
    ElseIf (c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Then
        cp.Tag = LCase$(ReadIdentifier())
        anything = True
    End If
    Do While mP < mN
        c = mC(mP)
        If c = 35 Then
            mP = mP + 1
            If cp.IdCount = 0 Then
                ReDim cp.Ids(0 To 1)
            ElseIf cp.IdCount > UBound(cp.Ids) Then
                ReDim Preserve cp.Ids(0 To cp.IdCount * 2)
            End If
            cp.Ids(cp.IdCount) = ReadIdentifier()
            cp.IdCount = cp.IdCount + 1
            anything = True
        ElseIf c = 46 Then
            mP = mP + 1
            If cp.ClsCount = 0 Then
                ReDim cp.Classes(0 To 1)
            ElseIf cp.ClsCount > UBound(cp.Classes) Then
                ReDim Preserve cp.Classes(0 To cp.ClsCount * 2)
            End If
            cp.Classes(cp.ClsCount) = ReadIdentifier()
            If cp.ClsCount = 0 Then
                ReDim cp.ClsPadded(0 To UBound(cp.Classes))
            ElseIf UBound(cp.ClsPadded) < UBound(cp.Classes) Then
                ReDim Preserve cp.ClsPadded(0 To UBound(cp.Classes))
            End If
            cp.ClsPadded(cp.ClsCount) = " " & cp.Classes(cp.ClsCount) & " "
            cp.ClsCount = cp.ClsCount + 1
            anything = True
        ElseIf c = 91 Then
            mP = mP + 1
            SkipSpace
            nm = LCase$(ReadIdentifier())
            SkipSpace
            op = ""
            value = ""
            c = Ch(mP)
            If c = 61 Then
                op = "="
            ElseIf (c = 126 Or c = 124 Or c = 94 Or c = 36 Or c = 42) And Ch(mP + 1) = 61 Then
                op = ChrW$(c) & "="
            End If
            If Len(op) > 0 Then
                mP = mP + Len(op)
                SkipSpace
                q = Ch(mP)
                If q = 34 Or q = 39 Then
                    mP = mP + 1
                    Do While mP < mN
                        If mC(mP) = q Then Exit Do
                        If mC(mP) = 92 Then
                            If mP + 1 < mN Then value = value & ChrW$(mC(mP + 1))
                            mP = mP + 2
                        Else
                            value = value & ChrW$(mC(mP))
                            mP = mP + 1
                        End If
                    Loop
                    mP = mP + 1
                Else
                    value = ReadIdentifier()
                End If
                SkipSpace
            End If
            If Ch(mP) <> 93 Then Fail
            mP = mP + 1
            If cp.AttrCount = 0 Then
                ReDim cp.Attrs(0 To 1)
            ElseIf cp.AttrCount > UBound(cp.Attrs) Then
                ReDim Preserve cp.Attrs(0 To cp.AttrCount * 2)
            End If
            cp.Attrs(cp.AttrCount).Name = nm
            cp.Attrs(cp.AttrCount).Op = op
            cp.Attrs(cp.AttrCount).Value = value
            If Left$(nm, 5) = "data-" Then
                cp.Attrs(cp.AttrCount).IsData = True
                cp.Attrs(cp.AttrCount).DataKey = modDom.DomCamelData(nm)
            End If
            cp.AttrCount = cp.AttrCount + 1
            anything = True
        ElseIf c = 58 Then
            mP = mP + 1
            If Ch(mP) = 58 Then mP = mP + 1
            If cp.PsCount = 0 Then
                ReDim cp.Pseudos(0 To 1)
            ElseIf cp.PsCount > UBound(cp.Pseudos) Then
                ReDim Preserve cp.Pseudos(0 To cp.PsCount * 2)
            End If
            nm = LCase$(ReadIdentifier())
            cp.Pseudos(cp.PsCount).Name = nm
            If Ch(mP) = 40 Then
                mP = mP + 1
                If nm = "not" Or nm = "is" Or nm = "where" Or nm = "has" Then
                    Set inner = New clsDomSelector
                    mP = inner.ParseFrom(mT, mP)
                    Set cp.Pseudos(cp.PsCount).Inner = inner
                Else
                    closeAt = InStr(mP + 1, mT, ")")
                    If closeAt = 0 Then Fail
                    cp.Pseudos(cp.PsCount).Arg = modDom.DomTrim(Mid$(mT, mP + 1, closeAt - 1 - mP))
                    mP = closeAt - 1
                End If
                SkipSpace
                If Ch(mP) <> 41 Then Fail
                mP = mP + 1
            End If
            CheckPseudo cp.Pseudos(cp.PsCount)
            cp.PsCount = cp.PsCount + 1
            anything = True
        Else
            Exit Do
        End If
    Loop
    If Not anything Then Fail
    If cp.ClsCount > 0 Or cp.IdCount > 0 Then mDeps = mDeps Or 2
    If cp.AttrCount > 0 Then mDeps = mDeps Or 4
End Sub

Private Sub CheckPseudo(ByRef ps As TPseudo)
    Select Case ps.Name
    Case "not", "is", "where", "has"
        If ps.Inner Is Nothing Then Fail
        If Not ps.Inner.Cacheable Then mCacheable = False
        mDeps = mDeps Or ps.Inner.Dependencies
    Case "nth-child"
        ParseNth ps
    Case "checked", "disabled", "enabled"
        mDeps = mDeps Or 4 Or 8
    Case "link", "any-link"
        mDeps = mDeps Or 4
    Case "scope", "first-child", "last-child", "only-child", "first-of-type", "hover", "active", "visited", "root"
    Case "empty", "focus", "focus-visible", "focus-within", "placeholder-shown"
        ' text, focus and value changes do not bump the structural version
        mCacheable = False
    Case Else
        ' unknown pseudo-classes fail when matched, like minidom
        mCacheable = False
    End Select
End Sub

Private Function AllDigits(ByVal s As String) As Boolean
    Dim i As Long, c As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1))
        If c < 48 Or c > 57 Then Exit Function
    Next
    AllDigits = True
End Function

Private Sub ParseNth(ByRef ps As TPseudo)
    Dim text As String, i As Long, c As String, at As Long, head As String, tail As String, sign As String
    For i = 1 To Len(ps.Arg)
        c = Mid$(ps.Arg, i, 1)
        If Not modDom.DomIsSpace(AscW(c)) Then text = text & c
    Next
    text = LCase$(text)
    If text = "odd" Then ps.NthKind = 1: Exit Sub
    If text = "even" Then ps.NthKind = 2: Exit Sub
    at = InStr(1, text, "n")
    If at > 0 Then
        head = Left$(text, at - 1)
        tail = Mid$(text, at + 1)
        sign = ""
        If Left$(head, 1) = "+" Or Left$(head, 1) = "-" Then sign = Left$(head, 1): head = Mid$(head, 2)
        If (Len(head) = 0 Or AllDigits(head)) And (Len(tail) = 0 Or ((Left$(tail, 1) = "+" Or Left$(tail, 1) = "-") And AllDigits(Mid$(tail, 2)))) Then
            ps.NthKind = 3
            If Len(head) = 0 Then ps.NthA = 1 Else ps.NthA = CLng(head)
            If sign = "-" Then ps.NthA = -ps.NthA
            If Len(tail) > 0 Then ps.NthB = CLng(tail)
            Exit Sub
        End If
    End If
    ' Number(text): only plain integers are modelled; anything else never matches.
    ps.NthKind = 0
    If Len(text) = 0 Then
        ps.NthB = 0
    ElseIf AllDigits(text) Then
        ps.NthB = CLng(text)
    ElseIf (Left$(text, 1) = "+" Or Left$(text, 1) = "-") And AllDigits(Mid$(text, 2)) Then
        ps.NthB = CLng(text)
    Else
        ps.NthKind = 4
    End If
End Sub

' ================================================================ matching

Friend Function MatchesNode(ByVal el As clsNode, ByVal scope As clsNode) As Boolean
    Dim i As Long
    For i = 0 To mCount - 1
        If MatchComplex(el, i, 0, scope) Then MatchesNode = True: Exit Function
    Next
End Function

Private Function MatchComplex(ByRef node As clsNode, ByVal ci As Long, ByVal pi As Long, ByRef scope As clsNode) As Boolean
    Dim other As clsNode
    If Not MatchCompound(node, mList(ci).Parts(pi), scope) Then Exit Function
    If pi = mList(ci).Count - 1 Then MatchComplex = True: Exit Function
    Select Case mList(ci).Parts(pi).Comb
    Case ">"
        Set other = node.ParentNode
        If Not other Is Nothing Then
            If other.NodeType = 1 Then MatchComplex = MatchComplex(other, ci, pi + 1, scope)
        End If
    Case " "
        Set other = node.ParentNode
        Do While Not other Is Nothing
            If other.NodeType <> 1 Then Exit Do
            If MatchComplex(other, ci, pi + 1, scope) Then MatchComplex = True: Exit Function
            Set other = other.ParentNode
        Loop
    Case "+"
        Set other = node.PreviousElementSibling
        If Not other Is Nothing Then MatchComplex = MatchComplex(other, ci, pi + 1, scope)
    Case "~"
        Set other = node.PreviousElementSibling
        Do While Not other Is Nothing
            If MatchComplex(other, ci, pi + 1, scope) Then MatchComplex = True: Exit Function
            Set other = other.PreviousElementSibling
        Loop
    End Select
End Function

Private Function MatchCompound(ByRef node As clsNode, ByRef cp As TCompound, ByRef scope As clsNode) As Boolean
    Dim i As Long, actual As String, expected As String
    If Len(cp.Tag) > 0 Then
        If Not node.TagIs(cp.Tag) Then Exit Function
    End If
    For i = 0 To cp.ClsCount - 1
        If Not node.HasClassRef(cp.Classes(i)) Then Exit Function
    Next
    For i = 0 To cp.IdCount - 1
        If Not node.HasIdAttr Then Exit Function
        If node.Id <> cp.Ids(i) Then Exit Function
    Next
    For i = 0 To cp.AttrCount - 1
        If cp.Attrs(i).IsData Then
            If Not node.DataLookup(cp.Attrs(i).DataKey, actual) Then Exit Function
        Else
            If Not node.AttrLookupPlain(cp.Attrs(i).Name, actual) Then Exit Function
        End If
        If Len(cp.Attrs(i).Op) > 0 Then
            expected = cp.Attrs(i).Value
            Select Case cp.Attrs(i).Op
            Case "="
                If actual <> expected Then Exit Function
            Case "~="
                If Not WordListHas(actual, expected) Then Exit Function
            Case "^="
                If Len(expected) = 0 Then Exit Function
                If Left$(actual, Len(expected)) <> expected Then Exit Function
            Case "$="
                If Len(expected) = 0 Then Exit Function
                If Len(actual) < Len(expected) Then Exit Function
                If Right$(actual, Len(expected)) <> expected Then Exit Function
            Case "*="
                If Len(expected) = 0 Then Exit Function
                If InStr(1, actual, expected, vbBinaryCompare) = 0 Then Exit Function
            Case "|="
                If actual <> expected And Left$(actual, Len(expected) + 1) <> expected & "-" Then Exit Function
            End Select
        End If
    Next
    For i = 0 To cp.PsCount - 1
        If Not MatchPseudo(node, cp.Pseudos(i), scope) Then Exit Function
    Next
    MatchCompound = True
End Function

' actual.split(/\s+/).indexOf(expected) >= 0 (empty edge tokens included, like JS).
Private Function WordListHas(ByRef actual As String, ByRef expected As String) As Boolean
    Dim i As Long, n As Long, start As Long, c As Long, inSpace As Boolean
    n = Len(actual)
    start = 1
    i = 1
    Do While i <= n
        c = AscW(Mid$(actual, i, 1))
        If modDom.DomIsSpace(c) Then
            If Mid$(actual, start, i - start) = expected Then WordListHas = True: Exit Function
            Do While i <= n
                If Not modDom.DomIsSpace(AscW(Mid$(actual, i, 1))) Then Exit Do
                i = i + 1
            Loop
            start = i
        Else
            i = i + 1
        End If
    Loop
    WordListHas = (Mid$(actual, start) = expected)
End Function

Private Function MatchPseudo(ByRef node As clsNode, ByRef ps As TPseudo, ByRef scope As clsNode) As Boolean
    Dim parent As clsNode, t As String, pos As Long, i As Long, other As clsNode, d As Long
    Select Case ps.Name
    Case "not": MatchPseudo = Not ps.Inner.MatchesNode(node, scope)
    Case "is", "where": MatchPseudo = ps.Inner.MatchesNode(node, scope)
    Case "has": MatchPseudo = ps.Inner.HasDescendantMatch(node)
    Case "scope": MatchPseudo = (node Is scope)
    Case "checked"
        If node.LocalName = "option" Then
            MatchPseudo = node.Selected
        ElseIf node.LocalName = "input" Then
            t = node.TypeName
            If t = "checkbox" Or t = "radio" Then MatchPseudo = node.Checked
        End If
    Case "disabled": MatchPseudo = node.Disabled
    Case "enabled": If node.IsFormElement() Then MatchPseudo = Not node.Disabled
    Case "empty"
        For i = 0 To node.KidCount - 1
            If node.Kid(i).NodeType <> 8 Then Exit Function
        Next
        MatchPseudo = True
    Case "first-child"
        Set parent = node.ParentNode
        If Not parent Is Nothing Then MatchPseudo = (parent.FirstElementChild Is node)
    Case "last-child"
        Set parent = node.ParentNode
        If Not parent Is Nothing Then MatchPseudo = (parent.LastElementChild Is node)
    Case "only-child"
        Set parent = node.ParentNode
        If Not parent Is Nothing Then MatchPseudo = (parent.ChildElementCount = 1)
    Case "nth-child"
        Set parent = node.ParentNode
        If parent Is Nothing Then
            pos = 1
        Else
            For i = 0 To node.IndexInParent
                If parent.Kid(i).NodeType = 1 Then pos = pos + 1
            Next
        End If
        Select Case ps.NthKind
        Case 1: MatchPseudo = (pos Mod 2 = 1)
        Case 2: MatchPseudo = (pos Mod 2 = 0)
        Case 0: MatchPseudo = (pos = ps.NthB)
        Case 3
            If ps.NthA = 0 Then
                MatchPseudo = (pos = ps.NthB)
            Else
                d = pos - ps.NthB
                MatchPseudo = ((d / ps.NthA) >= 0 And (d Mod ps.NthA) = 0)
            End If
        End Select
    Case "first-of-type"
        Set parent = node.ParentNode
        If Not parent Is Nothing Then
            For i = 0 To parent.KidCount - 1
                Set other = parent.Kid(i)
                If other.NodeType = 1 Then
                    If other.LocalName = node.LocalName Then MatchPseudo = (other Is node): Exit Function
                End If
            Next
        End If
    Case "focus", "focus-visible": MatchPseudo = (modDom.DomRawActive() Is node)
    Case "focus-within": MatchPseudo = node.Contains(modDom.DomRawActive())
    Case "placeholder-shown"
        If node.HasValueProperty() Then MatchPseudo = (Len(node.Value) = 0 And node.HasAttribute("placeholder"))
    Case "hover", "active", "visited": MatchPseudo = False
    Case "link", "any-link": MatchPseudo = (node.LocalName = "a" And node.HasAttrRaw("href"))
    Case "root": MatchPseudo = (node Is modDom.DocumentElement)
    Case Else: Err.Raise vbObjectError + 711, "clsDomSelector", "SyntaxError: Unsupported pseudo-class :" & ps.Name
    End Select
End Function

' :has(): some complex selector matches a descendant, with the subject as :scope.
Friend Function HasDescendantMatch(ByVal subject As clsNode) As Boolean
    HasDescendantMatch = WalkHas(subject, subject)
End Function

Private Function WalkHas(ByVal node As clsNode, ByVal subject As clsNode) As Boolean
    Dim i As Long, kid As clsNode, k As Long
    For i = 0 To node.KidCount - 1
        Set kid = node.Kid(i)
        If kid.NodeType = 1 Then
            For k = 0 To mCount - 1
                If MatchComplex(kid, k, 0, subject) Then WalkHas = True: Exit Function
            Next
            If WalkHas(kid, subject) Then WalkHas = True: Exit Function
        End If
    Next
End Function

' ================================================================ querying

' Like Collect over a flat list of elements in document order (modDom's snapshot):
' elements failing the tag/class/id tests of every rightmost compound are skipped
' without a call; the others get the full match. scope is Nothing (document root).
Friend Sub CollectFlat(ByRef nodes() As clsNode, ByRef tags() As String, ByRef classKeys() As String, _
        ByRef ids() As String, ByVal firstIndex As Long, ByVal lastIndex As Long, ByVal scope As clsNode, _
        ByVal firstOnly As Boolean, ByVal results As clsList)
    Dim i As Long, k As Long, j As Long, ok As Boolean
    For i = firstIndex To lastIndex
        For k = 0 To mCount - 1
            ok = True
            With mList(k).Parts(0)
                If Len(.Tag) > 0 Then ok = (tags(i) = .Tag)
                If ok Then
                    For j = 0 To .ClsCount - 1
                        If InStr(1, classKeys(i), .ClsPadded(j), vbBinaryCompare) = 0 Then ok = False: Exit For
                    Next
                End If
                If ok Then
                    For j = 0 To .IdCount - 1
                        If ids(i) <> .Ids(j) Then ok = False: Exit For
                    Next
                End If
            End With
            If ok Then
                If MatchComplex(nodes(i), k, 0, scope) Then
                    results.Add nodes(i)
                    If firstOnly Then Exit Sub
                    Exit For
                End If
            End If
        Next
    Next
End Sub

' Appends matching descendants of root (document order) to results.
Friend Sub Collect(ByVal root As clsNode, ByVal firstOnly As Boolean, ByVal results As clsList)
    Dim scope As clsNode
    If root.NodeType = 1 Then Set scope = root
    mFound = False
    Walk root, scope, firstOnly, results
End Sub

Private Sub Walk(ByVal node As clsNode, ByVal scope As clsNode, ByVal firstOnly As Boolean, ByVal results As clsList)
    Dim i As Long, n As Long, kid As clsNode, k As Long
    n = node.KidCount
    For i = 0 To n - 1
        Set kid = node.Kid(i)
        If kid.NodeType = 1 Then
            For k = 0 To mCount - 1
                If MatchComplex(kid, k, 0, scope) Then
                    results.Add kid
                    If firstOnly Then mFound = True: Exit Sub
                    Exit For
                End If
            Next
            If kid.KidCount > 0 Then
                Walk kid, scope, firstOnly, results
                If mFound Then Exit Sub
            End If
        End If
    Next
End Sub
