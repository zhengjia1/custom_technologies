VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRenderLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Lays out the live DOM the way the website's stylesheet would (a port of the reference
' engine excel/.build/obsolete/src/layout.js): block flow with sibling margin collapsing,
' inline text wrapping, flex, grid and simple tables (clsRenderGrid), form controls,
' absolute and fixed positioning, and the window's regions (title bar, library, page,
' overlays). Text is measured by modRender with real MSForms labels. The resulting
' fragment tree is painted into a display list by clsRenderPaint. CSS px throughout.
Option Explicit

Private Const SCROLLBAR As Double = 17
Private Const NOV As Double = -1E+30

Private Type TTok
    T As Long               ' 1 word, 2 space, 3 forced break, 4 absolute box, 5 atomic inline
    Text As String
    Slot As Long
    Click As clsNode
    Wrap As Boolean
    Anywhere As Boolean
    Collapsible As Boolean
    Box As clsRenderBox
    W As Double
    Frag As clsRenderBox
End Type

Private Type TFlex
    Box As clsRenderBox
    IsText As Boolean
    Index As Long
    Base As Double
    Hyp As Double
    MinW As Double
    MaxW As Double
    Grow As Double
    Shrink As Double
    Margins As Double
    Size As Double
    Frozen As Boolean
    MLAuto As Boolean
    MRAuto As Boolean
    AutoLeft As Double
    AutoRight As Double
End Type

Private mW As Double, mH As Double
Private mGrid As clsRenderGrid
Private mFixed() As clsRenderBox
Private mFixedCount As Long
Private mSeenFixed As clsMap
Private mTop As clsNode, mLibrary As clsNode, mMain As clsNode, mWorkspace As clsNode, mResizer As clsNode, mShell As clsNode
Private mMobile As Boolean
Private mRoot As clsRenderBox
Private mSerial As Long
' intrinsic sizes cached per pass: key -> index into the arrays
Private mIntrMap As clsMap
Private mIntrMin() As Double, mIntrMax() As Double, mIntrCount As Long
' measurement caches (persist across passes)
Private mWidths As clsMap
Private mLineHeights As clsMap
Private mFontKeys() As String       ' per style slot: font key, valid while the layout stamp matches
Private mFontKeyStamp() As Long

' Incremental layout. Before each pass the visible DOM is walked (StyleOf/PseudoOf on
' everything layout will style, so slots stay alive) and each node's state (version, layout
' stamps of its slots) is compared with the last walk: a node whose subtree changed loses
' its cached fragments and intrinsic sizes. LayoutBox then returns the cached fragment of
' an unchanged subtree laid out with the same arguments (fields its callers adjust are
' restored). Viewport, theme, root inline style (custom properties) and rule-table changes
' clear everything.
Private mPassNo As Long
Private mCacheOn As Boolean
Private mGlobalKey As String
Private mNoneIdx As Long, mListItemIdx As Long
Private mUidCap As Long
Private mWalkPass() As Long         ' by uid: pass of the last walk that reached the node
Private mStKnown() As Boolean       ' by uid: recorded state below is set
Private mStVer() As Long, mStSelf() As Long, mStBefore() As Long, mStAfter() As Long, mStParent() As Long
' by k = uid * 3 + which + 1: first cache entry, intrinsic sizes
Private mEHead() As Long
Private mIHas() As Boolean
Private mIMn() As Double, mIMx() As Double
' cache entries (1-based; 0 = none)
Private mECap As Long, mEFree() As Long, mEFreeCount As Long, mEUsed As Long
Private mENext() As Long, mEFrag() As clsRenderBox, mEFlags() As Long
Private mECbw() As Double, mEOptW() As Double, mEOptH() As Double, mECbh() As Double
Private mESave() As Double          ' 8 per entry: X Y H AutoTop AutoBottom AutoLeft AutoRight JustifyOffset
Private mEPos() As Boolean, mERole() As Long, mEClick() As clsNode, mEFixed() As Collection
Public CacheHits As Long, CacheMisses As Long, WalkedNodes As Long, DirtyNodes As Long

Private Sub Class_Initialize()
    Set mWidths = NewMap()
    Set mLineHeights = NewMap()
    ReDim mFontKeys(0 To 255)
    ReDim mFontKeyStamp(0 To 255)
    ClearLayoutCache
    Set mGrid = New clsRenderGrid
    Set mGrid.Engine = Me
End Sub

Private Sub Class_Terminate()
    If Not mGrid Is Nothing Then Set mGrid.Engine = Nothing
    Set mGrid = Nothing
End Sub

' ================================================================ entry

' Lays out the document for a viewport of width x height CSS px. The fragment tree
' (Root) and the fixed overlays (FixedCount/FixedFrag) are then ready for painting.
Public Sub Run(ByVal width As Double, ByVal height As Double)
    Dim root As clsRenderBox, i As Long
    mW = width
    If mW < 200 Then mW = 200
    mH = height
    If mH < 200 Then mH = 200
    modStyle.StyleBegin mW, mH
    ResetIntrinsic
    mFixedCount = 0
    Erase mFixed
    Set mSeenFixed = NewMap()
    Set mTop = modDom.Q("header.topbar")
    Set mLibrary = modDom.ById("library")
    Set mMain = modDom.Q("main")
    Set mWorkspace = modDom.Q(".workspace")
    Set mResizer = modDom.ById("sidebarResizer")
    Set mShell = modDom.Q(".app-shell")
    mMobile = False
    If Not mWorkspace Is Nothing Then mMobile = (modStyle.SV(modStyle.StyleOf(mWorkspace), P_DISPLAY) <> "grid")
    BeginIncremental
    Set root = ElementBox(modDom.DocumentElement)
    Set mRoot = LayoutBox(root, mW, mW, NOV, False, False, NOV)
    ' Fixed boxes found while walking (overlays) are laid out against the viewport.
    For i = 0 To mFixedCount - 1
        Set mFixed(i) = LayoutFixed(mFixed(i))
    Next i
End Sub

Public Property Get Root() As clsRenderBox
    Set Root = mRoot
End Property

Public Property Get FixedCount() As Long
    FixedCount = mFixedCount
End Property

Public Function FixedFrag(ByVal i As Long) As clsRenderBox
    Set FixedFrag = mFixed(i)
End Function

Public Property Get ViewWidth() As Double
    ViewWidth = mW
End Property

Public Property Get ViewHeight() As Double
    ViewHeight = mH
End Property

Public Property Get Mobile() As Boolean
    Mobile = mMobile
End Property

' ================================================================ small helpers

Private Function MaxD(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxD = a Else MaxD = b
End Function

Private Function MinD(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinD = a Else MinD = b
End Function

Friend Function ElementBox(ByVal el As clsNode) As clsRenderBox
    Dim b As New clsRenderBox
    b.Kind = 1
    Set b.Node = el
    b.Slot = modStyle.StyleOf(el)
    Set ElementBox = b
End Function

Private Function TextBox(ByVal owner As clsNode, ByVal slot As Long, ByVal text As String) As clsRenderBox
    Dim b As New clsRenderBox
    b.Kind = 3
    Set b.Node = owner
    b.Slot = slot
    b.Text = text
    Set TextBox = b
End Function

Friend Function RoleOf(ByVal el As clsNode) As Long
    If el Is Nothing Then Exit Function
    If mMobile Then
        ' Narrow windows scroll the whole page (the title bar scrolls away; the library
        ' button sticks to the top, see clsRenderPaint).
        If mShell Is Nothing Then
            If el Is mTop Then RoleOf = 1
            If el Is mWorkspace Then RoleOf = 4
        ElseIf el Is mShell Then
            RoleOf = 4
        End If
    ElseIf el Is mTop Then
        RoleOf = 1
    ElseIf el Is mLibrary Then
        RoleOf = 2
    ElseIf el Is mMain Then
        RoleOf = 3
    ElseIf el Is mResizer Then
        RoleOf = 5
    End If
End Function

Friend Function IsControlTag(ByVal el As clsNode) As Boolean
    Select Case el.LocalName
        Case "input", "select", "textarea": IsControlTag = True
    End Select
End Function

Friend Function IsControl(ByVal box As clsRenderBox) As Boolean
    If box.Kind = 1 Then IsControl = IsControlTag(box.Node)
End Function

Friend Function IsAbsolute(ByVal box As clsRenderBox) As Boolean
    If box.Kind = 3 Or box.Kind = 4 Then Exit Function
    IsAbsolute = (modStyle.SV(box.Slot, P_POSITION) = "absolute")
End Function

Friend Function DisplayOf(ByVal box As clsRenderBox, ByVal blockify As Boolean) As String
    Dim d As String, pos As String
    If box.Kind = 3 Then DisplayOf = "inline": Exit Function
    If box.Kind = 4 Then DisplayOf = "block": Exit Function
    d = modStyle.SV(box.Slot, P_DISPLAY)
    pos = modStyle.SV(box.Slot, P_POSITION)
    If blockify Or pos = "absolute" Or pos = "fixed" Then
        Select Case d
            Case "inline", "inline-block": d = "block"
            Case "inline-flex": d = "flex"
            Case "inline-grid": d = "grid"
            Case "inline-table": d = "table"
        End Select
    End If
    DisplayOf = d
End Function

Friend Function IsBlockLevel(ByVal d As String) As Boolean
    Select Case d
        Case "block", "flex", "grid", "list-item", "table", "flow-root": IsBlockLevel = True
    End Select
End Function

Private Function TextIsBlank(ByVal text As String) As Boolean
    Dim i As Long, c As String
    For i = 1 To Len(text)
        c = Mid$(text, i, 1)
        If c <> " " And c <> vbTab And c <> vbLf And c <> vbCr And c <> vbFormFeed Then Exit Function
    Next i
    TextIsBlank = True
End Function

' Edges (padding, border, margin) of a fragment. cbw = containing block width (NOV when
' unknown: percentages then resolve against 0).
Friend Sub SetEdges(ByVal f As clsRenderBox, ByVal s As Long, ByVal cbw As Double)
    Dim base As Double, a As Boolean, b As Boolean, c As Boolean, d As Boolean
    If f.Kind = 3 Or f.Kind = 4 Then Exit Sub
    If cbw = NOV Then base = 0 Else base = cbw
    f.PT = Pad(s, P_PT, base): f.PR = Pad(s, P_PR, base): f.PB = Pad(s, P_PB, base): f.PL = Pad(s, P_PL, base)
    f.BT = Bord(s, P_BTW, P_BTS): f.BR = Bord(s, P_BRW, P_BRS): f.BB = Bord(s, P_BBW, P_BBS): f.BL = Bord(s, P_BLW, P_BLS)
    ' (Object properties cannot be passed ByRef: the flags go through locals.)
    f.MT = Marg(s, P_MT, base, a): f.MR = Marg(s, P_MR, base, b)
    f.MB = Marg(s, P_MB, base, c): f.ML = Marg(s, P_ML, base, d)
    f.MTAuto = a: f.MRAuto = b: f.MBAuto = c: f.MLAuto = d
End Sub

Private Function Pad(ByVal s As Long, ByVal p As Long, ByVal base As Double) As Double
    Dim v As Double
    If modStyle.SLen(s, p, base, True, v) Then If v > 0 Then Pad = v
End Function

Private Function Bord(ByVal s As Long, ByVal pw As Long, ByVal ps As Long) As Double
    Dim v As Double, st As String
    st = modStyle.SV(s, ps)
    If st = "none" Or st = "hidden" Then Exit Function
    If modStyle.SLen(s, pw, 0, False, v) Then If v > 0 Then Bord = v
End Function

Private Function Marg(ByVal s As Long, ByVal p As Long, ByVal base As Double, ByRef isAuto As Boolean) As Double
    Dim v As Double
    isAuto = False
    If modStyle.SV(s, p) = "auto" Then isAuto = True: Exit Function
    If modStyle.SLen(s, p, base, True, v) Then Marg = v
End Function

' ================================================================ measuring

' Forgets measured widths (the window's fonts changed).
Public Sub ClearMeasureCache()
    mWidths.Clear
    mLineHeights.Clear
    ClearLayoutCache
End Sub

Private Function FontKey(ByVal s As Long) As String
    Dim stamp As Long
    stamp = modStyle.SLayoutStamp(s)
    If s > UBound(mFontKeys) Then
        ReDim Preserve mFontKeys(0 To s * 2 + 64)
        ReDim Preserve mFontKeyStamp(0 To s * 2 + 64)
    End If
    If mFontKeyStamp(s) <> stamp Or stamp = 0 Or Len(mFontKeys(s)) = 0 Then
        mFontKeys(s) = modStyle.SFont(s) & "|" & modStyle.SWeight(s) & "|" & modStyle.SItalic(s)
        mFontKeyStamp(s) = stamp
    End If
    FontKey = mFontKeys(s)
End Function

Friend Function Measure(ByVal text As String, ByVal s As Long) As Double
    Dim key As String, v As Variant
    key = FontKey(s) & "|" & text
    v = mWidths(key)
    If IsEmpty(v) Then
        v = modRender.MeasureText(text, modStyle.SFont(s), modStyle.SWeight(s), modStyle.SItalic(s))
        If mWidths.Count > 30000 Then mWidths.Clear
        mWidths(key) = v
    End If
    Measure = v
End Function

Friend Function SpaceWidth(ByVal s As Long) As Double
    SpaceWidth = Measure(" ", s)
End Function

Friend Function NaturalLineHeight(ByVal s As Long) As Double
    Dim key As String, v As Variant
    key = FontKey(s)
    v = mLineHeights(key)
    If IsEmpty(v) Then
        v = modRender.NaturalLineHeight(modStyle.SFont(s), modStyle.SWeight(s))
        mLineHeights(key) = v
    End If
    NaturalLineHeight = v
End Function

Friend Function CssLineHeight(ByVal s As Long) As Double
    Dim size As Double, t As String, v As Double
    t = modStyle.SV(s, P_LINEHEIGHT)
    size = modStyle.SFont(s)
    If t = "normal" Then
        CssLineHeight = NaturalLineHeight(s)
    ElseIf modStyle.SKind(s, P_LINEHEIGHT) = VK_NUM Then
        CssLineHeight = modStyle.SNum(s, P_LINEHEIGHT) * size
    ElseIf modStyle.SLen(s, P_LINEHEIGHT, size, True, v) Then
        CssLineHeight = v
    Else
        CssLineHeight = NaturalLineHeight(s)
    End If
End Function

' ================================================================ children

' The boxes an element's content generates, in order: ::before, child elements and text,
' ::after. Hidden boxes are skipped; position: fixed boxes are set aside as overlays;
' display: contents is flattened.
Friend Sub KidsOf(ByVal el As clsNode, ByVal slot As Long, ByRef kids() As clsRenderBox, ByRef n As Long)
    Dim c As clsNode, cs As Long, d As String, closedDetails As Boolean, b As clsRenderBox, key As String
    If IsControlTag(el) Then Exit Sub
    PushPseudo el, 1, kids, n
    closedDetails = (el.LocalName = "details" And Not el.HasAttribute("open"))
    Set c = el.FirstChild
    Do While Not c Is Nothing
        If c.NodeType = 3 Then
            If Not closedDetails And Len(c.Data) > 0 Then PushKid TextBox(el, slot, c.Data), kids, n
        ElseIf c.NodeType = 1 Then
            If Not (closedDetails And c.LocalName <> "summary") Then
                cs = modStyle.StyleOf(c)
                d = modStyle.SV(cs, P_DISPLAY)
                If d <> "none" And Not IsScreenReaderOnly(c, cs) Then
                    If modStyle.SV(cs, P_POSITION) = "fixed" Then
                        key = CStr(c.Uid)
                        If IsEmpty(mSeenFixed(key)) Then
                            mSeenFixed(key) = True
                            Set b = New clsRenderBox
                            b.Kind = 1: Set b.Node = c: b.Slot = cs
                            If mFixedCount = 0 Then
                                ReDim mFixed(0 To 3)
                            ElseIf mFixedCount > UBound(mFixed) Then
                                ReDim Preserve mFixed(0 To mFixedCount * 2)
                            End If
                            Set mFixed(mFixedCount) = b
                            mFixedCount = mFixedCount + 1
                        End If
                    ElseIf d = "contents" Then
                        KidsOf c, cs, kids, n
                    Else
                        Set b = New clsRenderBox
                        b.Kind = 1: Set b.Node = c: b.Slot = cs
                        PushKid b, kids, n
                    End If
                End If
            End If
        End If
        Set c = c.NextSibling
    Loop
    PushPseudo el, 2, kids, n
End Sub

Private Sub PushKid(ByVal b As clsRenderBox, ByRef kids() As clsRenderBox, ByRef n As Long)
    If n = 0 Then
        ReDim kids(0 To 7)
    ElseIf n > UBound(kids) Then
        ReDim Preserve kids(0 To n * 2)
    End If
    Set kids(n) = b
    n = n + 1
End Sub

Private Sub PushPseudo(ByVal el As clsNode, ByVal which As Long, ByRef kids() As clsRenderBox, ByRef n As Long)
    Dim ps As Long, text As String, ok As Boolean, b As clsRenderBox
    ps = modStyle.PseudoOf(el, which)
    If ps = 0 Then Exit Sub
    If modStyle.SV(ps, P_DISPLAY) = "none" Then Exit Sub
    text = PseudoText(modStyle.SV(ps, P_CONTENT), ok)
    If Not ok Then Exit Sub
    ' The collapse chevron is a rotated ">"; MSForms cannot rotate text, so an open card
    ' gets the down-pointing arrowhead glyph instead (a closed one keeps the ">").
    If text = ChrW$(&H203A) And InStr(1, modStyle.SV(ps, P_TRANSFORM), "rotate(90deg)") > 0 Then
        ' (Windows draws the arrowhead with the symbol font; a Mac uses a triangle Arial has.)
        If modRender.SymbolFallbackActive() Then text = ChrW$(&H2304) Else text = ChrW$(&H25BE)
    End If
    Set b = New clsRenderBox
    b.Kind = 2
    Set b.Node = el
    b.Slot = ps
    b.Which = which
    b.Text = text
    PushKid b, kids, n
End Sub

' The string of a content value such as "abc" or '\203a'.
Private Function PseudoText(ByVal content As String, ByRef ok As Boolean) As String
    Dim q As String, body As String, i As Long, c As String, out As String, hexText As String, j As Long
    content = Trim$(content)
    If Len(content) < 2 Then Exit Function
    q = Left$(content, 1)
    If (q <> """" And q <> "'") Or Right$(content, 1) <> q Then Exit Function
    body = Mid$(content, 2, Len(content) - 2)
    i = 1
    Do While i <= Len(body)
        c = Mid$(body, i, 1)
        If c = "\" And i < Len(body) Then
            j = i + 1
            hexText = ""
            Do While j <= Len(body) And Len(hexText) < 6
                If InStr(1, "0123456789abcdefABCDEF", Mid$(body, j, 1)) = 0 Then Exit Do
                hexText = hexText & Mid$(body, j, 1)
                j = j + 1
            Loop
            If Len(hexText) > 0 Then
                out = out & modDom.DomFromCodePoint(CLng("&H" & hexText))
                If j <= Len(body) Then If Mid$(body, j, 1) = " " Then j = j + 1
                i = j
            Else
                out = out & Mid$(body, i + 1, 1)
                i = i + 2
            End If
        Else
            out = out & c
            i = i + 1
        End If
    Loop
    ok = True
    PseudoText = out
End Function

Private Function IsScreenReaderOnly(ByVal el As clsNode, ByVal s As Long) As Boolean
    If el.HasClass("sr-only") Then IsScreenReaderOnly = True: Exit Function
    If modStyle.SV(s, P_POSITION) = "absolute" Then IsScreenReaderOnly = (Left$(Replace(modStyle.SV(s, P_CLIP), " ", ""), 6) = "rect(0")
End Function

' The element a click on inline content activates: a <summary>, or a <label> of a
' checkbox/radio (Nothing otherwise).
Friend Function LabelClick(ByVal el As clsNode) As clsNode
    Dim control As clsNode, t As String
    If el Is Nothing Then Exit Function
    If el.LocalName = "summary" Then Set LabelClick = el: Exit Function
    If el.LocalName <> "label" Then Exit Function
    Set control = el.LabelControl
    If control Is Nothing Then Exit Function
    t = control.TypeName
    If t = "checkbox" Or t = "radio" Then Set LabelClick = el
End Function

' ================================================================ inline content

Private Sub PushTok(ByRef toks() As TTok, ByRef tc As Long, ByVal t As Long, ByVal text As String, ByVal slot As Long, _
        ByVal click As clsNode, ByVal wrap As Boolean, ByVal anywhere As Boolean, ByVal collapsible As Boolean, ByVal box As clsRenderBox)
    If tc > UBound(toks) Then ReDim Preserve toks(0 To tc * 2 + 1)
    With toks(tc)
        .T = t
        .Text = text
        .Slot = slot
        Set .Click = click
        .Wrap = wrap
        .Anywhere = anywhere
        .Collapsible = collapsible
        Set .Box = box
        .W = 0
        Set .Frag = Nothing
    End With
    tc = tc + 1
End Sub

Private Function CollapseWhite(ByVal text As String, ByVal keepNewlines As Boolean) As String
    Dim i As Long, c As String, out As String, inSpace As Boolean, n As Long, buf() As String
    n = Len(text)
    If n = 0 Then Exit Function
    ReDim buf(1 To n)
    Dim k As Long
    For i = 1 To n
        c = Mid$(text, i, 1)
        If c = vbLf And keepNewlines Then
            ' pre-line: spaces around a newline disappear
            If k > 0 Then If buf(k) = " " Then k = k - 1
            k = k + 1: buf(k) = vbLf
            inSpace = True
        ElseIf c = " " Or c = vbTab Or c = vbLf Or c = vbCr Or c = vbFormFeed Then
            If Not inSpace Then k = k + 1: buf(k) = " ": inSpace = True
        Else
            k = k + 1: buf(k) = c
            inSpace = False
        End If
    Next i
    If k = 0 Then Exit Function
    ReDim Preserve buf(1 To k)
    out = Join(buf, "")
    CollapseWhite = out
End Function

Private Function TransformText(ByVal text As String, ByVal s As Long) As String
    Dim t As String, i As Long, prevWord As Boolean, c As String, isWord As Boolean
    t = modStyle.SV(s, P_TEXTTRANSFORM)
    Select Case t
        Case "uppercase": TransformText = UCase$(text)
        Case "lowercase": TransformText = LCase$(text)
        Case "capitalize"
            TransformText = text
            For i = 1 To Len(text)
                c = Mid$(text, i, 1)
                isWord = (c Like "[A-Za-z0-9_]")
                If isWord And Not prevWord Then Mid$(TransformText, i, 1) = UCase$(c)
                prevWord = isWord
            Next i
        Case Else: TransformText = text
    End Select
End Function

Private Sub AddText(ByVal text As String, ByVal s As Long, ByRef toks() As TTok, ByRef tc As Long, ByVal click As clsNode)
    Dim ws As String, collapse As Boolean, keepNewlines As Boolean, wrap As Boolean, anywhere As Boolean
    Dim segs() As String, parts() As String, i As Long, j As Long, ow As String, wb As String, lastT As Long
    Dim pos As Long, startPos As Long, seg As String
    ws = modStyle.SV(s, P_WHITESPACE)
    collapse = (ws = "normal" Or ws = "nowrap" Or ws = "pre-line")
    keepNewlines = (ws = "pre" Or ws = "pre-wrap" Or ws = "pre-line" Or ws = "break-spaces")
    wrap = (ws <> "nowrap" And ws <> "pre")
    ow = modStyle.SV(s, P_OVERFLOWWRAP)
    wb = modStyle.SV(s, P_WORDBREAK)
    anywhere = (ow = "anywhere" Or ow = "break-word" Or wb = "break-all")
    If ws = "normal" Or ws = "nowrap" Then
        text = CollapseWhite(text, False)
    ElseIf ws = "pre-line" Then
        text = CollapseWhite(Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf), True)
    Else
        text = Replace(Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf), vbTab, "    ")
    End If
    text = TransformText(text, s)
    If keepNewlines Then segs = Split(text, vbLf) Else ReDim segs(0 To 0): segs(0) = text
    For i = 0 To UBound(segs)
        If i > 0 Then PushTok toks, tc, 3, "", s, click, wrap, False, False, Nothing
        seg = segs(i)
        If collapse Then
            parts = Split(seg, " ")
            For j = 0 To UBound(parts)
                If j > 0 Then
                    lastT = 0
                    If tc > 0 Then
                        lastT = toks(tc - 1).T
                        If lastT = 2 And toks(tc - 1).Collapsible Then lastT = -1
                    End If
                    If tc > 0 And lastT <> 3 And lastT <> -1 Then PushTok toks, tc, 2, " ", s, click, wrap, False, True, Nothing
                End If
                If Len(parts(j)) > 0 Then PushTok toks, tc, 1, parts(j), s, click, wrap, anywhere, False, Nothing
            Next j
        Else
            ' Preserved spaces: each one is its own (non-collapsible) token.
            startPos = 1
            For pos = 1 To Len(seg)
                If Mid$(seg, pos, 1) = " " Then
                    If pos > startPos Then PushTok toks, tc, 1, Mid$(seg, startPos, pos - startPos), s, click, wrap, anywhere Or ws = "pre-wrap", False, Nothing
                    PushTok toks, tc, 2, " ", s, click, wrap, False, False, Nothing
                    startPos = pos + 1
                End If
            Next pos
            If Len(seg) >= startPos Then PushTok toks, tc, 1, Mid$(seg, startPos), s, click, wrap, anywhere Or ws = "pre-wrap", False, Nothing
        End If
    Next i
End Sub

Private Sub FlattenInline(ByRef kids() As clsRenderBox, ByVal n As Long, ByRef toks() As TTok, ByRef tc As Long, ByVal click As clsNode)
    Dim i As Long, k As clsRenderBox, d As String, sub_() As clsRenderBox, sn As Long, c2 As clsNode
    For i = 0 To n - 1
        Set k = kids(i)
        If k.Kind = 3 Then
            AddText k.Text, k.Slot, toks, tc, click
        ElseIf IsAbsolute(k) Then
            PushTok toks, tc, 4, "", k.Slot, Nothing, True, False, False, k
        Else
            d = DisplayOf(k, False)
            If k.Kind = 2 Then
                If d = "inline" Then AddText k.Text, k.Slot, toks, tc, click Else PushTok toks, tc, 5, "", k.Slot, click, True, False, False, k
            ElseIf k.Node.LocalName = "br" Then
                PushTok toks, tc, 3, "", k.Slot, click, True, False, False, Nothing
            ElseIf d = "inline" And Not IsControl(k) And k.Node.LocalName <> "button" Then
                sn = 0
                KidsOf k.Node, k.Slot, sub_, sn
                Set c2 = LabelClick(k.Node)
                If c2 Is Nothing Then Set c2 = click
                If sn > 0 Then FlattenInline sub_, sn, toks, tc, c2
            Else
                PushTok toks, tc, 5, "", k.Slot, click, True, False, False, k
            End If
        End If
    Next i
End Sub

' Breaks inline content into lines of width cw. Text runs and atomic inlines are added to
' target (runs with line/run indexes; atoms as child fragments); absolutely positioned
' boxes go to absTarget. Returns the content height.
Friend Function LayoutInline(ByRef kids() As clsRenderBox, ByVal n As Long, ByVal blockSlot As Long, ByVal cw As Double, _
        ByVal owner As clsRenderBox, ByVal target As clsRenderBox, ByVal absTarget As clsRenderBox) As Double
    Dim toks() As TTok, tc As Long, origCount As Long, click As clsNode, k As Long
    Dim items() As Long, ic As Long, curStart As Long, curW As Double, curContent As Long
    Dim lStart() As Long, lEnd() As Long, lW() As Double, lBr() As Long, lc As Long, brSlot As Long
    Dim w As Double, frag As clsRenderBox, opportunity As Boolean, lastIdx As Long
    Dim chunk As String, ch As String, chunkW As Double, chW As Double, ci As Long
    ReDim toks(0 To 15)
    ReDim items(0 To 15)
    ReDim lStart(0 To 7): ReDim lEnd(0 To 7): ReDim lW(0 To 7): ReDim lBr(0 To 7)
    If owner.Kind = 1 Then Set click = LabelClick(owner.Node)
    FlattenInline kids, n, toks, tc, click
    origCount = tc
    For k = 0 To origCount - 1
        Select Case toks(k).T
            Case 4
                If Not absTarget Is Nothing Then absTarget.AddAbsolute toks(k).Box
            Case 3
                brSlot = toks(k).Slot
                FinishLine toks, items, ic, curStart, curW, curContent, lStart, lEnd, lW, lBr, lc, brSlot
            Case 2
                If Not (ic = curStart And toks(k).Collapsible) Then
                    w = SpaceWidth(toks(k).Slot)
                    toks(k).W = w
                    PushItem items, ic, k
                    curW = curW + w
                End If
            Case Else
                If toks(k).T = 1 Then
                    w = Measure(toks(k).Text, toks(k).Slot)
                Else
                    Set frag = LayoutBox(toks(k).Box, cw, NOV, NOV, True, False, NOV)
                    Set toks(k).Frag = frag
                    w = frag.W + frag.ML + frag.MR
                End If
                toks(k).W = w
                opportunity = False
                If ic > curStart Then
                    lastIdx = items(ic - 1)
                    opportunity = (toks(lastIdx).T = 2 Or toks(lastIdx).T = 5 Or toks(k).T = 5)
                End If
                If curContent > 0 And opportunity And toks(k).Wrap And curW + w > cw + 0.01 Then
                    FinishLine toks, items, ic, curStart, curW, curContent, lStart, lEnd, lW, lBr, lc, 0
                End If
                If toks(k).T = 1 And w > cw And toks(k).Anywhere And Len(toks(k).Text) > 1 Then
                    ' Break a long word wherever it reaches the edge.
                    chunk = ""
                    chunkW = 0
                    For ci = 1 To Len(toks(k).Text)
                        ch = Mid$(toks(k).Text, ci, 1)
                        chW = Measure(ch, toks(k).Slot)
                        If Len(chunk) > 0 And curW + chunkW + chW > cw + 0.01 Then
                            AddChunk toks, tc, k, chunk, chunkW, items, ic
                            curW = curW + chunkW
                            curContent = curContent + 1
                            FinishLine toks, items, ic, curStart, curW, curContent, lStart, lEnd, lW, lBr, lc, 0
                            chunk = ch
                            chunkW = chW
                        Else
                            chunk = chunk & ch
                            chunkW = chunkW + chW
                        End If
                    Next ci
                    If Len(chunk) > 0 Then
                        AddChunk toks, tc, k, chunk, chunkW, items, ic
                        curW = curW + chunkW
                        curContent = curContent + 1
                    End If
                Else
                    PushItem items, ic, k
                    curW = curW + w
                    curContent = curContent + 1
                End If
        End Select
    Next k
    If ic > curStart Or lc = 0 Then FinishLine toks, items, ic, curStart, curW, curContent, lStart, lEnd, lW, lBr, lc, 0
    LayoutInline = PlaceLines(toks, items, lStart, lEnd, lW, lBr, lc, blockSlot, cw, target)
End Function

Private Sub PushItem(ByRef items() As Long, ByRef ic As Long, ByVal k As Long)
    If ic > UBound(items) Then ReDim Preserve items(0 To ic * 2 + 1)
    items(ic) = k
    ic = ic + 1
End Sub

Private Sub AddChunk(ByRef toks() As TTok, ByRef tc As Long, ByVal k As Long, ByVal chunk As String, ByVal w As Double, ByRef items() As Long, ByRef ic As Long)
    PushTok toks, tc, 1, chunk, toks(k).Slot, toks(k).Click, toks(k).Wrap, toks(k).Anywhere, False, Nothing
    toks(tc - 1).W = w
    PushItem items, ic, tc - 1
End Sub

Private Sub FinishLine(ByRef toks() As TTok, ByRef items() As Long, ByRef ic As Long, ByRef curStart As Long, ByRef curW As Double, _
        ByRef curContent As Long, ByRef lStart() As Long, ByRef lEnd() As Long, ByRef lW() As Double, ByRef lBr() As Long, ByRef lc As Long, ByVal brSlot As Long)
    Do While ic > curStart
        If toks(items(ic - 1)).T <> 2 Or Not toks(items(ic - 1)).Collapsible Then Exit Do
        curW = curW - toks(items(ic - 1)).W
        ic = ic - 1
    Loop
    If lc > UBound(lStart) Then
        ReDim Preserve lStart(0 To lc * 2): ReDim Preserve lEnd(0 To lc * 2)
        ReDim Preserve lW(0 To lc * 2): ReDim Preserve lBr(0 To lc * 2)
    End If
    lStart(lc) = curStart
    lEnd(lc) = ic
    lW(lc) = curW
    lBr(lc) = brSlot
    lc = lc + 1
    curStart = ic
    curW = 0
    curContent = 0
End Sub

Private Function PlaceLines(ByRef toks() As TTok, ByRef items() As Long, ByRef lStart() As Long, ByRef lEnd() As Long, ByRef lW() As Double, _
        ByRef lBr() As Long, ByVal lc As Long, ByVal blockSlot As Long, ByVal cw As Double, ByVal target As clsRenderBox) As Double
    Dim align As String, strut As Double, y As Double, li As Long, lineH As Double, i As Long, k As Long
    Dim x As Double, outer As Double, f As clsRenderBox, natural As Double, runIndex As Long
    Dim runOpen As Boolean, runX As Double, runW As Double, runText As String, runSlot As Long, runClick As clsNode, runKey As String
    Dim key As String, prevSlot As Long, prevClick As clsNode, t As String, lastLine As Long
    align = modStyle.SV(blockSlot, P_TEXTALIGN)
    strut = CssLineHeight(blockSlot)
    lastLine = lc - 1
    For li = 0 To lc - 1
        ' An empty trailing line only exists after a forced break.
        If li = lastLine And lStart(li) = lEnd(li) And lc > 1 Then
            If lBr(li - 1) = 0 Then Exit For
        End If
        If lEnd(li) > lStart(li) Then
            lineH = strut
        ElseIf lBr(li) > 0 Then
            lineH = CssLineHeight(lBr(li))
        Else
            lineH = strut
        End If
        For i = lStart(li) To lEnd(li) - 1
            k = items(i)
            If toks(k).T = 5 Then
                Set f = toks(k).Frag
                lineH = MaxD(lineH, f.H + f.MT + f.MB)
            Else
                lineH = MaxD(lineH, CssLineHeight(toks(k).Slot))
            End If
        Next i
        Select Case align
            Case "center": x = (cw - lW(li)) / 2
            Case "right", "end": x = cw - lW(li)
            Case Else: x = 0
        End Select
        If x < 0 Then x = 0
        runOpen = False
        runIndex = 0
        prevSlot = -1
        For i = lStart(li) To lEnd(li) - 1
            k = items(i)
            If toks(k).T = 5 Then
                If runOpen Then CloseRun target, runX, y, lineH, runW, runText, runSlot, runClick, li, runIndex: runOpen = False
                Set f = toks(k).Frag
                outer = f.H + f.MT + f.MB
                f.X = x + f.ML
                If outer >= lineH Then f.Y = y + f.MT Else f.Y = y + (lineH - outer) / 2 + f.MT
                Set f.ClickNode = toks(k).Click
                target.AddKid f
            Else
                If toks(k).T = 2 Then t = " " Else t = toks(k).Text
                If runOpen And toks(k).Slot = prevSlot And toks(k).Click Is prevClick Then
                    runText = runText & t
                    runW = runW + toks(k).W
                Else
                    key = FontKey(toks(k).Slot) & "|" & modStyle.SVI(toks(k).Slot, P_COLOR) & "|" & modStyle.SVI(toks(k).Slot, P_TEXTTRANSFORM) & "|" & ObjKey(toks(k).Click)
                    If runOpen And key = runKey Then
                        runText = runText & t
                        runW = runW + toks(k).W
                    Else
                        If runOpen Then CloseRun target, runX, y, lineH, runW, runText, runSlot, runClick, li, runIndex
                        runOpen = True
                        runKey = key
                        runX = x
                        runW = toks(k).W
                        runText = t
                        runSlot = toks(k).Slot
                        Set runClick = toks(k).Click
                    End If
                    prevSlot = toks(k).Slot
                    Set prevClick = toks(k).Click
                End If
            End If
            x = x + toks(k).W
        Next i
        If runOpen Then CloseRun target, runX, y, lineH, runW, runText, runSlot, runClick, li, runIndex
        y = y + lineH
    Next li
    PlaceLines = y
End Function

Private Function ObjKey(ByVal node As clsNode) As String
    If node Is Nothing Then ObjKey = "-" Else ObjKey = CStr(node.Uid)
End Function

Private Sub CloseRun(ByVal target As clsRenderBox, ByVal runX As Double, ByVal lineY As Double, ByVal lineH As Double, ByVal runW As Double, _
        ByVal runText As String, ByVal runSlot As Long, ByVal runClick As clsNode, ByVal lineIndex As Long, ByRef runIndex As Long)
    Dim natural As Double, n As Long
    n = Len(runText)
    Do While n > 0
        If Mid$(runText, n, 1) <> " " Then Exit Do
        n = n - 1
    Loop
    If n = 0 Then Exit Sub
    runText = Left$(runText, n)
    natural = NaturalLineHeight(runSlot)
    target.AddRun runX, lineY + (lineH - natural) / 2, runW, natural, runText, runSlot, runClick, lineIndex, runIndex
    runIndex = runIndex + 1
End Sub

' ================================================================ intrinsic widths

Private Sub ResetIntrinsic()
    Set mIntrMap = NewMap()
    mIntrCount = 0
    ReDim mIntrMin(0 To 255)
    ReDim mIntrMax(0 To 255)
End Sub

' Called when a size container's width changed during the pass: cached sizes may depend
' on styles that just changed.
Friend Sub ForgetIntrinsic()
    ResetIntrinsic
End Sub

Private Sub InlineIntrinsic(ByRef kids() As clsRenderBox, ByVal n As Long, ByRef mn As Double, ByRef mx As Double)
    Dim toks() As TTok, tc As Long, k As Long, lineW As Double, w As Double, imn As Double, imx As Double
    Dim e As New clsRenderBox, hasNoWrap As Boolean, run As Double
    ReDim toks(0 To 15)
    FlattenInline kids, n, toks, tc, Nothing
    mn = 0: mx = 0
    For k = 0 To tc - 1
        Select Case toks(k).T
            Case 3
                mx = MaxD(mx, lineW)
                lineW = 0
            Case 2
                toks(k).W = SpaceWidth(toks(k).Slot)
                lineW = lineW + toks(k).W
            Case 1
                w = Measure(toks(k).Text, toks(k).Slot)
                toks(k).W = w
                If toks(k).Wrap Then mn = MaxD(mn, w) Else hasNoWrap = True
                lineW = lineW + w
            Case 5
                Intrinsic toks(k).Box, imn, imx
                e.Kind = 1
                SetEdges e, toks(k).Box.Slot, NOV
                If toks(k).Box.Kind = 3 Then e.ML = 0: e.MR = 0
                w = imx + e.ML + e.MR
                mn = MaxD(mn, imn + e.ML + e.MR)
                lineW = lineW + w
        End Select
    Next k
    mx = MaxD(mx, lineW)
    If hasNoWrap Then
        For k = 0 To tc - 1
            If (toks(k).T = 1 Or toks(k).T = 2) And Not toks(k).Wrap Then
                run = run + toks(k).W
                mn = MaxD(mn, run)
            Else
                run = 0
            End If
        Next k
    End If
End Sub

Private Function ControlIntrinsic(ByVal box As clsRenderBox) As Double
    Dim el As clsNode, s As Long, opts As clsList, i As Long, widest As Double, e As New clsRenderBox, t As String
    Set el = box.Node
    s = box.Slot
    Select Case el.LocalName
        Case "select"
            Set opts = el.Options
            For i = 0 To opts.Count - 1
                t = Trim$(CollapseWhite(opts(i).TextContent, False))
                widest = MaxD(widest, Measure(t, s))
            Next i
            e.Kind = 1
            SetEdges e, s, NOV
            ControlIntrinsic = widest + MaxD(0, 22 - e.PR)
        Case "textarea"
            ControlIntrinsic = modStyle.SFont(s) * 13
        Case Else
            t = el.TypeName
            If t = "checkbox" Or t = "radio" Then ControlIntrinsic = 13 Else ControlIntrinsic = modStyle.SFont(s) * 11.5
    End Select
End Function

' Minimum and maximum content widths of a box including its padding and border.
Friend Sub Intrinsic(ByVal box As clsRenderBox, ByRef mn As Double, ByRef mx As Double)
    Dim key As String, v As Variant, s As Long, e As New clsRenderBox, padX As Double, borderBox As Boolean
    Dim specW As Double, imn As Double, imx As Double, minW As Double, maxW As Double, adjust As Double, idx As Long, ck As Long
    ck = CacheKey(box)
    If ck > 0 Then
        If mIHas(ck) Then mn = mIMn(ck): mx = mIMx(ck): Exit Sub
    ElseIf box.Kind = 1 Then
        key = CStr(box.Node.Uid)
    ElseIf box.Kind = 2 Then
        key = box.Node.Uid & ":" & box.Which
    End If
    If Len(key) > 0 Then
        v = mIntrMap(key)
        If Not IsEmpty(v) Then
            mn = mIntrMin(v): mx = mIntrMax(v)
            Exit Sub
        End If
    End If
    s = box.Slot
    e.Kind = box.Kind
    SetEdges e, s, NOV
    padX = e.PL + e.PR + e.BL + e.BR
    borderBox = (modStyle.SV(s, P_BOXSIZING) = "border-box")
    If box.Kind <> 3 And modStyle.SLen(s, P_WIDTH, 0, False, specW) Then
        If borderBox Then mn = MaxD(specW, padX) Else mn = specW + padX
        mx = mn
    Else
        If box.Kind = 2 Then
            imn = TextMin(box.Text, s)
            imx = Measure(box.Text, s)
        ElseIf IsControl(box) Then
            imx = ControlIntrinsic(box)
            ' Form controls are compressible: a percentage width contributes no minimum.
            If modStyle.SKind(s, P_WIDTH) = VK_PCT Then imn = 0 Else imn = imx
        Else
            ContentIntrinsic box, imn, imx
        End If
        mn = imn + padX
        mx = imx + padX
        If Not borderBox Then adjust = padX
        If Not modStyle.SIsRelative(s, P_MAXW) Then
            If modStyle.SLen(s, P_MAXW, 0, False, maxW) Then
                mn = MinD(mn, maxW + adjust)
                mx = MinD(mx, maxW + adjust)
            End If
        End If
        If Not modStyle.SIsRelative(s, P_MINW) And modStyle.SV(s, P_MINW) <> "auto" Then
            If modStyle.SLen(s, P_MINW, 0, False, minW) Then
                mn = MaxD(mn, minW + adjust)
                mx = MaxD(mx, minW + adjust)
            End If
        End If
    End If
    If mx < mn Then mx = mn
    If ck > 0 And mCacheOn Then
        mIHas(ck) = True: mIMn(ck) = mn: mIMx(ck) = mx
    ElseIf Len(key) > 0 Then
        If mIntrCount > UBound(mIntrMin) Then ReDim Preserve mIntrMin(0 To mIntrCount * 2): ReDim Preserve mIntrMax(0 To mIntrCount * 2)
        mIntrMin(mIntrCount) = mn
        mIntrMax(mIntrCount) = mx
        mIntrMap(key) = mIntrCount
        mIntrCount = mIntrCount + 1
    End If
End Sub

Friend Function TextMin(ByVal text As String, ByVal s As Long) As Double
    Dim parts() As String, i As Long
    parts = Split(CollapseWhite(text, False), " ")
    For i = 0 To UBound(parts)
        If Len(parts(i)) > 0 Then TextMin = MaxD(TextMin, Measure(parts(i), s))
    Next i
End Function

' Outer (margin-box) intrinsic sizes of a child box.
Friend Sub OuterIntrinsic(ByVal kid As clsRenderBox, ByRef mn As Double, ByRef mx As Double)
    Dim e As New clsRenderBox
    If kid.Kind = 3 Then
        mn = TextMin(kid.Text, kid.Slot)
        mx = Measure(Trim$(CollapseWhite(kid.Text, False)), kid.Slot)
        Exit Sub
    End If
    Intrinsic kid, mn, mx
    e.Kind = kid.Kind
    SetEdges e, kid.Slot, NOV
    mn = mn + e.ML + e.MR
    mx = mx + e.ML + e.MR
End Sub

' Children that take part in flex/grid layout (no blank text, no absolute boxes).
Friend Sub FlowKids(ByRef kids() As clsRenderBox, ByVal n As Long, ByRef out() As clsRenderBox, ByRef m As Long)
    Dim i As Long, k As clsRenderBox
    m = 0
    For i = 0 To n - 1
        Set k = kids(i)
        If k.Kind = 3 Then
            If Not (TextIsBlank(k.Text) And InStr(1, modStyle.SV(k.Slot, P_WHITESPACE), "pre") = 0) Then PushKid k, out, m
        ElseIf Not IsAbsolute(k) Then
            PushKid k, out, m
        End If
    Next i
End Sub

Private Sub ContentIntrinsic(ByVal box As clsRenderBox, ByRef mn As Double, ByRef mx As Double)
    Dim d As String, kids() As clsRenderBox, n As Long, fk() As clsRenderBox, m As Long, i As Long, s As Long
    Dim row As Boolean, gap As Double, imn As Double, imx As Double, sumMax As Double, sumMin As Double
    Dim run() As clsRenderBox, rn As Long, wrap As Boolean, hasContent As Boolean
    s = box.Slot
    d = DisplayOf(box, False)
    KidsOf box.Node, s, kids, n
    mn = 0: mx = 0
    If d = "flex" Or d = "inline-flex" Then
        row = (Left$(modStyle.SV(s, P_FLEXDIR), 3) = "row")
        FlowKids kids, n, fk, m
        If m = 0 Then Exit Sub
        If row Then
            If Not modStyle.SLen(s, P_COLGAP, 0, False, gap) Then gap = 0
        Else
            If Not modStyle.SLen(s, P_ROWGAP, 0, False, gap) Then gap = 0
        End If
        wrap = (modStyle.SV(s, P_FLEXWRAP) <> "nowrap")
        For i = 0 To m - 1
            OuterIntrinsic fk(i), imn, imx
            If row Then
                sumMax = sumMax + imx
                sumMin = sumMin + imn
                If wrap Then mn = MaxD(mn, imn)
            Else
                mn = MaxD(mn, imn)
                mx = MaxD(mx, imx)
            End If
        Next i
        If row Then
            mx = sumMax + gap * (m - 1)
            If Not wrap Then mn = sumMin + gap * (m - 1)
        End If
        Exit Sub
    End If
    If d = "grid" Or d = "inline-grid" Then
        FlowKids kids, n, fk, m
        mGrid.GridIntrinsic box, fk, m, mn, mx
        Exit Sub
    End If
    If d = "table" Or d = "inline-table" Then
        mGrid.TableIntrinsic box, mn, mx
        Exit Sub
    End If
    ' Block container: inline runs and blocks.
    For i = 0 To n - 1
        If IsBlockLevel(DisplayOf(kids(i), False)) And Not IsControl(kids(i)) Then
            If rn > 0 Then
                FlushIntrinsicRun run, rn, mn, mx
                rn = 0
            End If
            OuterIntrinsic kids(i), imn, imx
            mn = MaxD(mn, imn)
            mx = MaxD(mx, imx)
        ElseIf Not IsAbsolute(kids(i)) Then
            PushKid kids(i), run, rn
        End If
    Next i
    If rn > 0 Then FlushIntrinsicRun run, rn, mn, mx
End Sub

Private Sub FlushIntrinsicRun(ByRef run() As clsRenderBox, ByVal rn As Long, ByRef mn As Double, ByRef mx As Double)
    Dim i As Long, anyFound As Boolean, imn As Double, imx As Double
    For i = 0 To rn - 1
        If run(i).Kind <> 3 Then
            anyFound = True
        ElseIf Not TextIsBlank(run(i).Text) Then
            anyFound = True
        End If
        If anyFound Then Exit For
    Next i
    If Not anyFound Then Exit Sub
    InlineIntrinsic run, rn, imn, imx
    mn = MaxD(mn, imn)
    mx = MaxD(mx, imx)
End Sub

' ================================================================ boxes

' Lays out one box. cbw: containing block width; optW/optH: imposed border-box size (NOV
' for none); shrink: shrink-to-fit; blockify: flex/grid item; cbh: containing block
' height for percentages (NOV when indefinite). The fragment's x/y are set by the caller.
Friend Function LayoutBox(ByVal box As clsRenderBox, ByVal cbw As Double, ByVal optW As Double, ByVal optH As Double, _
        ByVal shrink As Boolean, ByVal blockify As Boolean, ByVal cbh As Double) As clsRenderBox
    Dim k As Long, e As Long, flags As Long, f As clsRenderBox, fixedBefore As Long
    k = CacheKey(box)
    If k > 0 Then
        If shrink Then flags = 1
        If blockify Then flags = flags + 2
        e = mEHead(k)
        Do While e > 0
            If mECbw(e) = cbw And mEOptW(e) = optW And mEOptH(e) = optH And mECbh(e) = cbh And mEFlags(e) = flags Then
                Set f = mEFrag(e)
                RestoreEntry e, f
                CacheHits = CacheHits + 1
                Set LayoutBox = f
                Exit Function
            End If
            e = mENext(e)
        Loop
        CacheMisses = CacheMisses + 1
        fixedBefore = mFixedCount
    End If
    Set f = LayoutBoxNow(box, cbw, optW, optH, shrink, blockify, cbh)
    If k > 0 And mCacheOn Then StoreEntry k, f, cbw, optW, optH, cbh, flags, fixedBefore
    Set LayoutBox = f
End Function

Private Function LayoutBoxNow(ByVal box As clsRenderBox, ByVal cbw As Double, ByVal optW As Double, ByVal optH As Double, _
        ByVal shrink As Boolean, ByVal blockify As Boolean, ByVal cbh As Double) As clsRenderBox
    Dim f As clsRenderBox, s As Long, d As String, padX As Double, padY As Double, borderBox As Boolean
    Dim adjustX As Double, adjustY As Double, control As Boolean, role As Long, el As clsNode
    Dim specW As Double, hasSpecW As Boolean, minW As Double, maxW As Double, hasMaxW As Boolean, w As Double, v As Double
    Dim available As Double, mn As Double, mx As Double, reserve As Double, cw As Double
    Dim specH As Double, hasSpecH As Boolean, minH As Double, maxH As Double, hasMaxH As Boolean
    Dim fixedH As Double, hasFixedH As Boolean, innerDef As Double, innerMin As Double, contentH As Double, h As Double
    Dim overflowY As String, overflowX As String, scrolls As Boolean, rightEdge As Double, i As Long, j As Long
    Dim k As clsRenderBox, a As clsRenderBox, scrollW As Double, hasCbw As Boolean, hasCbh As Boolean
    s = box.Slot
    Set f = box.NewFragment()
    d = DisplayOf(box, blockify)
    f.Display = d
    SetEdges f, s, cbw
    padX = f.PL + f.PR + f.BL + f.BR
    padY = f.PT + f.PB + f.BT + f.BB
    borderBox = (modStyle.SV(s, P_BOXSIZING) = "border-box")
    If Not borderBox Then adjustX = padX: adjustY = padY
    control = IsControl(box)
    If box.Kind = 1 Then
        Set el = box.Node
        role = RoleOf(el)
    End If
    f.Role = role
    hasCbw = (cbw <> NOV)
    hasCbh = (cbh <> NOV)
    hasSpecW = modStyle.SLen(s, P_WIDTH, cbw, hasCbw, specW)
    If hasSpecW Then specW = specW + adjustX
    If modStyle.SV(s, P_MINW) <> "auto" Then
        If modStyle.SLen(s, P_MINW, cbw, hasCbw, v) Then minW = v + adjustX
    End If
    hasMaxW = modStyle.SLen(s, P_MAXW, cbw, hasCbw, maxW)
    If hasMaxW Then maxW = maxW + adjustX
    If hasCbw Then available = cbw - f.ML - f.MR Else available = mW - f.ML - f.MR
    If available < 0 Then available = 0
    If optW <> NOV Then
        w = optW
    ElseIf hasSpecW Then
        w = specW
    ElseIf shrink Or Not IsBlockLevel(d) Or control Or d = "table" Then
        Intrinsic box, mn, mx
        w = MinD(MaxD(mn, available), mx)
    Else
        w = available
    End If
    ' min/max-width also bound imposed (stretched, flexed) widths.
    If hasMaxW Then w = MinD(w, maxW)
    w = MaxD(w, minW)
    w = MaxD(w, padX)
    If role = 3 Or role = 4 Then reserve = SCROLLBAR
    cw = MaxD(0, w - padX - reserve)
    hasSpecH = modStyle.SLen(s, P_HEIGHT, cbh, hasCbh, specH)
    If hasSpecH Then specH = specH + adjustY
    If modStyle.SV(s, P_MINH) <> "auto" Then
        If modStyle.SLen(s, P_MINH, cbh, hasCbh, v) Then minH = v + adjustY
    End If
    hasMaxH = modStyle.SLen(s, P_MAXH, cbh, hasCbh, maxH)
    If hasMaxH Then maxH = maxH + adjustY
    ' Textareas that auto-grow on the website size to their content between min and max.
    If Not el Is Nothing Then
        If el.LocalName = "textarea" And hasMaxH Then hasSpecH = False
    End If
    If optH <> NOV Then
        fixedH = optH
        hasFixedH = True
    ElseIf hasSpecH Then
        fixedH = MaxD(specH, minH)
        If hasMaxH Then fixedH = MinD(fixedH, maxH)
        hasFixedH = True
    End If
    If hasFixedH Then innerDef = MaxD(0, fixedH - padY) Else innerDef = NOV
    innerMin = MaxD(0, minH - padY)
    f.W = w
    f.CW = cw
    If Not el Is Nothing Then
        If modStyle.SV(s, P_CONTAINER) <> "normal" Then SetContainer el, cw
    End If
    contentH = LayoutContent(f, cw, innerDef, innerMin)
    If Not el Is Nothing Then If el.LocalName = "fieldset" Then contentH = PlaceLegend(f, contentH)
    If hasFixedH Then
        h = fixedH
    Else
        h = MaxD(contentH + padY, minH)
        If hasMaxH Then h = MinD(h, maxH)
    End If
    overflowY = modStyle.SV(s, P_OVERFLOWY)
    scrolls = (overflowY = "auto" Or overflowY = "scroll") And Not control
    If scrolls And contentH + padY > h + 0.5 And reserve = 0 Then
        reserve = SCROLLBAR
        ' The library's scrollbar may cover its right padding: browsers with overlay
        ' scrollbars (Windows 11, macOS) give its content the full width.
        If role = 2 Then reserve = MaxD(0, SCROLLBAR - f.PR)
        cw = MaxD(0, w - padX - reserve)
        f.CW = cw
        f.ClearKids
        contentH = LayoutContent(f, cw, innerDef, innerMin)
        If Not el Is Nothing Then If el.LocalName = "fieldset" Then contentH = PlaceLegend(f, contentH)
    End If
    If (scrolls Or role = 2 Or role = 3 Or role = 4) And contentH + padY > h + 0.5 Then f.Scrolls = True
    f.ScrollH = contentH + f.PT + f.PB
    ' Content wider than an overflow-x box (lp constraint rows) scrolls sideways.
    overflowX = modStyle.SV(s, P_OVERFLOWX)
    If Not control And (overflowX = "auto" Or overflowX = "scroll" Or overflowX = "hidden") Then
        rightEdge = 0
        For i = 0 To f.KidCount - 1
            Set k = f.Kid(i)
            If k.Kind = 4 Then
                For j = 0 To k.KidCount - 1
                    Set a = k.Kid(j)
                    rightEdge = MaxD(rightEdge, k.X + a.X + a.W + a.MR)
                Next j
            Else
                rightEdge = MaxD(rightEdge, k.X + k.W + k.MR)
            End If
        Next i
        scrollW = rightEdge - f.BL + f.PR
        If scrollW > w - f.BL - f.BR + 0.5 Then
            f.ScrollsX = True
            f.ScrollW = scrollW
            If overflowX <> "hidden" And Not hasFixedH Then h = h + SCROLLBAR
        End If
    End If
    f.H = h
    f.ContentH = contentH
    For i = 0 To f.AbsCount - 1
        PlaceAbsolute f.AbsBox(i), f
    Next i
    Set LayoutBoxNow = f
End Function

' A fieldset's first legend sits on the top border (centred on it; the border is drawn
' through its middle and left out behind it) instead of inside the padding. Returns the
' content height without the legend's in-flow share.
Private Function PlaceLegend(ByVal f As clsRenderBox, ByVal contentH As Double) As Double
    Dim lg As clsRenderBox, L As Double, up As Double, i As Long, mn As Double, mx As Double
    PlaceLegend = contentH
    f.HasLegend = False
    If f.KidCount = 0 Then Exit Function
    Set lg = f.Kid(0)
    If lg.Kind <> 1 Then Exit Function
    If lg.Node.LocalName <> "legend" Then Exit Function
    L = lg.H + lg.MT + lg.MB
    up = MinD(f.BT, L)
    For i = 1 To f.KidCount - 1
        f.Kid(i).Y = f.Kid(i).Y - up
    Next i
    lg.Y = MaxD(0, (f.BT - L) / 2) + lg.MT
    Intrinsic lg.Source, mn, mx
    f.HasLegend = True
    f.LegendLineY = MaxD(0, (L - f.BT) / 2)
    f.LegendX0 = lg.X
    f.LegendX1 = lg.X + MinD(mx, lg.W)
    PlaceLegend = contentH - up
End Function

Private Sub SetContainer(ByVal el As clsNode, ByVal cw As Double)
    If modStyle.SetContainerWidth(el, cw) Then
        ' Container queries changed mid-pass: styles walked before may change now.
        ResetIntrinsic
        ClearLayoutCache
        mCacheOn = False
    End If
End Sub

Private Function LayoutContent(ByVal f As clsRenderBox, ByVal cw As Double, ByVal innerDef As Double, ByVal innerMin As Double) As Double
    Dim ox As Double, oy As Double, kids() As clsRenderBox, n As Long, one(0 To 0) As clsRenderBox, i As Long, d As String
    ox = f.BL + f.PL
    oy = f.BT + f.PT
    If IsControl(f.Source) Then
        LayoutContent = ControlContentHeight(f.Source, cw)
        Exit Function
    End If
    If f.Kind = 2 Then
        If Len(f.Text) = 0 Then Exit Function
        Set one(0) = TextBox(f.Node, f.Slot, f.Text)
        LayoutContent = LayoutInline(one, 1, f.Slot, cw, f.Source, f, Nothing)
        f.LinesX = ox
        f.LinesY = oy
        Exit Function
    End If
    KidsOf f.Node, f.Slot, kids, n
    d = f.Display
    If d = "flex" Or d = "inline-flex" Then
        LayoutContent = LayoutFlex(f, kids, n, cw, innerDef, innerMin)
    ElseIf d = "grid" Or d = "inline-grid" Then
        LayoutContent = mGrid.LayoutGrid(f, kids, n, cw, innerDef, innerMin)
    ElseIf d = "table" Or d = "inline-table" Then
        LayoutContent = mGrid.LayoutTable(f, cw)
    Else
        LayoutContent = LayoutFlow(f, kids, n, cw)
    End If
    For i = 0 To f.KidCount - 1
        f.Kid(i).X = f.Kid(i).X + ox
        f.Kid(i).Y = f.Kid(i).Y + oy
    Next i
End Function

Private Function ControlContentHeight(ByVal box As clsRenderBox, ByVal cw As Double) As Double
    Dim el As clsNode, s As Long, natural As Double, lines As Long, t As String
    Set el = box.Node
    s = box.Slot
    natural = NaturalLineHeight(s)
    If el.LocalName = "textarea" Then
        lines = WrapCount(el.Value, cw - SCROLLBAR, s)
        If lines < 2 Then lines = 2
        ControlContentHeight = lines * natural + 2
        Exit Function
    End If
    t = el.TypeName
    If t = "checkbox" Or t = "radio" Then ControlContentHeight = 13 Else ControlContentHeight = natural
End Function

Private Function WrapCount(ByVal text As String, ByVal width As Double, ByVal s As Long) As Long
    Dim lines() As String, i As Long, words() As String, j As Long, cur As Double, count As Long, w As Double, sp As Double
    text = Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf)
    lines = Split(text, vbLf)
    sp = SpaceWidth(s)
    For i = 0 To UBound(lines)
        count = 1
        cur = 0
        words = Split(lines(i), " ")
        For j = 0 To UBound(words)
            If j > 0 Then cur = cur + sp
            If Len(words(j)) > 0 Then
                w = Measure(words(j), s)
                If cur + w > width And cur > 0 Then
                    count = count + 1
                    cur = w
                Else
                    cur = cur + w
                End If
            End If
        Next j
        WrapCount = WrapCount + count
    Next i
    If UBound(lines) < 0 Then WrapCount = 1
End Function

' ================================================================ block flow

Private Function LayoutFlow(ByVal f As clsRenderBox, ByRef kids() As clsRenderBox, ByVal n As Long, ByVal cw As Double) As Double
    Dim y As Double, pending As Double, first As Boolean, run() As clsRenderBox, rn As Long, groupIndex As Long
    Dim i As Long, k As clsRenderBox, d As String, child As clsRenderBox, gap As Double, x As Double, marker As clsRenderBox
    first = True
    Set marker = ListMarker(f)
    If Not marker Is Nothing Then PushKid marker, run, rn
    For i = 0 To n - 1
        Set k = kids(i)
        If IsAbsolute(k) Then
            f.AddAbsolute k
        Else
            d = DisplayOf(k, False)
            If Not IsBlockLevel(d) Or (IsControl(k) And d <> "block") Then
                PushKid k, run, rn
            Else
                If rn > 0 Then FlushFlowInline f, run, rn, cw, y, pending, first, groupIndex
                Set child = LayoutBox(k, cw, NOV, NOV, False, False, NOV)
                If first Then
                    gap = child.MT
                ElseIf pending >= 0 And child.MT >= 0 Then
                    gap = MaxD(pending, child.MT)
                Else
                    gap = pending + child.MT
                End If
                child.Y = y + gap
                x = child.ML
                If child.MLAuto And child.MRAuto Then
                    x = (cw - child.W) / 2
                ElseIf child.MLAuto Then
                    x = cw - child.W - child.MR
                End If
                child.X = x
                f.AddKid child
                y = child.Y + child.H
                pending = child.MB
                first = False
            End If
        End If
    Next i
    If rn > 0 Then FlushFlowInline f, run, rn, cw, y, pending, first, groupIndex
    ' A last margin stays inside only when padding or a border separates it.
    If f.PB > 0 Or f.BB > 0 Then y = y + pending
    LayoutFlow = y
End Function

Private Sub FlushFlowInline(ByVal f As clsRenderBox, ByRef run() As clsRenderBox, ByRef rn As Long, ByVal cw As Double, _
        ByRef y As Double, ByRef pending As Double, ByRef first As Boolean, ByRef groupIndex As Long)
    Dim i As Long, content As Boolean, anon As clsRenderBox, h As Double, k As clsRenderBox
    For i = 0 To rn - 1
        Set k = run(i)
        If k.Kind = 3 Then
            If Not (TextIsBlank(k.Text) And InStr(1, modStyle.SV(k.Slot, P_WHITESPACE), "pre") = 0) Then content = True
        Else
            content = True
        End If
        If content Then Exit For
    Next i
    If content Then
        Set anon = New clsRenderBox
        anon.Kind = 4
        Set anon.Node = f.Node
        anon.Slot = f.Slot
        anon.Which = f.Which
        Set anon.Owner = f
        anon.AnonIndex = CStr(groupIndex)
        groupIndex = groupIndex + 1
        y = y + pending
        pending = 0
        h = LayoutInline(run, rn, f.Slot, cw, f.Source, anon, f)
        anon.X = 0
        anon.Y = y
        anon.W = cw
        anon.H = h
        f.AddKid anon
        y = y + h
        first = False
    Else
        For i = 0 To rn - 1
            If IsAbsolute(run(i)) Then f.AddAbsolute run(i)
        Next i
    End If
    rn = 0
End Sub

Private Function ListMarker(ByVal f As clsRenderBox) As clsRenderBox
    Dim t As String, text As String, parent As clsNode, c As clsNode, index As Long
    If f.Display <> "list-item" Or f.Kind <> 1 Then Exit Function
    t = modStyle.SV(f.Slot, P_LISTSTYLE)
    If t = "none" Or Len(t) = 0 Then Exit Function
    Set parent = f.Node.ParentNode
    If t = "disclosure-closed" Or t = "disclosure-open" Then
        text = ChrW$(&H25B8) & " "
        If Not parent Is Nothing Then If parent.HasAttribute("open") Then text = ChrW$(&H25BE) & " "
    ElseIf t = "decimal" Then
        If Not parent Is Nothing Then
            Set c = parent.FirstElementChild
            Do While Not c Is Nothing
                If c.LocalName = "li" Then index = index + 1
                If c Is f.Node Then Exit Do
                Set c = c.NextElementSibling
            Loop
        End If
        text = index & ". "
    Else
        text = ChrW$(&H2022) & " "
    End If
    Set ListMarker = TextBox(f.Node, f.Slot, text)
End Function

' ================================================================ flex

Private Function AlignOf(ByVal box As clsRenderBox, ByVal alignItems As String) As String
    Dim self As String
    If Not box Is Nothing Then
        If box.Kind <> 3 Then self = modStyle.SV(box.Slot, P_ALIGNSELF)
    End If
    If Len(self) > 0 And self <> "auto" Then AlignOf = self Else AlignOf = alignItems
    If AlignOf = "normal" Then AlignOf = "stretch"
End Function

Friend Function LayoutAnonText(ByVal textBox As clsRenderBox, ByVal f As clsRenderBox, ByVal width As Double, ByVal indexLabel As String) As clsRenderBox
    Dim anon As New clsRenderBox, one(0 To 0) As clsRenderBox
    anon.Kind = 4
    Set anon.Node = f.Node
    anon.Slot = f.Slot
    anon.Which = f.Which
    Set anon.Owner = f
    anon.AnonIndex = indexLabel
    Set one(0) = textBox
    anon.H = LayoutInline(one, 1, f.Slot, width, f.Source, anon, Nothing)
    anon.W = width
    Set LayoutAnonText = anon
End Function

Private Function LayoutFlex(ByVal f As clsRenderBox, ByRef kids() As clsRenderBox, ByVal n As Long, ByVal cw As Double, _
        ByVal innerDef As Double, ByVal innerMin As Double) As Double
    Dim s As Long, row As Boolean, wrap As Boolean, colGap As Double, rowGap As Double, items() As TFlex, ic As Long
    Dim alignItems As String, justify As String, i As Long, k As clsRenderBox
    s = f.Slot
    row = (Left$(modStyle.SV(s, P_FLEXDIR), 3) = "row")
    wrap = (modStyle.SV(s, P_FLEXWRAP) <> "nowrap")
    If Not modStyle.SLen(s, P_COLGAP, cw, True, colGap) Then colGap = 0
    If Not modStyle.SLen(s, P_ROWGAP, innerDef, innerDef <> NOV, rowGap) Then rowGap = 0
    alignItems = modStyle.SV(s, P_ALIGNITEMS)
    justify = modStyle.SV(s, P_JUSTIFYCONTENT)
    ReDim items(0 To n)
    For i = 0 To n - 1
        Set k = kids(i)
        If IsAbsolute(k) Then
            f.AddAbsolute k
        ElseIf k.Kind = 3 Then
            If Not TextIsBlank(k.Text) Then
                Set items(ic).Box = k
                items(ic).IsText = True
                items(ic).Index = ic
                ic = ic + 1
            End If
        Else
            Set items(ic).Box = k
            items(ic).Index = ic
            ic = ic + 1
        End If
    Next i
    If ic = 0 Then Exit Function
    If row Then
        LayoutFlex = FlexRow(f, items, ic, cw, innerDef, innerMin, wrap, colGap, rowGap, alignItems, justify)
    Else
        LayoutFlex = FlexColumn(f, items, ic, cw, innerDef, innerMin, rowGap, alignItems, justify)
    End If
End Function

Private Function FlexColumn(ByVal f As clsRenderBox, ByRef items() As TFlex, ByVal ic As Long, ByVal cw As Double, ByVal innerDef As Double, _
        ByVal innerMin As Double, ByVal rowGap As Double, ByVal alignItems As String, ByVal justify As String) As Double
    Dim placed() As clsRenderBox, i As Long, child As clsRenderBox, align As String, e As New clsRenderBox, stretch As Boolean
    Dim v As Double, outer As Double, total As Double, container As Double, free As Double, offset As Double, between As Double
    Dim growSum As Double, autoCount As Long, share As Double, y As Double, g As Double
    ReDim placed(0 To ic - 1)
    For i = 0 To ic - 1
        If items(i).IsText Then
            Set child = LayoutAnonText(items(i).Box, f, cw, "f" & items(i).Index)
        Else
            align = AlignOf(items(i).Box, alignItems)
            e.Kind = 1
            SetEdges e, items(i).Box.Slot, cw
            stretch = (align = "stretch") And Not modStyle.SLen(items(i).Box.Slot, P_WIDTH, cw, True, v)
            If stretch Then
                Set child = LayoutBox(items(i).Box, cw, MaxD(0, cw - e.ML - e.MR), NOV, False, True, NOV)
            Else
                Set child = LayoutBox(items(i).Box, cw, NOV, NOV, True, True, NOV)
            End If
            outer = child.W + e.ML + e.MR
            Select Case align
                Case "center": child.X = e.ML + (cw - outer) / 2
                Case "flex-end", "end": child.X = e.ML + cw - outer
                Case Else: child.X = e.ML
            End Select
            If e.MLAuto And e.MRAuto Then child.X = (cw - child.W) / 2
        End If
        Set placed(i) = child
        total = total + child.H + child.MT + child.MB
    Next i
    total = total + rowGap * (ic - 1)
    If innerDef <> NOV Then container = innerDef Else container = MaxD(innerMin, total)
    free = container - total
    If free > 0 Then
        For i = 0 To ic - 1
            If placed(i).Kind <> 4 Then
                g = modStyle.SNum(placed(i).Slot, P_FLEXGROW)
                If g > 0 Then growSum = growSum + g
                If placed(i).MTAuto Then autoCount = autoCount + 1
                If placed(i).MBAuto Then autoCount = autoCount + 1
            End If
        Next i
        If growSum > 0 Then
            For i = 0 To ic - 1
                If placed(i).Kind <> 4 Then
                    g = modStyle.SNum(placed(i).Slot, P_FLEXGROW)
                    If g > 0 Then placed(i).H = placed(i).H + free * g / growSum
                End If
            Next i
        ElseIf autoCount > 0 Then
            share = free / autoCount
            For i = 0 To ic - 1
                If placed(i).MTAuto Then placed(i).AutoTop = share
                If placed(i).MBAuto Then placed(i).AutoBottom = share
            Next i
        ElseIf justify = "center" Then
            offset = free / 2
        ElseIf justify = "flex-end" Or justify = "end" Then
            offset = free
        ElseIf justify = "space-between" And ic > 1 Then
            between = free / (ic - 1)
        End If
    End If
    y = offset
    For i = 0 To ic - 1
        Set child = placed(i)
        y = y + child.AutoTop + child.MT
        child.Y = y
        y = y + child.H + child.MB + child.AutoBottom + rowGap + between
        f.AddKid child
    Next i
    FlexColumn = total
End Function

Private Function FlexRow(ByVal f As clsRenderBox, ByRef items() As TFlex, ByVal ic As Long, ByVal cw As Double, ByVal innerDef As Double, _
        ByVal innerMin As Double, ByVal wrap As Boolean, ByVal colGap As Double, ByVal rowGap As Double, ByVal alignItems As String, ByVal justify As String) As Double
    Dim i As Long, s As Long, e As New clsRenderBox, padX As Double, adjust As Double, basisText As String, base As Double, hasBase As Boolean
    Dim sw As Double, mn As Double, mx As Double, minW As Double, maxW As Double, v As Double
    Dim lineFirst() As Long, lineLast() As Long, lc As Long, used As Double, outer As Double, curFirst As Long
    Dim li As Long, y As Double
    For i = 0 To ic - 1
        If items(i).IsText Then
            OuterIntrinsic items(i).Box, mn, mx
            items(i).Base = mx: items(i).Hyp = mx: items(i).MinW = mn: items(i).MaxW = 1E+300
            items(i).Grow = 0: items(i).Shrink = 1: items(i).Margins = 0
        Else
            s = items(i).Box.Slot
            e.Kind = 1
            SetEdges e, s, cw
            items(i).Margins = e.ML + e.MR
            items(i).MLAuto = e.MLAuto
            items(i).MRAuto = e.MRAuto
            items(i).Grow = modStyle.SNum(s, P_FLEXGROW)
            items(i).Shrink = modStyle.SNum(s, P_FLEXSHRINK)
            padX = e.PL + e.PR + e.BL + e.BR
            If modStyle.SV(s, P_BOXSIZING) = "border-box" Then adjust = 0 Else adjust = padX
            basisText = modStyle.SV(s, P_FLEXBASIS)
            hasBase = False
            If basisText <> "auto" And basisText <> "content" Then
                hasBase = modStyle.SLen(s, P_FLEXBASIS, cw, True, base)
                If hasBase Then base = base + adjust
            End If
            If Not hasBase Then
                If modStyle.SLen(s, P_WIDTH, cw, True, sw) Then base = sw + adjust: hasBase = True
            End If
            Intrinsic items(i).Box, mn, mx
            If Not hasBase Then base = mx
            If modStyle.SV(s, P_MINW) = "auto" Then
                If modStyle.SV(s, P_OVERFLOWX) <> "visible" Then minW = 0 Else minW = mn
                If modStyle.SLen(s, P_WIDTH, cw, True, sw) Then minW = MinD(minW, sw + adjust)
            Else
                If modStyle.SLen(s, P_MINW, cw, True, v) Then minW = v + adjust Else minW = adjust
            End If
            If modStyle.SLen(s, P_MAXW, cw, True, maxW) Then maxW = maxW + adjust Else maxW = 1E+300
            items(i).Base = base
            items(i).MinW = MaxD(minW, padX)
            items(i).MaxW = maxW
            items(i).Hyp = MaxD(items(i).MinW, MinD(items(i).MaxW, base))
        End If
    Next i
    ' Lines.
    ReDim lineFirst(0 To ic): ReDim lineLast(0 To ic)
    curFirst = 0
    used = 0
    For i = 0 To ic - 1
        outer = items(i).Hyp + items(i).Margins
        If wrap And i > curFirst And used + colGap + outer > cw + 0.01 Then
            lineFirst(lc) = curFirst: lineLast(lc) = i - 1: lc = lc + 1
            curFirst = i
            used = 0
        End If
        If i > curFirst Then used = used + colGap
        used = used + outer
    Next i
    lineFirst(lc) = curFirst: lineLast(lc) = ic - 1: lc = lc + 1
    For li = 0 To lc - 1
        y = FlexLine(f, items, lineFirst(li), lineLast(li), lc, li, cw, innerDef, innerMin, colGap, rowGap, alignItems, justify, y)
    Next li
    FlexRow = y
End Function

' Resolves flexible lengths for one line and places its items; returns the next line's y.
Private Function FlexLine(ByVal f As clsRenderBox, ByRef items() As TFlex, ByVal a As Long, ByVal b As Long, ByVal lc As Long, ByVal li As Long, _
        ByVal cw As Double, ByVal innerDef As Double, ByVal innerMin As Double, ByVal colGap As Double, ByVal rowGap As Double, _
        ByVal alignItems As String, ByVal justify As String, ByVal y As Double) As Double
    Dim gaps As Double, i As Long, round_ As Long, free As Double, sum As Double, violated As Boolean, target As Double, active As Long
    Dim autoCount As Long, x As Double, between As Double, frags() As clsRenderBox, cross As Double, child As clsRenderBox
    Dim align As String, outer As Double, v As Double, count As Long
    count = b - a + 1
    gaps = colGap * (count - 1)
    For i = a To b
        items(i).Size = items(i).Hyp
        items(i).Frozen = False
        items(i).AutoLeft = 0
        items(i).AutoRight = 0
    Next i
    For round_ = 0 To 3
        free = cw - gaps
        active = 0
        For i = a To b
            free = free - items(i).Size - items(i).Margins
            If Not items(i).Frozen Then active = active + 1
        Next i
        If Abs(free) < 0.01 Or active = 0 Then Exit For
        sum = 0
        If free > 0 Then
            For i = a To b
                If Not items(i).Frozen Then sum = sum + items(i).Grow
            Next i
            If sum = 0 Then Exit For
            violated = False
            For i = a To b
                If Not items(i).Frozen Then
                    target = items(i).Size + free * items(i).Grow / sum
                    If target > items(i).MaxW Then
                        items(i).Size = items(i).MaxW: items(i).Frozen = True: violated = True
                    Else
                        items(i).Size = target
                    End If
                End If
            Next i
            If Not violated Then Exit For
        Else
            For i = a To b
                If Not items(i).Frozen Then sum = sum + items(i).Shrink * items(i).Base
            Next i
            If sum = 0 Then Exit For
            violated = False
            For i = a To b
                If Not items(i).Frozen Then
                    target = items(i).Size + free * items(i).Shrink * items(i).Base / sum
                    If target < items(i).MinW Then
                        items(i).Size = items(i).MinW: items(i).Frozen = True: violated = True
                    Else
                        items(i).Size = target
                    End If
                End If
            Next i
            If Not violated Then Exit For
        End If
    Next round_
    free = cw - gaps
    For i = a To b
        free = free - items(i).Size - items(i).Margins
        If items(i).MLAuto Then autoCount = autoCount + 1
        If items(i).MRAuto Then autoCount = autoCount + 1
    Next i
    ' Auto margins absorb the remaining space before justify-content does.
    If free > 0 And autoCount > 0 Then
        For i = a To b
            If items(i).MLAuto Then items(i).AutoLeft = free / autoCount
            If items(i).MRAuto Then items(i).AutoRight = free / autoCount
        Next i
    ElseIf free > 0 Then
        Select Case justify
            Case "center": x = free / 2
            Case "flex-end", "end", "right": x = free
            Case "space-between": If count > 1 Then between = free / (count - 1)
            Case "space-around": between = free / count: x = between / 2
            Case "space-evenly": between = free / (count + 1): x = between
        End Select
    End If
    ReDim frags(a To b)
    For i = a To b
        If items(i).IsText Then
            Set frags(i) = LayoutAnonText(items(i).Box, f, items(i).Size, "f" & items(i).Index)
        Else
            Set frags(i) = LayoutBox(items(i).Box, cw, items(i).Size, NOV, False, True, NOV)
        End If
        cross = MaxD(cross, frags(i).H + frags(i).MT + frags(i).MB)
    Next i
    If lc = 1 Then
        If innerDef <> NOV Then cross = MaxD(cross, innerDef) Else cross = MaxD(cross, innerMin)
    End If
    For i = a To b
        Set child = frags(i)
        x = x + items(i).AutoLeft
        child.X = x + child.ML
        If items(i).IsText Then
            align = alignItems
            If align = "normal" Then align = "stretch"
        Else
            align = AlignOf(items(i).Box, alignItems)
        End If
        outer = child.H + child.MT + child.MB
        If align = "stretch" And Not items(i).IsText Then
            If Not modStyle.SLen(items(i).Box.Slot, P_HEIGHT, 0, False, v) Then
                child.H = MaxD(child.H, cross - child.MT - child.MB)
            End If
            child.Y = y + child.MT
        ElseIf align = "center" Then
            child.Y = y + (cross - outer) / 2 + child.MT
        ElseIf align = "flex-end" Or align = "end" Then
            child.Y = y + cross - outer + child.MT
        Else
            child.Y = y + child.MT
        End If
        x = x + child.W + child.ML + child.MR + items(i).AutoRight + colGap + between
        f.AddKid child
    Next i
    If li < lc - 1 Then FlexLine = y + cross + rowGap Else FlexLine = y + cross
End Function

' ================================================================ positioned boxes

Private Sub PlaceAbsolute(ByVal box As clsRenderBox, ByVal containing As clsRenderBox)
    Dim s As Long, pw As Double, ph As Double, lft As Double, rgt As Double, tp As Double, btm As Double
    Dim hasL As Boolean, hasR As Boolean, hasT As Boolean, hasB As Boolean, e As New clsRenderBox, optW As Double, optH As Double
    Dim child As clsRenderBox
    s = box.Slot
    pw = containing.W - containing.BL - containing.BR
    ph = containing.H - containing.BT - containing.BB
    hasL = modStyle.SLen(s, P_LEFT, pw, True, lft)
    hasR = modStyle.SLen(s, P_RIGHT, pw, True, rgt)
    hasT = modStyle.SLen(s, P_TOP, ph, True, tp)
    hasB = modStyle.SLen(s, P_BOTTOM, ph, True, btm)
    e.Kind = box.Kind
    SetEdges e, s, pw
    optW = NOV: optH = NOV
    If hasL And hasR And modStyle.SV(s, P_WIDTH) = "auto" Then optW = MaxD(0, pw - lft - rgt - e.ML - e.MR)
    If hasT And hasB And modStyle.SV(s, P_HEIGHT) = "auto" Then optH = MaxD(0, ph - tp - btm - e.MT - e.MB)
    Set child = LayoutBox(box, pw, optW, optH, True, False, ph)
    If hasL Then
        child.X = containing.BL + lft + e.ML
    ElseIf hasR Then
        child.X = containing.BL + pw - rgt - child.W - e.MR
    Else
        child.X = containing.BL + e.ML
    End If
    If hasT Then
        child.Y = containing.BT + tp + e.MT
    ElseIf hasB Then
        child.Y = containing.BT + ph - btm - child.H - e.MB
    Else
        child.Y = containing.BT + e.MT
    End If
    child.Positioned = True
    containing.AddKid child
End Sub

Private Function LayoutFixed(ByVal box As clsRenderBox) As clsRenderBox
    Dim s As Long, lft As Double, rgt As Double, tp As Double, btm As Double
    Dim hasL As Boolean, hasR As Boolean, hasT As Boolean, hasB As Boolean, e As New clsRenderBox, optW As Double, optH As Double
    Dim f As clsRenderBox
    s = box.Slot
    hasL = modStyle.SLen(s, P_LEFT, mW, True, lft)
    hasR = modStyle.SLen(s, P_RIGHT, mW, True, rgt)
    hasT = modStyle.SLen(s, P_TOP, mH, True, tp)
    hasB = modStyle.SLen(s, P_BOTTOM, mH, True, btm)
    e.Kind = 1
    SetEdges e, s, mW
    optW = NOV: optH = NOV
    If hasL And hasR And modStyle.SV(s, P_WIDTH) = "auto" Then optW = MaxD(0, mW - lft - rgt - e.ML - e.MR)
    If hasT And hasB And modStyle.SV(s, P_HEIGHT) = "auto" Then optH = MaxD(0, mH - tp - btm - e.MT - e.MB)
    Set f = LayoutBox(box, mW, optW, optH, True, True, mH)
    If hasL Then
        f.X = lft + e.ML
    ElseIf hasR Then
        f.X = mW - rgt - f.W - e.MR
    Else
        f.X = e.ML
    End If
    If hasT Then
        f.Y = tp + e.MT
    ElseIf hasB Then
        f.Y = mH - btm - f.H - e.MB
    Else
        f.Y = e.MT
    End If
    Set LayoutFixed = f
End Function

' ================================================================ incremental layout

' Forgets every cached fragment and intrinsic size.
Public Sub ClearLayoutCache()
    Dim i As Long
    mUidCap = 1023
    ReDim mWalkPass(0 To mUidCap): ReDim mStKnown(0 To mUidCap)
    ReDim mStVer(0 To mUidCap): ReDim mStSelf(0 To mUidCap): ReDim mStBefore(0 To mUidCap)
    ReDim mStAfter(0 To mUidCap): ReDim mStParent(0 To mUidCap)
    ReDim mEHead(0 To mUidCap * 3 + 3)
    ReDim mIHas(0 To mUidCap * 3 + 3): ReDim mIMn(0 To mUidCap * 3 + 3): ReDim mIMx(0 To mUidCap * 3 + 3)
    mECap = 255
    mEUsed = 0
    mEFreeCount = 0
    ReDim mEFree(0 To 63)
    ReDim mENext(0 To mECap): ReDim mEFrag(0 To mECap): ReDim mEFlags(0 To mECap)
    ReDim mECbw(0 To mECap): ReDim mEOptW(0 To mECap): ReDim mEOptH(0 To mECap): ReDim mECbh(0 To mECap)
    ReDim mESave(0 To mECap * 8 + 7)
    ReDim mEPos(0 To mECap): ReDim mERole(0 To mECap): ReDim mEClick(0 To mECap): ReDim mEFixed(0 To mECap)
    mGlobalKey = ""
    i = 0
End Sub

Private Sub EnsureUidCap(ByVal uid As Long)
    Dim cap As Long
    If uid <= mUidCap Then Exit Sub
    cap = mUidCap
    Do While cap < uid
        cap = cap * 2 + 1
    Loop
    mUidCap = cap
    ReDim Preserve mWalkPass(0 To cap): ReDim Preserve mStKnown(0 To cap)
    ReDim Preserve mStVer(0 To cap): ReDim Preserve mStSelf(0 To cap): ReDim Preserve mStBefore(0 To cap)
    ReDim Preserve mStAfter(0 To cap): ReDim Preserve mStParent(0 To cap)
    ReDim Preserve mEHead(0 To cap * 3 + 3)
    ReDim Preserve mIHas(0 To cap * 3 + 3): ReDim Preserve mIMn(0 To cap * 3 + 3): ReDim Preserve mIMx(0 To cap * 3 + 3)
End Sub

' Starts a pass: clears the caches when something every box depends on changed, then
' walks the visible DOM to find the subtrees that changed since the last pass.
Private Sub BeginIncremental()
    Dim key As String, root As clsNode
    mPassNo = mPassNo + 1
    CacheHits = 0: CacheMisses = 0: WalkedNodes = 0: DirtyNodes = 0
    Set root = modDom.DocumentElement
    mCacheOn = Not root Is Nothing
    If Not mCacheOn Then Exit Sub
    key = mW & "|" & mH & "|" & mMobile & "|" & modStyle.StyleIsDark() & "|" & modStyle.StyleEpoch() & "|" & InlineStyleKey(root)
    If key <> mGlobalKey Then
        ClearLayoutCache
        mGlobalKey = key
    End If
    mNoneIdx = modStyle.Intern("none")
    mListItemIdx = modStyle.Intern("list-item")
    WalkNode root, 0
End Sub

' The element's inline declarations (the root's custom properties feed var()).
Private Function InlineStyleKey(ByVal el As clsNode) As String
    Dim names As clsList, i As Long, out As String
    Set names = el.StyleNames
    For i = 0 To names.Count - 1
        out = out & CStr(names(i)) & ":" & el.GetStyle(CStr(names(i))) & ";"
    Next i
    InlineStyleKey = out
End Function

' Walks a node laid out as part of the document (styles it, compares its state) and
' returns True when it or anything beneath it changed.
Private Function WalkNode(ByVal node As clsNode, ByVal parentVer As Long) As Boolean
    Dim uid As Long, s As Long, ver As Long, selfStamp As Long, bStamp As Long, aStamp As Long, pv As Long
    Dim dirty As Boolean, c As clsNode, t As Long, isCtl As Boolean, closed As Boolean, display As Long
    uid = node.Uid
    EnsureUidCap uid
    mWalkPass(uid) = mPassNo
    WalkedNodes = WalkedNodes + 1
    ver = node.Version
    s = modStyle.StyleOf(node)
    selfStamp = modStyle.SLayoutStamp(s)
    display = modStyle.SVI(s, P_DISPLAY)
    If display <> mNoneIdx Then
        isCtl = IsControlTag(node)
        If Not isCtl Then
            bStamp = modStyle.SLayoutStamp(modStyle.PseudoOf(node, 1))
            aStamp = modStyle.SLayoutStamp(modStyle.PseudoOf(node, 2))
            If node.LocalName = "details" Then closed = Not node.HasAttribute("open")
        End If
        ' List markers depend on the parent (item number, details open state).
        If display = mListItemIdx Then pv = parentVer
    End If
    If Not mStKnown(uid) Then
        dirty = True
    ElseIf mStVer(uid) <> ver Or mStSelf(uid) <> selfStamp Or mStBefore(uid) <> bStamp Or mStAfter(uid) <> aStamp Or mStParent(uid) <> pv Then
        dirty = True
    End If
    If dirty Then
        mStKnown(uid) = True
        mStVer(uid) = ver: mStSelf(uid) = selfStamp: mStBefore(uid) = bStamp: mStAfter(uid) = aStamp: mStParent(uid) = pv
    End If
    If display <> mNoneIdx Then
        Set c = node.FirstChild
        Do While Not c Is Nothing
            t = c.NodeType
            If t = 1 Then
                If isCtl Then
                    If WalkPlain(c) Then dirty = True
                ElseIf Not closed Then
                    If WalkNode(c, ver) Then dirty = True
                ElseIf c.LocalName = "summary" Then
                    If WalkNode(c, ver) Then dirty = True
                End If
            ElseIf t = 3 Then
                If Not closed Or isCtl Then
                    If WalkPlain(c) Then dirty = True
                End If
            End If
            Set c = c.NextSibling
        Loop
    End If
    If dirty Then
        DirtyNodes = DirtyNodes + 1
        InvalidateUid uid
    End If
    WalkNode = dirty
End Function

' Text nodes and the content of form controls: only their versions matter.
Private Function WalkPlain(ByVal node As clsNode) As Boolean
    Dim uid As Long, ver As Long, dirty As Boolean, c As clsNode
    uid = node.Uid
    EnsureUidCap uid
    mWalkPass(uid) = mPassNo
    ver = node.Version
    If Not mStKnown(uid) Then
        dirty = True
    ElseIf mStVer(uid) <> ver Then
        dirty = True
    End If
    If dirty Then mStKnown(uid) = True: mStVer(uid) = ver
    If node.NodeType = 1 Then
        Set c = node.FirstChild
        Do While Not c Is Nothing
            If WalkPlain(c) Then dirty = True
            Set c = c.NextSibling
        Loop
    End If
    WalkPlain = dirty
End Function

Private Sub InvalidateUid(ByVal uid As Long)
    Dim k As Long, w As Long, e As Long, nx As Long
    For w = 0 To 2
        k = uid * 3 + w + 1
        mIHas(k) = False
        e = mEHead(k)
        Do While e > 0
            nx = mENext(e)
            FreeEntry e
            e = nx
        Loop
        mEHead(k) = 0
    Next w
End Sub

' Cache index of a box this pass may cache (0 when it may not).
Private Function CacheKey(ByVal box As clsRenderBox) As Long
    Dim uid As Long
    If Not mCacheOn Then Exit Function
    If box.Kind <> 1 And box.Kind <> 2 Then Exit Function
    uid = box.Node.Uid
    If uid > mUidCap Or uid < 0 Then Exit Function
    If mWalkPass(uid) <> mPassNo Then Exit Function
    If box.Kind = 1 Then CacheKey = uid * 3 + 1 Else CacheKey = uid * 3 + box.Which + 1
End Function

Private Function NewEntry() As Long
    Dim e As Long
    If mEFreeCount > 0 Then
        mEFreeCount = mEFreeCount - 1
        NewEntry = mEFree(mEFreeCount)
        Exit Function
    End If
    mEUsed = mEUsed + 1
    e = mEUsed
    If e > mECap Then
        mECap = mECap * 2 + 1
        ReDim Preserve mENext(0 To mECap): ReDim Preserve mEFrag(0 To mECap): ReDim Preserve mEFlags(0 To mECap)
        ReDim Preserve mECbw(0 To mECap): ReDim Preserve mEOptW(0 To mECap): ReDim Preserve mEOptH(0 To mECap)
        ReDim Preserve mECbh(0 To mECap): ReDim Preserve mESave(0 To mECap * 8 + 7)
        ReDim Preserve mEPos(0 To mECap): ReDim Preserve mERole(0 To mECap): ReDim Preserve mEClick(0 To mECap)
        ReDim Preserve mEFixed(0 To mECap)
    End If
    NewEntry = e
End Function

Private Sub FreeEntry(ByVal e As Long)
    Set mEFrag(e) = Nothing
    Set mEClick(e) = Nothing
    Set mEFixed(e) = Nothing
    mENext(e) = 0
    If mEFreeCount > UBound(mEFree) Then ReDim Preserve mEFree(0 To mEFreeCount * 2)
    mEFree(mEFreeCount) = e
    mEFreeCount = mEFreeCount + 1
End Sub

' Remembers a freshly laid out fragment (at most two argument sets per box).
Private Sub StoreEntry(ByVal k As Long, ByVal f As clsRenderBox, ByVal cbw As Double, ByVal optW As Double, ByVal optH As Double, _
        ByVal cbh As Double, ByVal flags As Long, ByVal fixedBefore As Long)
    Dim e As Long, second As Long, i As Long, b As Long
    e = NewEntry()
    Set mEFrag(e) = f
    mECbw(e) = cbw: mEOptW(e) = optW: mEOptH(e) = optH: mECbh(e) = cbh: mEFlags(e) = flags
    b = e * 8
    mESave(b) = f.X: mESave(b + 1) = f.Y: mESave(b + 2) = f.H: mESave(b + 3) = f.AutoTop
    mESave(b + 4) = f.AutoBottom: mESave(b + 5) = f.AutoLeft: mESave(b + 6) = f.AutoRight: mESave(b + 7) = f.JustifyOffset
    mEPos(e) = f.Positioned
    mERole(e) = f.Role
    Set mEClick(e) = f.ClickNode
    If mFixedCount > fixedBefore Then
        Set mEFixed(e) = New Collection
        For i = fixedBefore To mFixedCount - 1
            mEFixed(e).Add mFixed(i)
        Next i
    End If
    second = mEHead(k)
    mENext(e) = second
    mEHead(k) = e
    If second > 0 Then
        If mENext(second) > 0 Then
            FreeEntry mENext(second)
            mENext(second) = 0
        End If
    End If
End Sub

Private Sub RestoreEntry(ByVal e As Long, ByVal f As clsRenderBox)
    Dim b As Long, box As Variant, fb As clsRenderBox, key As String
    b = e * 8
    f.X = mESave(b): f.Y = mESave(b + 1): f.H = mESave(b + 2): f.AutoTop = mESave(b + 3)
    f.AutoBottom = mESave(b + 4): f.AutoLeft = mESave(b + 5): f.AutoRight = mESave(b + 6): f.JustifyOffset = mESave(b + 7)
    f.Positioned = mEPos(e)
    f.Role = mERole(e)
    Set f.ClickNode = mEClick(e)
    ' Overlays found inside the subtree when it was laid out.
    If Not mEFixed(e) Is Nothing Then
        For Each box In mEFixed(e)
            Set fb = box
            key = CStr(fb.Node.Uid)
            If IsEmpty(mSeenFixed(key)) Then
                mSeenFixed(key) = True
                fb.Slot = modStyle.StyleOf(fb.Node)
                If mFixedCount = 0 Then
                    ReDim mFixed(0 To 3)
                ElseIf mFixedCount > UBound(mFixed) Then
                    ReDim Preserve mFixed(0 To mFixedCount * 2)
                End If
                Set mFixed(mFixedCount) = fb
                mFixedCount = mFixedCount + 1
            End If
        Next box
    End If
End Sub

' ================================================================ helpers for clsRenderGrid

Friend Function NoValue() As Double
    NoValue = NOV
End Function

Friend Sub TableRowsOf(ByVal table As clsNode, ByRef rows As Collection)
    Dim c As clsNode, cell As clsNode, cells As Collection, cs As Long
    Set c = table.FirstElementChild
    Do While Not c Is Nothing
        If c.LocalName = "tr" Then
            Set cells = New Collection
            Set cell = c.FirstElementChild
            Do While Not cell Is Nothing
                If cell.LocalName = "td" Or cell.LocalName = "th" Then
                    cs = modStyle.StyleOf(cell)
                    If modStyle.SV(cs, P_DISPLAY) <> "none" Then cells.Add ElementBox(cell)
                End If
                Set cell = cell.NextElementSibling
            Loop
            rows.Add cells
        ElseIf c.LocalName = "thead" Or c.LocalName = "tbody" Or c.LocalName = "tfoot" Then
            TableRowsOf c, rows
        End If
        Set c = c.NextElementSibling
    Loop
End Sub
