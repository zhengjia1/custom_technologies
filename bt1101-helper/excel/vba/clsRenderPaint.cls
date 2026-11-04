VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRenderPaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Turns the fragment tree of clsRenderLayout into the display list (clsRenderSpec): one
' region per frame (title bar, library, page, nested scrollers, fixed overlays) and the
' rect/text/hit/control items drawn in each, in paint order (port of the reference
' engine's paint pass). Colours are resolved to opaque RGB by blending translucent
' colours and opacity over the background beneath.
Option Explicit

Public Engine As clsRenderLayout
Public PathIds As Boolean           ' ids from element paths (tests) instead of node uids
Public ScrollTarget As clsNode      ' element to scroll into view (scrollIntoView)
Public ScrollTargetRegion As String ' its region and position there (after Paint)
Public ScrollTargetY As Double
Public FocusVisible As Boolean      ' focus came from the keyboard: draw :focus-visible outlines

Private mRegions As Collection
Private mPage As clsRenderSpec
Private mW As Double, mH As Double
Private mFocusNode As clsNode
Private mFocusId As String
Private mPageBg As Long
Private mPaintIndex As Long
Private mSymbolFallback As Boolean
Private mSticky As Collection       ' position: sticky elements met in a scroller region (Array(f, ax, ay, region, backdrop, opacity, tip, click))
Private mNoneIdx As Long            ' interned "none" (box-shadow test)
Private mStickyCopy As Boolean      ' painting a sticky element's pinned copy

Private Function MaxD(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxD = a Else MaxD = b
End Function

Private Function MinD(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinD = a Else MinD = b
End Function

' ================================================================ entry

Public Sub Paint()
    Dim root As clsRenderBox, i As Long, f As clsRenderBox, rgbValue As Long, alpha As Double, backdrop As Long
    Dim region As clsRenderSpec, fx As Double, fy As Double, s As Long, bodySlot As Long
    mW = Engine.ViewWidth
    mH = Engine.ViewHeight
    mSymbolFallback = modRender.SymbolFallbackActive()
    Set mRegions = New Collection
    Set mFocusNode = modDom.ActiveElement()
    mFocusId = ""
    mPaintIndex = 0
    ScrollTargetRegion = ""
    mPageBg = RGB(255, 255, 255)
    If Not modDom.Body Is Nothing Then
        bodySlot = modStyle.StyleOf(modDom.Body)
        If modStyle.SColour(bodySlot, P_BG, rgbValue, alpha) Then mPageBg = Blend(rgbValue, alpha, RGB(255, 255, 255))
    End If
    Set mPage = NewRegion("page", "", 0, 0, mW, mH, mH, mPageBg)
    Set root = Engine.Root
    mNoneIdx = modStyle.Intern("none")
    Set mSticky = New Collection
    Walk root, 0, 0, mPage, mPageBg, 1, "", Nothing
    For i = 1 To mSticky.Count
        PaintStickyCopy mSticky(i)
    Next i
    Set mSticky = Nothing
    For i = 0 To Engine.FixedCount - 1
        Set f = Engine.FixedFrag(i)
        s = f.Slot
        If OpacityOf(s) > 0.01 And modStyle.SV(s, P_VISIBILITY) <> "hidden" Then
            backdrop = mPageBg
            If modStyle.SColour(s, P_BG, rgbValue, alpha) Then backdrop = Blend(rgbValue, alpha, mPageBg)
            fx = f.X: fy = f.Y
            f.Role = 6
            Set region = OpenRegion(f, fx, fy, Nothing, backdrop)
            f.X = 0: f.Y = 0: f.Role = 0
            Walk f, fx, fy, region, backdrop, 1, "", Nothing
        End If
    Next i
End Sub

' Regions in paint order (the page region first when it has items, overlays last).
Public Function Regions() As Collection
    Dim out As New Collection, r As clsRenderSpec
    For Each r In mRegions
        If Not r.Overlay Then If Not (r Is mPage And r.Items.Count = 0) Then out.Add r
    Next r
    For Each r In mRegions
        If r.Overlay Then out.Add r
    Next r
    Set Regions = out
End Function

Public Property Get PageBg() As Long
    PageBg = mPageBg
End Property

Public Property Get FocusId() As String
    FocusId = mFocusId
End Property

' ================================================================ ids

Public Function IdOf(ByVal el As clsNode) As String
    Dim parts As String, node As clsNode, parent As clsNode, idx As Long, c As clsNode
    If el Is Nothing Then IdOf = "anon": Exit Function
    If Not PathIds Then IdOf = "u" & el.Uid: Exit Function
    Set node = el
    Do
        Set parent = node.ParentNode
        If parent Is Nothing Then Exit Do
        If parent.NodeType <> 1 Then Exit Do
        idx = 0
        Set c = parent.FirstElementChild
        Do While Not c Is Nothing
            If c Is node Then Exit Do
            idx = idx + 1
            Set c = c.NextElementSibling
        Loop
        If Len(parts) = 0 Then parts = CStr(idx) Else parts = idx & "." & parts
        Set node = parent
    Loop
    IdOf = "p" & parts
End Function

Private Function PseudoSuffix(ByVal which As Long) As String
    If which = 1 Then PseudoSuffix = ":before" Else PseudoSuffix = ":after"
End Function

' ================================================================ colours

Public Function Blend(ByVal rgbValue As Long, ByVal alpha As Double, ByVal backdrop As Long) As Long
    Dim r As Double, g As Double, b As Double
    If alpha >= 1 Then Blend = rgbValue: Exit Function
    If alpha < 0 Then alpha = 0
    r = (rgbValue And &HFF&) * alpha + (backdrop And &HFF&) * (1 - alpha)
    g = ((rgbValue \ &H100&) And &HFF&) * alpha + ((backdrop \ &H100&) And &HFF&) * (1 - alpha)
    b = ((rgbValue \ &H10000) And &HFF&) * alpha + ((backdrop \ &H10000) And &HFF&) * (1 - alpha)
    Blend = RGB(CLng(r), CLng(g), CLng(b))
End Function

Private Function OpacityOf(ByVal s As Long) As Double
    Dim t As String
    t = modStyle.SV(s, P_OPACITY)
    If Len(t) = 0 Then OpacityOf = 1 Else OpacityOf = modStyle.SNum(s, P_OPACITY)
    If Right$(t, 1) = "%" Then OpacityOf = OpacityOf / 100
End Function

' ================================================================ regions and items

Private Function NewRegion(ByVal id As String, ByVal parentId As String, ByVal x As Double, ByVal y As Double, ByVal w As Double, _
        ByVal h As Double, ByVal contentH As Double, ByVal bg As Long) As clsRenderSpec
    Dim r As New clsRenderSpec
    r.Kind = "region"
    r.Id = id
    r.Region = parentId
    r.X = x: r.Y = y: r.W = w: r.H = h
    r.ContentH = contentH
    r.ContentW = w
    r.Bg = bg
    Set r.Items = New Collection
    mRegions.Add r
    Set NewRegion = r
End Function

Private Function OpenRegion(ByVal f As clsRenderBox, ByVal ax As Double, ByVal ay As Double, ByVal parent As clsRenderSpec, ByVal backdrop As Long) As clsRenderSpec
    Dim r As clsRenderSpec, lft As Double, tp As Double, px As Double, py As Double, w As Double, h As Double
    Select Case f.Role
        Case 3
            lft = ax - f.ML: tp = ay - f.MT
            Set r = NewRegion("main", "", lft, tp, mW - lft, mH - tp, f.H + f.MT + f.MB, backdrop)
            r.OX = lft: r.OY = tp
        Case 4
            Set r = NewRegion("main", "", ax, ay, mW - ax, mH - ay, f.H, backdrop)
            r.OX = ax: r.OY = ay
        Case 2
            Set r = NewRegion("library", "", ax, ay, f.W, MinD(f.H, mH - ay), MaxD(f.H, f.ScrollH + f.BT + f.BB), backdrop)
            r.OX = ax: r.OY = ay
        Case 1
            Set r = NewRegion("top", "", ax, ay, mW - ax, f.H, f.H, backdrop)
            r.OX = ax: r.OY = ay
        Case 6
            Set r = NewRegion("ov:" & IdOf(f.Node), "", ax, ay, f.W, MinD(f.H, mH), MaxD(f.H, f.ScrollH), backdrop)
            r.OX = ax: r.OY = ay
            r.Overlay = True
        Case Else
            px = ax + f.BL: py = ay + f.BT
            w = f.W - f.BL - f.BR
            h = f.H - f.BT - f.BB
            If f.Scrolls Then
                Set r = NewRegion("sc:" & IdOf(f.Node), "", px, py, w, h, f.ScrollH, backdrop)
            Else
                Set r = NewRegion("sc:" & IdOf(f.Node), "", px, py, w, h, h, backdrop)
            End If
            r.OX = px: r.OY = py
            If f.ScrollsX Then r.ContentW = f.ScrollW
    End Select
    If Not parent Is Nothing And Not r.Overlay Then
        r.Region = parent.Id
        r.X = r.X - parent.OX
        r.Y = r.Y - parent.OY
    End If
    r.Scrolls = (r.ContentH > r.H + 0.5)
    If r.ContentW < r.W Then r.ContentW = r.W
    r.ScrollsX = (r.ContentW > r.W + 0.5 And modStyle.SV(f.Slot, P_OVERFLOWX) <> "hidden")
    Set OpenRegion = r
End Function

Private Sub PushItem(ByVal region As clsRenderSpec, ByVal it As clsRenderSpec)
    If mStickyCopy Then
        ' The pinned copy draws and takes clicks; focus stays with the element in the flow.
        If Not it.Node Is Nothing Then Exit Sub
        it.Id = "st|" & it.Id
    End If
    it.Region = region.Id
    it.X = it.X - region.OX
    it.Y = it.Y - region.OY
    mPaintIndex = mPaintIndex + 1
    it.PaintIndex = mPaintIndex
    region.Items.Add it
End Sub

Private Function NewItem(ByVal kind As String, ByVal id As String, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double) As clsRenderSpec
    Dim it As New clsRenderSpec
    it.Kind = kind
    it.Id = id
    it.X = x: it.Y = y: it.W = w: it.H = h
    Set NewItem = it
End Function

' ================================================================ walking fragments

Private Sub Walk(ByVal f As clsRenderBox, ByVal px As Double, ByVal py As Double, ByVal region As clsRenderSpec, ByVal backdrop As Long, _
        ByVal opacity As Double, ByVal tip As String, ByVal click As clsNode)
    Dim ax As Double, ay As Double, s As Long, el As clsNode, opacityHere As Double, hidden As Boolean, ownTip As String
    Dim lc As clsNode, roleRegion As Boolean, rgbValue As Long, alpha As Double, hasBg As Boolean, bg As Long, idBase As String
    Dim isButton As Boolean, startCount As Long, childBackdrop As Long, inner As clsRenderSpec, i As Long, k As clsRenderBox
    Dim childClick As clsNode, it As clsRenderSpec
    ax = px + f.X
    ay = py + f.Y
    If f.Kind = 4 Then
        PaintLines f, ax, ay, region, backdrop, opacity, click
        For i = 0 To f.KidCount - 1
            Set k = f.Kid(i)
            Set childClick = click
            If Not k.ClickNode Is Nothing Then Set childClick = k.ClickNode
            Walk k, ax + f.LinesX, ay + f.LinesY, region, backdrop, opacity, tip, childClick
        Next i
        Exit Sub
    End If
    s = f.Slot
    Set el = f.Node
    opacityHere = OpacityOf(s)
    If opacityHere <= 0.01 Then Exit Sub
    If opacityHere < 1 Then opacity = opacity * opacityHere
    hidden = (modStyle.SV(s, P_VISIBILITY) = "hidden")
    If f.Kind = 1 Then
        If el Is ScrollTarget Then
            If Len(ScrollTargetRegion) = 0 Then ScrollTargetRegion = region.Id: ScrollTargetY = ay - region.OY
        End If
        ownTip = TipFor(el)
        If Len(ownTip) > 0 Then tip = ownTip
        Set lc = Engine.LabelClick(el)
        If Not lc Is Nothing Then Set click = lc
    End If
    If f.Role = 5 Then
        Set it = NewItem("hit", IdOf(el) & ":hit", ax, ay, f.W, f.H)
        Set it.ClickNode = el
        it.Tip = AttrText(el, "title")
        it.Resizer = True
        PushItem region, it
        Exit Sub
    End If
    roleRegion = (f.Role >= 1 And f.Role <= 4)
    If roleRegion Then Set region = OpenRegion(f, ax, ay, Nothing, backdrop)
    If modStyle.SColour(s, P_BG, rgbValue, alpha) Then
        hasBg = True
        bg = Blend(rgbValue, alpha * opacity, backdrop)
    End If
    idBase = IdOf(el)
    If f.Kind = 2 Then idBase = idBase & PseudoSuffix(f.Which)
    If Not hidden And f.Kind = 1 Then
        If Engine.IsControlTag(el) Then
            PaintControl f, ax, ay, region, backdrop, opacity, tip
            If el Is mFocusNode Then mFocusId = IdOf(el)
            Exit Sub
        End If
    End If
    isButton = (Not hidden And f.Kind = 1)
    If isButton Then isButton = (el.LocalName = "button")
    startCount = region.Items.Count
    If Not hidden Then PaintShadows f, ax, ay, region, idBase, backdrop, opacity
    If Not hidden And (hasBg Or HasBorder(f)) Then PaintBox f, ax, ay, region, idBase, hasBg, bg, backdrop, opacity
    If hasBg Then childBackdrop = bg Else childBackdrop = backdrop
    If Not hidden And f.HasLines Then PaintLines f, ax, ay, region, childBackdrop, opacity, click
    Set inner = region
    If (f.Scrolls Or f.ScrollsX) And Not roleRegion And f.Role <> 6 Then Set inner = OpenRegion(f, ax, ay, region, childBackdrop)
    If isButton Then Set childClick = Nothing Else Set childClick = click
    For i = 0 To f.KidCount - 1
        If Not f.Kid(i).Positioned Then Walk f.Kid(i), ax, ay, inner, childBackdrop, opacity, tip, childClick
    Next i
    For i = 0 To f.KidCount - 1
        If f.Kid(i).Positioned Then Walk f.Kid(i), ax, ay, inner, childBackdrop, opacity, tip, childClick
    Next i
    If isButton Then
        CollapseButton f, el, ax, ay, region, startCount, backdrop, opacity, tip
        If el Is mFocusNode Then mFocusId = IdOf(el)
    ElseIf f.Kind = 1 Then
        If el.LocalName = "summary" Then
            Set it = NewItem("hit", IdOf(el) & ":hit", ax, ay, f.W, f.H)
            Set it.ClickNode = el
            it.Tip = tip
            PushItem region, it
        End If
    End If
    ' A sticky element in the page's scroller also gets a pinned copy (PaintStickyCopy).
    If f.Kind = 1 And Not hidden And Not roleRegion And Not mStickyCopy And Not PathIds And Not mSticky Is Nothing Then
        If region.Id = "main" And Len(region.Region) = 0 Then
            If modStyle.SV(s, P_POSITION) = "sticky" Then mSticky.Add Array(f, ax, ay, region, backdrop, opacity, tip, click)
        End If
    End If
End Sub

' position: sticky in the page's scroller: a copy of the element in an overlay region at the
' scroller's top edge (plus the element's top offset), which modRender shows only while
' the scroller is scrolled past the element's own position.
Private Sub PaintStickyCopy(ByVal v As Variant)
    Dim f As clsRenderBox, ax As Double, ay As Double, host As clsRenderSpec, r As clsRenderSpec, topInset As Double, click As clsNode
    Set f = v(0): ax = v(1): ay = v(2): Set host = v(3)
    If IsObject(v(7)) Then Set click = v(7)
    If Not modStyle.SLen(f.Slot, P_TOP, 0, False, topInset) Then Exit Sub
    Set r = NewRegion("st:" & IdOf(f.Node), "", ax, host.Y + topInset, f.W, f.H, f.H, CLng(v(4)))
    r.OX = ax: r.OY = ay
    r.Overlay = True
    r.StickyOf = host.Id
    r.StickyY = (ay - host.OY) - topInset
    mStickyCopy = True
    Walk f, ax - f.X, ay - f.Y, r, CLng(v(4)), CDbl(v(5)), CStr(v(6)), click
    mStickyCopy = False
End Sub

Private Function HasBorder(ByVal f As clsRenderBox) As Boolean
    HasBorder = (f.BT > 0 Or f.BR > 0 Or f.BB > 0 Or f.BL > 0)
End Function

Private Function AttrText(ByVal el As clsNode, ByVal name As String) As String
    Dim v As Variant
    v = el.GetAttribute(name)
    If Not IsNull(v) Then AttrText = CStr(v)
End Function

Private Function TipFor(ByVal el As clsNode) As String
    Dim tt As clsNode
    If el.HasClass("metric-info") Then
        Set tt = el.QuerySelector(".metric-tooltip")
        If Not tt Is Nothing Then TipFor = Application.WorksheetFunction.Trim(Replace(Replace(tt.TextContent, vbLf, " "), vbTab, " ")): Exit Function
    End If
    TipFor = AttrText(el, "title")
End Function

Private Function SideColour(ByVal s As Long, ByVal p As Long, ByVal over As Long, ByVal opacity As Double, ByRef c As Long) As Boolean
    Dim rgbValue As Long, alpha As Double
    If modStyle.SColour(s, p, rgbValue, alpha) Then
        c = Blend(rgbValue, alpha * opacity, over)
        SideColour = True
    End If
End Function

Private Sub PaintBox(ByVal f As clsRenderBox, ByVal ax As Double, ByVal ay As Double, ByVal region As clsRenderSpec, ByVal idBase As String, _
        ByVal hasBg As Boolean, ByVal bg As Long, ByVal backdrop As Long, ByVal opacity As Double)
    Dim widths(0 To 3) As Double, cols(0 To 3) As Long, has(0 To 3) As Boolean, i As Long, over As Long, uniform As Boolean
    Dim it As clsRenderSpec, s As Long
    s = f.Slot
    widths(0) = f.BT: widths(1) = f.BR: widths(2) = f.BB: widths(3) = f.BL
    If hasBg Then over = bg Else over = backdrop
    For i = 0 To 3
        If widths(i) > 0 Then has(i) = SideColour(s, P_BTC + i, over, opacity, cols(i))
    Next i
    If f.HasLegend Then
        ' A fieldset's border starts at the middle of its legend and skips the legend.
        PaintLegendBox f, ax, ay, region, idBase, hasBg, bg, widths, has, cols
        Exit Sub
    End If
    uniform = (widths(0) > 0 And has(0))
    For i = 1 To 3
        If widths(i) <> widths(0) Or has(i) <> has(0) Or cols(i) <> cols(0) Then uniform = False
    Next i
    ' MSForms draws a one-pixel border; thicker borders are drawn as side rects.
    If uniform And widths(0) <= 1.5 Then
        Set it = NewItem("rect", idBase & ":box", ax, ay, f.W, f.H)
        If hasBg Then it.Bg = bg
        it.Bc = cols(0)
        it.Bw = widths(0)
        PushItem region, it
        Exit Sub
    End If
    If hasBg Then
        Set it = NewItem("rect", idBase & ":box", ax, ay, f.W, f.H)
        it.Bg = bg
        PushItem region, it
    End If
    If widths(0) > 0 And has(0) Then Set it = NewItem("rect", idBase & ":bt", ax, ay, f.W, widths(0)): it.Bg = cols(0): PushItem region, it
    If widths(2) > 0 And has(2) Then Set it = NewItem("rect", idBase & ":bb", ax, ay + f.H - widths(2), f.W, widths(2)): it.Bg = cols(2): PushItem region, it
    If widths(3) > 0 And has(3) Then Set it = NewItem("rect", idBase & ":bl", ax, ay, widths(3), f.H): it.Bg = cols(3): PushItem region, it
    If widths(1) > 0 And has(1) Then Set it = NewItem("rect", idBase & ":br", ax + f.W - widths(1), ay, widths(1), f.H): it.Bg = cols(1): PushItem region, it
End Sub

Private Sub PaintLegendBox(ByVal f As clsRenderBox, ByVal ax As Double, ByVal ay As Double, ByVal region As clsRenderSpec, ByVal idBase As String, _
        ByVal hasBg As Boolean, ByVal bg As Long, ByRef widths() As Double, ByRef has() As Boolean, ByRef cols() As Long)
    Dim it As clsRenderSpec, top As Double, h As Double
    top = ay + f.LegendLineY
    h = f.H - f.LegendLineY
    If hasBg Then Set it = NewItem("rect", idBase & ":box", ax, top, f.W, h): it.Bg = bg: PushItem region, it
    If widths(0) > 0 And has(0) Then
        If f.LegendX0 > 0.5 Then Set it = NewItem("rect", idBase & ":bt0", ax, top, f.LegendX0, widths(0)): it.Bg = cols(0): PushItem region, it
        If f.LegendX1 < f.W - 0.5 Then Set it = NewItem("rect", idBase & ":bt", ax + f.LegendX1, top, f.W - f.LegendX1, widths(0)): it.Bg = cols(0): PushItem region, it
    End If
    If widths(2) > 0 And has(2) Then Set it = NewItem("rect", idBase & ":bb", ax, ay + f.H - widths(2), f.W, widths(2)): it.Bg = cols(2): PushItem region, it
    If widths(3) > 0 And has(3) Then Set it = NewItem("rect", idBase & ":bl", ax, top, widths(3), h): it.Bg = cols(3): PushItem region, it
    If widths(1) > 0 And has(1) Then Set it = NewItem("rect", idBase & ":br", ax + f.W - widths(1), top, widths(1), h): it.Bg = cols(1): PushItem region, it
End Sub

Private Sub PaintLines(ByVal f As clsRenderBox, ByVal ax As Double, ByVal ay As Double, ByVal region As clsRenderSpec, ByVal backdrop As Long, _
        ByVal opacity As Double, ByVal click As clsNode)
    Dim owner As clsRenderBox, base As String, prefix As String, ox As Double, oy As Double, i As Long, s As Long
    Dim rgbValue As Long, alpha As Double, colour As Long, it As clsRenderSpec, rc As clsNode, text As String
    If f.Kind = 4 Then Set owner = f.Owner Else Set owner = f
    base = IdOf(owner.Node)
    If owner.Kind = 2 Then base = base & PseudoSuffix(owner.Which)
    If f.Kind = 4 Then prefix = base & ":a" & f.AnonIndex & ":L" Else prefix = base & ":L"
    ox = ax + f.LinesX
    oy = ay + f.LinesY
    For i = 0 To f.RunCount - 1
        s = f.RunSlot(i)
        If modStyle.SV(s, P_VISIBILITY) <> "hidden" Then
            If modStyle.SColour(s, P_COLOR, rgbValue, alpha) Then colour = Blend(rgbValue, alpha * opacity, backdrop) Else colour = Blend(0, opacity, backdrop)
            Set rc = f.RunClick(i)
            If rc Is Nothing Then Set rc = click
            text = f.RunText(i)
            If mSymbolFallback Then
                If modRender.NeedsSymbolFont(text) Then
                    PaintSplitRun prefix & f.RunLine(i) & "." & f.RunIndex(i), ox + f.RunX(i), oy + f.RunY(i), f.RunH(i), text, s, colour, rc, region
                    GoTo NextRun
                End If
            End If
            Set it = NewItem("text", prefix & f.RunLine(i) & "." & f.RunIndex(i), ox + f.RunX(i), oy + f.RunY(i), f.RunW(i) + 2, f.RunH(i))
            FillText it, text, s, colour, rc
            PushItem region, it
        End If
NextRun:
    Next i
End Sub

Private Sub FillText(ByVal it As clsRenderSpec, ByVal text As String, ByVal s As Long, ByVal colour As Long, ByVal click As clsNode)
    it.Text = text
    it.Size = Round(modStyle.SFont(s), 2)
    it.Weight = modStyle.SWeight(s)
    it.Italic = modStyle.SItalic(s)
    it.Color = colour
    it.Align = "l"
    Set it.ClickNode = click
End Sub

' A run with characters the text font lacks: pieces alternate between the text font and
' the symbol font.
Private Sub PaintSplitRun(ByVal id As String, ByVal x As Double, ByVal y As Double, ByVal h As Double, ByVal text As String, ByVal s As Long, _
        ByVal colour As Long, ByVal click As clsNode, ByVal region As clsRenderSpec)
    Dim i As Long, startPos As Long, cur As Boolean, sym As Boolean, piece As String, n As Long, it As clsRenderSpec, w As Double
    startPos = 1
    cur = Not modStyle.GlyphCovered(AscW(Mid$(text, 1, 1)) And &HFFFF&)
    For i = 2 To Len(text) + 1
        If i <= Len(text) Then sym = Not modStyle.GlyphCovered(AscW(Mid$(text, i, 1)) And &HFFFF&) Else sym = Not cur
        If sym <> cur Then
            piece = Mid$(text, startPos, i - startPos)
            w = Engine.Measure(piece, s)
            If n = 0 Then Set it = NewItem("text", id, x, y, w + 2, h) Else Set it = NewItem("text", id & "~" & n, x, y, w + 2, h)
            FillText it, piece, s, colour, click
            it.Symbol = cur
            If piece = ChrW$(&H2304) Then CentreArrowhead it, x, y, w, h
            PushItem region, it
            x = x + w
            n = n + 1
            startPos = i
            cur = sym
        End If
    Next i
End Sub

' The open-card chevron (the symbol font's down arrowhead, whose ink is 0.37 em wide and
' centred 0.35 em from the left, 0.96 em from the top of a label): a little smaller than
' the 21px it stands in for and centred on its line box, like the website's rotated ">".
Private Sub CentreArrowhead(ByVal it As clsRenderSpec, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double)
    Dim sz As Double
    sz = Round(it.Size * 0.85, 1)
    it.Size = sz
    it.X = x + w / 2 - 0.35 * sz
    it.Y = y + h / 2 - 0.96 * sz + 1
    it.W = sz
    it.H = sz * 1.2
End Sub

Private Sub ControlColours(ByVal f As clsRenderBox, ByVal backdrop As Long, ByVal opacity As Double, ByRef bg As Long, ByRef fg As Long, ByRef bc As Long, ByRef bw As Double)
    Dim s As Long, rgbValue As Long, alpha As Double, i As Long, widths(0 To 3) As Double
    s = f.Slot
    bg = backdrop
    If modStyle.SColour(s, P_BG, rgbValue, alpha) Then bg = Blend(rgbValue, alpha * opacity, backdrop)
    fg = 0
    If modStyle.SColour(s, P_COLOR, rgbValue, alpha) Then fg = Blend(rgbValue, alpha * opacity, bg)
    widths(0) = f.BT: widths(1) = f.BR: widths(2) = f.BB: widths(3) = f.BL
    bc = -1: bw = 0
    For i = 0 To 3
        If widths(i) > bw Then bw = widths(i)
    Next i
    For i = 0 To 3
        If widths(i) > 0 Then
            If SideColour(s, P_BTC + i, bg, opacity, bc) Then Exit For
        End If
    Next i
    If bc < 0 Then bw = 0
End Sub

' Text boxes and lists sit inside a rect that draws the website's padding and border, so
' their text lines up with the page rather than the control's edge.
Private Sub PaintControl(ByVal f As clsRenderBox, ByVal ax As Double, ByVal ay As Double, ByVal region As clsRenderSpec, ByVal backdrop As Long, _
        ByVal opacity As Double, ByVal tip As String)
    Dim el As clsNode, s As Long, it As clsRenderSpec, bg As Long, fg As Long, bc As Long, bw As Double, toggle As Boolean, t As String
    Dim rgbValue As Long, alpha As Double, ownBg As Long, hasOwn As Boolean, natural As Double, lft As Double, rgt As Double
    Dim tp As Double, btm As Double, innerH As Double, opts As clsList, i As Long, caption As String, flags As String, p As Variant
    Set el = f.Node
    s = f.Slot
    ControlColours f, backdrop, opacity, bg, fg, bc, bw
    t = el.TypeName
    toggle = (el.LocalName = "input" And (t = "checkbox" Or t = "radio"))
    Set it = NewItem("", IdOf(el), ax, ay, f.W, f.H)
    Set it.Node = el
    it.Size = Round(modStyle.SFont(s), 2)
    it.Weight = modStyle.SWeight(s)
    it.Tip = tip
    it.Disabled = el.Disabled
    it.Bg = bg
    it.Color = fg
    If toggle Then
        If IsSwitch(f, el) Then
            it.Kind = "check"
            it.Checked = el.Checked
            PaintSwitch f, ax, ay, region, backdrop, opacity, tip, it
            Exit Sub
        End If
        it.Bc = bc
        it.Bw = bw
    Else
        If modStyle.SColour(s, P_BG, rgbValue, alpha) Then hasOwn = True: ownBg = Blend(rgbValue, alpha * opacity, backdrop)
        PaintShadows f, ax, ay, region, IdOf(el), backdrop, opacity
        PaintBox f, ax, ay, region, IdOf(el), hasOwn, ownBg, backdrop, opacity
        natural = Engine.NaturalLineHeight(s)
        lft = f.BL + MaxD(0, f.PL - 3)
        rgt = f.BR + MaxD(0, f.PR - 3)
        it.X = ax + lft
        it.W = MaxD(8, f.W - lft - rgt)
        If el.LocalName = "textarea" Then
            tp = f.BT + MaxD(0, f.PT - 3)
            btm = f.BB + MaxD(0, f.PB - 3)
            it.Y = ay + tp
            it.H = MaxD(natural + 6, f.H - tp - btm)
        Else
            innerH = f.H - f.BT - f.BB
            it.H = MinD(innerH, natural + 8)
            it.Y = ay + f.BT + (innerH - it.H) / 2
        End If
    End If
    If el.LocalName = "select" Then
        it.Kind = "select"
        Set opts = el.Options
        For i = 0 To opts.Count - 1
            Set p = opts(i)
            caption = Application.WorksheetFunction.Trim(Replace(Replace(p.TextContent, vbLf, " "), vbTab, " "))
            If i > 0 Then it.Options = it.Options & vbLf
            it.Options = it.Options & caption
            If p.Disabled Then flags = flags & "1" Else flags = flags & "0"
        Next i
        it.OptionFlags = flags
        it.SelIndex = el.SelectedIndex
        PushItem region, it
        ' The list's own drop button is hidden: a chevron sits in the right padding, as on
        ' the website, and opens the list when clicked.
        If f.PR >= 12 Then PaintChevron el, s, ax + f.W - f.BR - f.PR / 2, ay, f.H, fg, region
        PaintOutline f, ax, ay, region, IdOf(el), backdrop, opacity
        Exit Sub
    End If
    If toggle Then
        it.Kind = "check"
        it.IsRadio = (t = "radio")
        it.Checked = el.Checked
        PushItem region, it
        PaintOutline f, ax, ay, region, IdOf(el), backdrop, opacity
        Exit Sub
    End If
    If el.LocalName = "textarea" Then
        it.Kind = "textarea"
        it.Multiline = True
    Else
        it.Kind = "input"
    End If
    it.Text = el.Value
    it.Placeholder = el.Placeholder
    If it.Kind = "input" And Len(it.Placeholder) > 0 Then it.Placeholder = FitText(it.Placeholder, s, it.W - 8)
    it.ReadOnly = el.ReadOnly
    PushItem region, it
End Sub

' One line of text cut to a width with an ellipsis (a placeholder longer than its field;
' browsers clip it, but text measures a little wider here than in a browser).
Private Function FitText(ByVal text As String, ByVal s As Long, ByVal width As Double) As String
    Dim lo As Long, hi As Long, m As Long, ell As String, cut As String
    FitText = text
    If Engine.Measure(text, s) <= width Then Exit Function
    ell = ChrW$(&H2026)
    If Right$(text, 1) = ell Then text = Left$(text, Len(text) - 1)
    lo = 0: hi = Len(text)
    Do While lo < hi
        m = (lo + hi + 1) \ 2
        cut = RTrim$(Left$(text, m))
        If Engine.Measure(cut & ell, s) <= width Then lo = m Else hi = m - 1
    Loop
    cut = Left$(text, lo)
    Do While Len(cut) > 0 And InStr(1, " ,;:", Right$(cut, 1)) > 0
        cut = Left$(cut, Len(cut) - 1)
    Loop
    FitText = cut & ell
End Function

Private Sub PaintChevron(ByVal el As clsNode, ByVal s As Long, ByVal cx As Double, ByVal ay As Double, ByVal h As Double, ByVal fg As Long, _
        ByVal region As clsRenderSpec)
    Dim it As clsRenderSpec, t As String, w As Double, lh As Double
    t = ChrW$(&H25BE)
    w = Engine.Measure(t, s)
    lh = Engine.NaturalLineHeight(s)
    Set it = NewItem("text", IdOf(el) & ":chev", cx - w / 2, ay + (h - lh) / 2, w + 2, lh)
    FillText it, t, s, fg, el
    If mSymbolFallback Then it.Symbol = modRender.NeedsSymbolFont(t)
    PushItem region, it
End Sub

' Buttons are drawn like any other box (their face rect and text) and take clicks through
' those labels, or through a transparent hit label when nothing covers the face. A
' zero-size CommandButton (a "proxy": modRender gives it no area at the button's corner)
' keeps the button in the tab order and takes Enter and Space; snapshots keep the button's
' full geometry for it.
Private Sub CollapseButton(ByVal f As clsRenderBox, ByVal el As clsNode, ByVal ax As Double, ByVal ay As Double, ByVal region As clsRenderSpec, _
        ByVal startCount As Long, ByVal backdrop As Long, ByVal opacity As Double, ByVal tip As String)
    Dim i As Long, it As clsRenderSpec, s As Long, bg As Long, fg As Long, bc As Long, bw As Double, btn As clsRenderSpec
    Dim dimText As Boolean, faceId As String, covered As Boolean, label As String, ownTip As String, simple As Boolean, mid0 As Double
    Dim texts As Long, ta As String, d As String, prefix As String
    s = f.Slot
    prefix = IdOf(el) & ":"
    simple = True
    ControlColours f, backdrop, opacity, bg, fg, bc, bw
    ' A disabled button the stylesheet does not dim is greyed like a browser does.
    dimText = (el.Disabled And OpacityOf(s) >= 1)
    ownTip = tip
    If Len(ownTip) = 0 Then ownTip = AttrText(el, "aria-label")
    faceId = IdOf(el) & ":box"
    For i = startCount + 1 To region.Items.Count
        Set it = region.Items(i)
        If it.ClickNode Is Nothing Then Set it.ClickNode = el
        If Len(it.Tip) = 0 Then it.Tip = ownTip
        If it.Kind = "text" Then
            texts = texts + 1
            If texts = 1 Then
                mid0 = it.Y + it.H / 2
            ElseIf Abs(it.Y + it.H / 2 - mid0) >= 3 Then
                simple = False
            End If
            If dimText Then it.Color = Blend(it.Color, 0.45, bg)
            If Len(label) > 0 Then label = label & " "
            label = label & it.Text
        ElseIf it.Kind = "rect" And it.Id = faceId Then
            If it.X <= ax + 0.5 And it.Y <= ay + 0.5 And it.W >= f.W - 1 And it.H >= f.H - 1 Then covered = True
        ElseIf it.Kind = "rect" Then
            If Left$(it.Id, Len(prefix)) <> prefix Then simple = False
        Else
            simple = False
        End If
    Next i
    ' Content richer than one centred line (an icon, a title and a note) gets a hit label
    ' over it all, as the face labels may not cover the gaps.
    If simple Then
        ta = modStyle.SV(s, P_TEXTALIGN)
        d = modStyle.SV(s, P_DISPLAY)
        If ta <> "center" Then
            simple = False
            If InStr(1, d, "grid") > 0 Or InStr(1, d, "flex") > 0 Then
                simple = (InStr(1, modStyle.SV(s, P_JUSTIFYCONTENT), "center") > 0 Or InStr(1, modStyle.SV(s, P_JUSTIFYITEMS), "center") > 0)
            End If
        End If
    End If
    If Not covered Or Not simple Then
        Set it = NewItem("hit", IdOf(el) & ":hit", ax, ay, f.W, f.H)
        Set it.ClickNode = el
        it.Tip = ownTip
        PushItem region, it
    End If
    Set btn = NewItem("button", IdOf(el), ax, ay, f.W, f.H)
    Set btn.Node = el
    btn.Proxy = True
    label = Application.WorksheetFunction.Trim(label)
    If Len(label) = 0 Then label = AttrText(el, "aria-label")
    btn.Text = label
    btn.Size = Round(modStyle.SFont(s), 2)
    btn.Weight = modStyle.SWeight(s)
    btn.Color = fg
    btn.Disabled = el.Disabled
    btn.Tip = ownTip
    PushItem region, btn
    PaintOutline f, ax, ay, region, IdOf(el), backdrop, opacity
End Sub

' ================================================================ shadows, outlines, switches

' Box-shadow layers without blur: offset copies of the box (the title bar's document icon
' lines) and spread rings (focus rings). Blurred and inset layers are skipped.
Private Sub PaintShadows(ByVal f As clsRenderBox, ByVal ax As Double, ByVal ay As Double, ByVal region As clsRenderSpec, ByVal idBase As String, _
        ByVal backdrop As Long, ByVal opacity As Double)
    Dim t As String, layers() As String, toks() As String, i As Long, j As Long, n As Long, nums(0 To 3) As Double
    Dim colourText As String, rgbValue As Long, alpha As Double, colour As Long, it As clsRenderSpec, sp As Double, id As String, c As String
    If modStyle.SVI(f.Slot, P_BOXSHADOW) = mNoneIdx Then Exit Sub
    t = modStyle.SV(f.Slot, P_BOXSHADOW)
    If t = "none" Or Len(t) = 0 Then Exit Sub
    layers = Split(t, ",")
    For i = UBound(layers) To 0 Step -1
        toks = Split(Application.WorksheetFunction.Trim(layers(i)), " ")
        n = 0: colourText = ""
        For j = 0 To UBound(toks)
            c = LCase$(Left$(toks(j), 1))
            If LCase$(toks(j)) = "inset" Then
                n = -1
                Exit For
            ElseIf c = "#" Or (c >= "a" And c <= "z") Then
                colourText = toks(j)
            ElseIf n <= 3 Then
                nums(n) = Val(Replace(toks(j), "px", ""))
                n = n + 1
            End If
        Next j
        If n >= 2 And Len(colourText) > 0 Then
            If n < 3 Then nums(2) = 0
            If n < 4 Then nums(3) = 0
            If nums(2) = 0 Then
                If modStyle.ColourText(colourText, rgbValue, alpha) Then
                    colour = Blend(rgbValue, alpha * opacity, backdrop)
                    id = idBase & ":sh" & i
                    sp = nums(3)
                    If sp > 0 Then
                        PaintRing ax + nums(0), ay + nums(1), f.W, f.H, sp, sp, colour, region, id
                    ElseIf sp = 0 And (nums(0) <> 0 Or nums(1) <> 0) Then
                        Set it = NewItem("rect", id, ax + nums(0), ay + nums(1), f.W, f.H)
                        it.Bg = colour
                        PushItem region, it
                    End If
                End If
            End If
        End If
    Next i
End Sub

' Four rects of thickness th around a box, their outer edge gap px outside it.
Private Sub PaintRing(ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, ByVal gap As Double, ByVal th As Double, _
        ByVal colour As Long, ByVal region As clsRenderSpec, ByVal id As String)
    Dim it As clsRenderSpec, ox As Double, oy As Double, ow As Double, oh As Double
    ox = x - gap: oy = y - gap: ow = w + 2 * gap: oh = h + 2 * gap
    Set it = NewItem("rect", id & "t", ox, oy, ow, th): it.Bg = colour: PushItem region, it
    Set it = NewItem("rect", id & "b", ox, oy + oh - th, ow, th): it.Bg = colour: PushItem region, it
    Set it = NewItem("rect", id & "l", ox, oy + th, th, oh - 2 * th): it.Bg = colour: PushItem region, it
    Set it = NewItem("rect", id & "r", ox + ow - th, oy + th, th, oh - 2 * th): it.Bg = colour: PushItem region, it
End Sub

' The focused element's outline, when focus came from the keyboard (browsers show
' :focus-visible outlines only then).
Private Sub PaintOutline(ByVal f As clsRenderBox, ByVal ax As Double, ByVal ay As Double, ByVal region As clsRenderSpec, ByVal idBase As String, _
        ByVal backdrop As Long, ByVal opacity As Double)
    Dim s As Long, st As String, w As Double, off As Double, rgbValue As Long, alpha As Double
    If Not FocusVisible Then Exit Sub
    If Not f.Node Is mFocusNode Then Exit Sub
    s = f.Slot
    st = modStyle.SV(s, P_OUTLINES)
    If st = "none" Or st = "hidden" Or Len(st) = 0 Then Exit Sub
    If Not modStyle.SLen(s, P_OUTLINEW, 0, False, w) Then Exit Sub
    If w <= 0 Then Exit Sub
    If Not modStyle.SLen(s, P_OUTLINEOFF, 0, False, off) Then off = 0
    If Not modStyle.SColour(s, P_OUTLINEC, rgbValue, alpha) Then Exit Sub
    PaintRing ax, ay, f.W, f.H, off + w, w, Blend(rgbValue, alpha * opacity, backdrop), region, idBase & ":ol"
End Sub

' True for an appearance:none checkbox drawn as a switch (a pill wider than it is tall).
Private Function IsSwitch(ByVal f As clsRenderBox, ByVal el As clsNode) As Boolean
    If el.TypeName <> "checkbox" Then Exit Function
    IsSwitch = (f.W >= 24 And f.W > f.H * 1.3)
End Function

' A switch: its track (the input's own box) and its ::before knob, a transparent hit label
' that toggles it, and a zero-size CheckBox proxy that keeps Tab and Space.
Private Sub PaintSwitch(ByVal f As clsRenderBox, ByVal ax As Double, ByVal ay As Double, ByVal region As clsRenderSpec, ByVal backdrop As Long, _
        ByVal opacity As Double, ByVal tip As String, ByVal proxy As clsRenderSpec)
    Dim el As clsNode, s As Long, ps As Long, rgbValue As Long, alpha As Double, hasBg As Boolean, bg As Long, it As clsRenderSpec
    Dim kw As Double, kh As Double, mt As Double, ml As Double, tx As Double, t As String, p As Long, knobBg As Long, sz As Double
    Set el = f.Node
    s = f.Slot
    If modStyle.SColour(s, P_BG, rgbValue, alpha) Then hasBg = True: bg = Blend(rgbValue, alpha * opacity, backdrop)
    PaintShadows f, ax, ay, region, IdOf(el), backdrop, opacity
    PaintBox f, ax, ay, region, IdOf(el), hasBg, bg, backdrop, opacity
    If Not hasBg Then bg = backdrop
    ps = modStyle.PseudoOf(el, 1)
    If ps > 0 Then
        If modStyle.SV(ps, P_DISPLAY) <> "none" Then
            If modStyle.SLen(ps, P_WIDTH, 0, False, kw) And modStyle.SLen(ps, P_HEIGHT, 0, False, kh) Then
                If Not modStyle.SLen(ps, P_MT, 0, False, mt) Then mt = 0
                If Not modStyle.SLen(ps, P_ML, 0, False, ml) Then ml = 0
                t = modStyle.SV(ps, P_TRANSFORM)
                p = InStr(1, t, "translateX(", vbTextCompare)
                If p > 0 Then tx = Val(Mid$(t, p + 11))
                knobBg = RGB(255, 255, 255)
                If modStyle.SColour(ps, P_BG, rgbValue, alpha) Then knobBg = Blend(rgbValue, alpha * opacity, bg)
                If mSymbolFallback And Abs(kw - kh) < 0.5 Then
                    ' A round knob: the symbol font's black circle (its ink is 0.71 em wide,
                    ' 0.08 em from the left and 0.37 em from the top of a label).
                    sz = Round(kw / 0.71, 1)
                    Set it = NewItem("text", IdOf(el) & ":knob", ax + f.BL + f.PL + ml + tx - 0.08 * sz, ay + f.BT + f.PT + mt - 0.37 * sz, sz, sz * 1.15)
                    it.Text = ChrW$(&H25CF)
                    it.Size = sz
                    it.Weight = 400
                    it.Symbol = True
                    it.Color = knobBg
                    it.Align = "l"
                Else
                    Set it = NewItem("rect", IdOf(el) & ":knob", ax + f.BL + f.PL + ml + tx, ay + f.BT + f.PT + mt, kw, kh)
                    it.Bg = knobBg
                End If
                Set it.ClickNode = el
                it.Tip = tip
                PushItem region, it
            End If
        End If
    End If
    Set it = NewItem("hit", IdOf(el) & ":hit", ax, ay, f.W, f.H)
    Set it.ClickNode = el
    it.Tip = tip
    PushItem region, it
    proxy.Proxy = True
    PushItem region, proxy
    PaintOutline f, ax, ay, region, IdOf(el), backdrop, opacity
End Sub
