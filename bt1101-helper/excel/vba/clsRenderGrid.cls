VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRenderGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Grid and table layout for clsRenderLayout (port of the reference engine's layoutGrid,
' layoutTable and their intrinsic sizes): explicit column tracks (px, fr, auto, minmax,
' fit-content, repeat with counts and auto-fit/auto-fill), column spans, row-major auto
' placement, grid-template-rows (auto and fr), justify/align of items, and simple tables
' with border-collapse and percentage column widths.
Option Explicit

Private Type TTrack
    MinType As Long         ' 0 auto, 1 px, 2 fr
    MinV As Double
    MaxType As Long
    MaxV As Double
    Fit As Boolean
End Type

Public Engine As clsRenderLayout

Private Function MaxD(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxD = a Else MaxD = b
End Function

Private Function MinD(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinD = a Else MinD = b
End Function

' Splits on a separator (" " = any whitespace) outside parentheses.
Private Function SplitTop(ByVal text As String, ByVal sep As String, ByRef parts() As String) As Long
    Dim i As Long, depth As Long, c As String, start As Long, n As Long, piece As String, isSep As Boolean
    ReDim parts(0 To 7)
    start = 1
    For i = 1 To Len(text) + 1
        If i <= Len(text) Then c = Mid$(text, i, 1) Else c = ""
        If c = "(" Then depth = depth + 1
        If c = ")" Then depth = depth - 1
        If i > Len(text) Then
            isSep = True
        ElseIf depth = 0 Then
            If sep = " " Then isSep = (c = " " Or c = vbTab Or c = vbLf) Else isSep = (c = sep)
        Else
            isSep = False
        End If
        If isSep Then
            piece = Mid$(text, start, i - start)
            If sep <> " " Or Len(Trim$(piece)) > 0 Then
                If n > UBound(parts) Then ReDim Preserve parts(0 To n * 2)
                parts(n) = piece
                n = n + 1
            End If
            start = i + 1
        End If
    Next i
    SplitTop = n
End Function

Private Sub ParseBreadth(ByVal text As String, ByVal s As Long, ByVal cw As Double, ByRef typ As Long, ByRef v As Double)
    Dim t As String, px As Double
    t = LCase$(Trim$(text))
    typ = 0: v = 0
    If t = "auto" Or t = "min-content" Or t = "max-content" Then Exit Sub
    If Right$(t, 2) = "fr" Then typ = 2: v = Val(t): Exit Sub
    If modStyle.SLenText(t, s, cw, True, px) Then typ = 1: v = px
End Sub

Private Sub ParseTrack(ByVal text As String, ByVal s As Long, ByVal cw As Double, ByRef tr As TTrack)
    Dim t As String, args() As String, n As Long, px As Double
    t = Trim$(text)
    tr.Fit = False
    If LCase$(Left$(t, 7)) = "minmax(" And Right$(t, 1) = ")" Then
        n = SplitTop(Mid$(t, 8, Len(t) - 8), ",", args)
        ParseBreadth args(0), s, cw, tr.MinType, tr.MinV
        If n > 1 Then ParseBreadth args(1), s, cw, tr.MaxType, tr.MaxV Else tr.MaxType = tr.MinType: tr.MaxV = tr.MinV
        Exit Sub
    End If
    If LCase$(Left$(t, 12)) = "fit-content(" And Right$(t, 1) = ")" Then
        tr.MinType = 0: tr.MinV = 0
        tr.MaxType = 1
        If modStyle.SLenText(Mid$(t, 13, Len(t) - 13), s, cw, True, px) Then tr.MaxV = px Else tr.MaxV = 0
        tr.Fit = True
        Exit Sub
    End If
    ParseBreadth t, s, cw, tr.MinType, tr.MinV
    If tr.MinType = 2 Then
        tr.MaxType = 2: tr.MaxV = tr.MinV
        tr.MinType = 0: tr.MinV = 0
    Else
        tr.MaxType = tr.MinType: tr.MaxV = tr.MinV
    End If
End Sub

' Parses grid-template-columns. Returns the track count (-1 for none); isAuto is set for
' repeat(auto-fit/auto-fill, ...).
Private Function ParseTracks(ByVal value As String, ByVal s As Long, ByVal cw As Double, ByVal gap As Double, ByVal itemCount As Long, _
        ByVal measuring As Boolean, ByRef tracks() As TTrack, ByRef isAuto As Boolean) As Long
    Dim tokens() As String, tn As Long, i As Long, n As Long, inner As String, args() As String, an As Long, cnt As String
    Dim listText As String, list() As String, ln As Long, tr As TTrack, j As Long, r As Long, reps As Long, size As Double
    Dim listTracks() As TTrack
    isAuto = False
    value = Trim$(value)
    If value = "none" Or Len(value) = 0 Then ParseTracks = -1: Exit Function
    ReDim tracks(0 To 7)
    tn = SplitTop(value, " ", tokens)
    For i = 0 To tn - 1
        If LCase$(Left$(tokens(i), 7)) = "repeat(" And Right$(tokens(i), 1) = ")" Then
            inner = Mid$(tokens(i), 8, Len(tokens(i)) - 8)
            an = SplitTop(inner, ",", args)
            cnt = LCase$(Trim$(args(0)))
            listText = ""
            For j = 1 To an - 1
                If j > 1 Then listText = listText & ","
                listText = listText & args(j)
            Next j
            ln = SplitTop(Trim$(listText), " ", list)
            ReDim listTracks(0 To ln)
            For j = 0 To ln - 1
                ParseTrack list(j), s, cw, listTracks(j)
            Next j
            If cnt = "auto-fit" Or cnt = "auto-fill" Then
                isAuto = True
                If measuring Then
                    reps = 1
                    ln = 1
                Else
                    If listTracks(0).MinType = 1 Then
                        size = listTracks(0).MinV
                    ElseIf listTracks(0).MaxType = 1 Then
                        size = listTracks(0).MaxV
                    Else
                        size = 0
                    End If
                    If size > 0 Then reps = Int((cw + gap) / (size + gap)) Else reps = 1
                    If reps < 1 Then reps = 1
                    If cnt = "auto-fit" Then
                        If reps > itemCount Then reps = itemCount
                        If reps < 1 Then reps = 1
                    End If
                End If
            Else
                reps = Val(cnt)
                If reps < 1 Then reps = 1
            End If
            For r = 1 To reps
                For j = 0 To ln - 1
                    If n > UBound(tracks) Then ReDim Preserve tracks(0 To n * 2)
                    tracks(n) = listTracks(j)
                    n = n + 1
                Next j
            Next r
        Else
            ParseTrack tokens(i), s, cw, tr
            If n > UBound(tracks) Then ReDim Preserve tracks(0 To n * 2)
            tracks(n) = tr
            n = n + 1
        End If
    Next i
    ParseTracks = n
End Function

Private Sub GridSpan(ByVal s As Long, ByVal columns As Long, ByRef col As Long, ByRef span As Long)
    Dim st As String, en As String, a As Long, b As Long, bi As Long
    st = Trim$(modStyle.SV(s, P_GRIDCOLSTART))
    en = Trim$(modStyle.SV(s, P_GRIDCOLEND))
    col = -1
    span = 1
    If LCase$(Left$(st, 5)) = "span " Then span = Val(Mid$(st, 6))
    If LCase$(Left$(en, 5)) = "span " Then span = Val(Mid$(en, 6))
    If IsInteger(st) Then
        a = CLng(st)
        If a > 0 Then col = a - 1 Else col = columns + a
        If IsInteger(en) Then
            b = CLng(en)
            If b > 0 Then bi = b - 1 Else bi = columns + 1 + b
            span = bi - col
            If span < 1 Then span = 1
        End If
    End If
    If span > columns Then span = columns
    If span < 1 Then span = 1
    If col >= columns Then col = columns - 1
End Sub

Private Function IsInteger(ByVal t As String) As Boolean
    Dim i As Long, c As String
    If Len(t) = 0 Then Exit Function
    For i = 1 To Len(t)
        c = Mid$(t, i, 1)
        If Not (c >= "0" And c <= "9") And Not (i = 1 And c = "-" And Len(t) > 1) Then Exit Function
    Next i
    IsInteger = True
End Function

Private Function IsToggle(ByVal box As clsRenderBox) As Boolean
    Dim t As String
    If box Is Nothing Then Exit Function
    If box.Kind <> 1 Then Exit Function
    If box.Node.LocalName <> "input" Then Exit Function
    t = box.Node.TypeName
    IsToggle = (t = "checkbox" Or t = "radio")
End Function

' ================================================================ grid intrinsic sizes

Public Sub GridIntrinsic(ByVal box As clsRenderBox, ByRef fk() As clsRenderBox, ByVal m As Long, ByRef mn As Double, ByRef mx As Double)
    Dim s As Long, gap As Double, tracks() As TTrack, n As Long, isAuto As Boolean, i As Long, c As Long
    Dim mins() As Double, maxs() As Double, colMin() As Double, colMax() As Double, sumMax As Double
    s = box.Slot
    mn = 0: mx = 0
    If m = 0 Then Exit Sub
    ReDim mins(0 To m - 1): ReDim maxs(0 To m - 1)
    For i = 0 To m - 1
        Engine.OuterIntrinsic fk(i), mins(i), maxs(i)
        sumMax = sumMax + maxs(i)
    Next i
    If Not modStyle.SLen(s, P_COLGAP, 0, False, gap) Then gap = 0
    n = ParseTracks(modStyle.SV(s, P_GRIDCOLS), s, 0, gap, m, True, tracks, isAuto)
    If n <= 0 Or isAuto Then
        For i = 0 To m - 1
            mn = MaxD(mn, mins(i))
            If Not isAuto Then mx = MaxD(mx, maxs(i))
        Next i
        If isAuto Then mx = sumMax + gap * (m - 1)
        If mx < mn Then mx = mn
        Exit Sub
    End If
    ReDim colMin(0 To n - 1): ReDim colMax(0 To n - 1)
    For i = 0 To m - 1
        c = i Mod n
        colMin(c) = MaxD(colMin(c), mins(i))
        colMax(c) = MaxD(colMax(c), maxs(i))
    Next i
    For c = 0 To n - 1
        If tracks(c).MinType = 1 And tracks(c).MaxType = 1 Then
            colMin(c) = tracks(c).MinV
            colMax(c) = tracks(c).MinV
        ElseIf tracks(c).MinType = 1 Then
            colMin(c) = MaxD(colMin(c), tracks(c).MinV)
            colMax(c) = MaxD(colMax(c), tracks(c).MinV)
        End If
        mn = mn + colMin(c)
        mx = mx + colMax(c)
    Next c
    mn = mn + gap * (n - 1)
    mx = mx + gap * (n - 1)
End Sub

' ================================================================ grid layout

Public Function LayoutGrid(ByVal f As clsRenderBox, ByRef kids() As clsRenderBox, ByVal kn As Long, ByVal cw As Double, _
        ByVal innerDef As Double, ByVal innerMin As Double) As Double
    Dim NOV As Double, s As Long, colGap As Double, rowGap As Double, boxes() As clsRenderBox, isText() As Boolean, m As Long
    Dim i As Long, k As clsRenderBox, tracks() As TTrack, n As Long, isAuto As Boolean, implicit As Boolean
    Dim cellRow() As Long, cellCol() As Long, cellSpan() As Long, row As Long, col As Long, spanCol As Long, span As Long, rowCount As Long
    Dim sizes() As Double, limits() As Double, imn() As Double, imx() As Double, haveIntr() As Boolean
    s = f.Slot
    NOV = Engine.NoValue()
    If Not modStyle.SLen(s, P_COLGAP, cw, True, colGap) Then colGap = 0
    If Not modStyle.SLen(s, P_ROWGAP, innerDef, innerDef <> NOV, rowGap) Then rowGap = 0
    ReDim boxes(0 To kn): ReDim isText(0 To kn)
    For i = 0 To kn - 1
        Set k = kids(i)
        If Engine.IsAbsolute(k) Then
            f.AddAbsolute k
        ElseIf k.Kind = 3 Then
            If Len(Trim$(Replace(Replace(Replace(k.Text, vbLf, " "), vbTab, " "), vbCr, " "))) > 0 Then
                Set boxes(m) = k: isText(m) = True: m = m + 1
            End If
        Else
            Set boxes(m) = k: m = m + 1
        End If
    Next i
    If m = 0 Then Exit Function
    n = ParseTracks(modStyle.SV(s, P_GRIDCOLS), s, cw, colGap, m, False, tracks, isAuto)
    implicit = (n <= 0)
    If implicit Then
        n = 1
        ReDim tracks(0 To 0)
    End If
    ' Placement (row-major auto flow).
    ReDim cellRow(0 To m - 1): ReDim cellCol(0 To m - 1): ReDim cellSpan(0 To m - 1)
    For i = 0 To m - 1
        spanCol = -1: span = 1
        If Not isText(i) Then GridSpan boxes(i).Slot, n, spanCol, span
        If spanCol >= 0 Then
            If spanCol < col Then row = row + 1
            col = spanCol
        End If
        If col + span > n Then
            row = row + 1
            If spanCol >= 0 Then col = spanCol Else col = 0
        End If
        cellRow(i) = row: cellCol(i) = col: cellSpan(i) = span
        col = col + span
        If col >= n Then row = row + 1: col = 0
        If cellRow(i) + 1 > rowCount Then rowCount = cellRow(i) + 1
    Next i
    ' Column sizes.
    ReDim sizes(0 To n - 1): ReDim limits(0 To n - 1)
    ReDim imn(0 To m - 1): ReDim imx(0 To m - 1): ReDim haveIntr(0 To m - 1)
    If implicit Then
        sizes(0) = cw
        limits(0) = cw
    Else
        SizeColumns tracks, n, cw, colGap, boxes, m, cellCol, cellSpan, imn, imx, haveIntr, sizes, limits, modStyle.SV(s, P_JUSTIFYCONTENT)
    End If
    LayoutGrid = PlaceCells(f, boxes, isText, m, cellRow, cellCol, cellSpan, rowCount, sizes, n, colGap, rowGap, innerDef, innerMin)
End Function

Private Function ContentOf(ByVal c As Long, ByVal wantMax As Boolean, ByRef boxes() As clsRenderBox, ByVal m As Long, ByRef cellCol() As Long, _
        ByRef cellSpan() As Long, ByRef imn() As Double, ByRef imx() As Double, ByRef haveIntr() As Boolean) As Double
    Dim i As Long, v As Double
    For i = 0 To m - 1
        If cellCol(i) = c And cellSpan(i) = 1 Then
            If Not haveIntr(i) Then
                Engine.OuterIntrinsic boxes(i), imn(i), imx(i)
                haveIntr(i) = True
            End If
            If wantMax Then v = MaxD(v, imx(i)) Else v = MaxD(v, MinContribution(boxes(i), imn(i)))
        End If
    Next i
    ContentOf = v
End Function

' An item's minimum contribution to an auto track minimum: its min-width when one is set
' (min-width: 0 lets a 1fr track shrink below the content), else its min-content size.
Private Function MinContribution(ByVal box As clsRenderBox, ByVal minContent As Double) As Double
    Dim s As Long, v As Double
    MinContribution = minContent
    If box.Kind = 3 Then Exit Function
    s = box.Slot
    If modStyle.SV(s, P_MINW) = "auto" Then Exit Function
    If modStyle.SLen(s, P_MINW, 0, False, v) Then
        If v <= 0 Then MinContribution = 0
    End If
End Function

Private Sub SizeColumns(ByRef tracks() As TTrack, ByVal n As Long, ByVal cw As Double, ByVal colGap As Double, ByRef boxes() As clsRenderBox, _
        ByVal m As Long, ByRef cellCol() As Long, ByRef cellSpan() As Long, ByRef imn() As Double, ByRef imx() As Double, ByRef haveIntr() As Boolean, _
        ByRef sizes() As Double, ByRef limits() As Double, ByVal justify As String)
    Dim c As Long, gaps As Double, frCount As Long, flexible() As Boolean, round_ As Long, fixed As Double, frSum As Double, unit As Double
    Dim violators As Long, free As Double, growable As Long, share As Double, add As Double, autos As Long
    For c = 0 To n - 1
        Select Case tracks(c).MinType
            Case 1: sizes(c) = tracks(c).MinV
            Case 0: sizes(c) = ContentOf(c, False, boxes, m, cellCol, cellSpan, imn, imx, haveIntr)
            Case Else: sizes(c) = 0
        End Select
        Select Case tracks(c).MaxType
            Case 1: limits(c) = MaxD(sizes(c), tracks(c).MaxV)
            Case 0: limits(c) = MaxD(sizes(c), ContentOf(c, True, boxes, m, cellCol, cellSpan, imn, imx, haveIntr))
            Case Else: limits(c) = sizes(c)
        End Select
        If tracks(c).Fit Then limits(c) = MinD(ContentOf(c, True, boxes, m, cellCol, cellSpan, imn, imx, haveIntr), tracks(c).MaxV)
        If tracks(c).MaxType = 2 Then frCount = frCount + 1
    Next c
    gaps = colGap * (n - 1)
    If frCount > 0 Then
        ReDim flexible(0 To n - 1)
        For c = 0 To n - 1
            flexible(c) = (tracks(c).MaxType = 2)
        Next c
        For round_ = 0 To n
            fixed = 0: frSum = 0
            For c = 0 To n - 1
                If flexible(c) Then frSum = frSum + tracks(c).MaxV Else fixed = fixed + sizes(c)
            Next c
            If frSum = 0 Then frSum = 1
            unit = MaxD(0, (cw - gaps - fixed) / frSum)
            violators = 0
            For c = 0 To n - 1
                If flexible(c) Then
                    If tracks(c).MaxV * unit < sizes(c) Then
                        flexible(c) = False
                        violators = violators + 1
                    End If
                End If
            Next c
            If violators = 0 Then
                For c = 0 To n - 1
                    If flexible(c) Then sizes(c) = tracks(c).MaxV * unit
                Next c
                Exit For
            End If
            frCount = 0
            For c = 0 To n - 1
                If flexible(c) Then frCount = frCount + 1
            Next c
            If frCount = 0 Then Exit For
        Next round_
    Else
        free = cw - gaps
        For c = 0 To n - 1
            free = free - sizes(c)
        Next c
        For round_ = 0 To 2
            If free <= 0.01 Then Exit For
            growable = 0
            For c = 0 To n - 1
                If limits(c) > sizes(c) Then growable = growable + 1
            Next c
            If growable = 0 Then Exit For
            share = free / growable
            For c = 0 To n - 1
                If limits(c) > sizes(c) Then
                    add = MinD(share, limits(c) - sizes(c))
                    sizes(c) = sizes(c) + add
                    free = free - add
                End If
            Next c
        Next round_
        ' justify-content: normal stretches auto tracks into what is left.
        For c = 0 To n - 1
            If tracks(c).MaxType = 0 Then autos = autos + 1
        Next c
        If free > 0.01 And autos > 0 And (justify = "normal" Or justify = "stretch") Then
            For c = 0 To n - 1
                If tracks(c).MaxType = 0 Then sizes(c) = sizes(c) + free / autos
            Next c
        End If
    End If
End Sub

Private Function PlaceCells(ByVal f As clsRenderBox, ByRef boxes() As clsRenderBox, ByRef isText() As Boolean, ByVal m As Long, _
        ByRef cellRow() As Long, ByRef cellCol() As Long, ByRef cellSpan() As Long, ByVal rowCount As Long, ByRef sizes() As Double, _
        ByVal n As Long, ByVal colGap As Double, ByVal rowGap As Double, ByVal innerDef As Double, ByVal innerMin As Double) As Double
    Dim NOV As Double, s As Long, colX() As Double, acc As Double, c As Long, i As Long, width As Double, child As clsRenderBox
    Dim justifyItems As String, alignItems As String, js As String, e As New clsRenderBox, stretch As Boolean, v As Double, outer As Double
    Dim frags() As clsRenderBox, rowH() As Double, rowY() As Double, total As Double, target As Double, extra As Double
    Dim alignContent As String, r As Long, align As String, cs As Long
    NOV = Engine.NoValue()
    s = f.Slot
    ReDim colX(0 To n - 1)
    For c = 0 To n - 1
        colX(c) = acc
        acc = acc + sizes(c) + colGap
    Next c
    justifyItems = modStyle.SV(s, P_JUSTIFYITEMS)
    alignItems = modStyle.SV(s, P_ALIGNITEMS)
    ReDim frags(0 To m - 1)
    For i = 0 To m - 1
        width = colGap * (cellSpan(i) - 1)
        For c = cellCol(i) To cellCol(i) + cellSpan(i) - 1
            If c <= n - 1 Then width = width + sizes(c)
        Next c
        If isText(i) Then
            Set child = Engine.LayoutAnonText(boxes(i), f, width, "f" & i)
        Else
            cs = boxes(i).Slot
            js = modStyle.SV(cs, P_JUSTIFYSELF)
            If Len(js) = 0 Or js = "auto" Then js = justifyItems
            e.Kind = 1
            Engine.SetEdges e, cs, width
            stretch = (js = "normal" Or js = "stretch") And Not modStyle.SLen(cs, P_WIDTH, width, True, v) And Not IsToggle(boxes(i))
            If stretch Then
                Set child = Engine.LayoutBox(boxes(i), width, MaxD(0, width - e.ML - e.MR), NOV, False, True, NOV)
            Else
                Set child = Engine.LayoutBox(boxes(i), width, NOV, NOV, True, True, NOV)
            End If
            outer = child.W + e.ML + e.MR
            Select Case js
                Case "center": child.JustifyOffset = (width - outer) / 2
                Case "end", "right", "flex-end": child.JustifyOffset = width - outer
                Case Else: child.JustifyOffset = 0
            End Select
            If e.MLAuto And e.MRAuto Then child.JustifyOffset = (width - outer) / 2
        End If
        Set frags(i) = child
    Next i
    ' Rows.
    ReDim rowH(0 To rowCount - 1): ReDim rowY(0 To rowCount - 1)
    For i = 0 To m - 1
        rowH(cellRow(i)) = MaxD(rowH(cellRow(i)), frags(i).H + frags(i).MT + frags(i).MB)
    Next i
    For r = 0 To rowCount - 1
        total = total + rowH(r)
    Next r
    total = total + rowGap * MaxD(0, rowCount - 1)
    If innerDef <> NOV Then target = innerDef Else target = innerMin
    alignContent = modStyle.SV(s, P_ALIGNCONTENT)
    If target > total + 0.01 And (alignContent = "normal" Or alignContent = "stretch") And rowCount > 0 Then
        DistributeRows s, rowH, rowCount, target - total
    End If
    acc = 0
    For r = 0 To rowCount - 1
        rowY(r) = acc
        acc = acc + rowH(r) + rowGap
    Next r
    For i = 0 To m - 1
        Set child = frags(i)
        r = cellRow(i)
        If isText(i) Then
            align = "stretch"
        Else
            align = modStyle.SV(boxes(i).Slot, P_ALIGNSELF)
            If Len(align) = 0 Or align = "auto" Then align = alignItems
            If align = "normal" Then align = "stretch"
        End If
        outer = child.H + child.MT + child.MB
        child.X = colX(cellCol(i)) + child.JustifyOffset + child.ML
        If align = "stretch" And Not isText(i) And Not IsToggle(boxes(i)) And Not modStyle.SLen(boxes(i).Slot, P_HEIGHT, 0, False, v) Then
            child.H = MaxD(child.H, rowH(r) - child.MT - child.MB)
            child.Y = rowY(r) + child.MT
        ElseIf align = "center" Then
            child.Y = rowY(r) + (rowH(r) - outer) / 2 + child.MT
        ElseIf align = "end" Or align = "flex-end" Then
            child.Y = rowY(r) + rowH(r) - outer + child.MT
        Else
            child.Y = rowY(r) + child.MT
        End If
        f.AddKid child
    Next i
    PlaceCells = MaxD(0, acc - rowGap)
End Function

' Extra block space goes to fr rows of grid-template-rows, else equally to every row.
Private Sub DistributeRows(ByVal s As Long, ByRef rowH() As Double, ByVal rowCount As Long, ByVal extra As Double)
    Dim tokens() As String, tn As Long, r As Long, frSum As Double, frs() As Double, t As String
    ReDim frs(0 To rowCount - 1)
    t = modStyle.SV(s, P_GRIDROWS)
    If t <> "none" And Len(t) > 0 Then
        tn = SplitTop(t, " ", tokens)
        For r = 0 To rowCount - 1
            If r < tn Then
                If LCase$(Right$(Trim$(tokens(r)), 2)) = "fr" Then frs(r) = Val(Trim$(tokens(r)))
            End If
            frSum = frSum + frs(r)
        Next r
    End If
    For r = 0 To rowCount - 1
        If frSum > 0 Then
            rowH(r) = rowH(r) + extra * frs(r) / frSum
        Else
            rowH(r) = rowH(r) + extra / rowCount
        End If
    Next r
End Sub

' ================================================================ tables

Public Sub TableIntrinsic(ByVal box As clsRenderBox, ByRef mn As Double, ByRef mx As Double)
    Dim rows As New Collection, cells As Collection, cell As clsRenderBox, rowMin As Double, rowMax As Double, a As Double, b As Double
    Engine.TableRowsOf box.Node, rows
    mn = 0: mx = 0
    For Each cells In rows
        rowMin = 0: rowMax = 0
        For Each cell In cells
            Engine.Intrinsic cell, a, b
            rowMin = rowMin + a
            rowMax = rowMax + b
        Next cell
        mn = MaxD(mn, rowMin)
        mx = MaxD(mx, rowMax)
    Next cells
End Sub

Public Function LayoutTable(ByVal f As clsRenderBox, ByVal cw As Double) As Double
    Dim NOV As Double, s As Long, rows As New Collection, cells As Collection, cell As clsRenderBox, collapse As Boolean, spacing As Double
    Dim n As Long, c As Long, mins() As Double, maxs() As Double, pct() As Double, hasPct() As Boolean, a As Double, b As Double, wt As String
    Dim available As Double, widths() As Double, fixedTotal As Double, rest As Double, sumMax As Double, sumMin As Double, flexCount As Long
    Dim y As Double, x As Double, rowH As Double, child As clsRenderBox, frags() As clsRenderBox, i As Long, firstBB As Double
    NOV = Engine.NoValue()
    s = f.Slot
    Engine.TableRowsOf f.Node, rows
    If rows.Count = 0 Then Exit Function
    collapse = (modStyle.SV(s, P_BORDERCOLLAPSE) = "collapse")
    If Not collapse Then
        If Not modStyle.SLen(s, P_BORDERSPACING, 0, False, spacing) Then spacing = 2
    End If
    For Each cells In rows
        If cells.Count > n Then n = cells.Count
    Next cells
    If n = 0 Then Exit Function
    ReDim mins(0 To n - 1): ReDim maxs(0 To n - 1): ReDim pct(0 To n - 1): ReDim hasPct(0 To n - 1): ReDim widths(0 To n - 1)
    For Each cells In rows
        c = 0
        For Each cell In cells
            Engine.Intrinsic cell, a, b
            mins(c) = MaxD(mins(c), a)
            maxs(c) = MaxD(maxs(c), b)
            wt = modStyle.SV(cell.Slot, P_WIDTH)
            If Right$(wt, 1) = "%" Then pct(c) = Val(wt): hasPct(c) = True
            c = c + 1
        Next cell
    Next cells
    available = cw - spacing * (n + 1)
    For c = 0 To n - 1
        If hasPct(c) Then
            widths(c) = MaxD(mins(c), available * pct(c) / 100)
            fixedTotal = fixedTotal + widths(c)
        Else
            sumMax = sumMax + maxs(c)
            sumMin = sumMin + mins(c)
            flexCount = flexCount + 1
        End If
    Next c
    rest = available - fixedTotal
    For c = 0 To n - 1
        If Not hasPct(c) Then
            If sumMax <= rest Then
                If sumMax > 0 Then widths(c) = maxs(c) + (rest - sumMax) * maxs(c) / sumMax Else widths(c) = rest / flexCount
            ElseIf sumMax - sumMin > 0 Then
                widths(c) = mins(c) + MaxD(0, rest - sumMin) * (maxs(c) - mins(c)) / (sumMax - sumMin)
            Else
                widths(c) = mins(c)
            End If
        End If
    Next c
    y = spacing
    For Each cells In rows
        x = spacing
        rowH = 0
        ReDim frags(0 To cells.Count)
        i = 0
        For Each cell In cells
            Set child = Engine.LayoutBox(cell, widths(i), widths(i), NOV, False, False, NOV)
            child.X = x
            x = x + widths(i) + spacing
            If collapse Then x = x - child.BR
            If child.H > rowH Then rowH = child.H
            Set frags(i) = child
            i = i + 1
        Next cell
        firstBB = 0
        For c = 0 To i - 1
            frags(c).Y = y
            frags(c).H = rowH
            f.AddKid frags(c)
            If c = 0 Then firstBB = frags(c).BB
        Next c
        y = y + rowH + spacing
        If collapse And i > 0 Then y = y - firstBB
    Next cells
    LayoutTable = y
End Function
