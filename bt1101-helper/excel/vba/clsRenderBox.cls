VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRenderBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' A box for the layout engine (clsRenderLayout): before layout it describes what to lay
' out (an element, a ::before/::after box, a text node, or an anonymous line box); after
' layout it is a fragment with a position (relative to its parent fragment's border box),
' a size, its edges, child fragments and, for line boxes, text runs. CSS px throughout.
Option Explicit

Public Kind As Long             ' 1 element, 2 pseudo-element, 3 text, 4 anonymous line box
Public Node As clsNode          ' the element (the owner element for text, pseudo and line boxes)
Public Slot As Long             ' modStyle slot
Public Which As Long            ' pseudo-element: 1 ::before, 2 ::after
Public Text As String           ' text node data / pseudo-element content
Public Display As String
Public Source As clsRenderBox   ' fragment: the box it lays out

Public X As Double, Y As Double, W As Double, H As Double, CW As Double
Public PT As Double, PR As Double, PB As Double, PL As Double
Public BT As Double, BR As Double, BB As Double, BL As Double
Public MT As Double, MR As Double, MB As Double, ML As Double
Public MTAuto As Boolean, MRAuto As Boolean, MBAuto As Boolean, MLAuto As Boolean

Public Role As Long             ' 0 none, 1 title bar, 2 library, 3 main, 4 workspace, 5 resizer, 6 fixed
Public Scrolls As Boolean       ' content taller than the box (overflow auto/scroll or a role region)
Public ScrollH As Double
Public ScrollsX As Boolean
Public ScrollW As Double
Public ContentH As Double
Public Positioned As Boolean
Public HasLines As Boolean
Public LinesX As Double, LinesY As Double
Public AnonIndex As String
Public Owner As clsRenderBox    ' line box: the block fragment it belongs to
Public JustifyOffset As Double
Public AutoTop As Double, AutoBottom As Double, AutoLeft As Double, AutoRight As Double
Public ClickNode As clsNode     ' inline atom: the label/summary a click on it activates
Public HasLegend As Boolean     ' fieldset: its legend sits on the top border, which is drawn at
Public LegendLineY As Double    '   this offset and left out between LegendX0 and LegendX1
Public LegendX0 As Double, LegendX1 As Double

Private mKids() As clsRenderBox
Private mKidCount As Long
Private mAbs() As clsRenderBox
Private mAbsCount As Long

Private mRunCount As Long
Private mRunX() As Double, mRunY() As Double, mRunW() As Double, mRunH() As Double
Private mRunText() As String, mRunSlot() As Long, mRunLine() As Long, mRunIndex() As Long
Private mRunClick() As clsNode

' ---------------------------------------------------------------- child fragments

Public Sub AddKid(ByVal kid As clsRenderBox)
    If mKidCount = 0 Then
        ReDim mKids(0 To 7)
    ElseIf mKidCount > UBound(mKids) Then
        ReDim Preserve mKids(0 To mKidCount * 2)
    End If
    Set mKids(mKidCount) = kid
    mKidCount = mKidCount + 1
End Sub

Public Property Get KidCount() As Long
    KidCount = mKidCount
End Property

Public Function Kid(ByVal i As Long) As clsRenderBox
    Set Kid = mKids(i)
End Function

Public Sub ClearKids()
    mKidCount = 0
    Erase mKids
    mAbsCount = 0
    Erase mAbs
    ClearRuns
End Sub

' ---------------------------------------------------------------- absolutely positioned boxes

Public Sub AddAbsolute(ByVal box As clsRenderBox)
    If mAbsCount = 0 Then
        ReDim mAbs(0 To 3)
    ElseIf mAbsCount > UBound(mAbs) Then
        ReDim Preserve mAbs(0 To mAbsCount * 2)
    End If
    Set mAbs(mAbsCount) = box
    mAbsCount = mAbsCount + 1
End Sub

Public Property Get AbsCount() As Long
    AbsCount = mAbsCount
End Property

Public Function AbsBox(ByVal i As Long) As clsRenderBox
    Set AbsBox = mAbs(i)
End Function

' ---------------------------------------------------------------- text runs

Public Sub AddRun(ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, ByVal text As String, _
        ByVal slot As Long, ByVal click As clsNode, ByVal lineIndex As Long, ByVal runIndex As Long)
    If mRunCount = 0 Then
        ReDim mRunX(0 To 7): ReDim mRunY(0 To 7): ReDim mRunW(0 To 7): ReDim mRunH(0 To 7)
        ReDim mRunText(0 To 7): ReDim mRunSlot(0 To 7): ReDim mRunClick(0 To 7)
        ReDim mRunLine(0 To 7): ReDim mRunIndex(0 To 7)
    ElseIf mRunCount > UBound(mRunX) Then
        Dim n As Long
        n = mRunCount * 2
        ReDim Preserve mRunX(0 To n): ReDim Preserve mRunY(0 To n): ReDim Preserve mRunW(0 To n): ReDim Preserve mRunH(0 To n)
        ReDim Preserve mRunText(0 To n): ReDim Preserve mRunSlot(0 To n): ReDim Preserve mRunClick(0 To n)
        ReDim Preserve mRunLine(0 To n): ReDim Preserve mRunIndex(0 To n)
    End If
    mRunX(mRunCount) = x: mRunY(mRunCount) = y: mRunW(mRunCount) = w: mRunH(mRunCount) = h
    mRunText(mRunCount) = text: mRunSlot(mRunCount) = slot
    Set mRunClick(mRunCount) = click
    mRunLine(mRunCount) = lineIndex: mRunIndex(mRunCount) = runIndex
    mRunCount = mRunCount + 1
    HasLines = True
End Sub

Public Sub ClearRuns()
    mRunCount = 0
    HasLines = False
End Sub

Public Property Get RunCount() As Long
    RunCount = mRunCount
End Property

Public Function RunX(ByVal i As Long) As Double
    RunX = mRunX(i)
End Function

Public Function RunY(ByVal i As Long) As Double
    RunY = mRunY(i)
End Function

Public Function RunW(ByVal i As Long) As Double
    RunW = mRunW(i)
End Function

Public Function RunH(ByVal i As Long) As Double
    RunH = mRunH(i)
End Function

Public Function RunText(ByVal i As Long) As String
    RunText = mRunText(i)
End Function

Public Function RunSlot(ByVal i As Long) As Long
    RunSlot = mRunSlot(i)
End Function

Public Function RunClick(ByVal i As Long) As clsNode
    Set RunClick = mRunClick(i)
End Function

Public Function RunLine(ByVal i As Long) As Long
    RunLine = mRunLine(i)
End Function

Public Function RunIndex(ByVal i As Long) As Long
    RunIndex = mRunIndex(i)
End Function

' ---------------------------------------------------------------- copies

' A fragment for this box (layout fills in geometry).
Public Function NewFragment() As clsRenderBox
    Dim f As New clsRenderBox
    f.Kind = Kind
    Set f.Node = Node
    f.Slot = Slot
    f.Which = Which
    f.Text = Text
    Set f.Source = Me
    Set NewFragment = f
End Function

Public Sub CopyEdgesFrom(ByVal other As clsRenderBox)
    PT = other.PT: PR = other.PR: PB = other.PB: PL = other.PL
    BT = other.BT: BR = other.BR: BB = other.BB: BL = other.BL
    MT = other.MT: MR = other.MR: MB = other.MB: ML = other.ML
    MTAuto = other.MTAuto: MRAuto = other.MRAuto: MBAuto = other.MBAuto: MLAuto = other.MLAuto
End Sub
