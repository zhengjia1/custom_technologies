VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRenderSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' One record of the display list the paint pass produces: a region (a frame: title bar,
' library, page, nested scroller, overlay) or an item drawn with MSForms controls inside a
' region. Coordinates are CSS px relative to the region's content origin (regions: to
' their parent region's content origin, or the window). Colours are VBA RGB Longs, -1 =
' none.
Option Explicit

Public Kind As String           ' region, rect, text, hit, input, textarea, select, button, check
Public Id As String
Public Region As String         ' items: region id; regions: parent region id ("" = window)
Public X As Double, Y As Double, W As Double, H As Double
Public Text As String           ' text/button caption, input value
Public Placeholder As String
Public Size As Double           ' font size, px
Public Weight As Long           ' 400/600/700
Public Italic As Boolean
Public Symbol As Boolean        ' draw with the symbol font (glyphs the text font lacks)
Public Color As Long
Public Bg As Long
Public Bc As Long
Public Bw As Double
Public Align As String          ' l/c/r
Public Node As clsNode          ' controls: the element that receives input events
Public ClickNode As clsNode     ' text/hit/rect: the element a click activates
Public Tip As String
Public Disabled As Boolean
Public ReadOnly As Boolean
Public Checked As Boolean
Public IsRadio As Boolean
Public Options As String        ' select: option captions joined by vbLf
Public OptionFlags As String    ' select: "1" per disabled option, "0" otherwise
Public SelIndex As Long
Public Multiline As Boolean
Public Resizer As Boolean       ' hit: the sidebar resizer
Public Proxy As Boolean         ' button/check: a zero-size focus proxy (the element is drawn)
Public StickyOf As String       ' overlay region: copy of a sticky element in this scroller region
Public StickyY As Double        '   shown once the scroller is scrolled past this offset (CSS px)
' regions
Public Scrolls As Boolean
Public ScrollsX As Boolean
Public ContentW As Double
Public ContentH As Double
Public Overlay As Boolean
Public PaintIndex As Long
Public OX As Double, OY As Double   ' regions: absolute content origin (paint bookkeeping)
Public Items As Collection          ' regions: items in paint order (paint bookkeeping)

Private Sub Class_Initialize()
    Bg = -1
    Bc = -1
    Color = 0
    SelIndex = -1
End Sub

' Everything that affects how the item looks except its position.
Public Function Look() As String
    Look = Kind & "|" & Text & "|" & Placeholder & "|" & Size & "|" & Weight & "|" & Italic & "|" & Symbol & "|" & Color & "|" & _
        Bg & "|" & Bc & "|" & Bw & "|" & Align & "|" & Tip & "|" & Disabled & "|" & ReadOnly & "|" & Checked & "|" & _
        IsRadio & "|" & Proxy & "|" & OptionFlags & "|" & SelIndex & "|" & Multiline & "|" & Len(Options) & "|" & Options
End Function

' True when o looks the same (everything Look() lists).
Public Function SameLook(ByVal o As clsRenderSpec) As Boolean
    If Kind <> o.Kind Then Exit Function
    If Text <> o.Text Then Exit Function
    If Color <> o.Color Or Bg <> o.Bg Or Bc <> o.Bc Or Bw <> o.Bw Then Exit Function
    If Size <> o.Size Or Weight <> o.Weight Or Italic <> o.Italic Or Symbol <> o.Symbol Then Exit Function
    If Placeholder <> o.Placeholder Or Align <> o.Align Or Tip <> o.Tip Then Exit Function
    If Disabled <> o.Disabled Or ReadOnly <> o.ReadOnly Or Checked <> o.Checked Or IsRadio <> o.IsRadio Or Proxy <> o.Proxy Then Exit Function
    If SelIndex <> o.SelIndex Or Multiline <> o.Multiline Or OptionFlags <> o.OptionFlags Or Options <> o.Options Then Exit Function
    SameLook = True
End Function

' True when o has the same Geometry().
Public Function SameGeometry(ByVal o As clsRenderSpec) As Boolean
    If Proxy <> o.Proxy Then Exit Function
    If Round(X) <> Round(o.X) Or Round(Y) <> Round(o.Y) Then Exit Function
    SameGeometry = (Round(W) = Round(o.W) And Round(H) = Round(o.H))
End Function

' Position and size, rounded to whole px.
Public Function Geometry() As String
    Geometry = Round(X) & "," & Round(Y) & "," & Round(W) & "," & Round(H)
    If Proxy Then Geometry = Geometry & "p"
End Function
