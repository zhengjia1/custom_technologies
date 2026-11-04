VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRenderItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' One drawn item of the app window: its MSForms control(s), the display-list record it
' was last drawn from, and the event sink that reports what the user does to modRender
' (which turns it into DOM events).
Option Explicit

Public Id As String
Public Kind As String
Public RegionId As String
Public Spec As clsRenderSpec        ' the record last applied
Public LookKey As String            ' "ok" when Spec's look was applied completely
Public GeomKey As String            ' "ok" when Spec's geometry was applied
Public OptionsKey As String         ' select: options last loaded
Public Ctl As Object                ' the main control
Public Hint As Object               ' input/textarea: placeholder label (beneath the text box)
Public Container As Object          ' Controls collection holding the controls

Public WithEvents Lbl As MSForms.Label
Attribute Lbl.VB_VarHelpID = -1
Public WithEvents Txt As MSForms.TextBox
Attribute Txt.VB_VarHelpID = -1
Public WithEvents Btn As MSForms.CommandButton
Attribute Btn.VB_VarHelpID = -1
Public WithEvents Cmb As MSForms.ComboBox
Attribute Cmb.VB_VarHelpID = -1
Public WithEvents Chk As MSForms.CheckBox
Attribute Chk.VB_VarHelpID = -1
Public WithEvents Opt As MSForms.OptionButton
Attribute Opt.VB_VarHelpID = -1
Public WithEvents Frm As MSForms.Frame          ' a scroller frame whose scrolling moves sticky copies
Attribute Frm.VB_VarHelpID = -1

Public Sub Dispose()
    Set Lbl = Nothing: Set Txt = Nothing: Set Btn = Nothing
    Set Cmb = Nothing: Set Chk = Nothing: Set Opt = Nothing: Set Frm = Nothing
    Set Ctl = Nothing: Set Hint = Nothing: Set Container = Nothing
    Set Spec = Nothing
End Sub

' ---------------------------------------------------------------- labels (text, rect, hit)

Private Sub Lbl_Click()
    modRender.OnItemClick Me
End Sub

Private Sub Lbl_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    modRender.OnItemDblClick Me
End Sub

Private Sub Lbl_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    modRender.OnItemMouse Me, 0, Button, X, Y
End Sub

Private Sub Lbl_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button <> 0 Then modRender.OnItemMouse Me, 1, Button, X, Y
End Sub

Private Sub Lbl_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    modRender.OnItemMouse Me, 2, Button, X, Y
End Sub

' ---------------------------------------------------------------- text boxes

Private Sub Txt_Change()
    modRender.OnItemText Me
End Sub

Private Sub Txt_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If modRender.OnItemKeyDown(Me, KeyCode, Shift) Then KeyCode = 0
End Sub

Private Sub Txt_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    modRender.OnItemKeyUp Me
End Sub

Private Sub Txt_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    modRender.OnItemMouseFocus Me
End Sub

' ---------------------------------------------------------------- buttons

Private Sub Btn_Click()
    modRender.OnItemClick Me
End Sub

Private Sub Btn_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If modRender.OnItemKeyDown(Me, KeyCode, Shift) Then KeyCode = 0
End Sub

Private Sub Btn_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    modRender.OnItemKeyUp Me
End Sub

' ---------------------------------------------------------------- lists

Private Sub Cmb_Change()
    modRender.OnItemSelect Me
End Sub

Private Sub Cmb_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If modRender.OnItemKeyDown(Me, KeyCode, Shift) Then KeyCode = 0
End Sub

Private Sub Cmb_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    modRender.OnItemKeyUp Me
End Sub

Private Sub Cmb_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    modRender.OnItemMouseFocus Me
End Sub

Private Sub Cmb_DropButtonClick()
    modRender.OnItemFocusCheck Me
End Sub

' ---------------------------------------------------------------- toggles

Private Sub Chk_Click()
    modRender.OnItemClick Me
End Sub

Private Sub Chk_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If modRender.OnItemKeyDown(Me, KeyCode, Shift) Then KeyCode = 0
End Sub

Private Sub Chk_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    modRender.OnItemKeyUp Me
End Sub

Private Sub Opt_Click()
    modRender.OnItemClick Me
End Sub

Private Sub Opt_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If modRender.OnItemKeyDown(Me, KeyCode, Shift) Then KeyCode = 0
End Sub

Private Sub Opt_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    modRender.OnItemKeyUp Me
End Sub

' ---------------------------------------------------------------- scroller frames

Private Sub Frm_Scroll(ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, ByVal RequestDx As Single, _
        ByVal RequestDy As Single, ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle)
    modRender.OnFrameScroll Me, ActualDy
End Sub
