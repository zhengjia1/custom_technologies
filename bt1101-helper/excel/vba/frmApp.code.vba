' The app window. Everything it shows is created at run time by modRender from the DOM;
' this code only forwards the window's own events.
Option Explicit

Private Sub UserForm_Resize()
    modRender.FormResized
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    modRender.FormClosing
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If modRender.OnFormKeyDown(KeyCode, Shift) Then KeyCode = 0
End Sub
