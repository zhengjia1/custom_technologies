Attribute VB_Name = "modBoot"
' Opens and closes BT1101 Prose Lab. The website's markup, the stylesheet rules and the
' fonts are stored by the build on the very hidden "BTX" sheet: row 1 names each item,
' and the rows below hold its text in chunks.
Option Explicit

Private mRunning As Boolean

Public Sub StartApp()
    On Error GoTo Failed
    If mRunning Then
        CallOp "modRender.OpenAppWindow"
        Exit Sub
    End If
    LoadDocument StoredText("body")
    CallOp "InitApp"
    Settle 700
    mRunning = True
    CallOp "modRender.OpenAppWindow"
    Exit Sub
Failed:
    mRunning = False
    MsgBox "BT1101 Prose Lab could not start." & vbCrLf & vbCrLf & Err.Description, vbCritical, "BT1101 Prose Lab"
End Sub

Public Sub StopApp()
    On Error Resume Next
    If Not mRunning Then Exit Sub
    mRunning = False
    CallOp "modRender.CloseAppWindow"
    ResetDocument
End Sub

Public Function AppRunning() As Boolean
    AppRunning = mRunning
End Function

' Text stored by the build under a name in row 1 of the "BTX" sheet ("" when absent).
Public Function StoredText(ByVal name As String) As String
    Dim sheet As Worksheet, column As Long, last As Long, values As Variant, parts() As String, i As Long
    Set sheet = ThisWorkbook.Worksheets("BTX")
    For column = 1 To sheet.Cells(1, sheet.Columns.Count).End(xlToLeft).Column
        If CStr(sheet.Cells(1, column).Value2) = name Then
            last = sheet.Cells(sheet.Rows.Count, column).End(xlUp).Row
            If last < 2 Then Exit Function
            values = sheet.Range(sheet.Cells(2, column), sheet.Cells(last, column)).Value2
            If Not IsArray(values) Then
                StoredText = CStr(values)
                Exit Function
            End If
            ReDim parts(1 To last - 1)
            For i = 1 To last - 1
                parts(i) = CStr(values(i, 1))
            Next i
            StoredText = Join(parts, "")
            Exit Function
        End If
    Next column
End Function
