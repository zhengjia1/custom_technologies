' Opens the app window when the workbook opens (macros must be enabled).
Option Explicit

Private Sub Workbook_Open()
    ' Apostrophes in the file name are doubled inside the quoted workbook reference.
    Application.OnTime Now, "'" & Replace(Me.Name, "'", "''") & "'!modBoot.StartApp"
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    modBoot.StopApp
End Sub
