Attribute VB_Name = "testOpLp"
' Test helpers for the lpSolve operation (opLp, opLp2). The differential step scripts in
' test/steps/opLp run against the full shell (default boot InitApp), so no separate boot
' is needed; this module only offers direct checks of the model-text parser.
Option Explicit

' Parses model text (the smart parser's input) and returns a one-line summary such as
' "ok max X1,X2 [3,5] continuous | A: 1,1 <= 4" or "error <message>".
Public Function ParseSummary(ByVal text As String) As String
    Dim model As clsMap, constraints As clsList, i As Long, constraint As clsMap, parts As clsList
    Set model = opLp.ParseLpModelText(text)
    If Not CBool(model("ok")) Then
        ParseSummary = "error " & model("message")
        Exit Function
    End If
    Set parts = NewList()
    Set constraints = model("constraints")
    For i = 0 To constraints.Count - 1
        Set constraint = constraints(i)
        parts.Push constraint("name") & ": " & constraint("coefficients").Join(",") & " " & constraint("operator") & " " & JsNumberToString(CDbl(constraint("rhs")))
    Next i
    ParseSummary = "ok " & model("direction") & " " & model("variables").Join(",") & " [" & model("objective").Join(",") & "] " & _
        model("type") & " " & model("binaryPositions").Join(",") & " | " & parts.Join(" ; ")
End Function

' Runs every steps file named in the list file (like testDriver.RunStepsList) and appends
' the errors that event handlers reported through ReportError (they never surface as
' differing records on their own).
Public Function RunStepsWithErrors(ByVal listPath As String) As String
    Dim paths() As String, i As Long, out As String
    paths = Split(Replace(ReadUtf8File(listPath), vbCrLf, vbLf), vbLf)
    For i = 0 To UBound(paths)
        If Len(Trim$(paths(i))) > 0 Then
            LastErrors = ""
            out = out & vbLf & "=====FILE " & Trim$(paths(i)) & vbLf & Left$(testDriver.RunStepsFile(Trim$(paths(i))), 200) & vbLf & "ERRORS:" & vbLf & LastErrors
        End If
    Next i
    RunStepsWithErrors = out
End Function
