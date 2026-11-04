Attribute VB_Name = "testOpStats"
' Test-only helpers for the opStats / opAnova step scripts (test/steps/opStats,
' test/steps/opAnova). The scripts run with the shell's InitApp boot; this adds a check
' that no event handler reported an error while they ran (handler errors are trapped and
' logged by modShared.ReportError, so they never show in the step records themselves).
Option Explicit

' Runs testDriver.RunStepsList and appends the handler errors logged during the run.
Public Function RunStepsListWithErrors(ByVal listPath As String) As String
    Dim out As String
    LastErrors = ""
    out = RunStepsList(listPath)
    RunStepsListWithErrors = out & vbLf & "=====ERRORS=====" & vbLf & LastErrors
End Function
