Attribute VB_Name = "testVariants"
' Test-only: engine-equivalence runner for test/variants/equiv.js and the boot procedure
' for the answer-option step scripts (test/steps/modVariants).
Option Explicit

' Reads cases (one per line: seed, kind, escaped canonical, history count (-1 = none),
' escaped history entries; tab-separated), writes the six escaped outputs of each case to
' <path>.out and returns a summary.
Public Function RunEngineCases(ByVal path As String) As String
    Dim lines() As String, i As Long, parts() As String, out As clsList, history As clsList, n As Long, j As Long
    Dim variants As clsList, started As Double, line As String, errors As Long
    started = Timer
    lines = Split(Replace(ReadUtf8File(path), vbCrLf, vbLf), vbLf)
    Set out = NewList()
    For i = 0 To UBound(lines)
        line = lines(i)
        If Len(line) > 0 Then
            parts = Split(line, vbTab)
            Set history = Nothing
            n = CLng(parts(3))
            If n >= 0 Then
                Set history = NewList()
                For j = 1 To n
                    history.Push Unesc(parts(3 + j))
                Next j
            End If
            On Error GoTo CaseFailed
            JsSeedRandom CDbl(parts(0))
            If parts(1) = "code" Then
                Set variants = CodeAnswerVariants(Unesc(parts(2)), history)
            Else
                Set variants = ProseAnswerVariants(Unesc(parts(2)), history)
            End If
            For j = 0 To variants.Count - 1
                variants(j) = Esc(CStr(variants(j)))
            Next j
            out.Push variants.Join(vbTab)
            On Error GoTo 0
        End If
NextCase:
    Next i
    WriteUtf8File path & ".out", out.Join(vbLf)
    RunEngineCases = "cases " & out.Count & ", errors " & errors & ", " & Format$(Timer - started, "0.0") & " s"
    Exit Function
CaseFailed:
    errors = errors + 1
    out.Push "ERROR " & Err.Number & " " & Err.Description
    Resume NextCase
End Function

Private Function Esc(ByVal text As String) As String
    Esc = Replace(Replace(Replace(Replace(text, "\", "\\"), vbLf, "\n"), vbTab, "\t"), vbCr, "\r")
End Function

Private Function Unesc(ByVal text As String) As String
    Dim i As Long, ch As String, buffer As String, n As Long
    If InStr(1, text, "\", vbBinaryCompare) = 0 Then Unesc = text: Exit Function
    buffer = String$(Len(text), " ")
    i = 1
    Do While i <= Len(text)
        ch = Mid$(text, i, 1)
        If ch = "\" And i < Len(text) Then
            Select Case Mid$(text, i + 1, 1)
                Case "n": ch = vbLf
                Case "t": ch = vbTab
                Case "r": ch = vbCr
                Case Else: ch = Mid$(text, i + 1, 1)
            End Select
            i = i + 2
        Else
            i = i + 1
        End If
        n = n + 1
        Mid$(buffer, n, 1) = ch
    Loop
    Unesc = Left$(buffer, n)
End Function

' Math.random().toString(36) for 3000 values of a seeded mulberry32 stream, space-separated.
Public Function Base36Cases(ByVal seedText As String) As String
    Dim i As Long, out As clsList
    Set out = NewList()
    JsSeedRandom CDbl(seedText)
    For i = 1 To 3000
        out.Push VarRandomToBase36(JsRandom())
    Next i
    Base36Cases = out.Join(" ")
End Function

' Boot for the step scripts that exercises modVariants' own document listeners and
' mutation hook: the full startup, with the shell's copies of those listeners replaced.
Public Sub BootOwnListeners()
    Application.Run "InitApp"
    Doc.RemoveEventListener "click", "modApp2.OnDocClickAnswerVariants", True
    Doc.RemoveEventListener "paste", "modApp2.OnDocPaste", True
    Doc.RemoveEventListener "input", "modApp2.OnDocInputAnswerVariants"
    Doc.RemoveEventListener "change", "modApp2.OnDocChangeAnswerVariants"
    InstallAnswerVariantListeners
End Sub

' Stress: re-renders the PCA workflow (new output nodes each time) and the K-fold code,
' then reports the cache size, timing and any handler errors.
Public Function StressRender(ByVal bodyPath As String) As String
    Dim i As Long, started As Double, perEdit As Double, el As clsNode, report As String
    JsSeedRandom 1
    LoadDocument ReadUtf8File(bodyPath)
    Application.Run "InitApp"
    Settle 700
    Q("#tutorialList .operation[data-id=""pca""]").Click
    Settle 700
    Q("#loadPcaSampleBtn").Click
    Settle 700
    started = Timer
    For i = 1 To 150
        Set el = Q("#pcaDataset")
        el.Value = IIf(i Mod 2 = 0, "USArrests", "crime_data")
        el.DispatchEvent NewEvent("input", True)
        Settle 700
    Next i
    perEdit = (Timer - started) / 150
    report = "pca: cache nodes " & AnswerVariantCacheSize() & ", " & Format$(perEdit * 1000, "0") & " ms per edit"
    Q("#tutorialList .operation[data-id=""kfold""]").Click
    Settle 700
    Q("#loadKfoldSampleBtn").Click
    Settle 700
    started = Timer
    For i = 1 To 50
        Set el = Q("#kfoldK")
        el.Value = CStr(2 + (i Mod 9))
        el.DispatchEvent NewEvent("input", True)
        Settle 700
    Next i
    report = report & "; kfold: " & Format$((Timer - started) / 50 * 1000, "0") & " ms per edit, cache nodes " & AnswerVariantCacheSize()
    StressRender = report & vbLf & "errors: " & LastErrors
End Function
