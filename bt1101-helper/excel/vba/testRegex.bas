Attribute VB_Name = "testRegex"
' Test-only: differential and unit tests for modRegex / clsRegex / clsMatch.
'
'   node test/regex-cases.js                      writes .build/regex-cases.txt (Node's results)
'   node tools/vba.js run testRegex.RunRegexCases <absolute path> [pattern ids] --only modRegex.bas,clsRegex.cls,clsMatch.cls,testRegex.bas --no-form
'   node tools/vba.js run testRegex.RunRegexUnit --only ...
'   node tools/vba.js run testRegex.RunRegexPerf <absolute path> --only ...
'
' Self-contained (own UTF-8 reader) so it depends only on the regex modules.
Option Explicit

Private mSubjects() As String
Private mPatterns() As clsRegex
Private mPatSource() As String
Private mPatFlags() As String
Private mPatError() As String
Private mFails As Long
Private mReport As String

' =====================================================================
' Differential cases
' =====================================================================

' Replays the cases recorded by test/regex-cases.js. filter: optional comma-separated
' pattern ids to run. Returns a summary plus the first mismatches.
Public Function RunRegexCases(ByVal path As String, Optional ByVal filter As String = "") As String
    Dim text As String, lines() As String, i As Long, f() As String, tag As String
    Dim nCases As Long, nReplace As Long, started As Double, pid As Long, sid As Long
    Dim msg As String, nCompile As Long, wanted As String
    started = Timer
    text = ReadUtf8(path)
    lines = VBA.Split(text, vbLf)
    ReDim mSubjects(0 To 1023)
    ReDim mPatterns(0 To 255)
    ReDim mPatSource(0 To 255)
    ReDim mPatFlags(0 To 255)
    ReDim mPatError(0 To 255)
    mFails = 0
    mReport = ""
    If Len(filter) > 0 Then wanted = "," & filter & ","
    For i = 0 To UBound(lines)
        If Len(lines(i)) > 0 Then
            tag = Left$(lines(i), 1)
            f = VBA.Split(lines(i), vbTab)
            Select Case tag
            Case "S"
                sid = CLng(f(1))
                If sid > UBound(mSubjects) Then ReDim Preserve mSubjects(0 To sid * 2)
                mSubjects(sid) = Unesc(f(2))
            Case "P"
                pid = CLng(f(1))
                If pid > UBound(mPatterns) Then
                    ReDim Preserve mPatterns(0 To pid * 2)
                    ReDim Preserve mPatSource(0 To pid * 2)
                    ReDim Preserve mPatFlags(0 To pid * 2)
                    ReDim Preserve mPatError(0 To pid * 2)
                End If
                mPatSource(pid) = Unesc(f(3))
                mPatFlags(pid) = f(2)
                If mPatFlags(pid) = "-" Then mPatFlags(pid) = ""
                msg = CompilePattern(pid)
                If Len(msg) > 0 Then
                    nCompile = nCompile + 1
                    Fail pid, -1, "compile", "", msg
                End If
            Case "C", "R"
                pid = CLng(f(1))
                If Len(wanted) = 0 Or InStr(1, wanted, "," & pid & ",") > 0 Then
                    If Not mPatterns(pid) Is Nothing Then
                        If tag = "C" Then
                            RunCase pid, CLng(f(2)), f(3), f(4), f(5), f(6)
                            nCases = nCases + 1
                        Else
                            RunReplace pid, CLng(f(2)), Unesc(f(3)), Unesc(f(4))
                            nReplace = nReplace + 1
                        End If
                    End If
                End If
            End Select
        End If
    Next i
    RunRegexCases = "cases=" & nCases & " replace=" & nReplace & " compileErrors=" & nCompile & _
        " mismatches=" & mFails & " seconds=" & Format$(Timer - started, "0.0") & mReport
End Function

Private Function CompilePattern(ByVal pid As Long) As String
    On Error GoTo Failed
    Set mPatterns(pid) = Rx(mPatSource(pid), mPatFlags(pid))
    Exit Function
Failed:
    CompilePattern = Err.Description
End Function

Private Sub RunCase(ByVal pid As Long, ByVal sid As Long, ByVal wantExec As String, ByVal wantMiddle As String, _
        ByVal wantAll As String, ByVal wantSplit As String)
    Dim re As clsRegex, s As String, got As String, stage As String
    On Error GoTo Crashed
    Set re = mPatterns(pid)
    s = mSubjects(sid)
    stage = "exec"
    got = EncodeExec(re.Exec(s), s)
    If got <> wantExec Then Fail pid, sid, "exec", wantExec, got
    stage = "exec-middle"
    got = EncodeExec(re.Exec(s, Len(s) \ 2), s)
    If got <> wantMiddle Then Fail pid, sid, "exec from " & (Len(s) \ 2), wantMiddle, got
    stage = "matchAll"
    got = EncodeAll(re.MatchAll(s), s)
    If got <> wantAll Then Fail pid, sid, "matchAll", wantAll, got
    stage = "split"
    got = EncodeSplit(re.Split(s))
    If got <> UnescSplit(wantSplit) Then Fail pid, sid, "split", UnescSplit(wantSplit), got
    stage = "test"
    If re.Test(s) <> (wantExec <> "-") Then Fail pid, sid, "test", CStr(wantExec <> "-"), CStr(Not (wantExec <> "-"))
    Exit Sub
Crashed:
    Fail pid, sid, stage & " (run-time error)", "", Err.Number & ": " & Err.Description
End Sub

Private Sub RunReplace(ByVal pid As Long, ByVal sid As Long, ByVal template As String, ByVal want As String)
    Dim got As String
    On Error GoTo Crashed
    got = mPatterns(pid).Replace(mSubjects(sid), template)
    If got <> want Then Fail pid, sid, "replace " & Esc(template), want, got
    Exit Sub
Crashed:
    Fail pid, sid, "replace (run-time error)", "", Err.Number & ": " & Err.Description
End Sub

Private Sub Fail(ByVal pid As Long, ByVal sid As Long, ByVal what As String, ByVal want As String, ByVal got As String)
    mFails = mFails + 1
    If mFails > 40 Then Exit Sub
    mReport = mReport & vbLf & "#" & pid & " /" & Esc(mPatSource(pid)) & "/" & mPatFlags(pid) & " " & what
    If sid >= 0 Then mReport = mReport & " subject " & sid & " """ & Clip(Esc(mSubjects(sid)), 90) & """"
    mReport = mReport & vbLf & "    want: " & Clip(Esc(want), 300) & vbLf & "    got:  " & Clip(Esc(got), 300)
End Sub

Private Function EncodeExec(ByVal m As clsMatch, ByRef s As String) As String
    Dim g As Long, out As String
    If m Is Nothing Then
        EncodeExec = "-"
        Exit Function
    End If
    For g = 0 To m.GroupCount
        If g > 0 Then out = out & ";"
        out = out & m.GroupIndex(g) & "," & m.GroupEnd(g)
        If m.Matched(g) Then
            If m.Group(g) <> Mid$(s, m.GroupIndex(g) + 1, m.GroupEnd(g) - m.GroupIndex(g)) Then out = out & "!text"
        ElseIf m.Group(g) <> "" Or m.GroupIndex(g) <> -1 Then
            out = out & "!undefined"
        End If
    Next g
    If m.Value <> m.Group(0) Or m.Index <> m.GroupIndex(0) Or m.EndIndex <> m.GroupEnd(0) Then out = out & "!value"
    EncodeExec = out
End Function

Private Function EncodeAll(ByVal col As Collection, ByRef s As String) As String
    Dim parts() As String, i As Long
    If col.Count = 0 Then Exit Function
    ReDim parts(0 To col.Count - 1)
    For i = 1 To col.Count
        parts(i - 1) = EncodeExec(col(i), s)
    Next i
    EncodeAll = Join(parts, "|")
End Function

Private Function EncodeSplit(ByVal col As clsList) As String
    Dim parts() As String, i As Long
    If col.Count = 0 Then Exit Function
    ReDim parts(0 To col.Count - 1)
    For i = 0 To col.Count - 1
        If IsEmpty(col(i)) Then parts(i) = ChrW$(&H1E) Else parts(i) = col(i)
    Next i
    EncodeSplit = Join(parts, ChrW$(&H1F))
End Function

' The split field keeps U+001E/U+001F raw; items are escaped.
Private Function UnescSplit(ByVal field As String) As String
    Dim parts() As String, i As Long
    If Len(field) = 0 Then Exit Function
    parts = VBA.Split(field, ChrW$(&H1F))
    For i = 0 To UBound(parts)
        If parts(i) <> ChrW$(&H1E) Then parts(i) = Unesc(parts(i))
    Next i
    UnescSplit = Join(parts, ChrW$(&H1F))
End Function

' =====================================================================
' Unit tests (API details the differential cases do not cover)
' =====================================================================

Public Function RunRegexUnit() As String
    Dim re As clsRegex, m As clsMatch, col As Collection, pieces As clsList, s As String, i As Long, t As Double
    mFails = 0
    mReport = ""
    ' cache: exact keys (Collection keys would be case-insensitive)
    Check "cache same", Rx("ab", "g") Is Rx("ab", "g")
    Check "cache case-sensitive pattern", Not (Rx("ab", "g") Is Rx("AB", "g"))
    Check "cache flags", Not (Rx("ab", "") Is Rx("ab", "i"))
    For i = 1 To 1500
        Set re = Rx("x" & i, "")
    Next i
    Check "cache growth", Rx("x77", "").Test("ax77b") And Rx("x1500", "") Is Rx("x1500", "")
    ' properties
    Set re = Rx("(a)(b)?", "gim")
    Check "IsGlobal", re.IsGlobal
    Check "IgnoreCase", re.IgnoreCase
    Check "Multiline", re.Multiline
    Check "Unicode", Not re.Unicode
    Check "Source/Flags", re.Source = "(a)(b)?" And re.Flags = "gim"
    Check "GroupCount", re.GroupCount = 2
    ' match object
    Set m = Rx("(a)(b)?(c)").Exec("xxac")
    Check "Index", m.Index = 2
    Check "EndIndex", m.EndIndex = 4
    Check "Value", m.Value = "ac"
    Check "Group 1", m.Group(1) = "a" And m.Matched(1)
    Check "Group 2 undefined", m.Group(2) = "" And Not m.Matched(2) And m.GroupIndex(2) = -1
    Check "Group 3", m.Group(3) = "c" And m.GroupIndex(3) = 3 And m.GroupEnd(3) = 4
    Check "Group out of range", m.Group(9) = "" And Not m.Matched(9) And Not m.Matched(-1)
    Check "InputText", m.InputText = "xxac"
    Check "GroupCount on match", m.GroupCount = 3
    Set m = Rx("()").Exec("ab", 1)
    Check "empty group participates", m.Matched(1) And m.Group(1) = "" And m.Index = 1
    Set col = Rx("\d").MatchAll("a1b2c3")
    Check "MatchAll count", col.Count = 3
    Check "MatchAll shared input", col(3).InputText = "a1b2c3" And col(2).Value = "2" And col(3).Index = 5
    Check "Exec beyond end", Rx("").Exec("ab", 3) Is Nothing
    Check "Exec at end", Rx("$").Exec("ab", 2).Index = 2
    Check "Exec negative start", Rx("a").Exec("ab", -5).Index = 0
    Check "Exec no match", Rx("z").Exec("ab") Is Nothing
    ' split with limit
    Set pieces = Rx(",").Split("a,b,c", 2)
    Check "Split limit", pieces.Count = 2 And pieces(0) = "a" And pieces(1) = "b"
    Set pieces = Rx("(,)").Split("a,b", 2)
    Check "Split limit captures", pieces.Count = 2 And pieces(1) = ","
    Check "Split limit 0", Rx(",").Split("a,b", 0).Count = 0
    Set pieces = Rx("(x)?,").Split("a,b")
    Check "Split undefined capture", pieces.Count = 3 And IsEmpty(pieces(1)) And pieces(2) = "b"
    ' convenience functions
    Check "RxTest", RxTest("Hello", "^h", "i") And Not RxTest("Hello", "^h")
    Check "RxReplace", RxReplace("a-b-c", "-", "+", "g") = "a+b+c" And RxReplace("a-b-c", "-", "+") = "a+b-c"
    Check "RxMatch", RxMatch("x=12", "\d+").Value = "12" And RxMatch("x", "\d") Is Nothing
    s = "a.b*c+d?e^f$g{h}i(j)k|l[m]n\o/"
    Check "RxEscape", RxEscape("a.b*c") = "a\.b\*c" And Rx("^" & RxEscape(s) & "$").Test(s)
    ' invalid patterns and flags raise errors
    CheckError "(", ""
    CheckError "a)", ""
    CheckError "[a", ""
    CheckError "*a", ""
    CheckError "a{2,1}", ""
    CheckError "[z-a]", ""
    CheckError "(?<=a)b", ""
    CheckError "a**", ""
    CheckError "\", ""
    CheckError "a", "gg"
    CheckError "a", "x"
    CheckError "\p{Foo}", "u"
    ' Unicode text
    s = "H" & ChrW$(&H2080) & ": " & ChrW$(&H3BC) & " " & ChrW$(&H2265) & " 0; " & ChrW$(&H3C3) & " " & ChrW$(&H2264) & " 1"
    Check "unicode literal", Rx("H" & ChrW$(&H2080) & ":\s*" & ChrW$(&H3BC)).Test(s)
    Check "unicode class", Rx("[" & ChrW$(&H2264) & ChrW$(&H2265) & "]", "g").MatchAll(s).Count = 2
    Check "unicode replace", Rx("[" & ChrW$(&H2264) & ChrW$(&H2265) & "]", "g").Replace(s, "?") = _
        "H" & ChrW$(&H2080) & ": " & ChrW$(&H3BC) & " ? 0; " & ChrW$(&H3C3) & " ? 1"
    Check "sigma ignoreCase", Rx(ChrW$(&H3A3), "i").Test(s)
    Check "letters", Rx("\p{L}+", "gu").MatchAll(s).Count = 3
    ' long inputs: explicit stack, no "Out of stack space"
    s = String$(200000, "a") & "b"
    t = Timer
    Check "long (?:a|b)*", Rx("^(?:a|b)*$").Exec(s).EndIndex = 200001
    Check "long (a|b)* capture", Rx("^(a|b)*$").Exec(s).Group(1) = "b"
    Check "long \s+", Rx("\s+").Exec("x" & Space$(200000) & "y").EndIndex = 200001
    Check "long lazy", Rx("^[\s\S]*?b").Exec(s).EndIndex = 200001
    Check "long lookahead", Rx("a(?=a*b)").MatchAll(Left$(s, 2000) & "b").Count = 2000
    Check "long replace", Len(Rx("a", "g").Replace(s, "xy")) = 400001
    Check "long split", Rx("(?:)").Split(Left$(s, 5000)).Count = 5000
    RunRegexUnit = "unit failures=" & mFails & " long-input seconds=" & Format$(Timer - t, "0.00") & mReport
End Function

Private Sub Check(ByVal name As String, ByVal ok As Boolean)
    If ok Then Exit Sub
    mFails = mFails + 1
    mReport = mReport & vbLf & "FAIL " & name
End Sub

Private Sub CheckError(ByVal pattern As String, ByVal flags As String)
    Dim re As clsRegex
    On Error GoTo Raised
    Set re = RxCompile(pattern, flags)
    mFails = mFails + 1
    mReport = mReport & vbLf & "FAIL no error for /" & pattern & "/" & flags
    Exit Sub
Raised:
End Sub

' =====================================================================
' Timings
' =====================================================================

' Times representative patterns. path: the cases file (its R console samples are used).
Public Function RunRegexPerf(ByVal path As String) As String
    Dim text As String, lines() As String, i As Long, f() As String, sample As String, cand As String
    Dim big As String, out As String, num As String, rowRe As clsRegex, outLines() As String, k As Long, n As Long
    text = ReadUtf8(path)
    lines = VBA.Split(text, vbLf)
    For i = 0 To UBound(lines)
        If Left$(lines(i), 2) = "S" & vbTab Then
            f = VBA.Split(lines(i), vbTab)
            cand = Unesc(f(2))
            If InStr(1, cand, "Coefficients:") > 0 And InStr(1, cand, "Residual standard error") > 0 Then
                If Len(cand) > Len(sample) And Len(cand) <= 3000 Then sample = cand
            End If
        End If
    Next i
    Do While Len(big) < 50000
        big = big & sample & vbLf
    Loop
    big = Left$(big, 50000)
    out = "sample length=" & Len(sample) & " (lm summary); big length=" & Len(big)
    num = "[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?"
    out = out & vbLf & TimeTest("compile fresh coefficient-row pattern (RxCompile)", "compile", "^\s*(.+?)\s+(" & num & ")\s+(" & num & ")\s+(" & num & ")\s+(<=|>=|<|>)?\s*(" & num & ")(?:\s+\S+)?\s*$", "i", sample)
    Set rowRe = Rx("^\s*(.+?)\s+(" & num & ")\s+(" & num & ")\s+(" & num & ")\s+(<=|>=|<|>)?\s*(" & num & ")(?:\s+\S+)?\s*$", "i")
    outLines = VBA.Split(sample, vbLf)
    Dim t As Double, reps As Long
    reps = 0
    t = Timer
    Do
        For k = 0 To UBound(outLines)
            If rowRe.Test(outLines(k)) Then n = n + 1
        Next k
        reps = reps + 1
    Loop While Timer - t < 1
    out = out & vbLf & "coefficient-row regex over every line of the sample (" & (UBound(outLines) + 1) & " lines): " & Format$((Timer - t) * 1000 / reps, "0.00") & " ms"
    out = out & vbLf & TimeTest("detector test /call:\s*lm|residual standard error|coefficients:/i", "test", "call:\s*lm|residual standard error|coefficients:", "i", sample)
    out = out & vbLf & TimeTest("F-statistic exec", "exec", "F-statistic\s*:\s*(" & num & ")\s+on\s+(" & num & ")\s+and\s+(" & num & ")\s+DF\s*,?\s*p-value\s*:\s*(<=|>=|<|>|=)?\s*(" & num & ")", "i", sample)
    out = out & vbLf & TimeTest("p-value exec (no match)", "exec", "p(?:-|\s{0,3})value\s{0,3}(<=|>=|<|>|=)\s{0,3}([+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?)", "i", Replace(sample, "p-value", "q"))
    out = out & vbLf & TimeTest("/\s+/g replace with ' ' (2 KB)", "replace", "\s+", "g", sample)
    out = out & vbLf & TimeTest("/[\s\S]*?Residual/ exec (2 KB)", "exec", "[\s\S]*?Residual", "", sample)
    out = out & vbLf & TimeTest("/^.*$/gm matchAll (2 KB)", "all", "^.*$", "gm", sample)
    out = out & vbLf & TimeTest("/\s+/g replace (50 KB)", "replace", "\s+", "g", big)
    out = out & vbLf & TimeTest("/^.*$/gm matchAll (50 KB)", "all", "^.*$", "gm", big)
    out = out & vbLf & TimeTest("/[\s\S]*?ZZZ$/ exec success (50 KB)", "exec", "[\s\S]*?ZZZ$", "", big & "ZZZ")
    out = out & vbLf & TimeTest("/\bThe excluded (variables|variable)\b/g replace, no match (50 KB)", "replace", "\bThe excluded (variables|variable)\b", "g", big)
    out = out & vbLf & TimeTest("/\r?\n/ split (50 KB)", "split", "\r?\n", "", big)
    RunRegexPerf = out
End Function

Private Function TimeTest(ByVal label As String, ByVal action As String, ByVal pattern As String, ByVal flags As String, ByVal s As String) As String
    Dim t As Double, reps As Long, re As clsRegex, dummy As Variant
    Set re = Rx(pattern, flags)
    t = Timer
    Do
        Select Case action
        Case "compile": Set re = RxCompile(pattern, flags)
        Case "test": dummy = re.Test(s)
        Case "exec": Set dummy = re.Exec(s)
        Case "all": Set dummy = re.MatchAll(s)
        Case "replace": dummy = re.Replace(s, " ")
        Case "split": Set dummy = re.Split(s)
        End Select
        reps = reps + 1
    Loop While Timer - t < 1
    TimeTest = label & ": " & Format$((Timer - t) * 1000 / reps, "0.000") & " ms"
End Function

' =====================================================================
' Helpers
' =====================================================================

Private Function ReadUtf8(ByVal path As String) As String
    Dim handle As Integer, bytes() As Byte, size As Long
    handle = FreeFile
    Open path For Binary Access Read As #handle
    size = LOF(handle)
    If size = 0 Then
        Close #handle
        Exit Function
    End If
    ReDim bytes(0 To size - 1)
    Get #handle, , bytes
    Close #handle
    ReadUtf8 = DecodeUtf8(bytes)
End Function

Private Function DecodeUtf8(ByRef bytes() As Byte) As String
    Dim out() As Byte, n As Long, i As Long, last As Long, b As Long, code As Long, extra As Long
    last = UBound(bytes)
    ReDim out(0 To 2 * (last + 2))
    i = 0
    If last >= 2 Then
        If bytes(0) = &HEF And bytes(1) = &HBB And bytes(2) = &HBF Then i = 3
    End If
    Do While i <= last
        b = bytes(i)
        If b < &H80 Then
            code = b: extra = 0
        ElseIf b < &HE0 Then
            code = b And &H1F: extra = 1
        ElseIf b < &HF0 Then
            code = b And &HF: extra = 2
        Else
            code = b And &H7: extra = 3
        End If
        i = i + 1
        Do While extra > 0 And i <= last
            code = code * 64 + (bytes(i) And &H3F)
            i = i + 1
            extra = extra - 1
        Loop
        If code >= &H10000 Then
            code = code - &H10000
            PutUnit out, n, &HD800& + (code \ 1024)
            PutUnit out, n, &HDC00& + (code And 1023)
        Else
            PutUnit out, n, code
        End If
    Loop
    If n = 0 Then Exit Function
    ReDim Preserve out(0 To 2 * n - 1)
    DecodeUtf8 = out
End Function

Private Sub PutUnit(ByRef out() As Byte, ByRef n As Long, ByVal code As Long)
    out(2 * n) = code And &HFF
    out(2 * n + 1) = (code \ 256) And &HFF
    n = n + 1
End Sub

' Reverses the file escaping ("\\" and "\uXXXX").
Private Function Unesc(ByVal s As String) As String
    Dim p As Long, out As String, start As Long, c As String
    p = InStr(1, s, "\", vbBinaryCompare)
    If p = 0 Then
        Unesc = s
        Exit Function
    End If
    start = 1
    Do While p > 0
        out = out & Mid$(s, start, p - start)
        c = Mid$(s, p + 1, 1)
        If c = "u" Then
            out = out & ChrW$(CLng("&H" & Mid$(s, p + 2, 4)))
            start = p + 6
        Else
            out = out & c
            start = p + 2
        End If
        p = InStr(start, s, "\", vbBinaryCompare)
    Loop
    Unesc = out & Mid$(s, start)
End Function

' Escapes control characters and backslashes for the report.
Private Function Esc(ByVal s As String) As String
    Dim i As Long, c As Long, out As String
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1)) And &HFFFF&
        If c = 92 Then
            out = out & "\\"
        ElseIf c < 32 Or c > 126 Then
            out = out & "\u" & Right$("000" & LCase$(Hex$(c)), 4)
        Else
            out = out & ChrW$(c)
        End If
        If Len(out) > 2000 Then Exit For
    Next i
    Esc = out
End Function

Private Function Clip(ByVal s As String, ByVal n As Long) As String
    If Len(s) > n Then Clip = Left$(s, n) & "..." Else Clip = s
End Function
