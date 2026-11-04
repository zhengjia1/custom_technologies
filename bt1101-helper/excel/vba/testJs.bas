Attribute VB_Name = "testJs"
Option Explicit
Option Compare Binary

' Tests for modJs, clsList and clsMap.
'   RunCases(path)    differential cases written by test/js-cases.js (V8 results)
'   TestCollections() unit tests for clsList / clsMap
'   Bench()           timing of the main formatting/parsing helpers

Private mTags() As String
Private mTagCount() As Long
Private mTagFail() As Long
Private mTagTime() As Double
Private mNTags As Long

' ---------------------------------------------------------------- file reading

Private Function ReadUtf8(ByVal path As String) As String
    Dim f As Integer, b() As Byte, n As Long
    f = FreeFile
    Open path For Binary Access Read As #f
    n = LOF(f)
    If n > 0 Then
        ReDim b(0 To n - 1)
        Get #f, , b
    End If
    Close #f
    If n > 0 Then ReadUtf8 = DecodeUtf8(b)
End Function

Private Function DecodeUtf8(b() As Byte) As String
    Dim out() As Byte, i As Long, n As Long, u As Long, c As Long, cp As Long
    n = UBound(b) + 1
    ReDim out(0 To 2 * n + 3)
    i = 0
    If n >= 3 Then
        If b(0) = &HEF And b(1) = &HBB And b(2) = &HBF Then i = 3
    End If
    Do While i < n
        c = b(i)
        If c < &H80 Then
            cp = c
            i = i + 1
        ElseIf c < &HE0 Then
            cp = (c And &H1F) * 64& + (b(i + 1) And &H3F)
            i = i + 2
        ElseIf c < &HF0 Then
            cp = (c And &HF) * 4096& + (b(i + 1) And &H3F) * 64& + (b(i + 2) And &H3F)
            i = i + 3
        Else
            cp = (c And &H7) * 262144 + (b(i + 1) And &H3F) * 4096& + (b(i + 2) And &H3F) * 64& + (b(i + 3) And &H3F)
            i = i + 4
        End If
        If cp >= &H10000 Then
            cp = cp - &H10000
            PutUnit out, u, &HD800& + cp \ 1024
            PutUnit out, u, &HDC00& + (cp And 1023)
        Else
            PutUnit out, u, cp
        End If
    Loop
    If u = 0 Then Exit Function
    ReDim Preserve out(0 To 2 * u - 1)
    DecodeUtf8 = out
End Function

Private Sub PutUnit(out() As Byte, ByRef u As Long, ByVal cp As Long)
    out(2 * u) = cp And &HFF&
    out(2 * u + 1) = (cp \ 256) And &HFF&
    u = u + 1
End Sub

Private Function HexVal(ByVal h As String) As Long
    Dim i As Long, c As Long, v As Long
    For i = 1 To Len(h)
        c = AscW(Mid$(h, i, 1))
        Select Case c
        Case 48 To 57: c = c - 48
        Case 65 To 70: c = c - 55
        Case 97 To 102: c = c - 87
        End Select
        v = v * 16 + c
    Next
    HexVal = v
End Function

Private Function Unescape(ByVal s As String) As String
    Dim p As Long, q As Long, r As String
    p = 1
    q = InStr(p, s, "\u", vbBinaryCompare)
    If q = 0 Then Unescape = s: Exit Function
    Do While q > 0
        r = r & Mid$(s, p, q - p) & ChrW$(HexVal(Mid$(s, q + 2, 4)))
        p = q + 6
        q = InStr(p, s, "\u", vbBinaryCompare)
    Loop
    Unescape = r & Mid$(s, p)
End Function

Private Function Esc(ByVal s As String) As String
    Dim i As Long, c As Long, r As String
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1)) And &HFFFF&
        If c < 32 Or c = 92 Or c > 126 Then
            r = r & "\u" & Right$("000" & Hex$(c), 4)
        Else
            r = r & Mid$(s, i, 1)
        End If
    Next
    Esc = r
End Function

' ---------------------------------------------------------------- case evaluation

Private Function D(ByVal s As String) As Double
    If s = "#NaN" Then D = JsNaN() Else D = JsHexToDouble(Mid$(s, 2))
End Function

Private Function H(ByVal x As Double) As String
    If JsIsNaN(x) Then H = "#NaN" Else H = "#" & JsDoubleToHex(x)
End Function

Private Function B(ByVal v As Boolean) As String
    If v Then B = "true" Else B = "false"
End Function

Private Function L(ByVal s As String) As Long
    L = CLng(Val(s))
End Function

Private Function EvalCase(f() As String) As String
    Dim lst As clsList
    On Error GoTo Failed
    Select Case f(0)
    Case "S": EvalCase = JsNumberToString(D(f(1)))
    Case "F": EvalCase = JsToFixed(D(f(1)), L(f(2)))
    Case "P"
        If f(2) = "u" Then EvalCase = JsToPrecision(D(f(1))) Else EvalCase = JsToPrecision(D(f(1)), L(f(2)))
    Case "E"
        If f(2) = "u" Then EvalCase = JsToExponential(D(f(1))) Else EvalCase = JsToExponential(D(f(1)), L(f(2)))
    Case "L"
        If f(2) = "u" Then
            EvalCase = JsToLocaleString(D(f(1)))
        Else
            EvalCase = JsToLocaleString(D(f(1)), L(f(2)), L(f(3)), f(4) = "1")
        End If
    Case "R": EvalCase = H(JsMathRound(D(f(1))))
    Case "N": EvalCase = H(JsNumber(f(1)))
    Case "PF": EvalCase = H(JsParseFloat(f(1)))
    Case "PI"
        If f(2) = "u" Then EvalCase = H(JsParseInt(f(1))) Else EvalCase = H(JsParseInt(f(1), Val(f(2))))
    Case "ADD": EvalCase = H(JsAdd(D(f(1)), D(f(2))))
    Case "SUB": EvalCase = H(JsSub(D(f(1)), D(f(2))))
    Case "MUL": EvalCase = H(JsMul(D(f(1)), D(f(2))))
    Case "DIV": EvalCase = H(JsDiv(D(f(1)), D(f(2))))
    Case "MOD": EvalCase = H(JsMod(D(f(1)), D(f(2))))
    Case "POW": EvalCase = H(JsPow(D(f(1)), D(f(2))))
    Case "MAX": EvalCase = H(JsMax(D(f(1)), D(f(2))))
    Case "MIN": EvalCase = H(JsMin(D(f(1)), D(f(2))))
    Case "LT": EvalCase = B(JsLt(D(f(1)), D(f(2))))
    Case "EQ": EvalCase = B(JsEq(D(f(1)), D(f(2))))
    Case "GE": EvalCase = B(JsGe(D(f(1)), D(f(2))))
    Case "SQRT": EvalCase = H(JsSqrt(D(f(1))))
    Case "FLOOR": EvalCase = H(JsFloor(D(f(1))))
    Case "CEIL": EvalCase = H(JsCeil(D(f(1))))
    Case "TRUNC": EvalCase = H(JsTrunc(D(f(1))))
    Case "SIGN": EvalCase = H(JsSign(D(f(1))))
    Case "ABS": EvalCase = H(JsAbs(D(f(1))))
    Case "I32": EvalCase = CStr(JsToInt32(D(f(1))))
    Case "U32": EvalCase = H(JsToUint32(D(f(1))))
    Case "LOG": EvalCase = H(JsLog(D(f(1))))
    Case "EXP": EvalCase = H(JsExp(D(f(1))))
    Case "LOGR": EvalCase = H(JsLog(D(f(1))))
    Case "EXPR": EvalCase = H(JsExp(D(f(1))))
    Case "SINR": EvalCase = H(JsSin(D(f(1))))
    Case "COSR": EvalCase = H(JsCos(D(f(1))))
    Case "POWR": EvalCase = H(JsPow(D(f(1)), D(f(2))))
    Case "LOG10": EvalCase = H(JsLog10(D(f(1))))
    Case "LOG2": EvalCase = H(JsLog2(D(f(1))))
    Case "NAN": EvalCase = B(JsIsNaN(D(f(1))))
    Case "FIN": EvalCase = B(JsIsFinite(D(f(1))))
    Case "INT": EvalCase = B(JsIsInteger(D(f(1))))
    Case "RAND"
        Dim i As Long, r As String
        JsSeedRandom Val(f(1))
        For i = 1 To 8
            If i > 1 Then r = r & ","
            r = r & H(JsRandom())
        Next
        EvalCase = r
    Case "TRIM": EvalCase = JsTrim(f(1))
    Case "TRIMS": EvalCase = JsTrimStart(f(1))
    Case "TRIME": EvalCase = JsTrimEnd(f(1))
    Case "SLICE"
        If f(3) = "u" Then EvalCase = JsSlice(f(1), L(f(2))) Else EvalCase = JsSlice(f(1), L(f(2)), L(f(3)))
    Case "SUBSTRING"
        If f(3) = "u" Then EvalCase = JsSubstring(f(1), L(f(2))) Else EvalCase = JsSubstring(f(1), L(f(2)), L(f(3)))
    Case "INDEXOF"
        If f(3) = "u" Then EvalCase = CStr(JsIndexOf(f(1), f(2))) Else EvalCase = CStr(JsIndexOf(f(1), f(2), L(f(3))))
    Case "LASTINDEXOF"
        If f(3) = "u" Then EvalCase = CStr(JsLastIndexOf(f(1), f(2))) Else EvalCase = CStr(JsLastIndexOf(f(1), f(2), L(f(3))))
    Case "INCLUDES"
        If f(3) = "u" Then EvalCase = B(JsIncludes(f(1), f(2))) Else EvalCase = B(JsIncludes(f(1), f(2), L(f(3))))
    Case "STARTS"
        If f(3) = "u" Then EvalCase = B(JsStartsWith(f(1), f(2))) Else EvalCase = B(JsStartsWith(f(1), f(2), L(f(3))))
    Case "ENDS"
        If f(3) = "u" Then EvalCase = B(JsEndsWith(f(1), f(2))) Else EvalCase = B(JsEndsWith(f(1), f(2), L(f(3))))
    Case "PADS"
        If f(3) = "u" Then EvalCase = JsPadStart(f(1), L(f(2))) Else EvalCase = JsPadStart(f(1), L(f(2)), f(3))
    Case "PADE"
        If f(3) = "u" Then EvalCase = JsPadEnd(f(1), L(f(2))) Else EvalCase = JsPadEnd(f(1), L(f(2)), f(3))
    Case "REPEAT": EvalCase = JsRepeat(f(1), L(f(2)))
    Case "CHARAT": EvalCase = JsCharAt(f(1), L(f(2)))
    Case "CHARCODEAT": EvalCase = H(JsCharCodeAt(f(1), L(f(2))))
    Case "AT": EvalCase = JsAt(f(1), L(f(2)))
    Case "SPLIT"
        If f(2) = "u" Then
            If f(3) = "u" Then Set lst = JsSplit(f(1)) Else Set lst = JsSplit(f(1), , L(f(3)))
        Else
            If f(3) = "u" Then Set lst = JsSplit(f(1), f(2)) Else Set lst = JsSplit(f(1), f(2), L(f(3)))
        End If
        EvalCase = lst.Count & ":" & lst.Join(ChrW$(31))
    Case "REPLACEALL": EvalCase = JsReplaceAll(f(1), f(2), f(3))
    Case "LOWER": EvalCase = JsLower(f(1))
    Case "UPPER": EvalCase = JsUpper(f(1))
    Case Else: EvalCase = "?unknown tag"
    End Select
    Exit Function
Failed:
    If Err.Number = 5 Then EvalCase = "!RangeError" Else EvalCase = "!ERR" & Err.Number & " " & Err.Description
End Function

Private Function IsUlpTag(ByVal tag As String) As Boolean
    IsUlpTag = (tag = "EXPR" Or tag = "SINR" Or tag = "COSR" Or tag = "POWR")
End Function

Private Function WithinUlp(ByVal a As String, ByVal b As String) As Boolean
    Dim x As Double, y As Double
    If a = "#NaN" Or b = "#NaN" Then Exit Function
    x = D(a)
    y = D(b)
    If Not (JsIsFinite(x) And JsIsFinite(y)) Then Exit Function
    WithinUlp = (Abs(x - y) <= Abs(y) * 2.3E-16)
End Function

Private Function TagIndex(ByVal tag As String) As Long
    Dim i As Long
    For i = 0 To mNTags - 1
        If mTags(i) = tag Then TagIndex = i: Exit Function
    Next
    If mNTags = 0 Then
        ReDim mTags(0 To 99)
        ReDim mTagCount(0 To 99)
        ReDim mTagFail(0 To 99)
        ReDim mTagTime(0 To 99)
    End If
    mTags(mNTags) = tag
    TagIndex = mNTags
    mNTags = mNTags + 1
End Function

Public Function RunCases(ByVal path As String) As String
    Dim text As String, lines() As String, i As Long, j As Long, f() As String, got As String
    Dim expected As String, t As Double, ti As Long, total As Long, fails As Long, report As String
    Dim shown As Long, started As Double, ulps As Long
    mNTags = 0
    started = Timer
    text = ReadUtf8(path)
    lines = Split(text, vbLf)
    For i = 0 To UBound(lines)
        If Len(lines(i)) > 0 Then
            f = Split(lines(i), vbTab)
            For j = 0 To UBound(f)
                f(j) = Unescape(f(j))
            Next
            expected = f(UBound(f))
            ti = TagIndex(f(0))
            t = Timer
            got = EvalCase(f)
            mTagTime(ti) = mTagTime(ti) + (Timer - t)
            mTagCount(ti) = mTagCount(ti) + 1
            total = total + 1
            If got <> expected And IsUlpTag(f(0)) And Left$(got, 1) = "#" And Left$(expected, 1) = "#" Then
                If WithinUlp(got, expected) Then
                    ulps = ulps + 1
                    got = expected
                End If
            End If
            If got <> expected Then
                fails = fails + 1
                mTagFail(ti) = mTagFail(ti) + 1
                If shown < 80 Then
                    shown = shown + 1
                    report = report & "MISMATCH " & Esc(Join(f, " | ")) & vbLf & "     got " & Esc(got) & vbLf
                End If
            End If
        End If
    Next
    report = report & "cases " & total & ", mismatches " & fails & ", " & Format$(Timer - started, "0.0") & " s" & vbLf
    report = report & "(EXPR/SINR/COSR/POWR use VBA's Exp/Sin/Cos/^: " & ulps & " results differ from V8 by 1 ulp; not counted)" & vbLf
    For i = 0 To mNTags - 1
        report = report & mTags(i) & "=" & mTagCount(i)
        If mTagFail(i) > 0 Then report = report & " FAIL " & mTagFail(i)
        report = report & " (" & Format$(1000000# * mTagTime(i) / mTagCount(i), "0") & " us)  "
    Next
    RunCases = report
End Function

' ---------------------------------------------------------------- collections

Private Sub Check(ByRef log As String, ByRef n As Long, ByVal ok As Boolean, ByVal what As String)
    n = n + 1
    If Not ok Then log = log & "FAIL " & what & vbLf
End Sub

Public Function TestCollections() As String
    Dim log As String, n As Long, a As clsList, b As clsList, m As clsMap, v As Variant, i As Long
    Dim s As String, keys As clsList, t As Double

    Set a = NewList(1, "two", 3.5)
    Check log, n, a.Count = 3, "NewList count"
    Check log, n, a(0) = 1 And a(1) = "two" And a.Item(2) = 3.5, "default member get"
    Check log, n, IsEmpty(a(3)) And IsEmpty(a(-1)), "out of range is Empty"
    a(1) = "TWO"
    Check log, n, a(1) = "TWO", "default member let"
    Check log, n, a.Push(4, 5) = 5, "push returns length"
    Check log, n, a.Pop() = 5 And a.Count = 4, "pop"
    Check log, n, a.Shift() = 1 And a.Count = 3 And a(0) = "TWO", "shift"
    Check log, n, a.Unshift("x", "y") = 5 And a(0) = "x" And a(1) = "y" And a(2) = "TWO", "unshift"
    Check log, n, a.Join("|") = "x|y|TWO|3.5|4", "join numbers via JsNumberToString"
    Check log, n, a.Slice(1, -1).Join(",") = "y,TWO,3.5", "slice negative end"
    Check log, n, a.Slice(-2).Join(",") = "3.5,4", "slice negative start"
    Check log, n, a.Slice(3, 1).Count = 0, "slice empty"
    Set b = a.Splice(1, 2, "p", "q", "r")
    Check log, n, b.Join(",") = "y,TWO" And a.Join(",") = "x,p,q,r,3.5,4", "splice"
    Set b = a.Splice(4, 100)
    Check log, n, b.Join(",") = "3.5,4" And a.Join(",") = "x,p,q,r", "splice to end"
    Check log, n, a.IndexOf("q") = 2 And a.IndexOf("Q") = -1 And a.IndexOf("x", 1) = -1, "indexOf (case-sensitive)"
    Check log, n, a.Includes("r") And Not a.Includes("R"), "includes"
    a.Add 7
    Check log, n, a.IndexOf(7#) = 4 And a.IndexOf("7") = -1, "=== numbers vs strings"
    a.Add JsNaN()
    Check log, n, a.IndexOf(JsNaN()) = -1 And a.Includes(JsNaN()), "NaN indexOf/includes"
    Check log, n, a.Join(",") = "x,p,q,r,7,NaN", "join NaN"
    Check log, n, a.Reverse().Join(",") = "NaN,7,r,q,p,x", "reverse"
    Check log, n, JsIsNaN(a.First()), "first"
    Check log, n, a.Last() = "x" And a.At(-1) = "x" And JsIsNaN(a.At(0)) And IsEmpty(a.At(6)), "last/at"
    a(8) = "z"
    Check log, n, a.Count = 9 And IsEmpty(a(7)) And a.Join(",") = "NaN,7,r,q,p,x,,,z", "let past end extends"
    a.Clear
    Check log, n, a.Count = 0 And a.Join(",") = "" And UBound(a.ToArray()) = -1, "clear"

    ' objects
    Set a = New clsList
    Set b = NewList("inner")
    a.Add b
    Set a(1) = b
    Check log, n, a(0) Is b And a(1) Is b And a.IndexOf(b) = 0, "objects stored by reference"
    Check log, n, a(0)(0) = "inner", "nested default member"
    Set v = a.Pop()
    Check log, n, v Is b And a.Count = 1, "pop object"
    i = 0
    For Each v In a
        i = i + 1
        Check log, n, v Is b, "for each object"
    Next
    Check log, n, i = 1, "for each count"
    Set a = NewList(1, 2, 3)
    s = ""
    Dim w As Variant
    For Each v In a
        For Each w In a
            s = s & v & w & " "
        Next
    Next
    Check log, n, s = "11 12 13 21 22 23 31 32 33 ", "nested for each over one list"

    ' growth and for each
    Set a = New clsList
    For i = 0 To 9999
        a.Add i
    Next
    Check log, n, a.Count = 10000 And a(9999) = 9999, "10000 pushes"
    t = 0
    For Each v In a
        t = t + v
    Next
    Check log, n, t = 49995000, "for each sum"
    v = a.ToArray()
    Check log, n, UBound(v) = 9999 And v(1234) = 1234, "toArray"

    ' sorting
    Set a = NewList("b", "a", "B", "aa", "A", "", "b")
    Check log, n, a.SortStrings().Join(",") = ",A,B,a,aa,b,b", "sort strings by code unit"
    Set a = NewList(10, 9, 1, 100, -5, 2.5)
    Check log, n, a.Clone().SortStrings().Join(",") = "-5,1,10,100,2.5,9", "default sort of numbers is textual"
    Check log, n, a.SortNumbers().Join(",") = "-5,1,2.5,9,10,100", "numeric sort"
    Check log, n, a.SortNumbers(True).Join(",") = "100,10,9,2.5,1,-5", "numeric sort descending"
    Set a = NewList("w", "x", "y", "z", "v")
    Set keys = NewList(2, 1, 2, 1, 0)
    Check log, n, a.SortBy(keys).Join(",") = "v,x,z,w,y", "stable sortBy"
    Check log, n, keys.Join(",") = "2,1,2,1,0", "sortBy leaves keys"
    Set a = NewList("w", "x", "y", "z")
    Check log, n, a.SortBy(NewList("b", "a", "b", "a"), True).Join(",") = "w,y,x,z", "stable sortBy descending"
    Set a = NewList(1, 2)
    Set b = NewList(3)
    Check log, n, a.Concat(b).Join(",") = "1,2,3" And a.Count = 2, "concat"
    a.Insert 1, "mid"
    Check log, n, a.Join(",") = "1,mid,2", "insert"
    Check log, n, a.RemoveAt(0) = 1 And a.Join(",") = "mid,2", "removeAt"
    a.Swap 0, 1
    Check log, n, a.Join(",") = "2,mid", "swap"

    ' map
    Set m = NewMap()
    m("b") = 1
    m("a") = 2
    m("B") = 3
    Check log, n, m.Count = 3 And m("b") = 1 And m("B") = 3, "case-sensitive keys"
    Check log, n, m.Keys.Join(",") = "b,a,B" And m.Values.Join(",") = "1,2,3", "insertion order"
    m("b") = 10
    Check log, n, m.Keys.Join(",") = "b,a,B" And m("b") = 10, "update keeps position"
    Check log, n, IsEmpty(m("zz")) And Not m.Exists("zz") And m.Count = 3, "missing key reads Empty without adding"
    Check log, n, m.Remove("a") And Not m.Remove("a") And m.Count = 2 And m.Keys.Join(",") = "b,B", "remove"
    m.Add "a", "again"
    Check log, n, m.Keys.Join(",") = "b,B,a", "re-add goes last"
    Set m("obj") = b
    Check log, n, m("obj") Is b, "object values"
    m("") = "empty key"
    Check log, n, m("") = "empty key" And m.Exists(""), "empty key"
    s = ""
    For Each v In m
        s = s & v & ";"
    Next
    Check log, n, s = "b;B;a;obj;;", "for each keys"
    m.Clear
    Check log, n, m.Count = 0 And m.Keys.Count = 0, "map clear"

    t = Timer
    Set m = New clsMap
    For i = 0 To 4999
        m("key" & i) = i
    Next
    For i = 0 To 4999 Step 2
        m.Remove "key" & i
    Next
    Check log, n, m.Count = 2500 And m("key4999") = 4999 And IsEmpty(m("key4998")), "5000 keys"
    For i = 1 To 4999 Step 2
        If m("key" & i) <> i Then Check log, n, False, "lookup key" & i
    Next
    Check log, n, m.Keys(0) = "key1" And m.Keys(2499) = "key4999", "order after removals"
    m("KEY1") = "upper"
    Check log, n, m("key1") = 1 And m("KEY1") = "upper", "case-sensitive after rebuild"
    s = Format$((Timer - t) * 1000, "0") & " ms for 5000 inserts + 2500 removes + 2500 lookups"

    If Len(log) = 0 Then
        TestCollections = "OK " & n & " checks; " & s
    Else
        TestCollections = log & n & " checks; " & s
    End If
End Function

' ---------------------------------------------------------------- timing

Public Function Bench() As String
    Dim t As Double, i As Long, r As String, s As String, x As Double, N As Long
    N = 2000
    t = Timer
    For i = 1 To N
        s = JsToFixed(1234.5678 + i, 2)
    Next
    r = r & "JsToFixed(1234.5678+i, 2): " & Format$((Timer - t) / N * 1000000#, "0") & " us" & vbLf
    t = Timer
    For i = 1 To N
        s = JsToFixed(0.000123 * i, 3)
    Next
    r = r & "JsToFixed(0.000123*i, 3): " & Format$((Timer - t) / N * 1000000#, "0") & " us" & vbLf
    t = Timer
    For i = 1 To N
        s = JsToPrecision(0.04567 * i, 3)
    Next
    r = r & "JsToPrecision(0.04567*i, 3): " & Format$((Timer - t) / N * 1000000#, "0") & " us" & vbLf
    t = Timer
    For i = 1 To N
        s = JsNumberToString(i)
    Next
    r = r & "JsNumberToString(integer): " & Format$((Timer - t) / N * 1000000#, "0") & " us" & vbLf
    t = Timer
    For i = 1 To N
        s = JsNumberToString(i / 7)
    Next
    r = r & "JsNumberToString(i/7): " & Format$((Timer - t) / N * 1000000#, "0") & " us" & vbLf
    t = Timer
    For i = 1 To N
        s = JsNumberToString(1E-300 * i / 7)
    Next
    r = r & "JsNumberToString(1e-300*i/7): " & Format$((Timer - t) / N * 1000000#, "0") & " us" & vbLf
    t = Timer
    For i = 1 To N
        s = JsToLocaleString(i * 1.37, 0, 2, True)
    Next
    r = r & "JsToLocaleString(i*1.37, 0, 2): " & Format$((Timer - t) / N * 1000000#, "0") & " us" & vbLf
    t = Timer
    For i = 1 To N
        x = JsNumber("12.345")
    Next
    r = r & "JsNumber(""12.345""): " & Format$((Timer - t) / N * 1000000#, "0") & " us" & vbLf
    t = Timer
    For i = 1 To N
        x = JsNumber("0.12345678901234567")
    Next
    r = r & "JsNumber(""0.12345678901234567""): " & Format$((Timer - t) / N * 1000000#, "0") & " us" & vbLf
    t = Timer
    For i = 1 To N
        x = JsParseFloat("  3.14159 is pi")
    Next
    r = r & "JsParseFloat(""  3.14159 is pi""): " & Format$((Timer - t) / N * 1000000#, "0") & " us" & vbLf
    t = Timer
    For i = 1 To N
        x = JsMathRound(i * 0.37)
    Next
    r = r & "JsMathRound: " & Format$((Timer - t) / N * 1000000#, "0") & " us" & vbLf
    t = Timer
    For i = 1 To N
        x = JsRandom()
    Next
    r = r & "JsRandom: " & Format$((Timer - t) / N * 1000000#, "0") & " us" & vbLf
    t = Timer
    For i = 1 To N
        s = JsTrim("   some text  ")
    Next
    r = r & "JsTrim: " & Format$((Timer - t) / N * 1000000#, "0") & " us" & vbLf
    Bench = r
End Function

