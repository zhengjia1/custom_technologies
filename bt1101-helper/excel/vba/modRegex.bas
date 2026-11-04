Attribute VB_Name = "modRegex"
' JavaScript-compatible regular expressions in pure VBA (Windows and Mac Excel).
'
'   Set re = Rx(pattern, flags)      compiled once and cached by flags & pattern
'   re.Test / re.Exec / re.MatchAll / re.Replace / re.Split      (see clsRegex)
'
' Supported: literals, ., classes with ranges/negation/escapes, \d \D \s \S \w \W,
' \n \r \t \f \v \0 \xHH \uHHHH \cX, escaped metacharacters, ^ $ (m flag), \b \B,
' capturing and (?:) groups, backreferences \1.., alternation, * + ? {n} {n,} {n,m}
' greedy and lazy, (?=) (?!), flags g i m s u y, \p{L} / \P{L} under u.
' Case-insensitive matching uses JavaScript's non-unicode canonicalisation
' (String.prototype.toUpperCase, single code unit, never non-ASCII -> ASCII) for the
' whole Basic Multilingual Plane. Not supported: lookbehind, named groups, astral
' (surrogate-pair) code points as single characters in u mode.
'
' The compiler turns a pattern into a small program for the backtracking machine in
' clsRegex. The machine uses an explicit stack (no recursion except one level per
' nested lookahead), and quantified single-character atoms run as tight loops.
Option Explicit

' ---- Opcodes shared with clsRegex ----
Public Const RX_MATCH As Long = 0
Public Const RX_CHAR As Long = 1        ' A = code (canonical when ignoreCase)
Public Const RX_ANY As Long = 2         ' . without s flag
Public Const RX_CLASS As Long = 3       ' A = class index
Public Const RX_STR As Long = 4         ' A = literal offset, B = length
Public Const RX_BOL As Long = 5         ' A = 1 when multiline
Public Const RX_EOL As Long = 6         ' A = 1 when multiline
Public Const RX_WORDB As Long = 7
Public Const RX_NWORDB As Long = 8
Public Const RX_SPLIT As Long = 9       ' try A first, then B
Public Const RX_JMP As Long = 10        ' A = target
Public Const RX_SAVE As Long = 11       ' A = capture slot
Public Const RX_REPS As Long = 12       ' single-char repeat: A = atom kind, B = atom arg, C = min, D = max (-1), E = greedy
Public Const RX_MARK As Long = 13       ' A = register: remember the position
Public Const RX_CHECK As Long = 14      ' A = register: fail when no progress
Public Const RX_CLEAR As Long = 15      ' A..B = capture groups reset to undefined
Public Const RX_LOOK As Long = 16       ' A = 1 negative, B = body pc, C = continuation pc
Public Const RX_LOOKEND As Long = 17
Public Const RX_BACKREF As Long = 18    ' A = group number
Public Const RX_ANYALL As Long = 19     ' . with s flag

' Atom kinds for RX_REPS
Public Const RXA_CHAR As Long = 1
Public Const RXA_ANY As Long = 2
Public Const RXA_CLASS As Long = 3
Public Const RXA_ALL As Long = 4

' ---- Subject text shared with clsRegex (UTF-16 code units) ----
Public gRxCode() As Long
Public gRxLen As Long
Public gRxText As String
Private mTextReady As Boolean

' ---- Case folding (loaded on first ignoreCase compile) ----
Public gRxCanon() As Long
Public gRxCanonReady As Boolean
Public gRxFoldNext() As Long     ' circular lists of code units with one canonical value
Private mCased() As Long         ' sorted code units whose fold list is not just themselves
Private mCasedCount As Long

' ---- Cache (open addressing; keys compared exactly) ----
Private mCacheKeys() As String
Private mCacheObjs() As clsRegex
Private mCacheCap As Long
Private mCacheCount As Long

' ---- Parser state ----
Private Const N_EMPTY As Long = 0
Private Const N_CHAR As Long = 1
Private Const N_ANY As Long = 2
Private Const N_CLASS As Long = 3
Private Const N_SEQ As Long = 4
Private Const N_ALT As Long = 5
Private Const N_GROUP As Long = 6
Private Const N_REPEAT As Long = 7
Private Const N_BOL As Long = 8
Private Const N_EOL As Long = 9
Private Const N_WORDB As Long = 10
Private Const N_NWORDB As Long = 11
Private Const N_LOOK As Long = 12
Private Const N_NLOOK As Long = 13
Private Const N_BACKREF As Long = 14

Private pSrc() As Long
Private pLen As Long
Private pPos As Long
Private pPattern As String
Private pIgnoreCase As Boolean
Private pMultiline As Boolean
Private pDotAll As Boolean
Private pUnicode As Boolean
Private pGroups As Long
Private pTotalGroups As Long

' AST nodes
Private nKind() As Long
Private nA() As Long
Private nB() As Long
Private nC() As Long
Private nD() As Long
Private nE() As Long
Private nFirst() As Long
Private nLastChild() As Long
Private nNext() As Long
Private nCount As Long

' Character classes (all classes of the pattern being compiled)
Private cOff() As Long
Private cCnt() As Long
Private cCount As Long
Private cLo() As Long
Private cHi() As Long
Private cRanges As Long
' Range builder for the class being parsed
Private tLo() As Long
Private tHi() As Long
Private tN As Long

' Program
Private oOp() As Long
Private oA() As Long
Private oB() As Long
Private oC() As Long
Private oD() As Long
Private oE() As Long
Private oCount As Long
Private oLit() As Long
Private oLitCount As Long
Private oRegs As Long

Private Const MAX_PROGRAM As Long = 200000

' =====================================================================
' Public API
' =====================================================================

' Returns the compiled regular expression for pattern/flags (cached).
Public Function Rx(ByVal pattern As String, Optional ByVal flags As String = "") As clsRegex
    Dim key As String, h As Long
    key = flags & "/" & pattern
    If mCacheCap = 0 Then InitCache 512
    h = HashKey(key) And (mCacheCap - 1)
    Do
        If mCacheObjs(h) Is Nothing Then Exit Do
        If mCacheKeys(h) = key Then
            Set Rx = mCacheObjs(h)
            Exit Function
        End If
        h = (h + 1) And (mCacheCap - 1)
    Loop
    Set Rx = RxCompile(pattern, flags)
    mCacheKeys(h) = key
    Set mCacheObjs(h) = Rx
    mCacheCount = mCacheCount + 1
    If mCacheCount * 2 > mCacheCap Then GrowCache
End Function

Public Function RxTest(ByVal text As String, ByVal pattern As String, Optional ByVal flags As String = "") As Boolean
    RxTest = Rx(pattern, flags).Test(text)
End Function

Public Function RxReplace(ByVal text As String, ByVal pattern As String, ByVal replacement As String, Optional ByVal flags As String = "") As String
    RxReplace = Rx(pattern, flags).Replace(text, replacement)
End Function

' First match (from index 0) or Nothing.
Public Function RxMatch(ByVal text As String, ByVal pattern As String, Optional ByVal flags As String = "") As clsMatch
    Set RxMatch = Rx(pattern, flags).Exec(text)
End Function

' Escapes regular-expression metacharacters, like the website's escapeRegExp:
' text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")
Public Function RxEscape(ByVal text As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If InStr(1, ".*+?^${}()|[]\", ch, vbBinaryCompare) > 0 Then out = out & "\"
        out = out & ch
    Next i
    RxEscape = out
End Function

' Clears the compiled-pattern cache.
Public Sub RxClearCache()
    mCacheCap = 0
    mCacheCount = 0
    Erase mCacheKeys
    Erase mCacheObjs
End Sub

' Compiles without caching.
Public Function RxCompile(ByVal pattern As String, Optional ByVal flags As String = "") As clsRegex
    Dim re As clsRegex, root As Long, i As Long, c As Long
    Dim fG As Boolean, fI As Boolean, fM As Boolean, fS As Boolean, fU As Boolean, fY As Boolean
    For i = 1 To Len(flags)
        c = AscW(Mid$(flags, i, 1))
        Select Case c
        Case 103: If fG Then FlagError flags
            fG = True
        Case 105: If fI Then FlagError flags
            fI = True
        Case 109: If fM Then FlagError flags
            fM = True
        Case 115: If fS Then FlagError flags
            fS = True
        Case 117: If fU Then FlagError flags
            fU = True
        Case 121: If fY Then FlagError flags
            fY = True
        Case Else: FlagError flags
        End Select
    Next i
    pPattern = pattern
    pIgnoreCase = fI: pMultiline = fM: pDotAll = fS: pUnicode = fU
    If pIgnoreCase Then InitCanon
    LoadPattern pattern
    ResetCompiler
    pTotalGroups = CountGroups()
    pGroups = 0
    pPos = 0
    root = ParseDisjunction()
    If pPos < pLen Then
        If pSrc(pPos) = 41 Then ParseFail "Unmatched ')'" Else ParseFail "Unexpected character"
    End If
    EmitNode root
    Emit RX_MATCH
    Set re = New clsRegex
    re.Init pattern, flags, fG, fI, fM, fU, fY, pGroups, oOp, oA, oB, oC, oD, oE, oCount, _
        oLit, oLitCount, cOff, cCnt, cCount, cLo, cHi, cRanges, oRegs
    Set RxCompile = re
End Function

' Loads text into the shared code-unit buffer (no work when it is already loaded).
Public Sub RxLoadText(ByRef text As String)
    Dim b() As Byte, n As Long, i As Long, j As Long
    If mTextReady Then
        If LenB(text) = LenB(gRxText) Then
            If text = gRxText Then Exit Sub
        End If
    End If
    n = Len(text)
    gRxText = text
    gRxLen = n
    If n = 0 Then
        ReDim gRxCode(0 To 0)
        mTextReady = True
        Exit Sub
    End If
    b = text
    ReDim gRxCode(0 To n)
    j = 0
    For i = 0 To n - 1
        gRxCode(i) = b(j) Or (CLng(b(j + 1)) * 256&)
        j = j + 2
    Next i
    mTextReady = True
End Sub

' JS Canonicalize (non-unicode ignoreCase) of one code unit.
Public Function RxCanon(ByVal c As Long) As Long
    If Not gRxCanonReady Then InitCanon
    If c >= 0 And c <= 65535 Then RxCanon = gRxCanon(c) Else RxCanon = c
End Function

' =====================================================================
' Cache
' =====================================================================

Private Sub InitCache(ByVal cap As Long)
    mCacheCap = cap
    mCacheCount = 0
    ReDim mCacheKeys(0 To cap - 1)
    ReDim mCacheObjs(0 To cap - 1)
End Sub

Private Sub GrowCache()
    Dim oldKeys() As String, oldObjs() As clsRegex, oldCap As Long, i As Long, h As Long
    oldKeys = mCacheKeys
    oldCap = mCacheCap
    ReDim oldObjs(0 To oldCap - 1)
    For i = 0 To oldCap - 1
        Set oldObjs(i) = mCacheObjs(i)
    Next i
    InitCache oldCap * 2
    For i = 0 To oldCap - 1
        If Not oldObjs(i) Is Nothing Then
            h = HashKey(oldKeys(i)) And (mCacheCap - 1)
            Do While Not mCacheObjs(h) Is Nothing
                h = (h + 1) And (mCacheCap - 1)
            Loop
            mCacheKeys(h) = oldKeys(i)
            Set mCacheObjs(h) = oldObjs(i)
            mCacheCount = mCacheCount + 1
        End If
    Next i
End Sub

Private Function HashKey(ByRef key As String) As Long
    Dim b() As Byte, i As Long, h As Long
    If Len(key) = 0 Then Exit Function
    b = key
    h = 5381
    For i = 0 To UBound(b)
        h = ((h And &H3FFFFF) * 33 + b(i)) And &H7FFFFFF
    Next i
    HashKey = h
End Function

' =====================================================================
' Unicode tables
' =====================================================================

Private Sub InitCanon()
    Dim s As String, k As Long, start As Long, cnt As Long, delta As Long, stp As Long, j As Long, c As Long
    Dim first As Long, prev As Long
    If gRxCanonReady Then Exit Sub
    ReDim gRxCanon(0 To 65535)
    ReDim gRxFoldNext(0 To 65535)
    For c = 0 To 65535
        gRxCanon(c) = c
        gRxFoldNext(c) = c
    Next c
    s = CanonData()
    For k = 1 To Len(s) Step 11
        start = CLng("&H" & Mid$(s, k, 4))
        cnt = CLng("&H" & Mid$(s, k + 4, 2))
        delta = CLng("&H" & Mid$(s, k + 6, 4))
        If delta >= 32768 Then delta = delta - 65536
        stp = CLng(Mid$(s, k + 10, 1))
        For j = 0 To cnt - 1
            c = start + j * stp
            gRxCanon(c) = (c + delta + 65536) Mod 65536
        Next j
    Next k
    ' Circular lists of the code units sharing one canonical value.
    For c = 0 To 65535
        If gRxCanon(c) <> c Then
            first = gRxCanon(c)
            ' insert c after first in first's cycle
            prev = gRxFoldNext(first)
            gRxFoldNext(first) = c
            gRxFoldNext(c) = prev
        End If
    Next c
    ReDim mCased(0 To 4095)
    mCasedCount = 0
    For c = 0 To 65535
        If gRxFoldNext(c) <> c Then
            If mCasedCount > UBound(mCased) Then ReDim Preserve mCased(0 To mCasedCount * 2)
            mCased(mCasedCount) = c
            mCasedCount = mCasedCount + 1
        End If
    Next c
    gRxCanonReady = True
End Sub

Private Sub LoadPattern(ByRef pattern As String)
    Dim b() As Byte, i As Long, j As Long
    pLen = Len(pattern)
    ReDim pSrc(0 To pLen + 4)
    If pLen = 0 Then Exit Sub
    b = pattern
    j = 0
    For i = 0 To pLen - 1
        pSrc(i) = b(j) Or (CLng(b(j + 1)) * 256&)
        j = j + 2
    Next i
End Sub

' =====================================================================
' Errors
' =====================================================================

Private Sub FlagError(ByVal flags As String)
    Err.Raise 5, "Rx", "Invalid flags supplied to RegExp constructor '" & flags & "'"
End Sub

Private Sub ParseFail(ByVal message As String)
    Err.Raise 5, "Rx", "Invalid regular expression: /" & pPattern & "/: " & message
End Sub

' =====================================================================
' Parser
' =====================================================================

Private Sub ResetCompiler()
    nCount = 0
    ReDim nKind(0 To 63): ReDim nA(0 To 63): ReDim nB(0 To 63): ReDim nC(0 To 63)
    ReDim nD(0 To 63): ReDim nE(0 To 63): ReDim nFirst(0 To 63): ReDim nLastChild(0 To 63)
    ReDim nNext(0 To 63)
    cCount = 0: cRanges = 0
    ReDim cOff(0 To 7): ReDim cCnt(0 To 7): ReDim cLo(0 To 63): ReDim cHi(0 To 63)
    oCount = 0: oLitCount = 0: oRegs = 0
    ReDim oOp(0 To 63): ReDim oA(0 To 63): ReDim oB(0 To 63): ReDim oC(0 To 63)
    ReDim oD(0 To 63): ReDim oE(0 To 63): ReDim oLit(0 To 15)
End Sub

Private Function NewNode(ByVal kind As Long) As Long
    Dim n As Long
    If nCount > UBound(nKind) Then
        n = UBound(nKind) * 2 + 1
        ReDim Preserve nKind(0 To n): ReDim Preserve nA(0 To n): ReDim Preserve nB(0 To n)
        ReDim Preserve nC(0 To n): ReDim Preserve nD(0 To n): ReDim Preserve nE(0 To n)
        ReDim Preserve nFirst(0 To n): ReDim Preserve nLastChild(0 To n): ReDim Preserve nNext(0 To n)
    End If
    n = nCount
    nKind(n) = kind: nA(n) = 0: nB(n) = 0: nC(n) = 0: nD(n) = 0: nE(n) = 0
    nFirst(n) = -1: nLastChild(n) = -1: nNext(n) = -1
    nCount = nCount + 1
    NewNode = n
End Function

Private Sub AddChild(ByVal parent As Long, ByVal child As Long)
    If nFirst(parent) < 0 Then
        nFirst(parent) = child
    Else
        nNext(nLastChild(parent)) = child
    End If
    nLastChild(parent) = child
End Sub

Private Function CharNode(ByVal code As Long) As Long
    Dim n As Long
    n = NewNode(N_CHAR)
    nA(n) = code
    CharNode = n
End Function

' Counts capturing groups ahead of parsing (decides \N backreference vs octal).
Private Function CountGroups() As Long
    Dim i As Long, inClass As Boolean, n As Long
    i = 0
    Do While i < pLen
        Select Case pSrc(i)
        Case 92
            i = i + 1
        Case 91
            inClass = True
        Case 93
            inClass = False
        Case 40
            If Not inClass Then
                If pSrc(i + 1) <> 63 Then
                    n = n + 1
                ElseIf pSrc(i + 2) = 60 And pSrc(i + 3) <> 61 And pSrc(i + 3) <> 33 Then
                    n = n + 1
                End If
            End If
        End Select
        i = i + 1
    Loop
    CountGroups = n
End Function

Private Function PeekC() As Long
    If pPos < pLen Then PeekC = pSrc(pPos) Else PeekC = -1
End Function

Private Function ParseDisjunction() As Long
    Dim first As Long, alt As Long
    first = ParseAlternative()
    If PeekC() <> 124 Then
        ParseDisjunction = first
        Exit Function
    End If
    alt = NewNode(N_ALT)
    AddChild alt, first
    Do While PeekC() = 124
        pPos = pPos + 1
        AddChild alt, ParseAlternative()
    Loop
    ParseDisjunction = alt
End Function

Private Function ParseAlternative() As Long
    Dim seq As Long, c As Long
    seq = NewNode(N_SEQ)
    Do While pPos < pLen
        c = pSrc(pPos)
        If c = 124 Or c = 41 Then Exit Do
        AddChild seq, ParseTerm()
    Loop
    ParseAlternative = seq
End Function

Private Function ParseTerm() As Long
    Dim c As Long, c2 As Long, atom As Long, capBefore As Long, mn As Long, mx As Long
    c = pSrc(pPos)
    capBefore = pGroups
    Select Case c
    Case 94
        pPos = pPos + 1
        ParseTerm = NewNode(N_BOL)
        Exit Function
    Case 36
        pPos = pPos + 1
        ParseTerm = NewNode(N_EOL)
        Exit Function
    Case 92
        If pPos + 1 < pLen Then
            c2 = pSrc(pPos + 1)
            If c2 = 98 Or c2 = 66 Then
                pPos = pPos + 2
                If c2 = 98 Then ParseTerm = NewNode(N_WORDB) Else ParseTerm = NewNode(N_NWORDB)
                Exit Function
            End If
        End If
        atom = ParseAtomEscape()
    Case 40
        atom = ParseGroup()
        If pUnicode And (nKind(atom) = N_LOOK Or nKind(atom) = N_NLOOK) Then
            ParseTerm = atom
            Exit Function
        End If
    Case 46
        pPos = pPos + 1
        atom = NewNode(N_ANY)
    Case 91
        atom = ParseClass()
    Case 42, 43, 63
        ParseFail "Nothing to repeat"
    Case 123
        If pUnicode Then ParseFail "Lone quantifier brackets"
        If ParseBrace(mn, mx, False) Then ParseFail "Nothing to repeat"
        pPos = pPos + 1
        atom = CharNode(123)
    Case Else
        pPos = pPos + 1
        atom = CharNode(c)
    End Select
    ParseTerm = ParseQuantifier(atom, capBefore)
End Function

Private Function ParseQuantifier(ByVal atom As Long, ByVal capBefore As Long) As Long
    Dim mn As Long, mx As Long, greedy As Boolean, rep As Long
    ParseQuantifier = atom
    If pPos >= pLen Then Exit Function
    Select Case pSrc(pPos)
    Case 42
        mn = 0: mx = -1: pPos = pPos + 1
    Case 43
        mn = 1: mx = -1: pPos = pPos + 1
    Case 63
        mn = 0: mx = 1: pPos = pPos + 1
    Case 123
        If Not ParseBrace(mn, mx, True) Then
            If pUnicode Then ParseFail "Incomplete quantifier"
            Exit Function
        End If
    Case Else
        Exit Function
    End Select
    greedy = True
    If pPos < pLen Then
        If pSrc(pPos) = 63 Then
            greedy = False
            pPos = pPos + 1
        End If
    End If
    If mx >= 0 And mx < mn Then ParseFail "numbers out of order in {} quantifier"
    rep = NewNode(N_REPEAT)
    nFirst(rep) = atom
    nA(rep) = mn
    nB(rep) = mx
    nC(rep) = IIf(greedy, 1, 0)
    nD(rep) = capBefore + 1
    nE(rep) = pGroups
    ParseQuantifier = rep
End Function

' Parses {n}, {n,}, {n,m} at pPos. Returns False (and consumes nothing) when invalid.
Private Function ParseBrace(ByRef mn As Long, ByRef mx As Long, ByVal consume As Boolean) As Boolean
    Dim p As Long, v As Double, digits As Long, lo As Double, hi As Double
    p = pPos + 1
    v = 0: digits = 0
    Do While p < pLen
        If pSrc(p) < 48 Or pSrc(p) > 57 Then Exit Do
        v = v * 10 + (pSrc(p) - 48): digits = digits + 1: p = p + 1
    Loop
    If digits = 0 Then Exit Function
    lo = v: hi = v
    If p < pLen Then
        If pSrc(p) = 44 Then
            p = p + 1
            v = 0: digits = 0
            Do While p < pLen
                If pSrc(p) < 48 Or pSrc(p) > 57 Then Exit Do
                v = v * 10 + (pSrc(p) - 48): digits = digits + 1: p = p + 1
            Loop
            If digits = 0 Then hi = -1 Else hi = v
        End If
    End If
    If p >= pLen Then Exit Function
    If pSrc(p) <> 125 Then Exit Function
    ParseBrace = True
    If Not consume Then Exit Function
    If lo > 100000 Then lo = 100000
    If hi > 100000 Then hi = 100000
    mn = CLng(lo): mx = CLng(hi)
    pPos = p + 1
End Function

Private Function ParseGroup() As Long
    Dim n As Long, c As Long, idx As Long
    pPos = pPos + 1
    If PeekC() = 63 Then
        c = -1
        If pPos + 1 < pLen Then c = pSrc(pPos + 1)
        Select Case c
        Case 58
            pPos = pPos + 2
            n = NewNode(N_GROUP)
            nA(n) = 0
            AddChild n, ParseDisjunction()
        Case 61, 33
            pPos = pPos + 2
            If c = 61 Then n = NewNode(N_LOOK) Else n = NewNode(N_NLOOK)
            AddChild n, ParseDisjunction()
        Case 60
            If pPos + 2 < pLen Then
                If pSrc(pPos + 2) = 61 Or pSrc(pPos + 2) = 33 Then ParseFail "Lookbehind assertions are not supported"
            End If
            ParseFail "Named capture groups are not supported"
        Case Else
            ParseFail "Invalid group"
        End Select
    Else
        pGroups = pGroups + 1
        idx = pGroups
        n = NewNode(N_GROUP)
        nA(n) = idx
        AddChild n, ParseDisjunction()
    End If
    If PeekC() <> 41 Then ParseFail "Unterminated group"
    pPos = pPos + 1
    ParseGroup = n
End Function

' pPos at the backslash of an atom escape (not \b or \B).
Private Function ParseAtomEscape() As Long
    Dim c As Long, n As Long, v As Double, p As Long
    pPos = pPos + 1
    If pPos >= pLen Then ParseFail "\ at end of pattern"
    c = pSrc(pPos)
    Select Case c
    Case 100, 68, 115, 83, 119, 87
        pPos = pPos + 1
        TReset
        AddBuiltin c
        n = NewNode(N_CLASS)
        nA(n) = FinishClass(False)
        ParseAtomEscape = n
        Exit Function
    Case 112, 80
        If pUnicode Then
            TReset
            AddProperty
            n = NewNode(N_CLASS)
            nA(n) = FinishClass(False)
            ParseAtomEscape = n
            Exit Function
        End If
    Case 49 To 57
        p = pPos
        v = 0
        Do While p < pLen
            If pSrc(p) < 48 Or pSrc(p) > 57 Then Exit Do
            v = v * 10 + (pSrc(p) - 48)
            p = p + 1
        Loop
        If v <= pTotalGroups Then
            pPos = p
            n = NewNode(N_BACKREF)
            nA(n) = CLng(v)
            ParseAtomEscape = n
            Exit Function
        End If
        If pUnicode Then ParseFail "Invalid escape"
    End Select
    ParseAtomEscape = CharNode(ParseCharEscape(False))
End Function

' pPos at the character after a backslash. Returns the code unit and advances.
Private Function ParseCharEscape(ByVal inClass As Boolean) As Long
    Dim c As Long, v As Long, d As Long, k As Long
    c = pSrc(pPos)
    pPos = pPos + 1
    Select Case c
    Case 110: ParseCharEscape = 10
    Case 114: ParseCharEscape = 13
    Case 116: ParseCharEscape = 9
    Case 102: ParseCharEscape = 12
    Case 118: ParseCharEscape = 11
    Case 98
        If inClass Then ParseCharEscape = 8 Else ParseCharEscape = 98
    Case 45
        ParseCharEscape = 45
    Case 99
        If pPos < pLen Then
            d = pSrc(pPos)
            If (d >= 65 And d <= 90) Or (d >= 97 And d <= 122) Then
                pPos = pPos + 1
                ParseCharEscape = d Mod 32
                Exit Function
            End If
            If inClass And ((d >= 48 And d <= 57) Or d = 95) And Not pUnicode Then
                pPos = pPos + 1
                ParseCharEscape = d Mod 32
                Exit Function
            End If
        End If
        If pUnicode Then ParseFail "Invalid unicode escape"
        pPos = pPos - 1
        ParseCharEscape = 92
    Case 120
        v = HexAt(pPos, 2)
        If v >= 0 Then
            pPos = pPos + 2
            ParseCharEscape = v
        Else
            If pUnicode Then ParseFail "Invalid escape"
            ParseCharEscape = 120
        End If
    Case 117
        If pUnicode And PeekC() = 123 Then
            v = 0: k = pPos + 1
            Do While k < pLen
                d = HexDigit(pSrc(k))
                If d < 0 Then Exit Do
                v = v * 16 + d
                If v > &H10FFFF Then ParseFail "Invalid Unicode escape"
                k = k + 1
            Loop
            If k >= pLen Or k = pPos + 1 Then ParseFail "Invalid Unicode escape"
            If pSrc(k) <> 125 Then ParseFail "Invalid Unicode escape"
            pPos = k + 1
            If v > 65535 Then ParseFail "Astral code points are not supported"
            ParseCharEscape = v
            Exit Function
        End If
        v = HexAt(pPos, 4)
        If v >= 0 Then
            pPos = pPos + 4
            ParseCharEscape = v
        Else
            If pUnicode Then ParseFail "Invalid Unicode escape"
            ParseCharEscape = 117
        End If
    Case 48 To 55
        If c = 48 And (pPos >= pLen Or pSrc(pPos) < 48 Or pSrc(pPos) > 57) Then
            ParseCharEscape = 0
            Exit Function
        End If
        If pUnicode Then ParseFail "Invalid decimal escape"
        ' Legacy octal escape (up to three digits, at most 0377)
        v = c - 48
        If pPos < pLen Then
            d = pSrc(pPos)
            If d >= 48 And d <= 55 Then
                v = v * 8 + (d - 48)
                pPos = pPos + 1
                If c <= 51 And pPos < pLen Then
                    d = pSrc(pPos)
                    If d >= 48 And d <= 55 Then
                        v = v * 8 + (d - 48)
                        pPos = pPos + 1
                    End If
                End If
            End If
        End If
        ParseCharEscape = v
    Case Else
        If pUnicode Then
            If InStr(1, "^$\.*+?()[]{}|/", ChrW$(c), vbBinaryCompare) = 0 Then ParseFail "Invalid escape"
        End If
        ParseCharEscape = c
    End Select
End Function

Private Function HexDigit(ByVal c As Long) As Long
    Select Case c
    Case 48 To 57: HexDigit = c - 48
    Case 65 To 70: HexDigit = c - 55
    Case 97 To 102: HexDigit = c - 87
    Case Else: HexDigit = -1
    End Select
End Function

Private Function HexAt(ByVal p As Long, ByVal n As Long) As Long
    Dim i As Long, v As Long, d As Long
    If p + n > pLen Then
        HexAt = -1
        Exit Function
    End If
    For i = 0 To n - 1
        d = HexDigit(pSrc(p + i))
        If d < 0 Then
            HexAt = -1
            Exit Function
        End If
        v = v * 16 + d
    Next i
    HexAt = v
End Function

' ---- Character classes ----

Private Function ParseClass() As Long
    Dim negate As Boolean, lo As Long, hi As Long, isSet As Boolean, isSet2 As Boolean, n As Long
    pPos = pPos + 1
    If PeekC() = 94 Then
        negate = True
        pPos = pPos + 1
    End If
    TReset
    Do
        If pPos >= pLen Then ParseFail "Unterminated character class"
        If pSrc(pPos) = 93 Then
            pPos = pPos + 1
            Exit Do
        End If
        lo = ClassAtom(isSet)
        If pPos + 1 < pLen Then
            If pSrc(pPos) = 45 And pSrc(pPos + 1) <> 93 Then
                pPos = pPos + 1
                hi = ClassAtom(isSet2)
                If isSet Or isSet2 Then
                    If pUnicode Then ParseFail "Invalid character class"
                    If Not isSet Then TAdd lo, lo
                    TAdd 45, 45
                    If Not isSet2 Then TAdd hi, hi
                Else
                    If lo > hi Then ParseFail "Range out of order in character class"
                    TAdd lo, hi
                End If
                GoTo NextAtom
            End If
        End If
        If Not isSet Then TAdd lo, lo
NextAtom:
    Loop
    n = NewNode(N_CLASS)
    nA(n) = FinishClass(negate)
    ParseClass = n
End Function

' Reads one class atom. For \d \s \w (and negations, \p{L}) adds the ranges and sets isSet.
Private Function ClassAtom(ByRef isSet As Boolean) As Long
    Dim c As Long
    isSet = False
    c = pSrc(pPos)
    If c <> 92 Then
        pPos = pPos + 1
        ClassAtom = c
        Exit Function
    End If
    pPos = pPos + 1
    If pPos >= pLen Then ParseFail "\ at end of pattern"
    c = pSrc(pPos)
    Select Case c
    Case 100, 68, 115, 83, 119, 87
        pPos = pPos + 1
        AddBuiltin c
        isSet = True
        Exit Function
    Case 112, 80
        If pUnicode Then
            AddProperty
            isSet = True
            Exit Function
        End If
    End Select
    ClassAtom = ParseCharEscape(True)
End Function

' pPos at p/P of \p{...}
Private Sub AddProperty()
    Dim negate As Boolean, name As String, k As Long
    negate = (pSrc(pPos) = 80)
    pPos = pPos + 1
    If PeekC() <> 123 Then ParseFail "Invalid property name"
    k = pPos + 1
    Do While k < pLen
        If pSrc(k) = 125 Then Exit Do
        name = name & ChrW$(pSrc(k))
        k = k + 1
    Loop
    If k >= pLen Then ParseFail "Invalid property name"
    pPos = k + 1
    Select Case name
    Case "L", "Letter", "General_Category=L", "gc=L", "General_Category=Letter", "gc=Letter"
        AddLetters negate
    Case Else
        ParseFail "Unsupported property name \p{" & name & "}"
    End Select
End Sub

Private Sub AddLetters(ByVal negate As Boolean)
    Dim s As String, k As Long, lo As Long, hi As Long, prev As Long
    s = LetterData()
    prev = 0
    For k = 1 To Len(s) Step 8
        lo = CLng("&H" & Mid$(s, k, 4))
        hi = CLng("&H" & Mid$(s, k + 4, 4))
        If negate Then
            If lo > prev Then TAdd prev, lo - 1
            prev = hi + 1
        Else
            TAdd lo, hi
        End If
    Next k
    If negate And prev <= 65535 Then TAdd prev, 65535
End Sub

Private Sub AddBuiltin(ByVal c As Long)
    Select Case c
    Case 100 ' \d
        TAdd 48, 57
    Case 68  ' \D
        TAdd 0, 47: TAdd 58, 65535
    Case 119 ' \w
        TAdd 48, 57: TAdd 65, 90: TAdd 95, 95: TAdd 97, 122
    Case 87  ' \W
        TAdd 0, 47: TAdd 58, 64: TAdd 91, 94: TAdd 96, 96: TAdd 123, 65535
    Case 115 ' \s
        TAdd 9, 13: TAdd 32, 32: TAdd 160, 160: TAdd 5760, 5760: TAdd 8192, 8202
        TAdd 8232, 8233: TAdd 8239, 8239: TAdd 8287, 8287: TAdd 12288, 12288: TAdd 65279, 65279
    Case 83  ' \S
        TAdd 0, 8: TAdd 14, 31: TAdd 33, 159: TAdd 161, 5759: TAdd 5761, 8191
        TAdd 8203, 8231: TAdd 8234, 8238: TAdd 8240, 8286: TAdd 8288, 12287
        TAdd 12289, 65278: TAdd 65280, 65535
    End Select
End Sub

Private Sub TReset()
    tN = 0
    ReDim tLo(0 To 15)
    ReDim tHi(0 To 15)
End Sub

Private Sub TAdd(ByVal lo As Long, ByVal hi As Long)
    If tN > UBound(tLo) Then
        ReDim Preserve tLo(0 To tN * 2 + 1)
        ReDim Preserve tHi(0 To tN * 2 + 1)
    End If
    tLo(tN) = lo
    tHi(tN) = hi
    tN = tN + 1
End Sub

' Sorts and merges the builder ranges.
Private Sub TNormalize()
    Dim i As Long, j As Long, lo As Long, hi As Long, m As Long
    For i = 1 To tN - 1
        lo = tLo(i): hi = tHi(i)
        j = i - 1
        Do While j >= 0
            If tLo(j) <= lo Then Exit Do
            tLo(j + 1) = tLo(j): tHi(j + 1) = tHi(j)
            j = j - 1
        Loop
        tLo(j + 1) = lo: tHi(j + 1) = hi
    Next i
    If tN = 0 Then Exit Sub
    m = 0
    For i = 1 To tN - 1
        If tLo(i) <= tHi(m) + 1 Then
            If tHi(i) > tHi(m) Then tHi(m) = tHi(i)
        Else
            m = m + 1
            tLo(m) = tLo(i): tHi(m) = tHi(i)
        End If
    Next i
    tN = m + 1
End Sub

' Whether c lies in the first n (sorted, merged) builder ranges.
Private Function InRanges(ByVal c As Long, ByVal n As Long) As Boolean
    Dim lo As Long, hi As Long, md As Long
    lo = 0: hi = n - 1
    Do While lo <= hi
        md = (lo + hi) \ 2
        If c < tLo(md) Then
            hi = md - 1
        ElseIf c > tHi(md) Then
            lo = md + 1
        Else
            InRanges = True
            Exit Function
        End If
    Loop
End Function

' Index of the first cased code unit >= c (mCasedCount when none).
Private Function FirstCasedAtOrAbove(ByVal c As Long) As Long
    Dim lo As Long, hi As Long, md As Long
    lo = 0: hi = mCasedCount
    Do While lo < hi
        md = (lo + hi) \ 2
        If mCased(md) < c Then lo = md + 1 Else hi = md
    Loop
    FirstCasedAtOrAbove = lo
End Function

' Finalises the builder ranges into a class (case closure, negation); returns its index.
Private Function FinishClass(ByVal negate As Boolean) As Long
    Dim i As Long, c As Long, k As Long, n0 As Long, lo As Long, hi As Long, prev As Long
    Dim nLo() As Long, nHi() As Long, m As Long, j As Long
    TNormalize
    If pIgnoreCase Then
        n0 = tN
        For i = 0 To n0 - 1
            lo = tLo(i): hi = tHi(i)
            If hi > 65535 Then hi = 65535
            ' Only code units with case partners matter: walk them within lo..hi.
            j = FirstCasedAtOrAbove(lo)
            Do While j < mCasedCount
                c = mCased(j)
                If c > hi Then Exit Do
                k = gRxFoldNext(c)
                Do While k <> c
                    If Not InRanges(k, n0) Then TAdd k, k
                    k = gRxFoldNext(k)
                Loop
                j = j + 1
            Loop
        Next i
        TNormalize
    End If
    If negate Then
        ReDim nLo(0 To tN + 1)
        ReDim nHi(0 To tN + 1)
        m = 0: prev = 0
        For i = 0 To tN - 1
            If tLo(i) > prev Then
                nLo(m) = prev: nHi(m) = tLo(i) - 1: m = m + 1
            End If
            prev = tHi(i) + 1
        Next i
        If prev <= 65535 Then
            nLo(m) = prev: nHi(m) = 65535: m = m + 1
        End If
        tN = 0
        For i = 0 To m - 1
            TAdd nLo(i), nHi(i)
        Next i
    End If
    ' store
    If cCount > UBound(cOff) Then
        ReDim Preserve cOff(0 To cCount * 2 + 1)
        ReDim Preserve cCnt(0 To cCount * 2 + 1)
    End If
    cOff(cCount) = cRanges
    cCnt(cCount) = tN
    For i = 0 To tN - 1
        If cRanges > UBound(cLo) Then
            ReDim Preserve cLo(0 To cRanges * 2 + 1)
            ReDim Preserve cHi(0 To cRanges * 2 + 1)
        End If
        cLo(cRanges) = tLo(i)
        cHi(cRanges) = tHi(i)
        cRanges = cRanges + 1
    Next i
    FinishClass = cCount
    cCount = cCount + 1
End Function

' =====================================================================
' Code generation
' =====================================================================

Private Function Emit(ByVal op As Long, Optional ByVal a As Long = 0, Optional ByVal b As Long = 0, _
        Optional ByVal c As Long = 0, Optional ByVal d As Long = 0, Optional ByVal e As Long = 0) As Long
    Dim n As Long
    If oCount > UBound(oOp) Then
        If oCount >= MAX_PROGRAM Then ParseFail "Regular expression too large"
        n = UBound(oOp) * 2 + 1
        ReDim Preserve oOp(0 To n): ReDim Preserve oA(0 To n): ReDim Preserve oB(0 To n)
        ReDim Preserve oC(0 To n): ReDim Preserve oD(0 To n): ReDim Preserve oE(0 To n)
    End If
    oOp(oCount) = op: oA(oCount) = a: oB(oCount) = b: oC(oCount) = c: oD(oCount) = d: oE(oCount) = e
    Emit = oCount
    oCount = oCount + 1
End Function

Private Function Fold(ByVal c As Long) As Long
    If pIgnoreCase And c >= 0 And c <= 65535 Then Fold = gRxCanon(c) Else Fold = c
End Function

Private Sub EmitNode(ByVal n As Long)
    Dim pc As Long
    Select Case nKind(n)
    Case N_CHAR
        Emit RX_CHAR, Fold(nA(n))
    Case N_ANY
        If pDotAll Then Emit RX_ANYALL Else Emit RX_ANY
    Case N_CLASS
        Emit RX_CLASS, nA(n)
    Case N_SEQ
        EmitSeq n
    Case N_ALT
        EmitAlt n
    Case N_GROUP
        If nA(n) > 0 Then
            Emit RX_SAVE, 2 * nA(n)
            EmitNode nFirst(n)
            Emit RX_SAVE, 2 * nA(n) + 1
        Else
            EmitNode nFirst(n)
        End If
    Case N_REPEAT
        EmitRepeat n
    Case N_BOL
        Emit RX_BOL, IIf(pMultiline, 1, 0)
    Case N_EOL
        Emit RX_EOL, IIf(pMultiline, 1, 0)
    Case N_WORDB
        Emit RX_WORDB
    Case N_NWORDB
        Emit RX_NWORDB
    Case N_LOOK, N_NLOOK
        pc = Emit(RX_LOOK, IIf(nKind(n) = N_NLOOK, 1, 0))
        oB(pc) = pc + 1
        EmitNode nFirst(n)
        Emit RX_LOOKEND
        oC(pc) = oCount
    Case N_BACKREF
        Emit RX_BACKREF, nA(n)
    End Select
End Sub

Private Sub EmitSeq(ByVal n As Long)
    Dim ch As Long, run As Long, k As Long, first As Long
    ch = nFirst(n)
    Do While ch >= 0
        If nKind(ch) = N_CHAR Then
            first = ch
            run = 0
            k = ch
            Do While k >= 0
                If nKind(k) <> N_CHAR Then Exit Do
                run = run + 1
                k = nNext(k)
            Loop
            If run = 1 Then
                EmitNode ch
            Else
                If oLitCount + run > UBound(oLit) Then ReDim Preserve oLit(0 To (oLitCount + run) * 2)
                Emit RX_STR, oLitCount, run
                k = first
                Do While k >= 0
                    If nKind(k) <> N_CHAR Then Exit Do
                    oLit(oLitCount) = Fold(nA(k))
                    oLitCount = oLitCount + 1
                    k = nNext(k)
                Loop
            End If
            ch = k
        Else
            EmitNode ch
            ch = nNext(ch)
        End If
    Loop
End Sub

Private Sub EmitAlt(ByVal n As Long)
    Dim ch As Long, sp As Long, jumps() As Long, nj As Long, i As Long
    ReDim jumps(0 To 7)
    ch = nFirst(n)
    Do While ch >= 0
        If nNext(ch) < 0 Then
            EmitNode ch
        Else
            sp = Emit(RX_SPLIT)
            oA(sp) = sp + 1
            EmitNode ch
            If nj > UBound(jumps) Then ReDim Preserve jumps(0 To nj * 2 + 1)
            jumps(nj) = Emit(RX_JMP)
            nj = nj + 1
            oB(sp) = oCount
        End If
        ch = nNext(ch)
    Loop
    For i = 0 To nj - 1
        oA(jumps(i)) = oCount
    Next i
End Sub

Private Sub EmitRepeat(ByVal n As Long)
    Dim child As Long, mn As Long, mx As Long, greedy As Boolean, capLo As Long, capHi As Long
    Dim kind As Long, arg As Long, hasCaps As Boolean, canEmpty As Boolean, i As Long
    Dim sp As Long, body As Long, r As Long, splits() As Long, bodies() As Long, cnt As Long
    child = nFirst(n): mn = nA(n): mx = nB(n): greedy = (nC(n) = 1)
    capLo = nD(n): capHi = nE(n)
    If mx = 0 Then Exit Sub
    hasCaps = (capHi >= capLo)
    If Not hasCaps Then
        If SingleAtom(child, kind, arg) Then
            Emit RX_REPS, kind, arg, mn, mx, IIf(greedy, 1, 0)
            Exit Sub
        End If
    End If
    canEmpty = (MinLen(child) = 0)
    For i = 1 To mn
        If hasCaps Then Emit RX_CLEAR, capLo, capHi
        EmitNode child
    Next i
    If mx < 0 Then
        sp = Emit(RX_SPLIT)
        body = oCount
        If canEmpty Then
            r = oRegs: oRegs = oRegs + 1
            Emit RX_MARK, r
        End If
        If hasCaps Then Emit RX_CLEAR, capLo, capHi
        EmitNode child
        If canEmpty Then Emit RX_CHECK, r
        Emit RX_JMP, sp
        If greedy Then
            oA(sp) = body: oB(sp) = oCount
        Else
            oA(sp) = oCount: oB(sp) = body
        End If
    Else
        cnt = mx - mn
        ReDim splits(0 To cnt)
        ReDim bodies(0 To cnt)
        For i = 0 To cnt - 1
            splits(i) = Emit(RX_SPLIT)
            bodies(i) = oCount
            If canEmpty Then
                r = oRegs: oRegs = oRegs + 1
                Emit RX_MARK, r
            End If
            If hasCaps Then Emit RX_CLEAR, capLo, capHi
            EmitNode child
            If canEmpty Then Emit RX_CHECK, r
        Next i
        For i = 0 To cnt - 1
            If greedy Then
                oA(splits(i)) = bodies(i): oB(splits(i)) = oCount
            Else
                oA(splits(i)) = oCount: oB(splits(i)) = bodies(i)
            End If
        Next i
    End If
End Sub

Private Function SingleAtom(ByVal n As Long, ByRef kind As Long, ByRef arg As Long) As Boolean
    Select Case nKind(n)
    Case N_CHAR
        kind = RXA_CHAR: arg = Fold(nA(n)): SingleAtom = True
    Case N_ANY
        If pDotAll Then kind = RXA_ALL Else kind = RXA_ANY
        arg = 0: SingleAtom = True
    Case N_CLASS
        kind = RXA_CLASS: arg = nA(n): SingleAtom = True
    Case N_GROUP
        If nA(n) = 0 Then SingleAtom = SingleAtom(nFirst(n), kind, arg)
    Case N_SEQ
        If nFirst(n) >= 0 Then
            If nNext(nFirst(n)) < 0 Then SingleAtom = SingleAtom(nFirst(n), kind, arg)
        End If
    End Select
End Function

Private Function MinLen(ByVal n As Long) As Double
    Dim ch As Long, total As Double, best As Double, v As Double
    Select Case nKind(n)
    Case N_CHAR, N_ANY, N_CLASS
        MinLen = 1
    Case N_SEQ
        ch = nFirst(n)
        Do While ch >= 0
            total = total + MinLen(ch)
            ch = nNext(ch)
        Loop
        MinLen = total
    Case N_ALT
        best = -1
        ch = nFirst(n)
        Do While ch >= 0
            v = MinLen(ch)
            If best < 0 Or v < best Then best = v
            ch = nNext(ch)
        Loop
        If best < 0 Then best = 0
        MinLen = best
    Case N_GROUP
        MinLen = MinLen(nFirst(n))
    Case N_REPEAT
        MinLen = nA(n) * MinLen(nFirst(n))
    Case Else
        MinLen = 0
    End Select
End Function

Private Function CanonData() As String
    Dim s As String
    s = ""
    s = s & "00611AFFE0100B50102E7100E017FFE0100F807FFE0100FF0100791010118FFFF2013303FFFF2013A08FFFF2014B17FFFF2017A03FFFF201800100C31018302FFFF2018801FFFF1018C01FFFF1019201FFFF101950100611019901FFFF1019A0100A31019B01A6411019E010082101A103FFFF201A801FFFF101AD01FFFF101B001FFFF101B402FFFF201B901FFFF101BD01FFFF101BF010038101C501FFFF101C601FFFE101C801FFFF101C901FFFE101CB01FFFF101CC01FFFE101CE08FFFF201DD01FFB1101DF09FFFF201F201FFFF101F301FFFE101F501FFFF101F914FFFF2022309FFFF2023C01FFFF1023F022A3F1024201FFFF1024705FFFF20250012A1F10251012A1C10252012A1E1025301FF2E1025401FF321025602FF331025901FF361025B01FF351025C01A54F1026001FF331026101A54B1026301FF311026401A5671026501A5281026601A5441026801FF2F1026901FF2D1026A01A5441026B0129F71026C01A5411026F01FF2D102710129FD1027201FF2B1027501FF2A1027D0129E71028001FF261028201A5431028301FF261028701A52A1028801FF261028901FFBB1028A02FF271028C01FFB91029201FF251"
    s = s & "029D01A5151029E01A512103450100541037102FFFF2037701FFFF1037B030082103AC01FFDA103AD03FFDB103B111FFE0103C201FFE1103C309FFE0103CC01FFC0103CD02FFC1103D001FFC2103D101FFC7103D501FFD1103D601FFCA103D701FFF8103D90CFFFF203F001FFAA103F101FFB0103F2010007103F301FF8C103F501FFA0103F801FFFF103FB01FFFF1043020FFE01045010FFB01046111FFFF2048B1BFFFF204C207FFFF204CF01FFF1104D130FFFF2056126FFD0110D02B0BC0110FD030BC0113F806FFF811C8001E79211C8101E79311C8201E79C11C8302E79E11C8501E79D11C8601E7A411C8701E7DB11C880189C211C8A01FFFF11D79018A0411D7D010EE611D8E018A3811E014BFFFF21E9B01FFC511EA130FFFF21F0008000811F1006000811F2008000811F3008000811F4006000811F5104000821F6008000811F7002004A11F7204005611F7602006411F7802008011F7A02007011F7C02007E11FB002000811FBE01E3DB11FD002000811FE002000811FE50100071214E01FFE41217010FFF01218401FFFF124D01AFFE612C3030FFD012C6101FFFF12C6501D5D512C6601D5D812C6803FFFF22C7301FFFF1"
    s = s & "2C7601FFFF12C8132FFFF22CEC02FFFF22CF301FFFF12D0026E3A012D2701E3A012D2D01E3A01A64117FFFF2A6810EFFFF2A72307FFFF2A7331FFFFF2A77A02FFFF2A77F05FFFF2A78C01FFFF1A79102FFFF2A7940100301A7970AFFFF2A7B508FFFF2A7C802FFFF2A7CD01FFFF1A7D101FFFF1A7D703FFFF2A7F601FFFF1AB5301FC601AB705068301FF411AFFE01"
    CanonData = s
End Function

Private Function LetterData() As String
    Dim s As String
    s = ""
    s = s & "0041005A0061007A00AA00AA00B500B500BA00BA00C000D600D800F600F802C102C602D102E002E402EC02EC02EE02EE0370037403760377037A037D037F037F038603860388038A038C038C038E03A103A303F503F70481048A052F05310556055905590560058805D005EA05EF05F20620064A066E066F067106D306D506D506E506E606EE06EF06FA06FC06FF06FF071007100712072F074D07A507B107B107CA07EA07F407F507FA07FA08000815081A081A0824082408280828084008580860086A087008870889088E08A008C909040939093D093D0950095009580961097109800985098C098F0990099309A809AA09B009B209B209B609B909BD09BD09CE09CE09DC09DD09DF09E109F009F109FC09FC0A050A0A0A0F0A100A130A280A2A0A300A320A330A350A360A380A390A590A5C0A5E0A5E0A720A740A850A8D0A8F0A910A930AA80AAA0AB00AB20AB30AB50AB90ABD0ABD0AD00AD00AE00AE10AF90AF90B050B0C0B0F0B100B130B280B2A0B300B320B330B350B390B3D0B3D0B5C0B5D0B5F0B610B710B710B830B830B850B8A0B8E0B900B920B950B990B9A0B9C0B9C0B9E0B9F0BA30BA40BA80BAA0BAE0BB90BD00BD0"
    s = s & "0C050C0C0C0E0C100C120C280C2A0C390C3D0C3D0C580C5A0C5D0C5D0C600C610C800C800C850C8C0C8E0C900C920CA80CAA0CB30CB50CB90CBD0CBD0CDD0CDE0CE00CE10CF10CF20D040D0C0D0E0D100D120D3A0D3D0D3D0D4E0D4E0D540D560D5F0D610D7A0D7F0D850D960D9A0DB10DB30DBB0DBD0DBD0DC00DC60E010E300E320E330E400E460E810E820E840E840E860E8A0E8C0EA30EA50EA50EA70EB00EB20EB30EBD0EBD0EC00EC40EC60EC60EDC0EDF0F000F000F400F470F490F6C0F880F8C1000102A103F103F10501055105A105D1061106110651066106E107010751081108E108E10A010C510C710C710CD10CD10D010FA10FC1248124A124D1250125612581258125A125D12601288128A128D129012B012B212B512B812BE12C012C012C212C512C812D612D81310131213151318135A1380138F13A013F513F813FD1401166C166F167F1681169A16A016EA16F116F817001711171F1731174017511760176C176E1770178017B317D717D717DC17DC1820187818801884188718A818AA18AA18B018F51900191E1950196D19701974198019AB19B019C91A001A161A201A541AA71AA71B051B331B451B4C1B831BA0"
    s = s & "1BAE1BAF1BBA1BE51C001C231C4D1C4F1C5A1C7D1C801C8A1C901CBA1CBD1CBF1CE91CEC1CEE1CF31CF51CF61CFA1CFA1D001DBF1E001F151F181F1D1F201F451F481F4D1F501F571F591F591F5B1F5B1F5D1F5D1F5F1F7D1F801FB41FB61FBC1FBE1FBE1FC21FC41FC61FCC1FD01FD31FD61FDB1FE01FEC1FF21FF41FF61FFC20712071207F207F2090209C2102210221072107210A2113211521152119211D212421242126212621282128212A212D212F2139213C213F21452149214E214E218321842C002CE42CEB2CEE2CF22CF32D002D252D272D272D2D2D2D2D302D672D6F2D6F2D802D962DA02DA62DA82DAE2DB02DB62DB82DBE2DC02DC62DC82DCE2DD02DD62DD82DDE2E2F2E2F3005300630313035303B303C30413096309D309F30A130FA30FC30FF3105312F3131318E31A031BF31F031FF34004DBF4E00A48CA4D0A4FDA500A60CA610A61FA62AA62BA640A66EA67FA69DA6A0A6E5A717A71FA722A788A78BA7CDA7D0A7D1A7D3A7D3A7D5A7DCA7F2A801A803A805A807A80AA80CA822A840A873A882A8B3A8F2A8F7A8FBA8FBA8FDA8FEA90AA925A930A946A960A97CA984A9B2A9CFA9CFA9E0A9E4A9E6A9EFA9FAA9FE"
    s = s & "AA00AA28AA40AA42AA44AA4BAA60AA76AA7AAA7AAA7EAAAFAAB1AAB1AAB5AAB6AAB9AABDAAC0AAC0AAC2AAC2AADBAADDAAE0AAEAAAF2AAF4AB01AB06AB09AB0EAB11AB16AB20AB26AB28AB2EAB30AB5AAB5CAB69AB70ABE2AC00D7A3D7B0D7C6D7CBD7FBF900FA6DFA70FAD9FB00FB06FB13FB17FB1DFB1DFB1FFB28FB2AFB36FB38FB3CFB3EFB3EFB40FB41FB43FB44FB46FBB1FBD3FD3DFD50FD8FFD92FDC7FDF0FDFBFE70FE74FE76FEFCFF21FF3AFF41FF5AFF66FFBEFFC2FFC7FFCAFFCFFFD2FFD7FFDAFFDC"
    LetterData = s
End Function
