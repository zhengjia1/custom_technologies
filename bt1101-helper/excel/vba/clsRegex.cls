VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRegex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' A compiled JavaScript regular expression. Create with modRegex.Rx(pattern, flags).
' The object is stateless between calls (there is no lastIndex), so cached instances
' can be shared. Indices are 0-based UTF-16 code-unit offsets, as in JavaScript.
'
' Matching runs a backtracking machine with an explicit stack: no recursion except
' one level per nested lookahead, so long inputs cannot exhaust the VBA call stack.
Option Explicit

Private Const BT_BRANCH As Long = 1
Private Const BT_CAP As Long = 2
Private Const BT_REG As Long = 3
Private Const BT_GREEDY As Long = 4
Private Const BT_LAZY As Long = 5

Private Const PRE_NONE As Long = 0
Private Const PRE_LITERAL As Long = 1
Private Const PRE_CANON As Long = 2
Private Const PRE_CLASS As Long = 3
Private Const PRE_LINESTART As Long = 4
Private Const PRE_ANCHORED As Long = 5
Private Const PRE_SET As Long = 6

Private mSource As String
Private mFlags As String
Private mGlobal As Boolean
Private mIgnoreCase As Boolean
Private mMultiline As Boolean
Private mUnicode As Boolean
Private mSticky As Boolean
Private mGroups As Long

Private mOp() As Long
Private mA() As Long
Private mB() As Long
Private mC() As Long
Private mD() As Long
Private mGreedy() As Long
Private mLit() As Long
Private mClsOff() As Long
Private mClsCnt() As Long
Private mLo() As Long
Private mHi() As Long
Private mAscii() As Boolean

Private mPre As Long
Private mPreStr As String
Private mPreArg As Long
Private mFirstAscii() As Boolean
Private mFirstHigh As Boolean
Private mSetAscii() As Boolean
Private mSetHigh As Boolean
Private mFollowIdx() As Long      ' per pc: follow-set index of an RX_REPS, or -1
Private mFollow() As Boolean       ' follow set * 128 + ASCII code
Private mFollowHigh() As Boolean   ' follow set accepts non-ASCII

' Matching state
Private mCaps() As Long
Private mRegs() As Long
Private mStk() As Long
Private mTop As Long
Private mStkCap As Long
Private mEnd As Long

' Parts buffer for Replace
Private mParts() As String
Private mPartCount As Long

' =====================================================================
' Properties
' =====================================================================

Public Property Get Source() As String
    Source = mSource
End Property

Public Property Get Flags() As String
    Flags = mFlags
End Property

' The g flag (Global is a VBA keyword).
Public Property Get IsGlobal() As Boolean
    IsGlobal = mGlobal
End Property

Public Property Get IgnoreCase() As Boolean
    IgnoreCase = mIgnoreCase
End Property

Public Property Get Multiline() As Boolean
    Multiline = mMultiline
End Property

Public Property Get Unicode() As Boolean
    Unicode = mUnicode
End Property

Public Property Get Sticky() As Boolean
    Sticky = mSticky
End Property

' Number of capturing groups.
Public Property Get GroupCount() As Long
    GroupCount = mGroups
End Property

' =====================================================================
' Public methods
' =====================================================================

' True when the pattern matches anywhere in text (ignores the g flag: no lastIndex).
Public Function Test(ByVal text As String) As Boolean
    RxLoadText text
    Test = (Search(0) >= 0)
End Function

' First match at or after start (0-based), or Nothing.
Public Function Exec(ByVal text As String, Optional ByVal start As Long = 0) As clsMatch
    Dim s As Long
    RxLoadText text
    If start < 0 Then start = 0
    s = Search(start)
    If s < 0 Then Exit Function
    Set Exec = NewMatch(text, Nothing)
End Function

' Every match, like [...text.matchAll(/re/g)] (whether or not the g flag is set).
Public Function MatchAll(ByVal text As String) As Collection
    Dim col As Collection, pos As Long, s As Long, e As Long, m As clsMatch, owner As clsMatch
    Set col = New Collection
    RxLoadText text
    pos = 0
    Do While pos <= gRxLen
        s = Search(pos)
        If s < 0 Then Exit Do
        e = mCaps(1)
        Set m = NewMatch(text, owner)
        If owner Is Nothing Then Set owner = m
        col.Add m
        If e = s Then pos = AdvanceIndex(e) Else pos = e
    Loop
    Set MatchAll = col
End Function

' String.prototype.replace(re, replacement): every match with g, else the first.
' Supports $1..$99, $&, $$, $` and $' in the replacement.
Public Function Replace(ByVal text As String, ByVal replacement As String) As String
    Dim pos As Long, s As Long, e As Long, lastEnd As Long, hasDollar As Boolean, found As Boolean
    RxLoadText text
    hasDollar = (InStr(1, replacement, "$", vbBinaryCompare) > 0)
    mPartCount = 0
    ReDim mParts(0 To 15)
    pos = 0
    lastEnd = 0
    Do
        s = Search(pos)
        If s < 0 Then Exit Do
        e = mCaps(1)
        found = True
        If s > lastEnd Then AddPart Mid$(text, lastEnd + 1, s - lastEnd)
        If hasDollar Then
            AddPart Substitute(text, replacement, s, e)
        ElseIf Len(replacement) > 0 Then
            AddPart replacement
        End If
        lastEnd = e
        If Not mGlobal Then Exit Do
        If e = s Then pos = AdvanceIndex(e) Else pos = e
        If pos > gRxLen Then Exit Do
    Loop
    If Not found Then
        Replace = text
        Exit Function
    End If
    If lastEnd < Len(text) Then AddPart Mid$(text, lastEnd + 1)
    If mPartCount = 0 Then
        Replace = ""
        Exit Function
    End If
    ReDim Preserve mParts(0 To mPartCount - 1)
    Replace = Join(mParts, "")
    Erase mParts
End Function

' String.prototype.split(re, limit). Returns a Collection of String items; a capturing
' group that did not participate gives an Empty item (JavaScript undefined).
' limit < 0 means no limit.
Public Function Split(ByVal text As String, Optional ByVal limit As Long = -1) As clsList
    Dim col As clsList, size As Long, p As Long, q As Long, s As Long, e As Long, g As Long
    Set col = New clsList
    Set Split = col
    If limit = 0 Then Exit Function
    RxLoadText text
    size = gRxLen
    If size = 0 Then
        If SearchAt(0) < 0 Then col.Add text
        Exit Function
    End If
    p = 0: q = 0
    Do While q < size
        s = SearchAt(q)
        If s < 0 Or s >= size Then Exit Do
        e = mCaps(1)
        If e > size Then e = size
        If e = p Then
            q = AdvanceIndex(s)
        Else
            col.Add Mid$(text, p + 1, s - p)
            If col.Count = limit Then Exit Function
            p = e
            For g = 1 To mGroups
                If mCaps(2 * g) < 0 Or mCaps(2 * g + 1) < 0 Then
                    col.Add Empty
                Else
                    col.Add Mid$(text, mCaps(2 * g) + 1, mCaps(2 * g + 1) - mCaps(2 * g))
                End If
                If col.Count = limit Then Exit Function
            Next g
            q = p
        End If
    Loop
    col.Add Mid$(text, p + 1)
End Function

' =====================================================================
' Construction (called by modRegex.RxCompile)
' =====================================================================

Friend Sub Init(ByVal pattern As String, ByVal flagText As String, ByVal isGlobal As Boolean, _
        ByVal isIgnoreCase As Boolean, ByVal isMultiline As Boolean, ByVal isUnicode As Boolean, _
        ByVal isSticky As Boolean, ByVal groups As Long, _
        ByRef op() As Long, ByRef a() As Long, ByRef b() As Long, ByRef c() As Long, _
        ByRef d() As Long, ByRef e() As Long, ByVal opCount As Long, _
        ByRef lit() As Long, ByVal litCount As Long, ByRef clsOff() As Long, ByRef clsCnt() As Long, _
        ByVal clsCount As Long, ByRef lo() As Long, ByRef hi() As Long, ByVal rangeCount As Long, _
        ByVal regs As Long)
    Dim i As Long, k As Long, ch As Long, top As Long
    mSource = pattern
    mFlags = flagText
    mGlobal = isGlobal
    mIgnoreCase = isIgnoreCase
    mMultiline = isMultiline
    mUnicode = isUnicode
    mSticky = isSticky
    mGroups = groups
    mOp = op: mA = a: mB = b: mC = c: mD = d: mGreedy = e
    mLit = lit
    mClsOff = clsOff: mClsCnt = clsCnt: mLo = lo: mHi = hi
    If clsCount > 0 Then
        ReDim mAscii(0 To clsCount * 128 - 1)
        For i = 0 To clsCount - 1
            For k = mClsOff(i) To mClsOff(i) + mClsCnt(i) - 1
                If mLo(k) < 128 Then
                    top = mHi(k)
                    If top > 127 Then top = 127
                    For ch = mLo(k) To top
                        mAscii(i * 128 + ch) = True
                    Next ch
                End If
            Next k
        Next i
    Else
        ReDim mAscii(0 To 0)
    End If
    ReDim mCaps(0 To 2 * groups + 1)
    If regs < 1 Then regs = 1
    ReDim mRegs(0 To regs - 1)
    mStkCap = 256
    ReDim mStk(0 To mStkCap * 4 - 1)
    SetupPrefilter
    SetupFollowSets
End Sub

' Finds a cheap test for where a match can start.
Private Sub SetupPrefilter()
    Dim pc As Long, i As Long
    mPre = PRE_NONE
    pc = 0
    Do
        Select Case mOp(pc)
        Case RX_SAVE, RX_WORDB, RX_NWORDB
            pc = pc + 1
        Case Else
            Exit Do
        End Select
    Loop
    Select Case mOp(pc)
    Case RX_BOL
        If mA(pc) = 1 Then mPre = PRE_LINESTART Else mPre = PRE_ANCHORED
    Case RX_CHAR
        SetCharPrefilter mA(pc)
    Case RX_STR
        If mIgnoreCase Then
            mPre = PRE_CANON
            mPreArg = mLit(mA(pc))
        Else
            mPre = PRE_LITERAL
            mPreStr = ""
            For i = 0 To mB(pc) - 1
                mPreStr = mPreStr & ChrW$(mLit(mA(pc) + i))
            Next i
        End If
    Case RX_CLASS
        mPre = PRE_CLASS
        mPreArg = mA(pc)
    Case RX_REPS
        If mC(pc) >= 1 Then
            If mA(pc) = RXA_CHAR Then
                SetCharPrefilter mB(pc)
            ElseIf mA(pc) = RXA_CLASS Then
                mPre = PRE_CLASS
                mPreArg = mB(pc)
            End If
        End If
    End Select
    If mPre = PRE_NONE Or mPre = PRE_CANON Then
        If BuildFirstSet(0) Then
            mPre = PRE_SET
            mFirstAscii = mSetAscii
            mFirstHigh = mSetHigh
        End If
    End If
End Sub

' For every backtracking single-character repeat, the code units its continuation can
' start with: positions where the continuation must fail are skipped.
Private Sub SetupFollowSets()
    Dim pc As Long, n As Long, c As Long
    ReDim mFollowIdx(0 To UBound(mOp))
    For pc = 0 To UBound(mOp)
        mFollowIdx(pc) = -1
        If mOp(pc) = RX_REPS Then
            If mC(pc) <> mD(pc) Then
                If BuildFirstSet(pc + 1) Then
                    If n = 0 Then
                        ReDim mFollow(0 To 127)
                        ReDim mFollowHigh(0 To 0)
                    Else
                        ReDim Preserve mFollow(0 To (n + 1) * 128 - 1)
                        ReDim Preserve mFollowHigh(0 To n)
                    End If
                    For c = 0 To 127
                        mFollow(n * 128 + c) = mSetAscii(c)
                    Next c
                    mFollowHigh(n) = mSetHigh
                    mFollowIdx(pc) = n
                    n = n + 1
                End If
            End If
        End If
    Next pc
End Sub

' Whether the continuation of the repeat at pc can start at position p.
Private Function FollowOk(ByVal pc As Long, ByVal p As Long) As Boolean
    Dim f As Long, c As Long
    f = mFollowIdx(pc)
    If f < 0 Or p >= gRxLen Then
        FollowOk = True
        Exit Function
    End If
    c = gRxCode(p)
    If c < 128 Then FollowOk = mFollow(f * 128 + c) Else FollowOk = mFollowHigh(f)
End Function

' Collects into mSetAscii/mSetHigh every code unit that a match of the program from
' startPc can start with (ASCII exactly; non-ASCII as one flag). False when it can
' start with anything or match empty.
Private Function BuildFirstSet(ByVal startPc As Long) As Boolean
    Dim visited() As Boolean, stack() As Long, sp As Long, pc As Long, n As Long, k As Long
    n = UBound(mOp)
    ReDim visited(0 To n)
    ReDim stack(0 To 2 * n + 2)
    ReDim mSetAscii(0 To 127)
    mSetHigh = False
    stack(0) = startPc: sp = 1
    Do While sp > 0
        sp = sp - 1
        pc = stack(sp)
        If Not visited(pc) Then
            visited(pc) = True
            Select Case mOp(pc)
            Case RX_SAVE, RX_WORDB, RX_NWORDB, RX_BOL, RX_EOL, RX_MARK, RX_CHECK, RX_CLEAR
                stack(sp) = pc + 1: sp = sp + 1
            Case RX_SPLIT
                stack(sp) = mA(pc): stack(sp + 1) = mB(pc): sp = sp + 2
            Case RX_JMP
                stack(sp) = mA(pc): sp = sp + 1
            Case RX_CHAR
                AddFirstChar mA(pc)
            Case RX_STR
                AddFirstChar mLit(mA(pc))
            Case RX_CLASS
                AddFirstClass mA(pc)
            Case RX_REPS
                Select Case mA(pc)
                Case RXA_CHAR
                    AddFirstChar mB(pc)
                Case RXA_CLASS
                    AddFirstClass mB(pc)
                Case Else
                    Exit Function
                End Select
                If mC(pc) = 0 Then
                    stack(sp) = pc + 1: sp = sp + 1
                End If
            Case Else
                Exit Function
            End Select
        End If
    Loop
    BuildFirstSet = True
End Function

Private Sub AddFirstChar(ByVal code As Long)
    Dim k As Long
    If code < 128 Then mSetAscii(code) = True Else mSetHigh = True
    If mIgnoreCase Then
        k = gRxFoldNext(code)
        Do While k <> code
            If k < 128 Then mSetAscii(k) = True Else mSetHigh = True
            k = gRxFoldNext(k)
        Loop
    End If
End Sub

Private Sub AddFirstClass(ByVal idx As Long)
    Dim c As Long
    For c = 0 To 127
        If mAscii(idx * 128 + c) Then mSetAscii(c) = True
    Next c
    If mClsCnt(idx) > 0 Then
        If mHi(mClsOff(idx) + mClsCnt(idx) - 1) >= 128 Then mSetHigh = True
    End If
End Sub

Private Sub SetCharPrefilter(ByVal code As Long)
    If mIgnoreCase Then
        mPre = PRE_CANON
        mPreArg = code
    Else
        mPre = PRE_LITERAL
        mPreStr = ChrW$(code)
    End If
End Sub

' =====================================================================
' Matching
' =====================================================================

' Leftmost match starting at or after from. Returns its index (captures in mCaps) or -1.
Private Function Search(ByVal from As Long) As Long
    Dim s As Long, i As Long, c As Long
    Search = -1
    If from > gRxLen Then Exit Function
    For i = 0 To 2 * mGroups + 1
        mCaps(i) = -1
    Next i
    mTop = 0
    s = from
    If mUnicode And s > 0 And s < gRxLen Then
        ' Like V8: a start inside a surrogate pair steps back to the pair.
        If gRxCode(s) >= &HDC00& And gRxCode(s) <= &HDFFF& Then
            If gRxCode(s - 1) >= &HD800& And gRxCode(s - 1) <= &HDBFF& Then s = s - 1
        End If
    End If
    If mSticky Then
        If Run(0, s) Then
            mCaps(0) = s: mCaps(1) = mEnd
            Search = s
        End If
        Exit Function
    End If
    Do
        Select Case mPre
        Case PRE_LITERAL
            i = InStr(s + 1, gRxText, mPreStr, vbBinaryCompare)
            If i = 0 Then Exit Function
            s = i - 1
        Case PRE_CANON
            Do While s < gRxLen
                If gRxCanon(gRxCode(s)) = mPreArg Then Exit Do
                s = s + 1
            Loop
            If s >= gRxLen Then Exit Function
        Case PRE_CLASS
            Do While s < gRxLen
                c = gRxCode(s)
                If c < 128 Then
                    If mAscii(mPreArg * 128 + c) Then Exit Do
                ElseIf InClassSlow(mPreArg, c) Then
                    Exit Do
                End If
                s = s + 1
            Loop
            If s >= gRxLen Then Exit Function
        Case PRE_LINESTART
            Do While s > 0 And s < gRxLen
                c = gRxCode(s - 1)
                If c = 10 Or c = 13 Or c = 8232 Or c = 8233 Then Exit Do
                s = s + 1
            Loop
        Case PRE_ANCHORED
            If s > 0 Then Exit Function
        Case PRE_SET
            Do While s < gRxLen
                c = gRxCode(s)
                If c < 128 Then
                    If mFirstAscii(c) Then Exit Do
                ElseIf mFirstHigh Then
                    Exit Do
                End If
                s = s + 1
            Loop
            If s >= gRxLen Then Exit Function
        End Select
        If Run(0, s) Then
            mCaps(0) = s: mCaps(1) = mEnd
            Search = s
            Exit Function
        End If
        s = s + 1
        If s > gRxLen Then Exit Function
        If mPre = PRE_ANCHORED Then Exit Function
    Loop
End Function

' Search ignoring the sticky flag (used by Split).
Private Function SearchAt(ByVal from As Long) As Long
    Dim saved As Boolean
    saved = mSticky
    mSticky = False
    SearchAt = Search(from)
    mSticky = saved
End Function

Private Function AdvanceIndex(ByVal i As Long) As Long
    If mUnicode And i + 1 < gRxLen Then
        If gRxCode(i) >= &HD800& And gRxCode(i) <= &HDBFF& Then
            If gRxCode(i + 1) >= &HDC00& And gRxCode(i + 1) <= &HDFFF& Then
                AdvanceIndex = i + 2
                Exit Function
            End If
        End If
    End If
    AdvanceIndex = i + 1
End Function

Private Sub Push(ByVal kind As Long, ByVal a As Long, ByVal b As Long, ByVal c As Long)
    Dim t As Long
    If mTop >= mStkCap Then
        mStkCap = mStkCap * 2
        ReDim Preserve mStk(0 To mStkCap * 4 - 1)
    End If
    t = mTop * 4
    mStk(t) = kind
    mStk(t + 1) = a
    mStk(t + 2) = b
    mStk(t + 3) = c
    mTop = mTop + 1
End Sub

Private Function InClassSlow(ByVal idx As Long, ByVal c As Long) As Boolean
    Dim lo As Long, hi As Long, md As Long
    lo = mClsOff(idx)
    hi = lo + mClsCnt(idx) - 1
    Do While lo <= hi
        md = (lo + hi) \ 2
        If c < mLo(md) Then
            hi = md - 1
        ElseIf c > mHi(md) Then
            lo = md + 1
        Else
            InClassSlow = True
            Exit Function
        End If
    Loop
End Function

Private Function IsWordAt(ByVal p As Long) As Boolean
    Dim c As Long
    If p < 0 Or p >= gRxLen Then Exit Function
    c = gRxCode(p)
    If c < 48 Or c > 122 Then Exit Function
    IsWordAt = (c <= 57) Or (c >= 65 And c <= 90) Or c = 95 Or c >= 97
End Function

' Whether the single-character atom of the RX_REPS at pc matches the code unit at p.
Private Function AtomAt(ByVal pc As Long, ByVal p As Long) As Boolean
    Dim c As Long
    c = gRxCode(p)
    Select Case mA(pc)
    Case RXA_CHAR
        If mIgnoreCase Then c = gRxCanon(c)
        AtomAt = (c = mB(pc))
    Case RXA_ANY
        AtomAt = Not (c = 10 Or c = 13 Or c = 8232 Or c = 8233)
    Case RXA_ALL
        AtomAt = True
    Case RXA_CLASS
        If c < 128 Then AtomAt = mAscii(mB(pc) * 128 + c) Else AtomAt = InClassSlow(mB(pc), c)
    End Select
End Function

' Number of consecutive atom matches from pos (at most lim).
Private Function CountAtoms(ByVal pc As Long, ByVal pos As Long, ByVal lim As Long) As Long
    Dim n As Long, c As Long, arg As Long, base As Long, last As Long
    arg = mB(pc)
    last = pos + lim
    n = pos
    Select Case mA(pc)
    Case RXA_CHAR
        If mIgnoreCase Then
            Do While n < last
                If gRxCanon(gRxCode(n)) <> arg Then Exit Do
                n = n + 1
            Loop
        Else
            Do While n < last
                If gRxCode(n) <> arg Then Exit Do
                n = n + 1
            Loop
        End If
    Case RXA_ANY
        Do While n < last
            c = gRxCode(n)
            If c = 10 Or c = 13 Or c = 8232 Or c = 8233 Then Exit Do
            n = n + 1
        Loop
    Case RXA_ALL
        n = last
    Case RXA_CLASS
        base = arg * 128
        Do While n < last
            c = gRxCode(n)
            If c < 128 Then
                If Not mAscii(base + c) Then Exit Do
            ElseIf Not InClassSlow(arg, c) Then
                Exit Do
            End If
            n = n + 1
        Loop
    End Select
    CountAtoms = n - pos
End Function

' Runs the program from pc at pos. On success mEnd holds the end position and the
' stack keeps the choice points above the entry level; on failure the stack and
' captures are restored to the entry state. Cases are ordered by frequency.
Private Function Run(ByVal pc As Long, ByVal pos As Long) As Boolean
    Dim base As Long, c As Long, ok As Boolean, k As Long, n As Long, i As Long, p As Long
    Dim mx As Long, mn As Long, cnt As Long, lim As Long, t As Long, f As Long, found As Boolean
    base = mTop
    Do
        ok = True
        Select Case mOp(pc)
        Case RX_REPS
            mx = mD(pc): mn = mC(pc)
            lim = gRxLen - pos
            If mx >= 0 And mx < lim Then lim = mx
            If mGreedy(pc) = 1 Then
                cnt = CountAtoms(pc, pos, lim)
                If cnt < mn Then
                    ok = False
                Else
                    If cnt > mn Then Push BT_GREEDY, pc, pos + mn, pos + cnt
                    pos = pos + cnt
                    If FollowOk(pc, pos) Then pc = pc + 1 Else ok = False
                End If
            ElseIf mn > lim Then
                ok = False
            Else
                If mn > 0 Then cnt = CountAtoms(pc, pos, mn) Else cnt = 0
                If cnt < mn Then
                    ok = False
                Else
                    pos = pos + mn
                    If mx < 0 Or mn < mx Then Push BT_LAZY, pc, pos, mn
                    If FollowOk(pc, pos) Then pc = pc + 1 Else ok = False
                End If
            End If
        Case RX_CHAR
            If pos < gRxLen Then
                c = gRxCode(pos)
                If mIgnoreCase Then c = gRxCanon(c)
                If c = mA(pc) Then
                    pos = pos + 1: pc = pc + 1
                Else
                    ok = False
                End If
            Else
                ok = False
            End If
        Case RX_CLASS
            If pos < gRxLen Then
                c = gRxCode(pos)
                If c < 128 Then
                    ok = mAscii(mA(pc) * 128 + c)
                Else
                    ok = InClassSlow(mA(pc), c)
                End If
                If ok Then pos = pos + 1: pc = pc + 1
            Else
                ok = False
            End If
        Case RX_SPLIT
            Push BT_BRANCH, mB(pc), pos, 0
            pc = mA(pc)
        Case RX_SAVE
            k = mA(pc)
            Push BT_CAP, k, mCaps(k), 0
            mCaps(k) = pos
            pc = pc + 1
        Case RX_JMP
            pc = mA(pc)
        Case RX_STR
            n = mB(pc)
            If pos + n > gRxLen Then
                ok = False
            Else
                k = mA(pc)
                If mIgnoreCase Then
                    For i = 0 To n - 1
                        If gRxCanon(gRxCode(pos + i)) <> mLit(k + i) Then
                            ok = False
                            Exit For
                        End If
                    Next i
                Else
                    For i = 0 To n - 1
                        If gRxCode(pos + i) <> mLit(k + i) Then
                            ok = False
                            Exit For
                        End If
                    Next i
                End If
                If ok Then pos = pos + n: pc = pc + 1
            End If
        Case RX_MATCH, RX_LOOKEND
            mEnd = pos
            Run = True
            Exit Function
        Case RX_ANY
            If pos < gRxLen Then
                c = gRxCode(pos)
                If c = 10 Or c = 13 Or c = 8232 Or c = 8233 Then
                    ok = False
                Else
                    pos = pos + 1: pc = pc + 1
                End If
            Else
                ok = False
            End If
        Case RX_EOL
            If pos = gRxLen Then
                pc = pc + 1
            ElseIf mA(pc) = 1 Then
                c = gRxCode(pos)
                If c = 10 Or c = 13 Or c = 8232 Or c = 8233 Then pc = pc + 1 Else ok = False
            Else
                ok = False
            End If
        Case RX_BOL
            If pos = 0 Then
                pc = pc + 1
            ElseIf mA(pc) = 1 Then
                c = gRxCode(pos - 1)
                If c = 10 Or c = 13 Or c = 8232 Or c = 8233 Then pc = pc + 1 Else ok = False
            Else
                ok = False
            End If
        Case RX_WORDB
            If IsWordAt(pos - 1) <> IsWordAt(pos) Then pc = pc + 1 Else ok = False
        Case RX_NWORDB
            If IsWordAt(pos - 1) = IsWordAt(pos) Then pc = pc + 1 Else ok = False
        Case RX_ANYALL
            If pos < gRxLen Then
                pos = pos + 1: pc = pc + 1
            Else
                ok = False
            End If
        Case RX_CLEAR
            For i = 2 * mA(pc) To 2 * mB(pc) + 1
                If mCaps(i) <> -1 Then
                    Push BT_CAP, i, mCaps(i), 0
                    mCaps(i) = -1
                End If
            Next i
            pc = pc + 1
        Case RX_MARK
            k = mA(pc)
            Push BT_REG, k, mRegs(k), 0
            mRegs(k) = pos
            pc = pc + 1
        Case RX_CHECK
            If pos = mRegs(mA(pc)) Then ok = False Else pc = pc + 1
        Case RX_LOOK
            If Look(pc, pos) Then pc = mC(pc) Else ok = False
        Case RX_BACKREF
            k = mA(pc)
            If mCaps(2 * k) < 0 Or mCaps(2 * k + 1) < 0 Then
                pc = pc + 1
            Else
                p = mCaps(2 * k)
                n = mCaps(2 * k + 1) - p
                If pos + n > gRxLen Then
                    ok = False
                Else
                    For i = 0 To n - 1
                        If mIgnoreCase Then
                            If gRxCanon(gRxCode(p + i)) <> gRxCanon(gRxCode(pos + i)) Then ok = False
                        ElseIf gRxCode(p + i) <> gRxCode(pos + i) Then
                            ok = False
                        End If
                        If Not ok Then Exit For
                    Next i
                    If ok Then pos = pos + n: pc = pc + 1
                End If
            End If
        End Select
        If Not ok Then
            ' Backtrack to the most recent choice point, undoing captures on the way.
            Do
                If mTop <= base Then
                    Run = False
                    Exit Function
                End If
                mTop = mTop - 1
                t = mTop * 4
                Select Case mStk(t)
                Case BT_BRANCH
                    pc = mStk(t + 1)
                    pos = mStk(t + 2)
                    Exit Do
                Case BT_CAP
                    mCaps(mStk(t + 1)) = mStk(t + 2)
                Case BT_GREEDY
                    ' Give back one more character, skipping positions where the
                    ' continuation cannot start.
                    k = mStk(t + 1)
                    mn = mStk(t + 2)
                    pos = mStk(t + 3) - 1
                    f = mFollowIdx(k)
                    If f >= 0 Then
                        Do While pos >= mn
                            c = gRxCode(pos)
                            If c < 128 Then
                                If mFollow(f * 128 + c) Then Exit Do
                            ElseIf mFollowHigh(f) Then
                                Exit Do
                            End If
                            pos = pos - 1
                        Loop
                    End If
                    If pos >= mn Then
                        If pos > mn Then
                            mStk(t + 3) = pos
                            mTop = mTop + 1
                        End If
                        pc = k + 1
                        Exit Do
                    End If
                Case BT_LAZY
                    ' Take one more character (or more, while the continuation cannot start).
                    k = mStk(t + 1)
                    p = mStk(t + 2)
                    cnt = mStk(t + 3)
                    mx = mD(k)
                    found = False
                    f = mFollowIdx(k)
                    n = mA(k)
                    i = mB(k)
                    Do While p < gRxLen
                        If mx >= 0 And cnt >= mx Then Exit Do
                        c = gRxCode(p)
                        Select Case n
                        Case RXA_ALL
                        Case RXA_CLASS
                            If c < 128 Then
                                If Not mAscii(i * 128 + c) Then Exit Do
                            ElseIf Not InClassSlow(i, c) Then
                                Exit Do
                            End If
                        Case RXA_ANY
                            If c = 10 Or c = 13 Or c = 8232 Or c = 8233 Then Exit Do
                        Case Else
                            If mIgnoreCase Then c = gRxCanon(c)
                            If c <> i Then Exit Do
                        End Select
                        p = p + 1
                        cnt = cnt + 1
                        If f < 0 Or p >= gRxLen Then
                            found = True
                            Exit Do
                        End If
                        c = gRxCode(p)
                        If c < 128 Then
                            If mFollow(f * 128 + c) Then
                                found = True
                                Exit Do
                            End If
                        ElseIf mFollowHigh(f) Then
                            found = True
                            Exit Do
                        End If
                    Loop
                    If found Then
                        If mx < 0 Or cnt < mx Then
                            mStk(t + 2) = p
                            mStk(t + 3) = cnt
                            mTop = mTop + 1
                        End If
                        pc = k + 1
                        pos = p
                        Exit Do
                    End If
                Case BT_REG
                    mRegs(mStk(t + 1)) = mStk(t + 2)
                End Select
            Loop
        End If
    Loop
End Function

' Runs a lookahead at pc (RX_LOOK). Lookaheads are atomic: their choice points are
' discarded on success. Captures set inside a positive lookahead stay set (and are
' restored when backtracking past it); a negative lookahead never sets captures.
Private Function Look(ByVal pc As Long, ByVal pos As Long) As Boolean
    Dim snap() As Long, savedTop As Long, hit As Boolean, i As Long, savedEnd As Long
    snap = mCaps
    savedTop = mTop
    savedEnd = mEnd
    hit = Run(mB(pc), pos)
    mEnd = savedEnd
    If hit Then mTop = savedTop
    If mA(pc) = 0 Then
        If hit Then
            For i = 2 To 2 * mGroups + 1
                If mCaps(i) <> snap(i) Then Push BT_CAP, i, snap(i), 0
            Next i
        End If
        Look = hit
    Else
        If hit Then
            mCaps = snap
            Look = False
        Else
            Look = True
        End If
    End If
End Function

' =====================================================================
' Results
' =====================================================================

Private Function NewMatch(ByRef text As String, ByVal owner As clsMatch) As clsMatch
    Dim m As clsMatch
    Set m = New clsMatch
    m.Init text, mCaps, mGroups, owner
    Set NewMatch = m
End Function

Private Sub AddPart(ByRef part As String)
    If mPartCount > UBound(mParts) Then ReDim Preserve mParts(0 To mPartCount * 2 + 1)
    mParts(mPartCount) = part
    mPartCount = mPartCount + 1
End Sub

Private Function GroupText(ByRef text As String, ByVal g As Long) As String
    If mCaps(2 * g) >= 0 And mCaps(2 * g + 1) >= 0 Then
        GroupText = Mid$(text, mCaps(2 * g) + 1, mCaps(2 * g + 1) - mCaps(2 * g))
    End If
End Function

' GetSubstitution for one match (s..e).
Private Function Substitute(ByRef text As String, ByRef template As String, ByVal s As Long, ByVal e As Long) As String
    Dim i As Long, L As Long, ch As String, nx As Long, d1 As Long, d2 As Long, nn As Long, out As String
    Dim litStart As Long
    L = Len(template)
    i = 1
    litStart = 1
    Do While i <= L
        If Mid$(template, i, 1) = "$" And i < L Then
            nx = AscW(Mid$(template, i + 1, 1))
            Select Case nx
            Case 36 ' $$
                out = out & Mid$(template, litStart, i - litStart) & "$"
                i = i + 2: litStart = i
            Case 38 ' $&
                out = out & Mid$(template, litStart, i - litStart) & Mid$(text, s + 1, e - s)
                i = i + 2: litStart = i
            Case 96 ' $`
                out = out & Mid$(template, litStart, i - litStart) & Left$(text, s)
                i = i + 2: litStart = i
            Case 39 ' $'
                out = out & Mid$(template, litStart, i - litStart)
                If e < Len(text) Then out = out & Mid$(text, e + 1)
                i = i + 2: litStart = i
            Case 48 To 57
                d1 = nx - 48
                nn = -1
                If i + 2 <= L Then
                    d2 = AscW(Mid$(template, i + 2, 1))
                    If d2 >= 48 And d2 <= 57 Then
                        nn = d1 * 10 + (d2 - 48)
                        If nn < 1 Or nn > mGroups Then nn = -1
                    End If
                End If
                If nn >= 1 Then
                    out = out & Mid$(template, litStart, i - litStart) & GroupText(text, nn)
                    i = i + 3: litStart = i
                ElseIf d1 >= 1 And d1 <= mGroups Then
                    out = out & Mid$(template, litStart, i - litStart) & GroupText(text, d1)
                    i = i + 2: litStart = i
                Else
                    i = i + 1
                End If
            Case Else
                i = i + 1
            End Select
        Else
            i = i + 1
        End If
    Loop
    If litStart <= L Then out = out & Mid$(template, litStart)
    Substitute = out
End Function
