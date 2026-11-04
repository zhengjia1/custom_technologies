Attribute VB_Name = "modJs"
Option Explicit
Option Compare Binary
Option Private Module

' =====================================================================================
' modJs - JavaScript semantics for the BT1101 Prose Lab Excel edition (pure VBA,
' Windows and Mac). Number formatting/parsing matches V8 exactly; string helpers use
' 0-based UTF-16 indices like JavaScript.
'
' NaN AND INFINITY CONVENTION (measured on Excel 64-bit VBA7, Windows):
'   * Numbers are Doubles. NaN is the IEEE quiet NaN (bits 7FF8000000000000, from JsNaN())
'     and +/-Infinity are the IEEE infinities (JsInfinity(sign)). They are built from bit
'     patterns; VBA itself never produces them (overflow and 0/0 raise run-time errors).
'   * NEVER compare a Double that may be NaN with =, <>, <, >, <=, >=: VBA raises
'     run-time error 6 (Overflow) for any comparison with a NaN Double (and with a NaN in
'     a Variant it silently answers True for "=", which is just as wrong). Test with
'     JsIsNaN(x) first, or use JsEq/JsLt/JsLe/JsGt/JsGe/JsNe, which follow JS (false
'     whenever NaN is involved). Sgn(), CLng(), ^ also raise error 6 on NaN.
'   * Infinity compares correctly and +, -, *, / with one infinite operand work, but
'     VBA raises errors where IEEE/JS give NaN or Infinity: Inf - Inf, Inf / Inf,
'     0 / 0 (error 6), x / 0 (error 11), finite overflow such as 1E308 * 10 (error 6),
'     Sqr(-1) / Log(0) / Log(-1) (error 5), Exp(1000) (error 6), (-8) ^ (1/3) (error 5).
'     NaN arithmetic (+ - * / Int Fix Abs) propagates NaN without an error on Windows,
'     but do not rely on it (Mac VBA may differ).
'   * RULE FOR CALLERS: whenever an operand can be NaN or infinite, or the result can
'     overflow or divide by zero, use JsAdd/JsSub/JsMul/JsDiv/JsMod/JsPow/JsSqrt/JsLog/
'     JsLog10/JsLog2/JsExp/JsFloor/JsCeil/JsTrunc/JsAbs/JsSign/JsMax/JsMin and the JsEq/
'     JsLt/... comparisons. Plain VBA operators are fine for values known to be finite
'     with finite results (the common case after a JsIsFinite check).
'
' FLOATING-POINT PRECISION (measured): VBA evaluates the intermediates of an expression
' in x87 extended precision (64-bit mantissa) and rounds to Double only when a value is
' stored in a variable. So (1 + 2^-53) - 1 written as one expression gives 2^-53 in VBA
' but 0 in JS, and even a single stored operation is rounded twice (64 then 53 bits),
' which differs from JS in the last bit in roughly 1 of 2000 inexact operations. Where
' bit-exact agreement with the website matters (sums that are compared, values that are
' rounded at a .5 boundary), store each step in a Double variable. This module's own
' parsing/formatting is exact regardless (big-integer arithmetic, verified fast paths).
' Math functions: JsLog and JsLog10 are ports of V8's fdlibm code and are bit-identical
' to V8; JsExp/JsSin/JsCos/JsTan/JsAtan and non-special JsPow use VBA's Exp/Sin/Cos/
' Tan/Atn/^ and differ from V8 by one ulp in a few percent of cases. JsPow(10, k) for
' integer k reproduces V8 exactly (including V8's off-by-one-ulp powers of ten).
'   * VBA's Mod rounds its operands to integers: use JsMod for JavaScript's %.
'   * CStr/Format/Val/CDbl on numbers depend on the locale or have quirks: format with
'     JsNumberToString/JsToFixed/... and parse with JsNumber/JsParseFloat/JsParseInt.
'   * -0 is a real negative zero (VBA keeps it: -0 = 0 is True, CStr gives "-0").
'     JsNumberToString(-0) = "0" like JS.
'
' Exact formatting works from the exact binary value of the Double (big-integer decimal
' expansion), so halfway cases such as (1.005).toFixed(2) = "1.00" and (8.345).toFixed(2)
' = "8.35" match V8. Decimal-to-double conversion is correctly rounded (ties to even).
' =====================================================================================

Private Type TDbl
    d As Double
End Type

Private Type TWords
    lo As Long
    hi As Long
End Type

Private Const TWO32 As Double = 4294967296#
Private Const TWO31 As Double = 2147483648#
Private Const TWO52 As Double = 4503599627370496#
Private Const TWO53 As Double = 9007199254740992#
Private Const LIMB As Double = 10000000#
Private Const POW_CACHE As Long = 1100

Private mPow5(0 To POW_CACHE) As Variant
Private mPow2(0 To POW_CACHE) As Variant
Private mP10(0 To 22) As Double
Private mP10Ready As Boolean
Private mRandState As Double
Private mRandSeeded As Boolean
' fdlibm log constants (built from their bit patterns by LogInit)
Private mFdReady As Boolean
Private fdLn2Hi As Double, fdLn2Lo As Double, fdTwo54 As Double, fdThird As Double
Private fdLg1 As Double, fdLg2 As Double, fdLg3 As Double, fdLg4 As Double, fdLg5 As Double, fdLg6 As Double, fdLg7 As Double

' -------------------------------------------------------------------------------------
' Bits, NaN and Infinity
' -------------------------------------------------------------------------------------

Private Sub DblWords(ByVal x As Double, ByRef hi As Long, ByRef lo As Long)
    Dim a As TDbl, b As TWords
    a.d = x
    LSet b = a
    hi = b.hi
    lo = b.lo
End Sub

Private Function WordsDbl(ByVal hi As Long, ByVal lo As Long) As Double
    Dim a As TDbl, b As TWords
    b.hi = hi
    b.lo = lo
    LSet a = b
    WordsDbl = a.d
End Function

Private Function ULong(ByVal v As Long) As Double
    If v < 0 Then ULong = CDbl(v) + TWO32 Else ULong = v
End Function

' Returns 0 for +/-0, 1 for finite non-zero, 2 for +/-Infinity, 3 for NaN.
' For finite x: |x| = m * 2^e with m an integer (< 2^53); neg is the sign bit.
Private Function Decompose(ByVal x As Double, ByRef neg As Boolean, ByRef m As Double, ByRef e As Long) As Long
    Dim hi As Long, lo As Long, ex As Long, f As Double
    DblWords x, hi, lo
    neg = (hi < 0)
    ex = (hi And &H7FF00000) \ &H100000
    f = CDbl(hi And &HFFFFF) * TWO32 + ULong(lo)
    If ex = &H7FF Then
        If f = 0 Then Decompose = 2 Else Decompose = 3
    ElseIf ex = 0 Then
        m = f
        e = -1074
        If f = 0 Then Decompose = 0 Else Decompose = 1
    Else
        m = f + TWO52
        e = ex - 1075
        Decompose = 1
    End If
End Function

' Bit pattern helpers (also used by the tests): 16 upper-case hex digits.
Public Function JsDoubleToHex(ByVal x As Double) As String
    Dim hi As Long, lo As Long
    DblWords x, hi, lo
    JsDoubleToHex = Right$("0000000" & Hex$(hi), 8) & Right$("0000000" & Hex$(lo), 8)
End Function

Public Function JsHexToDouble(ByVal hex16 As String) As Double
    JsHexToDouble = WordsDbl(HexWord(Left$(hex16, 8)), HexWord(Mid$(hex16, 9, 8)))
End Function

Private Function HexWord(ByVal h As String) As Long
    Dim i As Long, v As Double, c As Long
    For i = 1 To Len(h)
        c = AscW(Mid$(h, i, 1))
        Select Case c
        Case 48 To 57: c = c - 48
        Case 65 To 70: c = c - 55
        Case 97 To 102: c = c - 87
        Case Else: c = 0
        End Select
        v = v * 16 + c
    Next
    If v >= TWO31 Then v = v - TWO32
    HexWord = CLng(v)
End Function

Public Function JsNaN() As Double
    JsNaN = WordsDbl(&H7FF80000, 0)
End Function

Public Function JsInfinity(Optional ByVal sign As Long = 1) As Double
    If sign < 0 Then JsInfinity = WordsDbl(&HFFF00000, 0) Else JsInfinity = WordsDbl(&H7FF00000, 0)
End Function

Public Function JsIsNaN(ByVal x As Double) As Boolean
    Dim hi As Long, lo As Long
    DblWords x, hi, lo
    If (hi And &H7FF00000) = &H7FF00000 Then JsIsNaN = ((hi And &HFFFFF) <> 0) Or (lo <> 0)
End Function

' Number.isFinite: False for NaN and +/-Infinity.
Public Function JsIsFinite(ByVal x As Double) As Boolean
    Dim hi As Long, lo As Long
    DblWords x, hi, lo
    JsIsFinite = ((hi And &H7FF00000) <> &H7FF00000)
End Function

' Number.isInteger
Public Function JsIsInteger(ByVal x As Double) As Boolean
    If JsIsFinite(x) Then JsIsInteger = (Int(x) = x)
End Function

' True for -0 (and other negative numbers): the IEEE sign bit.
Public Function JsSignBit(ByVal x As Double) As Boolean
    Dim hi As Long, lo As Long
    DblWords x, hi, lo
    JsSignBit = (hi < 0)
End Function

Public Function JsIsNegativeZero(ByVal x As Double) As Boolean
    Dim hi As Long, lo As Long
    DblWords x, hi, lo
    JsIsNegativeZero = (hi = &H80000000 And lo = 0)
End Function

Private Function SignOf(ByVal x As Double) As Long
    If JsSignBit(x) Then SignOf = -1 Else SignOf = 1
End Function

Private Function NegZero() As Double
    NegZero = WordsDbl(&H80000000, 0)
End Function

Private Function MaxDouble() As Double
    MaxDouble = WordsDbl(&H7FEFFFFF, -1)
End Function

' 2^k exactly for -1022 <= k <= 1023.
Private Function Pow2Dbl(ByVal k As Long) As Double
    Pow2Dbl = WordsDbl((k + 1023) * &H100000, 0)
End Function

Private Function NextUpPos(ByVal c As Double) As Double
    Dim hi As Long, lo As Long
    DblWords c, hi, lo
    If lo = -1 Then
        lo = 0
        hi = hi + 1
    ElseIf lo = &H7FFFFFFF Then
        lo = &H80000000
    Else
        lo = lo + 1
    End If
    NextUpPos = WordsDbl(hi, lo)
End Function

Private Function NextDownPos(ByVal c As Double) As Double
    Dim hi As Long, lo As Long
    DblWords c, hi, lo
    If lo = 0 Then
        lo = -1
        hi = hi - 1
    ElseIf lo = &H80000000 Then
        lo = &H7FFFFFFF
    Else
        lo = lo - 1
    End If
    NextDownPos = WordsDbl(hi, lo)
End Function

' -------------------------------------------------------------------------------------
' Safe arithmetic (JS results instead of VBA run-time errors)
' -------------------------------------------------------------------------------------

Public Function JsAdd(ByVal a As Double, ByVal b As Double) As Double
    If JsIsNaN(a) Or JsIsNaN(b) Then JsAdd = JsNaN(): Exit Function
    On Error GoTo Special
    JsAdd = a + b
    Exit Function
Special:
    If JsIsFinite(a) And JsIsFinite(b) Then JsAdd = JsInfinity(SignOf(a)) Else JsAdd = JsNaN()
End Function

Public Function JsSub(ByVal a As Double, ByVal b As Double) As Double
    If JsIsNaN(a) Or JsIsNaN(b) Then JsSub = JsNaN(): Exit Function
    JsSub = JsAdd(a, -b)
End Function

Public Function JsMul(ByVal a As Double, ByVal b As Double) As Double
    If JsIsNaN(a) Or JsIsNaN(b) Then JsMul = JsNaN(): Exit Function
    If (a = 0 And Not JsIsFinite(b)) Or (b = 0 And Not JsIsFinite(a)) Then JsMul = JsNaN(): Exit Function
    On Error GoTo Special
    JsMul = a * b
    Exit Function
Special:
    JsMul = JsInfinity(SignOf(a) * SignOf(b))
End Function

Public Function JsDiv(ByVal a As Double, ByVal b As Double) As Double
    If JsIsNaN(a) Or JsIsNaN(b) Then JsDiv = JsNaN(): Exit Function
    If b = 0 Then
        If a = 0 Then JsDiv = JsNaN() Else JsDiv = JsInfinity(SignOf(a) * SignOf(b))
        Exit Function
    End If
    If Not JsIsFinite(a) And Not JsIsFinite(b) Then JsDiv = JsNaN(): Exit Function
    On Error GoTo Special
    JsDiv = a / b
    Exit Function
Special:
    JsDiv = JsInfinity(SignOf(a) * SignOf(b))
End Function

' JavaScript a % b (exact fmod; the result has the sign of a).
Public Function JsMod(ByVal a As Double, ByVal b As Double) As Double
    Dim r As Double, d As Double, t As Double
    If JsIsNaN(a) Or JsIsNaN(b) Then JsMod = JsNaN(): Exit Function
    If Not JsIsFinite(a) Or b = 0 Then JsMod = JsNaN(): Exit Function
    If Not JsIsFinite(b) Or a = 0 Then JsMod = a: Exit Function
    r = Abs(a)
    d = Abs(b)
    If r < d Then JsMod = a: Exit Function
    t = d
    Do While t <= r / 2
        t = t * 2
    Loop
    Do While t >= d
        If r >= t Then r = r - t
        t = t / 2
    Loop
    If a < 0 Then
        If r = 0 Then JsMod = NegZero() Else JsMod = -r
    Else
        JsMod = r
    End If
End Function

Private Function IsOddInteger(ByVal x As Double) As Boolean
    If JsIsFinite(x) Then
        If Int(x) = x And Abs(x) < TWO53 Then IsOddInteger = (x - 2 * Int(x / 2) <> 0)
    End If
End Function

' Math.pow / **. Special cases follow ECMAScript; finite results come from VBA's ^
' (may differ from V8 in the last bit for non-trivial operands).
Public Function JsPow(ByVal a As Double, ByVal b As Double) As Double
    If JsIsNaN(b) Then JsPow = JsNaN(): Exit Function
    If b = 0 Then JsPow = 1: Exit Function
    If JsIsNaN(a) Then JsPow = JsNaN(): Exit Function
    If Not JsIsFinite(b) Then
        If Abs(a) = 1 Then
            JsPow = JsNaN()
        ElseIf (Abs(a) > 1) = (b > 0) Then
            JsPow = JsInfinity(1)
        Else
            JsPow = 0
        End If
        Exit Function
    End If
    If Not JsIsFinite(a) Or a = 0 Then
        Dim big As Boolean, negRes As Boolean
        big = (Not JsIsFinite(a)) = (b > 0)
        negRes = JsSignBit(a) And IsOddInteger(b)
        If big Then
            If negRes Then JsPow = JsInfinity(-1) Else JsPow = JsInfinity(1)
        Else
            If negRes Then JsPow = NegZero() Else JsPow = 0
        End If
        Exit Function
    End If
    If a < 0 And Int(b) <> b Then JsPow = JsNaN(): Exit Function
    If Int(b) = b And Abs(b) <= 1074 Then
        If a = 10 And Abs(b) <= 400 Then JsPow = Pow10V8(CLng(b)): Exit Function
        If a = 2 Then
            If b >= -1022 And b <= 1023 Then
                JsPow = Pow2Dbl(CLng(b))
            ElseIf b > 1023 Then
                JsPow = JsInfinity(1)
            Else
                JsPow = Pow2Dbl(CLng(b) + 100) / Pow2Dbl(100)
            End If
            Exit Function
        End If
    End If
    ' fdlibm's special cases (exact like V8)
    If b = 1 Then JsPow = a: Exit Function
    If b = -1 Then JsPow = JsDiv(1, a): Exit Function
    If b = 2 Then JsPow = JsMul(a, a): Exit Function
    If b = 0.5 Then JsPow = Sqr(a): Exit Function
    On Error GoTo Special
    JsPow = a ^ b
    Exit Function
Special:
    If a < 0 And IsOddInteger(b) Then JsPow = JsInfinity(-1) Else JsPow = JsInfinity(1)
End Function

' 10 ** k exactly as V8 computes it: the correctly rounded 1e<k>, except for the listed
' exponents where V8's pow (fdlibm) is one ulp above (+) or below (-).
Private Function Pow10V8(ByVal k As Long) As Double
    Const V8_POW10_ULPS As String = ",-307+,-306-,-303+,-278+,-268+,-246+,-222-,-207+,-166-,-163+,-148+," & _
        "-145+,-138-,-128-,-122-,-111-,-106+,-104+,-101-,-88+,-76+,-63-,-57+,-39+,-32-,-29+,-24+,-23+," & _
        "-21+,-20+,-17-,-5-,-4-,26-,29+,34+,36-,39+,42-,49+,58+,61+,70-,95-,96-,111+,126+,136-,142-,165+," & _
        "166+,169+,187+,191-,210+,215+,218-,225+,232-,242-,244-,260-,273+,274+,275+,279-,284-,290-,294-,"
    Dim r As Double
    r = DecToDouble("1", k + 1)
    If InStr(1, V8_POW10_ULPS, "," & CStr(k) & "+,", vbBinaryCompare) > 0 Then r = NextUpPos(r)
    If InStr(1, V8_POW10_ULPS, "," & CStr(k) & "-,", vbBinaryCompare) > 0 Then r = NextDownPos(r)
    Pow10V8 = r
End Function

Public Function JsSqrt(ByVal x As Double) As Double
    If JsIsNaN(x) Then JsSqrt = x: Exit Function
    If x = 0 Then JsSqrt = x: Exit Function
    If x < 0 Then JsSqrt = JsNaN(): Exit Function
    If Not JsIsFinite(x) Then JsSqrt = x: Exit Function
    JsSqrt = Sqr(x)
End Function

' Math.log: port of fdlibm / FreeBSD e_log.c, the algorithm V8 uses. VBA evaluates an
' expression's intermediates in x87 extended precision, so every operation is stored
' into a Double variable to get V8's double-precision results.
Public Function JsLog(ByVal x As Double) As Double
    Dim hx As Long, lx As Long, k As Long, i As Long, j As Long
    Dim f As Double, s As Double, z As Double, w As Double, R As Double, a As Double, b As Double
    Dim t1 As Double, t2 As Double, dk As Double, hfsq As Double
    If JsIsNaN(x) Then JsLog = x: Exit Function
    LogInit
    DblWords x, hx, lx
    If hx < &H100000 Then
        If ((hx And &H7FFFFFFF) Or lx) = 0 Then JsLog = JsInfinity(-1): Exit Function
        If hx < 0 Then JsLog = JsNaN(): Exit Function
        k = -54
        x = x * fdTwo54
        DblWords x, hx, lx
    End If
    If hx >= &H7FF00000 Then JsLog = x: Exit Function
    If hx = &H3FF00000 And lx = 0 Then JsLog = 0: Exit Function
    k = k + (hx \ &H100000) - 1023
    hx = hx And &HFFFFF
    i = (hx + &H95F64) And &H100000
    x = WordsDbl(hx Or (i Xor &H3FF00000), lx)
    k = k + i \ &H100000
    f = x - 1#
    dk = k
    If (&HFFFFF And (2 + hx)) < 3 Then
        If f = 0 Then
            If k <> 0 Then
                a = dk * fdLn2Hi
                b = dk * fdLn2Lo
                JsLog = a + b
            End If
            Exit Function
        End If
        a = fdThird * f
        a = 0.5 - a
        b = f * f
        R = b * a
        If k = 0 Then
            JsLog = f - R
        Else
            a = dk * fdLn2Lo
            a = R - a
            a = a - f
            b = dk * fdLn2Hi
            JsLog = b - a
        End If
        Exit Function
    End If
    a = 2# + f
    s = f / a
    z = s * s
    i = hx - &H6147A
    w = z * z
    j = &H6B851 - hx
    a = w * fdLg6
    a = fdLg4 + a
    a = w * a
    a = fdLg2 + a
    t1 = w * a
    a = w * fdLg7
    a = fdLg5 + a
    a = w * a
    a = fdLg3 + a
    a = w * a
    a = fdLg1 + a
    t2 = z * a
    i = i Or j
    R = t2 + t1
    If i > 0 Then
        a = 0.5 * f
        hfsq = a * f
        a = hfsq + R
        a = s * a
        If k = 0 Then
            a = hfsq - a
            JsLog = f - a
        Else
            b = dk * fdLn2Lo
            a = a + b
            a = hfsq - a
            a = a - f
            b = dk * fdLn2Hi
            JsLog = b - a
        End If
    Else
        a = f - R
        a = s * a
        If k = 0 Then
            JsLog = f - a
        Else
            b = dk * fdLn2Lo
            a = a - b
            a = a - f
            b = dk * fdLn2Hi
            JsLog = b - a
        End If
    End If
End Function

Private Sub LogInit()
    If mFdReady Then Exit Sub
    fdLn2Hi = WordsDbl(&H3FE62E42, &HFEE00000)
    fdLn2Lo = WordsDbl(&H3DEA39EF, &H35793C76)
    fdTwo54 = WordsDbl(&H43500000, 0)
    fdThird = WordsDbl(&H3FD55555, &H55555555)
    fdLg1 = WordsDbl(&H3FE55555, &H55555593)
    fdLg2 = WordsDbl(&H3FD99999, &H9997FA04)
    fdLg3 = WordsDbl(&H3FD24924, &H94229359)
    fdLg4 = WordsDbl(&H3FCC71C5, &H1D8E78AF)
    fdLg5 = WordsDbl(&H3FC74664, &H96CB03DE)
    fdLg6 = WordsDbl(&H3FC39A09, &HD078C69F)
    fdLg7 = WordsDbl(&H3FC2F112, &HDF3E5244)
    mFdReady = True
End Sub

' Math.log10: port of fdlibm e_log10.c as used by V8 (each operation stored, see JsLog).
Public Function JsLog10(ByVal x As Double) As Double
    Dim hx As Long, lx As Long, k As Long, i As Long, y As Double, z As Double, a As Double
    If JsIsNaN(x) Then JsLog10 = x: Exit Function
    LogInit
    DblWords x, hx, lx
    If hx < &H100000 Then
        If ((hx And &H7FFFFFFF) Or lx) = 0 Then JsLog10 = JsInfinity(-1): Exit Function
        If hx < 0 Then JsLog10 = JsNaN(): Exit Function
        k = -54
        x = x * fdTwo54
        DblWords x, hx, lx
    End If
    If hx >= &H7FF00000 Then JsLog10 = x: Exit Function
    k = k + (hx \ &H100000) - 1023
    If k < 0 Then i = 1
    hx = (hx And &HFFFFF) Or ((&H3FF - i) * &H100000)
    y = k + i
    x = WordsDbl(hx, lx)
    a = JsLog(x)
    a = WordsDbl(&H3FDBCB7B, &H1526E50E) * a
    z = y * WordsDbl(&H3D59FEF3, &H11F12B36)
    z = z + a
    a = y * WordsDbl(&H3FD34413, &H509F6000)
    JsLog10 = z + a
End Function

' Math.log2; exact for powers of two.
Public Function JsLog2(ByVal x As Double) As Double
    Dim r As Double, k As Long, m As Double, e As Long, neg As Boolean
    r = JsLog(x)
    If Not JsIsFinite(r) Then JsLog2 = r: Exit Function
    Decompose x, neg, m, e
    k = 0
    Do While m - 2 * Int(m / 2) = 0
        m = m / 2
        k = k + 1
    Loop
    If m = 1 Then JsLog2 = k + e Else JsLog2 = r / Log(2#)
End Function

' Math.sin / Math.cos / Math.tan / Math.atan / Math.exp use VBA's functions: results can
' differ from V8 in the last bit (V8 uses its own libm ports).
Public Function JsSin(ByVal x As Double) As Double
    If Not JsIsFinite(x) Then JsSin = JsNaN() Else JsSin = Sin(x)
End Function

Public Function JsCos(ByVal x As Double) As Double
    If Not JsIsFinite(x) Then JsCos = JsNaN() Else JsCos = Cos(x)
End Function

Public Function JsTan(ByVal x As Double) As Double
    If Not JsIsFinite(x) Then JsTan = JsNaN() Else JsTan = Tan(x)
End Function

Public Function JsAtan(ByVal x As Double) As Double
    If JsIsNaN(x) Then JsAtan = x Else JsAtan = Atn(x)
End Function

' Math.PI and Math.E (full precision; VBA literals carry only 15 digits).
Public Function JsPI() As Double
    JsPI = WordsDbl(&H400921FB, &H54442D18)
End Function

Public Function JsE() As Double
    JsE = WordsDbl(&H4005BF0A, &H8B145769)
End Function

' Number.EPSILON (2^-52)
Public Function JsEpsilon() As Double
    JsEpsilon = Pow2Dbl(-52)
End Function

Public Function JsExp(ByVal x As Double) As Double
    If JsIsNaN(x) Then JsExp = x: Exit Function
    If Not JsIsFinite(x) Then
        If x > 0 Then JsExp = x Else JsExp = 0
        Exit Function
    End If
    If x > 710 Then JsExp = JsInfinity(1): Exit Function
    If x < -746 Then JsExp = 0: Exit Function
    On Error GoTo Special
    JsExp = Exp(x)
    Exit Function
Special:
    JsExp = JsInfinity(1)
End Function

Public Function JsFloor(ByVal x As Double) As Double
    If Not JsIsFinite(x) Then JsFloor = x Else JsFloor = Int(x)
End Function

Public Function JsCeil(ByVal x As Double) As Double
    If Not JsIsFinite(x) Then JsCeil = x Else JsCeil = -Int(-x)
End Function

Public Function JsTrunc(ByVal x As Double) As Double
    If Not JsIsFinite(x) Then JsTrunc = x Else JsTrunc = Fix(x)
End Function

Public Function JsAbs(ByVal x As Double) As Double
    Dim hi As Long, lo As Long
    DblWords x, hi, lo
    JsAbs = WordsDbl(hi And &H7FFFFFFF, lo)
End Function

Public Function JsSign(ByVal x As Double) As Double
    If JsIsNaN(x) Then JsSign = x: Exit Function
    If x = 0 Then JsSign = x Else JsSign = Sgn(x)
End Function

' Math.max(...values): NaN if any value is NaN; -Infinity for no values.
Public Function JsMax(ParamArray values() As Variant) As Double
    Dim i As Long, r As Double, v As Double
    r = JsInfinity(-1)
    For i = LBound(values) To UBound(values)
        v = CDbl(values(i))
        If JsIsNaN(v) Then JsMax = v: Exit Function
        If v > r Or (v = 0 And r = 0 And Not JsSignBit(v)) Then r = v
    Next
    JsMax = r
End Function

' Math.min(...values): NaN if any value is NaN; +Infinity for no values.
Public Function JsMin(ParamArray values() As Variant) As Double
    Dim i As Long, r As Double, v As Double
    r = JsInfinity(1)
    For i = LBound(values) To UBound(values)
        v = CDbl(values(i))
        If JsIsNaN(v) Then JsMin = v: Exit Function
        If v < r Or (v = 0 And r = 0 And JsSignBit(v)) Then r = v
    Next
    JsMin = r
End Function

' Comparisons with JS semantics (False whenever NaN is involved; -0 equals 0).
Public Function JsEq(ByVal a As Double, ByVal b As Double) As Boolean
    If JsIsNaN(a) Or JsIsNaN(b) Then Exit Function
    JsEq = (a = b)
End Function

Public Function JsNe(ByVal a As Double, ByVal b As Double) As Boolean
    JsNe = Not JsEq(a, b)
End Function

Public Function JsLt(ByVal a As Double, ByVal b As Double) As Boolean
    If JsIsNaN(a) Or JsIsNaN(b) Then Exit Function
    JsLt = (a < b)
End Function

Public Function JsLe(ByVal a As Double, ByVal b As Double) As Boolean
    If JsIsNaN(a) Or JsIsNaN(b) Then Exit Function
    JsLe = (a <= b)
End Function

Public Function JsGt(ByVal a As Double, ByVal b As Double) As Boolean
    If JsIsNaN(a) Or JsIsNaN(b) Then Exit Function
    JsGt = (a > b)
End Function

Public Function JsGe(ByVal a As Double, ByVal b As Double) As Boolean
    If JsIsNaN(a) Or JsIsNaN(b) Then Exit Function
    JsGe = (a >= b)
End Function

' ToInt32 / ToUint32 (for porting bitwise operators).
Public Function JsToUint32(ByVal x As Double) As Double
    If Not JsIsFinite(x) Then Exit Function
    x = Fix(x)
    x = x - Int(x / TWO32) * TWO32
    If x = 0 Then x = 0
    JsToUint32 = x
End Function

Public Function JsToInt32(ByVal x As Double) As Long
    Dim u As Double
    u = JsToUint32(x)
    If u >= TWO31 Then u = u - TWO32
    JsToInt32 = CLng(u)
End Function

' -------------------------------------------------------------------------------------
' Big integers: little-endian Double limbs in base 10^7
' -------------------------------------------------------------------------------------

Private Sub BigMulSmall(a() As Double, ByRef n As Long, ByVal f As Double)
    ' f <= 2^21 keeps every intermediate below 2^53
    Dim i As Long, t As Double, c As Double
    For i = 0 To n - 1
        t = a(i) * f + c
        c = Int(t / LIMB)
        a(i) = t - c * LIMB
    Next
    Do While c > 0
        If n > UBound(a) Then ReDim Preserve a(0 To 2 * n + 7)
        t = c
        c = Int(t / LIMB)
        a(n) = t - c * LIMB
        n = n + 1
    Loop
End Sub

' 5^k or 2^k as limbs (cached).
Private Function PowLimbs(ByVal base As Long, ByVal k As Long, ByRef n As Long) As Double()
    Dim a() As Double, stepF As Double, stepK As Long, r As Long, f As Double, i As Long, got As Boolean
    If k <= POW_CACHE Then
        If base = 5 Then
            If Not IsEmpty(mPow5(k)) Then a = mPow5(k): got = True
        Else
            If Not IsEmpty(mPow2(k)) Then a = mPow2(k): got = True
        End If
        If got Then
            n = UBound(a) + 1
            PowLimbs = a
            Exit Function
        End If
    End If
    ReDim a(0 To 7)
    a(0) = 1
    n = 1
    If base = 5 Then
        stepF = 1953125#
        stepK = 9
    Else
        stepF = 1048576#
        stepK = 20
    End If
    r = k
    Do While r >= stepK
        BigMulSmall a, n, stepF
        r = r - stepK
    Loop
    If r > 0 Then
        f = 1
        For i = 1 To r
            f = f * base
        Next
        BigMulSmall a, n, f
    End If
    ReDim Preserve a(0 To n - 1)
    If k <= POW_CACHE Then
        If base = 5 Then mPow5(k) = a Else mPow2(k) = a
    End If
    PowLimbs = a
End Function

Private Function LimbsToString(r() As Double, ByVal n As Long) As String
    Dim s As String, head As String, part As String, i As Long, pos As Long
    Do While n > 1
        If r(n - 1) <> 0 Then Exit Do
        n = n - 1
    Loop
    head = CStr(CLng(r(n - 1)))
    s = head & String$((n - 1) * 7, "0")
    pos = Len(head) + 1
    For i = n - 2 To 0 Step -1
        If r(i) <> 0 Then
            part = CStr(CLng(r(i)))
            Mid$(s, pos + 7 - Len(part), Len(part)) = part
        End If
        pos = pos + 7
    Next
    LimbsToString = s
End Function

' Exact decimal digits of (m * mulA + addB) * 2^e as an integer string S, where the
' value is S / 10^k (k = 0 when e >= 0, else k = -e). m < 2^53, 1 <= mulA <= 4.
Private Function BigExact(ByVal m As Double, ByVal mulA As Long, ByVal addB As Long, ByVal e As Long, ByRef k As Long) As String
    Dim nv(0 To 3) As Double, i As Long, j As Long, q As Double, t As Double, c As Double
    Dim nn As Long, p() As Double, np As Long, r() As Double
    t = m
    For i = 0 To 2
        q = Int(t / LIMB)
        nv(i) = t - q * LIMB
        If nv(i) < 0 Then
            q = q - 1
            nv(i) = nv(i) + LIMB
        ElseIf nv(i) >= LIMB Then
            q = q + 1
            nv(i) = nv(i) - LIMB
        End If
        t = q
    Next
    c = addB
    For i = 0 To 3
        t = nv(i) * mulA + c
        c = Int(t / LIMB)
        nv(i) = t - c * LIMB
    Next
    nn = 4
    Do While nn > 1
        If nv(nn - 1) <> 0 Then Exit Do
        nn = nn - 1
    Loop
    If e >= 0 Then
        p = PowLimbs(2, e, np)
        k = 0
    Else
        p = PowLimbs(5, -e, np)
        k = -e
    End If
    ReDim r(0 To np + nn)
    For j = 0 To nn - 1
        If nv(j) <> 0 Then
            t = nv(j)
            For i = 0 To np - 1
                r(i + j) = r(i + j) + t * p(i)
            Next
        End If
    Next
    c = 0
    For i = 0 To np + nn
        t = r(i) + c
        c = Int(t / LIMB)
        r(i) = t - c * LIMB
    Next
    BigExact = LimbsToString(r, np + nn + 1)
End Function

Private Function StripTrailingZeros(ByVal s As String) As String
    Dim n As Long
    n = Len(s)
    Do While n > 0
        If Mid$(s, n, 1) <> "0" Then Exit Do
        n = n - 1
    Loop
    StripTrailingZeros = Left$(s, n)
End Function

' Adds one to a decimal digit string (may grow by one digit).
Private Function IncDigits(ByVal s As String) As String
    Dim i As Long, c As String
    For i = Len(s) To 1 Step -1
        c = Mid$(s, i, 1)
        If c = "9" Then
            Mid$(s, i, 1) = "0"
        Else
            Mid$(s, i, 1) = Chr$(Asc(c) + 1)
            IncDigits = s
            Exit Function
        End If
    Next
    IncDigits = "1" & s
End Function

' Exact |x| = 0.D * 10^P for finite non-zero x = m * 2^e (D without trailing zeros).
Private Sub ExactDecimal(ByVal m As Double, ByVal e As Long, ByRef D As String, ByRef P As Long)
    Dim k As Long, s As String
    Do While e < 0
        If m - 2 * Int(m / 2) <> 0 Then Exit Do
        m = m / 2
        e = e + 1
    Loop
    s = BigExact(m, 1, 0, e, k)
    P = Len(s) - k
    D = StripTrailingZeros(s)
End Sub

' Rounds 0.D * 10^P (half up on the given digits) to `keep` leading digit positions and
' returns those digits as an integer string ("" when the result is zero). The result
' may have keep + 1 digits after a carry.
Private Function RoundDigits(ByVal D As String, ByVal keep As Long) As String
    Dim r As String
    If keep < 0 Then Exit Function
    If keep = 0 Then
        If Left$(D, 1) >= "5" Then RoundDigits = "1"
        Exit Function
    End If
    If keep >= Len(D) Then
        RoundDigits = D & String$(keep - Len(D), "0")
        Exit Function
    End If
    r = Left$(D, keep)
    If Mid$(D, keep + 1, 1) >= "5" Then r = IncDigits(r)
    RoundDigits = r
End Function

' Decimal digits of an integer 0 <= v < 2^53 without locale effects.
Private Function IntToStr(ByVal v As Double) As String
    Dim h As Double, l As Double
    If v < 2147483647# Then IntToStr = CStr(CLng(v)): Exit Function
    h = Int(v / 1000000000#)
    l = v - h * 1000000000#
    If l < 0 Then
        h = h - 1
        l = l + 1000000000#
    ElseIf l >= 1000000000# Then
        h = h + 1
        l = l - 1000000000#
    End If
    IntToStr = CStr(CLng(h)) & Right$("00000000" & CStr(CLng(l)), 9)
End Function

' Shortest round-trip digits (V8 / ECMAScript Number::toString): |x| = 0.D * 10^n,
' D as short as possible, closest to x, ties to even digit.
Private Sub ShortestDigits(ByVal m As Double, ByVal e As Long, ByRef D As String, ByRef n As Long)
    Dim k As Long, X As String, L As String, H As String, W As Long, lowAdd As Long
    Dim inclusive As Boolean, lead As Long, p As Long, cut As Long, tail As String
    Dim F As String, C As String, head As String, okF As Boolean, okC As Boolean
    Dim cand As String, cmp As Long, i As Long
    inclusive = (m - 2 * Int(m / 2) = 0)
    If m = TWO52 And e > -1074 Then lowAdd = -1 Else lowAdd = -2
    X = BigExact(m, 4, 0, e - 2, k)
    L = BigExact(m, 4, lowAdd, e - 2, k)
    H = BigExact(m, 4, 2, e - 2, k)
    W = Len(H)
    If Len(X) < W Then X = String$(W - Len(X), "0") & X
    If Len(L) < W Then L = String$(W - Len(L), "0") & L
    lead = 1
    Do While Mid$(X, lead, 1) = "0"
        lead = lead + 1
    Loop
    cand = X
    For p = 1 To W - lead + 1
        cut = lead + p - 1
        If cut >= W Then Exit For
        tail = Mid$(X, cut + 1)
        If tail = String$(W - cut, "0") Then Exit For
        head = Left$(X, cut)
        F = head & String$(W - cut, "0")
        cmp = StrComp(F, L, vbBinaryCompare)
        okF = (cmp > 0) Or (cmp = 0 And inclusive)
        C = IncDigits(head)
        If Len(C) > cut Then
            okC = False
        Else
            C = C & String$(W - cut, "0")
            cmp = StrComp(C, H, vbBinaryCompare)
            okC = (cmp < 0) Or (cmp = 0 And inclusive)
        End If
        If okF Or okC Then
            If okF And okC Then
                cmp = StrComp(tail, "5" & String$(Len(tail) - 1, "0"), vbBinaryCompare)
                If cmp = 0 Then
                    If (Asc(Right$(head, 1)) - 48) Mod 2 = 0 Then cmp = -1 Else cmp = 1
                End If
                If cmp < 0 Then cand = F Else cand = C
            ElseIf okF Then
                cand = F
            Else
                cand = C
            End If
            Exit For
        End If
    Next
    i = 1
    Do While Mid$(cand, i, 1) = "0"
        i = i + 1
    Loop
    cand = Mid$(cand, i)
    n = Len(cand) - k
    D = StripTrailingZeros(cand)
End Sub

' Formats shortest digits per ECMAScript Number::toString (no sign).
Private Function FormatShortest(ByVal D As String, ByVal n As Long) As String
    Dim k As Long, ex As Long
    k = Len(D)
    If k <= n And n <= 21 Then
        FormatShortest = D & String$(n - k, "0")
    ElseIf 0 < n And n <= 21 Then
        FormatShortest = Left$(D, n) & "." & Mid$(D, n + 1)
    ElseIf -6 < n And n <= 0 Then
        FormatShortest = "0." & String$(-n, "0") & D
    Else
        ex = n - 1
        If k = 1 Then FormatShortest = D Else FormatShortest = Left$(D, 1) & "." & Mid$(D, 2)
        FormatShortest = FormatShortest & "e" & ExpSign(ex)
    End If
End Function

Private Function ExpSign(ByVal ex As Long) As String
    If ex < 0 Then ExpSign = "-" & CStr(-ex) Else ExpSign = "+" & CStr(ex)
End Function

' Shortest digits of a finite non-zero |x| (fast path for safe integers).
Private Sub ShortestOf(ByVal x As Double, ByVal m As Double, ByVal e As Long, ByRef D As String, ByRef n As Long)
    Dim ax As Double, s As String
    ax = Abs(x)
    If ax < TWO53 Then
        If Int(ax) = ax Then
            s = IntToStr(ax)
            n = Len(s)
            D = StripTrailingZeros(s)
            Exit Sub
        End If
    End If
    ShortestDigits m, e, D, n
End Sub

' -------------------------------------------------------------------------------------
' Number -> String
' -------------------------------------------------------------------------------------

' String(x) / x.toString()
Public Function JsNumberToString(ByVal x As Double) As String
    Dim cls As Long, neg As Boolean, m As Double, e As Long, D As String, n As Long
    cls = Decompose(x, neg, m, e)
    Select Case cls
    Case 3
        JsNumberToString = "NaN"
    Case 0
        JsNumberToString = "0"
    Case 2
        If neg Then JsNumberToString = "-Infinity" Else JsNumberToString = "Infinity"
    Case Else
        ShortestOf x, m, e, D, n
        If neg Then
            JsNumberToString = "-" & FormatShortest(D, n)
        Else
            JsNumberToString = FormatShortest(D, n)
        End If
    End Select
End Function

' x.toFixed(digits); digits 0..100 (else run-time error 5, like RangeError).
Public Function JsToFixed(ByVal x As Double, Optional ByVal digits As Long = 0) As String
    Dim cls As Long, neg As Boolean, m As Double, e As Long, D As String, P As Long, R As String
    cls = Decompose(x, neg, m, e)
    If digits < 0 Or digits > 100 Then Err.Raise 5, "JsToFixed", "toFixed() digits argument must be between 0 and 100"
    If cls = 3 Then JsToFixed = "NaN": Exit Function
    If cls = 2 Then JsToFixed = JsNumberToString(x): Exit Function
    If Abs(x) >= 1E+21 Then JsToFixed = JsNumberToString(x): Exit Function
    If cls = 0 Then
        R = "0"
        neg = False
    Else
        ExactDecimal m, e, D, P
        R = RoundDigits(D, P + digits)
        If R = "" Then R = "0"
    End If
    If digits > 0 Then
        If Len(R) <= digits Then R = String$(digits + 1 - Len(R), "0") & R
        R = Left$(R, Len(R) - digits) & "." & Right$(R, digits)
    End If
    If neg Then JsToFixed = "-" & R Else JsToFixed = R
End Function

' x.toPrecision(precision); omitted precision -> String(x). precision 1..100.
Public Function JsToPrecision(ByVal x As Double, Optional ByVal precision As Variant) As String
    Dim cls As Long, neg As Boolean, m As Double, e As Long, D As String, P As Long, R As String
    Dim pr As Long, ex As Long, s As String
    cls = Decompose(x, neg, m, e)
    If IsMissing(precision) Then JsToPrecision = JsNumberToString(x): Exit Function
    If cls >= 2 Then JsToPrecision = JsNumberToString(x): Exit Function
    pr = CLng(precision)
    If pr < 1 Or pr > 100 Then Err.Raise 5, "JsToPrecision", "toPrecision() argument must be between 1 and 100"
    If cls = 0 Then
        R = String$(pr, "0")
        ex = 0
        neg = False
    Else
        ExactDecimal m, e, D, P
        R = RoundDigits(D, pr)
        ex = P - 1
        If Len(R) > pr Then
            R = Left$(R, pr)
            ex = ex + 1
        End If
    End If
    If ex < -6 Or ex >= pr Then
        s = Left$(R, 1)
        If pr > 1 Then s = s & "." & Mid$(R, 2)
        s = s & "e" & ExpSign(ex)
    ElseIf ex = pr - 1 Then
        s = R
    ElseIf ex >= 0 Then
        s = Left$(R, ex + 1) & "." & Mid$(R, ex + 2)
    Else
        s = "0." & String$(-(ex + 1), "0") & R
    End If
    If neg Then JsToPrecision = "-" & s Else JsToPrecision = s
End Function

' x.toExponential(fractionDigits); omitted -> as many digits as needed (shortest).
Public Function JsToExponential(ByVal x As Double, Optional ByVal fractionDigits As Variant) As String
    Dim cls As Long, neg As Boolean, m As Double, e As Long, D As String, P As Long, R As String
    Dim f As Long, ex As Long, s As String
    cls = Decompose(x, neg, m, e)
    If cls >= 2 Then JsToExponential = JsNumberToString(x): Exit Function
    If Not IsMissing(fractionDigits) Then
        f = CLng(fractionDigits)
        If f < 0 Or f > 100 Then Err.Raise 5, "JsToExponential", "toExponential() argument must be between 0 and 100"
    End If
    If cls = 0 Then
        neg = False
        ex = 0
        If IsMissing(fractionDigits) Then R = "0" Else R = String$(f + 1, "0")
    ElseIf IsMissing(fractionDigits) Then
        ShortestOf x, m, e, R, P
        ex = P - 1
    Else
        ExactDecimal m, e, D, P
        R = RoundDigits(D, f + 1)
        ex = P - 1
        If Len(R) > f + 1 Then
            R = Left$(R, f + 1)
            ex = ex + 1
        End If
    End If
    s = Left$(R, 1)
    If Len(R) > 1 Then s = s & "." & Mid$(R, 2)
    s = s & "e" & ExpSign(ex)
    If neg Then JsToExponential = "-" & s Else JsToExponential = s
End Function

' Number(x).toLocaleString("en-US", {minimumFractionDigits, maximumFractionDigits,
' useGrouping}). maxFrac < 0 means the default max(minFrac, 3). Rounds the shortest
' decimal representation half away from zero, like ICU. NaN -> "NaN", Infinity -> U+221E (infinity sign).
Public Function JsToLocaleString(ByVal x As Double, Optional ByVal minFrac As Long = 0, Optional ByVal maxFrac As Long = -1, Optional ByVal useGrouping As Boolean = True) As String
    Dim cls As Long, neg As Boolean, m As Double, e As Long, D As String, n As Long, R As String
    Dim intPart As String, fracPart As String, i As Long, g As String, cut As Long
    If maxFrac < 0 Then
        maxFrac = minFrac
        If maxFrac < 3 Then maxFrac = 3
    End If
    If minFrac < 0 Or minFrac > 100 Or maxFrac > 100 Or maxFrac < minFrac Then Err.Raise 5, "JsToLocaleString", "maximumFractionDigits value is out of range"
    cls = Decompose(x, neg, m, e)
    If cls = 3 Then JsToLocaleString = "NaN": Exit Function
    If cls = 2 Then
        If neg Then JsToLocaleString = "-" & ChrW$(&H221E) Else JsToLocaleString = ChrW$(&H221E)
        Exit Function
    End If
    If cls = 0 Then
        R = ""
    Else
        ShortestOf x, m, e, D, n
        R = RoundDigits(D, n + maxFrac)
    End If
    If Len(R) <= maxFrac Then R = String$(maxFrac + 1 - Len(R), "0") & R
    intPart = Left$(R, Len(R) - maxFrac)
    fracPart = Right$(R, maxFrac)
    cut = Len(fracPart)
    Do While cut > minFrac
        If Mid$(fracPart, cut, 1) <> "0" Then Exit Do
        cut = cut - 1
    Loop
    fracPart = Left$(fracPart, cut)
    If useGrouping And Len(intPart) > 3 Then
        g = ""
        i = Len(intPart)
        Do While i > 3
            g = "," & Mid$(intPart, i - 2, 3) & g
            i = i - 3
        Loop
        intPart = Left$(intPart, i) & g
    End If
    If Len(fracPart) > 0 Then intPart = intPart & "." & fracPart
    If neg Then JsToLocaleString = "-" & intPart Else JsToLocaleString = intPart
End Function

' -------------------------------------------------------------------------------------
' String -> Number
' -------------------------------------------------------------------------------------

Private Function Pow10Exact(ByVal k As Long) As Double
    Dim i As Long
    If Not mP10Ready Then
        mP10(0) = 1
        For i = 1 To 22
            mP10(i) = mP10(i - 1) * 10
        Next
        mP10Ready = True
    End If
    Pow10Exact = mP10(k)
End Function

Private Function DigitsValue(ByVal D As String) As Double
    Dim i As Long, v As Double
    For i = 1 To Len(D)
        v = v * 10 + (AscW(Mid$(D, i, 1)) - 48)
    Next
    DigitsValue = v
End Function

Private Function CmpDec(ByRef D1 As String, ByVal P1 As Long, ByRef D2 As String, ByVal P2 As Long) As Long
    Dim l1 As Long, l2 As Long
    If P1 <> P2 Then
        If P1 > P2 Then CmpDec = 1 Else CmpDec = -1
        Exit Function
    End If
    l1 = Len(D1)
    l2 = Len(D2)
    If l1 = l2 Then
        CmpDec = StrComp(D1, D2, vbBinaryCompare)
    ElseIf l1 < l2 Then
        CmpDec = StrComp(D1 & String$(l2 - l1, "0"), D2, vbBinaryCompare)
    Else
        CmpDec = StrComp(D1, D2 & String$(l1 - l2, "0"), vbBinaryCompare)
    End If
End Function

' Moves a non-negative approximation c to the correctly rounded double of 0.D * 10^P.
Private Function RefineDecimal(ByRef D As String, ByVal P As Long, ByVal c As Double) As Double
    Dim m As Double, e As Long, neg As Boolean, k As Long, s As String, cmp As Long, odd As Boolean
    If Not JsIsFinite(c) Then c = MaxDouble()
    Do
        If Decompose(c, neg, m, e) = 0 Then
            m = 0
            e = -1074
        End If
        odd = (m - 2 * Int(m / 2)) <> 0
        s = BigExact(m, 2, 1, e - 1, k)
        cmp = CmpDec(D, P, s, Len(s) - k)
        If cmp > 0 Or (cmp = 0 And odd) Then
            c = NextUpPos(c)
            If Not JsIsFinite(c) Then Exit Do
        ElseIf m = 0 Then
            Exit Do
        Else
            If m = TWO52 And e > -1074 Then s = BigExact(m, 4, -1, e - 2, k) Else s = BigExact(m, 2, -1, e - 1, k)
            cmp = CmpDec(D, P, s, Len(s) - k)
            If cmp < 0 Or (cmp = 0 And odd) Then c = NextDownPos(c) Else Exit Do
        End If
    Loop
    RefineDecimal = c
End Function

Private Function IsOddMantissa(ByVal x As Double) As Boolean
    Dim hi As Long, lo As Long
    DblWords x, hi, lo
    IsOddMantissa = ((lo And 1) <> 0)
End Function

' Correctly rounded double (ties to even) of 0.D * 10^P; D has no leading zeros.
Private Function DecToDouble(ByVal D As String, ByVal P As Long) As Double
    Dim nd As Long, q As Long, v As Double, na As Long, a As Double, s As Long
    D = StripTrailingZeros(D)
    If Len(D) = 0 Then Exit Function
    If P > 310 Then DecToDouble = JsInfinity(1): Exit Function
    If P < -323 Then Exit Function
    nd = Len(D)
    q = P - nd
    If nd <= 15 And q >= -22 And q <= 22 Then
        v = DigitsValue(D)
        If q >= 0 Then a = v * Pow10Exact(q) Else a = v / Pow10Exact(-q)
        ' VBA rounds to 64-bit precision before rounding to double; a double-rounding
        ' error always lands on an even mantissa, so only those results are verified.
        If IsOddMantissa(a) Then DecToDouble = a Else DecToDouble = RefineDecimal(D, P, a)
        Exit Function
    End If
    na = nd
    If na > 17 Then na = 17
    a = DigitsValue(Left$(D, na))
    q = P - na
    Do While q <> 0
        s = q
        If s > 300 Then s = 300
        If s < -300 Then s = -300
        a = JsMul(a, 10# ^ s)
        q = q - s
    Loop
    DecToDouble = RefineDecimal(D, P, a)
End Function

' Value of a digit string in a power-of-two radix, correctly rounded (ties to even).
Private Function RadixPow2Value(ByVal digits As String, ByVal radix As Long) As Double
    Dim bitsPer As Long, i As Long, j As Long, v As Long, b As String, pos As Long
    Dim n As Long, first As Long, r As Double, roundUp As Boolean
    bitsPer = 1
    Do While 2 ^ bitsPer < radix
        bitsPer = bitsPer + 1
    Loop
    b = String$(Len(digits) * bitsPer, "0")
    pos = 1
    For i = 1 To Len(digits)
        v = DigitValue(AscW(Mid$(digits, i, 1)))
        For j = bitsPer - 1 To 0 Step -1
            If (v And (2 ^ j)) <> 0 Then Mid$(b, pos, 1) = "1"
            pos = pos + 1
        Next
    Next
    first = InStr(1, b, "1", vbBinaryCompare)
    If first = 0 Then Exit Function
    b = Mid$(b, first)
    n = Len(b)
    For i = 1 To n
        If i > 53 Then Exit For
        r = r * 2
        If Mid$(b, i, 1) = "1" Then r = r + 1
    Next
    If n <= 53 Then RadixPow2Value = r: Exit Function
    If Mid$(b, 54, 1) = "1" Then
        roundUp = (InStr(55, b, "1", vbBinaryCompare) > 0) Or (r - 2 * Int(r / 2) <> 0)
        If roundUp Then r = r + 1
    End If
    n = n - 53
    If n > 1100 Then RadixPow2Value = JsInfinity(1): Exit Function
    Do While n > 0
        If n > 1000 Then
            r = JsMul(r, Pow2Dbl(1000))
            n = n - 1000
        Else
            r = JsMul(r, Pow2Dbl(n))
            n = 0
        End If
    Loop
    RadixPow2Value = r
End Function

Private Function DigitValue(ByVal c As Long) As Long
    Select Case c
    Case 48 To 57: DigitValue = c - 48
    Case 65 To 90: DigitValue = c - 55
    Case 97 To 122: DigitValue = c - 87
    Case Else: DigitValue = 99
    End Select
End Function

Private Function IsDigitCode(ByVal c As Long) As Boolean
    IsDigitCode = (c >= 48 And c <= 57)
End Function

Private Function CodeAt(ByRef s As String, ByVal i As Long) As Long
    If i < 1 Or i > Len(s) Then
        CodeAt = -1
    Else
        CodeAt = AscW(Mid$(s, i, 1)) And &HFFFF&
    End If
End Function

' Longest StrDecimalLiteral prefix of s starting at pos (1-based): sign, Infinity,
' digits [. digits] [e [+-] digits]. endPos = pos when nothing matched (result NaN).
Private Function ParseDecimalPrefix(ByRef s As String, ByVal pos As Long, ByRef endPos As Long) As Double
    Dim i As Long, n As Long, neg As Boolean, c As Long, intStart As Long, intLen As Long
    Dim fracStart As Long, fracLen As Long, ex As Long, exNeg As Boolean, j As Long
    Dim D As String, P As Long, first As Long, v As Double
    n = Len(s)
    i = pos
    endPos = pos
    ParseDecimalPrefix = JsNaN()
    c = CodeAt(s, i)
    If c = 43 Or c = 45 Then
        neg = (c = 45)
        i = i + 1
    End If
    If Mid$(s, i, 8) = "Infinity" Then
        endPos = i + 8
        If neg Then ParseDecimalPrefix = JsInfinity(-1) Else ParseDecimalPrefix = JsInfinity(1)
        Exit Function
    End If
    intStart = i
    Do While IsDigitCode(CodeAt(s, i))
        i = i + 1
    Loop
    intLen = i - intStart
    If CodeAt(s, i) = 46 Then
        fracStart = i + 1
        j = fracStart
        Do While IsDigitCode(CodeAt(s, j))
            j = j + 1
        Loop
        fracLen = j - fracStart
        If intLen + fracLen > 0 Then i = j
    End If
    If intLen + fracLen = 0 Then Exit Function
    endPos = i
    c = CodeAt(s, i)
    If c = 101 Or c = 69 Then
        j = i + 1
        c = CodeAt(s, j)
        If c = 43 Or c = 45 Then
            exNeg = (c = 45)
            j = j + 1
        End If
        If IsDigitCode(CodeAt(s, j)) Then
            Do While IsDigitCode(CodeAt(s, j))
                If ex < 10000000 Then ex = ex * 10 + (CodeAt(s, j) - 48)
                j = j + 1
            Loop
            endPos = j
            If exNeg Then ex = -ex
        End If
    End If
    D = Mid$(s, intStart, intLen)
    If fracLen > 0 Then D = D & Mid$(s, fracStart, fracLen)
    P = intLen + ex
    first = 1
    Do While first <= Len(D)
        If Mid$(D, first, 1) <> "0" Then Exit Do
        first = first + 1
    Loop
    P = P - (first - 1)
    D = Mid$(D, first)
    v = DecToDouble(D, P)
    If neg Then v = -v
    ParseDecimalPrefix = v
End Function

' Number(text): ECMAScript StringToNumber. "" and whitespace -> 0; invalid -> NaN.
Public Function JsNumber(ByVal text As String) As Double
    Dim s As String, n As Long, endPos As Long, radix As Long, i As Long, body As String
    s = JsTrim(text)
    n = Len(s)
    If n = 0 Then JsNumber = 0: Exit Function
    If n > 2 And Left$(s, 1) = "0" Then
        Select Case Mid$(s, 2, 1)
        Case "x", "X": radix = 16
        Case "o", "O": radix = 8
        Case "b", "B": radix = 2
        End Select
        If radix > 0 Then
            body = Mid$(s, 3)
            For i = 1 To Len(body)
                If DigitValue(AscW(Mid$(body, i, 1)) And &HFFFF&) >= radix Then JsNumber = JsNaN(): Exit Function
            Next
            JsNumber = RadixPow2Value(body, radix)
            Exit Function
        End If
    End If
    JsNumber = ParseDecimalPrefix(s, 1, endPos)
    If endPos <> n + 1 Then JsNumber = JsNaN()
End Function

' parseFloat(text)
Public Function JsParseFloat(ByVal text As String) As Double
    Dim endPos As Long
    JsParseFloat = ParseDecimalPrefix(JsTrimStart(text), 1, endPos)
End Function

' parseInt(text, radix). Exact for radix 10 and powers of two; other radices are exact
' while the value stays below 2^53.
Public Function JsParseInt(ByVal text As String, Optional ByVal radix As Variant) As Double
    Dim s As String, neg As Boolean, r As Long, stripPrefix As Boolean, i As Long, c As Long
    Dim digits As String, v As Double, first As Long
    s = JsTrimStart(text)
    c = CodeAt(s, 1)
    If c = 45 Or c = 43 Then
        neg = (c = 45)
        s = Mid$(s, 2)
    End If
    stripPrefix = True
    If Not IsMissing(radix) Then r = JsToInt32(CDbl(radix))
    If r <> 0 Then
        If r < 2 Or r > 36 Then JsParseInt = JsNaN(): Exit Function
        If r <> 16 Then stripPrefix = False
    Else
        r = 10
    End If
    If stripPrefix Then
        If Left$(s, 2) = "0x" Or Left$(s, 2) = "0X" Then
            s = Mid$(s, 3)
            r = 16
        End If
    End If
    i = 1
    Do While i <= Len(s)
        If DigitValue(CodeAt(s, i)) >= r Then Exit Do
        i = i + 1
    Loop
    digits = Left$(s, i - 1)
    If Len(digits) = 0 Then JsParseInt = JsNaN(): Exit Function
    Select Case r
    Case 10
        first = 1
        Do While first < Len(digits)
            If Mid$(digits, first, 1) <> "0" Then Exit Do
            first = first + 1
        Loop
        digits = Mid$(digits, first)
        If digits = "0" Then v = 0 Else v = DecToDouble(digits, Len(digits))
    Case 2, 4, 8, 16, 32
        v = RadixPow2Value(digits, r)
    Case Else
        For i = 1 To Len(digits)
            v = JsAdd(JsMul(v, r), DigitValue(CodeAt(digits, i)))
        Next
    End Select
    If neg Then v = -v
    JsParseInt = v
End Function

' -------------------------------------------------------------------------------------
' Math.round and Math.random
' -------------------------------------------------------------------------------------

' Math.round: nearest integer, halves toward +Infinity; -0 for -0.5 <= x < 0.
Public Function JsMathRound(ByVal x As Double) As Double
    Dim f As Double
    If Not JsIsFinite(x) Then JsMathRound = x: Exit Function
    If Abs(x) >= TWO52 Then JsMathRound = x: Exit Function
    f = Int(x)
    If x - f >= 0.5 Then f = f + 1
    If f = 0 And JsSignBit(x) Then f = NegZero()
    JsMathRound = f
End Function

Private Function U32(ByVal x As Double) As Double
    U32 = x - Int(x / TWO32) * TWO32
End Function

Private Function ToSigned(ByVal u As Double) As Long
    If u >= TWO31 Then ToSigned = CLng(u - TWO32) Else ToSigned = CLng(u)
End Function

Private Function Xor32(ByVal a As Double, ByVal b As Double) As Double
    Xor32 = ULong(ToSigned(a) Xor ToSigned(b))
End Function

Private Function Or32(ByVal a As Double, ByVal b As Double) As Double
    Or32 = ULong(ToSigned(a) Or ToSigned(b))
End Function

Private Function Imul32(ByVal a As Double, ByVal b As Double) As Double
    Dim ah As Double, al As Double, bh As Double, bl As Double, mid As Double
    ah = Int(a / 65536#)
    al = a - ah * 65536#
    bh = Int(b / 65536#)
    bl = b - bh * 65536#
    mid = ah * bl + al * bh
    mid = mid - Int(mid / 65536#) * 65536#
    Imul32 = U32(al * bl + mid * 65536#)
End Function

' Seeds JsRandom (mulberry32 generator) for reproducible sequences.
Public Sub JsSeedRandom(ByVal seed As Double)
    mRandState = JsToUint32(seed)
    mRandSeeded = True
End Sub

' Math.random(): uniform in [0, 1). Unseeded, it seeds itself from the clock.
' mulberry32 with the 32-bit helpers (U32, Xor32, Or32, Imul32) written out inline: the
' answer-option generators draw thousands of numbers per edit. Same arithmetic, same bits.
Public Function JsRandom() As Double
    Dim a As Double, t As Double, u As Double, sa As Long, sb As Long, p As Double, q As Double
    Dim ph As Double, pl As Double, qh As Double, ql As Double, m As Double
    If Not mRandSeeded Then JsSeedRandom Timer * 1000003# + CDbl(Now) * 86400#
    ' a = (state + 0x6D2B79F5) >>> 0
    a = mRandState + 1831565813#
    If a >= TWO32 Then a = a - TWO32
    mRandState = a
    ' t = Math.imul(a ^ (a >>> 15), a | 1)
    If a >= TWO31 Then sa = CLng(a - TWO32) Else sa = CLng(a)
    u = Int(a / 32768#)
    sb = CLng(u)
    sb = sa Xor sb
    If sb < 0 Then p = CDbl(sb) + TWO32 Else p = sb
    sb = sa Or 1
    If sb < 0 Then q = CDbl(sb) + TWO32 Else q = sb
    ph = Int(p / 65536#): pl = p - ph * 65536#
    qh = Int(q / 65536#): ql = q - qh * 65536#
    m = ph * ql + pl * qh
    m = m - Int(m / 65536#) * 65536#
    t = pl * ql + m * 65536#
    t = t - Int(t / TWO32) * TWO32
    ' t ^= t + Math.imul(t ^ (t >>> 7), t | 61)
    If t >= TWO31 Then sa = CLng(t - TWO32) Else sa = CLng(t)
    u = Int(t / 128#)
    sb = CLng(u)
    sb = sa Xor sb
    If sb < 0 Then p = CDbl(sb) + TWO32 Else p = sb
    sb = sa Or 61
    If sb < 0 Then q = CDbl(sb) + TWO32 Else q = sb
    ph = Int(p / 65536#): pl = p - ph * 65536#
    qh = Int(q / 65536#): ql = q - qh * 65536#
    m = ph * ql + pl * qh
    m = m - Int(m / 65536#) * 65536#
    u = pl * ql + m * 65536#
    u = u - Int(u / TWO32) * TWO32
    u = t + u
    If u >= TWO32 Then u = u - TWO32
    If u >= TWO31 Then sb = CLng(u - TWO32) Else sb = CLng(u)
    sa = sa Xor sb
    If sa < 0 Then t = CDbl(sa) + TWO32 Else t = sa
    ' ((t ^ (t >>> 14)) >>> 0) / 2^32
    If t >= TWO31 Then sa = CLng(t - TWO32) Else sa = CLng(t)
    u = Int(t / 16384#)
    sb = CLng(u)
    sa = sa Xor sb
    If sa < 0 Then JsRandom = (CDbl(sa) + TWO32) / TWO32 Else JsRandom = sa / TWO32
End Function

' -------------------------------------------------------------------------------------
' Strings (0-based UTF-16 indices)
' -------------------------------------------------------------------------------------

' ECMAScript WhiteSpace and LineTerminator code units.
Public Function JsIsWhitespaceCode(ByVal c As Long) As Boolean
    Select Case c
    Case 9 To 13, 32, 160, 5760, 8192 To 8202, 8232, 8233, 8239, 8287, 12288, 65279
        JsIsWhitespaceCode = True
    End Select
End Function

Public Function JsTrim(ByVal text As String) As String
    Dim a As Long, b As Long
    a = 1
    b = Len(text)
    Do While a <= b
        If Not JsIsWhitespaceCode(AscW(Mid$(text, a, 1)) And &HFFFF&) Then Exit Do
        a = a + 1
    Loop
    Do While b >= a
        If Not JsIsWhitespaceCode(AscW(Mid$(text, b, 1)) And &HFFFF&) Then Exit Do
        b = b - 1
    Loop
    JsTrim = Mid$(text, a, b - a + 1)
End Function

Public Function JsTrimStart(ByVal text As String) As String
    Dim a As Long, n As Long
    a = 1
    n = Len(text)
    Do While a <= n
        If Not JsIsWhitespaceCode(AscW(Mid$(text, a, 1)) And &HFFFF&) Then Exit Do
        a = a + 1
    Loop
    JsTrimStart = Mid$(text, a)
End Function

Public Function JsTrimEnd(ByVal text As String) As String
    Dim b As Long
    b = Len(text)
    Do While b >= 1
        If Not JsIsWhitespaceCode(AscW(Mid$(text, b, 1)) And &HFFFF&) Then Exit Do
        b = b - 1
    Loop
    JsTrimEnd = Left$(text, b)
End Function

' Relative index (slice/at semantics): negative counts from the end, clamped to [0, n].
Private Function RelIndex(ByVal v As Variant, ByVal n As Long, ByVal dflt As Long) As Long
    Dim d As Double
    If IsMissing(v) Then RelIndex = dflt: Exit Function
    d = CDbl(v)
    If JsIsNaN(d) Then d = 0
    d = JsTrunc(d)
    If d < 0 Then
        d = d + n
        If d < 0 Then d = 0
    ElseIf d > n Then
        d = n
    End If
    RelIndex = CLng(d)
End Function

' Clamped index (substring semantics): negative -> 0, > n -> n.
Private Function ClampIndex(ByVal v As Variant, ByVal n As Long, ByVal dflt As Long) As Long
    Dim d As Double
    If IsMissing(v) Then ClampIndex = dflt: Exit Function
    d = CDbl(v)
    If JsIsNaN(d) Then d = 0
    If d < 0 Then d = 0
    If d > n Then d = n
    ClampIndex = CLng(JsTrunc(d))
End Function

' text.slice(start, [end])
Public Function JsSlice(ByVal text As String, ByVal start As Variant, Optional ByVal endIndex As Variant) As String
    Dim n As Long, a As Long, b As Long
    n = Len(text)
    a = RelIndex(start, n, 0)
    b = RelIndex(endIndex, n, n)
    If b > a Then JsSlice = Mid$(text, a + 1, b - a)
End Function

' text.substring(start, [end]) (negative -> 0, arguments swapped when start > end)
Public Function JsSubstring(ByVal text As String, ByVal start As Variant, Optional ByVal endIndex As Variant) As String
    Dim n As Long, a As Long, b As Long, t As Long
    n = Len(text)
    a = ClampIndex(start, n, 0)
    b = ClampIndex(endIndex, n, n)
    If a > b Then
        t = a
        a = b
        b = t
    End If
    If b > a Then JsSubstring = Mid$(text, a + 1, b - a)
End Function

' text.indexOf(search, [fromIndex]) -> 0-based, -1 when absent.
Public Function JsIndexOf(ByVal text As String, ByVal search As String, Optional ByVal fromIndex As Variant) As Long
    Dim n As Long, f As Long
    n = Len(text)
    f = ClampIndex(fromIndex, n, 0)
    If Len(search) = 0 Then JsIndexOf = f: Exit Function
    JsIndexOf = InStr(f + 1, text, search, vbBinaryCompare) - 1
End Function

' text.lastIndexOf(search, [fromIndex])
Public Function JsLastIndexOf(ByVal text As String, ByVal search As String, Optional ByVal fromIndex As Variant) As Long
    Dim n As Long, f As Long, lim As Long, d As Double
    n = Len(text)
    f = n
    If Not IsMissing(fromIndex) Then
        d = CDbl(fromIndex)
        If Not JsIsNaN(d) Then f = ClampIndex(d, n, n)
    End If
    If Len(search) = 0 Then JsLastIndexOf = f: Exit Function
    lim = f + Len(search)
    If lim > n Then lim = n
    JsLastIndexOf = InStrRev(Left$(text, lim), search, -1, vbBinaryCompare) - 1
End Function

Public Function JsIncludes(ByVal text As String, ByVal search As String, Optional ByVal fromIndex As Variant) As Boolean
    JsIncludes = (JsIndexOf(text, search, fromIndex) >= 0)
End Function

Public Function JsStartsWith(ByVal text As String, ByVal search As String, Optional ByVal position As Variant) As Boolean
    Dim p As Long
    p = ClampIndex(position, Len(text), 0)
    If p + Len(search) > Len(text) Then Exit Function
    JsStartsWith = (Mid$(text, p + 1, Len(search)) = search)
End Function

Public Function JsEndsWith(ByVal text As String, ByVal search As String, Optional ByVal endPosition As Variant) As Boolean
    Dim p As Long, st As Long
    p = ClampIndex(endPosition, Len(text), Len(text))
    st = p - Len(search)
    If st < 0 Then Exit Function
    JsEndsWith = (Mid$(text, st + 1, Len(search)) = search)
End Function

Public Function JsRepeat(ByVal text As String, ByVal count As Long) As String
    Dim r As String
    If count < 0 Then Err.Raise 5, "JsRepeat", "Invalid count value"
    Do While count > 0
        If count And 1 Then r = r & text
        count = count \ 2
        If count > 0 Then text = text & text
    Loop
    JsRepeat = r
End Function

Private Function PadFill(ByVal text As String, ByVal targetLength As Long, ByVal padString As String) As String
    Dim need As Long, fill As String
    need = targetLength - Len(text)
    If need <= 0 Or Len(padString) = 0 Then Exit Function
    fill = JsRepeat(padString, need \ Len(padString) + 1)
    PadFill = Left$(fill, need)
End Function

Public Function JsPadStart(ByVal text As String, ByVal targetLength As Long, Optional ByVal padString As String = " ") As String
    JsPadStart = PadFill(text, targetLength, padString) & text
End Function

Public Function JsPadEnd(ByVal text As String, ByVal targetLength As Long, Optional ByVal padString As String = " ") As String
    JsPadEnd = text & PadFill(text, targetLength, padString)
End Function

' text.charAt(i): "" when out of range.
Public Function JsCharAt(ByVal text As String, ByVal index As Long) As String
    If index >= 0 And index < Len(text) Then JsCharAt = Mid$(text, index + 1, 1)
End Function

' text.charCodeAt(i): NaN when out of range.
Public Function JsCharCodeAt(ByVal text As String, ByVal index As Long) As Double
    If index >= 0 And index < Len(text) Then
        JsCharCodeAt = AscW(Mid$(text, index + 1, 1)) And &HFFFF&
    Else
        JsCharCodeAt = JsNaN()
    End If
End Function

' text.at(i): negative counts from the end; "" (JS undefined) when out of range.
Public Function JsAt(ByVal text As String, ByVal index As Long) As String
    If index < 0 Then index = index + Len(text)
    If index >= 0 And index < Len(text) Then JsAt = Mid$(text, index + 1, 1)
End Function

' String.fromCharCode(...codes)
Public Function JsFromCharCode(ParamArray codes() As Variant) As String
    Dim i As Long, s As String
    For i = LBound(codes) To UBound(codes)
        s = s & ChrW$(CLng(JsToUint32(CDbl(codes(i))) - Int(JsToUint32(CDbl(codes(i))) / 65536#) * 65536#))
    Next
    JsFromCharCode = s
End Function

' text.split(separator, [limit]) with a string separator; omitted separator -> [text].
Public Function JsSplit(ByVal text As String, Optional ByVal separator As Variant, Optional ByVal limit As Variant) As clsList
    Dim lst As clsList, lim As Double, sep As String, p As Long, q As Long, i As Long
    Set lst = New clsList
    Set JsSplit = lst
    lim = 4294967295#
    If Not IsMissing(limit) Then lim = JsToUint32(CDbl(limit))
    If lim = 0 Then Exit Function
    If IsMissing(separator) Then lst.Add text: Exit Function
    sep = CStr(separator)
    If Len(sep) = 0 Then
        For i = 1 To Len(text)
            If lst.Count >= lim Then Exit Function
            lst.Add Mid$(text, i, 1)
        Next
        Exit Function
    End If
    p = 1
    Do
        q = InStr(p, text, sep, vbBinaryCompare)
        If q = 0 Then Exit Do
        lst.Add Mid$(text, p, q - p)
        If lst.Count >= lim Then Exit Function
        p = q + Len(sep)
    Loop
    lst.Add Mid$(text, p)
End Function

' text.replaceAll(search, replacement) for a literal search string ($ patterns in the
' replacement are not interpreted). An empty search inserts between every code unit.
Public Function JsReplaceAll(ByVal text As String, ByVal search As String, ByVal replacement As String) As String
    Dim i As Long, s As String
    If Len(search) > 0 Then
        JsReplaceAll = Replace(text, search, replacement, 1, -1, vbBinaryCompare)
    Else
        s = replacement
        For i = 1 To Len(text)
            s = s & Mid$(text, i, 1) & replacement
        Next
        JsReplaceAll = s
    End If
End Function

Private Function CaseMap(ByVal text As String, ByVal upper As Boolean) As String
    Dim b() As Byte, i As Long, c As Long, t As String
    If LenB(text) = 0 Then Exit Function
    b = text
    For i = 0 To UBound(b) - 1 Step 2
        If b(i + 1) = 0 And b(i) < 128 Then
            If upper Then
                If b(i) >= 97 And b(i) <= 122 Then b(i) = b(i) - 32
            Else
                If b(i) >= 65 And b(i) <= 90 Then b(i) = b(i) + 32
            End If
        Else
            c = CLng(b(i)) + 256& * b(i + 1)
            If upper Then t = UCase$(ChrW$(c)) Else t = LCase$(ChrW$(c))
            If Len(t) = 1 Then
                c = AscW(t) And &HFFFF&
                b(i) = c And &HFF&
                b(i + 1) = c \ 256
            End If
        End If
    Next
    CaseMap = b
End Function

' toLowerCase / toUpperCase (exact for ASCII; other letters use VBA's single-character
' mapping, so special cases such as final sigma or "ss" for U+00DF are not applied).
Public Function JsLower(ByVal text As String) As String
    JsLower = CaseMap(text, False)
End Function

Public Function JsUpper(ByVal text As String) As String
    JsUpper = CaseMap(text, True)
End Function

' Decodes \uXXXX escapes (handy for writing non-ASCII literals in ASCII sources).
Public Function JsU(ByVal text As String) As String
    Dim p As Long, h As String
    p = InStr(1, text, "\u", vbBinaryCompare)
    Do While p > 0
        h = Mid$(text, p + 2, 4)
        If Len(h) = 4 And IsHex4(h) Then
            text = Left$(text, p - 1) & ChrW$(HexWord(h)) & Mid$(text, p + 6)
        End If
        p = InStr(p + 1, text, "\u", vbBinaryCompare)
    Loop
    JsU = text
End Function

Private Function IsHex4(ByVal h As String) As Boolean
    Dim i As Long, c As Long
    For i = 1 To 4
        c = AscW(Mid$(h, i, 1))
        If Not ((c >= 48 And c <= 57) Or (c >= 65 And c <= 70) Or (c >= 97 And c <= 102)) Then Exit Function
    Next
    IsHex4 = True
End Function

' -------------------------------------------------------------------------------------
' Conversions and constructors
' -------------------------------------------------------------------------------------

' String(value): numbers via JsNumberToString, Booleans "true"/"false", Empty
' "undefined", Null "null", clsList joined with ",".
Public Function JsStr(ByVal value As Variant) As String
    If IsObject(value) Then
        If value Is Nothing Then
            JsStr = "null"
        ElseIf TypeOf value Is clsList Then
            JsStr = value.Join(",")
        Else
            JsStr = "[object " & TypeName(value) & "]"
        End If
        Exit Function
    End If
    Select Case VarType(value)
    Case vbEmpty: JsStr = "undefined"
    Case vbNull: JsStr = "null"
    Case vbBoolean: If value Then JsStr = "true" Else JsStr = "false"
    Case vbString: JsStr = value
    Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
        JsStr = JsNumberToString(CDbl(value))
    Case Else
        JsStr = CStr(value)
    End Select
End Function

' Number(value): strings via JsNumber, Booleans 1/0, Empty NaN, Null 0.
Public Function JsToNumber(ByVal value As Variant) As Double
    If IsObject(value) Then JsToNumber = JsNaN(): Exit Function
    Select Case VarType(value)
    Case vbEmpty: JsToNumber = JsNaN()
    Case vbNull: JsToNumber = 0
    Case vbBoolean: If value Then JsToNumber = 1 Else JsToNumber = 0
    Case vbString: JsToNumber = JsNumber(value)
    Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
        JsToNumber = CDbl(value)
    Case Else
        JsToNumber = JsNaN()
    End Select
End Function

Public Function NewList(ParamArray items() As Variant) As clsList
    Dim lst As clsList, i As Long
    Set lst = New clsList
    For i = LBound(items) To UBound(items)
        lst.Add items(i)
    Next
    Set NewList = lst
End Function

Public Function NewMap() As clsMap
    Set NewMap = New clsMap
End Function

' -------------------------------------------------------------------------------------
' Symbols printed by the website (ASCII sources build them with ChrW$)
' -------------------------------------------------------------------------------------

Public Function SYM_SUB0() As String: SYM_SUB0 = ChrW$(&H2080): End Function
Public Function SYM_SUB1() As String: SYM_SUB1 = ChrW$(&H2081): End Function
Public Function SYM_SUB2() As String: SYM_SUB2 = ChrW$(&H2082): End Function
Public Function SYM_SUBK() As String: SYM_SUBK = ChrW$(&H2096): End Function
Public Function SYM_H0() As String: SYM_H0 = "H" & ChrW$(&H2080): End Function
Public Function SYM_H1() As String: SYM_H1 = "H" & ChrW$(&H2081): End Function
Public Function SYM_SQUARED() As String: SYM_SQUARED = ChrW$(&HB2): End Function
Public Function SYM_MU() As String: SYM_MU = ChrW$(&H3BC): End Function
Public Function SYM_SIGMA() As String: SYM_SIGMA = ChrW$(&H3C3): End Function
Public Function SYM_RHO() As String: SYM_RHO = ChrW$(&H3C1): End Function
Public Function SYM_BETA() As String: SYM_BETA = ChrW$(&H3B2): End Function
Public Function SYM_CHI() As String: SYM_CHI = ChrW$(&H3C7): End Function
Public Function SYM_CHI2() As String: SYM_CHI2 = ChrW$(&H3C7) & ChrW$(&HB2): End Function
Public Function SYM_NE() As String: SYM_NE = ChrW$(&H2260): End Function
Public Function SYM_LE() As String: SYM_LE = ChrW$(&H2264): End Function
Public Function SYM_GE() As String: SYM_GE = ChrW$(&H2265): End Function
Public Function SYM_LE_SLANT() As String: SYM_LE_SLANT = ChrW$(&H2A7D): End Function
Public Function SYM_GE_SLANT() As String: SYM_GE_SLANT = ChrW$(&H2A7E): End Function
Public Function SYM_MINUS() As String: SYM_MINUS = ChrW$(&H2212): End Function
Public Function SYM_PLUSMINUS() As String: SYM_PLUSMINUS = ChrW$(&HB1): End Function
Public Function SYM_TIMES() As String: SYM_TIMES = ChrW$(&HD7): End Function
Public Function SYM_SQRT() As String: SYM_SQRT = ChrW$(&H221A): End Function
Public Function SYM_INFINITY() As String: SYM_INFINITY = ChrW$(&H221E): End Function
Public Function SYM_MIDDOT() As String: SYM_MIDDOT = ChrW$(&HB7): End Function
Public Function SYM_ARROW() As String: SYM_ARROW = ChrW$(&H2192): End Function
Public Function SYM_ELLIPSIS() As String: SYM_ELLIPSIS = ChrW$(&H2026): End Function
Public Function SYM_EMDASH() As String: SYM_EMDASH = ChrW$(&H2014): End Function
Public Function SYM_ENDASH() As String: SYM_ENDASH = ChrW$(&H2013): End Function
Public Function SYM_LSQUO() As String: SYM_LSQUO = ChrW$(&H2018): End Function
Public Function SYM_RSQUO() As String: SYM_RSQUO = ChrW$(&H2019): End Function
Public Function SYM_LDQUO() As String: SYM_LDQUO = ChrW$(&H201C): End Function
Public Function SYM_RDQUO() As String: SYM_RDQUO = ChrW$(&H201D): End Function
Public Function SYM_NBSP() As String: SYM_NBSP = ChrW$(&HA0): End Function
Public Function SYM_COMB_HAT() As String: SYM_COMB_HAT = ChrW$(&H302): End Function
Public Function SYM_COMB_BAR() As String: SYM_COMB_BAR = ChrW$(&H304): End Function
