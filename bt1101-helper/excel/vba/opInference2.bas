Attribute VB_Name = "opInference2"
' Tutorial 5 intervals: the distribution functions (inverseNormal, logGamma, betaFraction,
' regularizedBeta, studentTCdf, inverseStudentT), confidence/prediction intervals,
' sample sizes, their R code, pasted-result prose and the interval form.
' Ported from app-src.js lines 3453-3781 (event wiring: opInference.InitOpInference).
'
' Numerics reproduce V8 bit for bit: every floating-point operation is stored in a Double
' (VBA may keep expression intermediates in extended precision), constants are parsed
' from the website's literals with JsNumber or built from their bits, Math.log is
' modJs.JsLog (fdlibm) and Math.exp is FdExp below (fdlibm, as in V8). Math.sin (only
' reached by logGamma for z < 0.5, which the website never requests) uses VBA's Sin.
Option Explicit

Private Type TDblBox
    d As Double
End Type

Private Type TWordBox
    lo As Long
    hi As Long
End Type

Private Type TInterval
    calcType As String
    parameter As String
    stdev As String
    dataset As String
    label As String
    datasetProvided As Boolean
    labelProvided As Boolean
    confidence As Double
    estimate As Double
    estimateProvided As Boolean
    sd As Double
    sdProvided As Boolean
    n As Double
    nProvided As Boolean
    proportion As Double
    proportionProvided As Boolean
    proportionMode As String
    excludeMissing As Boolean
    population As String
    success As String
    conditionOperator As String
    threshold As Double
    thresholdProvided As Boolean
    margin As Double
    marginProvided As Boolean
End Type

' let intervalPastedProse = "";
Private mIntervalPastedProse As String
' The shell procedure that opens the known-population-SD modal (the website's
' showKnownPopulationSdAlert, in modApp); set by opInference.InitOpInference.
Public KnownSdAlertProcedure As String

Private mNumReady As Boolean
Private mNa(0 To 5) As Double, mNb(0 To 5) As Double, mNc(0 To 5) As Double, mNd(0 To 4) As Double
Private mNormLow As Double, mNormHigh As Double
Private mLg(0 To 7) As Double, mLgX0 As Double
Private mBfEps As Double, mBfTiny As Double
Private eOThreshold As Double, eUThreshold As Double, eLn2Hi As Double, eLn2Lo As Double, eInvLn2 As Double
Private eP1 As Double, eP2 As Double, eP3 As Double, eP4 As Double, eP5 As Double, eE As Double
Private eTwoM1000 As Double, eTwo1023 As Double

Public Sub ResetIntervalState()
    mIntervalPastedProse = ""
End Sub

' ---------------------------------------------------------------- numeric helpers

Private Sub InitNum()
    Dim i As Long, parts As Variant
    If mNumReady Then Exit Sub
    parts = Array("-39.6968302866538", "220.946098424521", "-275.928510446969", "138.357751867269", "-30.6647980661472", "2.50662827745924")
    For i = 0 To 5: mNa(i) = JsNumber(CStr(parts(i))): Next i
    parts = Array("-54.4760987982241", "161.585836858041", "-155.698979859887", "66.8013118877197", "-13.2806815528857", "1")
    For i = 0 To 5: mNb(i) = JsNumber(CStr(parts(i))): Next i
    parts = Array("-0.00778489400243029", "-0.322396458041136", "-2.40075827716184", "-2.54973253934373", "4.37466414146497", "2.93816398269878")
    For i = 0 To 5: mNc(i) = JsNumber(CStr(parts(i))): Next i
    parts = Array("0.00778469570904146", "0.32246712907004", "2.445134137143", "3.75440866190742", "1")
    For i = 0 To 4: mNd(i) = JsNumber(CStr(parts(i))): Next i
    mNormLow = JsNumber("0.02425")
    mNormHigh = 1 - mNormLow
    parts = Array("676.5203681218851", "-1259.1392167224028", "771.3234287776531", "-176.6150291621406", "12.507343278686905", "-0.13857109526572012", "9.984369578019571e-6", "1.5056327351493116e-7")
    For i = 0 To 7: mLg(i) = JsNumber(CStr(parts(i))): Next i
    mLgX0 = JsNumber("0.9999999999998099")
    mBfEps = JsNumber("3e-14")
    mBfTiny = JsNumber("1e-300")
    eOThreshold = JsHexToDouble("40862E42FEFA39EF")
    eUThreshold = JsHexToDouble("C0874910D52D3051")
    eLn2Hi = JsHexToDouble("3FE62E42FEE00000")
    eLn2Lo = JsHexToDouble("3DEA39EF35793C76")
    eInvLn2 = JsHexToDouble("3FF71547652B82FE")
    eP1 = JsHexToDouble("3FC555555555553E")
    eP2 = JsHexToDouble("BF66C16C16BEBD93")
    eP3 = JsHexToDouble("3F11566AAF25DE2C")
    eP4 = JsHexToDouble("BEBBBD41C5D26BF1")
    eP5 = JsHexToDouble("3E66376972BEA4D0")
    eE = JsHexToDouble("4005BF0A8B145769")
    eTwoM1000 = JsHexToDouble("0170000000000000")
    eTwo1023 = JsHexToDouble("7FE0000000000000")
    mNumReady = True
End Sub

Private Function HighWord(ByVal x As Double) As Long
    Dim a As TDblBox, b As TWordBox
    a.d = x
    LSet b = a
    HighWord = b.hi
End Function

Private Function FromWords(ByVal hi As Long, ByVal lo As Long) As Double
    Dim a As TDblBox, b As TWordBox
    b.hi = hi
    b.lo = lo
    LSet a = b
    FromWords = a.d
End Function

' Horner evaluation (((c0*x + c1)*x + c2)...), one rounding per operation like V8.
Private Function Horner(ByRef c() As Double, ByVal count As Long, ByVal x As Double) As Double
    Dim t As Double, i As Long
    t = c(0)
    For i = 1 To count - 1
        t = t * x
        t = t + c(i)
    Next i
    Horner = t
End Function

' Math.exp: port of fdlibm e_exp.c as used by V8 (base::ieee754::exp).
Public Function FdExp(ByVal x As Double) As Double
    Dim hx As Long, xsb As Long, k As Long
    Dim hi As Double, lo As Double, c As Double, t As Double, y As Double, a As Double, b As Double
    If JsIsNaN(x) Then FdExp = x: Exit Function
    InitNum
    hx = HighWord(x)
    If hx < 0 Then xsb = 1
    hx = hx And &H7FFFFFFF
    If hx >= &H40862E42 Then
        If hx >= &H7FF00000 Then
            If xsb = 0 Then FdExp = x Else FdExp = 0
            Exit Function
        End If
        If x > eOThreshold Then FdExp = JsInfinity(1): Exit Function
        If x < eUThreshold Then FdExp = 0: Exit Function
    End If
    If hx > &H3FD62E42 Then
        If hx < &H3FF0A2B2 Then
            If x = 1 Then FdExp = eE: Exit Function
            If xsb = 0 Then
                hi = x - eLn2Hi
                lo = eLn2Lo
                k = 1
            Else
                hi = x + eLn2Hi
                lo = -eLn2Lo
                k = -1
            End If
        Else
            a = eInvLn2 * x
            If xsb = 0 Then a = a + 0.5 Else a = a - 0.5
            k = Fix(a)
            t = k
            a = t * eLn2Hi
            hi = x - a
            lo = t * eLn2Lo
        End If
        x = hi - lo
    ElseIf hx < &H3E300000 Then
        FdExp = 1 + x
        Exit Function
    Else
        k = 0
    End If
    t = x * x
    a = t * eP5
    a = eP4 + a
    a = t * a
    a = eP3 + a
    a = t * a
    a = eP2 + a
    a = t * a
    a = eP1 + a
    a = t * a
    c = x - a
    If k = 0 Then
        a = x * c
        b = c - 2
        a = a / b
        a = a - x
        FdExp = 1 - a
        Exit Function
    End If
    a = x * c
    b = 2 - c
    a = a / b
    a = lo - a
    a = a - hi
    y = 1 - a
    If k >= -1021 Then
        If k = 1024 Then
            a = y * 2
            FdExp = JsMul(a, eTwo1023)
        Else
            FdExp = y * FromWords((k + 1023) * &H100000, 0)
        End If
    Else
        a = y * FromWords((k + 2023) * &H100000, 0)
        FdExp = a * eTwoM1000
    End If
End Function

' ---------------------------------------------------------------- distribution functions

Public Function InverseNormal(ByVal p As Double) As Double
    Dim q As Double, r As Double, num As Double, den As Double
    InitNum
    If JsIsNaN(p) Then InverseNormal = p: Exit Function
    If p <= 0 Or p >= 1 Then
        If p = 0 Then
            InverseNormal = JsInfinity(-1)
        ElseIf p = 1 Then
            InverseNormal = JsInfinity(1)
        Else
            InverseNormal = JsNaN()
        End If
        Exit Function
    End If
    If p < mNormLow Then
        q = JsLog(p)
        q = -2 * q
        q = Sqr(q)
        num = Horner(mNc, 6, q)
        den = Horner(mNd, 5, q)
        InverseNormal = num / den
        Exit Function
    End If
    If p > mNormHigh Then
        q = 1 - p
        q = JsLog(q)
        q = -2 * q
        q = Sqr(q)
        num = Horner(mNc, 6, q)
        den = Horner(mNd, 5, q)
        InverseNormal = -num / den
        Exit Function
    End If
    q = p - 0.5
    r = q * q
    num = Horner(mNa, 6, r)
    num = num * q
    den = Horner(mNb, 6, r)
    InverseNormal = num / den
End Function

Public Function LogGamma(ByVal z As Double) As Double
    Dim x As Double, t As Double, a As Double, b As Double, i As Long
    InitNum
    If JsIsNaN(z) Then LogGamma = z: Exit Function
    If z < 0.5 Then
        a = JsLog(JsPI())
        b = JsMul(JsPI(), z)
        b = JsSin(b)
        b = JsLog(b)
        a = JsSub(a, b)
        b = 1 - z
        LogGamma = JsSub(a, LogGamma(b))
        Exit Function
    End If
    z = z - 1
    x = mLgX0
    For i = 0 To 7
        t = z + i
        t = t + 1
        t = mLg(i) / t
        x = x + t
    Next i
    t = JsAdd(z, 8)
    t = JsSub(t, 0.5)
    a = 2 * JsPI()
    a = JsLog(a)
    a = 0.5 * a
    b = JsAdd(z, 0.5)
    b = JsMul(b, JsLog(t))
    a = JsAdd(a, b)
    a = JsSub(a, t)
    LogGamma = JsAdd(a, JsLog(x))
End Function

Public Function BetaFraction(ByVal a As Double, ByVal b As Double, ByVal x As Double) As Double
    Dim qab As Double, qap As Double, qam As Double, c As Double, d As Double, h As Double
    Dim m As Long, m2 As Double, aa As Double, t As Double, u As Double, delta As Double, dm As Double
    InitNum
    qab = a + b
    qap = a + 1
    qam = a - 1
    c = 1
    t = qab * x
    t = t / qap
    d = 1 - t
    If Abs(d) < mBfTiny Then d = mBfTiny
    d = 1 / d
    h = d
    For m = 1 To 200
        dm = m
        m2 = 2 * dm
        t = b - dm
        t = dm * t
        t = t * x
        u = qam + m2
        aa = a + m2
        u = u * aa
        aa = t / u
        t = aa * d
        d = 1 + t
        If Abs(d) < mBfTiny Then d = mBfTiny
        t = aa / c
        c = 1 + t
        If Abs(c) < mBfTiny Then c = mBfTiny
        d = 1 / d
        t = d * c
        h = h * t
        t = a + dm
        t = -t
        u = qab + dm
        t = t * u
        t = t * x
        u = a + m2
        aa = qap + m2
        u = u * aa
        aa = t / u
        t = aa * d
        d = 1 + t
        If Abs(d) < mBfTiny Then d = mBfTiny
        t = aa / c
        c = 1 + t
        If Abs(c) < mBfTiny Then c = mBfTiny
        d = 1 / d
        delta = d * c
        h = h * delta
        t = delta - 1
        If Abs(t) < mBfEps Then Exit For
    Next m
    BetaFraction = h
End Function

Public Function RegularizedBeta(ByVal x As Double, ByVal a As Double, ByVal b As Double) As Double
    Dim e As Double, t As Double, om As Double, front As Double, threshold As Double
    If Not JsIsNaN(x) Then
        If x <= 0 Then RegularizedBeta = 0: Exit Function
        If x >= 1 Then RegularizedBeta = 1: Exit Function
    End If
    If JsIsNaN(x) Or JsIsNaN(a) Or JsIsNaN(b) Then RegularizedBeta = JsNaN(): Exit Function
    t = JsAdd(a, b)
    e = LogGamma(t)
    e = JsSub(e, LogGamma(a))
    e = JsSub(e, LogGamma(b))
    t = JsMul(a, JsLog(x))
    e = JsAdd(e, t)
    om = 1 - x
    t = JsMul(b, JsLog(om))
    e = JsAdd(e, t)
    front = FdExp(e)
    threshold = JsAdd(a, 1)
    t = JsAdd(a, b)
    t = JsAdd(t, 2)
    threshold = JsDiv(threshold, t)
    If JsLt(x, threshold) Then
        t = JsMul(front, BetaFraction(a, b, x))
        RegularizedBeta = JsDiv(t, a)
    Else
        t = JsMul(front, BetaFraction(b, a, om))
        t = JsDiv(t, b)
        RegularizedBeta = JsSub(1, t)
    End If
End Function

Public Function StudentTCdf(ByVal t As Double, ByVal degrees As Double) As Double
    Dim x As Double, halfBeta As Double
    x = JsMul(t, t)
    x = JsAdd(degrees, x)
    x = JsDiv(degrees, x)
    halfBeta = JsMul(0.5, RegularizedBeta(x, JsDiv(degrees, 2), 0.5))
    If JsGe(t, 0) Then StudentTCdf = JsSub(1, halfBeta) Else StudentTCdf = halfBeta
End Function

Public Function InverseStudentT(ByVal p As Double, ByVal degrees As Double) As Double
    Dim low As Double, high As Double, mid As Double, i As Long
    If Not JsGt(degrees, 0) Or JsLe(p, 0) Or JsGe(p, 1) Then InverseStudentT = JsNaN(): Exit Function
    If JsLt(p, 0.5) Then
        mid = 1 - p
        InverseStudentT = -InverseStudentT(mid, degrees)
        Exit Function
    End If
    low = 0
    high = 1
    Do While JsLt(StudentTCdf(high, degrees), p) And high < 10000000#
        high = high * 2
    Loop
    For i = 0 To 99
        mid = low + high
        mid = mid / 2
        If JsLt(StudentTCdf(mid, degrees), p) Then low = mid Else high = mid
    Next i
    mid = low + high
    InverseStudentT = mid / 2
End Function

' ---------------------------------------------------------------- interval values and R code

Private Function FieldText(ByVal selector As String) As String
    FieldText = JsTrim(Q(selector).Value)
End Function

Private Function IntervalValues() As TInterval
    Dim v As TInterval, text As String
    text = FieldText("#intervalDataset")
    v.datasetProvided = Len(text) > 0
    If v.datasetProvided Then v.dataset = text Else v.dataset = "dataset"
    text = FieldText("#intervalLabel")
    v.labelProvided = Len(text) > 0
    If v.labelProvided Then v.label = text Else v.label = "column1"
    v.calcType = Q("#intervalCalcType").Value
    v.parameter = Q("#intervalParameter").Value
    v.stdev = Q("#intervalStdev").Value
    v.confidence = JsNumber(Q("#confidenceLevel").Value)
    text = FieldText("#intervalEstimate")
    v.estimate = JsNumber(text)
    v.estimateProvided = Len(text) > 0
    text = FieldText("#intervalSd")
    v.sd = JsNumber(text)
    v.sdProvided = Len(text) > 0
    text = FieldText("#intervalN")
    v.n = JsNumber(text)
    v.nProvided = Len(text) > 0
    text = FieldText("#intervalProportion")
    v.proportion = JsNumber(text)
    v.proportionProvided = Len(text) > 0
    v.proportionMode = Q("#intervalProportionMode").Value
    v.excludeMissing = Q("#intervalExcludeMissing").Checked
    v.population = FieldText("#intervalPopulation")
    If Len(v.population) = 0 Then v.population = "observations"
    v.success = FieldText("#intervalSuccess")
    If Len(v.success) = 0 Then v.success = "yes"
    v.conditionOperator = Q("#intervalConditionOperator").Value
    text = FieldText("#intervalConditionThreshold")
    v.threshold = JsNumber(text)
    v.thresholdProvided = Len(text) > 0
    text = FieldText("#intervalMargin")
    v.margin = JsNumber(text)
    v.marginProvided = Len(text) > 0
    IntervalValues = v
End Function

Private Function SafeLabel(ByRef v As TInterval) As String
    SafeLabel = RxReplace(v.label, "[^A-Za-z0-9]", "", "g")
    If Len(SafeLabel) = 0 Then SafeLabel = "column1"
End Function

Private Function SafeLevel(ByRef v As TInterval) As String
    SafeLevel = RxReplace(JsNumberToString(v.confidence), "[^0-9]", "", "g")
    If Len(SafeLevel) = 0 Then SafeLevel = "95"
End Function

' String(Number(x.toPrecision(12)))
Private Function RTail(ByVal x As Double) As String
    RTail = JsNumberToString(JsNumber(JsToPrecision(x, 12)))
End Function

Private Function BuildMeanIntervalRCode(ByRef v As TInterval, ByVal tailProbability As String) As String
    Dim dataRef As String, intervalType As String, upperName As String, lowerName As String, common As String
    Dim critical As String, spread As String
    dataRef = RColumnName(v.dataset) & "$" & RColumnName(v.label)
    If v.calcType = "prediction" Then intervalType = "PI" Else intervalType = "CI"
    upperName = "u" & intervalType & SafeLabel(v) & SafeLevel(v)
    lowerName = "l" & intervalType & SafeLabel(v) & SafeLevel(v)
    common = "values <- na.omit(" & dataRef & ")" & vbLf & "n <- length(values)" & vbLf & "tail_probability <- " & tailProbability
    If v.calcType = "prediction" Then
        BuildMeanIntervalRCode = common & vbLf & vbLf & _
            upperName & " <- mean(values) - qt(tail_probability, df = n - 1) * (sd(values) * sqrt(1 + 1/n))" & vbLf & _
            lowerName & " <- mean(values) + qt(tail_probability, df = n - 1) * (sd(values) * sqrt(1 + 1/n))" & vbLf & _
            "print(cbind(" & upperName & ", " & lowerName & "), digits = 4)"
        Exit Function
    End If
    If v.stdev = "known" Then
        critical = "qnorm(tail_probability)"
        spread = JsNumberToString(v.sd)
    Else
        critical = "qt(tail_probability, df = n - 1)"
        spread = "sd(values)"
    End If
    BuildMeanIntervalRCode = common & vbLf & vbLf & _
        upperName & " <- mean(values) - " & critical & " * " & spread & " / sqrt(n)" & vbLf & _
        lowerName & " <- mean(values) + " & critical & " * " & spread & " / sqrt(n)" & vbLf & _
        "print(cbind(" & upperName & ", " & lowerName & "), digits = 4)"
End Function

Private Function BuildProportionIntervalRCode(ByRef v As TInterval) As String
    Dim dataRef As String, upperName As String, lowerName As String, tailProbability As String
    Dim conditionData As String, validRows As String, sampleCalculation As String, t As Double, label As String
    label = RColumnName(v.label)
    dataRef = RColumnName(v.dataset) & "$" & label
    upperName = "uCI" & SafeLabel(v) & SafeLevel(v)
    lowerName = "lCI" & SafeLabel(v) & SafeLevel(v)
    t = JsDiv(v.confidence, 100)
    t = JsSub(1, t)
    t = JsDiv(t, 2)
    tailProbability = RTail(t)
    If v.excludeMissing Then
        conditionData = "valid_rows"
        validRows = "valid_rows <- " & RColumnName(v.dataset) & " %>%" & vbLf & "  dplyr::filter(!is.na(" & label & "))" & vbLf & vbLf
    Else
        conditionData = RColumnName(v.dataset)
    End If
    If v.proportionMode = "condition" Then
        sampleCalculation = validRows & "matching_rows <- " & conditionData & " %>%" & IIf(v.excludeMissing, vbLf & "  ", " ") & _
            "dplyr::filter(" & label & " " & v.conditionOperator & " " & JsNumberToString(v.threshold) & ")" & vbLf & _
            IIf(v.excludeMissing, vbLf, "") & "n <- nrow(" & conditionData & ")" & vbLf & "sample_proportion <- nrow(matching_rows) / n"
    Else
        sampleCalculation = "values <- na.omit(" & dataRef & ")" & vbLf & "success_value <- """ & RString(v.success) & """" & vbLf & _
            "successes <- sum(as.character(values) == success_value)" & vbLf & "n <- length(values)" & vbLf & "sample_proportion <- successes / n"
    End If
    BuildProportionIntervalRCode = sampleCalculation & vbLf & "tail_probability <- " & tailProbability & vbLf & _
        "margin_error <- qnorm(1 - tail_probability) * sqrt(sample_proportion * (1 - sample_proportion) / n)" & vbLf & vbLf & _
        upperName & " <- sample_proportion + margin_error" & vbLf & _
        lowerName & " <- sample_proportion - margin_error" & vbLf & _
        "print(cbind(" & lowerName & ", " & upperName & "), digits = 4)"
End Function

Private Function ProportionDescription(ByRef v As TInterval) As String
    Dim comparison As String
    If v.proportionMode <> "condition" Then
        ProportionDescription = v.population & " where " & v.label & " equals " & v.success
        Exit Function
    End If
    Select Case v.conditionOperator
        Case ">": comparison = "more than"
        Case ">=": comparison = "at least"
        Case "<": comparison = "less than"
        Case "<=": comparison = "at most"
        Case "==": comparison = "equal to"
        Case "!=": comparison = "not equal to"
        Case Else: comparison = v.conditionOperator
    End Select
    ProportionDescription = v.population & " with " & v.label & " being " & comparison & " " & FormatFinalNumber(v.threshold)
End Function

Private Function ConfidenceLabel(ByVal confidence As Double) As String
    If JsIsInteger(confidence) Then ConfidenceLabel = JsNumberToString(confidence) Else ConfidenceLabel = FormatFinalNumber(confidence)
End Function

Private Function ConfidenceValid(ByVal confidence As Double) As Boolean
    ConfidenceValid = JsGt(confidence, 0) And JsLt(confidence, 100)
End Function

Private Function ProportionCiProse(ByVal label As String, ByVal headline As String, ByRef v As TInterval) As String
    ProportionCiProse = "The " & label & "% CI for the proportion of " & ProportionDescription(v) & " is " & headline & _
        ". This interval contains the true population proportion with a " & label & "% level of confidence. More generally, with repeated random sampling from the same population, " & _
        label & "% of such constructed confidence intervals would contain the true population proportion of " & ProportionDescription(v) & "."
End Function

Private Function PredictionProse(ByVal label As String, ByVal headline As String, ByRef v As TInterval) As String
    PredictionProse = "For the " & label & "% Prediction Interval: " & headline & ", given the observed " & v.label & ", the " & v.label & _
        " for a new observation will lie within this interval with a " & label & "% level of confidence. More generally, with repeated random sampling from the same population (infinitely), " & _
        label & "% of such constructed predictive intervals would contain the new " & v.label & "."
End Function

Private Function MeanCiProse(ByVal label As String, ByVal headline As String, ByRef v As TInterval) As String
    MeanCiProse = "The " & label & "% CI for the population mean of " & v.label & " is " & headline & ". We are " & label & _
        JsU("% \u201Cconfident\u201D that this interval contains the true population mean of ") & v.label & _
        ". More generally, with repeated random sampling from the same population, " & label & "% of such constructed confidence intervals would contain the true population mean."
End Function

' parsePrintedInterval(text, minimumValues = 2): a clsMap { values: clsList of Double,
' lower, upper (Null with fewer than two values) }, or Nothing (the website's null).
Public Function ParsePrintedInterval(ByVal text As String, Optional ByVal minimumValues As Long = 2) As clsMap
    Dim lines As clsList, i As Long, line As String, hasRowPrefix As Boolean, cleaned As String
    Dim matches As Collection, m As Variant, values As clsList, x As Double, result As clsMap
    Set lines = Rx("\r?\n").Split(text)
    For i = 0 To lines.Count - 1
        line = CStr(lines(i))
        hasRowPrefix = Rx("^\s*\[\s*\d+\s*,?\s*\]\s*").Test(line)
        cleaned = Rx("^\s*\[\s*\d+\s*,?\s*\]\s*").Replace(line, "")
        If hasRowPrefix Or Rx("^\s*[-+]?(?:\d|\.)").Test(cleaned) Then
            Set values = NewList()
            Set matches = Rx("[-+]?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?", "gi").MatchAll(cleaned)
            For Each m In matches
                x = JsNumber(m.Value)
                If JsIsFinite(x) Then values.Push x
            Next m
            If values.Count >= minimumValues Then
                Set result = NewMap()
                Set result("values") = values
                If values.Count >= 2 Then
                    If CDbl(values(0)) < CDbl(values(1)) Then
                        result("lower") = CDbl(values(0))
                        result("upper") = CDbl(values(1))
                    Else
                        result("lower") = CDbl(values(1))
                        result("upper") = CDbl(values(0))
                    End If
                Else
                    result("lower") = Null
                    result("upper") = Null
                End If
                Set ParsePrintedInterval = result
                Exit Function
            End If
        End If
    Next i
    Set ParsePrintedInterval = Nothing
End Function

Public Function IntervalPastedProse() As String
    IntervalPastedProse = mIntervalPastedProse
End Function

Public Sub RenderIntervalPastedResult()
    Dim box As clsNode, text As String, v As TInterval, confidenceOk As Boolean, needsTwoLimits As Boolean
    Dim parsed As clsMap, parsedEnough As Boolean, proportionContextReady As Boolean, label As String
    Dim required As Double, marginText As String, headline As String, minimum As Long, values As clsList
    Set box = Q("#intervalResultProse")
    If box Is Nothing Then Exit Sub
    text = JsTrim(Q("#intervalResultInput").Value)
    v = IntervalValues()
    confidenceOk = ConfidenceValid(v.confidence)
    needsTwoLimits = (v.calcType <> "sample-size")
    If needsTwoLimits Then minimum = 2 Else minimum = 1
    Set parsed = ParsePrintedInterval(text, minimum)
    parsedEnough = Not parsed Is Nothing
    If v.parameter <> "proportion" Or v.calcType = "sample-size" Then
        proportionContextReady = True
    ElseIf v.proportionMode = "condition" Then
        proportionContextReady = v.thresholdProvided And JsIsFinite(v.threshold)
    Else
        proportionContextReady = Len(v.success) > 0
    End If
    mIntervalPastedProse = ""
    If Len(text) = 0 Then
        box.RemoveClass "is-error"
        If needsTwoLimits Then box.TextContent = "Paste the two printed interval limits to generate the prose." Else box.TextContent = "Paste the calculated sample size to generate the prose."
    ElseIf Not parsedEnough Or Not confidenceOk Or Not v.labelProvided Or Not proportionContextReady Then
        box.AddClass "is-error"
        If Not parsedEnough Then
            If needsTwoLimits Then box.TextContent = "I could not find two numeric interval limits in the pasted output." Else box.TextContent = "I could not find the calculated sample size in the pasted output."
        ElseIf Not proportionContextReady Then
            box.TextContent = "Complete the success definition before interpreting this proportion interval."
        Else
            box.TextContent = "Enter the column / parameter name and a confidence level between 0 and 100."
        End If
    Else
        label = ConfidenceLabel(v.confidence)
        If v.calcType = "sample-size" Then
            Set values = parsed("values")
            required = JsCeil(CDbl(values(0)))
            If JsGt(v.margin, 0) Then marginText = " and a margin of error no greater than " & FormatFinalNumber(v.margin)
            mIntervalPastedProse = "A minimum sample size of " & JsNumberToString(required) & " is required to estimate the population " & v.parameter & " of " & v.label & " with a " & label & "% confidence level" & marginText & "."
        Else
            headline = "[" & FormatFinalNumber(CDbl(parsed("lower"))) & ", " & FormatFinalNumber(CDbl(parsed("upper"))) & "]"
            If v.calcType = "prediction" Then
                mIntervalPastedProse = PredictionProse(label, headline, v)
            ElseIf v.parameter = "proportion" Then
                mIntervalPastedProse = ProportionCiProse(label, headline, v)
            Else
                mIntervalPastedProse = MeanCiProse(label, headline, v)
            End If
        End If
        box.RemoveClass "is-error"
        box.TextContent = mIntervalPastedProse
    End If
    Q("#copyIntervalPastedProseBtn").Disabled = (Len(mIntervalPastedProse) = 0)
End Sub

Public Sub SyncIntervalForm()
    Dim calcType As String, parameter As String, sampleSize As Boolean, prediction As Boolean, proportion As Boolean
    Dim conditionProportion As Boolean, stdev As clsNode, stdevNotApplicable As Boolean, previous As Variant
    Dim noteTitle As String, noteText As String, label As String
    calcType = Q("#intervalCalcType").Value
    If calcType = "prediction" And Q("#intervalParameter").Value = "proportion" Then Q("#intervalParameter").Value = "mean"
    parameter = Q("#intervalParameter").Value
    sampleSize = (calcType = "sample-size")
    prediction = (calcType = "prediction")
    proportion = (parameter = "proportion")
    conditionProportion = proportion And Not sampleSize And Q("#intervalProportionMode").Value = "condition"
    Set stdev = Q("#intervalStdev")
    Q("#intervalParameter").QuerySelector("option[value=""proportion""]").Disabled = prediction
    stdev.QuerySelector("option[value=""known""]").Disabled = prediction
    If prediction Then stdev.Value = "unknown"
    stdevNotApplicable = sampleSize Or proportion
    If stdevNotApplicable Then
        If stdev.Value = "na" Then
            previous = stdev.DataGet("previous")
            If IsNull(previous) Then previous = ""
            If Len(previous) = 0 Then previous = "unknown"
            stdev.DataSet "previous", CStr(previous)
        Else
            stdev.DataSet "previous", stdev.Value
        End If
        stdev.Value = "na"
    ElseIf stdev.Value = "na" Then
        previous = stdev.DataGet("previous")
        If IsNull(previous) Then previous = ""
        If Len(previous) = 0 Then previous = "unknown"
        stdev.Value = CStr(previous)
    End If
    stdev.Disabled = stdevNotApplicable
    Q("#stdevChoiceWrap").ToggleClass "inapplicable", stdevNotApplicable
    Q("#estimateField").ToggleClass "hidden", sampleSize Or proportion
    Q("#sampleSizeField").ToggleClass "hidden", sampleSize
    Q("#proportionField").ToggleClass "hidden", Not proportion
    Q("#proportionDefinitionField").ToggleClass "hidden", Not proportion Or sampleSize
    Q("#populationLabelField").ToggleClass "hidden", Not proportion Or sampleSize
    Q("#successValueField").ToggleClass "hidden", Not proportion Or sampleSize Or conditionProportion
    Q("#conditionOperatorField").ToggleClass "hidden", Not conditionProportion
    Q("#conditionThresholdField").ToggleClass "hidden", Not conditionProportion
    Q("#intervalExcludeMissingField").ToggleClass "hidden", calcType <> "confidence" Or Not conditionProportion
    Q("#marginField").ToggleClass "hidden", Not sampleSize
    Q("#stdevField").ToggleClass "hidden", proportion
    Q("#sampleSizeField").FirstChild.TextContent = IIf(proportion, "Sample size", "Sample size (optional)")
    Q("#proportionField").FirstChild.TextContent = IIf(sampleSize, "Expected proportion", "Sample proportion (optional)")
    If sampleSize Then
        label = "Planning standard deviation"
    ElseIf stdev.Value = "known" Then
        label = "Known population standard deviation"
    Else
        label = "Sample standard deviation (optional)"
    End If
    Q("#stdevField").FirstChild.TextContent = label
    If sampleSize Then
        noteTitle = "Sample size for a " & parameter
        If proportion Then
            noteText = JsU("Uses n \u2265 z\u00B2p(1\u2212p)/E\u00B2 and always rounds up.")
        Else
            noteText = JsU("Uses n \u2265 z\u00B2\u03C3\u00B2/E\u00B2 and always rounds up.")
        End If
    Else
        noteTitle = IIf(prediction, "Prediction interval", "Confidence interval") & " for a " & parameter
        If proportion Then
            If conditionProportion Then
                noteText = JsU("Filters rows satisfying the condition, then uses p\u0302 \u00B1 z\u221A(p\u0302(1\u2212p\u0302)/n).")
            Else
                noteText = JsU("Counts the selected success value, then uses p\u0302 \u00B1 z\u221A(p\u0302(1\u2212p\u0302)/n).")
            End If
        ElseIf stdev.Value = "known" Then
            noteText = "Uses a z critical value."
        Else
            noteText = JsU("Uses a t critical value with n \u2212 1 degrees of freedom.")
        End If
    End If
    Q("#intervalNote").InnerHtml = "<strong>" & EscapeHtml(noteTitle) & "</strong><span>" & EscapeHtml(noteText) & "</span>"
    If sampleSize Then
        Q("#intervalResultInput").Placeholder = "[1] 246"
    Else
        Q("#intervalResultInput").Placeholder = "     upper_limit lower_limit" & vbLf & "[1,]       20.33       19.32"
    End If
    RenderIntervalResult
End Sub

Public Sub RenderIntervalResult()
    Dim v As TInterval, confidenceOk As Boolean, alphaTail As Double, lowerTail As Double, rLowerTail As String, z As Double
    Dim prose As String, code As String, formula As String, headline As String, valid As Boolean, label As String
    Dim meanCodeReady As Boolean, proportionDefinitionReady As Boolean, proportionCodeReady As Boolean, sampleSizeCodeReady As Boolean
    Dim required As Double, t As Double, u As Double, margin As Double, lower As Double, upper As Double, critical As Double
    Dim dataRef As String, codeReady As Boolean, proseCopyButton As clsNode, picker As clsNode, html As String
    Dim criticalCode As String, sdCode As String, sdName As String, quantile As String
    v = IntervalValues()
    confidenceOk = ConfidenceValid(v.confidence)
    If confidenceOk Then
        t = v.confidence / 100
        t = 1 - t
        lowerTail = t / 2
        alphaTail = 1 - lowerTail
        rLowerTail = RTail(lowerTail)
    Else
        alphaTail = JsNaN()
        lowerTail = JsNaN()
    End If
    z = InverseNormal(alphaTail)
    valid = confidenceOk And v.labelProvided
    label = ConfidenceLabel(v.confidence)
    meanCodeReady = confidenceOk And v.calcType <> "sample-size" And v.parameter = "mean" And v.datasetProvided And v.labelProvided And _
        (v.stdev <> "known" Or (v.sdProvided And JsGt(v.sd, 0)))
    If v.proportionMode = "condition" Then
        proportionDefinitionReady = v.thresholdProvided And JsIsFinite(v.threshold)
    Else
        proportionDefinitionReady = Len(v.success) > 0
    End If
    proportionCodeReady = confidenceOk And v.calcType = "confidence" And v.parameter = "proportion" And v.datasetProvided And v.labelProvided And proportionDefinitionReady
    sampleSizeCodeReady = confidenceOk And v.calcType = "sample-size" And v.labelProvided And v.marginProvided And JsGt(v.margin, 0)
    If v.parameter = "mean" Then
        sampleSizeCodeReady = sampleSizeCodeReady And v.sdProvided And JsGt(v.sd, 0)
    Else
        sampleSizeCodeReady = sampleSizeCodeReady And v.proportionProvided And JsGe(v.proportion, 0) And JsLe(v.proportion, 1)
    End If

    If v.calcType = "sample-size" Then
        valid = valid And v.labelProvided And v.marginProvided And JsGt(v.margin, 0)
        If v.parameter = "mean" Then
            valid = valid And v.sdProvided And JsGt(v.sd, 0)
            If valid Then
                t = JsMul(z, z)
                t = JsMul(t, v.sd)
                t = JsMul(t, v.sd)
                u = JsMul(v.margin, v.margin)
                required = JsCeil(JsDiv(t, u))
            End If
            formula = JsU("n \u2265 z\u00B2\u03C3\u00B2 / E\u00B2")
            code = "ceiling((qnorm(" & JsNumberToString(alphaTail) & ")^2 * " & JsNumberToString(v.sd) & "^2) / " & JsNumberToString(v.margin) & "^2)"
        Else
            valid = valid And v.proportionProvided And JsGe(v.proportion, 0) And JsLe(v.proportion, 1)
            If valid Then
                t = JsMul(z, z)
                t = JsMul(t, v.proportion)
                t = JsMul(t, JsSub(1, v.proportion))
                u = JsMul(v.margin, v.margin)
                required = JsCeil(JsDiv(t, u))
            End If
            formula = JsU("n \u2265 z\u00B2p(1\u2212p) / E\u00B2")
            code = "ceiling((qnorm(" & JsNumberToString(alphaTail) & ")^2 * " & JsNumberToString(v.proportion) & " * (1-" & JsNumberToString(v.proportion) & ")) / " & JsNumberToString(v.margin) & "^2)"
        End If
        If valid Then
            headline = "Minimum n = " & JsNumberToString(required)
            prose = "A minimum sample size of " & JsNumberToString(required) & " is required to estimate the population " & v.parameter & " of " & v.label & " with a " & label & "% confidence level and a margin of error no greater than " & FormatFinalNumber(v.margin) & "."
        End If
    ElseIf v.parameter = "proportion" Then
        valid = valid And proportionDefinitionReady And v.nProvided And JsGt(v.n, 0) And JsIsInteger(v.n) And v.proportionProvided And JsGe(v.proportion, 0) And JsLe(v.proportion, 1)
        If valid Then
            t = JsSub(1, v.proportion)
            t = JsMul(v.proportion, t)
            t = JsDiv(t, v.n)
            t = JsSqrt(t)
            margin = JsMul(z, t)
            lower = JsMax(0, JsSub(v.proportion, margin))
            upper = JsMin(1, JsAdd(v.proportion, margin))
            headline = "[" & FormatFinalNumber(lower) & ", " & FormatFinalNumber(upper) & "]"
            formula = JsU("p\u0302 \u00B1 z\u221A(p\u0302(1\u2212p\u0302)/n)")
            prose = ProportionCiProse(label, headline, v)
            code = "sample_proportion <- " & JsNumberToString(v.proportion) & vbLf & "n <- " & JsNumberToString(v.n) & vbLf & _
                "margin_error <- qnorm(" & JsNumberToString(alphaTail) & ") * sqrt((sample_proportion * (1 - sample_proportion)) / n)" & vbLf & _
                "c(sample_proportion - margin_error, sample_proportion + margin_error)"
        End If
    Else
        valid = valid And v.nProvided And JsGt(v.n, 1) And JsIsInteger(v.n) And v.sdProvided And JsGt(v.sd, 0) And v.estimateProvided And JsIsFinite(v.estimate)
        If valid Then
            If v.calcType = "prediction" Then
                critical = InverseStudentT(alphaTail, v.n - 1)
            ElseIf v.stdev = "known" Then
                critical = z
            Else
                critical = InverseStudentT(alphaTail, v.n - 1)
            End If
            If v.calcType = "prediction" Then
                t = JsDiv(1, v.n)
                t = JsAdd(1, t)
                t = JsSqrt(t)
                t = JsMul(v.sd, t)
            Else
                t = JsDiv(v.sd, JsSqrt(v.n))
            End If
            margin = JsMul(critical, t)
            lower = JsSub(v.estimate, margin)
            upper = JsAdd(v.estimate, margin)
            headline = "[" & FormatFinalNumber(lower) & ", " & FormatFinalNumber(upper) & "]"
            If v.calcType = "prediction" Then
                quantile = "t"
            ElseIf v.stdev = "known" Then
                quantile = "z"
            Else
                quantile = "t"
            End If
            If v.calcType = "prediction" Then
                formula = JsU("x\u0304 \u00B1 t\u00B7s\u221A(1+1/n), df = n\u22121")
            Else
                formula = JsU("x\u0304 \u00B1 ") & quantile & JsU("\u00B7s/\u221An")
            End If
            dataRef = RColumnName(v.dataset) & "$" & RColumnName(v.label)
            If v.calcType = "prediction" Then
                prose = PredictionProse(label, headline, v)
                code = "values <- na.omit(" & dataRef & ")" & vbLf & "n <- length(values)" & vbLf & "sample_mean <- mean(values)" & vbLf & _
                    "sample_stdev <- sd(values)" & vbLf & "tail_probability <- " & rLowerTail & vbLf & vbLf & _
                    "upper_pi <- sample_mean - qt(tail_probability, df = n - 1) * (sample_stdev * sqrt(1 + 1/n))" & vbLf & _
                    "lower_pi <- sample_mean + qt(tail_probability, df = n - 1) * (sample_stdev * sqrt(1 + 1/n))" & vbLf & _
                    "print(cbind(upper_pi, lower_pi), digits = 4)"
            Else
                prose = MeanCiProse(label, headline, v)
                If v.stdev = "known" Then
                    criticalCode = "qnorm(tail_probability)"
                    sdCode = "population_stdev <- " & JsNumberToString(v.sd)
                    sdName = "population_stdev"
                Else
                    criticalCode = "qt(tail_probability, df = n - 1)"
                    sdCode = "sample_stdev <- sd(values)"
                    sdName = "sample_stdev"
                End If
                code = "values <- na.omit(" & dataRef & ")" & vbLf & "n <- length(values)" & vbLf & "sample_mean <- mean(values)" & vbLf & _
                    sdCode & vbLf & "tail_probability <- " & rLowerTail & vbLf & vbLf & _
                    "upper_ci <- sample_mean - " & criticalCode & " * " & sdName & " / sqrt(n)" & vbLf & _
                    "lower_ci <- sample_mean + " & criticalCode & " * " & sdName & " / sqrt(n)" & vbLf & _
                    "print(cbind(upper_ci, lower_ci), digits = 4)"
            End If
        End If
    End If

    If meanCodeReady Then code = BuildMeanIntervalRCode(v, rLowerTail)
    If proportionCodeReady Then code = BuildProportionIntervalRCode(v)
    If v.calcType = "sample-size" And Not sampleSizeCodeReady Then code = ""
    codeReady = Len(code) > 0

    Set proseCopyButton = Q("#copyIntervalProseBtn")
    proseCopyButton.Disabled = Not valid
    proseCopyButton.ToggleClass "hidden", Not valid
    Set picker = PickerForCopyButton(proseCopyButton)
    If Not picker Is Nothing Then picker.ToggleClass "hidden", Not valid
    Q("#copyIntervalCodeBtn").Disabled = Not codeReady
    If valid Then
        html = "<div class=""prose-card""><div class=""result-interval"">" & EscapeHtml(headline) & "</div><p id=""intervalGeneratedProse"">" & EscapeHtml(prose) & _
            "</p><div class=""formula-line"">" & EscapeHtml(formula) & "</div><pre class=""interval-code intervalGeneratedCode"">" & EscapeHtml(code) & "</pre></div>"
    ElseIf codeReady Then
        html = "<div class=""prose-card""><h3>R code ready</h3><p>"
        If v.calcType = "sample-size" Then
            html = html & "Run this code in RStudio, then paste the calculated sample size into the output box below."
        Else
            html = html & "Run this code in RStudio, then paste the two printed limits into the output box below. Any missing sample statistics will be calculated from the data."
        End If
        html = html & "</p><pre class=""interval-code intervalGeneratedCode"">" & EscapeHtml(code) & "</pre></div>"
    Else
        html = "<div class=""empty-state compact""><div class=""empty-state-inner""><strong>Complete the required values</strong><p>"
        If v.calcType = "sample-size" Then
            html = html & "Enter the parameter name, confidence level, desired margin of error, and the required planning value to generate the sample-size code."
        Else
            html = html & "Enter the dataset, column / parameter name, and a confidence level to generate R code. Manual sample statistics are only required for an immediate browser calculation."
        End If
        html = html & "</p></div></div>"
    End If
    Q("#intervalOutput").InnerHtml = html
    RenderIntervalPastedResult
End Sub

' ---------------------------------------------------------------- event handlers (wired by opInference.InitOpInference)

Public Sub OnIntervalSync(ByVal e As clsEvent)
    On Error GoTo Failed
    SyncIntervalForm
    Exit Sub
Failed:
    ReportError "opInference2.OnIntervalSync"
End Sub

Public Sub OnIntervalRender(ByVal e As clsEvent)
    On Error GoTo Failed
    RenderIntervalResult
    Exit Sub
Failed:
    ReportError "opInference2.OnIntervalRender"
End Sub

Public Sub OnIntervalPastedInput(ByVal e As clsEvent)
    On Error GoTo Failed
    RenderIntervalPastedResult
    Exit Sub
Failed:
    ReportError "opInference2.OnIntervalPastedInput"
End Sub

Public Sub OnIntervalStdevChange(ByVal e As clsEvent)
    On Error GoTo Failed
    If e.CurrentTarget.Value = "known" Then
        e.CurrentTarget.Value = "unknown"
        CallOp KnownSdAlertProcedure
        Exit Sub
    End If
    SyncIntervalForm
    Exit Sub
Failed:
    ReportError "opInference2.OnIntervalStdevChange"
End Sub

Public Sub OnLoadIntervalSample(ByVal e As clsEvent)
    Dim calcType As String, proportion As Boolean, conditionProportion As Boolean
    On Error GoTo Failed
    Q("#intervalExcludeMissing").Checked = False
    calcType = Q("#intervalCalcType").Value
    proportion = (Q("#intervalParameter").Value = "proportion")
    conditionProportion = proportion And calcType <> "sample-size" And Q("#intervalProportionMode").Value = "condition"
    Q("#intervalDataset").Value = IIf(conditionProportion, "dvd", "dataset")
    If conditionProportion Then
        Q("#intervalLabel").Value = "Amount"
    ElseIf calcType = "prediction" Then
        Q("#intervalLabel").Value = "nail length"
    Else
        Q("#intervalLabel").Value = "column1"
    End If
    Q("#confidenceLevel").Value = IIf(conditionProportion, "90", "95")
    Q("#intervalEstimate").Value = "50"
    Q("#intervalSd").Value = "4"
    If conditionProportion Then
        Q("#intervalN").Value = "213"
    ElseIf proportion Then
        Q("#intervalN").Value = "100"
    Else
        Q("#intervalN").Value = "64"
    End If
    Q("#intervalProportion").Value = IIf(conditionProportion, "0.251", "0.64")
    Q("#intervalPopulation").Value = "transactions"
    Q("#intervalSuccess").Value = "yes"
    Q("#intervalConditionOperator").Value = ">"
    Q("#intervalConditionThreshold").Value = IIf(conditionProportion, "22", "")
    Q("#intervalMargin").Value = IIf(proportion, "0.05", "0.5")
    RenderIntervalResult
    ShowToast "Sample loaded"
    Exit Sub
Failed:
    ReportError "opInference2.OnLoadIntervalSample"
End Sub

Public Sub OnLoadIntervalResultSample(ByVal e As clsEvent)
    Dim calcType As String, parameter As String, conditionProportion As Boolean
    On Error GoTo Failed
    Q("#intervalExcludeMissing").Checked = False
    calcType = Q("#intervalCalcType").Value
    parameter = Q("#intervalParameter").Value
    conditionProportion = parameter = "proportion" And calcType <> "sample-size" And Q("#intervalProportionMode").Value = "condition"
    Q("#intervalDataset").Value = IIf(conditionProportion, "dvd", "dataset")
    Q("#intervalLabel").Value = IIf(conditionProportion, "Amount", "column1")
    Q("#confidenceLevel").Value = IIf(conditionProportion, "90", "95")
    If calcType = "sample-size" Then
        Q("#intervalMargin").Value = IIf(parameter = "proportion", "0.05", "0.5")
        Q("#intervalSd").Value = "4"
        Q("#intervalProportion").Value = "0.64"
        Q("#intervalResultInput").Value = IIf(parameter = "proportion", "[1] 355", "[1] 246")
    ElseIf parameter = "proportion" Then
        Q("#intervalPopulation").Value = "transactions"
        Q("#intervalSuccess").Value = "yes"
        Q("#intervalConditionOperator").Value = ">"
        Q("#intervalConditionThreshold").Value = IIf(conditionProportion, "22", "")
        If conditionProportion Then
            Q("#intervalResultInput").Value = "     lCIAmount90 uCIAmount90" & vbLf & "[1,]          0.202       0.300"
        Else
            Q("#intervalResultInput").Value = "     lCIcolumn195 uCIcolumn195" & vbLf & "[1,]          0.546        0.734"
        End If
    ElseIf calcType = "prediction" Then
        Q("#intervalResultInput").Value = "     uPIcolumn195 lPIcolumn195" & vbLf & "[1,]            58.06            41.94"
    ElseIf Q("#intervalStdev").Value = "known" Then
        Q("#intervalSd").Value = "4"
        Q("#intervalResultInput").Value = "     uCIcolumn195 lCIcolumn195" & vbLf & "[1,]            50.98            49.02"
    Else
        Q("#intervalResultInput").Value = "     uCIcolumn195 lCIcolumn195" & vbLf & "[1,]            51.00            49.00"
    End If
    SyncIntervalForm
    RenderIntervalPastedResult
    ShowToast "Sample output loaded"
    Exit Sub
Failed:
    ReportError "opInference2.OnLoadIntervalResultSample"
End Sub

Public Sub OnCopyIntervalPastedProse(ByVal e As clsEvent)
    On Error GoTo Failed
    If Len(mIntervalPastedProse) = 0 Then Exit Sub
    If WriteAnswerText(mIntervalPastedProse) Then ShowToast "Prose copied" Else ShowToast "Clipboard access is unavailable"
    Exit Sub
Failed:
    ReportError "opInference2.OnCopyIntervalPastedProse"
End Sub

Public Sub OnCopyIntervalProse(ByVal e As clsEvent)
    Dim node As clsNode, prose As String
    On Error GoTo Failed
    Set node = Q("#intervalGeneratedProse")
    If node Is Nothing Then Exit Sub
    prose = node.TextContent
    If Len(prose) = 0 Then Exit Sub
    If WriteAnswerText(prose) Then ShowToast "Prose copied" Else ShowToast "Clipboard access is unavailable"
    Exit Sub
Failed:
    ReportError "opInference2.OnCopyIntervalProse"
End Sub

Public Sub OnCopyIntervalCode(ByVal e As clsEvent)
    Dim node As clsNode, code As String
    On Error GoTo Failed
    Set node = Q(".intervalGeneratedCode")
    If node Is Nothing Then Exit Sub
    code = node.TextContent
    If Len(code) = 0 Then Exit Sub
    If WriteAnswerText(code) Then ShowToast "R code copied" Else ShowToast "Clipboard access is unavailable"
    Exit Sub
Failed:
    ReportError "opInference2.OnCopyIntervalCode"
End Sub
