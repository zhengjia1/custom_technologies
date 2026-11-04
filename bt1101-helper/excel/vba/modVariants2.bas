Attribute VB_Name = "modVariants2"
' Answer options, part 2: the R-code variation engine (website app-src.js lines 211-656):
' rTokens, rStatements, protectedRNames, assignedRObjects, rObjectNameType, freshRName,
' renameRObjects, storeRReports, varyRStructure, randomPlotStylePlans, changeRPlotStyle,
' rArgumentCalls, shuffleRNamedArguments and codeAnswerVariants. Randomness only through
' modJs.JsRandom, consumed in exactly the website's order (Math.floor(Math.random() * n)).
' Token positions are 0-based UTF-16 offsets like the website's.
Option Explicit
Option Compare Binary

Private Type RToks
    Count As Long
    Txt() As String
    St() As Long          ' start offset (0-based)
    En() As Long          ' end offset (exclusive)
    Par() As Long         ' enclosing open token, -1 at top level
    OpenOf() As Long      ' closer -> its opener (-1 = undefined)
    CloseOf() As Long     ' opener -> its closer (-1 = undefined)
End Type

Private Type REdits
    Count As Long
    S() As Long
    E() As Long
    T() As String
End Type

Private Type RCall
    Valid As Boolean
    ArgCount As Long
    ArgLeft() As Long
    ArgRight() As Long
    ArgName() As String   ' "" = not a named argument
    ArgStart() As Long
    ArgEnd() As Long
    SlotCount As Long
    Slots() As Long
    Key As String
End Type

' tokens[i]?.text for a missing token (never equal to any token text).
Private Const UNDEF As String = vbNullChar & "undefined"

Private mReady As Boolean
Private mNamePools As clsMap        ' type -> Variant array of names
Private mPolicyPackages As clsMap   ' call -> " pkg pkg "
Private mPolicyNames As clsMap      ' call -> " name name "
Private mPureCalls As String
Private mEffectfulCalls As String

' Tokenization memo: the variant pipeline tokenizes the same texts many times (the
' canonical code once per option, unchanged intermediate results again); a few recent
' results are kept by exact text. Tokenize is pure, so a copy of a cached result is
' identical to a fresh one.
Private Const TOKEN_MEMO As Long = 8
Private mTokKey(0 To TOKEN_MEMO - 1) As String
Private mTokVal(0 To TOKEN_MEMO - 1) As RToks
Private mTokUsed(0 To TOKEN_MEMO - 1) As Boolean
Private mTokNext As Long

' assignedRObjects(canonical) for the duration of one VarCodeAnswerVariants call: every
' option renames the same canonical text and the DOM does not change meanwhile, so the
' (random-free) result is the same each time.
Private mAssignedFor As String
Private mAssignedActive As Boolean
Private mAssignedProtected As clsMap
Private mAssignedResult As clsList

' freshRName candidates per name type (pool names, dotted and camel forms; built once)
' and rObjectNameType results per name (pure functions of their inputs).
Private mCandidates As clsMap
Private mObjectTypes As clsMap

' Plot-style plans 1..5 (index 0 = null, the canonical answer).
Private mPlanPrimary(0 To 5) As String
Private mPlanAccent(0 To 5) As String
Private mPlanBrewer(0 To 5) As String
Private mPlanShape(0 To 5) As String

' ================================================================ randomness

' randomIndex(length): Math.floor(Math.random() * length); no draw when length <= 1.
Public Function VarRandomIndex(ByVal length As Long) As Long
    Dim r As Double
    If length <= 1 Then Exit Function
    r = JsRandom() * length
    VarRandomIndex = Int(r)
End Function

' shuffled(values) in place (Fisher-Yates from the end, like the website).
Public Sub VarShuffleStrings(ByRef values() As String, ByVal lo As Long, ByVal hi As Long)
    Dim index As Long, swap As Long, tmp As String
    For index = hi - lo To 1 Step -1
        swap = VarRandomIndex(index + 1)
        tmp = values(lo + index)
        values(lo + index) = values(lo + swap)
        values(lo + swap) = tmp
    Next index
End Sub

Private Sub ShuffleLongs(ByRef values() As Long, ByVal n As Long)
    Dim index As Long, swap As Long, tmp As Long
    For index = n - 1 To 1 Step -1
        swap = VarRandomIndex(index + 1)
        tmp = values(index)
        values(index) = values(swap)
        values(swap) = tmp
    Next index
End Sub

' ================================================================ data

Private Sub EnsureData()
    If mReady Then Exit Sub
    Set mCandidates = NewMap()
    Set mObjectTypes = NewMap()
    Set mNamePools = NewMap()
    Set mPolicyPackages = NewMap()
    Set mPolicyNames = NewMap()
    InitNamePools
    InitShufflePolicies
    mReady = True
End Sub

Private Sub AddNamePool(ByVal kind As String, ByVal names As String)
    mNamePools(kind) = Split(names, " ")
End Sub

Private Sub AddShufflePolicy(ByVal callName As String, ByVal packages As String, ByVal names As String)
    mPolicyPackages(callName) = " " & packages & " "
    mPolicyNames(callName) = " " & names & " "
End Sub

' BEGIN GENERATED DATA (node test/variants/gen-data.js names)
Private Sub InitNamePools()
    AddNamePool "model", "analysis_model fitted_model regression_model final_model student_model model_fit chosen_model estimated_model test_model full_model"
    AddNamePool "result", "analysis_result test_result result_output computed_result final_result answer_result test_output analysis_output results output"
    AddNamePool "table", "summary_table results_table output_table frequency_table analysis_table data_summary result_table final_table summary_results table_output"
    AddNamePool "data", "analysis_data working_data selected_data clean_data filtered_data plot_data student_data data_subset prepared_data input_data"
    AddNamePool "values", "data_values plot_values selected_values numeric_values working_values value_list sample_values observed_values input_values analysis_values"
    AddNamePool "vector", "value_vector result_vector coefficient_vector input_vector output_vector number_vector data_vector answer_vector working_vector vector_values"
    AddNamePool "matrix", "data_matrix result_matrix coefficient_matrix constraint_matrix analysis_matrix working_matrix input_matrix output_matrix matrix_values model_matrix"
    AddNamePool "range", "axis_range value_range plot_range data_range display_range number_range result_range graph_range limits_range working_range"
    AddNamePool "ticks", "axis_ticks tick_values plot_ticks number_ticks display_ticks graph_ticks tick_marks axis_marks scale_ticks label_ticks"
    AddNamePool "colours", "plot_colours bar_colours chart_colours group_colours display_colours colour_values graph_colours fill_colours series_colours chosen_colours"
    AddNamePool "loop", "i j k row_number item_number fold_number group_number step_number position index"
    AddNamePool "generic", "analysis_step working_result calculated_value student_result output_value current_result final_value computed_value answer_value result_item work_value analysis_item"
End Sub

Private Sub InitShufflePolicies()
    AddShufflePolicy "plot", "graphics", "x y height beside horiz horizontal range breaks plot freq labels names.arg main sub xlab ylab xlim ylim las cex cex.names cex.axis cex.lab cex.main col border fill pch lty lwd type pos offset at legend title radius notch outline add axes ann log"
    AddShufflePolicy "barplot", "graphics", "x y height beside horiz horizontal range breaks plot freq labels names.arg main sub xlab ylab xlim ylim las cex cex.names cex.axis cex.lab cex.main col border fill pch lty lwd type pos offset at legend title radius notch outline add axes ann log"
    AddShufflePolicy "boxplot", "graphics", "x y height beside horiz horizontal range breaks plot freq labels names.arg main sub xlab ylab xlim ylim las cex cex.names cex.axis cex.lab cex.main col border fill pch lty lwd type pos offset at legend title radius notch outline add axes ann log"
    AddShufflePolicy "hist", "graphics", "x y height beside horiz horizontal range breaks plot freq labels names.arg main sub xlab ylab xlim ylim las cex cex.names cex.axis cex.lab cex.main col border fill pch lty lwd type pos offset at legend title radius notch outline add axes ann log"
    AddShufflePolicy "pie", "graphics", "x y height beside horiz horizontal range breaks plot freq labels names.arg main sub xlab ylab xlim ylim las cex cex.names cex.axis cex.lab cex.main col border fill pch lty lwd type pos offset at legend title radius notch outline add axes ann log"
    AddShufflePolicy "mosaicplot", "graphics", "x y height beside horiz horizontal range breaks plot freq labels names.arg main sub xlab ylab xlim ylim las cex cex.names cex.axis cex.lab cex.main col border fill pch lty lwd type pos offset at legend title radius notch outline add axes ann log"
    AddShufflePolicy "biplot", "graphics", "x y height beside horiz horizontal range breaks plot freq labels names.arg main sub xlab ylab xlim ylim las cex cex.names cex.axis cex.lab cex.main col border fill pch lty lwd type pos offset at legend title radius notch outline add axes ann log"
    AddShufflePolicy "points", "graphics", "x y height beside horiz horizontal range breaks plot freq labels names.arg main sub xlab ylab xlim ylim las cex cex.names cex.axis cex.lab cex.main col border fill pch lty lwd type pos offset at legend title radius notch outline add axes ann log"
    AddShufflePolicy "lines", "graphics", "x y height beside horiz horizontal range breaks plot freq labels names.arg main sub xlab ylab xlim ylim las cex cex.names cex.axis cex.lab cex.main col border fill pch lty lwd type pos offset at legend title radius notch outline add axes ann log"
    AddShufflePolicy "text", "graphics", "x y height beside horiz horizontal range breaks plot freq labels names.arg main sub xlab ylab xlim ylim las cex cex.names cex.axis cex.lab cex.main col border fill pch lty lwd type pos offset at legend title radius notch outline add axes ann log"
    AddShufflePolicy "axis", "graphics", "x y height beside horiz horizontal range breaks plot freq labels names.arg main sub xlab ylab xlim ylim las cex cex.names cex.axis cex.lab cex.main col border fill pch lty lwd type pos offset at legend title radius notch outline add axes ann log"
    AddShufflePolicy "legend", "graphics", "x y height beside horiz horizontal range breaks plot freq labels names.arg main sub xlab ylab xlim ylim las cex cex.names cex.axis cex.lab cex.main col border fill pch lty lwd type pos offset at legend title radius notch outline add axes ann log"
    AddShufflePolicy "pairs.panels", "psych", "method hist.col col pch main smooth scale density ellipses lm cor digits cex"
    AddShufflePolicy "fviz_eig", "factoextra", "barfill barcolor linecolor palette data addlabels ncp choice geom main repel"
    AddShufflePolicy "fviz_cluster", "factoextra", "barfill barcolor linecolor palette data addlabels ncp choice geom main repel"
    AddShufflePolicy "lm", "stats", "formula data family subset weights na.action model x y qr contrasts offset singular.ok"
    AddShufflePolicy "glm", "stats", "formula data family subset weights na.action model x y qr contrasts offset singular.ok"
    AddShufflePolicy "aov", "stats", "formula data family subset weights na.action model x y qr contrasts offset singular.ok"
    AddShufflePolicy "t.test", "stats psych", "x y n p mu ratio data alternative conf.level paired var.equal correct method exact continuity adjust ci"
    AddShufflePolicy "prop.test", "stats psych", "x y n p mu ratio data alternative conf.level paired var.equal correct method exact continuity adjust ci"
    AddShufflePolicy "var.test", "stats psych", "x y n p mu ratio data alternative conf.level paired var.equal correct method exact continuity adjust ci"
    AddShufflePolicy "cor.test", "stats psych", "x y n p mu ratio data alternative conf.level paired var.equal correct method exact continuity adjust ci"
    AddShufflePolicy "corr.test", "stats psych", "x y n p mu ratio data alternative conf.level paired var.equal correct method exact continuity adjust ci"
    AddShufflePolicy "bartlett.test", "stats psych", "x y n p mu ratio data alternative conf.level paired var.equal correct method exact continuity adjust ci"
    AddShufflePolicy "predict", "stats", "newdata type interval level parm which ordered conf.level"
    AddShufflePolicy "confint", "stats", "newdata type interval level parm which ordered conf.level"
    AddShufflePolicy "TukeyHSD", "stats", "newdata type interval level parm which ordered conf.level"
    AddShufflePolicy "prcomp", "stats", "center scale. rank. tol centers nstart iter.max algorithm direction trace scope steps k"
    AddShufflePolicy "kmeans", "stats", "center scale. rank. tol centers nstart iter.max algorithm direction trace scope steps k"
    AddShufflePolicy "step", "stats", "center scale. rank. tol centers nstart iter.max algorithm direction trace scope steps k"
    AddShufflePolicy "mean", "base stats", "na.rm trim probs names type"
    AddShufflePolicy "sd", "base stats", "na.rm trim probs names type"
    AddShufflePolicy "var", "base stats", "na.rm trim probs names type"
    AddShufflePolicy "quantile", "base stats", "na.rm trim probs names type"
    AddShufflePolicy "median", "base stats", "na.rm trim probs names type"
    AddShufflePolicy "matrix", "base", "data nrow ncol byrow dimnames"
    AddShufflePolicy "data.frame", "base", "row.names check.rows check.names fix.empty.names stringsAsFactors"
    AddShufflePolicy "table", "base", "exclude useNA dnn deparse.level"
    AddShufflePolicy "cut", "base", "breaks labels include.lowest right dig.lab ordered_result"
    AddShufflePolicy "format", "base", "digits nsmall trim scientific justify width big.mark small.mark decimal.mark"
    AddShufflePolicy "round", "base", "digits nsmall trim scientific justify width big.mark small.mark decimal.mark"
    AddShufflePolicy "signif", "base", "digits nsmall trim scientific justify width big.mark small.mark decimal.mark"
    AddShufflePolicy "seq", "base", "from to by length.out along.with times each"
    AddShufflePolicy "seq.int", "base", "from to by length.out along.with times each"
    AddShufflePolicy "rep", "base", "from to by length.out along.with times each"
    AddShufflePolicy "hcl.colors", "grDevices", "n palette alpha rev fixup"
    AddShufflePolicy "kable", "knitr", "format digits row.names col.names align caption escape padding"
    AddShufflePolicy "lp", "lpSolve", "direction objective.in const.mat const.dir const.rhs transpose.const int.vec presolve compute.sens all.int all.bin scale dense.const num.bin.solns binary.vec use.rw"
    AddShufflePolicy "pivot_longer", "tidyr", "cols names_to values_to names_prefix names_sep names_pattern values_drop_na cols_vary"
    AddShufflePolicy "spread", "tidyr", "key value fill convert drop sep"
    AddShufflePolicy "aes", "ggplot2", "x y fill colour color size shape alpha title subtitle caption tag"
    AddShufflePolicy "labs", "ggplot2", "x y fill colour color size shape alpha title subtitle caption tag"
    AddShufflePolicy "geom_bar", "ggplot2", "mapping data stat position width fill colour color alpha na.rm show.legend inherit.aes orientation"
    AddShufflePolicy "geom_col", "ggplot2", "mapping data stat position width fill colour color alpha na.rm show.legend inherit.aes orientation"
    AddShufflePolicy "scale_y_continuous", "ggplot2", "name breaks labels limits expand trans transform minor_breaks position"
    AddShufflePolicy "scale_x_continuous", "ggplot2", "name breaks labels limits expand trans transform minor_breaks position"
    AddShufflePolicy "scale_fill_brewer", "ggplot2", "aesthetics palette name type direction breaks labels limits drop na.translate"
    AddShufflePolicy "discrete_scale", "ggplot2", "aesthetics palette name type direction breaks labels limits drop na.translate"
    AddShufflePolicy "element_text", "ggplot2", "family face colour color size hjust vjust angle lineheight margin debug inherit.blank"
    AddShufflePolicy "expansion", "ggplot2", "mult add"
    mPureCalls = " c list length lengths nrow ncol nchar colnames rownames names levels as.character as.numeric as.integer as.logical as.matrix as.data.frame factor is.finite is.na min max sum abs diff range pretty floor ceiling sqrt log log10 log2 exp seq_along seq_len unique sort rev head tail which all any pmin pmax paste paste0 rainbow coef residuals fitted starts_with ends_with everything where is.numeric cumsum colSums rowSums "
    mEffectfulCalls = " plot barplot boxplot hist pie mosaicplot biplot points lines text axis legend pairs.panels fviz_eig fviz_cluster kmeans step "
End Sub
' END GENERATED DATA

' ================================================================ tokens

' rTokens(value), memoized (see TOKEN_MEMO).
Private Sub Tokenize(ByVal value As String, ByRef t As RToks)
    Dim i As Long, n As Long
    n = Len(value)
    For i = 0 To TOKEN_MEMO - 1
        If mTokUsed(i) Then
            If Len(mTokKey(i)) = n Then
                If mTokKey(i) = value Then t = mTokVal(i): Exit Sub
            End If
        End If
    Next i
    TokenizeText value, t
    mTokKey(mTokNext) = value
    mTokVal(mTokNext) = t
    mTokUsed(mTokNext) = True
    mTokNext = (mTokNext + 1) Mod TOKEN_MEMO
End Sub

' rTokens(value): every token except whitespace and comments, with bracket nesting.
Private Sub TokenizeText(ByVal value As String, ByRef t As RToks)
    Dim cs() As Long, b() As Byte, n As Long, i As Long, p As Long, q As Long, c As Long
    Dim stack() As Long, sp As Long, cap As Long, idx As Long, tx As String
    n = Len(value)
    t.Count = 0
    cap = 16
    ReDim t.Txt(0 To cap - 1): ReDim t.St(0 To cap - 1): ReDim t.En(0 To cap - 1)
    ReDim t.Par(0 To cap - 1): ReDim t.OpenOf(0 To cap - 1): ReDim t.CloseOf(0 To cap - 1)
    If n = 0 Then Exit Sub
    b = value
    ReDim cs(0 To n)
    For i = 0 To n - 1
        cs(i) = b(2 * i) + 256& * b(2 * i + 1)
    Next i
    cs(n) = -1
    ReDim stack(0 To 15)
    sp = 0
    p = 0
    Do While p < n
        c = cs(p)
        If JsIsWhitespaceCode(c) Then
            q = p + 1
            Do While q < n
                If Not JsIsWhitespaceCode(cs(q)) Then Exit Do
                q = q + 1
            Loop
            p = q
        ElseIf c = 35 Then
            q = p + 1
            Do While q < n
                If cs(q) = 10 Then Exit Do
                q = q + 1
            Loop
            p = q
        Else
            q = -1
            If c = 34 Or c = 39 Or c = 96 Then q = StringEnd(cs, n, p, c)
            If q < 0 Then
                If IsDigitCode(c) Then
                    q = NumberEnd(cs, n, p)
                ElseIf c = 46 Then
                    If IsDigitCode(cs(p + 1)) Then q = NumberEnd(cs, n, p)
                End If
            End If
            If q < 0 Then
                If IsIdentStart(c) Then
                    q = p + 1
                    Do While IsIdentPart(cs(q))
                        q = q + 1
                    Loop
                End If
            End If
            If q < 0 And c = 37 Then q = PercentEnd(cs, n, p)
            If q < 0 Then q = p + OperatorLengthAt(cs, p)
            tx = Mid$(value, p + 1, q - p)
            idx = t.Count
            If idx > UBound(t.Txt) Then
                cap = cap * 2
                ReDim Preserve t.Txt(0 To cap - 1): ReDim Preserve t.St(0 To cap - 1): ReDim Preserve t.En(0 To cap - 1)
                ReDim Preserve t.Par(0 To cap - 1): ReDim Preserve t.OpenOf(0 To cap - 1): ReDim Preserve t.CloseOf(0 To cap - 1)
            End If
            t.Txt(idx) = tx
            t.St(idx) = p
            t.En(idx) = q
            If sp > 0 Then t.Par(idx) = stack(sp - 1) Else t.Par(idx) = -1
            t.OpenOf(idx) = -1
            t.CloseOf(idx) = -1
            t.Count = idx + 1
            If tx = "(" Or tx = "[" Or tx = "{" Then
                If sp > UBound(stack) Then ReDim Preserve stack(0 To sp * 2 + 1)
                stack(sp) = idx
                sp = sp + 1
            ElseIf tx = ")" Or tx = "]" Or tx = "}" Then
                If sp > 0 Then
                    sp = sp - 1
                    t.CloseOf(stack(sp)) = idx
                    t.OpenOf(idx) = stack(sp)
                End If
            End If
            p = q
        End If
    Loop
End Sub

Private Function IsDigitCode(ByVal c As Long) As Boolean
    IsDigitCode = (c >= 48 And c <= 57)
End Function

Private Function IsIdentStart(ByVal c As Long) As Boolean
    IsIdentStart = (c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Or c = 46
End Function

Private Function IsIdentPart(ByVal c As Long) As Boolean
    IsIdentPart = (c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Or c = 46 Or (c >= 48 And c <= 57) Or c = 95
End Function

' "(?:\\[\s\S]|[^"\\])*" and its ' and ` forms: end offset, or -1 when unterminated.
Private Function StringEnd(ByRef cs() As Long, ByVal n As Long, ByVal p As Long, ByVal quote As Long) As Long
    Dim i As Long
    i = p + 1
    Do
        If i >= n Then StringEnd = -1: Exit Function
        If cs(i) = 92 Then
            If i + 1 >= n Then StringEnd = -1: Exit Function
            i = i + 2
        ElseIf cs(i) = quote Then
            StringEnd = i + 1
            Exit Function
        Else
            i = i + 1
        End If
    Loop
End Function

' (?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?L?
Private Function NumberEnd(ByRef cs() As Long, ByVal n As Long, ByVal p As Long) As Long
    Dim q As Long, r As Long
    q = p
    If IsDigitCode(cs(q)) Then
        Do While IsDigitCode(cs(q))
            q = q + 1
        Loop
        If cs(q) = 46 Then
            q = q + 1
            Do While IsDigitCode(cs(q))
                q = q + 1
            Loop
        End If
    Else
        q = p + 1
        Do While IsDigitCode(cs(q))
            q = q + 1
        Loop
    End If
    If cs(q) = 101 Or cs(q) = 69 Then
        r = q + 1
        If cs(r) = 43 Or cs(r) = 45 Then r = r + 1
        If IsDigitCode(cs(r)) Then
            Do While IsDigitCode(cs(r))
                r = r + 1
            Loop
            q = r
        End If
    End If
    If cs(q) = 76 Then q = q + 1
    NumberEnd = q
End Function

' %[^%\n]*%
Private Function PercentEnd(ByRef cs() As Long, ByVal n As Long, ByVal p As Long) As Long
    Dim q As Long
    q = p + 1
    Do While q < n
        If cs(q) = 37 Then PercentEnd = q + 1: Exit Function
        If cs(q) = 10 Then Exit Do
        q = q + 1
    Loop
    PercentEnd = -1
End Function

' OperatorLength on the code array (cs(n) = -1 terminates it).
Private Function OperatorLengthAt(ByRef cs() As Long, ByVal p As Long) As Long
    Dim c0 As Long, c1 As Long, c2 As Long
    c0 = cs(p)
    c1 = cs(p + 1)
    If c1 >= 0 Then c2 = cs(p + 2) Else c2 = -1
    OperatorLengthAt = 1
    Select Case c0
        Case 58 ' :
            If c1 = 58 Then
                If c2 = 58 Then OperatorLengthAt = 3 Else OperatorLengthAt = 2
            End If
        Case 60 ' <
            If c1 = 60 And c2 = 45 Then
                OperatorLengthAt = 3
            ElseIf c1 = 45 Or c1 = 61 Then
                OperatorLengthAt = 2
            End If
        Case 45 ' -
            If c1 = 62 Then
                If c2 = 62 Then OperatorLengthAt = 3 Else OperatorLengthAt = 2
            End If
        Case 124 ' |
            If c1 = 62 Or c1 = 124 Then OperatorLengthAt = 2
        Case 62, 61, 33 ' > = !
            If c1 = 61 Then OperatorLengthAt = 2
        Case 38 ' &
            If c1 = 38 Then OperatorLengthAt = 2
    End Select
End Function

' :::|::|<<-|<-|->>|->|\|>|<=|>=|==|!=|&&|\|\||[\s\S]
Private Function OperatorLength(ByVal value As String, ByVal p As Long) As Long
    Dim s3 As String, s2 As String
    s3 = Mid$(value, p + 1, 3)
    If s3 = ":::" Or s3 = "<<-" Or s3 = "->>" Then OperatorLength = 3: Exit Function
    s2 = Left$(s3, 2)
    Select Case s2
        Case "::", "<-", "->", "|>", "<=", ">=", "==", "!=", "&&", "||": OperatorLength = 2
        Case Else: OperatorLength = 1
    End Select
End Function

Private Function TT(ByRef t As RToks, ByVal i As Long) As String
    If i < 0 Or i >= t.Count Then TT = UNDEF Else TT = t.Txt(i)
End Function

' rSymbol: /^[A-Za-z.][A-Za-z0-9._]*$/
Public Function VarIsSymbol(ByVal s As String) As Boolean
    Dim i As Long, c As Long, n As Long
    n = Len(s)
    If n = 0 Then Exit Function
    For i = 1 To n
        c = AscW(Mid$(s, i, 1))
        Select Case c
            Case 65 To 90, 97 To 122, 46
            Case 48 To 57, 95
                If i = 1 Then Exit Function
            Case Else
                Exit Function
        End Select
    Next i
    VarIsSymbol = True
End Function

Private Function IsSymTok(ByRef t As RToks, ByVal i As Long) As Boolean
    IsSymTok = VarIsSymbol(TT(t, i))
End Function

Private Function IsQuoted(ByVal s As String) As Boolean
    Dim c As String
    c = Left$(s, 1)
    IsQuoted = (c = """" Or c = "'")
End Function

Private Function SliceStr(ByVal value As String, ByVal s As Long, ByVal e As Long) As String
    If s < 0 Then s = 0
    If e > Len(value) Then e = Len(value)
    If e > s Then SliceStr = Mid$(value, s + 1, e - s)
End Function

Private Function GapHas(ByVal value As String, ByVal s As Long, ByVal e As Long, ByVal needle As String) As Boolean
    If e > s Then GapHas = InStr(1, Mid$(value, s + 1, e - s), needle, vbBinaryCompare) > 0
End Function

Private Function InSpaced(ByVal spaced As String, ByVal word As String) As Boolean
    If Len(word) = 0 Then Exit Function
    InSpaced = InStr(1, spaced, " " & word & " ", vbBinaryCompare) > 0
End Function

' rValueReference(tokens, index)
Private Function RValueRef(ByRef t As RToks, ByVal i As Long) As Boolean
    Dim before As String, after As String
    If Not IsSymTok(t, i) Then Exit Function
    before = TT(t, i - 1)
    after = TT(t, i + 1)
    If before = "$" Or before = "@" Or before = "::" Or before = ":::" Then Exit Function
    If after = "(" Or after = "::" Or after = ":::" Or after = "=" Then Exit Function
    RValueRef = True
End Function

' Public rTokens for other modules: clsList of clsMap { text, start, end, parent[, open][, close] }.
Public Function VarTokenList(ByVal value As String) As clsList
    Dim t As RToks, i As Long, m As clsMap, out As clsList
    Tokenize value, t
    Set out = NewList()
    For i = 0 To t.Count - 1
        Set m = NewMap()
        m("text") = t.Txt(i)
        m("start") = t.St(i)
        m("end") = t.En(i)
        m("parent") = t.Par(i)
        If t.OpenOf(i) >= 0 Then m("open") = t.OpenOf(i)
        If t.CloseOf(i) >= 0 Then m("close") = t.CloseOf(i)
        out.Push m
    Next i
    Set VarTokenList = out
End Function

' ================================================================ edits

Private Sub AddEdit(ByRef ed As REdits, ByVal s As Long, ByVal e As Long, ByVal text As String)
    If ed.Count = 0 Then
        ReDim ed.S(0 To 7): ReDim ed.E(0 To 7): ReDim ed.T(0 To 7)
    ElseIf ed.Count > UBound(ed.S) Then
        ReDim Preserve ed.S(0 To ed.Count * 2 - 1): ReDim Preserve ed.E(0 To ed.Count * 2 - 1): ReDim Preserve ed.T(0 To ed.Count * 2 - 1)
    End If
    ed.S(ed.Count) = s
    ed.E(ed.Count) = e
    ed.T(ed.Count) = text
    ed.Count = ed.Count + 1
End Sub

' editRText: stable sort by start then end, both descending; apply from the end.
Private Function ApplyEdits(ByVal value As String, ByRef ed As REdits) As String
    Dim order() As Long, i As Long, j As Long, k As Long, text As String
    If ed.Count = 0 Then ApplyEdits = value: Exit Function
    ReDim order(0 To ed.Count - 1)
    For i = 0 To ed.Count - 1
        k = i
        j = i - 1
        Do While j >= 0
            If Not EditBefore(ed, k, order(j)) Then Exit Do
            order(j + 1) = order(j)
            j = j - 1
        Loop
        order(j + 1) = k
    Next i
    text = value
    For i = 0 To ed.Count - 1
        k = order(i)
        text = SliceStr(text, 0, ed.S(k)) & ed.T(k) & Mid$(text, IIf(ed.E(k) + 1 < 1, 1, ed.E(k) + 1))
    Next i
    ApplyEdits = text
End Function

' True when edit a sorts strictly before edit b (b.start - a.start || b.end - a.end < 0).
Private Function EditBefore(ByRef ed As REdits, ByVal a As Long, ByVal b As Long) As Boolean
    If ed.S(a) <> ed.S(b) Then
        EditBefore = ed.S(a) > ed.S(b)
    Else
        EditBefore = ed.E(a) > ed.E(b)
    End If
End Function

' ================================================================ statements

Private Function Continues(ByVal s As String) As Boolean
    Select Case s
        Case "<-", "=", "~", "+", "-", "*", "/", "^", ",", "|>", "&", "|", "&&", "||": Continues = True
        Case Else: Continues = IsPercentOp(s)
    End Select
End Function

' %.*% as a whole token (. excludes line terminators).
Private Function IsPercentOp(ByVal s As String) As Boolean
    Dim i As Long, c As Long
    If Len(s) < 2 Then Exit Function
    If Left$(s, 1) <> "%" Or Right$(s, 1) <> "%" Then Exit Function
    For i = 2 To Len(s) - 1
        c = AscW(Mid$(s, i, 1)) And &HFFFF&
        If c = 10 Or c = 13 Or c = &H2028& Or c = &H2029& Then Exit Function
    Next i
    IsPercentOp = True
End Function

' rStatements(value): the code split before each top-level statement's line break.
Public Function VarRStatements(ByVal value As String) As clsList
    Dim t As RToks, out As clsList, i As Long, start As Long, isEnd As Boolean, tx As String, top As Boolean
    Tokenize value, t
    Set out = NewList()
    start = 0
    For i = 0 To t.Count - 1
        If i = t.Count - 1 Then
            isEnd = True
        Else
            isEnd = False
            tx = t.Txt(i)
            If Not (tx = "(" Or tx = "[" Or tx = "{") Then
                top = (t.Par(i) = -1)
                If Not top Then
                    If t.OpenOf(i) >= 0 Then top = (t.Par(t.OpenOf(i)) = -1)
                End If
                If top Then
                    If Not Continues(tx) Then
                        If Not Continues(t.Txt(i + 1)) Then isEnd = GapHas(value, t.En(i), t.St(i + 1), vbLf)
                    End If
                End If
            End If
        End If
        If isEnd Then
            out.Push SliceStr(value, start, t.En(i))
            start = t.En(i)
        End If
    Next i
    If start < Len(value) Then out.Push Mid$(value, start + 1)
    Set VarRStatements = out
End Function

' ================================================================ names

' value.match(/[A-Za-z.][A-Za-z0-9._]*/g) added to a set.
Private Sub AddIdentifiers(ByVal value As String, ByVal names As clsMap)
    Dim n As Long, p As Long, q As Long, b() As Byte, cs() As Long, i As Long
    n = Len(value)
    If n = 0 Then Exit Sub
    b = value
    ReDim cs(0 To n)
    For i = 0 To n - 1
        cs(i) = b(2 * i) + 256& * b(2 * i + 1)
    Next i
    cs(n) = -1
    p = 0
    Do While p < n
        If IsIdentStart(cs(p)) Then
            q = p + 1
            Do While IsIdentPart(cs(q))
                q = q + 1
            Loop
            names(Mid$(value, p + 1, q - p)) = True
            p = q
        Else
            p = p + 1
        End If
    Loop
End Sub

' protectedRNames(value, tokens): a set (clsMap name -> True).
Private Function ProtectedRNames(ByVal value As String, ByRef t As RToks) As clsMap
    Dim names As clsMap, i As Long, k As Long, at As Long, nm As String, before As String
    Dim seps() As Long, sepCount As Long, masked As Boolean, functionStart As Long, piped As Boolean
    Dim hasData As Boolean, formula As Boolean, a0 As Long, a1 As Long, after As String
    Dim bench As clsNode, inputs As clsList, el As clsNode
    Set names = NewMap()
    names("dataset") = True
    names("column1") = True
    names("column2") = True
    For i = 0 To t.Count - 1
        If IsSymTok(t, i) Then
            before = TT(t, i - 1)
            If before = "$" Or before = "@" Then names(t.Txt(i)) = True
        End If
        If t.Txt(i) = "(" And t.CloseOf(i) >= 0 Then
            nm = TT(t, i - 1)
            sepCount = 0
            ReDim seps(0 To 7)
            seps(0) = i: sepCount = 1
            For at = i + 1 To t.CloseOf(i) - 1
                If t.Txt(at) = "," And t.Par(at) = i Then
                    If sepCount > UBound(seps) Then ReDim Preserve seps(0 To sepCount * 2)
                    seps(sepCount) = at: sepCount = sepCount + 1
                End If
            Next at
            If sepCount > UBound(seps) Then ReDim Preserve seps(0 To sepCount)
            seps(sepCount) = t.CloseOf(i)
            Select Case nm
                Case "group_by", "select", "summarise", "summarize", "mutate", "filter", "count", "across", "complete", "spread": masked = True
                Case Else: masked = False
            End Select
            If TT(t, i - 2) = "::" Then functionStart = i - 3 Else functionStart = i - 1
            before = TT(t, functionStart - 1)
            piped = (before = "%>%" Or before = "|>")
            hasData = False
            For k = 0 To sepCount - 1
                a0 = seps(k) + 1
                a1 = seps(k + 1) - 1
                If a1 >= a0 + 1 Then
                    If t.Txt(a0) = "data" And t.Txt(a0 + 1) = "=" Then hasData = True
                End If
            Next k
            For k = 0 To sepCount - 1
                a0 = seps(k) + 1
                a1 = seps(k + 1) - 1
                formula = False
                If hasData Then
                    For at = a0 To a1
                        If t.Txt(at) = "~" Then formula = True: Exit For
                    Next at
                End If
                If formula Or (masked And (piped Or k <> 0)) Then
                    For at = a0 To a1
                        If RValueRef(t, at) Then
                            after = TT(t, at + 1)
                            If Not (after = "$" Or after = "@") Then names(t.Txt(at)) = True
                        End If
                    Next at
                End If
            Next k
        End If
    Next i
    ' Explicit object-name fields of the visible operation.
    Set bench = modVariants.VisibleOperationBench()
    If Not bench Is Nothing Then
        Set inputs = bench.QuerySelectorAll("input:not([type=""number""]):not([type=""checkbox""]), textarea.column-input")
        For i = 0 To inputs.Count - 1
            Set el = inputs(i)
            If Not RxTest(el.Id, "output|search|title|caption", "i") Then AddIdentifiers el.Value, names
        Next i
    End If
    Set ProtectedRNames = names
End Function

Private Function IsAssignContinuation(ByVal s As String) As Boolean
    Select Case s
        Case "<-", "=", "|>", "+", "-", "*", "/", "~": IsAssignContinuation = True
        Case Else: IsAssignContinuation = IsPercentOp(s)
    End Select
End Function

' assignedRObjects(value): temporary objects safe to rename, in first-assignment order.
' Callers only read the returned list (it may be the memoized one).
Private Function AssignedObjects(ByVal value As String) As clsList
    Dim t As RToks, names As clsMap, protectedSet As clsMap, i As Long, j As Long, endAt As Long
    Dim nxt As String, assignment As Boolean, isLoop As Boolean, priorRead As Boolean, selfRead As Boolean
    Dim parentText As String, memo As Boolean, result As clsList
    If mAssignedActive Then
        If mAssignedFor = value Then
            If Not mAssignedResult Is Nothing Then Set AssignedObjects = mAssignedResult: Exit Function
            memo = True
        End If
    End If
    Tokenize value, t
    Set names = NewMap()
    If memo Then Set protectedSet = mAssignedProtected Else Set protectedSet = ProtectedRNames(value, t)
    For i = 0 To t.Count - 1
        If IsSymTok(t, i) Then
            nxt = TT(t, i + 1)
            parentText = TT(t, t.Par(i))
            assignment = (nxt = "<-")
            If Not assignment And nxt = "=" Then assignment = (t.Par(i) = -1 Or parentText = "{")
            isLoop = (TT(t, i - 2) = "for" And TT(t, i - 1) = "(")
            If (assignment Or isLoop) And Not names.Exists(t.Txt(i)) And Not protectedSet.Exists(t.Txt(i)) Then
                priorRead = False
                For j = 0 To i - 1
                    If t.Txt(j) = t.Txt(i) Then
                        If RValueRef(t, j) Then priorRead = True: Exit For
                    End If
                Next j
                endAt = i + 2
                Do While endAt < t.Count
                    If endAt > i + 2 Then
                        If t.Par(endAt) = t.Par(i) Then
                            If GapHas(value, t.En(endAt - 1), t.St(endAt), vbLf) Then
                                If Not IsAssignContinuation(t.Txt(endAt - 1)) Then Exit Do
                            End If
                        End If
                    End If
                    endAt = endAt + 1
                Loop
                selfRead = False
                If Not isLoop Then
                    For j = i + 2 To endAt - 1
                        If t.Txt(j) = t.Txt(i) Then
                            If RValueRef(t, j) Then selfRead = True: Exit For
                        End If
                    Next j
                End If
                If Not priorRead And Not selfRead Then names(t.Txt(i)) = True
            End If
        End If
    Next i
    Set result = names.Keys()
    If memo Then Set mAssignedResult = result
    Set AssignedObjects = result
End Function

Private Function ObjectNameType(ByVal nm As String) As String
    Dim kind As String
    If mObjectTypes.Exists(nm) Then
        kind = mObjectTypes(nm)
    Else
        kind = ObjectNameTypeOf(nm)
        mObjectTypes(nm) = kind
    End If
    ObjectNameType = kind
End Function

Private Function ObjectNameTypeOf(ByVal nm As String) As String
    If RxTest(nm, "^(i|j|k|index|fold|row_num|step_num)$", "") Then
        ObjectNameTypeOf = "loop"
    ElseIf RxTest(nm, "matrix|constraint_mat", "i") Then
        ObjectNameTypeOf = "matrix"
    ElseIf RxTest(nm, "table|tab|freq|summary", "i") Then
        ObjectNameTypeOf = "table"
    ElseIf RxTest(nm, "data|frame|subset|filter|select|grouped|(?:^|[._])df(?:$|[._])|^X(?:_scaled)?$", "i") Then
        ObjectNameTypeOf = "data"
    ElseIf RxTest(nm, "model|fit|linreg|regression|anova|glm|km_obj|pca", "i") Then
        ObjectNameTypeOf = "model"
    ElseIf RxTest(nm, "colou?r|palette", "i") Then
        ObjectNameTypeOf = "colours"
    ElseIf RxTest(nm, "tick|mark", "i") Then
        ObjectNameTypeOf = "ticks"
    ElseIf RxTest(nm, "range|limit|bound", "i") Then
        ObjectNameTypeOf = "range"
    ElseIf RxTest(nm, "vector|coeff|rhs|dir|objective", "i") Then
        ObjectNameTypeOf = "vector"
    ElseIf RxTest(nm, "value|count|mean|mids|breaks|labels", "i") Then
        ObjectNameTypeOf = "values"
    ElseIf RxTest(nm, "result|output|solution|test", "i") Then
        ObjectNameTypeOf = "result"
    Else
        ObjectNameTypeOf = "generic"
    End If
End Function

' name.replace(/_([a-z])/g, (_, letter) => letter.toUpperCase())
Private Function CamelName(ByVal nm As String) As String
    Dim i As Long, c As Long, out As String
    i = 1
    Do While i <= Len(nm)
        If Mid$(nm, i, 1) = "_" And i < Len(nm) Then
            c = AscW(Mid$(nm, i + 1, 1))
            If c >= 97 And c <= 122 Then
                out = out & ChrW$(c - 32)
                i = i + 2
            Else
                out = out & "_"
                i = i + 1
            End If
        Else
            out = out & Mid$(nm, i, 1)
            i = i + 1
        End If
    Loop
    CamelName = out
End Function

' freshRName(type, reserved): adds the chosen name to reserved.
Public Function VarFreshRName(ByVal kind As String, ByVal reserved As clsMap) As String
    Dim pool As Variant, candidates() As String, n As Long, i As Long, nm As String, base As String, poolKind As String
    EnsureData
    If mNamePools.Exists(kind) Then poolKind = kind Else poolKind = "generic"
    pool = mNamePools(poolKind)
    n = UBound(pool) - LBound(pool) + 1
    If Not mCandidates.Exists(poolKind) Then
        ReDim candidates(0 To 3 * n - 1)
        For i = 0 To n - 1
            candidates(3 * i) = pool(LBound(pool) + i)
            candidates(3 * i + 1) = Replace(pool(LBound(pool) + i), "_", ".")
            candidates(3 * i + 2) = CamelName(pool(LBound(pool) + i))
        Next i
        mCandidates(poolKind) = candidates
    End If
    candidates = mCandidates(poolKind)
    VarShuffleStrings candidates, 0, 3 * n - 1
    For i = 0 To 3 * n - 1
        If Not reserved.Exists(candidates(i)) Then nm = candidates(i): Exit For
    Next i
    If Len(nm) = 0 Then
        base = pool(LBound(pool) + VarRandomIndex(n))
        Do
            nm = base & "_" & CStr(2 + VarRandomIndex(999))
        Loop While reserved.Exists(nm)
    End If
    reserved(nm) = True
    VarFreshRName = nm
End Function

' replaceSymbols inside renameRObjects.
Private Function ReplaceSymbols(ByVal text As String, ByVal mapping As clsMap, ByVal preserveLabels As Boolean) As String
    Dim t As RToks, ed As REdits, i As Long, before As String, after As String, parentText As String, replacement As String
    Tokenize text, t
    For i = 0 To t.Count - 1
        If mapping.Exists(t.Txt(i)) Then
            before = TT(t, i - 1)
            after = TT(t, i + 1)
            If Not (before = "$" Or before = "@" Or before = "::" Or before = ":::" Or after = "(" Or after = "::" Or after = ":::") Then
                parentText = TT(t, t.Par(i))
                If Not (after = "=" And parentText = "(") Then
                    replacement = mapping(t.Txt(i))
                    If preserveLabels And parentText = "(" And TT(t, t.Par(i) - 1) = "cbind" _
                        And (before = "(" Or before = ",") And (after = "," Or after = ")") Then replacement = t.Txt(i) & " = " & replacement
                    AddEdit ed, t.St(i), t.En(i), replacement
                End If
            End If
        End If
    Next i
    ReplaceSymbols = ApplyEdits(text, ed)
End Function

' \{\{|\}\}|\{([^{}]*)\} inside glue() strings: only interpolations are renamed.
Private Function GlueReplace(ByVal text As String, ByVal mapping As clsMap) As String
    Dim i As Long, n As Long, j As Long, ch As String, out As String, c As String, found As Boolean
    n = Len(text)
    i = 1
    Do While i <= n
        ch = Mid$(text, i, 1)
        If ch = "{" And Mid$(text, i + 1, 1) = "{" Then
            out = out & "{{": i = i + 2
        ElseIf ch = "}" And Mid$(text, i + 1, 1) = "}" Then
            out = out & "}}": i = i + 2
        ElseIf ch = "{" Then
            found = False
            For j = i + 1 To n
                c = Mid$(text, j, 1)
                If c = "}" Then found = True: Exit For
                If c = "{" Then Exit For
            Next j
            If found Then
                out = out & "{" & ReplaceSymbols(Mid$(text, i + 1, j - i - 1), mapping, False) & "}"
                i = j + 1
            Else
                out = out & ch: i = i + 1
            End If
        Else
            out = out & ch: i = i + 1
        End If
    Loop
    GlueReplace = out
End Function

Private Function RenameObjects(ByVal value As String, ByVal reserved As clsMap) As String
    Dim names As clsList, mapping As clsMap, i As Long, code As String, t As RToks, ed As REdits, replacement As String
    Set names = AssignedObjects(value)
    Set mapping = NewMap()
    For i = 0 To names.Count - 1
        mapping(CStr(names(i))) = VarFreshRName(ObjectNameType(CStr(names(i))), reserved)
    Next i
    code = ReplaceSymbols(value, mapping, True)
    Tokenize code, t
    For i = 0 To t.Count - 1
        If IsQuoted(t.Txt(i)) Then
            If TT(t, t.Par(i)) = "(" And TT(t, t.Par(i) - 1) = "glue" Then
                replacement = GlueReplace(t.Txt(i), mapping)
                If replacement <> t.Txt(i) Then AddEdit ed, t.St(i), t.En(i), replacement
            End If
        End If
    Next i
    RenameObjects = ApplyEdits(code, ed)
End Function

' ================================================================ structure

Private Function IsPrintableCall(ByVal nm As String) As Boolean
    Select Case nm
        Case "shapiro.test", "cor.test", "corr.test", "t.test", "prop.test", "var.test", "bartlett.test", "shapiro_test", _
             "welch_anova_test", "games_howell_test", "confint", "anova", "confusionMatrix": IsPrintableCall = True
    End Select
End Function

Private Function StoreReports(ByVal value As String, ByVal reserved As clsMap) As String
    Dim statements As clsList, k As Long, out As String
    Set statements = VarRStatements(value)
    For k = 0 To statements.Count - 1
        out = out & StoreReport(CStr(statements(k)), reserved)
    Next k
    StoreReports = out
End Function

' True when no printable call name occurs in the text at all (then StoreReport cannot
' change the statement, and it draws no random numbers).
Private Function NoPrintableName(ByRef text As String) As Boolean
    If InStr(1, text, "test", vbBinaryCompare) > 0 Then Exit Function
    If InStr(1, text, "anova", vbBinaryCompare) > 0 Then Exit Function
    If InStr(1, text, "confint", vbBinaryCompare) > 0 Then Exit Function
    If InStr(1, text, "confusionMatrix", vbBinaryCompare) > 0 Then Exit Function
    NoPrintableName = True
End Function

Private Function StoreReport(ByVal statement As String, ByVal reserved As clsMap) As String
    Dim t As RToks, i As Long, anyPrintable As Boolean, anyCall As Boolean, first As Long, expression As String, nm As String
    StoreReport = statement
    If NoPrintableName(statement) Then Exit Function
    Tokenize statement, t
    If t.Count = 0 Then Exit Function
    For i = 0 To t.Count - 1
        If t.Txt(i) = "<-" Or t.Txt(i) = "{" Or t.Txt(i) = "->" Then Exit Function
    Next i
    For i = 0 To t.Count - 1
        If IsSymTok(t, i) And TT(t, i + 1) = "(" Then
            anyCall = True
            If IsPrintableCall(t.Txt(i)) Then anyPrintable = True
            If t.Txt(i) = "print" Or t.Txt(i) = "cat" Or t.Txt(i) = "library" Then Exit Function
        End If
    Next i
    If Not anyCall Or Not anyPrintable Then Exit Function
    For i = 1 To t.Count - 1
        If t.Txt(i) = "=" And t.Par(i - 1) = -1 Then Exit Function
    Next i
    first = t.St(0)
    expression = JsTrimEnd(Mid$(statement, first + 1))
    nm = VarFreshRName("result", reserved)
    If VarRandomIndex(2) <> 0 Then
        StoreReport = Left$(statement, first) & nm & " <- " & expression & vbLf & "print(" & nm & ")"
    Else
        StoreReport = Left$(statement, first) & nm & " <- " & expression & vbLf & nm
    End If
End Function

Private Function VaryStructure(ByVal value As String) As String
    Dim t As RToks, ed As REdits, i As Long, assignments As Long
    Tokenize value, t
    assignments = VarRandomIndex(2)
    If assignments <> 0 Then
        For i = 1 To t.Count - 1
            If t.Txt(i) = "<-" And t.Par(i - 1) = -1 Then AddEdit ed, t.St(i), t.En(i), "="
        Next i
    End If
    VaryStructure = ApplyEdits(value, ed)
End Function

' ================================================================ plot styles

Private Sub RandomPlotStylePlans()
    Dim schemes(0 To 13) As String, shapes(0 To 6) As String, brewer(0 To 13) As String, i As Long, parts() As String
    shapes(0) = "15": shapes(1) = "16": shapes(2) = "17": shapes(3) = "18": shapes(4) = "8": shapes(5) = "4": shapes(6) = "3"
    VarShuffleStrings shapes, 0, 6
    brewer(0) = "Accent": brewer(1) = "Dark2": brewer(2) = "Paired": brewer(3) = "Pastel1": brewer(4) = "Pastel2"
    brewer(5) = "Set1": brewer(6) = "BrBG": brewer(7) = "PiYG": brewer(8) = "PRGn": brewer(9) = "PuOr"
    brewer(10) = "RdBu": brewer(11) = "RdGy": brewer(12) = "RdYlBu": brewer(13) = "Spectral"
    VarShuffleStrings brewer, 0, 13
    schemes(0) = "#2166AC|#B2182B": schemes(1) = "#00796B|#D95F02"
    schemes(2) = "#762A83|#1B7837": schemes(3) = "#A6611A|#018571"
    schemes(4) = "#C51B7D|#4D9221": schemes(5) = "#3B4CC0|#B40426"
    schemes(6) = "#00838F|#AD1457": schemes(7) = "#6A51A3|#E6550D"
    schemes(8) = "#33691E|#8E24AA": schemes(9) = "#D84315|#1565C0"
    schemes(10) = "#455A64|#C62828": schemes(11) = "#827717|#4527A0"
    schemes(12) = "#00695C|#C2185B": schemes(13) = "#795548|#0277BD"
    VarShuffleStrings schemes, 0, 13
    For i = 1 To 5
        parts = Split(schemes(i - 1), "|")
        mPlanPrimary(i) = parts(0)
        mPlanAccent(i) = parts(1)
        mPlanBrewer(i) = brewer(i - 1)
        mPlanShape(i) = shapes(i - 1)
    Next i
End Sub

Private Function DirectChildren(ByRef t As RToks, ByVal opener As Long, ByRef kids() As Long) As Long
    Dim at As Long, n As Long
    ReDim kids(0 To 7)
    For at = opener + 1 To t.CloseOf(opener) - 1
        If t.Par(at) = opener Then
            If n > UBound(kids) Then ReDim Preserve kids(0 To n * 2)
            kids(n) = at
            n = n + 1
        End If
    Next at
    DirectChildren = n
End Function

Private Function KidText(ByRef t As RToks, ByRef kids() As Long, ByVal n As Long, ByVal position As Long) As String
    If position < 0 Or position >= n Then KidText = UNDEF Else KidText = t.Txt(kids(position))
End Function

Private Function ChangePlotStyle(ByVal value As String, ByVal plan As Long) As String
    Dim t As RToks, ed As REdits, i As Long, colourMap As clsMap, kids() As Long, nk As Long, position As Long
    Dim prevName As String, tx As String, callName As String, args As String
    Dim namedPalette As Boolean, positionalPalette As Boolean
    If plan = 0 Then ChangePlotStyle = value: Exit Function
    Tokenize value, t
    Set colourMap = NewMap()
    For i = 0 To t.Count - 1
        tx = t.Txt(i)
        If tx = "rainbow" And TT(t, i + 1) = "(" Then
            If t.CloseOf(i + 1) >= 0 Then AddEdit ed, t.St(i), t.En(t.CloseOf(i + 1)), "RColorBrewer::brewer.pal(8, """ & mPlanBrewer(plan) & """)"
        Else
            If IsQuoted(tx) And TT(t, i - 1) = "=" Then
                Select Case TT(t, i - 2)
                    Case "col", "fill", "colour", "color", "barfill", "barcolor"
                        If Not colourMap.Exists(tx) Then
                            If colourMap.Count Mod 2 = 1 Then colourMap(tx) = mPlanAccent(plan) Else colourMap(tx) = mPlanPrimary(plan)
                        End If
                        AddEdit ed, t.St(i), t.En(i), """" & colourMap(tx) & """"
                End Select
            End If
            If tx = "19" And TT(t, i - 1) = "=" And TT(t, i - 2) = "pch" Then AddEdit ed, t.St(i), t.En(i), mPlanShape(plan)
            If (tx = "brewer.pal" Or tx = "scale_fill_brewer") And TT(t, i + 1) = "(" Then
                If t.CloseOf(i + 1) >= 0 Then
                    nk = DirectChildren(t, i + 1, kids)
                    If tx = "brewer.pal" Then prevName = "name" Else prevName = "palette"
                    For position = 0 To nk - 1
                        namedPalette = (KidText(t, kids, nk, position - 1) = "=" And KidText(t, kids, nk, position - 2) = prevName)
                        positionalPalette = (tx = "brewer.pal" And KidText(t, kids, nk, position - 1) = ",")
                        If IsQuoted(t.Txt(kids(position))) And (namedPalette Or positionalPalette) Then
                            AddEdit ed, t.St(kids(position)), t.En(kids(position)), """" & mPlanBrewer(plan) & """"
                        End If
                    Next position
                End If
            End If
        End If
    Next i
    value = ApplyEdits(value, ed)
    Tokenize value, t
    ed.Count = 0
    For i = 0 To t.Count - 1
        If t.Txt(i) = "(" And t.CloseOf(i) >= 0 Then
            callName = TT(t, i - 1)
            nk = DirectChildren(t, i, kids)
            args = ""
            Select Case callName
                Case "plot", "barplot", "boxplot", "hist", "pie", "mosaicplot", "biplot"
                    If Not KidHas(t, kids, nk, "col") Then
                        If callName = "hist" And KidHas(t, kids, nk, "plot") Then GoTo NextToken
                        If callName = "biplot" Then
                            args = ", col = c(""" & mPlanPrimary(plan) & """, """ & mPlanAccent(plan) & """)"
                        Else
                            args = ", col = """ & mPlanPrimary(plan) & """"
                        End If
                    End If
                Case "pairs.panels"
                    If Not KidHas(t, kids, nk, "col") Then args = ", col = """ & mPlanPrimary(plan) & """, hist.col = """ & mPlanAccent(plan) & """, pch = " & mPlanShape(plan)
                Case "fviz_eig"
                    If Not KidHas(t, kids, nk, "barfill") Then args = ", barfill = """ & mPlanPrimary(plan) & """, barcolor = """ & mPlanPrimary(plan) & """, linecolor = """ & mPlanAccent(plan) & """"
                Case "fviz_cluster"
                    If Not KidHas(t, kids, nk, "palette") Then args = ", palette = """ & mPlanBrewer(plan) & """"
            End Select
            If Len(args) > 0 Then AddEdit ed, t.St(t.CloseOf(i)), t.St(t.CloseOf(i)), args
        End If
NextToken:
    Next i
    ChangePlotStyle = ApplyEdits(value, ed)
End Function

Private Function KidHas(ByRef t As RToks, ByRef kids() As Long, ByVal n As Long, ByVal nm As String) As Boolean
    Dim k As Long
    For k = 0 To n - 1
        If t.Txt(kids(k)) = nm And KidText(t, kids, n, k + 1) = "=" Then KidHas = True: Exit Function
    Next k
End Function

' ================================================================ argument order

Private Sub ArgumentCalls(ByRef t As RToks, ByRef calls() As RCall)
    Dim occurrences As clsMap, index As Long, nm As String, before As String, ok As Boolean
    Dim seps() As Long, sepCount As Long, at As Long, slot As Long, leftAt As Long, rightAt As Long
    Dim c As RCall, j As Long, dup As Boolean, emptyArg As Boolean, sorted() As String, signature As String, occurrence As Long
    If t.Count = 0 Then
        ReDim calls(0 To 0)
        Exit Sub
    End If
    ReDim calls(0 To t.Count - 1)
    Set occurrences = NewMap()
    For index = 0 To t.Count - 1
        If t.Txt(index) = "(" And t.CloseOf(index) >= 0 And IsSymTok(t, index - 1) Then
            nm = t.Txt(index - 1)
            If mPolicyNames.Exists(nm) Then
                before = TT(t, index - 2)
                ok = Not (before = "$" Or before = "@" Or before = ":::")
                If ok And before = "::" Then ok = InSpaced(mPolicyPackages(nm), TT(t, index - 3))
                If ok Then
                    ReDim seps(0 To 7)
                    seps(0) = index: sepCount = 1
                    For at = index + 1 To t.CloseOf(index) - 1
                        If t.Txt(at) = "," And t.Par(at) = index Then
                            If sepCount > UBound(seps) Then ReDim Preserve seps(0 To sepCount * 2)
                            seps(sepCount) = at: sepCount = sepCount + 1
                        End If
                    Next at
                    If sepCount > UBound(seps) Then ReDim Preserve seps(0 To sepCount)
                    seps(sepCount) = t.CloseOf(index)
                    c.ArgCount = sepCount
                    ReDim c.ArgLeft(0 To sepCount - 1): ReDim c.ArgRight(0 To sepCount - 1): ReDim c.ArgName(0 To sepCount - 1)
                    ReDim c.ArgStart(0 To sepCount - 1): ReDim c.ArgEnd(0 To sepCount - 1): ReDim c.Slots(0 To sepCount - 1)
                    c.SlotCount = 0
                    dup = False
                    emptyArg = False
                    For slot = 0 To sepCount - 1
                        leftAt = seps(slot): rightAt = seps(slot + 1)
                        c.ArgLeft(slot) = leftAt
                        c.ArgRight(slot) = rightAt
                        c.ArgName(slot) = ""
                        If t.Par(leftAt + 1) = index And TT(t, leftAt + 2) = "=" Then c.ArgName(slot) = t.Txt(leftAt + 1)
                        c.ArgStart(slot) = t.En(leftAt)
                        c.ArgEnd(slot) = t.St(rightAt)
                        If rightAt = leftAt + 1 Then emptyArg = True
                        If Len(c.ArgName(slot)) > 0 Then
                            For j = 0 To slot - 1
                                If c.ArgName(j) = c.ArgName(slot) Then dup = True
                            Next j
                        End If
                    Next slot
                    If Not dup Then
                        For slot = 0 To sepCount - 1
                            If InSpaced(mPolicyNames(nm), c.ArgName(slot)) Then
                                c.Slots(c.SlotCount) = slot
                                c.SlotCount = c.SlotCount + 1
                            End If
                        Next slot
                        If c.SlotCount >= 2 And Not emptyArg Then
                            ReDim sorted(0 To c.SlotCount - 1)
                            For j = 0 To c.SlotCount - 1
                                sorted(j) = c.ArgName(c.Slots(j))
                            Next j
                            SortStringsBinary sorted, c.SlotCount
                            signature = nm & ":" & Join(sorted, ",")
                            occurrence = 0
                            If occurrences.Exists(signature) Then occurrence = occurrences(signature)
                            occurrences(signature) = occurrence + 1
                            c.Key = signature & ":" & CStr(occurrence)
                            c.Valid = True
                            calls(index) = c
                            c.Valid = False
                        End If
                    End If
                End If
            End If
        End If
    Next index
End Sub

Private Sub SortStringsBinary(ByRef values() As String, ByVal n As Long)
    Dim i As Long, j As Long, v As String
    For i = 1 To n - 1
        v = values(i)
        j = i - 1
        Do While j >= 0
            If StrComp(values(j), v, vbBinaryCompare) <= 0 Then Exit Do
            values(j + 1) = values(j)
            j = j - 1
        Loop
        values(j + 1) = v
    Next i
End Sub

Private Function CallOrderText(ByRef c As RCall, ByRef order() As Long) As String
    Dim i As Long, out As String
    For i = 0 To c.SlotCount - 1
        If i > 0 Then out = out & ","
        out = out & c.ArgName(order(i))
    Next i
    CallOrderText = out
End Function

Private Function SafeArguments(ByVal value As String, ByRef t As RToks, ByVal index As Long, ByVal closeAt As Long) As Boolean
    Dim at As Long, tx As String, prev As String, nm As String, hasPolicy As Boolean, before As String, packages As String
    Dim prevParent As Long
    For at = index + 1 To closeAt - 1
        tx = t.Txt(at)
        If GapHas(value, t.En(at - 1), t.St(at), "#") Then Exit Function
        Select Case tx
            Case "<-", "<<-", "->", "->>", "{", "...": Exit Function
        End Select
        If tx = "=" Then
            prevParent = t.Par(at - 1)
            If Not IsSymTok(t, prevParent - 1) Then Exit Function
        End If
        If tx = "(" Then
            prev = t.Txt(at - 1)
            If prev = ")" Or prev = "]" Then Exit Function
            Select Case Left$(prev, 1)
                Case "`", """", "'": Exit Function
            End Select
            If IsSymTok(t, at - 1) Then
                nm = prev
                hasPolicy = mPolicyNames.Exists(nm)
                If InSpaced(mEffectfulCalls, nm) Then Exit Function
                If Not hasPolicy And Not InSpaced(mPureCalls, nm) Then Exit Function
                before = TT(t, at - 2)
                If before = "$" Or before = "@" Or before = ":::" Then Exit Function
                If before = "::" Then
                    If hasPolicy Then packages = mPolicyPackages(nm) Else packages = " base stats "
                    If Not InSpaced(packages, TT(t, at - 3)) Then Exit Function
                End If
            End If
        End If
    Next at
    SafeArguments = Not GapHas(value, t.En(closeAt - 1), t.St(closeAt), "#")
End Function

Private Function RenderRange(ByVal value As String, ByRef t As RToks, ByRef rendered() As String, ByRef hasRendered() As Boolean, _
    ByVal leftAt As Long, ByVal rightAt As Long, ByVal startPos As Long, ByVal endPos As Long) As String
    Dim ed As REdits, at As Long
    at = leftAt
    Do While at < rightAt
        If hasRendered(at) Then
            AddEdit ed, t.St(at) - startPos, t.En(t.CloseOf(at)) - startPos, rendered(at)
            at = t.CloseOf(at)
        End If
        at = at + 1
    Loop
    RenderRange = ApplyEdits(SliceStr(value, startPos, endPos), ed)
End Function

Private Function LeadingSpace(ByVal s As String) As String
    Dim i As Long
    i = 1
    Do While i <= Len(s)
        If Not JsIsWhitespaceCode(AscW(Mid$(s, i, 1)) And &HFFFF&) Then Exit Do
        i = i + 1
    Loop
    LeadingSpace = Left$(s, i - 1)
End Function

Private Function TrailingSpace(ByVal s As String) As String
    Dim i As Long
    i = Len(s)
    Do While i >= 1
        If Not JsIsWhitespaceCode(AscW(Mid$(s, i, 1)) And &HFFFF&) Then Exit Do
        i = i - 1
    Loop
    TrailingSpace = Mid$(s, i + 1)
End Function

Private Function ShuffleNamedArguments(ByVal value As String, ByVal previous As String) As String
    Dim t As RToks, calls() As RCall, tp As RToks, pcalls() As RCall, previousOrders As clsMap
    Dim rendered() As String, hasRendered() As Boolean, index As Long, i As Long, slot As Long
    Dim parts() As String, order() As Long, tmp As Long, priorOrder As String, content As String
    Dim replacement() As String, isReplaced() As Boolean
    EnsureData
    Tokenize value, t
    ArgumentCalls t, calls
    Tokenize previous, tp
    ArgumentCalls tp, pcalls
    Set previousOrders = NewMap()
    For i = 0 To tp.Count - 1
        If pcalls(i).Valid Then previousOrders(pcalls(i).Key) = CallOrderText(pcalls(i), pcalls(i).Slots)
    Next i
    If t.Count = 0 Then ShuffleNamedArguments = value: Exit Function
    ReDim rendered(0 To t.Count - 1)
    ReDim hasRendered(0 To t.Count - 1)
    For index = t.Count - 1 To 0 Step -1
        If t.CloseOf(index) >= 0 Then
            content = ""
            If calls(index).Valid Then
                If Not SafeArguments(value, t, index, t.CloseOf(index)) Then GoTo PlainRender
                With calls(index)
                    ReDim parts(0 To .ArgCount - 1)
                    For slot = 0 To .ArgCount - 1
                        parts(slot) = RenderRange(value, t, rendered, hasRendered, .ArgLeft(slot) + 1, .ArgRight(slot), .ArgStart(slot), .ArgEnd(slot))
                    Next slot
                    ReDim order(0 To .SlotCount - 1)
                    For i = 0 To .SlotCount - 1
                        order(i) = .Slots(i)
                    Next i
                    ShuffleLongs order, .SlotCount
                    If previousOrders.Exists(.Key) Then
                        priorOrder = previousOrders(.Key)
                        If CallOrderText(calls(index), order) = priorOrder Then
                            tmp = order(0): order(0) = order(1): order(1) = tmp
                        End If
                    End If
                    ReDim replacement(0 To .ArgCount - 1)
                    ReDim isReplaced(0 To .ArgCount - 1)
                    For i = 0 To .SlotCount - 1
                        isReplaced(.Slots(i)) = True
                        replacement(.Slots(i)) = JsTrim(parts(order(i)))
                    Next i
                    For slot = 0 To .ArgCount - 1
                        If slot > 0 Then content = content & ","
                        If isReplaced(slot) Then
                            content = content & LeadingSpace(parts(slot)) & replacement(slot) & TrailingSpace(parts(slot))
                        Else
                            content = content & parts(slot)
                        End If
                    Next slot
                End With
                GoTo Store
            End If
PlainRender:
            content = RenderRange(value, t, rendered, hasRendered, index + 1, t.CloseOf(index), t.En(index), t.St(t.CloseOf(index)))
Store:
            rendered(index) = t.Txt(index) & content & t.Txt(t.CloseOf(index))
            hasRendered(index) = True
        End If
    Next index
    ShuffleNamedArguments = RenderRange(value, t, rendered, hasRendered, 0, t.Count, 0, Len(value))
End Function

' ================================================================ codeAnswerVariants

Private Function ListHas(ByVal list As clsList, ByVal text As String) As Boolean
    Dim i As Long
    For i = 0 To list.Count - 1
        If CStr(list(i)) = text Then ListHas = True: Exit Function
    Next i
End Function

' codeAnswerVariants(value, previousVariants): six versions, entry 0 = the canonical code.
Public Function VarCodeAnswerVariants(ByVal value As String, Optional ByVal previousVariants As clsList = Nothing) As clsList
    Dim variants As clsList, reserved As clsMap, protectedSet As clsMap, t As RToks, index As Long, code As String
    Dim key As Variant, previous As String, ks As clsList
    EnsureData
    Set variants = NewList()
    If Len(JsTrim(value)) = 0 Then
        For index = 0 To 5
            variants.Push value
        Next index
        Set VarCodeAnswerVariants = variants
        Exit Function
    End If
    Set reserved = NewMap()
    AddIdentifiers value, reserved
    Tokenize value, t
    Set protectedSet = ProtectedRNames(value, t)
    Set ks = protectedSet.Keys()
    For Each key In ks
        reserved(CStr(key)) = True
    Next key
    mAssignedFor = value
    Set mAssignedProtected = protectedSet
    Set mAssignedResult = Nothing
    mAssignedActive = True
    RandomPlotStylePlans
    variants.Push value
    For index = 1 To 5
        code = RenameObjects(value, reserved)
        code = StoreReports(code, reserved)
        code = VaryStructure(code)
        code = ChangePlotStyle(code, index)
        previous = ""
        If Not previousVariants Is Nothing Then
            If index < previousVariants.Count Then
                If Not IsEmpty(previousVariants(index)) Then previous = CStr(previousVariants(index))
            End If
        End If
        code = ShuffleNamedArguments(code, previous)
        If ListHas(variants, code) Then code = FallbackResult(code, reserved)
        If ListHas(variants, code) Then code = FallbackIntermediate(code, reserved)
        variants.Push code
    Next index
    mAssignedActive = False
    Set mAssignedProtected = Nothing
    Set mAssignedResult = Nothing
    Set VarCodeAnswerVariants = variants
End Function

' A standalone output expression saved under a fresh result name.
Private Function FallbackResult(ByVal code As String, ByVal reserved As clsMap) As String
    Dim statements As clsList, k As Long, t As RToks, i As Long, ok As Boolean, statement As String, first As Long, nm As String
    FallbackResult = code
    Set statements = VarRStatements(code)
    For k = 0 To statements.Count - 1
        statement = statements(k)
        Tokenize statement, t
        ok = (t.Count > 0)
        If ok Then
            For i = 0 To t.Count - 1
                Select Case t.Txt(i)
                    Case "<-", "=", "{", "->": ok = False: Exit For
                End Select
            Next i
        End If
        If ok Then ok = Not RxTest(JsTrim(statement), "^(library|set.seed|cat|print|par)\s*\(", "")
        If ok Then
            first = t.St(0)
            nm = VarFreshRName("result", reserved)
            statements(k) = Left$(statement, first) & nm & " <- " & Mid$(statement, first + 1) & vbLf & nm
            FallbackResult = statements.Join("")
            Exit Function
        End If
    Next k
End Function

' An explicitly named destination keeps its name; a local intermediate is introduced.
Private Function FallbackIntermediate(ByVal code As String, ByVal reserved As clsMap) As String
    Dim statements As clsList, k As Long, t As RToks, statement As String, nm As String
    FallbackIntermediate = code
    Set statements = VarRStatements(code)
    For k = 0 To statements.Count - 1
        statement = statements(k)
        Tokenize statement, t
        If IsSymTok(t, 0) And (TT(t, 1) = "<-" Or TT(t, 1) = "=") Then
            If t.Count < 3 Then Exit Function
            nm = VarFreshRName("result", reserved)
            statements(k) = Left$(statement, t.St(0)) & nm & " <- " & Mid$(statement, t.St(2) + 1) & vbLf & t.Txt(0) & " <- " & nm
            FallbackIntermediate = statements.Join("")
            Exit Function
        End If
    Next k
End Function

' ================================================================ Math.random().toString(36)

' V8's DoubleToRadixCString for radix 36 and a value in [0, 1): "0.xxxx".
Public Function VarRandomToBase36(ByVal value As Double) As String
    Const CHARS As String = "0123456789abcdefghijklmnopqrstuvwxyz"
    Dim fraction As Double, delta As Double, ulp As Double, e As Long, v As Double, digit As Long
    Dim buffer As String, integerPart As Long, pos As Long, c As Long
    If value <= 0 Then VarRandomToBase36 = "0": Exit Function
    v = value
    e = 0
    Do While v < 1
        v = v * 2
        e = e - 1
    Loop
    ulp = 2 ^ (e - 52)
    delta = 0.5 * ulp
    fraction = value
    If fraction >= delta Then
        Do
            fraction = fraction * 36
            delta = delta * 36
            digit = Int(fraction)
            buffer = buffer & Mid$(CHARS, digit + 1, 1)
            fraction = fraction - digit
            If fraction > 0.5 Or (fraction = 0.5 And (digit And 1) = 1) Then
                If fraction + delta > 1 Then
                    Do
                        pos = Len(buffer)
                        If pos = 0 Then integerPart = integerPart + 1: Exit Do
                        c = InStr(1, CHARS, Right$(buffer, 1), vbBinaryCompare) - 1
                        buffer = Left$(buffer, pos - 1)
                        If c + 1 < 36 Then buffer = buffer & Mid$(CHARS, c + 2, 1): Exit Do
                    Loop
                    Exit Do
                End If
            End If
        Loop While fraction >= delta
    End If
    If Len(buffer) = 0 Then
        VarRandomToBase36 = CStr(integerPart)
    Else
        VarRandomToBase36 = CStr(integerPart) & "." & buffer
    End If
End Function
