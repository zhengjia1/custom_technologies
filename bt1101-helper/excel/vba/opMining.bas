Attribute VB_Name = "opMining"
' Tutorial 7 data mining: forward/backward stepwise selection, K-fold cross-validation,
' PCA (code sections, variance explained, loading equations, downstream linear-model
' component significance) and K-means clustering.
' Ported from index.html (app-src.js lines 2541-2764 and 4468-4493).
Option Explicit

Private Const NUMBER_PATTERN As String = "[-+]?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?"

' The shell procedure that opens the "disable scaling" modal (the website's
' showKmeansScaleAlert). Tests replace it with a stand-in.
Public KmeansScaleAlertProc As String

Private mStepwiseDirection As String
Private mPcaCodeSections As clsList          ' of clsMap: title, code
Private mPcaVarianceProse As String
Private mPcaVarianceMode As String
Private mPcaLoadingsProse As String
Private mPcaLinearProseSections As clsList   ' of clsMap: component, text

' ---------------------------------------------------------------- startup

Public Sub InitOpMining()
    Dim list As clsList, i As Long
    mStepwiseDirection = "backward"
    Set mPcaCodeSections = NewList()
    mPcaVarianceProse = ""
    mPcaVarianceMode = "individual"
    mPcaLoadingsProse = ""
    Set mPcaLinearProseSections = NewList()
    KmeansScaleAlertProc = "ShowKmeansScaleAlert"

    ' K-means
    Set list = QA(".kmeans-input")
    For i = 0 To list.Count - 1
        list(i).AddEventListener "input", "opMining.OnKmeansInput"
    Next i
    Q("#kmeansScale").AddEventListener "change", "opMining.OnKmeansScaleChange"
    Q("#loadKmeansSampleBtn").AddEventListener "click", "opMining.OnLoadKmeansSample"
    Q("#copyKmeansCodeBtn").AddEventListener "click", "opMining.OnCopyKmeansCode"

    ' PCA
    Set list = QA(".pca-input")
    For i = 0 To list.Count - 1
        list(i).AddEventListener "input", "opMining.OnPcaInput"
    Next i
    Q("#pcaModelType").AddEventListener "change", "opMining.OnPcaModelTypeChange"
    Q("#loadPcaSampleBtn").AddEventListener "click", "opMining.OnLoadPcaSample"
    Q("#copyPcaCodeBtn").AddEventListener "click", "opMining.OnCopyPcaCode"
    Q("#pcaCodeSections").AddEventListener "click", "opMining.OnPcaCodeSectionsClick"
    Q("#pcaSummaryOutput").AddEventListener "input", "opMining.OnPcaSummaryInput"
    Q("#pcaVarianceMode").AddEventListener "click", "opMining.OnPcaVarianceModeClick"
    Q("#loadPcaVarianceSampleBtn").AddEventListener "click", "opMining.OnLoadPcaVarianceSample"
    Q("#copyPcaVarianceBtn").AddEventListener "click", "opMining.OnCopyPcaVariance"
    Q("#pcaLoadingsOutput").AddEventListener "input", "opMining.OnPcaLoadingsInput"
    Q("#loadPcaLoadingsSampleBtn").AddEventListener "click", "opMining.OnLoadPcaLoadingsSample"
    Q("#copyPcaLoadingsBtn").AddEventListener "click", "opMining.OnCopyPcaLoadings"
    Q("#pcaLinearOutput").AddEventListener "input", "opMining.OnPcaLinearInput"
    Q("#loadPcaLinearSampleBtn").AddEventListener "click", "opMining.OnLoadPcaLinearSample"
    Q("#copyPcaLinearBtn").AddEventListener "click", "opMining.OnCopyPcaLinear"
    Q("#pcaLinearProse").AddEventListener "click", "opMining.OnPcaLinearProseClick"
    RenderPcaCode
    RenderPcaVariance
    RenderPcaLoadings
    RenderPcaLinearInterpretation

    ' K-fold cross-validation
    Set list = QA(".kfold-input")
    For i = 0 To list.Count - 1
        list(i).AddEventListener "input", "opMining.OnKfoldInput"
    Next i
    Q("#loadKfoldSampleBtn").AddEventListener "click", "opMining.OnLoadKfoldSample"
    Q("#copyKfoldCodeBtn").AddEventListener "click", "opMining.OnCopyKfoldCode"

    ' Stepwise selection
    Set list = QA(".stepwise-input")
    For i = 0 To list.Count - 1
        list(i).AddEventListener "input", "opMining.OnStepwiseInput"
    Next i
    Q("#stepwiseDirection").AddEventListener "click", "opMining.OnStepwiseDirectionClick"
    Q("#loadStepwiseSampleBtn").AddEventListener "click", "opMining.OnLoadStepwiseSample"
    Q("#copyStepwiseCodeBtn").AddEventListener "click", "opMining.OnCopyStepwiseCode"
    SyncStepwiseDirection
End Sub

' ---------------------------------------------------------------- small helpers

Private Function TrimmedValue(ByVal selector As String) As String
    TrimmedValue = JsTrim(Q(selector).Value)
End Function

Private Function JoinColumnNames(ByVal names As clsList, ByVal separator As String) As String
    Dim i As Long, out As String
    For i = 0 To names.Count - 1
        If i > 0 Then out = out & separator
        out = out & RColumnName(CStr(names(i)))
    Next i
    JoinColumnNames = out
End Function

' text.match(/number/gi).map(Number) (an empty list when nothing matches).
Private Function NumbersIn(ByVal text As String) As clsList
    Dim m As clsMatch, out As clsList, found As Collection
    Set out = NewList()
    Set found = Rx(NUMBER_PATTERN, "gi").MatchAll(text)
    For Each m In found
        out.Push JsNumber(m.Value)
    Next m
    Set NumbersIn = out
End Function

' The delegated-click target as an element (text-node targets use their parent).
Private Function EventElement(ByVal e As clsEvent) As clsNode
    Dim node As clsNode
    Set node = e.Target
    If Not node Is Nothing Then
        If node.NodeType <> 1 Then Set node = node.ParentElement
    End If
    Set EventElement = node
End Function

Private Sub CopyText(ByVal text As String, ByVal successMessage As String)
    If WriteAnswerText(text) Then ShowToast successMessage Else ShowToast "Clipboard access is unavailable"
End Sub

' ---------------------------------------------------------------- stepwise selection

Public Sub RenderStepwiseCode()
    Dim dataset As String, outcome As String, terms As clsList, forward As Boolean, modelName As String
    Dim ready As Boolean, minus As String, i As Long, comment As String
    dataset = TrimmedValue("#stepwiseDataset")
    If Len(dataset) = 0 Then dataset = "dataset"
    outcome = TrimmedValue("#stepwiseOutcome")
    Set terms = PredictorList(Q("#stepwiseExclude").Value)
    forward = (mStepwiseDirection = "forward")
    modelName = TrimmedValue("#stepwiseModelName")
    If Len(modelName) = 0 Then modelName = IIf(forward, "model", "fullmodel")
    ready = Len(outcome) > 0 And (Not forward Or terms.Count > 0)
    If Not ready Then
        If forward Then
            Q("#stepwiseCode").TextContent = "Enter an outcome variable and at least one candidate predictor to generate the forward-selection code."
        Else
            Q("#stepwiseCode").TextContent = "Enter an outcome variable to generate the backward-selection code."
        End If
    ElseIf forward Then
        Q("#stepwiseCode").TextContent = modelName & " <- lm(" & RColumnName(outcome) & " ~ 1, data = " & RColumnName(dataset) & ") # null model" & vbLf & _
            "newmodel <- step(" & modelName & ", scope = ~ " & JoinColumnNames(terms, " + ") & ", direction = ""forward"", trace = 1)" & vbLf & _
            "summary(newmodel)$adj.r.squared"
    Else
        For i = 0 To terms.Count - 1
            minus = minus & " - " & RColumnName(CStr(terms(i)))
        Next i
        If terms.Count > 0 Then comment = "all columns except " & terms.Join(", ") Else comment = "full model using all columns"
        Q("#stepwiseCode").TextContent = modelName & " <- lm(" & RColumnName(outcome) & " ~ ." & minus & ", data = " & RColumnName(dataset) & ") # " & comment & vbLf & _
            "newmodel <- step(" & modelName & ", direction = ""backward"", trace = 1)" & vbLf & _
            "summary(newmodel)$adj.r.squared"
    End If
    Q("#copyStepwiseCodeBtn").Disabled = Not ready
End Sub

Public Sub SyncStepwiseDirection()
    Dim forward As Boolean, list As clsList, i As Long, button As clsNode, active As Boolean
    forward = (mStepwiseDirection = "forward")
    Q("#stepwiseSubtitle").TextContent = IIf(forward, "forward addition", "backward elimination")
    Q("#stepwiseTermsLabel").TextContent = IIf(forward, "Candidate predictors", "Columns to exclude")
    Q("#stepwiseExclude").Placeholder = IIf(forward, "var1, var2", "ID")
    Q("#stepwiseModelLabel").TextContent = IIf(forward, "Null-model object name", "Full-model object name")
    Q("#stepwiseGuideTitle").TextContent = IIf(forward, "Forward selection", "Backward selection")
    If forward Then
        Q("#stepwiseGuide").TextContent = "The procedure starts with a null model and adds the candidate term that produces the best AIC improvement at each step, stopping when no further addition improves the model."
    Else
        Q("#stepwiseGuide").TextContent = "The procedure starts with every eligible predictor and removes the term that produces the best AIC improvement at each step, stopping when no further removal improves the model."
    End If
    Set list = QA("[data-stepwise-direction]")
    For i = 0 To list.Count - 1
        Set button = list(i)
        active = (CStr(button.DataGet("stepwiseDirection")) = mStepwiseDirection)
        button.ToggleClass "active", active
        button.SetAttribute "aria-pressed", IIf(active, "true", "false")
    Next i
    RenderStepwiseCode
End Sub

Public Sub OnStepwiseInput(ByVal e As clsEvent)
    On Error GoTo Failed
    RenderStepwiseCode
    Exit Sub
Failed:
    ReportError "opMining.OnStepwiseInput"
End Sub

Public Sub OnStepwiseDirectionClick(ByVal e As clsEvent)
    Dim button As clsNode, previousDirection As String, modelInput As clsNode, previousDefault As String
    On Error GoTo Failed
    Set button = EventElement(e)
    If button Is Nothing Then Exit Sub
    Set button = button.Closest("[data-stepwise-direction]")
    If button Is Nothing Then Exit Sub
    previousDirection = mStepwiseDirection
    mStepwiseDirection = CStr(button.DataGet("stepwiseDirection"))
    Set modelInput = Q("#stepwiseModelName")
    previousDefault = IIf(previousDirection = "forward", "model", "fullmodel")
    If JsTrim(modelInput.Value) = previousDefault Then modelInput.Value = IIf(mStepwiseDirection = "forward", "model", "fullmodel")
    SyncStepwiseDirection
    Exit Sub
Failed:
    ReportError "opMining.OnStepwiseDirectionClick"
End Sub

Public Sub OnLoadStepwiseSample(ByVal e As clsEvent)
    Dim forward As Boolean
    On Error GoTo Failed
    forward = (mStepwiseDirection = "forward")
    Q("#stepwiseDataset").Value = IIf(forward, "dataset", "df1")
    Q("#stepwiseOutcome").Value = IIf(forward, "dependent_variable", "Salary")
    Q("#stepwiseExclude").Value = IIf(forward, "var1, var2", "Years, CAtBat, CHits, CHmRun, CRuns, CRBI, CWalks")
    Q("#stepwiseModelName").Value = IIf(forward, "model", "fullmodel")
    RenderStepwiseCode
    ShowToast IIf(forward, "Forward", "Backward") & " sample loaded"
    Exit Sub
Failed:
    ReportError "opMining.OnLoadStepwiseSample"
End Sub

Public Sub OnCopyStepwiseCode(ByVal e As clsEvent)
    On Error GoTo Failed
    If Q("#copyStepwiseCodeBtn").Disabled Then Exit Sub
    CopyText Q("#stepwiseCode").TextContent, "R code copied"
    Exit Sub
Failed:
    ReportError "opMining.OnCopyStepwiseCode"
End Sub

' ---------------------------------------------------------------- K-fold cross-validation

Public Sub RenderKfoldCode()
    Dim dataset As String, outcome As String, predictors As clsList, k As Double, ready As Boolean, formula As String
    dataset = TrimmedValue("#kfoldDataset")
    If Len(dataset) = 0 Then dataset = "dataset"
    outcome = TrimmedValue("#kfoldOutcome")
    Set predictors = PredictorList(Q("#kfoldPredictors").Value)
    k = JsNumber(Q("#kfoldK").Value)
    If Len(outcome) > 0 And predictors.Count > 0 And JsIsInteger(k) Then ready = (k >= 2)
    If ready Then
        formula = RColumnName(outcome) & " ~ " & JoinColumnNames(predictors, " + ")
        Q("#kfoldCode").TextContent = "library(boot)" & vbLf & vbLf & _
            "fit_glm <- glm(" & formula & ", data = " & RColumnName(dataset) & ")" & vbLf & vbLf & _
            "cv.err <- cv.glm(" & RColumnName(dataset) & ", fit_glm, K = " & JsNumberToString(k) & ")" & vbLf & _
            "cv.err$delta # raw and bias-adjusted cross-validation error estimates"
    Else
        Q("#kfoldCode").TextContent = "Complete the dependent variable and predictors, and use an integer K of at least 2."
    End If
    Q("#copyKfoldCodeBtn").Disabled = Not ready
End Sub

Public Sub OnKfoldInput(ByVal e As clsEvent)
    On Error GoTo Failed
    RenderKfoldCode
    Exit Sub
Failed:
    ReportError "opMining.OnKfoldInput"
End Sub

Public Sub OnLoadKfoldSample(ByVal e As clsEvent)
    On Error GoTo Failed
    Q("#kfoldDataset").Value = "dataset"
    Q("#kfoldOutcome").Value = "column1"
    Q("#kfoldPredictors").Value = "column2, column3, column4"
    Q("#kfoldK").Value = "5"
    RenderKfoldCode
    ShowToast "Sample loaded"
    Exit Sub
Failed:
    ReportError "opMining.OnLoadKfoldSample"
End Sub

Public Sub OnCopyKfoldCode(ByVal e As clsEvent)
    On Error GoTo Failed
    If Q("#copyKfoldCodeBtn").Disabled Then Exit Sub
    CopyText Q("#kfoldCode").TextContent, "R code copied"
    Exit Sub
Failed:
    ReportError "opMining.OnCopyKfoldCode"
End Sub

' ---------------------------------------------------------------- PCA samples

Public Function PcaVarianceSample() As String
    PcaVarianceSample = "[1] ""sdev""     ""rotation"" ""center""   ""scale""    ""x""        ""call""" & vbLf & _
        "Importance of components:" & vbLf & _
        "                          PC1    PC2     PC3     PC4" & vbLf & _
        "Standard deviation     1.5749 0.9949 0.59713 0.41645" & vbLf & _
        "Proportion of Variance 0.6201 0.2474 0.08914 0.04336" & vbLf & _
        "Cumulative Proportion  0.6201 0.8675 0.95664 1.00000"
End Function

Public Function PcaLoadingsSample() As String
    PcaLoadingsSample = "                PC1        PC2        PC3         PC4" & vbLf & _
        "Murder   -0.5358995 -0.4181809  0.3412327  0.64922780" & vbLf & _
        "Assault  -0.5831836 -0.1879856  0.2681484 -0.74340748" & vbLf & _
        "UrbanPop -0.2781909  0.8728062  0.3780158  0.13387773" & vbLf & _
        "Rape     -0.5434321  0.1673186 -0.8177779  0.08902432"
End Function

Public Function PcaLinearSample() As String
    Dim lq As String, rq As String
    lq = ChrW$(&H2018)
    rq = ChrW$(&H2019)
    PcaLinearSample = "Call:" & vbLf & _
        "lm(formula = GDP ~ pc1 + pc2, data = USArrests)" & vbLf & vbLf & _
        "Residuals:" & vbLf & _
        "       Min         1Q     Median         3Q        Max" & vbLf & _
        "-2.502e+11 -6.607e+10 -1.397e+10  4.653e+10  4.578e+11" & vbLf & vbLf & _
        "Coefficients:" & vbLf & _
        "              Estimate Std. Error t value Pr(>|t|)" & vbLf & _
        "(Intercept)  1.538e+11  1.854e+10   8.294 9.32e-11 ***" & vbLf & _
        "pc1         -4.344e+10  1.189e+10  -3.653 0.000652 ***" & vbLf & _
        "pc2          4.310e+10  1.883e+10   2.289 0.026592 *" & vbLf & _
        "---" & vbLf & _
        "Signif. codes:  0 " & lq & "***" & rq & " 0.001 " & lq & "**" & rq & " 0.01 " & lq & "*" & rq & " 0.05 " & lq & "." & rq & " 0.1 " & lq & " " & rq & " 1" & vbLf & vbLf & _
        "Residual standard error: 1.311e+11 on 47 degrees of freedom" & vbLf & _
        "Multiple R-squared:  0.2834, Adjusted R-squared:  0.2529" & vbLf & _
        "F-statistic: 9.292 on 2 and 47 DF, p-value: 0.0003978"
End Function

' ---------------------------------------------------------------- PCA code

Private Function NewSection(ByVal title As String, ByVal code As String) As clsMap
    Dim section As clsMap
    Set section = NewMap()
    section("title") = title
    section("code") = code
    Set NewSection = section
End Function

Public Sub RenderPcaCode()
    Dim dataset As String, variables As clsList, count As Double, modelType As String, outcome As String, threshold As Double
    Dim values As clsList, buttons As clsList, needsModel As Boolean, ready As Boolean, i As Long
    Dim data As String, selected As String, assignments As String, rhs As String, pcName As String, all As String
    dataset = TrimmedValue("#pcaDataset")
    If Len(dataset) = 0 Then dataset = "dataset"
    Set variables = PredictorList(Q("#pcaVariables").Value)
    count = JsNumber(Q("#pcaCount").Value)
    modelType = Q("#pcaModelType").Value
    outcome = TrimmedValue("#pcaOutcome")
    threshold = JsNumber(Q("#pcaThreshold").Value)
    Set values = QA("[data-pca-code-value]")
    Set buttons = QA("[data-copy-pca-section]")
    needsModel = (modelType <> "none")
    Q("#pcaOutcomeField").ToggleClass "hidden", Not needsModel
    Q("#pcaThresholdField").ToggleClass "hidden", modelType <> "binomial"
    Q("#pcaDownstreamCodeBlock").ToggleClass "hidden", Not needsModel
    If variables.Count >= 2 And JsIsInteger(count) Then
        If count >= 1 And count <= variables.Count And (Not needsModel Or Len(outcome) > 0) Then
            If modelType <> "binomial" Then
                ready = True
            Else
                ready = JsGe(threshold, 0) And JsLe(threshold, 1)
            End If
        End If
    End If
    If Not ready Then
        Set mPcaCodeSections = NewList()
        For i = 0 To values.Count - 1
            values(i).TextContent = IIf(i = 0, "Enter at least two PCA variables, choose a valid retained-component count, and complete any downstream model fields.", "")
        Next i
        For i = 0 To buttons.Count - 1
            buttons(i).Disabled = True
        Next i
        Q("#pcaCode").TextContent = ""
        Q("#copyPcaCodeBtn").Disabled = True
        Exit Sub
    End If
    data = RColumnName(dataset)
    For i = 0 To variables.Count - 1
        If i > 0 Then selected = selected & ", "
        selected = selected & """" & RString(CStr(variables(i))) & """"
    Next i
    For i = 1 To CLng(count)
        pcName = "PC" & CStr(i)
        If i > 1 Then
            assignments = assignments & vbLf
            rhs = rhs & " + "
        End If
        assignments = assignments & data & "$" & pcName & " <- pca_model$x[, """ & pcName & """]"
        rhs = rhs & pcName
    Next i
    Set mPcaCodeSections = NewList()
    mPcaCodeSections.Push NewSection("PCA data selection", "pca_data <- " & data & "[, c(" & selected & ")]")
    mPcaCodeSections.Push NewSection("PCA model", "pca_model <- prcomp(" & vbLf & "  pca_data," & vbLf & "  center = TRUE," & vbLf & "  scale. = TRUE" & vbLf & ")")
    mPcaCodeSections.Push NewSection("PCA summary", "summary(pca_model)")
    mPcaCodeSections.Push NewSection("PC loadings", "# Loadings for the first two components" & vbLf & "pca_model$rotation[, 1:2]")
    mPcaCodeSections.Push NewSection("PC scores", "# Add PC scores" & vbLf & assignments)
    mPcaCodeSections.Push NewSection("Scree plot", "# Scree plot" & vbLf & "library(factoextra)" & vbLf & "fviz_eig(pca_model)")
    mPcaCodeSections.Push NewSection("Biplot", "# Biplot" & vbLf & "biplot(pca_model, cex = c(0.5, 0.75))")
    If modelType = "linear" Then
        mPcaCodeSections.Push NewSection("Downstream linear model", "# Fit a linear regression model using retained components" & vbLf & _
            "model <- lm(" & RColumnName(outcome) & " ~ " & rhs & ", data = " & data & ")" & vbLf & "summary(model)")
    End If
    If modelType = "binomial" Then
        mPcaCodeSections.Push NewSection("Downstream binomial model", "# Check statistical significance of retained PCs with a binomial GLM" & vbLf & _
            "pc_model <- glm(" & RColumnName(outcome) & " ~ " & rhs & ", data = " & data & ", family = ""binomial"")" & vbLf & _
            "summary(pc_model)" & vbLf & vbLf & "# Predict classes" & vbLf & _
            "prob <- predict(pc_model, type = ""response"")" & vbLf & _
            "pred <- ifelse(prob >= " & JsNumberToString(threshold) & ", 1, 0)" & vbLf & _
            "sum(pred == 1) # Yes" & vbLf & "sum(pred == 0) # No")
    End If
    For i = 0 To values.Count - 1
        If i < mPcaCodeSections.Count Then values(i).TextContent = mPcaCodeSections(i)("code") Else values(i).TextContent = ""
    Next i
    For i = 0 To buttons.Count - 1
        buttons(i).Disabled = Not (i < mPcaCodeSections.Count)
    Next i
    For i = 0 To mPcaCodeSections.Count - 1
        If i > 0 Then all = all & vbLf & vbLf
        all = all & mPcaCodeSections(i)("code")
    Next i
    Q("#pcaCode").TextContent = all
    Q("#copyPcaCodeBtn").Disabled = False
End Sub

Public Function PcaOrdinal(ByVal number As Double) As String
    Dim remainder100 As Double, remainder10 As Double, text As String
    text = JsNumberToString(number)
    remainder100 = JsMod(number, 100)
    If remainder100 >= 11 And remainder100 <= 13 Then
        PcaOrdinal = text & "th"
        Exit Function
    End If
    remainder10 = JsMod(number, 10)
    If remainder10 = 1 Then
        PcaOrdinal = text & "st"
    ElseIf remainder10 = 2 Then
        PcaOrdinal = text & "nd"
    ElseIf remainder10 = 3 Then
        PcaOrdinal = text & "rd"
    Else
        PcaOrdinal = text & "th"
    End If
End Function

' ---------------------------------------------------------------- PCA variance explained

Private Sub SetVarianceMessage(ByVal message As String, Optional ByVal isError As Boolean = False)
    Dim prose As clsNode
    mPcaVarianceProse = ""
    Set prose = Q("#pcaVarianceProse")
    prose.TextContent = message
    prose.ToggleClass "is-error", isError
    Q("#copyPcaVarianceBtn").Disabled = True
End Sub

Public Sub RenderPcaVariance()
    Dim output As String, retained As Double, prose As clsNode, row As clsMatch, rowName As String
    Dim all As clsList, values As clsList, i As Long, n As Long, text As String
    output = Replace(Q("#pcaSummaryOutput").Value, vbCr, "")
    retained = JsNumber(Q("#pcaCount").Value)
    Set prose = Q("#pcaVarianceProse")
    If Not JsIsInteger(retained) Then
        SetVarianceMessage "Enter a whole number of components to retain.", True
        Exit Sub
    End If
    If retained < 1 Then
        SetVarianceMessage "Enter a whole number of components to retain.", True
        Exit Sub
    End If
    If Len(JsTrim(output)) = 0 Then
        SetVarianceMessage "Paste the PCA summary output to describe the retained components."
        Exit Sub
    End If
    If mPcaVarianceMode = "cumulative" Then
        Set row = Rx("Cumulative[ \t]*Proportion[ \t]+([^\n]+)", "i").Exec(output)
        rowName = "Cumulative Proportion"
    Else
        Set row = Rx("Proportion[ \t]*of[ \t]*Variance[ \t]+([^\n]+)", "i").Exec(output)
        rowName = "Proportion of Variance"
    End If
    If row Is Nothing Then
        SetVarianceMessage "I could not find the " & ChrW$(&H2018) & rowName & ChrW$(&H2019) & " row in the pasted output.", True
        Exit Sub
    End If
    Set all = NumbersIn(row.Group(1))
    Set values = NewList()
    For i = 0 To all.Count - 1
        If JsIsFinite(CDbl(all(i))) Then values.Push all(i)
    Next i
    n = values.Count
    If n < retained Then
        SetVarianceMessage "The output contains " & CStr(n) & " variance value" & IIf(n = 1, "", "s") & ", but " & JsNumberToString(retained) & " components were selected.", True
        Exit Sub
    End If
    If mPcaVarianceMode = "cumulative" Then
        If retained = 1 Then
            text = "The 1st principal component explains about " & PercentText(values(0)) & "% of the total variance."
        Else
            text = "Together, the first " & JsNumberToString(retained) & " principal components explain about " & PercentText(values(CLng(retained) - 1)) & "% of the total variance."
        End If
    Else
        For i = 0 To CLng(retained) - 1
            If i > 0 Then text = text & vbLf
            text = text & "The " & PcaOrdinal(i + 1) & " principal component explains about " & PercentText(values(i)) & "% of the proportion of variance."
        Next i
    End If
    mPcaVarianceProse = text
    prose.RemoveClass "is-error"
    prose.TextContent = mPcaVarianceProse
    Q("#copyPcaVarianceBtn").Disabled = False
End Sub

' (value * 100).toFixed(1), with the product rounded to a Double first.
Private Function PercentText(ByVal value As Double) As String
    Dim percent As Double
    percent = value * 100
    PercentText = JsToFixed(percent, 1)
End Function

' ---------------------------------------------------------------- PCA loading equations

Public Function FormatPcaLoading(ByVal value As Double) As String
    Dim absolute As Double
    absolute = Abs(value)
    If absolute = 0 Then
        FormatPcaLoading = "0"
    ElseIf absolute >= 1 Then
        FormatPcaLoading = JsToFixed(absolute, 2)
    Else
        FormatPcaLoading = JsToPrecision(absolute, 3)
    End If
End Function

Private Sub SetLoadingsMessage(ByVal message As String, Optional ByVal isError As Boolean = False)
    Dim prose As clsNode
    mPcaLoadingsProse = ""
    Set prose = Q("#pcaLoadingsProse")
    prose.TextContent = message
    prose.ToggleClass "is-error", isError
    Q("#copyPcaLoadingsBtn").Disabled = True
End Sub

Public Sub RenderPcaLoadings()
    Dim output As String, retained As Double, prose As clsNode, rawLines As clsList, lines As clsList, i As Long, j As Long
    Dim headerIndex As Long, components As clsList, m As clsMatch, nComponents As Long, rows As clsList, rowNames As clsList
    Dim match As clsMatch, values As clsList, ok As Boolean, name As String, equations As String, terms As String
    Dim value As Double, term As String, rowValues As clsList, found As Collection
    output = Replace(Q("#pcaLoadingsOutput").Value, vbCr, "")
    retained = JsNumber(Q("#pcaCount").Value)
    Set prose = Q("#pcaLoadingsProse")
    If Not JsIsInteger(retained) Then
        SetLoadingsMessage "Enter a whole number of components to retain.", True
        Exit Sub
    End If
    If retained < 1 Then
        SetLoadingsMessage "Enter a whole number of components to retain.", True
        Exit Sub
    End If
    If Len(JsTrim(output)) = 0 Then
        SetLoadingsMessage "Paste the loading matrix to write an equation for each retained component."
        Exit Sub
    End If
    Set rawLines = JsSplit(output, vbLf)
    Set lines = NewList()
    For i = 0 To rawLines.Count - 1
        If Len(JsTrim(CStr(rawLines(i)))) > 0 Then lines.Push rawLines(i)
    Next i
    headerIndex = -1
    For i = 0 To lines.Count - 1
        If Rx("\bPC\s*1\b", "i").Test(CStr(lines(i))) Then
            headerIndex = i
            Exit For
        End If
    Next i
    If headerIndex < 0 Then
        SetLoadingsMessage "I could not find a header containing PC1 in the pasted loading matrix.", True
        Exit Sub
    End If
    Set components = NewList()
    Set found = Rx("PC\s*\d+", "gi").MatchAll(CStr(lines(headerIndex)))
    For Each m In found
        components.Push JsUpper(RxReplace(m.Value, "\s", "", "g"))
    Next m
    nComponents = components.Count
    If nComponents < retained Then
        SetLoadingsMessage "The matrix contains " & CStr(nComponents) & " component" & IIf(nComponents = 1, "", "s") & ", but " & JsNumberToString(retained) & " components were selected.", True
        Exit Sub
    End If
    Set rows = NewList()
    Set rowNames = NewList()
    For i = headerIndex + 1 To lines.Count - 1
        Set match = Rx("^\s*(?:""([^""]+)""|'([^']+)'|(\S+))\s+(.+)$").Exec(CStr(lines(i)))
        If Not match Is Nothing Then
            Set values = NumbersIn(match.Group(4))
            ok = (values.Count >= nComponents)
            If ok Then
                For j = 0 To nComponents - 1
                    If Not JsIsFinite(CDbl(values(j))) Then ok = False
                Next j
            End If
            If ok Then
                If match.Matched(1) Then
                    name = match.Group(1)
                ElseIf match.Matched(2) Then
                    name = match.Group(2)
                Else
                    name = match.Group(3)
                End If
                rowNames.Push name
                rows.Push values
            End If
        End If
    Next i
    If rows.Count = 0 Then
        SetLoadingsMessage "I found the component headings but could not read any loading rows.", True
        Exit Sub
    End If
    For j = 0 To CLng(retained) - 1
        terms = ""
        For i = 0 To rows.Count - 1
            Set rowValues = rows(i)
            value = rowValues(j)
            term = FormatPcaLoading(value) & " " & ChrW$(&HD7) & " " & CStr(rowNames(i))
            If i = 0 Then
                If value < 0 Then terms = terms & "-" & term Else terms = terms & term
            Else
                If value < 0 Then terms = terms & " - " & term Else terms = terms & " + " & term
            End If
        Next i
        If j > 0 Then equations = equations & vbLf
        equations = equations & CStr(components(j)) & " = " & terms
    Next j
    mPcaLoadingsProse = equations
    prose.RemoveClass "is-error"
    prose.TextContent = mPcaLoadingsProse
    Q("#copyPcaLoadingsBtn").Disabled = False
End Sub

' ---------------------------------------------------------------- PCA downstream linear model

Public Function FormatPcaModelNumber(ByVal value As Double) As String
    Dim absolute As Double
    absolute = Abs(value)
    If absolute = 0 Then
        FormatPcaModelNumber = "0"
    ElseIf absolute >= 1 Then
        FormatPcaModelNumber = JsToFixed(value, 2)
    Else
        FormatPcaModelNumber = JsToPrecision(value, 3)
    End If
End Function

Private Sub SetLinearMessage(ByVal message As String, Optional ByVal isError As Boolean = False)
    Dim box As clsNode, container As clsNode
    Set mPcaLinearProseSections = NewList()
    Set box = CreateElement("div")
    box.ClassName = "pca-variance-prose" & IIf(isError, " is-error", "")
    box.TextContent = message
    Set container = Q("#pcaLinearProse")
    container.ReplaceChildren box
    Q("#copyPcaLinearBtn").Disabled = True
End Sub

Public Sub RenderPcaLinearInterpretation()
    Dim output As String, retained As Double, container As clsNode, formula As clsMatch, dependent As String
    Dim startMatch As clsMatch, coefficientText As String, endMatch As clsMatch, lines As clsList, i As Long, j As Long
    Dim match As clsMatch, numbers As clsList, ok As Boolean, rows As clsList, row As clsMap, opMatch As clsMatch
    Dim selected As clsList, significant As Variant, section As clsMap, text As String, index As Double
    Dim card As clsNode, head As clsNode, title As clsNode, button As clsNode, paragraph As clsNode
    output = Replace(Q("#pcaLinearOutput").Value, vbCr, "")
    retained = JsNumber(Q("#pcaCount").Value)
    Set container = Q("#pcaLinearProse")
    If Not JsIsInteger(retained) Then
        SetLinearMessage "Enter a whole number of components to retain.", True
        Exit Sub
    End If
    If retained < 1 Then
        SetLinearMessage "Enter a whole number of components to retain.", True
        Exit Sub
    End If
    If Len(JsTrim(output)) = 0 Then
        SetLinearMessage "Paste the linear-regression summary to interpret each retained principal component."
        Exit Sub
    End If
    Set formula = Rx("lm\s*\(\s*formula\s*=\s*([^~\n,]+?)\s*~", "i").Exec(output)
    If formula Is Nothing Then
        SetLinearMessage "I could not identify the dependent variable from the lm() call.", True
        Exit Sub
    End If
    dependent = RxReplace(JsTrim(formula.Group(1)), "^`|`$", "", "g")
    Set startMatch = Rx("Coefficients\s*:", "i").Exec(output)
    If startMatch Is Nothing Then
        SetLinearMessage "I could not find the Coefficients section.", True
        Exit Sub
    End If
    coefficientText = JsSlice(output, startMatch.Index)
    Set endMatch = Rx("\n---|\nSignif\.\s*codes", "i").Exec(coefficientText)
    If Not endMatch Is Nothing Then coefficientText = JsSlice(coefficientText, 0, endMatch.Index)
    Set rows = NewList()
    Set lines = JsSplit(coefficientText, vbLf)
    For i = 0 To lines.Count - 1
        Set match = Rx("^\s*`?(PC\s*(\d+))`?\s+(.+)$", "i").Exec(CStr(lines(i)))
        If Not match Is Nothing Then
            Set numbers = NumbersIn(match.Group(3))
            ok = (numbers.Count >= 4)
            If ok Then
                For j = 0 To 3
                    If Not JsIsFinite(CDbl(numbers(j))) Then ok = False
                Next j
            End If
            If ok Then
                Set row = NewMap()
                index = JsNumber(match.Group(2))
                row("component") = "PC" & JsNumberToString(index)
                row("index") = index
                row("t") = CDbl(numbers(2))
                row("p") = CDbl(numbers(3))
                Set opMatch = Rx("(<=|>=|<|>)\s*[-+]?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?\s*[* .]*$", "i").Exec(match.Group(3))
                If opMatch Is Nothing Then row("pOperator") = "=" Else row("pOperator") = opMatch.Group(1)
                InsertRowByIndex rows, row
            End If
        End If
    Next i
    Set selected = NewList()
    For i = 0 To rows.Count - 1
        Set row = rows(i)
        If row("index") <= retained Then selected.Push row
    Next i
    If selected.Count = 0 Then
        SetLinearMessage "I found the coefficient table but could not identify any PC coefficient rows.", True
        Exit Sub
    End If
    For i = 0 To selected.Count - 1
        Set row = selected(i)
        If IsNull(PValueDecision(row("pOperator"), row("p"))) Then
            SetLinearMessage "A component p-value does not establish significance at 0.05. Paste a more precise p-value.", True
            Exit Sub
        End If
    Next i
    Set mPcaLinearProseSections = NewList()
    For i = 0 To selected.Count - 1
        Set row = selected(i)
        significant = PValueDecision(row("pOperator"), row("p"))
        text = "For " & row("component") & ", the t-value = " & FormatPcaModelNumber(row("t")) & " and p-value " & row("pOperator") & " " & _
            FormatPValue(row("p")) & ", which is " & PValueComparison(row("pOperator"), row("p")) & " 0.05, thus it is " & _
            IIf(significant, "significant", "not significant") & " in explaining the dependent variable (" & dependent & _
            ") in this regression model at the 5% level of significance."
        Set section = NewMap()
        section("component") = row("component")
        section("text") = text
        mPcaLinearProseSections.Push section
    Next i
    container.ReplaceChildren
    For i = 0 To mPcaLinearProseSections.Count - 1
        Set section = mPcaLinearProseSections(i)
        Set card = CreateElement("section")
        card.ClassName = "sensitivity-prose-card"
        Set head = CreateElement("div")
        head.ClassName = "sensitivity-prose-head"
        Set title = CreateElement("h3")
        title.TextContent = section("component")
        Set button = CreateElement("button")
        button.TypeName = "button"
        button.ClassName = "sample-btn"
        button.DataSet "copyPcaLinear", CStr(i)
        button.TextContent = "Copy"
        Set paragraph = CreateElement("p")
        paragraph.TextContent = section("text")
        head.Append title, button
        card.Append head, paragraph
        container.AppendChild card
    Next i
    Q("#copyPcaLinearBtn").Disabled = False
End Sub

' Stable insertion by row("index") (the website's rows.sort((a, b) => a.index - b.index)).
Private Sub InsertRowByIndex(ByVal rows As clsList, ByVal row As clsMap)
    Dim position As Long, other As clsMap
    position = rows.Count
    Do While position > 0
        Set other = rows(position - 1)
        If other("index") <= row("index") Then Exit Do
        position = position - 1
    Loop
    If position = rows.Count Then rows.Push row Else rows.Splice position, 0, row
End Sub

' ---------------------------------------------------------------- PCA events

Private Sub RenderAllPca()
    RenderPcaCode
    RenderPcaVariance
    RenderPcaLoadings
    RenderPcaLinearInterpretation
End Sub

Public Sub OnPcaInput(ByVal e As clsEvent)
    On Error GoTo Failed
    RenderAllPca
    Exit Sub
Failed:
    ReportError "opMining.OnPcaInput"
End Sub

Public Sub OnPcaModelTypeChange(ByVal e As clsEvent)
    On Error GoTo Failed
    RenderPcaCode
    Exit Sub
Failed:
    ReportError "opMining.OnPcaModelTypeChange"
End Sub

Public Sub OnLoadPcaSample(ByVal e As clsEvent)
    On Error GoTo Failed
    Q("#pcaDataset").Value = "USArrests"
    Q("#pcaVariables").Value = "Murder, Assault, UrbanPop, Rape"
    Q("#pcaCount").Value = "2"
    Q("#pcaModelType").Value = "none"
    Q("#pcaOutcome").Value = ""
    Q("#pcaThreshold").Value = "0.5"
    RenderAllPca
    ShowToast "PCA sample loaded"
    Exit Sub
Failed:
    ReportError "opMining.OnLoadPcaSample"
End Sub

Public Sub OnCopyPcaCode(ByVal e As clsEvent)
    On Error GoTo Failed
    If Q("#copyPcaCodeBtn").Disabled Then Exit Sub
    CopyText Q("#pcaCode").TextContent, "All PCA code copied"
    Exit Sub
Failed:
    ReportError "opMining.OnCopyPcaCode"
End Sub

Public Sub OnPcaCodeSectionsClick(ByVal e As clsEvent)
    Dim button As clsNode, index As Double, section As clsMap
    On Error GoTo Failed
    Set button = EventElement(e)
    If button Is Nothing Then Exit Sub
    Set button = button.Closest("[data-copy-pca-section]")
    If button Is Nothing Then Exit Sub
    If button.Disabled Then Exit Sub
    index = JsNumber(CStr(button.DataGet("copyPcaSection")))
    If Not JsIsInteger(index) Then Exit Sub
    If index < 0 Or index >= mPcaCodeSections.Count Then Exit Sub
    Set section = mPcaCodeSections(CLng(index))
    CopyText section("code"), section("title") & " copied"
    Exit Sub
Failed:
    ReportError "opMining.OnPcaCodeSectionsClick"
End Sub

Public Sub OnPcaSummaryInput(ByVal e As clsEvent)
    On Error GoTo Failed
    RenderPcaVariance
    Exit Sub
Failed:
    ReportError "opMining.OnPcaSummaryInput"
End Sub

Public Sub OnPcaVarianceModeClick(ByVal e As clsEvent)
    Dim button As clsNode, list As clsList, i As Long, opt As clsNode, active As Boolean
    On Error GoTo Failed
    Set button = EventElement(e)
    If button Is Nothing Then Exit Sub
    Set button = button.Closest("[data-pca-variance-mode]")
    If button Is Nothing Then Exit Sub
    mPcaVarianceMode = CStr(button.DataGet("pcaVarianceMode"))
    Set list = QA("[data-pca-variance-mode]")
    For i = 0 To list.Count - 1
        Set opt = list(i)
        active = (opt Is button)
        opt.ToggleClass "active", active
        opt.SetAttribute "aria-pressed", IIf(active, "true", "false")
    Next i
    RenderPcaVariance
    Exit Sub
Failed:
    ReportError "opMining.OnPcaVarianceModeClick"
End Sub

Public Sub OnLoadPcaVarianceSample(ByVal e As clsEvent)
    On Error GoTo Failed
    Q("#pcaSummaryOutput").Value = PcaVarianceSample()
    Q("#pcaCount").Value = "2"
    RenderPcaVariance
    ShowToast "PCA output sample loaded"
    Exit Sub
Failed:
    ReportError "opMining.OnLoadPcaVarianceSample"
End Sub

Public Sub OnCopyPcaVariance(ByVal e As clsEvent)
    On Error GoTo Failed
    If Len(mPcaVarianceProse) = 0 Then Exit Sub
    CopyText mPcaVarianceProse, "Variance interpretation copied"
    Exit Sub
Failed:
    ReportError "opMining.OnCopyPcaVariance"
End Sub

Public Sub OnPcaLoadingsInput(ByVal e As clsEvent)
    On Error GoTo Failed
    RenderPcaLoadings
    Exit Sub
Failed:
    ReportError "opMining.OnPcaLoadingsInput"
End Sub

Public Sub OnLoadPcaLoadingsSample(ByVal e As clsEvent)
    On Error GoTo Failed
    Q("#pcaLoadingsOutput").Value = PcaLoadingsSample()
    Q("#pcaCount").Value = "2"
    RenderPcaLoadings
    ShowToast "PCA loadings sample loaded"
    Exit Sub
Failed:
    ReportError "opMining.OnLoadPcaLoadingsSample"
End Sub

Public Sub OnCopyPcaLoadings(ByVal e As clsEvent)
    On Error GoTo Failed
    If Len(mPcaLoadingsProse) = 0 Then Exit Sub
    CopyText mPcaLoadingsProse, "Loading equations copied"
    Exit Sub
Failed:
    ReportError "opMining.OnCopyPcaLoadings"
End Sub

Public Sub OnPcaLinearInput(ByVal e As clsEvent)
    On Error GoTo Failed
    RenderPcaLinearInterpretation
    Exit Sub
Failed:
    ReportError "opMining.OnPcaLinearInput"
End Sub

Public Sub OnLoadPcaLinearSample(ByVal e As clsEvent)
    On Error GoTo Failed
    Q("#pcaLinearOutput").Value = PcaLinearSample()
    Q("#pcaCount").Value = "2"
    RenderPcaLinearInterpretation
    ShowToast "Linear-model sample loaded"
    Exit Sub
Failed:
    ReportError "opMining.OnLoadPcaLinearSample"
End Sub

Public Sub OnCopyPcaLinear(ByVal e As clsEvent)
    Dim i As Long, text As String, section As clsMap
    On Error GoTo Failed
    If mPcaLinearProseSections.Count = 0 Then Exit Sub
    For i = 0 To mPcaLinearProseSections.Count - 1
        Set section = mPcaLinearProseSections(i)
        If i > 0 Then text = text & vbLf & vbLf
        text = text & section("text")
    Next i
    CopyText text, "All PC interpretations copied"
    Exit Sub
Failed:
    ReportError "opMining.OnCopyPcaLinear"
End Sub

Public Sub OnPcaLinearProseClick(ByVal e As clsEvent)
    Dim button As clsNode, index As Double, section As clsMap
    On Error GoTo Failed
    Set button = EventElement(e)
    If button Is Nothing Then Exit Sub
    Set button = button.Closest("[data-copy-pca-linear]")
    If button Is Nothing Then Exit Sub
    index = JsNumber(CStr(button.DataGet("copyPcaLinear")))
    If Not JsIsInteger(index) Then Exit Sub
    If index < 0 Or index >= mPcaLinearProseSections.Count Then Exit Sub
    Set section = mPcaLinearProseSections(CLng(index))
    CopyText section("text"), section("component") & " interpretation copied"
    Exit Sub
Failed:
    ReportError "opMining.OnPcaLinearProseClick"
End Sub

' ---------------------------------------------------------------- K-means

Public Sub RenderKmeansCode()
    Dim dataset As String, features As clsList, maxK As Double, k As Double, seed As Double, nstart As Double
    Dim scaleVariables As Boolean, ready As Boolean, data As String, selected As String, clusterData As String
    Dim scalingCode As String, code As String, codeNode As clsNode
    dataset = TrimmedValue("#kmeansDataset")
    If Len(dataset) = 0 Then dataset = "dataset"
    Set features = PredictorList(Q("#kmeansFeatures").Value)
    maxK = JsNumber(Q("#kmeansMaxK").Value)
    k = JsNumber(Q("#kmeansK").Value)
    seed = JsNumber(Q("#kmeansSeed").Value)
    nstart = JsNumber(Q("#kmeansNstart").Value)
    scaleVariables = Q("#kmeansScale").Checked
    Set codeNode = Q("#kmeansCode")
    If features.Count >= 2 And JsIsInteger(maxK) And JsIsInteger(k) And JsIsInteger(seed) And JsIsInteger(nstart) Then
        ready = (maxK >= 2 And k >= 1 And k <= maxK And nstart >= 1)
    End If
    If Not ready Then
        codeNode.TextContent = "Enter at least two features and valid integer settings; chosen K must not exceed maximum K."
        Q("#copyKmeansCodeBtn").Disabled = True
        Exit Sub
    End If
    data = RColumnName(dataset)
    selected = JoinColumnNames(features, ", ")
    If scaleVariables Then
        clusterData = "X_scaled"
        scalingCode = vbLf & vbLf & "# Standardize variables" & vbLf & "X_scaled <- scale(X)"
    Else
        clusterData = "X"
    End If
    code = "library(psych)" & vbLf & "library(factoextra)" & vbLf & vbLf & _
        "# Select features" & vbLf & "X <- " & data & " %>% dplyr::select(" & selected & ")" & vbLf & vbLf & _
        "# View pairwise correlation coefficients" & vbLf & "pairs.panels(X)" & scalingCode & vbLf & vbLf & _
        "# Elbow method" & vbLf & "set.seed(" & JsNumberToString(seed) & ")" & vbLf & _
        "wss <- rep(NA, " & JsNumberToString(maxK) & ")" & vbLf & _
        "for (i in 1:" & JsNumberToString(maxK) & ") {" & vbLf & _
        "  wss[i] <- kmeans(" & clusterData & ", centers = i, nstart = " & JsNumberToString(nstart) & ")$tot.withinss" & vbLf & _
        "}" & vbLf & _
        BoundedTickCode("x", "seq_along(wss)", False, True) & vbLf & _
        BoundedTickCode("y", "wss") & vbLf & _
        "plot(wss, type = ""b"", " & NumericAxisArguments("x", "x") & " " & NumericAxisArguments("y", "y") & " xlab = ""Number of Clusters""," & vbLf & _
        "     ylab = ""Total Within-Cluster Sum Squares""," & vbLf & _
        "     main = ""Elbow Plot for K-Means Clustering"")" & vbLf & _
        NumericAxisCode() & vbLf & _
        NumericAxisCode() & vbLf & vbLf
    code = code & "# Run K-means using the chosen K" & vbLf & "set.seed(" & JsNumberToString(seed) & ")" & vbLf & _
        "km_obj <- kmeans(" & clusterData & ", centers = " & JsNumberToString(k) & ", nstart = " & JsNumberToString(nstart) & ")" & vbLf & _
        "fviz_cluster(km_obj, data = " & clusterData & ")" & vbLf & vbLf & _
        "# Add cluster labels" & vbLf & data & " <- " & data & " %>% dplyr::mutate(cluster = km_obj$cluster)" & vbLf & vbLf & _
        "# Interpret cluster profiles" & vbLf & data & " %>%" & vbLf & _
        "  dplyr::group_by(cluster) %>%" & vbLf & _
        "  dplyr::summarise(dplyr::across(where(is.numeric), ~ mean(.x, na.rm = TRUE)))"
    codeNode.TextContent = code
    codeNode.TextContent = "library(dplyr)" & vbLf & codeNode.TextContent
    Q("#copyKmeansCodeBtn").Disabled = False
End Sub

Public Sub OnKmeansInput(ByVal e As clsEvent)
    On Error GoTo Failed
    RenderKmeansCode
    Exit Sub
Failed:
    ReportError "opMining.OnKmeansInput"
End Sub

Public Sub OnKmeansScaleChange(ByVal e As clsEvent)
    Dim box As clsNode
    On Error GoTo Failed
    Set box = e.CurrentTarget
    If Not box.Checked Then
        box.Checked = True
        CallOp KmeansScaleAlertProc
        Exit Sub
    End If
    RenderKmeansCode
    Exit Sub
Failed:
    ReportError "opMining.OnKmeansScaleChange"
End Sub

Public Sub OnLoadKmeansSample(ByVal e As clsEvent)
    On Error GoTo Failed
    Q("#kmeansDataset").Value = "dataset"
    Q("#kmeansFeatures").Value = "column1, column2, column3"
    Q("#kmeansMaxK").Value = "10"
    Q("#kmeansK").Value = "3"
    Q("#kmeansSeed").Value = "1"
    Q("#kmeansNstart").Value = "10"
    Q("#kmeansScale").Checked = True
    RenderKmeansCode
    ShowToast "Sample loaded"
    Exit Sub
Failed:
    ReportError "opMining.OnLoadKmeansSample"
End Sub

Public Sub OnCopyKmeansCode(ByVal e As clsEvent)
    On Error GoTo Failed
    If Q("#copyKmeansCodeBtn").Disabled Then Exit Sub
    CopyText Q("#kmeansCode").TextContent, "R code copied"
    Exit Sub
Failed:
    ReportError "opMining.OnCopyKmeansCode"
End Sub
