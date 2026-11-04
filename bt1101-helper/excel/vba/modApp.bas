Attribute VB_Name = "modApp"
' The application shell (website app-src.js lines 1-70, 963-1590, 3889-4258, 4743-end):
' operation groups and samples, local storage, the update check, warning modals and the
' startup sequence (InitApp). modApp2 holds setActive, the detector, the self-contained
' parsers and the workbench handlers; modApp3 holds number steppers, the axis-interval
' toggle, theme, sidebar width, typing-to-search and search ranking.
Option Explicit

' The edition's version (the release build rewrites this literal).
Public Const APP_VERSION As String = "021126_2300"
Private Const UPDATE_ENDPOINT As String = "https://zhengjia.dev/bt1101"
Private Const STORAGE_APP As String = "BT1101 Prose Lab"
Private Const STORAGE_SECTION As String = "storage"
Private Const STORAGE_ABSENT As String = "<<bt1101-absent>>"

' window.innerWidth (the renderer updates it and dispatches "resize" on modDom.Win).
Public WindowInnerWidth As Long
' Element the website scrolls into view (scrollIntoView); the renderer may consume it.
Public ScrollRequestTarget As clsNode

' groups: one entry per sidebar operation, in tutorial order.
Public OpCount As Long
Public OpIds() As String, OpLabels() As String, OpTerms() As String, OpGroup() As Long
Public GroupCount As Long
Public GroupTitles() As String, GroupShorts() As String
' samples (single-sample operations) and universalSamples (detector dropdown).
Public SampleCount As Long
Public SampleIds() As String, SampleCodes() As String, SampleInputs() As String
Private mSampleProse() As String
Public UniversalSampleCount As Long
Public UniversalRoutes() As String, UniversalLabels() As String, UniversalInputs() As String

' Storage (localStorage): SaveSetting/GetSetting, or an in-memory map in test mode.
Private mTestModeExplicit As Boolean
Private mTestMode As Boolean
Private mMemoryStore As clsMap

' Update check state.
Private mUpdateCheckState As String
Private mAvailableUpdateVersion As String
Private mAvailableUpdateHtml As String

' Return-focus targets of the warning modals.
Private mPlotAxisRangeReturnFocus As clsNode
Private mVariableRenameReturnFocus As clsNode
Private mKmeansScaleReturnFocus As clsNode
Private mRegressionUnitsReturnFocus As clsNode
Private mKnownPopulationSdReturnFocus As clsNode

' ================================================================ startup

' The website's whole startup, in its script order. Requires the document to be loaded
' (modBoot.StartApp / the test driver call LoadDocument first); opens no window.
' Operation modules' Init procedures run through TryInit, which skips a procedure that is
' not present in the project (so the shell also runs while other modules are missing).
Public Sub InitApp()
    If Not mTestModeExplicit Then
        mTestMode = TestBuildPresent()
        Set mMemoryStore = NewMap()
    End If
    If mMemoryStore Is Nothing Then Set mMemoryStore = NewMap()
    WindowInnerWidth = 1280
    Set ScrollRequestTarget = Nothing
    InitShared
    InitVariants
    ' <head> script: saved theme, else the system colour scheme (before first paint).
    RunHeadThemeScript
    ' Lines 1-70: groups, samples, universalSamples, state.
    LoadAppData
    ' Lines 963-1590: answer-option listeners, update check, modals, sample select,
    ' collapsibles, sidebar list, operation routing.
    modApp2.InitShellPartA
    ' Lines 1591-1695 (first statements of opStats; its later statements at 4704-4742
    ' run inside its single Init procedure here).
    TryInit "InitOpStats"
    ' Lines 3889-4258: textarea fitting, workbench input/sample/paste/copy handlers.
    modApp2.InitShellPartB
    ' Lines 4259-4742: the operation modules' top-level statements, in website order
    ' (opPlot 4259, opLp 4374, opMining 4468, opRegression 4494, opAnova 4516,
    ' opInference 4563, opStats 4704).
    TryInit "InitOpPlot"
    TryInit "InitOpLp"
    TryInit "InitOpMining"
    TryInit "InitOpRegression"
    TryInit "InitOpAnova"
    TryInit "InitOpInference"
    TryInit "InitOpStatsEvents"
    ' Lines 4743-end: number steppers, axis-interval toggles, theme, sidebar, search.
    modApp3.InitShellPartC
    ' The website never calls setActive at startup: state stays { universal, "Universal
    ' detector", "Smart entry point" } (InitShared) with the markup's active universal entry.
End Sub

' Runs an operation module's Init procedure; skips it when the module is absent.
Private Sub TryInit(ByVal procedure As String)
    On Error GoTo Failed
    Application.Run procedure
    Exit Sub
Failed:
    If Err.Number = 1004 Then
        Err.Clear
    Else
        ReportError "modApp.InitApp (" & procedure & ")"
    End If
End Sub

' Calls another module's procedure by name (like CallOp) and returns its result (Empty
' for a Sub). A procedure that is not in the project (its module is absent: error 1004)
' is skipped silently; other failures are reported with ReportError.
Public Function TryCallOp(ByVal procedure As String, ParamArray args() As Variant) As Variant
    On Error GoTo Failed
    Select Case UBound(args) - LBound(args) + 1
        Case 0: TryCallOp = Application.Run(procedure)
        Case 1: TryCallOp = Application.Run(procedure, args(0))
        Case 2: TryCallOp = Application.Run(procedure, args(0), args(1))
        Case Else: TryCallOp = Application.Run(procedure, args(0), args(1), args(2))
    End Select
    Exit Function
Failed:
    If Err.Number = 1004 Then Err.Clear Else ReportError "modApp.TryCallOp (" & procedure & ")"
    TryCallOp = Empty
End Function

' Like TryCallOp for procedures that return an object (Nothing when absent or failing).
Public Function TryCallOpObject(ByVal procedure As String, ByVal arg As Variant) As Object
    On Error GoTo Failed
    Set TryCallOpObject = Application.Run(procedure, arg)
    Exit Function
Failed:
    If Err.Number = 1004 Then Err.Clear Else ReportError "modApp.TryCallOpObject (" & procedure & ")"
    Set TryCallOpObject = Nothing
End Function

' ================================================================ data (lines 1-70)

Private Sub AddGroup(ByVal title As String, ByVal short As String)
    ReDim Preserve GroupTitles(0 To GroupCount)
    ReDim Preserve GroupShorts(0 To GroupCount)
    GroupTitles(GroupCount) = title
    GroupShorts(GroupCount) = short
    GroupCount = GroupCount + 1
End Sub

Private Sub AddOp(ByVal id As String, ByVal label As String, ByVal terms As String)
    ReDim Preserve OpIds(0 To OpCount)
    ReDim Preserve OpLabels(0 To OpCount)
    ReDim Preserve OpTerms(0 To OpCount)
    ReDim Preserve OpGroup(0 To OpCount)
    OpIds(OpCount) = id
    OpLabels(OpCount) = label
    OpTerms(OpCount) = terms
    OpGroup(OpCount) = GroupCount - 1
    OpCount = OpCount + 1
End Sub

Private Sub AddSample(ByVal id As String, ByVal code As String, ByVal inputText As String)
    ReDim Preserve SampleIds(0 To SampleCount)
    ReDim Preserve SampleCodes(0 To SampleCount)
    ReDim Preserve SampleInputs(0 To SampleCount)
    ReDim Preserve mSampleProse(0 To SampleCount)
    SampleIds(SampleCount) = id
    SampleCodes(SampleCount) = code
    SampleInputs(SampleCount) = inputText
    SampleCount = SampleCount + 1
End Sub

Private Sub AddUniversalSample(ByVal route As String, ByVal label As String, ByVal inputText As String)
    ReDim Preserve UniversalRoutes(0 To UniversalSampleCount)
    ReDim Preserve UniversalLabels(0 To UniversalSampleCount)
    ReDim Preserve UniversalInputs(0 To UniversalSampleCount)
    UniversalRoutes(UniversalSampleCount) = route
    UniversalLabels(UniversalSampleCount) = label
    UniversalInputs(UniversalSampleCount) = inputText
    UniversalSampleCount = UniversalSampleCount + 1
End Sub

' samples[id] index, or -1.
Public Function SampleIndex(ByVal id As String) As Long
    Dim i As Long
    For i = 0 To SampleCount - 1
        If SampleIds(i) = id Then SampleIndex = i: Exit Function
    Next i
    SampleIndex = -1
End Function

Public Function SampleProse(ByVal index As Long) As String
    SampleProse = mSampleProse(index)
End Function

' Index of an operation in groups, or -1.
Public Function OpIndex(ByVal id As String) As Long
    Dim i As Long
    For i = 0 To OpCount - 1
        If OpIds(i) = id Then OpIndex = i: Exit Function
    Next i
    OpIndex = -1
End Function

Private Sub LoadAppData()
    Dim t As String
    OpCount = 0: GroupCount = 0: SampleCount = 0: UniversalSampleCount = 0
    AddGroup "Tutorial 2", "Plotting & Tables"
    AddOp "ggplot", "Plotting & tables", "plot chart graph ggplot bar pie histogram boxplot scatter line frequency contingency pareto summarize means"
    AddGroup "Tutorial 4", "Statistical Measures & Probability"
    AddOp "summarize", "Summarize for stats", "summary summarize mean median min max variable name"
    AddOp "shape", "Skewness & kurtosis", "coefficient skewness kurtosis excess kurtosis asymmetry platykurtic leptokurtic describe"
    AddOp "wide-table", "Wide table format", "wide table pivot summary format"
    AddGroup "Tutorial 5", "Statistical Inference"
    AddOp "interval", "Confidence / prediction intervals", "confidence prediction interval mean proportion known unknown standard deviation sample size future observation"
    AddOp "one-sample", "One-sample mean / proportion", "one sample t test prop test z test conditional proportion filter threshold mean proportion mu hypothesis"
    AddOp "two-sample", "Two-sample mean / variance", "two sample independent paired t test var test fisher variance book dvd collectorbook outlier"
    AddOp "anova-many", ">2 samples: ANOVA", "anova aov more than two samples f value regions bartlett tukey"
    AddGroup "Tutorial 6", "Linear Regression"
    AddOp "lm-predict", "One-variable linear regression", "lm simple one variable linear regression scatterplot transform log sqrt inverse diagnostics"
    AddOp "multi-regression", "Multi-variable linear regression", "multiple multivariable lm predictors predict confidence interval vif multicollinearity"
    AddOp "model-significance", "Model significance", "f-statistic r-squared p-value statistical significance"
    AddOp "model-comparison", "Pairwise model comparison", "anova model comparison nested rss deviance"
    AddOp "dummy", "Dummy variables / reference groups", "dummy categorical factor reference group coefficients"
    AddOp "interaction", "Interaction terms", "interaction colon asterisk coefficient moderation"
    AddOp "glm", "GLM / binary variable", "glm binomial logistic logit odds binary"
    AddOp "confusion", "Confusion matrix", "confusion matrix accuracy sensitivity specificity precision"
    AddGroup "Tutorial 7", "Data Mining"
    AddOp "pca", "PCA & component significance", "pca prcomp principal component importance variance loadings"
    AddOp "stepwise", "Forward/Backward Stepwise Selection", "forward backward stepwise null full model add remove predictor aic model selection"
    AddOp "kmeans", "K-means clustering", "kmeans cluster centers withinss betweenss"
    AddOp "kfold", "K-fold cross-validation", "k fold cross validation cv rmse mae accuracy"
    AddGroup "Tutorial 8", "Linear Optimization"
    AddOp "lpsolve", "lpSolve / Sensitivity analysis", "lp solve objective constraints optimum solution linear optimization sensitivity shadow price dual allowable range"
    AddGroup "Others", "Others"
    AddOp "shapiro", JsU("Shapiro\u2013Wilk"), "shapiro.test normality W p-value"
    AddOp "pearson", "Pearson correlation test", "cor.test pearson correlation t df p-value"
    t = "##  Shapiro-Wilk normality test" & vbLf & "##" & vbLf & "## data:  dataset$column1" & vbLf
    t = t & "## W = 0.42617," & vbLf & "## p-value < 2.2e-16"
    AddSample "shapiro", "shapiro.test(dataset$column1)", t
    t = JsU("H\u2080: The variable is normally distributed.") & vbLf & vbLf & JsU("H\u2081: The variable is not normally distributed.") & vbLf & vbLf
    t = t & "Since p-value < 0.05, we reject the null hypothesis at the 5% level of significance. There is sufficient evidence to con" & "clude that dataset$column1 is not normally distributed."
    mSampleProse(SampleCount - 1) = t
    t = "##  Pearson's product-moment correlation" & vbLf & "##" & vbLf & "## data:  dataset$column1 and dataset$column2" & vbLf
    t = t & "## t = 0.46122, df = 209, p-value = 0.6451" & vbLf & "## alternative hypothesis: true correlation is not equal to 0" & vbLf & "## 95 percent confidence interval:" & vbLf
    t = t & "## -0.1036276 0.1662399" & vbLf & "## sample estimates:" & vbLf & "##        cor" & vbLf
    t = t & "## 0.03188728"
    AddSample "pearson", "cor.test(dataset$column1, dataset$column2)", t
    t = JsU("H\u2080: There is no correlation between dataset$column1 and dataset$column2 (\u03C1 = 0).") & vbLf & vbLf & JsU("H\u2081: There is a correlation between dataset$column1 and dataset$column2 (\u03C1 \u2260 0).") & vbLf & vbLf
    t = t & "A correlation test was conducted between dataset$column1 and dataset$column2. The test statistic is t = 0.461 with a p-v" & "alue of 0.645. Since the p-value is greater than 0.05, we fail to reject the null hypothesis at the 5% significance level. " & "There is no statistically significant correlation between the two variables."
    mSampleProse(SampleCount - 1) = t
    t = "##  Shapiro-Wilk normality test" & vbLf & "##" & vbLf & "## data:  dataset$column1" & vbLf
    t = t & "## W = 0.42617," & vbLf & "## p-value < 2.2e-16"
    AddUniversalSample "shapiro", JsU("Shapiro\u2013Wilk"), t
    t = "##  Pearson's product-moment correlation" & vbLf & "##" & vbLf & "## data:  dataset$column1 and dataset$column2" & vbLf
    t = t & "## t = 0.46122, df = 209, p-value = 0.6451" & vbLf & "## alternative hypothesis: true correlation is not equal to 0" & vbLf & "## 95 percent confidence interval:" & vbLf
    t = t & "## -0.1036276 0.1662399" & vbLf & "## sample estimates:" & vbLf & "##        cor" & vbLf
    t = t & "## 0.03188728"
    AddUniversalSample "pearson", "Pearson correlation", t
    t = "     uCIAmount99 lCIAmount99" & vbLf & "[1,]        20.33        19.32"
    AddUniversalSample "interval", "Confidence / prediction interval", t
    t = "z = -1.75" & vbLf & "p-value = 0.04"
    AddUniversalSample "one-sample", "One-sample mean / proportion", t
    t = "Welch Two Sample t-test" & vbLf & vbLf & "data: Amount by Product" & vbLf & "t = 8.03, df = 180.4, p-value = 3.34e-14"
    AddUniversalSample "two-sample", "Two-sample mean / variance", t
    t = "Call:" & vbLf & "glm(formula = column1 ~ column2 + column3, family = binomial, data = dataset)" & vbLf & vbLf & "Coefficients:"
    t = t & vbLf & "            Estimate Std. Error z value Pr(>|z|)" & vbLf & "(Intercept)  -1.4200     0.5100  -2.784  0.00537" & vbLf & "column2       0.6800     0.2400   2.833  0.00461"
    t = t & vbLf & vbLf & "Null deviance: 81.45  on 59  degrees of freedom" & vbLf & "Residual deviance: 69.12  on 57  degrees of freedom"
    AddUniversalSample "glm", "Binary GLM", t
    t = "Call:" & vbLf & "lm(formula = column2 ~ column1, data = dataset)" & vbLf & vbLf & "Residuals:"
    t = t & vbLf & "    Min      1Q  Median      3Q     Max" & vbLf & "-8.120  -2.340   0.110   2.190   7.850" & vbLf & vbLf
    t = t & "Coefficients:" & vbLf & "            Estimate Std. Error t value Pr(>|t|)" & vbLf & "(Intercept)   12.450      2.180   5.711 1.2e-06" & vbLf
    t = t & "column1        1.730      0.310   5.581 1.9e-06" & vbLf & vbLf & "Residual standard error: 4.21 on 48 degrees of freedom" & vbLf & "Multiple R-squared: 0.3936, Adjusted R-squared: 0.3810"
    t = t & vbLf & "F-statistic: 31.15 on 1 and 48 DF, p-value: 1.9e-06"
    AddUniversalSample "multi-regression", "Linear regression output", t
    t = "Confusion Matrix and Statistics" & vbLf & vbLf & "          Reference" & vbLf & "Prediction  No Yes"
    t = t & vbLf & "       No   42   6" & vbLf & "       Yes   8  44" & vbLf & vbLf
    t = t & "Sensitivity : 0.880" & vbLf & "Specificity : 0.840" & vbLf & "Accuracy : 0.860"
    AddUniversalSample "confusion", "Confusion matrix", t
    t = "Importance of components:" & vbLf & "                          PC1    PC2    PC3" & vbLf & "Standard deviation     1.421  0.884  0.541" & vbLf
    t = t & "Proportion of Variance 0.673  0.260  0.067" & vbLf & "Cumulative Proportion  0.673  0.933  1.000"
    AddUniversalSample "pca", "Principal component analysis", t
    t = "| Minimize total effort using | Effort = 6 $X_1$ + 7$X_2$ + |" & vbLf & "| decision variables $X_1$, $X_2$, $X_3$, | 9 $X_3$ + 8 $X_4$ + |" & vbLf & "| $X_4$, $X_5$, $X_6$ | 10 $X_5$ + 20 $X_6$ |" & vbLf
    t = t & "|---|---|" & vbLf & "| Subject to | |" & vbLf & "| Credit Requirement | $2X_1+4X_2+3X_3+3X_4+4X_5+5X_6\geq12$ |" & vbLf
    t = t & "| Weekly effort | $6X_1+7X_2+9X_3+8X_4+10X_5+20X_6\leq30$ |" & vbLf & "| CS2 Prereq | $X_3-X_4\geq0$ |" & vbLf & "| BT2 prereq | $(X_1+X_5)-X_2\geq0$ |" & vbLf
    t = t & "| DT prereq | $(X_1+X_3)-2X_6\geq0$ |" & vbLf & "| Binary, Integer, Non-negativity Constraints | X1, X2, X3 all binary, integers |"
    AddUniversalSample "lpsolve", "lpSolve optimisation model", t
End Sub

' ================================================================ storage (localStorage)

' Test mode: in-memory storage (never the real saved settings), no network, light system
' theme and no clipboard reading. Decided automatically on each InitApp (the test-only
' modules are present) unless set explicitly here; SetAppTestMode also clears the store.
Public Sub SetAppTestMode(ByVal enabled As Boolean)
    mTestModeExplicit = True
    mTestMode = enabled
    Set mMemoryStore = NewMap()
End Sub

Public Function AppTestMode() As Boolean
    AppTestMode = mTestMode
End Function

Private Function TestBuildPresent() As Boolean
    Dim result As Variant
    On Error GoTo Absent
    result = Application.Run("testDriver.EscapeStep", "")
    TestBuildPresent = True
    Exit Function
Absent:
    Err.Clear
    TestBuildPresent = False
End Function

' localStorage.getItem: Null when absent.
Public Function StorageGetItem(ByVal key As String) As Variant
    Dim value As String
    StorageGetItem = Null
    If mTestMode Then
        If mMemoryStore Is Nothing Then Set mMemoryStore = NewMap()
        If mMemoryStore.Exists(key) Then StorageGetItem = CStr(mMemoryStore(key))
        Exit Function
    End If
    On Error GoTo Failed
    value = GetSetting(STORAGE_APP, STORAGE_SECTION, key, STORAGE_ABSENT)
    If value <> STORAGE_ABSENT Then StorageGetItem = value
    Exit Function
Failed:
    Err.Clear
End Function

Public Sub StorageSetItem(ByVal key As String, ByVal value As String)
    If mTestMode Then
        If mMemoryStore Is Nothing Then Set mMemoryStore = NewMap()
        mMemoryStore(key) = value
        Exit Sub
    End If
    On Error Resume Next
    SaveSetting STORAGE_APP, STORAGE_SECTION, key, value
    Err.Clear
End Sub

Public Sub StorageRemoveItem(ByVal key As String)
    If mTestMode Then
        If Not mMemoryStore Is Nothing Then
            If mMemoryStore.Exists(key) Then mMemoryStore.Remove key
        End If
        Exit Sub
    End If
    On Error Resume Next
    DeleteSetting STORAGE_APP, STORAGE_SECTION, key
    Err.Clear
End Sub

' matchMedia("(prefers-color-scheme: dark)").matches.
Public Function SystemPrefersDark() As Boolean
    If mTestMode Then Exit Function
#If Mac Then
    SystemPrefersDark = False
#Else
    Dim shell As Object, value As Variant
    On Error GoTo Failed
    Set shell = CreateObject("WScript.Shell")
    value = shell.RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme")
    SystemPrefersDark = (CLng(value) = 0)
    Exit Function
Failed:
    Err.Clear
    SystemPrefersDark = False
#End If
End Function

' The <head> script.
Private Sub RunHeadThemeScript()
    Dim saved As Variant
    On Error GoTo Failed
    saved = StorageGetItem("bt1101-theme")
    If Not IsNull(saved) Then
        If saved = "dark" Or saved = "light" Then
            DocumentElement.DataSet "theme", CStr(saved)
            Exit Sub
        End If
    End If
    DocumentElement.DataSet "theme", IIf(SystemPrefersDark(), "dark", "light")
    Exit Sub
Failed:
    Err.Clear
    DocumentElement.DataSet "theme", "light"
End Sub

' ================================================================ update check

' Returns a clsMap { canonical, timestamp (ms since 1970, UTC) }, or Nothing when invalid.
Public Function ParseReleaseVersion(ByVal value As String) As clsMap
    Dim m As clsMatch, d As Long, mo As Long, y As Long, h As Long, mi As Long, result As clsMap
    Set m = RxMatch(JsTrim(value), "^v?(\d{2})(\d{2})(\d{2})_(\d{2})(\d{2})h?$", "i")
    If m Is Nothing Then Exit Function
    d = CLng(m.Group(1)): mo = CLng(m.Group(2)): y = 2000 + CLng(m.Group(3))
    h = CLng(m.Group(4)): mi = CLng(m.Group(5))
    If mo < 1 Or mo > 12 Or d < 1 Or h > 23 Or mi > 59 Then Exit Function
    If d > Day(DateSerial(y, mo + 1, 0)) Then Exit Function
    Set result = NewMap()
    result("canonical") = m.Group(1) & m.Group(2) & m.Group(3) & "_" & m.Group(4) & m.Group(5)
    result("timestamp") = CDbl(DateSerial(y, mo, d) - DateSerial(1970, 1, 1)) * 86400000# + h * 3600000# + mi * 60000#
    Set ParseReleaseVersion = result
End Function

Public Function DisplayReleaseVersion(ByVal version As String) As String
    DisplayReleaseVersion = "v" & version & "h"
End Function

Public Function ReleaseVersionFromFirstLine(ByVal html As String) As clsMap
    Dim firstLine As String, p As Long, marker As clsMatch
    p = InStr(1, html, vbLf, vbBinaryCompare)
    If p > 0 Then firstLine = Left$(html, p - 1) Else firstLine = html
    If Right$(firstLine, 1) = vbCr Then firstLine = Left$(firstLine, Len(firstLine) - 1)
    firstLine = JsTrim(firstLine)
    Set marker = RxMatch(firstLine, "^<!--\s*BT1101 Prose Lab version:\s*(v?\d{6}_\d{4}h?)\s*-->$", "i")
    If marker Is Nothing Then Exit Function
    Set ReleaseVersionFromFirstLine = ParseReleaseVersion(marker.Group(1))
End Function

' The Excel edition has no page hostname, so it is never the official hosted copy.
Public Function IsOfficialHostedCopy() As Boolean
    IsOfficialHostedCopy = False
End Function

Public Function PrepareStandaloneUpdateHtml(ByVal html As String) As String
    Dim appMarker As Long, appScriptStart As Long, appScriptOpenEnd As Long, appScriptClose As Long
    Dim scriptCloseTag As String, appScriptTag As String, appEnd As Long, beforeApp As String, app As String
    Dim cleanedTail As String, matches As Collection, m As clsMatch, out As String, last As Long, body As String
    appMarker = JsIndexOf(html, "const groups = [")
    If appMarker < 0 Then Err.Raise vbObjectError + 101, "modApp", "Update file does not contain the BT1101 application script."
    appScriptStart = JsLastIndexOf(html, "<script", appMarker)
    appScriptOpenEnd = JsIndexOf(html, ">", appScriptStart)
    scriptCloseTag = "</scr" & "ipt>"
    appScriptClose = JsIndexOf(html, scriptCloseTag, appMarker)
    If appScriptStart < 0 Or appScriptOpenEnd < 0 Or appScriptClose < 0 Then Err.Raise vbObjectError + 101, "modApp", "Update file does not contain the BT1101 application script."
    ' /\s+type=(["'])[A-Za-z0-9]+-text\/javascript\1/i (the back-reference written as two alternatives)
    appScriptTag = Rx("\s+type=""[A-Za-z0-9]+-text\/javascript""|\s+type='[A-Za-z0-9]+-text\/javascript'", "i").Replace(JsSlice(html, appScriptStart, appScriptOpenEnd + 1), "")
    appEnd = appScriptClose + Len(scriptCloseTag)
    beforeApp = JsSlice(html, 0, appScriptStart)
    app = appScriptTag & JsSlice(html, appScriptOpenEnd + 1, appEnd)
    cleanedTail = Rx("<script\b(?=[^>]*\bsrc=[""'][^""']*\/cdn-cgi\/scripts\/[^""']*\/rocket-loader\.min\.js[""'])[^>]*>\s*<\/script>", "gi").Replace(JsSlice(html, appEnd), "")
    Set matches = Rx("<script\b([^>]*)>([\s\S]*?)<\/script>", "gi").MatchAll(cleanedTail)
    For Each m In matches
        out = out & JsSlice(cleanedTail, last, m.Index)
        body = m.Group(2)
        If Len(JsTrim(m.Group(1))) = 0 And Len(body) < 5000 And InStr(1, body, "window.__CF$cv$params", vbBinaryCompare) > 0 _
            And InStr(1, body, "/cdn-cgi/challenge-platform/", vbBinaryCompare) > 0 Then
            ' dropped
        Else
            out = out & m.Value
        End If
        last = m.EndIndex
    Next m
    out = out & JsSlice(cleanedTail, last)
    PrepareStandaloneUpdateHtml = beforeApp & app & out
End Function

Public Sub SetUpdateCheckState(ByVal nextState As String, Optional ByVal version As String = "")
    Dim box As clsNode, button As clsNode, message As clsNode
    Set box = Q("#updateCheck")
    Set button = Q("#checkForUpdatesBtn")
    Set message = Q("#updateCheckMessage")
    mUpdateCheckState = nextState
    mAvailableUpdateVersion = version
    box.RemoveClass "is-checking", "is-current", "is-offline", "is-update"
    box.AddClass "is-" & nextState
    button.Disabled = (nextState = "checking")
    Select Case nextState
        Case "current": message.TextContent = "Well done! You're on the latest version!"
        Case "offline"
#If Mac Then
            ' No HTTP without helpers on Mac: the check cannot run, so offer the download page.
            If mTestMode Then
                message.TextContent = "No internet connection. Retry?"
            Else
                message.TextContent = "Automatic update checks aren't available on Mac " & ChrW$(&H2013) & " click to open the download page"
            End If
#Else
            message.TextContent = "No internet connection. Retry?"
#End If
        Case "update": message.TextContent = "Click here to update to the latest version!"
        Case Else: message.TextContent = JsU("Checking for updates\u2026")
    End Select
End Sub

Public Function UpdateCheckState() As String
    UpdateCheckState = mUpdateCheckState
End Function

' Fetches the hosted website file. Windows only (MSXML); False when unavailable, in test
' mode and on Mac (no HTTP without helpers there).
Private Function FetchUpdateHtml(ByRef html As String) As Boolean
    If mTestMode Then Exit Function
#If Mac Then
    FetchUpdateHtml = False
#Else
    Dim request As Object
    On Error GoTo Failed
    Set request = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    request.setTimeouts 3000, 3000, 4000, 6000
    request.Open "GET", UPDATE_ENDPOINT, False
    request.setRequestHeader "Cache-Control", "no-cache"
    request.setRequestHeader "Pragma", "no-cache"
    request.send
    If request.Status < 200 Or request.Status > 299 Then Exit Function
    html = request.responseText
    FetchUpdateHtml = True
    Exit Function
Failed:
    Err.Clear
    FetchUpdateHtml = False
#End If
End Function

Public Sub CheckForUpdates()
    SetUpdateCheckState "checking"
    mAvailableUpdateHtml = ""
    If IsOfficialHostedCopy() Then
        SetUpdateCheckState "current"
        Exit Sub
    End If
    ' The fetch is asynchronous in the website: here it runs after startup/the click.
#If Mac Then
    QueueMicrotask "modApp.FinishUpdateCheck"
#Else
    If mTestMode Then
        QueueMicrotask "modApp.FinishUpdateCheck"
    Else
        On Error GoTo Immediate
        Application.OnTime Now + TimeSerial(0, 0, 1), "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!modApp.FinishUpdateCheckOnTime"
    End If
#End If
    Exit Sub
Immediate:
    Err.Clear
    QueueMicrotask "modApp.FinishUpdateCheck"
End Sub

' Microtask/OnTime continuation of checkForUpdates (the fetch and version comparison).
Public Sub FinishUpdateCheck()
    Dim html As String, localRelease As clsMap, latestRelease As clsMap
    On Error GoTo Failed
    If Q("#updateCheck") Is Nothing Then Exit Sub
    If Not FetchUpdateHtml(html) Then
        SetUpdateCheckState "offline"
        Exit Sub
    End If
    Set localRelease = ParseReleaseVersion(APP_VERSION)
    Set latestRelease = ReleaseVersionFromFirstLine(html)
    If localRelease Is Nothing Or latestRelease Is Nothing Then
        SetUpdateCheckState "offline"
    ElseIf latestRelease("timestamp") <= localRelease("timestamp") Then
        SetUpdateCheckState "current"
    Else
        mAvailableUpdateHtml = html
        SetUpdateCheckState "update", CStr(latestRelease("canonical"))
    End If
    Exit Sub
Failed:
    ReportError "modApp.FinishUpdateCheck"
    On Error Resume Next
    SetUpdateCheckState "offline"
    Err.Clear
End Sub

' Application.OnTime entry: runs the check, settles and asks the renderer to redraw.
Public Sub FinishUpdateCheckOnTime()
    On Error GoTo Failed
    If Doc Is Nothing Then Exit Sub
    FinishUpdateCheck
    Settle 700
    On Error Resume Next
    Application.Run "modRender.RefreshAppWindow"
    Err.Clear
    Exit Sub
Failed:
    ReportError "modApp.FinishUpdateCheckOnTime"
End Sub

' Saves the update (the website downloads a Blob). Raises an error on failure.
Private Sub DownloadUpdate(ByVal version As String)
    Dim html As String, release As clsMap, standalone As String, target As Variant
    html = mAvailableUpdateHtml
    If Len(html) = 0 Then
        If Not FetchUpdateHtml(html) Then Err.Raise vbObjectError + 102, "modApp", "Update service is unavailable."
    End If
    Set release = ReleaseVersionFromFirstLine(html)
    If release Is Nothing Then Err.Raise vbObjectError + 103, "modApp", "Update service did not return a valid BT1101 Prose Lab file."
    If release("canonical") <> version Or Not RxTest(html, "<\!doctype html>", "i") Or Not RxTest(html, "<title>\s*BT1101 Prose Lab\s*<\/title>", "i") Then
        Err.Raise vbObjectError + 103, "modApp", "Update service did not return a valid BT1101 Prose Lab file."
    End If
    standalone = PrepareStandaloneUpdateHtml(html)
    If mTestMode Then Exit Sub
#If Mac Then
    target = Application.GetSaveAsFilename(InitialFileName:="BT1101-Prose-Lab-" & DisplayReleaseVersion(version) & ".html")
#Else
    target = Application.GetSaveAsFilename(InitialFileName:="BT1101-Prose-Lab-" & DisplayReleaseVersion(version) & ".html", FileFilter:="HTML files (*.html), *.html")
#End If
    If VarType(target) = vbBoolean Then Exit Sub
    WriteUtf8Text CStr(target), standalone
End Sub

' Writes text as UTF-8 without a BOM.
Public Sub WriteUtf8Text(ByVal path As String, ByVal text As String)
    Dim handle As Integer, bytes() As Byte, n As Long, i As Long, code As Long, low As Long
    ReDim bytes(0 To Len(text) * 3 + 3)
    i = 1
    Do While i <= Len(text)
        code = AscW(Mid$(text, i, 1)) And &HFFFF&
        If code >= &HD800& And code <= &HDBFF& And i < Len(text) Then
            low = AscW(Mid$(text, i + 1, 1)) And &HFFFF&
            If low >= &HDC00& And low <= &HDFFF& Then
                code = &H10000 + (code - &HD800&) * 1024 + (low - &HDC00&)
                i = i + 1
            End If
        End If
        If code < &H80 Then
            bytes(n) = code: n = n + 1
        ElseIf code < &H800 Then
            bytes(n) = &HC0 Or (code \ 64): bytes(n + 1) = &H80 Or (code And &H3F): n = n + 2
        ElseIf code < &H10000 Then
            bytes(n) = &HE0 Or (code \ 4096): bytes(n + 1) = &H80 Or ((code \ 64) And &H3F): bytes(n + 2) = &H80 Or (code And &H3F): n = n + 3
        Else
            bytes(n) = &HF0 Or (code \ 262144): bytes(n + 1) = &H80 Or ((code \ 4096) And &H3F)
            bytes(n + 2) = &H80 Or ((code \ 64) And &H3F): bytes(n + 3) = &H80 Or (code And &H3F): n = n + 4
        End If
        i = i + 1
    Loop
    handle = FreeFile
    If Len(Dir$(path)) > 0 Then Kill path
    Open path For Binary Access Write As #handle
    If n > 0 Then
        ReDim Preserve bytes(0 To n - 1)
        Put #handle, , bytes
    End If
    Close #handle
End Sub

Public Sub HandleUpdateCheckClick(ByVal e As clsEvent)
    Dim version As String
    On Error GoTo Failed
#If Mac Then
    If Not mTestMode And mUpdateCheckState = "offline" Then
        ThisWorkbook.FollowHyperlink UPDATE_ENDPOINT
        Exit Sub
    End If
#End If
    If mUpdateCheckState <> "update" Or Len(mAvailableUpdateVersion) = 0 Then
        CheckForUpdates
        Exit Sub
    End If
    version = mAvailableUpdateVersion
    Q("#checkForUpdatesBtn").Disabled = True
    On Error GoTo DownloadFailed
    DownloadUpdate version
    On Error GoTo Failed
    SetUpdateCheckState "update", version
    Exit Sub
DownloadFailed:
    Err.Clear
    On Error GoTo Failed
    SetUpdateCheckState "offline"
    Exit Sub
Failed:
    ReportError "modApp.HandleUpdateCheckClick"
End Sub

' ================================================================ warning modals

Public Function AlertMarkup() As String
    Dim t As String
    t = "" & vbLf & "      <div class=""variable-rename-alert hidden"" id=""variableRenameAlert"" role=""dialog"" aria-modal=""true"" aria-labelledby" & "=""variableRenameTitle"" aria-describedby=""variableRenameDescription"">" & vbLf & "        <div class=""variable-rename-dialog"">"
    t = t & vbLf & "          <div class=""variable-rename-kicker"">Required before submission</div>" & vbLf & "          <h2 id=""variableRenameTitle"">Replace the placeholder variable names</h2>" & vbLf & "          <p id=""variableRenameDescription"">The copied prose still uses <strong>X1, X2, X3, and so on</strong>. Replace "
    t = t & "every placeholder with the real decision-variable name before using or submitting the answer.</p>" & vbLf & "          <div class=""variable-rename-example""><strong>Do not leave:</strong> 700 X1, 0 X2, and 133.33 X3<br><strong>Wri" & "te instead:</strong> 700 lawn chairs, 0 benches, and 133.33 tables</div>" & vbLf & "          <div class=""variable-rename-actions""><button class=""primary-btn variable-rename-dismiss"" id=""dismissVariableRe"
    t = t & "nameAlert"" type=""button"">I understand</button></div>" & vbLf & "        </div>" & vbLf & "      </div>" & vbLf
    t = t & "      <div class=""variable-rename-alert hidden"" id=""kmeansScaleAlert"" role=""dialog"" aria-modal=""true"" aria-labelledby=""k" & "meansScaleAlertTitle"" aria-describedby=""kmeansScaleAlertDescription"">" & vbLf & "        <div class=""variable-rename-dialog"">" & vbLf & "          <div class=""variable-rename-kicker"">K-means scaling warning</div>"
    t = t & vbLf & "          <h2 id=""kmeansScaleAlertTitle"">Are you sure to disable scaling?</h2>" & vbLf & "          <p id=""kmeansScaleAlertDescription"">Disable scaling only when the variables are already directly comparable; o" & "therwise, scale them for distance-based methods such as k-means.</p>" & vbLf
    t = t & "          <div class=""variable-rename-actions""><button class=""sample-btn variable-rename-dismiss"" id=""keepKmeansScalingB" & "tn"" type=""button"">Keep scaling</button><button class=""primary-btn danger-btn variable-rename-dismiss"" id=""disableKmeansS" & "calingBtn"" type=""button"">Disable scaling</button></div>" & vbLf & "        </div>" & vbLf
    t = t & "      </div>" & vbLf & "      <div class=""variable-rename-alert hidden"" id=""regressionUnitsAlert"" role=""dialog"" aria-modal=""true"" aria-labelledb" & "y=""regressionUnitsAlertTitle"" aria-describedby=""regressionUnitsAlertDescription"">" & vbLf & "        <div class=""variable-rename-dialog"">"
    t = t & vbLf & "          <div class=""variable-rename-kicker"">Required unit check</div>" & vbLf & "          <h2 id=""regressionUnitsAlertTitle"">Change the generic units before submitting</h2>" & vbLf & JsU("          <p id=""regressionUnitsAlertDescription"">The copied interpretation currently says <strong>\u201Cfor every 1-unit inc")
    t = t & JsU("rease\u201D</strong> and describes the outcome in generic <strong>\u201Cunits\u201D</strong>. Replace both with the measurements used b") & "y your variables.</p>" & vbLf & JsU("          <div class=""variable-rename-example""><strong>Examples:</strong> \u201Cfor every 1-minute increase\u201D, \u201Cfor every $1 i") & JsU("ncrease\u201D, \u201Cfor every 1-year increase\u201D, or \u201Cfor every 1-percentage-point increase\u201D. Also replace the outcome\u2019s \u201Cunits\u201D wi") & "th minutes, dollars, hours, points, or the appropriate measurement.</div>"
    t = t & vbLf & "          <div class=""variable-rename-actions""><button class=""primary-btn variable-rename-dismiss"" id=""dismissRegression" & "UnitsAlert"" type=""button"">I understand</button></div>" & vbLf & "        </div>" & vbLf
    t = t & "      </div>" & vbLf & "      <div class=""variable-rename-alert hidden"" id=""plotAxisRangeAlert"" role=""dialog"" aria-modal=""true"" aria-labelledby=" & """plotAxisRangeTitle"" aria-describedby=""plotAxisRangeDescription plotAxisRangeFootnote"">" & vbLf & "        <div class=""variable-rename-dialog"">"
    t = t & vbLf & "          <div class=""variable-rename-kicker"">Check your chart before using it</div>" & vbLf & "          <h2 id=""plotAxisRangeTitle"">Make sure the axes cover the entire dataset</h2>" & vbLf & "          <p id=""plotAxisRangeDescription""><strong>Calculate axis intervals is off.</strong> After running the code, che"
    t = t & "ck that the ranges of both axes cover the entire dataset and that no bars, points or labels extend beyond the chart. Adj" & "ust the axis limits if needed. For stacked bars, check the full stacked totals.</p>" & vbLf & "          <div class=""variable-rename-actions""><button class=""primary-btn variable-rename-dismiss"" id=""dismissPlotAxisRa" & "ngeAlert"" type=""button"">I understand</button></div>" & vbLf
    t = t & "          <small id=""plotAxisRangeFootnote"" style=""display:block;margin-top:12px;font-size:13px"">Press Escape to dismiss" & " this warning.</small>" & vbLf & "        </div>" & vbLf & "      </div>"
    t = t & vbLf & "      <div class=""variable-rename-alert hidden"" id=""knownPopulationSdAlert"" role=""dialog"" aria-modal=""true"" aria-labelle" & "dby=""knownPopulationSdTitle"" aria-describedby=""knownPopulationSdDescription"">" & vbLf & "        <div class=""variable-rename-dialog"">" & vbLf
    t = t & "          <div class=""variable-rename-kicker"">Standard-deviation check</div>" & vbLf & "          <h2 id=""knownPopulationSdTitle"">Are you sure?</h2>" & vbLf & "          <p id=""knownPopulationSdDescription"">Are you sure you know the <strong>population</strong> (and not sample) st" & "andard deviation?</p>"
    t = t & vbLf & JsU("          <div class=""variable-rename-example"">Use \u201Cknown standard deviation\u201D only when the population standard deviatio") & JsU("n, \u03C3, is supplied. A standard deviation calculated from the sample is not the population standard deviation.</div>") & vbLf & "          <div class=""variable-rename-actions""><button class=""sample-btn variable-rename-dismiss"" id=""useUnknownSdBtn"" t" & "ype=""button"">Use unknown standard deviation</button><button class=""primary-btn danger-btn variable-rename-dismiss"" id=""c"
    t = t & "onfirmKnownSdBtn"" type=""button"">Yes, population SD is known</button></div>" & vbLf & "        </div>" & vbLf & "      </div>" & vbLf
    t = t & "    "
    AlertMarkup = t
End Function

Public Sub InsertShellMarkup()
    Q("#sensSolution").Closest("label").InsertAdjacentHtml "afterend", "<label>Objective value<input class=""column-input sensitivity-input"" id=""sensObjectiveValue"" placeholder=""2766.667""></label>"
    Body.InsertAdjacentHtml "beforeend", AlertMarkup()
End Sub

Public Sub RegisterAlertListeners()
    Q("#dismissPlotAxisRangeAlert").AddEventListener "click", "modApp.OnDismissPlotAxisRangeClick"
    Q("#plotAxisRangeAlert").AddEventListener "keydown", "modApp.OnPlotAxisRangeKeydown"
    Q("#dismissVariableRenameAlert").AddEventListener "click", "modApp.OnDismissVariableRenameClick"
    Q("#variableRenameAlert").AddEventListener "keydown", "modApp.OnVariableRenameKeydown"
    Q("#keepKmeansScalingBtn").AddEventListener "click", "modApp.OnKeepKmeansScalingClick"
    Q("#disableKmeansScalingBtn").AddEventListener "click", "modApp.OnDisableKmeansScalingClick"
    Q("#kmeansScaleAlert").AddEventListener "keydown", "modApp.OnKmeansScaleKeydown"
    Q("#dismissRegressionUnitsAlert").AddEventListener "click", "modApp.OnDismissRegressionUnitsClick"
    Q("#regressionUnitsAlert").AddEventListener "keydown", "modApp.OnRegressionUnitsKeydown"
    Q("#useUnknownSdBtn").AddEventListener "click", "modApp.OnUseUnknownSdClick"
    Q("#confirmKnownSdBtn").AddEventListener "click", "modApp.OnConfirmKnownSdClick"
    Q("#knownPopulationSdAlert").AddEventListener "keydown", "modApp.OnKnownPopulationSdKeydown"
End Sub

' containsRChart(code): uses modVariants.RTokens (tokens are clsMap { text } or strings).
Public Function ContainsRChart(ByVal code As String) As Boolean
    Dim tokens As clsList, i As Long, t As String
    Const CHART_CALLS As String = "|plot|barplot|hist|pie|boxplot|mosaicplot|ggplot|qplot|biplot|pairs|pairs.panels|fviz_eig|fviz_cluster|screeplot|qqnorm|qqplot|matplot|"
    Set tokens = RTokens(code)
    For i = 0 To tokens.Count - 2
        t = TokenText(tokens(i))
        If Len(t) > 0 Then
            If InStr(1, CHART_CALLS, "|" & t & "|", vbBinaryCompare) > 0 Then
                If TokenText(tokens(i + 1)) = "(" Then ContainsRChart = True: Exit Function
            End If
        End If
    Next i
End Function

Private Function TokenText(ByVal token As Variant) As String
    On Error GoTo Failed
    If IsObject(token) Then
        If TypeOf token Is clsMap Then
            TokenText = CStr(token("text"))
        Else
            TokenText = CStr(CallByName(token, "Text", VbGet))
        End If
    Else
        TokenText = CStr(token)
    End If
    Exit Function
Failed:
    Err.Clear
    TokenText = ""
End Function

' getClientRects().length > 0: connected and not inside a hidden element.
Private Function IsRendered(ByVal node As clsNode) As Boolean
    If node Is Nothing Then Exit Function
    If node.NodeType <> 1 Then Exit Function
    If Not node.IsConnected Then Exit Function
    If Not node.Closest(".hidden, [hidden]") Is Nothing Then Exit Function
    ' Like getClientRects(): content of a collapsed card (outside its head) and of a
    ' closed <details> (outside its summary) has no boxes.
    Dim child As clsNode, parent As clsNode
    Set child = node
    Set parent = node.ParentElement
    Do While Not parent Is Nothing
        If parent.HasClass("lp-collapsible") And parent.HasClass("is-collapsed") And Not child.HasClass("panel-head") Then Exit Function
        If parent.LocalName = "details" And Not parent.IsOpen And child.LocalName <> "summary" Then Exit Function
        Set child = parent
        Set parent = parent.ParentElement
    Loop
    IsRendered = True
End Function

Private Sub FocusIfPresent(ByVal node As clsNode)
    If Not node Is Nothing Then node.Focus
End Sub

Public Sub ShowPlotAxisRangeAlert(Optional ByVal button As clsNode = Nothing)
    If button Is Nothing Then Set mPlotAxisRangeReturnFocus = ActiveElement() Else Set mPlotAxisRangeReturnFocus = button
    Q("#plotAxisRangeAlert").RemoveClass "hidden"
    Body.AddClass "variable-rename-alert-open"
    RequestAnimationFrame "modApp.FocusPlotAxisRangeDismiss"
End Sub

Public Sub FocusPlotAxisRangeDismiss()
    On Error GoTo Failed
    If Not Q("#plotAxisRangeAlert").HasClass("hidden") Then Q("#dismissPlotAxisRangeAlert").Focus
    Exit Sub
Failed:
    ReportError "modApp.FocusPlotAxisRangeDismiss"
End Sub

Public Sub DismissPlotAxisRangeAlert()
    Dim target As clsNode
    Q("#plotAxisRangeAlert").AddClass "hidden"
    If Q(".variable-rename-alert:not(.hidden)") Is Nothing Then Body.RemoveClass "variable-rename-alert-open"
    Set target = mPlotAxisRangeReturnFocus
    If IsRendered(target) Then
        If Not target.Disabled Then
            target.Focus
            Exit Sub
        End If
    End If
    Q("#searchInput").Focus
End Sub

Public Sub OnDismissPlotAxisRangeClick(ByVal e As clsEvent)
    On Error GoTo Failed
    DismissPlotAxisRangeAlert
    Exit Sub
Failed:
    ReportError "modApp.OnDismissPlotAxisRangeClick"
End Sub

Public Sub OnPlotAxisRangeKeydown(ByVal e As clsEvent)
    On Error GoTo Failed
    If e.Key = "Escape" Then
        e.PreventDefault
        e.StopPropagation
        DismissPlotAxisRangeAlert
    ElseIf e.Key = "Tab" Then
        e.PreventDefault
        Q("#dismissPlotAxisRangeAlert").Focus
    End If
    Exit Sub
Failed:
    ReportError "modApp.OnPlotAxisRangeKeydown"
End Sub

Public Sub ShowVariableRenameAlert()
    Set mVariableRenameReturnFocus = ActiveElement()
    Q("#variableRenameAlert").RemoveClass "hidden"
    Body.AddClass "variable-rename-alert-open"
    RequestAnimationFrame "modApp.FocusVariableRenameDismiss"
End Sub

Public Sub FocusVariableRenameDismiss()
    On Error GoTo Failed
    Q("#dismissVariableRenameAlert").Focus
    Exit Sub
Failed:
    ReportError "modApp.FocusVariableRenameDismiss"
End Sub

Public Sub DismissVariableRenameAlert()
    Q("#variableRenameAlert").AddClass "hidden"
    Body.RemoveClass "variable-rename-alert-open"
    FocusIfPresent mVariableRenameReturnFocus
End Sub

Public Sub OnDismissVariableRenameClick(ByVal e As clsEvent)
    On Error GoTo Failed
    DismissVariableRenameAlert
    Exit Sub
Failed:
    ReportError "modApp.OnDismissVariableRenameClick"
End Sub

Public Sub OnVariableRenameKeydown(ByVal e As clsEvent)
    On Error GoTo Failed
    If e.Key = "Escape" Then
        e.PreventDefault
        Exit Sub
    End If
    If e.Key = "Tab" Then
        e.PreventDefault
        Q("#dismissVariableRenameAlert").Focus
    End If
    Exit Sub
Failed:
    ReportError "modApp.OnVariableRenameKeydown"
End Sub

Public Sub ShowKmeansScaleAlert()
    Set mKmeansScaleReturnFocus = ActiveElement()
    Q("#kmeansScaleAlert").RemoveClass "hidden"
    Body.AddClass "variable-rename-alert-open"
    RequestAnimationFrame "modApp.FocusKeepKmeansScaling"
End Sub

Public Sub FocusKeepKmeansScaling()
    On Error GoTo Failed
    Q("#keepKmeansScalingBtn").Focus
    Exit Sub
Failed:
    ReportError "modApp.FocusKeepKmeansScaling"
End Sub

Public Sub CloseKmeansScaleAlert(Optional ByVal disableScaling As Boolean = False)
    Q("#kmeansScaleAlert").AddClass "hidden"
    Body.RemoveClass "variable-rename-alert-open"
    Q("#kmeansScale").Checked = Not disableScaling
    TryCallOp "RenderKmeansCode"
    FocusIfPresent mKmeansScaleReturnFocus
End Sub

Public Sub OnKeepKmeansScalingClick(ByVal e As clsEvent)
    On Error GoTo Failed
    CloseKmeansScaleAlert False
    Exit Sub
Failed:
    ReportError "modApp.OnKeepKmeansScalingClick"
End Sub

Public Sub OnDisableKmeansScalingClick(ByVal e As clsEvent)
    On Error GoTo Failed
    CloseKmeansScaleAlert True
    Exit Sub
Failed:
    ReportError "modApp.OnDisableKmeansScalingClick"
End Sub

' Tab/Shift+Tab cycles between two buttons; Escape is swallowed.
Private Sub CycleTwoButtons(ByVal e As clsEvent, ByVal firstSelector As String, ByVal secondSelector As String)
    Dim buttons(0 To 1) As clsNode, index As Long, active As clsNode, nextIndex As Long
    If e.Key = "Escape" Then
        e.PreventDefault
        Exit Sub
    End If
    If e.Key <> "Tab" Then Exit Sub
    Set buttons(0) = Q(firstSelector)
    Set buttons(1) = Q(secondSelector)
    Set active = ActiveElement()
    index = -1
    If Not active Is Nothing Then
        If active Is buttons(0) Then index = 0 Else If active Is buttons(1) Then index = 1
    End If
    e.PreventDefault
    nextIndex = ((index + IIf(e.ShiftKey, -1, 1) + 2) Mod 2 + 2) Mod 2
    buttons(nextIndex).Focus
End Sub

Public Sub OnKmeansScaleKeydown(ByVal e As clsEvent)
    On Error GoTo Failed
    CycleTwoButtons e, "#keepKmeansScalingBtn", "#disableKmeansScalingBtn"
    Exit Sub
Failed:
    ReportError "modApp.OnKmeansScaleKeydown"
End Sub

Public Sub ShowRegressionUnitsAlert()
    Set mRegressionUnitsReturnFocus = ActiveElement()
    Q("#regressionUnitsAlert").RemoveClass "hidden"
    Body.AddClass "variable-rename-alert-open"
    RequestAnimationFrame "modApp.FocusRegressionUnitsDismiss"
End Sub

Public Sub FocusRegressionUnitsDismiss()
    On Error GoTo Failed
    Q("#dismissRegressionUnitsAlert").Focus
    Exit Sub
Failed:
    ReportError "modApp.FocusRegressionUnitsDismiss"
End Sub

Public Sub DismissRegressionUnitsAlert()
    Q("#regressionUnitsAlert").AddClass "hidden"
    Body.RemoveClass "variable-rename-alert-open"
    FocusIfPresent mRegressionUnitsReturnFocus
End Sub

Public Sub OnDismissRegressionUnitsClick(ByVal e As clsEvent)
    On Error GoTo Failed
    DismissRegressionUnitsAlert
    Exit Sub
Failed:
    ReportError "modApp.OnDismissRegressionUnitsClick"
End Sub

Public Sub OnRegressionUnitsKeydown(ByVal e As clsEvent)
    On Error GoTo Failed
    If e.Key = "Escape" Then
        e.PreventDefault
        Exit Sub
    End If
    If e.Key = "Tab" Then
        e.PreventDefault
        Q("#dismissRegressionUnitsAlert").Focus
    End If
    Exit Sub
Failed:
    ReportError "modApp.OnRegressionUnitsKeydown"
End Sub

Public Sub ShowKnownPopulationSdAlert()
    Set mKnownPopulationSdReturnFocus = ActiveElement()
    Q("#knownPopulationSdAlert").RemoveClass "hidden"
    Body.AddClass "variable-rename-alert-open"
    RequestAnimationFrame "modApp.FocusUseUnknownSd"
End Sub

Public Sub FocusUseUnknownSd()
    On Error GoTo Failed
    Q("#useUnknownSdBtn").Focus
    Exit Sub
Failed:
    ReportError "modApp.FocusUseUnknownSd"
End Sub

Public Sub CloseKnownPopulationSdAlert(Optional ByVal useKnown As Boolean = False)
    Dim stdev As clsNode
    Q("#knownPopulationSdAlert").AddClass "hidden"
    Body.RemoveClass "variable-rename-alert-open"
    Set stdev = Q("#intervalStdev")
    stdev.Value = IIf(useKnown, "known", "unknown")
    stdev.DataSet "previous", stdev.Value
    TryCallOp "SyncIntervalForm"
    FocusIfPresent mKnownPopulationSdReturnFocus
End Sub

Public Sub OnUseUnknownSdClick(ByVal e As clsEvent)
    On Error GoTo Failed
    CloseKnownPopulationSdAlert False
    Exit Sub
Failed:
    ReportError "modApp.OnUseUnknownSdClick"
End Sub

Public Sub OnConfirmKnownSdClick(ByVal e As clsEvent)
    On Error GoTo Failed
    CloseKnownPopulationSdAlert True
    Exit Sub
Failed:
    ReportError "modApp.OnConfirmKnownSdClick"
End Sub

Public Sub OnKnownPopulationSdKeydown(ByVal e As clsEvent)
    On Error GoTo Failed
    CycleTwoButtons e, "#useUnknownSdBtn", "#confirmKnownSdBtn"
    Exit Sub
Failed:
    ReportError "modApp.OnKnownPopulationSdKeydown"
End Sub
