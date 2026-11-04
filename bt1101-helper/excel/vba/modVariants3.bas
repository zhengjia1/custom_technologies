Attribute VB_Name = "modVariants3"
' Answer options, part 3: prose variation (website app-src.js lines 117-200): the phrase
' pools and proseAnswerVariants. Each pool's choices are drawn without replacement, so the
' five variations use different wordings; decisions, numbers and comparisons are kept.
Option Explicit
Option Compare Binary

Private mReady As Boolean
Private mPoolCount As Long
Private mPoolPattern() As String
Private mPoolFlags() As String
Private mPoolLiteral() As String     ' text every match contains ("" = always run the regex)
Private mPoolChoice() As String      ' (pool, 0..4)

Private Const PROTECTED_ENDING As String = "There is no statistically significant correlation between the two variables."

Private Sub EnsureData()
    If mReady Then Exit Sub
    ReDim mPoolPattern(0 To 79): ReDim mPoolFlags(0 To 79): ReDim mPoolLiteral(0 To 79)
    ReDim mPoolChoice(0 To 79, 0 To 4)
    InitProsePools
    mReady = True
End Sub

Private Sub AddProsePool(ByVal pattern As String, ByVal flags As String, ByVal literal As String, _
    ByVal c1 As String, ByVal c2 As String, ByVal c3 As String, ByVal c4 As String, ByVal c5 As String)
    mPoolPattern(mPoolCount) = pattern
    mPoolFlags(mPoolCount) = flags
    mPoolLiteral(mPoolCount) = literal
    mPoolChoice(mPoolCount, 0) = c1
    mPoolChoice(mPoolCount, 1) = c2
    mPoolChoice(mPoolCount, 2) = c3
    mPoolChoice(mPoolCount, 3) = c4
    mPoolChoice(mPoolCount, 4) = c5
    mPoolCount = mPoolCount + 1
End Sub

' proseAnswerVariants(value): six versions, entry 0 = the canonical prose. history is
' accepted for symmetry with the code generator (the website's prose generator ignores it).
Public Function VarProseAnswerVariants(ByVal value As String, Optional ByVal history As clsList = Nothing) As clsList
    Dim variants As clsList, ending As String, body As String, choices() As String, pool() As String
    Dim intro(0 To 9) As String, index As Long, rule As Long, k As Long, text As String, layout As Long, candidate As String
    EnsureData
    Set variants = NewList()
    If Len(JsTrim(value)) = 0 Then
        For index = 0 To 5
            variants.Push value
        Next index
        Set VarProseAnswerVariants = variants
        Exit Function
    End If
    If Right$(value, Len(PROTECTED_ENDING)) = PROTECTED_ENDING Then ending = PROTECTED_ENDING
    body = Left$(value, Len(value) - Len(ending))
    ' Independent phrase choices, without replacement.
    ReDim choices(0 To mPoolCount - 1, 0 To 4)
    ReDim pool(0 To 4)
    For rule = 0 To mPoolCount - 1
        For k = 0 To 4
            pool(k) = mPoolChoice(rule, k)
        Next k
        VarShuffleStrings pool, 0, 4
        For k = 0 To 4
            choices(rule, k) = pool(k)
        Next k
    Next rule
    variants.Push value
    intro(0) = "Interpretation": intro(1) = "Reading of the results": intro(2) = "Statistical interpretation"
    intro(3) = "Results explained": intro(4) = "Interpretation of the output": intro(5) = "Summary of the findings"
    intro(6) = "What the results indicate": intro(7) = "Analysis of the results": intro(8) = "Reported findings"
    intro(9) = "Result interpretation"
    VarShuffleStrings intro, 0, 9
    For index = 0 To 4
        text = body
        For rule = 0 To mPoolCount - 1
            If Len(mPoolLiteral(rule)) = 0 Then
                text = Rx(mPoolPattern(rule), mPoolFlags(rule)).Replace(text, choices(rule, index))
            ElseIf InStr(1, text, mPoolLiteral(rule), vbBinaryCompare) > 0 Then
                text = Rx(mPoolPattern(rule), mPoolFlags(rule)).Replace(text, choices(rule, index))
            End If
        Next rule
        layout = VarRandomIndex(3)
        If layout = 1 Then text = Rx("([.!?]) +(?=[A-Z][a-z])", "g").Replace(text, "$1" & vbLf)
        If layout = 2 Then text = Rx("\n{2,}", "g").Replace(text, vbLf)
        candidate = text & ending
        If candidate = value Or ListHasText(variants, candidate) Then candidate = intro(index) & ":" & vbLf & candidate
        variants.Push candidate
    Next index
    Set VarProseAnswerVariants = variants
End Function

Private Function ListHasText(ByVal list As clsList, ByVal text As String) As Boolean
    Dim i As Long
    For i = 0 To list.Count - 1
        If CStr(list(i)) = text Then ListHasText = True: Exit Function
    Next i
End Function

' BEGIN GENERATED DATA (node test/variants/gen-data.js prose)
Private Sub InitProsePools()
    mPoolCount = 0
    AddProsePool "\bwe do not have sufficient evidence to reject the null hypothesis\b", "g", "we do not have sufficient evidence to reject the null hypothesis", _
        "there is insufficient evidence to reject the null hypothesis", "the evidence does not justify rejecting the null hypothesis", "we lack sufficient evidence to reject the null hypothesis", _
        "the evidence is insufficient for rejection of the null hypothesis", "the results do not provide enough evidence to reject the null hypothesis"
    AddProsePool "\bwe have sufficient evidence to reject the null hypothesis\b", "g", "we have sufficient evidence to reject the null hypothesis", _
        "there is sufficient evidence to reject the null hypothesis", "the evidence justifies rejecting the null hypothesis", "we have enough evidence to reject the null hypothesis", _
        "the evidence is sufficient for rejection of the null hypothesis", "the results provide enough evidence to reject the null hypothesis"
    AddProsePool "\bThe excluded (variables|variable)\b", "g", "The excluded ", _
        "The $1 excluded from the restricted model", "The additional $1 in the full model", "The $1 omitted from the restricted model", _
        "The $1 added to the full model", "The $1 present only in the full model"
    AddProsePool "\bOptimal production is\b", "g", "Optimal production is", _
        "At the optimum, production is", "The optimal production quantities are", "The production plan at the optimum is", _
        "The optimal production plan is", "Production at the optimal solution is"
    AddProsePool "\bThe shadow price of the\b", "g", "The shadow price of the", _
        "The shadow price associated with the", "The reported shadow price for the", "The shadow-price value for the", _
        "The calculated shadow price for the", "The shadow price corresponding to the"
    AddProsePool "\bThis means that\b", "g", "This means that", _
        "This indicates that", "The interpretation is that", "In other words,", _
        "This tells us that", "Equivalently,"
    AddProsePool "\bThis applies provided\b", "g", "This applies provided", _
        "This holds as long as", "This interpretation applies if", "This remains valid provided", _
        "This holds provided that", "This statement applies as long as"
    AddProsePool "\bThe reduced cost of\b", "g", "The reduced cost of", _
        "The reduced cost associated with", "The reported reduced cost for", "The calculated reduced cost for", _
        "The reduced-cost value for", "The reduced cost corresponding to"
    AddProsePool "\bThe objective coefficient can vary between\b", "g", "The objective coefficient can vary between", _
        "The objective coefficient may range between", "The objective coefficient can lie between", "The allowable range for the objective coefficient is between", _
        "The objective coefficient may take values between", "The objective coefficient can take any value between"
    AddProsePool "\bSlack values \(unused resource\) are:", "g", "Slack values (unused resource) are:", _
        "The unused-resource quantities (slack) are:", "Unused resource, measured as slack, is:", "The resource slack values are:", _
        "Slack, representing unused resources, is:", "The remaining resource quantities (slack) are:"
    AddProsePool "\bA minimum sample size of (\d+) is required to\b", "g", "A minimum sample size of ", _
        "At least $1 observations are required to", "A sample of at least $1 observations is needed to", "The minimum required sample size is $1 to", _
        "A sample size no smaller than $1 is required to", "At least $1 observations are needed to"
    AddProsePool "\bThe coefficient of skewness is\b", "g", "The coefficient of skewness is", _
        "The skewness coefficient is", "Skewness has a coefficient of", "The reported skewness is", _
        "The calculated skewness is", "The skewness value is"
    AddProsePool "\bThis implies that,", "g", "This implies that,", _
        "This means that,", "The interpretation is that,", "This indicates that,", _
        "The estimate implies that,", "In terms of the fitted relationship,"
    AddProsePool "\bThe regression coefficient estimate of\b", "g", "The regression coefficient estimate of", _
        "The estimated regression coefficient for", "The fitted regression coefficient for", "The regression estimate for the coefficient of", _
        "The estimate of the regression coefficient for", "The reported regression coefficient for"
    AddProsePool "\bThis shows that approximately\b", "g", "This shows that approximately", _
        "This indicates that about", "This means that approximately", "The result indicates that around", _
        "This tells us that approximately", "The interpretation is that about"
    AddProsePool "\bThe Multiple R-squared value is\b", "g", "The Multiple R-squared value is", _
        "Multiple R-squared is", "The reported multiple R-squared is", "The model's multiple R-squared is", _
        "The multiple R-squared statistic is", "The value of multiple R-squared is"
    AddProsePool "\bThe coefficient for\b", "g", "The coefficient for", _
        "The estimated coefficient for", "The fitted coefficient for", "The model coefficient for", _
        "The reported coefficient for", "The coefficient estimate for"
    AddProsePool "\bThe odds ratio is\b", "g", "The odds ratio is", _
        "The estimated odds ratio is", "The corresponding odds ratio is", "The odds ratio is calculated as", _
        "The resulting odds ratio is", "The fitted odds ratio is"
    AddProsePool "\bThis corresponds to a\b", "g", "This corresponds to a", _
        "This is equivalent to a", "This represents a", "In percentage terms, this is a", _
        "This amounts to a", "Equivalently, this is a"
    AddProsePool "\bThe model correctly classified\b", "g", "The model correctly classified", _
        "The model's correct classifications comprised", "The correctly classified observations were", "Correct classifications consisted of", _
        "The model classified correctly", "The classifier correctly identified"
    AddProsePool "\bIt incorrectly classified\b", "g", "It incorrectly classified", _
        "It misclassified", "The model misclassified", "The incorrect classifications were", _
        "Its classification errors were", "The model incorrectly labelled"
    AddProsePool "\bThe accuracy is\b", "g", "The accuracy is", _
        "The model's accuracy is", "Overall accuracy is", "The accuracy score is", _
        "Classification accuracy is", "The proportion classified correctly is"
    AddProsePool "\bThe sensitivity is\b", "g", "The sensitivity is", _
        "The model's sensitivity is", "Sensitivity is", "The true-positive rate is", _
        "The sensitivity score is", "The recall for the positive class is"
    AddProsePool "\bThe specificity is\b", "g", "The specificity is", _
        "The model's specificity is", "Specificity is", "The true-negative rate is", _
        "The specificity score is", "The rate of correctly identifying actual negatives is"
    AddProsePool "\bThe precision is\b", "g", "The precision is", _
        "The model's precision is", "Precision is", "The positive predictive value is", _
        "The precision score is", "The proportion of positive predictions that are correct is"
    AddProsePool "\bmeaning that\b", "g", "meaning that", _
        "which means that", "indicating that", "showing that", _
        "that is,", "in other words,"
    AddProsePool "\bexplains about\b", "g", "explains about", _
        "accounts for approximately", "explains approximately", "accounts for about", _
        "captures approximately", "captures about"
    AddProsePool "\bCI for the population mean\b", "g", "CI for the population mean", _
        "confidence interval for the population mean", "CI estimating the population mean", "confidence interval estimating the population mean", _
        "CI for the true population mean", "confidence interval for the true population mean"
    AddProsePool "\bMore generally, with repeated random sampling\b", "g", "More generally, with repeated random sampling", _
        "Under repeated random sampling", "More broadly, if we repeatedly take random samples", "In repeated random sampling", _
        "Considering repeated random sampling", "Over repeated random sampling"
    AddProsePool "\bthus it is significant in explaining\b", "g", "thus it is significant in explaining", _
        "so it significantly explains", "hence it is significant in explaining", "therefore it significantly explains", _
        "so it makes a statistically significant contribution to explaining", "thus it has a significant role in explaining"
    AddProsePool "^(PC\d+) = ", "gm", "", _
        "$1 is calculated as ", "The score for $1 is ", "For $1, use ", _
        "$1 has the equation: $1 = ", "$1 is given by "
    AddProsePool "H" & ChrW$(&H2080) & ":\s*", "g", "H" & ChrW$(&H2080) & ":", _
        "H" & ChrW$(&H2080) & " (null hypothesis): ", "Null hypothesis, H" & ChrW$(&H2080) & ": ", "H" & ChrW$(&H2080) & " states: ", _
        "The null hypothesis (H" & ChrW$(&H2080) & ") is: ", "Under H" & ChrW$(&H2080) & ": "
    AddProsePool "H" & ChrW$(&H2081) & ":\s*", "g", "H" & ChrW$(&H2081) & ":", _
        "H" & ChrW$(&H2081) & " (alternative hypothesis): ", "Alternative hypothesis, H" & ChrW$(&H2081) & ": ", "H" & ChrW$(&H2081) & " states: ", _
        "The alternative hypothesis (H" & ChrW$(&H2081) & ") is: ", "Under H" & ChrW$(&H2081) & ": "
    AddProsePool "\bSince the p-value is\b", "g", "Since the p-value is", _
        "As the p-value is", "Because the p-value is", "Given that the p-value is", _
        "Since the reported p-value is", "As the observed p-value is"
    AddProsePool "\bThere is sufficient evidence to conclude that\b", "g", "There is sufficient evidence to conclude that", _
        "The evidence is sufficient to conclude that", "There is enough evidence to conclude that", "The results provide sufficient evidence that", _
        "The evidence supports the conclusion that", "There is sufficient support for the conclusion that"
    AddProsePool "\bThere is insufficient evidence to conclude that\b", "g", "There is insufficient evidence to conclude that", _
        "The evidence is insufficient to conclude that", "There is not enough evidence to conclude that", "The results do not provide sufficient evidence that", _
        "The evidence does not support the conclusion that", "There is insufficient support for the conclusion that"
    AddProsePool "\bwe fail to reject the null hypothesis\b", "g", "we fail to reject the null hypothesis", _
        "we do not reject the null hypothesis", "the null hypothesis is not rejected", "we cannot reject the null hypothesis", _
        "we lack sufficient evidence to reject the null hypothesis", "the decision is to not reject the null hypothesis"
    AddProsePool "\bwe reject the null hypothesis\b", "g", "we reject the null hypothesis", _
        "the null hypothesis is rejected", "our decision is to reject the null hypothesis", "we therefore reject the null hypothesis", _
        "rejection of the null hypothesis is warranted", "the test leads us to reject the null hypothesis"
    AddProsePool "\bwas conducted to\b", "g", "was conducted to", _
        "was performed to", "was carried out to", "was used to", _
        "was undertaken to", "was applied to"
    AddProsePool "\bwas conducted between\b", "g", "was conducted between", _
        "was performed between", "was carried out between", "was run between", _
        "was used to assess the relationship between", "was undertaken between"
    AddProsePool "\bThe test statistic is\b", "g", "The test statistic is", _
        "The observed test statistic is", "The test gives", "The calculated test statistic is", _
        "The reported test statistic is", "The value of the test statistic is"
    AddProsePool "\bwith a p-value of\b", "g", "with a p-value of", _
        "with p-value", "and the p-value is", "with an associated p-value of", _
        "and its p-value is", "with a reported p-value of"
    AddProsePool "\bAn increase of 1 unit in\b", "g", "An increase of 1 unit in", _
        "A 1-unit increase in", "A one-unit rise in", "An increase of one unit in", _
        "A rise of 1 unit in", "An upward change of 1 unit in"
    AddProsePool "\bwhile holding (all )?other predictors constant\b", "g", "while holding ", _
        "with the other predictors held constant", "while keeping the other predictors unchanged", "when all other predictors are held fixed", _
        "holding the remaining predictors constant", "with no change in the other predictors"
    AddProsePool "\bwas not statistically significant\b", "g", "was not statistically significant", _
        "did not reach statistical significance", "was statistically non-significant", "was not significant statistically", _
        "did not attain statistical significance", "did not show statistical significance"
    AddProsePool "\bwas statistically significant\b", "g", "was statistically significant", _
        "reached statistical significance", "attained statistical significance", "was significant statistically", _
        "showed statistical significance", "met the criterion for statistical significance"
    AddProsePool "\bTherefore,\s*", "g", "Therefore,", _
        "Thus, ", "Consequently, ", "Accordingly, ", _
        "It follows that ", "Hence, "
    AddProsePool "\bThe optimal solution is\b", "g", "The optimal solution is", _
        "The solution at the optimum is", "The optimum is achieved with", "The optimal allocation is", _
        "The quantities at the optimum are", "The optimal decision is"
    AddProsePool "\bThe optimal objective value is\b", "g", "The optimal objective value is", _
        "The objective value at the optimum is", "The optimum objective value is", "The objective achieves an optimal value of", _
        "At the optimum, the objective value is", "The optimal value of the objective is"
    AddProsePool "\bis fully utilised\b", "g", "is fully utilised", _
        "is used in full", "has no unused capacity", "is entirely utilised", _
        "has its full capacity utilised", "has all of its capacity used"
    AddProsePool "\bis not fully utilised\b", "g", "is not fully utilised", _
        "has unused capacity", "is not used in full", "has capacity remaining", _
        "has some unused capacity", "has some capacity left unused"
    AddProsePool "\bwhich is below\b", "g", "which is below", _
        "which is less than", "a value below", "which falls below", _
        "a value less than", "which lies below"
    AddProsePool "\bwhich is above\b", "g", "which is above", _
        "which is greater than", "a value above", "which exceeds", _
        "a value greater than", "which lies above"
    AddProsePool "\bThe distribution is\b", "g", "The distribution is", _
        "This distribution is", "The distribution can be described as", "The distribution is classified as", _
        "Its shape is described as", "The distribution's classification is"
    AddProsePool "\bThe model explains\b", "g", "The model explains", _
        "The model accounts for", "The fitted model explains", "This model explains", _
        "The fitted model accounts for", "The regression model explains"
End Sub
' END GENERATED DATA
