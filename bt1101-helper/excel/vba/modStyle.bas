Attribute VB_Name = "modStyle"
' CSS cascade for the renderer. The website's stylesheet arrives pre-processed by
' tools/css-rules.js (rule table text: modBoot.StoredText("css") in the release, the
' .build/css-rules.txt file in tests). Rules are bucketed by their subject key and
' matched with clsNode.Matches; computed styles live in numbered "slots" (one per element
' or pseudo-element) holding an interned value index per property, and are cached per
' node until the node, its parent, the theme, the viewport or the container width change.
'
' Use: StyleLoad text (once) -> per layout pass StyleBegin width, height -> StyleOf(node),
' PseudoOf(node, 1|2) -> SV/SLen/SColour/SFont... -> StyleEnd (frees unused slots).
Option Explicit

' Property ids: the order of PROPERTIES in tools/css-rules.js (checked on load).
Public Const P_DISPLAY As Long = 0, P_POSITION As Long = 1, P_TOP As Long = 2, P_RIGHT As Long = 3
Public Const P_BOTTOM As Long = 4, P_LEFT As Long = 5, P_ZINDEX As Long = 6, P_WIDTH As Long = 7
Public Const P_HEIGHT As Long = 8, P_MINW As Long = 9, P_MINH As Long = 10, P_MAXW As Long = 11
Public Const P_MAXH As Long = 12, P_BOXSIZING As Long = 13
Public Const P_MT As Long = 14, P_MR As Long = 15, P_MB As Long = 16, P_ML As Long = 17
Public Const P_PT As Long = 18, P_PR As Long = 19, P_PB As Long = 20, P_PL As Long = 21
Public Const P_BTW As Long = 22, P_BRW As Long = 23, P_BBW As Long = 24, P_BLW As Long = 25
Public Const P_BTS As Long = 26, P_BRS As Long = 27, P_BBS As Long = 28, P_BLS As Long = 29
Public Const P_BTC As Long = 30, P_BRC As Long = 31, P_BBC As Long = 32, P_BLC As Long = 33
Public Const P_COLOR As Long = 34, P_BG As Long = 35, P_OPACITY As Long = 36, P_VISIBILITY As Long = 37
Public Const P_FONTSIZE As Long = 38, P_FONTWEIGHT As Long = 39, P_FONTSTYLE As Long = 40
Public Const P_LINEHEIGHT As Long = 41, P_TEXTALIGN As Long = 42, P_TEXTTRANSFORM As Long = 43
Public Const P_WHITESPACE As Long = 44, P_OVERFLOWWRAP As Long = 45, P_WORDBREAK As Long = 46
Public Const P_LETTERSPACING As Long = 47, P_LISTSTYLE As Long = 48, P_OVERFLOWX As Long = 49
Public Const P_OVERFLOWY As Long = 50, P_FLEXDIR As Long = 51, P_FLEXWRAP As Long = 52
Public Const P_FLEXGROW As Long = 53, P_FLEXSHRINK As Long = 54, P_FLEXBASIS As Long = 55
Public Const P_ALIGNITEMS As Long = 56, P_ALIGNSELF As Long = 57, P_ALIGNCONTENT As Long = 58
Public Const P_JUSTIFYCONTENT As Long = 59, P_JUSTIFYITEMS As Long = 60, P_JUSTIFYSELF As Long = 61
Public Const P_ROWGAP As Long = 62, P_COLGAP As Long = 63, P_GRIDCOLS As Long = 64
Public Const P_GRIDCOLSTART As Long = 65, P_GRIDCOLEND As Long = 66, P_CONTAINER As Long = 67
Public Const P_CONTENT As Long = 68, P_TRANSFORM As Long = 69, P_CLIP As Long = 70
Public Const P_BORDERCOLLAPSE As Long = 71, P_BORDERSPACING As Long = 72, P_TEXTOVERFLOW As Long = 73
Public Const P_GRIDROWS As Long = 74
' Paint only: zero-blur box shadows and focus outlines (a rule table without them loads too).
Public Const P_BOXSHADOW As Long = 75, P_OUTLINEW As Long = 76, P_OUTLINES As Long = 77, P_OUTLINEC As Long = 78
Public Const P_OUTLINEOFF As Long = 79
Public Const NP As Long = 80
Private Const PROPERTY_NAMES As String = "display position top right bottom left z-index width height min-width min-height max-width max-height box-sizing " & _
    "margin-top margin-right margin-bottom margin-left padding-top padding-right padding-bottom padding-left " & _
    "border-top-width border-right-width border-bottom-width border-left-width border-top-style border-right-style border-bottom-style border-left-style " & _
    "border-top-color border-right-color border-bottom-color border-left-color color background-color opacity visibility " & _
    "font-size font-weight font-style line-height text-align text-transform white-space overflow-wrap word-break letter-spacing list-style-type " & _
    "overflow-x overflow-y flex-direction flex-wrap flex-grow flex-shrink flex-basis align-items align-self align-content justify-content justify-items justify-self " & _
    "row-gap column-gap grid-template-columns grid-column-start grid-column-end container-type content transform clip border-collapse border-spacing text-overflow grid-template-rows " & _
    "box-shadow outline-width outline-style outline-color outline-offset"
Private Const INITIAL_VALUES As String = "inline|static|auto|auto|auto|auto|auto|auto|auto|auto|auto|none|none|content-box|" & _
    "0|0|0|0|0|0|0|0|medium|medium|medium|medium|none|none|none|none|currentcolor|currentcolor|currentcolor|currentcolor|" & _
    "#000000|transparent|1|visible|16px|400|normal|normal|start|none|normal|normal|normal|normal|none|visible|visible|" & _
    "row|nowrap|0|1|auto|normal|auto|normal|normal|normal|auto|normal|normal|none|auto|auto|normal|normal|none|auto|separate|0|clip|none|" & _
    "none|medium|none|currentcolor|0"

' Value kinds.
Public Const VK_WORD As Long = 0, VK_PX As Long = 1, VK_PCT As Long = 2, VK_EM As Long = 3
Public Const VK_EXPR As Long = 4, VK_NUM As Long = 5

' ---------------------------------------------------------------- loaded tables
Private mLoaded As Boolean
Private mInherited(0 To NP - 1) As Boolean
Private mInitial(0 To NP - 1) As Long

' interned values
Private mVCount As Long
Private mVText() As String
Private mVKind() As Long
Private mVNum() As Double
Private mVColState() As Long        ' 0 unparsed, 1 colour, 2 none/transparent, 3 currentcolor
Private mVRgb() As Long
Private mVAlpha() As Double
Private mVMap As clsMap
Private mVKeyword() As Long         ' 0 none, 1 inherit, 2 unset, 3 initial/revert
' candidate rules per element signature (tag#id.classes), merged and sorted once
Private mSigMap As clsMap
Private mSigFirst() As Long, mSigLen() As Long, mSigCount As Long
Private mSigRules() As Long, mSigRuleCount As Long

' declarations and blocks
Private mDCount As Long
Private mDProp() As Long
Private mDImp() As Boolean
Private mDLight() As Long
Private mDDark() As Long
Private mBCount As Long
Private mBFirst() As Long
Private mBLen() As Long

' rules
Private mRCount As Long
Private mRBlock() As Long
Private mRCond() As Long
Private mRPseudo() As Long
Private mRSimple() As Boolean
Private mRSel() As String
Private mRDead() As Boolean

' conditions
Private mCCount As Long
Private mCParent() As Long
Private mCIsContainer() As Boolean
Private mCAnyContainer() As Boolean   ' the condition or an outer one is a container query
Private mCAltFirst() As Long
Private mCAltLen() As Long
Private mCMediaOk() As Boolean        ' per environment (media part only)
Private mCBit() As Long               ' bit of a container condition
Private mAltCount As Long
Private mAltMin() As Double
Private mAltMax() As Double
Private mAltNever() As Boolean

' buckets
Private mBucketMap As clsMap
Private mBkCount As Long
Private mBkFirst() As Long
Private mBkLen() As Long
Private mBkRules() As Long

' glyph coverage of the text font
Private mGlyphCount As Long
Private mGlyphLo() As Long
Private mGlyphHi() As Long

' ---------------------------------------------------------------- environment and slots
Private mViewW As Double, mViewH As Double
Private mDark As Boolean
Private mEnv As Long
Private mPass As Long
Private mPassStart As Long
Private mSerialSeq As Long

Private mSCap As Long
Private mSUsed As Long
Private mFree() As Long
Private mFreeCount As Long
Private mSInUse() As Boolean
Private mSNode() As clsNode
Private mSUid() As Long
Private mSPseudo() As Long
Private mSNodeVer() As Long
Private mSParentVer() As Long
Private mSParent() As Long
Private mSParentSerial() As Long
Private mSSerial() As Long
Private mSLay() As Long             ' layout stamp: changes when a layout-relevant value changed
Private mPaintOnly(0 To NP - 1) As Boolean
Private mEpoch As Long              ' changes when every computed style may have changed
Private mSEnv() As Long
Private mSPass() As Long
Private mSBefore() As Long
Private mSAfter() As Long
Private mSContainer() As Long       ' uid of the nearest size-container ancestor (0 none)
Private mSCondBits() As Long        ' container-condition bits used (-1 = none used)
Private mSFont() As Double
Private mSWeight() As Long
Private mVal() As Long

Private mUidSlot() As Long
Private mUidCap As Long
Private mContW() As Double          ' container content width by uid (-1 unknown)

' scratch for cascading
Private mNorm(0 To NP - 1) As Long
Private mImp(0 To NP - 1) As Long
Private mOld(0 To NP - 1) As Long
Private mCand() As Long
Private mCandCount As Long

' ================================================================ loading

Public Function StyleLoaded() As Boolean
    StyleLoaded = mLoaded
End Function

' Loads the rule table produced by tools/css-rules.js.
Public Sub StyleLoad(ByVal text As String)
    Dim lines() As String, i As Long, f() As String, names() As String, propIndex As Long
    Dim blockId As Long, initials() As String, keyIndex As Variant, k As Long
    Dim alts() As String, a As Long, parts() As String, ruleKeys() As Long, containerCount As Long
    ResetTables
    ReDim ruleKeys(0 To 255)
    names = Split(PROPERTY_NAMES, " ")
    initials = Split(INITIAL_VALUES, "|")
    If UBound(names) <> NP - 1 Or UBound(initials) <> NP - 1 Then Err.Raise vbObjectError + 801, "modStyle", "Property tables are inconsistent"
    For i = 0 To NP - 1
        mInitial(i) = Intern(initials(i))
    Next i
    mInherited(P_COLOR) = True: mInherited(P_FONTSIZE) = True: mInherited(P_FONTWEIGHT) = True
    mInherited(P_FONTSTYLE) = True: mInherited(P_LINEHEIGHT) = True: mInherited(P_TEXTALIGN) = True
    mInherited(P_TEXTTRANSFORM) = True: mInherited(P_WHITESPACE) = True: mInherited(P_OVERFLOWWRAP) = True
    mInherited(P_WORDBREAK) = True: mInherited(P_LETTERSPACING) = True: mInherited(P_LISTSTYLE) = True
    mInherited(P_VISIBILITY) = True
    text = Replace(text, vbCr, "")
    lines = Split(text, vbLf)
    For i = 0 To UBound(lines)
        If Len(lines(i)) > 0 Then
            f = Split(lines(i), vbTab)
            Select Case f(0)
                Case "V"
                    If Val(f(1)) <> 2 Then Err.Raise vbObjectError + 802, "modStyle", "Unsupported rule table version " & f(1)
                Case "P"
                    If propIndex > NP - 1 Then Err.Raise vbObjectError + 803, "modStyle", "Unexpected property " & f(1)
                    If f(1) <> names(propIndex) Then Err.Raise vbObjectError + 803, "modStyle", "Property " & propIndex & " is " & f(1) & ", expected " & names(propIndex)
                    propIndex = propIndex + 1
                Case "C"
                    k = CLng(f(1))
                    GrowConditions k
                    mCParent(k) = CLng(f(2))
                    mCIsContainer(k) = (f(3) = "c")
                    mCAnyContainer(k) = mCIsContainer(k)
                    If mCParent(k) > 0 Then If mCAnyContainer(mCParent(k)) Then mCAnyContainer(k) = True
                    alts = Split(f(4), ";")
                    mCAltFirst(k) = mAltCount
                    mCAltLen(k) = UBound(alts) + 1
                    For a = 0 To UBound(alts)
                        GrowAlts
                        If alts(a) = "x" Then
                            mAltNever(mAltCount) = True
                        Else
                            parts = Split(alts(a), ":")
                            mAltMin(mAltCount) = Val(parts(0))
                            mAltMax(mAltCount) = Val(parts(1))
                        End If
                        mAltCount = mAltCount + 1
                    Next a
                    If mCIsContainer(k) Then
                        If containerCount > 29 Then Err.Raise vbObjectError + 805, "modStyle", "Too many container queries"
                        mCBit(k) = 2 ^ containerCount
                        containerCount = containerCount + 1
                    End If
                Case "B"
                    blockId = CLng(f(1))
                    GrowBlocks blockId
                    mBFirst(blockId) = mDCount
                    mBLen(blockId) = 0
                Case "D"
                    GrowDecls
                    mDProp(mDCount) = CLng(f(1))
                    mDImp(mDCount) = (f(2) = "1")
                    mDLight(mDCount) = Intern(Unescape(f(3)))
                    If UBound(f) >= 4 Then mDDark(mDCount) = Intern(Unescape(f(4))) Else mDDark(mDCount) = mDLight(mDCount)
                    mDCount = mDCount + 1
                    mBLen(blockId) = mBLen(blockId) + 1
                Case "R"
                    GrowRules
                    mRBlock(mRCount) = CLng(f(1))
                    mRCond(mRCount) = CLng(f(2))
                    mRPseudo(mRCount) = CLng(f(3))
                    mRSimple(mRCount) = (f(6) = "1")
                    mRSel(mRCount) = Unescape(f(7))
                    keyIndex = mBucketMap(f(5))
                    If IsEmpty(keyIndex) Then
                        mBucketMap(f(5)) = mBkCount
                        GrowBuckets
                        mBkCount = mBkCount + 1
                        keyIndex = mBkCount - 1
                    End If
                    mBkLen(keyIndex) = mBkLen(keyIndex) + 1
                    If mRCount > UBound(ruleKeys) Then ReDim Preserve ruleKeys(0 To mRCount * 2)
                    ruleKeys(mRCount) = keyIndex
                    mRCount = mRCount + 1
                Case "G"
                    LoadGlyphs f(1)
            End Select
        End If
    Next i
    ' Tables from before the paint-only properties were added list a prefix of them.
    If propIndex < P_BOXSHADOW Then Err.Raise vbObjectError + 803, "modStyle", "The rule table lists " & propIndex & " properties, expected " & NP
    ' Fill the buckets (rules stay in cascade order inside each bucket).
    ReDim mBkRules(0 To mRCount)
    k = 0
    For i = 0 To mBkCount - 1
        mBkFirst(i) = k
        k = k + mBkLen(i)
        mBkLen(i) = 0
    Next i
    ReDim mRDead(0 To mRCount)
    For i = 0 To mRCount - 1
        k = ruleKeys(i)
        mBkRules(mBkFirst(k) + mBkLen(k)) = i
        mBkLen(k) = mBkLen(k) + 1
    Next i
    mLoaded = True
    mEnv = mEnv + 1
    mEpoch = mEpoch + 1
    ' Properties only the paint pass reads (a change does not invalidate cached layout).
    For i = 0 To NP - 1
        mPaintOnly(i) = False
    Next i
    mPaintOnly(P_BG) = True: mPaintOnly(P_OPACITY) = True: mPaintOnly(P_VISIBILITY) = True: mPaintOnly(P_ZINDEX) = True
    For i = P_BTC To P_BLC
        mPaintOnly(i) = True
    Next i
    For i = P_BOXSHADOW To P_OUTLINEOFF
        mPaintOnly(i) = True
    Next i
End Sub

Private Sub LoadGlyphs(ByVal text As String)
    Dim ranges() As String, i As Long, parts() As String
    If Len(text) = 0 Then Exit Sub
    ranges = Split(text, ",")
    mGlyphCount = UBound(ranges) + 1
    ReDim mGlyphLo(0 To mGlyphCount - 1)
    ReDim mGlyphHi(0 To mGlyphCount - 1)
    For i = 0 To mGlyphCount - 1
        parts = Split(ranges(i), "-")
        mGlyphLo(i) = CLng("&H" & parts(0))
        mGlyphHi(i) = CLng("&H" & parts(1))
    Next i
End Sub

' True when the text font has a glyph for the code point (always True without coverage data).
Public Function GlyphCovered(ByVal code As Long) As Boolean
    Dim lo As Long, hi As Long, mid As Long
    If mGlyphCount = 0 Or code < 128 Then GlyphCovered = True: Exit Function
    lo = 0: hi = mGlyphCount - 1
    Do While lo <= hi
        mid = (lo + hi) \ 2
        If code < mGlyphLo(mid) Then
            hi = mid - 1
        ElseIf code > mGlyphHi(mid) Then
            lo = mid + 1
        Else
            GlyphCovered = True
            Exit Function
        End If
    Loop
End Function

Private Function Unescape(ByVal text As String) As String
    If InStr(1, text, "\") = 0 Then Unescape = text: Exit Function
    text = Replace(text, "\\", ChrW$(1))
    text = Replace(text, "\t", vbTab)
    text = Replace(text, "\n", vbLf)
    Unescape = Replace(text, ChrW$(1), "\")
End Function

Private Sub ResetTables()
    mLoaded = False
    mVCount = 0
    ReDim mVText(0 To 255): ReDim mVKind(0 To 255): ReDim mVNum(0 To 255)
    ReDim mVColState(0 To 255): ReDim mVRgb(0 To 255): ReDim mVAlpha(0 To 255)
    ReDim mVKeyword(0 To 255)
    Set mVMap = NewMap()
    Set mSigMap = NewMap()
    mSigCount = 0: ReDim mSigFirst(0 To 63): ReDim mSigLen(0 To 63)
    mSigRuleCount = 0: ReDim mSigRules(0 To 1023)
    mDCount = 0: ReDim mDProp(0 To 255): ReDim mDImp(0 To 255): ReDim mDLight(0 To 255): ReDim mDDark(0 To 255)
    mBCount = 0: ReDim mBFirst(0 To 63): ReDim mBLen(0 To 63)
    mRCount = 0: ReDim mRBlock(0 To 255): ReDim mRCond(0 To 255): ReDim mRPseudo(0 To 255)
    ReDim mRSimple(0 To 255): ReDim mRSel(0 To 255)
    mCCount = 0: ReDim mCParent(0 To 15): ReDim mCIsContainer(0 To 15): ReDim mCAnyContainer(0 To 15)
    ReDim mCAltFirst(0 To 15): ReDim mCAltLen(0 To 15): ReDim mCMediaOk(0 To 15): ReDim mCBit(0 To 15)
    mAltCount = 0: ReDim mAltMin(0 To 15): ReDim mAltMax(0 To 15): ReDim mAltNever(0 To 15)
    Set mBucketMap = NewMap()
    mBkCount = 0: ReDim mBkFirst(0 To 63): ReDim mBkLen(0 To 63)
    mGlyphCount = 0
    ReDim mCand(0 To 63)
    ResetSlots
End Sub

Private Sub GrowConditions(ByVal k As Long)
    If k > mCCount Then mCCount = k
    If k > UBound(mCParent) Then
        ReDim Preserve mCParent(0 To k * 2): ReDim Preserve mCIsContainer(0 To k * 2)
        ReDim Preserve mCAnyContainer(0 To k * 2): ReDim Preserve mCAltFirst(0 To k * 2)
        ReDim Preserve mCAltLen(0 To k * 2): ReDim Preserve mCMediaOk(0 To k * 2): ReDim Preserve mCBit(0 To k * 2)
    End If
End Sub

Private Sub GrowAlts()
    If mAltCount > UBound(mAltMin) Then
        ReDim Preserve mAltMin(0 To mAltCount * 2): ReDim Preserve mAltMax(0 To mAltCount * 2)
        ReDim Preserve mAltNever(0 To mAltCount * 2)
    End If
End Sub

Private Sub GrowBlocks(ByVal k As Long)
    If k > mBCount Then mBCount = k
    If k > UBound(mBFirst) Then ReDim Preserve mBFirst(0 To k * 2): ReDim Preserve mBLen(0 To k * 2)
End Sub

Private Sub GrowDecls()
    If mDCount > UBound(mDProp) Then
        ReDim Preserve mDProp(0 To mDCount * 2): ReDim Preserve mDImp(0 To mDCount * 2)
        ReDim Preserve mDLight(0 To mDCount * 2): ReDim Preserve mDDark(0 To mDCount * 2)
    End If
End Sub

Private Sub GrowRules()
    If mRCount > UBound(mRBlock) Then
        ReDim Preserve mRBlock(0 To mRCount * 2): ReDim Preserve mRCond(0 To mRCount * 2)
        ReDim Preserve mRPseudo(0 To mRCount * 2): ReDim Preserve mRSimple(0 To mRCount * 2)
        ReDim Preserve mRSel(0 To mRCount * 2)
    End If
End Sub

Private Sub GrowBuckets()
    If mBkCount > UBound(mBkFirst) Then ReDim Preserve mBkFirst(0 To mBkCount * 2): ReDim Preserve mBkLen(0 To mBkCount * 2)
End Sub

' ================================================================ interned values

' Returns the index of a value text, adding and classifying it on first use.
Public Function Intern(ByVal text As String) As Long
    Dim found As Variant
    found = mVMap(text)
    If Not IsEmpty(found) Then Intern = found: Exit Function
    If mVCount > UBound(mVText) Then
        ReDim Preserve mVText(0 To mVCount * 2): ReDim Preserve mVKind(0 To mVCount * 2)
        ReDim Preserve mVNum(0 To mVCount * 2): ReDim Preserve mVColState(0 To mVCount * 2)
        ReDim Preserve mVRgb(0 To mVCount * 2): ReDim Preserve mVAlpha(0 To mVCount * 2)
        ReDim Preserve mVKeyword(0 To mVCount * 2)
    End If
    mVText(mVCount) = text
    Select Case text
        Case "inherit": mVKeyword(mVCount) = 1
        Case "unset": mVKeyword(mVCount) = 2
        Case "initial", "revert": mVKeyword(mVCount) = 3
        Case Else: mVKeyword(mVCount) = 0
    End Select
    Classify mVCount
    mVMap(text) = mVCount
    Intern = mVCount
    mVCount = mVCount + 1
End Function

Private Sub Classify(ByVal v As Long)
    Dim t As String, n As Double, unit As String
    t = LCase$(Trim$(mVText(v)))
    mVKind(v) = VK_WORD
    If Len(t) = 0 Then Exit Sub
    If InStr(1, t, "(") > 0 Or InStr(1, t, " ") > 0 Then
        If Left$(t, 5) = "calc(" Or Left$(t, 4) = "min(" Or Left$(t, 4) = "max(" Or Left$(t, 6) = "clamp(" Or Left$(t, 4) = "var(" Then mVKind(v) = VK_EXPR
        Exit Sub
    End If
    If Not ParseDimension(t, n, unit) Then Exit Sub
    Select Case unit
        Case "": mVKind(v) = VK_NUM: mVNum(v) = n
        Case "px": mVKind(v) = VK_PX: mVNum(v) = n
        Case "em": mVKind(v) = VK_EM: mVNum(v) = n
        Case "rem": mVKind(v) = VK_PX: mVNum(v) = n * 16
        Case "pt": mVKind(v) = VK_PX: mVNum(v) = n * 4 / 3
        Case "%": mVKind(v) = VK_PCT: mVNum(v) = n
        Case "vw", "vh", "vmin", "vmax", "ch": mVKind(v) = VK_EXPR
    End Select
End Sub

' "12px" -> 12, "px"; ".5" -> 0.5, ""; False when the text is not a number with a unit.
Private Function ParseDimension(ByRef t As String, ByRef n As Double, ByRef unit As String) As Boolean
    Dim i As Long, c As Long, digits As Long, length As Long
    length = Len(t)
    i = 1
    If i <= length Then
        c = AscW(Mid$(t, i, 1))
        If c = 43 Or c = 45 Then i = i + 1
    End If
    Do While i <= length
        c = AscW(Mid$(t, i, 1))
        If c >= 48 And c <= 57 Then
            digits = digits + 1
        ElseIf c <> 46 Then
            Exit Do
        End If
        i = i + 1
    Loop
    If digits = 0 Then Exit Function
    n = Val(Left$(t, i - 1))
    unit = Mid$(t, i)
    Select Case unit
        Case "", "px", "em", "rem", "pt", "%", "vw", "vh", "vmin", "vmax", "ch", "fr", "deg"
            ParseDimension = True
    End Select
End Function

Public Function ValueText(ByVal v As Long) As String
    ValueText = mVText(v)
End Function

' ================================================================ environment

' Starts a layout pass for a viewport (CSS px). Invalidates everything when the viewport,
' the theme or the rule table changed.
Public Sub StyleBegin(ByVal viewWidth As Double, ByVal viewHeight As Double)
    Dim dark As Boolean, theme As Variant, c As Long
    If Not mLoaded Then Err.Raise vbObjectError + 804, "modStyle", "Style rules are not loaded"
    theme = modDom.DocumentElement.GetAttribute("data-theme")
    If Not IsNull(theme) Then dark = (CStr(theme) = "dark")
    If viewWidth <> mViewW Or viewHeight <> mViewH Or dark <> mDark Then
        mViewW = viewWidth
        mViewH = viewHeight
        mDark = dark
        mEnv = mEnv + 1
        For c = 1 To mCCount
            mCMediaOk(c) = MediaHolds(c)
        Next c
    End If
    mPass = mPass + 1
    mPassStart = mPass
End Sub

' Frees the slots of elements that were not styled during this pass.
Public Sub StyleEnd()
    Dim s As Long
    For s = 1 To mSUsed
        If mSInUse(s) Then
            If mSPseudo(s) = 0 And mSPass(s) < mPassStart Then ReleaseSlot s
        End If
    Next s
End Sub

' Forgets every cached style (e.g. after the rule table changed).
Public Sub StyleInvalidateAll()
    mEnv = mEnv + 1
    mEpoch = mEpoch + 1
End Sub

' Recomputes an element's style on its next use (state its node version does not show,
' such as :focus).
Public Sub StyleRefreshNode(ByVal node As clsNode)
    Dim uid As Long, s As Long
    If node Is Nothing Then Exit Sub
    uid = node.Uid
    If uid <= 0 Or uid > mUidCap Then Exit Sub
    s = mUidSlot(uid)
    If s > 0 Then mSEnv(s) = -1: mSPass(s) = -1
End Sub

' Changes whenever the rule table was (re)loaded or every style was invalidated.
Public Function StyleEpoch() As Long
    StyleEpoch = mEpoch
End Function

' A stamp that changes whenever a layout-relevant computed value of slot s changes
' (everything except colours of backgrounds and borders, opacity, visibility, z-index).
Public Function SLayoutStamp(ByVal s As Long) As Long
    If s > 0 Then SLayoutStamp = mSLay(s)
End Function

Public Function StyleIsDark() As Boolean
    StyleIsDark = mDark
End Function

Public Property Get ViewportWidth() As Double
    ViewportWidth = mViewW
End Property

Public Property Get ViewportHeight() As Double
    ViewportHeight = mViewH
End Property

Private Function MediaHolds(ByVal c As Long) As Boolean
    ' The media part of condition c and its outer conditions (container parts pass).
    Dim a As Long, ok As Boolean
    If c = 0 Then MediaHolds = True: Exit Function
    If mCIsContainer(c) Then
        ok = True
    Else
        For a = mCAltFirst(c) To mCAltFirst(c) + mCAltLen(c) - 1
            If AltHolds(a, mViewW) Then ok = True: Exit For
        Next a
    End If
    If ok Then ok = MediaHolds(mCParent(c))
    MediaHolds = ok
End Function

Private Function AltHolds(ByVal a As Long, ByVal width As Double) As Boolean
    If mAltNever(a) Then Exit Function
    If mAltMin(a) >= 0 Then If width < mAltMin(a) Then Exit Function
    If mAltMax(a) >= 0 Then If width > mAltMax(a) Then Exit Function
    AltHolds = True
End Function

' Bits of the container conditions that hold for a container of this content width.
Private Function ContainerBits(ByVal width As Double) As Long
    Dim c As Long, a As Long, ok As Boolean, bits As Long
    For c = 1 To mCCount
        If mCIsContainer(c) Then
            ok = False
            For a = mCAltFirst(c) To mCAltFirst(c) + mCAltLen(c) - 1
                If AltHolds(a, width) Then ok = True: Exit For
            Next a
            If ok Then bits = bits Or mCBit(c)
        End If
    Next c
    ContainerBits = bits
End Function

Private Function ConditionHolds(ByVal c As Long, ByVal bits As Long) As Boolean
    If Not mCMediaOk(c) Then Exit Function
    Do While c > 0
        If mCIsContainer(c) Then
            If (bits And mCBit(c)) = 0 Then Exit Function
        End If
        c = mCParent(c)
    Loop
    ConditionHolds = True
End Function

' Records the content width of a size container (called by layout). When the container
' queries that hold change, every style is revalidated (returns True).
Public Function SetContainerWidth(ByVal node As clsNode, ByVal width As Double) As Boolean
    Dim old As Double
    EnsureUid node.Uid
    old = mContW(node.Uid)
    mContW(node.Uid) = width
    If old < 0 Then old = mViewW
    If ContainerBits(old) <> ContainerBits(width) Then
        mPass = mPass + 1
        SetContainerWidth = True
    End If
End Function

Private Function ContainerWidthOf(ByVal uid As Long) As Double
    If uid <= 0 Or uid > mUidCap Then ContainerWidthOf = mViewW: Exit Function
    If mContW(uid) < 0 Then ContainerWidthOf = mViewW Else ContainerWidthOf = mContW(uid)
End Function

' ================================================================ slots

Private Sub ResetSlots()
    mSCap = 255
    mSUsed = 0
    mFreeCount = 0
    ReDim mFree(0 To 63)
    ReDim mSInUse(0 To mSCap): ReDim mSNode(0 To mSCap): ReDim mSUid(0 To mSCap)
    ReDim mSPseudo(0 To mSCap): ReDim mSNodeVer(0 To mSCap): ReDim mSParentVer(0 To mSCap)
    ReDim mSParent(0 To mSCap): ReDim mSParentSerial(0 To mSCap): ReDim mSSerial(0 To mSCap): ReDim mSLay(0 To mSCap)
    ReDim mSEnv(0 To mSCap): ReDim mSPass(0 To mSCap): ReDim mSBefore(0 To mSCap)
    ReDim mSAfter(0 To mSCap): ReDim mSContainer(0 To mSCap): ReDim mSCondBits(0 To mSCap)
    ReDim mSFont(0 To mSCap): ReDim mSWeight(0 To mSCap)
    ReDim mVal(0 To (mSCap + 1) * NP - 1)
    mUidCap = 1023
    ReDim mUidSlot(0 To mUidCap)
    ReDim mContW(0 To mUidCap)
    Dim i As Long
    For i = 0 To mUidCap
        mContW(i) = -1
    Next i
End Sub

Private Sub EnsureUid(ByVal uid As Long)
    Dim old As Long, i As Long
    If uid <= mUidCap Then Exit Sub
    old = mUidCap
    Do While mUidCap < uid
        mUidCap = mUidCap * 2 + 1
    Loop
    ReDim Preserve mUidSlot(0 To mUidCap)
    ReDim Preserve mContW(0 To mUidCap)
    For i = old + 1 To mUidCap
        mContW(i) = -1
    Next i
End Sub

Private Function AllocSlot() As Long
    Dim s As Long
    If mFreeCount > 0 Then
        mFreeCount = mFreeCount - 1
        s = mFree(mFreeCount)
    Else
        mSUsed = mSUsed + 1
        s = mSUsed
        If s > mSCap Then
            mSCap = mSCap * 2 + 1
            ReDim Preserve mSInUse(0 To mSCap): ReDim Preserve mSNode(0 To mSCap): ReDim Preserve mSUid(0 To mSCap)
            ReDim Preserve mSPseudo(0 To mSCap): ReDim Preserve mSNodeVer(0 To mSCap): ReDim Preserve mSParentVer(0 To mSCap)
            ReDim Preserve mSParent(0 To mSCap): ReDim Preserve mSParentSerial(0 To mSCap): ReDim Preserve mSSerial(0 To mSCap): ReDim Preserve mSLay(0 To mSCap)
            ReDim Preserve mSEnv(0 To mSCap): ReDim Preserve mSPass(0 To mSCap): ReDim Preserve mSBefore(0 To mSCap)
            ReDim Preserve mSAfter(0 To mSCap): ReDim Preserve mSContainer(0 To mSCap): ReDim Preserve mSCondBits(0 To mSCap)
            ReDim Preserve mSFont(0 To mSCap): ReDim Preserve mSWeight(0 To mSCap)
            ReDim Preserve mVal(0 To (mSCap + 1) * NP - 1)
        End If
    End If
    mSInUse(s) = True
    mSBefore(s) = -1
    mSAfter(s) = -1
    mSSerial(s) = 0
    mSLay(s) = 0
    AllocSlot = s
End Function

Private Sub ReleaseSlot(ByVal s As Long)
    If Not mSInUse(s) Then Exit Sub
    If mSPseudo(s) = 0 Then
        If mSUid(s) > 0 And mSUid(s) <= mUidCap Then
            If mUidSlot(mSUid(s)) = s Then mUidSlot(mSUid(s)) = 0
        End If
        If mSBefore(s) > 0 Then ReleaseSlot mSBefore(s)
        If mSAfter(s) > 0 Then ReleaseSlot mSAfter(s)
    End If
    mSInUse(s) = False
    Set mSNode(s) = Nothing
    If mFreeCount > UBound(mFree) Then ReDim Preserve mFree(0 To mFreeCount * 2)
    mFree(mFreeCount) = s
    mFreeCount = mFreeCount + 1
End Sub

' The computed style slot of an element (validated or recomputed for this pass).
Public Function StyleOf(ByVal node As clsNode) As Long
    Dim uid As Long, s As Long, parent As clsNode, ps As Long, valid As Boolean
    uid = node.Uid
    EnsureUid uid
    s = mUidSlot(uid)
    If s > 0 Then
        If mSPass(s) = mPass Then StyleOf = s: Exit Function
    End If
    Set parent = node.ParentNode
    If Not parent Is Nothing Then
        If parent.NodeType = 1 Then ps = StyleOf(parent) Else Set parent = Nothing
    End If
    If s > 0 Then
        valid = (mSNodeVer(s) = node.Version And mSEnv(s) = mEnv And mSParent(s) = ps)
        If valid And ps > 0 Then valid = (mSParentSerial(s) = mSSerial(ps) And mSParentVer(s) = parent.Version)
        If valid And mSCondBits(s) >= 0 Then valid = (mSCondBits(s) = ContainerBits(ContainerWidthOf(mSContainer(s))))
        If valid Then
            mSPass(s) = mPass
            StyleOf = s
            Exit Function
        End If
        If mSBefore(s) > 0 Then ReleaseSlot mSBefore(s)
        If mSAfter(s) > 0 Then ReleaseSlot mSAfter(s)
        mSBefore(s) = -1
        mSAfter(s) = -1
    Else
        s = AllocSlot()
        mUidSlot(uid) = s
        Set mSNode(s) = node
        mSUid(s) = uid
        mSPseudo(s) = 0
    End If
    ComputeSlot s, node, parent, ps, 0
    mSPass(s) = mPass
    StyleOf = s
End Function

' The slot of an element's ::before (1) or ::after (2) box, or 0 when it has no content.
Public Function PseudoOf(ByVal node As clsNode, ByVal which As Long) As Long
    Dim es As Long, ps As Long, cur As Long
    es = StyleOf(node)
    If which = 1 Then cur = mSBefore(es) Else cur = mSAfter(es)
    If cur >= 0 Then
        If cur > 0 Then mSPass(cur) = mPass
        PseudoOf = cur
        Exit Function
    End If
    ps = AllocSlot()
    Set mSNode(ps) = node
    mSUid(ps) = node.Uid
    mSPseudo(ps) = which
    If ComputeSlot(ps, node, node, es, which) Then
        Select Case mVText(mVal(ps * NP + P_CONTENT))
            Case "normal", "none", ""
                ReleaseSlot ps
                ps = 0
        End Select
    Else
        ReleaseSlot ps
        ps = 0
    End If
    If ps > 0 Then mSPass(ps) = mPass
    If which = 1 Then mSBefore(es) = ps Else mSAfter(es) = ps
    PseudoOf = ps
End Function

' Collects the candidate rules for an element (bucket order merged into rule order).
Private Sub CollectCandidates(ByVal node As clsNode)
    Dim cls As String, parts() As String, i As Long, idText As String, sig As String, found As Variant, k As Long, n As Long
    idText = node.Id
    cls = node.ClassName
    sig = node.LocalName & "#" & idText & "." & cls
    found = mSigMap(sig)
    If Not IsEmpty(found) Then
        k = found
        n = mSigLen(k)
        If n > UBound(mCand) + 1 Then ReDim mCand(0 To n * 2)
        For i = 0 To n - 1
            mCand(i) = mSigRules(mSigFirst(k) + i)
        Next i
        mCandCount = n
        Exit Sub
    End If
    CollectCandidatesNow node
    ' Remember the merged list for elements with the same tag, id and classes.
    If mSigCount > UBound(mSigFirst) Then ReDim Preserve mSigFirst(0 To mSigCount * 2): ReDim Preserve mSigLen(0 To mSigCount * 2)
    If mSigRuleCount + mCandCount > UBound(mSigRules) Then ReDim Preserve mSigRules(0 To (mSigRuleCount + mCandCount) * 2)
    mSigFirst(mSigCount) = mSigRuleCount
    mSigLen(mSigCount) = mCandCount
    For i = 0 To mCandCount - 1
        mSigRules(mSigRuleCount + i) = mCand(i)
    Next i
    mSigRuleCount = mSigRuleCount + mCandCount
    mSigMap(sig) = mSigCount
    mSigCount = mSigCount + 1
End Sub

Private Sub CollectCandidatesNow(ByVal node As clsNode)
    Dim cls As String, parts() As String, i As Long, idText As String
    mCandCount = 0
    AddBucket "*"
    AddBucket node.LocalName
    idText = node.Id
    If Len(idText) > 0 Then AddBucket "#" & idText
    cls = node.ClassName
    If Len(cls) > 0 Then
        parts = Split(cls, " ")
        For i = 0 To UBound(parts)
            If Len(parts(i)) > 0 Then AddBucket "." & parts(i)
        Next i
    End If
    SortCandidates
End Sub

Private Sub AddBucket(ByVal key As String)
    Dim b As Variant, i As Long, n As Long
    b = mBucketMap(key)
    If IsEmpty(b) Then Exit Sub
    n = mBkLen(b)
    If mCandCount + n > UBound(mCand) Then ReDim Preserve mCand(0 To (mCandCount + n) * 2)
    For i = mBkFirst(b) To mBkFirst(b) + n - 1
        mCand(mCandCount) = mBkRules(i)
        mCandCount = mCandCount + 1
    Next i
End Sub

Private Sub SortCandidates()
    Dim i As Long, j As Long, v As Long, w As Long, k As Long
    For i = 1 To mCandCount - 1
        v = mCand(i)
        j = i - 1
        Do While j >= 0
            If mCand(j) <= v Then Exit Do
            mCand(j + 1) = mCand(j)
            j = j - 1
        Loop
        mCand(j + 1) = v
    Next i
    ' A class listed twice adds its bucket twice: drop duplicates.
    If mCandCount > 1 Then
        k = 1
        For i = 1 To mCandCount - 1
            w = mCand(i)
            If w <> mCand(k - 1) Then mCand(k) = w: k = k + 1
        Next i
        mCandCount = k
    End If
End Sub

' Cascades the matching rules (and inline styles) into slot s. Returns False for a pseudo
' element that no rule targets.
Private Function ComputeSlot(ByVal s As Long, ByVal node As clsNode, ByVal parent As clsNode, ByVal ps As Long, ByVal pseudo As Long) As Boolean
    Dim i As Long, r As Long, b As Long, d As Long, p As Long, v As Long, base As Long, pbase As Long
    Dim bits As Long, usesContainer As Boolean, container As Long, matched As Boolean, anyFound As Boolean
    Dim changed As Boolean, fontSet As Boolean, txt As String, layChanged As Boolean, kw As Long, oldFont As Double, oldWeight As Long
    ' The size container whose width container queries test.
    If ps > 0 Then
        If pseudo = 0 And mVText(mVal(ps * NP + P_CONTAINER)) <> "normal" Then container = mSUid(ps) Else container = mSContainer(ps)
        If pseudo <> 0 Then container = mSContainer(ps)
    End If
    bits = ContainerBits(ContainerWidthOf(container))
    base = s * NP
    For p = 0 To NP - 1
        mOld(p) = mVal(base + p)
        mNorm(p) = -1
        mImp(p) = -1
    Next p
    CollectCandidates node
    For i = 0 To mCandCount - 1
        r = mCand(i)
        If mRPseudo(r) = pseudo And Not mRDead(r) Then
            matched = True
            If mRCond(r) > 0 Then
                If mCAnyContainer(mRCond(r)) Then usesContainer = True
                matched = ConditionHolds(mRCond(r), bits)
            End If
            If matched And Not mRSimple(r) Then matched = SafeMatches(node, r)
            If matched Then
                anyFound = True
                b = mRBlock(r)
                For d = mBFirst(b) To mBFirst(b) + mBLen(b) - 1
                    If mDark Then v = mDDark(d) Else v = mDLight(d)
                    If mDImp(d) Then mImp(mDProp(d)) = v Else mNorm(mDProp(d)) = v
                Next d
            End If
        End If
    Next i
    If pseudo <> 0 And Not anyFound Then Exit Function
    If pseudo = 0 Then ApplyInline node
    ' Resolve inheritance and keywords.
    pbase = ps * NP
    For p = 0 To NP - 1
        v = mImp(p)
        If v < 0 Then v = mNorm(p)
        If v >= 0 Then
            kw = mVKeyword(v)
            If kw = 1 Or (kw = 2 And mInherited(p)) Then
                If ps > 0 Then v = mVal(pbase + p) Else v = mInitial(p)
            ElseIf kw >= 2 Then
                v = mInitial(p)
            End If
            If p = P_FONTSIZE And kw <> 1 Then fontSet = True
        ElseIf mInherited(p) And ps > 0 Then
            v = mVal(pbase + p)
        Else
            v = mInitial(p)
        End If
        mVal(base + p) = v
    Next p
    oldFont = mSFont(s)
    oldWeight = mSWeight(s)
    If ps > 0 And Not fontSet Then
        mSFont(s) = mSFont(ps)
    Else
        mSFont(s) = ComputeFontSize(s, ps)
    End If
    mSWeight(s) = WeightOf(mVText(mVal(base + P_FONTWEIGHT)))
    ' Bookkeeping for validation.
    mSNodeVer(s) = node.Version
    mSEnv(s) = mEnv
    mSParent(s) = ps
    If ps > 0 Then
        mSParentSerial(s) = mSSerial(ps)
        If pseudo = 0 Then mSParentVer(s) = parent.Version
    End If
    mSContainer(s) = container
    If usesContainer Then mSCondBits(s) = bits Else mSCondBits(s) = -1
    ' A recomputation that changed nothing keeps the serial, so children stay valid.
    ' (The resolved font size and weight count too: children inherit them.)
    If mSSerial(s) = 0 Or mSLay(s) = 0 Or oldFont <> mSFont(s) Or oldWeight <> mSWeight(s) Then
        changed = True
        layChanged = True
    Else
        For p = 0 To NP - 1
            If mOld(p) <> mVal(base + p) Then
                changed = True
                If Not mPaintOnly(p) Then layChanged = True: Exit For
            End If
        Next p
    End If
    If changed Then
        mSerialSeq = mSerialSeq + 1
        mSSerial(s) = mSerialSeq
    End If
    If layChanged Then
        mSerialSeq = mSerialSeq + 1
        mSLay(s) = mSerialSeq
    End If
    ComputeSlot = True
End Function

Private Function SafeMatches(ByVal node As clsNode, ByVal r As Long) As Boolean
    On Error GoTo Failed
    SafeMatches = node.Matches(mRSel(r))
    Exit Function
Failed:
    mRDead(r) = True
    SafeMatches = False
End Function

' Inline style declarations (node.SetStyle); custom properties stay on the node for var().
Private Sub ApplyInline(ByVal node As clsNode)
    Dim names As clsList, i As Long, name As String, value As String
    Set names = node.StyleNames
    If names.Count = 0 Then Exit Sub
    For i = 0 To names.Count - 1
        name = LCase$(CStr(names(i)))
        If Left$(name, 2) <> "--" Then
            value = Trim$(node.GetStyle(name))
            If Len(value) > 0 Then ExpandInline name, value
        End If
    Next i
End Sub

Private Sub ExpandInline(ByVal name As String, ByVal value As String)
    Dim t() As String, n As Long, important As Boolean
    If LCase$(Right$(value, 10)) = "!important" Then
        important = True
        value = Trim$(Left$(value, Len(value) - 10))
    End If
    t = Split(Application.WorksheetFunction.Trim(value), " ")
    n = UBound(t) + 1
    Select Case name
        Case "margin": SetFour P_MT, t, n, important
        Case "padding": SetFour P_PT, t, n, important
        Case "inset": SetFour P_TOP, t, n, important
        Case "gap"
            SetInline P_ROWGAP, t(0), important
            If n > 1 Then SetInline P_COLGAP, t(1), important Else SetInline P_COLGAP, t(0), important
        Case "overflow"
            SetInline P_OVERFLOWX, t(0), important
            If n > 1 Then SetInline P_OVERFLOWY, t(1), important Else SetInline P_OVERFLOWY, t(0), important
        Case "flex"
            If n = 1 And t(0) = "none" Then
                SetInline P_FLEXGROW, "0", important: SetInline P_FLEXSHRINK, "0", important: SetInline P_FLEXBASIS, "auto", important
            ElseIf n = 1 And IsNumeric(t(0)) Then
                SetInline P_FLEXGROW, t(0), important: SetInline P_FLEXSHRINK, "1", important: SetInline P_FLEXBASIS, "0px", important
            ElseIf n >= 2 Then
                SetInline P_FLEXGROW, t(0), important: SetInline P_FLEXSHRINK, t(1), important
                If n >= 3 Then SetInline P_FLEXBASIS, t(2), important Else SetInline P_FLEXBASIS, "0px", important
            End If
        Case "background": SetInline P_BG, value, important
        Case Else
            Dim p As Long
            p = PropertyIndex(name)
            If p >= 0 Then SetInline p, value, important
    End Select
End Sub

Private Sub SetFour(ByVal first As Long, ByRef t() As String, ByVal n As Long, ByVal important As Boolean)
    Dim a As String, b As String, c As String, d As String
    a = t(0)
    If n > 1 Then b = t(1) Else b = a
    If n > 2 Then c = t(2) Else c = a
    If n > 3 Then d = t(3) Else d = b
    SetInline first, a, important
    SetInline first + 1, b, important
    SetInline first + 2, c, important
    SetInline first + 3, d, important
End Sub

Private Sub SetInline(ByVal p As Long, ByVal value As String, ByVal important As Boolean)
    If important Then mImp(p) = Intern(value) Else mNorm(p) = Intern(value)
End Sub

Public Function PropertyIndex(ByVal name As String) As Long
    Dim names() As String, i As Long
    names = Split(PROPERTY_NAMES, " ")
    For i = 0 To UBound(names)
        If names(i) = name Then PropertyIndex = i: Exit Function
    Next i
    PropertyIndex = -1
End Function

Private Function ComputeFontSize(ByVal s As Long, ByVal ps As Long) As Double
    Dim parentSize As Double, v As Long, px As Double, t As String
    If ps > 0 Then parentSize = mSFont(ps) Else parentSize = 16
    v = mVal(s * NP + P_FONTSIZE)
    t = mVText(v)
    Select Case mVKind(v)
        Case VK_PX: ComputeFontSize = mVNum(v)
        Case VK_EM: ComputeFontSize = mVNum(v) * parentSize
        Case VK_PCT: ComputeFontSize = mVNum(v) * parentSize / 100
        Case VK_NUM: ComputeFontSize = mVNum(v)
        Case VK_EXPR
            If EvalText(t, s, parentSize, True, px, parentSize) Then ComputeFontSize = px Else ComputeFontSize = parentSize
        Case Else
            Select Case t
                Case "smaller": ComputeFontSize = parentSize / 1.2
                Case "larger": ComputeFontSize = parentSize * 1.2
                Case "xx-small": ComputeFontSize = 9
                Case "x-small": ComputeFontSize = 10
                Case "small": ComputeFontSize = 13
                Case "medium": ComputeFontSize = 16
                Case "large": ComputeFontSize = 18
                Case "x-large": ComputeFontSize = 24
                Case "xx-large": ComputeFontSize = 32
                Case Else: ComputeFontSize = parentSize
            End Select
    End Select
    If ComputeFontSize <= 0 Then ComputeFontSize = parentSize
End Function

Private Function WeightOf(ByVal t As String) As Long
    Dim n As Double
    Select Case t
        Case "bold", "bolder": WeightOf = 700
        Case "normal", "lighter", "": WeightOf = 400
        Case Else
            n = Val(t)
            If n >= 650 Then
                WeightOf = 700
            ElseIf n >= 550 Then
                WeightOf = 600
            Else
                WeightOf = 400
            End If
    End Select
End Function

' ================================================================ accessors

' Value text of a property.
Public Function SV(ByVal s As Long, ByVal p As Long) As String
    SV = mVText(mVal(s * NP + p))
End Function

' Value index of a property (equal indexes mean equal values).
Public Function SVI(ByVal s As Long, ByVal p As Long) As Long
    SVI = mVal(s * NP + p)
End Function

' Value kind (VK_*) of a property.
Public Function SKind(ByVal s As Long, ByVal p As Long) As Long
    SKind = mVKind(mVal(s * NP + p))
End Function

' Numeric value (flex-grow, opacity...): Val of the text.
Public Function SNum(ByVal s As Long, ByVal p As Long) As Double
    Dim v As Long
    v = mVal(s * NP + p)
    If mVKind(v) = VK_NUM Or mVKind(v) = VK_PX Then SNum = mVNum(v) Else SNum = Val(mVText(v))
End Function

Public Function SFont(ByVal s As Long) As Double
    SFont = mSFont(s)
End Function

Public Function SWeight(ByVal s As Long) As Long
    SWeight = mSWeight(s)
End Function

Public Function SItalic(ByVal s As Long) As Boolean
    Dim t As String
    t = mVText(mVal(s * NP + P_FONTSTYLE))
    SItalic = (t = "italic" Or t = "oblique")
End Function

Public Function SNode(ByVal s As Long) As clsNode
    Set SNode = mSNode(s)
End Function

Public Function SPseudo(ByVal s As Long) As Long
    SPseudo = mSPseudo(s)
End Function

Public Function SParent(ByVal s As Long) As Long
    SParent = mSParent(s)
End Function

' Evaluates a length property in px. base: the percentage basis (hasBase False when the
' basis is indefinite). Returns False for auto/none/normal or an unresolvable value.
Public Function SLen(ByVal s As Long, ByVal p As Long, ByVal base As Double, ByVal hasBase As Boolean, ByRef px As Double) As Boolean
    Dim v As Long
    v = mVal(s * NP + p)
    Select Case mVKind(v)
        Case VK_PX, VK_NUM: px = mVNum(v): SLen = True
        Case VK_EM: px = mVNum(v) * mSFont(s): SLen = True
        Case VK_PCT
            If hasBase Then px = base * mVNum(v) / 100: SLen = True
        Case VK_EXPR
            SLen = EvalText(mVText(v), s, base, hasBase, px, mSFont(s))
        Case Else
            Select Case mVText(v)
                Case "thin": px = 1: SLen = True
                Case "medium": px = 3: SLen = True
                Case "thick": px = 5: SLen = True
            End Select
    End Select
End Function

' True when a length property is a percentage (or an expression that needs a basis).
Public Function SIsRelative(ByVal s As Long, ByVal p As Long) As Boolean
    Dim v As Long
    v = mVal(s * NP + p)
    If mVKind(v) = VK_PCT Then
        SIsRelative = True
    ElseIf mVKind(v) = VK_EXPR Then
        SIsRelative = (InStr(1, mVText(v), "%") > 0)
    End If
End Function

' Evaluates arbitrary length text in the context of slot s (grid tracks, flex bases).
Public Function SLenText(ByVal text As String, ByVal s As Long, ByVal base As Double, ByVal hasBase As Boolean, ByRef px As Double) As Boolean
    SLenText = EvalText(text, s, base, hasBase, px, mSFont(s))
End Function

' Colour of a property as RGB (VBA Long) and alpha. False when transparent.
Public Function SColour(ByVal s As Long, ByVal p As Long, ByRef rgbValue As Long, ByRef alpha As Double) As Boolean
    Dim v As Long
    v = mVal(s * NP + p)
    If mVColState(v) = 0 Then ParseColourValue v
    Select Case mVColState(v)
        Case 1
            rgbValue = mVRgb(v)
            alpha = mVAlpha(v)
            SColour = (alpha > 0)
        Case 3
            If p <> P_COLOR Then SColour = SColour(s, P_COLOR, rgbValue, alpha)
    End Select
End Function

' Colour text such as "#0b57d033" (a box-shadow layer's colour) as RGB and alpha.
Public Function ColourText(ByVal text As String, ByRef rgbValue As Long, ByRef alpha As Double) As Boolean
    Dim v As Long
    v = Intern(text)
    If mVColState(v) = 0 Then ParseColourValue v
    If mVColState(v) = 1 Then
        rgbValue = mVRgb(v)
        alpha = mVAlpha(v)
        ColourText = (alpha > 0)
    End If
End Function

Private Sub ParseColourValue(ByVal v As Long)
    Dim t As String, r As Long, g As Long, b As Long, a As Double
    t = LCase$(Trim$(mVText(v)))
    mVColState(v) = 2
    If t = "currentcolor" Then mVColState(v) = 3: Exit Sub
    If Left$(t, 1) = "#" Then
        If Len(t) = 4 Then t = "#" & String$(2, Mid$(t, 2, 1)) & String$(2, Mid$(t, 3, 1)) & String$(2, Mid$(t, 4, 1))
        If Len(t) = 7 Or Len(t) = 9 Then
            r = CLng("&H" & Mid$(t, 2, 2)): g = CLng("&H" & Mid$(t, 4, 2)): b = CLng("&H" & Mid$(t, 6, 2))
            a = 1
            If Len(t) = 9 Then a = CLng("&H" & Mid$(t, 8, 2)) / 255
            mVRgb(v) = RGB(r, g, b)
            mVAlpha(v) = a
            mVColState(v) = 1
        End If
    ElseIf t = "white" Then
        mVRgb(v) = RGB(255, 255, 255): mVAlpha(v) = 1: mVColState(v) = 1
    ElseIf t = "black" Then
        mVRgb(v) = 0: mVAlpha(v) = 1: mVColState(v) = 1
    End If
End Sub

' ================================================================ expressions

' Evaluates a CSS length expression: numbers with px/em/rem/%/vw/vh/vmin/vmax/pt/ch,
' + - * /, calc() min() max() clamp() and var(--name, fallback) (custom properties set
' inline on the element or an ancestor).
Private Function EvalText(ByVal text As String, ByVal s As Long, ByVal base As Double, ByVal hasBase As Boolean, ByRef px As Double, ByVal em As Double) As Boolean
    Dim i As Long, ok As Boolean, t As String
    t = LCase$(text)
    i = 1
    ok = True
    px = ExSum(t, i, s, base, hasBase, em, ok)
    If ok Then
        ExSkip t, i
        If i <= Len(t) Then ok = False
    End If
    EvalText = ok
End Function

Private Sub ExSkip(ByRef t As String, ByRef i As Long)
    Do While i <= Len(t)
        If Mid$(t, i, 1) <> " " And Mid$(t, i, 1) <> vbTab And Mid$(t, i, 1) <> vbLf Then Exit Do
        i = i + 1
    Loop
End Sub

Private Function ExSum(ByRef t As String, ByRef i As Long, ByVal s As Long, ByVal base As Double, ByVal hasBase As Boolean, ByVal em As Double, ByRef ok As Boolean) As Double
    Dim value As Double, right As Double, op As String
    value = ExProduct(t, i, s, base, hasBase, em, ok)
    Do While ok
        ExSkip t, i
        op = Mid$(t, i, 1)
        If op <> "+" And op <> "-" Then Exit Do
        i = i + 1
        right = ExProduct(t, i, s, base, hasBase, em, ok)
        If op = "+" Then value = value + right Else value = value - right
    Loop
    ExSum = value
End Function

Private Function ExProduct(ByRef t As String, ByRef i As Long, ByVal s As Long, ByVal base As Double, ByVal hasBase As Boolean, ByVal em As Double, ByRef ok As Boolean) As Double
    Dim value As Double, right As Double, op As String
    value = ExAtom(t, i, s, base, hasBase, em, ok)
    Do While ok
        ExSkip t, i
        op = Mid$(t, i, 1)
        If op <> "*" And op <> "/" Then Exit Do
        i = i + 1
        right = ExAtom(t, i, s, base, hasBase, em, ok)
        If op = "*" Then
            value = value * right
        ElseIf right = 0 Then
            ok = False
        Else
            value = value / right
        End If
    Loop
    ExProduct = value
End Function

Private Function ExAtom(ByRef t As String, ByRef i As Long, ByVal s As Long, ByVal base As Double, ByVal hasBase As Boolean, ByVal em As Double, ByRef ok As Boolean) As Double
    Dim fname As String, args(0 To 7) As Double, n As Long, j As Long, c As String, start As Long
    Dim number As Double, unit As String, value As Double
    ExSkip t, i
    If i > Len(t) Then ok = False: Exit Function
    c = Mid$(t, i, 1)
    If c = "(" Then
        fname = "calc"
        i = i + 1
    ElseIf Mid$(t, i, 4) = "var(" Then
        i = i + 4
        ExAtom = ExVar(t, i, s, base, hasBase, em, ok)
        Exit Function
    ElseIf Mid$(t, i, 5) = "calc(" Then
        fname = "calc": i = i + 5
    ElseIf Mid$(t, i, 4) = "min(" Then
        fname = "min": i = i + 4
    ElseIf Mid$(t, i, 4) = "max(" Then
        fname = "max": i = i + 4
    ElseIf Mid$(t, i, 6) = "clamp(" Then
        fname = "clamp": i = i + 6
    End If
    If Len(fname) > 0 Then
        Do
            args(n) = ExSum(t, i, s, base, hasBase, em, ok)
            If Not ok Then Exit Function
            n = n + 1
            ExSkip t, i
            c = Mid$(t, i, 1)
            If c = "," And n < 8 Then
                i = i + 1
            ElseIf c = ")" Then
                i = i + 1
                Exit Do
            Else
                ok = False
                Exit Function
            End If
        Loop
        Select Case fname
            Case "min"
                value = args(0)
                For j = 1 To n - 1
                    If args(j) < value Then value = args(j)
                Next j
            Case "max"
                value = args(0)
                For j = 1 To n - 1
                    If args(j) > value Then value = args(j)
                Next j
            Case "clamp"
                If n < 3 Then ok = False: Exit Function
                value = args(1)
                If args(2) < value Then value = args(2)
                If args(0) > value Then value = args(0)
            Case Else
                value = args(0)
        End Select
        ExAtom = value
        Exit Function
    End If
    ' A number with an optional unit.
    start = i
    If c = "+" Or c = "-" Then i = i + 1
    Do While i <= Len(t)
        c = Mid$(t, i, 1)
        If (c < "0" Or c > "9") And c <> "." Then Exit Do
        i = i + 1
    Loop
    If i = start Then ok = False: Exit Function
    number = Val(Mid$(t, start, i - start))
    start = i
    Do While i <= Len(t)
        c = Mid$(t, i, 1)
        If Not ((c >= "a" And c <= "z") Or c = "%") Then Exit Do
        i = i + 1
    Loop
    unit = Mid$(t, start, i - start)
    Select Case unit
        Case "", "px": ExAtom = number
        Case "em": ExAtom = number * em
        Case "rem": ExAtom = number * 16
        Case "ch": ExAtom = number * em * 0.5
        Case "pt": ExAtom = number * 4 / 3
        Case "%"
            If hasBase Then ExAtom = number * base / 100 Else ok = False
        Case "vw": ExAtom = number * mViewW / 100
        Case "vh": ExAtom = number * mViewH / 100
        Case "vmin"
            If mViewW < mViewH Then ExAtom = number * mViewW / 100 Else ExAtom = number * mViewH / 100
        Case "vmax"
            If mViewW > mViewH Then ExAtom = number * mViewW / 100 Else ExAtom = number * mViewH / 100
        Case Else: ok = False
    End Select
End Function

' var(--name[, fallback]) with i just after "var(".
Private Function ExVar(ByRef t As String, ByRef i As Long, ByVal s As Long, ByVal base As Double, ByVal hasBase As Boolean, ByVal em As Double, ByRef ok As Boolean) As Double
    Dim depth As Long, j As Long, c As String, inner As String, comma As Long, name As String
    Dim fallback As String, found As String, px As Double
    depth = 1
    j = i
    Do While j <= Len(t)
        c = Mid$(t, j, 1)
        If c = "(" Then
            depth = depth + 1
        ElseIf c = ")" Then
            depth = depth - 1
            If depth = 0 Then Exit Do
        ElseIf c = "," And depth = 1 And comma = 0 Then
            comma = j
        End If
        j = j + 1
    Loop
    If depth <> 0 Then ok = False: Exit Function
    If comma > 0 Then
        name = Trim$(Mid$(t, i, comma - i))
        fallback = Trim$(Mid$(t, comma + 1, j - comma - 1))
    Else
        name = Trim$(Mid$(t, i, j - i))
    End If
    i = j + 1
    found = CustomProperty(s, name)
    If Len(found) = 0 Then found = fallback
    If Len(found) = 0 Then ok = False: Exit Function
    If EvalText(found, s, base, hasBase, px, em) Then ExVar = px Else ok = False
End Function

' A custom property set inline (style.setProperty) on the element or an ancestor.
Private Function CustomProperty(ByVal s As Long, ByVal name As String) As String
    Dim node As clsNode, value As String
    If s <= 0 Then Exit Function
    Set node = mSNode(s)
    Do While Not node Is Nothing
        If node.NodeType <> 1 Then Exit Do
        value = node.GetStyle(name)
        If Len(value) > 0 Then CustomProperty = value: Exit Function
        Set node = node.ParentNode
    Loop
End Function

' ================================================================ diagnostics

Public Function StyleStats() As String
    StyleStats = "rules " & mRCount & ", blocks " & mBCount & ", decls " & mDCount & ", values " & mVCount & _
        ", buckets " & mBkCount & ", slots " & mSUsed & " (free " & mFreeCount & "), conditions " & mCCount
End Function

' Every property of an element's computed style as "name: value" lines (tests).
Public Function StyleDump(ByVal node As clsNode) As String
    Dim names() As String, s As Long, p As Long, out As String
    names = Split(PROPERTY_NAMES, " ")
    s = StyleOf(node)
    For p = 0 To NP - 1
        out = out & names(p) & ": " & SV(s, p) & vbLf
    Next p
    StyleDump = out & "font: " & mSFont(s) & "px " & mSWeight(s)
End Function
