Attribute VB_Name = "modRender"
' The app window. Lays out the visible DOM with the website's stylesheet (modStyle,
' clsRenderLayout/clsRenderGrid), paints it into a display list (clsRenderPaint) and
' reconciles that list with MSForms controls in frmApp: one Frame per region (title bar,
' library, page, nested scrollers, overlays) and pooled controls per item (clsRenderItem).
' User actions on the controls become DOM events; after each action the DOM settles
' (modDom.Settle 700) and the window re-renders when modDom.Version changed. Longer DOM
' timers run from Application.OnTime.
'
' Smooth updates: layout is incremental (clsRenderLayout reuses the fragments of unchanged
' subtrees); controls are written only when their look, geometry or value changed; items
' that disappear are hidden and kept per frame for reuse instead of being removed (no
' control creation when switching operations once the pools are warm); only controls
' painted after a new or reused one are re-stacked. On Windows the first control change
' of an update suspends painting of the form's client window (WM_SETREDRAW) and the update
' ends with a single RedrawWindow of the client and all its frames (WillWrite/ResumePaint).
' A Mac has no such API: there the same minimal set of property writes is applied directly.
'
' Drawn controls: buttons and switches (appearance:none checkboxes) are painted like any
' box (face rect, text, track and knob labels) so they look like the website on both
' platforms; labels take their clicks. Keyboard access comes from a focus proxy per element:
' a zero-size CommandButton (buttons) or CheckBox (switches) at the element's corner, in
' paint (= tab) order. Tab reaches it, Enter/Space press it (Space natively: Btn_Click /
' Chk_Click), a click on the drawn labels focuses it; :focus-visible outlines are painted
' when focus came from the keyboard. Zero-size controls take focus on Windows (tested with
' posted Tab/Space keys); on a Mac this is the one assumption to check first.
' Sticky elements (position: sticky in the page scroller, e.g. the narrow-window library
' button) get a pinned copy in an overlay frame shown while the page is scrolled past them
' (frame Scroll event, wheel, keys). Cosmetic MSForms properties go through TrySet.
'
' Entry points: OpenAppWindow, RenderNow, CloseAppWindow, LayoutSnapshot (headless, tests).
' Windows-only conveniences (private font registration, resizable frame, mouse wheel)
' sit in #If Mac Then #Else blocks; on a Mac the window uses Arial and native behaviour.
Option Explicit

#If Mac Then
#Else
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowW" (ByVal className As LongPtr, ByVal windowName As LongPtr) As LongPtr
#If Win64 Then
Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrW" (ByVal hwnd As LongPtr, ByVal index As Long) As LongPtr
Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrW" (ByVal hwnd As LongPtr, ByVal index As Long, ByVal value As LongPtr) As LongPtr
#Else
Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongW" (ByVal hwnd As LongPtr, ByVal index As Long) As LongPtr
Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongW" (ByVal hwnd As LongPtr, ByVal index As Long, ByVal value As LongPtr) As LongPtr
#End If
Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function ScreenToClient Lib "user32" (ByVal hwnd As LongPtr, ByRef point As POINTAPI) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal dc As LongPtr) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal dc As LongPtr, ByVal index As Long) As Long
Private Declare PtrSafe Function AddFontResourceEx Lib "gdi32" Alias "AddFontResourceExW" (ByVal fileName As LongPtr, ByVal flags As Long, ByVal reserved As LongPtr) As Long
Private Declare PtrSafe Function RemoveFontResourceEx Lib "gdi32" Alias "RemoveFontResourceExW" (ByVal fileName As LongPtr, ByVal flags As Long, ByVal reserved As LongPtr) As Long
Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExW" (ByVal hookId As Long, ByVal procedure As LongPtr, ByVal module As LongPtr, ByVal threadId As Long) As LongPtr
Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hook As LongPtr) As Long
Private Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hook As LongPtr, ByVal code As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare PtrSafe Function GetAncestor Lib "user32" (ByVal hwnd As LongPtr, ByVal flags As Long) As LongPtr
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal destination As LongPtr, ByVal source As LongPtr, ByVal bytes As LongPtr)
Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal cmd As Long) As LongPtr
Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameW" (ByVal hwnd As LongPtr, ByVal buffer As LongPtr, ByVal size As Long) As Long
Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hwnd As LongPtr, ByVal msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function RedrawWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal rect As LongPtr, ByVal region As LongPtr, ByVal flags As Long) As Long
Private Const GW_CHILD As Long = 5
Private Const GW_HWNDNEXT As Long = 2
Private Const WM_SETREDRAW As Long = &HB
Private Const RDW_INVALIDATE As Long = &H1
Private Const RDW_ERASE As Long = &H4
Private Const RDW_ALLCHILDREN As Long = &H80
Private Const RDW_UPDATENOW As Long = &H100
Private Const RDW_FRAME As Long = &H400
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef counter As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef frequency As Currency) As Long
Private Const GWL_STYLE As Long = -16
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const FR_PRIVATE As Long = &H10
Private Const WH_MOUSE As Long = 7
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const GA_ROOT As Long = 2
#If Win64 Then
Private Const HOOK_OFFSET_HWND As Long = 8
Private Const HOOK_OFFSET_DATA As Long = 32
#Else
Private Const HOOK_OFFSET_HWND As Long = 8
Private Const HOOK_OFFSET_DATA As Long = 20
#End If
Private mHwnd As LongPtr
Private mClientHwnd As LongPtr
Private mHook As LongPtr
Private mDpi As Long
#End If

Private Const APP_TITLE As String = "BT1101 Prose Lab"
Private Const DEFAULT_WIDTH As Double = 1280
Private Const DEFAULT_HEIGHT As Double = 860
Private Const NESTED_FRAMES As Long = 10

' Timings of the last render (ms) and how many controls it created.
Public LastRenderMs As Double, LastLayoutMs As Double, LastPaintMs As Double, LastApplyMs As Double
Public LastControlsCreated As Long, LastControlsUpdated As Long, LastItemCount As Long, LastControlsReused As Long
' Timings of the last user action (ms): the DOM event handler, Settle, the whole action.
Public LastHandlerMs As Double, LastSettleMs As Double, LastActionMs As Double
Public LastMeasureCount As Long, LastMeasureMs As Double

Private mForm As Object                 ' frmApp instance (also the measuring host)
Private mShown As Boolean
Private mLayout As clsRenderLayout
Private mPaint As clsRenderPaint
Private mPx As Double                   ' points per CSS px
Private mItems As clsMap                ' item id -> clsRenderItem
Private mCtlItems As clsMap             ' control name -> clsRenderItem
Private mNodeItems As clsMap            ' node uid -> clsRenderItem (controls)
Private mFrames As clsMap               ' region id -> MSForms Frame
Private mFrameParent As clsMap          ' region id -> parent region id
Private mFrameSpecs As clsMap           ' region id -> clsRenderSpec last applied
Private mFramePool As clsMap           ' container key -> Collection of hidden frames
Private mCounter As Long
Private mApplying As Boolean
Private mBusy As Boolean
Private mRendering As Boolean
Private mRenderedVersion As Long
Private mRenderedW As Double, mRenderedH As Double
Private mFocusNode As clsNode
Private mSinks As clsMap                ' container key -> invisible control that takes focus while controls go
Private mFocusValue As String
Private mTimerAt As Date
Private mTimerSet As Boolean
Private mResizePending As Boolean
Private mDragging As Boolean
Private mDragX As Double
Private mStyleText As String
Private mIdle As clsMap                 ' "frame|kind|flavour" -> Collection of hidden, reusable items
Private mPaintSuspended As Boolean
Private mSuspendEnabled As Boolean
Private mEngineInit As Boolean
Private Const IDLE_LIMIT As Long = 250
Private mActionStart As Double
Private mStyledFocus As clsNode         ' the element styled as :focus in the last render
Private mFocusVisible As Boolean        ' the last interaction was a key: draw :focus-visible outlines
Private mPaintedFocusVisible As Boolean
Private mScrollSink As clsRenderItem    ' Scroll events of the frame that hosts sticky copies
Private mStickyTickDue As Boolean
Public LastScrollEventTop As Double, LastScrollEventDy As Double   ' diagnostics (tests)

' fonts and measuring
Private mFontsReady As Boolean
Private mFontRegular As String, mFontSemibold As String, mFontSymbol As String
Private mSemiboldIsBold As Boolean
Private mSymbolFallback As Boolean
Private mFontFiles As Collection
Private mFontData As clsMap             ' tests: weight -> base64 TTF
Private mMeasurers As clsMap            ' font key -> Label
Private mFontObjects As clsMap          ' font key -> StdFont (Windows)
Private mOverhead As clsMap             ' font key -> width overhead (pt)
Private mLineHeights As clsMap          ' font key -> natural line height (px)
Private mNoMeasure As Boolean
Private mMeasureBroken As Boolean       ' labels did not measure (a Mac quirk): widths come from a table
' Arial / Helvetica advance widths (1/1000 em) of the characters 32..126.
Private Const ARIAL_WIDTHS As String = "278278355556556889667191333333389584278333278278556556556556556556556556556556278278584584584556" & _
    "015667667722722667611778722278500667556833722778667778722667611722667944667667611278278278469556" & _
    "333556556500556556278556556222222500222833556556556556333500278556500722500500500334260334584"

' ================================================================ entry points

' Opens (or shows again) the app window and renders the document.
Public Sub OpenAppWindow()
    On Error GoTo Failed
    If mShown And Not mForm Is Nothing Then
        mForm.Show vbModeless
        Exit Sub
    End If
    EnsureEngine
    EnsureStyles
    EnsureFonts
    EnsureHost
    SizeWindow
    mForm.Show vbModeless
    mShown = True
    DoEvents
    WindowCreated
    UpdateViewport True
    RenderNow
    Exit Sub
Failed:
    ReportError "modRender.OpenAppWindow"
    MsgBox APP_TITLE & " could not open its window." & vbCrLf & vbCrLf & Err.Description, vbCritical, APP_TITLE
End Sub

' Closes the app window (the document and its state stay loaded).
Public Sub CloseAppWindow()
    Dim f As Object
    On Error Resume Next
    CancelTimer
    WindowClosing
    If Not mForm Is Nothing Then
        Set f = mForm
        Set mForm = Nothing
        mShown = False
        Unload f
    End If
    ResetControls
    RemoveFonts
    mFontsReady = False
    Set mMeasurers = Nothing
End Sub

' Milliseconds from a high-resolution clock (timings; Timer on a Mac).
Public Function ClockMs() As Double
    #If Mac Then
        ClockMs = CDbl(Timer) * 1000
    #Else
        Static freq As Currency
        Dim c As Currency
        If freq = 0 Then QueryPerformanceFrequency freq
        QueryPerformanceCounter c
        If freq <> 0 Then ClockMs = CDbl(c) / CDbl(freq) * 1000 Else ClockMs = CDbl(Timer) * 1000
    #End If
End Function

' Lays out and redraws the window now.
Public Sub RenderNow()
    Dim t0 As Double, t1 As Double, t2 As Double, t3 As Double
    On Error GoTo Failed
    If mForm Is Nothing Or Not mShown Then Exit Sub
    If mRendering Then Exit Sub
    mRendering = True
    t0 = ClockMs()
    LastMeasureCount = 0
    LastMeasureMs = 0
    mRenderedW = mForm.InsideWidth / mPx
    mRenderedH = mForm.InsideHeight / mPx
    RefreshFocusStyle
    mLayout.Run mRenderedW, mRenderedH
    t1 = ClockMs()
    mPaint.PathIds = False
    Set mPaint.ScrollTarget = modApp.ScrollRequestTarget
    mPaint.FocusVisible = mFocusVisible
    mPaintedFocusVisible = mFocusVisible
    mPaint.Paint
    t2 = ClockMs()
    ApplyDisplayList mPaint.Regions(), mPaint.PageBg
    modStyle.StyleEnd
    mRenderedVersion = modDom.Version
    ApplyScrollRequest
    ReflectFocus
    UpdateSticky 0
    ResumePaint
    t3 = ClockMs()
    LastLayoutMs = t1 - t0
    LastPaintMs = t2 - t1
    LastApplyMs = t3 - t2
    LastRenderMs = ClockMs() - t0
    mRendering = False
    ScheduleTimer
    Exit Sub
Failed:
    mRendering = False
    mApplying = False
    ResumePaint
    ReportError "modRender.RenderNow"
End Sub

' :focus styles follow the DOM's focused element.
Private Sub RefreshFocusStyle()
    Dim act As clsNode
    Set act = modDom.ActiveElement()
    If act Is mStyledFocus Then Exit Sub
    modStyle.StyleRefreshNode mStyledFocus
    modStyle.StyleRefreshNode act
    Set mStyledFocus = act
End Sub

Private Function FocusRestyleDue() As Boolean
    FocusRestyleDue = Not (modDom.ActiveElement() Is mStyledFocus) Or (mFocusVisible <> mPaintedFocusVisible)
End Function

' For code that changed the DOM outside a window event (Application.OnTime handlers):
' settles timers' follow-up work, redraws when needed and re-arms the timer.
Public Sub RefreshAppWindow()
    On Error GoTo Failed
    If mForm Is Nothing Or Not mShown Then Exit Sub
    If mBusy Or mApplying Or mRendering Then Exit Sub
    mBusy = True
    modDom.Settle 700
    If modDom.Version <> mRenderedVersion Then RenderNow
    mBusy = False
    ScheduleTimer
    Exit Sub
Failed:
    mBusy = False
    ReportError "modRender.RefreshAppWindow"
End Sub

' Applies a window size change now (tests; the window itself does this from OnTime).
Public Sub ResizeNow()
    On Error GoTo Failed
    If mForm Is Nothing Or Not mShown Then Exit Sub
    UpdateViewport False
    RenderNow
    Exit Sub
Failed:
    ReportError "modRender.ResizeNow"
End Sub

' The control drawn for an element (Nothing when none), for tests.
Public Function ControlFor(ByVal node As clsNode) As Object
    Dim item As clsRenderItem
    If mNodeItems Is Nothing Or node Is Nothing Then Exit Function
    Set item = Obj(mNodeItems, CStr(node.Uid))
    If Not item Is Nothing Then Set ControlFor = item.Ctl
End Function

' Test hooks: drive the window's event handlers for an element's control.
Public Function SimulateKey(ByVal node As clsNode, ByVal keyCode As Integer, ByVal shift As Integer) As Boolean
    Dim item As clsRenderItem
    Set item = Obj(mNodeItems, CStr(node.Uid))
    If item Is Nothing Then Exit Function
    On Error Resume Next
    item.Ctl.SetFocus
    On Error GoTo 0
    SimulateKey = OnItemKeyDown(item, keyCode, shift)
End Function

' Test hook: a click on the control drawn for an element (its own control, or a label or
' hit area whose click activates it), through the window's event path. False when the
' element has no control.
Public Function SimulateClick(ByVal node As clsNode) As Boolean
    Dim item As clsRenderItem, k As Variant, cand As clsRenderItem
    If mItems Is Nothing Or node Is Nothing Then Exit Function
    Set item = Obj(mNodeItems, CStr(node.Uid))
    If item Is Nothing Then
        Dim keys1 As clsList: Set keys1 = mItems.Keys
        For Each k In keys1
            Set cand = mItems(CStr(k))
            If Not cand.Spec Is Nothing Then
                If cand.Spec.ClickNode Is node Then
                    If cand.Kind = "hit" Then Set item = cand: Exit For
                    If item Is Nothing Then Set item = cand
                End If
            End If
        Next k
    End If
    If item Is Nothing Then Exit Function
    If item.Kind = "check" Then
        ' Clicking an MSForms toggle flips its value and raises Click.
        item.Ctl.Value = Not CBool(item.Ctl.Value)
    Else
        OnItemClick item
    End If
    SimulateClick = True
End Function

Public Sub SimulateResizerDrag(ByVal fromX As Double, ByVal toX As Double)
    Dim k As Variant, item As clsRenderItem, ax As Double, ay As Double
    Dim keys2 As clsList: Set keys2 = mItems.Keys
    For Each k In keys2
        Set item = mItems(CStr(k))
        If Not item.Spec Is Nothing Then
            If item.Spec.Resizer Then
                ' MSForms reports positions relative to the control, which moves as the
                ' library resizes.
                AbsoluteOrigin item.Ctl, ax, ay
                OnItemMouse item, 0, 1, fromX * mPx - ax, 1
                AbsoluteOrigin item.Ctl, ax, ay
                OnItemMouse item, 1, 1, (fromX + (toX - fromX) / 2) * mPx - ax, 1
                AbsoluteOrigin item.Ctl, ax, ay
                OnItemMouse item, 1, 1, toX * mPx - ax, 1
                AbsoluteOrigin item.Ctl, ax, ay
                OnItemMouse item, 2, 1, toX * mPx - ax, 1
                Exit Sub
            End If
        End If
    Next k
End Sub

' The DOM element whose control has the window's focus (Nothing when none).
Public Function FocusedNode() As clsNode
    Dim c As Object, item As clsRenderItem
    Set c = ActiveLeaf()
    If c Is Nothing Then Exit Function
    Set item = Obj(mCtlItems, c.Name)
    If item Is Nothing Then Exit Function
    Set FocusedNode = ItemNode(item)
End Function

' Window client size in CSS px.
Public Sub SetClientSize(ByVal width As Double, ByVal height As Double)
    If mForm Is Nothing Then Exit Sub
    mForm.Width = width * mPx + (mForm.Width - mForm.InsideWidth)
    mForm.Height = height * mPx + (mForm.Height - mForm.InsideHeight)
End Sub

' Diagnostics: the frames and how many controls each holds.
Public Function DescribeWindow() As String
    Dim k As Variant, fr As Object, out As String
    If mFrames Is Nothing Then Exit Function
    Dim keys3 As clsList: Set keys3 = mFrames.Keys
    For Each k In keys3
        Set fr = mFrames(CStr(k))
        If k = "page" Then GoTo NextFrame
        out = out & k & " [" & fr.Name & " in " & TypeName(fr.Parent) & "] " & fr.Left & "," & fr.Top & " " & fr.Width & "x" & fr.Height & _
            " visible " & fr.Visible & " controls " & fr.Controls.Count & " scroll " & fr.ScrollBars & "/" & fr.ScrollHeight & vbLf
NextFrame:
    Next k
    If Not mForm Is Nothing Then out = out & "form inside " & mForm.InsideWidth & "x" & mForm.InsideHeight & " controls " & mForm.Controls.Count & vbLf
    DescribeWindow = out
End Function

' The frame drawn for a region ("main", "library", "top", ...), Nothing when none (tests).
Public Function RegionFrame(ByVal id As String) As Object
    If mFrames Is Nothing Then Exit Function
    Set RegionFrame = Obj(mFrames, id)
End Function

Public Function AppWindowOpen() As Boolean
    AppWindowOpen = (mShown And Not mForm Is Nothing)
End Function

' The app window form (Nothing when closed). Tests use it to capture screenshots.
Public Function AppForm() As Object
    Set AppForm = mForm
End Function

' Uses this rule table instead of the stored one (tests).
Public Sub LoadStyleText(ByVal text As String)
    mStyleText = text
    modStyle.StyleLoad text
End Sub

' Uses these fonts (base64 TTF by weight "400"/"600"/"700") instead of the stored ones (tests).
Public Sub UseFontData(ByVal weight As String, ByVal base64 As String)
    If mFontData Is Nothing Then Set mFontData = NewMap()
    mFontData(weight) = base64
End Sub

' Computed boxes of the current DOM at width x height CSS px, without a window: one line
' per region ("R", id, parent, x, y, w, h, content height) and per item ("I", region, id,
' kind, x, y, w, h, text), tab-separated, ids from element paths. Measures text with the
' window's (hidden) form.
Public Function LayoutSnapshot(ByVal width As Double, ByVal height As Double) As String
    Dim regions As Collection, r As clsRenderSpec, it As clsRenderSpec, parts() As String, n As Long
    EnsureEngine
    EnsureStyles
    EnsureFonts
    EnsureHost
    mLayout.Run width, height
    mPaint.PathIds = True
    mPaint.FocusVisible = False
    mPaint.Paint
    mPaint.PathIds = False
    Set regions = mPaint.Regions()
    ReDim parts(0 To 63)
    For Each r In regions
        AddLine parts, n, "R" & vbTab & r.Id & vbTab & r.Region & vbTab & Num(r.X) & vbTab & Num(r.Y) & vbTab & Num(r.W) & vbTab & Num(r.H) & vbTab & Num(r.ContentH)
        For Each it In r.Items
            AddLine parts, n, "I" & vbTab & r.Id & vbTab & it.Id & vbTab & it.Kind & vbTab & Num(it.X) & vbTab & Num(it.Y) & vbTab & Num(it.W) & vbTab & Num(it.H) & vbTab & Clean(it.Text)
        Next it
    Next r
    modStyle.StyleEnd
    ReDim Preserve parts(0 To n)
    LayoutSnapshot = Join(parts, vbLf)
End Function

' Forgets the incremental layout caches (tests compare incremental and full layouts).
Public Sub ResetLayoutCache()
    If Not mLayout Is Nothing Then mLayout.ClearLayoutCache
End Sub

' Cache statistics of the last layout pass (tests).
Public Function LayoutCacheStats() As String
    If mLayout Is Nothing Then Exit Function
    LayoutCacheStats = "walked " & mLayout.WalkedNodes & ", dirty " & mLayout.DirtyNodes & ", hits " & mLayout.CacheHits & ", misses " & mLayout.CacheMisses
End Function

Private Sub AddLine(ByRef parts() As String, ByRef n As Long, ByVal line As String)
    If n > UBound(parts) Then ReDim Preserve parts(0 To n * 2)
    parts(n) = line
    n = n + 1
End Sub

Private Function Num(ByVal v As Double) As String
    Num = modJs.JsNumberToString(Round(v, 2))
End Function

Private Function Clean(ByVal t As String) As String
    Clean = Replace(Replace(Replace(t, vbTab, " "), vbCr, " "), vbLf, "\n")
End Function

' ================================================================ setup

' Turns batched painting (one redraw per update) on or off (tests compare both).
Public Sub SetPaintBatching(ByVal enabled As Boolean)
    mSuspendEnabled = enabled
End Sub

Private Sub EnsureEngine()
    If Not mEngineInit Then mSuspendEnabled = True: mEngineInit = True
    If mLayout Is Nothing Then
        Set mLayout = New clsRenderLayout
        Set mPaint = New clsRenderPaint
        Set mPaint.Engine = mLayout
    End If
    If mItems Is Nothing Then ResetControls
    SetPx
End Sub

' Points per CSS px: Windows draws 96 px per inch (0.75 pt); a Mac draws 1 px per point.
Private Sub SetPx()
    #If Mac Then
        mPx = 1
    #Else
        mPx = 0.75
    #End If
End Sub

Private Sub EnsureStyles()
    Dim text As String
    If modStyle.StyleLoaded() Then Exit Sub
    If Len(mStyleText) > 0 Then
        text = mStyleText
    Else
        text = modBoot.StoredText("css")
    End If
    If Len(text) = 0 Then Err.Raise vbObjectError + 830, "modRender", "The stylesheet rules are missing from this workbook"
    modStyle.StyleLoad text
End Sub

' The form instance: the window when shown, the measuring host otherwise.
Private Sub EnsureHost()
    If Not mForm Is Nothing Then Exit Sub
    On Error GoTo NoForm
    Set mForm = VBA.UserForms.Add("frmApp")
    mForm.Caption = APP_TITLE
    Set mMeasurers = Nothing
    Exit Sub
NoForm:
    Set mForm = Nothing
    mNoMeasure = True
End Sub

Private Sub SizeWindow()
    Dim w As Double, h As Double, maxW As Double, maxH As Double
    On Error Resume Next
    maxW = Application.UsableWidth * 0.96
    maxH = Application.UsableHeight * 0.96
    If maxW < 400 Then maxW = 960
    If maxH < 300 Then maxH = 645
    w = DEFAULT_WIDTH * mPx + (mForm.Width - mForm.InsideWidth)
    h = DEFAULT_HEIGHT * mPx + (mForm.Height - mForm.InsideHeight)
    If w > maxW Then w = maxW
    If h > maxH Then h = maxH
    mForm.Width = w
    mForm.Height = h
    mForm.Left = Application.Left + (Application.Width - w) / 2
    mForm.Top = Application.Top + (Application.Height - h) / 2
    If mForm.Left < 0 Then mForm.Left = 0
    If mForm.Top < 0 Then mForm.Top = 0
End Sub

Private Sub ResetControls()
    Dim k As Variant, it As clsRenderItem
    If Not mItems Is Nothing Then
        Dim keys4 As clsList: Set keys4 = mItems.Keys
        For Each k In keys4
            Set it = mItems(CStr(k))
            it.Dispose
        Next k
    End If
    If Not mIdle Is Nothing Then
        Dim keys5 As clsList: Set keys5 = mIdle.Keys
        For Each k In keys5
            For Each it In mIdle(CStr(k))
                it.Dispose
            Next it
        Next k
    End If
    Set mIdle = Nothing
    Set mItems = NewMap()
    Set mCtlItems = NewMap()
    Set mNodeItems = NewMap()
    Set mFrames = NewMap()
    Set mFrameParent = NewMap()
    Set mFrameSpecs = NewMap()
    Set mFocusNode = Nothing
    Set mSinks = Nothing
    Set mFramePool = Nothing
    If Not mScrollSink Is Nothing Then Set mScrollSink.Frm = Nothing
    Set mScrollSink = Nothing
    mRenderedVersion = -1
End Sub

' Sets window.innerWidth for the website's code and dispatches "resize" when it changed.
Private Sub UpdateViewport(ByVal force As Boolean)
    Dim w As Double, h As Double
    If mForm Is Nothing Then Exit Sub
    w = mForm.InsideWidth / mPx
    h = mForm.InsideHeight / mPx
    If Not force And Abs(w - mRenderedW) < 0.5 And Abs(h - mRenderedH) < 0.5 Then Exit Sub
    SetInnerWidth w
    On Error GoTo Failed
    modDom.Win.DispatchEvent modDom.NewEvent("resize")
    modDom.Settle 700
    Exit Sub
Failed:
    ReportError "modRender.UpdateViewport"
End Sub

' window.innerWidth lives in the shell (modApp).
Private Sub SetInnerWidth(ByVal w As Double)
    modApp.WindowInnerWidth = CLng(w)
End Sub

Private Function Qualified(ByVal procedure As String) As String
    Qualified = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & procedure
End Function

' ================================================================ fonts

Private Sub EnsureFonts()
    If mPx = 0 Then SetPx
    If mFontsReady Then Exit Sub
    mFontRegular = "Arial"
    mFontSemibold = "Arial"
    mSemiboldIsBold = True
    mFontSymbol = "Arial"
    mSymbolFallback = False
    #If Mac Then
    #Else
        mFontSymbol = "Segoe UI Symbol"
        mSymbolFallback = True
        If InstallPrivateFonts() Then
            mFontRegular = "Source Sans Pro"
            mFontSemibold = "Source Sans Pro SemiBold"
            mSemiboldIsBold = False
        End If
    #End If
    mFontsReady = True
    Set mMeasurers = Nothing
    Set mFontObjects = Nothing
    If Not mLayout Is Nothing Then mLayout.ClearMeasureCache
End Sub

' True when text runs drawn with the text font should switch to the symbol font for
' glyphs the text font lacks.
Public Function SymbolFallbackActive() As Boolean
    SymbolFallbackActive = mSymbolFallback
End Function

Public Function NeedsSymbolFont(ByVal text As String) As Boolean
    Dim i As Long, c As Long
    For i = 1 To Len(text)
        c = AscW(Mid$(text, i, 1)) And &HFFFF&
        If c >= 256 Then
            If Not modStyle.GlyphCovered(c) Then NeedsSymbolFont = True: Exit Function
        End If
    Next i
End Function

Public Function FontNames() As String
    FontNames = mFontRegular & " / " & mFontSemibold & " / " & mFontSymbol
End Function

#If Mac Then
#Else
' Registers the embedded Source Sans Pro faces for this Excel process only.
Private Function InstallPrivateFonts() As Boolean
    Dim folder As String, weights As Variant, i As Long, path As String, data() As Byte, handle As Integer, b64 As String
    Dim sep As String
    On Error GoTo Failed
    sep = Application.PathSeparator
    folder = Environ$("TEMP") & sep & "BT1101 Prose Lab"
    If Len(Dir$(folder, vbDirectory)) = 0 Then MkDir folder
    Set mFontFiles = New Collection
    weights = Array("400", "600", "700")
    For i = 0 To 2
        b64 = FontBase64(CStr(weights(i)))
        If Len(b64) = 0 Then GoTo Failed
        data = DecodeBase64(b64)
        path = folder & sep & "SourceSansPro-" & weights(i) & ".ttf"
        If Len(Dir$(path)) > 0 Then
            If FileLen(path) <> UBound(data) + 1 Then
                On Error Resume Next
                Kill path
                On Error GoTo Failed
            End If
        End If
        If Len(Dir$(path)) = 0 Then
            handle = FreeFile
            Open path For Binary Access Write As #handle
            Put #handle, , data
            Close #handle
        End If
        If AddFontResourceEx(StrPtr(path), FR_PRIVATE, 0) > 0 Then mFontFiles.Add path
    Next i
    InstallPrivateFonts = (mFontFiles.Count = 3)
    Exit Function
Failed:
    InstallPrivateFonts = False
End Function
#End If

Private Function FontBase64(ByVal weight As String) As String
    Dim v As Variant
    If Not mFontData Is Nothing Then
        v = mFontData(weight)
        If Not IsEmpty(v) Then FontBase64 = CStr(v): Exit Function
    End If
    On Error Resume Next
    FontBase64 = modBoot.StoredText("font" & weight)
    Err.Clear
End Function

Private Sub RemoveFonts()
    #If Mac Then
    #Else
        Dim path As Variant
        If mFontFiles Is Nothing Then Exit Sub
        For Each path In mFontFiles
            RemoveFontResourceEx StrPtr(CStr(path)), FR_PRIVATE, 0
        Next path
        Set mFontFiles = Nothing
    #End If
End Sub

' Pure-VBA base64 decoder.
Public Function DecodeBase64(ByVal text As String) As Byte()
    Dim table(0 To 255) As Long, i As Long, n As Long, out() As Byte, src() As Byte, v As Long, bits As Long
    Dim acc As Long, o As Long, c As Long, alphabet As String
    alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    For i = 0 To 255
        table(i) = -1
    Next i
    For i = 1 To 64
        table(Asc(Mid$(alphabet, i, 1))) = i - 1
    Next i
    src = text                          ' UTF-16 bytes: every other byte is the character
    n = (UBound(src) + 1) \ 2
    ReDim out(0 To (n * 3) \ 4 + 2)
    For i = 0 To n - 1
        c = src(i * 2)
        If src(i * 2 + 1) = 0 Then
            v = table(c)
            If v >= 0 Then
                acc = (acc * 64 + v) And &HFFFFFF
                bits = bits + 6
                If bits >= 8 Then
                    bits = bits - 8
                    out(o) = (acc \ (2 ^ bits)) And &HFF
                    o = o + 1
                End If
            End If
        End If
    Next i
    If o = 0 Then
        ReDim out(0 To 0)
    Else
        ReDim Preserve out(0 To o - 1)
    End If
    DecodeBase64 = out
End Function

' ================================================================ measuring

Private Function MeasureKey(ByVal size As Double, ByVal weight As Long, ByVal italic As Boolean, ByVal symbol As Boolean) As String
    MeasureKey = Round(size, 2) & "|" & weight & "|" & italic & "|" & symbol
End Function

Private Function Measurer(ByVal size As Double, ByVal weight As Long, ByVal italic As Boolean, ByVal symbol As Boolean, ByRef overhead As Double) As Object
    Dim key As String, lbl As Object, v As Variant, w10 As Double, w20 As Double
    key = MeasureKey(size, weight, italic, symbol)
    If mMeasurers Is Nothing Then
        Set mMeasurers = NewMap()
        Set mOverhead = NewMap()
        Set mLineHeights = NewMap()
    End If
    Set lbl = Obj(mMeasurers, key)
    If lbl Is Nothing Then
        mCounter = mCounter + 1
        #If Mac Then
            ' A hidden AutoSize label may not resize on a Mac: measure visibly, off-screen.
            Set lbl = mForm.Controls.Add("Forms.Label.1", "m" & mCounter, True)
            lbl.Left = -10000
            lbl.Top = -10000
        #Else
            Set lbl = mForm.Controls.Add("Forms.Label.1", "m" & mCounter, False)
        #End If
        lbl.AutoSize = False
        lbl.WordWrap = False
        ApplyFont lbl, size, weight, italic, symbol
        lbl.AutoSize = True
        lbl.Caption = String$(10, "x")
        w10 = lbl.Width
        lbl.Caption = String$(20, "x")
        w20 = lbl.Width
        ' Ten more x's must add about five em: otherwise labels do not measure here.
        If w20 - w10 < size * mPx * 2 Then mMeasureBroken = True
        mOverhead(key) = 2 * w10 - w20
        lbl.Caption = "Xg"
        mLineHeights(key) = lbl.Height / mPx
        Set mMeasurers(key) = lbl
    End If
    overhead = mOverhead(key)
    Set Measurer = lbl
End Function

' Width in CSS px of one line of text drawn with a Label in this font.
Public Function MeasureText(ByVal text As String, ByVal sizePx As Double, ByVal weight As Long, ByVal italic As Boolean) As Double
    Dim i As Long, startPos As Long, cur As Boolean, sym As Boolean
    If Len(text) = 0 Then Exit Function
    If Not mFontsReady Then EnsureFonts
    If mForm Is Nothing And Not mNoMeasure Then EnsureHost
    If mForm Is Nothing Or mMeasureBroken Then
        MeasureText = TableWidth(text, sizePx, weight)
        Exit Function
    End If
    If mSymbolFallback Then
        If NeedsSymbolFont(text) Then
            startPos = 1
            cur = Not modStyle.GlyphCovered(AscW(Mid$(text, 1, 1)) And &HFFFF&)
            For i = 2 To Len(text) + 1
                If i <= Len(text) Then sym = Not modStyle.GlyphCovered(AscW(Mid$(text, i, 1)) And &HFFFF&) Else sym = Not cur
                If sym <> cur Then
                    MeasureText = MeasureText + MeasurePiece(Mid$(text, startPos, i - startPos), sizePx, weight, italic, cur)
                    startPos = i
                    cur = sym
                End If
            Next i
            Exit Function
        End If
    End If
    MeasureText = MeasurePiece(text, sizePx, weight, italic, False)
End Function

' Estimated width (CSS px) from Arial's advance widths, when labels cannot measure.
Private Function TableWidth(ByVal text As String, ByVal sizePx As Double, ByVal weight As Long) As Double
    Dim i As Long, c As Long, units As Double
    For i = 1 To Len(text)
        c = AscW(Mid$(text, i, 1)) And &HFFFF&
        If c >= 32 And c <= 126 Then
            units = units + Val(Mid$(ARIAL_WIDTHS, (c - 32) * 3 + 1, 3))
        Else
            units = units + 556
        End If
    Next i
    ' "@" is 1015: the table holds its last three digits.
    units = units + 1000 * (Len(text) - Len(Replace(text, "@", "")))
    TableWidth = units / 1000 * sizePx
    If weight >= 600 Then TableWidth = TableWidth * 1.08
End Function

Private Function MeasurePiece(ByVal text As String, ByVal sizePx As Double, ByVal weight As Long, ByVal italic As Boolean, ByVal symbol As Boolean) As Double
    Dim lbl As Object, overhead As Double, w As Double, t0 As Double
    t0 = ClockMs()
    LastMeasureCount = LastMeasureCount + 1
    Set lbl = Measurer(sizePx, weight, italic, symbol, overhead)
    If Len(Trim$(text)) = 0 Then
        lbl.Caption = "x" & text & "x"
        w = lbl.Width
        lbl.Caption = "xx"
        w = w - lbl.Width
    Else
        lbl.Caption = text
        w = lbl.Width - overhead
    End If
    If w < 0 Then w = 0
    MeasurePiece = w / mPx
    LastMeasureMs = LastMeasureMs + ClockMs() - t0
End Function

' Height in CSS px a one-line Label needs for this font.
Public Function NaturalLineHeight(ByVal sizePx As Double, ByVal weight As Long) As Double
    Dim overhead As Double
    If Not mFontsReady Then EnsureFonts
    If mForm Is Nothing And Not mNoMeasure Then EnsureHost
    If mForm Is Nothing Then NaturalLineHeight = sizePx * 1.26: Exit Function
    Measurer sizePx, weight, False, False, overhead
    If mMeasureBroken Then NaturalLineHeight = sizePx * 1.15: Exit Function
    NaturalLineHeight = mLineHeights(MeasureKey(sizePx, weight, False, False))
End Function

' ================================================================ reconciling controls

' Sets a cosmetic or platform-specific control property, ignoring an error (a property
' Mac MSForms lacks must not abort the whole render).
Private Sub TrySet(ByVal ctl As Object, ByVal prop As String, ByVal value As Variant)
    On Error Resume Next
    CallByName ctl, prop, VbLet, value
    Err.Clear
End Sub

Private Function NewControl(ByVal container As Object, ByVal progId As String) As Object
    mCounter = mCounter + 1
    Set NewControl = container.Add(progId, "c" & mCounter, True)
End Function

Private Sub ApplyDisplayList(ByVal regions As Collection, ByVal pageBg As Long)
    Dim r As clsRenderSpec, it As clsRenderSpec, item As clsRenderItem, k As Variant
    Dim seenRegions As clsMap, seenItems As clsMap, pending As Collection, pool As Collection, key As String
    Dim newMin As clsMap, fr As Object, created As Long, updated As Long, total As Long, reused As Long, framesChanged As Boolean
    Dim frameCount As Long
    mApplying = True
    On Error GoTo Failed
    Set seenRegions = NewMap()
    Set seenItems = NewMap()
    Set pending = New Collection
    Set newMin = NewMap()
    If mIdle Is Nothing Then Set mIdle = NewMap()
    If mForm.BackColor <> pageBg Then WillWrite: mForm.BackColor = pageBg
    ' Regions (parents come before their children).
    frameCount = mFrames.Count
    For Each r In regions
        If ApplyRegion(r) Then framesChanged = True
        seenRegions(r.Id) = True
    Next r
    Dim keys6 As clsList: Set keys6 = mFrames.Keys
    For Each k In keys6
        If IsEmpty(seenRegions(CStr(k))) Then RemoveRegion CStr(k): framesChanged = True
    Next k
    ' Items that keep their id.
    For Each r In regions
        For Each it In r.Items
            total = total + 1
            seenItems(it.Id) = True
            Set item = Obj(mItems, it.Id)
            If Not item Is Nothing Then
                If item.Kind = it.Kind And item.RegionId = r.Id And Not (it.Kind = "check" And item.Spec.IsRadio <> it.IsRadio) Then
                    UpdateItem item, it
                    updated = updated + 1
                Else
                    RetireItem item
                    pending.Add it
                End If
            Else
                pending.Add it
            End If
        Next it
    Next r
    ' Items whose id disappeared wait, hidden, in their frame's pool for reuse.
    Dim keys7 As clsList: Set keys7 = mItems.Keys
    For Each k In keys7
        If IsEmpty(seenItems(CStr(k))) Then RetireItem mItems(CStr(k))
    Next k
    For Each it In pending
        Set item = Nothing
        Set pool = Obj(mIdle, IdleKey(ContainerKey(mFrames(it.Region)), it.Kind, PoolFlavour(it)))
        If Not pool Is Nothing Then
            If pool.Count > 0 Then
                Set item = pool(pool.Count)
                pool.Remove pool.Count
            End If
        End If
        If item Is Nothing Then
            WillWrite
            Set item = CreateItem(it)
            created = created + 1
        Else
            reused = reused + 1
        End If
        item.Id = it.Id
        item.RegionId = it.Region
        Set mItems(it.Id) = item
        UpdateItem item, it
        If Not item.Ctl.Visible Then WillWrite: item.Ctl.Visible = True
        If IsEmpty(newMin(it.Region)) Then
            newMin(it.Region) = it.PaintIndex
        ElseIf it.PaintIndex < newMin(it.Region) Then
            newMin(it.Region) = it.PaintIndex
        End If
    Next it
    TrimIdle
    ' New and reused controls go on top of their frame; whatever paints after them in the
    ' region is raised again so the z-order follows paint order. Focusable controls get
    ' the paint order as tab order.
    For Each r In regions
        If Not IsEmpty(newMin(r.Id)) Then
            Restack r, CLng(newMin(r.Id))
            SetTabOrder r
        End If
    Next r
    If created > 0 Or framesChanged Then SetFrameTabOrder regions
    ' MSForms does not clip a frame's repaint against frames above it: overlays stay on top.
    For Each r In regions
        If r.Overlay Then
            Set fr = mFrames(r.Id)
            If framesChanged Or mPaintSuspended Then
                WillWrite
                fr.ZOrder 0
            End If
        End If
    Next r
    LastControlsCreated = created
    LastControlsUpdated = updated
    LastControlsReused = reused
    LastItemCount = total
    mApplying = False
    Exit Sub
Failed:
    mApplying = False
    ReportError "modRender.ApplyDisplayList"
End Sub

Private Function IdleKey(ByVal container As String, ByVal kind As String, ByVal flavour As String) As String
    IdleKey = container & "|" & kind & "|" & flavour
End Function

' Takes an item off the window: hidden and kept in its frame's pool for reuse.
Private Sub RetireItem(ByVal item As clsRenderItem)
    Dim pool As Collection, key As String
    On Error GoTo Gone
    If Not IsEmpty(mItems(item.Id)) Then
        If mItems(item.Id) Is item Then mItems.Remove item.Id
    End If
    UnmapItem item
    If item.Ctl Is Nothing Then Exit Sub
    ParkFocusIfInside item.Ctl
    If item.Ctl.Visible Then WillWrite: item.Ctl.Visible = False
    If Not item.Hint Is Nothing Then
        If item.Hint.Visible Then WillWrite: item.Hint.Visible = False
    End If
    key = IdleKey(ContainerKey(item.Ctl.Parent), item.Kind, PoolFlavour(item.Spec))
    Set pool = Obj(mIdle, key)
    If pool Is Nothing Then
        Set pool = New Collection
        Set mIdle(key) = pool
    End If
    pool.Add item
    Exit Sub
Gone:
    ' The control no longer exists (its frame went away): forget it.
    On Error Resume Next
    item.Dispose
End Sub

' Keeps each pool to a bounded number of hidden controls.
Private Sub TrimIdle()
    Dim k As Variant, pool As Collection
    Dim keys8 As clsList: Set keys8 = mIdle.Keys
    For Each k In keys8
        Set pool = mIdle(CStr(k))
        Do While pool.Count > IDLE_LIMIT
            DeleteItem pool(pool.Count)
            pool.Remove pool.Count
        Loop
    Next k
End Sub

' Before the first control change of an update: suspends painting so the window is
' redrawn once, after the update (Windows; a Mac draws the changes as they come).
Private Sub WillWrite()
    If mPaintSuspended Then Exit Sub
    #If Mac Then
    #Else
        If mHwnd = 0 Or Not mSuspendEnabled Then Exit Sub
        mClientHwnd = GetWindow(mHwnd, GW_CHILD)
        If mClientHwnd = 0 Then Exit Sub
        If IsWindowVisible(mClientHwnd) = 0 Then Exit Sub
        SendMessage mClientHwnd, WM_SETREDRAW, 0, 0
        mPaintSuspended = True
    #End If
End Sub

' Ends a suspended update: painting resumes and the window is redrawn once.
Private Sub ResumePaint()
    If Not mPaintSuspended Then Exit Sub
    mPaintSuspended = False
    #If Mac Then
    #Else
        SendMessage mClientHwnd, WM_SETREDRAW, 1, 0
        RedrawWindow mClientHwnd, 0, 0, RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN Or RDW_FRAME Or RDW_UPDATENOW
    #End If
End Sub

Private Function PoolFlavour(ByVal it As clsRenderSpec) As String
    If it Is Nothing Then Exit Function
    If it.Kind = "check" And it.IsRadio Then PoolFlavour = "r"
End Function

Private Function ApplyRegion(ByVal r As clsRenderSpec) As Boolean
    Dim v As Variant, fr As Object, container As Object, old As clsRenderSpec, sb As Long, created As Boolean, maxTop As Double
    Set fr = Obj(mFrames, r.Id)
    If Not fr Is Nothing Then
        If CStr(mFrameParent(r.Id)) <> r.Region Then
            RemoveRegion r.Id
            Set fr = Nothing
        End If
    End If
    ' The page region is the form itself: frames (child windows) always draw above it.
    If r.Id = "page" Then
        If fr Is Nothing Then
            Set mFrames(r.Id) = mForm
            mFrameParent(r.Id) = ""
        End If
        If mForm.BackColor <> r.Bg Then WillWrite: mForm.BackColor = r.Bg
        Set mFrameSpecs(r.Id) = r
        Exit Function
    End If
    If fr Is Nothing Then
        If Len(r.Region) > 0 Then Set fr = PooledFrame(ContainerKey(mFrames(r.Region))) Else Set fr = PooledFrame("form")
        If Not fr Is Nothing Then
            WillWrite
            fr.Visible = True
            Set mFrames(r.Id) = fr
            mFrameParent(r.Id) = r.Region
            created = True
        End If
    End If
    If fr Is Nothing Then
        If Len(r.Region) > 0 Then Set container = mFrames(r.Region).Controls Else Set container = mForm.Controls
        WillWrite
        On Error GoTo AddFailed
        Set fr = NewControl(container, "Forms.Frame.1")
        On Error GoTo 0
        fr.Caption = ""
        fr.BorderStyle = 0
        TrySet fr, "SpecialEffect", 0
        TrySet fr, "KeepScrollBarsVisible", 0
        fr.ScrollBars = 0
        Set mFrames(r.Id) = fr
        mFrameParent(r.Id) = r.Region
        created = True
        ' Frames added to a frame after it has been in use can fail (RPC_E_DISCONNECTED):
        ' a new top-level region frame gets its nested scroller frames up front.
        If Len(r.Region) = 0 And Not r.Overlay Then PrecreateFrames fr, NESTED_FRAMES
    Else
        Set old = Obj(mFrameSpecs, r.Id)
    End If
    If created Or old Is Nothing Then
        WillWrite
        fr.Move r.X * mPx, r.Y * mPx, Pos(r.W * mPx), Pos(r.H * mPx)
        ApplyRegion = True
    ElseIf old.Geometry() <> r.Geometry() Then
        WillWrite
        fr.Move r.X * mPx, r.Y * mPx, Pos(r.W * mPx), Pos(r.H * mPx)
        ApplyRegion = True
    End If
    If fr.BackColor <> r.Bg Then WillWrite: fr.BackColor = r.Bg
    sb = 0
    If r.Scrolls Then sb = sb + 2
    If r.ScrollsX Then sb = sb + 1
    If fr.ScrollBars <> sb Then WillWrite: fr.ScrollBars = sb
    If r.Scrolls Then
        If Abs(fr.ScrollHeight - r.ContentH * mPx) > 0.5 Then WillWrite: fr.ScrollHeight = r.ContentH * mPx
        ' MSForms keeps the old ScrollTop when the content got shorter: clamp it.
        maxTop = fr.ScrollHeight - fr.InsideHeight
        If maxTop < 0 Then maxTop = 0
        If fr.ScrollTop > maxTop + 0.5 Then WillWrite: fr.ScrollTop = maxTop
    Else
        ' A content height left from an earlier, longer page would still let MSForms scroll
        ' the frame (to a focused control, or by code) into blank space.
        If fr.ScrollTop <> 0 Then WillWrite: fr.ScrollTop = 0
        If fr.ScrollHeight <> 0 Then WillWrite: fr.ScrollHeight = 0
    End If
    If r.ScrollsX Then
        If Abs(fr.ScrollWidth - r.ContentW * mPx) > 0.5 Then WillWrite: fr.ScrollWidth = r.ContentW * mPx
    Else
        If fr.ScrollLeft <> 0 Then WillWrite: fr.ScrollLeft = 0
        If fr.ScrollWidth <> 0 Then WillWrite: fr.ScrollWidth = 0
    End If
    Set mFrameSpecs(r.Id) = r
    Exit Function
AddFailed:
    Err.Raise Err.Number, "modRender.ApplyRegion", Err.Description & " [frame for region " & r.Id & " in '" & r.Region & "']"
End Function

' MSForms becomes unstable (and Excel can crash) when the focused control or a frame
' holding it is removed: focus moves to an invisible sink first, and back to the DOM's
' focused element after the update (ReflectFocus).
Private Sub ParkFocusIfInside(ByVal target As Object)
    Dim par As Object, act As Object, sink As Object, key As String, real As Object
    On Error GoTo Done
    ' Every frame remembers its own active control even while focus is elsewhere; when
    ' that control is removed the frame refuses new frames (RPC_E_DISCONNECTED) and Excel
    ' can crash. The container's active control moves to a sink of its own first.
    Set par = target.Parent
    Set act = par.ActiveControl
    If Not act Is target Then Exit Sub
    key = ContainerKey(par)
    If mSinks Is Nothing Then Set mSinks = NewMap()
    Set sink = Obj(mSinks, key)
    If sink Is Nothing Then
        mCounter = mCounter + 1
        Set sink = par.Controls.Add("Forms.CommandButton.1", "sink" & mCounter, True)
        sink.Move -60, -60, 8, 8
        sink.TabStop = False
        TrySet sink, "TakeFocusOnClick", False
        Set mSinks(key) = sink
    End If
    Set real = ActiveLeaf()
    sink.SetFocus
    ' The window's focus left the element (ReflectFocus restores it after the update).
    If Not real Is Nothing Then
        If Not ActiveLeaf() Is real Then Set mFocusNode = Nothing
    End If
Done:
End Sub

Private Sub PrecreateFrames(ByVal parent As Object, ByVal count As Long)
    Dim i As Long, f As Object
    On Error GoTo Done
    For i = 1 To count
        Set f = NewControl(parent.Controls, "Forms.Frame.1")
        f.Caption = ""
        f.BorderStyle = 0
        TrySet f, "SpecialEffect", 0
        TrySet f, "KeepScrollBarsVisible", 0
        f.ScrollBars = 0
        f.Move 0, 0, 10, 10
        f.Visible = False
        PoolFrame ContainerKey(parent), f
    Next i
Done:
End Sub

' Frame pools are per containing control (a frame's name, or "form").
Private Function ContainerKey(ByVal container As Object) As String
    If TypeName(container) = "Frame" Then ContainerKey = container.Name Else ContainerKey = "form"
End Function

Private Sub PoolFrame(ByVal parentId As String, ByVal fr As Object)
    Dim pool As Collection
    If mFramePool Is Nothing Then Set mFramePool = NewMap()
    Set pool = Obj(mFramePool, parentId)
    If pool Is Nothing Then
        Set pool = New Collection
        Set mFramePool(parentId) = pool
    End If
    pool.Add fr
End Sub

Private Function PooledFrame(ByVal parentId As String) As Object
    Dim pool As Collection
    If mFramePool Is Nothing Then Exit Function
    Set pool = Obj(mFramePool, parentId)
    If pool Is Nothing Then Exit Function
    If pool.Count = 0 Then Exit Function
    Set PooledFrame = pool(pool.Count)
    pool.Remove pool.Count
End Function

Private Sub RemoveRegion(ByVal id As String)
    Dim k As Variant, item As clsRenderItem, fr As Object, child As Variant
    If IsEmpty(mFrames(id)) Then Exit Sub
    ' Nested regions go first.
    Dim keys9 As clsList: Set keys9 = mFrameParent.Keys
    For Each child In keys9
        If CStr(mFrameParent(CStr(child))) = id Then RemoveRegion CStr(child)
    Next child
    Set fr = mFrames(id)
    If id <> "page" Then ParkFocusIfInside fr
    Dim keys10 As clsList: Set keys10 = mItems.Keys
    For Each k In keys10
        Set item = mItems(CStr(k))
        If item.RegionId = id Then RetireItem item
    Next k
    ' MSForms fails (and Excel may crash) when frames are removed from a frame and new ones
    ' added to it later: frames are hidden and reused instead of removed.
    If id <> "page" Then
        On Error Resume Next
        WillWrite
        fr.Visible = False
        fr.ScrollBars = 0
        fr.ScrollTop = 0
        fr.ScrollLeft = 0
        On Error GoTo 0
        PoolFrame ContainerKey(fr.Parent), fr
    End If
    mFrames.Remove id
    mFrameParent.Remove id
    If Not IsEmpty(mFrameSpecs(id)) Then mFrameSpecs.Remove id
End Sub

' An object stored in a map (Nothing when absent). Reading it into a Variant would
' evaluate the object's default member instead.
Private Function Obj(ByVal m As clsMap, ByVal key As String) As Object
    If m.Exists(key) Then Set Obj = m(key)
End Function

Private Function Pos(ByVal v As Double) As Double
    If v > 0 Then Pos = v
End Function

Private Function CreateItem(ByVal it As clsRenderSpec) As clsRenderItem
    Dim item As New clsRenderItem, container As Object
    On Error GoTo Failed
    item.Kind = it.Kind
    item.RegionId = it.Region
    Set container = mFrames(it.Region).Controls
    Set item.Container = container
    Select Case it.Kind
        Case "input", "textarea"
            Set item.Hint = NewControl(container, "Forms.Label.1")
            item.Hint.Caption = ""
            item.Hint.BackStyle = 0
            item.Hint.WordWrap = (it.Kind = "textarea")
            Set item.Ctl = NewControl(container, "Forms.TextBox.1")
            Set item.Txt = item.Ctl
            With item.Txt
                TrySet item.Txt, "BackStyle", 0
                .BorderStyle = 0
                TrySet item.Txt, "SpecialEffect", 0
                TrySet item.Txt, "SelectionMargin", False
                TrySet item.Txt, "IntegralHeight", False
                .MultiLine = (it.Kind = "textarea")
                .WordWrap = (it.Kind = "textarea")
                TrySet item.Txt, "EnterKeyBehavior", (it.Kind = "textarea")
                TrySet item.Txt, "TabKeyBehavior", False
                If it.Kind = "textarea" Then TrySet item.Txt, "ScrollBars", 2
            End With
        Case "select"
            Set item.Ctl = NewControl(container, "Forms.ComboBox.1")
            Set item.Cmb = item.Ctl
            With item.Cmb
                .Style = 2
                TrySet item.Cmb, "SpecialEffect", 0
                .BorderStyle = 0
                TrySet item.Cmb, "ShowDropButtonWhen", 0
                TrySet item.Cmb, "ListRows", 16
            End With
        Case "button"
            Set item.Ctl = NewControl(container, "Forms.CommandButton.1")
            Set item.Btn = item.Ctl
            TrySet item.Btn, "TakeFocusOnClick", True
            item.Btn.WordWrap = False
        Case "check"
            If it.IsRadio Then
                Set item.Ctl = NewControl(container, "Forms.OptionButton.1")
                Set item.Opt = item.Ctl
                item.Opt.GroupName = "g" & mCounter
            Else
                Set item.Ctl = NewControl(container, "Forms.CheckBox.1")
                Set item.Chk = item.Ctl
            End If
            item.Ctl.Caption = ""
            item.Ctl.BackStyle = 0
            TrySet item.Ctl, "SpecialEffect", 0
        Case Else
            Set item.Ctl = NewControl(container, "Forms.Label.1")
            Set item.Lbl = item.Ctl
            With item.Lbl
                .Caption = ""
                .BackStyle = 0
                .WordWrap = False
                .AutoSize = False
            End With
    End Select
    Set mCtlItems(item.Ctl.Name) = item
    Set CreateItem = item
    Exit Function
Failed:
    Err.Raise Err.Number, "modRender.CreateItem", Err.Description & " [" & it.Kind & " " & it.Id & " in " & it.Region & "; frame parent " & CStr(mFrameParent(it.Region)) & "]"
End Function

Private Sub UnmapItem(ByVal item As clsRenderItem)
    Dim v As Variant
    If Not item.Spec Is Nothing Then
        If Not item.Spec.Node Is Nothing Then
            If Obj(mNodeItems, CStr(item.Spec.Node.Uid)) Is item Then mNodeItems.Remove CStr(item.Spec.Node.Uid)
        End If
    End If
End Sub

Private Sub DeleteItem(ByVal item As clsRenderItem)
    On Error Resume Next
    If Not IsEmpty(mItems(item.Id)) Then
        If mItems(item.Id) Is item Then mItems.Remove item.Id
    End If
    UnmapItem item
    If Not item.Ctl Is Nothing Then
        ParkFocusIfInside item.Ctl
        If Not IsEmpty(mCtlItems(item.Ctl.Name)) Then mCtlItems.Remove item.Ctl.Name
        item.Container.Remove item.Ctl.Name
    End If
    If Not item.Hint Is Nothing Then item.Container.Remove item.Hint.Name
    item.Dispose
End Sub

' Raises the region's items painted at or after paint index fromIndex, in paint order.
Private Sub Restack(ByVal r As clsRenderSpec, ByVal fromIndex As Long)
    Dim it As clsRenderSpec, item As clsRenderItem
    On Error Resume Next
    For Each it In r.Items
        If it.PaintIndex >= fromIndex Then
            Set item = Obj(mItems, it.Id)
            If Not item Is Nothing Then
                WillWrite
                If Not item.Hint Is Nothing Then item.Hint.ZOrder 0
                item.Ctl.ZOrder 0
            End If
        End If
    Next it
End Sub

Private Sub SetTabOrder(ByVal r As clsRenderSpec)
    Dim it As clsRenderSpec, item As clsRenderItem, n As Long
    On Error Resume Next
    For Each it In r.Items
        Select Case it.Kind
            Case "input", "textarea", "select", "button", "check"
                Set item = Obj(mItems, it.Id)
                If Not item Is Nothing Then
                    If item.Ctl.TabIndex <> n Then item.Ctl.TabIndex = n
                    n = n + 1
                End If
        End Select
    Next it
End Sub

' Frames follow their regions' order (title bar, library, page) so Tab walks the window
' top to bottom.
Private Sub SetFrameTabOrder(ByVal regions As Collection)
    Dim r As clsRenderSpec, fr As Object, n As Long
    On Error Resume Next
    For Each r In regions
        If r.Id <> "page" And Len(r.Region) = 0 Then
            Set fr = Obj(mFrames, r.Id)
            If Not fr Is Nothing Then
                fr.TabIndex = n
                n = n + 1
            End If
        End If
    Next r
End Sub

' Writes only what changed since the record last applied to the item's controls (fields are
' compared directly: building look strings for every item cost more than the comparison).
Private Sub UpdateItem(ByVal item As clsRenderItem, ByVal it As clsRenderSpec)
    Dim old As clsRenderSpec, lookSame As Boolean, geomSame As Boolean
    Set old = item.Spec
    If Not old Is Nothing Then
        If item.LookKey = "ok" Then lookSame = it.SameLook(old)
        If item.GeomKey = "ok" Then geomSame = it.SameGeometry(old)
    End If
    Set item.Spec = it
    If Not it.Node Is Nothing Then Set mNodeItems(CStr(it.Node.Uid)) = item
    If Not lookSame Then
        item.LookKey = ""
        ApplyLook item, it
        item.LookKey = "ok"
    End If
    If Not geomSame Then
        item.GeomKey = ""
        ApplyGeometry item, it
        item.GeomKey = "ok"
    End If
    SyncValue item, it
End Sub

Private Sub ApplyFont(ByVal ctl As Object, ByVal size As Double, ByVal weight As Long, ByVal italic As Boolean, ByVal symbol As Boolean)
    Dim name As String, bold As Boolean
    If size <= 0 Then size = 13
    If symbol Then
        name = mFontSymbol
    ElseIf weight >= 650 Then
        name = mFontRegular
        bold = True
    ElseIf weight >= 550 Then
        name = mFontSemibold
        bold = mSemiboldIsBold
    Else
        name = mFontRegular
    End If
    #If Mac Then
        With ctl.Font
            If .Name <> name Then .Name = name
            .Size = size * mPx
            .Bold = bold
            .Italic = italic
        End With
    #Else
        ' A new control's Font may be a shared object whose Bold/Italic do not change when
        ' set; assigning a font object of our own avoids that. Fonts are cached per style
        ' and never modified after assignment.
        Dim key As String, sf As StdFont
        key = name & "|" & Round(size * mPx, 2) & "|" & bold & "|" & italic
        If mFontObjects Is Nothing Then Set mFontObjects = NewMap()
        If mFontObjects.Exists(key) Then
            Set sf = mFontObjects(key)
        Else
            Set sf = New StdFont
            sf.Name = name
            sf.Size = size * mPx
            sf.Bold = bold
            sf.Italic = italic
            Set mFontObjects(key) = sf
        End If
        Set ctl.Font = sf
    #End If
End Sub

Private Sub ApplyLook(ByVal item As clsRenderItem, ByVal it As clsRenderSpec)
    Dim ctl As Object
    WillWrite
    Set ctl = item.Ctl
    Select Case it.Kind
        Case "text"
            ctl.Caption = it.Text
            ctl.ForeColor = it.Color
            ApplyFont ctl, it.Size, it.Weight, it.Italic, it.Symbol
            ctl.BackStyle = 0
            ctl.BorderStyle = 0
            ctl.TextAlign = 1
            TrySet ctl, "ControlTipText", it.Tip
        Case "rect"
            ctl.Caption = ""
            If it.Bg >= 0 Then ctl.BackStyle = 1: ctl.BackColor = it.Bg Else ctl.BackStyle = 0
            If it.Bc >= 0 Then ctl.BorderStyle = 1: ctl.BorderColor = it.Bc Else ctl.BorderStyle = 0
            TrySet ctl, "ControlTipText", it.Tip
        Case "hit"
            ctl.Caption = ""
            ctl.BackStyle = 0
            ctl.BorderStyle = 0
            If it.Resizer Then TrySet ctl, "MousePointer", 9 Else TrySet ctl, "MousePointer", 0
            TrySet ctl, "ControlTipText", it.Tip
        Case "input", "textarea"
            ApplyFont ctl, it.Size, it.Weight, False, False
            ctl.ForeColor = it.Color
            ' Disabled controls stay enabled (MSForms would emboss them, ignoring the colours)
            ' but refuse input and focus.
            ctl.Locked = it.ReadOnly Or it.Disabled
            If Not ctl.Enabled Then ctl.Enabled = True
            ctl.TabStop = Not it.Disabled
            TrySet ctl, "ControlTipText", it.Tip
            item.Hint.Caption = it.Placeholder
            item.Hint.ForeColor = MutedColour(it.Color, it.Bg)
            ApplyFont item.Hint, it.Size, it.Weight, False, False
        Case "select"
            ApplyFont ctl, it.Size, it.Weight, False, False
            #If Mac Then
                ' Mac lists may draw a system background whatever BackColor says: they keep
                ' the system text colour so dark-theme text never sits on a light list.
            #Else
                If it.Bg >= 0 Then ctl.BackColor = it.Bg
                ctl.ForeColor = it.Color
            #End If
            ctl.Locked = it.Disabled
            If Not ctl.Enabled Then ctl.Enabled = True
            ctl.TabStop = Not it.Disabled
            TrySet ctl, "ControlTipText", it.Tip
            If item.OptionsKey <> it.Options & "|" & it.OptionFlags Then
                LoadOptions item, it
                item.OptionsKey = it.Options & "|" & it.OptionFlags
            End If
        Case "button"
            ' A focus proxy: the button itself is drawn with labels.
            If Len(ctl.Caption) > 0 Then ctl.Caption = ""
            If Not ctl.Enabled Then ctl.Enabled = True
            ctl.TabStop = Not it.Disabled
            TrySet ctl, "ControlTipText", it.Tip
        Case "check"
            ctl.Locked = it.Disabled
            If Not ctl.Enabled Then ctl.Enabled = True
            ctl.TabStop = Not it.Disabled
            TrySet ctl, "ControlTipText", it.Tip
    End Select
End Sub

Private Function MutedColour(ByVal fg As Long, ByVal bg As Long) As Long
    If bg < 0 Then bg = RGB(255, 255, 255)
    MutedColour = mPaint.Blend(fg, 0.55, bg)
End Function

Private Sub LoadOptions(ByVal item As clsRenderItem, ByVal it As clsRenderSpec)
    Dim caps() As String, i As Long
    item.Cmb.Clear
    If Len(it.Options) = 0 Then Exit Sub
    caps = Split(it.Options, vbLf)
    For i = 0 To UBound(caps)
        If Mid$(it.OptionFlags, i + 1, 1) = "1" Then
            item.Cmb.AddItem caps(i) & " (not available)"
        Else
            item.Cmb.AddItem caps(i)
        End If
    Next i
End Sub

Private Sub ApplyGeometry(ByVal item As clsRenderItem, ByVal it As clsRenderSpec)
    Dim x As Double, y As Double, w As Double, h As Double, lh As Double
    WillWrite
    x = it.X: y = it.Y: w = it.W: h = it.H
    Select Case it.Kind
        Case "text"
            ' A little slack absorbs rounding so the last glyph is never clipped.
            item.Ctl.Move x * mPx, y * mPx, Pos((w + 4) * mPx), Pos((h + 2) * mPx)
        Case "check", "button"
            If it.Proxy Then
                ' Focus proxies take no area (and no clicks): the element is drawn.
                item.Ctl.Move x * mPx, y * mPx, 0, 0
            ElseIf it.Kind = "check" Then
                If w < 16 Then w = 16
                item.Ctl.Move (x - 1) * mPx, (y + (h - 16) / 2) * mPx, w * mPx, 16 * mPx
            Else
                item.Ctl.Move x * mPx, y * mPx, Pos(w * mPx), Pos(h * mPx)
            End If
        Case "input", "textarea"
            item.Ctl.Move x * mPx, y * mPx, Pos(w * mPx), Pos(h * mPx)
            If it.Kind = "textarea" Then
                item.Hint.Move (x + 3) * mPx, (y + 3) * mPx, Pos((w - 8) * mPx), Pos((h - 6) * mPx)
            Else
                lh = NaturalLineHeight(it.Size, it.Weight)
                item.Hint.Move (x + 3) * mPx, (y + (h - lh) / 2) * mPx, Pos((w - 6) * mPx), Pos(lh * mPx)
            End If
        Case Else
            item.Ctl.Move x * mPx, y * mPx, Pos(w * mPx), Pos(h * mPx)
    End Select
End Sub

' Control state the user can change: re-applied whenever it differs from the DOM.
Private Sub SyncValue(ByVal item As clsRenderItem, ByVal it As clsRenderSpec)
    Dim cur As String, want As String, hint As Boolean
    Select Case it.Kind
        Case "input", "textarea"
            cur = NormaliseText(item.Txt.Text)
            If cur <> it.Text Then
                WillWrite
                If it.Kind = "textarea" Then want = Replace(it.Text, vbLf, vbCrLf) Else want = it.Text
                item.Txt.Text = want
                item.Txt.SelStart = Len(want)
            End If
            hint = (Len(it.Text) = 0 And Len(it.Placeholder) > 0)
            If item.Hint.Visible <> hint Then WillWrite: item.Hint.Visible = hint
        Case "select"
            If item.Cmb.ListIndex <> it.SelIndex Then
                If it.SelIndex < item.Cmb.ListCount Then WillWrite: item.Cmb.ListIndex = it.SelIndex
            End If
        Case "check"
            If CBool(item.Ctl.Value) <> it.Checked Then WillWrite: item.Ctl.Value = it.Checked
    End Select
End Sub

Private Function NormaliseText(ByVal t As String) As String
    NormaliseText = Replace(Replace(t, vbCrLf, vbLf), vbCr, vbLf)
End Function

' ================================================================ sticky copies

' Shows each sticky element's pinned copy (an overlay frame) only while its scroller is
' scrolled past the element; pending = a scroll distance not yet applied (Scroll event).
Private Sub UpdateSticky(ByVal pending As Double)
    Dim keys As clsList, k As Variant, r As clsRenderSpec, fr As Object, host As Object, show As Boolean, found As Boolean, sinkHost As Object
    On Error GoTo Done
    If mFrameSpecs Is Nothing Then Exit Sub
    Set keys = mFrameSpecs.Keys
    For Each k In keys
        Set r = mFrameSpecs(CStr(k))
        If Len(r.StickyOf) > 0 Then
            Set fr = Obj(mFrames, CStr(k))
            Set host = Obj(mFrames, r.StickyOf)
            If Not fr Is Nothing And Not host Is Nothing Then
                found = True
                Set sinkHost = host
                show = ((host.ScrollTop + pending) / mPx > r.StickyY + 0.5)
                If fr.Visible <> show Then fr.Visible = show
            End If
        End If
    Next k
    ' Scrollbar drags scroll the frame without our code: its Scroll event updates the copies.
    If found Then
        If mScrollSink Is Nothing Then Set mScrollSink = New clsRenderItem
        If Not mScrollSink.Frm Is sinkHost Then Set mScrollSink.Frm = sinkHost
    ElseIf Not mScrollSink Is Nothing Then
        Set mScrollSink.Frm = Nothing
    End If
Done:
End Sub

' A scroller frame scrolled (scrollbar, keys, wheel). The event comes before ScrollTop
' changes (Windows): the copies follow ScrollTop + dy now and ScrollTop once Excel is idle.
Public Sub OnFrameScroll(ByVal item As clsRenderItem, ByVal dy As Single)
    On Error Resume Next
    LastScrollEventTop = item.Frm.ScrollTop
    LastScrollEventDy = dy
    If mApplying Or mRendering Then Exit Sub
    UpdateSticky dy
    If Not mStickyTickDue Then
        mStickyTickDue = True
        Application.OnTime Now, Qualified("modRender.StickyTick")
    End If
End Sub

Public Sub StickyTick()
    On Error Resume Next
    mStickyTickDue = False
    If mForm Is Nothing Or Not mShown Then Exit Sub
    UpdateSticky 0
End Sub

' ================================================================ user actions

Private Function ItemNode(ByVal item As clsRenderItem) As clsNode
    If item.Spec Is Nothing Then Exit Function
    Set ItemNode = item.Spec.Node
End Function

Private Function BeginAction() As Boolean
    If mApplying Or mBusy Or mForm Is Nothing Then Exit Function
    mBusy = True
    mActionStart = ClockMs()
    BeginAction = True
End Function

' Runs the DOM's queued work, re-renders when the DOM changed and arms the timer.
Private Sub AfterAction()
    Dim t1 As Double
    On Error GoTo Failed
    t1 = ClockMs()
    LastHandlerMs = t1 - mActionStart
    modDom.Settle 700
    LastSettleMs = ClockMs() - t1
    If modDom.Version <> mRenderedVersion Or FocusRestyleDue() Then
        RenderNow
    Else
        LastRenderMs = 0: LastLayoutMs = 0: LastPaintMs = 0: LastApplyMs = 0
        ReflectFocus
    End If
    ScheduleTimer
    LastActionMs = ClockMs() - mActionStart
    Exit Sub
Failed:
    ReportError "modRender.AfterAction"
End Sub

Public Sub OnItemClick(ByVal item As clsRenderItem)
    Dim began As Boolean
    On Error GoTo Failed
    Dim node As clsNode
    If Not BeginAction() Then Exit Sub
    began = True
    On Error GoTo Failed
    FocusProxyFor item
    SyncFocus
    Select Case item.Kind
        Case "button", "check"
            Set node = ItemNode(item)
            If Not node Is Nothing Then
                If Not node.Disabled Then node.Click
            End If
        Case Else
            If Not item.Spec Is Nothing Then
                If Not item.Spec.Resizer And Not item.Spec.ClickNode Is Nothing Then
                    If item.Spec.ClickNode.LocalName = "select" Then
                        ' A select's chevron opens its list.
                        DropDownFor item.Spec.ClickNode
                    ElseIf Not IsDisabledControl(item.Spec.ClickNode) Then
                        item.Spec.ClickNode.Click
                    End If
                End If
            End If
    End Select
    AfterAction
    mBusy = False
    Exit Sub
Failed:
    If began Then mBusy = False
    ReportError "modRender.OnItemClick"
End Sub

' A click on a drawn button or switch (its labels) focuses it, as a browser does on
' mousedown: its focus proxy takes the window's focus. Mouse focus shows no outline.
Private Sub FocusProxyFor(ByVal item As clsRenderItem)
    Dim target As clsRenderItem, node As clsNode
    If item.Spec Is Nothing Then Exit Sub
    If item.Spec.Proxy Then Exit Sub          ' Space on the proxy: a key press
    Set node = item.Spec.ClickNode
    If node Is Nothing Then Exit Sub
    If item.Spec.Resizer Then Exit Sub
    mFocusVisible = False
    Set target = Obj(mNodeItems, CStr(node.Uid))
    If target Is Nothing Then Exit Sub
    If target.Spec Is Nothing Then Exit Sub
    If Not target.Spec.Proxy Or target.Spec.Disabled Then Exit Sub
    On Error Resume Next
    If Not ActiveLeaf() Is target.Ctl Then target.Ctl.SetFocus
End Sub

Private Function IsDisabledControl(ByVal node As clsNode) As Boolean
    Select Case node.LocalName
        Case "button", "input", "select", "textarea": IsDisabledControl = node.Disabled
    End Select
End Function

Private Sub DropDownFor(ByVal node As clsNode)
    Dim target As clsRenderItem
    If node.Disabled Then Exit Sub
    Set target = Obj(mNodeItems, CStr(node.Uid))
    If target Is Nothing Then Exit Sub
    If target.Cmb Is Nothing Then Exit Sub
    On Error Resume Next
    target.Cmb.SetFocus
    target.Cmb.DropDown
End Sub

Public Sub OnItemDblClick(ByVal item As clsRenderItem)
    Dim began As Boolean
    On Error GoTo Failed
    If Not BeginAction() Then Exit Sub
    began = True
    On Error GoTo Failed
    If Not item.Spec Is Nothing Then
        If item.Spec.Resizer And Not item.Spec.ClickNode Is Nothing Then
            item.Spec.ClickNode.DispatchEvent modDom.NewMouseEvent("dblclick")
        ElseIf Not item.Spec.ClickNode Is Nothing Then
            item.Spec.ClickNode.Click
        End If
    End If
    AfterAction
    mBusy = False
    Exit Sub
Failed:
    If began Then mBusy = False
    ReportError "modRender.OnItemDblClick"
End Sub

' Mouse on labels: drives the sidebar resizer (phase 0 down, 1 move, 2 up).
Public Sub OnItemMouse(ByVal item As clsRenderItem, ByVal phase As Long, ByVal button As Integer, ByVal x As Single, ByVal y As Single)
    Dim began As Boolean
    On Error GoTo Failed
    Dim node As clsNode, clientX As Double, evType As String, ax As Double, ay As Double
    If item.Spec Is Nothing Then Exit Sub
    If Not item.Spec.Resizer Then Exit Sub
    If phase = 1 And Not mDragging Then Exit Sub
    If phase = 0 And button <> 1 Then Exit Sub
    AbsoluteOrigin item.Ctl, ax, ay
    clientX = (ax + x) / mPx
    If phase = 1 And Abs(clientX - mDragX) < 3 Then Exit Sub
    If Not BeginAction() Then Exit Sub
    began = True
    On Error GoTo Failed
    mDragX = clientX
    Set node = item.Spec.ClickNode
    Select Case phase
        Case 0: evType = "pointerdown": mDragging = True
        Case 1: evType = "pointermove"
        Case Else: evType = "pointerup": mDragging = False
    End Select
    If Not node Is Nothing Then node.DispatchEvent modDom.NewMouseEvent(evType, clientX)
    AfterAction
    mBusy = False
    Exit Sub
Failed:
    If began Then mBusy = False
    ReportError "modRender.OnItemMouse"
End Sub

Public Sub OnItemText(ByVal item As clsRenderItem)
    Dim began As Boolean
    On Error GoTo Failed
    Dim node As clsNode, text As String
    If item.Txt Is Nothing Then Exit Sub
    If Not item.Hint Is Nothing And Not item.Spec Is Nothing Then item.Hint.Visible = (Len(item.Txt.Text) = 0 And Len(item.Spec.Placeholder) > 0)
    If Not BeginAction() Then Exit Sub
    began = True
    On Error GoTo Failed
    SyncFocus
    Set node = ItemNode(item)
    text = NormaliseText(item.Txt.Text)
    If Not node Is Nothing Then
        If node.Value <> text Then
            node.Value = text
            node.DispatchEvent modDom.NewEvent("input", True)
        End If
    End If
    AfterAction
    mBusy = False
    Exit Sub
Failed:
    If began Then mBusy = False
    ReportError "modRender.OnItemText"
End Sub

Public Sub OnItemSelect(ByVal item As clsRenderItem)
    Dim began As Boolean
    On Error GoTo Failed
    Dim node As clsNode, idx As Long
    If item.Cmb Is Nothing Then Exit Sub
    If Not BeginAction() Then Exit Sub
    began = True
    On Error GoTo Failed
    SyncFocus
    Set node = ItemNode(item)
    idx = item.Cmb.ListIndex
    If Not node Is Nothing Then
        If idx >= 0 And Mid$(item.Spec.OptionFlags, idx + 1, 1) = "1" Then
            mApplying = True
            If item.Spec.SelIndex < item.Cmb.ListCount Then item.Cmb.ListIndex = item.Spec.SelIndex
            mApplying = False
        ElseIf idx <> node.SelectedIndex Then
            node.SelectedIndex = idx
            node.DispatchEvent modDom.NewEvent("input", True)
            node.DispatchEvent modDom.NewEvent("change", True)
        End If
    End If
    AfterAction
    mBusy = False
    Exit Sub
Failed:
    mApplying = False
    If began Then mBusy = False
    ReportError "modRender.OnItemSelect"
End Sub

' Keys the website listens for go to the focused element as keydown events. Returns True
' when the key should be cancelled (the handler called preventDefault or moved focus).
Public Function OnItemKeyDown(ByVal item As clsRenderItem, ByVal keyCode As Integer, ByVal shift As Integer) As Boolean
    Dim began As Boolean
    On Error GoTo Failed
    Dim name As String, textual As Boolean, target As clsNode, evt As clsEvent, before As clsNode, cancel As Boolean
    textual = (item.Kind = "input" Or item.Kind = "textarea" Or item.Kind = "select")
    mFocusVisible = True
    name = KeyName(keyCode, shift, textual)
    If Len(name) = 0 Then Exit Function
    If Not BeginAction() Then Exit Function
    began = True
    On Error GoTo Failed
    SyncFocus
    Set target = mFocusNode
    If target Is Nothing Then Set target = modDom.Body
    Set before = modDom.ActiveElement()
    Set evt = modDom.NewKeyEvent("keydown", name, (shift And 1) = 1, (shift And 2) = 2, (shift And 4) = 4)
    target.DispatchEvent evt
    cancel = evt.DefaultPrevented
    If Not modDom.ActiveElement() Is before Then cancel = True
    ' Up/Down in a number field steps it, as a browser does (min, max and step honoured).
    If Not cancel And (keyCode = 38 Or keyCode = 40) And item.Kind = "input" Then
        If target.LocalName = "input" And LCase$(target.TypeName) = "number" And Not target.ReadOnly And Not target.Disabled Then
            If keyCode = 38 Then modApp3.StepNumberInput target, 1 Else modApp3.StepNumberInput target, -1
            If target Is mFocusNode Then mFocusValue = target.Value
            cancel = True
        End If
    End If
    ' Page Up/Down (and Home/End outside text) scroll the page when the website did not
    ' use the key (keyboard scrolling where a frame ignores the wheel or trackpad).
    If Not cancel And (((keyCode = 33 Or keyCode = 34) And item.Kind <> "textarea" And item.Kind <> "select") Or ((keyCode = 35 Or keyCode = 36) And Not textual)) Then
        If KeyScroll(item.RegionId, keyCode) Then cancel = True
    End If
    ' Enter on a focused button presses it, as in a browser.
    If Not cancel And item.Kind = "button" And keyCode = 13 Then
        If Not ItemNode(item) Is Nothing Then If Not ItemNode(item).Disabled Then ItemNode(item).Click
        cancel = True
    End If
    AfterAction
    mBusy = False
    OnItemKeyDown = cancel
    Exit Function
Failed:
    If began Then mBusy = False
    ReportError "modRender.OnItemKeyDown"
End Function

Private Function KeyName(ByVal keyCode As Integer, ByVal shift As Integer, ByVal textual As Boolean) As String
    Select Case keyCode
        Case 9: KeyName = "Tab"
        Case 13: KeyName = "Enter"
        Case 27: KeyName = "Escape"
        Case 38: KeyName = "ArrowUp"
        Case 40: KeyName = "ArrowDown"
        Case 37: If Not textual Then KeyName = "ArrowLeft"
        Case 39: If Not textual Then KeyName = "ArrowRight"
        Case 36: If Not textual Then KeyName = "Home"
        Case 35: If Not textual Then KeyName = "End"
        Case 33: KeyName = "PageUp"
        Case 34: KeyName = "PageDown"
        Case 65 To 90
            ' Not with Ctrl, Alt or (where MSForms reports it) Cmd: Ctrl/Cmd+C on a button
            ' must not start a search.
            If Not textual And (shift And 14) = 0 Then
                If (shift And 1) = 1 Then KeyName = Chr$(keyCode) Else KeyName = LCase$(Chr$(keyCode))
            End If
    End Select
End Function

' Scrolls the region (or the nearest scrolling region around it) for Page Up/Down,
' Home/End. True when a frame scrolled.
Private Function KeyScroll(ByVal regionId As String, ByVal keyCode As Integer) As Boolean
    Dim fr As Object, id As String, target As Double, maxTop As Double, guard As Long
    On Error GoTo Done
    id = regionId
    Do While Len(id) > 0 And guard < 20
        Set fr = Obj(mFrames, id)
        If Not fr Is Nothing Then
            If (fr.ScrollBars And 2) <> 0 Then Exit Do
        End If
        Set fr = Nothing
        If IsEmpty(mFrameParent(id)) Then Exit Do
        id = CStr(mFrameParent(id))
        guard = guard + 1
    Loop
    If fr Is Nothing Then Set fr = Obj(mFrames, "main")
    If fr Is Nothing Then Exit Function
    If (fr.ScrollBars And 2) = 0 Then Exit Function
    maxTop = fr.ScrollHeight - fr.InsideHeight
    If maxTop < 0 Then maxTop = 0
    Select Case keyCode
        Case 33: target = fr.ScrollTop - fr.InsideHeight * 0.875
        Case 34: target = fr.ScrollTop + fr.InsideHeight * 0.875
        Case 36: target = 0
        Case 35: target = maxTop
    End Select
    If target > maxTop Then target = maxTop
    If target < 0 Then target = 0
    fr.ScrollTop = target
    UpdateSticky 0
    KeyScroll = True
Done:
End Function

' A mouse press in a text box or list: focus that shows no outline.
Public Sub OnItemMouseFocus(ByVal item As clsRenderItem)
    mFocusVisible = False
    OnItemFocusCheck item
End Sub

Public Sub OnItemKeyUp(ByVal item As clsRenderItem)
    mFocusVisible = True
    OnItemFocusCheck item
End Sub

' Focus may have moved (Tab, a click into a field): report it to the DOM.
Public Sub OnItemFocusCheck(ByVal item As clsRenderItem)
    Dim began As Boolean
    On Error GoTo Failed
    If Not BeginAction() Then Exit Sub
    began = True
    On Error GoTo Failed
    If SyncFocus() Then AfterAction
    mBusy = False
    Exit Sub
Failed:
    If began Then mBusy = False
    ReportError "modRender.OnItemFocusCheck"
End Sub

' Keys pressed while no control has focus (the form itself).
Public Function OnFormKeyDown(ByVal keyCode As Integer, ByVal shift As Integer) As Boolean
    Dim began As Boolean
    On Error GoTo Failed
    Dim name As String, evt As clsEvent
    name = KeyName(keyCode, shift, False)
    If Len(name) = 0 Then Exit Function
    If Not BeginAction() Then Exit Function
    began = True
    On Error GoTo Failed
    Set evt = modDom.NewKeyEvent("keydown", name, (shift And 1) = 1, (shift And 2) = 2, (shift And 4) = 4)
    modDom.ActiveElement().DispatchEvent evt
    OnFormKeyDown = evt.DefaultPrevented
    If Not OnFormKeyDown And keyCode >= 33 And keyCode <= 36 Then OnFormKeyDown = KeyScroll("main", keyCode)
    AfterAction
    mBusy = False
    Exit Function
Failed:
    If began Then mBusy = False
    ReportError "modRender.OnFormKeyDown"
End Function

' ================================================================ focus

Private Function ActiveLeaf() As Object
    Dim c As Object
    On Error GoTo Done
    Set c = mForm.ActiveControl
    Do While Not c Is Nothing
        If TypeName(c) <> "Frame" Then Exit Do
        Set c = c.ActiveControl
    Loop
    Set ActiveLeaf = c
Done:
End Function

' Compares the focused control with the DOM's focus; on a real change dispatches change
' (edited text fields), blur/focusout and focus/focusin. Returns True when focus moved.
Private Function SyncFocus() As Boolean
    Dim c As Object, v As Variant, item As clsRenderItem, node As clsNode
    Set c = ActiveLeaf()
    If c Is Nothing Then Exit Function
    Set item = Obj(mCtlItems, c.Name)
    If item Is Nothing Then Exit Function
    Set node = ItemNode(item)
    If node Is Nothing Then Exit Function
    If node Is mFocusNode Then Exit Function
    ChangeFocus node
    SyncFocus = True
End Function

Private Sub ChangeFocus(ByVal node As clsNode)
    Dim old As clsNode, tag As String
    Set old = mFocusNode
    If Not old Is Nothing Then
        If old.IsConnected Then
            tag = old.LocalName
            If (tag = "input" Or tag = "textarea") And old.Value <> mFocusValue Then old.DispatchEvent modDom.NewEvent("change", True)
            old.DispatchEvent modDom.NewEvent("blur")
            old.DispatchEvent modDom.NewEvent("focusout", True)
        End If
    End If
    Set mFocusNode = node
    mFocusValue = node.Value
    node.Focus
    node.DispatchEvent modDom.NewEvent("focus")
    node.DispatchEvent modDom.NewEvent("focusin", True)
End Sub

' element.scrollIntoView({ block: "start" }) requests from the shell.
Private Sub ApplyScrollRequest()
    Dim fr As Object, target As Double
    On Error GoTo Done
    If modApp.ScrollRequestTarget Is Nothing Then Exit Sub
    Set modApp.ScrollRequestTarget = Nothing
    If Len(mPaint.ScrollTargetRegion) = 0 Then Exit Sub
    Set fr = Obj(mFrames, mPaint.ScrollTargetRegion)
    If fr Is Nothing Then Exit Sub
    If (fr.ScrollBars And 2) = 0 Then Exit Sub
    target = (mPaint.ScrollTargetY - 12) * mPx
    If target > fr.ScrollHeight - fr.InsideHeight Then target = fr.ScrollHeight - fr.InsideHeight
    If target < 0 Then target = 0
    fr.ScrollTop = target
Done:
End Sub

' Moves the window's focus to the DOM's focused element when the website moved it.
Private Sub ReflectFocus()
    Dim active As clsNode, v As Variant, item As clsRenderItem
    On Error GoTo Done
    Set active = modDom.ActiveElement()
    If active Is Nothing Then Exit Sub
    If active Is mFocusNode Then Exit Sub
    Set item = Obj(mNodeItems, CStr(active.Uid))
    If item Is Nothing Then Exit Sub
    Select Case item.Kind
        Case "input", "textarea", "select", "button", "check"
        Case Else: Exit Sub
    End Select
    If item.Spec Is Nothing Then Exit Sub
    If item.Spec.Disabled Then Exit Sub
    Set mFocusNode = active
    mFocusValue = active.Value
    EnsureVisible item
    mApplying = True
    item.Ctl.SetFocus
    If Not item.Txt Is Nothing Then item.Txt.SelStart = Len(item.Txt.Text)
    mApplying = False
    Exit Sub
Done:
    mApplying = False
End Sub

Private Sub EnsureVisible(ByVal item As clsRenderItem)
    Dim fr As Object, v As Variant, tp As Double, bottom As Double, target As Double
    Set fr = Obj(mFrames, item.RegionId)
    If fr Is Nothing Then Exit Sub
    If (fr.ScrollBars And 2) = 0 Then Exit Sub
    tp = item.Ctl.Top
    bottom = tp + item.Ctl.Height
    If Not item.Spec Is Nothing Then
        If item.Spec.Proxy Then bottom = tp + item.Spec.H * mPx
    End If
    If tp < fr.ScrollTop Then
        target = tp - 24
        If target < 0 Then target = 0
        fr.ScrollTop = target
    ElseIf bottom > fr.ScrollTop + fr.InsideHeight Then
        target = bottom - fr.InsideHeight + 24
        If target > fr.ScrollHeight - fr.InsideHeight Then target = fr.ScrollHeight - fr.InsideHeight
        If target < 0 Then target = 0
        fr.ScrollTop = target
    End If
End Sub

' Position of a control's top-left corner in form client points.
Private Sub AbsoluteOrigin(ByVal ctl As Object, ByRef x As Double, ByRef y As Double)
    Dim node As Object
    x = ctl.Left: y = ctl.Top
    On Error GoTo Done
    Set node = ctl.Parent
    Do While TypeName(node) = "Frame"
        x = x + node.Left - node.ScrollLeft
        y = y + node.Top - node.ScrollTop
        Set node = node.Parent
    Loop
Done:
End Sub

' ================================================================ timers and window events

Private Sub ScheduleTimer()
    Dim d As Long, secs As Long, due As Date
    On Error GoTo Failed
    If mForm Is Nothing Then Exit Sub
    d = modDom.NextTimerDelay()
    If d < 0 Then Exit Sub
    secs = (d + 999) \ 1000
    If secs < 1 Then secs = 1
    due = Now + TimeSerial(0, 0, secs)
    If mTimerSet Then
        If mTimerAt <= due Then Exit Sub
        CancelTimer
    End If
    Application.OnTime due, Qualified("modRender.TimerTick")
    mTimerAt = due
    mTimerSet = True
    Exit Sub
Failed:
    ReportError "modRender.ScheduleTimer"
End Sub

Private Sub CancelTimer()
    If Not mTimerSet Then Exit Sub
    On Error Resume Next
    Application.OnTime mTimerAt, Qualified("modRender.TimerTick"), , False
    mTimerSet = False
End Sub

' Application.OnTime: runs the DOM's next due timer(s) and redraws.
Public Sub TimerTick()
    On Error GoTo Failed
    mTimerSet = False
    If mForm Is Nothing Or Not mShown Then Exit Sub
    If mBusy Or mApplying Then
        ScheduleTimer
        Exit Sub
    End If
    mBusy = True
    modDom.AdvanceClock
    modDom.Settle 700
    If modDom.Version <> mRenderedVersion Then RenderNow
    mBusy = False
    ScheduleTimer
    Exit Sub
Failed:
    mBusy = False
    ReportError "modRender.TimerTick"
End Sub

' frmApp UserForm_Resize.
Public Sub FormResized()
    On Error GoTo Failed
    If mForm Is Nothing Or Not mShown Then Exit Sub
    If mResizePending Then Exit Sub
    mResizePending = True
    Application.OnTime Now, Qualified("modRender.ResizeTick")
    Exit Sub
Failed:
    ReportError "modRender.FormResized"
End Sub

Public Sub ResizeTick()
    On Error GoTo Failed
    mResizePending = False
    If mForm Is Nothing Or Not mShown Then Exit Sub
    If mBusy Then
        FormResized
        Exit Sub
    End If
    mBusy = True
    UpdateViewport False
    RenderNow
    mBusy = False
    Exit Sub
Failed:
    mBusy = False
    ReportError "modRender.ResizeTick"
End Sub

' frmApp UserForm_QueryClose: the window closes; the document stays loaded.
Public Sub FormClosing()
    On Error Resume Next
    CancelTimer
    WindowClosing
    mShown = False
    Set mForm = Nothing
    ResetControls
    RemoveFonts
    mFontsReady = False
    Set mMeasurers = Nothing
End Sub

' ================================================================ Windows window chrome and wheel

Private Sub WindowCreated()
    #If Mac Then
    #Else
        Dim style As LongPtr, dc As LongPtr, unique As String
        On Error Resume Next
        unique = APP_TITLE & " " & ObjPtr(mForm)
        mForm.Caption = unique
        mHwnd = FindWindow(StrPtr("ThunderDFrame"), StrPtr(unique))
        mForm.Caption = APP_TITLE
        If mHwnd <> 0 Then
            style = GetWindowLongPtr(mHwnd, GWL_STYLE)
            SetWindowLongPtr mHwnd, GWL_STYLE, style Or WS_THICKFRAME Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX
            DrawMenuBar mHwnd
        End If
        dc = GetDC(0)
        mDpi = GetDeviceCaps(dc, 88)
        ReleaseDC 0, dc
        If mDpi <= 0 Then mDpi = 96
        If mHook = 0 Then mHook = SetWindowsHookEx(WH_MOUSE, AddressOf WheelHookProc, 0, GetCurrentThreadId())
    #End If
End Sub

Private Sub WindowClosing()
    #If Mac Then
    #Else
        If mHook <> 0 Then UnhookWindowsHookEx mHook
        mHook = 0
        mHwnd = 0
    #End If
End Sub

#If Mac Then
#Else
' Diagnostics: the app window's child windows (class, visibility), indented by depth.
Public Function DescribeChildWindows() As String
    DescribeChildWindows = ChildWindowLines(mHwnd, 0)
End Function

Private Function ChildWindowLines(ByVal parent As LongPtr, ByVal depth As Long) As String
    Dim h As LongPtr, buf As String, n As Long, out As String
    If parent = 0 Then Exit Function
    h = GetWindow(parent, GW_CHILD)
    Do While h <> 0
        buf = String$(128, vbNullChar)
        n = GetClassName(h, StrPtr(buf), 128)
        out = out & String$(depth * 2, " ") & Left$(buf, n) & " visible " & IsWindowVisible(h) & vbLf
        out = out & ChildWindowLines(h, depth + 1)
        h = GetWindow(h, GW_HWNDNEXT)
    Loop
    ChildWindowLines = out
End Function

' Window handle of the app window (0 when closed); tests capture it for screenshots.
Public Function AppWindowHandle() As LongPtr
    AppWindowHandle = mHwnd
End Function

' MSForms frames ignore the mouse wheel: a mouse hook on Excel's thread turns wheel
' messages aimed at the window into region scrolling.
Private Function WheelHookProc(ByVal code As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Dim x As Long, y As Long, target As LongPtr, data As Long
    On Error Resume Next
    If code >= 0 And wParam = WM_MOUSEWHEEL And mHwnd <> 0 And Not mForm Is Nothing Then
        CopyMemory VarPtr(x), lParam, 4
        CopyMemory VarPtr(y), lParam + 4, 4
        CopyMemory VarPtr(target), lParam + HOOK_OFFSET_HWND, LenB(target)
        CopyMemory VarPtr(data), lParam + HOOK_OFFSET_DATA, 4
        If GetAncestor(target, GA_ROOT) = mHwnd Then
            If OnWheel(x, y, data \ 65536) Then
                WheelHookProc = 1
                Exit Function
            End If
        End If
    End If
    WheelHookProc = CallNextHookEx(mHook, code, wParam, lParam)
End Function

' Scrolls the innermost scrollable region under the pointer.
Private Function OnWheel(ByVal screenX As Long, ByVal screenY As Long, ByVal delta As Long) As Boolean
    Dim p As POINTAPI, x As Double, y As Double, k As Variant, fr As Object, best As Object, fx As Double, fy As Double
    Dim bestDepth As Long, depth As Long, node As Object, target As Double
    p.x = screenX: p.y = screenY
    ScreenToClient mHwnd, p
    x = p.x * 72 / mDpi
    y = p.y * 72 / mDpi
    bestDepth = -1
    Dim keys11 As clsList: Set keys11 = mFrames.Keys
    For Each k In keys11
        Set fr = mFrames(CStr(k))
        If (fr.ScrollBars And 2) <> 0 And fr.Visible Then
            AbsoluteOrigin fr, fx, fy
            If x >= fx And x < fx + fr.Width And y >= fy And y < fy + fr.Height Then
                depth = 0
                Set node = fr.Parent
                Do While TypeName(node) = "Frame"
                    depth = depth + 1
                    Set node = node.Parent
                Loop
                If depth >= bestDepth Then Set best = fr: bestDepth = depth
            End If
        End If
    Next k
    If best Is Nothing Then Exit Function
    target = best.ScrollTop - delta / 120 * 54
    If target > best.ScrollHeight - best.InsideHeight Then target = best.ScrollHeight - best.InsideHeight
    If target < 0 Then target = 0
    best.ScrollTop = target
    UpdateSticky 0
    OnWheel = True
End Function
#End If
