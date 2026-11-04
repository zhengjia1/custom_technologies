VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' A DOM event (port of minidom's Event/MouseEvent/KeyboardEvent/InputEvent/CustomEvent).
' Create with modDom.NewEvent / NewKeyEvent / NewMouseEvent and dispatch with
' node.DispatchEvent. Handlers receive this object.

Public EventType As String
Public Bubbles As Boolean
Public Cancelable As Boolean
Public Key As String
Public Code As String
Public ShiftKey As Boolean
Public CtrlKey As Boolean
Public AltKey As Boolean
Public MetaKey As Boolean
Public IsComposing As Boolean
Public ClientX As Double
Public ClientY As Double
Public Button As Long
Public InputType As String
Public Detail As Variant

Private mDefaultPrevented As Boolean
Private mTarget As clsNode
Private mCurrent As clsNode
Private mPhase As Long
Private mStop As Boolean
Private mStopNow As Boolean

Public Property Get DefaultPrevented() As Boolean
    DefaultPrevented = mDefaultPrevented
End Property

Public Property Get Target() As clsNode
    Set Target = mTarget
End Property

Public Property Get CurrentTarget() As clsNode
    Set CurrentTarget = mCurrent
End Property

' 0 none, 1 capturing, 2 at target, 3 bubbling.
Public Property Get EventPhase() As Long
    EventPhase = mPhase
End Property

Public Sub PreventDefault()
    If Cancelable Then mDefaultPrevented = True
End Sub

Public Sub StopPropagation()
    mStop = True
End Sub

Public Sub StopImmediatePropagation()
    mStop = True
    mStopNow = True
End Sub

' ---- used by clsNode.DispatchEvent

Friend Sub BeginDispatch(ByVal target As clsNode)
    Set mTarget = target
    mStop = False
    mStopNow = False
End Sub

Friend Sub SetCurrent(ByVal node As clsNode)
    Set mCurrent = node
End Sub

Friend Sub SetPhase(ByVal phase As Long)
    mPhase = phase
End Sub

Friend Property Get Stopped() As Boolean
    Stopped = mStop
End Property

Friend Property Get StoppedNow() As Boolean
    StoppedNow = mStopNow
End Property
