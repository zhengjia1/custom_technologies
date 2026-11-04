VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' One regular-expression match (like a JavaScript exec() result). Indices are 0-based.
' Group(0) is the whole match. A group that did not participate reads as "" with
' Matched(n) = False (JavaScript undefined).
Option Explicit

Private mIndex As Long
Private mEnd As Long
Private mGroups As Long
Private mValues() As String
Private mStarts() As Long
Private mEnds() As Long
Private mInput As String
Private mOwner As clsMatch

' Called by clsRegex. Matches from one MatchAll share the first match's Input (owner).
Friend Sub Init(ByRef text As String, ByRef caps() As Long, ByVal groups As Long, ByVal owner As clsMatch)
    Dim g As Long
    mGroups = groups
    mIndex = caps(0)
    mEnd = caps(1)
    ReDim mValues(0 To groups)
    ReDim mStarts(0 To groups)
    ReDim mEnds(0 To groups)
    For g = 0 To groups
        If caps(2 * g) >= 0 And caps(2 * g + 1) >= 0 Then
            mStarts(g) = caps(2 * g)
            mEnds(g) = caps(2 * g + 1)
            mValues(g) = Mid$(text, mStarts(g) + 1, mEnds(g) - mStarts(g))
        Else
            mStarts(g) = -1
            mEnds(g) = -1
        End If
    Next g
    If owner Is Nothing Then
        mInput = text
    Else
        Set mOwner = owner
    End If
End Sub

' 0-based start of the match.
Public Property Get Index() As Long
    Index = mIndex
End Property

' 0-based end of the match (exclusive).
Public Property Get EndIndex() As Long
    EndIndex = mEnd
End Property

' The matched text (Group(0)).
Public Property Get Value() As String
    Value = mValues(0)
End Property

' Number of capturing groups in the pattern.
Public Property Get GroupCount() As Long
    GroupCount = mGroups
End Property

' Text of group n ("" when it did not participate or n is out of range).
Public Property Get Group(ByVal n As Long) As String
    If n >= 0 And n <= mGroups Then Group = mValues(n)
End Property

' False when group n did not participate (JavaScript undefined) or n is out of range.
Public Property Get Matched(ByVal n As Long) As Boolean
    If n >= 0 And n <= mGroups Then Matched = (mStarts(n) >= 0)
End Property

' 0-based start of group n, or -1 when it did not participate.
Public Property Get GroupIndex(ByVal n As Long) As Long
    GroupIndex = -1
    If n >= 0 And n <= mGroups Then GroupIndex = mStarts(n)
End Property

' 0-based end (exclusive) of group n, or -1 when it did not participate.
Public Property Get GroupEnd(ByVal n As Long) As Long
    GroupEnd = -1
    If n >= 0 And n <= mGroups Then GroupEnd = mEnds(n)
End Property

' The string that was searched (JavaScript match.input; Input is a VBA keyword).
Public Property Get InputText() As String
    If mOwner Is Nothing Then InputText = mInput Else InputText = mOwner.InputText
End Property
