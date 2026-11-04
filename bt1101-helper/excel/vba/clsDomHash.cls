VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDomHash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Case-sensitive string-keyed hash table (open addressing) used by the DOM for the
' id index and the parsed-selector cache. Values are Variants (objects allowed).
' There is no single-key removal: the DOM only ever clears the whole table.

Private mKeys() As String
Private mVals() As Variant
Private mUsed() As Boolean
Private mCap As Long
Private mCount As Long

Private Sub Class_Initialize()
    Allocate 64
End Sub

Private Sub Allocate(ByVal capacity As Long)
    mCap = capacity
    ReDim mKeys(0 To mCap - 1)
    ReDim mVals(0 To mCap - 1)
    ReDim mUsed(0 To mCap - 1)
    mCount = 0
End Sub

Public Property Get Count() As Long
    Count = mCount
End Property

Public Sub Clear()
    Allocate 64
End Sub

Private Function HashOf(ByRef key As String) As Long
    Dim b() As Byte, i As Long, h As Long, n As Long
    If LenB(key) = 0 Then HashOf = 0: Exit Function
    b = key
    n = UBound(b)
    h = 5381
    For i = 0 To n
        h = ((h * 31) + b(i)) And &HFFFFFF
    Next
    HashOf = h
End Function

Private Function Slot(ByRef key As String) As Long
    Dim i As Long
    i = HashOf(key) And (mCap - 1)
    Do While mUsed(i)
        If mKeys(i) = key Then Slot = i: Exit Function
        i = (i + 1) And (mCap - 1)
    Loop
    Slot = i
End Function

Public Function Exists(ByVal key As String) As Boolean
    Exists = mUsed(Slot(key))
End Function

' Returns the stored value (Empty when absent); found reports presence.
Public Function Lookup(ByVal key As String, ByRef found As Boolean) As Variant
    Dim i As Long
    i = Slot(key)
    found = mUsed(i)
    If found Then
        If IsObject(mVals(i)) Then Set Lookup = mVals(i) Else Lookup = mVals(i)
    End If
End Function

Public Sub Store(ByVal key As String, ByRef value As Variant)
    Dim i As Long
    i = Slot(key)
    If Not mUsed(i) Then
        If (mCount + 1) * 2 > mCap Then
            Grow
            i = Slot(key)
        End If
        mUsed(i) = True
        mKeys(i) = key
        mCount = mCount + 1
    End If
    If IsObject(value) Then Set mVals(i) = value Else mVals(i) = value
End Sub

Private Sub Grow()
    Dim oldKeys() As String, oldVals() As Variant, oldUsed() As Boolean, oldCap As Long, i As Long, j As Long
    oldKeys = mKeys: oldVals = mVals: oldUsed = mUsed: oldCap = mCap
    Allocate oldCap * 2
    For i = 0 To oldCap - 1
        If oldUsed(i) Then
            j = Slot(oldKeys(i))
            mUsed(j) = True
            mKeys(j) = oldKeys(i)
            If IsObject(oldVals(i)) Then Set mVals(j) = oldVals(i) Else mVals(j) = oldVals(i)
            mCount = mCount + 1
        End If
    Next
End Sub
