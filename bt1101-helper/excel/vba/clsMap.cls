VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Binary

' Insertion-ordered map from case-sensitive String keys to Variants (values or objects),
' like a JavaScript Map with string keys. Item(key) is the default member:
' m("a") = 1, Set m("b") = obj, x = m("a"). Reading a missing key returns Empty.
' Setting an existing key keeps its position. For Each enumerates the keys.
' Pass keys as Strings (a Double key would be converted with the locale-dependent CStr).
' Implementation: separate-chaining hash table over entry arrays kept in insertion order.

Private mSnaps As Collection
Private mKeys() As String
Private mVals() As Variant
Private mHash() As Long
Private mNext() As Long
Private mLive() As Boolean
Private mBucket() As Long
Private mUsed As Long
Private mCount As Long
Private mMask As Long

Private Sub Class_Initialize()
    Allocate 8
End Sub

Private Sub Allocate(ByVal cap As Long)
    Dim i As Long
    ReDim mKeys(0 To cap - 1)
    ReDim mVals(0 To cap - 1)
    ReDim mHash(0 To cap - 1)
    ReDim mNext(0 To cap - 1)
    ReDim mLive(0 To cap - 1)
    ReDim mBucket(0 To cap - 1)
    For i = 0 To cap - 1
        mBucket(i) = -1
    Next
    mMask = cap - 1
    mUsed = 0
    mCount = 0
End Sub

Private Function HashKey(ByRef key As String) As Long
    Dim b() As Byte, i As Long, h As Long
    h = 5381
    If LenB(key) > 0 Then
        b = key
        For i = 0 To UBound(b)
            h = (h And &HFFFFFF) * 31 + b(i)
        Next
    End If
    h = h And &HFFFFFF
    HashKey = h Xor (h \ 4096)
End Function

Private Function Find(ByRef key As String, ByVal h As Long) As Long
    Dim i As Long
    i = mBucket(h And mMask)
    Do While i >= 0
        If mHash(i) = h Then
            If mKeys(i) = key Then Find = i: Exit Function
        End If
        i = mNext(i)
    Loop
    Find = -1
End Function

' Rebuilds the table with room for at least `need` entries, dropping removed ones.
Private Sub Rebuild(ByVal need As Long)
    Dim oKeys() As String, oVals() As Variant, oHash() As Long, oLive() As Boolean
    Dim oUsed As Long, cap As Long, i As Long, j As Long
    oKeys = mKeys
    oVals = mVals
    oHash = mHash
    oLive = mLive
    oUsed = mUsed
    cap = 8
    Do While cap < need * 2
        cap = cap * 2
    Loop
    Allocate cap
    For i = 0 To oUsed - 1
        If oLive(i) Then
            j = mUsed
            mKeys(j) = oKeys(i)
            If IsObject(oVals(i)) Then Set mVals(j) = oVals(i) Else mVals(j) = oVals(i)
            mHash(j) = oHash(i)
            mLive(j) = True
            mNext(j) = mBucket(oHash(i) And mMask)
            mBucket(oHash(i) And mMask) = j
            mUsed = mUsed + 1
            mCount = mCount + 1
        End If
    Next
End Sub

Private Sub StoreValue(ByRef key As String, ByRef value As Variant)
    Dim h As Long, i As Long
    h = HashKey(key)
    i = Find(key, h)
    If i < 0 Then
        If mUsed > mMask Then
            Rebuild mCount + 1
        End If
        i = mUsed
        mKeys(i) = key
        mHash(i) = h
        mLive(i) = True
        mNext(i) = mBucket(h And mMask)
        mBucket(h And mMask) = i
        mUsed = mUsed + 1
        mCount = mCount + 1
    End If
    If IsObject(value) Then Set mVals(i) = value Else mVals(i) = value
End Sub

Public Property Get Item(ByVal key As String) As Variant
Attribute Item.VB_UserMemId = 0
    Dim i As Long
    i = Find(key, HashKey(key))
    If i < 0 Then Exit Property
    If IsObject(mVals(i)) Then Set Item = mVals(i) Else Item = mVals(i)
End Property

Public Property Let Item(ByVal key As String, ByVal value As Variant)
    StoreValue key, value
End Property

Public Property Set Item(ByVal key As String, ByVal value As Variant)
    StoreValue key, value
End Property

' map.set(key, value): adds or replaces (an existing key keeps its position).
Public Sub Add(ByVal key As String, ByVal value As Variant)
    StoreValue key, value
End Sub

Public Function Exists(ByVal key As String) As Boolean
    Exists = (Find(key, HashKey(key)) >= 0)
End Function

' map.delete(key): True when the key existed.
Public Function Remove(ByVal key As String) As Boolean
    Dim h As Long, i As Long, prev As Long
    h = HashKey(key)
    prev = -1
    i = mBucket(h And mMask)
    Do While i >= 0
        If mHash(i) = h Then
            If mKeys(i) = key Then Exit Do
        End If
        prev = i
        i = mNext(i)
    Loop
    If i < 0 Then Exit Function
    If prev < 0 Then mBucket(h And mMask) = mNext(i) Else mNext(prev) = mNext(i)
    mLive(i) = False
    mKeys(i) = vbNullString
    mVals(i) = Empty
    mCount = mCount - 1
    Remove = True
    If mUsed > 32 And mCount * 4 < mUsed Then Rebuild mCount
End Function

Public Property Get Count() As Long
    Count = mCount
End Property

Public Sub Clear()
    Allocate 8
End Sub

' Keys in insertion order.
Public Function Keys() As clsList
    Dim r As clsList, i As Long
    Set r = New clsList
    For i = 0 To mUsed - 1
        If mLive(i) Then r.Add mKeys(i)
    Next
    Set Keys = r
End Function

' Values in insertion order.
Public Function Values() As clsList
    Dim r As clsList, i As Long
    Set r = New clsList
    For i = 0 To mUsed - 1
        If mLive(i) Then r.Add mVals(i)
    Next
    Set Values = r
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Dim c As Collection, i As Long
    Set c = New Collection
    For i = 0 To mUsed - 1
        If mLive(i) Then c.Add mKeys(i)
    Next
    ' A Collection's enumerator does not keep the Collection alive, so the snapshot is
    ' held here; the last few are kept so that nested loops over this object work.
    If mSnaps Is Nothing Then Set mSnaps = New Collection
    If mSnaps.Count >= 4 Then mSnaps.Remove 1
    mSnaps.Add c
    Set NewEnum = c.[_NewEnum]
End Function
