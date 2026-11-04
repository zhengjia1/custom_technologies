VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Binary

' Growable 0-based list of Variants (values or objects), like a JavaScript array.
' Item(i) is the default member: lst(0), lst(0) = 5, Set lst(1) = obj.
' Reading an index outside 0..Count-1 returns Empty (JS undefined); writing past the end
' extends the list with Empty. For Each works through NewEnum (it snapshots the items
' into a Collection, so changes made during the loop are not seen).

Private mSnaps As Collection
Private mItems() As Variant
Private mCount As Long

Private Sub Class_Initialize()
    ReDim mItems(0 To 7)
End Sub

Private Sub EnsureCapacity(ByVal need As Long)
    Dim cap As Long
    cap = UBound(mItems) + 1
    If need <= cap Then Exit Sub
    Do While cap < need
        cap = cap * 2
    Loop
    ReDim Preserve mItems(0 To cap - 1)
End Sub

Private Sub PutAt(ByVal i As Long, ByRef value As Variant)
    If IsObject(value) Then Set mItems(i) = value Else mItems(i) = value
End Sub

Private Sub MoveSlot(ByVal dst As Long, ByVal src As Long)
    If IsObject(mItems(src)) Then Set mItems(dst) = mItems(src) Else mItems(dst) = mItems(src)
End Sub

Public Property Get Item(ByVal index As Long) As Variant
Attribute Item.VB_UserMemId = 0
    If index < 0 Or index >= mCount Then Exit Property
    If IsObject(mItems(index)) Then Set Item = mItems(index) Else Item = mItems(index)
End Property

Public Property Let Item(ByVal index As Long, ByVal value As Variant)
    StoreAt index, value
End Property

Public Property Set Item(ByVal index As Long, ByVal value As Variant)
    StoreAt index, value
End Property

Private Sub StoreAt(ByVal index As Long, ByRef value As Variant)
    If index < 0 Then Err.Raise 9, "clsList", "Negative list index"
    If index >= mCount Then
        EnsureCapacity index + 1
        mCount = index + 1
    End If
    PutAt index, value
End Sub

Public Property Get Count() As Long
    Count = mCount
End Property

Public Property Get Length() As Long
    Length = mCount
End Property

' arr.at(i): negative counts from the end; Empty when out of range.
Public Function At(ByVal index As Long) As Variant
    If index < 0 Then index = index + mCount
    If index < 0 Or index >= mCount Then Exit Function
    If IsObject(mItems(index)) Then Set At = mItems(index) Else At = mItems(index)
End Function

Public Function First() As Variant
    If mCount = 0 Then Exit Function
    If IsObject(mItems(0)) Then Set First = mItems(0) Else First = mItems(0)
End Function

Public Function Last() As Variant
    If mCount = 0 Then Exit Function
    If IsObject(mItems(mCount - 1)) Then Set Last = mItems(mCount - 1) Else Last = mItems(mCount - 1)
End Function

Public Sub Add(ByVal value As Variant)
    If mCount > UBound(mItems) Then EnsureCapacity mCount + 1
    PutAt mCount, value
    mCount = mCount + 1
End Sub

' arr.push(...values): returns the new length.
Public Function Push(ParamArray values() As Variant) As Long
    Dim i As Long
    For i = LBound(values) To UBound(values)
        Add values(i)
    Next
    Push = mCount
End Function

' arr.pop(): Empty when the list is empty.
Public Function Pop() As Variant
    If mCount = 0 Then Exit Function
    mCount = mCount - 1
    If IsObject(mItems(mCount)) Then Set Pop = mItems(mCount) Else Pop = mItems(mCount)
    mItems(mCount) = Empty
End Function

' arr.shift(): Empty when the list is empty.
Public Function Shift() As Variant
    If mCount = 0 Then Exit Function
    If IsObject(mItems(0)) Then Set Shift = mItems(0) Else Shift = mItems(0)
    RemoveRange 0, 1
End Function

' arr.unshift(...values): returns the new length.
Public Function Unshift(ParamArray values() As Variant) As Long
    Dim i As Long, k As Long
    k = UBound(values) - LBound(values) + 1
    If k > 0 Then
        OpenGap 0, k
        For i = 0 To k - 1
            PutAt i, values(LBound(values) + i)
        Next
    End If
    Unshift = mCount
End Function

' Inserts value before position index (0..Count).
Public Sub Insert(ByVal index As Long, ByVal value As Variant)
    If index < 0 Then index = 0
    If index > mCount Then index = mCount
    OpenGap index, 1
    PutAt index, value
End Sub

' Removes and returns the item at index (Empty when out of range).
Public Function RemoveAt(ByVal index As Long) As Variant
    If index < 0 Or index >= mCount Then Exit Function
    If IsObject(mItems(index)) Then Set RemoveAt = mItems(index) Else RemoveAt = mItems(index)
    RemoveRange index, 1
End Function

Public Sub Swap(ByVal i As Long, ByVal j As Long)
    Dim t As Variant
    If i < 0 Or j < 0 Or i >= mCount Or j >= mCount Then Err.Raise 9, "clsList", "Index out of range"
    If IsObject(mItems(i)) Then Set t = mItems(i) Else t = mItems(i)
    MoveSlot i, j
    If IsObject(t) Then Set mItems(j) = t Else mItems(j) = t
End Sub

Private Sub OpenGap(ByVal index As Long, ByVal k As Long)
    Dim i As Long
    EnsureCapacity mCount + k
    For i = mCount - 1 To index Step -1
        MoveSlot i + k, i
    Next
    mCount = mCount + k
End Sub

Private Sub RemoveRange(ByVal index As Long, ByVal k As Long)
    Dim i As Long
    For i = index + k To mCount - 1
        MoveSlot i - k, i
    Next
    For i = mCount - k To mCount - 1
        mItems(i) = Empty
    Next
    mCount = mCount - k
End Sub

Private Function RelIndex(ByVal v As Variant, ByVal dflt As Long) As Long
    Dim d As Long
    If IsMissing(v) Then RelIndex = dflt: Exit Function
    d = CLng(v)
    If d < 0 Then
        d = d + mCount
        If d < 0 Then d = 0
    ElseIf d > mCount Then
        d = mCount
    End If
    RelIndex = d
End Function

' arr.slice([start], [end]) -> new list (negative indices count from the end).
Public Function Slice(Optional ByVal start As Variant, Optional ByVal endIndex As Variant) As clsList
    Dim r As clsList, a As Long, b As Long, i As Long
    Set r = New clsList
    a = RelIndex(start, 0)
    b = RelIndex(endIndex, mCount)
    For i = a To b - 1
        r.Add mItems(i)
    Next
    Set Slice = r
End Function

' arr.splice(start, deleteCount, ...items) -> list of removed items. Pass Count (or any
' large number) as deleteCount to remove everything from start; negative counts are 0.
Public Function Splice(ByVal start As Long, ByVal deleteCount As Long, ParamArray items() As Variant) As clsList
    Dim r As clsList, i As Long, k As Long
    Set r = New clsList
    start = RelIndex(start, 0)
    If deleteCount < 0 Then deleteCount = 0
    If deleteCount > mCount - start Then deleteCount = mCount - start
    For i = start To start + deleteCount - 1
        r.Add mItems(i)
    Next
    If deleteCount > 0 Then RemoveRange start, deleteCount
    k = UBound(items) - LBound(items) + 1
    If k > 0 Then
        OpenGap start, k
        For i = 0 To k - 1
            PutAt start + i, items(LBound(items) + i)
        Next
    End If
    Set Splice = r
End Function

' Strict equality (===): numbers numerically (NaN never equal), strings binary,
' objects by identity, different kinds never equal.
Private Function StrictEq(ByRef a As Variant, ByRef b As Variant) As Boolean
    If IsObject(a) Or IsObject(b) Then
        If IsObject(a) And IsObject(b) Then StrictEq = (a Is b)
        Exit Function
    End If
    Select Case KindOf(a)
    Case 1
        If KindOf(b) = 1 Then
            If Not (modJs.JsIsNaN(CDbl(a)) Or modJs.JsIsNaN(CDbl(b))) Then StrictEq = (CDbl(a) = CDbl(b))
        End If
    Case 2
        If KindOf(b) = 2 Then StrictEq = (StrComp(a, b, vbBinaryCompare) = 0)
    Case 3
        If KindOf(b) = 3 Then StrictEq = (a = b)
    Case Else
        StrictEq = (VarType(a) = VarType(b))
    End Select
End Function

' 1 number, 2 string, 3 boolean, 0 Empty, 4 Null, 5 other.
Private Function KindOf(ByRef v As Variant) As Long
    Select Case VarType(v)
    Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal: KindOf = 1
    Case vbString: KindOf = 2
    Case vbBoolean: KindOf = 3
    Case vbEmpty: KindOf = 0
    Case vbNull: KindOf = 4
    Case Else: KindOf = 5
    End Select
End Function

' arr.indexOf(value, [fromIndex]) using ===; -1 when absent.
Public Function IndexOf(ByVal value As Variant, Optional ByVal fromIndex As Variant) As Long
    Dim i As Long
    For i = RelIndex(fromIndex, 0) To mCount - 1
        If StrictEq(mItems(i), value) Then IndexOf = i: Exit Function
    Next
    IndexOf = -1
End Function

Public Function LastIndexOf(ByVal value As Variant) As Long
    Dim i As Long
    For i = mCount - 1 To 0 Step -1
        If StrictEq(mItems(i), value) Then LastIndexOf = i: Exit Function
    Next
    LastIndexOf = -1
End Function

' arr.includes(value): like indexOf but NaN matches NaN.
Public Function Includes(ByVal value As Variant) As Boolean
    Dim i As Long
    If Not IsObject(value) Then
        If KindOf(value) = 1 Then
            If modJs.JsIsNaN(CDbl(value)) Then
                For i = 0 To mCount - 1
                    If Not IsObject(mItems(i)) Then
                        If KindOf(mItems(i)) = 1 Then
                            If modJs.JsIsNaN(CDbl(mItems(i))) Then Includes = True: Exit Function
                        End If
                    End If
                Next
                Exit Function
            End If
        End If
    End If
    Includes = (IndexOf(value) >= 0)
End Function

' arr.join(separator): numbers via JsNumberToString, Empty/Null as "".
Public Function Join(Optional ByVal separator As String = ",") As String
    Dim parts() As String, i As Long
    If mCount = 0 Then Exit Function
    ReDim parts(0 To mCount - 1)
    For i = 0 To mCount - 1
        If IsObject(mItems(i)) Then
            parts(i) = modJs.JsStr(mItems(i))
        ElseIf IsEmpty(mItems(i)) Or IsNull(mItems(i)) Then
            parts(i) = ""
        Else
            parts(i) = modJs.JsStr(mItems(i))
        End If
    Next
    Join = VBA.Join(parts, separator)
End Function

' Reverses in place; returns Me.
Public Function Reverse() As clsList
    Dim i As Long, j As Long
    i = 0
    j = mCount - 1
    Do While i < j
        Swap i, j
        i = i + 1
        j = j - 1
    Loop
    Set Reverse = Me
End Function

' arr.concat(other) -> new list.
Public Function Concat(ByVal other As clsList) As clsList
    Dim r As clsList, i As Long
    Set r = Clone()
    If Not other Is Nothing Then
        For i = 0 To other.Count - 1
            r.Add other.Item(i)
        Next
    End If
    Set Concat = r
End Function

Public Function Clone() As clsList
    Dim r As clsList, i As Long
    Set r = New clsList
    For i = 0 To mCount - 1
        r.Add mItems(i)
    Next
    Set Clone = r
End Function

Public Sub Clear()
    ReDim mItems(0 To 7)
    mCount = 0
End Sub

' 0-based Variant array copy (Array() with UBound -1 when empty).
Public Function ToArray() As Variant
    Dim a() As Variant, i As Long
    If mCount = 0 Then ToArray = Array(): Exit Function
    ReDim a(0 To mCount - 1)
    For i = 0 To mCount - 1
        If IsObject(mItems(i)) Then Set a(i) = mItems(i) Else a(i) = mItems(i)
    Next
    ToArray = a
End Function

' ---- sorting (stable merge sort) ----

' Compares sort keys: numbers numerically (NaN after all numbers), strings by UTF-16
' code units, numbers before strings, Empty last.
Private Function CompareKeys(ByRef a As Variant, ByRef b As Variant) As Long
    Dim ka As Long, kb As Long, na As Boolean, nb As Boolean, x As Double, y As Double
    ka = KindOf(a)
    kb = KindOf(b)
    If ka = 0 Or kb = 0 Then
        If ka = kb Then
            CompareKeys = 0
        ElseIf ka = 0 Then
            CompareKeys = 1
        Else
            CompareKeys = -1
        End If
        Exit Function
    End If
    If ka = 1 And kb = 1 Then
        x = CDbl(a)
        y = CDbl(b)
        na = modJs.JsIsNaN(x)
        nb = modJs.JsIsNaN(y)
        If na Or nb Then
            If na And nb Then
                CompareKeys = 0
            ElseIf na Then
                CompareKeys = 1
            Else
                CompareKeys = -1
            End If
        ElseIf x < y Then
            CompareKeys = -1
        ElseIf x > y Then
            CompareKeys = 1
        End If
        Exit Function
    End If
    If ka = 1 Then CompareKeys = -1: Exit Function
    If kb = 1 Then CompareKeys = 1: Exit Function
    CompareKeys = StrComp(CStr(a), CStr(b), vbBinaryCompare)
End Function

Private Sub SortWithKeys(keys() As Variant, ByVal descending As Boolean)
    Dim idx() As Long, tmp() As Long, i As Long, w As Long, lo As Long, mid As Long, hi As Long
    Dim a As Long, b As Long, k As Long, c As Long, items() As Variant
    If mCount < 2 Then Exit Sub
    ReDim idx(0 To mCount - 1)
    ReDim tmp(0 To mCount - 1)
    For i = 0 To mCount - 1
        idx(i) = i
    Next
    w = 1
    Do While w < mCount
        lo = 0
        Do While lo < mCount
            mid = lo + w
            If mid > mCount Then mid = mCount
            hi = lo + 2 * w
            If hi > mCount Then hi = mCount
            a = lo
            b = mid
            k = lo
            Do While a < mid And b < hi
                c = CompareKeys(keys(idx(b)), keys(idx(a)))
                If descending Then c = -c
                If c < 0 Then
                    tmp(k) = idx(b)
                    b = b + 1
                Else
                    tmp(k) = idx(a)
                    a = a + 1
                End If
                k = k + 1
            Loop
            Do While a < mid
                tmp(k) = idx(a)
                a = a + 1
                k = k + 1
            Loop
            Do While b < hi
                tmp(k) = idx(b)
                b = b + 1
                k = k + 1
            Loop
            lo = hi
        Loop
        For i = 0 To mCount - 1
            idx(i) = tmp(i)
        Next
        w = w * 2
    Loop
    ReDim items(0 To mCount - 1)
    For i = 0 To mCount - 1
        If IsObject(mItems(idx(i))) Then Set items(i) = mItems(idx(i)) Else items(i) = mItems(idx(i))
    Next
    For i = 0 To mCount - 1
        PutAt i, items(i)
    Next
End Sub

' Default JS sort(): compares String(item) by UTF-16 code units (stable). Returns Me.
Public Function SortStrings(Optional ByVal descending As Boolean = False) As clsList
    Dim keys() As Variant, i As Long
    If mCount > 0 Then
        ReDim keys(0 To mCount - 1)
        For i = 0 To mCount - 1
            If IsEmpty(mItems(i)) Then keys(i) = Empty Else keys(i) = modJs.JsStr(mItems(i))
        Next
        SortWithKeys keys, descending
    End If
    Set SortStrings = Me
End Function

' sort((a, b) => a - b) for numeric items (stable; NaN last). Returns Me.
Public Function SortNumbers(Optional ByVal descending As Boolean = False) As clsList
    Dim keys() As Variant, i As Long
    If mCount > 0 Then
        ReDim keys(0 To mCount - 1)
        For i = 0 To mCount - 1
            keys(i) = modJs.JsToNumber(mItems(i))
        Next
        SortWithKeys keys, descending
    End If
    Set SortNumbers = Me
End Function

' Stable sort of the items by the parallel keys list (numbers or strings; the keys list
' itself is not changed). Returns Me.
Public Function SortBy(ByVal keys As clsList, Optional ByVal descending As Boolean = False) As clsList
    Dim k() As Variant, i As Long
    If keys.Count <> mCount Then Err.Raise 5, "clsList.SortBy", "keys must have one entry per item"
    If mCount > 0 Then
        ReDim k(0 To mCount - 1)
        For i = 0 To mCount - 1
            k(i) = keys.Item(i)
        Next
        SortWithKeys k, descending
    End If
    Set SortBy = Me
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Dim c As Collection, i As Long
    Set c = New Collection
    For i = 0 To mCount - 1
        c.Add mItems(i)
    Next
    ' A Collection's enumerator does not keep the Collection alive, so the snapshot is
    ' held here; the last few are kept so that nested loops over this object work.
    If mSnaps Is Nothing Then Set mSnaps = New Collection
    If mSnaps.Count >= 4 Then mSnaps.Remove 1
    mSnaps.Add c
    Set NewEnum = c.[_NewEnum]
End Function
