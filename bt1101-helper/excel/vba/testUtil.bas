Attribute VB_Name = "testUtil"
' Test-only helpers: reading UTF-8 files and building compact mismatch reports.
Option Explicit

' Reads a UTF-8 (optionally BOM-prefixed) file into a VBA string.
Public Function ReadUtf8File(ByVal path As String) As String
    Dim handle As Integer, bytes() As Byte, size As Long
    handle = FreeFile
    Open path For Binary Access Read As #handle
    size = LOF(handle)
    If size = 0 Then
        Close #handle
        Exit Function
    End If
    ReDim bytes(0 To size - 1)
    Get #handle, , bytes
    Close #handle
    ReadUtf8File = DecodeUtf8(bytes)
End Function

Public Function DecodeUtf8(ByRef bytes() As Byte) As String
    Dim out() As Integer, n As Long, i As Long, last As Long, b As Long, code As Long, extra As Long
    last = UBound(bytes)
    ReDim out(0 To last + 1)
    i = LBound(bytes)
    If last - i >= 2 Then
        If bytes(i) = &HEF And bytes(i + 1) = &HBB And bytes(i + 2) = &HBF Then i = i + 3
    End If
    Do While i <= last
        b = bytes(i)
        If b < &H80 Then
            code = b: extra = 0
        ElseIf b < &HE0 Then
            code = b And &H1F: extra = 1
        ElseIf b < &HF0 Then
            code = b And &HF: extra = 2
        Else
            code = b And &H7: extra = 3
        End If
        Do While extra > 0 And i < last
            i = i + 1
            code = code * 64 + (bytes(i) And &H3F)
            extra = extra - 1
        Loop
        If code >= &H10000 Then
            code = code - &H10000
            out(n) = UnsignedToInteger(&HD800& + code \ 1024): n = n + 1
            out(n) = UnsignedToInteger(&HDC00& + (code And 1023)): n = n + 1
        Else
            out(n) = UnsignedToInteger(code): n = n + 1
        End If
        i = i + 1
    Loop
    If n = 0 Then Exit Function
    DecodeUtf8 = String$(n, vbNullChar)
    For i = 0 To n - 1
        Mid$(DecodeUtf8, i + 1, 1) = ChrW$(out(i))
    Next i
End Function

Private Function UnsignedToInteger(ByVal value As Long) As Integer
    If value >= 32768 Then UnsignedToInteger = CInt(value - 65536) Else UnsignedToInteger = CInt(value)
End Function

' Writes a VBA string as UTF-8 (no BOM).
Public Sub WriteUtf8File(ByVal path As String, ByVal text As String)
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
            ReDim Preserve bytes(0 To UBound(bytes) + 4)
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
