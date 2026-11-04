Attribute VB_Name = "modClipboard"
' Plain-text clipboard for Windows (Win32 API) and Mac (MSForms DataObject).
Option Explicit

#If Mac Then
#Else
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal format As Long, ByVal memory As LongPtr) As LongPtr
Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal format As Long) As LongPtr
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal flags As Long, ByVal bytes As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal memory As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal memory As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal memory As LongPtr) As Long
Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal text As LongPtr) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal destination As LongPtr, ByVal source As LongPtr, ByVal bytes As LongPtr)
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)
Private Const CF_UNICODETEXT As Long = 13
Private Const GMEM_MOVEABLE As Long = &H2
#End If

' The last text written, so tests and the app can check what was copied.
Public LastClipboardText As String

' Returns True when the text reached the clipboard. Line endings become the platform's own.
Public Function ClipboardWrite(ByVal text As String) As Boolean
    LastClipboardText = text
    On Error GoTo Failed
#If Mac Then
    Dim data As New MSForms.DataObject
    data.SetText Replace(text, vbCrLf, vbLf)
    data.PutInClipboard
    ClipboardWrite = True
#Else
    Dim memory As LongPtr, pointer As LongPtr, bytes As LongPtr, attempt As Long, opened As Boolean
    text = Replace(Replace(text, vbCrLf, vbLf), vbLf, vbCrLf) & vbNullChar
    bytes = LenB(text)
    memory = GlobalAlloc(GMEM_MOVEABLE, bytes)
    If memory = 0 Then Exit Function
    pointer = GlobalLock(memory)
    CopyMemory pointer, StrPtr(text), bytes
    GlobalUnlock memory
    For attempt = 1 To 10
        If OpenClipboard(0) <> 0 Then opened = True: Exit For
        Sleep 20
    Next attempt
    If Not opened Then
        GlobalFree memory
        Exit Function
    End If
    EmptyClipboard
    If SetClipboardData(CF_UNICODETEXT, memory) = 0 Then GlobalFree memory Else ClipboardWrite = True
    CloseClipboard
#End If
    Exit Function
Failed:
    ClipboardWrite = False
End Function

' Returns the clipboard's text with LF line endings ("" when it holds no text).
Public Function ClipboardRead() As String
    On Error GoTo Failed
#If Mac Then
    Dim data As New MSForms.DataObject
    data.GetFromClipboard
    ClipboardRead = data.GetText
#Else
    Dim memory As LongPtr, pointer As LongPtr, length As Long, attempt As Long, opened As Boolean
    For attempt = 1 To 10
        If OpenClipboard(0) <> 0 Then opened = True: Exit For
        Sleep 20
    Next attempt
    If Not opened Then Exit Function
    memory = GetClipboardData(CF_UNICODETEXT)
    If memory <> 0 Then
        pointer = GlobalLock(memory)
        If pointer <> 0 Then
            length = lstrlenW(pointer)
            ClipboardRead = String$(length, vbNullChar)
            If length > 0 Then CopyMemory StrPtr(ClipboardRead), pointer, CLngPtr(length) * 2
            GlobalUnlock memory
        End If
    End If
    CloseClipboard
#End If
    ClipboardRead = Replace(Replace(ClipboardRead, vbCrLf, vbLf), vbCr, vbLf)
    Exit Function
Failed:
    ClipboardRead = ""
End Function
