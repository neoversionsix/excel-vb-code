'This copies the text stored in the TextBox called 'OUTPUT' into the Clipboard
' DOES not require any reference libraries in excel
Option Explicit

Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function lstrcpy Lib "Kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "Kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "Kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "Kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "Kernel32" (ByVal hMem As Long) As Long

Private Const GHND = &H42
Private Const CF_TEXT = 1
Private Const MAXSIZE = 4096

Sub CopyTextboxToClipboard()
    Dim text As String
    text = ActiveSheet.Shapes("OUTPUT").TextFrame.Characters.text
    ClipBoard_SetData text
End Sub

Public Function ClipBoard_SetData(MyString As String)
    Dim hGlobalMemory As Long, lpGlobalMemory As Long
    Dim hClipMemory As Long, X As Long

    ' Allocate moveable global memory.
    '-------------------------------------------
    hGlobalMemory = GlobalAlloc(GHND, Len(MyString) + 1)

    ' Lock the block to get a far pointer
    ' to this memory.
    lpGlobalMemory = GlobalLock(hGlobalMemory)

    ' Copy the string to this global memory.
    lpGlobalMemory = lstrcpy(lpGlobalMemory, MyString)

    ' Unlock the memory.
    If GlobalUnlock(hGlobalMemory) <> 0 Then
        MsgBox "Could not unlock memory location. Copy aborted."
        GoTo OutOfHere2
    End If

    ' Open the Clipboard to copy data to.
    If OpenClipboard(0&) = 0 Then
        MsgBox "Could not open the Clipboard. Copy aborted."
        Exit Function
    End If

    ' Clear the Clipboard.
    X = EmptyClipboard()

    ' Copy the data to the Clipboard.
    hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory)

OutOfHere2:
    If CloseClipboard() = 0 Then
        MsgBox "Could not close Clipboard."
    End If
End Function