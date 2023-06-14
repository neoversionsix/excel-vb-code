' Generates code for updating. For every code value a "base code" is generated after implanting each code value in the spreadsheet.

Option Explicit

' Clipboard API functions
Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long


' Clipboard constants
Private Const GHND = &H42
Private Const CF_TEXT = 1
Private Const MAXSIZE = 4096

Sub GenerateTextAndSaveToClipboard()

    Dim wsBase As Worksheet
    Dim wsGenerator As Worksheet
    Dim tbl As ListObject
    Dim rngCodes As Range
    Dim cell As Range
    Dim starterText As String
    Dim generatedText As String
    Dim allGeneratedText As String

    ' Set worksheets and table
    Set wsBase = ThisWorkbook.Sheets("BASE CODE")
    Set wsGenerator = ThisWorkbook.Sheets("GENERATOR")
    Set tbl = wsGenerator.ListObjects("Table1")

    ' Get starter text
    starterText = wsBase.Range("A2").Value

    ' Get range of CODES column
    Set rngCodes = tbl.ListColumns("CODES").DataBodyRange

    ' Initialize string that will hold all generated text
    allGeneratedText = ""

    ' Iterate over each cell in the CODES column
    For Each cell In rngCodes
        ' Replace "swapme123" with current cell's value
        generatedText = Replace(starterText, "swapme123", cell.Value)
        
        ' Append to allGeneratedText
        allGeneratedText = allGeneratedText & generatedText & vbNewLine
    Next cell

    ' Save allGeneratedText to clipboard using Windows API
    Dim hGlobalMemory As Long
    Dim lpGlobalMemory As Long
    Dim hClipMemory As Long
    Dim X As Long

    ' Allocate movable global memory.
    '-------------------------------------------
    hGlobalMemory = GlobalAlloc(GHND, Len(allGeneratedText) + 1)

    ' Lock the block to get a far pointer
    ' to this memory.
    lpGlobalMemory = GlobalLock(hGlobalMemory)

    ' Copy the string to this global memory.
    lpGlobalMemory = lstrcpy(lpGlobalMemory, allGeneratedText)

    ' Unlock the memory.
    If GlobalUnlock(hGlobalMemory) <> 0 Then
        MsgBox "Could not unlock memory location. Copy aborted."
        GoTo OutOfHere2
    End If

    ' Open the Clipboard to copy data to.
    If OpenClipboard(0&) = 0 Then
        MsgBox "Could not open the Clipboard. Copy aborted."
        Exit Sub
    End If

    ' Clear the Clipboard.
    X = EmptyClipboard()

    ' Copy the data to the Clipboard.
    hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory)

    ' Close Clipboard
    If CloseClipboard() = 0 Then
        MsgBox "Could not close Clipboard."
    End If

OutOfHere2:
    ' Free memory
    If hGlobalMemory <> 0 Then GlobalFree hGlobalMemory

End Sub


