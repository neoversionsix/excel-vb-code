Option Explicit

Sub GenerateTextAndOutput()

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rng As Range
    Dim row As Range
    Dim inputText As String
    Dim outputText As String
    Dim cell As Range
    Dim fileName As String
    Dim command As String
    
    ' Set worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Replace "Sheet1" with your sheet's name if different
    
    ' Set table
    Set tbl = ws.ListObjects("Table1")
    
    ' Get input text
    inputText = ws.Shapes("INPUT").TextFrame2.TextRange.text
    
    ' Initialize output text
    outputText = ""
    
    ' Iterate over each row in the table
    For Each row In tbl.DataBodyRange.Rows
        ' Replace each placeholder in the input text with the corresponding value in the row
        Dim tempText As String
        tempText = inputText
        For Each cell In row.Cells
            tempText = Replace(tempText, tbl.HeaderRowRange.Cells(cell.Column - tbl.HeaderRowRange.Column + 1).Value, cell.Value)
        Next cell
        
        ' Append to output text
        outputText = outputText & tempText & vbNewLine
    Next row
    
    ' Output final text
    ws.Shapes("OUTPUT").TextFrame2.TextRange.text = outputText
    
    ' Write output to temporary text file
    fileName = Environ$("TEMP") & "\clipboard.txt"
    Open fileName For Output As #1
    Print #1, outputText
    Close #1
    
    ' Load text file contents into clipboard
    command = "cmd /c clip < """ & fileName & """"
    Call Shell(command, vbHide)
    
    'Copy the text in 'OUTPUT' Textbox to the clipboard
    ' Note: This requires the vba for 'CopyTextboxToClipboard to be saved as another module
    Call CopyTextboxToClipboard

End Sub

