' This is for writing UTF8 text to the textbox
Sub WriteTextFileUTF8(fileName As String, outputText As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 'Specify stream type - we want To save text/string data.
    stream.Charset = "utf-8" 'Specify charset For the source text data.
    stream.Open 'Open the stream And write binary data To the object
    stream.WriteText outputText
    stream.SaveToFile fileName, 2 'Save binary data To disk
    stream.Close

    ' Re-open the file and remove the BOM
    Dim fileNum As Integer
    fileNum = FreeFile
    Open fileName For Binary As #fileNum
    Dim fileContents As String
    fileContents = Input$(LOF(fileNum), #fileNum)
    Close #fileNum
    ' Check if the file starts with the BOM and remove it
    If Left$(fileContents, 3) = Chr$(&HEF) & Chr$(&HBB) & Chr$(&HBF) Then
        fileContents = Mid$(fileContents, 4)
        ' Save the file again without the BOM
        Open fileName For Output As #fileNum
        Print #fileNum, fileContents;
        Close #fileNum
    End If
End Sub
