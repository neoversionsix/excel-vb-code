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
End Sub