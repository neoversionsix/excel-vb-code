'This requires the reference ligrary FM20.ddl (Microsoft Forms 2.0)
' To find the library browser here (C:\Windows\System32) or here (C:\Windows\SysWOW64)
Sub CopyTextboxToClipboard()

    Dim MyData As DataObject
    Dim text As String
    
    Set MyData = New DataObject
    text = ActiveSheet.Shapes("Textbox1").TextFrame.Characters.Text
    
    MyData.SetText text
    MyData.PutInClipboard

End Sub
