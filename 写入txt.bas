Attribute VB_Name = "ģ��1"
Sub write_in_txt()
    Dim text_file
    Dim fs
    Dim text
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set text_file = fs.CreateTextfile(ThisWorkbook.Path & "\hello_world.txt", True)
    
    text = "Hello world ��"
    text_file.writeline (text)
    text_file.Close
End Sub
