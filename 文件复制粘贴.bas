Attribute VB_Name = "Ä£¿é1"
Sub file_copy_paste()
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.copyfile "D:\data.txt", "E:\data.txt"
End Sub
