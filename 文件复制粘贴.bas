Attribute VB_Name = "ģ��1"
Sub file_copy_paste()
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.copyfile "D:\data.txt", "E:\data.txt"
End Sub
