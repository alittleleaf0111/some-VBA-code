Attribute VB_Name = "ģ��1"
Public Function select_count(c As Range, a As Range)
    Dim myrange As Range, n As Range, i As Integer
    For Each n In c
        If n.Value = a.Value Then
            select_count = select_count + 1
        End If
    Next n
End Function
