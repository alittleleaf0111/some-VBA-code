Attribute VB_Name = "Ä£¿é1"
Public Function CountColor(arr As Range, c As Range)
    Dim rng As Range
    For Each rng In arr
        If rng.Interior.Color = c.Interior.Color Then
            CountColor = CountColor + 1
        End If
    Next rng
End Function
