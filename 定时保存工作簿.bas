Attribute VB_Name = "ģ��1"
Sub ontime()
    Application.ontime Now() + TimeValue("00:00:05"), "wbsave"
End Sub

Sub wbsave()
    ThisWorkbook.Save
    Call ontime
End Sub
