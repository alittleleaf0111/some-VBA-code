Attribute VB_Name = "Ä£¿é1"
Sub ontime()
    Application.ontime Now() + TimeValue("00:00:01"), "nowtime"
End Sub
Sub nowtime()
    Sheet1.Range("L1") = Format(Now(), "yyyy-mm-dd hh:mm:ss")
    Call ontime
End Sub
