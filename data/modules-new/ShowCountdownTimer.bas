Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim sec As Long : sec = CLng(Val(CStr(param))) : If sec < 1 Then sec = 10
    Dim i As Long
    For i = sec To 0 Step -1
        Application.StatusBar = "Geri sayim: " & i
        Application.Wait Now + TimeValue("00:00:01")
    Next i
    Application.StatusBar = False
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1) : ws.Range("A1").Value = "Bitti" : ws.Range("B1").Value = sec
    Set DynamicFunc = Nothing
End Function
