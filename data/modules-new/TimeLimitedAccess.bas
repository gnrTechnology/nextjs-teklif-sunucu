Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim parts() As String : parts = Split(CStr(param), "-")
    Dim h1 As Long : h1 = CLng(Val(parts(0)))
    Dim h2 As Long : h2 = IIf(UBound(parts) >= 1, CLng(Val(parts(1))), 18)
    Dim cur As Long : cur = Hour(Now)
    Dim ok As Boolean : ok = (cur >= h1 And cur < h2)
    If Not ok Then targetWb.Windows(1).Visible = False
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Saat" : ws.Range("B1").Value = cur
    ws.Range("A2").Value = "Erisim" : ws.Range("B2").Value = ok
    Set DynamicFunc = Nothing
End Function
