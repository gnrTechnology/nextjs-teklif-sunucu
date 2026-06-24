Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim v As String : v = InputBox(CStr(param), "Teklif")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1) : ws.Range("A1").Value = "Girdi" : ws.Range("B1").Value = v
    Set DynamicFunc = v
End Function
