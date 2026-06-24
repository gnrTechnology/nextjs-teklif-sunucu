Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim addr As String : addr = Trim$(CStr(param))
    If Len(addr) = 0 Then addr = targetWb.ActiveSheet.UsedRange.Address
    targetWb.ActiveSheet.Shapes.AddChart2 227, xlColumnClustered, 300, 10, 400, 250
    targetWb.ActiveSheet.ChartObjects(1).Chart.SetSourceData targetWb.ActiveSheet.Range(addr)
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Grafik" : ws.Range("B1").Value = "OK"
    Set DynamicFunc = Nothing
End Function
