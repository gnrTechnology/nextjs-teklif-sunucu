Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim addr As String : addr = Trim$(CStr(param))
    If Len(addr) = 0 Then addr = targetWb.ActiveSheet.UsedRange.Address
    Dim pc As PivotCache
    Set pc = targetWb.PivotCaches.Create(xlDatabase, targetWb.ActiveSheet.Range(addr))
    Dim pt As PivotTable
    Set pt = pc.CreatePivotTable(targetWb.ActiveSheet.Range("H1"), "PivotAuto")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Pivot" : ws.Range("B1").Value = pt.Name
    Set DynamicFunc = Nothing
End Function
