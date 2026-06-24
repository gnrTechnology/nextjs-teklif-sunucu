Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim addr As String : addr = Trim$(CStr(param))
    If Len(addr) = 0 Then addr = targetWb.ActiveSheet.UsedRange.Address
    targetWb.ActiveSheet.Range(addr).Value = targetWb.ActiveSheet.Range(addr).Value
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Aralik" : ws.Range("B1").Value = addr
    Set DynamicFunc = Nothing
End Function
