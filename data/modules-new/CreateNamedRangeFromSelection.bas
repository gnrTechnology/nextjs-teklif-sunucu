Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim nm As String : nm = Trim$(CStr(param))
    targetWb.Names.Add Name:=nm, RefersTo:=targetWb.ActiveSheet.Selection
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Ad" : ws.Range("B1").Value = nm
    Set DynamicFunc = Nothing
End Function
