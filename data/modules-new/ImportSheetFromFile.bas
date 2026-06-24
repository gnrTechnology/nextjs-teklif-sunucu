Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim src As String : src = Trim$(CStr(param))
    Dim wb As Workbook : Set wb = Workbooks.Open(src, ReadOnly:=True)
    wb.Sheets(1).Copy After:=targetWb.Sheets(targetWb.Sheets.Count)
    wb.Close False
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Kaynak" : ws.Range("B1").Value = src
    Set DynamicFunc = Nothing
End Function
