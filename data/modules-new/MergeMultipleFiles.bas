Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim list As String : list = CStr(param)
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    Dim r As Long : r = 1
    Dim f As Variant
    For Each f In Split(list, ";")
        Dim wb As Workbook : Set wb = Workbooks.Open(Trim$(CStr(f)), ReadOnly:=True)
        wb.Sheets(1).UsedRange.Copy ws.Cells(r, 1)
        r = r + wb.Sheets(1).UsedRange.Rows.Count
        wb.Close False
    Next f
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Birlestirildi" : ws.Range("B1").Value = r
    Set DynamicFunc = Nothing
End Function
