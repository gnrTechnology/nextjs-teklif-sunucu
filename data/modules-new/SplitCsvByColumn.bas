Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim csvPath As String : csvPath = Trim$(CStr(param))
    Dim wb As Workbook : Set wb = Workbooks.Open(csvPath, ReadOnly:=True)
    wb.Sheets(1).UsedRange.Copy targetWb.Sheets(1).Range("A1")
    Application.CutCopyMode = False
    wb.Close False
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "CSV ice aktarildi" : ws.Range("B1").Value = csvPath
    Set DynamicFunc = Nothing
End Function
