Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim parts() As String : parts = Split(CStr(param), "|", 2)
    Dim csv As String : csv = Trim$(parts(0))
    Dim api As String : api = IIf(UBound(parts) >= 1, Trim$(parts(1)), "")
    Workbooks.Open csv : ActiveSheet.UsedRange.Copy targetWb.Sheets(1).Range("A1")
    ActiveWorkbook.Close False
    If Len(api) > 0 Then Application.Run "TableToJsonAndPost", targetWb, api
    Set DynamicFunc = Nothing
End Function
