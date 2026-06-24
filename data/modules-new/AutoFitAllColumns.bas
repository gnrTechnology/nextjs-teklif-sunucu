Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim sh As Worksheet
    For Each sh In targetWb.Worksheets
        sh.Columns.AutoFit
    Next sh
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Sayfa" : ws.Range("B1").Value = targetWb.Worksheets.Count
    Set DynamicFunc = Nothing
End Function
