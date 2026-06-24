Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim sh As Worksheet, pt As PivotTable, n As Long : n = 0
    For Each sh In targetWb.Worksheets
        For Each pt In sh.PivotTables
            pt.RefreshTable : n = n + 1
        Next pt
    Next sh
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Yenilenen pivot" : ws.Range("B1").Value = n
    Set DynamicFunc = Nothing
End Function
