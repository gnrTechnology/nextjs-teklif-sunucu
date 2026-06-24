Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim n As Long : n = 0
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If Not wb.IsAddin Then wb.Save : n = n + 1
    Next wb
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Kaydedilen" : ws.Range("B1").Value = n
    Set DynamicFunc = Nothing
End Function
