Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim objWMI As Object, col As Object, obj As Object
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set col = objWMI.ExecQuery("SELECT Caption, Bias, StandardName FROM Win32_TimeZone")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1:C1").Value = Array("Saat Dilimi", "UTC Farkı (dk)", "Standart Ad")
    ws.Range("A1:C1").Font.Bold = True
    Dim r As Long : r = 2
    For Each obj In col
        ws.Cells(r, 1).Value = obj.Caption
        ws.Cells(r, 2).Value = obj.Bias
        ws.Cells(r, 3).Value = obj.StandardName
        r = r + 1
    Next obj
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function