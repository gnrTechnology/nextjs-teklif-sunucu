Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim objWMI As Object, col As Object, obj As Object
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set col = objWMI.ExecQuery("SELECT Manufacturer, Name, Version, ReleaseDate FROM Win32_BIOS")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1:D1").Value = Array("Üretici", "BIOS Adı", "Sürüm", "Tarih")
    ws.Range("A1:D1").Font.Bold = True
    Dim r As Long : r = 2
    For Each obj In col
        ws.Cells(r, 1).Value = obj.Manufacturer
        ws.Cells(r, 2).Value = obj.Name
        ws.Cells(r, 3).Value = obj.Version
        Dim rawDate As String : rawDate = CStr(obj.ReleaseDate)
        ws.Cells(r, 4).Value = Left(rawDate, 4) & "-" & Mid(rawDate, 5, 2) & "-" & Mid(rawDate, 7, 2)
        r = r + 1
    Next obj
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function