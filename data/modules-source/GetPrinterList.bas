Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim objWMI As Object, col As Object, obj As Object
    Set objWMI = GetObject("winmgmts:\\.ootcimv2")
    Set col = objWMI.ExecQuery("SELECT Name, Default, PrinterStatus, DriverName FROM Win32_Printer")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1:D1").Value = Array("Yazıcı Adı", "Varsayılan", "Durum", "Sürücü")
    ws.Range("A1:D1").Font.Bold = True
    Dim r As Long : r = 2
    For Each obj In col
        ws.Cells(r, 1).Value = obj.Name
        ws.Cells(r, 2).Value = IIf(obj.Default, "Evet", "Hayır")
        ws.Cells(r, 3).Value = obj.PrinterStatus
        ws.Cells(r, 4).Value = obj.DriverName
        r = r + 1
    Next obj
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function