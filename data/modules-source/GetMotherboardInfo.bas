Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim objWMI As Object, col As Object, obj As Object
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set col = objWMI.ExecQuery("SELECT Manufacturer, Product, SerialNumber, Version FROM Win32_BaseBoard")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1:D1").Value = Array("Üretici", "Model", "Seri No", "Sürüm")
    ws.Range("A1:D1").Font.Bold = True
    Dim r As Long : r = 2
    For Each obj In col
        ws.Cells(r, 1).Value = obj.Manufacturer
        ws.Cells(r, 2).Value = obj.Product
        ws.Cells(r, 3).Value = obj.SerialNumber
        ws.Cells(r, 4).Value = obj.Version
        r = r + 1
    Next obj
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function