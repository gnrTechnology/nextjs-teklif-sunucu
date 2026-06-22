Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim objWMI As Object, col As Object, obj As Object
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set col = objWMI.ExecQuery("SELECT Name, Manufacturer, Status, DeviceID FROM Win32_SoundDevice")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1:D1").Value = Array("Ses Aygıtı", "Üretici", "Durum", "Device ID")
    ws.Range("A1:D1").Font.Bold = True
    Dim r As Long : r = 2
    For Each obj In col
        ws.Cells(r, 1).Value = obj.Name
        ws.Cells(r, 2).Value = IIf(IsNull(obj.Manufacturer), "-", obj.Manufacturer)
        ws.Cells(r, 3).Value = obj.Status
        ws.Cells(r, 4).Value = obj.DeviceID
        r = r + 1
    Next obj
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function