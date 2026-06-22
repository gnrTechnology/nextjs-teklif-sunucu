Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim objWMI As Object, col As Object, obj As Object
    Set objWMI = GetObject("winmgmts:\\.ootcimv2")
    Set col = objWMI.ExecQuery("SELECT DeviceID, VolumeName, Size, FreeSpace, FileSystem FROM Win32_LogicalDisk WHERE DriveType=3")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1:E1").Value = Array("Sürücü", "Etiket", "Toplam (GB)", "Boş (GB)", "Dosya Sistemi")
    ws.Range("A1:E1").Font.Bold = True
    Dim r As Long : r = 2
    For Each obj In col
        ws.Cells(r, 1).Value = obj.DeviceID
        ws.Cells(r, 2).Value = obj.VolumeName
        ws.Cells(r, 3).Value = Format(obj.Size / 1073741824, "0.00")
        ws.Cells(r, 4).Value = Format(obj.FreeSpace / 1073741824, "0.00")
        ws.Cells(r, 5).Value = obj.FileSystem
        r = r + 1
    Next obj
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function