Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim objWMI As Object, col As Object, obj As Object
    Set objWMI = GetObject("winmgmts:\\.ootcimv2")
    Set col = objWMI.ExecQuery("SELECT Name, ProcessId, WorkingSetSize, KernelModeTime FROM Win32_Process ORDER BY WorkingSetSize DESC")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1:D1").Value = Array("Process Adı", "PID", "Bellek (MB)", "CPU Zamanı (s)")
    ws.Range("A1:D1").Font.Bold = True
    Dim r As Long : r = 2
    For Each obj In col
        ws.Cells(r, 1).Value = obj.Name
        ws.Cells(r, 2).Value = obj.ProcessId
        ws.Cells(r, 3).Value = Format(obj.WorkingSetSize / 1048576, "0.00")
        ws.Cells(r, 4).Value = Format(obj.KernelModeTime / 10000000, "0.0")
        r = r + 1
        If r > 102 Then Exit For ' max 100 satır
    Next obj
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function