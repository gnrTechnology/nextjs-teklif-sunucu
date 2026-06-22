Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim objWMI As Object, col As Object, obj As Object
    Set objWMI = GetObject("winmgmts:\\.ootcimv2")
    Set col = objWMI.ExecQuery("SELECT Name, AdapterRAM, DriverVersion, VideoProcessor FROM Win32_VideoController")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1:D1").Value = Array("Ekran Kartı", "VRAM (MB)", "Sürücü", "İşlemci")
    ws.Range("A1:D1").Font.Bold = True
    Dim r As Long : r = 2
    For Each obj In col
        ws.Cells(r, 1).Value = obj.Name
        On Error Resume Next
        ws.Cells(r, 2).Value = Format(obj.AdapterRAM / 1048576, "0")
        On Error GoTo 0
        ws.Cells(r, 3).Value = obj.DriverVersion
        ws.Cells(r, 4).Value = obj.VideoProcessor
        r = r + 1
    Next obj
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function