Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim objWMI As Object, col As Object, obj As Object
    Set objWMI = GetObject("winmgmts:\\.ootcimv2")
    Set col = objWMI.ExecQuery("SELECT TotalVisibleMemorySize, FreePhysicalMemory FROM Win32_OperatingSystem")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1:C1").Value = Array("Toplam RAM (GB)", "Boş RAM (GB)", "Kullanılan RAM (GB)")
    ws.Range("A1:C1").Font.Bold = True
    For Each obj In col
        Dim total As Double : total = obj.TotalVisibleMemorySize / 1048576
        Dim free  As Double : free  = obj.FreePhysicalMemory / 1048576
        ws.Range("A2:C2").Value = Array(Format(total, "0.00"), Format(free, "0.00"), Format(total - free, "0.00"))
    Next obj
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function