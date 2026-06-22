Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim objWMI As Object, col As Object, obj As Object
    Set objWMI = GetObject("winmgmts:\\.ootcimv2")
    Set col = objWMI.ExecQuery("SELECT Caption, Version, BuildNumber, OSArchitecture FROM Win32_OperatingSystem")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1:D1").Value = Array("İşletim Sistemi", "Sürüm", "Build", "Mimari")
    ws.Range("A1:D1").Font.Bold = True
    Dim r As Long : r = 2
    For Each obj In col
        ws.Cells(r, 1).Value = obj.Caption
        ws.Cells(r, 2).Value = obj.Version
        ws.Cells(r, 3).Value = obj.BuildNumber
        ws.Cells(r, 4).Value = obj.OSArchitecture
        r = r + 1
    Next obj
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function