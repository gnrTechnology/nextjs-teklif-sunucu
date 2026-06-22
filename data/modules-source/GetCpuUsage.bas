Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim wmi As Object : Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Dim col As Object, obj As Object
    Set col = wmi.ExecQuery("SELECT LoadPercentage FROM Win32_Processor")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Çekirdek" : ws.Range("B1").Value = "CPU %"
    ws.Range("A1:B1").Font.Bold = True
    Dim r As Long : r = 2
    For Each obj In col
        ws.Cells(r, 1).Value = "CPU " & (r - 1)
        ws.Cells(r, 2).Value = obj.LoadPercentage & " %"
        r = r + 1
    Next
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function