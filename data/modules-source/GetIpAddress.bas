Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim objWMI As Object, col As Object, obj As Object
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set col = objWMI.ExecQuery("SELECT IPAddress, Description FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1:B1").Value = Array("IP Adresi", "Adaptör")
    ws.Range("A1:B1").Font.Bold = True
    Dim r As Long : r = 2
    For Each obj In col
        If Not IsNull(obj.IPAddress) Then
            Dim ip As Variant
            For Each ip In obj.IPAddress
                ws.Cells(r, 1).Value = ip
                ws.Cells(r, 2).Value = obj.Description
                r = r + 1
            Next ip
        End If
    Next obj
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function