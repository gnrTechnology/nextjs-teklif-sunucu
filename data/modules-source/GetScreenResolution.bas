Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim objWMI As Object, col As Object, obj As Object
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set col = objWMI.ExecQuery("SELECT CurrentHorizontalResolution, CurrentVerticalResolution, Name FROM Win32_VideoController")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1:C1").Value = Array("Ekran Kartı", "Genişlik (px)", "Yükseklik (px)")
    ws.Range("A1:C1").Font.Bold = True
    Dim r As Long : r = 2
    For Each obj In col
        On Error Resume Next
        If obj.CurrentHorizontalResolution > 0 Then
            ws.Cells(r, 1).Value = obj.Name
            ws.Cells(r, 2).Value = obj.CurrentHorizontalResolution
            ws.Cells(r, 3).Value = obj.CurrentVerticalResolution
            r = r + 1
        End If
        On Error GoTo 0
    Next obj
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function