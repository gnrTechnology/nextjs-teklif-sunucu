Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim json As String : json = CStr(param)
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "JSON icerik (basit)"
    ws.Range("B1").Value = Left$(json, 32000)
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function
