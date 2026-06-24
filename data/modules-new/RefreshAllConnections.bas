Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim cn As WorkbookConnection
    For Each cn In targetWb.Connections
        On Error Resume Next : cn.Refresh : On Error GoTo 0
    Next cn
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Baglanti" : ws.Range("B1").Value = targetWb.Connections.Count
    Set DynamicFunc = Nothing
End Function
