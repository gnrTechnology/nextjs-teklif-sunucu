Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Sonuc"
    ws.Range("B1").Value = "Named Pipe — D39/D40 veya TeklifPipeCom"
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function
