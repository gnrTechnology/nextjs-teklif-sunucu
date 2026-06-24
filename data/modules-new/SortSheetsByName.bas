Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim i As Long, j As Long
    For i = 1 To targetWb.Worksheets.Count - 1
        For j = i + 1 To targetWb.Worksheets.Count
            If targetWb.Worksheets(j).Name < targetWb.Worksheets(i).Name Then
                targetWb.Worksheets(j).Move Before:=targetWb.Worksheets(i)
            End If
        Next j
    Next i
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Siralandi" : ws.Range("B1").Value = Now
    Set DynamicFunc = Nothing
End Function
