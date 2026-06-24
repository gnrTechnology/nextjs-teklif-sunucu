Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    With targetWb.ActiveWindow
        .FreezePanes = False
        .SplitColumn = 1
        .SplitRow = 1
        .FreezePanes = True
    End With
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Durum" : ws.Range("B1").Value = "Donduruldu"
    Set DynamicFunc = Nothing
End Function
