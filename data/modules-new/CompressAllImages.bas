Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim shp As Shape, n As Long : n = 0
    For Each shp In targetWb.ActiveSheet.Shapes
        If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Then n = n + 1
    Next shp
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1) : ws.Range("A1").Value = "Resim" : ws.Range("B1").Value = n & " (sikistirma Office API ile sinirli)"
    Set DynamicFunc = Nothing
End Function
