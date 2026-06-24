Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1) : ws.Cells.ClearContents
    ws.Range("A1").Value = "Modul" : ws.Range("B1").Value = "Satir"
    Dim comp As Object, r As Long : r = 2
    For Each comp In targetWb.VBProject.VBComponents
        ws.Cells(r, 1).Value = comp.Name
        ws.Cells(r, 2).Value = comp.CodeModule.CountOfLines
        r = r + 1
    Next
    Set DynamicFunc = Nothing
End Function
