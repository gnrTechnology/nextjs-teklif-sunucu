Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim sumName As String : sumName = "Ozet"
    On Error Resume Next
    Application.DisplayAlerts = False
    targetWb.Worksheets(sumName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Dim wsSum As Worksheet : Set wsSum = targetWb.Worksheets.Add(Before:=targetWb.Sheets(1))
    wsSum.Name = sumName
    wsSum.Range("A1").Value = "Sayfa" : wsSum.Range("B1").Value = "A1"
    Dim sh As Worksheet, r As Long : r = 2
    For Each sh In targetWb.Worksheets
        If sh.Name <> sumName Then
            wsSum.Cells(r, 1).Value = sh.Name
            wsSum.Cells(r, 2).Value = sh.Range("A1").Value
            r = r + 1
        End If
    Next sh
    Set DynamicFunc = Nothing
End Function
