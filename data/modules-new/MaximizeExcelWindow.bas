Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Application.WindowState = xlMaximized
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1) : ws.Range("A1").Value = "Durum" : ws.Range("B1").Value = "Maximized"
    Set DynamicFunc = Nothing
End Function
