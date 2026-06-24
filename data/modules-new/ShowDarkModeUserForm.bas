Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Not" : ws.Range("B1").Value = "UserForm tasarimi IDE'de yapilmali"
    Set DynamicFunc = Nothing
End Function
