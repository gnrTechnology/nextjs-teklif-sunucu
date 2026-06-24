Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim d As Variant : d = Application.InputBox("Tarih (gg.aa.yyyy)", "Takvim", Date, Type:=2)
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1) : ws.Range("A1").Value = "Tarih" : ws.Range("B1").Value = d
    Set DynamicFunc = d
End Function
