Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    targetWb.Windows(1).Visible = False
    Application.Wait Now + TimeValue("00:00:02")
    targetWb.Windows(1).Visible = True
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1) : ws.Range("A1").Value = "Not" : ws.Range("B1").Value = "Tam engel icin D109 TeklifNotifyCom gerekir"
    Set DynamicFunc = Nothing
End Function
