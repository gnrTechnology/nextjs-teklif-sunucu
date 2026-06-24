Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim lic As String : lic = GetSetting("ilhan", "Settings", "license", "")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Lisans" : ws.Range("B1").Value = lic
    ws.Range("A2").Value = "Aktif" : ws.Range("B2").Value = (LCase$(lic) = "true" Or lic = "1")
    Set DynamicFunc = Nothing
End Function
