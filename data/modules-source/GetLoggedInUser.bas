Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1:B1").Value = Array("Alan", "Değer")
    ws.Range("A1:B1").Font.Bold = True
    ws.Range("A2:B2").Value = Array("Kullanıcı Adı", Environ("USERNAME"))
    ws.Range("A3:B3").Value = Array("Bilgisayar Adı", Environ("COMPUTERNAME"))
    ws.Range("A4:B4").Value = Array("Kullanıcı Profil Yolu", Environ("USERPROFILE"))
    ws.Range("A5:B5").Value = Array("AppData", Environ("APPDATA"))
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function