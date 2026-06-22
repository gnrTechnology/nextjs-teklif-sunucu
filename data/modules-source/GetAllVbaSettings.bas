Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim sections() As String : sections = Split("ilhan,scngnr,sercan", ",")
    Dim keys() As String : keys = Split("mac,mdip,TBveren,teklifYolu,startingAddin,ihlalDosyaYolu,ihlalDosyaAdi,apiBaseUrl", ",")
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Bölüm" : ws.Range("B1").Value = "Anahtar" : ws.Range("C1").Value = "Değer"
    ws.Range("A1:C1").Font.Bold = True
    Dim r As Long : r = 2
    Dim sec As Variant, k As Variant
    For Each sec In sections
        For Each k In keys
            Dim v As String : v = GetSetting(CStr(sec), "Settings", CStr(k), "(yok)")
            ws.Cells(r,1).Value = sec : ws.Cells(r,2).Value = k : ws.Cells(r,3).Value = v : r=r+1
        Next k
    Next sec
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function