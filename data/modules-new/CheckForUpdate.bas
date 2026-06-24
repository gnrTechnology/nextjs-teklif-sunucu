Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim localVer As String, modName As String
    Dim parts() As String : parts = Split(CStr(param), "|", 2)
    modName = Trim$(parts(0))
    localVer = IIf(UBound(parts) >= 1, Trim$(parts(1)), "1")
    Application.Run "GetLatestModuleVersion", targetWb, modName
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A4").Value = "Yerel" : ws.Range("B4").Value = localVer
    ws.Range("A5").Value = "Guncelleme" : ws.Range("B5").Value = "Sunucu yanitini karsilastirin"
    Set DynamicFunc = Nothing
End Function
