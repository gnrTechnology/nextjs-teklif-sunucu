Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim url As String : url = Trim(CStr(param))
    If url = "" Then url = "https://www.google.com"
    CreateObject("WScript.Shell").Run url
    targetWb.Sheets(1).Range("A1").Value = "✅ Açıldı: " & url
    Set DynamicFunc = Nothing
End Function