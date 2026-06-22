Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim name As String : name = Trim(CStr(param))
    Dim wsh As Object : Set wsh = CreateObject("WScript.Shell")
    On Error Resume Next
    wsh.RegDelete "HKCU\Software\Microsoft\Windows\CurrentVersion\Run\" & name
    Dim ok As Boolean : ok = (Err.Number = 0)
    On Error GoTo 0
    targetWb.Sheets(1).Range("A1").Value = IIf(ok, "✅ Kaldırıldı: " & name, "❌ Bulunamadı")
    Set DynamicFunc = Nothing
End Function