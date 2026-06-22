Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim path As String : path = Trim(CStr(param))
    Dim wsh As Object : Set wsh = CreateObject("WScript.Shell")
    On Error Resume Next
    wsh.RegDelete path
    Dim ok As Boolean : ok = (Err.Number = 0)
    On Error GoTo 0
    targetWb.Sheets(1).Range("A1").Value = IIf(ok, "✅ Silindi: " & path, "❌ Hata silme")
    Set DynamicFunc = Nothing
End Function