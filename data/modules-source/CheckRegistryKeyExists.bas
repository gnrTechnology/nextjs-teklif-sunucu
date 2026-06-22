Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim path As String : path = Trim(CStr(param))
    Dim wsh As Object : Set wsh = CreateObject("WScript.Shell")
    Dim exists As Boolean : exists = False
    On Error Resume Next
    Dim v As String : v = wsh.RegRead(path) : exists = (Err.Number = 0)
    On Error GoTo 0
    targetWb.Sheets(1).Range("A1").Value = IIf(exists, "✅ Mevcut: " & path, "❌ Bulunamadı: " & path)
    Set DynamicFunc = Nothing
End Function