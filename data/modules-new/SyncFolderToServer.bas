Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim folderPath As String, apiUrl As String
    Dim parts() As String : parts = Split(CStr(param), "|", 2)
    folderPath = Trim$(parts(0))
    apiUrl = IIf(UBound(parts) >= 1, Trim$(parts(1)), GetSetting("ilhan", "Settings", "apiBaseUrl", "") & "module-output/")
    Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    Dim folder As Object : Set folder = fso.GetFolder(folderPath)
    Dim f As Object, n As Long : n = 0
    For Each f In folder.Files
        n = n + 1
    Next f
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Klasor" : ws.Range("B1").Value = folderPath
    ws.Range("A2").Value = "Dosya sayisi" : ws.Range("B2").Value = n
    ws.Range("A3").Value = "API" : ws.Range("B3").Value = apiUrl
    Set DynamicFunc = Nothing
End Function
