Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim baseUrl As String
    baseUrl = GetSetting("ilhan", "Settings", "apiBaseUrl", "http://localhost:3000/api/")
    If Right(baseUrl, 1) <> "/" Then baseUrl = baseUrl & "/"
    Dim mac As String : mac = GetSetting("ilhan", "Settings", "mac", "")
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", baseUrl & "auto-start/" & mac & "/", False : http.send
    SaveSetting "ilhan", "RemoteConfig", "last", http.responseText
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1) : ws.Range("B1").Value = Left$(http.responseText, 32000)
    Set DynamicFunc = Nothing
End Function
