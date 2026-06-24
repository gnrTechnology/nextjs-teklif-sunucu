Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim mac As String : mac = GetSetting("ilhan", "Settings", "mac", "")
    Dim baseUrl As String : baseUrl = GetSetting("ilhan", "Settings", "apiBaseUrl", "http://localhost:3000/api/")
    If Right(baseUrl, 1) <> "/" Then baseUrl = baseUrl & "/"
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", baseUrl & "license/" & mac & "/", False
    http.send
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "MAC" : ws.Range("B1").Value = mac
    ws.Range("A2").Value = "Durum" : ws.Range("B2").Value = http.Status
    ws.Range("A3").Value = "Yanit" : ws.Range("B3").Value = Left$(http.responseText, 2000)
    Set DynamicFunc = Nothing
End Function
