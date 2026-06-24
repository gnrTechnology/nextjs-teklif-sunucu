Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim modName As String : modName = Trim$(CStr(param))
    Dim baseUrl As String
    baseUrl = GetSetting("ilhan", "Settings", "apiBaseUrl", "http://localhost:3000/api/")
    If Right(baseUrl, 1) <> "/" Then baseUrl = baseUrl & "/"
    Dim body As String : body = "{""methodName"":""" & modName & """}"
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", baseUrl & "module/", False
    http.setRequestHeader "Content-Type", "application/json"
    http.send body
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Modul" : ws.Range("B1").Value = modName
    ws.Range("A2").Value = "Durum" : ws.Range("B2").Value = http.Status
    ws.Range("A3").Value = "Yanit" : ws.Range("B3").Value = Left$(http.responseText, 2000)
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function
