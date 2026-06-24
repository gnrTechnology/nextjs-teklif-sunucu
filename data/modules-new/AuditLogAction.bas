Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim detail As String : detail = CStr(param)
    Dim mac As String : mac = GetSetting("ilhan", "Settings", "mac", "")
    Dim baseUrl As String : baseUrl = GetSetting("ilhan", "Settings", "apiBaseUrl", "http://localhost:3000/api/")
    If Right(baseUrl, 1) <> "/" Then baseUrl = baseUrl & "/"
    Dim body As String
    body = "{""eventType"":""audit"",""macAdresi"":""" & mac & """,""detail"":""" & Replace(detail, Chr(34), "'") & """}"
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", baseUrl & "activity/", False
    http.setRequestHeader "Content-Type", "application/json"
    http.send body
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Log" : ws.Range("B1").Value = http.Status
    Set DynamicFunc = Nothing
End Function
