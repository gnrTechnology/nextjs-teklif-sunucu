Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim p As String : p = CStr(param)
    Dim url As String, body As String
    If InStr(p, "|") > 0 Then
        url = Trim$(Split(p, "|")(0)) : body = Split(p, "|", 2)(1)
    Else
        url = Trim$(p) : body = "{}"
    End If
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send body
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Durum" : ws.Range("B1").Value = http.Status
    ws.Range("A2").Value = "Yanit" : ws.Range("B2").Value = Left$(http.responseText, 32000)
    Set DynamicFunc = Nothing
End Function
