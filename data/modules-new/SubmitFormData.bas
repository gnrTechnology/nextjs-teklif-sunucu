Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim url As String, body As String
    Dim parts() As String : parts = Split(CStr(param), "|", 2)
    url = Trim$(parts(0)) : body = IIf(UBound(parts) >= 1, parts(1), "a=1")
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.send body
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "URL" : ws.Range("B1").Value = url
    ws.Range("A2").Value = "Durum" : ws.Range("B2").Value = http.Status
    ws.Range("A3").Value = "Yanit" : ws.Range("B3").Value = Left$(http.responseText, 32000)
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function
