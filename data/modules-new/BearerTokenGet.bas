Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim url As String, token As String
    Dim parts() As String : parts = Split(CStr(param), "|", 2)
    url = Trim$(parts(0))
    token = IIf(UBound(parts) >= 1, Trim$(parts(1)), "")
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", url, False
    If Len(token) > 0 Then http.setRequestHeader "Authorization", "Bearer " & token
    http.setTimeouts 5000, 10000, 30000, 30000
    http.send
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "URL" : ws.Range("B1").Value = url
    ws.Range("A2").Value = "Durum" : ws.Range("B2").Value = http.Status
    ws.Range("A3").Value = "Yanit" : ws.Range("B3").Value = Left$(http.responseText, 32000)
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function
