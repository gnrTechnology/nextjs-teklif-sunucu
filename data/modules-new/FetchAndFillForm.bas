Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim url As String : url = Trim$(CStr(param))
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", url, False : http.send
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "API" : ws.Range("B1").Value = url
    ws.Range("A2").Value = "JSON" : ws.Range("B2").Value = Left$(http.responseText, 32000)
    Set DynamicFunc = Nothing
End Function
