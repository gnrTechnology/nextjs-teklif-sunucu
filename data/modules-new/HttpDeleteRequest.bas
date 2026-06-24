Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim url As String
    url = Trim$(CStr(param))
    If Len(url) = 0 Then url = "http://localhost:3000/api/"

    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "DELETE", url, False
    http.setTimeouts 5000, 10000, 30000, 30000
    http.send

    Dim ws As Worksheet
    Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "URL"
    ws.Range("B1").Value = url
    ws.Range("A2").Value = "Durum"
    ws.Range("B2").Value = http.Status
    ws.Range("A3").Value = "Yanit"
    ws.Range("B3").Value = Left$(http.responseText, 32000)
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function
