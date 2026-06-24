Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim url As String : url = Trim$(CStr(param))
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", url, False
    http.setTimeouts 5000, 10000, 30000, 30000
    http.Option(6) = True
    http.send
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Baslangic" : ws.Range("B1").Value = url
    ws.Range("A2").Value = "Son URL" : ws.Range("B2").Value = http.getResponseHeader("Location")
    If Len(ws.Range("B2").Value) = 0 Then ws.Range("B2").Value = url
    ws.Range("A3").Value = "Durum" : ws.Range("B3").Value = http.Status
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function
