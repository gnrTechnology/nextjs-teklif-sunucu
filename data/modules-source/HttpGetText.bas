Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim url As String : url = Trim(CStr(param))
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", url, False
    http.setTimeouts 5000, 10000, 30000, 30000
    http.send
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    If http.Status = 200 Then
        ws.Range("A1").Value = http.responseText
    Else
        ws.Range("A1").Value = "HTTP " & http.Status
    End If
    Set DynamicFunc = Nothing
End Function