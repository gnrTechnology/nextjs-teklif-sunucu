Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim url As String : url = Trim(CStr(param))
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    On Error Resume Next
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "HEAD", url, False
    http.setTimeouts 3000, 5000, 10000, 10000
    http.send
    Dim status As Long : status = http.Status
    On Error GoTo 0
    ws.Range("A1").Value = "URL"    : ws.Range("B1").Value = url
    ws.Range("A2").Value = "Durum"  : ws.Range("B2").Value = IIf(status >= 200 And status < 400, "✅ Erişilebilir (HTTP " & status & ")", "❌ Erişilemiyor (" & status & ")")
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function