Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim url As String, user As String, pass As String
    ParseAuthParam CStr(param), url, user, pass
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", url, False
    http.setRequestHeader "Authorization", "Basic " & Base64Encode(user & ":" & pass)
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
Private Sub ParseAuthParam(p As String, ByRef url As String, ByRef user As String, ByRef pass As String)
    Dim parts() As String : parts = Split(p, "|")
    url = Trim$(parts(0)) : user = "" : pass = ""
    If UBound(parts) >= 1 Then user = Trim$(parts(1))
    If UBound(parts) >= 2 Then pass = Trim$(parts(2))
End Sub
Private Function Base64Encode(s As String) As String
    Dim dm As Object : Set dm = CreateObject("MSXML2.DOMDocument")
    Dim el As Object : Set el = dm.createElement("b64")
    el.DataType = "bin.base64"
    el.nodeTypedValue = StrConv(s, vbFromUnicode)
    Base64Encode = Replace(Replace(el.Text, vbCr, ""), vbLf, "")
End Function
