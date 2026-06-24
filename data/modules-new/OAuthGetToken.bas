Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim tokenUrl As String, cid As String, secret As String, scope As String
    Dim parts() As String : parts = Split(CStr(param), "|")
    tokenUrl = Trim$(parts(0))
    cid = IIf(UBound(parts) >= 1, Trim$(parts(1)), "")
    secret = IIf(UBound(parts) >= 2, Trim$(parts(2)), "")
    scope = IIf(UBound(parts) >= 3, Trim$(parts(3)), "")
    Dim body As String
    body = "grant_type=client_credentials&client_id=" & cid & "&client_secret=" & secret
    If Len(scope) > 0 Then body = body & "&scope=" & scope
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", tokenUrl, False
    http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.send body
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Durum" : ws.Range("B1").Value = http.Status
    ws.Range("A2").Value = "Token" : ws.Range("B2").Value = Left$(http.responseText, 32000)
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function
