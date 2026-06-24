Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim webhook As String, msg As String
    Dim parts() As String : parts = Split(CStr(param), "|", 2)
    webhook = Trim$(parts(0)) : msg = IIf(UBound(parts) >= 1, parts(1), "Teklif bildirimi")
    Dim body As String
    body = "{""type"":""message"",""attachments"":[{""contentType"":""application/vnd.microsoft.card.adaptive"",""content"":{""type"":""AdaptiveCard"",""body"":[{""type"":""TextBlock"",""text"":""" & JsonEsc(msg) & """}]}}]}"
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", webhook, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send body
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "Durum" : ws.Range("B1").Value = http.Status
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function
Private Function JsonEsc(s As String) As String
    JsonEsc = Replace(Replace(CStr(s), Chr(92), Chr(92) & Chr(92)), Chr(34), Chr(92) & Chr(34))
End Function
