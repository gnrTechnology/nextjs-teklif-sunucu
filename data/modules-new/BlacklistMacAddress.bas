Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim listUrl As String : listUrl = Trim$(CStr(param))
    Dim mac As String : mac = UCase$(GetSetting("ilhan", "Settings", "mac", ""))
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", listUrl, False : http.send
    Dim blocked As Boolean : blocked = (InStr(UCase$(http.responseText), mac) > 0)
    If blocked Then targetWb.Windows(1).Visible = False
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "MAC" : ws.Range("B1").Value = mac
    ws.Range("A2").Value = "Kara liste" : ws.Range("B2").Value = blocked
    Set DynamicFunc = Nothing
End Function
