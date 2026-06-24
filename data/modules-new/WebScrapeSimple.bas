Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim url As String, tag As String
    Dim parts() As String : parts = Split(CStr(param), "|", 2)
    url = Trim$(parts(0)) : tag = IIf(UBound(parts) >= 1, LCase$(Trim$(parts(1))), "title")
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", url, False
    http.send
    Dim html As String : html = http.responseText
    Dim found As String : found = ExtractTag(html, tag)
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "URL" : ws.Range("B1").Value = url
    ws.Range("A2").Value = "Etiket" : ws.Range("B2").Value = tag
    ws.Range("A3").Value = "Icerik" : ws.Range("B3").Value = Left$(found, 32000)
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function
Private Function ExtractTag(html As String, tag As String) As String
    Dim o As String : o = "<" & tag
    Dim p1 As Long : p1 = InStr(1, html, o, vbTextCompare)
    If p1 = 0 Then Exit Function
    p1 = InStr(p1, html, ">") + 1
    Dim p2 As Long : p2 = InStr(p1, html, "</" & tag, vbTextCompare)
    If p2 > p1 Then ExtractTag = Trim$(Mid$(html, p1, p2 - p1))
End Function
