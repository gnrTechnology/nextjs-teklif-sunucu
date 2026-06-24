Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Dim base As String : base = Trim$(CStr(param)) : If Len(base) = 0 Then base = "USD"
    Dim url As String : url = "https://api.exchangerate-api.com/v4/latest/" & base
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", url, False : http.send
    Dim resp As String : resp = http.responseText
    Dim ws As Worksheet : Set ws = targetWb.Sheets(1)
    ws.Cells.ClearContents
    ws.Range("A1").Value = "Para" : ws.Range("B1").Value = "Kur (1 " & base & ")"
    Dim cur() As String : cur = Split("TRY,EUR,GBP,CHF,JPY", ",")
    Dim i As Long, r As Long : r = 2
    For i = LBound(cur) To UBound(cur)
        ws.Cells(r, 1).Value = cur(i)
        ws.Cells(r, 2).Value = JsonNum(resp, cur(i))
        r = r + 1
    Next i
    ws.Range(r, 1).Value = "Tarih" : ws.Range(r, 2).Value = Now
    ws.Columns.AutoFit
    Set DynamicFunc = Nothing
End Function
Private Function JsonNum(json As String, key As String) As Double
    Dim sk As String : sk = Chr(34) & key & Chr(34) & ":"
    Dim p As Long : p = InStr(json, sk)
    If p = 0 Then Exit Function
    JsonNum = CDbl(Val(Mid$(json, p + Len(sk), 12)))
End Function
